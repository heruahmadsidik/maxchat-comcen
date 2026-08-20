require('dotenv').config();
const express = require('express');
const axios   = require('axios');
const fs      = require('fs');
const path    = require('path');
const compression = require('compression');
const session = require('express-session');
const XLSX    = require('xlsx');
const app     = express();

app.use(compression());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'nikmatmanalagiyangkamudustakan', // ⚠️ sebaiknya diisi lewat .env
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 1000 * 60 * 60 * 12 } // idle timeout: 12 jam
}));

// Folder "public" (berisi logo.png dsb) — dilayani SEBELUM middleware login,
// supaya logo tetap tampil di halaman /login (sebelum user login).
app.use(express.static(path.join(__dirname, 'public')));

// Proteksi semua route KECUALI /login dan /webhook (webhook dipanggil Maxchat server, bukan browser)
function requireLogin(req, res, next) {
  if (req.path === '/login' || req.path === '/webhook') return next();
  if (req.session && req.session.loggedIn) return next();
  if (req.method === 'GET' && !req.path.startsWith('/api')) {
    return res.redirect('/login');
  }
  return res.status(401).json({ error: 'Sesi habis, silakan login kembali' });
}
app.use(requireLogin);

// Cek password konfirmasi tambahan sebelum menyimpan perubahan pengaturan sensitif
// (API token, webhook URL, username/password login) — mencegah operator biasa
// iseng mengubah pengaturan walau sudah login.
function checkSettingsPassword(req, res, cfg) {
  const { confirmPassword } = req.body;
  if (!cfg.settingsPassword || !confirmPassword || confirmPassword !== cfg.settingsPassword) {
    res.status(403).json({ error: 'Password konfirmasi salah atau belum diisi' });
    return false;
  }
  return true;
}

const DELAY_MS  = 1000;
const TPL_FILE  = path.join(__dirname, 'templates.json');
const CFG_FILE  = path.join(__dirname, 'config.json');
const DB_FILE   = path.join(__dirname, 'storage', 'history.db');

// ============================================================
// PROGRESS TRACKING BLAST (live progress bar di frontend)
// ============================================================
// Disimpan in-memory saja (cukup untuk 1 instance server). Key = blastId
// yang dibuat frontend saat mulai kirim, dipoll tiap ~400ms selama proses
// pengiriman berjalan di server (karena ada jeda DELAY_MS antar nomor).
const blastProgress = new Map(); // blastId -> { total, sent, success, failed, done }
function setBlastProgress(id, patch) {
  if (!id) return;
  const cur = blastProgress.get(id) || { total: 0, sent: 0, success: 0, failed: 0, done: false };
  blastProgress.set(id, { ...cur, ...patch });
}

if (!fs.existsSync(path.join(__dirname, 'storage'))) fs.mkdirSync(path.join(__dirname, 'storage'));

// ============================================================
// DATABASE SQLITE
// ============================================================
const Database = require('better-sqlite3');
let db;
const stmt = {};   

const CONTACT_SUMMARY_SQL = `
  WITH last_msg AS (
    SELECT c.phone, c.time, c.status, c.msgType, c.message,
           s.templateName, s.previewText,
           ROW_NUMBER() OVER (PARTITION BY c.phone ORDER BY c.time DESC, c.id DESC) AS rn
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
  ),
  last_name AS (
    SELECT phone, name,
           ROW_NUMBER() OVER (PARTITION BY phone ORDER BY time DESC, id DESC) AS rn
    FROM contacts WHERE name != ''
  ),
  last_non_umum_cat AS (
    SELECT c.phone, s.woCategory,
           ROW_NUMBER() OVER (PARTITION BY c.phone ORDER BY c.time DESC, c.id DESC) AS rn
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE s.woCategory IS NOT NULL AND s.woCategory != 'umum'
  )
  SELECT
    lm.phone,
    COALESCE(ln.name, '')       AS name,
    COALESCE(cc.woCategory, lnuc.woCategory, 'umum') AS woCategory,
    lm.time                     AS lastTime,
    CASE WHEN lm.templateName = '📥 PESAN MASUK' THEN lm.message ELSE lm.previewText END AS lastText,
    CASE WHEN lm.templateName = '📥 PESAN MASUK' THEN 'in' ELSE 'out' END AS lastDirection,
    lm.status                   AS lastStatus,
    lm.msgType                  AS lastMsgType,
    CASE WHEN lm.templateName = '📥 PESAN MASUK' AND (rs.lastReadTime IS NULL OR lm.time > rs.lastReadTime)
         THEN 1 ELSE 0 END      AS isUnread
  FROM last_msg lm
  LEFT JOIN last_name ln  ON ln.phone = lm.phone AND ln.rn = 1
  LEFT JOIN contact_category cc ON cc.phone = lm.phone
  LEFT JOIN last_non_umum_cat lnuc ON lnuc.phone = lm.phone AND lnuc.rn = 1
  LEFT JOIN read_state rs ON rs.phone = lm.phone
  WHERE lm.rn = 1
`;

function initDb() {
  db = new Database(DB_FILE);
  db.pragma('journal_mode = WAL');      
  db.pragma('foreign_keys = ON');
  db.pragma('synchronous = NORMAL');    

  db.exec(`
    CREATE TABLE IF NOT EXISTS sessions (
      id           TEXT PRIMARY KEY,
      time         TEXT NOT NULL,
      templateName TEXT,
      previewText  TEXT,
      woCategory   TEXT DEFAULT 'umum',
      total        INTEGER DEFAULT 0,
      success      INTEGER DEFAULT 0,
      failed       INTEGER DEFAULT 0
    )
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS contacts (
      id            INTEGER PRIMARY KEY AUTOINCREMENT,
      sessionId     TEXT NOT NULL,
      phone         TEXT NOT NULL,
      name          TEXT DEFAULT '',
      status        TEXT DEFAULT 'success',
      message       TEXT,
      messageId     TEXT DEFAULT '',
      nomerLapor    TEXT DEFAULT '',
      msgType       TEXT DEFAULT 'text',
      attachmentUrl TEXT DEFAULT '',
      time          TEXT NOT NULL,
      FOREIGN KEY (sessionId) REFERENCES sessions(id)
    )
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS read_state (
      phone        TEXT PRIMARY KEY,
      lastReadTime TEXT NOT NULL
    )
  `);
  db.exec(`
    CREATE TABLE IF NOT EXISTS contact_category (
      phone      TEXT PRIMARY KEY,
      woCategory TEXT NOT NULL DEFAULT 'umum',
      updatedAt  TEXT NOT NULL
    )
  `);
  // Tabel ringkasan: 1 baris per nomor kontak, selalu berisi kondisi TERBARU.
  // Dibuat supaya menu chat / filter kategori / hitung unread TIDAK PERLU lagi
  // scan seluruh riwayat pesan (contacts+sessions) setiap kali diminta.
  db.exec(`
    CREATE TABLE IF NOT EXISTS contact_summary (
      phone         TEXT PRIMARY KEY,
      name          TEXT DEFAULT '',
      woCategory    TEXT DEFAULT 'umum',
      lastTime      TEXT NOT NULL,
      lastText      TEXT DEFAULT '',
      lastDirection TEXT DEFAULT 'out',
      lastStatus    TEXT DEFAULT '',
      lastMsgType   TEXT DEFAULT 'text',
      isUnread      INTEGER DEFAULT 0
    )
  `);

  // ------------------------------------------------------------
  // KONTAK & LABEL — menu "Kontak": daftar semua nomor yang pernah dikirimi
  // pesan (WO Marking maupun Individu digabung jadi satu daftar), dengan
  // label bebas berupa teks (multi-label per kontak, tanpa perlu bikin daftar
  // label global dulu — tinggal ketik langsung di kontaknya), dan nama kontak
  // yang bisa diisi manual kalau belum ada dari riwayat chat.
  // ------------------------------------------------------------
  db.exec(`
    CREATE TABLE IF NOT EXISTS contact_tags (
      phone TEXT NOT NULL,
      tag   TEXT NOT NULL,
      PRIMARY KEY (phone, tag)
    )
  `);
  // Daftar MASTER label — label harus dibuat dulu di sini (lewat "Kelola
  // Label") sebelum bisa ditempelkan ke kontak manapun. Beda dari versi lama
  // yang bebas ketik teks apa saja langsung di kontak.
  db.exec(`
    CREATE TABLE IF NOT EXISTS labels (
      id        TEXT PRIMARY KEY,
      name      TEXT NOT NULL UNIQUE,
      color     TEXT DEFAULT '#2563eb',
      order_num INTEGER DEFAULT 0,
      createdAt TEXT
    )
  `);
  // Jaga-jaga untuk database lama yang tabel labels-nya sudah ada TANPA
// kolom createdAt (dibuat sebelum fitur ini ada) — tambahkan kolomnya,
// abaikan error kalau ternyata sudah ada.
  try { db.exec(`ALTER TABLE labels ADD COLUMN createdAt TEXT`); } catch (e) { /* kolom sudah ada */ }
  db.exec(`
    CREATE TABLE IF NOT EXISTS contact_meta (
      phone         TEXT PRIMARY KEY,
      customName    TEXT DEFAULT '',
      updatedAt     TEXT,
      manuallyAdded INTEGER DEFAULT 0
    )
  `);
  // Jaga-jaga untuk database lama yang tabel contact_meta-nya sudah ada TANPA
  // kolom manuallyAdded (dibuat sebelum fitur "Tambah Kontak" ini ada) — coba
  // tambahkan kolomnya, abaikan error kalau ternyata sudah ada.
  try { db.exec(`ALTER TABLE contact_meta ADD COLUMN manuallyAdded INTEGER DEFAULT 0`); } catch (e) { /* kolom sudah ada */ }
  db.exec(`CREATE INDEX IF NOT EXISTS idx_contact_tags_tag ON contact_tags(tag)`);

  // Tabel Quick Reply: daftar balasan cepat yang bisa dipilih dari menu chat,
  // dikelola (tambah/edit/hapus/urutkan) dari menu Pengaturan.
  db.exec(`
    CREATE TABLE IF NOT EXISTS quick_replies (
      id        TEXT PRIMARY KEY,
      text      TEXT NOT NULL,
      order_num INTEGER DEFAULT 0
    )
  `);

  db.exec(`CREATE INDEX IF NOT EXISTS idx_contacts_phone   ON contacts(phone)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_contacts_session ON contacts(sessionId)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_sessions_time    ON sessions(time DESC)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_contacts_phone_time ON contacts(phone, time DESC)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_summary_time      ON contact_summary(lastTime DESC)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_summary_unread    ON contact_summary(isUnread, lastTime DESC)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_summary_category  ON contact_summary(woCategory, lastTime DESC)`);
  db.exec(`CREATE INDEX IF NOT EXISTS idx_quick_replies_order ON quick_replies(order_num ASC)`);

  // ------------------------------------------------------------
  // FTS5: index pencarian teks (nama kontak, isi pesan, preview template)
  // supaya kolom pencarian tidak perlu lagi LIKE '%...%' menyisir seluruh
  // tabel contacts+sessions. rowid FTS = contacts.id, disinkronkan otomatis
  // lewat trigger setiap ada INSERT/UPDATE/DELETE di tabel contacts.
  // Nomor telepon TIDAK dimasukkan ke FTS (tokenizer FTS memecah per kata,
  // kurang cocok untuk pencarian potongan angka) — pencarian nomor tetap
  // pakai LIKE seperti sebelumnya, tapi sekarang hanya menyisir kolom
  // phone saja (jauh lebih ringan daripada join+scan message/previewText).
  // Kalau modul FTS5 ternyata tidak tersedia di build better-sqlite3 yang
  // dipakai, aplikasi otomatis fallback ke pencarian LIKE lama (lihat
  // ftsAvailable di endpoint /api/contacts/search).
  let ftsAvailable = false;
  try {
    db.exec(`
      CREATE VIRTUAL TABLE IF NOT EXISTS contacts_fts USING fts5(
        name, message, previewText,
        tokenize = 'unicode61 remove_diacritics 2'
      )
    `);
    db.exec(`
      CREATE TRIGGER IF NOT EXISTS contacts_fts_ai AFTER INSERT ON contacts BEGIN
        INSERT INTO contacts_fts(rowid, name, message, previewText)
        VALUES (
          new.id,
          new.name,
          new.message,
          (SELECT previewText FROM sessions WHERE id = new.sessionId)
        );
      END;
    `);
    db.exec(`
      CREATE TRIGGER IF NOT EXISTS contacts_fts_ad AFTER DELETE ON contacts BEGIN
        DELETE FROM contacts_fts WHERE rowid = old.id;
      END;
    `);
    db.exec(`
      CREATE TRIGGER IF NOT EXISTS contacts_fts_au AFTER UPDATE ON contacts BEGIN
        UPDATE contacts_fts SET name = new.name, message = new.message WHERE rowid = new.id;
      END;
    `);
    ftsAvailable = true;
  } catch (err) {
    console.error('⚠️  FTS5 tidak tersedia, pencarian akan pakai LIKE biasa (lebih lambat):', err.message);
  }
  db.ftsAvailable = ftsAvailable;

  stmt.count                = db.prepare('SELECT COUNT(*) AS c FROM sessions');
  stmt.latestTime           = db.prepare('SELECT MAX(time) AS t FROM sessions');
  stmt.loadSessions         = db.prepare(`SELECT id, time, templateName, previewText, woCategory, total, success, failed FROM sessions ORDER BY time DESC`);
  stmt.checkSessionExists   = db.prepare('SELECT 1 FROM sessions WHERE id = ?');
  stmt.insertSession        = db.prepare(`INSERT OR REPLACE INTO sessions (id, time, templateName, previewText, woCategory, total, success, failed) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`);
  stmt.insertContact        = db.prepare(`INSERT INTO contacts (sessionId, phone, name, status, message, messageId, nomerLapor, msgType, attachmentUrl, time) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
  stmt.updateContact        = db.prepare(`UPDATE contacts SET status = ? WHERE messageId = ? AND messageId != ''`);
  stmt.deleteContactsBySess = db.prepare(`DELETE FROM contacts WHERE sessionId = ?`);
  stmt.deleteSession        = db.prepare(`DELETE FROM sessions WHERE id = ?`);
  stmt.deleteAllContacts    = db.prepare(`DELETE FROM contacts`);
  stmt.deleteAllSessions    = db.prepare(`DELETE FROM sessions`);
  stmt.searchSessions       = db.prepare(`SELECT DISTINCT sessionId AS sid FROM contacts WHERE phone LIKE ? OR message LIKE ? OR name LIKE ?`);

  stmt.markRead = db.prepare(`
    INSERT INTO read_state (phone, lastReadTime) VALUES (?, ?)
    ON CONFLICT(phone) DO UPDATE SET lastReadTime = excluded.lastReadTime
  `);
  stmt.latestMsgTimeForPhone = db.prepare(`SELECT MAX(time) AS t FROM contacts WHERE phone = ?`);
  stmt.getReadState          = db.prepare(`SELECT lastReadTime FROM read_state WHERE phone = ?`);

  // Hitung unread langsung dari tabel ringkasan (indexed, instan) —
  // menggantikan query lama yang window-function full-scan tiap 3 detik polling.
  stmt.summaryUnreadCount = db.prepare(`SELECT COUNT(*) AS c FROM contact_summary WHERE isUnread = 1`);

  // Statement untuk menghitung ulang 1 baris ringkasan per nomor, dipakai
  // setiap ada pesan baru / kontak dihapus / kategori berubah. Semua pakai
  // index (phone, time DESC) jadi tetap cepat walau riwayat sudah jutaan baris.
  stmt.lastMsgForPhone = db.prepare(`
    SELECT c.time, c.status, c.msgType, c.message, s.templateName, s.previewText
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE c.phone = ?
    ORDER BY c.time DESC, c.id DESC
    LIMIT 1
  `);
  stmt.lastNameForPhone = db.prepare(`
    SELECT name FROM contacts WHERE phone = ? AND name != '' ORDER BY time DESC, id DESC LIMIT 1
  `);
  stmt.lastNonUmumCatForPhone = db.prepare(`
    SELECT s.woCategory
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE c.phone = ? AND s.woCategory IS NOT NULL AND s.woCategory != 'umum'
    ORDER BY c.time DESC, c.id DESC LIMIT 1
  `);
  stmt.deleteSummaryForPhone = db.prepare(`DELETE FROM contact_summary WHERE phone = ?`);
  stmt.replaceSummaryRow = db.prepare(`
    INSERT INTO contact_summary (phone, name, woCategory, lastTime, lastText, lastDirection, lastStatus, lastMsgType, isUnread)
    VALUES (@phone, @name, @woCategory, @lastTime, @lastText, @lastDirection, @lastStatus, @lastMsgType, @isUnread)
    ON CONFLICT(phone) DO UPDATE SET
      name          = @name,
      woCategory    = @woCategory,
      lastTime      = @lastTime,
      lastText      = @lastText,
      lastDirection = @lastDirection,
      lastStatus    = @lastStatus,
      lastMsgType   = @lastMsgType,
      isUnread      = @isUnread
  `);

  stmt.contactMessagesPage = db.prepare(`
    SELECT c.phone, c.name, c.status, c.message, c.messageId, c.nomerLapor,
           c.msgType, c.attachmentUrl, c.time,
           s.templateName, s.previewText
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE c.phone = ? AND c.time < ?
    ORDER BY c.time DESC
    LIMIT ?
  `);
  stmt.contactMessagesLatest = db.prepare(`
    SELECT c.phone, c.name, c.status, c.message, c.messageId, c.nomerLapor,
           c.msgType, c.attachmentUrl, c.time,
           s.templateName, s.previewText
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE c.phone = ?
    ORDER BY c.time DESC
    LIMIT ?
  `);
  stmt.contactMessagesAfter = db.prepare(`
    SELECT c.phone, c.name, c.status, c.message, c.messageId, c.nomerLapor,
           c.msgType, c.attachmentUrl, c.time,
           s.templateName, s.previewText
    FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE c.phone = ? AND c.time > ?
    ORDER BY c.time ASC
  `);

  stmt.setContactCategory = db.prepare(`
    INSERT INTO contact_category (phone, woCategory, updatedAt) VALUES (?, ?, ?)
    ON CONFLICT(phone) DO UPDATE SET woCategory = excluded.woCategory, updatedAt = excluded.updatedAt
  `);
  stmt.getContactCategory = db.prepare(`SELECT woCategory FROM contact_category WHERE phone = ?`);

  stmt.findSessionsByRange  = db.prepare(`SELECT id FROM sessions WHERE time >= ? AND time <= ?`);
  stmt.deleteContactsByRange = db.prepare(`DELETE FROM contacts WHERE sessionId IN (SELECT id FROM sessions WHERE time >= ? AND time <= ?)`);
  stmt.deleteSessionsByRange = db.prepare(`DELETE FROM sessions WHERE time >= ? AND time <= ?`);

  // Quick Reply — daftar balasan cepat.
  stmt.loadQuickReplies     = db.prepare(`SELECT id, text, order_num AS orderNum FROM quick_replies ORDER BY order_num ASC, rowid ASC`);
  stmt.maxQuickReplyOrder   = db.prepare(`SELECT MAX(order_num) AS m FROM quick_replies`);
  stmt.insertQuickReply     = db.prepare(`INSERT INTO quick_replies (id, text, order_num) VALUES (?, ?, ?)`);
  stmt.updateQuickReplyText = db.prepare(`UPDATE quick_replies SET text = ? WHERE id = ?`);
  stmt.updateQuickReplyOrd  = db.prepare(`UPDATE quick_replies SET order_num = ? WHERE id = ?`);
  stmt.deleteQuickReply     = db.prepare(`DELETE FROM quick_replies WHERE id = ?`);
  stmt.checkQuickReplyExists = db.prepare(`SELECT 1 FROM quick_replies WHERE id = ?`);

  // Kontak & Label (label bebas berupa teks, ditempel langsung ke kontak)
  stmt.loadDistinctTags       = db.prepare(`SELECT DISTINCT tag FROM contact_tags ORDER BY tag COLLATE NOCASE ASC`);
  stmt.attachContactTag       = db.prepare(`INSERT OR IGNORE INTO contact_tags (phone, tag) VALUES (?, ?)`);
  stmt.detachContactTag       = db.prepare(`DELETE FROM contact_tags WHERE phone = ? AND tag = ?`);
  stmt.deleteContactTagsByPhone = db.prepare(`DELETE FROM contact_tags WHERE phone = ?`);
  stmt.loadContactTagsByPhone = db.prepare(`SELECT tag FROM contact_tags WHERE phone = ? ORDER BY tag COLLATE NOCASE ASC`);

  // Master label — dibuat manual dulu, baru bisa ditempelkan ke kontak.
  stmt.loadLabels             = db.prepare(`SELECT * FROM labels ORDER BY order_num ASC, name COLLATE NOCASE ASC`);
  stmt.getLabelById           = db.prepare(`SELECT * FROM labels WHERE id = ?`);
  stmt.findLabelByName        = db.prepare(`SELECT * FROM labels WHERE name = ? COLLATE NOCASE`);
  stmt.insertLabel            = db.prepare(`INSERT INTO labels (id, name, color, order_num, createdAt) VALUES (?, ?, ?, ?, ?)`);
  stmt.updateLabel            = db.prepare(`UPDATE labels SET name = ?, color = ? WHERE id = ?`);
  stmt.deleteLabelRow         = db.prepare(`DELETE FROM labels WHERE id = ?`);
  stmt.renameContactTagUsage  = db.prepare(`UPDATE contact_tags SET tag = ? WHERE tag = ?`);
  stmt.deleteContactTagsByName = db.prepare(`DELETE FROM contact_tags WHERE tag = ?`);
  stmt.getContactMeta         = db.prepare(`SELECT customName FROM contact_meta WHERE phone = ?`);
  stmt.setContactMeta         = db.prepare(`
    INSERT INTO contact_meta (phone, customName, updatedAt) VALUES (?, ?, ?)
    ON CONFLICT(phone) DO UPDATE SET customName = excluded.customName, updatedAt = excluded.updatedAt
  `);
  stmt.insertManualContact    = db.prepare(`
    INSERT INTO contact_meta (phone, customName, updatedAt, manuallyAdded) VALUES (?, ?, ?, 1)
    ON CONFLICT(phone) DO UPDATE SET customName = excluded.customName, updatedAt = excluded.updatedAt, manuallyAdded = 1
  `);
  stmt.deleteManualContact    = db.prepare(`DELETE FROM contact_meta WHERE phone = ? AND manuallyAdded = 1`);
  stmt.checkContactHasHistory = db.prepare(`SELECT 1 FROM contact_summary WHERE phone = ?`);

  const legacyFile = path.join(__dirname, 'storage', 'history.json');
  if (fs.existsSync(legacyFile)) {
    try {
      const legacy = JSON.parse(fs.readFileSync(legacyFile, 'utf8'));
      let migrated = 0;

      const migrateAll = db.transaction((items) => {
        for (const sess of items) {
          if (stmt.checkSessionExists.get(sess.id)) continue;

          stmt.insertSession.run(
            sess.id,
            sess.time,
            sess.templateName || '',
            sess.previewText  || '',
            sess.woCategory   || 'umum',
            sess.summary?.total   || 0,
            sess.summary?.success || 0,
            sess.summary?.failed  || 0
          );
          for (const c of (sess.contacts || [])) {
            stmt.insertContact.run(
              sess.id,
              c.phone         || '',
              c.name          || '',
              c.status        || 'success',
              c.message       || '',
              c.messageId     || '',
              c.nomerLapor    || '',
              c.msgType       || 'text',
              c.attachmentUrl || '',
              c.time          || sess.time
            );
          }
          migrated++;
        }
      });
      migrateAll(legacy);

      fs.renameSync(legacyFile, legacyFile + '.migrated');
      console.log(`✅ Migrasi history.json selesai: ${migrated} sesi diimpor. File lama → history.json.migrated`);
    } catch (err) {
      console.error('⚠️  Migrasi history.json gagal (data asli tidak diubah):', err.message);
    }
  }

  const missingRows = db.prepare(`
    SELECT phone, woCategory FROM (
      SELECT c.phone, s.woCategory, c.time, c.id,
             ROW_NUMBER() OVER (PARTITION BY c.phone ORDER BY c.time DESC, c.id DESC) AS rn
      FROM contacts c JOIN sessions s ON c.sessionId = s.id
      WHERE s.woCategory IS NOT NULL AND s.woCategory != 'umum'
        AND c.phone NOT IN (SELECT phone FROM contact_category)
    ) WHERE rn = 1
  `).all();
  if (missingRows.length) {
    const seed = db.transaction((items) => {
      for (const r of items) stmt.setContactCategory.run(r.phone, r.woCategory, new Date().toISOString());
    });
    seed(missingRows);
    console.log(`✅ Re-backfill contact_category: ${missingRows.length} nomor diproses`);
  }

  backfillContactSummary();
  backfillContactsFts();
}

function historyCount() {
  return stmt.count.get().c || 0;
}

// ============================================================
// BACKUP OTOMATIS DATABASE
// ============================================================
// Backup penuh database dibuat otomatis setiap hari jam 04:00 WIB (dini hari,
// trafik paling sepi) memakai Online Backup API bawaan SQLite lewat
// better-sqlite3 (db.backup()) — aman dijalankan sambil server tetap jalan,
// tidak perlu mematikan aplikasi ataupun mengunci database. Backup lama
// otomatis dibersihkan, hanya BACKUP_KEEP_COUNT file terbaru yang disimpan
// supaya folder storage tidak membengkak tanpa batas.
const BACKUP_DIR        = path.join(__dirname, 'storage', 'backups');
const BACKUP_KEEP_COUNT = 2;  // simpan 7 backup terakhir (≈ 1 minggu kalau harian), sisanya (file terlama) otomatis dihapus
const BACKUP_HOUR_WIB   = 4;  // jam 04:00 WIB

function formatBytes(bytes) {
  if (!bytes || bytes < 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let i = 0, val = bytes;
  while (val >= 1024 && i < units.length - 1) { val /= 1024; i++; }
  return (i === 0 ? val : val.toFixed(val < 10 ? 2 : 1)) + ' ' + units[i];
}

// Indonesia (WIB) tidak mengenal daylight saving, jadi cukup offset tetap
// +7 jam dari UTC — tidak perlu library timezone tambahan.
function getWIBNow() { return new Date(Date.now() + 7 * 60 * 60 * 1000); }
function pad2(n) { return String(n).padStart(2, '0'); }
function wibDateStr(d) { return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()); }

function listBackupFiles() {
  if (!fs.existsSync(BACKUP_DIR)) return [];
  return fs.readdirSync(BACKUP_DIR)
    .filter(f => f.endsWith('.db'))
    .map(f => {
      const full = path.join(BACKUP_DIR, f);
      const st = fs.statSync(full);
      return { name: f, sizeBytes: st.size, mtime: st.mtime.toISOString() };
    })
    .sort((a, b) => a.name.localeCompare(b.name)); // nama file diawali tanggal → urut nama = urut waktu
}

function backupAlreadyDoneForDate(dateStr) {
  return listBackupFiles().some(f => f.name.includes(dateStr));
}

async function runDatabaseBackup(trigger) {
  if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });

  const wibNow   = getWIBNow();
  const dateStr  = wibDateStr(wibNow);
  const timeStr  = pad2(wibNow.getUTCHours()) + pad2(wibNow.getUTCMinutes()) + pad2(wibNow.getUTCSeconds());
  const fileName = `history-${dateStr}_${timeStr}.db`;
  const destPath = path.join(BACKUP_DIR, fileName);

  await db.backup(destPath);

  // Simpan hanya BACKUP_KEEP_COUNT backup terbaru, hapus sisanya.
  const files  = listBackupFiles();
  const excess = files.length - BACKUP_KEEP_COUNT;
  if (excess > 0) {
    files.slice(0, excess).forEach(f => {
      try { fs.unlinkSync(path.join(BACKUP_DIR, f.name)); } catch (e) {}
    });
  }

  console.log(`✅ Backup database (${trigger}) selesai: ${fileName}`);
  return fileName;
}

// Dicek setiap menit — begitu masuk jam 04:00–04:04 WIB dan belum ada backup
// untuk tanggal (WIB) hari ini, backup dijalankan sekali. Jendela 5 menit
// dipakai jaga-jaga andai interval sempat telat atau server baru saja restart.
let backupCheckInFlight = false;
async function maybeRunScheduledBackup() {
  if (backupCheckInFlight) return;
  const wibNow = getWIBNow();
  const hour   = wibNow.getUTCHours();
  const minute = wibNow.getUTCMinutes();
  if (hour !== BACKUP_HOUR_WIB || minute >= 5) return;

  const dateStr = wibDateStr(wibNow);
  if (backupAlreadyDoneForDate(dateStr)) return;

  backupCheckInFlight = true;
  try {
    await runDatabaseBackup('otomatis');
  } catch (err) {
    console.error('❌ Backup otomatis gagal:', err.message);
  } finally {
    backupCheckInFlight = false;
  }
}

// Template/sesi yang BUKAN broadcast (balasan manual, file, sync pesan keluar,
// pesan masuk webhook) — dipakai untuk mengecualikan balasan dari hitungan
// "Total Broadcast" di berbagai tempat pada dashboard, supaya angka yang
// ditampilkan murni mencerminkan pesan blast/template, bukan balasan chat.
const NON_BROADCAST_TEMPLATES = ['📥 PESAN MASUK', '💬 Balasan Manual', '📎 File Terkirim', '📤 PESAN KELUAR (SYNC)', 'Balasan Manual'];
const NON_BROADCAST_PLACEHOLDERS = NON_BROADCAST_TEMPLATES.map(() => '?').join(',');

function getDbInfo() {
  let dbSizeBytes = 0;
  // Ukuran database dihitung dari file utama + WAL/SHM (mode journal WAL
  // dipakai di sini) supaya angka yang ditampilkan mencerminkan pemakaian
  // disk yang sebenarnya, bukan cuma file .db utama.
  [DB_FILE, DB_FILE + '-wal', DB_FILE + '-shm'].forEach(f => {
    try { dbSizeBytes += fs.statSync(f).size; } catch (e) {}
  });

  const sessionsCount   = stmt.count.get().c || 0;
  // Total Broadcast = jumlah pesan yang benar-benar dikirim lewat blast
  // template, TIDAK termasuk balasan manual/file/sync/pesan masuk.
  const broadcastCount  = db.prepare(`
    SELECT COUNT(*) AS c
    FROM contacts c
    JOIN sessions s ON s.id = c.sessionId
    WHERE s.templateName NOT IN (${NON_BROADCAST_PLACEHOLDERS})
  `).get(...NON_BROADCAST_TEMPLATES).c || 0;

  const backups     = listBackupFiles().reverse(); // terbaru duluan
  const lastBackup  = backups[0] || null;

  return {
    dbSizeBytes,
    dbSizeFormatted: formatBytes(dbSizeBytes),
    sessionsCount,
    broadcastCount,
    backupScheduleHour: BACKUP_HOUR_WIB,
    backupKeepCount: BACKUP_KEEP_COUNT,
    backupCount: backups.length,
    backups: backups.slice(0, 10).map(b => ({ name: b.name, sizeFormatted: formatBytes(b.sizeBytes), mtime: b.mtime })),
    lastBackup: lastBackup ? { name: lastBackup.name, sizeFormatted: formatBytes(lastBackup.sizeBytes), mtime: lastBackup.mtime } : null,
  };
}

// ============================================================
// DASHBOARD — statistik riwayat broadcast, nomor sering lapor, trend WO
// ============================================================
// Semua pengelompokan tanggal memakai WIB (UTC+7), konsisten dengan
// wibDateStr() di atas — tapi dilakukan langsung di level SQL memakai
// date(kolom_waktu, '+7 hours') supaya tidak perlu menarik seluruh baris
// ke JS dulu baru dikelompokkan.
// NON_BROADCAST_TEMPLATES (balasan manual/file/sync/pesan masuk) selalu
// dikecualikan di sini, supaya statistik "broadcast" murni pesan blast.

const BROADCAST_COST_PER_MESSAGE = 365; // Rp per broadcast per nomor

// Riwayat total nomor yang di-broadcast (blast template) per tanggal ATAU per
// bulan (groupBy: 'day' | 'month') — balasan manual/file/sync/pesan masuk
// TIDAK dihitung. Biaya dihitung Rp 365 per pesan broadcast yang BERHASIL
// terkirim saja (nomor yang gagal tidak dihitung biayanya). Kolom totalBerhasil/totalGagal dihitung dari status pengiriman
// tiap kontak (status 'failed' = gagal, selain itu dianggap berhasil).
function getBroadcastDailyHistory(days, groupBy) {
  const dateExpr = groupBy === 'month'
    ? `strftime('%Y-%m', c.time, '+7 hours')`
    : `date(c.time, '+7 hours')`;
  const rows = db.prepare(`
    SELECT ${dateExpr}            AS date,
           COUNT(*)                 AS totalBroadcast,
           COUNT(DISTINCT c.phone)  AS totalNomor,
           SUM(CASE WHEN c.status = 'failed' THEN 1 ELSE 0 END) AS totalGagal,
           SUM(CASE WHEN c.status != 'failed' THEN 1 ELSE 0 END) AS totalBerhasil
    FROM contacts c
    JOIN sessions s ON s.id = c.sessionId
    WHERE s.templateName NOT IN (${NON_BROADCAST_PLACEHOLDERS})
    GROUP BY date
    ORDER BY date DESC
    LIMIT ?
  `).all(...NON_BROADCAST_TEMPLATES, days);
  return rows.map(r => ({ ...r, biaya: (r.totalBerhasil || 0) * BROADCAST_COST_PER_MESSAGE }));
}

// Format waktu ISO (UTC) → string tanggal+jam WIB "dd/mm/yyyy HH:MM", dipakai
// khusus untuk kolom "Tanggal" pada file export supaya enak dibaca di Excel.
function formatExportDateWIB(iso) {
  const d   = new Date(iso);
  const wib = new Date(d.getTime() + 7 * 60 * 60 * 1000);
  const dd  = pad2(wib.getUTCDate());
  const mm  = pad2(wib.getUTCMonth() + 1);
  const yyyy = wib.getUTCFullYear();
  const hh  = pad2(wib.getUTCHours());
  const mi  = pad2(wib.getUTCMinutes());
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
}

// Ambil data mentah broadcast (per nomor, per pengiriman) dalam rentang
// tanggal tertentu, untuk keperluan export ke XLSX/CSV. Balasan manual/file/
// sync/pesan masuk TIDAK diikutkan — murni pesan blast/template.
function getBroadcastExportRows(startIso, endIso) {
  return db.prepare(`
    SELECT c.time AS tanggal,
           c.phone AS nomor,
           COALESCE(NULLIF(s.previewText, ''), s.templateName) AS pesan,
           c.status AS status
    FROM contacts c
    JOIN sessions s ON s.id = c.sessionId
    WHERE s.templateName NOT IN (${NON_BROADCAST_PLACEHOLDERS})
      AND c.time >= ? AND c.time <= ?
    ORDER BY c.time ASC
  `).all(...NON_BROADCAST_TEMPLATES, startIso, endIso);
}

// Nomor yang PALING SERING DIBROADCAST (bukan lagi "sering lapor" — sekarang
// murni menghitung pengiriman broadcast/blast template, TIDAK termasuk
// balasan manual/file/sync/pesan masuk). 1x kirim broadcast = 1x hitungan.
// "kaliBroadcast" = total berapa kali nomor ini menerima broadcast dalam
// rentang waktu, "dayCount" = jumlah HARI berbeda dia menerima broadcast.
function getFrequentReporters(minDays, limit, days) {
  const sinceIso = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
  const rows = db.prepare(`
    SELECT c.phone,
           COALESCE((SELECT cs.name FROM contact_summary cs WHERE cs.phone = c.phone), '') AS name,
           COUNT(DISTINCT date(c.time, '+7 hours')) AS dayCount,
           COUNT(*)     AS kaliBroadcast,
           MAX(c.time)  AS lastTime,
           MIN(c.time)  AS firstTime
    FROM contacts c
    JOIN sessions s ON s.id = c.sessionId
    WHERE c.time >= ?
      AND s.templateName NOT IN (${NON_BROADCAST_PLACEHOLDERS})
    GROUP BY c.phone
    HAVING dayCount >= ?
    ORDER BY dayCount DESC, kaliBroadcast DESC, lastTime DESC
    LIMIT ?
  `).all(sinceIso, ...NON_BROADCAST_TEMPLATES, minDays, limit);
  return rows;
}

// Trend harian JUMLAH NOMOR UNIK (bukan total pesan) WO Marking vs WO
// Individu, dalam rentang N hari terakhir. Balasan manual/file/sync/pesan
// masuk dikecualikan — murni menghitung nomor yang di-broadcast.
function getWoTrend(days) {
  const sinceIso = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
  const rows = db.prepare(`
    SELECT date(c.time, '+7 hours')        AS date,
           s.woCategory                    AS cat,
           COUNT(DISTINCT c.phone)         AS cnt
    FROM contacts c
    JOIN sessions s ON s.id = c.sessionId
    WHERE s.templateName NOT IN (${NON_BROADCAST_PLACEHOLDERS})
      AND s.woCategory IN ('marking','individu')
      AND c.time >= ?
    GROUP BY date, cat
    ORDER BY date ASC
  `).all(...NON_BROADCAST_TEMPLATES, sinceIso);

  // Susun jadi deret harian lengkap (tanggal tanpa data tetap muncul dengan nilai 0)
  // supaya garis chart tidak "meloncat" saat ada hari yang kosong.
  const byDate = {};
  rows.forEach(r => {
    byDate[r.date] = byDate[r.date] || { marking: 0, individu: 0 };
    byDate[r.date][r.cat] = r.cnt;
  });
  const result = [];
  const today = new Date(Date.now() + 7 * 60 * 60 * 1000);
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today.getTime() - i * 24 * 60 * 60 * 1000);
    const dateStr = d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
    result.push({ date: dateStr, marking: byDate[dateStr]?.marking || 0, individu: byDate[dateStr]?.individu || 0 });
  }
  return result;
}

// ============================================================
// CONTACT SUMMARY (tabel ringkasan untuk menu chat)
// ============================================================

// Hitung ulang 1 baris ringkasan untuk satu nomor, berdasarkan data
// TERBARU di tabel contacts/sessions/contact_category/read_state.
// Semua query di dalamnya pakai index (phone, time DESC) sehingga
// tetap cepat meski riwayat pesan sudah sangat banyak.
function recomputeSummaryForPhone(phone) {
  if (!phone) return;

  const lm = stmt.lastMsgForPhone.get(phone);
  if (!lm) {
    // Tidak ada lagi pesan untuk nomor ini (misalnya setelah dihapus) → buang dari ringkasan
    stmt.deleteSummaryForPhone.run(phone);
    return;
  }

  const nameRow = stmt.lastNameForPhone.get(phone);
  const meta     = stmt.getContactMeta.get(phone);
  const cc      = stmt.getContactCategory.get(phone);
  const lnuc    = cc ? null : stmt.lastNonUmumCatForPhone.get(phone);
  const woCategory = cc?.woCategory || lnuc?.woCategory || 'umum';

  const direction = lm.templateName === '📥 PESAN MASUK' ? 'in' : 'out';
  const text      = direction === 'in' ? (lm.message || '') : (lm.previewText || '');

  const rs = stmt.getReadState.get(phone);
  const isUnread = (direction === 'in' && (!rs || lm.time > rs.lastReadTime)) ? 1 : 0;

  stmt.replaceSummaryRow.run({
    phone,
    name: meta?.customName || nameRow?.name || '',
    woCategory,
    lastTime: lm.time,
    lastText: text,
    lastDirection: direction,
    lastStatus: lm.status || '',
    lastMsgType: lm.msgType || 'text',
    isUnread,
  });
}

// Isi awal contact_summary dari riwayat lama — HANYA jalan sekali (saat
// tabel masih kosong tapi sudah ada data lama di contacts). Query berat
// (CONTACT_SUMMARY_SQL) sengaja dipakai di sini karena cuma jalan 1x saat
// startup, bukan tiap kali menu chat dibuka.
function backfillContactSummary() {
  const already = db.prepare('SELECT COUNT(*) AS c FROM contact_summary').get().c;
  if (already > 0) return;

  const totalContacts = db.prepare('SELECT COUNT(*) AS c FROM contacts').get().c;
  if (!totalContacts) return;

  console.log('⏳ Membangun contact_summary dari riwayat lama (sekali jalan, mohon tunggu)...');
  const rows = db.prepare(CONTACT_SUMMARY_SQL).all();
  const insertAll = db.transaction((items) => {
    for (const r of items) {
      stmt.replaceSummaryRow.run({
        phone: r.phone,
        name: r.name || '',
        woCategory: r.woCategory || 'umum',
        lastTime: r.lastTime,
        lastText: r.lastText || '',
        lastDirection: r.lastDirection || 'out',
        lastStatus: r.lastStatus || '',
        lastMsgType: r.lastMsgType || 'text',
        isUnread: r.isUnread ? 1 : 0,
      });
    }
  });
  insertAll(rows);
  console.log(`✅ contact_summary awal terisi: ${rows.length} kontak`);
}

// Isi awal contacts_fts dari riwayat lama — HANYA jalan sekali (saat tabel FTS
// masih kosong tapi sudah ada data lama di contacts). Setelah ini, index FTS
// selalu disinkronkan otomatis lewat trigger di setiap insert pesan baru.
function backfillContactsFts() {
  if (!db.ftsAvailable) return;

  const ftsCount = db.prepare('SELECT COUNT(*) AS c FROM contacts_fts').get().c;
  if (ftsCount > 0) return;

  const totalContacts = db.prepare('SELECT COUNT(*) AS c FROM contacts').get().c;
  if (!totalContacts) return;

  console.log('⏳ Membangun index pencarian (FTS) dari riwayat lama (sekali jalan, mohon tunggu)...');
  db.prepare(`
    INSERT INTO contacts_fts(rowid, name, message, previewText)
    SELECT c.id, c.name, c.message, s.previewText
    FROM contacts c LEFT JOIN sessions s ON c.sessionId = s.id
  `).run();
  console.log('✅ Index pencarian (FTS) selesai dibangun');
}

function getLastCategory(phone) {
  const row = stmt.getContactCategory.get(phone);
  return row?.woCategory || 'umum';
}

function setContactCategory(phone, category) {
  if (!phone || !category) return;
  stmt.setContactCategory.run(phone, category, new Date().toISOString());
}

function historyLatestTime() {
  return stmt.latestTime.get().t || '';
}

function loadHistory() {
  const sessions = stmt.loadSessions.all();
  if (!sessions.length) return [];

  const result = sessions.map(r => ({
    id:           r.id,
    time:         r.time,
    templateName: r.templateName,
    previewText:  r.previewText,
    woCategory:   r.woCategory,
    summary: { total: r.total, success: r.success, failed: r.failed },
    contacts: [],
  }));

  const ids = result.map(s => s.id);
  if (!ids.length) return result;

  const placeholders = ids.map(() => '?').join(',');
  const contacts = db.prepare(
    `SELECT sessionId, phone, name, status, message, messageId, nomerLapor, msgType, attachmentUrl, time
     FROM contacts WHERE sessionId IN (${placeholders})`
  ).all(...ids);

  const sesMap = {};
  result.forEach(s => { sesMap[s.id] = s; });

  contacts.forEach(r => {
    const sess = sesMap[r.sessionId];
    if (!sess) return;
    sess.contacts.push({
      phone:         r.phone,
      name:          r.name,
      status:        r.status,
      message:       r.message,
      messageId:     r.messageId,
      nomerLapor:    r.nomerLapor,
      msgType:       r.msgType,
      attachmentUrl: r.attachmentUrl,
      time:          r.time,
    });
  });
  return result;
}

function appendSession(sess) {
  const affectedPhones = new Set();

  const tx = db.transaction(() => {
    stmt.insertSession.run(
      sess.id,
      sess.time,
      sess.templateName || '',
      sess.previewText  || '',
      sess.woCategory   || 'umum',
      sess.summary?.total   || 0,
      sess.summary?.success || 0,
      sess.summary?.failed  || 0
    );

    for (const c of (sess.contacts || [])) {
      stmt.insertContact.run(
        sess.id,
        c.phone         || '',
        c.name          || '',
        c.status        || 'success',
        c.message       || '',
        c.messageId     || '',
        c.nomerLapor    || '',
        c.msgType       || 'text',
        c.attachmentUrl || '',
        c.time          || sess.time
      );
      if (c.phone) affectedPhones.add(c.phone);
    }
  });
  tx();

  // Update ringkasan HANYA untuk nomor yang baru saja dapat pesan
  // (bukan scan semua kontak) → ini yang bikin menu chat tetap ringan.
  affectedPhones.forEach(recomputeSummaryForPhone);
}

function updateContactStatus(messageId, newStatus) {
  stmt.updateContact.run(newStatus, messageId);
}

// ============================================================
// Config & Templates
// ============================================================
function loadData()   { return JSON.parse(fs.readFileSync(TPL_FILE, 'utf8')); }
function saveData(d)  { fs.writeFileSync(TPL_FILE, JSON.stringify(d, null, 2), 'utf8'); }
function loadConfig() {
  const def = {
    apiToken: process.env.DEFAULT_API_TOKEN || '',
    webhookUrl: process.env.DEFAULT_WEBHOOK_URL || '',
    loginUsername: process.env.DEFAULT_LOGIN_USERNAME || 'admin',
    loginPassword: process.env.DEFAULT_LOGIN_PASSWORD || 'admin123',
    settingsPassword: process.env.DEFAULT_SETTINGS_PASSWORD || ''
  }; // ⚠️ nilai ini hanya dipakai saat config.json belum ada (first-run)
  return fs.existsSync(CFG_FILE) ? { ...def, ...JSON.parse(fs.readFileSync(CFG_FILE, 'utf8')) } : def;
}
function saveConfig(c) { fs.writeFileSync(CFG_FILE, JSON.stringify(c, null, 2), 'utf8'); }

// ============================================================
// DEDUPLICATION
// ============================================================
const recentlySentByApp = new Map();
function markSentByApp(phone, text) {
  const key = `${phone}|${String(text).trim().toLowerCase()}`;
  recentlySentByApp.set(key, Date.now());
  setTimeout(() => recentlySentByApp.delete(key), 15000);
}
function isSentByApp(phone, text) {
  const key = `${phone}|${String(text).trim().toLowerCase()}`;
  return recentlySentByApp.has(key);
}

function normalizePhone(raw) {
  let no = raw.trim().replace(/[\\s\\-\\.]/g, '');
  if (no.startsWith('+62'))    no = no.slice(1);
  else if (no.startsWith('0')) no = '62' + no.slice(1);
  else if (no.startsWith('8')) no = '62' + no;
  return no;
}

// ============================================================
// Maxchat API helpers
// ============================================================
async function sendTemplateMessage(to, templateId, params, token) {
  const r = await axios.post('https://app.maxchat.id/api/messages/push',
    { to, msgType: 'text', templateId, values: { body: params } },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 10000 }
  );
  return r.data;
}
async function sendFreeTextMessage(to, text, token) {
  const r = await axios.post('https://app.maxchat.id/api/messages/reply',
    { channel: 'whatsapp', msgType: 'text', to, text },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 10000 }
  );
  return r.data;
}
async function sendMediaMessage(to, fileUrl, msgType, caption, token) {
  const r = await axios.post('https://app.maxchat.id/api/messages/reply',
    { channel: 'whatsapp', msgType, to, attachmentUrl: fileUrl, text: caption || '' },
    { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 15000 }
  );
  return r.data;
}

// ============================================================
// ROUTES
// ============================================================
app.get('/api/data',   (req, res) => res.json(loadData()));
app.get('/api/config', (req, res) => {
  const cfg = loadConfig();
  const { loginPassword, settingsPassword, ...safeCfg } = cfg; // sembunyikan password dari response
  res.json(safeCfg);
});

app.post('/api/config', (req, res) => {
  const cfg = loadConfig();
  if (!checkSettingsPassword(req, res, cfg)) return;
  const { apiToken, webhookUrl } = req.body;
  saveConfig({ ...cfg, apiToken: apiToken || '', webhookUrl: webhookUrl || '' });
  res.json({ success: true });
});

app.post('/api/login-credentials', (req, res) => {
  const cfg = loadConfig();
  if (!checkSettingsPassword(req, res, cfg)) return;
  const { loginUsername, loginPassword } = req.body;
  if (!loginUsername) return res.json({ error: 'Username tidak boleh kosong' });
  cfg.loginUsername = loginUsername;
  if (loginPassword) cfg.loginPassword = loginPassword; // hanya update kalau diisi
  saveConfig(cfg);
  res.json({ success: true });
});

app.get('/api/system/db-info', (req, res) => {
  try {
    res.json(getDbInfo());
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/system/backup-now', async (req, res) => {
  try {
    const fileName = await runDatabaseBackup('manual');
    res.json({ success: true, fileName, ...getDbInfo() });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

app.get('/api/dashboard/broadcast-history', (req, res) => {
  try {
    const groupBy = req.query.groupBy === 'month' ? 'month' : 'day';
    const maxLimit = groupBy === 'month' ? 24 : 365;
    const days = Math.min(Math.max(parseInt(req.query.days) || 14, 1), maxLimit);
    res.json({ days, groupBy, rows: getBroadcastDailyHistory(days, groupBy) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Helper reusable untuk kirim file export (XLSX/CSV) dari data tabular —
// dipakai baik oleh export ringkasan (rekap per tanggal/bulan) maupun
// export detail (per nomor per pengiriman).
function sendExportFile(res, { format, fileBase, sheetName, headers, dataRows, colWidths }) {
  const ws = XLSX.utils.json_to_sheet(dataRows, { header: headers });
  if (colWidths) ws['!cols'] = colWidths.map(w => ({ wch: w }));

  if (format === 'csv') {
    const csv = XLSX.utils.sheet_to_csv(ws);
    res.setHeader('Content-Type', 'text/csv; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${fileBase}.csv"`);
    res.send('\uFEFF' + csv); // BOM supaya Excel baca karakter UTF-8 (mis. emoji template) dengan benar
  } else {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileBase}.xlsx"`);
    res.send(buf);
  }
}

// Export REKAP (ringkasan per tanggal/bulan) dari kartu "Riwayat Nomor
// Diblas per Tanggal" — kolom: Periode, Total, Berhasil, Gagal, Nomor
// Unik, Biaya. Mengikuti filter rentang & groupBy yang sedang aktif di UI.
app.get('/api/dashboard/broadcast-history-export', (req, res) => {
  try {
    const groupBy = req.query.groupBy === 'month' ? 'month' : 'day';
    const format  = req.query.format === 'csv' ? 'csv' : 'xlsx';
    const maxLimit = groupBy === 'month' ? 24 : 365;
    const days = Math.min(Math.max(parseInt(req.query.days) || 14, 1), maxLimit);

    const rows = getBroadcastDailyHistory(days, groupBy);
    if (!rows.length) return res.status(404).json({ error: 'Tidak ada data pada rentang ini' });

    const periodeLabel = groupBy === 'month' ? 'Bulan' : 'Tanggal';
    const dataRows = rows.slice().reverse().map(r => ({
      [periodeLabel]:  r.date,
      'Total':         r.totalBroadcast,
      'Berhasil':      r.totalBerhasil || 0,
      'Gagal':         r.totalGagal || 0,
      'Nomor Unik':    r.totalNomor,
      'Biaya (Rp)':    r.biaya,
    }));

    sendExportFile(res, {
      format,
      fileBase: `rekap_broadcast_${groupBy}_${days}${groupBy === 'month' ? 'bulan' : 'hari'}`,
      sheetName: 'Rekap Broadcast',
      headers: [periodeLabel, 'Total', 'Berhasil', 'Gagal', 'Nomor Unik', 'Biaya (Rp)'],
      dataRows,
      colWidths: [16, 10, 10, 10, 12, 14],
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Export data broadcast mentah (per nomor, per pengiriman) ke XLSX atau CSV,
// kolom: Tanggal, Nomor Telepon, Isi Pesan / Template, Status Pengiriman.
app.get('/api/dashboard/broadcast-export', (req, res) => {
  try {
    const { start, end } = req.query;
    const format = req.query.format === 'csv' ? 'csv' : 'xlsx';
    if (!start || !end) return res.status(400).json({ error: 'Parameter start dan end wajib diisi' });

    const startIso = new Date(start + 'T00:00:00.000').toISOString();
    const endIso   = new Date(end   + 'T23:59:59.999').toISOString();
    if (isNaN(Date.parse(startIso)) || isNaN(Date.parse(endIso))) {
      return res.status(400).json({ error: 'Format tanggal tidak valid' });
    }

    const rawRows = getBroadcastExportRows(startIso, endIso);
    if (!rawRows.length) return res.status(404).json({ error: 'Tidak ada data broadcast pada rentang tanggal ini' });

    const dataRows = rawRows.map(r => ({
      'Tanggal':               formatExportDateWIB(r.tanggal),
      'Nomor Telepon':         r.nomor,
      'Isi Pesan / Template':  r.pesan || '',
      'Status Pengiriman':     r.status === 'failed' ? 'Gagal' : (r.status === 'success' ? 'Berhasil' : r.status),
    }));

    sendExportFile(res, {
      format,
      fileBase: `broadcast_${start}_sd_${end}`,
      sheetName: 'Broadcast',
      headers: ['Tanggal', 'Nomor Telepon', 'Isi Pesan / Template', 'Status Pengiriman'],
      dataRows,
      colWidths: [18, 16, 45, 16],
    });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/dashboard/frequent-reporters', (req, res) => {
  try {
    const minDays = Math.max(parseInt(req.query.minDays) || 3, 2);
    const limit   = Math.min(Math.max(parseInt(req.query.limit) || 50, 1), 200);
    const days    = Math.min(Math.max(parseInt(req.query.days) || 90, 7), 90); // maks 90 hari
    res.json({ minDays, days, rows: getFrequentReporters(minDays, limit, days) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/dashboard/wo-trend', (req, res) => {
  try {
    const days = Math.min(Math.max(parseInt(req.query.days) || 30, 7), 180);
    res.json({ days, rows: getWoTrend(days) });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/history',       (req, res) => res.json(loadHistory()));
app.get('/api/history/count', (req, res) => res.json({ count: historyCount() }));

app.get('/api/history/latest-time', (req, res) => {
  res.json({ latestTime: historyLatestTime(), unreadCount: stmt.summaryUnreadCount.get().c || 0 });
});

app.delete('/api/history/:id', (req, res) => {
  const { id } = req.params;
  const affectedPhones = db.prepare('SELECT DISTINCT phone FROM contacts WHERE sessionId = ?').all(id).map(r => r.phone);
  const tx = db.transaction(() => {
    stmt.deleteContactsBySess.run(id);
    stmt.deleteSession.run(id);
  });
  tx();
  affectedPhones.forEach(recomputeSummaryForPhone);
  res.json({ success: true });
});

app.delete('/api/history', (req, res) => {
  const tx = db.transaction(() => {
    stmt.deleteAllContacts.run();
    stmt.deleteAllSessions.run();
    db.exec('DELETE FROM read_state');       
    db.exec('DELETE FROM contact_category'); 
    db.exec('DELETE FROM contact_summary');
  });
  tx();
  res.json({ success: true });
});

app.delete('/api/history/range', (req, res) => {
  const { start, end } = req.query;
  if (!start || !end) return res.status(400).json({ error: 'Parameter start dan end wajib diisi' });

  const startIso = new Date(start + 'T00:00:00.000').toISOString();
  const endIso   = new Date(end   + 'T23:59:59.999').toISOString();
  if (isNaN(Date.parse(startIso)) || isNaN(Date.parse(endIso))) {
    return res.status(400).json({ error: 'Format tanggal tidak valid' });
  }

  const affectedPhones = db.prepare(`
    SELECT DISTINCT c.phone FROM contacts c JOIN sessions s ON c.sessionId = s.id
    WHERE s.time >= ? AND s.time <= ?
  `).all(startIso, endIso).map(r => r.phone);

  let deletedCount = 0;
  const tx = db.transaction(() => {
    deletedCount = stmt.findSessionsByRange.all(startIso, endIso).length;
    stmt.deleteContactsByRange.run(startIso, endIso);
    stmt.deleteSessionsByRange.run(startIso, endIso);
  });
  tx();
  affectedPhones.forEach(recomputeSummaryForPhone);
  res.json({ success: true, deleted: deletedCount });
});

app.get('/api/contacts', (req, res) => {
  const t0 = Date.now();
  const limit    = Math.min(parseInt(req.query.limit, 10) || 50, 200);
  const before   = req.query.before || null;
  const category = ['marking', 'individu'].includes(req.query.category) ? req.query.category : null;
  const unreadOnly = req.query.unread === '1';

  // Dibaca langsung dari contact_summary (1 baris per kontak, selalu up to date) —
  // tidak lagi menghitung ulang dari seluruh riwayat pesan tiap request, jadi
  // menu chat, filter Marking/Individu, dan tab Belum Terbaca sama-sama cepat.
  let sql = `
    SELECT phone, name, woCategory, lastTime, lastText, lastDirection, lastStatus, lastMsgType, isUnread
    FROM contact_summary
    WHERE 1 = 1
  `;
  const params = [];
  if (category)    { sql += ` AND woCategory = ?`; params.push(category); }
  if (unreadOnly)  { sql += ` AND isUnread = 1`; }
  if (before)      { sql += ` AND lastTime < ?`; params.push(before); }
  // Pesan belum terbaca selalu ditampilkan paling atas dulu, baru sisanya
  // urut waktu terbaru — berlaku untuk semua tab (Semua, WO Marking, WO Individu).
  sql += ` ORDER BY isUnread DESC, lastTime DESC LIMIT ?`;
  params.push(limit + 1);

  const tQuery0 = Date.now();
  const rawRows = db.prepare(sql).all(...params);
  const tQuery1 = Date.now();

  const rows = rawRows.map(r => ({ ...r, isUnread: !!r.isUnread }));

  const hasMore = rows.length > limit;
  if (hasMore) rows.pop();

  // DIAGNOSTIK SEMENTARA — hapus lagi setelah sumber lag ketemu.
  // Kalau "query" di bawah ini besar (misal >200ms), masalah masih di database.
  // Kalau "query" kecil tapi "total" besar, masalahnya ada di tempat lain (jaringan/render).
  console.log(`[api/contacts] category=${category || '-'} unread=${unreadOnly} query=${tQuery1 - tQuery0}ms total=${Date.now() - t0}ms rows=${rows.length}`);

  res.json({ contacts: rows, hasMore, oldestTime: rows.length ? rows[rows.length - 1].lastTime : null });
});

app.post('/api/contacts/:phone/read', (req, res) => {
  const { phone } = req.params;
  const latest = stmt.latestMsgTimeForPhone.get(phone);
  const readTime = latest?.t || new Date().toISOString();
  stmt.markRead.run(phone, readTime);
  recomputeSummaryForPhone(phone); // supaya badge unread di menu chat langsung hilang
  res.json({ success: true, lastReadTime: readTime });
});

app.get('/api/contacts/:phone/messages', (req, res) => {
  const { phone } = req.params;
  const limit  = Math.min(parseInt(req.query.limit, 10) || 10, 500);
  const before = req.query.before || null;
  const after  = req.query.after || null;

  if (after) {
    const rows = stmt.contactMessagesAfter.all(phone, after); 
    const messages = rows.map(r => ({
      text: r.templateName === '📥 PESAN MASUK' ? r.message : r.previewText,
      time: r.time,
      status: r.status,
      templateName: r.templateName,
      direction: r.templateName === '📥 PESAN MASUK' ? 'in' : 'out',
      nomerLapor: r.nomerLapor || '',
      msgType: r.msgType || 'text',
      attachmentUrl: r.attachmentUrl || '',
    }));
    return res.json({ messages, hasMore: false, oldestTime: null });
  }

  const rows = before
    ? stmt.contactMessagesPage.all(phone, before, limit + 1)
    : stmt.contactMessagesLatest.all(phone, limit + 1);

  const hasMore = rows.length > limit;
  if (hasMore) rows.pop();
  rows.reverse(); 

  const messages = rows.map(r => ({
    text: r.templateName === '📥 PESAN MASUK' ? r.message : r.previewText,
    time: r.time,
    status: r.status,
    templateName: r.templateName,
    direction: r.templateName === '📥 PESAN MASUK' ? 'in' : 'out',
    nomerLapor: r.nomerLapor || '',
    msgType: r.msgType || 'text',
    attachmentUrl: r.attachmentUrl || '',
  }));

  res.json({ messages, hasMore, oldestTime: rows.length ? rows[0].time : null });
});

// Ubah kata kunci pencarian jadi ekspresi query FTS5: tiap kata dijadikan
// prefix match ("kata"*) lalu digabung OR, supaya "budi sant" tetap nemu
// "Budi Santoso". Karakter kutip dibuang supaya query FTS tidak pernah error
// walau user mengetik simbol aneh.
function ftsSanitize(term) {
  const words = term
    .split(/\s+/)
    .map(w => w.replace(/["']/g, ''))
    .filter(w => w.length > 0);
  if (!words.length) return null;
  return words.map(w => `"${w}"*`).join(' OR ');
}

app.get('/api/contacts/search', (req, res) => {
  const q = (req.query.q || '').trim();
  const before = req.query.before || null;
  const SEARCH_RESULT_LIMIT = 50;

  if (!q) {
    let sql = `
      SELECT phone, name, woCategory, lastTime, lastText, lastDirection, lastStatus, lastMsgType, isUnread
      FROM contact_summary WHERE 1 = 1
    `;
    const params = [];
    if (before) { sql += ` AND lastTime < ?`; params.push(before); }
    sql += ` ORDER BY lastTime DESC LIMIT ?`;
    params.push(SEARCH_RESULT_LIMIT + 1);

    const rawRows = db.prepare(sql).all(...params);
    const hasMore = rawRows.length > SEARCH_RESULT_LIMIT;
    if (hasMore) rawRows.pop();
    const rows = rawRows.map(r => ({ ...r, isUnread: !!r.isUnread }));
    return res.json({ contacts: rows, hasMore, oldestTime: rows.length ? rows[rows.length - 1].lastTime : null });
  }

  const like = `%${q}%`;

  // Kalau query terlihat seperti nomor telepon (mis. 081283456789, 6281283456789,
  // 81283456789), normalisasi dulu pakai fungsi yang sama dipakai saat simpan
  // kontak, supaya ketiga format itu ketemu walau nomor disimpan dalam format 62xxx.
  const digitsOnly = q.replace(/[^0-9]/g, '');
  const looksLikePhone = /^[0-9+][0-9+\s-]*$/.test(q) && digitsOnly.length >= 3;
  const normalizedLike = looksLikePhone ? `%${normalizePhone(q)}%` : null;

  // 1) Nomor telepon: tetap LIKE, tapi sekarang hanya menyisir kolom phone saja
  //    (bukan join+scan message/previewText seperti sebelumnya) — jauh lebih ringan.
  const phoneSql = normalizedLike
    ? `SELECT DISTINCT phone FROM contacts WHERE phone LIKE ? OR phone LIKE ?`
    : `SELECT DISTINCT phone FROM contacts WHERE phone LIKE ?`;
  const phoneParams = normalizedLike ? [like, normalizedLike] : [like];
  const phoneMatches = db.prepare(phoneSql).all(...phoneParams).map(r => r.phone);

  // 2) Nama / isi pesan / preview template: pakai index FTS5 kalau tersedia
  //    (jauh lebih cepat dibanding LIKE '%...%' menyisir semua baris). Kalau
  //    FTS5 tidak tersedia di build better-sqlite3 yang dipakai, fallback ke
  //    LIKE lama supaya pencarian tetap jalan (hanya lebih lambat).
  let textMatches = [];
  if (db.ftsAvailable) {
    const ftsExpr = ftsSanitize(q);
    if (ftsExpr) {
      try {
        textMatches = db.prepare(`
          SELECT DISTINCT c.phone
          FROM contacts_fts f
          JOIN contacts c ON c.id = f.rowid
          WHERE contacts_fts MATCH ?
          LIMIT 500
        `).all(ftsExpr).map(r => r.phone);
      } catch (err) {
        // Query FTS tidak valid (jarang terjadi) — abaikan, hasil dari LIKE nomor tetap dipakai.
        console.error('⚠️  FTS search error, dilewati:', err.message);
      }
    }
  } else {
    textMatches = db.prepare(`
      SELECT DISTINCT c.phone
      FROM contacts c
      LEFT JOIN sessions s ON c.sessionId = s.id
      WHERE c.name LIKE ? OR c.message LIKE ? OR s.previewText LIKE ?
    `).all(like, like, like).map(r => r.phone);
  }

  const matchedPhones = [...new Set([...phoneMatches, ...textMatches])];
  if (!matchedPhones.length) return res.json({ contacts: [], hasMore: false, oldestTime: null });

  // Diambil per-halaman (50 kontak per klik "muat lebih banyak"), terbaru
  // duluan, sama seperti pola pagination "before" di /api/contacts — supaya
  // pesan/kontak yang lebih lama tetap bisa dijangkau lewat tombol muat lagi,
  // bukan cuma 50 kontak paling baru yang ke-cut permanen.
  const placeholders = matchedPhones.map(() => '?').join(',');
  let summarySql = `
    SELECT phone, name, woCategory, lastTime, lastText, lastDirection, lastStatus, lastMsgType, isUnread
    FROM contact_summary WHERE phone IN (${placeholders})
  `;
  const summaryParams = [...matchedPhones];
  if (before) { summarySql += ` AND lastTime < ?`; summaryParams.push(before); }
  summarySql += ` ORDER BY lastTime DESC LIMIT ?`;
  summaryParams.push(SEARCH_RESULT_LIMIT + 1);

  const rawRows = db.prepare(summarySql).all(...summaryParams);
  const hasMore = rawRows.length > SEARCH_RESULT_LIMIT;
  if (hasMore) rawRows.pop();
  const rows = rawRows.map(r => ({ ...r, isUnread: !!r.isUnread }));

  res.json({ contacts: rows, hasMore, oldestTime: rows.length ? rows[rows.length - 1].lastTime : null });
});

app.get('/api/history/search', (req, res) => {
  const q = (req.query.q || '').trim();
  if (!q) return res.json([]);

  const like = `%${q}%`;
  const matches = stmt.searchSessions.all(like, like, like);
  if (!matches.length) return res.json([]);

  const ids = matches.map(r => r.sid);
  const placeholders = ids.map(() => '?').join(',');
  const sessions = db.prepare(
    `SELECT id, time, templateName, previewText, woCategory, total, success, failed
     FROM sessions WHERE id IN (${placeholders}) ORDER BY time DESC`
  ).all(...ids);

  const contacts = db.prepare(
    `SELECT sessionId, phone, name, status, message, messageId, nomerLapor, msgType, attachmentUrl, time
     FROM contacts WHERE sessionId IN (${placeholders})`
  ).all(...ids);

  const result = sessions.map(r => ({
    id: r.id, time: r.time, templateName: r.templateName, previewText: r.previewText,
    woCategory: r.woCategory, summary: { total: r.total, success: r.success, failed: r.failed },
    contacts: [],
  }));
  const sesMap = {};
  result.forEach(s => { sesMap[s.id] = s; });
  contacts.forEach(r => {
    const sess = sesMap[r.sessionId];
    if (sess) sess.contacts.push(r);
  });
  res.json(result);
});

app.post('/api/blast', async (req, res) => {
  const { numbers, templateId, params, templateName, previewText, nomerLapor, blastId } = req.body;
  if (!numbers?.length) return res.json({ error: 'Nomor tidak boleh kosong' });
  if (!templateId)      return res.json({ error: 'Template belum dipilih' });

  if (blastId) setBlastProgress(blastId, { total: numbers.length, sent: 0, success: 0, failed: 0, done: false });

  const cfg   = loadConfig();
  const token = cfg.apiToken || '';

  if (templateId === '__reply__') {
    const results = numbers.map(no => ({ to: normalizePhone(no), status: 'success', message: 'Balasan manual' }));
    const byCategory = {};
    results.forEach(r => {
      const cat = getLastCategory(r.to);
      (byCategory[cat] = byCategory[cat] || []).push(r);
    });
    Object.entries(byCategory).forEach(([cat, group], idx) => {
      appendSession({
        id: 'sess_' + Date.now() + '_' + idx, time: new Date().toISOString(),
        templateName: templateName || 'Balasan Manual',
        previewText:  previewText  || '',
        woCategory:   cat,
        summary: { total: group.length, success: group.length, failed: 0 },
        contacts: group.map(r => ({ phone: r.to, status: r.status, message: r.message, time: new Date().toISOString() })),
      });
    });
    if (blastId) setBlastProgress(blastId, { sent: results.length, success: results.length, failed: 0, done: true });
    return res.json({ results, summary: { total: results.length, success: results.length, failed: 0 } });
  }

  if (!token) {
    if (blastId) setBlastProgress(blastId, { done: true });
    return res.json({ error: 'API Token belum dikonfigurasi di menu Pengaturan' });
  }

  const tplDef = loadData().templates.find(t => t.templateId === templateId);
  let resolvedPreview = (tplDef && tplDef.preview) ? tplDef.preview : (previewText || '');
  if (params?.length) params.forEach(p => { resolvedPreview = resolvedPreview.split('{{' + p.index + '}}').join(p.text || ''); });
  if (nomerLapor) resolvedPreview = resolvedPreview.replace(/__nomer_lapor__/g, nomerLapor);

  const results = [];
  for (let i = 0; i < numbers.length; i++) {
    const no = normalizePhone(numbers[i]);
    if (!no) continue;
    setContactCategory(no, tplDef?.woCategory || 'umum'); 
    try {
      markSentByApp(no, resolvedPreview);
      const data = await sendTemplateMessage(no, templateId, params || [], token);
      results.push({ to: no, status: 'success', message: data.content || 'Terkirim', messageId: data.id || data.messageId || '' });
    } catch (err) {
      results.push({ to: no, status: 'failed', message: err.response?.data?.message || err.message || 'Gagal' });
    }
    if (blastId) {
      const successSoFar = results.filter(r => r.status === 'success').length;
      setBlastProgress(blastId, { sent: results.length, success: successSoFar, failed: results.length - successSoFar });
    }
    if (i < numbers.length - 1) await new Promise(r => setTimeout(r, DELAY_MS));
  }
  const success = results.filter(r => r.status === 'success').length;

  appendSession({
    id: 'sess_' + Date.now(), time: new Date().toISOString(),
    templateName: templateName || templateId,
    woCategory:   tplDef?.woCategory || 'umum',
    previewText:  resolvedPreview || previewText || '',
    summary: { total: results.length, success, failed: results.length - success },
    contacts: results.map(r => ({
      phone: r.to, nomerLapor: nomerLapor || '', status: r.status,
      message: r.message, messageId: r.messageId || '', time: new Date().toISOString(),
    })),
  });

  if (blastId) {
    setBlastProgress(blastId, { done: true });
    setTimeout(() => blastProgress.delete(blastId), 60000); // bersihkan dari memori setelah 1 menit
  }

  res.json({ results, summary: { total: results.length, success, failed: results.length - success } });
});

// Dipoll dari frontend selama proses blast berjalan di server, supaya progress
// bar bisa menampilkan angka x/y yang benar-benar live (bukan tebakan client).
app.get('/api/blast/progress/:id', (req, res) => {
  const p = blastProgress.get(req.params.id);
  if (!p) return res.json({ found: false });
  res.json({ found: true, ...p });
});

app.post('/api/reply', async (req, res) => {
  const { to, text } = req.body;
  if (!to || !text) return res.json({ error: 'Nomor dan isi pesan wajib diisi' });
  const cfg   = loadConfig();
  const token = cfg.apiToken || '';
  if (!token) return res.json({ error: 'API Token belum dikonfigurasi di menu Pengaturan' });

  const phone = normalizePhone(to);
  try {
    markSentByApp(phone, text);
    const data = await sendFreeTextMessage(phone, text, token);
    appendSession({
      id: 'sess_' + Date.now(), time: new Date().toISOString(),
      templateName: '💬 Balasan Manual',
      previewText:  text,
      woCategory:   getLastCategory(phone), 
      summary: { total: 1, success: 1, failed: 0 },
      contacts: [{ phone, status: 'success', message: 'Terkirim', messageId: data.id || data.messageId || '', time: new Date().toISOString() }],
    });
    res.json({ success: true });
  } catch (err) {
    res.json({ error: err.response?.data?.message || err.message || 'Gagal mengirim' });
  }
});

const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 16 * 1024 * 1024 } });

app.post('/api/reply-media-upload', (req, res) => {
  upload.single('file')(req, res, async (err) => {
    if (err) return res.json({ error: 'Upload gagal: ' + err.message });
    if (!req.file) return res.json({ error: 'File tidak ditemukan' });

    const cfg   = loadConfig();
    const token = cfg.apiToken || '';
    if (!token) return res.json({ error: 'API Token belum dikonfigurasi' });

    const phone   = normalizePhone(req.body.to || '');
    const caption = req.body.caption || '';
    if (!phone) return res.json({ error: 'Nomor tidak valid' });

    try {
      const b64     = req.file.buffer.toString('base64');
      const mime    = req.file.mimetype;
      const isVideo = mime.startsWith('video/');
      const apiRes  = await axios.post('https://app.maxchat.id/api/v1/messages/send-media', {
        to:      phone + '@c.us',
        type:    isVideo ? 'video' : 'image',
        media:   { base64: b64, filename: req.file.originalname, mimeType: mime },
        caption,
      }, { headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' } });

      markSentByApp(phone, caption || req.file.originalname);
      appendSession({
        id: 'sess_' + Date.now(), time: new Date().toISOString(),
        templateName: '📎 File Terkirim',
        previewText:  caption || req.file.originalname,
        woCategory:   getLastCategory(phone), 
        summary: { total: 1, success: 1, failed: 0 },
        contacts: [{ phone, status: 'success', message: caption || req.file.originalname, messageId: apiRes.data.id || '', time: new Date().toISOString() }],
      });
      res.json({ success: true });
    } catch (e) {
      res.json({ error: e.response?.data?.message || e.message || 'Gagal kirim file' });
    }
  });
});

app.post('/api/reply-media', async (req, res) => {
  const { to, fileUrl, msgType, caption } = req.body;
  if (!to || !fileUrl) return res.json({ error: 'Nomor dan URL file wajib diisi' });
  const cfg   = loadConfig();
  const token = cfg.apiToken || '';
  if (!token) return res.json({ error: 'API Token belum dikonfigurasi' });
  const phone = normalizePhone(to);
  try {
    markSentByApp(phone, caption || fileUrl);
    const data = await sendMediaMessage(phone, fileUrl, msgType || 'image', caption || '', token);
    appendSession({
      id: 'sess_' + Date.now(), time: new Date().toISOString(),
      templateName: '📤 PESAN KELUAR (SYNC)',
      previewText:  caption || '[' + (msgType || 'file') + ']',
      woCategory:   getLastCategory(phone), 
      summary: { total: 1, success: 1, failed: 0 },
      contacts: [{
        phone, status: 'success',
        message: caption || '[' + (msgType || 'file') + ']',
        msgType: msgType || 'image', attachmentUrl: fileUrl,
        messageId: data.id || data.messageId || '', time: new Date().toISOString(),
      }],
    });
    res.json({ success: true, data });
  } catch (err) {
    console.error('reply-media error:', err.response?.data || err.message);
    res.json({ error: err.response?.data?.message || err.message });
  }
});

app.post('/api/templates/reorder', (req, res) => {
  const { ids } = req.body;
  const data = loadData();
  ids.forEach(item => {
    const t = data.templates.find(x => x.id === item.id);
    if (t) t.order = item.order;
  });
  saveData(data);
  res.json({ success: true });
});

app.post('/api/templates', (req, res) => {
  const { name, templateId, hasParams, preview, params, woCategory } = req.body;
  if (!name || !templateId) return res.json({ error: 'Nama dan Template ID wajib diisi' });
  const data = loadData();
  const tpl  = {
    id: 'tpl_' + Date.now(), name, templateId, channel: 'whatsapp',
    hasParams: !!hasParams, preview: preview || '', params: params || [],
    order: data.templates.length + 1, woCategory: woCategory || 'umum',
  };
  data.templates.push(tpl);
  saveData(data);
  res.json({ success: true, template: tpl });
});

app.put('/api/templates/:id', (req, res) => {
  const { name, templateId, hasParams, preview, params, woCategory } = req.body;
  if (!name || !templateId) return res.json({ error: 'Nama dan Template ID wajib diisi' });
  const data = loadData();
  const idx  = data.templates.findIndex(t => t.id === req.params.id);
  if (idx === -1) return res.json({ error: 'Template tidak ditemukan' });
  data.templates[idx] = { ...data.templates[idx], name, templateId, hasParams: !!hasParams, preview: preview || '', params: params || [], woCategory: woCategory || 'umum' };
  saveData(data);
  res.json({ success: true, template: data.templates[idx] });
});

app.delete('/api/templates/:id', (req, res) => {
  const data = loadData();
  data.templates = data.templates.filter(t => t.id !== req.params.id);
  saveData(data);
  res.json({ success: true });
});

// ============================================================
// QUICK REPLY
// ============================================================
app.get('/api/quick-replies', (req, res) => {
  res.json(stmt.loadQuickReplies.all());
});

app.post('/api/quick-replies', (req, res) => {
  const text = (req.body.text || '').trim();
  if (!text) return res.json({ error: 'Isi quick reply wajib diisi' });
  const maxOrder = stmt.maxQuickReplyOrder.get().m || 0;
  const id = 'qr_' + Date.now();
  stmt.insertQuickReply.run(id, text, maxOrder + 1);
  res.json({ success: true, quickReply: { id, text, orderNum: maxOrder + 1 } });
});

app.put('/api/quick-replies/:id', (req, res) => {
  const text = (req.body.text || '').trim();
  if (!text) return res.json({ error: 'Isi quick reply wajib diisi' });
  if (!stmt.checkQuickReplyExists.get(req.params.id)) return res.json({ error: 'Quick reply tidak ditemukan' });
  stmt.updateQuickReplyText.run(text, req.params.id);
  res.json({ success: true });
});

app.delete('/api/quick-replies/:id', (req, res) => {
  stmt.deleteQuickReply.run(req.params.id);
  res.json({ success: true });
});

app.post('/api/quick-replies/reorder', (req, res) => {
  const { ids } = req.body;
  if (!Array.isArray(ids)) return res.json({ error: 'Data urutan tidak valid' });
  const tx = db.transaction((items) => {
    items.forEach((id, idx) => stmt.updateQuickReplyOrd.run(idx + 1, id));
  });
  tx(ids);
  res.json({ success: true });
});

// ============================================================
// KONTAK & LABEL
// ============================================================

// Daftar semua teks label yang pernah ditempelkan ke kontak manapun (dipakai
// untuk chip filter "Semua nomor yang pernah dikirimi pesan..." di menu Kontak).
app.get('/api/tags', (req, res) => {
  res.json(stmt.loadDistinctTags.all().map(r => r.tag));
});

// ------------------------------------------------------------
// MASTER LABEL — label sekarang harus dibuat manual dulu lewat "Kelola
// Label" (masing-masing punya tab sendiri di menu Kontak). Baru setelah
// dibuat, label itu bisa ditempelkan/dilepas ke kontak — TIDAK BISA lagi
// ketik teks bebas langsung di kontak seperti versi sebelumnya.
// ------------------------------------------------------------
app.get('/api/labels', (req, res) => {
  res.json(stmt.loadLabels.all());
});

app.post('/api/labels', (req, res) => {
  const name  = (req.body.name  || '').trim();
  const color = (req.body.color || '#2563eb').trim();
  if (!name) return res.json({ error: 'Nama label wajib diisi' });
  if (stmt.findLabelByName.get(name)) return res.json({ error: 'Label dengan nama ini sudah ada' });
  const id = 'lbl_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const maxOrder = (db.prepare('SELECT COALESCE(MAX(order_num), 0) AS m FROM labels').get() || {}).m || 0;
  stmt.insertLabel.run(id, name, color, maxOrder + 1, new Date().toISOString());
  res.json({ success: true, label: { id, name, color, order_num: maxOrder + 1 } });
});

app.put('/api/labels/:id', (req, res) => {
  const existing = stmt.getLabelById.get(req.params.id);
  if (!existing) return res.json({ error: 'Label tidak ditemukan' });
  const name  = (req.body.name  || '').trim() || existing.name;
  const color = (req.body.color || existing.color);
  const dup = stmt.findLabelByName.get(name);
  if (dup && dup.id !== existing.id) return res.json({ error: 'Label dengan nama ini sudah ada' });
  if (name !== existing.name) stmt.renameContactTagUsage.run(name, existing.name);
  stmt.updateLabel.run(name, color, existing.id);
  res.json({ success: true, label: { id: existing.id, name, color } });
});

app.delete('/api/labels/:id', (req, res) => {
  const existing = stmt.getLabelById.get(req.params.id);
  if (!existing) return res.json({ error: 'Label tidak ditemukan' });
  stmt.deleteContactTagsByName.run(existing.name);
  stmt.deleteLabelRow.run(existing.id);
  res.json({ success: true });
});

// Label yang sedang ditempel di satu nomor kontak (dipakai juga di menu Chat
// supaya bisa lihat/atur label langsung sambil ngobrol).
app.get('/api/kontak/:phone/tags', (req, res) => {
  res.json(stmt.loadContactTagsByPhone.all(req.params.phone).map(r => r.tag));
});

// Tempelkan SATU label yang SUDAH ADA di daftar master ke satu kontak.
// Tidak bisa lagi ketik teks bebas — nama label harus persis cocok dengan
// salah satu label yang sudah dibuat lewat "Kelola Label".
app.post('/api/kontak/:phone/tags', (req, res) => {
  const tag = (req.body.tag || '').trim();
  if (!tag) return res.json({ error: 'Label tidak boleh kosong' });
  if (!stmt.findLabelByName.get(tag)) return res.json({ error: 'Label ini belum dibuat. Buat dulu lewat "Kelola Label".' });
  stmt.attachContactTag.run(req.params.phone, tag);
  res.json({ success: true, tag });
});

app.delete('/api/kontak/:phone/tags/:tag', (req, res) => {
  stmt.detachContactTag.run(req.params.phone, req.params.tag);
  res.json({ success: true });
});

app.put('/api/kontak/:phone/name', (req, res) => {
  const name = (req.body.name || '').trim();
  stmt.setContactMeta.run(req.params.phone, name, new Date().toISOString());
  recomputeSummaryForPhone(req.params.phone);
  res.json({ success: true });
});

// Daftar kontak untuk menu "Kontak" — SEMUA nomor yang pernah dikirimi pesan
// (WO Marking maupun Individu digabung, sesuai contact_summary yang sudah
// tidak membedakan kategori) DIGABUNG dengan kontak yang ditambahkan manual
// lewat tombol "Tambah Kontak" (belum tentu pernah dikirimi pesan sama
// sekali) — lengkap dengan label bebas yang menempel di tiap nomor.
app.get('/api/kontak', (req, res) => {
  const q      = (req.query.q || '').trim();
  const tag    = req.query.tag || null;
  const limit  = Math.min(Math.max(parseInt(req.query.limit, 10)  || 100, 1), 500);
  const offset = Math.max(parseInt(req.query.offset, 10) || 0, 0);

  // Gabungan: semua nomor di contact_summary (pernah dikirimi pesan) UNION
  // nomor yang ditambahkan manual TAPI belum ada riwayat pesannya sama sekali.
  const combinedBase = `
    SELECT phone, name, lastTime, 0 AS isManualOnly FROM contact_summary
    UNION ALL
    SELECT cm.phone, cm.customName AS name, NULL AS lastTime, 1 AS isManualOnly
    FROM contact_meta cm
    WHERE cm.manuallyAdded = 1 AND cm.phone NOT IN (SELECT phone FROM contact_summary)
  `;

  let baseWhere = ' WHERE 1 = 1';
  const baseParams = [];
  if (tag) { baseWhere += ` AND phone IN (SELECT phone FROM contact_tags WHERE tag = ?)`; baseParams.push(tag); }
  if (q)   { baseWhere += ` AND (phone LIKE ? OR name LIKE ?)`; baseParams.push('%'+q+'%', '%'+q+'%'); }

  const rows = db.prepare(`
    SELECT * FROM (${combinedBase}) combined
    ${baseWhere}
    ORDER BY (name = '') ASC, name COLLATE NOCASE ASC, lastTime DESC
    LIMIT ? OFFSET ?
  `).all(...baseParams, limit, offset);

  const total = db.prepare(`SELECT COUNT(*) AS c FROM (${combinedBase}) combined ${baseWhere}`).get(...baseParams).c || 0;

  let tagsByPhone = {};
  if (rows.length) {
    const phones = rows.map(r => r.phone);
    const placeholders = phones.map(() => '?').join(',');
    const tagRows = db.prepare(`SELECT phone, tag FROM contact_tags WHERE phone IN (${placeholders}) ORDER BY tag COLLATE NOCASE ASC`).all(...phones);
    tagRows.forEach(r => { (tagsByPhone[r.phone] = tagsByPhone[r.phone] || []).push(r.tag); });
  }

  res.json({
    contacts: rows.map(r => ({ ...r, isManualOnly: !!r.isManualOnly, labels: tagsByPhone[r.phone] || [] })),
    total,
    hasMore: offset + rows.length < total,
  });
});

// Tambah kontak manual — nomor yang BELUM PERNAH dikirimi pesan sama sekali,
// disiapkan lebih dulu (nama & label) sebelum broadcast pertama dilakukan.
app.post('/api/kontak', (req, res) => {
  const phone = normalizePhone((req.body.phone || '').trim());
  const name  = (req.body.name  || '').trim();
  if (!phone) return res.json({ error: 'Nomor wajib diisi' });
  stmt.insertManualContact.run(phone, name, new Date().toISOString());
  res.json({ success: true, contact: { phone, name, lastTime: null, isManualOnly: true, labels: [] } });
});

// Hapus kontak manual — HANYA boleh untuk nomor yang belum punya riwayat
// pesan/broadcast sama sekali (kontak yang sudah pernah chat tetap harus
// tersimpan sebagai bagian dari riwayat, dihapus lewat menu Pengaturan → Hapus Riwayat).
app.delete('/api/kontak/:phone', (req, res) => {
  const { phone } = req.params;
  if (stmt.checkContactHasHistory.get(phone)) {
    return res.json({ error: 'Kontak ini sudah punya riwayat chat/broadcast, tidak bisa dihapus dari menu Kontak. Hapus riwayatnya dulu lewat Pengaturan → Hapus Riwayat.' });
  }
  const tx = db.transaction(() => {
    stmt.deleteContactTagsByPhone.run(phone);
    stmt.deleteManualContact.run(phone);
  });
  tx();
  res.json({ success: true });
});

// ============================================================
// WEBHOOK
// ============================================================
app.post('/webhook', (req, res) => {
  const payload = req.body;
  // Log ringkas saja — JANGAN dump seluruh payload (JSON.stringify + console.log
  // synchronous di bawah PM2 bisa memblokir event loop tiap webhook masuk,
  // apalagi kalau ada lampiran media base64 yang besar. Ini penyebab utama
  // menu chat/tab kategori terasa "berat" walau query database sendiri sudah 0ms.
  console.log(`=== WEBHOOK MASUK === event=${payload.event || payload.type || '-'} from=${payload.from || payload.data?.from || '-'} to=${payload.to || payload.data?.to || '-'} msgType=${payload.msgType || payload.type || '-'}`);

  try {
    const rawTs     = payload.timestamp || payload.created_at || Date.now();
    const timestamp = String(rawTs).length <= 10 ? rawTs * 1000 : Number(rawTs);

    if (payload.type === 'status_update' || payload.event === 'message_status' || (payload.status && !payload.text && !payload.data)) {
      const msgId  = payload.message_id || payload.id || payload.msgId;
      const status = payload.status || payload.delivery_status;
      if (msgId && status) {
        const mapped = status === 'read' ? 'read' : (status === 'delivered' ? 'delivered' : null);
        if (mapped) updateContactStatus(msgId, mapped);
      }
      return res.status(200).send('OK');
    }

    const isOutgoing = payload.event === 'message_out' ||
                       payload.direction === 'out'      ||
                       payload.type === 'outbound'      ||
                       !!payload.to;

    if (isOutgoing) {
      const phoneRaw = payload.to || payload.data?.to || payload.to?.[0] || '';
      const phone    = normalizePhone(String(phoneRaw).replace('@c.us', ''));
      const message  = payload.text || payload.data?.text || payload.data?.body || payload.body || '';

      if (phone && message && isSentByApp(phone, message)) {
        console.log(`⏭️  Webhook KELUAR dilewati — duplikat dari app (${phone}): ${message.slice(0, 40)}...`);
        return res.status(200).send('OK');
      }

      if (phone && message) {
        appendSession({
          id: 'sess_' + Date.now(), time: new Date(timestamp).toISOString(),
          templateName: '📤 PESAN KELUAR (SYNC)',
          previewText:  message,
          woCategory:   getLastCategory(phone), 
          summary: { total: 1, success: 1, failed: 0 },
          contacts: [{ phone, status: 'success', message, messageId: payload.id || payload.data?.id || '', time: new Date(timestamp).toISOString() }],
        });
        console.log(`✅ Pesan KELUAR (dari web Maxchat) tersimpan ke ${phone}: ${message}`);
      }
      return res.status(200).send('OK');
    }

    if (payload.from || payload.data?.from) {
      const phone       = normalizePhone(String(payload.from || payload.data?.from).replace('@c.us', ''));
      const contactName = payload.username || payload.contact?.name || '';
      const msgType     = payload.msgType || payload.type || 'text';
      const attachUrl   = payload.attachmentUrl || payload.data?.attachmentUrl || '';
      const isMedia     = !!attachUrl || ['image', 'video', 'audio', 'document', 'sticker', 'voice'].includes(msgType);
      const textRaw     = typeof payload.text === 'string' ? payload.text : (payload.text?.body || payload.data?.text || payload.body || '');

      let locationText = '';
      if (msgType === 'location' && (payload.location || payload.data?.location)) {
        try {
          const rawLoc = payload.location || payload.data?.location;
          const loc    = typeof rawLoc === 'string' ? JSON.parse(rawLoc) : rawLoc;
          const mapsUrl = (loc.latitude && loc.longitude) ? `https://maps.google.com/?q=${loc.latitude},${loc.longitude}` : '';
          locationText = `📍 ${loc.name || 'Lokasi'}${loc.address ? ' - ' + loc.address : ''}${mapsUrl ? ' | ' + mapsUrl : ''}`;
        } catch (e) {
          locationText = '📍 [Lokasi]';
        }
      }

      const message     = textRaw || locationText || (isMedia ? '[' + msgType + ']' : '');

      if (phone && message) {
        appendSession({
          id: 'sess_' + Date.now(), time: new Date(timestamp).toISOString(),
          templateName: '📥 PESAN MASUK',
          previewText:  message,
          woCategory:   getLastCategory(phone), 
          summary: { total: 1, success: 1, failed: 0 },
          contacts: [{
            phone, name: contactName, status: 'success', message,
            msgType, attachmentUrl: attachUrl,
            messageId: payload.id || payload.data?.id || '',
            time: new Date(timestamp).toISOString(),
          }],
        });
        console.log(`✅ Pesan MASUK [${msgType}] dari ${contactName} (${phone}): ${message}${attachUrl ? ' | ' + attachUrl : ''}`);
      }
      return res.status(200).send('OK');
    }

  } catch (err) {
    console.error('❌ Webhook error:', err);
  }
  res.status(200).send('OK');
});

// ============================================================
// FRONTEND
// ============================================================
app.get('/login', (req, res) => {
  if (req.session?.loggedIn) return res.redirect('/');
  res.send(`<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Login — Maxchat Command Center</title>
<link rel="icon" type="image/png" href="/logo-blue.png">
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:linear-gradient(135deg,#002d6b,#004aad);height:100vh;display:flex;align-items:center;justify-content:center}
.box{background:#fff;border-radius:12px;padding:32px;width:320px;box-shadow:0 10px 40px rgba(0,0,0,.25)}
.logo{width:48px;height:48px;border-radius:10px;display:flex;align-items:center;justify-content:center;margin:0 auto 12px;overflow:hidden}
.logo img{width:100%;height:100%;object-fit:contain}
h1{font-size:15px;text-align:center;color:#1a2332;margin-bottom:2px}
p{font-size:11px;text-align:center;color:#8a94a6;margin-bottom:18px}
label{font-size:10px;font-weight:700;text-transform:uppercase;color:#6b7a8d;display:block;margin-bottom:4px}
input{width:100%;border:1.5px solid #dde3ec;border-radius:7px;padding:9px 11px;font-size:13px;margin-bottom:12px;outline:none}
input:focus{border-color:#0062cc}
button{width:100%;padding:10px;border:none;border-radius:8px;background:linear-gradient(135deg,#003d82,#0062cc);color:#fff;font-weight:700;font-size:13px;cursor:pointer}
.err{background:#fee2e2;color:#991b1b;font-size:11px;padding:8px 10px;border-radius:6px;margin-bottom:12px;display:none}
</style></head>
<body>
<div class="box">
  <div class="logo"><img src="/logo-blue.png" alt="Logo"></div>
  <h1>Maxchat Command Center</h1>
  <p>PLN UP2D Banten</p>
  <div class="err" id="err">Username atau password salah.</div>
  <form method="POST" action="/login">
    <label>Username</label>
    <input type="text" name="username" required autofocus>
    <label>Password</label>
    <input type="password" name="password" required>
    <button type="submit">Masuk</button>
  </form>
</div>
<script>
if (location.search.includes('failed')) document.getElementById('err').style.display='block';
</script>
</body></html>`);
});

app.post('/login', (req, res) => {
  const cfg = loadConfig();
  const { username, password } = req.body;
  if (username === cfg.loginUsername && password === cfg.loginPassword) {
    req.session.loggedIn = true;
    req.session.username = username;
    return res.redirect('/');
  }
  res.redirect('/login?failed=1');
});

app.post('/logout', (req, res) => {
  req.session.destroy(() => res.redirect('/login'));
});

app.get('/', (req, res) => res.send(`<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Maxchat — Command Center PLN UP2D Banten</title>
<link rel="icon" type="image/png" href="/logo-blue.png">
<script src="https://unpkg.com/lucide@latest/dist/umd/lucide.js"></script>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}

i[data-lucide]{display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;line-height:0}
i[data-lucide] svg{width:1em;height:1em;stroke-width:2}

body.light{
  --bg:#eef2f7;--sidebar-from:#002d6b;--sidebar-to:#004aad;
  --main-bg:#eef2f7;--topbar-bg:#fff;--topbar-border:#e2e8f0;--topbar-text:#1a2332;--topbar-sub:#8a94a6;
  --card-bg:#fff;--card-shadow:0 1px 3px rgba(0,0,0,.07);--card-border:transparent;
  --text:#1a2332;--text-muted:#6b7a8d;--text-faint:#8a94a6;
  --input-bg:#fff;--input-border:#dde3ec;--input-border-focus:#0062cc;--input-text:#1a2332;--textarea-bg:#f8fafc;
  --dd-bg:#fff;--dd-border:#dde3ec;--dd-btn-bg:#f8fafc;--dd-item-hover:#f0f7ff;
  --params-bg:#fffbeb;--params-border:#fde68a;--params-hd:#92400e;
  --counter-bg:#eef2f7;--counter-text:#003d82;
  --log-th:#f0f4f8;--log-border:#f5f7fa;
  --tpl-item-bg:#fafbfc;--tpl-item-border:#eef2f7;
  --pb-bg:#fafbfc;--pb-border:#e2e8f0;
  --tag-bg:#dbeafe;--tag-text:#1d4ed8;--tag-del:#60a5fa;
  --toggle-bg:#f8fafc;--toggle-border:#dde3ec;--toggle-active-bg:#eff6ff;--toggle-active-border:#0062cc;--toggle-active-text:#0062cc;
  --empty-text:#b0bac8;--hr-color:#eef2f7;
  --modal-bg:#fff;--modal-overlay:rgba(0,0,0,.45);--toast-bg:#1a2332;
  --wa-bg:#efeae2;--wa-bubble-out:#d9fdd3;--wa-bubble-in:#fff;--wa-time:#667781;--wa-tick:#53bdeb;
  --wa-header-bg:#008069;--wa-input-bg:#f0f2f5;--wa-input-inner:#fff;
  --chat-sidebar-bg:#fff;--chat-sidebar-border:#f0f0f0;--chat-item-hover:#f5f5f5;--chat-item-active:#e8f5e9;
  --settings-input-bg:#f8fafc;
}
body.dark{
  --bg:#0f1117;--sidebar-from:#0a1628;--sidebar-to:#0d1f3c;
  --main-bg:#0f1117;--topbar-bg:#161b27;--topbar-border:#252d3d;--topbar-text:#e2e8f0;--topbar-sub:#64748b;
  --card-bg:#161b27;--card-shadow:0 1px 3px rgba(0,0,0,.3);--card-border:#252d3d;
  --text:#e2e8f0;--text-muted:#94a3b8;--text-faint:#64748b;
  --input-bg:#1e2536;--input-border:#2d3748;--input-border-focus:#3b82f6;--input-text:#e2e8f0;--textarea-bg:#1a2032;
  --dd-bg:#1e2536;--dd-border:#2d3748;--dd-btn-bg:#1a2032;--dd-item-hover:#243048;
  --params-bg:#1f1a0e;--params-border:#4a3800;--params-hd:#fbbf24;
  --counter-bg:#1e2536;--counter-text:#93c5fd;
  --log-th:#1a2032;--log-border:#1e2536;
  --tpl-item-bg:#1a2032;--tpl-item-border:#252d3d;
  --pb-bg:#1a2032;--pb-border:#2d3748;
  --tag-bg:#1e3a5f;--tag-text:#93c5fd;--tag-del:#3b82f6;
  --toggle-bg:#1a2032;--toggle-border:#2d3748;--toggle-active-bg:#1e3a5f;--toggle-active-border:#3b82f6;--toggle-active-text:#60a5fa;
  --empty-text:#475569;--hr-color:#252d3d;
  --modal-bg:#161b27;--modal-overlay:rgba(0,0,0,.7);--toast-bg:#0a1628;
  --wa-bg:#0b141a;--wa-bubble-out:#005c4b;--wa-bubble-in:#202c33;--wa-time:#8696a0;--wa-tick:#53bdeb;
  --wa-header-bg:#1f2c34;--wa-input-bg:#1f2c34;--wa-input-inner:#2a3942;
  --chat-sidebar-bg:#111b21;--chat-sidebar-border:#2a3942;--chat-item-hover:#202c33;--chat-item-active:#2a3942;
  --settings-input-bg:#1e2536;
}

body{font-family:'Segoe UI',system-ui,sans-serif;background:var(--bg);color:var(--text);display:flex;flex-direction:column;transition:background .2s,color .2s}
.layout{display:flex;flex:1;overflow:hidden}

.sidebar{width:152px;min-width:152px;max-width:152px;flex-shrink:0;background:linear-gradient(180deg,var(--sidebar-from) 0%,var(--sidebar-to) 100%);display:flex;flex-direction:column;white-space:nowrap;overflow-x:hidden}
.sidebar-brand{padding:14px 12px 11px;border-bottom:1px solid rgba(255,255,255,.1);display:flex;flex-direction:column;align-items:flex-start}
.sidebar-brand .logo{width:34px;height:34px;border-radius:8px;display:flex;align-items:center;justify-content:center;margin-bottom:6px;flex-shrink:0;overflow:hidden}
.sidebar-brand .logo img{width:100%;height:100%;object-fit:contain}
.sidebar-brand h1{font-size:11px;font-weight:800;color:#fff;line-height:1.35;white-space:nowrap}
.sidebar-brand p{font-size:9px;color:rgba(255,255,255,.45);margin-top:1px;white-space:nowrap}
.sidebar-nav{padding:7px 6px;flex:1}
.nav-item{display:flex;align-items:center;gap:8px;padding:7px 8px;border-radius:7px;cursor:pointer;color:rgba(255,255,255,.6);font-size:12px;font-weight:500;transition:all .15s;margin-bottom:2px;border:none;background:transparent;width:100%;text-align:left;white-space:nowrap}
.nav-item:hover{background:rgba(255,255,255,.1);color:#fff}
.nav-item.active{background:rgba(255,255,255,.18);color:#fff;font-weight:700}
.nav-item .icon,i[data-lucide].icon{font-size:14px;width:16px;height:16px;color:inherit}
.nav-item svg{width:15px!important;height:15px!important}
.nav-toggle svg{width:15px!important;height:15px!important}
.nav-group{margin-bottom:2px}
.nav-toggle{display:flex;align-items:center;gap:8px;padding:7px 8px;border-radius:7px;cursor:pointer;color:rgba(255,255,255,.6);font-size:12px;font-weight:500;transition:all .15s;border:none;background:transparent;width:100%;text-align:left;white-space:nowrap}
.nav-toggle:hover{background:rgba(255,255,255,.1);color:#fff}
.nav-toggle.active{background:rgba(255,255,255,.18);color:#fff;font-weight:700}
.nav-chevron{margin-left:auto;transition:transform .2s;flex-shrink:0;width:13px!important;height:13px!important}
.nav-chevron svg{width:13px!important;height:13px!important}
.nav-chevron.open{transform:rotate(180deg)}
.nav-subitems{display:none;flex-direction:column;margin-top:2px}
.nav-subitems.open{display:flex}
.nav-subitem{display:flex;align-items:center;gap:8px;padding:6px 8px;border-radius:6px;cursor:pointer;color:rgba(255,255,255,.5);font-size:11.5px;font-weight:500;transition:all .15s;border:none;background:transparent;width:100%;text-align:left;white-space:nowrap;margin-bottom:1px}
.nav-subitem svg{width:15px!important;height:15px!important}
.nav-subitem:hover{background:rgba(255,255,255,.1);color:#fff}
.nav-subitem.active{background:rgba(255,255,255,.18);color:#fff;font-weight:700}
.settings-section-panel{display:flex;flex-direction:column;gap:12px}
.sidebar-footer{padding:8px 6px;border-top:1px solid rgba(255,255,255,.1)}
.theme-toggle{display:flex;align-items:center;gap:9px;padding:7px 10px;border-radius:7px;cursor:pointer;color:rgba(255,255,255,.6);font-size:11px;font-weight:500;border:none;background:transparent;width:100%;white-space:nowrap;transition:all .15s}
.theme-toggle:hover{background:rgba(255,255,255,.1);color:#fff}

.main{flex:1;display:flex;flex-direction:column;overflow:hidden;background:var(--main-bg)}
.topbar{background:var(--topbar-bg);border-bottom:1px solid var(--topbar-border);padding:0 18px;height:44px;display:flex;align-items:center;gap:8px;flex-shrink:0}
.topbar h2{font-size:13px;font-weight:700;color:var(--topbar-text)}
.topbar p{font-size:11px;color:var(--topbar-sub);margin-left:auto}
.content{flex:1;overflow-y:auto;padding:13px 16px}
.page{display:none}.page.active{display:block}

.blast-grid{display:grid;grid-template-columns:300px 1fr;gap:12px}
.col{display:flex;flex-direction:column;gap:11px}

.card{background:var(--card-bg);border-radius:10px;padding:14px;box-shadow:var(--card-shadow);border:1px solid var(--card-border, transparent)}
.card-hd{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--text-muted);margin-bottom:9px;display:flex;align-items:center;gap:6px}
.card-hd i[data-lucide]{font-size:12px;color:var(--text-muted)}

.dd-wrap{position:relative}
.dd-btn{width:100%;display:flex;align-items:center;justify-content:space-between;padding:7px 10px;border:1.5px solid var(--dd-border);border-radius:7px;background:var(--dd-btn-bg);cursor:pointer;font-size:12px;color:var(--text);transition:border-color .2s;text-align:left;gap:8px}
.dd-btn:hover{border-color:#93c5fd}
.dd-btn.open{border-color:var(--input-border-focus);background:var(--input-bg)}
.dd-btn .arrow{font-size:9px;color:var(--text-faint);transition:transform .2s;flex-shrink:0}
.dd-btn.open .arrow{transform:rotate(180deg)}
.dd-placeholder{color:var(--text-faint);flex:1;font-size:12px}
.dd-selected{font-weight:600;flex:1;font-size:12px;color:var(--text)}
.dd-menu{position:absolute;top:calc(100% + 3px);left:0;right:0;background:var(--dd-bg);border:1.5px solid var(--dd-border);border-radius:9px;box-shadow:0 6px 20px rgba(0,0,0,.15);z-index:200;overflow:hidden;display:none}
.dd-menu.open{display:block}
.dd-item{padding:9px 11px;cursor:pointer;border-bottom:1px solid var(--log-border);transition:background .15s}
.dd-item:last-child{border-bottom:none}
.dd-item:hover,.dd-item.active{background:var(--dd-item-hover)}
.dd-item-name{font-size:12px;font-weight:700;color:var(--text)}
.dd-item-id{font-size:10px;color:var(--text-faint);font-family:monospace;margin-top:1px;word-break:break-all}
.dd-item-badges{display:flex;gap:4px;margin-top:4px;flex-wrap:wrap}

.badge{display:inline-block;padding:2px 6px;border-radius:20px;font-size:9px;font-weight:700}
.badge-blue{background:#dbeafe;color:#1d4ed8}
.badge-yellow{background:#fef9c3;color:#854d0e}
.badge-gray{background:#f1f5f9;color:#64748b}
.badge-green{background:#dcfce7;color:#166534}
.badge-red{background:#fee2e2;color:#991b1b}

.field-label{display:block;font-size:10px;font-weight:600;color:var(--text-muted);margin-bottom:3px;text-transform:uppercase;letter-spacing:.04em}
.field-hint{font-size:10px;color:var(--text-faint);margin-top:3px}
.field-row{display:grid;grid-template-columns:1fr 1fr;gap:9px}
.fg{margin-bottom:10px}
select,input[type=text],input[type=password]{width:100%;border:1.5px solid var(--input-border);border-radius:7px;padding:6px 9px;font-size:12px;color:var(--input-text);background:var(--input-bg);outline:none;transition:border-color .2s}
select:focus,input[type=text]:focus,input[type=password]:focus{border-color:var(--input-border-focus)}
textarea{width:100%;border:1.5px solid var(--input-border);border-radius:7px;padding:8px 10px;font-family:'Consolas',monospace;font-size:12px;resize:none;background:var(--textarea-bg);color:var(--input-text);outline:none;transition:border-color .2s;height:110px}
textarea:focus{border-color:var(--input-border-focus);background:var(--input-bg)}

.params-box{background:var(--params-bg);border:1.5px solid var(--params-border);border-radius:9px;padding:11px}
.params-hd{font-size:10px;font-weight:700;color:var(--params-hd);text-transform:uppercase;letter-spacing:.05em;margin-bottom:8px;display:flex;align-items:center;gap:5px}
.param-item{margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid var(--params-border)}
.param-item:last-child{margin-bottom:0;padding-bottom:0;border-bottom:none}

.counter-row{display:flex;align-items:center;justify-content:space-between;margin-top:4px}
.counter-badge{background:var(--counter-bg);border-radius:5px;padding:2px 8px;font-size:11px;font-weight:700;color:var(--counter-text)}

.btn{padding:6px 13px;border-radius:7px;font-size:11px;font-weight:600;cursor:pointer;border:none;transition:opacity .15s,background .15s;color:var(--text);display:inline-flex;align-items:center;gap:5px}
.btn:hover{opacity:.85}
.btn:active{transform:scale(.98)}
.btn-primary{background:linear-gradient(135deg,#003d82,#0062cc);color:#fff}
.btn-danger{background:#fee2e2;color:#dc2626}
.btn-warning{background:#fef9c3;color:#854d0e}
.btn-ghost{background:var(--log-th);color:var(--text-muted)}
.btn-sm{padding:3px 9px;font-size:10px}
.btn-blast{width:100%;padding:10px;font-size:13px;font-weight:700;border-radius:8px;border:none;cursor:pointer;background:linear-gradient(135deg,#003d82,#0062cc);color:#fff;display:flex;align-items:center;justify-content:center;gap:7px;transition:opacity .2s;margin-top:8px}
.btn-blast:hover{opacity:.9}
.btn-blast:disabled{opacity:.5;cursor:not-allowed}

.wa-phone-wrap{display:flex;justify-content:center;align-items:flex-start;padding:4px 0}
.wa-phone{width:100%;max-width:280px;aspect-ratio:9/16;border-radius:14px;overflow:hidden;display:flex;flex-direction:column;box-shadow:0 4px 20px rgba(0,0,0,.18)}
.wa-screen{background:var(--wa-bg);flex:1;display:flex;flex-direction:column;min-height:0;position:relative}
.wa-header{background:var(--wa-header-bg);padding:9px 10px;display:flex;align-items:center;gap:9px;flex-shrink:0;color:#fff}
.wa-back{font-size:18px;color:#fff;cursor:pointer}
.wa-avatar{width:32px;height:32px;border-radius:50%;background:#25d366;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#fff;flex-shrink:0}
.wa-header-info{flex:1;min-width:0}
.wa-header-name{font-size:13px;font-weight:600;color:#fff;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.wa-header-status{font-size:10px;color:rgba(255,255,255,.85)}
.wa-action{font-size:16px;color:#fff;cursor:pointer}
.wa-action + .wa-action{margin-left:2px}
.wa-msgs{flex:1;padding:10px 8px;display:flex;flex-direction:column;gap:6px;overflow-y:auto;background-image:radial-gradient(rgba(0,0,0,.04) 1px, transparent 1px);background-size:18px 18px}
body.dark .wa-msgs{background-image:radial-gradient(rgba(255,255,255,.025) 1px, transparent 1px)}
.wa-date-chip{text-align:center;margin:4px 0 8px}
.wa-date-chip span{background:rgba(255,255,255,.9);color:#54656f;font-size:10px;font-weight:500;padding:3px 10px;border-radius:8px;box-shadow:0 1px 1px rgba(0,0,0,.13)}
body.dark .wa-date-chip span{background:rgba(30,40,50,.85);color:#cbd5d8}
.wa-bubble-wrap{display:flex;justify-content:flex-end}
.wa-bubble{position:relative;background:var(--wa-bubble-out);border-radius:8px 8px 0 8px;padding:6px 9px 6px 9px;max-width:85%;min-width:80px;box-shadow:0 1px 1px rgba(0,0,0,.13);font-size:12.5px;color:#111b21;line-height:1.45;white-space:pre-wrap;word-break:break-word}
body.dark .wa-bubble{color:#e9edef}
.wa-bubble::before{content:"";position:absolute;right:-7px;top:0;width:8px;height:13px;background:var(--wa-bubble-out);clip-path:polygon(0 0,100% 0,0 100%)}
.wa-bubble-text{display:inline}
.wa-bubble-meta{position:absolute;right:6px;bottom:3px;display:inline-flex;align-items:center;gap:2px;background:linear-gradient(to left, var(--wa-bubble-out) 60%, transparent);padding-left:14px}
body.dark .wa-bubble-meta{background:linear-gradient(to left, var(--wa-bubble-out) 60%, transparent)}
.wa-bubble-time{font-size:9px;color:var(--wa-time);font-weight:500}
.wa-bubble-tick{font-size:12px;color:var(--wa-tick)}
.wa-input-bar{background:var(--wa-input-bg);padding:6px 8px;display:flex;align-items:center;gap:6px;border-top:1px solid rgba(0,0,0,.06);flex-shrink:0}
body.dark .wa-input-bar{border-top-color:rgba(255,255,255,.04)}
.wa-input-icon{font-size:18px;color:#54656f;cursor:pointer}
body.dark .wa-input-icon{color:#8696a0}
.wa-input-bar-inner{flex:1;background:var(--wa-input-inner);border-radius:20px;padding:6px 12px;font-size:11px;color:#54656f}
body.dark .wa-input-bar-inner{color:#8696a0}
.wa-input-bar-btn{width:34px;height:34px;border-radius:50%;background:#00a884;display:flex;align-items:center;justify-content:center;font-size:16px;color:#fff;flex-shrink:0;cursor:pointer}
.preview-empty{color:var(--empty-text);font-style:italic;font-size:12px;padding:16px;text-align:center}

.progress-wrap{display:none;margin-top:9px}
.progress-bg{background:var(--input-border);border-radius:99px;height:6px;overflow:hidden;position:relative}
.progress-fill{
  height:100%;position:relative;overflow:hidden;border-radius:99px;width:0%;
  background:linear-gradient(90deg,#003d82,#0062cc);
  transition:width .45s cubic-bezier(.22,.61,.36,1);
}
.progress-fill::after{
  content:'';position:absolute;top:0;bottom:0;left:0;width:70px;
  background:linear-gradient(90deg,transparent,rgba(255,255,255,.45),transparent);
  animation:progressShimmer 1.35s linear infinite;
}
.progress-fill.is-done::after{animation:none;opacity:0}
.progress-fill.is-done{background:linear-gradient(90deg,#0d9155,#16a34a)}
@keyframes progressShimmer{
  0%{transform:translateX(-90px)}
  100%{transform:translateX(calc(100% + 90px))}
}
.progress-lbl{font-size:10.5px;color:var(--text-faint);margin-top:6px;text-align:center;display:flex;align-items:center;justify-content:center;gap:6px;letter-spacing:.1px}
.progress-lbl b{color:var(--text);font-weight:700}
.progress-pct{color:var(--text-faint);opacity:.8;font-variant-numeric:tabular-nums}
.progress-dot{width:6px;height:6px;border-radius:50%;background:#0062cc;flex-shrink:0;animation:progressPulse 1.1s ease-in-out infinite}
.progress-dot.is-done{background:#16a34a;animation:none}
@keyframes progressPulse{
  0%,100%{opacity:.35;transform:scale(.8)}
  50%{opacity:1;transform:scale(1.15)}
}

.summary-row{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:8px}
.sum-box{border-radius:7px;padding:9px;text-align:center}
.sum-box .val{font-size:20px;font-weight:800;line-height:1}
.sum-box .lbl{font-size:9px;margin-top:3px;font-weight:700;text-transform:uppercase}
.box-total{background:var(--log-th);color:var(--text)}
.box-ok{background:#dcfce7;color:#166534}
.box-err{background:#fee2e2;color:#991b1b}
.log-wrap{overflow:auto;max-height:220px}
.log-table{width:100%;border-collapse:collapse;font-size:11px}
.log-table th{background:var(--log-th);padding:5px 8px;text-align:left;font-weight:700;color:var(--text-muted);font-size:9px;text-transform:uppercase;position:sticky;top:0}
.log-table td{padding:5px 8px;border-bottom:1px solid var(--log-border);color:var(--text)}
.log-table tr:last-child td{border-bottom:none}
.result-split{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.result-col{border:1.5px solid var(--card-border, var(--input-border));border-radius:8px;overflow:hidden}
.result-col-err{border-color:#fca5a5}
.result-col-ok{border-color:#86efac}
.result-col-hd{padding:6px 9px;font-size:10px;font-weight:700;text-transform:uppercase;display:flex;align-items:center;gap:5px}
.result-col-hd-err{background:#fef2f2;color:#dc2626}
.result-col-hd-ok{background:#f0fdf4;color:#16a34a}
body.dark .result-col-hd-err{background:#2a1414;color:#f87171}
body.dark .result-col-hd-ok{background:#10241a;color:#4ade80}
@media (max-width:760px){.result-split{grid-template-columns:1fr}}
.s-ok{color:#16a34a;font-weight:700}.s-err{color:#dc2626;font-weight:700}

.settings-sections{display:flex;flex-direction:column;gap:12px}
.settings-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;align-items:stretch}
.tpl-list-item{display:flex;align-items:flex-start;gap:8px;padding:9px 11px;border:1.5px solid var(--tpl-item-border);border-radius:8px;margin-bottom:6px;background:var(--tpl-item-bg)}
.tpl-list-item:last-child{margin-bottom:0}
.tpl-list-info{flex:1;min-width:0}
.tpl-list-name{font-size:12px;font-weight:700;color:var(--text)}
.tpl-list-id{font-size:10px;color:var(--text-faint);font-family:monospace;word-break:break-all;margin-top:1px}
.tpl-list-actions{display:flex;gap:4px;flex-shrink:0;flex-direction:column}
.empty{text-align:center;color:var(--empty-text);font-size:12px;padding:14px}
.settings-input-group{margin-bottom:10px}
.settings-input-group label{display:block;font-size:10px;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:4px}
.settings-input-group input{background:var(--settings-input-bg)}
.settings-token-row{display:flex;gap:6px;align-items:center}
.settings-token-row input{flex:1}
.qr-add-row{display:flex;gap:6px;margin-bottom:10px}
.qr-add-row input{flex:1;background:var(--settings-input-bg)}
.qr-list-wrap{max-height:320px;overflow-y:auto}

.card-sub{font-size:10.5px;color:var(--text-faint);margin:-5px 0 12px;line-height:1.5}
.card-danger{border-color:#fca5a5}
body.dark .card-danger{border-color:#7f1d1d}
.card-danger .card-hd{color:#dc2626}
.card-danger .card-hd i[data-lucide]{color:#dc2626}
.danger-row{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;padding:11px 12px;border-radius:8px;background:var(--tpl-item-bg);border:1px solid var(--tpl-item-border)}
.danger-row + .danger-row{margin-top:8px}
.danger-row-info{flex:1;min-width:200px}
.danger-row-title{font-size:12px;font-weight:700;color:var(--text)}
.danger-row-desc{font-size:10.5px;color:var(--text-faint);margin-top:2px;line-height:1.5}

.status-pill{display:inline-flex;align-items:center;gap:5px;padding:3px 9px;border-radius:20px;font-size:10px;font-weight:700}
.status-pill.on{background:#dcfce7;color:#166534}
.status-pill.off{background:#fee2e2;color:#991b1b}
.status-pill .dot{width:6px;height:6px;border-radius:50%;background:currentColor;flex-shrink:0}

.db-stats-row{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-bottom:12px}
.db-stat-box{position:relative;border-radius:7px;background:var(--log-th);overflow:hidden}
.db-stat-box::before{content:'';display:block;padding-top:100%}
.db-stat-box-inner{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:6px;text-align:center}
.db-stat-box .val{font-size:16px;font-weight:800;line-height:1.2;color:var(--text);word-break:break-word}
.db-stat-box .lbl{font-size:9px;margin-top:4px;font-weight:700;text-transform:uppercase;color:var(--text-muted)}

.backup-info-box{border:1px solid var(--tpl-item-border);border-radius:8px;padding:10px 12px;background:var(--tpl-item-bg);margin-bottom:12px}
.backup-info-row{display:flex;align-items:center;justify-content:space-between;gap:10px;font-size:11px;color:var(--text)}
.backup-info-row + .backup-info-row{margin-top:7px;padding-top:7px;border-top:1px solid var(--hr-color)}
.backup-info-row .k{color:var(--text-faint);font-size:10.5px}
.backup-info-row .v{font-weight:600;text-align:right}

.template-grid{display:grid;grid-template-columns:1fr 320px;gap:12px}

.chat-layout{display:grid;grid-template-columns:260px 1fr;gap:0;height:calc(100vh - 44px - 26px);border-radius:10px;overflow:hidden;box-shadow:var(--card-shadow)}
.chat-contacts{background:var(--chat-sidebar-bg);border-right:1px solid var(--chat-sidebar-border);display:flex;flex-direction:column;overflow:hidden}
.chat-contacts-hd{padding:12px;border-bottom:1px solid var(--chat-sidebar-border);font-size:11px;font-weight:700;color:var(--text);display:flex;align-items:center;justify-content:space-between;gap:6px}

.chat-tabs{display:flex;padding:8px 8px 0;gap:4px}
.chat-tab{flex:1;padding:5px;border:1px solid var(--input-border);background:var(--input-bg);color:var(--text-muted);border-radius:6px;font-size:10px;font-weight:600;cursor:pointer;transition:all .15s;text-align:center}
.chat-tab:hover{background:var(--chat-item-hover)}
.chat-tab.active{background:var(--sidebar-to);color:#fff;border-color:var(--sidebar-to)}
.chat-tab span{background:#ef4444;color:#fff;border-radius:10px;padding:1px 5px;font-size:8px;margin-left:3px}

.chat-contacts-search{padding:8px;position:relative}
.chat-contacts-search input{width:100%;border:1.5px solid var(--input-border);border-radius:20px;padding:5px 28px 5px 12px;font-size:11px;background:var(--input-bg);color:var(--input-text);outline:none}
.search-clear{position:absolute;right:16px;top:50%;transform:translateY(-50%);border:none;background:transparent;cursor:pointer;color:var(--text-faint);display:none;align-items:center;justify-content:center}
.search-clear:hover{color:#dc2626}

.chat-contacts-list{flex:1;overflow-y:auto}
.chat-contact-item{padding:10px 12px;cursor:pointer;border-bottom:1px solid var(--chat-sidebar-border);transition:background .15s;display:flex;align-items:center;gap:9px}
.chat-contact-item:hover{background:var(--chat-item-hover)}
.chat-contact-item.active{background:var(--chat-item-active)}
.chat-contact-avatar{width:34px;height:34px;border-radius:50%;background:linear-gradient(135deg,#003d82,#0062cc);display:flex;align-items:center;justify-content:center;color:#fff;font-size:11px;font-weight:700;flex-shrink:0}
.chat-contact-info{flex:1;min-width:0;display:flex;flex-direction:column;gap:2px}
.chat-contact-name-text{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:12px;font-weight:600;color:var(--text)}
.chat-contact-last{font-size:10px;color:var(--text-faint);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;display:flex;align-items:center;gap:3px}
.chat-contact-right{align-self:stretch;display:flex;flex-direction:column;justify-content:space-between;align-items:flex-end;flex-shrink:0;padding:1px 0}
.right-top{display:flex;align-items:center;gap:4px}
.right-bot{display:flex;align-items:center;gap:4px}
.chat-time-text{font-size:9px;color:var(--text-faint)}
.chat-unread-dot{width:9px;height:9px;background:#ef4444;border-radius:50%;flex-shrink:0}
.chat-contact-item.has-unread .chat-contact-name-text{font-weight:800;color:var(--text)}
.chat-contact-item.has-unread .chat-contact-last{color:var(--text);font-weight:600}

.chat-main{background:var(--wa-bg);display:flex;flex-direction:column;overflow:hidden}
.chat-main-header{background:var(--sidebar-to);padding:8px 12px;display:flex;align-items:center;gap:8px;flex-shrink:0;color:#fff}
.chat-main-avatar{width:32px;height:32px;border-radius:50%;background:#fff;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#003d82}
.chat-main-info{flex:1;min-width:0}
.chat-header-labels{font-size:10px;color:rgba(255,255,255,.85);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:220px;flex-shrink:1;margin-right:2px}
.chat-main-name{font-size:12px;font-weight:700;color:#fff}
.chat-main-sub{font-size:9px;color:rgba(255,255,255,.65)}
.chat-header-phone{font-family:'Courier New',monospace;font-size:11px;font-weight:700;color:#fff;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:6px;padding:4px 10px;user-select:all;-webkit-user-select:all;white-space:nowrap;flex-shrink:0}
.chat-copy-btn{display:flex;align-items:center;justify-content:center;width:24px;height:24px;flex-shrink:0;background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);border-radius:6px;color:#fff;cursor:pointer;padding:0;transition:background .15s}
.chat-copy-btn:hover{background:rgba(255,255,255,.28)}
.chat-copy-btn:active{background:rgba(255,255,255,.4)}
.chat-messages{flex:1;overflow-y:auto;padding:10px 12px;display:flex;flex-direction:column;gap:6px;background-image:radial-gradient(rgba(0,0,0,.04) 1px, transparent 1px);background-size:20px 20px}
body.dark .chat-messages{background-image:radial-gradient(rgba(255,255,255,.025) 1px, transparent 1px)}
.chat-empty,.chat-no-contact{flex:1;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:8px;color:var(--text-faint)}
.chat-empty-icon,.chat-no-contact i[data-lucide]{font-size:40px;opacity:.4}

.chat-date-sep{display:flex;justify-content:center;margin:10px 0 12px}
.chat-date-sep span{background:var(--wa-bubble-in, var(--card-bg));color:var(--wa-time, var(--text-muted));font-size:10.5px;font-weight:600;padding:4px 12px;border-radius:8px;box-shadow:0 1px 1px rgba(0,0,0,.13)}
.chat-msg-bubble{display:flex;width:100%;margin-bottom:4px}
.chat-bubble-in{background:var(--wa-bubble-in);border-radius:10px 10px 10px 2px;border:1px solid var(--input-border);color:var(--text);padding:6px 10px 4px 10px;max-width:80%;word-break:break-word;white-space:pre-wrap;font-size:12px;line-height:1.4;box-shadow:0 1px 1px rgba(0,0,0,.13)}
.chat-bubble-out{background:linear-gradient(135deg,#003d82,#0062cc);color:#fff;border-radius:10px 10px 2px 10px;padding:6px 10px 4px 10px;max-width:80%;word-break:break-word;white-space:pre-wrap;font-size:12px;line-height:1.4;box-shadow:0 1px 1px rgba(0,0,0,.13)}
.chat-msg-bubble.pending .chat-bubble-out{opacity:.72}
.chat-bubble-meta{display:flex;justify-content:flex-end;align-items:center;gap:3px;margin-top:2px;font-size:9px;opacity:.7;line-height:1}
.chat-bubble-status{font-size:10px}
.chat-bubble-status.ok{color:var(--wa-tick)}
.chat-bubble-status.err{color:#f87171}
.chat-bubble-status.pending{color:rgba(255,255,255,.75)}

.chat-compose{background:var(--topbar-bg);border-top:1px solid var(--chat-sidebar-border);padding:10px 12px 12px;display:flex;align-items:center;gap:8px;flex-shrink:0;position:relative}
.chat-compose-wrap{flex:1;background:var(--input-bg);border:1.5px solid var(--input-border);border-radius:20px;padding:6px 14px;display:flex;align-items:center;gap:6px;transition:border-color .2s;min-height:36px}
.chat-compose-wrap:focus-within{border-color:var(--input-border-focus)}
.chat-compose-textarea{flex:1;border:none;outline:none;background:transparent;font-size:12px;color:var(--input-text);resize:none;font-family:inherit;line-height:1.5;height:20px;max-height:100px;min-height:20px;overflow-y:auto;padding:0;margin:0;display:block}
.chat-compose-send{width:36px;height:36px;border-radius:50%;background:linear-gradient(135deg,#003d82,#0062cc);border:none;color:#fff;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:opacity .15s;align-self:center}
.chat-compose-send:hover{opacity:.85}
.chat-compose-send:disabled{opacity:.4;cursor:not-allowed}
.chat-compose-hint{font-size:9px;color:var(--text-faint);padding:2px 6px 0;text-align:center}
.chat-attach-btn{width:34px;height:34px;border-radius:50%;background:var(--input-bg);border:1.5px solid var(--input-border);color:var(--text-muted);cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all .15s}
.chat-attach-btn:hover{background:var(--chat-item-hover);color:var(--text)}
.chat-attach-btn.active{background:var(--toggle-active-bg);border-color:var(--toggle-active-border);color:var(--toggle-active-text)}
.chat-compose.drag-over .chat-compose-wrap{border-color:#3b82f6;background:rgba(59,130,246,.07);box-shadow:0 0 0 3px rgba(59,130,246,.15)}

.qr-menu{position:absolute;bottom:calc(100% + 8px);left:12px;width:290px;max-height:290px;overflow-y:auto;background:var(--dd-bg);border:1.5px solid var(--dd-border);border-radius:10px;box-shadow:0 6px 20px rgba(0,0,0,.18);z-index:150}
.qr-menu-hd{padding:8px 11px;font-size:9.5px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--text-muted);border-bottom:1px solid var(--log-border)}
.qr-menu-item{padding:8px 11px;font-size:11.5px;color:var(--text);cursor:pointer;border-bottom:1px solid var(--log-border);white-space:pre-wrap;line-height:1.4}
.qr-menu-item:last-child{border-bottom:none}
.qr-menu-item:hover{background:var(--dd-item-hover)}
.qr-menu-empty{padding:14px;text-align:center;font-size:11px;color:var(--empty-text)}

.param-builder{border:1.5px solid var(--pb-border);border-radius:8px;padding:10px;margin-bottom:7px;background:var(--pb-bg);position:relative}
.param-builder-hd{display:flex;align-items:center;gap:8px;margin-bottom:8px}
.param-builder-title{font-size:11px;font-weight:700;color:#3b82f6}
.param-builder .remove-param{margin-left:auto;cursor:pointer;color:#dc2626;font-size:17px;line-height:1;background:none;border:none;padding:0}
.input-type-toggle{display:flex;gap:5px;margin-top:3px}
.toggle-btn{padding:3px 10px;border-radius:5px;font-size:10px;font-weight:600;cursor:pointer;border:1.5px solid var(--toggle-border);background:var(--toggle-bg);color:var(--text-muted);transition:all .15s}
.toggle-btn.active{border-color:var(--toggle-active-border);background:var(--toggle-active-bg);color:var(--toggle-active-text)}
.tag-wrap{display:flex;flex-wrap:wrap;gap:4px;padding:5px;border:1.5px solid var(--input-border);border-radius:7px;background:var(--textarea-bg);min-height:35px}
.tag{display:flex;align-items:center;gap:3px;background:var(--tag-bg);color:var(--tag-text);border-radius:4px;padding:2px 6px;font-size:10px;font-weight:600}
.tag-del{cursor:pointer;font-size:11px;color:var(--tag-del);line-height:1}
.tag-del:hover{color:#dc2626}
.tag-input{border:none;outline:none;background:transparent;font-size:11px;min-width:80px;flex:1;padding:1px 3px;color:var(--input-text)}

.modal-overlay{position:fixed;inset:0;background:var(--modal-overlay);z-index:500;display:flex;align-items:center;justify-content:center;display:none}
.modal-overlay.open{display:flex}
.modal{background:var(--modal-bg);border-radius:12px;padding:22px;width:680px;max-width:95vw;max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.25)}
.modal-hd{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px}
.modal-title{font-size:14px;font-weight:700;color:var(--text)}
.modal-close{cursor:pointer;font-size:19px;color:var(--text-faint);background:none;border:none;line-height:1;padding:0;display:flex;align-items:center}
.modal-close:hover{color:#dc2626}
.modal-footer{display:flex;justify-content:flex-end;gap:7px;margin-top:14px;padding-top:12px;border-top:1px solid var(--hr-color)}
.unsaved-dot{display:inline-block;width:6px;height:6px;background:#f59e0b;border-radius:50%;margin-left:5px;vertical-align:middle}

@keyframes spin{to{transform:rotate(360deg)}}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.spinner{width:14px;height:14px;border:2px solid rgba(255,255,255,.4);border-top-color:#fff;border-radius:50%;animation:spin .7s linear infinite;display:none}
.toast{position:fixed;bottom:16px;right:16px;background:var(--toast-bg);color:#fff;padding:8px 14px;border-radius:7px;font-size:11px;font-weight:600;opacity:0;transform:translateY(8px);transition:opacity .3s,transform .3s;z-index:999;pointer-events:none}
.toast.show{opacity:1;transform:translateY(0)}
.toast.ok{background:#166534}.toast.err{background:#991b1b}
hr{border:none;border-top:1px solid var(--hr-color);margin:11px 0}
.nav-badge{display:inline-block;width:7px;height:7px;background:#ef4444;border-radius:50%;margin-left:4px;animation:pulse 1.5s infinite}
.export-quick-btn.active{background:var(--toggle-active-bg);border:1px solid var(--toggle-active-border);color:var(--toggle-active-text)}
</style>
</head>
<body class="light">
<div class="layout">

<aside class="sidebar">
  <div class="sidebar-brand">
    <div class="logo"><img src="/logo.png" alt="Logo"></div>
    <h1>Maxchat Comcen</h1>
    <p>PLN UP2D Banten</p>
  </div>
  <nav class="sidebar-nav">
    <button class="nav-item active" onclick="showPage('broadcast',this)"><i data-lucide="radio" class="icon"></i>Broadcast</button>
    <button class="nav-item" onclick="showPage('template',this)"><i data-lucide="clipboard-list" class="icon"></i>Template</button>
    <button class="nav-item" onclick="showPage('chat',this)"><i data-lucide="message-circle" class="icon"></i>Chat</button>
    <button class="nav-item" onclick="showPage('kontak',this)"><i data-lucide="contact" class="icon"></i>Kontak</button>
    <div class="nav-group">
      <button class="nav-toggle" id="navPengaturanToggle" onclick="toggleSettingsMenu()">
        <i data-lucide="settings" class="icon"></i>Pengaturan
        <i data-lucide="chevron-down" class="icon nav-chevron" id="settingsChevron"></i>
      </button>
      <div class="nav-subitems" id="settingsSubmenu">
        <button class="nav-subitem" onclick="openSettingsSection('api',this)"><i data-lucide="key" class="icon"></i>API Maxchat</button>
        <button class="nav-subitem" onclick="openSettingsSection('username',this)"><i data-lucide="user-cog" class="icon"></i>Username</button>
        <button class="nav-subitem" onclick="openSettingsSection('quickreply',this)"><i data-lucide="message-square-quote" class="icon"></i>Balas Cepat</button>
        <button class="nav-subitem" onclick="openSettingsSection('history',this)"><i data-lucide="calendar-x" class="icon"></i>Hapus Riwayat</button>
      </div>
    </div>
    <button class="nav-item" onclick="showPage('dashboard',this)"><i data-lucide="layout-dashboard" class="icon"></i>Dashboard</button>
    <button class="nav-item" onclick="doLogout()"><i data-lucide="log-out" class="icon"></i>Logout</button>
  </nav>
  <div class="sidebar-footer">
    <button class="theme-toggle" onclick="toggleTheme()"><i data-lucide="moon" id="themeIcon" class="icon"></i><span id="themeLabel">Mode Gelap</span></button>
  </div>
</aside>

<div class="main">
  <div class="topbar">
    <h2 id="topbarTitle">Broadcast</h2>
    <p id="topbarSub">Pilih template lalu kirim ke banyak nomor</p>
  </div>
  <div class="content">

    <div class="page active" id="page-broadcast">
      <div class="blast-grid">
        <div class="col">
          <div class="card">
            <div class="card-hd"><i data-lucide="layout-template"></i>Template</div>
            <div class="dd-wrap" id="tplWrap">
              <button class="dd-btn" id="tplBtn" onclick="toggleDD()">
                <span class="dd-placeholder" id="tplBtnLabel">— Pilih template —</span>
                <span class="arrow">▼</span>
              </button>
              <div class="dd-menu" id="tplMenu"></div>
            </div>
          </div>
          <div class="card" id="paramCard" style="display:none">
            <div class="card-hd"><i data-lucide="sliders-horizontal"></i>Parameter</div>
            <div class="params-box">
              <div class="params-hd"><i data-lucide="settings-2" style="width:11px;height:11px"></i>Isi Parameter</div>
              <div id="paramFields"></div>
            </div>
          </div>
          <div class="card">
            <div class="card-hd"><i data-lucide="users"></i>Nomor Tujuan</div>
            <textarea id="numbers" placeholder="08123456789&#10;628123456789&#10;&#10;Format nomer lapor (spasi = pemisah):&#10;G56555625 081278887722&#10;G56555626 081234567890&#10;&#10;Juga bisa dipisah koma atau titik koma:&#10;08123456789, 08987654321" oninput="updateCounter()"></textarea>
            <div class="counter-row">
              <span style="font-size:10px;color:var(--text-faint);display:flex;align-items:center;gap:4px"><i data-lucide="check-circle" style="width:11px;height:11px"></i>Normalisasi 628xx · "KODE NOMER" didukung</span>
              <span class="counter-badge" id="counter">0 nomor</span>
            </div>
            <button class="btn-blast" id="btnBlast" onclick="startBlast()">
              <div class="spinner" id="spinner"></div>
              <i data-lucide="rocket" style="width:15px;height:15px"></i>
              <span id="btnText">Kirim Pesan</span>
            </button>
            <div class="progress-wrap" id="progressWrap">
              <div class="progress-bg"><div class="progress-fill" id="progressFill"></div></div>
              <div class="progress-lbl" id="progressLbl">Mengirim...</div>
            </div>
          </div>
        </div>

        <div class="col">
          <div class="card">
            <div class="card-hd"><i data-lucide="smartphone"></i>Preview Pesan (tampilan WhatsApp)</div>
            <div id="previewArea">
              <div class="preview-empty" id="previewEmpty">← Pilih template untuk melihat preview</div>
              <div class="wa-phone-wrap" id="waPhoneWrap" style="display:none">
                <div class="wa-phone">
                  <div class="wa-screen">
                    <div class="wa-header">
                      <i data-lucide="chevron-left" class="wa-back"></i>
                      <div class="wa-avatar">PLN</div>
                      <div class="wa-header-info">
                        <div class="wa-header-name">Command Center PLN</div>
                        <div class="wa-header-status">online</div>
                      </div>
                      <i data-lucide="video" class="wa-action"></i>
                      <i data-lucide="phone" class="wa-action"></i>
                      <i data-lucide="more-vertical" class="wa-action"></i>
                    </div>
                    <div class="wa-msgs">
                      <div class="wa-date-chip"><span id="waDateChip">Hari ini</span></div>
                      <div class="wa-bubble-wrap">
                        <div class="wa-bubble">
                          <span class="wa-bubble-text" id="waBubbleText">—</span>
                          <span class="wa-bubble-meta">
                            <span class="wa-bubble-time" id="waTimeChip">00:00</span>
                            <i data-lucide="check-check" class="wa-bubble-tick"></i>
                          </span>
                        </div>
                      </div>
                    </div>
                    <div class="wa-input-bar">
                      <i data-lucide="smile" class="wa-input-icon"></i>
                      <div class="wa-input-bar-inner">Ketik pesan...</div>
                      <i data-lucide="paperclip" class="wa-input-icon"></i>
                      <i data-lucide="camera" class="wa-input-icon"></i>
                      <div class="wa-input-bar-btn"><i data-lucide="mic" style="color:#fff"></i></div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div class="card" id="resultCard" style="display:none">
            <div class="card-hd"><i data-lucide="list-checks"></i>Hasil Pengiriman</div>
            <div class="summary-row">
              <div class="sum-box box-total"><div class="val" id="sumTotal">0</div><div class="lbl">Total</div></div>
              <div class="sum-box box-ok"><div class="val" id="sumOk">0</div><div class="lbl">Berhasil</div></div>
              <div class="sum-box box-err"><div class="val" id="sumErr">0</div><div class="lbl">Gagal</div></div>
            </div>
            <div class="result-split">
              <div class="result-col result-col-err">
                <div class="result-col-hd result-col-hd-err"><i data-lucide="x-circle" style="width:12px;height:12px"></i>Gagal (<span id="failCount">0</span>)</div>
                <div class="log-wrap">
                  <table class="log-table">
                    <thead><tr><th>#</th><th>Nomor</th><th>Keterangan</th></tr></thead>
                    <tbody id="logBodyFail"></tbody>
                  </table>
                  <div class="empty" id="logFailEmpty" style="display:none;padding:14px;text-align:center;font-size:11px">Tidak ada yang gagal 🎉</div>
                </div>
              </div>
              <div class="result-col result-col-ok">
                <div class="result-col-hd result-col-hd-ok"><i data-lucide="check-circle" style="width:12px;height:12px"></i>Berhasil (<span id="okCount">0</span>)</div>
                <div class="log-wrap">
                  <table class="log-table">
                    <thead><tr><th>#</th><th>Nomor</th><th>Keterangan</th></tr></thead>
                    <tbody id="logBodySuccess"></tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
          <div class="card" id="emptyResult" style="flex:1;display:flex;align-items:center;justify-content:center;min-height:100px">
            <div style="text-align:center;color:var(--empty-text)"><i data-lucide="inbox" style="font-size:28px;margin-bottom:5px"></i><div style="font-size:11px;margin-top:6px">Hasil pengiriman akan tampil di sini</div></div>
          </div>
        </div>
      </div>
    </div>

    <div class="page" id="page-template">
      <div class="template-grid">
        <div class="card">
          <div class="card-hd"><i data-lucide="list"></i>Daftar Template (Urutkan dengan ▲▼)</div>
          <div id="templateTplList"></div>
          <hr>
          <button class="btn btn-primary" style="width:100%;justify-content:center" onclick="openModal()"><i data-lucide="plus" style="width:13px;height:13px"></i>Tambah Template Baru</button>
        </div>
        <div class="card">
          <div class="card-hd"><i data-lucide="info"></i>Tentang Template</div>
          <div style="font-size:12px;color:var(--text-muted);line-height:1.8">
            <p>Setiap template dapat memiliki parameter dinamis yang disesuaikan kebutuhannya.</p>
            <br>
            <p><strong style="color:var(--text)">Tipe Input Parameter:</strong></p>
            <p>• <strong>Manual</strong> — pengguna mengetik nilai secara bebas</p>
            <p>• <strong>Dropdown</strong> — pengguna memilih dari daftar opsi</p>
            <br>
            <p>Parameter dapat ditandai <strong>Wajib</strong> atau <strong>Opsional</strong>.</p>
            <p>Template ID didapat dari dashboard Maxchat.</p>
          </div>
        </div>
      </div>
    </div>

    <div class="page" id="page-chat">
      <div class="chat-layout">
        <div class="chat-contacts">
          <div class="chat-contacts-hd">
            <span style="display:flex;align-items:center;gap:6px"><i data-lucide="messages-square" style="width:13px;height:13px"></i>Riwayat Chat</span>
            <span id="chatContactCount" style="font-size:10px;color:var(--text-faint)">0 kontak</span>
          </div>
          
          <div class="chat-tabs">
            <button class="chat-tab active" id="tabAll" onclick="filterTab('all', this)">Semua</button>
            <button class="chat-tab" id="tabMarking" onclick="filterTab('marking', this)">WO Marking</button>
            <button class="chat-tab" id="tabIndividu" onclick="filterTab('individu', this)">WO Individu</button>
          </div>
          <div class="chat-tabs" style="padding-top:4px">
            <button class="chat-tab" id="tabUnread" onclick="filterTab('unread', this)">Belum Dibaca <span id="unreadCount" style="display:none">0</span></button>
          </div>

          <div class="chat-contacts-search">
            <input type="text" id="chatSearch" placeholder="Cari pesan atau nomor..." oninput="filterContacts()">
            <button class="search-clear" id="searchClear" onclick="clearSearch()">
              <i data-lucide="x" style="width:14px;height:14px"></i>
            </button>
          </div>

          <div class="chat-contacts-list" id="chatContactsList">
            <div class="empty">Belum ada riwayat pengiriman.</div>
          </div>
        </div>
        <div class="chat-main" id="chatMain">
          <div class="chat-no-contact">
            <i data-lucide="message-circle"></i>
            <div>Pilih kontak untuk melihat riwayat pesan</div>
          </div>
        </div>
      </div>
    </div>

    <div class="page" id="page-kontak">
      <div class="card">
        <div class="card-hd" style="justify-content:space-between;display:flex;align-items:center">
          <span style="display:flex;align-items:center;gap:6px"><i data-lucide="contact"></i>Daftar Kontak</span>
          <div style="display:flex;gap:6px">
            <button class="btn btn-primary btn-sm" onclick="openAddKontakModal()"><i data-lucide="user-plus" style="width:11px;height:11px"></i>Tambah Kontak</button>
            <button class="btn btn-ghost btn-sm" onclick="openLabelManageModal()"><i data-lucide="tag" style="width:11px;height:11px"></i>Kelola Label</button>
          </div>
        </div>
        <div class="card-sub">Semua nomor yang pernah dikirimi pesan (WO Marking maupun WO Individu digabung jadi satu daftar), ditambah kontak yang kamu tambahkan manual. Buat label lewat "Kelola Label" — tiap label punya tab sendiri di bawah — lalu tempelkan ke kontak mana saja, terpisah dari kategori WO.</div>
        <div style="display:flex;gap:8px;margin-bottom:9px;flex-wrap:wrap;align-items:center">
          <input type="text" id="kontakSearch" placeholder="Cari nama atau nomor..." style="flex:1;min-width:180px" oninput="debounceLoadKontak()">
          <span id="kontakCount" style="font-size:10px;color:var(--text-faint);white-space:nowrap">0 kontak</span>
        </div>
        <div id="kontakLabelChips" class="tabs" style="display:flex;gap:5px;flex-wrap:wrap;margin-bottom:11px"></div>
        <div style="max-height:560px;overflow-y:auto">
          <table class="log-table">
            <thead><tr><th>Nama</th><th>Nomor</th><th>Label</th><th>Aksi</th></tr></thead>
            <tbody id="kontakTableBody"><tr><td colspan="4" class="empty">Memuat...</td></tr></tbody>
          </table>
        </div>
        <div style="text-align:center;padding:9px 0 0" id="kontakLoadMoreWrap"></div>
      </div>
    </div>

<div class="modal-overlay" id="labelModalOverlay">
  <div class="modal" style="width:420px">
    <div class="modal-hd">
      <span class="modal-title">Kelola Label</span>
      <button class="modal-close" onclick="closeLabelManageModal()"><i data-lucide="x" style="width:18px;height:18px"></i></button>
    </div>
    <div class="qr-add-row">
      <input type="text" id="newLabelName" placeholder="Nama label baru...">
      <input type="color" id="newLabelColor" value="#2563eb" style="width:40px;height:34px;padding:2px;border:1.5px solid var(--input-border);border-radius:7px;cursor:pointer;background:var(--input-bg)">
      <button class="btn btn-primary btn-sm" onclick="addLabel()"><i data-lucide="plus" style="width:11px;height:11px"></i></button>
    </div>
    <div class="field-hint" style="margin-bottom:10px">Buat label di sini dulu — baru bisa ditempelkan ke kontak (dari menu Kontak maupun saat chat). Tiap label akan muncul sebagai tab sendiri.</div>
    <div id="labelManageList"></div>
  </div>
</div>

<div class="qr-menu" id="labelPickerMenu" style="display:none;position:fixed;top:0;bottom:auto;width:230px"></div>


    <div class="page" id="page-settings">
      <div class="settings-sections">

        <div class="settings-section-panel" id="settings-section-api">
          <div class="card">
            <div class="card-hd"><i data-lucide="key"></i>Konfigurasi API Maxchat</div>
            <div class="card-sub">Token &amp; webhook yang menghubungkan aplikasi ini ke akun Maxchat Anda.</div>
            <div class="settings-input-group">
              <label>REST API Token (Authorization Key)</label>
              <div class="settings-token-row">
                <input type="password" id="cfgApiToken" placeholder="Masukkan API token dari Maxchat...">
                <button class="btn btn-ghost btn-sm" onclick="toggleTokenVisibility()" id="tokenToggleBtn"><i data-lucide="eye" style="width:12px;height:12px"></i></button>
              </div>
              <div class="field-hint">Token digunakan sebagai Bearer Authorization ke Maxchat API</div>
            </div>
            <div class="settings-input-group">
              <label>Webhook URL (Pesan Masuk & Status)</label>
              <input type="text" id="cfgWebhookUrl" placeholder="https://yourdomain.com/webhook/incoming">
              <div class="field-hint">URL yang akan menerima notifikasi pesan masuk/keluar/status dari Maxchat</div>
            </div>
            <button class="btn btn-primary" onclick="saveConfig()"><i data-lucide="save" style="width:12px;height:12px"></i>Simpan Konfigurasi</button>
          </div>

          <div class="card">
            <div class="card-hd"><i data-lucide="help-circle"></i>Cara Mendapatkan API Token</div>
            <div style="font-size:12px;color:var(--text-muted);line-height:1.8">
              <p>1. Login ke <strong style="color:var(--text)">app.maxchat.id</strong></p>
              <p>2. Buka menu <strong style="color:var(--text)">Settings → API</strong></p>
              <p>3. Copy nilai <strong style="color:var(--text)">API Token / Secret Key</strong></p>
              <p>4. Paste di kolom token di atas lalu simpan</p>
              <br>
              <p><strong style="color:var(--text)">Webhook URL</strong> digunakan agar Maxchat dapat mengirim notifikasi pesan masuk ke server Anda secara real-time.</p>
              <br>
              <p style="font-size:10px;color:var(--text-faint)">Database: SQLite (storage/history.db)</p>
            </div>
          </div>
        </div>

        <div class="settings-section-panel" id="settings-section-username" style="display:none">
          <div class="card">
            <div class="card-hd"><i data-lucide="user-cog"></i>Ubah Username &amp; Password Login</div>
            <div class="card-sub">Kredensial untuk masuk ke aplikasi Command Center ini.</div>
            <div class="settings-input-group">
              <label>Username</label>
              <input type="text" id="cfgLoginUsername" placeholder="Username login">
            </div>
            <div class="settings-input-group">
              <label>Password Baru</label>
              <input type="password" id="cfgLoginPassword" placeholder="Kosongkan jika tidak diubah">
            </div>
            <button class="btn btn-primary" onclick="saveLoginCredentials()"><i data-lucide="save" style="width:12px;height:12px"></i>Simpan Kredensial</button>
          </div>
        </div>

        <div class="settings-section-panel" id="settings-section-quickreply" style="display:none">
          <div class="card">
            <div class="card-hd"><i data-lucide="message-square-quote"></i>Quick Reply</div>
            <div class="card-sub">Daftar balasan cepat yang muncul saat menekan ikon Quick Reply di menu Chat.</div>
            <div class="qr-add-row">
              <input type="text" id="newQuickReplyText" placeholder="Tulis balasan cepat baru...">
              <button class="btn btn-primary btn-sm" onclick="addQuickReply()"><i data-lucide="plus" style="width:11px;height:11px"></i>Tambah</button>
            </div>
            <div class="field-hint" style="margin-bottom:10px">Urutkan dengan tombol ▲▼.</div>
            <div class="qr-list-wrap" id="quickReplyList"></div>
          </div>
        </div>

        <div class="settings-section-panel" id="settings-section-history" style="display:none">
          <div class="card">
            <div class="card-hd"><i data-lucide="calendar-x"></i>Hapus Riwayat Berdasarkan Tanggal</div>
            <div class="card-sub">Menghapus riwayat pesan (masuk &amp; keluar) pada rentang tanggal tertentu saja.</div>
            <div class="settings-input-group">
              <label>Dari Tanggal</label>
              <input type="date" id="cfgRangeStart">
            </div>
            <div class="settings-input-group">
              <label>Sampai Tanggal</label>
              <input type="date" id="cfgRangeEnd">
            </div>
            <button class="btn btn-danger btn-sm" onclick="deleteHistoryByRange()">
              <i data-lucide="calendar-x" style="width:11px;height:11px"></i>Hapus Riwayat di Rentang Ini
            </button>
          </div>

          <div class="card card-danger">
            <div class="card-hd"><i data-lucide="alert-triangle"></i>Zona Berbahaya</div>
            <div class="card-sub">Tindakan berikut menghapus data secara permanen dan tidak bisa dibatalkan.</div>
            <div class="danger-row">
              <div class="danger-row-info">
                <div class="danger-row-title">Hapus Semua Riwayat</div>
                <div class="danger-row-desc">Menghapus seluruh riwayat pesan blast &amp; chat masuk/keluar dari database. Backup otomatis harian tetap tersimpan terpisah dan tidak ikut terhapus.</div>
              </div>
              <button class="btn btn-danger btn-sm" onclick="clearAllHistory()">
                <i data-lucide="trash-2" style="width:11px;height:11px"></i>Hapus Semua Riwayat
              </button>
            </div>
          </div>
        </div>

      </div>
    </div>

    <div class="page" id="page-dashboard">
      <div class="settings-sections">
        <div class="settings-grid">
          <div class="card">
            <div class="card-hd"><i data-lucide="database"></i>Sistem &amp; Database</div>
            <div class="card-sub">Ukuran database saat ini dan status backup otomatis harian.</div>
            <div class="db-stats-row">
              <div class="db-stat-box"><div class="db-stat-box-inner"><div class="val" id="dbSizeVal">–</div><div class="lbl">Ukuran DB</div></div></div>
              <div class="db-stat-box"><div class="db-stat-box-inner"><div class="val" id="dbContactsVal">–</div><div class="lbl">Total Broadcast</div></div></div>
              <div class="db-stat-box"><div class="db-stat-box-inner"><div class="val" id="dbSessionsVal">–</div><div class="lbl">Sesi Blast</div></div></div>
            </div>
            <div class="backup-info-box">
              <div class="backup-info-row">
                <span class="k">Backup Otomatis</span>
                <span class="v"><span class="status-pill on"><span class="dot"></span>Aktif — 04:00 WIB</span></span>
              </div>
              <div class="backup-info-row">
                <span class="k">Backup Terakhir</span>
                <span class="v" id="lastBackupVal">Belum pernah</span>
              </div>
              <div class="backup-info-row">
                <span class="k">File Backup Tersimpan</span>
                <span class="v" id="backupCountVal">0 file</span>
              </div>
            </div>
            <div style="display:flex;gap:8px">
              <button class="btn btn-primary btn-sm" onclick="runManualBackup()" id="manualBackupBtn"><i data-lucide="hard-drive-download" style="width:11px;height:11px"></i>Backup Sekarang</button>
              <button class="btn btn-ghost btn-sm" onclick="refreshSystemInfo()"><i data-lucide="refresh-cw" style="width:11px;height:11px"></i>Refresh</button>
            </div>
          </div>

          <div class="card">
            <div class="card-hd"><i data-lucide="history"></i>Riwayat Nomor Diblas per Tanggal</div>
            <div class="card-sub">Jumlah pesan broadcast, nomor unik, status berhasil/gagal, &amp; estimasi biaya (Rp 365/broadcast, dihitung dari yang <strong>berhasil</strong> saja) — balasan manual tidak dihitung.</div>
            <div style="display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap;margin-bottom:10px">
              <div class="settings-input-group" style="margin-bottom:0;flex:1;min-width:160px">
                <label>Rentang</label>
                <select id="bcHistoryRange" onchange="loadBroadcastHistory()">
                  <option value="7">7 hari terakhir</option>
                  <option value="14" selected>14 hari terakhir</option>
                  <option value="30">30 hari terakhir</option>
                  <option value="90">90 hari terakhir</option>
                </select>
              </div>
              <div class="input-type-toggle" style="margin-bottom:2px">
                <button type="button" class="toggle-btn active" id="bcGroupDayBtn" onclick="setBcGroupBy('day')">Harian</button>
                <button type="button" class="toggle-btn" id="bcGroupMonthBtn" onclick="setBcGroupBy('month')">Bulanan</button>
              </div>
            </div>
            <div style="max-height:280px;overflow-y:auto" id="bcHistoryWrap">
              <table class="log-table">
                <thead><tr><th>Tanggal</th><th>Total</th><th>Berhasil</th><th>Gagal</th><th>Nomor Unik</th><th>Biaya</th></tr></thead>
                <tbody id="bcHistoryBody"><tr><td colspan="6" class="empty">Memuat...</td></tr></tbody>
              </table>
            </div>
            <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:10px;flex-wrap:wrap">
              <div class="field-hint" id="bcHistoryTotalCost" style="margin:0"></div>
              <div style="display:flex;gap:6px;flex-shrink:0">
                <button class="btn btn-ghost btn-sm" onclick="downloadBroadcastHistoryExport('xlsx')" title="Download rekap sebagai Excel">
                  <i data-lucide="file-spreadsheet" style="width:11px;height:11px"></i>Excel
                </button>
                <button class="btn btn-ghost btn-sm" onclick="downloadBroadcastHistoryExport('csv')" title="Download rekap sebagai CSV">
                  <i data-lucide="file-text" style="width:11px;height:11px"></i>CSV
                </button>
              </div>
            </div>
          </div>

          <div class="card">
            <div class="card-hd"><i data-lucide="file-down"></i>Export Data Broadcast</div>
            <div class="card-sub">Unduh detail per nomor (tanggal, nomor telepon, isi pesan/template, status pengiriman) untuk rentang tanggal tertentu.</div>

            <div class="fg" style="margin-bottom:10px">
              <label class="field-label">Pilih Cepat</label>
              <div style="display:flex;gap:5px;flex-wrap:wrap">
                <button type="button" class="btn btn-ghost btn-sm export-quick-btn" onclick="setExportQuickRange('today', this)">Hari Ini</button>
                <button type="button" class="btn btn-ghost btn-sm export-quick-btn" onclick="setExportQuickRange('7', this)">7 Hari</button>
                <button type="button" class="btn btn-ghost btn-sm export-quick-btn" onclick="setExportQuickRange('30', this)">30 Hari</button>
                <button type="button" class="btn btn-ghost btn-sm export-quick-btn" onclick="setExportQuickRange('month', this)">Bulan Ini</button>
              </div>
            </div>

            <div class="field-row" style="margin-bottom:12px">
              <div class="fg">
                <label class="field-label">Dari Tanggal</label>
                <input type="date" id="exportRangeStart" onchange="clearExportQuickActive()">
              </div>
              <div class="fg">
                <label class="field-label">Sampai Tanggal</label>
                <input type="date" id="exportRangeEnd" onchange="clearExportQuickActive()">
              </div>
            </div>

            <div class="fg" style="margin-bottom:14px">
              <label class="field-label">Format File</label>
              <div class="input-type-toggle" id="exportFormatToggle">
                <button type="button" class="toggle-btn active" data-fmt="xlsx" onclick="setExportFormat('xlsx', this)">
                  <i data-lucide="file-spreadsheet" style="width:11px;height:11px;vertical-align:-2px;margin-right:3px"></i>Excel (.xlsx)
                </button>
                <button type="button" class="toggle-btn" data-fmt="csv" onclick="setExportFormat('csv', this)">
                  <i data-lucide="file-text" style="width:11px;height:11px;vertical-align:-2px;margin-right:3px"></i>CSV (.csv)
                </button>
              </div>
            </div>

            <button class="btn-blast" style="margin-top:0" onclick="downloadBroadcastExport()">
              <i data-lucide="download" style="width:14px;height:14px"></i>
              <span>Download Data Broadcast</span>
            </button>
            <div class="field-hint" style="text-align:center;margin-top:8px">Kolom: Tanggal · Nomor Telepon · Isi Pesan/Template · Status Pengiriman</div>
          </div>

          <div class="card">
            <div class="card-hd"><i data-lucide="alert-circle"></i>Nomor yang Sering Dibroadcast</div>
            <div class="card-sub">Nomor yang paling sering menerima pesan broadcast/blast (bukan balasan manual), dihitung per pengiriman dan per hari berbeda dalam rentang waktu tertentu.</div>
            <div style="display:flex;gap:10px;flex-wrap:wrap">
              <div class="settings-input-group" style="margin-bottom:10px;flex:1;min-width:140px">
                <label>Rentang Cek</label>
                <select id="frDaysRange" onchange="loadFrequentReporters()">
                  <option value="30">30 hari terakhir</option>
                  <option value="60">60 hari terakhir</option>
                  <option value="90" selected>90 hari terakhir</option>
                </select>
              </div>
              <div class="settings-input-group" style="margin-bottom:10px;flex:1;min-width:140px">
                <label>Minimal Hari Dibroadcast</label>
                <select id="frMinDays" onchange="loadFrequentReporters()">
                  <option value="2">≥ 2 hari</option>
                  <option value="3" selected>≥ 3 hari</option>
                  <option value="5">≥ 5 hari</option>
                  <option value="7">≥ 7 hari</option>
                </select>
              </div>
            </div>
            <div style="max-height:280px;overflow-y:auto" id="frWrap">
              <table class="log-table">
                <thead><tr><th>Nomor</th><th>Nama</th><th>Jumlah Hari</th><th>Kali Dibroadcast</th><th>Broadcast Terakhir</th></tr></thead>
                <tbody id="frBody"><tr><td colspan="5" class="empty">Memuat...</td></tr></tbody>
              </table>
            </div>
          </div>

          <div class="card" style="grid-column:1 / -1">
            <div class="card-hd"><i data-lucide="line-chart"></i>Trend Harian WO Marking vs WO Individu</div>
            <div class="card-sub">Perbandingan jumlah NOMOR UNIK yang di-broadcast untuk WO Marking dan WO Individu setiap hari.</div>
            <div class="settings-input-group" style="margin-bottom:10px;max-width:240px">
              <label>Rentang</label>
              <select id="woTrendRange" onchange="loadWoTrendChart()">
                <option value="7">7 hari terakhir</option>
                <option value="14">14 hari terakhir</option>
                <option value="30" selected>30 hari terakhir</option>
                <option value="60">60 hari terakhir</option>
                <option value="90">90 hari terakhir</option>
              </select>
            </div>
            <div id="woTrendChartWrap" style="width:100%;overflow-x:auto"></div>
          </div>
        </div>
      </div>
    </div>

  </div>
</div>
</div>

<div class="modal-overlay" id="modalOverlay">
  <div class="modal">
    <div class="modal-hd">
      <span class="modal-title" id="modalTitle">Tambah Template Baru</span>
      <button class="modal-close" onclick="closeModal()"><i data-lucide="x" style="width:18px;height:18px"></i></button>
    </div>
    <input type="hidden" id="editingId">
    <div class="field-row" style="margin-bottom:10px">
      <div class="fg">
        <label class="field-label">Nama Template</label>
        <input type="text" id="mTplName" placeholder="contoh: info_gangguan">
      </div>
      <div class="fg">
        <label class="field-label">Template ID (dari Maxchat)</label>
        <input type="text" id="mTplId" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">
      </div>
    </div>
    <div class="fg" style="margin-bottom:10px">
      <label class="field-label">Kategori WO</label>
      <select id="mWoCategory">
        <option value="umum">Umum (tampil di semua PC)</option>
        <option value="individu">WO Individu</option>
        <option value="marking">WO Marking</option>
      </select>
      <div class="field-hint">Menentukan PC mana yang akan menerima balasan customer dari nomor yang mendapat template ini.</div>
    </div>
    <div class="fg" style="margin-bottom:10px">
      <label class="field-label">Preview Isi Pesan</label>
      <textarea id="mTplPreview" style="height:85px" placeholder="Tulis isi pesan. Gunakan {{1}} {{2}} dst untuk parameter."></textarea>
    </div>
    <div style="display:flex;align-items:center;gap:7px;margin-bottom:12px">
      <input type="checkbox" id="mHasParams" style="width:14px;height:14px;cursor:pointer" onchange="toggleParamBuilder()">
      <label for="mHasParams" style="font-size:12px;cursor:pointer;font-weight:600;color:var(--text)">Template ini memiliki parameter</label>
    </div>
    <div id="paramBuilderSection" style="display:none">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:7px">
        <span style="font-size:11px;font-weight:700;color:#3b82f6;display:flex;align-items:center;gap:5px"><i data-lucide="settings-2" style="width:12px;height:12px"></i>Daftar Parameter</span>
        <button class="btn btn-ghost btn-sm" onclick="addParamBuilder()"><i data-lucide="plus" style="width:11px;height:11px"></i>Tambah Parameter</button>
      </div>
      <div id="paramBuilderList"></div>
    </div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeModal()">Batal</button>
      <button class="btn btn-primary" id="modalSaveBtn" onclick="saveTemplate()"><i data-lucide="save" style="width:12px;height:12px"></i>Simpan Template</button>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<div class="modal-overlay" id="addKontakModalOverlay">
  <div class="modal" style="width:380px">
    <div class="modal-hd">
      <span class="modal-title">Tambah Kontak Baru</span>
      <button class="modal-close" onclick="closeAddKontakModal()"><i data-lucide="x" style="width:18px;height:18px"></i></button>
    </div>
    <div class="fg" style="margin-bottom:10px">
      <label class="field-label">Nomor WhatsApp</label>
      <input type="text" id="newKontakPhone" placeholder="08123456789">
    </div>
    <div class="fg" style="margin-bottom:10px">
      <label class="field-label">Nama (opsional)</label>
      <input type="text" id="newKontakName" placeholder="Nama kontak...">
    </div>
    <div class="field-hint" style="margin-bottom:12px">Untuk kontak yang belum pernah dikirimi pesan — bisa disiapkan nama & label-nya lebih dulu sebelum broadcast pertama.</div>
    <div class="modal-footer">
      <button class="btn btn-ghost" onclick="closeAddKontakModal()">Batal</button>
      <button class="btn btn-primary" onclick="saveNewKontak()"><i data-lucide="save" style="width:12px;height:12px"></i>Simpan Kontak</button>
    </div>
  </div>
</div>

<div id="lightboxOverlay" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,.88);z-index:9999;align-items:center;justify-content:center;cursor:zoom-out" onclick="closeLightbox()">
  <img id="lightboxImg" src="" alt="" style="max-width:92vw;max-height:92vh;border-radius:8px;object-fit:contain;box-shadow:0 8px 40px rgba(0,0,0,.6)">
  <video id="lightboxVideo" src="" controls autoplay style="display:none;max-width:92vw;max-height:92vh;border-radius:8px;box-shadow:0 8px 40px rgba(0,0,0,.6)" onclick="event.stopPropagation()"></video>
  <button onclick="closeLightbox()" style="position:fixed;top:16px;right:20px;background:rgba(255,255,255,.15);border:none;color:#fff;font-size:22px;width:36px;height:36px;border-radius:50%;cursor:pointer;display:flex;align-items:center;justify-content:center">✕</button>
</div>

<script>
let appData = { templates: [] };
window.appData = appData;
let selectedTpl = null;
let editingDirty = false;
let selectedContact = null;
let allContacts = [];
let lastHistoryCount = 0;
let lastLatestTime = '';
let activeChatTab = 'all';
const CONTACTS_PAGE_SIZE = 50; 
let contactsHasMore = false;
let contactsOldestTime = null;
let loadingOlderContacts = false;
let globalUnreadCount = 0; 
let tabContacts = [];      
let tabHasMore = false;
let tabOldestTime = null;
let loadingTabContacts = false;
let searchContacts = [];
let searchHasMore = false;
let searchOldestTime = null;
let loadingSearchMore = false;
let currentSearchQuery = '';
let currentRenderedContacts = []; 
let quickReplies = []; 

function refreshIcons(){ if (window.lucide) lucide.createIcons(); }

async function init() {
  const savedTheme = localStorage.getItem('mcTheme') || 'light';
  document.body.className = savedTheme;
  updateThemeUI(savedTheme);

  // Ambil semua data awal SEKALIGUS (paralel) — bukan satu-satu berurutan,
  // supaya loading awal tidak kena "pajak" latency tunnel berkali-kali.
  const [dataRes, cfgRes, contactsRes, latestTimeRes, quickRepliesRes] = await Promise.all([
    fetch('/api/data').then(r => r.json()),
    fetch('/api/config').then(r => r.json()),
    fetch('/api/contacts?limit=' + CONTACTS_PAGE_SIZE).then(r => r.json()),
    fetch('/api/history/latest-time').then(r => r.json()).catch(() => null),
    fetch('/api/quick-replies').then(r => r.json()).catch(() => []),
  ]);

  appData = dataRes;
  window.appData = appData;
  if (!appData.templates) appData.templates = [];
  appData.templates.sort((a,b) => (a.order||999) - (b.order||999));
  renderDropdown();
  renderTemplatePage();

  document.getElementById('cfgApiToken').value    = cfgRes.apiToken    || '';
  document.getElementById('cfgWebhookUrl').value  = cfgRes.webhookUrl  || '';
  document.getElementById('cfgLoginUsername').value = cfgRes.loginUsername || '';

  quickReplies = quickRepliesRes || [];
  renderQuickReplyList();

  allContacts = contactsRes.contacts || [];
  contactsHasMore = contactsRes.hasMore;
  contactsOldestTime = contactsRes.oldestTime;
  const countElInit = document.getElementById('chatContactCount');
  if (countElInit) countElInit.textContent = allContacts.length + ' kontak';
  lastHistoryCount = allContacts.length;

  if (latestTimeRes) {
    lastLatestTime = latestTimeRes.latestTime || '';
    globalUnreadCount = latestTimeRes.unreadCount || 0;
    updateUnreadCount();
  }

  // Kunci supaya polling tidak tumpang tindih saat jaringan/tunnel lambat —
  // kalau polling sebelumnya belum selesai, siklus berikutnya dilewati dulu.
  let pollInFlight = false;
  setInterval(async () => {
    if (pollInFlight) return;
    pollInFlight = true;
    try {
      const { latestTime, unreadCount } = await fetch('/api/history/latest-time').then(r => r.json());
      globalUnreadCount = unreadCount || 0;
      updateUnreadCount();

      if (latestTime !== lastLatestTime) {
        const isFirst = lastLatestTime === '';
        lastLatestTime = latestTime;
        const chatPage = document.getElementById('page-chat');

        if (chatPage?.classList.contains('active')) {
          // Hanya fetch daftar "Semua" kalau tab itu yang sedang aktif.
          // Untuk tab lain (Marking/Individu/Unread), filterContacts() di
          // bawah sudah otomatis fetch datanya sendiri lewat loadTabContacts —
          // kalau di sini kita fetch juga, jadi DUA request per siklus polling
          // untuk data yang sama. Ini sumber utama "lambat" yang terlihat di
          // Network tab (banyak request beruntun tiap kali ada pesan baru).
          if (activeChatTab === 'all') {
            await loadChatHistory(true);
          }
          filterContacts();
          if (selectedContact) {
            refreshChatMessages(selectedContact.phone);
          }
        } else {
          // Chat page tidak aktif: cukup segarkan allContacts di background
          // (murah, cuma 1 request) supaya saat menu Chat dibuka nanti datanya
          // tidak basi — tapi TIDAK perlu fetch tab kategori yang sedang tidak
          // dilihat siapa pun.
          await loadChatHistory(true);
          if (!isFirst) showIncomingBadge();
        }
      }
    } catch(e) {
    } finally {
      pollInFlight = false;
    }
  }, 5000);

  refreshIcons();
}

function showIncomingBadge() {
  const navChat = document.querySelector('.nav-item[onclick*="chat"]');
  if (!navChat) return;
  const chatPage = document.getElementById('page-chat');
  if (chatPage?.classList.contains('active')) return;
  if (!navChat.querySelector('.nav-badge')) {
    const badge = document.createElement('span');
    badge.className = 'nav-badge';
    navChat.appendChild(badge);
  }
}

async function doLogout() {
  if (!confirm('Yakin ingin logout?')) return;
  await fetch('/logout', { method: 'POST' });
  window.location.href = '/login';
}

function toggleTheme() {
  const isDark = document.body.classList.contains('dark');
  const newTheme = isDark ? 'light' : 'dark';
  document.body.className = newTheme;
  localStorage.setItem('mcTheme', newTheme);
  updateThemeUI(newTheme);
  refreshIcons();
}
function updateThemeUI(theme) {
  const ic = document.getElementById('themeIcon');
  if (ic) ic.setAttribute('data-lucide', theme === 'dark' ? 'sun' : 'moon');
  document.getElementById('themeLabel').textContent = theme === 'dark' ? 'Mode Terang' : 'Mode Gelap';
  refreshIcons();
}

function openMedia(el, kind) {
  const src  = el.src || el.currentSrc || '';
  const type = kind || (el.tagName === 'VIDEO' ? 'video' : 'image');
  const lb    = document.getElementById('lightboxOverlay');
  const img   = document.getElementById('lightboxImg');
  const video = document.getElementById('lightboxVideo');
  if (!lb || !img || !video) { window.open(src, '_blank'); return; }
  if (type === 'video') {
    img.style.display = 'none'; img.src = '';
    video.style.display = 'block';
    video.src = src;
    video.currentTime = 0;
    video.play().catch(() => {}); // browser bisa menolak autoplay bersuara, tidak masalah — tombol play tetap ada
  } else {
    video.pause(); video.removeAttribute('src'); video.load(); video.style.display = 'none';
    img.style.display = 'block';
    img.src = src;
  }
  lb.style.display = 'flex';
  document.body.style.overflow = 'hidden';
}
function openVideoLightbox(btn) {
  const wrap = btn.closest('.video-thumb-wrap');
  const vid  = wrap?.querySelector('video');
  if (vid) openMedia(vid, 'video');
}
function closeLightbox() {
  const lb    = document.getElementById('lightboxOverlay');
  const video = document.getElementById('lightboxVideo');
  if (video) { video.pause(); video.removeAttribute('src'); video.load(); } // hentikan pemutaran & lepas koneksi network saat ditutup
  if (lb) lb.style.display = 'none';
  document.body.style.overflow = '';
}
function mediaFallback(el) {
  const src = el.src || el.currentSrc || '';
  const fname = src.split('/').pop().split('?')[0] || 'file';
  const link = document.createElement('a');
  link.href = src; link.target = '_blank';
  link.style.cssText = 'color:#58a6ff;font-size:11px;text-decoration:underline';
  link.textContent = '🔗 Buka file: ' + fname;
  el.parentNode.replaceChild(link, el);
}

function toast(msg, type='') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(() => t.className = 'toast', 2600);
}

function copyPhoneToClipboard(phone) {
  const cleaned = String(phone).replace(/\\s+/g, '');
  const doCopy = () => toast('📋 Nomor disalin: ' + cleaned, 'ok');
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(cleaned).then(doCopy).catch(() => {
      const ta = document.createElement('textarea');
      ta.value = cleaned; document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy'); doCopy(); } catch(e) { toast('Gagal menyalin nomor', 'err'); }
      document.body.removeChild(ta);
    });
  } else {
    const ta = document.createElement('textarea');
    ta.value = cleaned; document.body.appendChild(ta); ta.select();
    try { document.execCommand('copy'); doCopy(); } catch(e) { toast('Gagal menyalin nomor', 'err'); }
    document.body.removeChild(ta);
  }
}

function showPage(name, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item, .nav-toggle, .nav-subitem').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + name).classList.add('active');
  el.classList.add('active');
  const map = {
    broadcast: ['Broadcast','Pilih template lalu kirim ke banyak nomor'],
    template:  ['Template','Kelola template pesan WhatsApp'],
    chat:      ['Chat','Riwayat pesan masuk & keluar per kontak'],
    kontak:    ['Kontak','Semua kontak (WO Marking & Individu) beserta label'],
    dashboard: ['Dashboard','Statistik sistem, database, dan trend WO']
  };
  document.getElementById('topbarTitle').textContent = map[name][0];
  document.getElementById('topbarSub').textContent   = map[name][1];
  if (name === 'template') renderTemplatePage();
  if (name === 'dashboard') { refreshSystemInfo(); loadBroadcastHistory(); loadFrequentReporters(); loadWoTrendChart(); }
  if (name === 'kontak') loadKontakPage();
  if (name === 'chat') {
    document.querySelectorAll('.nav-badge').forEach(b => b.remove());
    if (selectedContact) markContactRead(selectedContact.phone);
    filterContacts();
  }
  refreshIcons();
}

// ------------------------------------------------------------
// Menu tree "Pengaturan" — toggle buka/tutup submenu + pindah antar
// section (API Maxchat / Username / Balas Cepat / Hapus Riwayat) tanpa
// meninggalkan halaman-halaman lain (satu page-settings, section ditoggle
// lewat display:none/block, bukan lewat showPage biasa).
// ------------------------------------------------------------
const SETTINGS_SECTION_TITLES = {
  api:        ['Pengaturan · API Maxchat', 'Konfigurasi API Token dan Webhook'],
  username:   ['Pengaturan · Username', 'Ubah username & password login'],
  quickreply: ['Pengaturan · Balas Cepat', 'Kelola daftar balasan cepat di menu Chat'],
  history:    ['Pengaturan · Hapus Riwayat', 'Hapus riwayat pesan berdasarkan tanggal atau semuanya'],
};

function toggleSettingsMenu() {
  const sub  = document.getElementById('settingsSubmenu');
  const chev = document.getElementById('settingsChevron');
  const wasOpen  = sub.classList.contains('open');
  const willOpen = !wasOpen;
  sub.classList.toggle('open', willOpen);
  chev.classList.toggle('open', willOpen);
  // Kalau baru dibuka dan halaman Pengaturan belum aktif, langsung
  // tampilkan section pertama (API Maxchat) supaya klik pertama langsung berguna.
  if (willOpen && !document.getElementById('page-settings').classList.contains('active')) {
    openSettingsSection('api', sub.querySelector('.nav-subitem'));
  }
}

function openSettingsSection(section, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item, .nav-toggle, .nav-subitem').forEach(n => n.classList.remove('active'));
  document.getElementById('page-settings').classList.add('active');
  document.getElementById('navPengaturanToggle').classList.add('active');
  document.getElementById('settingsSubmenu').classList.add('open');
  document.getElementById('settingsChevron').classList.add('open');
  if (el) el.classList.add('active');

  document.querySelectorAll('.settings-section-panel').forEach(p => p.style.display = 'none');
  const panel = document.getElementById('settings-section-' + section);
  if (panel) panel.style.display = '';

  const titles = SETTINGS_SECTION_TITLES[section] || ['Pengaturan', ''];
  document.getElementById('topbarTitle').textContent = titles[0];
  document.getElementById('topbarSub').textContent   = titles[1];

  if (section === 'quickreply') renderQuickReplyList();
  refreshIcons();
}

function renderDropdown() {
  const menu = document.getElementById('tplMenu');
  appData.templates.sort((a,b) => (a.order||999) - (b.order||999));
  if (!appData.templates.length) { menu.innerHTML='<div class="empty" style="padding:10px">Belum ada template.</div>'; refreshIcons(); return; }
  menu.innerHTML = appData.templates.map(t => \`
    <div class="dd-item \${selectedTpl?.id===t.id?'active':''}" onclick="selectTpl('\${t.id}')">
      <div class="dd-item-name">\${t.name}</div>
      <div class="dd-item-id">\${t.templateId}</div>
      <div class="dd-item-badges">
        <span class="badge badge-blue">WhatsApp</span>
        \${t.hasParams && t.params?.length ? '<span class="badge badge-yellow">'+t.params.length+' param</span>' : '<span class="badge badge-gray">Tanpa Parameter</span>'}
      </div>
    </div>
  \`).join('');
  refreshIcons();
}

function toggleDD() {
  document.getElementById('tplBtn').classList.toggle('open');
  document.getElementById('tplMenu').classList.toggle('open');
}
document.addEventListener('click', e => {
  if (!document.getElementById('tplWrap').contains(e.target)) {
    document.getElementById('tplBtn').classList.remove('open');
    document.getElementById('tplMenu').classList.remove('open');
  }
});

// Tutup dropdown Quick Reply kalau klik di luar tombol/menu-nya.
document.addEventListener('click', e => {
  const menu = document.getElementById('quickReplyMenu');
  const btn  = document.getElementById('quickReplyBtn');
  if (!menu || menu.style.display !== 'block') return;
  if (menu.contains(e.target) || (btn && btn.contains(e.target))) return;
  menu.style.display = 'none';
});

function selectTpl(id) {
  selectedTpl = appData.templates.find(t => t.id === id);
  const lbl = document.getElementById('tplBtnLabel');
  lbl.className = 'dd-selected'; lbl.textContent = selectedTpl.name;
  document.getElementById('tplBtn').classList.remove('open');
  document.getElementById('tplMenu').classList.remove('open');
  renderDropdown();
  if (selectedTpl.hasParams && selectedTpl.params?.length) {
    document.getElementById('paramCard').style.display = 'block';
    renderParamFields();
  } else {
    document.getElementById('paramCard').style.display = 'none';
  }
  renderPreview();
  refreshIcons();
}

function renderParamFields() {
  const wrap = document.getElementById('paramFields');
  if (!selectedTpl?.params?.length) { wrap.innerHTML=''; return; }
  wrap.innerHTML = selectedTpl.params.map(p => \`
    <div class="param-item">
      <label class="field-label">\${p.label} (Parameter \${p.index})
        \${p.required ? '<span class="badge badge-red" style="margin-left:3px">Wajib</span>' : '<span class="badge badge-gray" style="margin-left:3px">Opsional</span>'}
      </label>
      \${p.inputType === 'dropdown'
        ? \`<select id="blastParam_\${p.index}" onchange="renderPreview()">
            <option value="\${p.default||''}">\${p.default ? '— '+p.default+' (default) —' : '— Pilih —'}</option>
            \${(p.options||[]).map(o=>\`<option value="\${o}">\${o}</option>\`).join('')}
           </select>\`
        : p.inputType === 'nomer_lapor'
          ? \`<div style="background:var(--counter-bg);border:1px dashed var(--input-border);border-radius:7px;padding:6px 10px;font-size:11px;color:var(--text-muted);display:flex;align-items:center;gap:6px">
               <i data-lucide="smartphone" style="width:13px;height:13px"></i>Otomatis dari nomer lapor di kolom Nomor Tujuan
             </div>
             <input type="hidden" id="blastParam_\${p.index}" value="__nomer_lapor__">\`
          : \`<input type="text" id="blastParam_\${p.index}" placeholder="Isi \${p.label}..." oninput="renderPreview()">\`
      }
    </div>
  \`).join('');
  refreshIcons();
}

function getParamValues() {
  if (!selectedTpl?.params?.length) return [];
  return selectedTpl.params.map(p => {
    let val = '';
    const el = document.getElementById('blastParam_' + p.index);
    if (el) val = el.value.trim();
    if (!val && p.default) val = p.default;
    return { index: p.index, type: 'text', text: val };
  });
}

function renderPreview() {
  if (!selectedTpl) return;
  document.getElementById('previewEmpty').style.display = 'none';
  document.getElementById('waPhoneWrap').style.display  = 'flex';
  let txt = (selectedTpl.preview || '');
  if (selectedTpl.hasParams && selectedTpl.params?.length) {
    selectedTpl.params.forEach(p => {
      const el  = document.getElementById('blastParam_' + p.index);
      let val = (el ? el.value.trim() : '');
      if (!val && p.default) val = p.default;
      if (val === '__nomer_lapor__' || (!val && p.inputType === 'nomer_lapor')) val = '[Nomer Lapor]';
      else if (!val) val = '{{'+p.index+'}}';
      txt = txt.split('{{'+p.index+'}}').join(val);
    });
  }
  document.getElementById('waBubbleText').textContent = txt;
  const now = new Date();
  document.getElementById('waDateChip').textContent = 'Hari ini';
  document.getElementById('waTimeChip').textContent = now.getHours().toString().padStart(2,'0')+':'+now.getMinutes().toString().padStart(2,'0');
  refreshIcons();
}

function parseNumbersWithLapor(raw) {
  const lines = raw.split(/\\n|,|;/).map(l => l.trim()).filter(Boolean);
  return lines.map(line => {
    const parts = line.trim().split(/\\s+/);
    if (parts.length >= 2) return { phone: parts[parts.length - 1], nomerLapor: parts.slice(0, parts.length - 1).join(' ') };
    return { phone: parts[0], nomerLapor: '' };
  }).filter(x => x.phone);
}
function parseNumbers(raw) { return parseNumbersWithLapor(raw).map(x => x.phone); }

function updateCounter() {
  const nums = parseNumbers(document.getElementById('numbers').value);
  document.getElementById('counter').textContent = nums.length + ' nomor';
}

function genBlastId() {
  return 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

// Satu fungsi terpusat buat update progress bar, dipakai baik oleh jalur
// kirim-1-per-1 (nomer lapor) maupun jalur polling live ke server (blast massal).
function renderProgress(sent, total, phase) {
  const fill = document.getElementById('progressFill');
  const lbl  = document.getElementById('progressLbl');
  if (!fill || !lbl) return;
  const safeTotal = total || 1;
  const pct  = Math.max(0, Math.min(100, Math.round((sent / safeTotal) * 100)));
  const done = sent >= total && total > 0;
  fill.style.width = pct + '%';
  fill.classList.toggle('is-done', done);
  const phaseTxt = phase || (done ? 'Selesai' : 'Mengirim pesan');
  lbl.innerHTML = \`<span class="progress-dot \${done ? 'is-done' : ''}"></span>\${phaseTxt} · <b>\${sent}</b>/\${total} · <span class="progress-pct">\${pct}%</span>\`;
}

async function startBlast() {
  if (!selectedTpl) { toast('Pilih template terlebih dahulu.','err'); return; }
  const rawText = document.getElementById('numbers').value;
  const raw = parseNumbers(rawText);
  const rawWithLapor = parseNumbersWithLapor(rawText);
  if (!raw.length) { toast('Masukkan minimal 1 nomor.','err'); return; }
  const hasNomerLapor = selectedTpl.hasParams && selectedTpl.params?.some(p => p.inputType === 'nomer_lapor');
  if (selectedTpl.hasParams && selectedTpl.params?.length) {
    for (const p of selectedTpl.params) {
      if (p.required && p.inputType !== 'nomer_lapor') {
        const el  = document.getElementById('blastParam_' + p.index);
        const val = (el ? el.value.trim() : '') || p.default || '';
        if (!val) { toast(\`"\${p.label}" wajib diisi.\`,'err'); return; }
      }
    }
  }
  const params = getParamValues();
  const previewText = document.getElementById('waBubbleText').textContent || selectedTpl.preview;
  const btn = document.getElementById('btnBlast');
  btn.disabled = true;
  document.getElementById('spinner').style.display = 'block';
  document.getElementById('btnText').textContent = 'Mengirim...';
  document.getElementById('progressWrap').style.display = 'block';
  document.getElementById('resultCard').style.display = 'none';
  document.getElementById('emptyResult').style.display = 'flex';
  renderProgress(0, hasNomerLapor ? rawWithLapor.length : raw.length, 'Menyiapkan...');
  let pollTimer = null;
  try {
    if (hasNomerLapor) {
      const results = [];
      for (let i = 0; i < rawWithLapor.length; i++) {
        const { phone, nomerLapor } = rawWithLapor[i];
        const perParams = params.map(p => {
          const tplParam = selectedTpl.params.find(tp => tp.index === p.index);
          if (tplParam?.inputType === 'nomer_lapor') return { ...p, text: nomerLapor || p.default || '' };
          return p;
        });
        const res = await fetch('/api/blast', {
          method:'POST', headers:{'Content-Type':'application/json'},
          body: JSON.stringify({ numbers:[phone], templateId:selectedTpl.templateId, params:perParams, templateName: selectedTpl.name, previewText, nomerLapor })
        });
        const data = await res.json();
        if (data.results) results.push(...data.results);
        renderProgress(i + 1, rawWithLapor.length);
        if (i < rawWithLapor.length - 1) await new Promise(r => setTimeout(r, 200));
      }
      finishBlast(results);
    } else {
      // Kirim massal: server memproses satu request panjang (dengan jeda antar
      // nomor), jadi progresnya dipoll live dari /api/blast/progress/:id supaya
      // angka x/y di UI benar-benar mengikuti apa yang terjadi di server saat itu.
      const blastId = genBlastId();
      pollTimer = setInterval(async () => {
        try {
          const pr = await fetch('/api/blast/progress/' + blastId);
          const p  = await pr.json();
          if (p.found) renderProgress(p.sent, p.total || raw.length, p.done ? 'Menyelesaikan...' : 'Mengirim pesan');
        } catch (_) { /* biarkan, polling berikutnya coba lagi */ }
      }, 350);

      const res  = await fetch('/api/blast', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body: JSON.stringify({ numbers:raw, templateId:selectedTpl.templateId, params, templateName: selectedTpl.name, previewText, blastId })
      });
      clearInterval(pollTimer); pollTimer = null;
      const data = await res.json();
      if (data.error) { toast(data.error,'err'); return; }
      renderProgress(raw.length, raw.length);
      finishBlast(data.results, data.summary);
    }
    await loadChatHistory(true);
  } catch(err) {
    toast('Error: '+err.message,'err');
  } finally {
    if (pollTimer) clearInterval(pollTimer);
    btn.disabled = false;
    document.getElementById('spinner').style.display = 'none';
    document.getElementById('btnText').textContent = 'Kirim Pesan';
    refreshIcons();
  }
}

function finishBlast(results, summary) {
  const total   = summary ? summary.total   : results.length;
  const success = summary ? summary.success : results.filter(r => r.status === 'success').length;
  const failed  = total - success;
  renderProgress(total, total, 'Selesai');
  document.getElementById('sumTotal').textContent = total;
  document.getElementById('sumOk').textContent    = success;
  document.getElementById('sumErr').textContent   = failed;
  const failResults = results.filter(r => r.status !== 'success');
  const okResults   = results.filter(r => r.status === 'success');
  document.getElementById('failCount').textContent = failResults.length;
  document.getElementById('okCount').textContent   = okResults.length;
  const rowsHtml = list => list.map((r,i) => \`<tr><td style="color:var(--text-faint)">\${i+1}</td><td style="font-family:monospace">\${r.to}</td><td style="color:var(--text-faint);font-size:10px">\${r.message}</td></tr>\`).join('');
  document.getElementById('logBodyFail').innerHTML    = rowsHtml(failResults);
  document.getElementById('logBodySuccess').innerHTML = rowsHtml(okResults);
  document.getElementById('logFailEmpty').style.display = failResults.length ? 'none' : 'block';
  document.getElementById('emptyResult').style.display = 'none';
  document.getElementById('resultCard').style.display = 'block';
  refreshIcons();
  toast('Selesai! '+success+' berhasil, '+failed+' gagal.','ok');
}

function renderTemplatePage() {
  const wrap = document.getElementById('templateTplList');
  appData.templates.sort((a,b) => (a.order||999) - (b.order||999));
  if (!appData.templates.length) { wrap.innerHTML='<div class="empty">Belum ada template.</div>'; refreshIcons(); return; }
  wrap.innerHTML = appData.templates.map((t, i) => {
    const woCatLabel = t.woCategory === 'individu' ? '<span class="badge badge-blue" style="margin-left:4px">WO Individu</span>' : t.woCategory === 'marking' ? '<span class="badge badge-yellow" style="margin-left:4px">WO Marking</span>' : '<span class="badge badge-gray" style="margin-left:4px">Umum</span>';
    return \`
    <div class="tpl-list-item">
      <div class="tpl-list-info">
        <div class="tpl-list-name">\${t.name}
          \${t.hasParams && t.params?.length ? '<span class="badge badge-yellow" style="margin-left:4px">'+t.params.length+' param</span>' : ''}
          \${woCatLabel}
        </div>
        <div class="tpl-list-id">\${t.templateId}</div>
      </div>
      <div class="tpl-list-actions">
        <div style="display:flex;gap:2px;justify-content:flex-end">
          <button class="btn btn-ghost btn-sm" onclick="moveTemplate('\${t.id}', -1)" \${i===0?'disabled':''}>▲</button>
          <button class="btn btn-ghost btn-sm" onclick="moveTemplate('\${t.id}', 1)" \${i===appData.templates.length-1?'disabled':''}>▼</button>
        </div>
        <div style="display:flex;gap:2px">
          <button class="btn btn-warning btn-sm" onclick="openModal('\${t.id}')">Edit</button>
          <button class="btn btn-danger btn-sm" onclick="deleteTpl('\${t.id}')">Hapus</button>
        </div>
      </div>
    </div>
  \`;}).join('');
  refreshIcons();
}

async function moveTemplate(id, dir) {
  const arr    = appData.templates;
  const sorted = arr.slice().sort((a,b) => (a.order||0) - (b.order||0));
  const i = sorted.findIndex(t => t.id === id);
  if (i === -1) return;
  const j = i + dir;
  if (j < 0 || j >= sorted.length) return;
  const tmp = sorted[i]; sorted[i] = sorted[j]; sorted[j] = tmp;
  sorted.forEach((t, idx) => { t.order = idx + 1; });
  await fetch('/api/templates/reorder', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ ids: sorted.map(t=>({id: t.id, order: t.order})) }) });
  appData.templates = sorted;
  renderTemplatePage();
  renderDropdown();
}

// ============================================================
// QUICK REPLY (frontend)
// ============================================================

function escapeHtmlText(s) {
  return (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// Sama seperti escapeHtmlText, tapi juga meng-escape tanda kutip — dipakai saat
// menaruh teks bebas (mis. keterangan/label kontak) di dalam atribut HTML.
function escapeHtmlAttr(s) {
  return escapeHtmlText(s).replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function renderQuickReplyList() {
  const wrap = document.getElementById('quickReplyList');
  if (!wrap) return;
  if (!quickReplies.length) { wrap.innerHTML = '<div class="empty">Belum ada quick reply. Tambahkan di atas.</div>'; refreshIcons(); return; }
  wrap.innerHTML = quickReplies.map((q, i) => \`
    <div class="tpl-list-item">
      <div class="tpl-list-info">
        <div class="tpl-list-name" style="font-weight:600;font-size:12px;white-space:pre-wrap;word-break:break-word">\${escapeHtmlText(q.text)}</div>
      </div>
      <div class="tpl-list-actions">
        <div style="display:flex;gap:2px;justify-content:flex-end">
          <button class="btn btn-ghost btn-sm" onclick="moveQuickReply('\${q.id}', -1)" \${i===0?'disabled':''}>▲</button>
          <button class="btn btn-ghost btn-sm" onclick="moveQuickReply('\${q.id}', 1)" \${i===quickReplies.length-1?'disabled':''}>▼</button>
        </div>
        <div style="display:flex;gap:2px">
          <button class="btn btn-warning btn-sm" onclick="editQuickReply('\${q.id}')">Edit</button>
          <button class="btn btn-danger btn-sm" onclick="deleteQuickReply('\${q.id}')">Hapus</button>
        </div>
      </div>
    </div>
  \`).join('');
  refreshIcons();
}

async function addQuickReply() {
  const inp = document.getElementById('newQuickReplyText');
  const text = inp.value.trim();
  if (!text) { toast('Isi quick reply tidak boleh kosong.','err'); return; }
  const res  = await fetch('/api/quick-replies', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ text }) });
  const data = await res.json();
  if (data.error) { toast(data.error,'err'); return; }
  inp.value = '';
  quickReplies.push(data.quickReply);
  renderQuickReplyList();
  toast('Quick reply ditambahkan!','ok');
}

async function editQuickReply(id) {
  const q = quickReplies.find(x => x.id === id);
  if (!q) return;
  const newText = prompt('Edit isi quick reply:', q.text);
  if (newText === null) return;
  const trimmed = newText.trim();
  if (!trimmed) { toast('Isi tidak boleh kosong.','err'); return; }
  const res  = await fetch('/api/quick-replies/' + encodeURIComponent(id), { method:'PUT', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ text: trimmed }) });
  const data = await res.json();
  if (data.error) { toast(data.error,'err'); return; }
  q.text = trimmed;
  renderQuickReplyList();
  toast('Quick reply diperbarui.','ok');
}

async function deleteQuickReply(id) {
  if (!confirm('Hapus quick reply ini?')) return;
  await fetch('/api/quick-replies/' + encodeURIComponent(id), { method: 'DELETE' });
  quickReplies = quickReplies.filter(q => q.id !== id);
  renderQuickReplyList();
  toast('Quick reply dihapus.','ok');
}

async function moveQuickReply(id, dir) {
  const i = quickReplies.findIndex(q => q.id === id);
  if (i === -1) return;
  const j = i + dir;
  if (j < 0 || j >= quickReplies.length) return;
  const arr = quickReplies.slice();
  const tmp = arr[i]; arr[i] = arr[j]; arr[j] = tmp;
  quickReplies = arr;
  renderQuickReplyList();
  await fetch('/api/quick-replies/reorder', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ ids: arr.map(q => q.id) }) });
}

function toggleQuickReplyMenu() {
  const menu = document.getElementById('quickReplyMenu');
  if (!menu) return;
  const isOpen = menu.style.display === 'block';
  if (isOpen) { menu.style.display = 'none'; return; }
  renderQuickReplyMenu();
  menu.style.display = 'block';
}

function renderQuickReplyMenu() {
  const menu = document.getElementById('quickReplyMenu');
  if (!menu) return;
  if (!quickReplies.length) {
    menu.innerHTML = '<div class="qr-menu-hd">Quick Reply</div><div class="qr-menu-empty">Belum ada quick reply.<br>Tambahkan di menu Pengaturan.</div>';
    return;
  }
  menu.innerHTML = '<div class="qr-menu-hd">Pilih Quick Reply</div>' + quickReplies.map((q, i) => \`<div class="qr-menu-item" data-qr-idx="\${i}">\${escapeHtmlText(q.text)}</div>\`).join('');
  menu.querySelectorAll('.qr-menu-item').forEach(el => {
    el.addEventListener('click', () => {
      const idx = parseInt(el.dataset.qrIdx, 10);
      insertQuickReply(quickReplies[idx]?.text || '');
    });
  });
}

function insertQuickReply(text) {
  const ta = document.getElementById('chatComposeText');
  if (ta) {
    ta.value = text;
    autoResizeCompose(ta);
    ta.focus();
  }
  const menu = document.getElementById('quickReplyMenu');
  if (menu) menu.style.display = 'none';
}

async function loadChatHistory(isPoll) {
  const data = await fetch('/api/contacts?limit=' + CONTACTS_PAGE_SIZE).then(r => r.json());
  const freshList = data.contacts || [];

  if (!isPoll || !allContacts.length) {
    allContacts = freshList;
    contactsHasMore = data.hasMore;
    contactsOldestTime = data.oldestTime;
  } else {
    const merged = [...allContacts];
    freshList.forEach(c => {
      const idx = merged.findIndex(x => x.phone === c.phone);
      if (idx >= 0) merged[idx] = c; else merged.unshift(c);
    });
    merged.sort((a, b) => (b.isUnread ? 1 : 0) - (a.isUnread ? 1 : 0) || new Date(b.lastTime) - new Date(a.lastTime));
    allContacts = merged;
    if (allContacts.length <= freshList.length) {
      contactsHasMore = data.hasMore;
      contactsOldestTime = data.oldestTime;
    }
  }

  const countEl = document.getElementById('chatContactCount');
  if (countEl) countEl.textContent = allContacts.length + ' kontak';
  updateUnreadCount();
}

async function loadOlderContacts() {
  if (loadingOlderContacts || !contactsHasMore) return;
  loadingOlderContacts = true;
  const btn = document.getElementById('loadOlderContactsBtn');
  if (btn) { btn.innerHTML = 'Memuat...'; btn.disabled = true; }
  try {
    const data = await fetch('/api/contacts?before=' + encodeURIComponent(contactsOldestTime) + '&limit=' + CONTACTS_PAGE_SIZE).then(r => r.json());
    allContacts = [...allContacts, ...(data.contacts || [])];
    contactsHasMore = data.hasMore;
    contactsOldestTime = data.oldestTime || contactsOldestTime;
    const countEl = document.getElementById('chatContactCount');
    if (countEl) countEl.textContent = allContacts.length + ' kontak';
    filterContacts();
  } catch (e) {
    toast('Gagal memuat kontak lama: ' + e.message, 'err');
  } finally {
    loadingOlderContacts = false;
  }
}

async function markContactRead(phone) {
  const c = allContacts.find(x => x.phone === phone) || tabContacts.find(x => x.phone === phone);
  if (c && !c.isUnread) return; 
  if (c) {
    c.isUnread = 0;
    if (globalUnreadCount > 0) globalUnreadCount--; 
    updateUnreadCount();
  }
  try {
    await fetch('/api/contacts/' + encodeURIComponent(phone) + '/read', { method: 'POST' });
  } catch (e) { /* tidak fatal */ }
}

function updateUnreadCount() {
  const el = document.getElementById('unreadCount');
  if (!el) return;
  if (globalUnreadCount > 0) { el.textContent = globalUnreadCount; el.style.display = 'inline-block'; }
  else el.style.display = 'none';
}

async function loadTabContacts(reset) {
  if (activeChatTab === 'all' || loadingTabContacts) return;
  loadingTabContacts = true;
  const btn = document.getElementById('loadOlderContactsBtn');
  if (btn && !reset) { btn.innerHTML = 'Memuat...'; btn.disabled = true; }
  try {
    const params = new URLSearchParams({ limit: CONTACTS_PAGE_SIZE });
    if (activeChatTab === 'marking')  params.set('category', 'marking');
    if (activeChatTab === 'individu') params.set('category', 'individu');
    if (activeChatTab === 'unread')   params.set('unread', '1');
    if (!reset && tabOldestTime) params.set('before', tabOldestTime);

    const data = await fetch('/api/contacts?' + params.toString()).then(r => r.json());
    tabContacts = reset ? (data.contacts || []) : [...tabContacts, ...(data.contacts || [])];
    tabHasMore = data.hasMore;
    tabOldestTime = data.oldestTime;
    renderChatContacts(tabContacts);
  } catch (e) {
    toast('Gagal memuat data: ' + e.message, 'err');
  } finally {
    loadingTabContacts = false;
  }
}

function filterTab(tab, el) {
  activeChatTab = tab;
  document.querySelectorAll('.chat-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  filterContacts();
}

function clearSearch() {
  document.getElementById('chatSearch').value = '';
  document.getElementById('searchClear').style.display = 'none';
  filterContacts();
}

let searchDebounceTimer = null;
let isSearchActive = false;

function filterContacts() {
  const searchEl = document.getElementById('chatSearch');
  const clearEl  = document.getElementById('searchClear');
  if (!searchEl || !clearEl) return;
  const q = searchEl.value.trim().toLowerCase();
  clearEl.style.display = q ? 'flex' : 'none';

  if (!q) {
    isSearchActive = false;
    currentSearchQuery = '';
    searchContacts = [];
    searchHasMore = false;
    searchOldestTime = null;
    if (activeChatTab === 'all') {
      renderChatContacts(allContacts);
    } else {
      loadTabContacts(true);
    }
    return;
  }

  // Urutan belum-terbaca di atas juga dijaga saat sedang mencari (hasil pencarian
  // tidak lewat query /api/contacts yang sudah ORDER BY isUnread, jadi disortir
  // ulang di sini). Sort JS stabil, jadi urutan waktu di dalam grup tetap terjaga.
  const byUnreadFirst = (list) => [...list].sort((a, b) => (b.isUnread ? 1 : 0) - (a.isUnread ? 1 : 0));

  const applyTabFilter = (list) => {
    if (activeChatTab === 'unread')        return list.filter(c => c.isUnread);
    if (activeChatTab === 'individu')      return byUnreadFirst(list.filter(c => c.woCategory === 'individu'));
    if (activeChatTab === 'marking')       return byUnreadFirst(list.filter(c => c.woCategory === 'marking'));
    return byUnreadFirst(list);
  };

  isSearchActive = true;
  const quickMatch = allContacts.filter(c =>
    c.phone.includes(q) || (c.name && c.name.toLowerCase().includes(q))
  );
  renderChatContacts(applyTabFilter(quickMatch));

  clearTimeout(searchDebounceTimer);
  searchDebounceTimer = setTimeout(async () => {
    if (document.getElementById('chatSearch').value.trim().toLowerCase() !== q) return;
    try {
      const data = await fetch('/api/contacts/search?q=' + encodeURIComponent(q)).then(r => r.json());
      if (document.getElementById('chatSearch').value.trim().toLowerCase() !== q) return;
      currentSearchQuery = q;
      searchContacts = data.contacts || [];
      searchHasMore = data.hasMore;
      searchOldestTime = data.oldestTime;
      renderChatContacts(applyTabFilter(searchContacts));
    } catch (e) { }
  }, 350);
}

async function loadOlderSearchResults() {
  if (loadingSearchMore || !searchHasMore || !currentSearchQuery) return;
  loadingSearchMore = true;
  const btn = document.getElementById('loadOlderContactsBtn');
  if (btn) { btn.innerHTML = 'Memuat...'; btn.disabled = true; }
  try {
    const params = new URLSearchParams({ q: currentSearchQuery });
    if (searchOldestTime) params.set('before', searchOldestTime);
    const data = await fetch('/api/contacts/search?' + params.toString()).then(r => r.json());

    // Kalau kata kunci sudah berubah selagi fetch berjalan (user lanjut ngetik),
    // hasil basi ini dibuang supaya tidak nyampur dengan pencarian yang baru.
    const searchEl = document.getElementById('chatSearch');
    if (!searchEl || searchEl.value.trim().toLowerCase() !== currentSearchQuery) return;

    searchContacts = [...searchContacts, ...(data.contacts || [])];
    searchHasMore = data.hasMore;
    searchOldestTime = data.oldestTime || searchOldestTime;

    const byUnreadFirst = (list) => [...list].sort((a, b) => (b.isUnread ? 1 : 0) - (a.isUnread ? 1 : 0));
    const applyTabFilter = (list) => {
      if (activeChatTab === 'unread')        return list.filter(c => c.isUnread);
      if (activeChatTab === 'individu')      return byUnreadFirst(list.filter(c => c.woCategory === 'individu'));
      if (activeChatTab === 'marking')       return byUnreadFirst(list.filter(c => c.woCategory === 'marking'));
      return byUnreadFirst(list);
    };
    renderChatContacts(applyTabFilter(searchContacts));
  } catch (e) {
    toast('Gagal memuat hasil pencarian sebelumnya: ' + e.message, 'err');
  } finally {
    loadingSearchMore = false;
  }
}

function renderChatContacts(contacts) {
  const list = contacts || allContacts;
  currentRenderedContacts = list; 
  const wrap = document.getElementById('chatContactsList');
  if (!list.length) {
    wrap.innerHTML = \`<div class="empty" style="padding:20px;text-align:center;color:var(--text-faint)">\${activeChatTab === 'unread' ? 'Tidak ada pesan belum dibaca' : activeChatTab === 'marking' ? 'Tidak ada kontak WO Marking' : activeChatTab === 'individu' ? 'Tidak ada kontak WO Individu' : 'Belum ada riwayat.'}</div>\`;
    refreshIcons(); return;
  }
  const showLoadMore = isSearchActive
    ? searchHasMore
    : (activeChatTab === 'all' ? contactsHasMore : tabHasMore);
  const loadMoreFn   = isSearchActive
    ? 'loadOlderSearchResults()'
    : (activeChatTab === 'all' ? 'loadOlderContacts()' : 'loadTabContacts(false)');
  const loadMoreLabel = isSearchActive ? 'Muat hasil pencarian sebelumnya' : 'Muat 50 kontak sebelumnya';
  const loadMoreHtml = showLoadMore
    ? \`<div style="text-align:center;padding:8px 0 12px"><button class="btn btn-ghost btn-sm" id="loadOlderContactsBtn" onclick="\${loadMoreFn}" style="display:inline-flex;align-items:center;gap:4px"><i data-lucide="chevron-down" style="width:12px;height:12px"></i> \${loadMoreLabel}</button></div>\`
    : '';
  wrap.innerHTML = list.map(c => {
    const t = c.lastTime ? new Date(c.lastTime) : null;
    const timeStr = t ? t.getHours().toString().padStart(2,'0')+':'+t.getMinutes().toString().padStart(2,'0') : '';
    const initials = c.phone.slice(-4);
    const lastText = c.lastText || '';
    const preview  = lastText.slice(0,30) + (lastText.length > 30 ? '…' : '');
    const displayName = c.name ? c.name : c.phone;
    const isLastIncoming = c.lastDirection === 'in';
    const isUnread = !!c.isUnread;
    const catBadge = c.woCategory === 'individu' ? '<span class="badge badge-blue">Individu</span>' : c.woCategory === 'marking' ? '<span class="badge badge-yellow">Marking</span>' : '';
    return \`<div class="chat-contact-item \${selectedContact?.phone===c.phone?'active':''} \${isUnread?'has-unread':''}" onclick="selectChatContact('\${c.phone}')">
      <div class="chat-contact-avatar" style="\${isLastIncoming?'background:linear-gradient(135deg,#16a34a,#22c55e)':''}">\${initials}</div>
      <div class="chat-contact-info">
        <span class="chat-contact-name-text">\${displayName}</span>
        <div class="chat-contact-last">
          \${isLastIncoming ? '<i data-lucide="corner-down-left" style="width:10px;height:10px"></i>' : '<i data-lucide="corner-up-right" style="width:10px;height:10px"></i>'}
          \${preview||'—'}
        </div>
      </div>
      <div class="chat-contact-right">
        <div class="right-top">\${catBadge}</div>
        <div class="right-bot">
          \${isUnread ? '<span class="chat-unread-dot"></span>' : ''}
          <span class="chat-time-text">\${timeStr}</span>
        </div>
      </div>
    </div>\`;
  }).join('') + loadMoreHtml;
  refreshIcons();
}

let currentMsgs = [];
let currentMsgsHasMore = false;
let currentMsgsOldest = null;
let loadingOlderMsgs = false;
const MSG_PAGE_SIZE = 10; 

function renderMessageBubbles(msgs) {
  let lastDateKey = '';
  return msgs.map(m => {
    const t  = new Date(m.time);
    const ts = t.getHours().toString().padStart(2,'0')+':'+t.getMinutes().toString().padStart(2,'0');
    const isPending = m.status === 'pending';
    const ok = m.status === 'success' || m.status === 'delivered' || m.status === 'read';
    const isIn = m.direction === 'in';
    const escaped = (m.text||'').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\\n/g,'<br>');
    const url   = m.attachmentUrl || '';
    const mtype = m.msgType || 'text';

    const dateKey = t.getFullYear() + '-' + (t.getMonth()+1) + '-' + t.getDate();
    let dateSepHtml = '';
    if (dateKey !== lastDateKey) {
      lastDateKey = dateKey;
      dateSepHtml = '<div class="chat-date-sep"><span>' + formatDateSeparator(t) + '</span></div>';
    }

    let mediaHtml = '';
    if (url) {
      const safeUrl = url.replace(/"/g, '&quot;');
      if (mtype === 'image' || mtype === 'sticker' || /\\.(jpg|jpeg|png|gif|webp)(\\?|$)/i.test(url)) {
        mediaHtml = '<div style="margin-bottom:4px"><img src="' + safeUrl + '" alt="foto" style="max-width:220px;max-height:220px;border-radius:6px;display:block;cursor:pointer;object-fit:cover" onclick="openMedia(this)" onerror="mediaFallback(this)"></div>';
      } else if (mtype === 'video' || /\\.(mp4|mov|avi|mkv)(\\?|$)/i.test(url)) {
        mediaHtml = '<div class="video-thumb-wrap" style="margin-bottom:4px;position:relative;display:inline-block">'
          + '<video src="' + safeUrl + '" controls style="max-width:220px;max-height:180px;border-radius:6px;display:block" onerror="mediaFallback(this)"></video>'
          + '<button type="button" onclick="event.stopPropagation();openVideoLightbox(this)" title="Perbesar" style="position:absolute;top:5px;right:5px;width:24px;height:24px;border-radius:6px;border:none;background:rgba(0,0,0,.55);color:#fff;display:flex;align-items:center;justify-content:center;cursor:pointer;font-size:13px;line-height:1;padding:0">⛶</button>'
          + '</div>';
      } else if (mtype === 'audio' || mtype === 'voice' || /\\.(ogg|mp3|m4a|aac|wav)(\\?|$)/i.test(url)) {
        mediaHtml = '<div style="margin-bottom:4px"><audio src="' + safeUrl + '" controls style="max-width:220px;height:36px" onerror="mediaFallback(this)"></audio></div>';
      } else {
        const fname = url.split('/').pop().split('?')[0] || 'file';
        mediaHtml = '<div style="margin-bottom:4px"><a href="' + safeUrl + '" target="_blank" style="display:inline-flex;align-items:center;gap:6px;padding:6px 10px;border-radius:6px;background:rgba(0,0,0,.15);color:inherit;text-decoration:none;font-size:11px">📄 ' + fname + '</a></div>';
      }
    }
    const captionHtml = (escaped && escaped !== '[' + mtype + ']') ? '<div>' + escaped + '</div>' : (!url ? '<div>' + escaped + '</div>' : '');
    const metaSpan = '<span>' + ts + (m.nomerLapor ? ' · 📋 ' + m.nomerLapor : '') + (m.templateName && !isIn ? ' · ' + m.templateName : '') + '</span>';
    const tickSpan = !isIn ? '<span class="chat-bubble-status ' + (isPending ? 'pending' : (ok?'ok':'err')) + '">' + (isPending ? '🕐' : (ok?'✓✓':'✗')) + '</span>' : '';
    return dateSepHtml
      + '<div class="chat-msg-bubble' + (isPending ? ' pending' : '') + '" style="justify-content:' + (isIn?'flex-start':'flex-end') + '">'
      + '<div class="' + (isIn ? 'chat-bubble-in' : 'chat-bubble-out') + '">'
      + mediaHtml + captionHtml
      + '<div class="chat-bubble-meta">' + metaSpan + tickSpan + '</div>'
      + '</div></div>';
  }).join('');
}

function renderChatMsgArea(preserveScrollFromTop) {
  const area = document.getElementById('chatMsgArea');
  if (!area) return;
  const prevHeight = preserveScrollFromTop ? area.scrollHeight : 0;
  const loadMoreHtml = currentMsgsHasMore
    ? '<div style="text-align:center;padding:4px 0 10px"><button class="btn btn-ghost btn-sm" id="loadOlderBtn" onclick="loadOlderMessages()" style="display:inline-flex;align-items:center;gap:4px"><i data-lucide="chevron-down" style="width:12px;height:12px"></i> Muat ' + MSG_PAGE_SIZE + ' pesan sebelumnya</button></div>'
    : (currentMsgs.length ? '<div style="text-align:center;padding:4px 0 10px;font-size:10px;color:var(--text-faint)">Awal percakapan</div>' : '');
  area.innerHTML = loadMoreHtml + renderMessageBubbles(currentMsgs);
  refreshIcons();
  if (preserveScrollFromTop) area.scrollTop = area.scrollHeight - prevHeight;
  else area.scrollTop = area.scrollHeight;
}

async function loadOlderMessages() {
  if (loadingOlderMsgs || !currentMsgsHasMore || !selectedContact) return;
  loadingOlderMsgs = true;
  const btn = document.getElementById('loadOlderBtn');
  if (btn) { btn.innerHTML = 'Memuat...'; btn.disabled = true; }
  try {
    const phone = selectedContact.phone;
    const data = await fetch('/api/contacts/' + encodeURIComponent(phone) + '/messages?before=' + encodeURIComponent(currentMsgsOldest) + '&limit=' + MSG_PAGE_SIZE).then(r => r.json());
    if (selectedContact?.phone !== phone) return; 
    currentMsgs = [...(data.messages || []), ...currentMsgs];
    currentMsgsHasMore = data.hasMore;
    currentMsgsOldest  = data.oldestTime || currentMsgsOldest;
    renderChatMsgArea(true);
  } catch (e) {
    toast('Gagal memuat pesan lama: ' + e.message, 'err');
  } finally {
    loadingOlderMsgs = false;
  }
}

function formatDateSeparator(d) {
  const now = new Date();
  const strip = (x) => new Date(x.getFullYear(), x.getMonth(), x.getDate());
  const diffDays = Math.round((strip(now) - strip(d)) / 86400000);
  if (diffDays === 0) return 'Hari ini';
  if (diffDays === 1) return 'Kemarin';
  return d.toLocaleDateString('id-ID', { weekday: diffDays < 7 ? 'long' : undefined, day: 'numeric', month: 'long', year: 'numeric' });
}

// Update status "aktif" (kontak yang lagi dibuka) & "sudah dibaca" pada list
// kontak yang SEDANG ditampilkan (bisa berupa hasil pencarian, tab tertentu,
// atau semua kontak) — tanpa memicu ulang pencarian/fetch ke server, supaya
// klik kontak tidak bikin daftar kontak kedip.
function refreshContactListHighlight(phone) {
  [currentRenderedContacts, allContacts, tabContacts, searchContacts].forEach(list => {
    const item = list?.find(c => c.phone === phone);
    if (item) item.isUnread = 0;
  });
  renderChatContacts(currentRenderedContacts);
}

async function selectChatContact(phone) {
  selectedContact = currentRenderedContacts.find(c => c.phone === phone) || allContacts.find(c => c.phone === phone);
  if (selectedContact) markContactRead(phone);
  refreshContactListHighlight(phone);
  const main = document.getElementById('chatMain');
  if (!selectedContact) { main.innerHTML='<div class="chat-no-contact"><i data-lucide="message-circle"></i><div>Pilih kontak</div></div>'; refreshIcons(); return; }

  const displayName = selectedContact.name || phone;
  main.innerHTML = \`
    <div class="chat-main-header">
      <div class="chat-main-avatar">\${phone.slice(-4)}</div>
      <span class="chat-header-phone" id="chatHeaderPhone">\${phone}</span>
      <button class="chat-copy-btn" id="chatCopyBtn" title="Salin nomor" onclick="copyPhoneToClipboard('\${phone}')">
        <i data-lucide="copy" style="width:13px;height:13px"></i>
      </button>
      <div class="chat-main-info">
        <div class="chat-main-name">\${displayName}</div>
      </div>
      <span class="chat-header-labels" id="chatMainSub"></span>
      <button class="chat-copy-btn open-label-picker-btn" data-phone="\${phone}" title="Atur Label">
        <i data-lucide="tag" style="width:13px;height:13px"></i>
      </button>
    </div>
    <div class="chat-messages" id="chatMsgArea"><div class="empty" style="padding:20px;text-align:center;color:var(--text-faint)">Memuat pesan…</div></div>
    <div class="chat-compose" id="chatComposeBar">
      <button class="chat-attach-btn" onclick="document.getElementById('filePickerInput').click()" title="Kirim foto/video">
        <i data-lucide="paperclip" style="width:16px;height:16px"></i>
      </button>
      <input type="file" id="filePickerInput" accept="image/*,video/mp4,video/quicktime" style="display:none" onchange="handleFileAttach(this)">
      <button class="chat-attach-btn" id="quickReplyBtn" onclick="toggleQuickReplyMenu()" title="Quick Reply">
        <i data-lucide="message-square-quote" style="width:16px;height:16px"></i>
      </button>
      <div class="qr-menu" id="quickReplyMenu" style="display:none"></div>
      <div class="chat-compose-wrap" id="chatComposeWrap">
        <textarea class="chat-compose-textarea" id="chatComposeText" placeholder="Ketik pesan balasan..." rows="1"
          oninput="autoResizeCompose(this)"
          onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();sendChatReply();}"></textarea>
      </div>
      <button class="chat-compose-send" onclick="sendChatReply()" title="Kirim (Enter)"><i data-lucide="send" style="width:15px;height:15px"></i></button>
    </div>
    <div id="filePreviewBar" style="display:none;padding:6px 12px 4px;background:var(--topbar-bg);border-top:1px solid var(--chat-sidebar-border);align-items:center;gap:8px;font-size:11px;color:var(--text-muted)">
      <span id="filePreviewName" style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap"></span>
      <button onclick="clearFileAttach()" style="background:none;border:none;color:#ef4444;cursor:pointer;font-size:14px;padding:0 4px">✕</button>
    </div>
  \`;
  refreshIcons();
  setupDragDrop();

  await loadMasterLabels();
  let chatContactLabels = [];
  try { chatContactLabels = await fetch('/api/kontak/' + encodeURIComponent(phone) + '/tags').then(r => r.json()); }
  catch (e) { chatContactLabels = []; }
  if (selectedContact?.phone !== phone) return;
  renderChatHeaderLabels(phone, chatContactLabels);
  const labelBtn = main.querySelector('.open-label-picker-btn[data-phone="' + CSS.escape(phone) + '"]');
  if (labelBtn) {
    labelBtn.onclick = () => {
      const currentTags = chatContactLabels;
      openLabelPickerMenu(phone, labelBtn, currentTags, (tags) => {
        chatContactLabels = tags;
        renderChatHeaderLabels(phone, tags);
        const c = kontakContacts.find(x => x.phone === phone);
        if (c) c.labels = tags;
      });
    };
  }

  currentMsgs = [];
  currentMsgsHasMore = false;
  currentMsgsOldest = null;

  try {
    const data = await fetch('/api/contacts/' + encodeURIComponent(phone) + '/messages?limit=' + MSG_PAGE_SIZE).then(r => r.json());
    if (selectedContact?.phone !== phone) return; 
    currentMsgs = data.messages || [];
    currentMsgsHasMore = data.hasMore;
    currentMsgsOldest  = data.oldestTime;
    renderChatMsgArea(false);

    const area = document.getElementById('chatMsgArea');
    if (area) {
      area.onscroll = () => { if (area.scrollTop < 40) loadOlderMessages(); };
    }
  } catch (e) {
    const area = document.getElementById('chatMsgArea');
    if (area) area.innerHTML = '<div class="empty" style="padding:20px;text-align:center;color:var(--text-faint)">Gagal memuat pesan.</div>';
  }

  const ta = document.getElementById('chatComposeText');
  if (ta) setTimeout(() => ta.focus(), 80);
}

async function refreshChatMessages(phone) {
  if (!selectedContact || selectedContact.phone !== phone) return;
  if (!document.getElementById('chatMsgArea')) return; 

  // Jangan hitung pesan optimis (pending) sebagai "pesan terakhir" saat menentukan
  // titik "after" — pesan itu belum tentu sudah tersimpan di server dengan waktu
  // yang persis sama, jadi polling tetap pakai waktu pesan TERAKHIR YANG SUDAH
  // dikonfirmasi server supaya tidak ada pesan asli yang terlewat.
  const confirmedMsgs = currentMsgs.filter(m => m.status !== 'pending');
  const lastTime = confirmedMsgs.length ? confirmedMsgs[confirmedMsgs.length - 1].time : null;
  try {
    const url = lastTime
      ? '/api/contacts/' + encodeURIComponent(phone) + '/messages?after=' + encodeURIComponent(lastTime)
      : '/api/contacts/' + encodeURIComponent(phone) + '/messages?limit=' + MSG_PAGE_SIZE;
    const data = await fetch(url).then(r => r.json());
    if (selectedContact?.phone !== phone) return; 
    const newMsgs = data.messages || [];
    if (!newMsgs.length) return; 

    // Buang bubble optimis (pending) yang sudah tergantikan oleh data asli dari server,
    // supaya pesan tidak tampil dobel begitu balasan sungguhan datang lewat polling.
    const stillPending = currentMsgs.filter(m => m.status === 'pending');

    if (lastTime) {
      currentMsgs = [...confirmedMsgs, ...newMsgs, ...stillPending];
    } else {
      currentMsgs = [...newMsgs, ...stillPending];
      currentMsgsHasMore = data.hasMore;
      currentMsgsOldest  = data.oldestTime;
    }
    renderChatMsgArea(false); 
    markContactRead(phone);   
  } catch (e) { }
}

async function clearAllHistory() {
  if (!confirm('Hapus SEMUA riwayat chat? Tindakan ini tidak bisa dibatalkan.')) return;
  await fetch('/api/history', { method: 'DELETE' });
  allContacts = [];
  tabContacts = [];
  tabHasMore = false;
  tabOldestTime = null;
  globalUnreadCount = 0;
  selectedContact = null;
  contactsHasMore = false;
  contactsOldestTime = null;
  lastHistoryCount = 0;
  lastLatestTime = '';
  filterContacts();
  updateUnreadCount();
  document.getElementById('chatMain').innerHTML = '<div class="chat-no-contact"><i data-lucide="message-circle"></i><div>Pilih kontak untuk melihat riwayat pesan</div></div>';
  refreshIcons();
  toast('Semua riwayat dihapus.', 'ok');
}

async function deleteHistoryByRange() {
  const start = document.getElementById('cfgRangeStart').value;
  const end   = document.getElementById('cfgRangeEnd').value;
  if (!start || !end) { toast('Pilih tanggal mulai dan sampai terlebih dahulu.', 'err'); return; }
  if (start > end) { toast('Tanggal mulai tidak boleh lebih besar dari tanggal sampai.', 'err'); return; }

  const fmt = (d) => new Date(d + 'T00:00:00').toLocaleDateString('id-ID', { day:'numeric', month:'long', year:'numeric' });
  if (!confirm(\`Hapus semua riwayat pesan dari \${fmt(start)} sampai \${fmt(end)}? Tindakan ini tidak bisa dibatalkan.\`)) return;

  try {
    const res = await fetch(\`/api/history/range?start=\${start}&end=\${end}\`, { method: 'DELETE' });
    const data = await res.json();
    if (!res.ok || !data.success) { toast(data.error || 'Gagal menghapus riwayat.', 'err'); return; }

    lastHistoryCount = 0;
    lastLatestTime = '';
    await loadChatHistory();
    filterContacts();
    if (selectedContact) {
      const updated = allContacts.find(c => c.phone === selectedContact.phone);
      if (updated) { selectedContact = updated; selectChatContact(selectedContact.phone); }
      else {
        selectedContact = null;
        document.getElementById('chatMain').innerHTML = '<div class="chat-no-contact"><i data-lucide="message-circle"></i><div>Pilih kontak untuk melihat riwayat pesan</div></div>';
        refreshIcons();
      }
    }
    toast(\`\${data.deleted} riwayat berhasil dihapus.\`, 'ok');
  } catch (err) {
    toast('Gagal menghapus riwayat: ' + err.message, 'err');
  }
}

async function saveConfig() {
  const apiToken   = document.getElementById('cfgApiToken').value.trim();
  const webhookUrl = document.getElementById('cfgWebhookUrl').value.trim();
  const confirmPassword = prompt('Masukkan password pengaturan untuk menyimpan perubahan:');
  if (!confirmPassword) return;
  const res  = await fetch('/api/config', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ apiToken, webhookUrl, confirmPassword }) });
  const data = await res.json();
  if (data.error) { toast(data.error, 'err'); return; }
  toast('Konfigurasi berhasil disimpan!','ok');
}

async function saveLoginCredentials() {
  const loginUsername = document.getElementById('cfgLoginUsername').value.trim();
  const loginPassword = document.getElementById('cfgLoginPassword').value; // boleh kosong
  if (!loginUsername) { toast('Username tidak boleh kosong.', 'err'); return; }
  const confirmPassword = prompt('Masukkan password pengaturan untuk menyimpan perubahan:');
  if (!confirmPassword) return;
  const res  = await fetch('/api/login-credentials', {
    method: 'POST', headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ loginUsername, loginPassword, confirmPassword })
  });
  const data = await res.json();
  if (data.error) { toast(data.error, 'err'); return; }
  document.getElementById('cfgLoginPassword').value = '';
  toast('Kredensial login berhasil diperbarui!', 'ok');
}

// Format ISO time (UTC) jadi string tanggal+jam WIB yang enak dibaca,
// mis. "11 Jul 2026, 04:00 WIB".
function formatWIBTimestamp(isoStr) {
  if (!isoStr) return '–';
  const d = new Date(isoStr);
  const wib = new Date(d.getTime() + 7 * 60 * 60 * 1000);
  const months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const dd = wib.getUTCDate();
  const mm = months[wib.getUTCMonth()];
  const yyyy = wib.getUTCFullYear();
  const hh = String(wib.getUTCHours()).padStart(2,'0');
  const mi = String(wib.getUTCMinutes()).padStart(2,'0');
  return \`\${dd} \${mm} \${yyyy}, \${hh}:\${mi} WIB\`;
}

async function refreshSystemInfo() {
  try {
    const data = await fetch('/api/system/db-info').then(r => r.json());
    document.getElementById('dbSizeVal').textContent     = data.dbSizeFormatted || '–';
    document.getElementById('dbContactsVal').textContent = (data.broadcastCount ?? '–').toLocaleString?.('id-ID') ?? data.broadcastCount;
    document.getElementById('dbSessionsVal').textContent = (data.sessionsCount ?? '–').toLocaleString?.('id-ID') ?? data.sessionsCount;
    document.getElementById('lastBackupVal').textContent = data.lastBackup ? formatWIBTimestamp(data.lastBackup.mtime) + ' (' + data.lastBackup.sizeFormatted + ')' : 'Belum pernah';
    document.getElementById('backupCountVal').textContent = (data.backupCount || 0) + ' file (maks ' + (data.backupKeepCount || 7) + ')';
  } catch (e) {
    // Diamkan — kartu cukup menampilkan "–" kalau gagal fetch, tidak perlu toast mengganggu.
  }
}

async function runManualBackup() {
  const btn = document.getElementById('manualBackupBtn');
  if (btn) { btn.disabled = true; btn.innerHTML = '<i data-lucide="loader-2" style="width:11px;height:11px"></i>Membackup...'; refreshIcons(); }
  try {
    const res = await fetch('/api/system/backup-now', { method: 'POST' });
    const data = await res.json();
    if (!res.ok || !data.success) { toast(data.error || 'Backup gagal dijalankan.', 'err'); return; }
    toast('Backup database berhasil dibuat: ' + data.fileName, 'ok');
    await refreshSystemInfo();
  } catch (err) {
    toast('Gagal membuat backup: ' + err.message, 'err');
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = '<i data-lucide="hard-drive-download" style="width:11px;height:11px"></i>Backup Sekarang'; refreshIcons(); }
  }
}

function formatWIBDateShort(dateStr) {
  // dateStr format 'YYYY-MM-DD' (sudah dalam basis WIB dari query SQL)
  const [y, m, d] = dateStr.split('-').map(Number);
  const months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  return \`\${d} \${months[m - 1]}\`;
}

function formatRupiah(n) {
  return 'Rp ' + Math.round(n).toLocaleString('id-ID');
}

let bcGroupBy = 'day';
function setBcGroupBy(mode) {
  bcGroupBy = mode;
  document.getElementById('bcGroupDayBtn').classList.toggle('active', mode === 'day');
  document.getElementById('bcGroupMonthBtn').classList.toggle('active', mode === 'month');
  const rangeSel = document.getElementById('bcHistoryRange');
  if (rangeSel) {
    rangeSel.innerHTML = mode === 'month'
      ? '<option value="3">3 bulan terakhir</option><option value="6" selected>6 bulan terakhir</option><option value="12">12 bulan terakhir</option><option value="24">24 bulan terakhir</option>'
      : '<option value="7">7 hari terakhir</option><option value="14" selected>14 hari terakhir</option><option value="30">30 hari terakhir</option><option value="90">90 hari terakhir</option>';
  }
  loadBroadcastHistory();
}

function formatBcPeriode(str) {
  // Harian: 'YYYY-MM-DD' → '3 Agu'. Bulanan: 'YYYY-MM' → 'Agu 2026'.
  const months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  const parts = str.split('-').map(Number);
  if (parts.length === 3) { const [y,m,d] = parts; return \`\${d} \${months[m-1]}\`; }
  const [y,m] = parts; return \`\${months[m-1]} \${y}\`;
}

async function loadBroadcastHistory() {
  const body = document.getElementById('bcHistoryBody');
  const totalEl = document.getElementById('bcHistoryTotalCost');
  const days = document.getElementById('bcHistoryRange')?.value || (bcGroupBy === 'month' ? 6 : 14);
  try {
    const data = await fetch('/api/dashboard/broadcast-history?days=' + days + '&groupBy=' + bcGroupBy).then(r => r.json());
    const rows = data.rows || [];
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="6" class="empty">Belum ada data.</td></tr>';
      if (totalEl) totalEl.textContent = '';
      return;
    }
    body.innerHTML = rows.map(r => \`
      <tr>
        <td>\${formatBcPeriode(r.date)}</td>
        <td>\${r.totalBroadcast.toLocaleString('id-ID')}</td>
        <td style="color:#166534;font-weight:600">\${(r.totalBerhasil||0).toLocaleString('id-ID')}</td>
        <td style="color:#991b1b;font-weight:600">\${(r.totalGagal||0).toLocaleString('id-ID')}</td>
        <td>\${r.totalNomor.toLocaleString('id-ID')}</td>
        <td>\${formatRupiah(r.biaya)}</td>
      </tr>
    \`).join('');
    const totalCost = rows.reduce((a, r) => a + r.biaya, 0);
    const label = bcGroupBy === 'month' ? (rows.length + ' bulan terakhir') : (rows.length + ' hari terakhir');
    if (totalEl) totalEl.textContent = 'Total biaya ' + label + ' (hanya broadcast berhasil): ' + formatRupiah(totalCost) + ' (Rp 365/broadcast)';
  } catch (e) {
    body.innerHTML = '<tr><td colspan="6" class="empty">Gagal memuat data.</td></tr>';
    if (totalEl) totalEl.textContent = '';
  }
}

let exportFormatSelected = 'xlsx';

function setExportFormat(fmt, btn) {
  exportFormatSelected = fmt;
  document.querySelectorAll('#exportFormatToggle .toggle-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
}

function clearExportQuickActive() {
  document.querySelectorAll('.export-quick-btn').forEach(b => b.classList.remove('active'));
}

function pad2Client(n) { return String(n).padStart(2, '0'); }
function toDateInputStr(d) { return d.getFullYear() + '-' + pad2Client(d.getMonth() + 1) + '-' + pad2Client(d.getDate()); }

function setExportQuickRange(range, btnEl) {
  const today = new Date();
  let start, end;
  if (range === 'today') {
    start = end = today;
  } else if (range === '7') {
    end = today; start = new Date(today); start.setDate(start.getDate() - 6);
  } else if (range === '30') {
    end = today; start = new Date(today); start.setDate(start.getDate() - 29);
  } else if (range === 'month') {
    end = today; start = new Date(today.getFullYear(), today.getMonth(), 1);
  } else return;

  document.getElementById('exportRangeStart').value = toDateInputStr(start);
  document.getElementById('exportRangeEnd').value   = toDateInputStr(end);
  clearExportQuickActive();
  if (btnEl) btnEl.classList.add('active');
}

function downloadBroadcastHistoryExport(format) {
  const days = document.getElementById('bcHistoryRange')?.value || (bcGroupBy === 'month' ? 6 : 14);
  const url = '/api/dashboard/broadcast-history-export?days=' + encodeURIComponent(days)
    + '&groupBy=' + encodeURIComponent(bcGroupBy) + '&format=' + encodeURIComponent(format);
  window.location.href = url;
}

function downloadBroadcastExport() {
  const start  = document.getElementById('exportRangeStart').value;
  const end    = document.getElementById('exportRangeEnd').value;
  const format = exportFormatSelected;
  if (!start || !end) { toast('Pilih tanggal mulai dan sampai terlebih dahulu.', 'err'); return; }
  if (start > end) { toast('Tanggal mulai tidak boleh lebih besar dari tanggal sampai.', 'err'); return; }
  const url = '/api/dashboard/broadcast-export?start=' + encodeURIComponent(start) + '&end=' + encodeURIComponent(end) + '&format=' + encodeURIComponent(format);
  window.location.href = url;
}

async function loadFrequentReporters() {
  const body = document.getElementById('frBody');
  const minDays = document.getElementById('frMinDays')?.value || 3;
  const days    = document.getElementById('frDaysRange')?.value || 90;
  try {
    const data = await fetch('/api/dashboard/frequent-reporters?minDays=' + minDays + '&days=' + days + '&limit=50').then(r => r.json());
    const rows = data.rows || [];
    if (!rows.length) { body.innerHTML = '<tr><td colspan="5" class="empty">Belum ada nomor yang memenuhi kriteria ini.</td></tr>'; return; }
    body.innerHTML = rows.map(r => \`
      <tr>
        <td>\${r.phone}</td>
        <td>\${r.name || '<span style="color:var(--text-faint)">–</span>'}</td>
        <td><span class="badge badge-yellow">\${r.dayCount} hari</span></td>
        <td><span class="badge badge-blue">\${r.kaliBroadcast}x</span></td>
        <td>\${formatWIBTimestamp(r.lastTime)}</td>
      </tr>
    \`).join('');
  } catch (e) {
    body.innerHTML = '<tr><td colspan="5" class="empty">Gagal memuat data.</td></tr>';
  }
}

async function loadWoTrendChart() {
  const wrap = document.getElementById('woTrendChartWrap');
  const days = document.getElementById('woTrendRange')?.value || 30;
  wrap.innerHTML = '<div class="empty" style="padding:20px">Memuat chart...</div>';
  try {
    const data = await fetch('/api/dashboard/wo-trend?days=' + days).then(r => r.json());
    renderWoTrendChart(wrap, data.rows || []);
  } catch (e) {
    wrap.innerHTML = '<div class="empty" style="padding:20px">Gagal memuat chart.</div>';
  }
}

function renderWoTrendChart(wrap, rows) {
  if (!rows.length) { wrap.innerHTML = '<div class="empty" style="padding:20px">Belum ada data.</div>'; return; }

  const isDark   = document.body.classList.contains('dark');
  const colMark  = '#d97706'; // amber — WO Marking
  const colInd   = '#2563eb'; // biru — WO Individu
  const gridCol  = isDark ? '#2a3942' : '#e2e8f0';
  const textCol  = isDark ? '#8696a0' : '#64748b';

  const W = Math.max(rows.length * 46, 560);
  const H = 220;
  const padL = 34, padR = 12, padT = 16, padB = 30;
  const plotW = W - padL - padR;
  const plotH = H - padT - padB;

  const maxVal = Math.max(1, ...rows.map(r => Math.max(r.marking, r.individu)));
  const niceMax = Math.ceil(maxVal / 5) * 5 || 5;

  const xStep = rows.length > 1 ? plotW / (rows.length - 1) : 0;
  const xAt = i => padL + xStep * i;
  const yAt = v => padT + plotH - (v / niceMax) * plotH;

  const buildPath = key => rows.map((r, i) => \`\${i === 0 ? 'M' : 'L'} \${xAt(i).toFixed(1)} \${yAt(r[key]).toFixed(1)}\`).join(' ');

  // Garis grid horizontal (5 level) + label sumbu-Y
  let gridLines = '';
  for (let g = 0; g <= 5; g++) {
    const val = (niceMax / 5) * g;
    const y = yAt(val);
    gridLines += \`<line x1="\${padL}" y1="\${y.toFixed(1)}" x2="\${W - padR}" y2="\${y.toFixed(1)}" stroke="\${gridCol}" stroke-width="1"/>\`;
    gridLines += \`<text x="\${padL - 6}" y="\${(y + 3).toFixed(1)}" font-size="9" fill="\${textCol}" text-anchor="end">\${Math.round(val)}</text>\`;
  }

  // Label sumbu-X — supaya tidak penuh sesak, tampilkan maksimal ~10 label terpisah rata
  const labelEvery = Math.max(1, Math.ceil(rows.length / 10));
  let xLabels = '';
  rows.forEach((r, i) => {
    if (i % labelEvery !== 0 && i !== rows.length - 1) return;
    xLabels += \`<text x="\${xAt(i).toFixed(1)}" y="\${H - 8}" font-size="9" fill="\${textCol}" text-anchor="middle">\${formatWIBDateShort(r.date)}</text>\`;
  });

  const dotsMark = rows.map((r, i) => \`<circle cx="\${xAt(i).toFixed(1)}" cy="\${yAt(r.marking).toFixed(1)}" r="2.5" fill="\${colMark}"><title>\${formatWIBDateShort(r.date)}: \${r.marking} nomor WO Marking</title></circle>\`).join('');
  const dotsInd  = rows.map((r, i) => \`<circle cx="\${xAt(i).toFixed(1)}" cy="\${yAt(r.individu).toFixed(1)}" r="2.5" fill="\${colInd}"><title>\${formatWIBDateShort(r.date)}: \${r.individu} nomor WO Individu</title></circle>\`).join('');

  const totalMark = rows.reduce((a, r) => a + r.marking, 0);
  const totalInd  = rows.reduce((a, r) => a + r.individu, 0);

  wrap.innerHTML = \`
    <div style="display:flex;gap:16px;align-items:center;margin-bottom:8px;font-size:11px">
      <span style="display:flex;align-items:center;gap:5px"><span style="width:10px;height:10px;border-radius:50%;background:\${colMark};display:inline-block"></span>WO Marking (akumulasi \${totalMark.toLocaleString('id-ID')} nomor)</span>
      <span style="display:flex;align-items:center;gap:5px"><span style="width:10px;height:10px;border-radius:50%;background:\${colInd};display:inline-block"></span>WO Individu (akumulasi \${totalInd.toLocaleString('id-ID')} nomor)</span>
    </div>
    <svg viewBox="0 0 \${W} \${H}" style="width:100%;min-width:\${W}px;height:\${H}px;display:block">
      \${gridLines}
      \${xLabels}
      <path d="\${buildPath('marking')}" fill="none" stroke="\${colMark}" stroke-width="2"/>
      <path d="\${buildPath('individu')}" fill="none" stroke="\${colInd}" stroke-width="2"/>
      \${dotsMark}
      \${dotsInd}
    </svg>
  \`;
}

function autoResizeCompose(el) {
  el.style.height = 'auto';
  el.style.height = Math.min(el.scrollHeight, 100) + 'px';
}

let pendingFile = null;

function setupDragDrop() {
  const bar = document.getElementById('chatComposeBar');
  if (!bar) return;
  bar.addEventListener('dragover', e => { e.preventDefault(); bar.classList.add('drag-over'); });
  bar.addEventListener('dragleave', () => bar.classList.remove('drag-over'));
  bar.addEventListener('drop', e => { e.preventDefault(); bar.classList.remove('drag-over'); const file = e.dataTransfer.files[0]; if (file) attachFile(file); });
}

function handleFileAttach(input) {
  const file = input.files[0];
  if (file) attachFile(file);
  input.value = '';
}

function attachFile(file) {
  const allowed = ['image/jpeg','image/png','image/gif','image/webp','video/mp4','video/quicktime'];
  if (!allowed.includes(file.type)) { toast('Format tidak didukung. Gunakan JPG, PNG, atau MP4.', 'err'); return; }
  if (file.size > 16 * 1024 * 1024) { toast('File terlalu besar (maks 16MB)', 'err'); return; }
  pendingFile = file;
  const bar  = document.getElementById('filePreviewBar');
  const name = document.getElementById('filePreviewName');
  if (bar && name) { name.textContent = '📎 ' + file.name + ' (' + (file.size/1024).toFixed(0) + ' KB)'; bar.style.display = 'flex'; }
  const ta = document.getElementById('chatComposeText');
  if (ta) ta.placeholder = 'Tambah caption (opsional)...';
}

function clearFileAttach() {
  pendingFile = null;
  const bar = document.getElementById('filePreviewBar');
  if (bar) bar.style.display = 'none';
  const ta = document.getElementById('chatComposeText');
  if (ta) { ta.placeholder = 'Ketik pesan balasan...'; ta.focus(); }
}

async function uploadAndSendFile(phone, file, caption) {
  const formData = new FormData();
  formData.append('file', file);
  formData.append('to', phone);
  formData.append('caption', caption || '');
  const res = await fetch('/api/reply-media-upload', { method:'POST', body: formData });
  return await res.json();
}

// Balasan chat ditampilkan LANGSUNG di layar begitu tombol kirim ditekan
// (optimistic UI) — tidak lagi menunggu status "terkirim" dari server dulu.
// Bubble sementara (status "pending") langsung dirender, lalu request
// dikirim ke belakang layar; kalau sukses, reload halaman chat akan
// menggantikannya dengan data asli dari database (termasuk status centang).
// Kalau gagal, bubble sementara itu dihapus lagi dan pesan tidak hilang dari
// kolom ketik supaya bisa dicoba ulang.
async function sendChatReply() {
  if (!selectedContact) return;
  const textarea = document.getElementById('chatComposeText');
  if (!textarea) return;
  const text  = textarea.value.trim();
  if (!text && !pendingFile) return;
  const phone = selectedContact.phone;
  const btn   = document.querySelector('.chat-compose-send');
  if (btn) btn.disabled = true;

  const fileToSend = pendingFile;
  const isMediaMsg = !!fileToSend;
  const mediaKind  = isMediaMsg ? (fileToSend.type.startsWith('video/') ? 'video' : 'image') : 'text';

  // Tampilkan bubble pesan SEKARANG JUGA, sebelum menunggu jawaban API.
  const tempId = 'temp_' + Date.now() + '_' + Math.random().toString(36).slice(2);
  const optimisticMsg = {
    text: text || (isMediaMsg ? '[' + mediaKind + ']' : ''),
    time: new Date().toISOString(),
    status: 'pending',
    templateName: isMediaMsg ? '📎 File Terkirim' : '💬 Balasan Manual',
    direction: 'out',
    nomerLapor: '',
    msgType: mediaKind,
    attachmentUrl: isMediaMsg ? URL.createObjectURL(fileToSend) : '',
    _tempId: tempId,
  };
  currentMsgs.push(optimisticMsg);
  renderChatMsgArea(false);

  // Bersihkan kolom ketik/lampiran segera supaya terasa responsif.
  textarea.value = ''; textarea.style.height = 'auto';
  if (isMediaMsg) clearFileAttach();

  function removeOptimistic() {
    currentMsgs = currentMsgs.filter(m => m._tempId !== tempId);
    renderChatMsgArea(false);
  }

  try {
    if (isMediaMsg) {
      const data = await uploadAndSendFile(phone, fileToSend, text);
      if (data.error) { toast('Gagal kirim file: ' + data.error, 'err'); removeOptimistic(); return; }
      toast('File terkirim!', 'ok');
    } else {
      const res  = await fetch('/api/reply', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ to: phone, text }) });
      const data = await res.json();
      if (data.error) { toast('Gagal kirim: ' + data.error, 'err'); removeOptimistic(); return; }
      toast('Pesan terkirim!', 'ok');
    }
    await loadChatHistory(true);
    if (selectedContact?.phone === phone) {
      // User masih membuka kontak ini — aman untuk refresh tampilan pesannya
      // (biar status centang/terkirim ter-update dari data asli server).
      const updated = allContacts.find(c => c.phone === phone);
      if (updated) { selectedContact = updated; selectChatContact(phone); }
    } else {
      // User sudah pindah ke kontak lain sebelum pesan ini selesai terkirim —
      // JANGAN paksa balik ke kontak lama. Cukup segarkan daftar kontak di
      // sidebar (preview pesan terakhir & waktu) tanpa mengganggu layar aktif.
      filterContacts();
    }
  } catch(err) {
    toast('Gagal kirim: ' + err.message, 'err');
    removeOptimistic();
  } finally {
    if (btn) btn.disabled = false;
    const ta = document.getElementById('chatComposeText');
    if (ta) ta.focus();
  }
}

function toggleTokenVisibility() {
  const inp = document.getElementById('cfgApiToken');
  const btn = document.getElementById('tokenToggleBtn');
  if (inp.type === 'password') { inp.type = 'text'; btn.innerHTML = '<i data-lucide="eye-off" style="width:12px;height:12px"></i>'; }
  else { inp.type = 'password'; btn.innerHTML = '<i data-lucide="eye" style="width:12px;height:12px"></i>'; }
  refreshIcons();
}

function openModal(editId) {
  editingDirty = false;
  const overlay = document.getElementById('modalOverlay');
  overlay.classList.add('open');
  document.getElementById('paramBuilderList').innerHTML = '';
  paramBuilderCount = 0;
  if (editId) {
    const t = appData.templates.find(x => x.id === editId);
    document.getElementById('modalTitle').textContent    = 'Edit Template';
    document.getElementById('editingId').value           = editId;
    document.getElementById('mTplName').value            = t.name;
    document.getElementById('mTplId').value              = t.templateId;
    document.getElementById('mTplPreview').value         = t.preview || '';
    document.getElementById('mWoCategory').value         = t.woCategory || 'umum';
    document.getElementById('mHasParams').checked        = t.hasParams;
    toggleParamBuilder();
    if (t.hasParams && t.params?.length) t.params.forEach(p => addParamBuilder(p));
    setTimeout(() => {
      overlay.querySelectorAll('input,textarea,select').forEach(el => {
        el.addEventListener('input', markDirty);
        el.addEventListener('change', markDirty);
      });
    }, 100);
  } else {
    document.getElementById('modalTitle').textContent    = 'Tambah Template Baru';
    document.getElementById('editingId').value           = '';
    document.getElementById('mTplName').value            = '';
    document.getElementById('mTplId').value              = '';
    document.getElementById('mTplPreview').value         = '';
    document.getElementById('mWoCategory').value         = 'umum';
    document.getElementById('mHasParams').checked        = false;
    toggleParamBuilder();
  }
  updateModalSaveBtn();
  refreshIcons();
}

function markDirty() { editingDirty = true; updateModalSaveBtn(); }
function updateModalSaveBtn() {
  const btn = document.getElementById('modalSaveBtn');
  if (editingDirty) {
    btn.innerHTML = '<i data-lucide="alert-triangle" style="width:12px;height:12px"></i>Simpan Perubahan <span class="unsaved-dot"></span>';
    btn.style.background = 'linear-gradient(135deg,#b45309,#d97706)';
  } else {
    btn.innerHTML = '<i data-lucide="save" style="width:12px;height:12px"></i>Simpan Template';
    btn.style.background = '';
  }
  refreshIcons();
}

function closeModal() {
  if (editingDirty && !confirm('Ada perubahan yang belum disimpan. Yakin ingin menutup?')) return;
  document.getElementById('modalOverlay').classList.remove('open');
  editingDirty = false;
}
document.getElementById('modalOverlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modalOverlay')) closeModal();
});

// ============================================================
// KONTAK & LABEL (frontend)
// ============================================================
// Label sekarang MASTER: dibuat manual dulu lewat "Kelola Label", masing-
// masing jadi tab tersendiri di menu Kontak. Saat menempelkan label ke
// kontak, yang muncul cuma nama label yang sudah dibuat (tidak bisa lagi
// ketik teks bebas langsung di kontak).
let kontakContacts = [];
let kontakLabels = [];           // master label: [{id,name,color,order_num}]
let kontakActiveLabelFilter = null;
let kontakOffset = 0;
let kontakHasMore = false;
let kontakSearchTimer = null;
const KONTAK_PAGE_SIZE = 100;

async function loadKontakPage() {
  await loadMasterLabels();
  renderKontakLabelChips();
  kontakOffset = 0;
  await loadKontakList(true);
}

async function loadMasterLabels() {
  try { kontakLabels = await fetch('/api/labels').then(r => r.json()); }
  catch (e) { kontakLabels = []; }
}

async function ensureMasterLabelsLoaded() {
  if (kontakLabels.length) return;
  await loadMasterLabels();
}

// Tampilkan label yang ditempel di kontak yang lagi dibuka di menu Chat —
// supaya bisa lihat & atur label langsung sambil ngobrol, tanpa pindah ke
// menu Kontak.
function renderChatHeaderLabels(phone, tags) {
  const subEl = document.getElementById('chatMainSub');
  if (!subEl) return;
  subEl.innerHTML = (tags && tags.length)
    ? tags.map(t => {
        const color = labelColorFor(t);
        return \`<span class="badge" style="background:\${color}22;color:\${color};margin-right:4px">\${escapeHtmlText(t)}</span>\`;
      }).join('')
    : '';
}

// Tab "Semua" + satu tab per label master (bukan lagi cuma daftar teks yang
// pernah dipakai) — supaya label yang belum ditempel ke siapapun tetap
// muncul tabnya.
function renderKontakLabelChips() {
  const wrap = document.getElementById('kontakLabelChips');
  if (!wrap) return;
  const allActive = !kontakActiveLabelFilter;
  let html = \`<span class="badge \${allActive?'badge-blue':'badge-gray'}" style="cursor:pointer" data-tag-filter="">Semua</span>\`;
  html += kontakLabels.map(l => {
    const active = kontakActiveLabelFilter === l.name;
    return \`<span class="badge" style="cursor:pointer;background:\${active ? l.color : l.color+'22'};color:\${active ? '#fff' : l.color}" data-tag-filter="\${escapeHtmlAttr(l.name)}">\${escapeHtmlText(l.name)}</span>\`;
  }).join('');
  wrap.innerHTML = html || '<span style="font-size:10px;color:var(--text-faint)">Belum ada label. Buat lewat "Kelola Label".</span>';
  wrap.querySelectorAll('[data-tag-filter]').forEach(el => {
    el.onclick = () => setKontakLabelFilter(el.dataset.tagFilter || null);
  });
}

function setKontakLabelFilter(tag) {
  kontakActiveLabelFilter = tag;
  renderKontakLabelChips();
  kontakOffset = 0;
  loadKontakList(true);
}

function debounceLoadKontak() {
  clearTimeout(kontakSearchTimer);
  kontakSearchTimer = setTimeout(() => { kontakOffset = 0; loadKontakList(true); }, 300);
}

async function loadKontakList(reset) {
  const q = document.getElementById('kontakSearch')?.value.trim() || '';
  const params = new URLSearchParams({ limit: KONTAK_PAGE_SIZE, offset: reset ? 0 : kontakOffset });
  if (kontakActiveLabelFilter) params.set('tag', kontakActiveLabelFilter);
  if (q) params.set('q', q);
  try {
    const data = await fetch('/api/kontak?' + params.toString()).then(r => r.json());
    kontakContacts = reset ? (data.contacts || []) : [...kontakContacts, ...(data.contacts || [])];
    kontakHasMore = data.hasMore;
    kontakOffset = kontakContacts.length;
    const countEl = document.getElementById('kontakCount');
    if (countEl) countEl.textContent = data.total + ' kontak';
    renderKontakTable();
  } catch (e) {
    toast('Gagal memuat kontak: ' + e.message, 'err');
  }
}

function labelColorFor(name) {
  const l = kontakLabels.find(x => x.name === name);
  return l ? l.color : '#6b7a8d';
}

function renderKontakTable() {
  const body = document.getElementById('kontakTableBody');
  const loadMoreWrap = document.getElementById('kontakLoadMoreWrap');
  if (!body) return;
  if (!kontakContacts.length) {
    body.innerHTML = '<tr><td colspan="4" class="empty">Belum ada kontak.</td></tr>';
    if (loadMoreWrap) loadMoreWrap.innerHTML = '';
    refreshIcons();
    return;
  }
  body.innerHTML = kontakContacts.map(c => {
    const labelChips = (c.labels||[]).length
      ? c.labels.map(tag => {
          const color = labelColorFor(tag);
          return \`<span class="badge" style="display:inline-flex;align-items:center;gap:4px;background:\${color}22;color:\${color}">\${escapeHtmlText(tag)}<span class="tag-chip-remove" data-phone="\${escapeHtmlAttr(c.phone)}" data-tag="\${escapeHtmlAttr(tag)}" title="Lepas label ini" style="cursor:pointer;font-weight:700;line-height:1">×</span></span>\`;
        }).join(' ')
      : '<span style="color:var(--text-faint);font-size:10px">— belum ada label —</span>';
    const deleteBtn = c.isManualOnly
      ? \`<button class="btn btn-danger btn-sm" onclick="deleteManualKontak('\${c.phone}')" title="Hapus Kontak"><i data-lucide="trash-2" style="width:11px;height:11px"></i></button>\`
      : '';
    return \`<tr>
      <td>\${c.name ? escapeHtmlText(c.name) : '<span style="color:var(--text-faint)">— belum ada nama —</span>'}</td>
      <td style="font-family:monospace">\${c.phone}</td>
      <td>\${labelChips}</td>
      <td style="white-space:nowrap">
        <button class="btn btn-ghost btn-sm" onclick="editKontakName('\${c.phone}')" title="Ubah Nama"><i data-lucide="pencil" style="width:11px;height:11px"></i></button>
        <button class="btn btn-ghost btn-sm open-label-picker-btn" data-phone="\${escapeHtmlAttr(c.phone)}" title="Atur Label"><i data-lucide="tag" style="width:11px;height:11px"></i></button>
        \${deleteBtn}
      </td>
    </tr>\`;
  }).join('');
  if (loadMoreWrap) {
    loadMoreWrap.innerHTML = kontakHasMore
      ? '<button class="btn btn-ghost btn-sm" onclick="loadKontakList(false)">Muat lebih banyak</button>'
      : '';
  }
  body.querySelectorAll('.tag-chip-remove').forEach(el => {
    el.onclick = () => removeKontakLabel(el.dataset.phone, el.dataset.tag);
  });
  body.querySelectorAll('.open-label-picker-btn').forEach(btn => {
    btn.onclick = () => {
      const phone = btn.dataset.phone;
      const c = kontakContacts.find(x => x.phone === phone);
      openLabelPickerMenu(phone, btn, c?.labels || [], (tags) => {
        if (c) c.labels = tags;
        renderKontakTable();
      });
    };
  });
  refreshIcons();
}

async function removeKontakLabel(phone, tag) {
  try {
    await fetch('/api/kontak/' + encodeURIComponent(phone) + '/tags/' + encodeURIComponent(tag), { method: 'DELETE' });
    const c = kontakContacts.find(x => x.phone === phone);
    if (c) c.labels = (c.labels||[]).filter(l => l !== tag);
    renderKontakTable();
  } catch (err) {
    toast('Gagal melepas label: ' + err.message, 'err');
  }
}

// ------------------------------------------------------------
// Picker label bersama (dipakai di menu Kontak DAN menu Chat) — cuma
// menampilkan label yang sudah dibuat di "Kelola Label", dicentang/lepas
// per klik, tidak ada input teks bebas sama sekali.
// ------------------------------------------------------------
let labelPickerContext = null;

function openLabelPickerMenu(phone, anchorEl, currentTags, onChange) {
  const menu = document.getElementById('labelPickerMenu');
  if (!menu) return;
  labelPickerContext = { phone, tags: new Set(currentTags || []), onChange };
  menu.innerHTML = '<div class="qr-menu-hd">Label — ' + phone + '</div><div class="qr-menu-empty">Memuat…</div>';
  const rect = anchorEl.getBoundingClientRect();
  menu.style.top  = (rect.bottom + 6) + 'px';
  menu.style.left = Math.min(rect.left, window.innerWidth - 246) + 'px';
  menu.style.display = 'block';
  loadMasterLabels().then(() => {
    if (labelPickerContext?.phone === phone) renderLabelPickerMenu();
  });
}

function renderLabelPickerMenu() {
  const menu = document.getElementById('labelPickerMenu');
  if (!menu || !labelPickerContext) return;
  const { phone, tags } = labelPickerContext;
  menu.innerHTML = '<div class="qr-menu-hd">Label — ' + phone + '</div>' +
    (kontakLabels.length ? kontakLabels.map(l => \`
      <div class="qr-menu-item" style="display:flex;align-items:center;gap:8px;cursor:pointer" data-label-name="\${escapeHtmlAttr(l.name)}">
        <span style="width:14px;height:14px;border-radius:4px;border:1.5px solid \${tags.has(l.name) ? l.color : 'var(--input-border)'};background:\${tags.has(l.name) ? l.color : 'transparent'};display:inline-flex;align-items:center;justify-content:center;color:#fff;font-size:10px;flex-shrink:0">\${tags.has(l.name) ? '✓' : ''}</span>
        <span style="width:9px;height:9px;border-radius:50%;background:\${l.color};display:inline-block;flex-shrink:0"></span>
        <span>\${escapeHtmlText(l.name)}</span>
      </div>
    \`).join('') : '<div class="qr-menu-empty">Belum ada label.<br>Buat dulu lewat "Kelola Label".</div>');
  menu.querySelectorAll('[data-label-name]').forEach(el => {
    el.onclick = () => toggleLabelForContact(el.dataset.labelName);
  });
}

async function toggleLabelForContact(name) {
  const ctx = labelPickerContext;
  if (!ctx) return;
  const { phone, tags } = ctx;
  const isActive = tags.has(name);
  try {
    if (isActive) {
      await fetch('/api/kontak/' + encodeURIComponent(phone) + '/tags/' + encodeURIComponent(name), { method: 'DELETE' });
      tags.delete(name);
    } else {
      const res  = await fetch('/api/kontak/' + encodeURIComponent(phone) + '/tags', {
        method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ tag: name })
      });
      const data = await res.json();
      if (data.error) { toast(data.error, 'err'); return; }
      tags.add(name);
    }
    renderLabelPickerMenu();
    if (ctx.onChange) ctx.onChange([...tags]);
  } catch (err) {
    toast('Gagal update label: ' + err.message, 'err');
  }
}

document.addEventListener('click', e => {
  const menu = document.getElementById('labelPickerMenu');
  if (!menu || menu.style.display !== 'block') return;
  if (menu.contains(e.target) || e.target.closest('.open-label-picker-btn')) return;
  menu.style.display = 'none';
});

async function editKontakName(phone) {
  const c = kontakContacts.find(x => x.phone === phone);
  const newName = prompt('Nama untuk ' + phone + ':', c?.name || '');
  if (newName === null) return;
  const trimmed = newName.trim();
  try {
    await fetch('/api/kontak/' + encodeURIComponent(phone) + '/name', {
      method: 'PUT', headers: {'Content-Type':'application/json'}, body: JSON.stringify({ name: trimmed })
    });
    if (c) c.name = trimmed;
    renderKontakTable();
    toast('Nama kontak diperbarui.', 'ok');
  } catch (e) {
    toast('Gagal memperbarui nama: ' + e.message, 'err');
  }
}

function openAddKontakModal() {
  document.getElementById('newKontakPhone').value = '';
  document.getElementById('newKontakName').value = '';
  document.getElementById('addKontakModalOverlay').classList.add('open');
  setTimeout(() => document.getElementById('newKontakPhone')?.focus(), 80);
}
function closeAddKontakModal() {
  document.getElementById('addKontakModalOverlay').classList.remove('open');
}
document.getElementById('addKontakModalOverlay').addEventListener('click', e => {
  if (e.target === document.getElementById('addKontakModalOverlay')) closeAddKontakModal();
});

async function saveNewKontak() {
  const phoneRaw = document.getElementById('newKontakPhone').value.trim();
  const name     = document.getElementById('newKontakName').value.trim();
  if (!phoneRaw) { toast('Nomor wajib diisi.', 'err'); return; }
  try {
    const res  = await fetch('/api/kontak', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ phone: phoneRaw, name }) });
    const data = await res.json();
    if (data.error) { toast(data.error, 'err'); return; }
    closeAddKontakModal();
    kontakOffset = 0;
    await loadKontakList(true);
    toast('Kontak ditambahkan!', 'ok');
  } catch (e) {
    toast('Gagal menambah kontak: ' + e.message, 'err');
  }
}

async function deleteManualKontak(phone) {
  if (!confirm('Hapus kontak ' + phone + '? Kontak ini belum punya riwayat pesan.')) return;
  try {
    const res  = await fetch('/api/kontak/' + encodeURIComponent(phone), { method: 'DELETE' });
    const data = await res.json();
    if (data.error) { toast(data.error, 'err'); return; }
    kontakContacts = kontakContacts.filter(c => c.phone !== phone);
    renderKontakTable();
    toast('Kontak dihapus.', 'ok');
  } catch (e) {
    toast('Gagal menghapus kontak: ' + e.message, 'err');
  }
}

// ------------------------------------------------------------
// Kelola Label — modal untuk membuat/mengubah/menghapus label MASTER.
// ------------------------------------------------------------
function openLabelManageModal() {
  document.getElementById('labelModalOverlay').classList.add('open');
  renderLabelManageList();
}
function closeLabelManageModal() {
  document.getElementById('labelModalOverlay').classList.remove('open');
}
document.getElementById('labelModalOverlay').addEventListener('click', e => {
  if (e.target === document.getElementById('labelModalOverlay')) closeLabelManageModal();
});

function renderLabelManageList() {
  const wrap = document.getElementById('labelManageList');
  if (!wrap) return;
  if (!kontakLabels.length) { wrap.innerHTML = '<div class="empty">Belum ada label. Tambahkan di atas.</div>'; return; }
  wrap.innerHTML = kontakLabels.map(l => \`
    <div class="tpl-list-item">
      <div class="tpl-list-info" style="display:flex;align-items:center;gap:8px">
        <span style="width:14px;height:14px;border-radius:4px;background:\${l.color};flex-shrink:0"></span>
        <span class="tpl-list-name">\${escapeHtmlText(l.name)}</span>
      </div>
      <div class="tpl-list-actions">
        <div style="display:flex;gap:2px">
          <button class="btn btn-warning btn-sm" onclick="editLabel('\${l.id}')">Edit</button>
          <button class="btn btn-danger btn-sm" onclick="deleteLabel('\${l.id}')">Hapus</button>
        </div>
      </div>
    </div>
  \`).join('');
}

async function addLabel() {
  const name  = document.getElementById('newLabelName').value.trim();
  const color = document.getElementById('newLabelColor').value || '#2563eb';
  if (!name) { toast('Nama label wajib diisi.', 'err'); return; }
  const res  = await fetch('/api/labels', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ name, color }) });
  const data = await res.json();
  if (data.error) { toast(data.error, 'err'); return; }
  kontakLabels.push(data.label);
  document.getElementById('newLabelName').value = '';
  renderLabelManageList();
  renderKontakLabelChips();
  toast('Label ditambahkan! Tab-nya sudah muncul di menu Kontak.', 'ok');
}

async function editLabel(id) {
  const l = kontakLabels.find(x => x.id === id);
  if (!l) return;
  const newName = prompt('Nama label:', l.name);
  if (newName === null) return;
  const trimmed = newName.trim();
  if (!trimmed) { toast('Nama tidak boleh kosong.', 'err'); return; }
  const res  = await fetch('/api/labels/' + id, { method:'PUT', headers:{'Content-Type':'application/json'}, body: JSON.stringify({ name: trimmed, color: l.color }) });
  const data = await res.json();
  if (data.error) { toast(data.error, 'err'); return; }
  const oldName = l.name;
  l.name = trimmed;
  kontakContacts.forEach(c => { c.labels = (c.labels||[]).map(t => t === oldName ? trimmed : t); });
  if (kontakActiveLabelFilter === oldName) kontakActiveLabelFilter = trimmed;
  renderLabelManageList();
  renderKontakLabelChips();
  renderKontakTable();
  toast('Label diperbarui.', 'ok');
}

async function deleteLabel(id) {
  const l = kontakLabels.find(x => x.id === id);
  if (!l) return;
  if (!confirm('Hapus label "' + l.name + '"? Label ini akan lepas dari semua kontak yang memakainya.')) return;
  await fetch('/api/labels/' + id, { method: 'DELETE' });
  kontakLabels = kontakLabels.filter(x => x.id !== id);
  kontakContacts.forEach(c => { c.labels = (c.labels||[]).filter(t => t !== l.name); });
  if (kontakActiveLabelFilter === l.name) kontakActiveLabelFilter = null;
  renderLabelManageList();
  renderKontakLabelChips();
  renderKontakTable();
  toast('Label dihapus.', 'ok');
}

function toggleParamBuilder() {
  const show = document.getElementById('mHasParams').checked;
  document.getElementById('paramBuilderSection').style.display = show ? 'block' : 'none';
  markDirty();
}

let paramBuilderCount = 0;
function addParamBuilder(existing) {
  paramBuilderCount++;
  const idx = paramBuilderCount;
  const p   = existing || { label:'', inputType:'manual', required:true, options:[], default:'' };
  const wrap = document.getElementById('paramBuilderList');
  const div  = document.createElement('div');
  div.className = 'param-builder'; div.id = 'pb_' + idx;
  div.innerHTML = \`
    <div class="param-builder-hd">
      <span class="param-builder-title">Parameter \${existing ? existing.index : idx}</span>
      <button class="remove-param" onclick="removeParamBuilder('pb_\${idx}')">×</button>
    </div>
    <div class="field-row" style="margin-bottom:7px">
      <div class="fg">
        <label class="field-label">Nama Parameter</label>
        <input type="text" class="pb-label" placeholder="contoh: Nama UP3" value="\${p.label}">
      </div>
      <div class="fg">
        <label class="field-label">Tipe Input</label>
        <div class="input-type-toggle">
          <button type="button" class="toggle-btn \${p.inputType==='manual'?'active':''}" onclick="setInputType(this,'pb_\${idx}','manual')">Manual</button>
          <button type="button" class="toggle-btn \${p.inputType==='dropdown'?'active':''}" onclick="setInputType(this,'pb_\${idx}','dropdown')">Dropdown</button>
          <button type="button" class="toggle-btn \${p.inputType==='nomer_lapor'?'active':''}" onclick="setInputType(this,'pb_\${idx}','nomer_lapor')">Nomer Lapor</button>
        </div>
      </div>
    </div>
    <div class="field-row" style="margin-bottom:7px">
      <div class="fg">
        <label class="field-label">Nilai Default (opsional)</label>
        <input type="text" class="pb-default" placeholder="Kosongkan jika tidak ada" value="\${p.default||''}">
      </div>
      <div class="fg" style="display:flex;align-items:center;gap:6px;padding-top:16px">
        <input type="checkbox" class="pb-required" id="req_\${idx}" style="width:13px;height:13px;cursor:pointer" \${p.required?'checked':''}>
        <label for="req_\${idx}" style="font-size:11px;cursor:pointer;font-weight:600;color:var(--text)">Wajib diisi</label>
      </div>
    </div>
    <div class="pb-dd-section fg" style="display:\${p.inputType==='dropdown'?'block':'none'}">
      <label class="field-label">Pilihan Dropdown</label>
      <div class="tag-wrap" id="ddTags_\${idx}">
        <input class="tag-input" placeholder="Ketik pilihan, tekan Enter..." onkeydown="handleDDTag(event, \${idx})">
      </div>
      <div class="field-hint">Tekan Enter untuk tambah · × untuk hapus</div>
    </div>
  \`;
  wrap.appendChild(div);
  if (p.options?.length) p.options.forEach(opt => addDDTag(idx, opt));
  div.querySelectorAll('input,select').forEach(el => el.addEventListener('input', markDirty));
}

function removeParamBuilder(pbId) { document.getElementById(pbId)?.remove(); markDirty(); }
function setInputType(btn, pbId, type) {
  const pb = document.getElementById(pbId);
  pb.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  pb.querySelector('.pb-dd-section').style.display = type === 'dropdown' ? 'block' : 'none';
  let hint = pb.querySelector('.pb-nomer-lapor-hint');
  if (type === 'nomer_lapor') {
    if (!hint) {
      hint = document.createElement('div');
      hint.className = 'pb-nomer-lapor-hint field-hint';
      hint.style.cssText = 'background:#eff6ff;border:1px solid #93c5fd;border-radius:6px;padding:5px 9px;margin-top:5px;color:#1d4ed8;font-size:10px';
      hint.innerHTML = '📱 Nilai akan otomatis diambil dari <strong>kode/nomer lapor</strong> yang ada di kolom Nomor Tujuan (bagian pertama sebelum spasi).';
      pb.appendChild(hint);
    }
  } else { hint?.remove(); }
  markDirty();
}
function handleDDTag(e, idx) {
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const val = e.target.value.trim();
  if (val) { addDDTag(idx, val); e.target.value = ''; markDirty(); }
}
function addDDTag(idx, val) {
  const wrap = document.getElementById('ddTags_' + idx);
  const inp  = wrap.querySelector('.tag-input');
  const tag  = document.createElement('div');
  tag.className = 'tag'; tag.dataset.val = val;
  tag.innerHTML = \`\${val}<span class="tag-del" onclick="this.parentElement.remove();markDirty()">×</span>\`;
  wrap.insertBefore(tag, inp);
}

async function saveTemplate() {
  const name    = document.getElementById('mTplName').value.trim();
  const tplId   = document.getElementById('mTplId').value.trim();
  const preview = document.getElementById('mTplPreview').value.trim();
  const hasPrm  = document.getElementById('mHasParams').checked;
  const editId  = document.getElementById('editingId').value;
  if (!name || !tplId) { toast('Nama dan Template ID wajib diisi.','err'); return; }
  const params = [];
  if (hasPrm) {
    document.querySelectorAll('.param-builder').forEach((pb, i) => {
      const label    = pb.querySelector('.pb-label').value.trim() || ('Parameter ' + (i+1));
      const required = pb.querySelector('.pb-required').checked;
      const defVal   = pb.querySelector('.pb-default').value.trim();
      const activeBtn = pb.querySelector('.toggle-btn.active');
      const activeTxt = activeBtn?.textContent.trim().toLowerCase();
      const inputType = activeTxt === 'dropdown' ? 'dropdown' : (activeTxt === 'nomer lapor' ? 'nomer_lapor' : 'manual');
      const options  = [];
      pb.querySelectorAll('.tag[data-val]').forEach(t => { if (t.dataset.val && !options.includes(t.dataset.val)) options.push(t.dataset.val); });
      params.push({ index: i+1, label, inputType, required, options:[...new Set(options)], default: defVal });
    });
  }
  const body = { name, templateId: tplId, hasParams: hasPrm, preview, params, woCategory: document.getElementById('mWoCategory').value || 'umum' };
  let res, data;
  if (editId) {
    res  = await fetch('/api/templates/' + editId, { method:'PUT', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body) });
    data = await res.json();
    if (data.success) {
      const idx = appData.templates.findIndex(t => t.id === editId);
      if (idx !== -1) appData.templates[idx] = data.template;
      if (selectedTpl?.id === editId) { selectedTpl = data.template; renderParamFields(); renderPreview(); }
      toast('Template berhasil diperbarui!','ok');
    }
  } else {
    res  = await fetch('/api/templates', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body) });
    data = await res.json();
    if (data.success) { appData.templates.push(data.template); toast('Template berhasil ditambahkan!','ok'); }
  }
  if (data.error) { toast(data.error,'err'); return; }
  editingDirty = false;
  document.getElementById('modalOverlay').classList.remove('open');
  renderDropdown();
  renderTemplatePage();
}

async function deleteTpl(id) {
  if (!confirm('Hapus template ini?')) return;
  await fetch('/api/templates/' + id, { method: 'DELETE' });
  appData.templates = appData.templates.filter(t => t.id !== id);
  if (selectedTpl?.id === id) {
    selectedTpl = null;
    const lbl = document.getElementById('tplBtnLabel');
    lbl.className = 'dd-placeholder';
    lbl.textContent = '— Pilih template —';
    document.getElementById('paramCard').style.display = 'none';
    document.getElementById('previewEmpty').style.display = 'block';
    document.getElementById('waPhoneWrap').style.display = 'none';
  }
  renderDropdown();
  renderTemplatePage();
  toast('Template dihapus.','ok');
}

init();
</script>
</body>
</html>`));

const PORT = 3002;

try {
  initDb();
  console.log('✅ Database SQLite (better-sqlite3) siap');
} catch (err) {
  console.error('❌ Gagal inisialisasi database:', err);
  process.exit(1);
}

// Cek jadwal backup otomatis tiap menit. Dicek juga langsung saat startup,
// jaga-jaga kalau server baru nyala/restart tepat di jendela 04:00-04:04 WIB.
maybeRunScheduledBackup();
setInterval(maybeRunScheduledBackup, 60 * 1000);

const server = app.listen(PORT, () => {
  console.log('✅ Berjalan di http://localhost:' + PORT);
});
// Default Node.js keepAliveTimeout (5s) lebih pendek dari idle timeout koneksi
// yang dijaga cloudflared, jadi rawan race condition: cloudflared coba pakai
// ulang koneksi yang baru saja ditutup Node → request tampak "nyantol" ~5 detik
// di "Dalam Antrean" sebelum akhirnya dicoba ulang lewat koneksi baru.
// Dinaikkan supaya lebih besar dari idle timeout tunnel pada umumnya.
server.keepAliveTimeout = 65000;
server.headersTimeout   = 66000; // wajib lebih besar dari keepAliveTimeout

process.on('exit',    () => { try { db?.close(); } catch (_) {} });
process.on('SIGINT',  () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));