/**
 * ============================================================================
 * UNITRACK — Setup.gs  (fail BARU, jangan ubah Code.gs)
 * ============================================================================
 * Cara guna:
 *   1. Tambah fail ini dalam projek Apps Script (Files > + > Script > namakan Setup).
 *   2. Run  setupSystemPreview()  -> baca log, TIADA apa diubah.
 *   3. Run  setupSystem()         -> backup dulu, baru tambah tab/header yang kurang.
 *   Selamat run berulang kali (idempotent).
 *
 * Jaminan keselamatan data:
 *   - Tiada sel dipadam, tiada tab dipadam, tiada header dialih atau ditukar nama.
 *   - Header baru hanya ditambah di HUJUNG baris header. Kalau ruang itu ada data, ditolak.
 *   - Salinan penuh (backup) dibuat SEBELUM sebarang perubahan. Backup gagal = setup berhenti.
 *   - Selepas siap, isi 4 tab asal dibandingkan semula dengan keadaan asal (hash). Tak sama = ralat.
 *   - Tab Cert ialah sumber tunggal staf (termasuk lajur IC, IC Attachment, Designation).
 *     API menapis IC daripada respons awam. PASTIKAN tab Cert TIDAK dipublish ke web.
 *   - Arkib padam (Deleted_Rows) dan log kerja (Job_Log) disimpan dalam Spreadsheet PRIVATE berasingan.
 * ============================================================================
 */

const SETUP_PROP_PRIVATE_SS = 'PRIVATE_SS_ID';
const SETUP_PROP_BACKUP_FOLDER = 'BACKUP_FOLDER_ID';
const SETUP_PRIVATE_SS_NAME = 'UniTrack_Private (JANGAN PUBLISH)';
const SETUP_BACKUP_FOLDER_NAME = 'UniTrack_Backups';
const SETUP_BACKUP_KEEP = 30; // simpan 30 salinan terkini setiap fail

// Lajur yang mesti wujud (ditambah di hujung baris header jika kurang).
const SETUP_PUBLIC_EXTRA_HEADERS = {
  chargeman: ['ID'],                                  // padankan Chargeman dengan Staff ID (VESxxx)
  cert: ['IC', 'IC Attachment', 'Designation']        // data sulit staf (ditapis daripada API awam)
};

// Tab dalam spreadsheet PRIVATE.
const SETUP_PRIVATE_SCHEMA = {
  Deleted_Rows: ['Tarikh', 'Kategori', 'Tab Asal', 'No Baris', 'Data (JSON)'],
  Job_Log:      ['Tarikh', 'Projek', 'Dibuat Oleh', 'Senarai ID', 'Bil Orang', 'PDF Link']
};

/** Run ini dulu. Tiada apa diubah, hanya laporan. */
function setupSystemPreview() {
  return setupRun_(true);
}

/** Run ini untuk betul-betul bina tab dan header. */
function setupSystem() {
  return setupRun_(false);
}

function setupRun_(dry) {
  const log = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka projek ini melalui Extensions > Apps Script dari dalam Google Sheet.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // halang doPost menulis semasa setup berjalan
  try {
    log.push(dry ? '=== PREVIEW (tiada perubahan dibuat) ===' : '=== SETUP SISTEM ===');

    // 1. Rakam keadaan asal 4 tab.
    const before = setupFingerprint_(ss);
    Object.keys(before).forEach(function (k) {
      log.push(before[k]
        ? 'Tab ' + k + ': ' + before[k].rows + ' baris, ' + before[k].cols + ' lajur (dirakam)'
        : 'AMARAN: tab ' + k + ' (gid ' + TAB_GIDS[k] + ') tidak dijumpai. Tidak disentuh.');
    });

    // 2. Backup. Gagal = berhenti (fail closed).
    if (dry) {
      log.push('[PREVIEW] Backup penuh akan dibuat dahulu ke folder ' + SETUP_BACKUP_FOLDER_NAME + '.');
    } else {
      const bk = backupNow();
      log.push('Backup siap sebelum perubahan: ' + bk);
    }

    // 3. Tab awam: tambah header yang kurang (tanpa sentuh data).
    Object.keys(SETUP_PUBLIC_EXTRA_HEADERS).forEach(function (key) {
      const sh = setupSheetByGid_(ss, TAB_GIDS[key]);
      if (!sh) return;
      setupEnsureHeaders_(sh, setupHeaderRow_(sh), SETUP_PUBLIC_EXTRA_HEADERS[key], dry, log, key, false);
    });

    // 3b. Amaran: lajur IC dalam tab awam dan PIN lalai.
    setupWarnPublicIc_(ss, log);
    if (!PropertiesService.getScriptProperties().getProperty('ADMIN_PIN')) {
      log.push('AMARAN: Script Property ADMIN_PIN belum ditetapkan, PIN masih nilai lalai dalam kod. Tetapkan di Project Settings > Script Properties.');
    }

    // 4. Senarai staf daripada tab Cert (untuk seed Staff dan padankan Chargeman).
    const people = setupReadPeople_(ss);
    log.push('Staf dijumpai dalam tab Cert: ' + people.length);

    // 5. Isi lajur ID Chargeman hanya bila nama sama tepat dan unik (sel kosong sahaja).
    setupFillChargemanId_(ss, people, dry, log);

    // 5b. Lajur IC sebagai teks (kekalkan sifar di depan). Format sahaja, nilai tidak diubah.
    setupFormatIcColumn_(ss, dry, log);

    // 6. Spreadsheet private (Deleted_Rows, Job_Log).
    setupPrivate_(ss, people, dry, log);

    // 7. Backup harian automatik.
    if (!dry) {
      try {
        log.push(setupInstallBackupTrigger_());
      } catch (err) {
        log.push('AMARAN: trigger backup harian gagal dipasang (' + err.message + '). Run installDailyBackupTrigger() secara manual.');
      }
    } else {
      log.push('[PREVIEW] Trigger backup harian (2 pagi) akan dipasang.');
    }

    // 8. Sahkan data asal tidak berubah.
    const problems = setupVerify_(ss, before);
    if (problems.length) {
      log.push('RALAT: DATA ASAL BERUBAH -> ' + problems.join('; ') + '. Restore daripada backup di folder ' + SETUP_BACKUP_FOLDER_NAME + '.');
      Logger.log(log.join('\n'));
      throw new Error('Pengesahan data gagal: ' + problems.join('; '));
    }
    log.push('PENGESAHAN: isi asal 4 tab tidak berubah. OK.');
  } finally {
    lock.releaseLock();
  }

  const teks = log.join('\n');
  Logger.log(teks);
  return teks;
}

/* ----------------------------------------------------------------------------
 * HELPER SHEET
 * -------------------------------------------------------------------------- */

function setupSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === Number(gid)) return sheets[i];
  }
  return null;
}

/** Logik sama dengan locateHeaderRow() dalam Code.gs supaya baris header konsisten. */
function setupHeaderRow_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return null;
  const scan = Math.min(lastRow, 5);
  const values = sheet.getRange(1, 1, scan, lastCol).getDisplayValues();
  const kata = ['plate', 'expired', 'expiry', 'chargeman', 'cidb', 'name'];
  for (let r = 0; r < values.length; r++) {
    const joined = values[r].join(' ').toLowerCase();
    if (kata.some(function (k) { return joined.indexOf(k) !== -1; })) {
      return { row: r + 1, headers: values[r].map(function (h) { return String(h).trim(); }) };
    }
  }
  return { row: 1, headers: values[0].map(function (h) { return String(h).trim(); }) };
}

/**
 * Tambah header yang kurang di hujung baris header. Tidak pernah menulis ke sel yang ada data.
 * info = null bermakna tab kosong (header ditulis di baris 1).
 */
function setupEnsureHeaders_(sheet, info, wanted, dry, log, label, styleHeader) {
  const row = info ? info.row : 1;
  const headers = info ? info.headers.slice() : [];
  const lower = headers.map(function (h) { return h.toLowerCase(); });
  const missing = wanted.filter(function (w) { return lower.indexOf(w.toLowerCase()) === -1; });

  if (!missing.length) {
    if (styleHeader && !dry) setupStyleHeader_(sheet, row, Math.max(headers.length, wanted.length));
    return { row: row, headers: headers, added: [], blocked: false };
  }

  let lastUsed = 0;
  headers.forEach(function (h, i) { if (h !== '') lastUsed = i + 1; });
  const startCol = lastUsed + 1;

  // Ruang di bawah header untuk lajur baru mesti kosong.
  const lastRow = sheet.getLastRow();
  const availCols = Math.min(missing.length, Math.max(0, sheet.getMaxColumns() - startCol + 1));
  if (lastRow > row && availCols > 0) {
    const bawah = sheet.getRange(row + 1, startCol, lastRow - row, availCols).getValues();
    const adaData = bawah.some(function (r) {
      return r.some(function (c) { return String(c).trim() !== ''; });
    });
    if (adaData) {
      log.push('AMARAN [' + label + ']: lajur ' + startCol + ' dan seterusnya sudah ada data tanpa header. Header ' +
        missing.join(', ') + ' TIDAK ditambah. Semak manual.');
      return { row: row, headers: headers, added: [], blocked: true };
    }
  }

  const newHeaders = headers.slice(0, lastUsed).concat(missing);
  if (dry) {
    log.push('[PREVIEW] [' + label + '] header akan ditambah di lajur ' + startCol + ': ' + missing.join(', '));
    return { row: row, headers: newHeaders, added: missing, blocked: false };
  }

  const perlu = startCol + missing.length - 1 - sheet.getMaxColumns();
  if (perlu > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), perlu);
  const target = sheet.getRange(row, startCol, 1, missing.length);
  target.setValues([missing]);
  if (styleHeader) {
    setupStyleHeader_(sheet, row, newHeaders.length);
  } else if (startCol > 1) {
    try {
      sheet.getRange(row, startCol - 1).copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    } catch (e) { /* format tidak kritikal */ }
  }
  log.push('Header ditambah [' + label + '] di lajur ' + startCol + ': ' + missing.join(', '));
  return { row: row, headers: newHeaders, added: missing, blocked: false };
}

function setupStyleHeader_(sheet, row, ncols) {
  sheet.getRange(row, 1, 1, ncols).setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  sheet.setFrozenRows(row);
}

function setupWriteRows_(sheet, startRow, values) {
  if (!values.length) return;
  const perlu = startRow + values.length - 1 - sheet.getMaxRows();
  if (perlu > 0) sheet.insertRowsAfter(sheet.getMaxRows(), perlu);
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}

function setupNorm_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}

/* ----------------------------------------------------------------------------
 * PENGESAHAN DATA (hash)
 * -------------------------------------------------------------------------- */

function setupHashRows_(values, skipCol, hdrIdx, blankCols) {
  const rows = values.map(function (r, ri) {
    let x = r;
    // Sel header yang asalnya kosong boleh diisi setup (header baru), jadi diabaikan.
    if (ri === hdrIdx && blankCols.length) {
      x = r.map(function (c, ci) { return blankCols.indexOf(ci) !== -1 ? '' : c; });
    }
    return skipCol < 0 ? x : x.filter(function (_, i) { return i !== skipCol; });
  });
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(rows));
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function setupFingerprint_(ss) {
  const out = {};
  Object.keys(TAB_GIDS).forEach(function (key) {
    const sh = setupSheetByGid_(ss, TAB_GIDS[key]);
    if (!sh) { out[key] = null; return; }
    const r = sh.getLastRow();
    const c = sh.getLastColumn();
    // Lajur ID Chargeman dikecualikan kerana langkah 5 boleh mengisi sel kosongnya.
    const info = setupHeaderRow_(sh);
    let skipCol = -1;
    if (key === 'chargeman' && info) {
      skipCol = info.headers.map(function (h) { return h.toLowerCase(); }).indexOf('id');
      if (skipCol === -1) { // lajur ID akan ditambah di hujung header
        let lastUsed = 0;
        info.headers.forEach(function (h, i) { if (h !== '') lastUsed = i + 1; });
        skipCol = lastUsed; // indeks 0-based lajur selepas header terakhir
      }
    }
    const hdrIdx = info ? info.row - 1 : -1;
    const blankCols = [];
    if (info) for (let i = 0; i < c; i++) { if (!info.headers[i]) blankCols.push(i); }
    out[key] = {
      rows: r, cols: c, skipCol: skipCol, hdrIdx: hdrIdx, blankCols: blankCols,
      hash: r && c ? setupHashRows_(sh.getRange(1, 1, r, c).getValues(), skipCol, hdrIdx, blankCols) : ''
    };
  });
  return out;
}

function setupVerify_(ss, before) {
  const problems = [];
  Object.keys(before).forEach(function (key) {
    const b = before[key];
    if (!b) return;
    const sh = setupSheetByGid_(ss, TAB_GIDS[key]);
    if (!sh) { problems.push('tab ' + key + ' hilang'); return; }
    if (sh.getLastRow() < b.rows) problems.push('tab ' + key + ' kurang baris');
    const now = b.rows && b.cols ? setupHashRows_(sh.getRange(1, 1, b.rows, b.cols).getValues(), b.skipCol, b.hdrIdx, b.blankCols) : '';
    if (now !== b.hash) problems.push('isi tab ' + key + ' tidak sama dengan asal');
  });
  return problems;
}

/**
 * Lapor lajur sulit (IC dan lampirannya) dalam tab Awam dan ingatkan tentang "Publish to web".
 * Lajur IC dalam tab Cert dijangka. API menapisnya, tetapi CSV "Publish to web" TIDAK menapis.
 * Tidak mengubah apa-apa.
 */
function setupWarnPublicIc_(ss, log) {
  Object.keys(TAB_GIDS).forEach(function (key) {
    const sh = setupSheetByGid_(ss, TAB_GIDS[key]);
    if (!sh) return;
    const info = setupHeaderRow_(sh);
    if (!info) return;
    info.headers.forEach(function (h, i) {
      if (!h || !staffIsSensitiveKey_(h)) return;
      const last = sh.getLastRow();
      let isi = 0;
      if (last > info.row) {
        sh.getRange(info.row + 1, i + 1, last - info.row, 1).getValues().forEach(function (r) {
          if (String(r[0]).trim() !== '') isi++;
        });
      }
      log.push('PERHATIAN: tab "' + key + '" ada lajur sulit "' + h + '" (lajur ' + (i + 1) + ', ' + isi +
        ' sel berisi). API menapisnya, TAPI pastikan tab ini TIDAK dipublish ke web (File > Share > Publish to web).');
    });
  });
  const lain = ss.getSheetByName('Staff');
  if (lain) {
    log.push('PERHATIAN: ada tab "Staff" dalam spreadsheet utama. Sistem tidak membaca atau menulis tab itu. ' +
      'Jika ia mengandungi IC, pastikan ia tidak dipublish.');
  }
}

/* ----------------------------------------------------------------------------
 * STAF DAN PADANAN CHARGEMAN
 * -------------------------------------------------------------------------- */

function setupReadPeople_(ss) {
  const sh = setupSheetByGid_(ss, TAB_GIDS.cert);
  if (!sh) return [];
  const info = setupHeaderRow_(sh);
  if (!info) return [];
  const hl = info.headers.map(function (h) { return h.toLowerCase(); });
  const iId = hl.indexOf('id');
  const iName = hl.indexOf('name');
  if (iId === -1 || iName === -1) return [];
  const first = info.row + 1;
  const last = sh.getLastRow();
  if (last < first) return [];
  const vals = sh.getRange(first, 1, last - info.row, sh.getLastColumn()).getValues();
  const seen = {};
  const out = [];
  vals.forEach(function (r) {
    const id = String(r[iId]).trim().toUpperCase();
    if (!id || seen[id]) return;
    seen[id] = true;
    out.push({ id: id, name: String(r[iName]).trim() });
  });
  return out;
}

function setupFillChargemanId_(ss, people, dry, log) {
  const sh = setupSheetByGid_(ss, TAB_GIDS.chargeman);
  if (!sh) return;
  const info = setupHeaderRow_(sh);
  if (!info) return;
  const hl = info.headers.map(function (h) { return h.toLowerCase(); });
  let iName = hl.indexOf('chargeman name');
  if (iName === -1) iName = hl.indexOf('name');
  if (iName === -1) { log.push('AMARAN: lajur nama tidak dijumpai dalam tab Chargeman. ID tidak diisi.'); return; }
  const iId = hl.indexOf('id');
  if (iId === -1 && !dry) { return; } // header ditolak tadi, jangan tulis apa-apa

  const first = info.row + 1;
  const last = sh.getLastRow();
  if (last < first) return;
  const n = last - info.row;
  const names = sh.getRange(first, iName + 1, n, 1).getValues();
  const ids = iId >= 0
    ? sh.getRange(first, iId + 1, n, 1).getValues()
    : names.map(function () { return ['']; });

  const peta = {};
  people.forEach(function (p) {
    const k = setupNorm_(p.name);
    (peta[k] = peta[k] || []).push(p.id);
  });

  let diisi = 0;
  const tiadaPadanan = [];
  for (let i = 0; i < n; i++) {
    if (String(ids[i][0]).trim() !== '') continue; // jangan ganti ID sedia ada
    const nm = String(names[i][0]).trim();
    if (!nm) continue;
    const calon = peta[setupNorm_(nm)];
    if (calon && calon.length === 1) {
      diisi++;
      if (!dry) sh.getRange(first + i, iId + 1).setValue(calon[0]); // satu sel sahaja
    } else {
      tiadaPadanan.push('baris ' + (first + i) + ': ' + nm + (calon && calon.length > 1 ? ' (nama berganda)' : ''));
    }
  }
  log.push((dry ? '[PREVIEW] ' : '') + 'ID Chargeman ' + (dry ? 'akan diisi' : 'diisi') + ' automatik: ' + diisi + ' baris.');
  if (tiadaPadanan.length) {
    log.push('Perlu isi ID manual (nama tak sama tepat dengan tab Cert): ' + tiadaPadanan.join(' | '));
  }
}

/* ----------------------------------------------------------------------------
 * SPREADSHEET PRIVATE
 * -------------------------------------------------------------------------- */

function setupPrivate_(ss, people, dry, log) {
  const props = PropertiesService.getScriptProperties();
  const sedia = props.getProperty(SETUP_PROP_PRIVATE_SS);
  let pss = null;

  if (sedia) {
    try {
      pss = SpreadsheetApp.openById(sedia);
    } catch (err) {
      // Jangan cipta yang baru senyap-senyap: arkib lama akan jadi yatim.
      throw new Error('PRIVATE_SS_ID wujud (' + sedia + ') tetapi fail tidak boleh dibuka. Semak akses/fail itu. Setup dihentikan.');
    }
  } else if (dry) {
    log.push('[PREVIEW] Spreadsheet private "' + SETUP_PRIVATE_SS_NAME + '" akan dicipta dengan tab: ' +
      Object.keys(SETUP_PRIVATE_SCHEMA).join(', ') + '.');
    return;
  } else {
    pss = SpreadsheetApp.create(SETUP_PRIVATE_SS_NAME);
    props.setProperty(SETUP_PROP_PRIVATE_SS, pss.getId());
    log.push('Spreadsheet private dicipta: ' + pss.getUrl());
    try {
      const induk = DriveApp.getFileById(ss.getId()).getParents();
      if (induk.hasNext()) DriveApp.getFileById(pss.getId()).moveTo(induk.next());
    } catch (e) {
      log.push('Nota: gagal alih fail private ke folder yang sama (' + e.message + '). Tidak kritikal.');
    }
  }

  const baru = {};
  Object.keys(SETUP_PRIVATE_SCHEMA).forEach(function (nama) {
    let sh = pss.getSheetByName(nama);
    if (!sh) {
      if (dry) { log.push('[PREVIEW] Tab private akan dicipta: ' + nama); return; }
      sh = pss.insertSheet(nama);
      baru[nama] = true;
      log.push('Tab private dicipta: ' + nama);
    }
    setupEnsureHeaders_(sh, setupHeaderRow_(sh) || null, SETUP_PRIVATE_SCHEMA[nama], dry, log, nama, true);
  });

  // Tab lama (Staff, Users) daripada versi setup terdahulu: dibiarkan, tidak dipadam, tidak digunakan lagi.
  ['Staff', 'Users'].forEach(function (lama) {
    if (pss.getSheetByName(lama)) {
      log.push('Nota: tab "' + lama + '" lama dalam spreadsheet private tidak digunakan lagi (tidak dipadam). Data staf kini dalam tab Cert.');
    }
  });

  if (dry) return;

  // Buang tab lalai kosong hanya dalam fail private yang kita cipta sendiri.
  if (Object.keys(baru).length === Object.keys(SETUP_PRIVATE_SCHEMA).length) {
    ['Sheet1', 'Helaian1'].forEach(function (lazim) {
      const d = pss.getSheetByName(lazim);
      if (d && d.getLastRow() === 0 && d.getLastColumn() === 0 && pss.getSheets().length > 1) pss.deleteSheet(d);
    });
  }
}

/** Format teks pada lajur IC dalam tab Cert (tidak mengubah nilai). */
function setupFormatIcColumn_(ss, dry, log) {
  const sh = setupSheetByGid_(ss, TAB_GIDS.cert);
  if (!sh) return;
  const info = setupHeaderRow_(sh);
  if (!info) return;
  const i = info.headers.map(function (h) { return h.toLowerCase(); }).indexOf('ic');
  if (i === -1) return; // belum wujud (dry-run) atau ditolak
  if (dry) { log.push('[PREVIEW] Lajur IC dalam tab Cert akan diformat sebagai teks.'); return; }
  sh.getRange(info.row + 1, i + 1, Math.max(sh.getMaxRows() - info.row, 1), 1).setNumberFormat('@');
  log.push('Lajur IC (tab Cert) diformat sebagai teks.');
}

/* ----------------------------------------------------------------------------
 * BACKUP
 * -------------------------------------------------------------------------- */

function setupBackupFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(SETUP_PROP_BACKUP_FOLDER);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* cipta semula di bawah */ }
  }
  const folder = DriveApp.createFolder(SETUP_BACKUP_FOLDER_NAME);
  props.setProperty(SETUP_PROP_BACKUP_FOLDER, folder.getId());
  return folder;
}

/** Salin penuh spreadsheet utama (dan private jika ada). Boleh run manual bila-bila masa. */
function backupNow() {
  const folder = setupBackupFolder_();
  const tz = Session.getScriptTimeZone() || 'Asia/Kuala_Lumpur';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  const fail = [SpreadsheetApp.getActiveSpreadsheet()];
  const pid = PropertiesService.getScriptProperties().getProperty(SETUP_PROP_PRIVATE_SS);
  if (pid) {
    try {
      fail.push(SpreadsheetApp.openById(pid));
    } catch (err) {
      // Jangan teruskan senyap-senyap: data private mungkin tidak ter-backup.
      throw new Error('PRIVATE_SS_ID wujud (' + pid + ') tetapi fail tidak boleh dibuka. Semak akses/fail itu. Dihentikan.');
    }
  }

  const dibuat = [];
  fail.forEach(function (s) {
    const prefix = s.getName() + ' | BACKUP ';
    DriveApp.getFileById(s.getId()).makeCopy(prefix + stamp, folder);
    dibuat.push(prefix + stamp);
    setupPruneBackups_(folder, prefix);
  });
  return dibuat.join(', ');
}

/** Hanya fail bernama "<nama> | BACKUP ..." dalam folder backup. Dihantar ke Trash (boleh pulih 30 hari). */
function setupPruneBackups_(folder, prefix) {
  const senarai = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getName().indexOf(prefix) === 0) senarai.push(f);
  }
  senarai.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  senarai.slice(SETUP_BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });
}

function setupInstallBackupTrigger_() {
  const ada = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'backupNow'; });
  if (ada) return 'Trigger backup harian sudah wujud.';
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(2).create();
  return 'Trigger backup harian dipasang (sekitar 2-3 pagi).';
}

function installDailyBackupTrigger() {
  return setupInstallBackupTrigger_();
}
