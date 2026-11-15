/**
 * ============================================================================
 * UNITRACK COMPLIANCE — GOOGLE APPS SCRIPT BACKEND (Code.gs) — VERSI PENUH
 * ============================================================================
 * Ganti SELURUH isi Code.gs lama dengan fail ini. Fail lain yang kekal: Setup.gs.
 * (Jangan ada fail Staff.gs berasingan: kandungannya sudah digabung di bawah.)
 *
 * Fungsi:
 * 1. Membaca 4 tab mengikut GID (Kereta, Motosikal, Chargeman, Cert).
 * 2. Menapis lajur IC / NRIC / lampiran IC daripada semua respons awam.
 * 3. staffUpdateCert: staf kemas kini tarikh dan muat naik sijil sendiri (tanpa PIN),
 *    tetapi HANYA lajur sijil yang dibenarkan.
 * 4. Operasi Admin (PIN): addRow / updateRow / deleteRow. Padam diarkibkan dahulu.
 * 5. Data sulit (no. IC, gambar IC, Designation) disimpan dalam tab Cert (lajur IC, IC Attachment, Designation).
 *    Ia TIDAK pernah keluar melalui getAllData/getStaffCert (ditapis). Hanya melalui PIN admin:
 *    getPrivateStaff, getIcImages, updateStaffProfile (Senarai Kerja), logJob.
 * 6. Arkib padam dan log kerja disimpan dalam spreadsheet PRIVATE (dicipta oleh setupSystem).
 *
 * PIN: nilai lalai di bawah ("admin123"). Boleh ditimpa dengan Script Property ADMIN_PIN
 * (Project Settings > Script Properties) tanpa ubah kod.
 */

// 1. TETAPAN
const ADMIN_PIN = "admin123"; // PIN lalai. Tukar di sini, atau guna Script Property ADMIN_PIN.
const DRIVE_UPLOAD_FOLDER_NAME = "UniTrack_Uploaded_Documents";

const TAB_GIDS = {
  kereta: 1406036794,
  motosikal: 906657935,
  chargeman: 1914957719,
  cert: 2314958
};

const DEFAULT_HEADERS = {
  kereta: ["No", "Plate No.", "Location", "Condition", "Expired Date", "Document"],
  motosikal: ["No", "Plate No.", "Location", "Condition", "Expired Date", "Document"],
  chargeman: ["No", "Solar Farm", "Chargeman Name", "Competency No.", "Expiry Date", "Document"],
  cert: ["ID", "NAME", "CIDB", "CIDB Cert", "NIOSH", "Niosh Cert", "Working at Height", "WaH Cert", "First Aid", "First Aid Cert"]
};

// Lajur yang boleh diubah oleh staf sendiri (portal kendiri, tanpa PIN).
const STAFF_CERT_COLUMNS = {
  dates: ["CIDB", "NIOSH", "Working at Height", "First Aid"],
  files: ["CIDB Cert", "Niosh Cert", "WaH Cert", "First Aid Cert"]
};

const MAX_UPLOAD_BYTES = 8 * 1024 * 1024;
const ALLOWED_UPLOAD_MIME = /^(application\/pdf|image\/(png|jpeg|jpg|webp|gif|heic|heif))$/i;

/**
 * ============================================================================
 * ENDPOINT GET
 * ============================================================================
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getAllData";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "getAllData") {
      const resultData = {
        kereta: readSheetByGid(ss, TAB_GIDS.kereta, "kereta"),
        motosikal: readSheetByGid(ss, TAB_GIDS.motosikal, "motosikal"),
        chargeman: readSheetByGid(ss, TAB_GIDS.chargeman, "chargeman"),
        cert: readSheetByGid(ss, TAB_GIDS.cert, "cert")
      };

      return createJsonResponse({
        status: "success",
        timestamp: new Date().toISOString(),
        data: resultData
      });
    }

    if (action === "getStaffCert") {
      const staffId = (e && e.parameter && e.parameter.staffId) ? String(e.parameter.staffId).trim().toUpperCase() : "";
      const certRecords = readSheetByGid(ss, TAB_GIDS.cert, "cert");
      const found = certRecords.find(function(r) {
        const idVal = String(getMatchingPropertyValue(r, "ID") || "").trim().toUpperCase();
        return idVal === staffId;
      });

      if (!found) {
        return createJsonResponse({ status: "error", message: "Staff ID tidak dijumpai." });
      }
      return createJsonResponse({ status: "success", data: found });
    }

    return createJsonResponse({ status: "error", message: "Action GET tidak dikenali." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * ============================================================================
 * ENDPOINT POST
 * ============================================================================
 * - staffUpdateCert: tanpa PIN, tetapi lajur dan jenis fail dihadkan.
 * - getPrivateStaff / getIcImages: PIN admin.
 * - addRow / updateRow / deleteRow: PIN admin.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Tiada data dihantar dalam permintaan POST.");
    }

    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // LALUAN 1: staf kemas kini sijil sendiri
    if (action === "staffUpdateCert") {
      return handleStaffUpdateCert(ss, payload);
    }

    // LALUAN 2: semua yang lain wajib PIN admin
    if (!adminPinOk_(payload.adminPin)) {
      return createJsonResponse({
        status: "error",
        message: "Akses Ditolak: Kata Laluan / PIN Admin tidak sah!"
      });
    }

    // Data sulit staf (halaman Senarai Kerja)
    if (action === "getPrivateStaff") {
      return createJsonResponse({ status: "success", data: staffPrivateList_() });
    }
    if (action === "getIcImages") {
      return createJsonResponse({ status: "success", data: staffIcImages_(payload.staffIds) });
    }
    if (action === "updateStaffProfile") {
      return createJsonResponse({ status: "success", data: updateStaffProfile_(payload) });
    }
    if (action === "logJob") {
      return createJsonResponse({ status: "success", data: logJob_(payload) });
    }

    // Operasi baris (dashboard admin)
    const category = payload.category;
    const gid = Number(payload.gid || TAB_GIDS[category]);
    const gidDibenarkan = Object.keys(TAB_GIDS).some(function(k) { return TAB_GIDS[k] === gid; });
    if (!gidDibenarkan) {
      throw new Error("Tab dengan GID " + gid + " tidak dibenarkan atau tidak dijumpai.");
    }
    const sheet = getSheetByGid(ss, gid);
    if (!sheet) {
      throw new Error("Tab helaian dengan GID " + gid + " tidak dijumpai dalam Google Sheet ini.");
    }

    // Muat naik dokumen (jika ada)
    let uploadedFileUrl = "";
    if (payload.uploadFile && payload.uploadFile.base64 && payload.uploadFile.fileName) {
      uploadedFileUrl = saveBase64FileToDrive(
        payload.uploadFile.base64,
        payload.uploadFile.fileName,
        payload.uploadFile.mimeType,
        String(category || "DOC").toUpperCase()
      );
    }

    // Muat naik berbilang fail sijil (borang Cert dashboard): { "CIDB Cert": {fileName, mimeType, base64}, ... }
    // Semua fail disahkan DAHULU; tiada apa dimuat naik jika satu sahaja tidak sah.
    const uploadedUrls = {};
    if (payload.uploadFiles && typeof payload.uploadFiles === "object" && Object.keys(payload.uploadFiles).length) {
      if (gid !== TAB_GIDS.cert) throw new Error("Muat naik berbilang fail hanya untuk tab Cert.");
      const pending = {};
      Object.keys(payload.uploadFiles).forEach(function(k) {
        const col = canonicalKey_(k, STAFF_CERT_COLUMNS.files);
        if (!col) throw new Error("Lajur fail '" + k + "' tidak dibenarkan.");
        const f = payload.uploadFiles[k];
        if (f) { validateUpload_(f); pending[col] = f; }
      });
      const idLabel = String((payload.rowData && payload.rowData.ID) || "CERT");
      Object.keys(pending).forEach(function(col) {
        const f = pending[col];
        uploadedUrls[col] = saveBase64FileToDrive(f.base64, f.fileName, f.mimeType, idLabel + "_" + col.replace(/\s+/g, "_"));
      });
    }

    // Buang lajur IC jika ada
    const cleanRowData = stripNricFields(payload.rowData || {});
    if (uploadedFileUrl) {
      const targetCol = payload.targetFileColumn || "Document";
      cleanRowData[targetCol] = uploadedFileUrl;
    }
    Object.keys(uploadedUrls).forEach(function(col) { cleanRowData[col] = uploadedUrls[col]; });

    if (action === "addRow") {
      appendRecordToSheet(sheet, category, cleanRowData);
      return createJsonResponse({
        status: "success",
        message: "Rekod baharu berjaya ditambah ke Google Sheet.",
        uploadedFileUrl: uploadedFileUrl,
        uploadedUrls: uploadedUrls
      });
    }

    if (action === "updateRow") {
      const rowIndex = Number(payload.rowIndex);
      updateRecordInSheet(sheet, category, rowIndex, cleanRowData);
      return createJsonResponse({
        status: "success",
        message: "Rekod berjaya dikemas kini di dalam Google Sheet.",
        uploadedFileUrl: uploadedFileUrl,
        uploadedUrls: uploadedUrls
      });
    }

    if (action === "deleteRow") {
      const rowIndex = Number(payload.rowIndex);
      deleteRecordFromSheet(sheet, category, rowIndex);
      return createJsonResponse({
        status: "success",
        message: "Rekod telah dipadamkan (salinan disimpan dalam arkib Deleted_Rows)."
      });
    }

    return createJsonResponse({ status: "error", message: "Action POST tidak sah." });
  } catch (err) {
    return createJsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * KEMAS KINI KENDIRI STAF (TANPA PIN) — DIHADKAN
 * ============================================================================
 * Hanya 4 lajur tarikh dan 4 lajur fail sijil. Tarikh mesti DD/MM/YYYY (atau kosong).
 * Hanya sel yang berubah ditulis. Tarikh kosong TIDAK menimpa tarikh sedia ada.
 */
function handleStaffUpdateCert(ss, payload) {
  const staffId = String(payload.staffId || "").trim().toUpperCase();
  if (!staffId) {
    return createJsonResponse({ status: "error", message: "Staff ID diperlukan untuk mengemas kini sijil." });
  }

  const certSheet = getSheetByGid(ss, TAB_GIDS.cert);
  if (!certSheet) {
    throw new Error("Tab Cert (GID: " + TAB_GIDS.cert + ") tidak dijumpai.");
  }

  const updatesIn = payload.updates || {};
  const uploadsIn = payload.uploads || {};
  const updates = {};
  const uploads = {};

  // Tapis dan sahkan input SEBELUM sebarang tulisan
  Object.keys(updatesIn).forEach(function(k) {
    const col = canonicalKey_(k, STAFF_CERT_COLUMNS.dates);
    if (!col) throw new Error("Lajur '" + k + "' tidak dibenarkan untuk kemas kini kendiri.");
    const v = String(updatesIn[k] === null || updatesIn[k] === undefined ? "" : updatesIn[k]).trim();
    if (v !== "" && !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v) && !/^\d{4}-\d{2}-\d{2}$/.test(v)) {
      throw new Error("Format tarikh tidak sah untuk " + col + " (guna DD/MM/YYYY).");
    }
    updates[col] = v;
  });
  Object.keys(uploadsIn).forEach(function(k) {
    const col = canonicalKey_(k, STAFF_CERT_COLUMNS.files);
    if (!col) throw new Error("Lajur fail '" + k + "' tidak dibenarkan untuk kemas kini kendiri.");
    if (uploadsIn[k]) {
      validateUpload_(uploadsIn[k]);
      uploads[col] = uploadsIn[k];
    }
  });

  const headerInfo = locateHeaderRow(certSheet, "cert");
  const wanted = {};
  Object.keys(updates).forEach(function(k) { wanted[k] = true; });
  Object.keys(uploads).forEach(function(k) { wanted[k] = true; });
  const headers = ensureColumnsExist(certSheet, headerInfo, wanted);

  let idColIndex = headers.findIndex(function(h) { return String(h).toLowerCase().trim() === "id"; });
  if (idColIndex === -1) idColIndex = 0;

  const lastRow = certSheet.getLastRow();
  const lastCol = Math.max(certSheet.getLastColumn(), headers.length);
  const allValues = certSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  let targetRow = -1;
  let currentRow = [];
  for (let r = headerInfo.headerRowNumber; r < allValues.length; r++) {
    if (String(allValues[r][idColIndex] || "").trim().toUpperCase() === staffId) {
      targetRow = r + 1;
      currentRow = allValues[r];
      break;
    }
  }
  if (targetRow === -1) {
    return createJsonResponse({
      status: "error",
      message: "Rekod Staff ID '" + staffId + "' tidak dijumpai di dalam Lajur A Tab Cert."
    });
  }

  // Muat naik fail dahulu (sudah disahkan)
  const uploadedUrls = {};
  Object.keys(uploads).forEach(function(col) {
    const f = uploads[col];
    uploadedUrls[col] = saveBase64FileToDrive(
      f.base64, f.fileName, f.mimeType, staffId + "_" + col.replace(/\s+/g, "_")
    );
  });

  // Tulis hanya sel yang berubah
  const changed = [];
  Object.keys(uploadedUrls).forEach(function(col) {
    const idx = headerIndex_(headers, col);
    certSheet.getRange(targetRow, idx + 1).setValue(uploadedUrls[col]);
    changed.push(col);
  });
  Object.keys(updates).forEach(function(col) {
    const idx = headerIndex_(headers, col);
    const sedia = String(currentRow[idx] === undefined ? "" : currentRow[idx]).trim();
    const baru = updates[col];
    if (baru === "" && sedia !== "") return;   // jangan padam tarikh sedia ada
    if (baru === sedia) return;
    certSheet.getRange(targetRow, idx + 1).setValue(baru);
    changed.push(col);
  });

  return createJsonResponse({
    status: "success",
    message: "Maklumat sijil bagi Staff ID " + staffId + " berjaya dikemas kini.",
    staffId: staffId,
    changed: changed,
    uploadedUrls: uploadedUrls
  });
}

function canonicalKey_(key, list) {
  const t = String(key).toLowerCase().trim();
  for (let i = 0; i < list.length; i++) {
    if (list[i].toLowerCase() === t) return list[i];
  }
  return "";
}

function headerIndex_(headers, name) {
  const t = String(name).toLowerCase().trim();
  return headers.findIndex(function(h) { return String(h).toLowerCase().trim() === t; });
}

/**
 * Sahkan fail muat naik: ada data, jenis dibenarkan, saiz dihadkan.
 */
function validateUpload_(f) {
  if (!f || !f.base64 || !f.fileName) throw new Error("Data fail tidak lengkap.");
  const mime = String(f.mimeType || "application/pdf");
  if (!ALLOWED_UPLOAD_MIME.test(mime)) throw new Error("Jenis fail tidak dibenarkan (" + mime + "). Guna PDF atau gambar.");
  // Anggaran saiz selepas decode: panjang base64 * 3/4
  const anggaran = Math.floor(String(f.base64).length * 3 / 4);
  if (anggaran > MAX_UPLOAD_BYTES) throw new Error("Fail terlalu besar (had " + Math.round(MAX_UPLOAD_BYTES / 1048576) + "MB).");
}

/**
 * ============================================================================
 * PENGESAHAN PIN ADMIN
 * ============================================================================
 * Baca Script Property ADMIN_PIN dahulu, jatuh balik ke const ADMIN_PIN.
 * 10 percubaan salah = dikunci 15 minit (kaunter global, kerana IP tidak tersedia).
 */
function adminPinOk_(pin) {
  const cache = CacheService.getScriptCache();
  const gagal = Number(cache.get("PIN_GAGAL") || 0);
  if (gagal >= 10) throw new Error("Terlalu banyak percubaan PIN salah. Cuba semula selepas 15 minit.");
  const sah = String(PropertiesService.getScriptProperties().getProperty("ADMIN_PIN") || ADMIN_PIN || "");
  const diberi = String(pin === undefined || pin === null ? "" : pin).trim();
  if (!diberi || !sah || diberi !== sah) {
    cache.put("PIN_GAGAL", String(gagal + 1), 900);
    return false;
  }
  cache.remove("PIN_GAGAL");
  return true;
}

/**
 * ============================================================================
 * DATA SULIT STAF (tab Cert: lajur IC, IC Attachment, Designation)
 * ============================================================================
 * Dibaca TANPA tapisan (hanya selepas PIN sah). File ID gambar IC hanya diambil dari sheet,
 * tidak pernah dari browser.
 */
const IC_UPLOAD_FOLDER_NAME = "UniTrack_IC_Private"; // fail IC kekal peribadi (tiada perkongsian pautan)
const IC_MAX_BYTES = 5 * 1024 * 1024;
const STAFF_MAX_IDS = 20;
const STAFF_MAX_BYTES = 3 * 1024 * 1024; // gambar IC lebih besar daripada ini -> guna thumbnail

/** Semua baris tab Cert (tanpa tapis IC) bersama nombor baris sebenar. */
function certRawRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByGid(ss, TAB_GIDS.cert);
  if (!sheet) throw new Error("Tab Cert tidak dijumpai.");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headerInfo = locateHeaderRow(sheet, "cert");
  const vals = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = vals[headerInfo.headerRowNumber - 1].map(function(h) { return String(h).trim(); });
  const out = [];
  for (let r = headerInfo.headerRowNumber; r < vals.length; r++) {
    const o = {};
    headers.forEach(function(h, c) { if (h) o[h] = String(vals[r][c] === undefined ? "" : vals[r][c]).trim(); });
    if (String(getMatchingPropertyValue(o, "ID") || "").trim()) out.push({ row: r + 1, data: o });
  }
  return out;
}

function staffFileId_(url) {
  const s = String(url || "");
  const m = s.match(/\/d\/([-\w]{20,})/) || s.match(/[?&]id=([-\w]{20,})/);
  return m ? m[1] : "";
}

/** Senarai staf lengkap (termasuk no. IC) untuk admin. Tiada gambar. */
function staffPrivateList_() {
  return certRawRows_().map(function(x) {
    const d = x.data;
    return {
      id: String(getMatchingPropertyValue(d, "ID")).trim().toUpperCase(),
      name: String(getMatchingPropertyValue(d, "NAME") || "").trim(),
      ic: String(getMatchingPropertyValue(d, "IC") || "").trim(),
      designation: String(getMatchingPropertyValue(d, "Designation") || "").trim(),
      aktif: true,
      hasIcImage: !!staffFileId_(getMatchingPropertyValue(d, "IC Attachment"))
    };
  });
}

/** Gambar IC bagi senarai ID. Hasil: { VES007: {mime, base64, lowres} | {error} } */
function staffIcImages_(ids) {
  const senarai = [];
  (ids || []).forEach(function(x) {
    const id = String(x || "").trim().toUpperCase();
    if (id && senarai.indexOf(id) === -1) senarai.push(id);
  });
  if (senarai.length > STAFF_MAX_IDS) throw new Error("Maksimum " + STAFF_MAX_IDS + " ID setiap permintaan.");

  const peta = {};
  certRawRows_().forEach(function(x) {
    peta[String(getMatchingPropertyValue(x.data, "ID")).trim().toUpperCase()] = staffFileId_(getMatchingPropertyValue(x.data, "IC Attachment"));
  });

  const out = {};
  senarai.forEach(function(id) {
    const fid = peta[id];
    if (!fid) { out[id] = { error: "TIADA_LINK" }; return; }
    try {
      const file = DriveApp.getFileById(fid);
      let blob = file.getBlob();
      let mime = String(blob.getContentType() || "");
      let bytes = blob.getBytes();
      let lowres = false;
      if (mime.indexOf("image/") !== 0 || bytes.length > STAFF_MAX_BYTES) {
        const thumb = file.getThumbnail(); // PDF atau fail besar: thumbnail resolusi rendah
        if (!thumb) { out[id] = { error: "TIADA_PRATONTON" }; return; }
        blob = thumb;
        mime = String(blob.getContentType() || "image/png");
        bytes = blob.getBytes();
        lowres = true;
      }
      out[id] = { mime: mime, base64: Utilities.base64Encode(bytes), lowres: lowres };
    } catch (err) {
      out[id] = { error: "GAGAL_BACA" };
    }
  });
  return out;
}

/**
 * Admin kemas kini profil staf: Designation, no. IC, salinan IC (fail disimpan peribadi dalam Drive).
 * Hanya medan yang dihantar diubah. Semua input disahkan SEBELUM sebarang tulisan.
 */
function updateStaffProfile_(p) {
  const staffId = String(p.staffId || "").trim().toUpperCase();
  if (!staffId) throw new Error("Staff ID diperlukan.");

  const wanted = {};
  let designation, ic;
  if (p.designation !== undefined && p.designation !== null) {
    designation = String(p.designation).trim();
    if (designation.length > 80) throw new Error("Designation terlalu panjang (maksimum 80 aksara).");
    wanted["Designation"] = true;
  }
  if (p.ic !== undefined && p.ic !== null) {
    ic = String(p.ic).trim();
    if (ic !== "" && !/^[\d-]{6,20}$/.test(ic)) throw new Error("Format no. IC tidak sah (nombor dan sengkang sahaja).");
    wanted["IC"] = true;
  }
  if (p.icFile) {
    validateIcUpload_(p.icFile);
    wanted["IC Attachment"] = true;
  }
  if (!Object.keys(wanted).length) throw new Error("Tiada medan untuk dikemas kini.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByGid(ss, TAB_GIDS.cert);
  if (!sheet) throw new Error("Tab Cert tidak dijumpai.");

  // Cari baris dahulu, sebelum ubah header atau muat naik apa-apa
  const found = certRawRows_().filter(function(x) {
    return String(getMatchingPropertyValue(x.data, "ID")).trim().toUpperCase() === staffId;
  })[0];
  if (!found) throw new Error("Staff ID '" + staffId + "' tidak dijumpai dalam tab Cert.");

  const headerInfo = locateHeaderRow(sheet, "cert");
  const headers = ensureColumnsExist(sheet, headerInfo, wanted);
  const changed = [];

  if (designation !== undefined) {
    sheet.getRange(found.row, headerIndex_(headers, "Designation") + 1).setValue(sheetSafeValue_(designation));
    changed.push("Designation");
  }
  if (ic !== undefined) {
    const cell = sheet.getRange(found.row, headerIndex_(headers, "IC") + 1);
    cell.setNumberFormat("@");   // kekalkan sifar di depan
    cell.setValue(ic);
    changed.push("IC");
  }
  if (p.icFile) {
    const url = saveIcFileToDrive_(p.icFile, staffId);
    sheet.getRange(found.row, headerIndex_(headers, "IC Attachment") + 1).setValue(url);
    changed.push("IC Attachment");
  }
  return { staffId: staffId, changed: changed };
}

function validateIcUpload_(f) {
  if (!f || !f.base64 || !f.fileName) throw new Error("Data fail IC tidak lengkap.");
  const mime = String(f.mimeType || "");
  if (!/^(application\/pdf|image\/(png|jpeg|jpg|webp|heic|heif))$/i.test(mime)) {
    throw new Error("Jenis fail IC tidak dibenarkan (" + mime + "). Guna JPG, PNG atau PDF.");
  }
  if (Math.floor(String(f.base64).length * 3 / 4) > IC_MAX_BYTES) {
    throw new Error("Fail IC terlalu besar (had " + Math.round(IC_MAX_BYTES / 1048576) + "MB).");
  }
}

/** Fail IC disimpan dalam folder peribadi. TIADA setSharing: hanya pemilik skrip boleh buka. */
function saveIcFileToDrive_(f, staffId) {
  let folder;
  const folders = DriveApp.getFoldersByName(IC_UPLOAD_FOLDER_NAME);
  folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(IC_UPLOAD_FOLDER_NAME);
  const bytes = Utilities.base64Decode(f.base64);
  if (bytes.length > IC_MAX_BYTES) throw new Error("Fail IC terlalu besar.");
  const selamat = String(f.fileName).replace(/[^\w.\- ]/g, "_").slice(0, 80);
  const nama = String(staffId).replace(/[^\w\-]/g, "_") + "_IC_" + new Date().getTime() + "_" + selamat;
  const file = folder.createFile(Utilities.newBlob(bytes, f.mimeType, nama));
  return file.getUrl();
}

/** Rekod listing yang dijana ke tab Job_Log (spreadsheet private). Tidak kritikal: gagal = tidak menjejas. */
function logJob_(p) {
  const pid = PropertiesService.getScriptProperties().getProperty("PRIVATE_SS_ID");
  if (!pid) return { logged: false, sebab: "Spreadsheet private belum disediakan" };
  const ids = (Array.isArray(p.staffIds) ? p.staffIds : []).slice(0, 200)
    .map(function(x) { return String(x).trim().toUpperCase().replace(/[^A-Z0-9_\-]/g, ""); })
    .filter(function(x) { return x; });
  const pss = SpreadsheetApp.openById(pid);
  let tab = pss.getSheetByName("Job_Log");
  if (!tab) {
    tab = pss.insertSheet("Job_Log");
    tab.getRange(1, 1, 1, 6).setValues([["Tarikh", "Projek", "Dibuat Oleh", "Senarai ID", "Bil Orang", "PDF Link"]]);
  }
  tab.appendRow([new Date(), sheetSafeValue_(String(p.title || "").slice(0, 160)), "Admin (PIN)", ids.join(", "), ids.length, ""]);
  return { logged: true };
}

/**
 * ============================================================================
 * BANTUAN SHEET
 * ============================================================================
 */
function getSheetByGid(spreadsheet, gid) {
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === Number(gid)) {
      return sheets[i];
    }
  }
  return null;
}

function locateHeaderRow(sheet, category) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    const defaultHeaders = DEFAULT_HEADERS[category] || ["No"];
    sheet.appendRow(defaultHeaders);
    return { headerRowNumber: 1, headers: defaultHeaders };
  }

  const scanRows = Math.min(lastRow, 5);
  const values = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    const joined = values[r].join(" ").toLowerCase();
    if (
      joined.indexOf("plate") !== -1 ||
      joined.indexOf("expired") !== -1 ||
      joined.indexOf("expiry") !== -1 ||
      joined.indexOf("chargeman") !== -1 ||
      joined.indexOf("cidb") !== -1 ||
      joined.indexOf("name") !== -1
    ) {
      return {
        headerRowNumber: r + 1,
        headers: values[r].map(function(h) { return String(h).trim(); })
      };
    }
  }

  return {
    headerRowNumber: 1,
    headers: values[0].map(function(h) { return String(h).trim(); })
  };
}

function ensureColumnsExist(sheet, headerInfo, rowDataObj) {
  const headers = headerInfo.headers.slice();
  const headerRowNumber = headerInfo.headerRowNumber;
  const incomingKeys = Object.keys(rowDataObj);

  incomingKeys.forEach(function(key) {
    if (!key) return;
    const foundIndex = headers.findIndex(function(h) {
      return h.toLowerCase().trim() === key.toLowerCase().trim();
    });

    if (foundIndex === -1) {
      headers.push(key);
      sheet.getRange(headerRowNumber, headers.length).setValue(sheetSafeValue_(key));
    }
  });

  return headers;
}

/**
 * Elak formula injection: nilai yang bermula dengan = + - @ (dan bukan nombor biasa)
 * ditulis sebagai teks. Nombor, tarikh dan URL tidak terjejas.
 */
function sheetSafeValue_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v !== "string") return v;
  if (/^[=+\-@]/.test(v) && !/^[-+]?[\d\s().,\/-]+$/.test(v)) return "'" + v;
  return v;
}

function readSheetByGid(spreadsheet, gid, category) {
  const sheet = getSheetByGid(spreadsheet, gid);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];

  const headerInfo = locateHeaderRow(sheet, category);
  const headerRowIdx = headerInfo.headerRowNumber - 1;
  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = allValues[headerRowIdx];

  const records = [];
  for (let r = headerRowIdx + 1; r < allValues.length; r++) {
    const rowCells = allValues[r];
    const hasContent = rowCells.some(function(c) { return String(c).trim() !== ""; });
    if (!hasContent) continue;

    const rawObj = {};
    for (let c = 0; c < headers.length; c++) {
      const colName = String(headers[c] || ("Column_" + (c + 1))).trim();
      rawObj[colName] = rowCells[c] !== undefined ? String(rowCells[c]).trim() : "";
    }

    records.push(stripNricFields(rawObj));
  }

  return records;
}

function appendRecordToSheet(sheet, category, rowDataObj) {
  const headerInfo = locateHeaderRow(sheet, category);
  const updatedHeaders = ensureColumnsExist(sheet, headerInfo, rowDataObj);

  const newRowArray = updatedHeaders.map(function(colName) {
    return sheetSafeValue_(getMatchingPropertyValue(rowDataObj, colName));
  });

  sheet.appendRow(newRowArray);
}

/**
 * Kemas kini: hanya sel yang dihantar ditulis. Sel lain, format dan formula tidak disentuh.
 */
function updateRecordInSheet(sheet, category, dataRowIndex, rowDataObj) {
  const headerInfo = locateHeaderRow(sheet, category);
  const updatedHeaders = ensureColumnsExist(sheet, headerInfo, rowDataObj);
  const rowNumber = findActualSheetRowNumber(sheet, headerInfo.headerRowNumber, dataRowIndex);

  updatedHeaders.forEach(function(colName, idx) {
    if (hasMatchingProperty(rowDataObj, colName)) {
      sheet.getRange(rowNumber, idx + 1).setValue(sheetSafeValue_(getMatchingPropertyValue(rowDataObj, colName)));
    }
  });
}

/**
 * Padam: baris disalin ke tab Deleted_Rows (spreadsheet private) DAHULU.
 * Jika arkib gagal, baris TIDAK dipadam.
 */
function deleteRecordFromSheet(sheet, category, dataRowIndex) {
  const headerInfo = locateHeaderRow(sheet, category);
  const rowNumber = findActualSheetRowNumber(sheet, headerInfo.headerRowNumber, dataRowIndex);
  archiveRowBeforeDelete_(category, sheet, headerInfo, rowNumber);
  sheet.deleteRow(rowNumber);
}

function archiveRowBeforeDelete_(category, sheet, headerInfo, rowNumber) {
  const pid = PropertiesService.getScriptProperties().getProperty("PRIVATE_SS_ID");
  if (!pid) {
    throw new Error("Arkib padam belum disediakan. Run setupSystem() dalam Apps Script dahulu, kemudian cuba lagi. Tiada apa dipadam.");
  }
  const pss = SpreadsheetApp.openById(pid);
  let tab = pss.getSheetByName("Deleted_Rows");
  if (!tab) {
    tab = pss.insertSheet("Deleted_Rows");
    tab.getRange(1, 1, 1, 5).setValues([["Tarikh", "Kategori", "Tab Asal", "No Baris", "Data (JSON)"]]);
  }
  const hdr = headerInfo.headers;
  const vals = sheet.getRange(rowNumber, 1, 1, hdr.length).getDisplayValues()[0];
  const obj = {};
  hdr.forEach(function(h, i) { if (h) obj[h] = vals[i]; });
  tab.appendRow([new Date(), String(category || ""), sheet.getName(), rowNumber, JSON.stringify(obj)]);
}

function findActualSheetRowNumber(sheet, headerRowNumber, targetDataIndex) {
  if (!isFinite(targetDataIndex) || targetDataIndex < 0) {
    throw new Error("Indeks baris tidak sah.");
  }
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  let currentValidIndex = 0;
  for (let r = headerRowNumber; r < allValues.length; r++) {
    const rowCells = allValues[r];
    const hasContent = rowCells.some(function(c) { return String(c).trim() !== ""; });
    if (!hasContent) continue;

    if (currentValidIndex === Number(targetDataIndex)) {
      return r + 1;
    }
    currentValidIndex++;
  }

  throw new Error("Baris sasaran untuk dikemas kini/dipadam tidak dijumpai.");
}

/**
 * ============================================================================
 * MUAT NAIK KE GOOGLE DRIVE
 * ============================================================================
 */
function saveBase64FileToDrive(base64Data, fileName, mimeType, prefixLabel) {
  validateUpload_({ base64: base64Data, fileName: fileName, mimeType: mimeType });

  let folder;
  const folders = DriveApp.getFoldersByName(DRIVE_UPLOAD_FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_UPLOAD_FOLDER_NAME);
  }

  const decodedBytes = Utilities.base64Decode(base64Data);
  if (decodedBytes.length > MAX_UPLOAD_BYTES) throw new Error("Fail terlalu besar.");
  const selamat = String(fileName).replace(/[^\w.\- ]/g, "_").slice(0, 80);
  const safeName = String(prefixLabel).replace(/[^\w\-]/g, "_") + "_" + new Date().getTime() + "_" + selamat;
  const blob = Utilities.newBlob(decodedBytes, mimeType || "application/pdf", safeName);
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Polisi organisasi mungkin menghadkan perkongsian awam
  }

  return file.getUrl();
}

/**
 * ============================================================================
 * PEMATUHAN PDPA: buang lajur IC / NRIC / No. KP / MyKad dan lampirannya
 * ============================================================================
 */
function staffIsSensitiveKey_(k) {
  return /(^ic($|[\s_\-.:])|nric|no\.?\s*kp|kad\s*pengenalan|mykad)/i.test(String(k).trim());
}

function stripNricFields(obj) {
  if (!obj || typeof obj !== "object") return {};
  const cleaned = {};
  Object.keys(obj).forEach(function(k) {
    if (!staffIsSensitiveKey_(k)) {
      cleaned[k] = obj[k];
    }
  });
  return cleaned;
}

function hasMatchingProperty(obj, targetKey) {
  const keys = Object.keys(obj || {});
  const t = String(targetKey).toLowerCase().trim();
  return keys.some(function(k) { return k.toLowerCase().trim() === t; });
}

function getMatchingPropertyValue(obj, targetKey) {
  const keys = Object.keys(obj || {});
  const t = String(targetKey).toLowerCase().trim();
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase().trim() === t) {
      return obj[keys[i]];
    }
  }
  return "";
}

function createJsonResponse(payloadObj) {
  return ContentService
    .createTextOutput(JSON.stringify(payloadObj))
    .setMimeType(ContentService.MimeType.JSON);
}
