/**
 * SYSTEM IKR - PT. BELACOM TEKNOLOGI INDONESIA
 * File: Code.gs
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BELACOM IKR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: Konversi Drive URL biasa menjadi Direct Stream Link
function convertToDirectLink(driveUrl) {
  if (!driveUrl || driveUrl === '-' || driveUrl.trim() === '') return '';
  var match = driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || driveUrl.match(/id=([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return 'https://lh3.googleusercontent.com/d/' + match[1];
  }
  return driveUrl;
}

// Helper Global Filter Checker untuk Baris WO
function applyFilters_(row, startDate, endDate, statusFilter, teamFilter) {
  let team = row[5] ? String(row[5]) : "Tanpa Team";
  let tglVisit = row[7] ? new Date(row[7]) : null;
  let status = row[8] ? String(row[8]) : "On Schedule";

  // 1. Filter Periode (Berdasarkan Tgl Visit)
  if (startDate && tglVisit && tglVisit < startDate) return false;
  if (endDate && tglVisit) {
    let endOfDay = new Date(endDate);
    endOfDay.setHours(23, 59, 59, 999);
    if (tglVisit > endOfDay) return false;
  }

  // 2. Filter Status WO
  if (statusFilter && status !== statusFilter) return false;

  // 3. Filter Team Aktivator
  if (teamFilter && team !== teamFilter) return false;

  return true;
}

// --- SISTEM LOGIN & MANAJEMEN USER ---
function checkUserLogin(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Username", "Password", "Role", "Nama Lengkap"]);
    sheet.appendRow(["admin", "admin123", "Admin", "Administrator Utama"]);
    sheet.appendRow(["leader", "leader123", "Leader", "Leader Operasional"]);
    sheet.appendRow(["team", "team123", "Team IKR", "Teknisi Lapangan"]);
  }

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    let u = String(rows[i][0]).trim();
    let p = String(rows[i][1]).trim();
    let role = String(rows[i][2]).trim();
    let nama = String(rows[i][3] || u).trim();

    if (u.toLowerCase() === String(username).toLowerCase() && p === password) {
      return { success: true, username: u, role: role, nama: nama };
    }
  }
  return { success: false, message: "Username atau password salah!" };
}

function getUsersList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map((r, idx) => ({
    rowIndex: idx + 2,
    username: r[0] ? String(r[0]) : "",
    password: r[1] ? String(r[1]) : "",
    role: r[2] ? String(r[2]) : "",
    nama: r[3] ? String(r[3]) : ""
  }));
}

function saveUser(userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Users");
  if (!sheet) {
    sheet = ss.insertSheet("Users");
    sheet.appendRow(["Username", "Password", "Role", "Nama Lengkap"]);
  }

  const rows = sheet.getDataRange().getValues();
  if (userData.rowIndex) {
    sheet.getRange(parseInt(userData.rowIndex), 1, 1, 4).setValues([[
      userData.username, userData.password, userData.role, userData.nama
    ]]);
  } else {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).toLowerCase() === String(userData.username).toLowerCase()) {
        return { success: false, message: "Username sudah terdaftar!" };
      }
    }
    sheet.appendRow([userData.username, userData.password, userData.role, userData.nama]);
  }
  return { success: true };
}

function deleteUser(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan" };
  sheet.deleteRow(parseInt(rowIndex));
  return { success: true };
}

function updatePassword(username, oldPassword, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  if (!sheet) return { success: false, message: "Sheet Users tidak ditemukan" };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      if (String(rows[i][1]).trim() !== String(oldPassword).trim()) {
        return { success: false, message: "Password lama tidak sesuai!" };
      }
      sheet.getRange(i + 1, 2).setValue(newPassword);
      return { success: true };
    }
  }
  return { success: false, message: "User tidak ditemukan" };
}

// 1. Dashboard Data WO (Dengan Filter)
function getDashboardData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  
  const defaultEmpty = {
    totalWO: 0,
    statusCounts: { "On Schedule": 0, "Done": 0, "On Progress": 0, "Reschedule": 0, "Pending": 0, "Hold To Sales": 0, "Hold To Planning": 0, "Hold To Maintenance": 0 },
    teamPerf: {},
    paymentCounts: { Paid: 0, Unpaid: 0 }
  };

  if (!sheetWO) return defaultEmpty;

  const data = sheetWO.getDataRange().getValues();
  if (data.length <= 1) return defaultEmpty;

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";

  let statusCounts = {
    "On Schedule": 0, "Done": 0, "On Progress": 0, "Reschedule": 0,
    "Pending": 0, "Hold To Sales": 0, "Hold To Planning": 0, "Hold To Maintenance": 0
  };
  let teamPerf = {};
  let paymentCounts = { "Paid": 0, "Unpaid": 0 };
  let totalFilteredWO = 0;

  for (let i = 1; i < data.length; i++) {
    let r = data[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    let team = r[5] ? String(r[5]) : "Tanpa Team";
    let status = r[8] ? String(r[8]) : "On Schedule";
    let payStatus = r[10] ? String(r[10]) : "Unpaid";

    totalFilteredWO++;

    if (statusCounts[status] !== undefined) statusCounts[status]++;
    else statusCounts[status] = 1;

    if (!teamPerf[team]) teamPerf[team] = { total: 0, done: 0 };
    teamPerf[team].total++;
    if (status === "Done") teamPerf[team].done++;

    if (payStatus === "Paid") paymentCounts.Paid++;
    else paymentCounts.Unpaid++;
  }

  return {
    totalWO: totalFilteredWO,
    statusCounts: statusCounts,
    teamPerf: teamPerf,
    paymentCounts: paymentCounts
  };
}

// Ambil list unik Team Aktivator untuk dropdown filter
function getUniqueTeamsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  if (!sheetWO) return [];
  const rows = sheetWO.getDataRange().getValues();
  let teams = new Set();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][5]) teams.add(String(rows[i][5]));
  }
  return Array.from(teams).sort();
}

// 2. Data Dashboard Tagihan & Analisis Pembayaran (Dengan Filter) - harga per WO via tanggal efektif
function getBillingDashboardData(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");

  let summary = {
    totalWO: 0, totalNominalTagihan: 0, totalPembayaranTeam: 0, totalRevenueBelacom: 0,
    paidCount: 0, paidNominal: 0, unpaidCount: 0, unpaidNominal: 0
  };
  let teamBilling = {};

  if (!sheetWO) return { summary: summary, teamBilling: [] };
  const rows = sheetWO.getDataRange().getValues();
  if (rows.length <= 1) return { summary: summary, teamBilling: [] };

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";
  const priceTable = loadPriceTable_();

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    const price = priceForWoRow_(priceTable, r);
    const team = r[5] ? String(r[5]) : "Tanpa Team";
    const isPaid = (r[10] ? String(r[10]) : "Unpaid") === "Paid";

    summary.totalWO++;
    summary.totalNominalTagihan += price.tagihan;
    summary.totalPembayaranTeam += price.team;
    summary.totalRevenueBelacom += price.revenue;
    if (isPaid) { summary.paidCount++; summary.paidNominal += price.tagihan; }
    else { summary.unpaidCount++; summary.unpaidNominal += price.tagihan; }

    if (!teamBilling[team]) {
      teamBilling[team] = { teamName: team, totalWO: 0, paidWO: 0, unpaidWO: 0, totalTagihan: 0, totalTeamPay: 0, totalRevenue: 0 };
    }
    const tb = teamBilling[team];
    tb.totalWO++;
    tb.totalTagihan += price.tagihan;
    tb.totalTeamPay += price.team;
    tb.totalRevenue += price.revenue;
    if (isPaid) tb.paidWO++; else tb.unpaidWO++;
  }

  return { summary: summary, teamBilling: Object.keys(teamBilling).map(k => teamBilling[k]) };
}

// 3. Ambil Semua Work Orders & Kalkulasi Tagihan Finansial (Dengan Filter) - harga via tanggal efektif
function getWorkOrders(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Work_Orders");
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";
  const priceTable = loadPriceTable_();

  let results = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    const price = priceForWoRow_(priceTable, r);
    results.push({
      rowIndex: i + 1,
      noTiket: r[0] ? String(r[0]) : "",
      namaCustomer: r[1] ? String(r[1]) : "",
      idCustomer: r[2] ? String(r[2]) : "",
      noTelp: r[3] ? String(r[3]) : "",
      alamat: r[4] ? String(r[4]) : "",
      team: r[5] ? String(r[5]) : "",
      pic: r[6] ? String(r[6]) : "",
      noTelpPic: r[13] ? String(r[13]) : "",
      tglVisit: r[7] ? (r[7] instanceof Date ? Utilities.formatDate(r[7], "GMT+7", "yyyy-MM-dd") : String(r[7])) : "",
      status: r[8] ? String(r[8]) : "On Schedule",
      tglBayar: r[9] ? (r[9] instanceof Date ? Utilities.formatDate(r[9], "GMT+7", "yyyy-MM-dd") : String(r[9])) : "",
      statusBayar: r[10] ? String(r[10]) : "Unpaid",
      nominalTagihan: price.tagihan,
      pembayaranTeam: price.team,
      revenueBelacom: price.revenue
    });
  }
  return results;
}

// Update Status Pembayaran WO
function updatePaymentStatus(noTiket, statusBayar, tglBayar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Work_Orders");
  if (!sheet) return { success: false, message: "Sheet Work_Orders tidak ditemukan" };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(noTiket)) {
      const payDate = tglBayar ? new Date(tglBayar) : new Date();
      sheet.getRange(i + 1, 10).setValue(payDate);
      sheet.getRange(i + 1, 11).setValue(statusBayar);
      sheet.getRange(i + 1, 13).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, message: "Tiket tidak ditemukan" };
}

// 4. Ambil List Teams dari Sheet "Teams"
function getTeamsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Teams");
  if (!sheet) {
    sheet = ss.insertSheet("Teams");
    sheet.appendRow(["Nama Team", "PIC / Teknisi", "No Telp PIC"]);
    sheet.appendRow(["BLC-IKR01", "Ahmad & Budi", "081234567890"]);
    sheet.appendRow(["BLC-IKR02", "Candra & Dedi", "081298765432"]);
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).filter(r => r[0]).map(r => ({
    teamName: String(r[0]),
    picName: String(r[1] || ''),
    noTelpPic: String(r[2] || '')
  }));
}

// 5. Ambil List Report IKR (Dengan Filter)
function getReportIKRList(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  const sheetRep = ss.getSheetByName("Report_IKR");
  if (!sheetWO) return [];

  const woRows = sheetWO.getDataRange().getValues();
  if (woRows.length <= 1) return [];

  const repRows = sheetRep ? sheetRep.getDataRange().getValues() : [];

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";

  let results = [];
  for (let i = 1; i < woRows.length; i++) {
    const r = woRows[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    const noTiket = String(r[0]);
    const rep = repRows.find(rp => String(rp[0]) === noTiket);

    results.push({
      noTiket: noTiket,
      namaCustomer: r[1] ? String(r[1]) : "",
      idCustomer: r[2] ? String(r[2]) : "",
      noTelp: r[3] ? String(r[3]) : "",
      alamat: r[4] ? String(r[4]) : "",
      team: r[5] ? String(r[5]) : "",
      pic: r[6] ? String(r[6]) : "",
      noTelpPic: r[13] ? String(r[13]) : "",
      tglVisit: r[7] ? (r[7] instanceof Date ? Utilities.formatDate(r[7], "GMT+7", "yyyy-MM-dd") : String(r[7])) : "",
      status: r[8] ? String(r[8]) : "On Schedule",
      hasReport: !!rep
    });
  }

  return results;
}

// 6. Simpan / Edit Work Order
function saveWorkOrder(woData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Work_Orders");
  if (!sheet) {
    sheet = ss.insertSheet("Work_Orders");
    sheet.appendRow([
      "No Tiket", "Nama Customer", "ID Customer", "No Telp", "Alamat",
      "Team", "PIC", "Tgl Visit", "Status", "Tgl Bayar", "Status Bayar", "Created At", "Updated At", "No Telp PIC"
    ]);
  }

  const now = new Date();

  if (woData.rowIndex) {
    sheet.getRange(parseInt(woData.rowIndex), 1, 1, 9).setValues([[
      woData.noTiket, woData.namaCustomer, woData.idCustomer, woData.noTelp,
      woData.alamat, woData.team, woData.pic, woData.tglVisit, woData.status
    ]]);
    sheet.getRange(parseInt(woData.rowIndex), 13).setValue(now);
    sheet.getRange(parseInt(woData.rowIndex), 14).setValue(woData.noTelpPic || '');
  } else {
    sheet.appendRow([
      woData.noTiket, woData.namaCustomer, woData.idCustomer, woData.noTelp,
      woData.alamat, woData.team, woData.pic, woData.tglVisit, woData.status || "On Schedule", "", "Unpaid", now, now, woData.noTelpPic || ''
    ]);
  }
  return { success: true };
}

// 7. Upload Foto ke Google Drive
function uploadImageToDrive(base64Data, fileName) {
  try {
    const prop = PropertiesService.getScriptProperties();
    let folderId = prop.getProperty("PHOTO_FOLDER_ID");
    let folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();

    const contentType = base64Data.substring(base64Data.indexOf(":") + 1, base64Data.indexOf(";"));
    const bytes = Utilities.base64Decode(base64Data.substr(base64Data.indexOf("base64,") + 7));
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    throw new Error("Gagal upload foto: " + e.message);
  }
}

// 8. Simpan Data Report IKR
function saveReportIKRData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  
  if (sheetWO) {
    const woRows = sheetWO.getDataRange().getValues();
    for (let i = 1; i < woRows.length; i++) {
      if (String(woRows[i][0]) === String(data.noTiket)) {
        // Update Status WO jika ada
        if (data.statusWO) {
          sheetWO.getRange(i + 1, 9).setValue(data.statusWO);
        }
        // Update ID Customer jika diisi
        if (data.idCustomer !== undefined) {
          sheetWO.getRange(i + 1, 3).setValue(data.idCustomer);
        }
        sheetWO.getRange(i + 1, 13).setValue(new Date());
        break;
      }
    }
  }

  let sheet = ss.getSheetByName("Report_IKR");
  if (!sheet) {
    sheet = ss.insertSheet("Report_IKR");
    sheet.appendRow([
      "No Tiket", "Status Pelanggan", "Tikor Client", "Tikor ODN", "BTS", "ID ODN",
      "Port", "Port Kosong", "ONT Brand", "SN ONT", "Kabel DW", "S-Clamp", "Segitiga Ulir",
      "Power FAT", "SN PON", "Power Indoor",
      "Foto Power OPM FAT", "Foto Power OPM Indoor", "Foto S-Clamp FAT", "Foto S-Clamp Rumah",
      "Foto Before ODP", "Foto After ODP", "Foto ONT Set", "Foto ONT SN",
      "Foto Speedtest", "Foto CPE ON", "Foto BA Customer", "Foto Customer BA", "Updated At"
    ]);
  }

  const rows = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.noTiket)) {
      rowIndex = i + 1;
      break;
    }
  }

  const now = new Date();
  const rowValues = [
    data.noTiket, "Aktif", data.tikorClient, data.tikorOdn, data.bts, data.idOdn,
    data.portUsed, data.portEmpty, data.ontBrand, data.snOnt, data.kabel, data.sclamp, data.segitiga,
    data.powerFat, data.snPon, data.powerIndoor,
    convertToDirectLink(data.imgPowerFat || ''),
    convertToDirectLink(data.imgPowerIndoor || ''),
    convertToDirectLink(data.imgSclampFat || ''),
    convertToDirectLink(data.imgSclampRumah || ''),
    convertToDirectLink(data.imgBeforeOdp || ''),
    convertToDirectLink(data.imgAfterOdp || ''),
    convertToDirectLink(data.imgOntSet || ''),
    convertToDirectLink(data.imgOntSn || ''),
    convertToDirectLink(data.imgSpeedtest || ''),
    convertToDirectLink(data.imgCpeOn || ''),
    convertToDirectLink(data.imgBaCustomer || ''),
    convertToDirectLink(data.imgCustomerBa || ''),
    now
  ];

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return { success: true };
}

// 9. Ambil Detail Report IKR
function getReportDetail(noTiket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  const sheetRep = ss.getSheetByName("Report_IKR");

  let currentStatusWO = "On Schedule";
  let currentIdCustomer = "";
  
  if (sheetWO) {
    const woRows = sheetWO.getDataRange().getValues();
    const wo = woRows.find(r => String(r[0]) === String(noTiket));
    if (wo) {
      if (wo[8]) currentStatusWO = String(wo[8]);
      if (wo[2]) currentIdCustomer = String(wo[2]);
    }
  }

  if (!sheetRep) return { noTiket: noTiket, statusWO: currentStatusWO, idCustomer: currentIdCustomer };

  const rows = sheetRep.getDataRange().getValues();
  const rep = rows.find(r => String(r[0]) === String(noTiket));
  if (!rep) return { noTiket: noTiket, statusWO: currentStatusWO, idCustomer: currentIdCustomer };

  return {
    noTiket: rep[0],
    statusWO: currentStatusWO,
    idCustomer: currentIdCustomer,
    tikorClient: rep[2] || '',
    tikorOdn: rep[3] || '',
    bts: rep[4] || '',
    idOdn: rep[5] || '',
    portUsed: rep[6] || '',
    portEmpty: rep[7] || '',
    ontBrand: rep[8] || '',
    snOnt: rep[9] || '',
    kabel: rep[10] || '',
    sclamp: rep[11] || '',
    segitiga: rep[12] || '',
    powerFat: rep[13] || '',
    snPon: rep[14] || '',
    powerIndoor: rep[15] || '',
    imgPowerFat: convertToDirectLink(rep[16] || ''),
    imgPowerIndoor: convertToDirectLink(rep[17] || ''),
    imgSclampFat: convertToDirectLink(rep[18] || ''),
    imgSclampRumah: convertToDirectLink(rep[19] || ''),
    imgBeforeOdp: convertToDirectLink(rep[20] || ''),
    imgAfterOdp: convertToDirectLink(rep[21] || ''),
    imgOntSet: convertToDirectLink(rep[22] || ''),
    imgOntSn: convertToDirectLink(rep[23] || ''),
    imgSpeedtest: convertToDirectLink(rep[24] || ''),
    imgCpeOn: convertToDirectLink(rep[25] || ''),
    imgBaCustomer: convertToDirectLink(rep[26] || ''),
    imgCustomerBa: convertToDirectLink(rep[27] || '')
  };
}

// 9. Ambil Detail Report IKR
function getReportDetail(noTiket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  const sheetRep = ss.getSheetByName("Report_IKR");

  let currentStatusWO = "On Schedule";
  let currentIdCustomer = "";
  
  // Ambil Status WO dan ID Customer dari sheet Work_Orders (Sumber Utama WO)
  if (sheetWO) {
    const woRows = sheetWO.getDataRange().getValues();
    const wo = woRows.find(r => String(r[0]) === String(noTiket));
    if (wo) {
      if (wo[8]) currentStatusWO = String(wo[8]); // Kolom ke-9 (Index 8) adalah Status WO
      if (wo[2]) currentIdCustomer = String(wo[2]); // Kolom ke-3 (Index 2) adalah ID Customer
    }
  }

  if (!sheetRep) return { noTiket: noTiket, statusWO: currentStatusWO, idCustomer: currentIdCustomer };

  const rows = sheetRep.getDataRange().getValues();
  const rep = rows.find(r => String(r[0]) === String(noTiket));
  
  if (!rep) return { noTiket: noTiket, statusWO: currentStatusWO, idCustomer: currentIdCustomer };

  return {
    noTiket: rep[0],
    statusWO: currentStatusWO,
    idCustomer: currentIdCustomer, // Mengirimkan ID Customer yang sudah ada dari Work_Orders
    tikorClient: rep[2] || '',
    tikorOdn: rep[3] || '',
    bts: rep[4] || '',
    idOdn: rep[5] || '',
    portUsed: rep[6] || '',
    portEmpty: rep[7] || '',
    ontBrand: rep[8] || '',
    snOnt: rep[9] || '',
    kabel: rep[10] || '',
    sclamp: rep[11] || '',
    segitiga: rep[12] || '',
    powerFat: rep[13] || '',
    snPon: rep[14] || '',
    powerIndoor: rep[15] || '',
    imgPowerFat: convertToDirectLink(rep[16] || ''),
    imgPowerIndoor: convertToDirectLink(rep[17] || ''),
    imgSclampFat: convertToDirectLink(rep[18] || ''),
    imgSclampRumah: convertToDirectLink(rep[19] || ''),
    imgBeforeOdp: convertToDirectLink(rep[20] || ''),
    imgAfterOdp: convertToDirectLink(rep[21] || ''),
    imgOntSet: convertToDirectLink(rep[22] || ''),
    imgOntSn: convertToDirectLink(rep[23] || ''),
    imgSpeedtest: convertToDirectLink(rep[24] || ''),
    imgCpeOn: convertToDirectLink(rep[25] || ''),
    imgBaCustomer: convertToDirectLink(rep[26] || ''),
    imgCustomerBa: convertToDirectLink(rep[27] || '')
  };
}

// 10. WA Template
function generateWATemplate(noTiket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  const sheetReport = ss.getSheetByName("Report_IKR");

  const woData = sheetWO.getDataRange().getValues().find(r => String(r[0]) === String(noTiket));
  const repData = sheetReport ? sheetReport.getDataRange().getValues().find(r => String(r[0]) === String(noTiket)) : null;

  if (!woData) return "Data WO tidak ditemukan";

  const tglFormatted = woData[7] ? Utilities.formatDate(new Date(woData[7]), "GMT+7", "dd.MM.yyyy") : "-";
  const tikorPelanggan = repData ? (repData[2] || '') : '';
  const urlMap = tikorPelanggan ? `https://maps.google.com/?q=${tikorPelanggan}` : '-';
  const noTelpPicVal = woData[13] ? String(woData[13]) : '-';

  let msg = `========================================\n`;
  msg += `📄 *BERITA ACARA IKR BELACOM*\n`;
  msg += `🗓 *Tanggal* : ${tglFormatted}\n`;
  msg += `👥 *Tim Aktivator* : ${woData[5] || '-'}\n`;
  msg += `👤 *PIC Teknisi* : ${woData[6] || '-'}\n`;
  msg += `📞 *No. Telp Teknisi* : ${noTelpPicVal}\n`;
  msg += `📊 *Status* : ${woData[8] || '-'}\n`;
  msg += `========================================\n\n`;

  msg += `📌 *DATA PELANGGAN*\n`;
  msg += `• *WO Number* : ${woData[0]}\n`;
  msg += `• *ID Layanan* : ${woData[2] || '-'}\n`;
  msg += `• *Nama Klien* : ${woData[1] || '-'}\n`;
  msg += `• *No. Telp* : ${woData[3] || '-'}\n`;
  msg += `• *Alamat* : ${woData[4] || '-'}\n`;
  msg += `• *Koordinat Klien* : ${urlMap}\n\n`;

  msg += `📡 *DATA ODN / FAT & BTS*\n`;
  msg += `• *BTS* : ${repData ? (repData[4] || '-') : '-'}\n`;
  msg += `• *ID ODN* : ${repData ? (repData[5] || 'No Label') : 'No Label'}\n`;
  msg += `• *Koordinat ODN* : ${repData ? (repData[3] || '-') : '-'}\n`;
  msg += `• *Port Digunakan* : ${repData ? (repData[6] || '-') : '-'}\n`;
  msg += `• *Port Kosong* : ${repData ? (repData[7] || '-') : '-'}\n`;
  msg += `• *Power FAT* : ${repData ? (repData[13] || '-') : '-'} dBm\n`;
  msg += `• *Power Indoor* : ${repData ? (repData[15] || '-') : '-'} dBm\n\n`;

  msg += `🔌 *SPESIFIKASI ONT*\n`;
  msg += `• *Brand ONT* : ${repData ? (repData[8] || '-') : '-'}\n`;
  msg += `• *SN PON* : ${repData ? (repData[14] || '-') : '-'}\n`;
  msg += `• *SN ONT* : ${repData ? (repData[9] || '-') : '-'}\n\n`;

  msg += `📦 *PEMAKAIAN MATERIAL*\n`;
  msg += `• *Kabel DW Preecon* : ${repData ? (repData[10] || '0') : '0'} M\n`;
  msg += `• *Segitiga Ulir* : ${repData ? (repData[12] || '0') : '0'} pcs\n`;
  msg += `• *S-Clamp* : ${repData ? (repData[11] || '0') : '0'} pcs\n`;
  msg += `========================================`;

  return msg;
}

// 11. Generate PDF Report Engine
function generatePDFReport(noTiket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  const sheetRep = ss.getSheetByName("Report_IKR");

  if (!sheetWO) throw new Error("Sheet Work_Orders tidak ditemukan!");

  const wo = sheetWO.getDataRange().getValues().find(r => String(r[0]) === String(noTiket));
  const rep = sheetRep ? sheetRep.getDataRange().getValues().find(r => String(r[0]) === String(noTiket)) : null;

  if (!wo) throw new Error("Data Work Order tidak ditemukan!");

  const prop = PropertiesService.getScriptProperties();
  const belacomLogoId = prop.getProperty("LOGO_BELACOM_ID");
  const remalaLogoId = prop.getProperty("LOGO_REMALA_ID");

  const imgLogoBelacom = getDriveImageBase64ById_(belacomLogoId);
  const imgLogoRemala = getDriveImageBase64ById_(remalaLogoId);

// Ambil Data Material dari Report IKR
  const kabelDW = rep ? (rep[10] || '0') : '0';
  const sClamp = rep ? (rep[11] || '0') : '0';
  const segitigaUlir = rep ? (rep[12] || '0') : '0';

  const imgPowerFat = rep ? convertToDirectLink(rep[16]) : '';
  const imgPowerIndoor = rep ? convertToDirectLink(rep[17]) : '';
  const imgSclampFat = rep ? convertToDirectLink(rep[18]) : '';
  const imgSclampRumah = rep ? convertToDirectLink(rep[19]) : '';
  const imgBeforeOdp = rep ? convertToDirectLink(rep[20]) : '';
  const imgAfterOdp = rep ? convertToDirectLink(rep[21]) : '';
  const imgOntSet = rep ? convertToDirectLink(rep[22]) : '';
  const imgOntSn = rep ? convertToDirectLink(rep[23]) : '';
  const imgSpeedtest = rep ? convertToDirectLink(rep[24]) : '';
  const imgCpeOn = rep ? convertToDirectLink(rep[25]) : '';
  const imgBaCustomer = rep ? convertToDirectLink(rep[26]) : '';
  const imgCustomerBa = rep ? convertToDirectLink(rep[27]) : '';
  const noTelpPicVal = wo[13] ? String(wo[13]) : '-';

  let htmlContent = `
    <html>
    <head>
      <style>
        @page { size: A4; margin-top: 12mm; margin-bottom: 15mm; margin-left: 8mm; margin-right: 8mm; }
        body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; font-size: 8pt; color: #111; margin: 0; padding: 0; }
        table.main-layout { width: 100%; border-collapse: collapse; }
        thead { display: table-header-group; }
        tfoot { display: table-footer-group; }
        tbody { display: table-row-group; }
        .kop-table { width: 100%; border-collapse: collapse; margin-bottom: 6px; border-bottom: 2px solid #C8102E; }
        .kop-table td { vertical-align: middle; padding: 2px 4px; }
        .title-kop { text-align: center; color: #C8102E; font-size: 11pt; font-weight: bold; }
        .footer-table { width: 100%; border-collapse: collapse; border-top: 1px solid #aaa; margin-top: 8px; font-size: 7.5pt; color: #555; }
        .footer-table td { padding: 4px 0; vertical-align: middle; }
        .sec-header { background-color: #A6A6A6; color: #000; font-weight: bold; font-size: 8.5pt; padding: 3px 6px; margin-top: 6px; margin-bottom: 4px; text-transform: uppercase; page-break-after: avoid; }
        table.tbl { width: 100%; border-collapse: collapse; margin-bottom: 6px; page-break-inside: avoid; }
        table.tbl td, table.tbl th { border: 1px solid #777; padding: 3px 5px; font-size: 8pt; vertical-align: middle; }
        table.tbl th { background-color: #E6E6E6; text-align: center; font-weight: bold; }
        .bg-gray { background-color: #F2F2F2; font-weight: bold; }
        .img-grid { width: 100%; border-collapse: collapse; margin-top: 2px; }
        .img-row { page-break-inside: avoid; }
        .img-grid td { border: 1px solid #999; padding: 3px; text-align: center; width: 50%; vertical-align: top; page-break-inside: avoid; }
        .img-container { width: 100%; height: 120px; display: flex; align-items: center; justify-content: center; overflow: hidden; background: #fafafa; }
        .img-container img { max-width: 100%; max-height: 120px; object-fit: contain; }
        .img-caption { font-size: 7.5pt; font-weight: bold; margin-bottom: 3px; background: #eee; padding: 2px; border: 1px solid #ccc; }
        .no-break { page-break-inside: avoid; }
      </style>
    </head>
    <body>
      <table class="main-layout">
        <thead>
          <tr>
            <td>
              <table class="kop-table">
                <tr>
                  <td width="32%" align="left">${imgLogoBelacom ? `<img src="${imgLogoBelacom}" style="max-height:40px;"/>` : '<strong style="color:#C8102E;">PT. BELACOM TEKNOLOGI INDONESIA</strong>'}</td>
                  <td width="36%" class="title-kop">BERITA ACARA & REPORT IKR</td>
                  <td width="32%" align="right">${imgLogoRemala ? `<img src="${imgLogoRemala}" style="max-height:40px;"/>` : '<strong style="color:#0056B3;">RA REMALA ABADI</strong>'}</td>
                </tr>
              </table>
            </td>
          </tr>
        </thead>
        <tfoot>
          <tr>
            <td>
              <table class="footer-table">
                <tr>
                  <td align="left"><strong>PT. Belacom Teknologi Indonesia</strong> | No. WO: <strong>${wo[0] || '-'}</strong></td>
                  <td align="right">Dokumen Berita Acara & Report IKR Resmi</td>
                </tr>
              </table>
            </td>
          </tr>
        </tfoot>
        <tbody>
          <tr>
            <td>
              <div class="no-break">
                <div class="sec-header">INFORMASI WORK ORDER & PELANGGAN</div>
                <table class="tbl">
                  <tr><td width="18%" class="bg-gray">TIKET/WORK ORDER</td><td width="32%">${wo[0] || '-'}</td><td width="18%" class="bg-gray">TEAM / PIC</td><td width="32%">${wo[5] || '-'} (${wo[6] || '-'})</td></tr>
                  <tr><td class="bg-gray">NO TELP TEKNISI</td><td>${noTelpPicVal}</td><td class="bg-gray">TANGGAL VISIT</td><td>${wo[7] ? Utilities.formatDate(new Date(wo[7]), "GMT+7", "dd/MM/yyyy") : "-"}</td></tr>
                  <tr><td class="bg-gray">NAMA CUSTOMER</td><td>${wo[1] || '-'}</td><td class="bg-gray">STATUS PELANGGAN</td><td>${rep ? (rep[1] || 'Aktif') : 'Aktif'}</td></tr>
                  <tr><td class="bg-gray">ID / NO TELP</td><td>${wo[2] || '-'} / ${wo[3] || '-'}</td><td class="bg-gray">ALAMAT</td><td>${wo[4] || '-'}</td></tr>
                </table>
              </div>
              <div class="no-break">
                <div class="sec-header">DATA TEKNIS</div>
                <table class="tbl">
                  <tr><td width="25%" class="bg-gray">LAT_LONG CLIENT</td><td width="25%">${rep ? (rep[2] || '-') : '-'}</td><td width="25%" class="bg-gray">LAT_LONG ODN</td><td width="25%">${rep ? (rep[3] || '-') : '-'}</td></tr>
                  <tr><td class="bg-gray">BTS</td><td>${rep ? (rep[4] || '-') : '-'}</td><td class="bg-gray">PORT ODN / KOSONG</td><td>${rep ? (rep[6] || '-') : '-'} / ${rep ? (rep[7] || '-') : '-'}</td></tr>
                  <tr><td class="bg-gray">ID ODN</td><td>${rep ? (rep[5] || '-') : '-'}</td><td class="bg-gray">POWER FAT / INDOOR</td><td>${rep ? (rep[13] || '-') : '-'} dBm / ${rep ? (rep[15] || '-') : '-'} dBm</td></tr>
                  <tr><td class="bg-gray">TYPE ONT / BRAND</td><td>${rep ? (rep[8] || '-') : '-'}</td><td class="bg-gray">SN PON / SN ONT</td><td>${rep ? (rep[14] || '-') : '-'} / ${rep ? (rep[9] || '-') : '-'}</td></tr>
                </table>
              </div>

              <!-- BAGIAN BARU: PEMAKAIAN MATERIAL -->
              <div class="no-break">
                <div class="sec-header">PEMAKAIAN MATERIAL</div>
                <table class="tbl">
                  <tr>
                    <td width="33%" class="bg-gray">KABEL DW PRECON</td>
                    <td width="33%" class="bg-gray">S-CLAMP</td>
                    <td width="34%" class="bg-gray">SEGITIGA ULIR</td>
                  </tr>
                  <tr>
                    <td>${kabelDW} Meter</td>
                    <td>${sClamp} Pcs</td>
                    <td>${segitigaUlir} Pcs</td>
                  </tr>
                </table>
              </div>

              <div class="sec-header">FOTO DOKUMENTASI</div>
              <table class="img-grid">
                <tr class="img-row"><td><div class="img-caption">1. POWER OPM FAT</div><div class="img-container">${imgPowerFat ? `<img src="${imgPowerFat}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td><td><div class="img-caption">2. POWER OPM INDOOR</div><div class="img-container">${imgPowerIndoor ? `<img src="${imgPowerIndoor}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td></tr>
                <tr class="img-row"><td><div class="img-caption">3. S-CLAMP FAT</div><div class="img-container">${imgSclampFat ? `<img src="${imgSclampFat}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td><td><div class="img-caption">4. S-CLAMP RUMAH</div><div class="img-container">${imgSclampRumah ? `<img src="${imgSclampRumah}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td></tr>
                <tr class="img-row"><td><div class="img-caption">5. BEFORE ODP</div><div class="img-container">${imgBeforeOdp ? `<img src="${imgBeforeOdp}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td><td><div class="img-caption">6. AFTER ODP</div><div class="img-container">${imgAfterOdp ? `<img src="${imgAfterOdp}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td></tr>
                <tr class="img-row"><td><div class="img-caption">7. ONT SET</div><div class="img-container">${imgOntSet ? `<img src="${imgOntSet}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td><td><div class="img-caption">8. ONT SN</div><div class="img-container">${imgOntSn ? `<img src="${imgOntSn}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td></tr>
                <tr class="img-row"><td><div class="img-caption">9. SPEEDTEST</div><div class="img-container">${imgSpeedtest ? `<img src="${imgSpeedtest}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td><td><div class="img-caption">10. CPE ON</div><div class="img-container">${imgCpeOn ? `<img src="${imgCpeOn}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td></tr>
                <tr class="img-row"><td><div class="img-caption">11. BA CUSTOMER</div><div class="img-container">${imgBaCustomer ? `<img src="${imgBaCustomer}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td><td><div class="img-caption">12. CUSTOMER BA</div><div class="img-container">${imgCustomerBa ? `<img src="${imgCustomerBa}"/>` : `<span style="color:#aaa;">[Tanpa Foto]</span>`}</div></td></tr>
              </table>
            </td>
          </tr>
        </tbody>
      </table>
    </body>
    </html>
  `;

  const blob = Utilities.newBlob(htmlContent, "text/html", `Report_IKR_${noTiket}.html`);
  const pdfBlob = blob.getAs("application/pdf").setName(`Report_IKR_${noTiket}.pdf`);

  let folderId = prop.getProperty("DOC_FOLDER_ID");
  let folder;
  try {
    folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  } catch (e) {
    folder = DriveApp.getRootFolder();
  }

  const pdfFile = folder.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return pdfFile.getUrl();
}

function getDriveImageBase64ById_(fileId) {
  if (!fileId) return '';
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return '';
  }
}

function getTeamStatusPerformance(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  if (!sheetWO) return { teams: [], statuses: [], datasets: [] };

  const data = sheetWO.getDataRange().getValues();
  if (data.length <= 1) return { teams: [], statuses: [], datasets: [] };

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";

  let teamSet = new Set();
  let statusList = [
    "Done", "On Progress", "On Schedule", "Reschedule", 
    "Pending", "Hold To Sales", "Hold To Planning", "Hold To Maintenance"
  ];
  
  let mapData = {};

  for (let i = 1; i < data.length; i++) {
    let r = data[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    let team = r[5] ? String(r[5]) : "Tanpa Team";
    let status = r[8] ? String(r[8]) : "On Schedule";

    teamSet.add(team);
    if (!mapData[team]) {
      mapData[team] = {};
      statusList.forEach(st => mapData[team][st] = 0);
    }
    if (mapData[team][status] !== undefined) {
      mapData[team][status]++;
    } else {
      mapData[team][status] = 1;
    }
  }

  let teams = Array.from(teamSet);
  const statusColors = {
    "Done": "#198754", "On Progress": "#0d6efd", "On Schedule": "#0dcaf0",
    "Reschedule": "#ffc107", "Pending": "#dc3545", "Hold To Sales": "#6c757d",
    "Hold To Planning": "#adb5bd", "Hold To Maintenance": "#343a40"
  };

  let datasets = statusList.map(status => {
    let counts = teams.map(team => (mapData[team] ? mapData[team][status] || 0 : 0));
    return { label: status, data: counts, backgroundColor: statusColors[status] || "#6c757d" };
  });

  return { teams: teams, datasets: datasets };
}

function getTeamsManagementList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Teams");
  if (!sheet) {
    sheet = ss.insertSheet("Teams");
    sheet.appendRow(["Nama Team", "PIC / Teknisi", "No Telp PIC"]);
    sheet.appendRow(["BLC-IKR01", "Ahmad & Budi", "081234567890"]);
    sheet.appendRow(["BLC-IKR02", "Candra & Dedi", "081298765432"]);
  }
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  return rows.slice(1).map((r, idx) => ({
    rowIndex: idx + 2,
    teamName: r[0] ? String(r[0]) : "",
    picName: r[1] ? String(r[1]) : "",
    noTelpPic: r[2] ? String(r[2]) : ""
  }));
}

function saveTeamData(teamData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Teams");
  if (!sheet) {
    sheet = ss.insertSheet("Teams");
    sheet.appendRow(["Nama Team", "PIC / Teknisi", "No Telp PIC"]);
  }

  const rows = sheet.getDataRange().getValues();
  if (teamData.rowIndex) {
    sheet.getRange(parseInt(teamData.rowIndex), 1, 1, 3).setValues([[
      teamData.teamName, teamData.picName, teamData.noTelpPic
    ]]);
  } else {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).trim().toLowerCase() === String(teamData.picName).trim().toLowerCase()) {
        return { success: false, message: "Nama PIC / Teknisi sudah terdaftar!" };
      }
    }
    sheet.appendRow([teamData.teamName, teamData.picName, teamData.noTelpPic]);
  }
  return { success: true };
}

function deleteTeamRow(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Teams");
  if (!sheet) return { success: false, message: "Sheet Teams tidak ditemukan" };
  sheet.deleteRow(parseInt(rowIndex));
  return { success: true };
}

// --- FITUR EXPORT EXCEL (URUTAN KOLOM SESUAI REQUEST) ---

// 1. Export List Work Order (WO)
function exportWorkOrdersToExcel(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Work_Orders");
  if (!sheet) return "";
  
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return "";

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";

  let csvRows = [];
  
  // Header kolom Work Order tersusun ke samping sesuai permintaan
  csvRows.push([
    "No Tiket",
    "Nama Customer",
    "ID Customer",
    "No Telp",
    "Alamat",
    "Team",
    "PIC",
    "Tgl Visit",
    "Status",
    "Tgl Bayar",
    "Status Bayar"
  ].map(field => `"${field}"`).join(","));

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    let tglVisitVal = r[7] ? (r[7] instanceof Date ? Utilities.formatDate(r[7], "GMT+7", "yyyy-MM-dd") : String(r[7])) : "";
    let tglBayarVal = r[9] ? (r[9] instanceof Date ? Utilities.formatDate(r[9], "GMT+7", "yyyy-MM-dd") : String(r[9])) : "";
    let statusBayarVal = r[10] ? String(r[10]) : "Unpaid";

    let row = [
      r[0] ? String(r[0]) : "", // No Tiket
      r[1] ? String(r[1]) : "", // Nama Customer
      r[2] ? String(r[2]) : "", // ID Customer
      r[3] ? String(r[3]) : "", // No Telp
      r[4] ? String(r[4]) : "", // Alamat
      r[5] ? String(r[5]) : "", // Team
      r[6] ? String(r[6]) : "", // PIC
      tglVisitVal,              // Tgl Visit
      r[8] ? String(r[8]) : "", // Status
      tglBayarVal,              // Tgl Bayar
      statusBayarVal            // Status Bayar
    ];

    csvRows.push(row.map(val => `"${String(val).replace(/"/g, '""')}"`).join(","));
  }

  const csvString = csvRows.join("\n");
  const blob = Utilities.newBlob(csvString, "text/csv;charset=utf-8;", `List_Work_Order_${Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss")}.csv`);
  return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
}

// 2. Export List Report IKR
function exportReportIKRToExcel(filters) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWO = ss.getSheetByName("Work_Orders");
  const sheetRep = ss.getSheetByName("Report_IKR");
  if (!sheetWO) return "";

  const woRows = sheetWO.getDataRange().getValues();
  if (woRows.length <= 1) return "";

  const repRows = sheetRep ? sheetRep.getDataRange().getValues() : [];

  filters = filters || {};
  let startDate = filters.startDate ? new Date(filters.startDate) : null;
  let endDate = filters.endDate ? new Date(filters.endDate) : null;
  let statusFilter = filters.status || "";
  let teamFilter = filters.team || "";

  let csvRows = [];
  
  // Header kolom Report IKR tersusun ke samping persis sesuai permintaan
  csvRows.push([
    "No Tiket",
    "Status Pelanggan",
    "Tikor Client",
    "Tikor ODN",
    "BTS",
    "ID ODN",
    "Port",
    "Port Kosong",
    "ONT Brand",
    "SN ONT",
    "Kabel DW",
    "S-Clamp",
    "Segitiga Ulir",
    "Power FAT",
    "SN PON",
    "Power Indoor",
    "Foto Power OPM FAT",
    "Foto Power OPM Indoor",
    "Foto S-Clamp FAT",
    "Foto S-Clamp Rumah",
    "Foto Before ODP",
    "Foto After ODP",
    "Foto ONT Set",
    "Foto ONT SN",
    "Foto Speedtest",
    "Foto CPE ON",
    "Foto BA Customer",
    "Foto Customer BA"
  ].map(field => `"${field}"`).join(","));

  for (let i = 1; i < woRows.length; i++) {
    const r = woRows[i];
    if (!applyFilters_(r, startDate, endDate, statusFilter, teamFilter)) continue;

    const noTiket = String(r[0]);
    const rep = repRows.find(rp => String(rp[0]) === noTiket);

    let row = [
      noTiket,                                // No Tiket
      rep ? (rep[1] || "Aktif") : "Aktif",    // Status Pelanggan
      rep ? (rep[2] || "") : "",              // Tikor Client
      rep ? (rep[3] || "") : "",              // Tikor ODN
      rep ? (rep[4] || "") : "",              // BTS
      rep ? (rep[5] || "") : "",              // ID ODN
      rep ? (rep[6] || "") : "",              // Port
      rep ? (rep[7] || "") : "",              // Port Kosong
      rep ? (rep[8] || "") : "",              // ONT Brand
      rep ? (rep[9] || "") : "",              // SN ONT
      rep ? (rep[10] || "") : "",             // Kabel DW
      rep ? (rep[11] || "") : "",             // S-Clamp
      rep ? (rep[12] || "") : "",             // Segitiga Ulir
      rep ? (rep[13] || "") : "",             // Power FAT
      rep ? (rep[14] || "") : "",             // SN PON
      rep ? (rep[15] || "") : "",             // Power Indoor
      rep ? (rep[16] || "") : "",             // Foto Power OPM FAT
      rep ? (rep[17] || "") : "",             // Foto Power OPM Indoor
      rep ? (rep[18] || "") : "",             // Foto S-Clamp FAT
      rep ? (rep[19] || "") : "",             // Foto S-Clamp Rumah
      rep ? (rep[20] || "") : "",             // Foto Before ODP
      rep ? (rep[21] || "") : "",             // Foto After ODP
      rep ? (rep[22] || "") : "",             // Foto ONT Set
      rep ? (rep[23] || "") : "",             // Foto ONT SN
      rep ? (rep[24] || "") : "",             // Foto Speedtest
      rep ? (rep[25] || "") : "",             // Foto CPE ON
      rep ? (rep[26] || "") : "",             // Foto BA Customer
      rep ? (rep[27] || "") : ""              // Foto Customer BA
    ];

    csvRows.push(row.map(val => `"${String(val).replace(/"/g, '""')}"`).join(","));
  }

  const csvString = csvRows.join("\n");
  const blob = Utilities.newBlob(csvString, "text/csv;charset=utf-8;", `List_Report_IKR_${Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmmss")}.csv`);
  return "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
}

// 11. Hapus Work Order berdasarkan No Tiket
function deleteWorkOrder(noTiket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Work_Orders");
  if (!sheet) return { success: false, message: "Sheet Work_Orders tidak ditemukan" };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(noTiket).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, message: "Work Order tidak ditemukan" };
}


// =====================================================================
// PENGATURAN HARGA DENGAN TANGGAL EFEKTIF
// Sheet "Pengaturan_Harga": Tgl Efektif | Tagihan WO | Bayar Team IKR | Revenue Belacom | Diubah Oleh | Diubah Pada | Catatan
// Aturan: harga sebuah WO = baris dengan Tgl Efektif terbesar yang <= Tgl Visit WO
//         (jika Tgl Visit kosong dipakai Created At). Riwayat tidak dihapus, sehingga
//         laporan bulan lalu tetap konsisten walau harga hari ini berubah.
// =====================================================================
const PRICE_SHEET_ = "Pengaturan_Harga";
const PRICE_BASE_DATE_ = "2000-01-01";

function toDateKey_(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return isNaN(v) ? "" : Utilities.formatDate(v, "GMT+7", "yyyy-MM-dd");
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  const d = new Date(s);
  return isNaN(d) ? "" : Utilities.formatDate(d, "GMT+7", "yyyy-MM-dd");
}

function getOrCreatePriceSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PRICE_SHEET_);
  if (!sheet) {
    sheet = ss.insertSheet(PRICE_SHEET_);
    sheet.appendRow(["Tgl Efektif", "Tagihan WO", "Bayar Team IKR", "Revenue Belacom", "Diubah Oleh", "Diubah Pada", "Catatan"]);
    sheet.getRange("A:A").setNumberFormat("@");
    // Baris dasar = tarif lama, supaya data historis tidak berubah
    sheet.appendRow([PRICE_BASE_DATE_, 100000, 75000, 25000, "system", new Date(), "Tarif awal"]);
  }
  return sheet;
}

function loadPriceTable_() {
  const rows = getOrCreatePriceSheet_().getDataRange().getValues().slice(1);
  return rows
    .map(r => ({
      eff: toDateKey_(r[0]),
      tagihan: Number(r[1]) || 0,
      team: Number(r[2]) || 0,
      revenue: Number(r[3]) || 0,
      by: r[4] ? String(r[4]) : "",
      at: r[5] instanceof Date ? Utilities.formatDate(r[5], "GMT+7", "yyyy-MM-dd HH:mm") : "",
      catatan: r[6] ? String(r[6]) : ""
    }))
    .filter(p => p.eff)
    .sort((a, b) => (a.eff < b.eff ? -1 : a.eff > b.eff ? 1 : 0));
}

function resolvePrice_(table, dateKey) {
  const key = dateKey || Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  let found = table[0];
  for (let i = 0; i < table.length; i++) {
    if (table[i].eff <= key) found = table[i]; else break;
  }
  return found || { tagihan: 0, team: 0, revenue: 0 };
}

function priceForWoRow_(table, r) {
  return resolvePrice_(table, toDateKey_(r[7]) || toDateKey_(r[11]));
}

function getPriceSettings() {
  const table = loadPriceTable_();
  const today = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  let activeEff = resolvePrice_(table, today).eff;
  return table.map(p => ({
    tglEfektif: p.eff,
    tagihan: p.tagihan,
    team: p.team,
    revenue: p.revenue,
    diubahOleh: p.by,
    diubahPada: p.at,
    catatan: p.catatan,
    isBase: p.eff === table[0].eff,
    status: p.eff > today ? "Terjadwal" : (p.eff === activeEff ? "Berlaku" : "Riwayat")
  })).reverse();
}

function savePriceSetting(data) {
  const eff = toDateKey_(data.tglEfektif);
  if (!eff) return { success: false, message: "Tanggal efektif tidak valid." };
  const t = Number(data.tagihan), p = Number(data.team), v = Number(data.revenue);
  if ([t, p, v].some(n => isNaN(n) || n < 0)) return { success: false, message: "Semua nominal harus angka >= 0." };

  const sheet = getOrCreatePriceSheet_();
  const rows = sheet.getDataRange().getValues();
  const rowValues = [eff, t, p, v, data.updatedBy || "", new Date(), data.catatan || ""];

  for (let i = 1; i < rows.length; i++) {
    if (toDateKey_(rows[i][0]) === eff) {
      sheet.getRange(i + 1, 1, 1, 7).setValues([rowValues]);
      return { success: true, updated: true };
    }
  }
  sheet.appendRow(rowValues);
  if (sheet.getLastRow() > 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).sort({ column: 1, ascending: true });
  return { success: true, updated: false };
}

function deletePriceSetting(tglEfektif) {
  const eff = toDateKey_(tglEfektif);
  const sheet = getOrCreatePriceSheet_();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 2) return { success: false, message: "Minimal harus ada satu pengaturan harga." };
  const table = loadPriceTable_();
  if (eff === table[0].eff) return { success: false, message: "Harga dasar (paling awal) tidak boleh dihapus." };
  for (let i = 1; i < rows.length; i++) {
    if (toDateKey_(rows[i][0]) === eff) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { success: false, message: "Data harga tidak ditemukan." };
}

// Pratinjau dampak: berapa WO yang memakai harga ini bila berlaku mulai tglEfektif
function getPriceImpact(tglEfektif) {
  const eff = toDateKey_(tglEfektif);
  const table = loadPriceTable_();
  const next = table.map(p => p.eff).filter(e => e > eff)[0] || "";
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Work_Orders");
  let affected = 0, paid = 0;
  if (sheet && eff) {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const k = toDateKey_(rows[i][7]) || toDateKey_(rows[i][11]);
      if (k && k >= eff && (!next || k < next)) {
        affected++;
        if (String(rows[i][10]) === "Paid") paid++;
      }
    }
  }
  return { affectedWO: affected, paidWO: paid, from: eff, to: next };
}
