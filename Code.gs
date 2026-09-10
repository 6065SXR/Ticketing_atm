/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: Code.gs
 * Fungsi: Master Backend Controller, Web App Entry Point, User Authentication,
 *         Routing getSectionData, dan Integrasi Modul Ekstensi.
 */

function doGet(e) {
  try {
    var ss = getSpreadsheet();
    if (ss) {
      findOrCreateUsersSheet(ss);
      findOrCreateFSESheet(ss);
    }
  } catch (err) {
    Logger.log('Init sheet warning: ' + err.toString());
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ATM Ticketing System - Hardiansyah Fam')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SPREADSHEET_ID');
  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log('Gagal membuka spreadsheet via ID: ' + e.toString());
    }
  }
  return null;
}

function findSheet(ss, possibleNames) {
  if (!ss) return null;
  var sheets = ss.getSheets();
  if (!sheets || sheets.length === 0) return null;

  var isLookingForCounter = false;
  for (var p = 0; p < possibleNames.length; p++) {
    if (possibleNames[p].toLowerCase().indexOf('counter') !== -1) {
      isLookingForCounter = true;
      break;
    }
  }

  // 1. Exact match (Case-Insensitive)
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (!isLookingForCounter && name.indexOf('counter') !== -1) continue;
    for (var j = 0; j < possibleNames.length; j++) {
      if (name === possibleNames[j].trim().toLowerCase()) {
        return sheets[i];
      }
    }
  }

  // 2. Substring match
  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName().trim().toLowerCase();
    if (!isLookingForCounter && sName.indexOf('counter') !== -1) continue;
    for (var m = 0; m < possibleNames.length; m++) {
      var target = possibleNames[m].trim().toLowerCase();
      if (target === 'sheet1' || target === 'sheet 1') continue;
      if (sName.indexOf(target) !== -1 || target.indexOf(sName) !== -1) {
        return sheets[k];
      }
    }
  }

  return null;
}

function findTicketsSheet(ss) {
  if (!ss) return null;
  var sheets = ss.getSheets();
  if (!sheets || sheets.length === 0) return null;

  var ticketNames = [
    'tickets', 'ticket', 'tiket', 'data tiket', 'data_tiket', 
    'datatickets', 'dataticket', 'data_ticket', 'datatiket',
    'ticketing', 'trouble ticket', 'trouble_ticket', 'trouble tickets',
    'problem ticket', 'incident', 'incidents', 'gangguan', 'daftar tiket', 
    'list tiket', 'master tiket', 'tiket atm', 'ticket atm'
  ];

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (name.indexOf('counter') !== -1) continue;
    for (var j = 0; j < ticketNames.length; j++) {
      if (name === ticketNames[j]) {
        return sheets[i];
      }
    }
  }

  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName().trim().toLowerCase();
    if (sName.indexOf('counter') !== -1) continue;
    if (sName === 'sheet1' || sName === 'sheet 1') continue;
    if (sName.indexOf('tiket') !== -1 || sName.indexOf('ticket') !== -1) {
      return sheets[k];
    }
  }

  for (var m = 0; m < sheets.length; m++) {
    var sh = sheets[m];
    var shName = sh.getName().trim().toLowerCase();
    if (shName.indexOf('counter') !== -1) continue;
    if (sh.getLastRow() > 0 && sh.getLastColumn() > 0) {
      var maxCols = Math.min(sh.getLastColumn(), 20);
      var headerVals = sh.getRange(1, 1, 1, maxCols).getDisplayValues()[0];
      var headerJoined = headerVals.join(' ').toLowerCase();
      if (
        headerJoined.indexOf('problem') !== -1 || 
        headerJoined.indexOf('keluhan') !== -1 || 
        headerJoined.indexOf('kendala') !== -1 ||
        headerJoined.indexOf('masalah') !== -1 ||
        (headerJoined.indexOf('bank') !== -1 && (headerJoined.indexOf('tiket') !== -1 || headerJoined.indexOf('ticket') !== -1 || headerJoined.indexOf('sn') !== -1))
      ) {
        return sh;
      }
    }
  }

  for (var n = 0; n < sheets.length; n++) {
    if (sheets[n].getName().trim().toLowerCase().indexOf('counter') === -1) {
      return sheets[n];
    }
  }

  return sheets[0];
}

function findOrCreateUsersSheet(ss) {
  var sheet = findSheet(ss, ['Users', 'User', 'Data Users', 'Data_Users', 'Akun', 'Pengguna']);
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    var headers = ['User_ID', 'Nama_Lengkap', 'No_HP', 'Password', 'Role', 'Wilayah_Tugas', 'Status_Akun', 'Created_By', 'Created_At'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var initialUsers = [
      ['USR-001', 'Monitoring Officer', '081200000002', 'monitoring123', 'Monitoring', 'NOC Pusat', 'Aktif', 'System', nowFormatted],
      ['USR-002', 'Super Administrator', '081100000001', 'admin123', 'Admin', 'Head Office', 'Aktif', 'System', nowFormatted],
      ['USR-003', 'Rizky Hardiansyah', '081299881122', 'fse123', 'FSE', 'Jakarta Pusat & Barat', 'Aktif', 'System', nowFormatted],
      ['USR-004', 'Fajar Ramadhan', '081377665544', 'fse123', 'FSE', 'Jakarta Selatan & Depok', 'Aktif', 'System', nowFormatted],
      ['USR-005', 'Dimas Kurniawan', '081733445566', 'fse123', 'FSE', 'Jakarta Timur & Bekasi', 'Aktif', 'System', nowFormatted]
    ];
    sheet.getRange(2, 1, initialUsers.length, initialUsers[0].length).setValues(initialUsers);
    SpreadsheetApp.flush();
  }
  return sheet;
}

function authenticateUser(noHp, password) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak dapat diakses' };

    var sheet = findOrCreateUsersSheet(ss);
    var cleanNoHp = String(noHp || '').replace(/[^0-9]/g, '').trim();
    var cleanPass = String(password || '').trim();

    if (!cleanNoHp || !cleanPass) {
      return { status: 'error', message: 'Nomor Handphone dan Password wajib diisi!' };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: 'error', message: 'Database pengguna masih kosong' };
    }

    var lastCol = Math.max(sheet.getLastColumn(), 9);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var colMap = { id: 0, nama: 1, noHp: 2, pass: 3, role: 4, wilayah: 5, status: 6 };
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'userid' || h === 'id') colMap.id = c;
      else if (h === 'namalengkap' || h === 'nama') colMap.nama = c;
      else if (h === 'nohp' || h === 'hp' || h === 'phone' || h === 'telepon') colMap.noHp = c;
      else if (h === 'password' || h === 'pass' || h === 'katasandi') colMap.pass = c;
      else if (h === 'role' || h === 'peran' || h === 'tipe') colMap.role = c;
      else if (h === 'wilayahtugas' || h === 'wilayah' || h === 'area') colMap.wilayah = c;
      else if (h === 'statusakun' || h === 'status') colMap.status = c;
    }

    var matchedUser = null;
    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      var rowHp = String(row[colMap.noHp] || '').replace(/[^0-9]/g, '').trim();
      var rowPass = String(row[colMap.pass] || '').trim();
      var rowStatus = String(row[colMap.status] || 'Aktif').trim();

      if (rowHp === cleanNoHp && rowPass === cleanPass) {
        if (rowStatus.toLowerCase() === 'nonaktif' || rowStatus.toLowerCase() === 'inactive') {
          return { status: 'error', message: 'Akun Anda berstatus Nonaktif. Silakan hubungi Monitoring Officer.' };
        }

        var roleRaw = String(row[colMap.role] || 'Monitoring').trim();
        var normalizedRole = 'monitoring';
        if (roleRaw.toLowerCase().indexOf('admin') !== -1) normalizedRole = 'admin';
        else if (roleRaw.toLowerCase().indexOf('fse') !== -1 || roleRaw.toLowerCase().indexOf('engineer') !== -1 || roleRaw.toLowerCase().indexOf('teknisi') !== -1) normalizedRole = 'fse';
        else normalizedRole = 'monitoring';

        var roleTitles = { admin: 'System Administrator', monitoring: 'Monitoring Officer', fse: 'Field Service Engineer' };
        var initials = { admin: 'AD', monitoring: 'MO', fse: 'FS' };

        matchedUser = {
          userId: String(row[colMap.id] || ''),
          name: String(row[colMap.nama] || 'Pengguna ATM'),
          noHp: rowHp,
          roleKey: normalizedRole,
          roleTitle: roleTitles[normalizedRole] || 'Petugas Monitoring',
          initials: initials[normalizedRole] || 'US',
          wilayah: String(row[colMap.wilayah] || 'DKI Jakarta'),
          indicator: normalizedRole.toUpperCase() + ' OPS'
        };
        break;
      }
    }

    if (matchedUser) {
      return { status: 'success', user: matchedUser };
    } else {
      return { status: 'error', message: 'Nomor Handphone atau Password tidak sesuai!' };
    }
  } catch (err) {
    return { status: 'error', message: 'Gagal autentikasi: ' + err.toString() };
  }
}

function getDropdownData() {
  try {
    var custRes = getCustomersData();
    var machRes = getMachinesData();
    var fseRes = getFSEData();

    var custList = (custRes && custRes.data) ? custRes.data : [];
    var banks = [];
    for (var i = 0; i < custList.length; i++) {
      var bName = custList[i].namaBank;
      if (bName && banks.indexOf(bName) === -1) {
        banks.push(bName);
      }
    }

    if (banks.length === 0 && machRes && machRes.data) {
      for (var j = 0; j < machRes.data.length; j++) {
        var mBank = machRes.data[j].bank;
        if (mBank && banks.indexOf(mBank) === -1) {
          banks.push(mBank);
        }
      }
    }

    return {
      status: 'success',
      data: {
        banks: banks,
        machines: (machRes && machRes.data) ? machRes.data : [],
        engineers: (fseRes && fseRes.data) ? fseRes.data : []
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getSectionData(sectionName) {
  try {
    if (sectionName === 'FSE') {
      return getFSEData();
    }
    if (sectionName === 'DatabaseMesin') {
      return getMachinesData();
    }
    if (sectionName === 'DatabaseCustomer') {
      return getCustomersData();
    }
    if (sectionName === 'DashboardMetrics') {
      return getDashboardMetrics();
    }
    if (sectionName === 'ReportingFSE') {
      return getFSEReportingData();
    }
    if (sectionName === 'DataPM') {
      return getPMData();
    }
    return { status: 'error', message: 'Seksi [' + sectionName + '] tidak dikenali' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
