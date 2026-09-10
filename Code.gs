/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: Code.gs
 * Fungsi: Master Backend Controller, Web App Entry Point, User Authentication,
 *         Routing getSectionData, Telemetri Personal FSE & Verifikasi Dokumen.
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
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
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

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (!isLookingForCounter && name.indexOf('counter') !== -1) continue;
    for (var j = 0; j < possibleNames.length; j++) {
      if (name === possibleNames[j].trim().toLowerCase()) {
        return sheets[i];
      }
    }
  }

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
      if (name === ticketNames[j]) return sheets[i];
    }
  }

  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName().trim().toLowerCase();
    if (sName.indexOf('counter') !== -1 || sName === 'sheet1') continue;
    if (sName.indexOf('tiket') !== -1 || sName.indexOf('ticket') !== -1) return sheets[k];
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
      else if (h === 'nohp' || h === 'hp') colMap.noHp = c;
      else if (h === 'password' || h === 'pass') colMap.pass = c;
      else if (h === 'role' || h === 'peran') colMap.role = c;
      else if (h === 'wilayahtugas' || h === 'wilayah') colMap.wilayah = c;
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
        else if (roleRaw.toLowerCase().indexOf('fse') !== -1 || roleRaw.toLowerCase().indexOf('engineer') !== -1) normalizedRole = 'fse';
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

function getFSEPersonalSummary(fseName, fseUserId) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak dapat diakses' };

    var now = new Date();
    var currentMonth = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM');
    var targetName = String(fseName || '').toLowerCase().trim();

    // 1. Ambil jumlah kelolaan mesin dari sheet FSE
    var fseSheet = findSheet(ss, ['FSE', 'Data_FSE', 'Master_FSE']);
    var jumlahMesin = 5;
    var wilayahFSE = 'Jabodetabek';

    if (fseSheet && fseSheet.getLastRow() > 1) {
      var fData = fseSheet.getRange(2, 1, fseSheet.getLastRow() - 1, fseSheet.getLastColumn()).getDisplayValues();
      for (var f = 0; f < fData.length; f++) {
        var uId = String(fData[f][0] || '').trim();
        var uName = String(fData[f][1] || '').toLowerCase().trim();
        if ((fseUserId && uId === fseUserId) || (targetName && uName === targetName)) {
          wilayahFSE = String(fData[f][5] || 'Jabodetabek');
          jumlahMesin = parseInt(fData[f][7], 10) || 0;
          break;
        }
      }
    }

    // 2. Ambil tiket CM dari sheet Tickets (bulan ini & list aktif)
    var ticketSheet = findTicketsSheet(ss);
    var totalCM = 0;
    var assignedTickets = [];
    var closedDocuments = [];

    if (ticketSheet && ticketSheet.getLastRow() > 1) {
      var tRows = ticketSheet.getRange(2, 1, ticketSheet.getLastRow() - 1, Math.max(ticketSheet.getLastColumn(), 16)).getDisplayValues();
      for (var t = 0; t < tRows.length; t++) {
        var row = tRows[t];
        var idTiket = String(row[0] || '').trim();
        var bank = String(row[1] || '').trim();
        var sn = String(row[2] || '').trim();
        var idMesin = String(row[3] || '').trim();
        var tgl = String(row[4] || '').trim();
        var jam = String(row[5] || '').trim();
        var tipe = String(row[6] || '').trim();
        var eng = String(row[7] || '').toLowerCase().trim();
        var prob = String(row[8] || '').trim();
        var note = String(row[9] || '').trim();
        var status = String(row[10] || 'Open').trim();
        var closedAt = String(row[12] || '').trim();
        var hardcopyStatus = String(row[13] || 'Belum Dikirim').trim();
        var hardcopyRecAt = String(row[14] || '').trim();
        var hardcopyRecBy = String(row[15] || '').trim();

        if (targetName && (eng.indexOf(targetName) !== -1 || targetName.indexOf(eng) !== -1)) {
          var isCM = (tipe.toLowerCase().indexOf('cm') !== -1 || tipe.toLowerCase().indexOf('corrective') !== -1);
          var isThisMonth = (tgl.indexOf(currentMonth) === 0 || tgl.indexOf(currentMonth.replace('-', '/')) === 0);
          if (isCM && isThisMonth) totalCM++;

          var ticketObj = {
            idTiket: idTiket, bank: bank, serialNumber: sn, idMesin: idMesin,
            tanggalTiket: tgl, jamTiket: jam, tipeTiket: tipe, namaEngineer: row[7],
            problem: prob, note: note, statusTiket: status,
            closedAt: closedAt, hardcopyStatus: hardcopyStatus,
            hardcopyReceivedAt: hardcopyRecAt, hardcopyReceivedBy: hardcopyRecBy
          };

          if (status.toLowerCase() === 'closed') {
            closedDocuments.push(ticketObj);
          } else {
            assignedTickets.push(ticketObj);
          }
        }
      }
    }

    // 3. Ambil jadwal PM FSE bulan ini dari sheet Data_PM
    var pmSheet = findSheet(ss, ['Data_PM', 'Data PM', 'DataPM', 'PM']);
    var totalPMTarget = 0;
    var totalPMDone = 0;

    if (pmSheet && pmSheet.getLastRow() > 1) {
      var pmRows = pmSheet.getRange(2, 1, pmSheet.getLastRow() - 1, Math.max(pmSheet.getLastColumn(), 10)).getDisplayValues();
      for (var p = 0; p < pmRows.length; p++) {
        var pr = pmRows[p];
        var pEng = String(pr[8] || '').toLowerCase().trim(); // Engineer PIC
        var pStat = String(pr[9] || '').toLowerCase().trim();
        var pTgl = String(pr[6] || '');

        if (targetName && (pEng.indexOf(targetName) !== -1 || targetName.indexOf(pEng) !== -1)) {
          if (pTgl.indexOf(currentMonth) === 0) {
            totalPMTarget++;
            if (pStat === 'completed' || pStat === 'selesai') totalPMDone++;
          }
        }
      }
    }

    assignedTickets.reverse();
    closedDocuments.reverse();

    return {
      status: 'success',
      data: {
        wilayah: wilayahFSE,
        jumlahMesin: jumlahMesin,
        totalCM: totalCM,
        totalPM: totalPMTarget > 0 ? (totalPMDone + '/' + totalPMTarget) : '0',
        assignedTickets: assignedTickets,
        closedDocuments: closedDocuments
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
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
      if (bName && banks.indexOf(bName) === -1) banks.push(bName);
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
    if (sectionName === 'FSE') return getFSEData();
    if (sectionName === 'DatabaseMesin') return getMachinesData();
    if (sectionName === 'DatabaseCustomer') return getCustomersData();
    if (sectionName === 'DashboardMetrics') return getDashboardMetrics();
    if (sectionName === 'ReportingFSE') return getFSEReportingData();
    if (sectionName === 'DataPM') return getPMData();
    return { status: 'error', message: 'Seksi [' + sectionName + '] tidak dikenali' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
