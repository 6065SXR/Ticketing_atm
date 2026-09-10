/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: code.gs
 * Fungsi: Master Backend Controller, Web App Entry Point, User Authentication (No HP & Password),
 *         FSE Account Registration by Monitoring, Ticket Lifecycle Engine, PM Controller, Master Data & Diagnostics.
 */

function doGet(e) {
  // Otomatis pastikan sheet Users sudah ada di Spreadsheet saat Web App dimuat
  try {
    var ss = getSpreadsheet();
    if (ss) {
      findOrCreateUsersSheet(ss);
    }
  } catch (err) {
    Logger.log('Init Users sheet warning: ' + err.toString());
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

function registerFSE(payload, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak ditemukan' };

    var creatorRole = (currentMonitoringUser && currentMonitoringUser.roleKey) ? currentMonitoringUser.roleKey.toLowerCase() : '';
    if (creatorRole !== 'monitoring' && creatorRole !== 'admin') {
      return { status: 'error', message: 'Hanya Role Monitoring (atau Admin) yang berhak mendaftarkan akun FSE baru!' };
    }

    var nama = String(payload.nama || '').trim();
    var noHp = String(payload.noHp || '').replace(/[^0-9]/g, '').trim();
    var password = String(payload.password || '').trim();
    var wilayah = String(payload.wilayah || 'DKI Jakarta').trim();

    if (!nama || !noHp || !password) {
      return { status: 'error', message: 'Nama Lengkap, No HP, dan Password wajib diisi!' };
    }

    var userSheet = findOrCreateUsersSheet(ss);
    var uLastRow = userSheet.getLastRow();
    if (uLastRow > 1) {
      var allUsers = userSheet.getRange(2, 1, uLastRow - 1, Math.max(userSheet.getLastColumn(), 5)).getDisplayValues();
      for (var u = 0; u < allUsers.length; u++) {
        var existingHp = String(allUsers[u][2] || '').replace(/[^0-9]/g, '').trim();
        if (existingHp === noHp) {
          return { status: 'error', message: 'Nomor HP [' + noHp + '] sudah terdaftar untuk pengguna lain!' };
        }
      }
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var nextUserSeq = (uLastRow > 1) ? ('00' + uLastRow).slice(-3) : '001';
    var newUserId = 'USR-' + nextUserSeq;
    var creatorName = (currentMonitoringUser && currentMonitoringUser.name) ? currentMonitoringUser.name : 'Monitoring Officer';

    userSheet.appendRow([
      newUserId, nama, noHp, password, 'FSE', wilayah, 'Aktif', creatorName, nowFormatted
    ]);

    var fseSheet = findSheet(ss, ['Engineers', 'Engineer', 'FSE', 'Data FSE', 'Teknisi']);
    var nextEngId = 'ENG-0' + (uLastRow);
    if (!fseSheet) {
      fseSheet = ss.insertSheet('Engineers');
      fseSheet.appendRow(['ID_Engineer', 'Nama_Engineer', 'Wilayah_Tugas', 'No_HP', 'Status_Aktif']);
    }

    var fLastRow = fseSheet.getLastRow();
    nextEngId = 'ENG-0' + Math.max(fLastRow, 1);

    fseSheet.appendRow([
      nextEngId, nama, wilayah, noHp, 'Aktif'
    ]);

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Akun FSE [' + nama + '] berhasil didaftarkan dan aktif!',
      fse: {
        id: nextEngId,
        nama: nama,
        noHp: noHp,
        wilayah: wilayah,
        status: 'Aktif'
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function generateNewTicketNumber() {
  try {
    var ss = getSpreadsheet();
    var counterSheet = findSheet(ss, ['Counter_Ticket', 'CounterTicket', 'Counter']);
    
    if (!counterSheet) {
      counterSheet = ss.insertSheet('Counter_Ticket');
      counterSheet.appendRow(['Tanggal', 'Last_Sequence']);
    }

    var now = new Date();
    var currentYYMMDD = Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd');
    var nextSequence = 1;

    var lastRow = counterSheet.getLastRow();
    if (lastRow > 1) {
      var counterData = counterSheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
      var foundToday = false;

      for (var i = 0; i < counterData.length; i++) {
        var recordedDate = counterData[i][0].toString().trim();
        if (recordedDate === currentYYMMDD) {
          foundToday = true;
          var currentSeq = parseInt(counterData[i][1], 10) || 0;
          nextSequence = currentSeq + 1;
          counterSheet.getRange(i + 2, 2).setValue(nextSequence);
          break;
        }
      }

      if (!foundToday) {
        counterSheet.appendRow([currentYYMMDD, 1]);
        nextSequence = 1;
      }
    } else {
      counterSheet.appendRow([currentYYMMDD, 1]);
      nextSequence = 1;
    }

    SpreadsheetApp.flush();

    var seqFormatted = ('0000' + nextSequence).slice(-4);
    var fullTicketId = currentYYMMDD + '-' + seqFormatted;
    var dateFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
    var timeFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');

    return {
      status: 'success',
      ticketId: fullTicketId,
      dateString: dateFormatted,
      timeString: timeFormatted
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function createTicket(payload) {
  try {
    var ss = getSpreadsheet();
    var ticketSheet = findTicketsSheet(ss);
    if (!ticketSheet) {
      ticketSheet = ss.insertSheet('Tickets');
      ticketSheet.appendRow([
        'ID_Tiket', 'Bank', 'Serial_Number', 'ID_Mesin', 
        'Tanggal_Tiket', 'Jam_Tiket', 'Tipe_Tiket', 'Nama_Engineer', 
        'Problem', 'Note', 'Status_Tiket', 'Updated_At'
      ]);
    }

    var now = new Date();
    var updatedAt = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var idTiket = String(payload.idTiket || '').trim();
    var bank = String(payload.bank || '').trim();
    var serialNumber = String(payload.serialNumber || '').trim();
    var idMesin = String(payload.idMesin || '').trim();
    var tanggalTiket = String(payload.tanggalTiket || '').trim();
    var jamTiket = String(payload.jamTiket || '').trim();
    var tipeTiket = String(payload.tipeTiket || '').trim();
    var namaEngineer = String(payload.namaEngineer || '').trim();
    var problem = String(payload.problem || '').trim();
    var note = String(payload.note || '').trim();
    var statusTiket = 'Open';

    var newRow = [
      idTiket, bank, serialNumber, idMesin,
      tanggalTiket, jamTiket, tipeTiket, namaEngineer,
      problem, note, statusTiket, updatedAt
    ];

    ticketSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Tiket ' + idTiket + ' berhasil disimpan!',
      ticket: {
        idTiket: idTiket,
        bank: bank,
        serialNumber: serialNumber,
        idMesin: idMesin,
        tanggalTiket: tanggalTiket,
        jamTiket: jamTiket,
        tipeTiket: tipeTiket,
        namaEngineer: namaEngineer,
        problem: problem,
        note: note,
        statusTiket: statusTiket,
        updatedAt: updatedAt
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getTicketsPaginated(page, pageSize, searchQuery, statusFilter, bankFilter) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        status: 'success',
        tickets: [],
        page: 1,
        pageSize: pageSize || 10,
        totalItems: 0,
        totalPages: 0
      };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
    var headerRowIdx = 0;
    var headers = allData[0];
    var isHeader0 = false;
    for (var h0 = 0; h0 < headers.length; h0++) {
      var strH0 = String(headers[h0] || '').toLowerCase();
      if (strH0.indexOf('tiket') !== -1 || strH0.indexOf('bank') !== -1 || strH0.indexOf('id') !== -1 || strH0.indexOf('problem') !== -1) {
        isHeader0 = true;
        break;
      }
    }
    if (!isHeader0 && allData.length > 2) {
      for (var h1 = 0; h1 < allData[1].length; h1++) {
        var strH1 = String(allData[1][h1] || '').toLowerCase();
        if (strH1.indexOf('tiket') !== -1 || strH1.indexOf('bank') !== -1 || strH1.indexOf('id') !== -1 || strH1.indexOf('problem') !== -1) {
          headerRowIdx = 1;
          headers = allData[1];
          break;
        }
      }
    }

    var colMap = {
      idTiket: -1, bank: -1, serialNumber: -1, idMesin: -1,
      tanggalTiket: -1, jamTiket: -1, tipeTiket: -1, namaEngineer: -1,
      problem: -1, note: -1, statusTiket: -1, updatedAt: -1
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'nomortiket' || h === 'ticketid' || h === 'id' || h === 'no') colMap.idTiket = c;
      else if (h === 'bank' || h === 'namabank' || h === 'customer' || h === 'mitra') colMap.bank = c;
      else if (h === 'serialnumber' || h === 'sn' || h === 'noserial' || h === 'noseri') colMap.serialNumber = c;
      else if (h === 'idmesin' || h === 'machineid' || h === 'mesin' || h === 'atmid' || h === 'nomesin') colMap.idMesin = c;
      else if (h === 'tanggaltiket' || h === 'tanggal' || h === 'date' || h === 'tgl') colMap.tanggalTiket = c;
      else if (h === 'jamtiket' || h === 'jam' || h === 'time' || h === 'waktu') colMap.jamTiket = c;
      else if (h === 'tipetiket' || h === 'tipe' || h === 'type' || h === 'jenis') colMap.tipeTiket = c;
      else if (h === 'namaengineer' || h === 'engineer' || h === 'fse' || h === 'teknisi' || h === 'pic') colMap.namaEngineer = c;
      else if (h === 'problem' || h === 'masalah' || h === 'keluhan' || h === 'kendala') colMap.problem = c;
      else if (h === 'note' || h === 'catatan' || h === 'keterangan' || h === 'notes') colMap.note = c;
      else if (h === 'statustiket' || h === 'status' || h === 'kondisi') colMap.statusTiket = c;
      else if (h === 'updatedat' || h === 'update' || h === 'waktuupdate' || h === 'lastupdate') colMap.updatedAt = c;
    }

    var maxIdx = Math.max(headers.length - 1, 0);
    if (colMap.idTiket === -1) colMap.idTiket = 0;
    if (colMap.bank === -1) colMap.bank = Math.min(1, maxIdx);
    if (colMap.serialNumber === -1) colMap.serialNumber = Math.min(2, maxIdx);
    if (colMap.idMesin === -1) colMap.idMesin = Math.min(3, maxIdx);
    if (colMap.tanggalTiket === -1) colMap.tanggalTiket = Math.min(4, maxIdx);
    if (colMap.jamTiket === -1) colMap.jamTiket = Math.min(5, maxIdx);
    if (colMap.tipeTiket === -1) colMap.tipeTiket = Math.min(6, maxIdx);
    if (colMap.namaEngineer === -1) colMap.namaEngineer = Math.min(7, maxIdx);
    if (colMap.problem === -1) colMap.problem = Math.min(8, maxIdx);
    if (colMap.note === -1) colMap.note = Math.min(9, maxIdx);
    if (colMap.statusTiket === -1) colMap.statusTiket = Math.min(10, maxIdx);
    if (colMap.updatedAt === -1) colMap.updatedAt = Math.min(11, maxIdx);

    var rawData = allData.slice(headerRowIdx + 1);
    var filtered = [];
    var searchLower = (searchQuery || '').toLowerCase().trim();
    var statusFilterVal = (statusFilter || '').trim();
    var bankFilterVal = (bankFilter || '').trim();

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      var hasData = false;
      for (var k = 0; k < row.length; k++) {
        if (String(row[k] || '').trim() !== '') {
          hasData = true;
          break;
        }
      }
      if (!hasData) continue;

      var idVal = colMap.idTiket >= 0 ? String(row[colMap.idTiket] || '').trim() : '';
      if (!idVal) idVal = 'TK-' + (i + 1);

      var item = {
        idTiket: idVal,
        bank: colMap.bank >= 0 ? String(row[colMap.bank] || '-').trim() : '-',
        serialNumber: colMap.serialNumber >= 0 ? String(row[colMap.serialNumber] || '-').trim() : '-',
        idMesin: colMap.idMesin >= 0 ? String(row[colMap.idMesin] || '-').trim() : '-',
        tanggalTiket: colMap.tanggalTiket >= 0 ? String(row[colMap.tanggalTiket] || '-').trim() : '-',
        jamTiket: colMap.jamTiket >= 0 ? String(row[colMap.jamTiket] || '-').trim() : '-',
        tipeTiket: colMap.tipeTiket >= 0 ? String(row[colMap.tipeTiket] || 'CM').trim() : 'CM',
        namaEngineer: colMap.namaEngineer >= 0 ? String(row[colMap.namaEngineer] || '-').trim() : '-',
        problem: colMap.problem >= 0 ? String(row[colMap.problem] || '-').trim() : '-',
        note: colMap.note >= 0 ? String(row[colMap.note] || '').trim() : '',
        statusTiket: colMap.statusTiket >= 0 ? String(row[colMap.statusTiket] || 'Open').trim() : 'Open',
        updatedAt: colMap.updatedAt >= 0 ? String(row[colMap.updatedAt] || '-').trim() : '-'
      };

      if (!item.statusTiket) item.statusTiket = 'Open';

      if (statusFilterVal && statusFilterVal !== 'ALL') {
        if (item.statusTiket.toLowerCase() !== statusFilterVal.toLowerCase()) continue;
      }

      if (bankFilterVal && bankFilterVal !== 'ALL') {
        var b1 = item.bank.toLowerCase();
        var b2 = bankFilterVal.toLowerCase();
        if (b1 !== b2 && b1.indexOf(b2) === -1 && b2.indexOf(b1) === -1) continue;
      }

      if (searchLower !== '') {
        var match = (
          item.idTiket.toLowerCase().indexOf(searchLower) !== -1 ||
          item.bank.toLowerCase().indexOf(searchLower) !== -1 ||
          item.serialNumber.toLowerCase().indexOf(searchLower) !== -1 ||
          item.idMesin.toLowerCase().indexOf(searchLower) !== -1 ||
          item.problem.toLowerCase().indexOf(searchLower) !== -1 ||
          item.namaEngineer.toLowerCase().indexOf(searchLower) !== -1
        );
        if (!match) continue;
      }

      filtered.push(item);
    }

    filtered.reverse(); // Tiket terbaru di posisi paling atas

    var totalItems = filtered.length;
    var validPageSize = Math.max(1, parseInt(pageSize, 10) || 10);
    var totalPages = Math.ceil(totalItems / validPageSize) || 1;
    var currentPage = Math.min(Math.max(1, parseInt(page, 10) || 1), totalPages);

    var startIndex = (currentPage - 1) * validPageSize;
    var endIndex = Math.min(startIndex + validPageSize, totalItems);
    var pagedData = filtered.slice(startIndex, endIndex);

    return {
      status: 'success',
      tickets: pagedData,
      page: currentPage,
      pageSize: validPageSize,
      totalItems: totalItems,
      totalPages: totalPages
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function updateTicketStatus(ticketId, newStatus, additionalNote) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    if (!sheet) return { status: 'error', message: 'Sheet Tickets tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data tiket' };

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var idCol = 0;
    var statusCol = 10;
    var noteCol = 9;
    var updateCol = 11;

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'nomortiket' || h === 'ticketid' || h === 'id') idCol = c;
      else if (h === 'statustiket' || h === 'status' || h === 'statusticket' || h === 'kondisi') statusCol = c;
      else if (h === 'note' || h === 'catatan' || h === 'keterangan') noteCol = c;
      else if (h === 'updatedat' || h === 'update' || h === 'waktuupdate') updateCol = c;
    }

    var targetRowIndex = -1;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][idCol]).trim() === String(ticketId).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { status: 'error', message: 'Tiket ' + ticketId + ' tidak ditemukan di database' };
    }

    var now = new Date();
    var updatedAt = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    sheet.getRange(targetRowIndex, statusCol + 1).setValue(newStatus);
    if (updateCol < lastCol) {
      sheet.getRange(targetRowIndex, updateCol + 1).setValue(updatedAt);
    }

    if (additionalNote && String(additionalNote).trim() !== '' && noteCol < lastCol) {
      var curNote = sheet.getRange(targetRowIndex, noteCol + 1).getDisplayValue();
      var combinedNote = curNote ? (curNote + ' | [' + newStatus + ']: ' + additionalNote) : ('[' + newStatus + ']: ' + additionalNote);
      sheet.getRange(targetRowIndex, noteCol + 1).setValue(combinedNote);
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      ticketId: ticketId,
      newStatus: newStatus,
      updatedAt: updatedAt
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getPMData() {
  try {
    var ss = getSpreadsheet();
    var sheetPM = findSheet(ss, [
      'Data_PM', 'Data PM', 'DataPM', 'PM', 
      'Preventive Maintenance', 'Preventive_Maintenance', 
      'Jadwal PM', 'Jadwal_PM', 'Preventive'
    ]);
    var pmList = [];

    if (sheetPM && sheetPM.getLastRow() > 1) {
      var lastRow = sheetPM.getLastRow();
      var lastCol = Math.max(sheetPM.getLastColumn(), 1);
      var allData = sheetPM.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var headers = allData[0];

      var colMap = {
        idPM: -1, idMesin: -1, serialNumber: -1, bank: -1, lokasi: -1,
        periode: -1, tanggalRencana: -1, tanggalRealisasi: -1, engineer: -1,
        status: -1, catatan: -1
      };

      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === 'idpm' || h === 'id' || h === 'nopm' || h === 'nomorpm') colMap.idPM = c;
        else if (h === 'idmesin' || h === 'atmid' || h === 'machineid' || h === 'nomesin') colMap.idMesin = c;
        else if (h === 'serialnumber' || h === 'sn' || h === 'noserial' || h === 'noseri') colMap.serialNumber = c;
        else if (h === 'bank' || h === 'namabank' || h === 'customer' || h === 'mitra') colMap.bank = c;
        else if (h === 'lokasi' || h === 'lokasiatm' || h === 'alamat' || h === 'area') colMap.lokasi = c;
        else if (h === 'periode' || h === 'siklus' || h === 'kuartal' || h === 'bulan') colMap.periode = c;
        else if (h === 'tanggalrencana' || h === 'jadwalpm' || h === 'jadwal' || h === 'tglrencana' || h === 'target') colMap.tanggalRencana = c;
        else if (h === 'tanggalrealisasi' || h === 'tglselesai' || h === 'realisasi' || h === 'tglpm') colMap.tanggalRealisasi = c;
        else if (h === 'engineer' || h === 'namaengineer' || h === 'fse' || h === 'teknisi' || h === 'pic') colMap.engineer = c;
        else if (h === 'status' || h === 'statuspm' || h === 'kondisi') colMap.status = c;
        else if (h === 'catatan' || h === 'keterangan' || h === 'note' || h === 'remark') colMap.catatan = c;
      }

      var rows = allData.slice(1);
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!r.some(function(cell) { return String(cell).trim() !== ''; })) continue;

        pmList.push({
          idPM: colMap.idPM >= 0 ? String(r[colMap.idPM] || '').trim() : ('PM-0' + (i + 1)),
          idMesin: colMap.idMesin >= 0 ? String(r[colMap.idMesin] || '-').trim() : '-',
          serialNumber: colMap.serialNumber >= 0 ? String(r[colMap.serialNumber] || '-').trim() : '-',
          bank: colMap.bank >= 0 ? String(r[colMap.bank] || '-').trim() : '-',
          lokasi: colMap.lokasi >= 0 ? String(r[colMap.lokasi] || '-').trim() : '-',
          periode: colMap.periode >= 0 ? String(r[colMap.periode] || 'Q3-2026').trim() : 'Q3-2026',
          tanggalRencana: colMap.tanggalRencana >= 0 ? String(r[colMap.tanggalRencana] || '-').trim() : '-',
          tanggalRealisasi: colMap.tanggalRealisasi >= 0 ? String(r[colMap.tanggalRealisasi] || '-').trim() : '-',
          engineer: colMap.engineer >= 0 ? String(r[colMap.engineer] || '-').trim() : '-',
          status: colMap.status >= 0 ? String(r[colMap.status] || 'Scheduled').trim() : 'Scheduled',
          catatan: colMap.catatan >= 0 ? String(r[colMap.catatan] || '').trim() : ''
        });
      }
    }

    return { status: 'success', data: pmList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getFSEData() {
  try {
    var ss = getSpreadsheet();
    var sheetFSE = findSheet(ss, ['Engineers', 'Engineer', 'FSE', 'Data FSE', 'Teknisi']);
    var fseList = [];

    if (sheetFSE && sheetFSE.getLastRow() > 1) {
      var lastRow = sheetFSE.getLastRow();
      var lastCol = Math.max(sheetFSE.getLastColumn(), 1);
      var allData = sheetFSE.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var headers = allData[0];

      var colMap = { id: 0, nama: 1, wilayah: 2, noHp: 3, status: 4 };
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === 'idengineer' || h === 'id' || h === 'idfse') colMap.id = c;
        else if (h === 'namaengineer' || h === 'nama' || h === 'namafse' || h === 'namateknisi') colMap.nama = c;
        else if (h === 'wilayahtugas' || h === 'wilayah' || h === 'area') colMap.wilayah = c;
        else if (h === 'nohp' || h === 'hp' || h === 'telepon' || h === 'phone') colMap.noHp = c;
        else if (h === 'statusaktif' || h === 'status') colMap.status = c;
      }

      var rows = allData.slice(1);
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!r.some(function(cell) { return String(cell).trim() !== ''; })) continue;
        fseList.push({
          id: String(r[colMap.id] || ('ENG-0' + (i + 1))).trim(),
          nama: String(r[colMap.nama] || '').trim(),
          wilayah: String(r[colMap.wilayah] || '-').trim(),
          noHp: String(r[colMap.noHp] || '-').trim(),
          status: String(r[colMap.status] || 'Aktif').trim()
        });
      }
    }

    return { status: 'success', data: fseList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getMachinesData() {
  try {
    var ss = getSpreadsheet();
    var sheetMach = findSheet(ss, ['Machines', 'Machine', 'Data Mesin', 'Mesin', 'ATM']);
    var machList = [];

    if (sheetMach && sheetMach.getLastRow() > 1) {
      var lastRow = sheetMach.getLastRow();
      var lastCol = Math.max(sheetMach.getLastColumn(), 1);
      var allData = sheetMach.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var headers = allData[0];

      var colMap = { idMesin: 0, serialNumber: 1, bank: 2, lokasi: 3, tipe: 4, status: 5 };
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[\s_-]/g, '');
        if (h === 'idmesin' || h === 'id' || h === 'machineid') colMap.idMesin = c;
        else if (h === 'serialnumber' || h === 'sn' || h === 'noserial') colMap.serialNumber = c;
        else if (h === 'namabank' || h === 'bank') colMap.bank = c;
        else if (h === 'lokasiatm' || h === 'lokasi' || h === 'alamat') colMap.lokasi = c;
        else if (h === 'tipemesin' || h === 'tipe' || h === 'model' || h === 'merk') colMap.tipe = c;
        else if (h === 'statusmesin' || h === 'status') colMap.status = c;
      }

      var rows = allData.slice(1);
      for (var j = 0; j < rows.length; j++) {
        var r = rows[j];
        if (!r.some(function(cell) { return String(cell).trim() !== ''; })) continue;
        machList.push({
          idMesin: String(r[colMap.idMesin] || ('ATM-0' + (j + 1))).trim(),
          serialNumber: String(r[colMap.serialNumber] || '').trim(),
          bank: String(r[colMap.bank] || '').trim(),
          lokasi: String(r[colMap.lokasi] || '-').trim(),
          tipe: String(r[colMap.tipe] || '-').trim(),
          status: String(r[colMap.status] || 'Active').trim()
        });
      }
    }

    return { status: 'success', data: machList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getCustomersData() {
  try {
    var ss = getSpreadsheet();
    var sheetCust = findSheet(ss, ['Customers', 'Customer', 'Data Customer', 'Data Bank', 'Bank']);
    var custList = [];

    if (sheetCust && sheetCust.getLastRow() > 1) {
      var lastRow = sheetCust.getLastRow();
      var lastCol = Math.max(sheetCust.getLastColumn(), 1);
      var allData = sheetCust.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var headers = allData[0];

      var colMap = { idCustomer: 0, namaBank: 1, kontakPic: 2, teleponPic: 3, emailBank: 4, alamat: 5 };
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[\s_-]/g, '');
        if (h === 'idcustomer' || h === 'id' || h === 'custid') colMap.idCustomer = c;
        else if (h === 'namabank' || h === 'bank') colMap.namaBank = c;
        else if (h === 'kontakpic' || h === 'pic' || h === 'namapic') colMap.kontakPic = c;
        else if (h === 'teleponpic' || h === 'telepon' || h === 'phone' || h === 'nohp') colMap.teleponPic = c;
        else if (h === 'emailbank' || h === 'email') colMap.emailBank = c;
        else if (h === 'alamat' || h === 'kantor' || h === 'address') colMap.alamat = c;
      }

      var rows = allData.slice(1);
      for (var k = 0; k < rows.length; k++) {
        var r = rows[k];
        if (!r.some(function(cell) { return String(cell).trim() !== ''; })) continue;
        custList.push({
          idCustomer: String(r[colMap.idCustomer] || ('CUST-0' + (k + 1))).trim(),
          namaBank: String(r[colMap.namaBank] || '').trim(),
          kontakPic: String(r[colMap.kontakPic] || '-').trim(),
          teleponPic: String(r[colMap.teleponPic] || '-').trim(),
          emailBank: String(r[colMap.emailBank] || '-').trim(),
          alamat: String(r[colMap.alamat] || '-').trim()
        });
      }
    }

    return { status: 'success', data: custList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getDashboardMetrics() {
  try {
    var ss = getSpreadsheet();
    var ticketSheet = findTicketsSheet(ss);
    
    var metrics = {
      totalTickets: 0,
      openTickets: 0,
      handlingTickets: 0,
      solvingTickets: 0,
      cancelTickets: 0,
      slaRate: '98.4%'
    };

    if (ticketSheet && ticketSheet.getLastRow() > 1) {
      var lastRow = ticketSheet.getLastRow();
      var lastCol = Math.max(ticketSheet.getLastColumn(), 1);
      var allData = ticketSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      
      var headerRowIdx = 0;
      var headers = allData[0];
      for (var h = 0; h < headers.length; h++) {
        var strH = String(headers[h] || '').toLowerCase();
        if (strH.indexOf('tiket') !== -1 || strH.indexOf('bank') !== -1 || strH.indexOf('problem') !== -1) {
          break;
        }
        if (h === headers.length - 1 && allData.length > 2) {
          headerRowIdx = 1;
          headers = allData[1];
        }
      }

      var statusColIdx = -1;
      for (var c = 0; c < headers.length; c++) {
        var hn = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (hn === 'statustiket' || hn === 'status' || hn === 'statusticket' || hn === 'kondisi' || hn === 'state') {
          statusColIdx = c;
          break;
        }
      }
      if (statusColIdx === -1) {
        statusColIdx = Math.min(10, headers.length - 1);
      }

      var rows = allData.slice(headerRowIdx + 1);
      var countValid = 0;
      for (var t = 0; t < rows.length; t++) {
        var r = rows[t];
        var hasContent = false;
        for (var cell = 0; cell < r.length; cell++) {
          if (String(r[cell] || '').trim() !== '') {
            hasContent = true;
            break;
          }
        }
        if (!hasContent) continue;
        countValid++;

        var st = String(r[statusColIdx] || '').trim().toLowerCase();
        if (st === 'open' || st === 'respon' || st === 'appointment' || st === 'baru' || st === 'pending' || st === 'antre') {
          metrics.openTickets++;
        } else if (st === 'handling' || st === 'take a part' || st === 'proses' || st === 'in progress') {
          metrics.handlingTickets++;
        } else if (st === 'solving' || st === 'solved' || st === 'selesai' || st === 'close' || st === 'closed' || st === 'done') {
          metrics.solvingTickets++;
        } else if (st === 'cancel' || st === 'batal') {
          metrics.cancelTickets++;
        } else {
          metrics.openTickets++;
        }
      }
      metrics.totalTickets = countValid;
    }

    return { status: 'success', data: metrics };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getFSEReportingData() {
  try {
    var ss = getSpreadsheet();
    var fseRes = getFSEData();
    var engineers = (fseRes && fseRes.data) ? fseRes.data : [];

    var ticketSheet = findTicketsSheet(ss);
    var ticketData = [];
    if (ticketSheet && ticketSheet.getLastRow() > 1) {
      var lastCol = Math.max(ticketSheet.getLastColumn(), 1);
      ticketData = ticketSheet.getRange(2, 1, ticketSheet.getLastRow() - 1, lastCol).getDisplayValues();
    }

    var reportList = [];
    for (var i = 0; i < engineers.length; i++) {
      var eng = engineers[i];
      var totalAssigned = 0;
      var totalSolving = 0;

      for (var j = 0; j < ticketData.length; j++) {
        var rowStr = ticketData[j].join(' ').toLowerCase();
        if (eng.nama && rowStr.indexOf(eng.nama.toLowerCase()) !== -1) {
          totalAssigned++;
          if (rowStr.indexOf('solving') !== -1 || rowStr.indexOf('selesai') !== -1 || rowStr.indexOf('solved') !== -1) {
            totalSolving++;
          }
        }
      }

      var rate = totalAssigned > 0 ? Math.round((totalSolving / totalAssigned) * 100) : 100;
      reportList.push({
        id: eng.id,
        nama: eng.nama,
        wilayah: eng.wilayah,
        noHp: eng.noHp,
        totalAssigned: totalAssigned,
        totalSolving: totalSolving,
        completionRate: rate + '%',
        avgResolutionTime: '2.4 Jam'
      });
    }

    return { status: 'success', data: reportList };
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

function checkSystemHealth() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak dapat dihubungi.' };

    var sheets = ss.getSheets();
    var sheetSummary = [];
    for (var i = 0; i < sheets.length; i++) {
      sheetSummary.push({
        name: sheets[i].getName(),
        rows: sheets[i].getLastRow(),
        cols: sheets[i].getLastColumn()
      });
    }

    return {
      status: 'healthy',
      spreadsheetName: ss.getName(),
      totalSheets: sheets.length,
      sheets: sheetSummary,
      serverTime: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
