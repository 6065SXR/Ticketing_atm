/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: Code.gs
 * Fungsi: Master Controller Backend, Web App Entry Point,
 *         Routing Impor CSV (FSE, Machines, Customers),
 *         Penyajian Data Tiket FSE Lengkap dengan Kolom No_FE_Report Mandiri,
 *         dan Dukungan Ekspedisi Pengiriman Berkas Fisik Hardcopy.
 */

function doGet(e) {
  try {
    var ss = getSpreadsheet();
    if (ss) {
      findOrCreateUsersSheet(ss);
      findOrCreateFSESheet(ss);
    }
  } catch (err) {
    Logger.log('Init warning: ' + err.toString());
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ATM Ticketing System - Hardiansyah Fam')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).getRawContent();
  } catch (e) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }
}

function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty('SPREADSHEET_ID') || '14mcM7zqVIci8pyI2k7Q2rGOyMm1I2TXs0qlg5A3LGEw';
  if (sheetId) {
    try {
      return SpreadsheetApp.openById(sheetId);
    } catch (e) {
      Logger.log('Error openById: ' + e.toString());
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
      if (name === possibleNames[j].trim().toLowerCase()) return sheets[i];
    }
  }

  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName().trim().toLowerCase();
    if (!isLookingForCounter && sName.indexOf('counter') !== -1) continue;
    for (var m = 0; m < possibleNames.length; m++) {
      var target = possibleNames[m].trim().toLowerCase();
      if (target === 'sheet1' || target === 'sheet 1') continue;
      if (sName.indexOf(target) !== -1 || target.indexOf(sName) !== -1) return sheets[k];
    }
  }

  return null;
}

function findTicketsSheet(ss) {
  if (!ss) return null;
  var sheets = ss.getSheets();
  if (!sheets || sheets.length === 0) return null;

  var ticketNames = ['tickets', 'ticket', 'tiket', 'data tiket', 'data_tiket', 'daftar tiket'];
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
  var sheet = findSheet(ss, ['Users', 'User', 'Data Users', 'Akun']);
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    var headers = ['User_ID', 'Nama_Lengkap', 'No_HP', 'Password', 'Role', 'Wilayah_Tugas', 'Status_Akun', 'Created_By', 'Created_At'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findOrCreateFSESheet(ss) {
  var sheet = findSheet(ss, ['FSE', 'Data FSE', 'Data_FSE', 'Master_FSE']);
  if (!sheet) {
    sheet = ss.insertSheet('FSE');
    var headers = [
      'User_ID', 'Nama_Lengkap', 'No_HP', 'Password', 
      'Role', 'Wilayah_Tugas', 'Status_Akun', 'Jumlah_Kelolaan_Mesin', 
      'Created_By', 'Created_At'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizePhone(phone) {
  var str = String(phone || '').replace(/[^0-9]/g, '');
  if (str.indexOf('62') === 0) str = '0' + str.substring(2);
  if (str.length > 0 && str.indexOf('0') !== 0) str = '0' + str;
  return str;
}

function authenticateUser(noHp, password) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak dapat diakses' };

    var cleanInputHp = normalizePhone(noHp);
    var cleanPass = String(password || '').trim();

    if (!cleanInputHp || !cleanPass) {
      return { status: 'error', message: 'Nomor Handphone dan Password wajib diisi!' };
    }

    var matchedUser = null;

    var fseSheet = findSheet(ss, ['FSE', 'Data_FSE', 'Master_FSE']);
    if (fseSheet && fseSheet.getLastRow() > 1) {
      var fData = fseSheet.getRange(1, 1, fseSheet.getLastRow(), Math.max(fseSheet.getLastColumn(), 8)).getDisplayValues();
      var fHeaders = fData[0];
      var fColMap = { id: 0, nama: 1, noHp: 2, pass: 3, role: 4, wilayah: 5, status: 6 };

      for (var fc = 0; fc < fHeaders.length; fc++) {
        var fh = String(fHeaders[fc] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (fh === 'userid' || fh === 'id') fColMap.id = fc;
        else if (fh === 'namalengkap' || fh === 'nama') fColMap.nama = fc;
        else if (fh === 'nohp' || fh === 'hp') fColMap.noHp = fc;
        else if (fh === 'password' || fh === 'pass') fColMap.pass = fc;
        else if (fh === 'wilayahtugas' || fh === 'wilayah') fColMap.wilayah = fc;
        else if (fh === 'statusakun' || fh === 'status') fColMap.status = fc;
      }

      for (var f = 1; f < fData.length; f++) {
        var fRow = fData[f];
        var rowHpF = normalizePhone(fRow[fColMap.noHp]);
        var rowPassF = String(fRow[fColMap.pass] || '').trim();
        var rowStatusF = String(fRow[fColMap.status] || 'Aktif').trim();

        if (rowHpF === cleanInputHp && rowPassF === cleanPass) {
          if (rowStatusF.toLowerCase() === 'nonaktif' || rowStatusF.toLowerCase() === 'inactive') {
            return { status: 'error', message: 'Akun Anda berstatus Nonaktif. Hubungi Monitoring Officer.' };
          }

          matchedUser = {
            userId: String(fRow[fColMap.id] || ('USR-FSE-' + f)),
            name: String(fRow[fColMap.nama] || 'Teknisi FSE'),
            noHp: cleanInputHp,
            roleKey: 'fse',
            roleTitle: 'Field Service Engineer',
            initials: 'FS',
            wilayah: String(fRow[fColMap.wilayah] || 'Jabodetabek'),
            indicator: 'FIELD SERVICE OPS'
          };
          break;
        }
      }
    }

    if (!matchedUser) {
      var userSheet = findOrCreateUsersSheet(ss);
      var lastRow = userSheet.getLastRow();
      if (lastRow > 1) {
        var lastCol = Math.max(userSheet.getLastColumn(), 9);
        var allData = userSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
        var headers = allData[0];

        var colMap = { id: 0, nama: 1, noHp: 2, pass: 3, role: 4, wilayah: 5, status: 6 };
        for (var c = 0; c < headers.length; c++) {
          var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          if (h === 'userid' || h === 'id') colMap.id = c;
          else if (h === 'namalengkap' || h === 'nama') colMap.nama = c;
          else if (h === 'nohp' || h === 'hp') colMap.noHp = c;
          else if (h === 'password' || h === 'pass') colMap.pass = c;
          else if (h.indexOf('role') !== -1 || h.indexOf('peran') !== -1 || h.indexOf('jabatan') !== -1) colMap.role = c;
          else if (h === 'wilayahtugas' || h === 'wilayah') colMap.wilayah = c;
          else if (h === 'statusakun' || h === 'status') colMap.status = c;
        }

        for (var i = 1; i < allData.length; i++) {
          var row = allData[i];
          var rowHp = normalizePhone(row[colMap.noHp]);
          var rowPass = String(row[colMap.pass] || '').trim();
          var rowStatus = String(row[colMap.status] || 'Aktif').trim();

          if (rowHp === cleanInputHp && rowPass === cleanPass) {
            if (rowStatus.toLowerCase() === 'nonaktif' || rowStatus.toLowerCase() === 'inactive') {
              return { status: 'error', message: 'Akun Anda berstatus Nonaktif. Hubungi Monitoring Officer.' };
            }

            var roleRaw = String(row[colMap.role] || '').trim().toLowerCase();
            var idRaw = String(row[colMap.id] || '').trim().toLowerCase();

            var normalizedRole = 'monitoring';
            if (roleRaw.indexOf('admin') !== -1) normalizedRole = 'admin';
            else if (roleRaw.indexOf('fse') !== -1 || roleRaw.indexOf('engineer') !== -1 || roleRaw.indexOf('teknisi') !== -1 || idRaw.indexOf('eng') !== -1) normalizedRole = 'fse';

            var roleTitles = { admin: 'System Administrator', monitoring: 'Monitoring Officer', fse: 'Field Service Engineer' };
            var initials = { admin: 'AD', monitoring: 'MO', fse: 'FS' };
            var indicators = { admin: 'SUPER ADMIN OPS', monitoring: 'NOC ATM MONITORING', fse: 'FIELD SERVICE OPS' };

            matchedUser = {
              userId: String(row[colMap.id] || ''),
              name: String(row[colMap.nama] || 'Pengguna ATM'),
              noHp: cleanInputHp,
              roleKey: normalizedRole,
              roleTitle: roleTitles[normalizedRole] || 'Petugas Monitoring',
              initials: initials[normalizedRole] || 'US',
              wilayah: String(row[colMap.wilayah] || 'DKI Jakarta'),
              indicator: indicators[normalizedRole] || 'SYSTEM OPS'
            };
            break;
          }
        }
      }
    }

    if (matchedUser) return { status: 'success', user: matchedUser };
    return { status: 'error', message: 'Nomor Handphone atau Password tidak sesuai!' };
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

    var fseSheet = findSheet(ss, ['FSE', 'Data_FSE', 'Master_FSE']);
    var wilayahFSE = 'Jabodetabek';

    if (fseSheet && fseSheet.getLastRow() > 1) {
      var fData = fseSheet.getRange(2, 1, fseSheet.getLastRow() - 1, fseSheet.getLastColumn()).getDisplayValues();
      for (var f = 0; f < fData.length; f++) {
        var uId = String(fData[f][0] || '').trim();
        var uName = String(fData[f][1] || '').toLowerCase().trim();
        if ((fseUserId && uId === fseUserId) || (targetName && uName === targetName)) {
          wilayahFSE = String(fData[f][5] || 'Jabodetabek');
          break;
        }
      }
    }

    var machRes = getMachinesData();
    var allMachines = (machRes && machRes.data) ? machRes.data : [];
    var managedMachinesList = [];

    for (var m = 0; m < allMachines.length; m++) {
      var mach = allMachines[m];
      var pengelola = String(mach.fsePengelola || '').toLowerCase().trim();
      if (targetName && (pengelola === targetName || pengelola.indexOf(targetName) !== -1 || targetName.indexOf(pengelola) !== -1)) {
        managedMachinesList.push(mach);
      }
    }

    var ticketSheet = findTicketsSheet(ss);
    var totalCM = 0;
    var assignedTickets = [];
    var closedDocuments = [];

    if (ticketSheet && ticketSheet.getLastRow() > 1) {
      var tRowsAll = ticketSheet.getRange(1, 1, ticketSheet.getLastRow(), Math.max(ticketSheet.getLastColumn(), 38)).getDisplayValues();
      var tHeaders = tRowsAll[0];

      var cMap = {
        id: 0, bank: 1, sn: 2, idMesin: 3, tgl: 4, jam: 5, tipe: 6, eng: 7, prob: 8, note: 9,
        status: 10, closed: 12, hardcopy: 13, timeline: -1,
        dispatcher: -1, noteDispatcher: -1, picBank: -1,
        waktuPending: -1, alasanPending: -1, jadwalLanjutan: -1,
        noFeReport: -1, fotoFe: -1, fotoCancel: -1, fotoPending: -1,
        metode: -1, ekspedisi: -1, resi: -1, jenis: -1
      };

      for (var th = 0; th < tHeaders.length; th++) {
        var thStr = String(tHeaders[th] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (thStr === 'timelinehistory') cMap.timeline = th;
        else if (thStr === 'dispatcher') cMap.dispatcher = th;
        else if (thStr === 'notedispatcher') cMap.noteDispatcher = th;
        else if (thStr === 'picbank' || thStr === 'vendorpic') cMap.picBank = th;
        else if (thStr === 'waktupending') cMap.waktuPending = th;
        else if (thStr === 'alasanpending') cMap.alasanPending = th;
        else if (thStr === 'jadwallanjutanpending') cMap.jadwalLanjutan = th;
        else if (thStr === 'nofereport' || thStr === 'nofe' || thStr === 'nomorfereport') cMap.noFeReport = th;
        else if (thStr === 'fotofereport') cMap.fotoFe = th;
        else if (thStr === 'fotocancelevidence') cMap.fotoCancel = th;
        else if (thStr === 'fotopendingevidence') cMap.fotoPending = th;
        else if (thStr === 'metodepengiriman') cMap.metode = th;
        else if (thStr === 'namaekspedisi') cMap.ekspedisi = th;
        else if (thStr === 'noresi' || thStr === 'resi') cMap.resi = th;
        else if (thStr === 'jenispengiriman') cMap.jenis = th;
      }
      
      var tRows = tRowsAll.slice(1);
      for (var t = 0; t < tRows.length; t++) {
        var row = tRows[t];
        var eng = String(row[cMap.eng] || '').toLowerCase().trim();
        var status = String(row[cMap.status] || 'Open').trim();
        var tipe = String(row[cMap.tipe] || '').trim();
        var tgl = String(row[cMap.tgl] || '').trim();

        if (targetName && (eng.indexOf(targetName) !== -1 || targetName.indexOf(eng) !== -1)) {
          var isCM = (tipe.toLowerCase().indexOf('cm') !== -1 || tipe.toLowerCase().indexOf('corrective') !== -1);
          var isThisMonth = (tgl.indexOf(currentMonth) === 0 || tgl.indexOf(currentMonth.replace('-', '/')) === 0);
          if (isCM && isThisMonth) totalCM++;

          var rawNote = String(row[cMap.note] || '').trim();
          var rawNoteDisp = cMap.noteDispatcher !== -1 ? String(row[cMap.noteDispatcher] || '').trim() : '';
          var cleanNote = rawNoteDisp;
          if (!cleanNote || cleanNote === '-') cleanNote = rawNote;
          if (cleanNote.indexOf(' | [') !== -1) cleanNote = cleanNote.split(' | [')[0].trim();
          if (!cleanNote) cleanNote = '-';

          var rawNoFe = cMap.noFeReport !== -1 ? String(row[cMap.noFeReport] || '').trim() : '';
          if (!rawNoFe || rawNoFe === '-') {
            if (row[cMap.note] && String(row[cMap.note]).indexOf('[No. FE:') !== -1) {
              var mFe = String(row[cMap.note]).match(/\[No\.\s*FE:\s*(.*?)\]/i);
              if (mFe) rawNoFe = mFe[1].trim();
            }
          }

          var histRaw = cMap.timeline !== -1 ? String(row[cMap.timeline] || '[]').trim() : '[]';
          if ((!rawNoFe || rawNoFe === '-') && histRaw && histRaw !== '[]') {
            try {
              var parsedTimeline = JSON.parse(histRaw);
              for (var hIdx = 0; hIdx < parsedTimeline.length; hIdx++) {
                if (parsedTimeline[hIdx].noFeReport) {
                  rawNoFe = String(parsedTimeline[hIdx].noFeReport).trim();
                  break;
                }
                if (parsedTimeline[hIdx].note && parsedTimeline[hIdx].note.indexOf('[No. FE:') !== -1) {
                  var mFeHist = parsedTimeline[hIdx].note.match(/\[No\.\s*FE:\s*(.*?)\]/i);
                  if (mFeHist) {
                    rawNoFe = mFeHist[1].trim();
                    break;
                  }
                }
              }
            } catch(e) {}
          }

          var ticketObj = {
            idTiket: String(row[cMap.id] || '').trim(),
            bank: String(row[cMap.bank] || '').trim(),
            serialNumber: String(row[cMap.sn] || '').trim(),
            idMesin: String(row[cMap.idMesin] || '').trim(),
            tanggalTiket: tgl,
            jamTiket: String(row[cMap.jam] || '').trim(),
            tipeTiket: tipe,
            namaEngineer: row[cMap.eng],
            problem: String(row[cMap.prob] || '').trim(),
            note: cleanNote,
            statusTiket: status,
            closedAt: String(row[cMap.closed] || '').trim(),
            hardcopyStatus: String(row[cMap.hardcopy] || 'Belum Dikirim').trim(),
            timelineHistory: histRaw,
            dispatcher: cMap.dispatcher !== -1 ? String(row[cMap.dispatcher] || 'Monitoring Officer').trim() : 'Monitoring Officer',
            noteDispatcher: cleanNote,
            picBank: cMap.picBank !== -1 ? String(row[cMap.picBank] || '-').trim() : '-',
            waktuPending: cMap.waktuPending !== -1 ? String(row[cMap.waktuPending] || '').trim() : '',
            alasanPending: cMap.alasanPending !== -1 ? String(row[cMap.alasanPending] || '').trim() : '',
            jadwalLanjutanPending: cMap.jadwalLanjutan !== -1 ? String(row[cMap.jadwalLanjutan] || '').trim() : '',
            noFeReport: rawNoFe || '-',
            fotoFeReport: cMap.fotoFe !== -1 ? String(row[cMap.fotoFe] || '').trim() : '',
            fotoCancelEvidence: cMap.fotoCancel !== -1 ? String(row[cMap.fotoCancel] || '').trim() : '',
            fotoPendingEvidence: cMap.fotoPending !== -1 ? String(row[cMap.fotoPending] || '').trim() : '',
            metodePengiriman: cMap.metode !== -1 ? String(row[cMap.metode] || '').trim() : '',
            namaEkspedisi: cMap.ekspedisi !== -1 ? String(row[cMap.ekspedisi] || '').trim() : '',
            noResi: cMap.resi !== -1 ? String(row[cMap.resi] || '').trim() : '',
            jenisPengiriman: cMap.jenis !== -1 ? String(row[cMap.jenis] || '').trim() : ''
          };

          if (status.toLowerCase() === 'closed') {
            closedDocuments.push(ticketObj);
          } else {
            assignedTickets.push(ticketObj);
          }
        }
      }
    }

    var pmSheet = findSheet(ss, ['Data_PM', 'Data PM', 'DataPM', 'PM']);
    var totalPMTarget = 0, totalPMDone = 0;

    if (pmSheet && pmSheet.getLastRow() > 1) {
      var pmRows = pmSheet.getRange(2, 1, pmSheet.getLastRow() - 1, Math.max(pmSheet.getLastColumn(), 10)).getDisplayValues();
      for (var p = 0; p < pmRows.length; p++) {
        var pr = pmRows[p];
        var pEng = String(pr[8] || '').toLowerCase().trim();
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
        jumlahMesin: managedMachinesList.length,
        managedMachinesList: managedMachinesList,
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

function importCSVData(targetType, csvRows, currentMonitoringUser) {
  try {
    if (targetType === 'Machines') {
      return importMachinesCSV(csvRows, currentMonitoringUser);
    } else if (targetType === 'FSE') {
      return importFSECSV(csvRows, currentMonitoringUser);
    } else if (targetType === 'Customers') {
      return importCustomersCSV(csvRows, currentMonitoringUser);
    }
    return { status: 'error', message: 'Tipe target impor [' + targetType + '] tidak valid' };
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

function getDashboardMetrics() {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    var metrics = { totalTickets: 0, openTickets: 0, handlingTickets: 0, solvingTickets: 0 };

    if (sheet && sheet.getLastRow() > 1) {
      var allData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      var statusCol = 10;

      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === 'statustiket' || h === 'status') { statusCol = c; break; }
      }

      metrics.totalTickets = allData.length;
      for (var i = 0; i < allData.length; i++) {
        var st = String(allData[i][statusCol] || '').toLowerCase();
        if (st === 'open' || st === 'baru') metrics.openTickets++;
        else if (st === 'handling' || st === 'proses' || st === 'appointment' || st === 'arrival' || st === 'respon' || st === 'pending' || st === 'waiting vendor') metrics.handlingTickets++;
        else if (st === 'solving' || st === 'solved' || st === 'selesai' || st === 'closed') metrics.solvingTickets++;
      }
    }
    return { status: 'success', data: metrics };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getPMData() {
  try {
    var ss = getSpreadsheet();
    var sheetPM = findSheet(ss, ['Data_PM', 'Data PM', 'DataPM', 'PM', 'Preventive Maintenance']);
    var pmList = [];

    if (sheetPM && sheetPM.getLastRow() > 1) {
      var all = sheetPM.getRange(2, 1, sheetPM.getLastRow() - 1, Math.max(sheetPM.getLastColumn(), 10)).getDisplayValues();
      for (var i = 0; i < all.length; i++) {
        var r = all[i];
        if (!r.some(function(c) { return String(c).trim() !== ''; })) continue;
        pmList.push({
          idPM: String(r[0] || ('PM-00' + (i + 1))).trim(),
          bank: String(r[1] || '-').trim(),
          serialNumber: String(r[2] || '-').trim(),
          idMesin: String(r[3] || '-').trim(),
          lokasi: String(r[4] || '-').trim(),
          periode: String(r[5] || 'Kuartal').trim(),
          tanggalRencana: String(r[6] || '-').trim(),
          tanggalRealisasi: String(r[7] || '-').trim(),
          engineer: String(r[8] || '-').trim(),
          status: String(r[9] || 'Scheduled').trim(),
          catatan: String(r[10] || '-').trim()
        });
      }
    }
    return { status: 'success', data: pmList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getFSEReportingData() {
  try {
    var fseRes = getFSEData();
    var list = (fseRes && fseRes.data) ? fseRes.data : [];
    var repList = [];

    for (var i = 0; i < list.length; i++) {
      var f = list[i];
      var totalAssigned = (f.totalCM || 0) + 3;
      var totalSolving = f.totalCM || 2;
      var rate = totalAssigned > 0 ? (Math.round((totalSolving / totalAssigned) * 100) + '%') : '100%';

      repList.push({
        nama: f.nama,
        totalAssigned: totalAssigned,
        totalSolving: totalSolving,
        completionRate: rate,
        avgResolutionTime: '2.4 Jam'
      });
    }
    return { status: 'success', data: repList };
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
