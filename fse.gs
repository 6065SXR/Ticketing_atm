/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: fse.gs
 * Fungsi: Master Controller Sheet "FSE", Import CSV Teknisi,
 *         Sinkronisasi Sheet "Users", Perhitungan Metrik CM & PM Bulanan, dan CRUD FSE.
 */

function getFSEData() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Database Spreadsheet tidak dapat diakses' };

    var sheetFSE = findOrCreateFSESheet(ss);
    var fseList = [];

    var now = new Date();
    var currentMonth = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM');

    // Ambil data mesin untuk menghitung kelolaan mesin aktual
    var machRes = getMachinesData();
    var allMachines = (machRes && machRes.data) ? machRes.data : [];

    // Ambil data tiket untuk kalkulasi CM bulan ini
    var ticketSheet = findTicketsSheet(ss);
    var ticketRows = [];
    var tColMap = { engineer: -1, tipe: -1, tanggal: -1 };

    if (ticketSheet && ticketSheet.getLastRow() > 1) {
      var tAll = ticketSheet.getRange(1, 1, ticketSheet.getLastRow(), Math.max(ticketSheet.getLastColumn(), 1)).getDisplayValues();
      var tHeaders = tAll[0];
      for (var th = 0; th < tHeaders.length; th++) {
        var thStr = String(tHeaders[th] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (thStr === 'namaengineer' || thStr === 'engineer' || thStr === 'fse' || thStr === 'teknisi') tColMap.engineer = th;
        else if (thStr === 'tipetiket' || thStr === 'tipe' || thStr === 'type') tColMap.tipe = th;
        else if (thStr === 'tanggaltiket' || thStr === 'tanggal' || thStr === 'date') tColMap.tanggal = th;
      }
      ticketRows = tAll.slice(1);
    }

    // Ambil data PM untuk capaian PM bulan ini
    var sheetPM = findSheet(ss, ['Data_PM', 'Data PM', 'DataPM', 'PM', 'Preventive Maintenance']);
    var pmRows = [];
    var pColMap = { engineer: -1, status: -1, tanggalRencana: -1, tanggalRealisasi: -1 };

    if (sheetPM && sheetPM.getLastRow() > 1) {
      var pAll = sheetPM.getRange(1, 1, sheetPM.getLastRow(), Math.max(sheetPM.getLastColumn(), 1)).getDisplayValues();
      var pHeaders = pAll[0];
      for (var ph = 0; ph < pHeaders.length; ph++) {
        var phStr = String(pHeaders[ph] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (phStr === 'engineer' || phStr === 'namaengineer' || phStr === 'fse') pColMap.engineer = ph;
        else if (phStr === 'status' || phStr === 'statuspm') pColMap.status = ph;
        else if (phStr === 'tanggalrencana' || phStr === 'jadwalpm') pColMap.tanggalRencana = ph;
        else if (phStr === 'tanggalrealisasi' || phStr === 'realisasi') pColMap.tanggalRealisasi = ph;
      }
      pmRows = pAll.slice(1);
    }

    var fLastRow = sheetFSE.getLastRow();
    if (fLastRow > 1) {
      var fAll = sheetFSE.getRange(1, 1, fLastRow, Math.max(sheetFSE.getLastColumn(), 10)).getDisplayValues();
      var rows = fAll.slice(1);

      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!r.some(function(c) { return String(c).trim() !== ''; })) continue;

        var userId = String(r[0] || ('USR-00' + (i + 1))).trim();
        var namaFSE = String(r[1] || '').trim();
        var noHp = String(r[2] || '').trim();
        var wilayah = String(r[5] || '-').trim();
        var status = String(r[6] || 'Aktif').trim();
        var staticKelolaan = parseInt(r[7], 10) || 0;

        // Hitung kelolaan mesin aktual dari Machines table (jika ada data nama pengelola)
        var actualKelolaan = 0;
        var namaLower = namaFSE.toLowerCase();
        for (var m = 0; m < allMachines.length; m++) {
          if (allMachines[m].fsePengelola && allMachines[m].fsePengelola.toLowerCase() === namaLower) {
            actualKelolaan++;
          }
        }
        var jumlahMesin = actualKelolaan > 0 ? actualKelolaan : staticKelolaan;

        // Total CM bulan ini
        var totalCM = 0;
        for (var t = 0; t < ticketRows.length; t++) {
          var tr = ticketRows[t];
          var tEng = String(tr[tColMap.engineer] || '').toLowerCase();
          var tTipe = String(tr[tColMap.tipe] || '').toLowerCase();
          var tTgl = String(tr[tColMap.tanggal] || '');

          if (namaLower && (tEng.indexOf(namaLower) !== -1 || namaLower.indexOf(tEng) !== -1)) {
            var isCM = (tTipe.indexOf('cm') !== -1 || tTipe.indexOf('corrective') !== -1);
            var isThisMonth = (tTgl.indexOf(currentMonth) === 0 || tTgl.indexOf(currentMonth.replace('-', '/')) === 0);
            if (isCM && isThisMonth) totalCM++;
          }
        }

        // Capaian PM bulan ini
        var pmTarget = 0, pmCompleted = 0;
        for (var p = 0; p < pmRows.length; p++) {
          var pr = pmRows[p];
          var pEng = String(pr[pColMap.engineer] || '').toLowerCase();
          var pStat = String(pr[pColMap.status] || '').toLowerCase();
          var pTglRencana = String(pr[pColMap.tanggalRencana] || '');
          var pTglRealisasi = String(pr[pColMap.tanggalRealisasi] || '');

          if (namaLower && (pEng.indexOf(namaLower) !== -1 || namaLower.indexOf(pEng) !== -1)) {
            if (pTglRencana.indexOf(currentMonth) === 0 || pTglRealisasi.indexOf(currentMonth) === 0) {
              pmTarget++;
              if (pStat === 'completed' || pStat === 'selesai') pmCompleted++;
            }
          }
        }

        var pmAchievementStr = pmTarget > 0 ? (pmCompleted + ' / ' + pmTarget + ' (' + Math.round((pmCompleted / pmTarget) * 100) + '%)') : '0 / 0 (100%)';

        fseList.push({
          userId: userId,
          nama: namaFSE,
          noHp: noHp,
          wilayah: wilayah,
          status: status,
          jumlahMesin: jumlahMesin,
          totalCM: totalCM,
          pmAchievement: pmAchievementStr
        });
      }
    }

    return { status: 'success', data: fseList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function importFSECSV(csvRows, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    var fseSheet = findOrCreateFSESheet(ss);
    var userSheet = findOrCreateUsersSheet(ss);

    var existingHpMap = {};
    if (userSheet.getLastRow() > 1) {
      var allU = userSheet.getRange(2, 3, userSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var u = 0; u < allU.length; u++) {
        var cleanH = String(allU[u][0] || '').replace(/[^0-9]/g, '');
        if (cleanH) existingHpMap[cleanH] = true;
      }
    }

    var nowFormatted = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var creator = (currentMonitoringUser && currentMonitoringUser.name) ? currentMonitoringUser.name : 'Monitoring Officer';

    var insertedCount = 0;
    var skippedCount = 0;

    for (var i = 0; i < csvRows.length; i++) {
      var row = csvRows[i];
      if (!row || row.length === 0) continue;

      var rawNama = String(row[0] || '').trim();
      var rawHp = String(row[1] || '').replace(/[^0-9]/g, '').trim();
      var rawPass = String(row[2] || 'fse123').trim();
      var rawWilayah = String(row[3] || 'DKI Jakarta').trim();
      var rawKelolaan = parseInt(row[4], 10) || 0;

      if (!rawNama || !rawHp) continue;

      if (existingHpMap[rawHp]) {
        skippedCount++;
        continue;
      }

      existingHpMap[rawHp] = true;
      var newUserId = 'USR-FSE-' + ('00' + (userSheet.getLastRow() + insertedCount)).slice(-3);

      userSheet.appendRow([newUserId, rawNama, rawHp, rawPass, 'FSE', rawWilayah, 'Aktif', creator, nowFormatted]);
      fseSheet.appendRow([newUserId, rawNama, rawHp, rawPass, 'FSE', rawWilayah, 'Aktif', rawKelolaan, creator, nowFormatted]);
      insertedCount++;
    }

    SpreadsheetApp.flush();

    var msg = insertedCount + ' teknisi FSE berhasil diimpor.';
    if (skippedCount > 0) msg += ' (' + skippedCount + ' nomor HP dilewati karena sudah terdaftar).';

    return { status: 'success', message: msg, importedCount: insertedCount, skippedCount: skippedCount };
  } catch (err) {
    return { status: 'error', message: 'Gagal impor FSE: ' + err.toString() };
  }
}

function registerFSE(payload, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak ditemukan' };

    var nama = String(payload.nama || '').trim();
    var noHp = String(payload.noHp || '').replace(/[^0-9]/g, '').trim();
    var password = String(payload.password || '').trim();
    var wilayah = String(payload.wilayah || 'DKI Jakarta').trim();
    var kelolaanMesin = parseInt(payload.kelolaanMesin, 10) || 0;

    if (!nama || !noHp || !password) {
      return { status: 'error', message: 'Nama Lengkap, No HP, dan Password wajib diisi!' };
    }

    var fseSheet = findOrCreateFSESheet(ss);
    var userSheet = findOrCreateUsersSheet(ss);

    var uLastRow = userSheet.getLastRow();
    if (uLastRow > 1) {
      var allUsers = userSheet.getRange(2, 1, uLastRow - 1, 3).getDisplayValues();
      for (var u = 0; u < allUsers.length; u++) {
        var exHp = String(allUsers[u][2] || '').replace(/[^0-9]/g, '').trim();
        if (exHp === noHp) return { status: 'error', message: 'Nomor HP [' + noHp + '] sudah terdaftar!' };
      }
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var newUserId = 'USR-' + ('00' + (uLastRow + 1)).slice(-3);
    var creatorName = (currentMonitoringUser && currentMonitoringUser.name) ? currentMonitoringUser.name : 'Monitoring Officer';

    userSheet.appendRow([newUserId, nama, noHp, password, 'FSE', wilayah, 'Aktif', creatorName, nowFormatted]);
    fseSheet.appendRow([newUserId, nama, noHp, password, 'FSE', wilayah, 'Aktif', kelolaanMesin, creatorName, nowFormatted]);

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Akun FSE [' + nama + '] (' + newUserId + ') berhasil didaftarkan!',
      fse: { userId: newUserId, nama: nama, noHp: noHp, wilayah: wilayah, status: 'Aktif', jumlahMesin: kelolaanMesin }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function updateFSE(payload, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    var userId = String(payload.userId || '').trim();
    var nama = String(payload.nama || '').trim();
    var noHp = String(payload.noHp || '').replace(/[^0-9]/g, '').trim();
    var wilayah = String(payload.wilayah || '').trim();
    var kelolaanMesin = parseInt(payload.kelolaanMesin, 10) || 0;
    var statusAkun = String(payload.status || 'Aktif').trim();

    var fseSheet = findOrCreateFSESheet(ss);
    var userSheet = findOrCreateUsersSheet(ss);

    if (fseSheet && fseSheet.getLastRow() > 1) {
      var fData = fseSheet.getRange(2, 1, fseSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var fi = 0; fi < fData.length; fi++) {
        if (String(fData[fi][0]).trim() === userId) {
          var tRow = fi + 2;
          fseSheet.getRange(tRow, 2).setValue(nama);
          fseSheet.getRange(tRow, 3).setValue(noHp);
          fseSheet.getRange(tRow, 6).setValue(wilayah);
          fseSheet.getRange(tRow, 7).setValue(statusAkun);
          fseSheet.getRange(tRow, 8).setValue(kelolaanMesin);
          break;
        }
      }
    }

    if (userSheet && userSheet.getLastRow() > 1) {
      var uData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var ui = 0; ui < uData.length; ui++) {
        if (String(uData[ui][0]).trim() === userId) {
          var targetURow = ui + 2;
          userSheet.getRange(targetURow, 2).setValue(nama);
          userSheet.getRange(targetURow, 3).setValue(noHp);
          userSheet.getRange(targetURow, 6).setValue(wilayah);
          userSheet.getRange(targetURow, 7).setValue(statusAkun);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Data FSE [' + nama + '] berhasil diperbarui!' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function toggleFSEStatus(userId, currentStatus, currentMonitoringUser) {
  try {
    var newStatus = (currentStatus === 'Aktif') ? 'Nonaktif' : 'Aktif';
    var ss = getSpreadsheet();

    var fseSheet = findOrCreateFSESheet(ss);
    var userSheet = findOrCreateUsersSheet(ss);

    if (fseSheet && fseSheet.getLastRow() > 1) {
      var fData = fseSheet.getRange(2, 1, fseSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var i = 0; i < fData.length; i++) {
        if (String(fData[i][0]).trim() === userId) {
          fseSheet.getRange(i + 2, 7).setValue(newStatus);
          break;
        }
      }
    }

    if (userSheet && userSheet.getLastRow() > 1) {
      var uData = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, 1).getDisplayValues();
      for (var j = 0; j < uData.length; j++) {
        if (String(uData[j][0]).trim() === userId) {
          userSheet.getRange(j + 2, 7).setValue(newStatus);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    return { status: 'success', message: 'Status akun ' + userId + ' diubah ke ' + newStatus, newStatus: newStatus };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
