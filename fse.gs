/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: fse.gs
 * Fungsi: Master Controller Sheet "FSE", Sinkronisasi Sheet "Users", Perhitungan Metrik CM & PM Bulanan, dan CRUD FSE.
 */

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

    // Sinkronisasi awal dari sheet Users yang berstatus FSE
    var userSheet = findSheet(ss, ['Users', 'User']);
    if (userSheet && userSheet.getLastRow() > 1) {
      var allUsers = userSheet.getRange(2, 1, userSheet.getLastRow() - 1, userSheet.getLastColumn()).getDisplayValues();
      var fseRows = [];
      for (var u = 0; u < allUsers.length; u++) {
        var r = allUsers[u];
        var role = String(r[4] || '').toUpperCase();
        if (role === 'FSE') {
          fseRows.push([
            r[0], // User_ID
            r[1], // Nama_Lengkap
            r[2], // No_HP
            r[3], // Password
            'FSE',
            r[5] || 'DKI Jakarta', // Wilayah_Tugas
            r[6] || 'Aktif',       // Status_Akun
            5,                     // Default Jumlah Kelolaan Mesin
            r[7] || 'System',      // Created_By
            r[8] || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')
          ]);
        }
      }
      if (fseRows.length > 0) {
        sheet.getRange(2, 1, fseRows.length, fseRows[0].length).setValues(fseRows);
      }
    }
    SpreadsheetApp.flush();
  }
  return sheet;
}

/**
 * Mengambil data FSE lengkap dengan:
 * - User_ID (ID Karyawan)
 * - Kelolaan Mesin
 * - Total Tiket Corrective Maintenance (CM) dalam 1 bulan kalender berjalan
 * - Pencapaian Preventive Maintenance (PM) dalam 1 bulan kalender berjalan
 */
function getFSEData() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Database Spreadsheet tidak dapat diakses' };

    var sheetFSE = findOrCreateFSESheet(ss);
    var fseList = [];

    // Tentukan bulan berjalan (format YYYY-MM)
    var now = new Date();
    var currentMonth = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM');

    // 1. Ambil data Tiket untuk kalkulasi CM bulan ini
    var ticketSheet = findTicketsSheet(ss);
    var ticketRows = [];
    var tColMap = { engineer: -1, tipe: -1, tanggal: -1 };

    if (ticketSheet && ticketSheet.getLastRow() > 1) {
      var tAll = ticketSheet.getRange(1, 1, ticketSheet.getLastRow(), Math.max(ticketSheet.getLastColumn(), 1)).getDisplayValues();
      var tHeaders = tAll[0];
      for (var th = 0; th < tHeaders.length; th++) {
        var thStr = String(tHeaders[th] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (thStr === 'namaengineer' || thStr === 'engineer' || thStr === 'fse' || thStr === 'teknisi') tColMap.engineer = th;
        else if (thStr === 'tipetiket' || thStr === 'tipe' || thStr === 'type' || thStr === 'jenis') tColMap.tipe = th;
        else if (thStr === 'tanggaltiket' || thStr === 'tanggal' || thStr === 'date' || thStr === 'tgl') tColMap.tanggal = th;
      }
      ticketRows = tAll.slice(1);
    }

    // 2. Ambil data PM untuk kalkulasi capaian PM bulan ini
    var sheetPM = findSheet(ss, ['Data_PM', 'Data PM', 'DataPM', 'PM', 'Preventive Maintenance']);
    var pmRows = [];
    var pColMap = { engineer: -1, status: -1, tanggalRencana: -1, tanggalRealisasi: -1, periode: -1 };

    if (sheetPM && sheetPM.getLastRow() > 1) {
      var pAll = sheetPM.getRange(1, 1, sheetPM.getLastRow(), Math.max(sheetPM.getLastColumn(), 1)).getDisplayValues();
      var pHeaders = pAll[0];
      for (var ph = 0; ph < pHeaders.length; ph++) {
        var phStr = String(pHeaders[ph] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (phStr === 'engineer' || phStr === 'namaengineer' || phStr === 'fse' || phStr === 'pic') pColMap.engineer = ph;
        else if (phStr === 'status' || phStr === 'statuspm') pColMap.status = ph;
        else if (phStr === 'tanggalrencana' || phStr === 'jadwalpm' || phStr === 'target') pColMap.tanggalRencana = ph;
        else if (phStr === 'tanggalrealisasi' || phStr === 'realisasi') pColMap.tanggalRealisasi = ph;
        else if (phStr === 'periode' || phStr === 'siklus') pColMap.periode = ph;
      }
      pmRows = pAll.slice(1);
    }

    // 3. Baca sheet FSE
    var fLastRow = sheetFSE.getLastRow();
    if (fLastRow > 1) {
      var fAll = sheetFSE.getRange(1, 1, fLastRow, Math.max(sheetFSE.getLastColumn(), 10)).getDisplayValues();
      var fHeaders = fAll[0];
      var fColMap = {
        userId: 0, nama: 1, noHp: 2, password: 3, role: 4,
        wilayah: 5, status: 6, jumlahMesin: 7
      };

      for (var fc = 0; fc < fHeaders.length; fc++) {
        var fch = String(fHeaders[fc] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (fch === 'userid' || fch === 'id') fColMap.userId = fc;
        else if (fch === 'namalengkap' || fch === 'nama') fColMap.nama = fc;
        else if (fch === 'nohp' || fch === 'hp') fColMap.noHp = fc;
        else if (fch === 'password' || fch === 'pass') fColMap.password = fc;
        else if (fch === 'wilayahtugas' || fch === 'wilayah') fColMap.wilayah = fc;
        else if (fch === 'statusakun' || fch === 'status') fColMap.status = fc;
        else if (fch === 'jumlahkelolaanmesin' || fch === 'kelolaanmesin' || fch === 'mesin') fColMap.jumlahMesin = fc;
      }

      var rows = fAll.slice(1);
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (!r.some(function(c) { return String(c).trim() !== ''; })) continue;

        var userId = String(r[fColMap.userId] || ('USR-00' + (i + 1))).trim();
        var namaFSE = String(r[fColMap.nama] || '').trim();
        var noHp = String(r[fColMap.noHp] || '').trim();
        var wilayah = String(r[fColMap.wilayah] || '-').trim();
        var status = String(r[fColMap.status] || 'Aktif').trim();
        var jumlahMesin = parseInt(r[fColMap.jumlahMesin], 10) || 0;

        // Kalkulasi Total Tiket CM (Bulan Ini)
        var totalCM = 0;
        var namaLower = namaFSE.toLowerCase();
        for (var t = 0; t < ticketRows.length; t++) {
          var tr = ticketRows[t];
          var tEng = String(tr[tColMap.engineer] || '').toLowerCase();
          var tTipe = String(tr[tColMap.tipe] || '').toLowerCase();
          var tTgl = String(tr[tColMap.tanggal] || '');

          if (namaLower && (tEng.indexOf(namaLower) !== -1 || namaLower.indexOf(tEng) !== -1)) {
            var isCM = (tTipe.indexOf('cm') !== -1 || tTipe.indexOf('corrective') !== -1);
            var isThisMonth = (tTgl.indexOf(currentMonth) === 0 || tTgl.indexOf(currentMonth.replace('-', '/')) === 0);
            if (isCM && isThisMonth) {
              totalCM++;
            }
          }
        }

        // Kalkulasi Capaian PM (Bulan Ini)
        var pmTarget = 0;
        var pmCompleted = 0;
        for (var p = 0; p < pmRows.length; p++) {
          var pr = pmRows[p];
          var pEng = String(pr[pColMap.engineer] || '').toLowerCase();
          var pStat = String(pr[pColMap.status] || '').toLowerCase();
          var pTglRencana = String(pr[pColMap.tanggalRencana] || '');
          var pTglRealisasi = String(pr[pColMap.tanggalRealisasi] || '');

          if (namaLower && (pEng.indexOf(namaLower) !== -1 || namaLower.indexOf(pEng) !== -1)) {
            var matchPMMonth = (
              pTglRencana.indexOf(currentMonth) === 0 || 
              pTglRealisasi.indexOf(currentMonth) === 0
            );
            if (matchPMMonth) {
              pmTarget++;
              if (pStat === 'completed' || pStat === 'selesai') {
                pmCompleted++;
              }
            }
          }
        }

        var pmAchievementStr = '-';
        if (pmTarget > 0) {
          var pmPercent = Math.round((pmCompleted / pmTarget) * 100);
          pmAchievementStr = pmCompleted + ' / ' + pmTarget + ' (' + pmPercent + '%)';
        } else {
          pmAchievementStr = '0 / 0 (100%)';
        }

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

/**
 * Pendaftaran FSE Baru (Menyimpan serentak ke sheet FSE, Users, dan Engineers)
 */
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
    var kelolaanMesin = parseInt(payload.kelolaanMesin, 10) || 0;

    if (!nama || !noHp || !password) {
      return { status: 'error', message: 'Nama Lengkap, No HP, dan Password wajib diisi!' };
    }

    var fseSheet = findOrCreateFSESheet(ss);
    var userSheet = findOrCreateUsersSheet(ss);

    // Cek duplikasi Nomor HP di Users
    var uLastRow = userSheet.getLastRow();
    if (uLastRow > 1) {
      var allUsers = userSheet.getRange(2, 1, uLastRow - 1, Math.max(userSheet.getLastColumn(), 3)).getDisplayValues();
      for (var u = 0; u < allUsers.length; u++) {
        var exHp = String(allUsers[u][2] || '').replace(/[^0-9]/g, '').trim();
        if (exHp === noHp) {
          return { status: 'error', message: 'Nomor HP [' + noHp + '] sudah terdaftar pada pengguna lain!' };
        }
      }
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var nextUserSeq = (uLastRow > 1) ? ('00' + uLastRow).slice(-3) : '001';
    var newUserId = 'USR-' + nextUserSeq;
    var creatorName = (currentMonitoringUser && currentMonitoringUser.name) ? currentMonitoringUser.name : 'Monitoring Officer';

    // 1. Simpan ke sheet Users
    userSheet.appendRow([
      newUserId, nama, noHp, password, 'FSE', wilayah, 'Aktif', creatorName, nowFormatted
    ]);

    // 2. Simpan ke sheet FSE
    fseSheet.appendRow([
      newUserId, nama, noHp, password, 'FSE', wilayah, 'Aktif', kelolaanMesin, creatorName, nowFormatted
    ]);

    // 3. Simpan ke sheet Engineers (Kompatibilitas data lama)
    var engSheet = findSheet(ss, ['Engineers', 'Engineer']);
    if (engSheet) {
      engSheet.appendRow([newUserId, nama, wilayah, noHp, 'Aktif']);
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Akun FSE [' + nama + '] (' + newUserId + ') berhasil didaftarkan!',
      fse: {
        userId: newUserId,
        nama: nama,
        noHp: noHp,
        wilayah: wilayah,
        status: 'Aktif',
        jumlahMesin: kelolaanMesin,
        totalCM: 0,
        pmAchievement: '0 / 0 (100%)'
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Pembaruan Data FSE (Update Serentak ke sheet FSE, Users, dan Engineers)
 */
function updateFSE(payload, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak ditemukan' };

    var editorRole = (currentMonitoringUser && currentMonitoringUser.roleKey) ? currentMonitoringUser.roleKey.toLowerCase() : '';
    if (editorRole !== 'monitoring' && editorRole !== 'admin') {
      return { status: 'error', message: 'Hanya Role Monitoring (atau Admin) yang berhak mengubah data FSE!' };
    }

    var userId = String(payload.userId || '').trim();
    var nama = String(payload.nama || '').trim();
    var noHp = String(payload.noHp || '').replace(/[^0-9]/g, '').trim();
    var wilayah = String(payload.wilayah || '').trim();
    var kelolaanMesin = parseInt(payload.kelolaanMesin, 10) || 0;
    var statusAkun = String(payload.status || 'Aktif').trim();
    var newPassword = String(payload.password || '').trim();

    if (!userId || !nama || !noHp) {
      return { status: 'error', message: 'User ID, Nama Lengkap, dan No HP wajib disertakan!' };
    }

    var fseSheet = findOrCreateFSESheet(ss);
    var userSheet = findOrCreateUsersSheet(ss);

    // 1. Perbarui pada sheet FSE
    var fLastRow = fseSheet.getLastRow();
    if (fLastRow > 1) {
      var fData = fseSheet.getRange(2, 1, fLastRow - 1, fseSheet.getLastColumn()).getDisplayValues();
      for (var fi = 0; fi < fData.length; fi++) {
        if (String(fData[fi][0]).trim() === userId) {
          var targetRow = fi + 2;
          fseSheet.getRange(targetRow, 2).setValue(nama);
          fseSheet.getRange(targetRow, 3).setValue(noHp);
          if (newPassword) fseSheet.getRange(targetRow, 4).setValue(newPassword);
          fseSheet.getRange(targetRow, 6).setValue(wilayah);
          fseSheet.getRange(targetRow, 7).setValue(statusAkun);
          fseSheet.getRange(targetRow, 8).setValue(kelolaanMesin);
          break;
        }
      }
    }

    // 2. Perbarui pada sheet Users
    var uLastRow = userSheet.getLastRow();
    if (uLastRow > 1) {
      var uData = userSheet.getRange(2, 1, uLastRow - 1, userSheet.getLastColumn()).getDisplayValues();
      for (var ui = 0; ui < uData.length; ui++) {
        if (String(uData[ui][0]).trim() === userId) {
          var uRow = ui + 2;
          userSheet.getRange(uRow, 2).setValue(nama);
          userSheet.getRange(uRow, 3).setValue(noHp);
          if (newPassword) userSheet.getRange(uRow, 4).setValue(newPassword);
          userSheet.getRange(uRow, 6).setValue(wilayah);
          userSheet.getRange(uRow, 7).setValue(statusAkun);
          break;
        }
      }
    }

    // 3. Perbarui pada sheet Engineers (Sinkronisasi Nama, Wilayah, HP)
    var engSheet = findSheet(ss, ['Engineers', 'Engineer']);
    if (engSheet && engSheet.getLastRow() > 1) {
      var eData = engSheet.getRange(2, 1, engSheet.getLastRow() - 1, engSheet.getLastColumn()).getDisplayValues();
      for (var ei = 0; ei < eData.length; ei++) {
        if (String(eData[ei][0]).trim() === userId || String(eData[ei][1]).trim().toLowerCase() === nama.toLowerCase()) {
          var eRow = ei + 2;
          engSheet.getRange(eRow, 2).setValue(nama);
          engSheet.getRange(eRow, 3).setValue(wilayah);
          engSheet.getRange(eRow, 4).setValue(noHp);
          engSheet.getRange(eRow, 5).setValue(statusAkun);
          break;
        }
      }
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Data FSE [' + nama + '] (' + userId + ') berhasil diperbarui!',
      fse: {
        userId: userId,
        nama: nama,
        noHp: noHp,
        wilayah: wilayah,
        status: statusAkun,
        jumlahMesin: kelolaanMesin
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Nonaktifkan / Ubah Status FSE (Soft Delete)
 */
function toggleFSEStatus(userId, currentStatus, currentMonitoringUser) {
  try {
    var newStatus = (currentStatus === 'Aktif') ? 'Nonaktif' : 'Aktif';
    var ss = getSpreadsheet();
    if (!ss) return { status: 'error', message: 'Spreadsheet tidak ditemukan' };

    var editorRole = (currentMonitoringUser && currentMonitoringUser.roleKey) ? currentMonitoringUser.roleKey.toLowerCase() : '';
    if (editorRole !== 'monitoring' && editorRole !== 'admin') {
      return { status: 'error', message: 'Hanya Role Monitoring (atau Admin) yang berhak mengubah status akun FSE!' };
    }

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

    return {
      status: 'success',
      message: 'Status akun ' + userId + ' berhasil diubah menjadi ' + newStatus,
      newStatus: newStatus
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
