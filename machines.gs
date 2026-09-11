/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: machines.gs
 * Fungsi: Logika Master Database Mesin ATM (Asset Register, FSE Pengelola, 
 *         serta Import CSV dengan Validasi Anti-Duplikasi Serial Number).
 */

function getMachinesData() {
  try {
    var ss = getSpreadsheet();
    var sheetMach = findSheet(ss, ['Machines', 'Machine', 'Data Mesin', 'Mesin', 'ATM']);
    var machList = [];

    if (sheetMach && sheetMach.getLastRow() > 1) {
      var lastRow = sheetMach.getLastRow();
      var lastCol = Math.max(sheetMach.getLastColumn(), 7);
      var allData = sheetMach.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var headers = allData[0];

      var colMap = { idMesin: 0, serialNumber: 1, bank: 2, lokasi: 3, tipe: 4, status: 5, fsePengelola: 6 };
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === 'idmesin' || h === 'id' || h === 'machineid') colMap.idMesin = c;
        else if (h === 'serialnumber' || h === 'sn' || h === 'noserial') colMap.serialNumber = c;
        else if (h === 'namabank' || h === 'bank') colMap.bank = c;
        else if (h === 'lokasiatm' || h === 'lokasi' || h === 'alamat') colMap.lokasi = c;
        else if (h === 'tipemesin' || h === 'tipe' || h === 'model') colMap.tipe = c;
        else if (h === 'statusmesin' || h === 'status') colMap.status = c;
        else if (h === 'fsepengelola' || h === 'pengelola' || h === 'fse' || h === 'teknisi') colMap.fsePengelola = c;
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
          status: String(r[colMap.status] || 'Active').trim(),
          fsePengelola: String(r[colMap.fsePengelola] || '-').trim()
        });
      }
    }

    return { status: 'success', data: machList };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function importMachinesCSV(csvRows, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    var sheetMach = findSheet(ss, ['Machines', 'Machine', 'Data Mesin', 'Mesin', 'ATM']);
    if (!sheetMach) return { status: 'error', message: 'Sheet Machines tidak ditemukan' };

    var existingSNMap = {};
    if (sheetMach.getLastRow() > 1) {
      var existingData = sheetMach.getRange(2, 2, sheetMach.getLastRow() - 1, 1).getDisplayValues();
      for (var k = 0; k < existingData.length; k++) {
        var snVal = String(existingData[k][0] || '').toUpperCase().trim();
        if (snVal) existingSNMap[snVal] = true;
      }
    }

    var insertedRows = [];
    var skippedDuplicates = 0;

    for (var i = 0; i < csvRows.length; i++) {
      var row = csvRows[i];
      if (!row || row.length === 0) continue;

      var rawId = String(row[0] || '').trim();
      var rawSN = String(row[1] || '').toUpperCase().trim();
      var rawBank = String(row[2] || '').trim();
      var rawLokasi = String(row[3] || '').trim();
      var rawTipe = String(row[4] || 'NCR / Diebold').trim();
      var rawStatus = String(row[5] || 'Active').trim();
      var rawFSE = String(row[6] || '-').trim();

      if (!rawSN) continue;

      if (existingSNMap[rawSN]) {
        skippedDuplicates++;
        continue;
      }

      existingSNMap[rawSN] = true;
      insertedRows.push([rawId, rawSN, rawBank, rawLokasi, rawTipe, rawStatus, rawFSE]);
    }

    if (insertedRows.length > 0) {
      var startRow = sheetMach.getLastRow() + 1;
      sheetMach.getRange(startRow, 1, insertedRows.length, insertedRows[0].length).setValues(insertedRows);
      SpreadsheetApp.flush();
    }

    var msg = insertedRows.length + ' mesin ATM berhasil diimpor.';
    if (skippedDuplicates > 0) {
      msg += ' (' + skippedDuplicates + ' mesin dilewati karena Serial Number ganda/sudah ada).';
    }

    return {
      status: 'success',
      message: msg,
      importedCount: insertedRows.length,
      skippedCount: skippedDuplicates
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal impor mesin: ' + err.toString() };
  }
}
