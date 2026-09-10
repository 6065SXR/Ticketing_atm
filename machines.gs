/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: machines.gs
 * Fungsi: Logika Master Database Mesin ATM (Asset Register & Filter Serial Number).
 */

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
