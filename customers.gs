/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: customers.gs
 * Fungsi: Logika Master Database Customer (Mitra Bank & Kontak PIC) serta Import CSV.
 */

function getCustomersData() {
  try {
    var ss = getSpreadsheet();
    var sheetCust = findSheet(ss, ['Customers', 'Customer', 'Data Customer', 'Data Bank', 'Bank']);
    var custList = [];

    if (sheetCust && sheetCust.getLastRow() > 1) {
      var lastRow = sheetCust.getLastRow();
      var lastCol = Math.max(sheetCust.getLastColumn(), 6);
      var allData = sheetCust.getRange(1, 1, lastRow, lastCol).getDisplayValues();
      var headers = allData[0];

      var colMap = { idCustomer: 0, namaBank: 1, kontakPic: 2, teleponPic: 3, emailBank: 4, alamat: 5 };
      for (var c = 0; c < headers.length; c++) {
        var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
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

function importCustomersCSV(csvRows, currentMonitoringUser) {
  try {
    var ss = getSpreadsheet();
    var sheetCust = findSheet(ss, ['Customers', 'Customer', 'Data Customer', 'Data Bank', 'Bank']);
    if (!sheetCust) return { status: 'error', message: 'Sheet Customers tidak ditemukan' };

    var existingBankMap = {};
    if (sheetCust.getLastRow() > 1) {
      var allB = sheetCust.getRange(2, 2, sheetCust.getLastRow() - 1, 1).getDisplayValues();
      for (var b = 0; b < allB.length; b++) {
        var bName = String(allB[b][0] || '').toLowerCase().trim();
        if (bName) existingBankMap[bName] = true;
      }
    }

    var insertedRows = [];
    var skippedCount = 0;

    for (var i = 0; i < csvRows.length; i++) {
      var row = csvRows[i];
      if (!row || row.length === 0) continue;

      var rawId = String(row[0] || '').trim() || ('CUST-00' + (sheetCust.getLastRow() + insertedRows.length));
      var rawBank = String(row[1] || '').trim();
      var rawPic = String(row[2] || '-').trim();
      var rawPhone = String(row[3] || '-').trim();
      var rawEmail = String(row[4] || '-').trim();
      var rawAlamat = String(row[5] || '-').trim();

      if (!rawBank) continue;

      if (existingBankMap[rawBank.toLowerCase()]) {
        skippedCount++;
        continue;
      }

      existingBankMap[rawBank.toLowerCase()] = true;
      insertedRows.push([rawId, rawBank, rawPic, rawPhone, rawEmail, rawAlamat]);
    }

    if (insertedRows.length > 0) {
      var startRow = sheetCust.getLastRow() + 1;
      sheetCust.getRange(startRow, 1, insertedRows.length, insertedRows[0].length).setValues(insertedRows);
      SpreadsheetApp.flush();
    }

    var msg = insertedRows.length + ' mitra Bank berhasil diimpor.';
    if (skippedCount > 0) msg += ' (' + skippedCount + ' bank dilewati karena sudah terdaftar).';

    return { status: 'success', message: msg, importedCount: insertedRows.length, skippedCount: skippedCount };
  } catch (err) {
    return { status: 'error', message: 'Gagal impor customer: ' + err.toString() };
  }
}
