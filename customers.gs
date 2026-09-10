/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: customers.gs
 * Fungsi: Logika Master Database Customer (Mitra Bank & Kontak PIC).
 */

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
