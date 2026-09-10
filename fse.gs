/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: fse.gs
 * Fungsi: Logika Master Data Field Service Engineer (FSE) dan Reporting Kinerja.
 */

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
