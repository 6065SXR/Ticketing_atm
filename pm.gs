/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: pm.gs
 * Fungsi: Logika Master & Monitoring Jadwal Preventive Maintenance (Data PM).
 */

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
        idPM: -1,
        idMesin: -1,
        serialNumber: -1,
        bank: -1,
        lokasi: -1,
        periode: -1,
        tanggalRencana: -1,
        tanggalRealisasi: -1,
        engineer: -1,
        status: -1,
        catatan: -1
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
        else if (h === 'tanggalrealisasi' || h === 'tglselesai' || h === 'tglpelaksanaan' || h === 'realisasi' || h === 'tglpm') colMap.tanggalRealisasi = c;
        else if (h === 'engineer' || h === 'namaengineer' || h === 'fse' || h === 'teknisi' || h === 'pic') colMap.engineer = c;
        else if (h === 'status' || h === 'statuspm' || h === 'kondisi' || h === 'state') colMap.status = c;
        else if (h === 'catatan' || h === 'keterangan' || h === 'note' || h === 'notes' || h === 'remark') colMap.catatan = c;
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
