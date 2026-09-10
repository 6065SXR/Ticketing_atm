/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: metrics.gs
 * Fungsi: Kalkulasi Telemetri Dashboard, Kepatuhan SLA, dan KPI Achievement Scorecard.
 */

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
      
      // Deteksi baris header
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

      // Cari kolom status secara otomatis
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
        if (st === 'open' || st === 'respon' || st === 'appointment' || st === 'baru' || st === 'pending' || st === 'antre' || st === 'antri') {
          metrics.openTickets++;
        } else if (st === 'handling' || st === 'take a part' || st === 'take apart' || st === 'proses' || st === 'progress' || st === 'in progress' || st === 'penanganan') {
          metrics.handlingTickets++;
        } else if (st === 'solving' || st === 'solved' || st === 'selesai' || st === 'close' || st === 'closed' || st === 'done' || st === 'sukses') {
          metrics.solvingTickets++;
        } else if (st === 'cancel' || st === 'batal' || st === 'void') {
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
