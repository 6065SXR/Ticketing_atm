/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: tickets.gs
 * Fungsi: Logika Manajemen Tiket (Auto-Increment ID Harian, CRUD, Server Pagination, dan Lifecycle Action).
 */

function generateNewTicketNumber() {
  try {
    var ss = getSpreadsheet();
    var counterSheet = findSheet(ss, ['Counter_Ticket', 'CounterTicket', 'Counter']);
    
    if (!counterSheet) {
      counterSheet = ss.insertSheet('Counter_Ticket');
      counterSheet.appendRow(['Tanggal', 'Last_Sequence']);
    }

    var now = new Date();
    var currentYYMMDD = Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd');
    var nextSequence = 1;

    var lastRow = counterSheet.getLastRow();
    if (lastRow > 1) {
      var counterData = counterSheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
      var foundToday = false;

      for (var i = 0; i < counterData.length; i++) {
        var recordedDate = counterData[i][0].toString().trim();
        if (recordedDate === currentYYMMDD) {
          foundToday = true;
          var currentSeq = parseInt(counterData[i][1], 10) || 0;
          nextSequence = currentSeq + 1;
          counterSheet.getRange(i + 2, 2).setValue(nextSequence);
          break;
        }
      }

      if (!foundToday) {
        counterSheet.appendRow([currentYYMMDD, 1]);
        nextSequence = 1;
      }
    } else {
      counterSheet.appendRow([currentYYMMDD, 1]);
      nextSequence = 1;
    }

    SpreadsheetApp.flush();

    var seqFormatted = ('0000' + nextSequence).slice(-4);
    var fullTicketId = currentYYMMDD + '-' + seqFormatted;
    var dateFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
    var timeFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');

    return {
      status: 'success',
      ticketId: fullTicketId,
      dateString: dateFormatted,
      timeString: timeFormatted
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function createTicket(payload) {
  try {
    var ss = getSpreadsheet();
    var ticketSheet = findTicketsSheet(ss);
    if (!ticketSheet) {
      ticketSheet = ss.insertSheet('Tickets');
      ticketSheet.appendRow([
        'ID_Tiket', 'Bank', 'Serial_Number', 'ID_Mesin', 
        'Tanggal_Tiket', 'Jam_Tiket', 'Tipe_Tiket', 'Nama_Engineer', 
        'Problem', 'Note', 'Status_Tiket', 'Updated_At'
      ]);
    }

    var now = new Date();
    var updatedAt = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var idTiket = String(payload.idTiket || '').trim();
    var bank = String(payload.bank || '').trim();
    var serialNumber = String(payload.serialNumber || '').trim();
    var idMesin = String(payload.idMesin || '').trim();
    var tanggalTiket = String(payload.tanggalTiket || '').trim();
    var jamTiket = String(payload.jamTiket || '').trim();
    var tipeTiket = String(payload.tipeTiket || '').trim();
    var namaEngineer = String(payload.namaEngineer || '').trim();
    var problem = String(payload.problem || '').trim();
    var note = String(payload.note || '').trim();
    var statusTiket = 'Open';

    var newRow = [
      idTiket, bank, serialNumber, idMesin,
      tanggalTiket, jamTiket, tipeTiket, namaEngineer,
      problem, note, statusTiket, updatedAt
    ];

    ticketSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Tiket berhasil disimpan ke database!',
      ticket: {
        idTiket: idTiket,
        bank: bank,
        serialNumber: serialNumber,
        idMesin: idMesin,
        tanggalTiket: tanggalTiket,
        jamTiket: jamTiket,
        tipeTiket: tipeTiket,
        namaEngineer: namaEngineer,
        problem: problem,
        note: note,
        statusTiket: statusTiket,
        updatedAt: updatedAt
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getTicketsPaginated(page, pageSize, searchQuery, statusFilter, bankFilter) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    
    if (!sheet) {
      return {
        status: 'error',
        message: 'Sheet Tiket tidak ditemukan di Spreadsheet.',
        tickets: [],
        page: 1,
        pageSize: pageSize || 10,
        totalItems: 0,
        totalPages: 0
      };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return {
        status: 'success',
        tickets: [],
        page: 1,
        pageSize: pageSize || 10,
        totalItems: 0,
        totalPages: 0
      };
    }

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
    // Deteksi baris header (Baris 1 atau Baris 2)
    var headerRowIdx = 0;
    var headers = allData[0];
    var isHeader0 = false;
    for (var h0 = 0; h0 < headers.length; h0++) {
      var strH0 = String(headers[h0] || '').toLowerCase();
      if (strH0.indexOf('tiket') !== -1 || strH0.indexOf('bank') !== -1 || strH0.indexOf('id') !== -1 || strH0.indexOf('problem') !== -1) {
        isHeader0 = true;
        break;
      }
    }
    if (!isHeader0 && allData.length > 2) {
      for (var h1 = 0; h1 < allData[1].length; h1++) {
        var strH1 = String(allData[1][h1] || '').toLowerCase();
        if (strH1.indexOf('tiket') !== -1 || strH1.indexOf('bank') !== -1 || strH1.indexOf('id') !== -1 || strH1.indexOf('problem') !== -1) {
          headerRowIdx = 1;
          headers = allData[1];
          break;
        }
      }
    }

    // Pemetaan indeks kolom fleksibel (mendukung berbagai format bahasa)
    var colMap = {
      idTiket: -1,
      bank: -1,
      serialNumber: -1,
      idMesin: -1,
      tanggalTiket: -1,
      jamTiket: -1,
      tipeTiket: -1,
      namaEngineer: -1,
      problem: -1,
      note: -1,
      statusTiket: -1,
      updatedAt: -1
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'nomortiket' || h === 'ticketid' || h === 'id' || h === 'no' || h === 'ticketno') {
        if (colMap.idTiket === -1) colMap.idTiket = c;
      }
      else if (h === 'bank' || h === 'namabank' || h === 'customer' || h === 'namacustomer' || h === 'mitra') {
        if (colMap.bank === -1) colMap.bank = c;
      }
      else if (h === 'serialnumber' || h === 'sn' || h === 'noserial' || h === 'noseri' || h === 'serialno' || h === 'seri') {
        if (colMap.serialNumber === -1) colMap.serialNumber = c;
      }
      else if (h === 'idmesin' || h === 'machineid' || h === 'mesin' || h === 'atmid' || h === 'idatm' || h === 'tid' || h === 'nomesin') {
        if (colMap.idMesin === -1) colMap.idMesin = c;
      }
      else if (h === 'tanggaltiket' || h === 'tanggal' || h === 'date' || h === 'tgl' || h === 'tgltiket' || h === 'tglbuat') {
        if (colMap.tanggalTiket === -1) colMap.tanggalTiket = c;
      }
      else if (h === 'jamtiket' || h === 'jam' || h === 'time' || h === 'waktu' || h === 'jambuat') {
        if (colMap.jamTiket === -1) colMap.jamTiket = c;
      }
      else if (h === 'tipetiket' || h === 'tipe' || h === 'type' || h === 'jenistiket' || h === 'jenis' || h === 'kategori') {
        if (colMap.tipeTiket === -1) colMap.tipeTiket = c;
      }
      else if (h === 'namaengineer' || h === 'engineer' || h === 'fse' || h === 'teknisi' || h === 'namafse' || h === 'namateknisi' || h === 'petugas' || h === 'pic') {
        if (colMap.namaEngineer === -1) colMap.namaEngineer = c;
      }
      else if (h === 'problem' || h === 'masalah' || h === 'keluhan' || h === 'kendala' || h === 'kerusakan' || h === 'deskripsi' || h === 'issue' || h === 'keteranganproblem') {
        if (colMap.problem === -1) colMap.problem = c;
      }
      else if (h === 'note' || h === 'catatan' || h === 'keterangan' || h === 'notes' || h === 'remark' || h === 'remarks') {
        if (colMap.note === -1) colMap.note = c;
      }
      else if (h === 'statustiket' || h === 'status' || h === 'statusticket' || h === 'kondisi' || h === 'state') {
        if (colMap.statusTiket === -1) colMap.statusTiket = c;
      }
      else if (h === 'updatedat' || h === 'update' || h === 'waktuupdate' || h === 'lastupdate' || h === 'tglupdate' || h === 'modified') {
        if (colMap.updatedAt === -1) colMap.updatedAt = c;
      }
    }

    // Fallback otomatis jika header berbeda
    var maxIdx = Math.max(headers.length - 1, 0);
    if (colMap.idTiket === -1) colMap.idTiket = 0;
    if (colMap.bank === -1) colMap.bank = Math.min(1, maxIdx);
    if (colMap.serialNumber === -1) colMap.serialNumber = Math.min(2, maxIdx);
    if (colMap.idMesin === -1) colMap.idMesin = Math.min(3, maxIdx);
    if (colMap.tanggalTiket === -1) colMap.tanggalTiket = Math.min(4, maxIdx);
    if (colMap.jamTiket === -1) colMap.jamTiket = Math.min(5, maxIdx);
    if (colMap.tipeTiket === -1) colMap.tipeTiket = Math.min(6, maxIdx);
    if (colMap.namaEngineer === -1) colMap.namaEngineer = Math.min(7, maxIdx);
    if (colMap.problem === -1) colMap.problem = Math.min(8, maxIdx);
    if (colMap.note === -1) colMap.note = Math.min(9, maxIdx);
    if (colMap.statusTiket === -1) colMap.statusTiket = Math.min(10, maxIdx);
    if (colMap.updatedAt === -1) colMap.updatedAt = Math.min(11, maxIdx);

    var rawData = allData.slice(headerRowIdx + 1);
    var filtered = [];
    var searchLower = (searchQuery || '').toLowerCase().trim();
    var statusFilterVal = (statusFilter || '').trim();
    var bankFilterVal = (bankFilter || '').trim();

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      
      var hasData = false;
      for (var k = 0; k < row.length; k++) {
        if (String(row[k] || '').trim() !== '') {
          hasData = true;
          break;
        }
      }
      if (!hasData) continue;

      var idVal = colMap.idTiket >= 0 ? String(row[colMap.idTiket] || '').trim() : '';
      if (!idVal) idVal = 'TK-' + (i + 1);

      var item = {
        idTiket: idVal,
        bank: colMap.bank >= 0 ? String(row[colMap.bank] || '-').trim() : '-',
        serialNumber: colMap.serialNumber >= 0 ? String(row[colMap.serialNumber] || '-').trim() : '-',
        idMesin: colMap.idMesin >= 0 ? String(row[colMap.idMesin] || '-').trim() : '-',
        tanggalTiket: colMap.tanggalTiket >= 0 ? String(row[colMap.tanggalTiket] || '-').trim() : '-',
        jamTiket: colMap.jamTiket >= 0 ? String(row[colMap.jamTiket] || '-').trim() : '-',
        tipeTiket: colMap.tipeTiket >= 0 ? String(row[colMap.tipeTiket] || 'CM').trim() : 'CM',
        namaEngineer: colMap.namaEngineer >= 0 ? String(row[colMap.namaEngineer] || '-').trim() : '-',
        problem: colMap.problem >= 0 ? String(row[colMap.problem] || '-').trim() : '-',
        note: colMap.note >= 0 ? String(row[colMap.note] || '').trim() : '',
        statusTiket: colMap.statusTiket >= 0 ? String(row[colMap.statusTiket] || 'Open').trim() : 'Open',
        updatedAt: colMap.updatedAt >= 0 ? String(row[colMap.updatedAt] || '-').trim() : '-'
      };

      if (!item.statusTiket) item.statusTiket = 'Open';

      if (statusFilterVal && statusFilterVal !== 'ALL') {
        if (item.statusTiket.toLowerCase() !== statusFilterVal.toLowerCase()) continue;
      }

      if (bankFilterVal && bankFilterVal !== 'ALL') {
        var b1 = item.bank.toLowerCase();
        var b2 = bankFilterVal.toLowerCase();
        if (b1 !== b2 && b1.indexOf(b2) === -1 && b2.indexOf(b1) === -1) continue;
      }

      if (searchLower !== '') {
        var match = (
          item.idTiket.toLowerCase().indexOf(searchLower) !== -1 ||
          item.bank.toLowerCase().indexOf(searchLower) !== -1 ||
          item.serialNumber.toLowerCase().indexOf(searchLower) !== -1 ||
          item.idMesin.toLowerCase().indexOf(searchLower) !== -1 ||
          item.problem.toLowerCase().indexOf(searchLower) !== -1 ||
          item.namaEngineer.toLowerCase().indexOf(searchLower) !== -1
        );
        if (!match) continue;
      }

      filtered.push(item);
    }

    filtered.reverse(); // Tiket terbaru di posisi paling atas

    var totalItems = filtered.length;
    var validPageSize = Math.max(1, parseInt(pageSize, 10) || 10);
    var totalPages = Math.ceil(totalItems / validPageSize) || 1;
    var currentPage = Math.min(Math.max(1, parseInt(page, 10) || 1), totalPages);

    var startIndex = (currentPage - 1) * validPageSize;
    var endIndex = Math.min(startIndex + validPageSize, totalItems);
    var pagedData = filtered.slice(startIndex, endIndex);

    return {
      status: 'success',
      tickets: pagedData,
      page: currentPage,
      pageSize: validPageSize,
      totalItems: totalItems,
      totalPages: totalPages
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function updateTicketStatus(ticketId, newStatus, additionalNote) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    if (!sheet) return { status: 'error', message: 'Sheet Tickets tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data tiket' };

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var idCol = 0;
    var statusCol = 10;
    var noteCol = 9;
    var updateCol = 11;

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'nomortiket' || h === 'ticketid' || h === 'id') idCol = c;
      else if (h === 'statustiket' || h === 'status' || h === 'statusticket' || h === 'kondisi') statusCol = c;
      else if (h === 'note' || h === 'catatan' || h === 'keterangan') noteCol = c;
      else if (h === 'updatedat' || h === 'update' || h === 'waktuupdate') updateCol = c;
    }

    var targetRowIndex = -1;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][idCol]).trim() === String(ticketId).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { status: 'error', message: 'Tiket ' + ticketId + ' tidak ditemukan di database' };
    }

    var now = new Date();
    var updatedAt = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    sheet.getRange(targetRowIndex, statusCol + 1).setValue(newStatus);
    if (updateCol < lastCol) {
      sheet.getRange(targetRowIndex, updateCol + 1).setValue(updatedAt);
    }

    if (additionalNote && String(additionalNote).trim() !== '' && noteCol < lastCol) {
      var curNote = sheet.getRange(targetRowIndex, noteCol + 1).getDisplayValue();
      var combinedNote = curNote ? (curNote + ' | [' + newStatus + ']: ' + additionalNote) : ('[' + newStatus + ']: ' + additionalNote);
      sheet.getRange(targetRowIndex, noteCol + 1).setValue(combinedNote);
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      ticketId: ticketId,
      newStatus: newStatus,
      updatedAt: updatedAt
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
