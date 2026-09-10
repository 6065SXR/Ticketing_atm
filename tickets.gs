/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: tickets.gs
 * Fungsi: Logika Tiket Harian, Server Pagination, Strict FSE Action Engine,
 *         Penutupan Tiket (Closed) oleh Monitoring, dan Verifikasi Dokumen Fisik oleh Admin.
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
        'Problem', 'Note', 'Status_Tiket', 'Updated_At',
        'Closed_At', 'Hardcopy_Status', 'Hardcopy_Received_At', 'Hardcopy_Received_By'
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
      problem, note, statusTiket, updatedAt,
      '', 'Belum Dikirim', '', ''
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
        updatedAt: updatedAt,
        closedAt: '',
        hardcopyStatus: 'Belum Dikirim'
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
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        status: 'success',
        tickets: [],
        page: 1,
        pageSize: pageSize || 10,
        totalItems: 0,
        totalPages: 0
      };
    }

    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 16);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    
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
      headerRowIdx = 1;
      headers = allData[1];
    }

    var colMap = {
      idTiket: 0, bank: 1, serialNumber: 2, idMesin: 3,
      tanggalTiket: 4, jamTiket: 5, tipeTiket: 6, namaEngineer: 7,
      problem: 8, note: 9, statusTiket: 10, updatedAt: 11,
      closedAt: 12, hardcopyStatus: 13, hardcopyReceivedAt: 14, hardcopyReceivedBy: 15
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'ticketid' || h === 'id') colMap.idTiket = c;
      else if (h === 'bank' || h === 'namabank' || h === 'customer') colMap.bank = c;
      else if (h === 'serialnumber' || h === 'sn' || h === 'noserial') colMap.serialNumber = c;
      else if (h === 'idmesin' || h === 'atmid' || h === 'machineid') colMap.idMesin = c;
      else if (h === 'tanggaltiket' || h === 'tanggal' || h === 'date') colMap.tanggalTiket = c;
      else if (h === 'jamtiket' || h === 'jam' || h === 'time') colMap.jamTiket = c;
      else if (h === 'tipetiket' || h === 'tipe' || h === 'type') colMap.tipeTiket = c;
      else if (h === 'namaengineer' || h === 'engineer' || h === 'fse' || h === 'teknisi') colMap.namaEngineer = c;
      else if (h === 'problem' || h === 'masalah' || h === 'kendala') colMap.problem = c;
      else if (h === 'note' || h === 'catatan' || h === 'notes') colMap.note = c;
      else if (h === 'statustiket' || h === 'status' || h === 'kondisi') colMap.statusTiket = c;
      else if (h === 'updatedat' || h === 'update' || h === 'waktuupdate') colMap.updatedAt = c;
      else if (h === 'closedat' || h === 'tglclosed' || h === 'waktuclosed') colMap.closedAt = c;
      else if (h === 'hardcopystatus' || h === 'statushardcopy' || h === 'dokumenfisik') colMap.hardcopyStatus = c;
      else if (h === 'hardcopyreceivedat' || h === 'tglditerima' || h === 'waktuterima') colMap.hardcopyReceivedAt = c;
      else if (h === 'hardcopyreceivedby' || h === 'penerima' || h === 'adminpenerima') colMap.hardcopyReceivedBy = c;
    }

    var rawData = allData.slice(headerRowIdx + 1);
    var filtered = [];
    var searchLower = (searchQuery || '').toLowerCase().trim();
    var statusFilterVal = (statusFilter || '').trim();
    var bankFilterVal = (bankFilter || '').trim();

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      if (!row.some(function(k) { return String(k).trim() !== ''; })) continue;

      var idVal = String(row[colMap.idTiket] || ('TK-' + (i + 1))).trim();

      var item = {
        idTiket: idVal,
        bank: String(row[colMap.bank] || '-').trim(),
        serialNumber: String(row[colMap.serialNumber] || '-').trim(),
        idMesin: String(row[colMap.idMesin] || '-').trim(),
        tanggalTiket: String(row[colMap.tanggalTiket] || '-').trim(),
        jamTiket: String(row[colMap.jamTiket] || '-').trim(),
        tipeTiket: String(row[colMap.tipeTiket] || 'CM').trim(),
        namaEngineer: String(row[colMap.namaEngineer] || '-').trim(),
        problem: String(row[colMap.problem] || '-').trim(),
        note: String(row[colMap.note] || '').trim(),
        statusTiket: String(row[colMap.statusTiket] || 'Open').trim(),
        updatedAt: String(row[colMap.updatedAt] || '-').trim(),
        closedAt: String(row[colMap.closedAt] || '').trim(),
        hardcopyStatus: String(row[colMap.hardcopyStatus] || 'Belum Dikirim').trim(),
        hardcopyReceivedAt: String(row[colMap.hardcopyReceivedAt] || '').trim(),
        hardcopyReceivedBy: String(row[colMap.hardcopyReceivedBy] || '').trim()
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

    filtered.reverse();

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

function updateTicketStatus(ticketId, newStatus, additionalNote, userContext) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    if (!sheet) return { status: 'error', message: 'Sheet Tickets tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data tiket' };

    var lastCol = Math.max(sheet.getLastColumn(), 16);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var idCol = 0, statusCol = 10, noteCol = 9, updateCol = 11, closedCol = 12;

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'id') idCol = c;
      else if (h === 'statustiket' || h === 'status') statusCol = c;
      else if (h === 'note' || h === 'catatan') noteCol = c;
      else if (h === 'updatedat' || h === 'update') updateCol = c;
      else if (h === 'closedat') closedCol = c;
    }

    var targetRowIndex = -1;
    var currentStatus = '';

    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][idCol]).trim() === String(ticketId).trim()) {
        targetRowIndex = i + 1;
        currentStatus = String(allData[i][statusCol]).trim();
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { status: 'error', message: 'Tiket ' + ticketId + ' tidak ditemukan di database' };
    }

    var role = (userContext && userContext.roleKey) ? userContext.roleKey.toLowerCase() : '';

    // Validasi Rules Khusus Role FSE:
    if (role === 'fse') {
      var curLower = currentStatus.toLowerCase();
      var newLower = newStatus.toLowerCase();

      // Rule 1: Jika sudah appointment, FSE tidak bisa kembali ke respon
      if (curLower === 'appointment' && newLower === 'respon') {
        return { status: 'error', message: 'Tiket yang sudah di tahap Appointment tidak dapat dikembalikan ke Respon!' };
      }

      // Rule 2: Jika sudah solving / solved, FSE tidak bisa kembali ke action mana pun
      if (curLower === 'solving' || curLower === 'solved') {
        return { status: 'error', message: 'Tiket sudah berstatus Solved dan telah dikunci. Menunggu verifikasi Closed oleh Monitoring.' };
      }

      // Rule 3: Jika cancel, hanya bisa diubah melalui role monitoring
      if (curLower === 'cancel' || curLower === 'batal') {
        return { status: 'error', message: 'Tiket yang dibatalkan (Cancel) hanya dapat diaktifkan kembali oleh Petugas Monitoring.' };
      }

      // Rule 4: FSE tidak boleh langsung menaikkan status menjadi Closed (wewenang Monitoring)
      if (newLower === 'closed') {
        return { status: 'error', message: 'Status Closed hanya dapat diterbitkan oleh Monitoring Officer setelah verifikasi pekerjaan selesai.' };
      }
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    sheet.getRange(targetRowIndex, statusCol + 1).setValue(newStatus);
    sheet.getRange(targetRowIndex, updateCol + 1).setValue(nowFormatted);

    // Jika status dinaikkan menjadi Closed, catat timestamp Closed_At
    if (newStatus.toLowerCase() === 'closed') {
      sheet.getRange(targetRowIndex, closedCol + 1).setValue(nowFormatted);
    }

    if (additionalNote && String(additionalNote).trim() !== '') {
      var curNote = sheet.getRange(targetRowIndex, noteCol + 1).getDisplayValue();
      var combinedNote = curNote ? (curNote + ' | [' + newStatus + ']: ' + additionalNote) : ('[' + newStatus + ']: ' + additionalNote);
      sheet.getRange(targetRowIndex, noteCol + 1).setValue(combinedNote);
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      ticketId: ticketId,
      newStatus: newStatus,
      updatedAt: nowFormatted
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function markHardcopyReceived(ticketId, adminUser) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    if (!sheet) return { status: 'error', message: 'Sheet Tickets tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 16);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var idCol = 0, hardcopyCol = 13, rcvDateCol = 14, rcvByCol = 15;

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket') idCol = c;
      else if (h === 'hardcopystatus' || h === 'statushardcopy') hardcopyCol = c;
      else if (h === 'hardcopyreceivedat' || h === 'tglditerima') rcvDateCol = c;
      else if (h === 'hardcopyreceivedby' || h === 'penerima') rcvByCol = c;
    }

    var targetRow = -1;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][idCol]).trim() === String(ticketId).trim()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return { status: 'error', message: 'Tiket ' + ticketId + ' tidak ditemukan' };
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var adminName = (adminUser && adminUser.name) ? adminUser.name : 'Administrator';

    sheet.getRange(targetRow, hardcopyCol + 1).setValue('Diterima');
    sheet.getRange(targetRow, rcvDateCol + 1).setValue(nowFormatted);
    sheet.getRange(targetRow, rcvByCol + 1).setValue(adminName);

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Hardcopy tiket ' + ticketId + ' berhasil diverifikasi & diterima oleh ' + adminName,
      receivedAt: nowFormatted,
      receivedBy: adminName
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
