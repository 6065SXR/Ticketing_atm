/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: tickets.gs
 * Fungsi: Logika Manajemen Tiket, Form Aksi Lapangan FSE (Appointment, Handling, Solved, Cancel, Pending),
 *         Penyimpanan Nomor FE Report ke Kolom Mandiri No_FE_Report,
 *         Penyimpanan Bukti Foto ke Google Drive (Direct URL Display),
 *         Preservasi Note Asli Dispatcher, serta Pengiriman Kolektif Berkas Fisik (Batch Dispatch).
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
      var targetHeaders = HEADERS.Tickets;
      ticketSheet.appendRow(targetHeaders);
    }

    var now = new Date();
    var updatedAt = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var idTiket = String(payload.idTiket || '').trim();
    if (!idTiket) {
      var genRes = generateNewTicketNumber();
      idTiket = (genRes && genRes.status === 'success') ? genRes.ticketId : ('TK-' + Date.now());
    }

    var bank = String(payload.bank || '').trim();
    var serialNumber = String(payload.serialNumber || '').trim();
    var idMesin = String(payload.idMesin || '').trim();
    var tanggalTiket = String(payload.tanggalTiket || '').trim() || Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
    var jamTiket = String(payload.jamTiket || '').trim() || Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');
    var tipeTiket = String(payload.tipeTiket || 'Corrective Maintenance (CM)').trim();
    var namaEngineer = String(payload.namaEngineer || '').trim();
    var problem = String(payload.problem || '').trim();
    var noteInput = String(payload.note || '').trim();
    var statusTiket = String(payload.statusTiket || 'Open').trim();
    var dispatcher = String(payload.dispatcher || 'Monitoring Officer').trim();
    var noteDispatcher = String(payload.noteDispatcher || noteInput || '-').trim();
    var picBank = String(payload.picBank || '').trim();
    var noFeReport = String(payload.noFeReport || '').trim();

    if (!picBank && bank) {
      var custRes = getCustomersData();
      if (custRes && custRes.data) {
        for (var k = 0; k < custRes.data.length; k++) {
          if (custRes.data[k].namaBank === bank) {
            picBank = custRes.data[k].kontakPic;
            break;
          }
        }
      }
    }
    if (!picBank) picBank = '-';

    var initialHistory = [
      {
        status: statusTiket,
        time: updatedAt,
        note: noteDispatcher,
        user: dispatcher
      }
    ];

    var waktuHandling = (statusTiket.toLowerCase() === 'handling') ? updatedAt : '';

    var newRow = [
      idTiket, bank, serialNumber, idMesin,
      tanggalTiket, jamTiket, tipeTiket, namaEngineer,
      problem, noteDispatcher, statusTiket, updatedAt,
      '', 'Belum Dikirim', '', '',
      '', '', '', waktuHandling,
      '', '', JSON.stringify(initialHistory),
      dispatcher, noteDispatcher, picBank,
      '', '', '', noFeReport, '', '', '',
      '', '', '', ''
    ];

    ticketSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Tiket [' + idTiket + '] berhasil diterbitkan!',
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
        note: noteDispatcher,
        statusTiket: statusTiket,
        dispatcher: dispatcher,
        noteDispatcher: noteDispatcher,
        picBank: picBank,
        noFeReport: noFeReport
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
      return { status: 'error', message: 'Sheet Tickets tidak ditemukan', tickets: [], page: 1, totalItems: 0, totalPages: 0 };
    }

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { status: 'success', tickets: [], page: 1, pageSize: pageSize || 10, totalItems: 0, totalPages: 0 };
    }

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var colMap = {
      idTiket: -1, bank: -1, serialNumber: -1, idMesin: -1,
      tanggalTiket: -1, jamTiket: -1, tipeTiket: -1, namaEngineer: -1,
      problem: -1, note: -1, statusTiket: -1, updatedAt: -1,
      closedAt: -1, hardcopyStatus: -1, hardcopyReceivedAt: -1, hardcopyReceivedBy: -1,
      timelineHistory: -1, dispatcher: -1, noteDispatcher: -1, picBank: -1,
      waktuPending: -1, alasanPending: -1, jadwalLanjutanPending: -1,
      noFeReport: -1, fotoFeReport: -1, fotoCancelEvidence: -1, fotoPendingEvidence: -1,
      metodePengiriman: -1, namaEkspedisi: -1, noResi: -1, jenisPengiriman: -1
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket' || h === 'nomortiket' || h === 'ticketid' || h === 'id') colMap.idTiket = c;
      else if (h === 'bank' || h === 'namabank' || h === 'customer') colMap.bank = c;
      else if (h === 'serialnumber' || h === 'sn' || h === 'noserial') colMap.serialNumber = c;
      else if (h === 'idmesin' || h === 'machineid' || h === 'mesin') colMap.idMesin = c;
      else if (h === 'tanggaltiket' || h === 'tanggal' || h === 'date') colMap.tanggalTiket = c;
      else if (h === 'jamtiket' || h === 'jam' || h === 'time') colMap.jamTiket = c;
      else if (h === 'tipetiket' || h === 'tipe' || h === 'type') colMap.tipeTiket = c;
      else if (h === 'namaengineer' || h === 'engineer' || h === 'fse' || h === 'teknisi') colMap.namaEngineer = c;
      else if (h === 'problem' || h === 'masalah' || h === 'keluhan') colMap.problem = c;
      else if (h === 'note' || h === 'catatan' || h === 'keterangan') colMap.note = c;
      else if (h === 'statustiket' || h === 'status') colMap.statusTiket = c;
      else if (h === 'updatedat' || h === 'update') colMap.updatedAt = c;
      else if (h === 'closedat') colMap.closedAt = c;
      else if (h === 'hardcopystatus') colMap.hardcopyStatus = c;
      else if (h === 'hardcopyreceivedat') colMap.hardcopyReceivedAt = c;
      else if (h === 'hardcopyreceivedby') colMap.hardcopyReceivedBy = c;
      else if (h === 'timelinehistory') colMap.timelineHistory = c;
      else if (h === 'dispatcher') colMap.dispatcher = c;
      else if (h === 'notedispatcher') colMap.noteDispatcher = c;
      else if (h === 'picbank' || h === 'vendorpic' || h === 'pic') colMap.picBank = c;
      else if (h === 'waktupending') colMap.waktuPending = c;
      else if (h === 'alasanpending') colMap.alasanPending = c;
      else if (h === 'jadwallanjutanpending') colMap.jadwalLanjutanPending = c;
      else if (h === 'nofereport' || h === 'nofe' || h === 'nomorfereport') colMap.noFeReport = c;
      else if (h === 'fotofereport') colMap.fotoFeReport = c;
      else if (h === 'fotocancelevidence') colMap.fotoCancelEvidence = c;
      else if (h === 'fotopendingevidence' || h === 'fotopending' || h === 'buktipending') colMap.fotoPendingEvidence = c;
      else if (h === 'metodepengiriman') colMap.metodePengiriman = c;
      else if (h === 'namaekspedisi') colMap.namaEkspedisi = c;
      else if (h === 'noresi' || h === 'resi') colMap.noResi = c;
      else if (h === 'jenispengiriman') colMap.jenisPengiriman = c;
    }

    var rawData = allData.slice(1);
    var filtered = [];
    var searchLower = (searchQuery || '').toLowerCase().trim();
    var statusFilterVal = (statusFilter || '').trim();
    var bankFilterVal = (bankFilter || '').trim();

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      if (!row.some(function(cell) { return String(cell).trim() !== ''; })) continue;

      var rawNote = colMap.note >= 0 ? String(row[colMap.note] || '').trim() : '';
      var rawNoteDisp = colMap.noteDispatcher >= 0 ? String(row[colMap.noteDispatcher] || '').trim() : '';
      
      var cleanNote = rawNoteDisp;
      if (!cleanNote || cleanNote === '-') cleanNote = rawNote;
      if (cleanNote.indexOf(' | [') !== -1) {
        cleanNote = cleanNote.split(' | [')[0].trim();
      }

      var rawNoFe = colMap.noFeReport >= 0 ? String(row[colMap.noFeReport] || '').trim() : '';
      if (!rawNoFe || rawNoFe === '-') {
        if (cleanNote && cleanNote.indexOf('[No. FE:') !== -1) {
          var mFe = cleanNote.match(/\[No\.\s*FE:\s*(.*?)\]/i);
          if (mFe) rawNoFe = mFe[1].trim();
        }
      }

      var histStr = colMap.timelineHistory >= 0 ? String(row[colMap.timelineHistory] || '[]').trim() : '[]';
      if ((!rawNoFe || rawNoFe === '-') && histStr && histStr !== '[]') {
        try {
          var parsedHist = JSON.parse(histStr);
          for (var hx = 0; hx < parsedHist.length; hx++) {
            if (parsedHist[hx].noFeReport) {
              rawNoFe = String(parsedHist[hx].noFeReport).trim();
              break;
            }
            if (parsedHist[hx].note && parsedHist[hx].note.indexOf('[No. FE:') !== -1) {
              var mFeH = parsedHist[hx].note.match(/\[No\.\s*FE:\s*(.*?)\]/i);
              if (mFeH) {
                rawNoFe = mFeH[1].trim();
                break;
              }
            }
          }
        } catch(e) {}
      }

      var item = {
        idTiket: colMap.idTiket >= 0 ? String(row[colMap.idTiket] || '').trim() : ('TK-' + (i + 1)),
        bank: colMap.bank >= 0 ? String(row[colMap.bank] || '-').trim() : '-',
        serialNumber: colMap.serialNumber >= 0 ? String(row[colMap.serialNumber] || '-').trim() : '-',
        idMesin: colMap.idMesin >= 0 ? String(row[colMap.idMesin] || '-').trim() : '-',
        tanggalTiket: colMap.tanggalTiket >= 0 ? String(row[colMap.tanggalTiket] || '-').trim() : '-',
        jamTiket: colMap.jamTiket >= 0 ? String(row[colMap.jamTiket] || '-').trim() : '-',
        tipeTiket: colMap.tipeTiket >= 0 ? String(row[colMap.tipeTiket] || 'CM').trim() : 'CM',
        namaEngineer: colMap.namaEngineer >= 0 ? String(row[colMap.namaEngineer] || '-').trim() : '-',
        problem: colMap.problem >= 0 ? String(row[colMap.problem] || '-').trim() : '-',
        note: cleanNote || '-',
        statusTiket: colMap.statusTiket >= 0 ? String(row[colMap.statusTiket] || 'Open').trim() : 'Open',
        updatedAt: colMap.updatedAt >= 0 ? String(row[colMap.updatedAt] || '-').trim() : '-',
        closedAt: colMap.closedAt >= 0 ? String(row[colMap.closedAt] || '').trim() : '',
        hardcopyStatus: colMap.hardcopyStatus >= 0 ? String(row[colMap.hardcopyStatus] || 'Belum Dikirim').trim() : 'Belum Dikirim',
        hardcopyReceivedAt: colMap.hardcopyReceivedAt >= 0 ? String(row[colMap.hardcopyReceivedAt] || '').trim() : '',
        hardcopyReceivedBy: colMap.hardcopyReceivedBy >= 0 ? String(row[colMap.hardcopyReceivedBy] || '').trim() : '',
        timelineHistory: histStr,
        dispatcher: colMap.dispatcher >= 0 ? String(row[colMap.dispatcher] || 'Monitoring Officer').trim() : 'Monitoring Officer',
        noteDispatcher: cleanNote || '-',
        picBank: colMap.picBank >= 0 ? String(row[colMap.picBank] || '-').trim() : '-',
        waktuPending: colMap.waktuPending >= 0 ? String(row[colMap.waktuPending] || '').trim() : '',
        alasanPending: colMap.alasanPending >= 0 ? String(row[colMap.alasanPending] || '').trim() : '',
        jadwalLanjutanPending: colMap.jadwalLanjutanPending >= 0 ? String(row[colMap.jadwalLanjutanPending] || '').trim() : '',
        noFeReport: rawNoFe || '-',
        fotoFeReport: colMap.fotoFeReport >= 0 ? String(row[colMap.fotoFeReport] || '').trim() : '',
        fotoCancelEvidence: colMap.fotoCancelEvidence >= 0 ? String(row[colMap.fotoCancelEvidence] || '').trim() : '',
        fotoPendingEvidence: colMap.fotoPendingEvidence >= 0 ? String(row[colMap.fotoPendingEvidence] || '').trim() : '',
        metodePengiriman: colMap.metodePengiriman >= 0 ? String(row[colMap.metodePengiriman] || '').trim() : '',
        namaEkspedisi: colMap.namaEkspedisi >= 0 ? String(row[colMap.namaEkspedisi] || '').trim() : '',
        noResi: colMap.noResi >= 0 ? String(row[colMap.noResi] || '').trim() : '',
        jenisPengiriman: colMap.jenisPengiriman >= 0 ? String(row[colMap.jenisPengiriman] || '').trim() : ''
      };

      if (statusFilterVal && statusFilterVal !== 'ALL') {
        if (item.statusTiket.toLowerCase() !== statusFilterVal.toLowerCase()) continue;
      }

      if (bankFilterVal && bankFilterVal !== 'ALL') {
        if (item.bank.toLowerCase().indexOf(bankFilterVal.toLowerCase()) === -1) continue;
      }

      if (searchLower !== '') {
        var match = (
          item.idTiket.toLowerCase().indexOf(searchLower) !== -1 ||
          item.bank.toLowerCase().indexOf(searchLower) !== -1 ||
          item.serialNumber.toLowerCase().indexOf(searchLower) !== -1 ||
          item.idMesin.toLowerCase().indexOf(searchLower) !== -1 ||
          item.problem.toLowerCase().indexOf(searchLower) !== -1 ||
          item.namaEngineer.toLowerCase().indexOf(searchLower) !== -1 ||
          item.noFeReport.toLowerCase().indexOf(searchLower) !== -1
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

    return {
      status: 'success',
      tickets: filtered.slice(startIndex, endIndex),
      page: currentPage,
      pageSize: validPageSize,
      totalItems: totalItems,
      totalPages: totalPages
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function saveBase64ImageToDrive(base64Data, fileName, folderName) {
  try {
    if (!base64Data || String(base64Data).trim() === '') return '';
    var split = base64Data.split(',');
    var contentType = 'image/jpeg';
    var rawBase64 = base64Data;
    if (split.length > 1) {
      var match = split[0].match(/:(.*?);/);
      if (match) contentType = match[1];
      rawBase64 = split[1];
    }
    var decoded = Utilities.base64Decode(rawBase64);
    var blob = Utilities.newBlob(decoded, contentType, fileName || ('bukti_' + Date.now() + '.jpg'));
    
    var folderTargetName = folderName || 'ATM_Ticketing_Uploads';
    var folderIter = DriveApp.getFoldersByName(folderTargetName);
    var folder = folderIter.hasNext() ? folderIter.next() : DriveApp.createFolder(folderTargetName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileId = file.getId();
    return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1200';
  } catch (err) {
    Logger.log('Drive Upload Exception: ' + err.toString());
    return 'File_Local_Fallback_' + Date.now();
  }
}

function updateTicketStatus(ticketId, newStatus, actionPayload, userName) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    if (!sheet) return { status: 'error', message: 'Sheet Tickets tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data tiket' };

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var col = {
      id: -1, status: -1, update: -1, closed: -1, hardcopy: -1,
      respon: -1, appt: -1, arrival: -1, handling: -1, wait: -1, solving: -1,
      timeline: -1, waktuPending: -1, alasanPending: -1, jadwalLanjutan: -1,
      noFeReport: -1, fotoFe: -1, fotoCancel: -1, fotoPending: -1
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket') col.id = c;
      else if (h === 'statustiket' || h === 'status') col.status = c;
      else if (h === 'updatedat' || h === 'update') col.update = c;
      else if (h === 'closedat') col.closed = c;
      else if (h === 'hardcopystatus') col.hardcopy = c;
      else if (h === 'wakturespon') col.respon = c;
      else if (h === 'waktuappointment') col.appt = c;
      else if (h === 'waktuarrival') col.arrival = c;
      else if (h === 'waktuhandling') col.handling = c;
      else if (h === 'waktuwaitingvendor') col.wait = c;
      else if (h === 'waktusolving') col.solving = c;
      else if (h === 'timelinehistory') col.timeline = c;
      else if (h === 'waktupending') col.waktuPending = c;
      else if (h === 'alasanpending') col.alasanPending = c;
      else if (h === 'jadwallanjutanpending') col.jadwalLanjutan = c;
      else if (h === 'nofereport' || h === 'nofe' || h === 'nomorfereport') col.noFeReport = c;
      else if (h === 'fotofereport') col.fotoFe = c;
      else if (h === 'fotocancelevidence') col.fotoCancel = c;
      else if (h === 'fotopendingevidence' || h === 'fotopending' || h === 'buktipending') col.fotoPending = c;
    }

    var targetRowIndex = -1;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][col.id >= 0 ? col.id : 0]).trim() === String(ticketId).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return { status: 'error', message: 'Tiket ' + ticketId + ' tidak ditemukan di database' };
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    var p = (typeof actionPayload === 'object' && actionPayload !== null) ? actionPayload : { note: String(actionPayload || '') };
    var actionNote = p.note || '';
    var feReportValue = p.noFeReport || '';

    if (col.status !== -1) sheet.getRange(targetRowIndex, col.status + 1).setValue(newStatus);
    if (col.update !== -1) sheet.getRange(targetRowIndex, col.update + 1).setValue(nowFormatted);

    var statLower = newStatus.toLowerCase();

    if (statLower === 'respon' && col.respon !== -1) {
      sheet.getRange(targetRowIndex, col.respon + 1).setValue(nowFormatted);
    } else if ((statLower === 'appointment' || statLower === 're-appointment') && col.appt !== -1) {
      sheet.getRange(targetRowIndex, col.appt + 1).setValue(nowFormatted);
      if (p.permintaanAppt && p.tglJanjian) {
        actionNote = '[Permintaan: ' + p.permintaanAppt + ' | Jadwal: ' + p.tglJanjian + ' ' + (p.jamJanjian || '') + '] ' + (p.note || '');
      }
    } else if (statLower === 'arrival' && col.arrival !== -1) {
      sheet.getRange(targetRowIndex, col.arrival + 1).setValue(nowFormatted);
    } else if (statLower === 'handling' && col.handling !== -1) {
      sheet.getRange(targetRowIndex, col.handling + 1).setValue(nowFormatted);
    } else if (statLower === 'waiting vendor' && col.wait !== -1) {
      sheet.getRange(targetRowIndex, col.wait + 1).setValue(nowFormatted);
    } else if (statLower === 'pending') {
      if (col.waktuPending !== -1) sheet.getRange(targetRowIndex, col.waktuPending + 1).setValue(nowFormatted);
      if (col.alasanPending !== -1 && p.alasanPending) sheet.getRange(targetRowIndex, col.alasanPending + 1).setValue(p.alasanPending);
      if (col.jadwalLanjutan !== -1 && p.tglJamLanjutan) sheet.getRange(targetRowIndex, col.jadwalLanjutan + 1).setValue(p.tglJamLanjutan);
      
      if (col.fotoPending !== -1 && p.fotoBase64) {
        var linkPending = saveBase64ImageToDrive(p.fotoBase64, 'Evidence_Pending_Lokasi_' + ticketId + '.jpg', 'ATM_Pending_Evidences');
        sheet.getRange(targetRowIndex, col.fotoPending + 1).setValue(linkPending);
      }
      actionNote = '[Pending: ' + (p.alasanPending || '-') + ' | Lanjutan: ' + (p.tglJamLanjutan || '-') + '] ' + (p.note || '');
    } else if (statLower === 'solving' || statLower === 'solved') {
      if (col.solving !== -1) sheet.getRange(targetRowIndex, col.solving + 1).setValue(nowFormatted);
      
      if (!feReportValue && p.note && p.note.indexOf('[No. FE:') !== -1) {
        var mFe = p.note.match(/\[No\.\s*FE:\s*(.*?)\]/i);
        if (mFe) feReportValue = mFe[1].trim();
      }

      if (col.noFeReport !== -1 && feReportValue) {
        sheet.getRange(targetRowIndex, col.noFeReport + 1).setValue(feReportValue);
      }

      if (col.fotoFe !== -1 && p.fotoBase64) {
        var linkFe = saveBase64ImageToDrive(p.fotoBase64, 'FE_Report_' + ticketId + '.jpg', 'ATM_FE_Reports');
        sheet.getRange(targetRowIndex, col.fotoFe + 1).setValue(linkFe);
      }
    } else if (statLower === 'cancel') {
      if (col.fotoCancel !== -1 && p.fotoBase64) {
        var linkCancel = saveBase64ImageToDrive(p.fotoBase64, 'Evidence_Cancel_' + ticketId + '.jpg', 'ATM_Cancel_Evidences');
        sheet.getRange(targetRowIndex, col.fotoCancel + 1).setValue(linkCancel);
      }
    } else if (statLower === 'closed') {
      if (col.closed !== -1) sheet.getRange(targetRowIndex, col.closed + 1).setValue(nowFormatted);
      if (col.hardcopy !== -1) sheet.getRange(targetRowIndex, col.hardcopy + 1).setValue('Belum Dikirim');
    }

    if (col.timeline !== -1) {
      var curHistStr = sheet.getRange(targetRowIndex, col.timeline + 1).getDisplayValue();
      var historyArr = [];
      if (curHistStr) {
        try { historyArr = JSON.parse(curHistStr); } catch(e) {}
      }
      var timelineItem = {
        status: newStatus,
        time: nowFormatted,
        note: actionNote || '-',
        user: userName || 'FSE'
      };
      if (feReportValue) {
        timelineItem.noFeReport = feReportValue;
      }
      historyArr.push(timelineItem);
      sheet.getRange(targetRowIndex, col.timeline + 1).setValue(JSON.stringify(historyArr));
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      ticketId: ticketId,
      newStatus: newStatus,
      updatedAt: nowFormatted,
      note: actionNote,
      noFeReport: feReportValue
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function submitBatchDispatchShipping(ticketIds, shippingPayload, userName) {
  try {
    var ss = getSpreadsheet();
    var sheet = findTicketsSheet(ss);
    if (!sheet) return { status: 'error', message: 'Sheet Tickets tidak ditemukan' };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'error', message: 'Tidak ada data tiket' };

    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var col = {
      id: -1, hardcopy: -1, timeline: -1,
      metode: -1, ekspedisi: -1, resi: -1, jenis: -1
    };

    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket') col.id = c;
      else if (h === 'hardcopystatus') col.hardcopy = c;
      else if (h === 'timelinehistory') col.timeline = c;
      else if (h === 'metodepengiriman') col.metode = c;
      else if (h === 'namaekspedisi') col.ekspedisi = c;
      else if (h === 'noresi' || h === 'resi') col.resi = c;
      else if (h === 'jenispengiriman') col.jenis = c;
    }

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var p = shippingPayload || {};
    var metode = String(p.metode || 'Kirim Sendiri ke Kantor').trim();
    var ekspedisi = String(p.namaEkspedisi || '-').trim();
    var resi = String(p.noResi || '-').trim();
    var jenis = String(p.jenisPengiriman || '-').trim();

    var logDesc = (metode === 'Kirim Sendiri ke Kantor') 
      ? 'Hardcopy dikirim langsung oleh teknisi ke kantor' 
      : ('Hardcopy dikirim via ' + ekspedisi + ' (' + jenis + ') | Resi: ' + resi);

    var updatedCount = 0;

    for (var i = 1; i < allData.length; i++) {
      var rowTicketId = String(allData[i][col.id >= 0 ? col.id : 0]).trim();
      if (ticketIds.indexOf(rowTicketId) !== -1) {
        var rowIndex = i + 1;
        if (col.hardcopy !== -1) sheet.getRange(rowIndex, col.hardcopy + 1).setValue('Sedang Dikirim');
        if (col.metode !== -1) sheet.getRange(rowIndex, col.metode + 1).setValue(metode);
        if (col.ekspedisi !== -1) sheet.getRange(rowIndex, col.ekspedisi + 1).setValue(ekspedisi);
        if (col.resi !== -1) sheet.getRange(rowIndex, col.resi + 1).setValue(resi);
        if (col.jenis !== -1) sheet.getRange(rowIndex, col.jenis + 1).setValue(jenis);

        if (col.timeline !== -1) {
          var curHistStr = sheet.getRange(rowIndex, col.timeline + 1).getDisplayValue();
          var historyArr = [];
          if (curHistStr) {
            try { historyArr = JSON.parse(curHistStr); } catch(e) {}
          }
          historyArr.push({
            status: 'Dispatch Hardcopy',
            time: nowFormatted,
            note: logDesc,
            user: userName || 'FSE'
          });
          sheet.getRange(rowIndex, col.timeline + 1).setValue(JSON.stringify(historyArr));
        }

        updatedCount++;
      }
    }

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: updatedCount + ' berkas fisik berhasil diproses untuk pengiriman!',
      updatedCount: updatedCount,
      shippingData: {
        metode: metode,
        namaEkspedisi: ekspedisi,
        noResi: resi,
        jenisPengiriman: jenis,
        dispatchedAt: nowFormatted
      }
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
    var lastCol = sheet.getLastColumn();
    var allData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    var headers = allData[0];

    var idCol = 0, statusCol = 13, timeCol = 14, byCol = 15;
    for (var c = 0; c < headers.length; c++) {
      var h = String(headers[c] || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      if (h === 'idtiket' || h === 'notiket') idCol = c;
      else if (h === 'hardcopystatus') statusCol = c;
      else if (h === 'hardcopyreceivedat') timeCol = c;
      else if (h === 'hardcopyreceivedby') byCol = c;
    }

    var targetRowIndex = -1;
    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][idCol]).trim() === String(ticketId).trim()) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) return { status: 'error', message: 'Tiket tidak ditemukan' };

    var now = new Date();
    var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    var receiver = (adminUser && adminUser.name) ? adminUser.name : 'Admin Ops';

    sheet.getRange(targetRowIndex, statusCol + 1).setValue('Diterima');
    sheet.getRange(targetRowIndex, timeCol + 1).setValue(nowFormatted);
    sheet.getRange(targetRowIndex, byCol + 1).setValue(receiver);

    SpreadsheetApp.flush();

    return {
      status: 'success',
      message: 'Dokumen hardcopy Tiket ' + ticketId + ' terverifikasi!',
      receivedAt: nowFormatted,
      receivedBy: receiver
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
