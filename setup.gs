/**
 * Hardiansyah Fam - ATM Problem Ticketing System
 * File: setup.gs
 * Fungsi: Inisialisasi Database, Safe Migrate (Tanpa menghapus data),
 *         Pencatatan Milestone Downtime, Pending Pause Timestamps,
 *         Audit Dispatcher, PIC Bank, Kolom Mandiri No_FE_Report,
 *         serta Kolom Pelacakan Ekspedisi Pengiriman Dokumen Fisik.
 */

var SHEET_NAMES = {
  TICKETS: 'Tickets',
  CUSTOMERS: 'Customers',
  MACHINES: 'Machines',
  ENGINEERS: 'Engineers',
  COUNTER: 'Counter_Ticket',
  USERS: 'Users',
  FSE: 'FSE'
};

var HEADERS = {
  Tickets: [
    'ID_Tiket', 'Bank', 'Serial_Number', 'ID_Mesin', 
    'Tanggal_Tiket', 'Jam_Tiket', 'Tipe_Tiket', 'Nama_Engineer', 
    'Problem', 'Note', 'Status_Tiket', 'Updated_At',
    'Closed_At', 'Hardcopy_Status', 'Hardcopy_Received_At', 'Hardcopy_Received_By',
    'Waktu_Respon', 'Waktu_Appointment', 'Waktu_Arrival', 'Waktu_Handling', 
    'Waktu_Waiting_Vendor', 'Waktu_Solving', 'Timeline_History',
    'Dispatcher', 'Note_Dispatcher', 'PIC_Bank',
    'Waktu_Pending', 'Alasan_Pending', 'Jadwal_Lanjutan_Pending',
    'No_FE_Report', 'Foto_FE_Report', 'Foto_Cancel_Evidence', 'Foto_Pending_Evidence',
    'Metode_Pengiriman', 'Nama_Ekspedisi', 'No_Resi', 'Jenis_Pengiriman'
  ],
  Customers: [
    'ID_Customer', 'Nama_Bank', 'Kontak_PIC', 
    'Telepon_PIC', 'Email_Bank', 'Alamat'
  ],
  Machines: [
    'ID_Mesin', 'Serial_Number', 'Nama_Bank', 
    'Lokasi_ATM', 'Tipe_Mesin', 'Status_Mesin', 'FSE_Pengelola'
  ],
  Engineers: [
    'ID_Engineer', 'Nama_Engineer', 'Wilayah_Tugas', 
    'No_HP', 'Status_Aktif'
  ],
  Counter_Ticket: [
    'Tanggal', 'Last_Sequence'
  ],
  Users: [
    'User_ID', 'Nama_Lengkap', 'No_HP', 'Password', 
    'Role', 'Wilayah_Tugas', 'Status_Akun', 'Created_By', 'Created_At'
  ],
  FSE: [
    'User_ID', 'Nama_Lengkap', 'No_HP', 'Password', 
    'Role', 'Wilayah_Tugas', 'Status_Akun', 'Jumlah_Kelolaan_Mesin', 
    'Created_By', 'Created_At'
  ]
};

function runSafeMigration() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logMessages = [];

  for (var key in SHEET_NAMES) {
    var sheetName = SHEET_NAMES[key];
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      logMessages.push('Sheet [' + sheetName + '] berhasil dibuat.');
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var targetHeaders = HEADERS[sheetName];

    if (lastRow === 0 || lastCol === 0) {
      sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
      sheet.getRange(1, 1, 1, targetHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      logMessages.push('Header diinisialisasi untuk sheet [' + sheetName + '].');

      populateInitialDummyData(sheet, sheetName, ss);
    } else {
      var currentHeaderRange = sheet.getRange(1, 1, 1, targetHeaders.length);
      currentHeaderRange.setValues([targetHeaders]);
      sheet.getRange(1, 1, 1, targetHeaders.length).setFontWeight('bold').setBackground('#0f766e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      logMessages.push('Header diverifikasi & diperbarui dengan aman untuk [' + sheetName + '].');
    }
  }

  SpreadsheetApp.flush();
  Logger.log('Safe Migration Selesai:\n' + logMessages.join('\n'));
  return { success: true, logs: logMessages };
}

function populateInitialDummyData(sheet, sheetName, ss) {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
  var timeStr = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');
  var yymmdd = Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd');
  var nowFormatted = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  if (sheetName === SHEET_NAMES.USERS) {
    var usersData = [
      ['USR-001', 'Monitoring Officer', '081200000002', 'monitoring123', 'Monitoring', 'NOC Pusat', 'Aktif', 'System', nowFormatted],
      ['USR-002', 'Super Administrator', '081100000001', 'admin123', 'Admin', 'Head Office', 'Aktif', 'System', nowFormatted],
      ['USR-003', 'Rizky Hardiansyah', '081299881122', 'fse123', 'FSE', 'Jakarta Pusat & Barat', 'Aktif', 'System', nowFormatted],
      ['USR-004', 'Fajar Ramadhan', '081377665544', 'fse123', 'FSE', 'Jakarta Selatan & Depok', 'Aktif', 'System', nowFormatted],
      ['USR-005', 'Dimas Kurniawan', '081733445566', 'fse123', 'FSE', 'Jakarta Timur & Bekasi', 'Aktif', 'System', nowFormatted]
    ];
    sheet.getRange(2, 1, usersData.length, usersData[0].length).setValues(usersData);
  }

  if (sheetName === SHEET_NAMES.FSE) {
    var fseData = [
      ['USR-003', 'Rizky Hardiansyah', '081299881122', 'fse123', 'FSE', 'Jakarta Pusat & Barat', 'Aktif', 3, 'System', nowFormatted],
      ['USR-004', 'Fajar Ramadhan', '081377665544', 'fse123', 'FSE', 'Jakarta Selatan & Depok', 'Aktif', 2, 'System', nowFormatted],
      ['USR-005', 'Dimas Kurniawan', '081733445566', 'fse123', 'FSE', 'Jakarta Timur & Bekasi', 'Aktif', 3, 'System', nowFormatted]
    ];
    sheet.getRange(2, 1, fseData.length, fseData[0].length).setValues(fseData);
  }

  if (sheetName === SHEET_NAMES.CUSTOMERS) {
    var customersData = [
      ['CUST-001', 'Bank Central Asia (BCA)', 'Bambang Sudibyo', '081234567890', 'pic.atm@bca.co.id', 'Jl. M.H. Thamrin No. 1, Jakarta Pusat'],
      ['CUST-002', 'Bank Mandiri', 'Siti Rahmawati', '081398765432', 'atm.ops@bankmandiri.co.id', 'Jl. Gatot Subroto Kav. 36-38, Jakarta Selatan'],
      ['CUST-003', 'Bank Rakyat Indonesia (BRI)', 'Eko Prasetyo', '081122334455', 'monitoring.atm@bri.co.id', 'Jl. Jenderal Sudirman No. 44, Jakarta Pusat'],
      ['CUST-004', 'Bank Negara Indonesia (BNI)', 'Dewi Lestari', '081566778899', 'fse.coordination@bni.co.id', 'Jl. Jenderal Sudirman Kav. 1, Jakarta Pusat'],
      ['CUST-005', 'Bank CIMB Niaga', 'Andi Pratama', '081799887766', 'atm.vendor@cimbniaga.co.id', 'Graha CIMB Niaga, Jl. Jend. Sudirman, Jakarta']
    ];
    sheet.getRange(2, 1, customersData.length, customersData[0].length).setValues(customersData);
  }

  if (sheetName === SHEET_NAMES.MACHINES) {
    var machinesData = [
      ['ATM-BCA-01', 'SN-BCA-99201', 'Bank Central Asia (BCA)', 'KCP Grand Indonesia Lt. 1', 'Diebold Nixdorf ProCash 280', 'Active', 'Rizky Hardiansyah'],
      ['ATM-BCA-02', 'SN-BCA-99202', 'Bank Central Asia (BCA)', 'SPBU Kuningan Rasuna Said', 'NCR Personas 77', 'Active', 'Rizky Hardiansyah'],
      ['ATM-MDR-01', 'SN-MDR-44011', 'Bank Mandiri', 'Plaza Senayan Ground Floor', 'Hyosung Monimax 5600', 'Active', 'Fajar Ramadhan'],
      ['ATM-MDR-02', 'SN-MDR-44012', 'Bank Mandiri', 'Stasiun Gambir Hall Utama', 'Diebold Nixdorf CS 7700', 'Active', 'Rizky Hardiansyah'],
      ['ATM-BRI-01', 'SN-BRI-11055', 'Bank Rakyat Indonesia (BRI)', 'Pasar Tanah Abang Blok A', 'NCR SelfServ 26', 'Active', 'Dimas Kurniawan'],
      ['ATM-BRI-02', 'SN-BRI-11056', 'Bank Rakyat Indonesia (BRI)', 'RS Cipto Mangunkusumo Lobby', 'Hyosung Monimax 7600T', 'Active', 'Dimas Kurniawan'],
      ['ATM-BNI-01', 'SN-BNI-77881', 'Bank Negara Indonesia (BNI)', 'Mall Kelapa Gading 3 Lt. Dasar', 'Diebold Nixdorf CS 5500', 'Active', 'Fajar Ramadhan'],
      ['ATM-CIMB-01', 'SN-CIMB-3341', 'Bank CIMB Niaga', 'Bandara Soekarno Hatta Terminal 3', 'NCR SelfServ 6632', 'Active', 'Dimas Kurniawan']
    ];
    sheet.getRange(2, 1, machinesData.length, machinesData[0].length).setValues(machinesData);
  }

  if (sheetName === SHEET_NAMES.COUNTER) {
    var counterData = [[yymmdd, 3]];
    sheet.getRange(2, 1, 1, 2).setValues(counterData);
  }

  if (sheetName === SHEET_NAMES.TICKETS) {
    var hist1 = JSON.stringify([
      { status: 'Open', time: dateStr + ' 08:15:20', note: 'Tiket diterbitkan oleh Dispatcher', user: 'Monitoring' },
      { status: 'Handling', time: dateStr + ' ' + timeStr, note: 'Tiba di lokasi dan mulai inspeksi', user: 'Rizky Hardiansyah' }
    ]);
    var hist2 = JSON.stringify([
      { status: 'Open', time: dateStr + ' 09:30:10', note: 'PM Rutin Kuartal 3', user: 'Monitoring' },
      { status: 'Solving', time: dateStr + ' 11:00:00', note: '[No. FE: FE-2026-0911-001] Pembersihan dan kalibrasi selesai', user: 'Fajar Ramadhan', noFeReport: 'FE-2026-0911-001' },
      { status: 'Closed', time: dateStr + ' 11:30:00', note: 'Verifikasi monitoring selesai', user: 'Monitoring' }
    ]);
    var hist3 = JSON.stringify([
      { status: 'Open', time: dateStr + ' 10:05:44', note: 'Tiket diterbitkan', user: 'Monitoring' },
      { status: 'Appointment', time: dateStr + ' ' + timeStr, note: 'Permintaan: FSE | Janjian jam 13:00', user: 'Dimas Kurniawan' }
    ]);

    var ticketsData = [
      [
        yymmdd + '-0001', 'Bank Central Asia (BCA)', 'SN-BCA-99201', 'ATM-BCA-01',
        dateStr, '08:15:20', 'Corrective Maintenance (CM)', 'Rizky Hardiansyah',
        'Mesin Card Reader error code 42 (Card Jammed)', 'Prioritas tinggi, nasabah antre.', 'Handling', dateStr + ' ' + timeStr,
        '', 'Belum Dikirim', '', '',
        dateStr + ' 08:20:00', dateStr + ' 08:35:00', dateStr + ' 09:00:00', dateStr + ' ' + timeStr, '', '', hist1,
        'Monitoring Officer', 'Mohon respon segera karena antrean ramai', 'Bambang Sudibyo',
        '', '', '', '', '', '', '',
        '', '', '', ''
      ],
      [
        yymmdd + '-0002', 'Bank Mandiri', 'SN-MDR-44011', 'ATM-MDR-01',
        dateStr, '09:30:10', 'Preventive Maintenance (PM)', 'Fajar Ramadhan',
        'Jadwal pembersihan modul dispenser & check sensor', 'PM Rutin Kuartal 3', 'Closed', dateStr + ' ' + timeStr,
        dateStr + ' 11:30:00', 'Belum Dikirim', '', '',
        dateStr + ' 09:35:00', dateStr + ' 09:50:00', dateStr + ' 10:15:00', dateStr + ' 10:20:00', '', dateStr + ' 11:00:00', hist2,
        'Monitoring Officer', 'Pemeriksaan rutin modul dispenser', 'Siti Rahmawati',
        '', '', '', 'FE-2026-0911-001', '', '', '',
        '', '', '', ''
      ],
      [
        yymmdd + '-0003', 'Bank Rakyat Indonesia (BRI)', 'SN-BRI-11055', 'ATM-BRI-01',
        dateStr, '10:05:44', 'Corrective Maintenance (CM)', 'Dimas Kurniawan',
        'Receipt Printer Out of Paper & Cutter Jam', 'Kertas cadangan dibawa FSE', 'Appointment', dateStr + ' ' + timeStr,
        '', 'Belum Dikirim', '', '',
        dateStr + ' 10:10:00', dateStr + ' ' + timeStr, '', '', '', '', hist3,
        'Monitoring Officer', 'Koordinasikan dengan PIC Pasar Tanah Abang', 'Eko Prasetyo',
        '', '', '', '', '', '', '',
        '', '', '', ''
      ]
    ];
    sheet.getRange(2, 1, ticketsData.length, ticketsData[0].length).setValues(ticketsData);
  }
}
