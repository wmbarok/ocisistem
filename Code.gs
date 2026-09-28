// =========================================================================
// GOOGLE APPS SCRIPT BACKEND - PT. OUTVENTURE CAKRAWALA INDONESIA
// Database: DB_FLIGHT & DB_HOTEL (FIXED VERSION)
// =========================================================================

const SPREADSHEET_ID = "1acEtuCeWPwSMR9lqILIH_r7Cyi8a8mztklJQxm4VS80";

const SHEETS = {
  FLIGHT:  { name: "DB_FLIGHT",           headerRow: 2, dataStart: 4, lastCol: 23 },
  HOTEL:   { name: "DB_HOTEL",            headerRow: 2, dataStart: 4, lastCol: 21 },
  USERS:   { name: "USERS",               headerRow: 1, dataStart: 2, lastCol: 5  },
  CLIENTS: { name: "CLIENTS",             headerRow: 1, dataStart: 2, lastCol: 6  },
  MAP:     { name: "CLIENT_INVOICE_MAP",  headerRow: 1, dataStart: 2, lastCol: 3  },
  LOG:     { name: "LOG_PRINT",           headerRow: 1, dataStart: 2, lastCol: 5  }
};

const COL_F = {
  idUnik:1, inv:2, pax:3, jabatan:4, hariTanggal:5,
  invoiceDate:6, nameBP:7, maskapai:8, flightNo:9,
  timeDep:10, depAt:11, timeArr:12, arrAt:13, kelas:14,
  pnr:15, ticketNo:16, invoiceNo:17,
  harga:18, addOn:19, totalAmount:20,
  tglBayar:21, nominalBayar:22, sisa:23
};

const COL_H = {
  idUnik:1, inv:2, pax:3, hariTanggal:4, invoiceDate:5,
  invoiceNo:6, bookingRef:7, propertyName:8, propertyAddress:9,
  typeRoom:10, meal:11, checkIn:12, checkOut:13, extraBed:14,
  roomQty:15, harga:16, addOn:17, totalAmount:18,
  tglBayar:19, nominalBayar:20, sisa:21
};

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ============ WEB APP ============
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getData') {
    try {
      var db = readDatabase();
      return jsonOut({
        status: "success",
        flights:   db.flights,
        hotels:    db.hotels,
        users:     db.users,
        clients:   db.clients,
        clientMap: db.clientMap,
        logPrint:  db.logPrint
      });
    } catch (err) {
      return jsonOut({ status: "error", message: err.toString() });
    }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PT. Outventure Cakrawala Indonesia - Sistem Tiketing & Travel')
    .setFaviconUrl('https://drive.google.com/uc?id=19YVSFif9noD5_eJ-gaie7TRPFsY8v1TI&export=download&format=png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

function doPost(e) {
  try {
    var data;
    if (e.parameter && e.parameter.data) data = JSON.parse(e.parameter.data);
    else if (e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
    else return jsonOut({status:"error", message:"No data received"});

    var action = data.action || "create";
    var type = data.type;

    if (type === "FLIGHT") {
      if (action === "create")       return createFlightDB(data);
      if (action === "update")       return updateFlightDB(data);
      if (action === "delete")       return deleteFlightDB(data.idUnik);
      if (action === "markPayment")  return markPaymentFlightDB(data);
    } else if (type === "HOTEL") {
      if (action === "create")       return createHotelDB(data);
      if (action === "update")       return updateHotelDB(data);
      if (action === "delete")       return deleteHotelDB(data.idUnik);
      if (action === "markPayment")  return markPaymentHotelDB(data);
    } else if (type === "USER") {
      return handleUserAction(data);
    } else if (type === "CLIENT") {
      return handleClientAction(data);
    } else if (type === "CLIENT_MAP") {
      return handleClientMapAction(data);
    } else if (type === "LOG_PRINT") {
      return handleLogPrint(data);
    }

    return jsonOut({status:"success", message:"Data tersimpan!"});
  } catch (err) {
    return jsonOut({status:"error", message: err.toString()});
  }
}

// ============ READ DATABASE ============
function readDatabase() {
  return {
    flights:   readDBFlight(),
    hotels:    readDBHotel(),
    users:     readDBUsers(),
    clients:   readDBClients(),
    clientMap: readDBClientMap(),
    logPrint:  readDBLogPrint()
  };
}

// ============ DB_FLIGHT ============
function readDBFlight() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SHEETS.FLIGHT.name);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.FLIGHT.dataStart) return [];

  var numRows = lastRow - SHEETS.FLIGHT.dataStart + 1;
  var data = sheet.getRange(SHEETS.FLIGHT.dataStart, 1, numRows, SHEETS.FLIGHT.lastCol).getValues();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var idUnik = String(row[COL_F.idUnik - 1] || '').trim();
    var inv = String(row[COL_F.inv - 1] || '').trim();
    var invNo = String(row[COL_F.invoiceNo - 1] || '').trim();
    var pax = String(row[COL_F.pax - 1] || '').trim();

    if (!idUnik && !pax && !inv && !invNo) continue;
    if (!idUnik && !pax) continue;

    var harga = parseRupiah(row[COL_F.harga - 1]);
    var addOn = parseRupiah(row[COL_F.addOn - 1]);
    var total = parseRupiah(row[COL_F.totalAmount - 1]);
    var nominal = parseRupiah(row[COL_F.nominalBayar - 1]);
    var sisa = parseRupiah(row[COL_F.sisa - 1]);
    var tglBayar = fmtDateCell(row[COL_F.tglBayar - 1]);

    var status = 'UNPAID';
    if (sisa <= 0 && nominal > 0) status = 'PAID';
    else if (nominal > 0 && sisa > 0) status = 'PARTIAL';

    result.push({
      _row: SHEETS.FLIGHT.dataStart + i,
      idUnik: idUnik,
      inv: inv,
      invoice: invNo,
      pax: pax,
      nameBP: String(row[COL_F.nameBP - 1] || '').trim(),
      hariTanggal: parseDateCell(row[COL_F.hariTanggal - 1]),
      hariTanggalRaw: fmtDateCell(row[COL_F.hariTanggal - 1]),
      invoiceDate: parseDateCell(row[COL_F.invoiceDate - 1]),
      invoiceDateRaw: fmtDateCell(row[COL_F.invoiceDate - 1]),
      maskapai: String(row[COL_F.maskapai - 1] || '').trim(),
      flightNo: String(row[COL_F.flightNo - 1] || '').trim(),
      timeDep: String(row[COL_F.timeDep - 1] || '').trim(),
      depAt: String(row[COL_F.depAt - 1] || '').trim(),
      timeArr: String(row[COL_F.timeArr - 1] || '').trim(),
      arrAt: String(row[COL_F.arrAt - 1] || '').trim(),
      kelas: String(row[COL_F.kelas - 1] || '').trim() || 'ECONOMY',
      pnr: String(row[COL_F.pnr - 1] || '').trim(),
      ticketNo: String(row[COL_F.ticketNo - 1] || '').trim(),
      harga: harga,
      addOn: addOn,
      totalAmount: total || (harga + addOn),
      tglBayar: tglBayar,
      nominalBayar: nominal,
      sisa: sisa,
      paymentStatus: status
    });
  }
  return result;
}

// ============ DB_HOTEL ============
function readDBHotel() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SHEETS.HOTEL.name);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.HOTEL.dataStart) return [];

  var numRows = lastRow - SHEETS.HOTEL.dataStart + 1;
  var data = sheet.getRange(SHEETS.HOTEL.dataStart, 1, numRows, SHEETS.HOTEL.lastCol).getValues();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var idUnik = String(row[COL_H.idUnik - 1] || '').trim();
    var inv = String(row[COL_H.inv - 1] || '').trim();
    var invNo = String(row[COL_H.invoiceNo - 1] || '').trim();
    var pax = String(row[COL_H.pax - 1] || '').trim();

    if (!idUnik && !pax && !inv && !invNo) continue;
    if (!idUnik && !pax) continue;

    var harga = parseRupiah(row[COL_H.harga - 1]);
    var addOn = parseRupiah(row[COL_H.addOn - 1]);
    var total = parseRupiah(row[COL_H.totalAmount - 1]);
    var nominal = parseRupiah(row[COL_H.nominalBayar - 1]);
    var sisa = parseRupiah(row[COL_H.sisa - 1]);
    var tglBayar = fmtDateCell(row[COL_H.tglBayar - 1]);

    var status = 'UNPAID';
    if (sisa <= 0 && nominal > 0) status = 'PAID';
    else if (nominal > 0 && sisa > 0) status = 'PARTIAL';

    result.push({
      _row: SHEETS.HOTEL.dataStart + i,
      idUnik: idUnik,
      inv: inv,
      invoice: invNo,
      pax: pax,
      hariTanggal: parseDateCell(row[COL_H.hariTanggal - 1]),
      hariTanggalRaw: fmtDateCell(row[COL_H.hariTanggal - 1]),
      invoiceDate: parseDateCell(row[COL_H.invoiceDate - 1]),
      invoiceDateRaw: fmtDateCell(row[COL_H.invoiceDate - 1]),
      bookingRef: String(row[COL_H.bookingRef - 1] || '').trim(),
      propertyName: String(row[COL_H.propertyName - 1] || '').trim(),
      propertyAddress: String(row[COL_H.propertyAddress - 1] || '').trim(),
      typeRoom: String(row[COL_H.typeRoom - 1] || '').trim(),
      meal: String(row[COL_H.meal - 1] || '').trim(),
      checkIn: parseDateCell(row[COL_H.checkIn - 1]),
      checkInRaw: fmtDateCell(row[COL_H.checkIn - 1]),
      checkOut: parseDateCell(row[COL_H.checkOut - 1]),
      checkOutRaw: fmtDateCell(row[COL_H.checkOut - 1]),
      extraBed: String(row[COL_H.extraBed - 1] || '').trim(),
      roomQty: Number(row[COL_H.roomQty - 1]) || 1,
      harga: harga,
      addOn: addOn,
      totalAmount: total || (harga + addOn),
      tglBayar: tglBayar,
      nominalBayar: nominal,
      sisa: sisa,
      paymentStatus: status
    });
  }
  return result;
}

// ============ DB_USERS ============
function readDBUsers() {
  var ss = getSS();
  var sheet = ss.getSheetByName(SHEETS.USERS.name);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.USERS.dataStart) return [];

  var data = sheet.getRange(SHEETS.USERS.dataStart, 1, lastRow - 1, SHEETS.USERS.lastCol).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var username = String(row[0] || '').trim();
    if (!username) continue;
    result.push({
      _row: SHEETS.USERS.dataStart + i,
      username: username.toLowerCase(),
      password: String(row[1] || ''),
      name: String(row[2] || ''),
      role: String(row[3] || 'viewer').toLowerCase(),
      status: String(row[4] || 'Aktif').trim()
    });
  }
  return result;
}

// ============ DB_CLIENTS ============
function readDBClients() {
  var sheet = getOrCreateSheet(SHEETS.CLIENTS.name, ["ID", "NAMA_CLIENT", "EMAIL", "PHONE", "ALAMAT", "CREATED_AT"]);
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.CLIENTS.dataStart) return [];
  var data = sheet.getRange(SHEETS.CLIENTS.dataStart, 1, lastRow - 1, SHEETS.CLIENTS.lastCol).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var id = String(row[0] || '').trim();
    var nama = String(row[1] || '').trim();
    if (!id && !nama) continue;
    result.push({
      _row: SHEETS.CLIENTS.dataStart + i,
      id: id,
      nama: nama,
      email: String(row[2] || '').trim(),
      phone: String(row[3] || '').trim(),
      alamat: String(row[4] || '').trim(),
      createdAt: fmtDateCell(row[5])
    });
  }
  return result;
}

// ============ DB_CLIENT_INVOICE_MAP ============
function readDBClientMap() {
  var sheet = getOrCreateSheet(SHEETS.MAP.name, ["INVOICE_NO", "CLIENT_ID", "CREATED_AT"]);
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.MAP.dataStart) return [];
  var data = sheet.getRange(SHEETS.MAP.dataStart, 1, lastRow - 1, SHEETS.MAP.lastCol).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var invNo = String(row[0] || '').trim();
    var cliId = String(row[1] || '').trim();
    if (!invNo) continue;
    result.push({
      _row: SHEETS.MAP.dataStart + i,
      invoiceNo: invNo,
      clientId: cliId,
      createdAt: fmtDateCell(row[2])
    });
  }
  return result;
}

// ============ DB_LOG_PRINT ============
function readDBLogPrint() {
  var sheet = getOrCreateSheet(SHEETS.LOG.name, ["TIMESTAMP", "USERNAME", "ID_UNIK", "JENIS_DOKUMEN", "REPRINT_KE"]);
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.LOG.dataStart) return [];
  var data = sheet.getRange(SHEETS.LOG.dataStart, 1, lastRow - 1, SHEETS.LOG.lastCol).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    result.push({
      _row: SHEETS.LOG.dataStart + i,
      timestamp: fmtDateCell(row[0]),
      username: String(row[1] || '').trim(),
      idUnik: String(row[2] || '').trim(),
      jenisDokumen: String(row[3] || '').trim(),
      reprintKe: Number(row[4]) || 0
    });
  }
  return result;
}

// ============ SHEET AUTO-CREATE ============
function getOrCreateSheet(name, headers) {
  var ss = getSS();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#28583e")
      .setFontColor("#fff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============ FIND FIRST EMPTY ROW ============
function findFirstEmptyRow(sheet, dataStart, idCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStart) return dataStart;
  var numRows = lastRow - dataStart + 1;
  var values = sheet.getRange(dataStart, idCol, numRows, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || '').trim();
    if (v === '') return dataStart + i;
  }
  return lastRow + 1;
}

// ============ CREATE FLIGHT ============
function createFlightDB(d) {
  var sheet = getSS().getSheetByName(SHEETS.FLIGHT.name);
  if (!sheet) throw new Error("Sheet DB_FLIGHT tidak ditemukan");

  var inv = String(d.inv || '').trim();
  var invoiceDate = String(d.invoiceDateRaw || d.invoiceDate || '').trim();
  var invoiceNo = generateInvoiceNo(inv, invoiceDate);
  var idUnik = generateIDUnik(inv);

  var harga = Number(d.harga) || 0;
  var addOn = Number(d.addOn) || 0;
  var total = harga + addOn;

  var targetRow = findFirstEmptyRow(sheet, SHEETS.FLIGHT.dataStart, COL_F.idUnik);

  // Set format kolom INV & INVOICE_NO jadi text dulu (biar 002 tidak jadi 2)
  sheet.getRange(targetRow, COL_F.inv).setNumberFormat('@');
  sheet.getRange(targetRow, COL_F.invoiceNo).setNumberFormat('@');

  var row = new Array(SHEETS.FLIGHT.lastCol).fill('');
  row[COL_F.idUnik - 1]      = idUnik;
  row[COL_F.inv - 1]         = inv;
  row[COL_F.pax - 1]         = String(d.pax || '').toUpperCase();
  row[COL_F.jabatan - 1]     = '';
  row[COL_F.hariTanggal - 1] = String(d.hariTanggalRaw || d.hariTanggal || '');
  row[COL_F.invoiceDate - 1] = invoiceDate;
  row[COL_F.nameBP - 1]      = String(d.nameBP || '').toUpperCase();
  row[COL_F.maskapai - 1]    = String(d.maskapai || '').toUpperCase();
  row[COL_F.flightNo - 1]    = String(d.flightNo || '');
  row[COL_F.timeDep - 1]     = String(d.timeDep || '');
  row[COL_F.depAt - 1]       = String(d.depAt || '');
  row[COL_F.timeArr - 1]     = String(d.timeArr || '');
  row[COL_F.arrAt - 1]       = String(d.arrAt || '');
  row[COL_F.kelas - 1]       = String(d.kelas || 'ECONOMY').toUpperCase();
  row[COL_F.pnr - 1]         = String(d.pnr || '').toUpperCase();
  row[COL_F.ticketNo - 1]    = String(d.ticketNo || '');
  row[COL_F.invoiceNo - 1]   = invoiceNo;
  row[COL_F.harga - 1]       = formatRupiahCell(harga);
  row[COL_F.addOn - 1]       = formatRupiahCell(addOn);
  row[COL_F.totalAmount - 1] = formatRupiahCell(total);
  row[COL_F.sisa - 1]        = formatRupiahCell(total);

  sheet.getRange(targetRow, 1, 1, SHEETS.FLIGHT.lastCol).setValues([row]);

  return jsonOut({
    status: "success",
    message: "Flight tersimpan",
    idUnik: idUnik,
    invoiceNo: invoiceNo,
    row: targetRow
  });
}

// ============ CREATE HOTEL ============
function createHotelDB(d) {
  var sheet = getSS().getSheetByName(SHEETS.HOTEL.name);
  if (!sheet) throw new Error("Sheet DB_HOTEL tidak ditemukan");

  var inv = String(d.inv || '').trim();
  var invoiceDate = String(d.invoiceDateRaw || d.invoiceDate || '').trim();
  var invoiceNo = generateInvoiceNo(inv, invoiceDate);
  var idUnik = generateIDUnik(inv);

  var harga = Number(d.harga) || 0;
  var addOn = Number(d.addOn) || 0;
  var roomQty = Number(d.roomQty) || 1;
  var total = (harga * roomQty) + addOn;

  var targetRow = findFirstEmptyRow(sheet, SHEETS.HOTEL.dataStart, COL_H.idUnik);
  sheet.getRange(targetRow, COL_H.inv).setNumberFormat('@');
  sheet.getRange(targetRow, COL_H.invoiceNo).setNumberFormat('@');

  var row = new Array(SHEETS.HOTEL.lastCol).fill('');
  row[COL_H.idUnik - 1]         = idUnik;
  row[COL_H.inv - 1]            = inv;
  row[COL_H.pax - 1]            = String(d.pax || '').toUpperCase();
  row[COL_H.hariTanggal - 1]    = String(d.hariTanggalRaw || d.hariTanggal || '');
  row[COL_H.invoiceDate - 1]    = invoiceDate;
  row[COL_H.invoiceNo - 1]      = invoiceNo;
  row[COL_H.bookingRef - 1]     = String(d.bookingRef || '-');
  row[COL_H.propertyName - 1]   = String(d.propertyName || '').toUpperCase();
  row[COL_H.propertyAddress - 1]= String(d.propertyAddress || '');
  row[COL_H.typeRoom - 1]       = String(d.typeRoom || '').toUpperCase();
  row[COL_H.meal - 1]           = String(d.meal || 'Breakfast');
  row[COL_H.checkIn - 1]        = String(d.checkInRaw || d.checkIn || '');
  row[COL_H.checkOut - 1]       = String(d.checkOutRaw || d.checkOut || '');
  row[COL_H.extraBed - 1]       = String(d.extraBed || '-');
  row[COL_H.roomQty - 1]        = roomQty;
  row[COL_H.harga - 1]          = formatRupiahCell(harga);
  row[COL_H.addOn - 1]          = formatRupiahCell(addOn);
  row[COL_H.totalAmount - 1]    = formatRupiahCell(total);
  row[COL_H.sisa - 1]           = formatRupiahCell(total);

  sheet.getRange(targetRow, 1, 1, SHEETS.HOTEL.lastCol).setValues([row]);

  return jsonOut({
    status: "success",
    message: "Hotel tersimpan",
    idUnik: idUnik,
    invoiceNo: invoiceNo,
    row: targetRow
  });
}

// ============ UPDATE FLIGHT ============
function updateFlightDB(d) {
  var sheet = getSS().getSheetByName(SHEETS.FLIGHT.name);
  var r = findRowByIdUnik(sheet, d._originalIdUnik || d.idUnik, SHEETS.FLIGHT.dataStart, COL_F.idUnik, SHEETS.FLIGHT.lastCol);
  if (r < 0) return jsonOut({status:"error", message:"Data tidak ditemukan"});

  var inv = String(d.inv || '').trim();
  var invoiceDate = String(d.invoiceDateRaw || d.invoiceDate || '').trim();
  var invoiceNo = generateInvoiceNo(inv, invoiceDate);
  var harga = Number(d.harga) || 0;
  var addOn = Number(d.addOn) || 0;
  var total = harga + addOn;

  sheet.getRange(r, COL_F.inv).setNumberFormat('@').setValue(inv);
  sheet.getRange(r, COL_F.pax).setValue(String(d.pax || '').toUpperCase());
  sheet.getRange(r, COL_F.hariTanggal).setValue(String(d.hariTanggalRaw || d.hariTanggal || ''));
  sheet.getRange(r, COL_F.invoiceDate).setValue(invoiceDate);
  sheet.getRange(r, COL_F.nameBP).setValue(String(d.nameBP || '').toUpperCase());
  sheet.getRange(r, COL_F.maskapai).setValue(String(d.maskapai || '').toUpperCase());
  sheet.getRange(r, COL_F.flightNo).setValue(String(d.flightNo || ''));
  sheet.getRange(r, COL_F.timeDep).setValue(String(d.timeDep || ''));
  sheet.getRange(r, COL_F.depAt).setValue(String(d.depAt || ''));
  sheet.getRange(r, COL_F.timeArr).setValue(String(d.timeArr || ''));
  sheet.getRange(r, COL_F.arrAt).setValue(String(d.arrAt || ''));
  sheet.getRange(r, COL_F.kelas).setValue(String(d.kelas || 'ECONOMY').toUpperCase());
  sheet.getRange(r, COL_F.pnr).setValue(String(d.pnr || '').toUpperCase());
  sheet.getRange(r, COL_F.ticketNo).setValue(String(d.ticketNo || ''));
  sheet.getRange(r, COL_F.invoiceNo).setNumberFormat('@').setValue(invoiceNo);
  sheet.getRange(r, COL_F.harga).setValue(formatRupiahCell(harga));
  sheet.getRange(r, COL_F.addOn).setValue(formatRupiahCell(addOn));
  sheet.getRange(r, COL_F.totalAmount).setValue(formatRupiahCell(total));

  var nominal = parseRupiah(sheet.getRange(r, COL_F.nominalBayar).getValue());
  var newSisa = total - nominal;
  if (newSisa < 0) newSisa = 0;
  sheet.getRange(r, COL_F.sisa).setValue(formatRupiahCell(newSisa));

  return jsonOut({status:"success", message:"Flight di-update", invoiceNo: invoiceNo});
}

// ============ UPDATE HOTEL ============
function updateHotelDB(d) {
  var sheet = getSS().getSheetByName(SHEETS.HOTEL.name);
  var r = findRowByIdUnik(sheet, d._originalIdUnik || d.idUnik, SHEETS.HOTEL.dataStart, COL_H.idUnik, SHEETS.HOTEL.lastCol);
  if (r < 0) return jsonOut({status:"error", message:"Data tidak ditemukan"});

  var inv = String(d.inv || '').trim();
  var invoiceDate = String(d.invoiceDateRaw || d.invoiceDate || '').trim();
  var invoiceNo = generateInvoiceNo(inv, invoiceDate);
  var harga = Number(d.harga) || 0;
  var addOn = Number(d.addOn) || 0;
  var roomQty = Number(d.roomQty) || 1;
  var total = (harga * roomQty) + addOn;

  sheet.getRange(r, COL_H.inv).setNumberFormat('@').setValue(inv);
  sheet.getRange(r, COL_H.pax).setValue(String(d.pax || '').toUpperCase());
  sheet.getRange(r, COL_H.hariTanggal).setValue(String(d.hariTanggalRaw || d.hariTanggal || ''));
  sheet.getRange(r, COL_H.invoiceDate).setValue(invoiceDate);
  sheet.getRange(r, COL_H.invoiceNo).setNumberFormat('@').setValue(invoiceNo);
  sheet.getRange(r, COL_H.bookingRef).setValue(String(d.bookingRef || '-'));
  sheet.getRange(r, COL_H.propertyName).setValue(String(d.propertyName || '').toUpperCase());
  sheet.getRange(r, COL_H.propertyAddress).setValue(String(d.propertyAddress || ''));
  sheet.getRange(r, COL_H.typeRoom).setValue(String(d.typeRoom || '').toUpperCase());
  sheet.getRange(r, COL_H.meal).setValue(String(d.meal || 'Breakfast'));
  sheet.getRange(r, COL_H.checkIn).setValue(String(d.checkInRaw || d.checkIn || ''));
  sheet.getRange(r, COL_H.checkOut).setValue(String(d.checkOutRaw || d.checkOut || ''));
  sheet.getRange(r, COL_H.extraBed).setValue(String(d.extraBed || '-'));
  sheet.getRange(r, COL_H.roomQty).setValue(roomQty);
  sheet.getRange(r, COL_H.harga).setValue(formatRupiahCell(harga));
  sheet.getRange(r, COL_H.addOn).setValue(formatRupiahCell(addOn));
  sheet.getRange(r, COL_H.totalAmount).setValue(formatRupiahCell(total));

  var nominal = parseRupiah(sheet.getRange(r, COL_H.nominalBayar).getValue());
  var newSisa = total - nominal;
  if (newSisa < 0) newSisa = 0;
  sheet.getRange(r, COL_H.sisa).setValue(formatRupiahCell(newSisa));

  return jsonOut({status:"success", message:"Hotel di-update", invoiceNo: invoiceNo});
}

// ============ DELETE ============
function deleteFlightDB(idUnik) {
  var sheet = getSS().getSheetByName(SHEETS.FLIGHT.name);
  var r = findRowByIdUnik(sheet, idUnik, SHEETS.FLIGHT.dataStart, COL_F.idUnik, SHEETS.FLIGHT.lastCol);
  if (r > 0) { sheet.deleteRow(r); return jsonOut({status:"success", message:"Flight dihapus"}); }
  return jsonOut({status:"error", message:"Data tidak ditemukan"});
}

function deleteHotelDB(idUnik) {
  var sheet = getSS().getSheetByName(SHEETS.HOTEL.name);
  var r = findRowByIdUnik(sheet, idUnik, SHEETS.HOTEL.dataStart, COL_H.idUnik, SHEETS.HOTEL.lastCol);
  if (r > 0) { sheet.deleteRow(r); return jsonOut({status:"success", message:"Hotel dihapus"}); }
  return jsonOut({status:"error", message:"Data tidak ditemukan"});
}

// ============ MARK PAYMENT ============
function markPaymentFlightDB(data) {
  var sheet = getSS().getSheetByName(SHEETS.FLIGHT.name);
  var idList = data.idUnikList || [data.idUnik];
  if (!idList.length) return jsonOut({status:"error", message:"Tidak ada data dipilih"});

  var isPaid = (data.status !== "UNPAID");
  var now = new Date();
  var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var count = 0;

  for (var i = 0; i < idList.length; i++) {
    var r = findRowByIdUnik(sheet, idList[i], SHEETS.FLIGHT.dataStart, COL_F.idUnik, SHEETS.FLIGHT.lastCol);
    if (r < 0) continue;
    var total = parseRupiah(sheet.getRange(r, COL_F.totalAmount).getValue());
    if (isPaid) {
      sheet.getRange(r, COL_F.tglBayar).setNumberFormat('@').setValue(nowStr);
      sheet.getRange(r, COL_F.nominalBayar).setValue(formatRupiahCell(total));
      sheet.getRange(r, COL_F.sisa).setValue(formatRupiahCell(0));
    } else {
      sheet.getRange(r, COL_F.tglBayar).setValue('');
      sheet.getRange(r, COL_F.nominalBayar).setValue('');
      sheet.getRange(r, COL_F.sisa).setValue(formatRupiahCell(total));
    }
    count++;
  }
  return jsonOut({status:"success", message: count + " flight diupdate", count: count});
}

function markPaymentHotelDB(data) {
  var sheet = getSS().getSheetByName(SHEETS.HOTEL.name);
  var idList = data.idUnikList || [data.idUnik];
  if (!idList.length) return jsonOut({status:"error", message:"Tidak ada data dipilih"});

  var isPaid = (data.status !== "UNPAID");
  var now = new Date();
  var nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var count = 0;

  for (var i = 0; i < idList.length; i++) {
    var r = findRowByIdUnik(sheet, idList[i], SHEETS.HOTEL.dataStart, COL_H.idUnik, SHEETS.HOTEL.lastCol);
    if (r < 0) continue;
    var total = parseRupiah(sheet.getRange(r, COL_H.totalAmount).getValue());
    if (isPaid) {
      sheet.getRange(r, COL_H.tglBayar).setNumberFormat('@').setValue(nowStr);
      sheet.getRange(r, COL_H.nominalBayar).setValue(formatRupiahCell(total));
      sheet.getRange(r, COL_H.sisa).setValue(formatRupiahCell(0));
    } else {
      sheet.getRange(r, COL_H.tglBayar).setValue('');
      sheet.getRange(r, COL_H.nominalBayar).setValue('');
      sheet.getRange(r, COL_H.sisa).setValue(formatRupiahCell(total));
    }
    count++;
  }
  return jsonOut({status:"success", message: count + " hotel diupdate", count: count});
}

// ============ USER MANAGEMENT ============
function handleUserAction(data) {
  var sheet = getOrCreateSheet(SHEETS.USERS.name, ["USERNAME", "PASSWORD", "NAMA_LENGKAP", "ROLE", "STATUS"]);
  var action = data.action;

  if (action === "create") {
    if (findUserRow(sheet, data.username) > 0) return jsonOut({status:"error", message:"Username sudah dipakai"});
    sheet.appendRow([
      String(data.username || '').toLowerCase(),
      String(data.password || ''),
      String(data.name || ''),
      String(data.role || 'viewer'),
      String(data.status || 'Aktif')
    ]);
  } else if (action === "update") {
    var r = findUserRow(sheet, data._originalUsername || data.username);
    if (r < 1) return jsonOut({status:"error", message:"User tidak ditemukan"});
    sheet.getRange(r, 1, 1, 5).setValues([[
      String(data.username || '').toLowerCase(),
      String(data.password || ''),
      String(data.name || ''),
      String(data.role || 'viewer'),
      String(data.status || 'Aktif')
    ]]);
  } else if (action === "delete") {
    var r2 = findUserRow(sheet, data.username);
    if (r2 > 0) sheet.deleteRow(r2);
  } else if (action === "changePassword") {
    var r3 = findUserRow(sheet, data.username);
    if (r3 < 1) return jsonOut({status:"error", message:"User tidak ditemukan"});
    if (String(sheet.getRange(r3, 2).getValue()) !== String(data.oldPassword||"")) {
      return jsonOut({status:"error", message:"Password lama salah"});
    }
    sheet.getRange(r3, 2).setValue(data.newPassword||"");
  }
  return jsonOut({status:"success", message:"User updated"});
}

// ============ CLIENT MANAGEMENT ============
function handleClientAction(data) {
  var sheet = getOrCreateSheet(SHEETS.CLIENTS.name, ["ID", "NAMA_CLIENT", "EMAIL", "PHONE", "ALAMAT", "CREATED_AT"]);
  var action = data.action;

  if (action === "create") {
    var newId = getNextClientId(sheet);
    sheet.appendRow([
      newId,
      String(data.nama || '').toUpperCase(),
      String(data.email || ''),
      String(data.phone || ''),
      String(data.alamat || ''),
      fmtDateTime(new Date())
    ]);
    return jsonOut({status:"success", message:"Client ditambahkan", id: newId});
  } else if (action === "update") {
    var r = findRowByCol(sheet, data.id, 1);
    if (r < 1) return jsonOut({status:"error", message:"Client tidak ditemukan"});
    sheet.getRange(r, 1, 1, 6).setValues([[
      data.id,
      String(data.nama || '').toUpperCase(),
      String(data.email || ''),
      String(data.phone || ''),
      String(data.alamat || ''),
      sheet.getRange(r, 6).getValue() || fmtDateTime(new Date())
    ]]);
  } else if (action === "delete") {
    var r2 = findRowByCol(sheet, data.id, 1);
    if (r2 > 0) sheet.deleteRow(r2);
  }
  return jsonOut({status:"success", message:"Client updated"});
}

function getNextClientId(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < SHEETS.CLIENTS.dataStart) return "C001";
  var values = sheet.getRange(SHEETS.CLIENTS.dataStart, 1, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] || '').trim();
    var m = v.match(/^C?(\d+)$/i);
    if (m) {
      var n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  }
  return "C" + String(max + 1).padStart(3, '0');
}

// ============ CLIENT INVOICE MAP ============
function handleClientMapAction(data) {
  var sheet = getOrCreateSheet(SHEETS.MAP.name, ["INVOICE_NO", "CLIENT_ID", "CREATED_AT"]);
  var action = data.action;

  if (action === "assign") {
    var invNo = String(data.invoiceNo || '').trim();
    if (!invNo) return jsonOut({status:"error", message:"Invoice No kosong"});

    var r = findRowByCol(sheet, invNo, 1);
    if (r > 0) {
      sheet.getRange(r, 2).setValue(String(data.clientId || ''));
      return jsonOut({status:"success", message:"Client updated untuk invoice ini"});
    }
    sheet.appendRow([invNo, String(data.clientId || ''), fmtDateTime(new Date())]);
    return jsonOut({status:"success", message:"Client di-assign"});
  } else if (action === "unassign") {
    var r2 = findRowByCol(sheet, data.invoiceNo, 1);
    if (r2 > 0) sheet.deleteRow(r2);
    return jsonOut({status:"success", message:"Client dihapus dari invoice"});
  }
  return jsonOut({status:"error", message:"Action tidak dikenal"});
}

// ============ LOG PRINT ============
function handleLogPrint(data) {
  var sheet = getOrCreateSheet(SHEETS.LOG.name, ["TIMESTAMP", "USERNAME", "ID_UNIK", "JENIS_DOKUMEN", "REPRINT_KE"]);
  var idUnik = String(data.idUnik || '').trim();
  var jenis = String(data.jenisDokumen || '').trim();
  var username = String(data.username || '').trim();

  if (!idUnik || !jenis) return jsonOut({status:"error", message:"Data log tidak lengkap"});

  var lastRow = sheet.getLastRow();
  var reprintKe = 1;
  if (lastRow >= SHEETS.LOG.dataStart) {
    var values = sheet.getRange(SHEETS.LOG.dataStart, 1, lastRow - 1, 5).getValues();
    var count = 0;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][2]).trim() === idUnik && String(values[i][3]).trim() === jenis) {
        count++;
      }
    }
    reprintKe = count + 1;
  }

  sheet.appendRow([fmtDateTime(new Date()), username, idUnik, jenis, reprintKe]);
  return jsonOut({status:"success", message:"Log tercatat", reprintKe: reprintKe});
}

// ============ UTILITIES ============
function generateInvoiceNo(inv, invoiceDate) {
  var invStr = String(inv || '').trim();
  var dateStr = String(invoiceDate || '').trim();

  // Kalau invoiceDate berupa Date object (dari input user yang auto-convert)
  // Kita sudah handle di form (formatDateDash di FE), jadi ini untuk safety
  if (dateStr.match(/GMT|Waktu/)) {
    // Fallback: coba parse via Date
    var d = new Date(dateStr);
    if (!isNaN(d.getTime())) {
      var dd = String(d.getDate()).padStart(2,'0');
      var mm = String(d.getMonth()+1).padStart(2,'0');
      var yyyy = d.getFullYear();
      return invStr + dd + mm + yyyy;
    }
  }

  var m = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    var dd2 = String(m[1]).padStart(2,'0');
    var mm2 = String(m[2]).padStart(2,'0');
    var yyyy2 = m[3];
    return invStr + dd2 + mm2 + yyyy2;
  }
  var m2 = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) {
    return invStr + m2[3] + m2[2] + m2[1];
  }
  return invStr;
}

function generateIDUnik(inv) {
  var invStr = String(inv || '').trim();
  if (!invStr) return '';
  var ss = getSS();
  var countF = countInvInSheet(ss.getSheetByName(SHEETS.FLIGHT.name), COL_F.inv, invStr);
  var countH = countInvInSheet(ss.getSheetByName(SHEETS.HOTEL.name),  COL_H.inv, invStr);
  var total = countF + countH;
  return (total + 1) + '-' + invStr;
}

function countInvInSheet(sheet, colInv, invValue) {
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  var dataStart = (sheet.getName() === SHEETS.FLIGHT.name) ? SHEETS.FLIGHT.dataStart : SHEETS.HOTEL.dataStart;
  if (lastRow < dataStart) return 0;
  var values = sheet.getRange(dataStart, colInv, lastRow - dataStart + 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === invValue) count++;
  }
  return count;
}

function findRowByIdUnik(sheet, idUnik, dataStart, colIdUnik, lastCol) {
  if (!sheet) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStart) return -1;
  var values = sheet.getRange(dataStart, colIdUnik, lastRow - dataStart + 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(idUnik).trim()) {
      return dataStart + i;
    }
  }
  return -1;
}

function findRowByCol(sheet, value, col) {
  if (!sheet) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  var values = sheet.getRange(1, col, lastRow, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(value).trim()) {
      return i + 1;
    }
  }
  return -1;
}

function findUserRow(sheet, username) {
  var v = sheet.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) return i + 1;
  }
  return -1;
}

function parseRupiah(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^\d]/g, '');
  return parseInt(s || '0', 10);
}

function formatRupiahCell(n) {
  var num = Math.floor(Number(n) || 0);
  return 'Rp' + num.toLocaleString('id-ID');
}

// Parse date cell → return "YYYY-MM-DD"
function parseDateCell(v) {
  if (!v && v !== 0) return '';
  // Date object
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  // String
  var s = String(v).trim();
  if (!s) return '';

  // String format "27/09/2026"
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    return m[3] + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[1]).padStart(2,'0');
  }

  // String format "2026-09-27"
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2) return s;

  // String format "Wednesday, August 05, 2026" atau "Friday, 07 August 2026"
  var m3 = parseHariTanggal(s);
  if (m3) return m3;

  // Fallback: Date.parse
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

// Format date cell untuk display/return ke FE → return "YYYY-MM-DD" atau original string
function fmtDateCell(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).trim();
}

// Alias untuk backward compat
function parseSlashDate(v) { return parseDateCell(v); }

var MONTH_MAP = {
  january:1, february:2, march:3, april:4, may:5, june:6,
  july:7, august:8, september:9, october:10, november:11, december:12,
  januari:1, februari:2, maret:3, mei:5, juni:6, juli:7, agustus:8,
  oktober:10, november:11, desember:12
};

function parseHariTanggal(v) {
  if (!v) return '';
  var s = String(v).trim();

  // Pattern A: "Wednesday, August 05, 2026"
  var mA = s.match(/^\w+,?\s+(\w+)\s+(\d{1,2}),?\s+(\d{4})$/);
  if (mA) {
    var m1 = MONTH_MAP[mA[1].toLowerCase()];
    if (m1) return mA[3] + '-' + String(m1).padStart(2,'0') + '-' + String(mA[2]).padStart(2,'0');
  }

  // Pattern B: "Friday, 07 August 2026"
  var mB = s.match(/^\w+,?\s+(\d{1,2})\s+(\w+)\s+(\d{4})$/);
  if (mB) {
    var m2 = MONTH_MAP[mB[2].toLowerCase()];
    if (m2) return mB[3] + '-' + String(m2).padStart(2,'0') + '-' + String(mB[1]).padStart(2,'0');
  }

  // Pattern C: "Wed Aug 05 2026 14:00:00 GMT+0700"
  var mC = s.match(/^\w+\s+(\w+)\s+(\d{1,2})\s+(\d{4})/);
  if (mC) {
    var m3 = MONTH_MAP[mC[1].toLowerCase()];
    if (m3) return mC[3] + '-' + String(m3).padStart(2,'0') + '-' + String(mC[2]).padStart(2,'0');
  }

  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

function fmtDateTime(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}