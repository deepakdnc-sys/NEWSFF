function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

function getAppState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) {
    initializeSheets();
    return getAppState();
  }

  const settings = readSheetMap('Config');
  const menu = readSheetMap('Menu');
  const offers = readSheetMap('Offers');
  const bookings = readSheetMap('Bookings');
  const vouchers = readSheetMap('Vouchers');

  return {
    restaurant: {
      name: 'SONI FAST FOOD',
      location: 'Moti Chowk, Taxi Stand, Jodhpur, Rajasthan',
      phone: '+91 8385963066',
      whatsapp: 'https://wa.me/918385963066',
      maps: 'https://maps.google.com/?q=Moti+Chowk+Taxi+Stand+Jodhpur+Rajasthan',
      upiId: settings.upiId || '8385963066@upi',
      upiName: settings.upiName || 'SONI FAST FOOD',
      announcement: settings.announcement || 'Freshly prepared meals for tomorrow dine-in only.',
      bookingOpen: String(settings.bookingOpen || 'true').toLowerCase() === 'true',
      defaultPinHash: settings.pinHash || '',
      timeSlots: JSON.parse(settings.timeSlots || '[]')
    },
    menu: menu,
    offers: offers,
    bookings: bookings,
    vouchers: vouchers,
    stats: computeDashboardStats(),
    now: new Date().toISOString()
  };
}

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, 'Config', [
    ['setting', 'value'],
    ['announcement', 'Freshly prepared meals for tomorrow dine-in only.'],
    ['bookingOpen', 'true'],
    ['upiId', '8385963066@upi'],
    ['upiName', 'SONI FAST FOOD'],
    ['timeSlots', JSON.stringify(['7:00 PM', '8:00 PM', '9:00 PM'])],
    ['pinHash', ''],
    ['pinChanged', 'false'],
    ['mode', 'customer']
  ]);

  ensureSheet(ss, 'Menu', [
    ['id', 'name', 'category', 'price', 'offerPrice', 'stock', 'totalStock', 'reservedStock', 'image', 'featured', 'active', 'description'],
    ['pizza_1', 'Classic Pizza', 'Pizza', 250, 220, 10, 10, 0, 'https://images.unsplash.com/photo-1513104890138-7c749659a591?auto=format&fit=crop&w=900&q=80', 'true', 'true', 'Loaded with cheese and herbs.'],
    ['fries_1', 'Crispy Fries', 'Fries', 120, 99, 15, 15, 0, 'https://images.unsplash.com/photo-1576107232684-1279f390859f?auto=format&fit=crop&w=900&q=80', 'true', 'true', 'Golden crisp fries.'],
    ['momo_1', 'Steamed Momos', 'Momo', 160, 140, 12, 12, 0, 'https://images.unsplash.com/photo-1604908176997-125f25cc6f3d?auto=format&fit=crop&w=900&q=80', 'true', 'true', 'Soft dumplings with spicy sauce.'],
    ['sandwich_1', 'Club Sandwich', 'Sandwich', 180, 160, 9, 9, 0, 'https://images.unsplash.com/photo-1528735602780-2552fd46c7af?auto=format&fit=crop&w=900&q=80', 'true', 'true', 'A hearty loaded sandwich.'],
    ['burger_1', 'Veg Burger', 'Burger', 170, 150, 8, 8, 0, 'https://images.unsplash.com/photo-1568901346375-23c9450c58cd?auto=format&fit=crop&w=900&q=80', 'true', 'true', 'Classic grilled veg burger.'],
    ['others_1', 'Combo Meal', 'Others', 260, 230, 6, 6, 0, 'https://images.unsplash.com/photo-1559847844-5315695dadae?auto=format&fit=crop&w=900&q=80', 'false', 'true', 'A complete meal combo.']
  ]);

  ensureSheet(ss, 'Offers', [
    ['id', 'title', 'subtitle', 'discount', 'active'],
    ['offer_1', 'Weekend Feast', 'Enjoy 10% off on selected combos.', '10%', 'true'],
    ['offer_2', 'Midnight Cravings', 'Free fries with burger orders today.', 'Free Fries', 'true']
  ]);

  ensureSheet(ss, 'Bookings', [
    ['id', 'bookingDate', 'slot', 'customerName', 'mobile', 'itemId', 'itemName', 'quantity', 'amount', 'upiId', 'utr', 'paymentProofUrl', 'paymentProofFileId', 'status', 'rejectionReason', 'createdAt', 'updatedAt', 'voucherTokenHash', 'voucherStatus', 'redeemedAt', 'adminVerifiedAt'],
  ]);

  ensureSheet(ss, 'Vouchers', [
    ['id', 'bookingId', 'tokenHash', 'status', 'expiresAt', 'createdAt', 'redeemedAt', 'revokedAt', 'reason'],
  ]);

  ensureSheet(ss, 'AdminAudit', [
    ['id', 'action', 'by', 'at', 'details'],
  ]);
}

function ensureSheet(ss, name, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return sheet;
}

function readSheetMap(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1).filter(row => row.some(cell => cell !== ''));
  return rows.map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  });
}

function getTomorrowString() {
  const now = new Date();
  const indiaTime = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Kolkata' }));
  const tomorrow = new Date(indiaTime);
  tomorrow.setDate(indiaTime.getDate() + 1);
  return Utilities.formatDate(tomorrow, 'Asia/Kolkata', 'yyyy-MM-dd');
}

function normalizePhone(phone) {
  return String(phone || '').replace(/\D/g, '');
}

function sha256Hex(value) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
  let txt = '';
  raw.forEach(byte => {
    txt += ('0' + (byte & 0xFF).toString(16)).slice(-2);
  });
  return txt;
}

function generateBookingId() {
  return 'SONI-' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
}

function getMenuByCategory() {
  const menu = readSheetMap('Menu').filter(item => String(item.active || '').toLowerCase() === 'true');
  const byCategory = {};
  menu.forEach(item => {
    const category = item.category || 'Others';
    if (!byCategory[category]) byCategory[category] = [];
    byCategory[category].push(item);
  });
  return byCategory;
}

function getAvailableStock(item) {
  const total = Number(item.totalStock || item.stock || 0);
  const reserved = Number(item.reservedStock || 0);
  const sold = Number(item.soldStock || 0);
  return Math.max(total - reserved - sold, 0);
}

function getActiveAdminSession() {
  const session = Session.getActiveUser();
  if (!session) return null;
  const payload = PropertiesService.getUserProperties().getProperty('soni_admin_session');
  if (!payload) return null;
  const parsed = JSON.parse(payload);
  if (!parsed || !parsed.expiresAt || Number(parsed.expiresAt) < Date.now()) {
    PropertiesService.getUserProperties().deleteProperty('soni_admin_session');
    return null;
  }
  return parsed;
}

function adminLogin(pin) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const map = {};
  data[0].forEach((header, idx) => map[header] = data[1] ? data[1][idx] : '');
  const hash = sha256Hex(String(pin || ''));
  const correctHash = map.pinHash || '';
  if (!correctHash || hash !== correctHash) {
    return { success: false, message: 'Invalid admin PIN.' };
  }
  const expiresAt = Date.now() + (30 * 60 * 1000);
  PropertiesService.getUserProperties().setProperty('soni_admin_session', JSON.stringify({ expiresAt, user: 'admin' }));
  return { success: true, expiresAt };
}

function adminSetPin(oldPin, newPin) {
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const map = {};
  data.forEach(row => {
    if (row[0]) map[row[0]] = row[1];
  });
  const currentHash = map.pinHash || '';
  if (currentHash && sha256Hex(String(oldPin || '')) !== currentHash) {
    return { success: false, message: 'Current PIN is incorrect.' };
  }
  const nextHash = sha256Hex(String(newPin || ''));
  updateConfigValue('pinHash', nextHash);
  updateConfigValue('pinChanged', 'true');
  return { success: true, message: 'PIN updated securely.' };
}

function updateConfigValue(key, value) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(key)) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

function reserveStockForBooking(itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Menu');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rowIndex = values.findIndex(row => String(row[0]) === String(itemId));
  if (rowIndex < 0) return false;
  const row = values[rowIndex];
  const reserved = Number(row[headers.indexOf('reservedStock')] || 0);
  const total = Number(row[headers.indexOf('totalStock')] || 0);
  if (total - reserved <= 0) return false;
  sheet.getRange(rowIndex + 1, headers.indexOf('reservedStock') + 1).setValue(reserved + 1);
  return true;
}

function releaseReservedStock(itemId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Menu');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rowIndex = values.findIndex(row => String(row[0]) === String(itemId));
  if (rowIndex < 0) return false;
  const row = values[rowIndex];
  const reserved = Number(row[headers.indexOf('reservedStock')] || 0);
  if (reserved <= 0) return false;
  sheet.getRange(rowIndex + 1, headers.indexOf('reservedStock') + 1).setValue(Math.max(0, reserved - 1));
  return true;
}

function createBooking(payload) {
  const itemId = payload.itemId;
  const mobile = normalizePhone(payload.mobile);
  const sku = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Menu').getDataRange().getValues().find(row => String(row[0]) === String(itemId));
  if (!sku) return { success: false, message: 'Selected item not found.' };

  const stockAvailable = Number(sku[5] || 0) - Number(sku[7] || 0);
  if (stockAvailable <= 0) return { success: false, message: 'This item is out of stock.' };

  const bookings = readSheetMap('Bookings');
  const active = bookings.filter(b => {
    const samePhone = normalizePhone(b.mobile || '') === mobile;
    const open = ['PENDING_VERIFICATION', 'APPROVED', 'PENDING'].includes(String(b.status || ''));
    return samePhone && open;
  });
  if (active.length > 0) return { success: false, message: 'You already have one active booking.' };

  const tomorrow = getTomorrowString();
  const slot = payload.slot || '';
  if (!slot) return { success: false, message: 'Please select a valid time slot.' };

  const bookingId = generateBookingId();
  const itemName = sku[1];
  const amount = Number(sku[3] || 0);
  const bookingRow = [
    bookingId,
    tomorrow,
    slot,
    payload.customerName,
    mobile,
    itemId,
    itemName,
    1,
    amount,
    payload.upiId || '8385963066@upi',
    payload.utr || '',
    payload.paymentProofUrl || '',
    payload.paymentProofFileId || '',
    'PENDING_VERIFICATION',
    '',
    new Date().toISOString(),
    new Date().toISOString(),
    '',
    'NOT_ISSUED',
    '',
    ''
  ];

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  sheet.appendRow(bookingRow);

  if (!reserveStockForBooking(itemId)) {
    const lastRow = sheet.getLastRow();
    const target = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
    target.setBackground('#FDE2E2');
    return { success: false, message: 'Stock reservation failed. Please try again.' };
  }

  return { success: true, bookingId, message: 'Booking saved successfully.' };
}

function submitPayment(payload) {
  const bookingId = payload.bookingId;
  const utr = String(payload.utr || '').trim();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const rowIndex = rows.findIndex(row => String(row[0]) === String(bookingId));
  if (rowIndex <= 0) return { success: false, message: 'Booking not found.' };
  if (!/^[0-9]{12}$/.test(utr)) return { success: false, message: 'UTR must be a 12-digit number.' };
  const row = rows[rowIndex];
  const status = String(row[headers.indexOf('status')]).trim();
  if (status !== 'PENDING_VERIFICATION') return { success: false, message: 'Booking is no longer pending verification.' };

  sheet.getRange(rowIndex + 1, headers.indexOf('utr') + 1).setValue(utr);
  sheet.getRange(rowIndex + 1, headers.indexOf('paymentProofUrl') + 1).setValue(payload.paymentProofUrl || '');
  sheet.getRange(rowIndex + 1, headers.indexOf('paymentProofFileId') + 1).setValue(payload.paymentProofFileId || '');
  sheet.getRange(rowIndex + 1, headers.indexOf('updatedAt') + 1).setValue(new Date().toISOString());
  return { success: true, message: 'Payment submitted for verification.' };
}

function adminApproveBooking(bookingId, action, reason) {
  const session = getActiveAdminSession();
  if (!session) return { success: false, message: 'Admin session expired.' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const rowIndex = rows.findIndex(row => String(row[0]) === String(bookingId));
  if (rowIndex <= 0) return { success: false, message: 'Booking not found.' };
  const originalStatus = String(rows[rowIndex][headers.indexOf('status')]);

  if (action === 'approve') {
    sheet.getRange(rowIndex + 1, headers.indexOf('status') + 1).setValue('APPROVED');
    sheet.getRange(rowIndex + 1, headers.indexOf('adminVerifiedAt') + 1).setValue(new Date().toISOString());
    const token = generateVoucherToken();
    const tokenHash = sha256Hex(token);
    const voucherSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vouchers');
    voucherSheet.appendRow([Utilities.getUuid(), bookingId, tokenHash, 'ACTIVE', new Date(Date.now() + 1000 * 60 * 60 * 24 * 7).toISOString(), new Date().toISOString(), '', '']);
    sheet.getRange(rowIndex + 1, headers.indexOf('voucherTokenHash') + 1).setValue(tokenHash);
    sheet.getRange(rowIndex + 1, headers.indexOf('voucherStatus') + 1).setValue('ACTIVE');
    return { success: true, bookingId, voucherToken: token, message: 'Booking approved and voucher created.' };
  }

  sheet.getRange(rowIndex + 1, headers.indexOf('status') + 1).setValue('REJECTED');
  sheet.getRange(rowIndex + 1, headers.indexOf('rejectionReason') + 1).setValue(reason || 'Payment rejected by admin.');
  const itemId = rows[rowIndex][headers.indexOf('itemId')];
  releaseReservedStock(itemId);
  return { success: true, message: 'Booking rejected and stock released.' };
}

function generateVoucherToken() {
  return Utilities.getUuid().replace(/-/g, '') + '-' + Math.random().toString(36).slice(2, 12);
}

function redeemVoucher(token) {
  const hash = sha256Hex(String(token || ''));
  const vouchers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vouchers');
  const values = vouchers.getDataRange().getValues();
  const headers = values[0];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const storedHash = row[headers.indexOf('tokenHash')];
    if (String(storedHash) === String(hash)) {
      const status = String(row[headers.indexOf('status')]).trim();
      if (status !== 'ACTIVE') return { success: false, message: 'Voucher is not active.' };
      const expiresAt = row[headers.indexOf('expiresAt')];
      if (expiresAt && new Date(expiresAt).getTime() < Date.now()) {
        vouchers.getRange(i + 1, headers.indexOf('status') + 1).setValue('EXPIRED');
        return { success: false, message: 'Voucher has expired.' };
      }
      vouchers.getRange(i + 1, headers.indexOf('status') + 1).setValue('REDEEMED');
      vouchers.getRange(i + 1, headers.indexOf('redeemedAt') + 1).setValue(new Date().toISOString());
      return { success: true, message: 'Voucher redeemed successfully.' };
    }
  }
  return { success: false, message: 'Voucher not found.' };
}

function computeDashboardStats() {
  const bookings = readSheetMap('Bookings');
  const menu = readSheetMap('Menu');
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  const pending = bookings.filter(b => b.status === 'PENDING_VERIFICATION').length;
  const approved = bookings.filter(b => b.status === 'APPROVED').length;
  const rejected = bookings.filter(b => b.status === 'REJECTED').length;
  const todayBookings = bookings.filter(b => String(b.createdAt || '').startsWith(today)).length;
  const totalApprovedRevenue = bookings.filter(b => b.status === 'APPROVED').reduce((sum, b) => sum + Number(b.amount || 0), 0);
  const todayRevenue = bookings.filter(b => b.status === 'APPROVED' && String(b.createdAt || '').startsWith(today)).reduce((sum, b) => sum + Number(b.amount || 0), 0);
  const redeemed = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vouchers').getDataRange().getValues().filter(r => String(r[3] || '').toLowerCase() === 'redeemed').length;
  const remainingStock = menu.reduce((sum, item) => sum + (Number(item.stock || 0) - Number(item.reservedStock || 0)), 0);
  return {
    pending,
    approved,
    totalApprovedRevenue,
    rejected,
    todayBookings,
    todayRevenue,
    redeemed,
    remainingStock
  };
}

function adminLogout() {
  PropertiesService.getUserProperties().deleteProperty('soni_admin_session');
  return { success: true, message: 'Logged out.' };
}

function getBookingsForAdmin() {
  const session = getActiveAdminSession();
  if (!session) return { success: false, message: 'Unauthorized' };
  return { success: true, bookings: readSheetMap('Bookings') };
}

function saveMenuItem(item) {
  const session = getActiveAdminSession();
  if (!session) return { success: false, message: 'Unauthorized' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Menu');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idx = values.findIndex(row => String(row[0]) === String(item.id));
  const row = [
    item.id,
    item.name,
    item.category,
    Number(item.price || 0),
    Number(item.offerPrice || 0),
    Number(item.stock || 0),
    Number(item.stock || 0),
    Number(item.reservedStock || 0),
    item.image,
    item.featured ? 'true' : 'false',
    item.active ? 'true' : 'false',
    item.description || ''
  ];
  if (idx >= 0) {
    sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, message: 'Menu item saved.' };
}

function saveOffer(offer) {
  const session = getActiveAdminSession();
  if (!session) return { success: false, message: 'Unauthorized' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Offers');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idx = values.findIndex(row => String(row[0]) === String(offer.id));
  const row = [offer.id, offer.title, offer.subtitle, offer.discount, offer.active ? 'true' : 'false'];
  if (idx >= 0) {
    sheet.getRange(idx + 1, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  return { success: true, message: 'Offer saved.' };
}

function saveSettings(settings) {
  const session = getActiveAdminSession();
  if (!session) return { success: false, message: 'Unauthorized' };
  updateConfigValue('announcement', String(settings.announcement || ''));
  updateConfigValue('bookingOpen', String(settings.bookingOpen === true || settings.bookingOpen === 'true'));
  updateConfigValue('upiId', settings.upiId || '8385963066@upi');
  updateConfigValue('upiName', settings.upiName || 'SONI FAST FOOD');
  updateConfigValue('timeSlots', JSON.stringify(settings.timeSlots || []));
  return { success: true, message: 'Settings updated.' };
}

function uploadToDrive(base64Data, fileName, mimeType) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = DriveApp.createFile(blob);
  return { id: file.getId(), url: file.getUrl() };
}

function getDashboardState() {
  const session = getActiveAdminSession();
  if (!session) return { success: false, message: 'Unauthorized' };
  return { success: true, state: getAppState() };
}

function checkInitialSetup() {
  const config = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (!config) return { needsSetup: true };
  const values = config.getDataRange().getValues();
  const map = {};
  values.forEach(row => { if (row[0]) map[String(row[0])] = row[1]; });
  return { needsSetup: !(map.pinHash && String(map.pinChanged || 'false') === 'true') };
}

function verifyCustomerBooking(bookingId) {
  const rows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Bookings').getDataRange().getValues();
  const headers = rows[0];
  const row = rows.find(r => String(r[0]) === String(bookingId));
  if (!row) return { success: false, message: 'Booking not found.' };
  return { success: true, booking: Object.fromEntries(headers.map((h, i) => [h, row[i]])) };
}

function getPublicMenu() {
  return getMenuByCategory();
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const action = payload.action || 'noop';
    if (action === 'createBooking') return ContentService.createTextOutput(JSON.stringify(createBooking(payload))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'submitPayment') return ContentService.createTextOutput(JSON.stringify(submitPayment(payload))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'adminLogin') return ContentService.createTextOutput(JSON.stringify(adminLogin(payload.pin))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'adminSetPin') return ContentService.createTextOutput(JSON.stringify(adminSetPin(payload.oldPin, payload.newPin))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'adminApproveBooking') return ContentService.createTextOutput(JSON.stringify(adminApproveBooking(payload.bookingId, payload.actionType, payload.reason))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'saveMenuItem') return ContentService.createTextOutput(JSON.stringify(saveMenuItem(payload.item))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'saveOffer') return ContentService.createTextOutput(JSON.stringify(saveOffer(payload.offer))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'saveSettings') return ContentService.createTextOutput(JSON.stringify(saveSettings(payload.settings))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'redeemVoucher') return ContentService.createTextOutput(JSON.stringify(redeemVoucher(payload.token))).setMimeType(ContentService.MimeType.JSON);
    if (action === 'uploadProof') return ContentService.createTextOutput(JSON.stringify(uploadToDrive(payload.base64, payload.fileName, payload.mimeType))).setMimeType(ContentService.MimeType.JSON);
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Unsupported action.' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
