// ============================================================\
//  EmpowerKick Tattoo Store — Code.gs
//  Google Apps Script Backend
//  Optimized: Server-side CacheService + Batch reads
// ============================================================\

var SHEET_ID   = '1pnxhVy6dLhAy6-TpHGRXGjFT6ULCV7sqBY-aHAcwzJs';
const CACHE_KEY  = 'ek_products_v2';
const CACHE_TTL  = 360; 

// ============================================================\
//  ROUTER — doGet
// ============================================================\
function doGet(e) {
  const param = e && e.parameter ? e.parameter : {};

  if (param.api === 'products') {
    return getProductsJson();
  }
  if (param.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('แผงควบคุมแอดมิน — EmpowerKick')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('EmpowerKick Tattoo Store')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================\
//  PRODUCTS MANAGEMENT (ดึง / เพิ่ม / แก้ไข / ลบ)
// ============================================================\

function getProducts() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY);
    if (cached) return JSON.parse(cached);

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Products');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const headers = data[0];
    const products = [];

    for (let i = 1; i < data.length; i++) {
      let row = data[i];
      let p = {};
      headers.forEach((h, idx) => { p[h] = row[idx]; });
      products.push(p);
    }

    cache.put(CACHE_KEY, JSON.stringify(products), CACHE_TTL);
    return products;
  } catch(e) {
    return [];
  }
}

function getProductsJson() {
  const data = getProducts();
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันแปลงลิงก์รูปภาพให้เป็น Direct Link อัตโนมัติก่อนลงชีต
function cleanImageUrl(url) {
  if (!url) return "";
  url = url.trim();
  if (url.includes("drive.google.com")) {
    var driveMatch = url.match(/\/file\/d\/([\w-]+)/);
    if (driveMatch && driveMatch[1]) return "https://lh3.googleusercontent.com/d/...0" + driveMatch[1];
  } else if (url.includes("ibb.co/") && !url.includes("i.ibb.co/")) {
    var imgbbMatch = url.match(/ibb\.co\/([\w-]+)/);
    if (imgbbMatch && imgbbMatch[1]) return "https://i.ibb.co/" + imgbbMatch[1] + "/image.png";
  }
  return url;
}

// เพิ่มสินค้าใหม่
function addProduct(p) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Products');
  const id = 'TAT-' + new Date().getTime();
  
  const finalImgUrl = cleanImageUrl(p.image_url);

  sheet.appendRow([
    id,
    p.name,
    p.category,
    p.size,
    Number(p.price) || 0,
    Number(p.stock) || 0,
    finalImgUrl,
    'ready'
  ]);

  clearCache();
  return 'เพิ่มสินค้าสำเร็จ รหัส: ' + id;
}

// อัปเดตสต๊อกสินค้า (เพิ่ม/ลด หรือแก้ไขค่าต่างๆ)
function updateProduct(id, updatedData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Products');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      if (updatedData.stock !== undefined) sheet.getRange(i + 1, 6).setValue(Number(updatedData.stock));
      if (updatedData.price !== undefined) sheet.getRange(i + 1, 5).setValue(Number(updatedData.price));
      if (updatedData.status !== undefined) sheet.getRange(i + 1, 8).setValue(updatedData.status);
      clearCache();
      return 'อัปเดตข้อมูลสินค้าสำเร็จ!';
    }
  }
  throw new Error('ไม่พบสินค้าที่ต้องการอัปเดต');
}

// ลบสินค้าออก
function deleteProduct(id) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Products');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      clearCache();
      return 'ลบสินค้าเรียบร้อยแล้ว';
    }
  }
  throw new Error('ไม่พบสินค้าที่ต้องการลบ');
}

// ============================================================\
//  ORDERS MANAGEMENT (ดูออเดอร์)
// ============================================================\

function saveOrder(orderData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Orders');
  const orderId = 'ORD-' + new Date().getTime();

  sheet.appendRow([
    new Date(),
    orderId,
    orderData.items,
    Number(orderData.total) || 0,
    orderData.paymentMethod  || '',
    orderData.customerPhone  || '', 
    orderData.shippingAddress || '', 
    'รอการยืนยัน'
  ]);

  return 'สั่งซื้อสำเร็จ! เลขที่: ' + orderId;
}

function getOrders() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Orders');
  const data  = sheet.getDataRange().getDisplayValues();
  return data.length > 1 ? data.slice(1).reverse() : []; 
}

// ตัวอย่างที่ 1: การเพิ่มสินค้า
function addProduct(p) {
  // ... (โค้ดเดิม)
  sheet.appendRow([ id, p.name, p.category, p.size, Number(p.price) || 0, Number(p.stock) || 0, finalImgUrl, 'ready' ]);
  
  SpreadsheetApp.flush(); // 👈 เพิ่มบรรทัดนี้ (บังคับให้ชีตเขียนข้อมูลทันทีก่อนข้ามไปทำอย่างอื่น)
  clearCache();
  return 'เพิ่มสินค้าสำเร็จ รหัส: ' + id;
}

// ตัวอย่างที่ 2: การเซฟออเดอร์
function saveOrder(orderData) {
  // ... (โค้ดเดิม)
  sheet.appendRow([ new Date(), orderId, orderData.items, Number(orderData.total) || 0, orderData.paymentMethod || '', orderData.customerPhone || '', orderData.shippingAddress || '', 'รอการยืนยัน' ]);

  SpreadsheetApp.flush(); // 👈 เพิ่มบรรทัดนี้เช่นกัน
  return 'สั่งซื้อสำเร็จ! เลขที่: ' + orderId;
}

function clearCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

