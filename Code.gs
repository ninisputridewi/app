/* =========================================================
   THE TANGERINE ROOM — Google Apps Script backend
   Pakai Google Sheet sebagai database. Tidak butuh Supabase
   atau hosting terpisah — Apps Script yang menyajikan website
   DAN menyimpan datanya sekaligus.

   CARA PAKAI: lihat SETUP_GUIDE_GoogleSheets.md
========================================================= */

const SHEET_HEADERS = {
  Products:            ['id','name','category','unit_price','cost_price','active','created_at'],
  Customers:           ['id','name','contact_person','phone','email','address','notes','active','created_at'],
  Invoices:            ['id','invoice_number','customer_id','invoice_date','due_date','status','subtotal','discount','total','amount_paid','notes','created_at'],
  InvoiceItems:        ['id','invoice_id','product_id','product_name','quantity','unit_price','cost_price','line_total'],
  Ingredients:         ['id','name','category','unit','active','created_at'],
  IngredientPurchases: ['id','purchase_date','ingredient_id','ingredient_name','supplier','quantity','unit','unit_cost','total_cost','notes','created_at'],
  AssetPurchases:      ['id','purchase_date','item_name','category','supplier','quantity','unit_cost','total_cost','notes','created_at'],
  MaintenanceRecords:  ['id','maintenance_date','item_name','description','vendor','cost','notes','created_at'],
  Employees:           ['id','name','role','payment_type','rate','active','notes','created_at'],
  PayrollPayments:     ['id','employee_id','employee_name','period_start','period_end','amount','date_paid','status','notes','created_at']
};

/* =========================================================
   WEB APP ENTRY POINT
========================================================= */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('The Tangerine Room — Business Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   ONE-TIME SETUP
   Run this once from the Apps Script editor (select
   "setupSheets" in the function dropdown, click Run) to
   create all tabs + starter data. Safe to run again later —
   it will NOT duplicate seed data if the tabs already have rows.
========================================================= */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEET_HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    const isNew = !sheet;
    if (!sheet) sheet = ss.insertSheet(name);
    if (isNew || sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, SHEET_HEADERS[name].length).setValues([SHEET_HEADERS[name]]);
      sheet.setFrozenRows(1);
    }
  });

  // remove the default empty "Sheet1" if present
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) {
    ss.deleteSheet(def);
  }

  if (readSheet('Products').length === 0) seedProducts();
  if (readSheet('Ingredients').length === 0) seedIngredients();

  Logger.log('Setup selesai. Cek tab-tab baru di spreadsheet ini.');
}

function seedProducts() {
  const rows = [
    ['Lemon Cake', 'Cake', 45000, 20000],
    ['Banana Bread', 'Bread', 40000, 18000],
    ['Classic Tiramisu', 'Tiramisu', 55000, 25000],
    ['Biscoff Tiramisu', 'Tiramisu', 60000, 28000],
    ['Peanut Cake', 'Cake', 45000, 20000],
    ['Basque Burnt Cheesecake', 'Cheesecake', 65000, 30000],
    ['Pumpkin Pie', 'Pie', 50000, 22000],
    ['Donut', 'Pastry', 15000, 6000]
  ];
  rows.forEach(r => appendRow('Products', { name: r[0], category: r[1], unit_price: r[2], cost_price: r[3], active: true }));
}

function seedIngredients() {
  const rows = [
    ['All Purpose Flour', 'Flour & Starch', 'kg'],
    ['Almond Flour', 'Flour & Starch', 'kg'],
    ['Glutinous Rice Flour', 'Flour & Starch', 'kg'],
    ['Ube Flour', 'Flour & Starch', 'kg'],
    ['Cornstarch', 'Flour & Starch', 'kg'],
    ['Cocoa Powder', 'Chocolate & Cocoa', 'kg'],
    ['Baking Powder', 'Leavening', 'kg'],
    ['Baking Soda', 'Leavening', 'kg'],
    ['Icing Sugar', 'Sugar & Sweetener', 'kg'],
    ['Granulated Sugar', 'Sugar & Sweetener', 'kg'],
    ['Demerara Sugar', 'Sugar & Sweetener', 'kg'],
    ['Salt', 'Seasoning', 'kg'],
    ['Maldon Sea Salt', 'Seasoning', 'kg'],
    ['Oolong Tea (Dilmah)', 'Flavoring & Misc', 'box'],
    ['Nescafe Coffee Powder', 'Flavoring & Misc', 'kg'],
    ['Lotus Biscoff', 'Flavoring & Misc', 'kg'],
    ['Gelatin Mass', 'Setting Agent', 'kg'],
    ['Gelatin Leaf', 'Setting Agent', 'pack'],
    ['Agar Agar', 'Setting Agent', 'kg'],
    ['Pectin NH', 'Setting Agent', 'kg'],
    ['Peanut', 'Nuts', 'kg'],
    ['Peanut Butter', 'Nuts', 'kg'],
    ['Ube Flavor', 'Flavoring & Misc', 'bottle'],
    ['Elle Unsalted Butter', 'Dairy & Cream', 'kg'],
    ['Anchor Unsalted Butter', 'Dairy & Cream', 'kg'],
    ['Sour Cream', 'Dairy & Cream', 'kg'],
    ['Green Fields Whip Cream', 'Dairy & Cream', 'liter'],
    ['Ever Whip Cream', 'Dairy & Cream', 'liter'],
    ['Mascarpone', 'Dairy & Cream', 'kg'],
    ['Cream Cheese', 'Dairy & Cream', 'kg'],
    ['Parmesan Cheese', 'Dairy & Cream', 'kg'],
    ['Brie Cheese', 'Dairy & Cream', 'kg'],
    ['Full Fat Milk', 'Dairy & Cream', 'liter'],
    ['Fresh Peach', 'Fruit', 'kg'],
    ['Ube', 'Fruit', 'kg'],
    ['Frozen Blueberry', 'Fruit', 'kg'],
    ['Peach Concentrate', 'Fruit', 'bottle'],
    ['Callebaut White Chocolate', 'Chocolate & Cocoa', 'kg'],
    ['Callebaut 823 Milk Chocolate', 'Chocolate & Cocoa', 'kg'],
    ['Egg', 'Dairy & Cream', 'kg'],
    ['Egg White', 'Dairy & Cream', 'kg'],
    ['Egg Yolk', 'Dairy & Cream', 'kg'],
    ['Baking Paper', 'Packaging', 'roll'],
    ['Water', 'General', 'liter'],
    ['Glucose Syrup', 'Sugar & Sweetener', 'kg'],
    ['Oil', 'General', 'liter'],
    ['Cinnamon Powder', 'Flavoring & Misc', 'kg'],
    ['Ambon Banana', 'Fruit', 'kg'],
    ['Lemon', 'Fruit', 'kg'],
    ['Palmia Margarine', 'Dairy & Cream', 'kg'],
    ['Adore Chocolate 58%', 'Chocolate & Cocoa', 'kg'],
    ['Adore Chocolate 40%', 'Chocolate & Cocoa', 'kg'],
    ['Adore Chocolate 70%', 'Chocolate & Cocoa', 'kg'],
    ['Choco Chips VH', 'Chocolate & Cocoa', 'kg'],
    ['Carrot', 'Fruit', 'kg'],
    ['VH Dark Chocolate', 'Chocolate & Cocoa', 'kg'],
    ['Essence Chocolate', 'Flavoring & Misc', 'bottle']
  ];
  rows.forEach(r => appendRow('Ingredients', { name: r[0], category: r[1], unit: r[2], active: true }));
}

/* =========================================================
   GENERIC SHEET HELPERS
========================================================= */
function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" belum ada. Jalankan fungsi setupSheets() dulu dari Apps Script editor.');
  return sheet;
}

function formatCell_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v;
}

function readSheet(sheetName) {
  const sheet = getSheet_(sheetName);
  const headers = SHEET_HEADERS[sheetName];
  const idCol = headers.indexOf('id');
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    if (!values[r][idCol]) continue; // skip blank rows
    const obj = {};
    headers.forEach((h, c) => { obj[h] = formatCell_(values[r][c]); });
    rows.push(obj);
  }
  return rows;
}

function appendRow(sheetName, obj) {
  const sheet = getSheet_(sheetName);
  const headers = SHEET_HEADERS[sheetName];
  if (!obj.id) obj.id = Utilities.getUuid();
  if (headers.indexOf('created_at') >= 0 && !obj.created_at) obj.created_at = new Date();
  if (headers.indexOf('active') >= 0 && obj.active === undefined) obj.active = true;
  const row = headers.map(h => (obj[h] !== undefined && obj[h] !== null) ? obj[h] : '');
  sheet.appendRow(row);
  return obj;
}

function updateRow(sheetName, id, updates) {
  const sheet = getSheet_(sheetName);
  const headers = SHEET_HEADERS[sheetName];
  const idCol = headers.indexOf('id');
  const values = sheet.getDataRange().getValues();
  for (let r = 1; r < values.length; r++) {
    if (values[r][idCol] === id) {
      headers.forEach((h, c) => {
        if (updates[h] !== undefined) sheet.getRange(r + 1, c + 1).setValue(updates[h]);
      });
      return true;
    }
  }
  return false;
}

function deleteRowGeneric(sheetName, id) {
  const sheet = getSheet_(sheetName);
  const headers = SHEET_HEADERS[sheetName];
  const idCol = headers.indexOf('id');
  const values = sheet.getDataRange().getValues();
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r][idCol] === id) sheet.deleteRow(r + 1);
  }
  return true;
}

/* =========================================================
   PUBLIC API — called from Index.html via google.script.run
========================================================= */
function getAllData() {
  return {
    products: readSheet('Products'),
    customers: readSheet('Customers'),
    employees: readSheet('Employees'),
    ingredients: readSheet('Ingredients'),
    invoices: readSheet('Invoices'),
    invoiceItems: readSheet('InvoiceItems'),
    ingredientPurchases: readSheet('IngredientPurchases'),
    assetPurchases: readSheet('AssetPurchases'),
    maintenance: readSheet('MaintenanceRecords'),
    payroll: readSheet('PayrollPayments')
  };
}

// ---- Invoices ----
function addInvoice(invoicePayload, items) {
  const inv = appendRow('Invoices', invoicePayload);
  items.forEach(it => {
    it.invoice_id = inv.id;
    appendRow('InvoiceItems', it);
  });
  return inv;
}
function markInvoicePaid(id, total) {
  return updateRow('Invoices', id, { status: 'paid', amount_paid: total });
}
function deleteInvoiceCascade(id) {
  const sheet = getSheet_('InvoiceItems');
  const headers = SHEET_HEADERS['InvoiceItems'];
  const fkCol = headers.indexOf('invoice_id');
  const values = sheet.getDataRange().getValues();
  for (let r = values.length - 1; r >= 1; r--) {
    if (values[r][fkCol] === id) sheet.deleteRow(r + 1);
  }
  deleteRowGeneric('Invoices', id);
  return true;
}

// ---- Products ----
function saveProduct(payload) {
  if (payload.id) { updateRow('Products', payload.id, payload); return payload; }
  return appendRow('Products', payload);
}

// ---- Customers ----
function saveCustomer(payload) {
  if (payload.id) { updateRow('Customers', payload.id, payload); return payload; }
  return appendRow('Customers', payload);
}

// ---- Ingredient master + purchases ----
function addIngredientMaster(payload) { return appendRow('Ingredients', payload); }
function addIngredientPurchase(payload) { return appendRow('IngredientPurchases', payload); }

// ---- Tools/Utensil purchases + Maintenance ----
function addAssetPurchase(payload) { return appendRow('AssetPurchases', payload); }
function addMaintenance(payload) { return appendRow('MaintenanceRecords', payload); }

// ---- Employees + Payroll ----
function addEmployee(payload) { return appendRow('Employees', payload); }
function addPayroll(payload) { return appendRow('PayrollPayments', payload); }

// ---- Generic delete (used by Settings / list rows) ----
function deleteRowByTable(tableName, id) {
  const map = {
    products: 'Products',
    customers: 'Customers',
    employees: 'Employees',
    ingredients: 'Ingredients',
    ingredient_purchases: 'IngredientPurchases',
    asset_purchases: 'AssetPurchases',
    maintenance_records: 'MaintenanceRecords',
    payroll_payments: 'PayrollPayments'
  };
  const sheetName = map[tableName];
  if (!sheetName) throw new Error('Tabel tidak dikenal: ' + tableName);
  return deleteRowGeneric(sheetName, id);
}
