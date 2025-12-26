// =======================
// VENDOR FUNCTIONS
// =======================

function getVendors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendors");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // Filter out empty rows
  const filtered = data.filter(row => row.some(cell => cell !== ""));

  // Map each row to an object
  return filtered.map(r => ({
    id: r[0],       // VendorID
    name: r[1],     // Vendor Name
    contact: r[2],  // Contact Person
    email: r[3],    // Email
    phone: r[4],    // Phone
    address: r[5]   // Address
  }));
}

// Add Vendor
function addVendor(vendor) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendors");

  // Ensure all fields are present
  const row = [
    vendor.id || "",
    vendor.name || "",
    vendor.contact || "",
    vendor.email || "",
    vendor.phone || "",
    vendor.address || ""
  ];

  sheet.appendRow(row);

  return vendor;
}

//Delete Vendor
function deleteVendorById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendors");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function getVendorById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vendors");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  for (let r of data) {
    if (r[0] == id) return { id: r[0], name: r[1], contact: r[2], email: r[3], phone: r[4], address: r[5] };
  }
  return null;
}



// =======================
// PURCHASE ORDER FUNCTIONS
// =======================

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Purchase Orders");
}

// Fetch Purchase Orders
function getPurchaseOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PurchaseOrders");

  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  return data.map(row => ({
    poid: row[0],
    vendorId: row[1],
    item: row[2],
    quantity: row[3],
    price: row[4],
    status: row[5],
    date: Utilities.formatDate(
      new Date(row[6]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    )
  }));
}

function deletePurchaseOrder(poid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PurchaseOrders");

  if (!sheet) return false;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // POID column

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === poid) {
      sheet.deleteRow(i + 2); // +2 because data starts at row 2
      return true;
    }
  }

  return false;
}





// Add Purchase Order
function addPurchaseOrder(id, vendorID, item, qty, price) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PurchaseOrders");
  if (!sheet) throw new Error("PurchaseOrders sheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] == vendorID && data[i][2].toLowerCase() === item.toLowerCase()) {
        throw new Error("Duplicate PO: This vendor already has the same item.");
      }
    }
  }

  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  sheet.appendRow([id, vendorID, item, qty, price, "Pending", date]);

  return { id, vendorID, item, qty, price, status: "Pending", date };
}


function updatePOStatus(id, newStatus) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PurchaseOrders");
  if (!sheet) throw new Error("PurchaseOrders sheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("No POs found.");

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.getRange(i + 2, 6).setValue(newStatus);
      return { id, status: newStatus };
    }
  }

  throw new Error("PO not found.");
}



// =======================
// SERVE FRONTEND
// =======================

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Purchase Order Management")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
