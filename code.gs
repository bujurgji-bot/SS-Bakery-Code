// ================= WEB APP ENTRY POINT =================
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SS Bakery House POS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================= AUTHENTICATION & RBAC =================
function authenticateUser(pin) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Staff");
    if (!sheet) return { success: false, message: 'Sheet "Staff" not found.' };
    const data = sheet.getDataRange().getValues();
    const inputPin = String(pin).trim();
    for (let i = 1; i < data.length; i++) {
      const sheetPin = String(data[i][2]).trim();
      const status = String(data[i][4]).trim();
      if (sheetPin === inputPin && status.toLowerCase() === "active") {
        return { success: true, staffName: data[i][1], role: data[i][3] };
      }
    }
    return { success: false, message: 'Invalid PIN or Inactive Staff Member' };
  } catch (error) {
    return { success: false, message: 'Server Error: ' + error.toString() };
  }
}

function logLogoutEvent(userName) {
  try {
    const sheet = getOrCreateSheet("AuditLog", ["Timestamp", "User", "Action"]);
    sheet.appendRow([new Date(), userName, "Logout"]);
  } catch (e) { console.error("Failed to log logout event", e); }
}

// ================= PRODUCTS & INVENTORY =================
function getProducts() {
  try {
    const sheet = getOrCreateSheet("Products", ["ID","Name","Category","Supplier","Price","UnitType","Stock","Image"]);
    const data = sheet.getDataRange().getValues();
    const products = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] || data[i][1]) {
        products.push({
          id: String(data[i][0]), name: String(data[i][1]),
          category: String(data[i][2] || 'General'), subCategorySupplier: String(data[i][3] || ''),
          price: Number(data[i][4] || 0), unitType: String(data[i][5] || 'PCS'),
          stock: Number(data[i][6] || 0), image: String(data[i][7] || '')
        });
      }
    }
    return products;
  } catch (error) { console.error("Error in getProducts:", error); return []; }
}

function addNewProduct(payload) {
  try {
    const sheet = getOrCreateSheet("Products", ["ID","Name","Category","Supplier","Price","UnitType","Stock","Image"]);
    const prodId = 'PROD-' + Math.floor(1000 + Math.random() * 9000);
    const imageUrl = payload.imageBase64 ? payload.imageBase64 : (payload.image || '');
    sheet.appendRow([prodId, payload.name, payload.category, payload.supplier, payload.price, payload.unitType, payload.stock, imageUrl]);
    return { success: true, productId: prodId };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function recordSupplierPurchase(payload) {
  try {
    const purSheet = getOrCreateSheet("Purchases", ["Timestamp","Supplier","Product","Qty","UnitCost","LoggedBy"]);
    purSheet.appendRow([new Date(), payload.supplier, payload.product, payload.qty, payload.unitCost, payload.loggedBy]);
    const prodSheet = getOrCreateSheet("Products");
    const data = prodSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(payload.product).toLowerCase()) {
        prodSheet.getRange(i + 1, 7).setValue(Number(data[i][6] || 0) + Number(payload.qty));
        break;
      }
    }
    return { success: true };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// ================= SALES & TRANSACTIONS =================
function saveOrder(payload) {
  try {
    const salesSheet = getOrCreateSheet("Sales", ["InvoiceID","Timestamp","Items","Subtotal","Total","PaymentMode","Phone","Notes","BilledBy"]);
    const itemsSheet = getOrCreateSheet("Sales_Items", ["InvoiceID","Timestamp","Item Name","Qty","Unit Price","Total Price","Billed By"]);
    const invoiceId = 'INV-' + Math.floor(100000 + Math.random() * 900000);
    const timestamp = new Date();
    const itemsSummary = payload.items.map(it => `${it.name} (${it.qty}x@₹${it.price})`).join(', ');
    salesSheet.appendRow([invoiceId, timestamp, itemsSummary, payload.subtotal, payload.total, payload.paymentMode, payload.customerPhone || '', payload.notes || '', payload.billedBy]);
    payload.items.forEach(item => {
      itemsSheet.appendRow([invoiceId, timestamp, item.name, item.qty, item.price, item.qty * item.price, payload.billedBy]);
    });
    deductInventoryStock(payload.items);
    return { success: true, invoiceId: invoiceId };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function deductInventoryStock(cartItems) {
  try {
    const sheet = getOrCreateSheet("Products");
    const data = sheet.getDataRange().getValues();
    cartItems.forEach(item => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === item.id || String(data[i][1]).toLowerCase() === item.name.toLowerCase()) {
          sheet.getRange(i + 1, 7).setValue(Math.max(0, Number(data[i][6] || 0) - Number(item.qty)));
          break;
        }
      }
    });
  } catch (e) { console.error("Failed to deduct stock:", e); }
}

function getRecentTransactions() {
  try {
    const sheet = getOrCreateSheet("Sales");
    const data = sheet.getDataRange().getValues();
    const transactions = [];
    for (let i = data.length - 1; i >= 1 && transactions.length < 50; i--) {
      if (data[i][0]) {
        transactions.push({
          invoiceId: String(data[i][0]),
          timestamp: data[i][1] instanceof Date ? data[i][1].toLocaleString() : String(data[i][1]),
          items: String(data[i][2]), total: Number(data[i][4] || 0),
          paymentMode: String(data[i][5]), billedBy: String(data[i][8])
        });
      }
    }
    return transactions;
  } catch (error) { console.error("Error in getRecentTransactions:", error); return []; }
}

// ================= FUTURE CAKE ORDERS =================
function saveFutureOrder(payload) {
  try {
    const sheet = getOrCreateSheet("FutureOrders", [
      "OrderID","Customer","Phone","DeliveryDateTime","Occasion","Instructions",
      "CakeType","Size","Flavour","Theme","Addons","Total","Advance","Balance","Status","Staff"
    ]);
    const orderId = 'ORD-' + Math.floor(1000 + Math.random() * 9000);
    sheet.appendRow([orderId, payload.name, payload.phone, payload.dateTime, payload.occasion || '', payload.instructions || '', payload.cakeType, payload.size, payload.flavour, payload.theme || '', payload.addons || '', payload.total, payload.advance, payload.total - payload.advance, 'Pending', payload.staff]);
    return { success: true, orderId: orderId };
  } catch (error) { return { success: false, message: error.toString() }; }
}

function getFutureOrders() {
  try {
    const sheet = getOrCreateSheet("FutureOrders");
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    const balColIdx = headers.indexOf("BalanceCollectedOn");
    const orders = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        orders.push({
          id: String(data[i][0]), name: String(data[i][1]), phone: String(data[i][2]),
          dateTime: String(data[i][3]), cakeType: String(data[i][6]), size: String(data[i][7]),
          flavour: String(data[i][8]), total: Number(data[i][11] || 0),
          advance: Number(data[i][12] || 0), balance: Number(data[i][13] || 0),
          status: String(data[i][14] || 'Pending'),
          balanceCollectedOn: balColIdx >= 0 && data[i][balColIdx] ? (data[i][balColIdx] instanceof Date ? data[i][balColIdx].toLocaleString('en-IN') : String(data[i][balColIdx])) : '',
          balanceCollectedBy: balColIdx >= 0 ? String(data[i][balColIdx + 1] || '') : '',
          balancePaymentMode: balColIdx >= 0 ? String(data[i][balColIdx + 2] || '') : ''
        });
      }
    }
    return orders;
  } catch (error) { console.error("Error in getFutureOrders:", error); return []; }
}

function updateFutureOrderStatus(orderId, status) {
  try {
    const sheet = getOrCreateSheet("FutureOrders");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(orderId)) {
        sheet.getRange(i + 1, 15).setValue(status);
        return { success: true };
      }
    }
    return { success: false, message: 'Order ID not found' };
  } catch (error) { return { success: false, message: error.toString() }; }
}

// NEW: Collect balance with full audit trail
function collectFutureOrderBalance(payload) {
  try {
    const sheet = getOrCreateSheet("FutureOrders");
    const data = sheet.getDataRange().getValues();
    const headers = data[0] || [];
    let balColIdx = headers.indexOf("BalanceCollectedOn");

    // Add audit columns to sheet header if they don't exist yet
    if (balColIdx === -1) {
      const nextCol = headers.length + 1;
      sheet.getRange(1, nextCol).setValue("BalanceCollectedOn");
      sheet.getRange(1, nextCol + 1).setValue("BalanceCollectedBy");
      sheet.getRange(1, nextCol + 2).setValue("BalancePaymentMode");
      sheet.getRange(1, 1, 1, nextCol + 2).setFontWeight("bold");
      balColIdx = nextCol - 1; // convert to 0-based
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.orderId)) {
        sheet.getRange(i + 1, 14).setValue(0);               // Clear Balance (col N)
        sheet.getRange(i + 1, 15).setValue('Fully Paid');    // Update Status (col O)
        sheet.getRange(i + 1, balColIdx + 1).setValue(new Date());          // Timestamp
        sheet.getRange(i + 1, balColIdx + 2).setValue(payload.collectedBy); // Who collected
        sheet.getRange(i + 1, balColIdx + 3).setValue(payload.paymentMode); // How collected
        return { success: true };
      }
    }
    return { success: false, message: 'Order not found' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// ================= ANALYTICS DASHBOARD (with Date Filters) =================
function getDashboardKPIs(startDateStr, endDateStr) {
  try {
    const sheet = getOrCreateSheet("Sales");
    const data = sheet.getDataRange().getValues();
    let totalSales = 0, totalOrders = 0;
    const hourlySales = {};
    const paymentBreakdown = {};

    const startDate = startDateStr ? new Date(startDateStr) : null;
    const endDate = endDateStr ? new Date(endDateStr) : null;
    if (endDate) endDate.setHours(23, 59, 59, 999);

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const rowDate = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
      if (isNaN(rowDate)) continue;
      if (startDate && rowDate < startDate) continue;
      if (endDate && rowDate > endDate) continue;

      const amount = Number(data[i][4] || 0);
      totalSales += amount;
      totalOrders++;

      // Hourly breakdown
      const hour = rowDate.getHours();
      const label = hour.toString().padStart(2, '0') + ':00';
      hourlySales[label] = (hourlySales[label] || 0) + amount;

      // Payment mode split
      const mode = String(data[i][5] || 'Cash');
      paymentBreakdown[mode] = (paymentBreakdown[mode] || 0) + amount;
    }

    return {
      totalSales, totalOrders,
      avgBill: totalOrders > 0 ? totalSales / totalOrders : 0,
      hourlySales, paymentBreakdown
    };
  } catch (error) {
    return { totalSales: 0, totalOrders: 0, avgBill: 0, hourlySales: {}, paymentBreakdown: {} };
  }
}

// ================= HELPER FUNCTIONS =================
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }
  }
  return sheet;
}

// ================= ADVANCED P&L & WASTAGE (with Date Filters) =================
function recordWastage(payload) {
  try {
    const sheet = getOrCreateSheet("Wastage", ["Timestamp","Product","Qty","Reason","LoggedBy"]);
    sheet.appendRow([new Date(), payload.product, payload.qty, payload.reason, payload.loggedBy]);
    const prodSheet = getOrCreateSheet("Products");
    const data = prodSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(payload.product).toLowerCase()) {
        prodSheet.getRange(i + 1, 7).setValue(Math.max(0, Number(data[i][6] || 0) - Number(payload.qty)));
        break;
      }
    }
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getAdvancedAnalytics(startDateStr, endDateStr) {
  try {
    const prodData = getOrCreateSheet("Products").getDataRange().getValues();
    const purData = getOrCreateSheet("Purchases").getDataRange().getValues();
    const salesItemsData = getOrCreateSheet("Sales_Items", ["InvoiceID","Timestamp","Item Name","Qty","Unit Price","Total Price","Billed By"]).getDataRange().getValues();
    const wastageData = getOrCreateSheet("Wastage", ["Timestamp","Product","Qty","Reason","LoggedBy"]).getDataRange().getValues();

    const startDate = startDateStr ? new Date(startDateStr) : null;
    const endDate = endDateStr ? new Date(endDateStr) : null;
    if (endDate) endDate.setHours(23, 59, 59, 999);

    const productMetrics = {};

    // Step 1: Seed from Products sheet
    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][1]) {
        productMetrics[String(prodData[i][1])] = {
          name: String(prodData[i][1]), price: Number(prodData[i][4] || 0),
          unitCost: 0, salesQty: 0, salesRevenue: 0, wastageQty: 0, purchasedQty: 0, totalPurchaseCost: 0
        };
      }
    }

    // Step 2: Unit cost from Purchases (all-time — cost doesn't change by date)
    for (let i = 1; i < purData.length; i++) {
      if (purData[i][2]) {
        const pName = String(purData[i][2]);
        const qty = Number(purData[i][3] || 0), cost = Number(purData[i][4] || 0);
        if (!productMetrics[pName]) productMetrics[pName] = { name: pName, price: 0, unitCost: 0, salesQty: 0, salesRevenue: 0, wastageQty: 0, purchasedQty: 0, totalPurchaseCost: 0 };
        productMetrics[pName].purchasedQty += qty;
        productMetrics[pName].totalPurchaseCost += qty * cost;
      }
    }
    for (const k in productMetrics) {
      const p = productMetrics[k];
      if (p.purchasedQty > 0) p.unitCost = p.totalPurchaseCost / p.purchasedQty;
    }

    // Step 3: Sales from Sales_Items (DATE FILTERED)
    let totalRevenue = 0;
    for (let i = 1; i < salesItemsData.length; i++) {
      const rowDate = salesItemsData[i][1] instanceof Date ? salesItemsData[i][1] : new Date(salesItemsData[i][1]);
      if (startDate && !isNaN(rowDate) && rowDate < startDate) continue;
      if (endDate && !isNaN(rowDate) && rowDate > endDate) continue;
      const pName = String(salesItemsData[i][2] || "").trim();
      if (!pName) continue;
      const qty = Number(salesItemsData[i][3] || 0), price = Number(salesItemsData[i][4] || 0), lineTotal = Number(salesItemsData[i][5] || 0);
      if (!productMetrics[pName]) productMetrics[pName] = { name: pName, price: price, unitCost: 0, salesQty: 0, salesRevenue: 0, wastageQty: 0, purchasedQty: 0, totalPurchaseCost: 0 };
      productMetrics[pName].salesQty += qty;
      productMetrics[pName].salesRevenue += lineTotal;
      totalRevenue += lineTotal;
    }

    // Step 4: Wastage (DATE FILTERED)
    let totalWastageCost = 0;
    for (let i = 1; i < wastageData.length; i++) {
      const rowDate = wastageData[i][0] instanceof Date ? wastageData[i][0] : new Date(wastageData[i][0]);
      if (startDate && !isNaN(rowDate) && rowDate < startDate) continue;
      if (endDate && !isNaN(rowDate) && rowDate > endDate) continue;
      const pName = String(wastageData[i][1]);
      const qty = Number(wastageData[i][2] || 0);
      if (productMetrics[pName]) {
        productMetrics[pName].wastageQty += qty;
        totalWastageCost += qty * productMetrics[pName].unitCost;
      }
    }

    // Step 5: Build product list
    let totalCOGS = 0, productList = [];
    for (const k in productMetrics) {
      const p = productMetrics[k];
      totalCOGS += p.salesQty * p.unitCost;
      p.margin = p.price > 0 ? ((p.price - p.unitCost) / p.price) * 100 : 0;
      if (p.salesQty > 0 || p.wastageQty > 0) productList.push(p);
    }

    productList.sort((a, b) => b.salesQty - a.salesQty);
    const topSelling = productList.slice(0, 5);
    const bottomSelling = productList.filter(p => p.salesQty > 0).sort((a, b) => a.salesQty - b.salesQty).slice(0, 5);
    const wastageLines = productList.filter(p => p.wastageQty > 0).sort((a, b) => (b.wastageQty * b.unitCost) - (a.wastageQty * a.unitCost));

    return {
      pnl: { revenue: totalRevenue, cogs: totalCOGS, wastageCost: totalWastageCost, netProfit: totalRevenue - totalCOGS - totalWastageCost },
      topSelling, bottomSelling, wastageLines
    };
  } catch (e) { console.error("Analytics error:", e); return null; }
}

// ================= ONE-TIME OLD DATA MIGRATOR =================
function fixOldSalesData() {
  const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales");
  if (!salesSheet) { Logger.log("Sales sheet not found!"); return; }
  const itemsSheet = getOrCreateSheet("Sales_Items", ["InvoiceID","Timestamp","Item Name","Qty","Unit Price","Total Price","Billed By"]);
  if (itemsSheet.getLastRow() > 1) itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, 7).clearContent();
  const data = salesSheet.getDataRange().getValues();
  let rowsProcessed = 0;
  for (let i = 1; i < data.length; i++) {
    const invoiceId = data[i][0], timestamp = data[i][1], itemsStr = String(data[i][2] || ""), billedBy = String(data[i][8] || "");
    if (!invoiceId) continue;
    if (itemsStr.trim().startsWith("[")) {
      try {
        JSON.parse(itemsStr).forEach(item => {
          const price = Number(item.price || item.originalPrice || 0), qty = Number(item.qty || 0);
          itemsSheet.appendRow([invoiceId, timestamp, item.name, qty, price, qty * price, billedBy]);
        });
        rowsProcessed++;
      } catch (e) {}
    } else {
      const regex = /(.*?)\s+\((\d+)x@₹([\d.]+)\)/g;
      let match, found = false;
      while ((match = regex.exec(itemsStr)) !== null) {
        const qty = Number(match[2]), price = Number(match[3]);
        itemsSheet.appendRow([invoiceId, timestamp, match[1].trim(), qty, price, qty * price, billedBy]);
        found = true;
      }
      if (found) rowsProcessed++;
    }
  }
  SpreadsheetApp.getUi().alert("✅ Done! " + rowsProcessed + " past orders migrated to Sales_Items.");
}