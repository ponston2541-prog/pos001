// ใส่ ID ของ Google Sheet
const SPREADSHEET_ID = "1L19bgi_elq4cT-_eVPVc8Iymypm29CFbniQSv_t7F5I";

function getSS() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ตรวจสอบและดึง Sheet ถ้าไม่มีให้สร้างให้อัตโนมัติ
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

// รองรับ GET Request
function doGet(e) {
  const action = e ? e.parameter.action : null;
  let result = {};

  if (action === "getProducts") {
    result = getProducts();
  } else if (action === "getDailySalesSummary") {
    const date = e.parameter.date;
    result = getDailySalesSummary(date);
  } else {
    result = { status: "error", message: "Invalid action" };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// รองรับ POST Request
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result = {};

    if (action === "addProduct") {
      result = addProduct(data.payload);
    } else if (action === "processSale") {
      result = processSale(data.payload);
    } else {
      result = { status: "error", message: "Invalid action" };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------- ฟังก์ชันประมวลผล -------------------

function getProducts() {
  try {
    const ss = getSS();
    const sheet = getOrCreateSheet(ss, "Products", ["Product_ID", "Name", "Category", "Price", "Stock", "Barcode"]);
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const products = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] !== "" && row[0] !== null) {
        products.push({
          Product_ID: String(row[0]),
          Name: String(row[1] || ''),
          Category: String(row[2] || 'ทั่วไป'),
          Price: Number(row[3]) || 0,
          Stock: Number(row[4]) || 0,
          Barcode: String(row[5] || '')
        });
      }
    }
    return products;
  } catch (e) {
    return [];
  }
}

function addProduct(p) {
  try {
    const ss = getSS();
    const sheet = getOrCreateSheet(ss, "Products", ["Product_ID", "Name", "Category", "Price", "Stock", "Barcode"]);

    const id = "P" + new Date().getTime().toString().slice(-6);
    sheet.appendRow([id, p.name, p.category, Number(p.price), Number(p.stock), p.barcode]);
    return { success: true, message: "เพิ่มสินค้าเรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

function processSale(payload) {
  try {
    const ss = getSS();
    const salesSheet = getOrCreateSheet(ss, "Sales", ["Sale_ID", "Date", "TotalAmount", "Discount", "FinalAmount", "PaymentMethod"]);
    // เพิ่มคอลัมน์ Name ลงในตาราง SaleDetails
    const detailsSheet = getOrCreateSheet(ss, "SaleDetails", ["Sale_ID", "Product_ID", "Name", "Quantity", "Subtotal"]);
    const productsSheet = getOrCreateSheet(ss, "Products", ["Product_ID", "Name", "Category", "Price", "Stock", "Barcode"]);

    const saleId = "INV" + new Date().getTime().toString().slice(-6);
    const now = new Date();

    salesSheet.appendRow([
      saleId,
      now,
      payload.totalAmount,
      payload.discount,
      payload.finalAmount,
      payload.paymentMethod === 'TRANSFER' ? 'โอนเงิน/QR' : 'เงินสด'
    ]);

    const pData = productsSheet.getDataRange().getValues();

    payload.cart.forEach(item => {
      // บันทึก Name เพิ่มลงในคอลัมน์ที่ 3
      const pId = item.Product_ID || item.id;
      const pName = item.Name || item.name || '';
      detailsSheet.appendRow([saleId, pId, pName, item.Quantity, item.Subtotal]);

      for (let i = 1; i < pData.length; i++) {
        if (String(pData[i][0]) === String(pId)) {
          const currentStock = Number(pData[i][4]) || 0;
          const newStock = Math.max(0, currentStock - item.Quantity);
          productsSheet.getRange(i + 1, 5).setValue(newStock);
          break;
        }
      }
    });

    return { success: true, saleId: saleId };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getDailySalesSummary(targetDateStr) {
  try {
    const ss = getSS();
    const salesSheet = getOrCreateSheet(ss, "Sales", ["Sale_ID", "Date", "TotalAmount", "Discount", "FinalAmount", "PaymentMethod"]);
    const detailsSheet = getOrCreateSheet(ss, "SaleDetails", ["Sale_ID", "Product_ID", "Name", "Quantity", "Subtotal"]);
    const productsSheet = getOrCreateSheet(ss, "Products", ["Product_ID", "Name", "Category", "Price", "Stock", "Barcode"]);

    let totalSales = 0, totalDiscount = 0, netRevenue = 0, totalOrders = 0;
    const matchedSaleIds = [];
    const recentSales = [];

    if (salesSheet && salesSheet.getLastRow() > 1) {
      const salesData = salesSheet.getDataRange().getValues();
      const tz = ss.getSpreadsheetTimeZone();

      for (let i = 1; i < salesData.length; i++) {
        const row = salesData[i];
        if (!row[1]) continue;
        
        const saleDate = new Date(row[1]);
        const formattedDate = Utilities.formatDate(saleDate, tz, "yyyy-MM-dd");

        if (formattedDate === targetDateStr) {
          matchedSaleIds.push(String(row[0]));
          totalSales += Number(row[2]) || 0;
          totalDiscount += Number(row[3]) || 0;
          netRevenue += Number(row[4]) || 0;
          totalOrders++;

          const timeStr = Utilities.formatDate(saleDate, tz, "HH:mm");
          recentSales.push({
            saleId: String(row[0]),
            time: timeStr,
            finalAmount: Number(row[4]) || 0,
            paymentMethod: String(row[5] || 'เงินสด')
          });
        }
      }
    }

    const prodCountMap = {};
    if (matchedSaleIds.length > 0 && detailsSheet && detailsSheet.getLastRow() > 1) {
      const detailsData = detailsSheet.getDataRange().getValues();
      for (let i = 1; i < detailsData.length; i++) {
        const row = detailsData[i];
        const sId = String(row[0]);
        const pId = String(row[1]);
        // รองรับทั้งโครงสร้างเก่าและโครงสร้างใหม่ที่แทรกคอลัมน์ Name
        const qty = Number(row[3]) || Number(row[2]) || 0;

        if (matchedSaleIds.includes(sId)) {
          prodCountMap[pId] = (prodCountMap[pId] || 0) + qty;
        }
      }
    }

    const prodNameMap = {};
    if (productsSheet && productsSheet.getLastRow() > 1) {
      const prodData = productsSheet.getDataRange().getValues();
      for (let i = 1; i < prodData.length; i++) {
        prodNameMap[String(prodData[i][0])] = String(prodData[i][1]);
      }
    }

    const topProducts = Object.keys(prodCountMap).map(pId => ({
      name: prodNameMap[pId] || pId,
      quantity: prodCountMap[pId]
    })).sort((a, b) => b.quantity - a.quantity).slice(0, 5);

    return {
      netRevenue,
      totalSales,
      totalDiscount,
      totalOrders,
      topProducts,
      recentSales: recentSales.reverse()
    };
  } catch(e) {
    return { netRevenue: 0, totalSales: 0, totalDiscount: 0, totalOrders: 0, topProducts: [], recentSales: [] };
  }
}
