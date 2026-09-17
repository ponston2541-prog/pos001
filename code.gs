// ใส่ ID ของ Google Sheet ที่ดึงมาจาก URL เรียบร้อยแล้ว
const SPREADSHEET_ID = "1L19bgi_elq4cT-_eVPVc8Iymypm29CFbniQSv_t7F5I"; 

function getSS() {
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Smart POS System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. ดึงข้อมูลสินค้าทั้งหมด
function getProducts() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName("Products");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; // มีแค่ Header
    
    const products = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] !== "" && row[0] !== null && row[0] !== undefined) {
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
    Logger.log("Error in getProducts: " + e.toString());
    throw new Error("ไม่สามารถดึงข้อมูลสินค้าได้: " + e.toString());
  }
}

// 2. บันทึกสินค้าใหม่
function addProduct(p) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName("Products");
    if (!sheet) throw new Error("ไม่พบแผ่นงานชื่อ 'Products'");

    const id = "P" + new Date().getTime().toString().slice(-6);
    sheet.appendRow([id, p.name, p.category, Number(p.price), Number(p.stock), p.barcode]);
    return { success: true, message: "เพิ่มสินค้าเรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + e.toString() };
  }
}

// 3. บันทึกการขาย
function processSale(payload) {
  try {
    const ss = getSS();
    const salesSheet = ss.getSheetByName("Sales");
    const detailsSheet = ss.getSheetByName("SaleDetails");
    const productsSheet = ss.getSheetByName("Products");

    if (!salesSheet || !detailsSheet || !productsSheet) {
      throw new Error("แผ่นงาน Sales, SaleDetails หรือ Products ไม่ครบถ้วน");
    }

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
      detailsSheet.appendRow([saleId, item.Product_ID, item.Quantity, item.Subtotal]);

      for (let i = 1; i < pData.length; i++) {
        if (String(pData[i][0]) === String(item.Product_ID)) {
          const currentStock = Number(pData[i][4]) || 0;
          const newStock = Math.max(0, currentStock - item.Quantity);
          productsSheet.getRange(i + 1, 5).setValue(newStock);
          break;
        }
      }
    });

    return { success: true, saleId: saleId };
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดในการบันทึกขาย: " + e.toString());
  }
}

// 4. ดึงข้อมูลสรุปยอดขายประจำวัน
function getDailySalesSummary(targetDateStr) {
  try {
    const ss = getSS();
    const salesSheet = ss.getSheetByName("Sales");
    const detailsSheet = ss.getSheetByName("SaleDetails");
    const productsSheet = ss.getSheetByName("Products");

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
        const qty = Number(row[2]) || 0;

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