/**
 * ====================================================================
 * ERP 報價系統 (Google Apps Script 後端)
 * 專案代號: qt-test01
 * ====================================================================
 */

const SHEET_NAMES = {
  QUOTATIONS: 'Quotations',
  QUOTATION_ITEMS: 'Quotation_Items',
  CUSTOMERS: 'Customers',
  PRODUCTS: 'Products'
};

/**
 * 網頁進入點 (GET)
 * 支援：
 * 1. 網頁 UI 渲染
 * 2. API 請求 (供 Cloudflare Pages / 外部前端 fetch)
 */
function doGet(e) {
  // 支援外部 API 呼叫 (REST API)
  if (e && e.parameter && e.parameter.action) {
    let result = { success: false, error: '未知操作' };
    if (e.parameter.action === 'getInitialData') {
      result = getInitialData();
    } else if (e.parameter.action === 'getRecentQuotations') {
      const limit = parseInt(e.parameter.limit || '15', 10);
      result = getRecentQuotations(limit);
    } else if (e.parameter.action === 'getQuotationDetail') {
      result = getQuotationDetail(e.parameter.id);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 如果網址帶有 print 參數，提供專門列印頁面
  if (e && e.parameter && e.parameter.print) {
    const template = HtmlService.createTemplateFromFile('print');
    template.quotationId = e.parameter.print;
    return template.evaluate()
      .setTitle('列印報價單 - ' + e.parameter.print)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 預設呈現報價單系統主介面
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ERP 報價系統 - QT TEST01')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * API 寫入端點 (POST)
 * 供 Cloudflare Pages / 外部表單跨域提交
 */
function doPost(e) {
  try {
    let payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      payload = e.parameter;
    }

    let result = { success: false, error: '無效資料' };
    if (payload) {
      result = submitQuotation(payload);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 初始化試算表結構與範例資料
 * 請在 Apps Script 編輯器中手動執行一次本函式
 */
function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 客戶資料表
  let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!custSheet) {
    custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
    custSheet.appendRow(['客戶代號', '客戶名稱', '統一編號', '聯絡人', '電話', 'Email', '地址']);
    custSheet.getRange('A1:G1').setBackground('#4f46e5').setFontColor('#ffffff').setFontWeight('bold');
    // 預設範例資料
    custSheet.appendRow(['CUST001', '宏達數位科技有限公司', '12345678', '陳經理', '02-2345-6789', 'contact@example.com', '台北市中山區南京東路二段']);
    custSheet.appendRow(['CUST002', '頂尖系統整合顧問', '87654321', '林專案', '04-8765-4321', 'sales@topconsult.com', '台中市西區台灣大道二段']);
  }

  // 2. 產品品項表
  let prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  if (!prodSheet) {
    prodSheet = ss.insertSheet(SHEET_NAMES.PRODUCTS);
    prodSheet.appendRow(['產品編號', '品名/服務項目', '規格說明', '單位', '參考單價', '備註']);
    prodSheet.getRange('A1:F1').setBackground('#0284c7').setFontColor('#ffffff').setFontWeight('bold');
    // 預設範例資料
    prodSheet.appendRow(['PROD001', 'ERP 系統模組授權', '雲端標準版 (年租)', '套', 36000, '含 5 位使用者']);
    prodSheet.appendRow(['PROD002', '系統客製化開發與串接', '每工時報價', '人/天', 16000, '客製流程與報表']);
    prodSheet.appendRow(['PROD003', '企業教育訓練', '半天實體培訓 (4小時)', '場', 12000, '附教材與影音存檔']);
  }

  // 3. 報價單主檔
  let qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
  if (!qSheet) {
    qSheet = ss.insertSheet(SHEET_NAMES.QUOTATIONS);
    qSheet.appendRow([
      '報價單號', '報價日期', '有效日期', '客戶代號', '客戶名稱', '聯絡人', 
      '幣別', '未稅金額', '營業稅(5%)', '含稅總額', '業務經辦', '狀態', '付款條件/備註', '建立時間'
    ]);
    qSheet.getRange('A1:N1').setBackground('#16a34a').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 4. 報價單明細檔
  let itemSheet = ss.getSheetByName(SHEET_NAMES.QUOTATION_ITEMS);
  if (!itemSheet) {
    itemSheet = ss.insertSheet(SHEET_NAMES.QUOTATION_ITEMS);
    itemSheet.appendRow([
      '明細ID', '報價單號', '項次', '品名項目', '規格說明', '數量', '單位', '單價', '折扣(%)', '小計', '備註'
    ]);
    itemSheet.getRange('A1:K1').setBackground('#0d9488').setFontColor('#ffffff').setFontWeight('bold');
  }

  return '資料庫工作表初始化成功！';
}

/**
 * 前端初始化載入所需資料 (客戶、品項清單、新單號)
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 取得客戶列表
  const custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  let customers = [];
  if (custSheet && custSheet.getLastRow() > 1) {
    const custValues = custSheet.getRange(2, 1, custSheet.getLastRow() - 1, 7).getValues();
    customers = custValues.map(row => ({
      id: row[0],
      name: row[1],
      taxId: row[2],
      contact: row[3],
      phone: row[4],
      email: row[5],
      address: row[6]
    }));
  }

  // 取得產品列表
  const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCTS);
  let products = [];
  if (prodSheet && prodSheet.getLastRow() > 1) {
    const prodValues = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 6).getValues();
    products = prodValues.map(row => ({
      id: row[0],
      name: row[1],
      spec: row[2],
      unit: row[3],
      price: row[4],
      remark: row[5]
    }));
  }

  // 產生新報價單號
  const nextNumber = generateNextQuotationNumber();

  return {
    success: true,
    data: {
      customers: customers,
      products: products,
      nextQuotationNumber: nextNumber,
      today: Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd')
    }
  };
}

/**
 * 自動產生報價單號：QT-YYYYMMDD-001
 */
function generateNextQuotationNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
  const prefix = 'QT-' + todayStr + '-';

  if (!qSheet || qSheet.getLastRow() <= 1) {
    return prefix + '001';
  }

  const ids = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, 1).getValues().flat();
  let maxSeq = 0;
  ids.forEach(id => {
    if (typeof id === 'string' && id.startsWith(prefix)) {
      const seq = parseInt(id.replace(prefix, ''), 10);
      if (!isNaN(seq) && seq > maxSeq) {
        maxSeq = seq;
      }
    }
  });

  const nextSeq = String(maxSeq + 1).padStart(3, '0');
  return prefix + nextSeq;
}

/**
 * 儲存報價單 (前端調用)
 */
function submitQuotation(payload) {
  try {
    const lock = LockService.getScriptLock();
    // 等候最長 30 秒鎖定，避免多工衝突
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
    let itemSheet = ss.getSheetByName(SHEET_NAMES.QUOTATION_ITEMS);

    if (!qSheet || !itemSheet) {
      initDatabase();
      qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
      itemSheet = ss.getSheetByName(SHEET_NAMES.QUOTATION_ITEMS);
    }

    const quotationNumber = payload.quotationNumber || generateNextQuotationNumber();
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');

    // 1. 寫入報價主檔
    qSheet.appendRow([
      quotationNumber,
      payload.quoteDate,
      payload.validUntil,
      payload.customerId || '',
      payload.customerName,
      payload.contactPerson,
      payload.currency || 'TWD',
      payload.subtotal,
      payload.tax,
      payload.total,
      payload.salesPerson || '',
      payload.status || '草稿',
      payload.notes || '',
      nowStr
    ]);

    // 2. 寫入明細檔
    if (payload.items && payload.items.length > 0) {
      payload.items.forEach((item, index) => {
        const itemId = quotationNumber + '-I' + (index + 1);
        itemSheet.appendRow([
          itemId,
          quotationNumber,
          index + 1,
          item.name,
          item.spec || '',
          item.qty,
          item.unit || '',
          item.price,
          item.discount || 0,
          item.subtotal,
          item.remark || ''
        ]);
      });
    }

    lock.releaseLock();

    return {
      success: true,
      quotationNumber: quotationNumber,
      message: '報價單已成功建立儲存！'
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * 取得最近報價單紀錄
 */
function getRecentQuotations(limit) {
  limit = limit || 15;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);

  if (!qSheet || qSheet.getLastRow() <= 1) {
    return { success: true, list: [] };
  }

  const lastRow = qSheet.getLastRow();
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;

  const data = qSheet.getRange(startRow, 1, numRows, 14).getValues();
  // 反轉排序，最新的在前面
  data.reverse();

  const list = data.map(r => ({
    quotationNumber: r[0],
    date: Utilities.formatDate(new Date(r[1]), 'Asia/Taipei', 'yyyy-MM-dd'),
    customerName: r[4],
    contactPerson: r[5],
    total: r[9],
    salesPerson: r[10],
    status: r[11]
  }));

  return { success: true, list: list };
}

/**
 * 取得單張報價單詳情 (供預覽或列印)
 */
function getQuotationDetail(quotationNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
  const itemSheet = ss.getSheetByName(SHEET_NAMES.QUOTATION_ITEMS);

  if (!qSheet || !itemSheet) {
    return { success: false, error: '資料表尚未初始化' };
  }

  // 尋找主檔
  const qValues = qSheet.getDataRange().getValues();
  let master = null;
  for (let i = 1; i < qValues.length; i++) {
    if (qValues[i][0] === quotationNumber) {
      master = {
        quotationNumber: qValues[i][0],
        quoteDate: Utilities.formatDate(new Date(qValues[i][1]), 'Asia/Taipei', 'yyyy-MM-dd'),
        validUntil: Utilities.formatDate(new Date(qValues[i][2]), 'Asia/Taipei', 'yyyy-MM-dd'),
        customerId: qValues[i][3],
        customerName: qValues[i][4],
        contactPerson: qValues[i][5],
        currency: qValues[i][6],
        subtotal: qValues[i][7],
        tax: qValues[i][8],
        total: qValues[i][9],
        salesPerson: qValues[i][10],
        status: qValues[i][11],
        notes: qValues[i][12],
        createdAt: qValues[i][13]
      };
      break;
    }
  }

  if (!master) {
    return { success: false, error: '找不到報價單：' + quotationNumber };
  }

  // 尋找明細檔
  const itemValues = itemSheet.getDataRange().getValues();
  const items = [];
  for (let j = 1; j < itemValues.length; j++) {
    if (itemValues[j][1] === quotationNumber) {
      items.push({
        seq: itemValues[j][2],
        name: itemValues[j][3],
        spec: itemValues[j][4],
        qty: itemValues[j][5],
        unit: itemValues[j][6],
        price: itemValues[j][7],
        discount: itemValues[j][8],
        subtotal: itemValues[j][9],
        remark: itemValues[j][10]
      });
    }
  }

  return {
    success: true,
    master: master,
    items: items
  };
}
