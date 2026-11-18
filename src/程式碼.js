/**
 * ====================================================================
 * 台灣精密機械 ERP 報價系統 (Google Apps Script 後端)
 * 專案名稱: qt-test01
 * 連結試算表 ID: 1wGCNpypD_n0s8NbLSSgtpQesO8BGHXNCRcp0HAbR75k
 * 安全規範遵循: .agent/rules/quote_safety.md
 * ====================================================================
 */

// 試算表 ID 與工作表名稱定義
const SPREADSHEET_ID = '1wGCNpypD_n0s8NbLSSgtpQesO8BGHXNCRcp0HAbR75k';

const SHEET_NAMES = {
  QUOTATIONS: 'Quotations',         // 報價單主檔
  QUOTATION_ITEMS: 'Quotation_Items', // 報價明細檔
  CUSTOMERS: 'Customers',           // 客戶資料庫
  INVENTORY: 'Inventory'            // 機台與加工件庫存表 (含底價防護)
};

// 規則 2: LINE 查詢合格業務白名單
const ALLOWED_SALES_IDS = [
  'U_SALES_01', // 林士傑
  'U_SALES_02', // 陳雅芳
  'U_DIRECTOR'  // 業務主管
];

/**
 * 取得試算表實例
 */
function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== '貼上你的_SCRIPT_ID') {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      console.warn('無法以 openById 開啟試算表，改用 getActiveSpreadsheet:', e);
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Web App 入口 (支援 Web 介面渲染與 Cloudflare API 請求)
 */
function doGet(e) {
  // 外部 API 讀取請求 (供 Cloudflare Pages / 前端呼叫)
  if (e && e.parameter && e.parameter.action) {
    let result = { success: false, error: '未知操作' };
    
    if (e.parameter.action === 'getInitialData') {
      result = getInitialData();
    } else if (e.parameter.action === 'getInventory') {
      // 規則 1: 庫存查詢必須剔除 cost_price
      result = getSafeInventory();
    } else if (e.parameter.action === 'getRecentQuotations') {
      const limit = parseInt(e.parameter.limit || '15', 10);
      result = getRecentQuotations(limit);
    } else if (e.parameter.action === 'getQuotationDetail') {
      result = getQuotationDetail(e.parameter.id);
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 預設呈現內建 Web App
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('台灣精密機械 ERP 報價系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Web App 寫入入口 (支援前端與 LINE Webhook POST 請求)
 */
function doPost(e) {
  try {
    let payload;
    if (e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      payload = e.parameter;
    }

    // 處理 LINE Webhook 訊息事件
    if (payload && payload.events) {
      return handleLineWebhook(payload);
    }

    // 處理報價單提交
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
 * 初始化試算表結構與範例資料 (自動建表)
 */
function initDatabase() {
  const ss = getSpreadsheet();

  // 1. 客戶資料庫
  let custSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMERS);
  if (!custSheet) {
    custSheet = ss.insertSheet(SHEET_NAMES.CUSTOMERS);
    custSheet.appendRow(['客戶代號', '客戶名稱', '統一編號', '對接窗口', '連絡電話', '電子信箱', '送貨廠區地址']);
    custSheet.getRange('A1:G1').setBackground('#1e293b').setFontColor('#ffffff').setFontWeight('bold');
    custSheet.appendRow(['CUST-001', '大同五軸精密機械工業股份有限公司', '11223344', '張協理', '04-2350-8888', 'sales@tatung-cnc.com.tw', '台中市西屯區工業區一路168號']);
    custSheet.appendRow(['CUST-002', '永進智動化科技股份有限公司', '55667788', '陳廠長', '04-2681-1234', 'factory@ycm-auto.com.tw', '台中市大甲區幼獅工業區幼九路']);
    custSheet.appendRow(['CUST-003', '亞威夾治具製造廠', '99887766', '林工程師', '03-598-7788', 'awea_fixture@gmail.com', '新竹縣湖口鄉新竹工業區中華路']);
  }

  // 2. 機台庫存與牌價表 (內部含 cost_price 底價，對外必須過濾)
  let invSheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  if (!invSheet) {
    invSheet = ss.insertSheet(SHEET_NAMES.INVENTORY);
    invSheet.appendRow(['料號', '品名/機台規格', '機構說明', '單位', '參考單價', '可用庫存', '進貨成本(cost_price)']);
    invSheet.getRange('A1:G1').setBackground('#0f766e').setFontColor('#ffffff').setFontWeight('bold');
    invSheet.appendRow(['MC-CNC-5X', 'CNC 高速立式五軸同動加工機', 'BT40 / 12,000rpm 直結主軸 / 附三軸光學尺', '台', 4200000, 3, 3100000]);
    invSheet.appendRow(['TL-ATC-30', '伺服自動換刀機構 (ATC 30T 刀庫)', 'BT40 刀臂式伺服換刀系統', '組', 380000, 8, 240000]);
    invSheet.appendRow(['FIX-HYD-01', '客製化四軸同動油壓夾治具', '高重複定位精度 ±0.003mm', '套', 260000, 5, 150000]);
    invSheet.appendRow(['SRV-TEST', '新機裝機定位與動態精度試車驗收', 'ISO 230-2 循圓球桿儀檢驗報告', '場', 45000, 99, 12000]);
  }

  // 3. 報價單主檔
  let qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
  if (!qSheet) {
    qSheet = ss.insertSheet(SHEET_NAMES.QUOTATIONS);
    qSheet.appendRow([
      '報價單編號', '報價日期', '有效日期', '客戶全銜', '統一編號', '對接窗口',
      '連絡電話', '送貨廠區', '幣別', '未稅金額', '營業稅(5%)', '含稅總額',
      '經辦業務', '狀態', '交易條款與備註', '開立時間'
    ]);
    qSheet.getRange('A1:P1').setBackground('#1e3a8a').setFontColor('#ffffff').setFontWeight('bold');
  }

  // 4. 報價明細檔
  let itemSheet = ss.getSheetByName(SHEET_NAMES.QUOTATION_ITEMS);
  if (!itemSheet) {
    itemSheet = ss.insertSheet(SHEET_NAMES.QUOTATION_ITEMS);
    itemSheet.appendRow([
      '明細編號', '報價單號', '項次', '機台/加工品名', '規格型號', '數量', '單位', '單價', '折扣(%)', '小計金額'
    ]);
    itemSheet.getRange('A1:J1').setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');
  }

  return '台灣機械 ERP 資料庫工作表初始化成功！';
}

/**
 * 規則 1: 庫存查詢必須嚴格剔除 cost_price 欄位，嚴禁洩漏底價
 */
function getSafeInventory() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.INVENTORY);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, list: [] };
  }

  // 讀取第 1 欄至第 6 欄 (完全避開第 7 欄 cost_price 進貨成本)
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const safeList = rows.map(r => ({
    id: r[0],
    name: r[1],
    spec: r[2],
    unit: r[3],
    price: r[4],
    stock: r[5]
    // 嚴格禁止包含 cost_price
  }));

  return { success: true, list: safeList };
}

/**
 * 規則 2: LINE 查詢驗證業務白名單 (ALLOWED_SALES_IDS)
 */
function verifySalesUser(userId) {
  if (!userId) return false;
  return ALLOWED_SALES_IDS.includes(userId);
}

/**
 * LINE Webhook 訊息分發處理器
 */
function handleLineWebhook(payload) {
  const events = payload.events || [];
  events.forEach(event => {
    if (event.type === 'message' && event.message.type === 'text') {
      const senderId = event.source.userId;
      const userMessage = event.message.text.trim();
      
      // 驗證業務權限
      if (!verifySalesUser(senderId)) {
        console.warn(`未授權存取嘗試: LINE ID ${senderId}`);
        // 拒絕回應或回覆未授權訊息
        return;
      }

      // 白名單合格業務，回覆機台庫存或報價資訊 (依然剔除底價)
      console.log(`業務 ${senderId} 查詢: ${userMessage}`);
    }
  });

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 前端載入基礎資料 (客戶清單、安全庫存品項、自動生成單號)
 */
function getInitialData() {
  const ss = getSpreadsheet();
  
  // 1. 客戶資料
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

  // 2. 機台庫存 (已過濾 cost_price)
  const inventoryRes = getSafeInventory();
  const products = inventoryRes.list;

  // 3. 自動生成報價單號
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
 * 自動依年月日生成報價單流水號 (QT-YYYYMMDD-001)
 */
function generateNextQuotationNumber() {
  const ss = getSpreadsheet();
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

  return prefix + String(maxSeq + 1).padStart(3, '0');
}

/**
 * 儲存報價單主檔與明細檔
 */
function submitQuotation(payload) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    const ss = getSpreadsheet();
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
      payload.validUntil || '',
      payload.customerName,
      payload.customerTaxId || '',
      payload.contactPerson,
      payload.customerPhone || '',
      payload.customerAddress || '',
      payload.currency || 'TWD',
      payload.subtotal,
      payload.tax,
      payload.total,
      payload.salesPerson || '',
      payload.status || '正式報價',
      payload.notes || '',
      nowStr
    ]);

    // 2. 寫入品項明細
    if (payload.items && payload.items.length > 0) {
      payload.items.forEach((item, index) => {
        const itemId = quotationNumber + '-' + (index + 1);
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
          item.subtotal
        ]);
      });
    }

    lock.releaseLock();

    return {
      success: true,
      quotationNumber: quotationNumber,
      message: '報價單已成功寫入試算表！'
    };
  } catch (err) {
    return {
      success: false,
      error: err.toString()
    };
  }
}

/**
 * 查詢最近開立的報價單總覽
 */
function getRecentQuotations(limit) {
  limit = limit || 15;
  const ss = getSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);

  if (!qSheet || qSheet.getLastRow() <= 1) {
    return { success: true, list: [] };
  }

  const lastRow = qSheet.getLastRow();
  const startRow = Math.max(2, lastRow - limit + 1);
  const numRows = lastRow - startRow + 1;

  const data = qSheet.getRange(startRow, 1, numRows, 14).getValues();
  data.reverse();

  const list = data.map(r => ({
    quotationNumber: r[0],
    date: Utilities.formatDate(new Date(r[1]), 'Asia/Taipei', 'yyyy-MM-dd'),
    customerName: r[3],
    contactPerson: r[5],
    total: r[11],
    salesPerson: r[12],
    status: r[13]
  }));

  return { success: true, list: list };
}

/**
 * 查詢特定報價單詳情 (供預覽或列印)
 */
function getQuotationDetail(quotationNumber) {
  const ss = getSpreadsheet();
  const qSheet = ss.getSheetByName(SHEET_NAMES.QUOTATIONS);
  const itemSheet = ss.getSheetByName(SHEET_NAMES.QUOTATION_ITEMS);

  if (!qSheet || !itemSheet) {
    return { success: false, error: '資料庫尚未初始化' };
  }

  const qValues = qSheet.getDataRange().getValues();
  let master = null;
  for (let i = 1; i < qValues.length; i++) {
    if (qValues[i][0] === quotationNumber) {
      master = {
        quotationNumber: qValues[i][0],
        quoteDate: Utilities.formatDate(new Date(qValues[i][1]), 'Asia/Taipei', 'yyyy-MM-dd'),
        validUntil: qValues[i][2] ? Utilities.formatDate(new Date(qValues[i][2]), 'Asia/Taipei', 'yyyy-MM-dd') : '',
        customerName: qValues[i][3],
        customerTaxId: qValues[i][4],
        contactPerson: qValues[i][5],
        customerPhone: qValues[i][6],
        customerAddress: qValues[i][7],
        currency: qValues[i][8],
        subtotal: qValues[i][9],
        tax: qValues[i][10],
        total: qValues[i][11],
        salesPerson: qValues[i][12],
        status: qValues[i][13],
        notes: qValues[i][14]
      };
      break;
    }
  }

  if (!master) {
    return { success: false, error: '找不到報價單編號：' + quotationNumber };
  }

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
        subtotal: itemValues[j][9]
      });
    }
  }

  return {
    success: true,
    master: master,
    items: items
  };
}
