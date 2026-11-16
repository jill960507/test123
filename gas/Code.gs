/**
 * 山餃 SAN JIAO｜線上訂單後端（Google Apps Script）
 *
 * 網站結帳時把訂單 POST 到這支程式部署出來的「網頁應用程式」網址，
 * 程式會依「菜單」工作表重新計算金額、產生訂單編號，寫入「訂單」與「訂單明細」，
 * 並寄送新訂單通知信。安裝步驟請看同資料夾的 README.md。
 */

const SHEET = {
  ORDERS: '訂單',
  ITEMS: '訂單明細',
  MENU: '菜單',
  SETTINGS: '設定',
  SUMMARY: '每日摘要'
};

const STATUSES = ['新訂單', '製作中', '可取餐', '已完成', '已取消'];
const FULFILLMENTS = ['外帶', '內用', '宅配'];

const ORDER_HEADERS = ['建立時間', '訂單編號', '狀態', '取餐方式', '姓名', '電話', '地址', '備註',
  '品項', '件數', '小計', '運費', '總金額', '狀態更新時間', '客戶端編號'];
const ITEM_HEADERS = ['建立時間', '訂單編號', '品項編號', '品名', '單價', '數量', '小計'];
const MENU_HEADERS = ['品項編號', '分類', '品名', '價格', '上架'];
const SUMMARY_HEADERS = ['日期', '訂單數', '取消數', '營業額', '外帶', '內用', '宅配', '熱銷品項'];

// 第一次執行 setup() 時寫入「菜單」的預設內容（與網站上的菜單相同）
const DEFAULT_MENU = [
  ['p1', '鍋貼', '經典鮮肉鍋貼', 75],
  ['p2', '鍋貼', '香辣川味鍋貼', 85],
  ['p3', '鍋貼', '鮮蝦韭黃鍋貼', 105],
  ['p4', '水餃', '手作高麗菜水餃', 80],
  ['p5', '水餃', '鮮蝦玉米水餃', 105],
  ['p6', '水餃', '香菇蔬菜水餃', 85],
  ['p7', '湯飲', '暖心酸辣湯', 45],
  ['p8', '湯飲', '清燉蘿蔔湯', 40],
  ['p9', '湯飲', '古早味紅茶', 30],
  ['p10', '小菜', '涼拌小黃瓜', 45],
  ['p11', '小菜', '秘製香辣豆干', 45],
  ['p12', '小菜', '芝麻涼拌時蔬', 50],
  ['gift50', '禮盒', '招牌招財鍋貼禮盒 50 入', 399]
];

// 「設定」工作表的項目：[名稱, 預設值, 說明]
const DEFAULT_SETTINGS = [
  ['接單中', '是', '填「否」可暫停網站接單（例如店休、備料不足）'],
  ['通知信箱', '', '新訂單通知寄到這裡，多個信箱用逗號分隔；留空則不寄'],
  ['宅配運費', 0, '宅配訂單加收的運費（元）'],
  ['免運門檻', 1000, '宅配小計達到這個金額就免運費（元）'],
  ['每日報表信箱', '', '每日營業摘要寄到這裡；留空則只寫入「每日摘要」工作表']
];

/* ───────────────────────── 網頁應用程式入口 ───────────────────────── */

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action !== 'createOrder') throw userError_('不支援的操作');
    return json_(createOrder_(body));
  } catch (err) {
    if (!err.userMessage) console.error(err && err.stack || err);
    return json_({ ok: false, error: err.userMessage || '系統忙碌，請稍後再試' });
  }
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'menu') {
    const menu = Object.values(readMenu_()).filter(m => m.active)
      .map(m => ({ id: m.id, category: m.category, name: m.name, price: m.price }));
    return json_({ ok: true, menu });
  }
  return json_({ ok: true, service: '山餃訂單 API', accepting: readSettings_().accepting });
}

/* ───────────────────────── 建立訂單 ───────────────────────── */

function createOrder_(body) {
  // 機器人陷阱：真人看不到這個欄位，有填就假裝成功但不寫入
  if (body.website) return { ok: true, orderNo: 'SJ-0', total: 0 };

  const order = validateOrder_(body);
  const settings = readSettings_();
  if (!settings.accepting) throw userError_('目前暫停線上接單，請稍後再試或直接來電');

  // 依「菜單」工作表重新計算，不採用前端送來的價格
  const menu = readMenu_();
  const lines = order.items.map(it => {
    const m = menu[it.id];
    if (!m || !m.active) throw userError_('部分餐點目前無法訂購，請重新整理頁面後再試');
    return { id: m.id, name: m.name, price: m.price, qty: it.qty, amount: m.price * it.qty };
  });
  const subtotal = lines.reduce((s, l) => s + l.amount, 0);
  const qty = lines.reduce((s, l) => s + l.qty, 0);
  const shipping = order.fulfillment === '宅配' && subtotal < settings.freeShipping ? settings.shippingFee : 0;
  const total = subtotal + shipping;

  const cache = CacheService.getScriptCache();
  const cacheKey = order.clientOrderId ? 'cid:' + order.clientOrderId : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let result;
  try {
    // 同一筆訂單重送（網路不穩、連點）時，回傳第一次的結果，不重複建單
    const previous = cacheKey && cache.get(cacheKey);
    if (previous) return JSON.parse(previous);

    const now = new Date();
    const orderNo = nextOrderNo_(now);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summary = lines.map(l => `${l.name} ×${l.qty}`).join('、');
    sheet_(ss, SHEET.ORDERS, ORDER_HEADERS).appendRow([
      now, orderNo, STATUSES[0], order.fulfillment, safe_(order.name), "'" + order.phone,
      safe_(order.address), safe_(order.note), summary, qty, subtotal, shipping, total, now, order.clientOrderId
    ]);
    const items = sheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
    items.getRange(items.getLastRow() + 1, 1, lines.length, ITEM_HEADERS.length)
      .setValues(lines.map(l => [now, orderNo, l.id, l.name, l.price, l.qty, l.amount]));
    SpreadsheetApp.flush();

    result = { ok: true, orderNo, subtotal, shipping, total, items: lines.map(l => ({ name: l.name, qty: l.qty, amount: l.amount })) };
    if (cacheKey) cache.put(cacheKey, JSON.stringify(result), 21600);
  } finally {
    lock.releaseLock();
  }

  try {
    notifyNewOrder_(settings, result, order);
  } catch (err) {
    console.error('通知信寄送失敗', err); // 寄信失敗不影響訂單
  }
  return result;
}

function validateOrder_(body) {
  const str = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const order = {
    fulfillment: str(body.fulfillment, 10),
    name: str(body.name, 50),
    phone: str(body.phone, 20),
    address: str(body.address, 200),
    note: str(body.note, 200),
    clientOrderId: str(body.clientOrderId, 64).replace(/[^\w-]/g, '')
  };
  if (FULFILLMENTS.indexOf(order.fulfillment) < 0) throw userError_('請選擇取餐方式');
  if (!order.name) throw userError_('請填寫聯絡姓名');
  if (!/^[0-9+() -]{8,20}$/.test(order.phone)) throw userError_('聯絡電話格式不正確');
  if (order.fulfillment === '宅配' && order.address.length < 5) throw userError_('請填寫完整的宅配地址');
  if (order.fulfillment !== '宅配') order.address = '';

  if (!Array.isArray(body.items) || !body.items.length) throw userError_('購物車是空的');
  if (body.items.length > 50) throw userError_('品項過多，請分開下單');
  const merged = {};
  body.items.forEach(it => {
    const id = str(it && it.id, 32);
    const qty = Number(it && it.qty);
    if (!id || !Number.isInteger(qty) || qty < 1 || qty > 99) throw userError_('餐點數量不正確');
    merged[id] = Math.min(99, (merged[id] || 0) + qty);
  });
  order.items = Object.keys(merged).map(id => ({ id, qty: merged[id] }));
  return order;
}

// 訂單編號：SJ-20261116-001，每天從 001 重新開始
function nextOrderNo_(now) {
  const day = Utilities.formatDate(now, tz_(), 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = 'seq-' + day;
  const seq = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(seq));
  return `SJ-${day}-${String(seq).padStart(3, '0')}`;
}

function notifyNewOrder_(settings, result, order) {
  if (!settings.notifyEmail) return;
  const rows = result.items.map(i => `・${i.name} ×${i.qty}　NT$ ${i.amount}`).join('\n');
  const text = [
    `訂單編號：${result.orderNo}`,
    `取餐方式：${order.fulfillment}`,
    `姓名：${order.name}`,
    `電話：${order.phone}`,
    order.address ? `地址：${order.address}` : '',
    order.note ? `備註：${order.note}` : '',
    '',
    rows,
    '',
    `小計：NT$ ${result.subtotal}`,
    result.shipping ? `運費：NT$ ${result.shipping}` : '',
    `總金額：NT$ ${result.total}`,
    '',
    `開啟訂單表：${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`
  ].filter(line => line !== '').join('\n');
  MailApp.sendEmail({
    to: settings.notifyEmail,
    subject: `【山餃新訂單】${result.orderNo}｜${order.fulfillment}｜NT$ ${result.total}`,
    body: text,
    name: '山餃線上訂單'
  });
}

/* ───────────────────────── 工作表內的操作 ───────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('山餃訂單')
    .addItem('初始化／修復工作表', 'setup')
    .addItem('產生今日營業摘要', 'dailySummary')
    .addItem('寄送測試通知信', 'sendTestEmail')
    .addSeparator()
    .addItem('開啟每晚自動報表（21:00）', 'installDailyTrigger')
    .addItem('關閉每晚自動報表', 'removeDailyTrigger')
    .addToUi();
}

// 在「訂單」工作表手動改狀態時，自動記錄狀態更新時間
function onEdit(e) {
  const range = e && e.range;
  if (!range || range.getSheet().getName() !== SHEET.ORDERS) return;
  const statusCol = ORDER_HEADERS.indexOf('狀態') + 1;
  if (range.getColumn() > statusCol || range.getLastColumn() < statusCol || range.getRow() < 2) return;
  const updatedCol = ORDER_HEADERS.indexOf('狀態更新時間') + 1;
  range.getSheet().getRange(range.getRow(), updatedCol, range.getNumRows(), 1)
    .setValues(Array.from({ length: range.getNumRows() }, () => [new Date()]));
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Taipei');

  const orders = sheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  const col = name => ORDER_HEADERS.indexOf(name) + 1;
  const body = (c) => orders.getRange(2, c, orders.getMaxRows() - 1, 1);
  body(col('建立時間')).setNumberFormat('yyyy/mm/dd hh:mm');
  body(col('狀態更新時間')).setNumberFormat('yyyy/mm/dd hh:mm');
  body(col('電話')).setNumberFormat('@');
  ['小計', '運費', '總金額'].forEach(n => body(col(n)).setNumberFormat('"NT$ "#,##0'));
  body(col('狀態')).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUSES, true).setAllowInvalid(false).build());
  const colors = { '新訂單': '#fde2c8', '製作中': '#fff2b3', '可取餐': '#d6ecd2', '已完成': '#e6e6e6', '已取消': '#f4c7c3' };
  const statusRange = body(col('狀態'));
  orders.setConditionalFormatRules(Object.keys(colors).map(s => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(s).setBackground(colors[s]).setRanges([statusRange]).build()));
  orders.setColumnWidth(col('品項'), 280);
  orders.setColumnWidth(col('地址'), 220);
  orders.hideColumns(col('客戶端編號'));

  const items = sheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
  items.getRange(2, 1, items.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');

  const menu = sheet_(ss, SHEET.MENU, MENU_HEADERS);
  if (menu.getLastRow() < 2) {
    menu.getRange(2, 1, DEFAULT_MENU.length, 4).setValues(DEFAULT_MENU);
    menu.getRange(2, 5, DEFAULT_MENU.length, 1).insertCheckboxes().check();
  }
  menu.getRange(2, 4, menu.getMaxRows() - 1, 1).setNumberFormat('#,##0');

  const settings = sheet_(ss, SHEET.SETTINGS, ['項目', '值', '說明']);
  const existing = settings.getLastRow() > 1
    ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
  const missing = DEFAULT_SETTINGS.filter(s => existing.indexOf(s[0]) < 0).map(s => s.slice());
  missing.forEach(s => { if (s[0] === '通知信箱' || s[0] === '每日報表信箱') s[1] = Session.getEffectiveUser().getEmail(); });
  if (missing.length) settings.getRange(settings.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
  settings.setColumnWidth(2, 240);
  settings.setColumnWidth(3, 420);

  sheet_(ss, SHEET.SUMMARY, SUMMARY_HEADERS);
  ss.setActiveSheet(orders);
  toast_('工作表已準備完成');
}

/* ───────────────────────── 每日營業摘要 ───────────────────────── */

function dailySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd');
  const orders = sheet_(ss, SHEET.ORDERS, ORDER_HEADERS);
  const rows = orders.getLastRow() > 1 ? orders.getRange(2, 1, orders.getLastRow() - 1, ORDER_HEADERS.length).getValues() : [];
  const c = name => ORDER_HEADERS.indexOf(name);
  const todays = rows.filter(r => r[c('建立時間')] instanceof Date && Utilities.formatDate(r[c('建立時間')], tz_(), 'yyyy/MM/dd') === today);
  const valid = todays.filter(r => r[c('狀態')] !== '已取消');
  const cancelledNos = todays.filter(r => r[c('狀態')] === '已取消').map(r => r[c('訂單編號')]);
  const revenue = valid.reduce((s, r) => s + Number(r[c('總金額')] || 0), 0);
  const byType = t => valid.filter(r => r[c('取餐方式')] === t).length;

  const items = sheet_(ss, SHEET.ITEMS, ITEM_HEADERS);
  const itemRows = items.getLastRow() > 1 ? items.getRange(2, 1, items.getLastRow() - 1, ITEM_HEADERS.length).getValues() : [];
  const sold = {};
  itemRows.forEach(r => {
    if (!(r[0] instanceof Date) || Utilities.formatDate(r[0], tz_(), 'yyyy/MM/dd') !== today) return;
    if (cancelledNos.indexOf(r[1]) >= 0) return;
    sold[r[3]] = (sold[r[3]] || 0) + Number(r[5] || 0);
  });
  const top = Object.keys(sold).sort((a, b) => sold[b] - sold[a]).slice(0, 3).map(n => `${n} ×${sold[n]}`).join('、') || '—';

  const row = [today, valid.length, cancelledNos.length, revenue, byType('外帶'), byType('內用'), byType('宅配'), top];
  const summary = sheet_(ss, SHEET.SUMMARY, SUMMARY_HEADERS);
  const dates = summary.getLastRow() > 1 ? summary.getRange(2, 1, summary.getLastRow() - 1, 1).getDisplayValues().map(r => r[0]) : [];
  const at = dates.indexOf(today);
  if (at >= 0) summary.getRange(at + 2, 1, 1, row.length).setValues([row]);
  else summary.appendRow(row);
  summary.getRange(2, 1, Math.max(summary.getLastRow() - 1, 1), 1).setNumberFormat('yyyy/mm/dd');
  summary.getRange(2, 4, Math.max(summary.getLastRow() - 1, 1), 1).setNumberFormat('"NT$ "#,##0');

  const to = readSettings_().reportEmail;
  if (to) {
    MailApp.sendEmail({
      to,
      subject: `【山餃營業摘要】${today}｜${valid.length} 筆｜NT$ ${revenue}`,
      body: `日期：${today}\n訂單數：${valid.length}（取消 ${cancelledNos.length}）\n營業額：NT$ ${revenue}\n外帶 ${row[4]}／內用 ${row[5]}／宅配 ${row[6]}\n熱銷：${top}\n\n${ss.getUrl()}`,
      name: '山餃線上訂單'
    });
  }
  toast_(`今日 ${valid.length} 筆訂單，營業額 NT$ ${revenue}`);
}

function installDailyTrigger() {
  removeDailyTrigger();
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(21).inTimezone(tz_()).create();
  toast_('已開啟：每晚 21:00 自動產生營業摘要');
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailySummary')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function sendTestEmail() {
  const settings = readSettings_();
  if (!settings.notifyEmail) return toast_('請先在「設定」工作表填寫通知信箱');
  notifyNewOrder_(settings, {
    orderNo: 'SJ-TEST-000', subtotal: 150, shipping: 0, total: 150,
    items: [{ name: '經典鮮肉鍋貼（測試）', qty: 2, amount: 150 }]
  }, { fulfillment: '外帶', name: '測試顧客', phone: '0912-345-678', address: '', note: '這是一封測試信' });
  toast_('測試通知信已寄出：' + settings.notifyEmail);
}

/* ───────────────────────── 工具函式 ───────────────────────── */

function readMenu_() {
  const sheet = sheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET.MENU, MENU_HEADERS);
  const menu = {};
  if (sheet.getLastRow() < 2) return menu;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, MENU_HEADERS.length).getValues().forEach(r => {
    const id = String(r[0]).trim();
    const price = Number(r[3]);
    if (!id || !(price >= 0)) return;
    menu[id] = { id, category: String(r[1]), name: String(r[2]), price, active: r[4] === true || r[4] === '是' };
  });
  return menu;
}

function readSettings_() {
  const sheet = sheet_(SpreadsheetApp.getActiveSpreadsheet(), SHEET.SETTINGS, ['項目', '值', '說明']);
  const map = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => { map[String(r[0]).trim()] = r[1]; });
  }
  const num = (v, d) => (v === '' || v == null || isNaN(Number(v)) ? d : Number(v));
  return {
    accepting: String(map['接單中'] == null ? '是' : map['接單中']).trim() !== '否',
    notifyEmail: String(map['通知信箱'] || '').trim(),
    shippingFee: num(map['宅配運費'], 0),
    freeShipping: num(map['免運門檻'], Infinity),
    reportEmail: String(map['每日報表信箱'] || '').trim()
  };
}

function sheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#34251d').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 避免顧客輸入的文字被試算表當成公式執行
function safe_(v) {
  const s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function userError_(message) {
  const err = new Error(message);
  err.userMessage = message;
  return err;
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function tz_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Taipei';
}

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, '山餃訂單', 5); } catch (e) {}
}
