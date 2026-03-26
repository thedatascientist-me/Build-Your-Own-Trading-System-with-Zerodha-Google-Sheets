function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Alladin')
    .addItem('Initialize orders sheets', 'initializeOrdersSheets')
    .addItem('Refresh orders', 'refreshOrders')
    .addItem('Build orders dashboard', 'buildOrdersDashboard')
    .addToUi();
}

function initializeOrdersSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Config', 'Orders Raw', 'Orders Dashboard', 'Logs'];

  requiredSheets.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
  });

  ensureConfigRow_('api_key', '');
  ensureConfigRow_('access_token', '');
  ensureConfigRow_('last_orders_refresh', '');
  ensureConfigRow_('status', '');

  logMessage_('INFO', 'Orders sheets initialized');
  SpreadsheetApp.getUi().alert('Orders sheets initialized successfully.');
}

function refreshOrders() {
  try {
    const config = getConfig_();
    if (!config.api_key || !config.access_token) {
      throw new Error('Missing api_key or access_token in Config sheet.');
    }

    const response = UrlFetchApp.fetch('https://api.kite.trade/orders', {
      method: 'get',
      headers: {
        'X-Kite-Version': '3',
        'Authorization': 'token ' + config.api_key + ':' + config.access_token
      },
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const text = response.getContentText();
    if (code !== 200) {
      throw new Error('API error ' + code + ': ' + text);
    }

    const payload = JSON.parse(text);
    if (payload.status !== 'success' || !payload.data) {
      throw new Error('Unexpected API response: ' + text);
    }

    writeOrdersRaw_(payload.data);
    updateConfigValue_('last_orders_refresh', new Date());
    updateConfigValue_('status', 'Orders refreshed successfully');
    buildOrdersDashboard();

    logMessage_('INFO', 'Orders refreshed successfully. Rows: ' + payload.data.length);
    SpreadsheetApp.getUi().alert('Orders refreshed successfully.');
  } catch (err) {
    updateConfigValue_('status', 'Error: ' + err.message);
    logMessage_('ERROR', err.message);
    SpreadsheetApp.getUi().alert('Refresh failed: ' + err.message);
  }
}

function writeOrdersRaw_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Orders Raw');
  sh.clear();

  const headers = [
    'order_id', 'status', 'order_timestamp', 'exchange', 'tradingsymbol',
    'transaction_type', 'product', 'order_type', 'quantity',
    'filled_quantity', 'pending_quantity', 'cancelled_quantity',
    'price', 'average_price', 'validity', 'variety', 'tag', 'status_message'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = data.map(item => [
    item.order_id || '',
    item.status || '',
    item.order_timestamp || '',
    item.exchange || '',
    item.tradingsymbol || '',
    item.transaction_type || '',
    item.product || '',
    item.order_type || '',
    item.quantity || 0,
    item.filled_quantity || 0,
    item.pending_quantity || 0,
    item.cancelled_quantity || 0,
    item.price || 0,
    item.average_price || 0,
    item.validity || '',
    item.variety || '',
    item.tag || '',
    item.status_message || ''
  ]);

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.autoResizeColumns(1, headers.length);
}

function buildOrdersDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('Orders Raw');
  const dashSheet = ss.getSheetByName('Orders Dashboard');

  const values = rawSheet.getDataRange().getValues();
  dashSheet.clear();

  if (values.length < 2) {
    dashSheet.getRange('A1').setValue('No order data available. Run Refresh orders first.');
    return;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const idx = indexMap_(headers);

  const totalOrders = rows.length;
  const buyOrders = rows.filter(r => String(r[idx.transaction_type]).toUpperCase() === 'BUY').length;
  const sellOrders = rows.filter(r => String(r[idx.transaction_type]).toUpperCase() === 'SELL').length;
  const completeOrders = rows.filter(r => String(r[idx.status]).toUpperCase() === 'COMPLETE').length;
  const rejectedOrders = rows.filter(r => String(r[idx.status]).toUpperCase() === 'REJECTED').length;
  const cancelledOrders = rows.filter(r => String(r[idx.status]).toUpperCase() === 'CANCELLED').length;
  const openOrders = rows.filter(r => isOpenStatus_(String(r[idx.status]))).length;
  const filledQty = rows.reduce((s, r) => s + Number(r[idx.filled_quantity] || 0), 0);
  const executedValue = rows.reduce((s, r) => s + Number(r[idx.average_price] || 0) * Number(r[idx.filled_quantity] || 0), 0);

  dashSheet.getRange('A1').setValue('Orders Dashboard');
  dashSheet.getRange('A3:B10').setValues([
    ['Total orders', totalOrders],
    ['Buy orders', buyOrders],
    ['Sell orders', sellOrders],
    ['Complete orders', completeOrders],
    ['Open / pending orders', openOrders],
    ['Rejected orders', rejectedOrders],
    ['Cancelled orders', cancelledOrders],
    ['Filled quantity total', filledQty]
  ]);
  dashSheet.getRange('D3:E3').setValues([['Executed value', executedValue]]);

  const outputHeaders = [
    'Order ID', 'Status', 'Timestamp', 'Symbol', 'Txn', 'Product',
    'Type', 'Qty', 'Filled', 'Pending', 'Cancelled', 'Price',
    'Avg Price', 'Status Message'
  ];

  const out = rows.map(r => [
    r[idx.order_id], r[idx.status], r[idx.order_timestamp], r[idx.tradingsymbol],
    r[idx.transaction_type], r[idx.product], r[idx.order_type], r[idx.quantity],
    r[idx.filled_quantity], r[idx.pending_quantity], r[idx.cancelled_quantity],
    r[idx.price], r[idx.average_price], r[idx.status_message]
  ]);

  const startRow = 13;
  dashSheet.getRange(startRow, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  dashSheet.getRange(startRow + 1, 1, out.length, outputHeaders.length).setValues(out);

  dashSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  dashSheet.getRange('A3:A10').setFontWeight('bold');
  dashSheet.getRange('D3').setFontWeight('bold');
  dashSheet.getRange(startRow, 1, 1, outputHeaders.length).setFontWeight('bold');
  dashSheet.getRange('B3:B10').setNumberFormat('#,##0');
  dashSheet.getRange('E3').setNumberFormat('#,##0.00');
  dashSheet.getRange(startRow + 1, 8, out.length, 5).setNumberFormat('#,##0.00');

  const statusRange = dashSheet.getRange(startRow + 1, 2, Math.max(out.length, 1), 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('COMPLETE').setBackground('#d9ead3').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('REJECTED').setBackground('#f4cccc').setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('CANCELLED').setBackground('#fff2cc').setRanges([statusRange]).build()
  ];
  dashSheet.setConditionalFormatRules(rules);
  dashSheet.autoResizeColumns(1, outputHeaders.length);
}

function isOpenStatus_(status) {
  const s = String(status || '').toUpperCase();
  const openStatuses = ['OPEN', 'OPEN PENDING', 'TRIGGER PENDING', 'PUT ORDER REQ RECEIVED', 'MODIFY PENDING', 'MODIFY VALIDATION PENDING', 'AMO REQ RECEIVED', 'VALIDATION PENDING'];
  return openStatuses.indexOf(s) !== -1;
}

function getConfig_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const config = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key && key !== 'Setting') config[key] = values[i][1];
  }
  return config;
}

function ensureConfigRow_(key, defaultValue) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  if (sh.getLastRow() === 0) {
    sh.getRange('A1:B1').setValues([['Setting', 'Value']]);
  }
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) return;
  }
  sh.appendRow([key, defaultValue]);
}

function updateConfigValue_(key, value) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function logMessage_(level, message) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Level', 'Message']);
  }
  sh.appendRow([new Date(), level, message]);
}

function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}
