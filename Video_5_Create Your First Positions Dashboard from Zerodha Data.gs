function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Alladin')
    .addItem('Initialize Video 5 sheets', 'initializeVideo5Sheets')
    .addItem('Refresh positions', 'refreshPositions')
    .addItem('Build positions dashboard', 'buildPositionsDashboard')
    .addToUi();
}
 
function initializeVideo5Sheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Config', 'Positions Raw', 'Positions Dashboard', 'Logs'];
 
  requiredSheets.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
 
    if (name === 'Config' && sh.getLastRow() === 0) {
      sh.getRange('A1:B5').setValues([
        ['Setting', 'Value'],
        ['api_key', ''],
        ['access_token', ''],
        ['last_positions_refresh', ''],
        ['status', '']
      ]);
    }
  });
 
  logMessage_('INFO', 'Video 5 sheets initialized');
  SpreadsheetApp.getUi().alert('Video 5 sheets initialized successfully.');
}
 
function refreshPositions() {
  try {
    const config = getConfig_();
    if (!config.api_key || !config.access_token) {
      throw new Error('Missing api_key or access_token in Config sheet.');
    }
 
    const response = UrlFetchApp.fetch('https://api.kite.trade/portfolio/positions', {
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
 
    writePositionsRaw_(payload.data);
    updateConfigValue_('last_positions_refresh', new Date());
    updateConfigValue_('status', 'Positions refreshed successfully');
    buildPositionsDashboard();
 
    logMessage_('INFO', 'Positions refreshed successfully');
    SpreadsheetApp.getUi().alert('Positions refreshed successfully.');
  } catch (err) {
    updateConfigValue_('status', 'Error: ' + err.message);
    logMessage_('ERROR', err.message);
    SpreadsheetApp.getUi().alert('Refresh failed: ' + err.message);
  }
}
 
function writePositionsRaw_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Positions Raw');
  if (!sh) throw new Error('Positions Raw sheet not found.');
 
  sh.clear();
 
  const headers = [
    'position_set', 'tradingsymbol', 'exchange', 'product', 'quantity',
    'overnight_quantity', 'multiplier', 'average_price', 'last_price',
    'close_price', 'value', 'pnl', 'm2m', 'unrealised', 'realised',
    'buy_quantity', 'buy_price', 'buy_value', 'sell_quantity', 'sell_price',
    'sell_value', 'day_buy_quantity', 'day_buy_price', 'day_buy_value',
    'day_sell_quantity', 'day_sell_price', 'day_sell_value'
  ];
 
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
 
  const rows = [];
  ['net', 'day'].forEach(positionSet => {
    const arr = data[positionSet] || [];
    arr.forEach(item => {
      rows.push([
        positionSet,
        item.tradingsymbol || '',
        item.exchange || '',
        item.product || '',
        item.quantity || 0,
        item.overnight_quantity || 0,
        item.multiplier || 0,
        item.average_price || 0,
        item.last_price || 0,
        item.close_price || 0,
        item.value || 0,
        item.pnl || 0,
        item.m2m || 0,
        item.unrealised || 0,
        item.realised || 0,
        item.buy_quantity || 0,
        item.buy_price || 0,
        item.buy_value || 0,
        item.sell_quantity || 0,
        item.sell_price || 0,
        item.sell_value || 0,
        item.day_buy_quantity || 0,
        item.day_buy_price || 0,
        item.day_buy_value || 0,
        item.day_sell_quantity || 0,
        item.day_sell_price || 0,
        item.day_sell_value || 0
      ]);
    });
  });
 
  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
 
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.autoResizeColumns(1, headers.length);
}
 
function buildPositionsDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('Positions Raw');
  const dashSheet = ss.getSheetByName('Positions Dashboard');
  if (!rawSheet || !dashSheet) throw new Error('Required sheets not found.');
 
  const values = rawSheet.getDataRange().getValues();
  if (values.length < 2) {
    dashSheet.clear();
    dashSheet.getRange('A1').setValue('No positions data available. Run Refresh positions first.');
    return;
  }
 
  const headers = values[0];
  const rows = values.slice(1);
  const idx = indexMap_(headers);
 
  const netRows = rows.filter(r => String(r[idx.position_set]) === 'net');
 
  const totalOpen = netRows.filter(r => Number(r[idx.quantity] || 0) !== 0).length;
  const totalPnl = netRows.reduce((s, r) => s + Number(r[idx.pnl] || 0), 0);
  const totalM2M = netRows.reduce((s, r) => s + Number(r[idx.m2m] || 0), 0);
  const grossExposure = netRows.reduce((s, r) => s + Math.abs(Number(r[idx.value] || 0)), 0);
 
  const output = netRows.map(r => [
    r[idx.tradingsymbol],
    r[idx.exchange],
    r[idx.product],
    Number(r[idx.quantity] || 0),
    Number(r[idx.average_price] || 0),
    Number(r[idx.last_price] || 0),
    Number(r[idx.value] || 0),
    Number(r[idx.pnl] || 0),
    Number(r[idx.m2m] || 0),
    Number(r[idx.unrealised] || 0),
    Number(r[idx.realised] || 0),
    Number(r[idx.day_buy_quantity] || 0),
    Number(r[idx.day_sell_quantity] || 0)
  ]);
 
  dashSheet.clear();
 
  dashSheet.getRange('A1').setValue('Positions Dashboard');
  dashSheet.getRange('A3:B6').setValues([
    ['Open positions', totalOpen],
    ['Gross exposure', grossExposure],
    ['Total P&L', totalPnl],
    ['Total M2M', totalM2M]
  ]);
 
  const hdr = [
    'Tradingsymbol', 'Exchange', 'Product', 'Quantity', 'Average Price',
    'Last Price', 'Value', 'P&L', 'M2M', 'Unrealised', 'Realised',
    'Day Buy Qty', 'Day Sell Qty'
  ];
 
  const startRow = 9;
  dashSheet.getRange(startRow, 1, 1, hdr.length).setValues([hdr]);
  if (output.length > 0) {
    dashSheet.getRange(startRow + 1, 1, output.length, hdr.length).setValues(output);
  }
 
  dashSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  dashSheet.getRange('A3:A6').setFontWeight('bold');
  dashSheet.getRange(startRow, 1, 1, hdr.length).setFontWeight('bold');
  dashSheet.getRange('B4:B6').setNumberFormat('#,##0.00');
 
  if (output.length > 0) {
    dashSheet.getRange(startRow + 1, 5, output.length, 9).setNumberFormat('#,##0.00');
    const pnlRange = dashSheet.getRange(startRow + 1, 8, output.length, 1);
    const rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#d9ead3')
        .setRanges([pnlRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#f4cccc')
        .setRanges([pnlRange])
        .build()
    ];
    dashSheet.setConditionalFormatRules(rules);
  }
 
  dashSheet.autoResizeColumns(1, hdr.length);
  logMessage_('INFO', 'Positions dashboard built successfully');
}
 
function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Config');
  if (!sh) throw new Error('Config sheet not found.');
 
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const key = String(data[i][0] || '').trim();
    const value = data[i][1];
    if (key) config[key] = value;
  }
  return config;
}
 
function updateConfigValue_(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Config');
  if (!sh) throw new Error('Config sheet not found.');
 
  const data = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}
 
function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}
 
function logMessage_(level, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Logs');
  if (!sh) sh = ss.insertSheet('Logs');
 
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Level', 'Message']);
  }
 
  sh.appendRow([new Date(), level, message]);
}
