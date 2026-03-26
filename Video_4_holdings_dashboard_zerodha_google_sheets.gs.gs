function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Trading OS')
    .addItem('Refresh holdings', 'refreshHoldings')
    .addItem('Build holdings dashboard', 'buildHoldingsDashboard')
    .addSeparator()
    .addItem('Initialize sheets', 'initializeTradingSheets')
    .addToUi();
}

function initializeTradingSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = ['Config', 'Holdings Raw', 'Holdings Dashboard', 'Logs'];

  requiredSheets.forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
    }
    if (name === 'Config' && sh.getLastRow() === 0) {
      sh.getRange('A1:B5').setValues([
        ['Setting', 'Value'],
        ['api_key', ''],
        ['access_token', ''],
        ['last_holdings_refresh', ''],
        ['status', '']
      ]);
    }
  });

  logMessage_('INFO', 'Initialization complete');
  SpreadsheetApp.getUi().alert('Sheets initialized successfully.');
}

function refreshHoldings() {
  try {
    const config = getConfig_();
    if (!config.api_key || !config.access_token) {
      throw new Error('Missing api_key or access_token in Config sheet.');
    }

    const url = 'https://api.kite.trade/portfolio/holdings';

    const options = {
      method: 'get',
      headers: {
        'X-Kite-Version': '3',
        'Authorization': 'token ' + config.api_key + ':' + config.access_token
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const text = response.getContentText();

    if (code !== 200) {
      throw new Error('API error ' + code + ': ' + text);
    }

    const payload = JSON.parse(text);

    if (payload.status !== 'success' || !payload.data) {
      throw new Error('Unexpected API response: ' + text);
    }

    writeHoldingsRaw_(payload.data);
    updateConfigValue_('last_holdings_refresh', new Date());
    updateConfigValue_('status', 'Holdings refreshed successfully');
    buildHoldingsDashboard();

    logMessage_('INFO', 'Holdings refreshed successfully. Rows: ' + payload.data.length);
    SpreadsheetApp.getUi().alert('Holdings refreshed successfully.');
  } catch (err) {
    updateConfigValue_('status', 'Error: ' + err.message);
    logMessage_('ERROR', err.message);
    SpreadsheetApp.getUi().alert('Refresh failed: ' + err.message);
  }
}

function buildHoldingsDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('Holdings Raw');
  const dashSheet = ss.getSheetByName('Holdings Dashboard');

  if (!rawSheet || !dashSheet) {
    throw new Error('Required sheets not found.');
  }

  const values = rawSheet.getDataRange().getValues();
  if (values.length < 2) {
    dashSheet.clear();
    dashSheet.getRange('A1').setValue('No holdings data available. Run Refresh holdings first.');
    return;
  }

  const headers = values[0];
  const rows = values.slice(1);

  const idx = indexMap_(headers);

  const enriched = rows.map(r => {
    const quantity = Number(r[idx.quantity] || 0);
    const avgPrice = Number(r[idx.average_price] || 0);
    const lastPrice = Number(r[idx.last_price] || 0);
    const invested = quantity * avgPrice;
    const current = quantity * lastPrice;
    const pnl = Number(r[idx.pnl] || 0);
    const dayChangePct = Number(r[idx.day_change_percentage] || 0);

    return [
      r[idx.tradingsymbol],
      r[idx.exchange],
      quantity,
      avgPrice,
      lastPrice,
      invested,
      current,
      pnl,
      dayChangePct
    ];
  });

  const totalInvested = enriched.reduce((s, r) => s + Number(r[5] || 0), 0);
  const totalCurrent = enriched.reduce((s, r) => s + Number(r[6] || 0), 0);
  const totalPnl = enriched.reduce((s, r) => s + Number(r[7] || 0), 0);
  const holdingsCount = enriched.length;

  dashSheet.clear();

  // KPI area
  dashSheet.getRange('A1').setValue('Holdings Dashboard');
  dashSheet.getRange('A3:B6').setValues([
    ['Number of holdings', holdingsCount],
    ['Total invested value', totalInvested],
    ['Current market value', totalCurrent],
    ['Overall P&L', totalPnl]
  ]);

  // Table header
  const startRow = 9;
  const outputHeaders = [
    'Tradingsymbol', 'Exchange', 'Quantity', 'Average Price', 'Last Price',
    'Invested Value', 'Current Value', 'P&L', 'Day Change %'
  ];

  dashSheet.getRange(startRow, 1, 1, outputHeaders.length).setValues([outputHeaders]);
  if (enriched.length > 0) {
    dashSheet.getRange(startRow + 1, 1, enriched.length, outputHeaders.length).setValues(enriched);
  }

  // Formatting
  dashSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  dashSheet.getRange('A3:A6').setFontWeight('bold');
  dashSheet.getRange(startRow, 1, 1, outputHeaders.length).setFontWeight('bold');

  dashSheet.getRange('B4:B6').setNumberFormat('#,##0.00');
  if (enriched.length > 0) {
    dashSheet.getRange(startRow + 1, 4, enriched.length, 6).setNumberFormat('#,##0.00');
  }

  // Conditional formatting on P&L column
  const pnlRange = dashSheet.getRange(startRow + 1, 8, Math.max(enriched.length, 1), 1);
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

  dashSheet.autoResizeColumns(1, outputHeaders.length);
  logMessage_('INFO', 'Holdings dashboard built successfully.');
}

function writeHoldingsRaw_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Holdings Raw');
  if (!sh) throw new Error('Holdings Raw sheet not found.');

  sh.clear();

  const headers = [
    'tradingsymbol',
    'exchange',
    'instrument_token',
    'isin',
    'product',
    'quantity',
    'used_quantity',
    't1_quantity',
    'realised_quantity',
    'authorised_quantity',
    'opening_quantity',
    'collateral_quantity',
    'average_price',
    'last_price',
    'close_price',
    'pnl',
    'day_change',
    'day_change_percentage',
    'discrepancy'
  ];

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);

  const rows = data.map(item => [
    item.tradingsymbol || '',
    item.exchange || '',
    item.instrument_token || '',
    item.isin || '',
    item.product || '',
    item.quantity || 0,
    item.used_quantity || 0,
    item.t1_quantity || 0,
    item.realised_quantity || 0,
    item.authorised_quantity || 0,
    item.opening_quantity || 0,
    item.collateral_quantity || 0,
    item.average_price || 0,
    item.last_price || 0,
    item.close_price || 0,
    item.pnl || 0,
    item.day_change || 0,
    item.day_change_percentage || 0,
    item.discrepancy || false
  ]);

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sh.autoResizeColumns(1, headers.length);
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
    if (key) {
      config[key] = value;
    }
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

function logMessage_(level, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Logs');
  if (!sh) {
    sh = ss.insertSheet('Logs');
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Level', 'Message']);
  }

  sh.appendRow([new Date(), level, message]);
}

function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[h] = i;
  });
  return map;
}
