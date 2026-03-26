function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Trading OS')
    .addItem('Initialize setup', 'initializeSetup')
    .addItem('Set API secret', 'setApiSecret')
    .addSeparator()
    .addItem('Show login URL', 'showLoginUrl')
    .addItem('Generate access token', 'generateAccessToken')
    .addItem('Test connection', 'testConnection')
    .addSeparator()
    .addItem('Clear saved tokens', 'clearSavedTokens')
    .addToUi();
}
 
function initializeSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scriptProps = PropertiesService.getScriptProperties();
 
  scriptProps.setProperty('SPREADSHEET_ID', ss.getId());
 
  let config = ss.getSheetByName('Config');
  if (!config) config = ss.insertSheet('Config');
 
  let logs = ss.getSheetByName('Logs');
  if (!logs) logs = ss.insertSheet('Logs');
 
  if (config.getLastRow() === 0) {
    config.getRange('A1:B7').setValues([
      ['Setting', 'Value'],
      ['api_key', ''],
      ['redirect_url', ''],
      ['request_token', ''],
      ['access_token', ''],
      ['last_auth_time', ''],
      ['status', '']
    ]);
    config.getRange('A1:B1').setFontWeight('bold');
    config.autoResizeColumns(1, 2);
  }
 
  if (logs.getLastRow() === 0) {
    logs.appendRow(['Timestamp', 'Level', 'Message']);
    logs.getRange('A1:C1').setFontWeight('bold');
    logs.autoResizeColumns(1, 3);
  }
 
  logMessage_('INFO', 'Setup initialized');
  SpreadsheetApp.getUi().alert(
    'Setup initialized.\n\nNext:\n1. Enter api_key in Config\n2. Set API secret from the menu\n3. Deploy this script as a web app\n4. Paste the web app URL into Config and your Kite app redirect URL'
  );
}
 
function setApiSecret() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Set Zerodha API Secret',
    'Paste your api_secret here. It will be stored in User Properties, not in the sheet.',
    ui.ButtonSet.OK_CANCEL
  );
 
  if (response.getSelectedButton() !== ui.Button.OK) return;
 
  const apiSecret = response.getResponseText().trim();
  if (!apiSecret) {
    ui.alert('API secret cannot be blank.');
    return;
  }
 
  PropertiesService.getUserProperties().setProperty('API_SECRET', apiSecret);
  logMessage_('INFO', 'API secret saved to User Properties');
  ui.alert('API secret saved successfully.');
}
 
function showLoginUrl() {
  const config = getConfig_();
  const apiKey = (config.api_key || '').trim();
  const redirectUrl = (config.redirect_url || '').trim();
 
  if (!apiKey || !redirectUrl) {
    SpreadsheetApp.getUi().alert(
      'Please fill api_key and redirect_url in the Config sheet first.'
    );
    return;
  }
 
  const loginUrl =
    'https://kite.zerodha.com/connect/login?v=3' +
    '&api_key=' + encodeURIComponent(apiKey);
 
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:16px;">' +
      '<h3>Zerodha Login URL</h3>' +
      '<p>Make sure this redirect URL is registered in your Kite app:</p>' +
      '<pre style="white-space:pre-wrap;background:#f4f4f4;padding:8px;">' + escapeHtml_(redirectUrl) + '</pre>' +
      '<p>Open this link, complete login, and Zerodha will redirect to your Apps Script web app with the request_token.</p>' +
      '<p><a href="' + loginUrl + '" target="_blank">Open Zerodha Login</a></p>' +
    '</div>'
  ).setWidth(500).setHeight(260);
 
  SpreadsheetApp.getUi().showModalDialog(html, 'Zerodha Login');
}
 
function doGet(e) {
  const requestToken = e && e.parameter ? e.parameter.request_token : '';
  const status = e && e.parameter ? e.parameter.status : '';
 
  if (requestToken) {
    PropertiesService.getUserProperties().setProperty('REQUEST_TOKEN', requestToken);
    updateSheetConfigFromWebApp_('request_token', requestToken);
    updateSheetConfigFromWebApp_('status', 'Request token received');
    logMessageFromWebApp_('INFO', 'Request token captured via web app');
  }
 
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:24px;">' +
      '<h2>Zerodha authorization received</h2>' +
      '<p>Status: ' + escapeHtml_(status || 'unknown') + '</p>' +
      '<p>Request token: ' + escapeHtml_(requestToken || 'not found') + '</p>' +
      '<p>You can now go back to the Google Sheet and click <b>Trading OS → Generate access token</b>.</p>' +
    '</div>'
  );
}
 
function generateAccessToken() {
  try {
    const config = getConfig_();
    const apiKey = (config.api_key || '').trim();
    const requestTokenFromSheet = (config.request_token || '').trim();
    const requestTokenFromProps = PropertiesService.getUserProperties().getProperty('REQUEST_TOKEN') || '';
    const requestToken = requestTokenFromProps || requestTokenFromSheet;
    const apiSecret = PropertiesService.getUserProperties().getProperty('API_SECRET') || '';
 
    if (!apiKey) throw new Error('Missing api_key in Config sheet.');
    if (!requestToken) throw new Error('Missing request_token. Complete Zerodha login first.');
    if (!apiSecret) throw new Error('Missing api_secret. Use Trading OS → Set API secret.');
 
    const checksum = sha256Hex_(apiKey + requestToken + apiSecret);
 
    const response = UrlFetchApp.fetch('https://api.kite.trade/session/token', {
      method: 'post',
      headers: {
        'X-Kite-Version': '3'
      },
      payload: {
        api_key: apiKey,
        request_token: requestToken,
        checksum: checksum
      },
      muteHttpExceptions: true
    });
 
    const code = response.getResponseCode();
    const text = response.getContentText();
 
    if (code !== 200) {
      throw new Error('Token exchange failed. HTTP ' + code + ': ' + text);
    }
 
    const json = JSON.parse(text);
    if (json.status !== 'success' || !json.data || !json.data.access_token) {
      throw new Error('Unexpected token response: ' + text);
    }
 
    setConfig_('access_token', json.data.access_token);
    setConfig_('request_token', requestToken);
    setConfig_('last_auth_time', new Date());
    setConfig_('status', 'Access token generated successfully');
 
    logMessage_('INFO', 'Access token generated successfully for user_id: ' + (json.data.user_id || 'unknown'));
 
    SpreadsheetApp.getUi().alert('Access token generated and saved in Config sheet.');
  } catch (err) {
    setConfig_('status', 'Error: ' + err.message);
    logMessage_('ERROR', err.message);
    SpreadsheetApp.getUi().alert(err.message);
  }
}
 
function testConnection() {
  try {
    const config = getConfig_();
    const apiKey = (config.api_key || '').trim();
    const accessToken = (config.access_token || '').trim();
 
    if (!apiKey || !accessToken) {
      throw new Error('Missing api_key or access_token in Config sheet.');
    }
 
    const response = UrlFetchApp.fetch('https://api.kite.trade/user/profile', {
      method: 'get',
      headers: {
        'X-Kite-Version': '3',
        'Authorization': 'token ' + apiKey + ':' + accessToken
      },
      muteHttpExceptions: true
    });
 
    const code = response.getResponseCode();
    const text = response.getContentText();
 
    if (code !== 200) {
      throw new Error('Connection test failed. HTTP ' + code + ': ' + text);
    }
 
    const json = JSON.parse(text);
    if (json.status !== 'success' || !json.data) {
      throw new Error('Unexpected response: ' + text);
    }
 
    setConfig_('status', 'Connected as ' + (json.data.user_name || json.data.user_id || 'user'));
    logMessage_('INFO', 'Connected successfully to Zerodha profile endpoint');
 
    SpreadsheetApp.getUi().alert(
      'Connection successful.\n\nUser: ' + (json.data.user_name || '') +
      '\nUser ID: ' + (json.data.user_id || '') +
      '\nBroker: ' + (json.data.broker || '')
    );
  } catch (err) {
    setConfig_('status', 'Error: ' + err.message);
    logMessage_('ERROR', err.message);
    SpreadsheetApp.getUi().alert(err.message);
  }
}
 
function clearSavedTokens() {
  PropertiesService.getUserProperties().deleteProperty('REQUEST_TOKEN');
  setConfig_('request_token', '');
  setConfig_('access_token', '');
  setConfig_('status', 'Tokens cleared');
  logMessage_('INFO', 'Saved tokens cleared');
  SpreadsheetApp.getUi().alert('Saved request_token and access_token cleared.');
}
 
function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Config');
  if (!sh) throw new Error('Config sheet not found.');
 
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  const config = {};
 
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const value = values[i][1];
    if (key) config[key] = value;
  }
 
  return config;
}
 
function setConfig_(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Config');
  if (!sh) throw new Error('Config sheet not found.');
 
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
 
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
 
  sh.appendRow([key, value]);
}
 
function updateSheetConfigFromWebApp_(key, value) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) return;
 
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sh = ss.getSheetByName('Config');
  if (!sh) return;
 
  const values = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
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
 
function logMessageFromWebApp_(level, message) {
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) return;
 
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sh = ss.getSheetByName('Logs');
  if (!sh) sh = ss.insertSheet('Logs');
 
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Level', 'Message']);
  }
 
  sh.appendRow([new Date(), level, message]);
}
 
function sha256Hex_(text) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return raw.map(function(byte) {
    const v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}
 
function escapeHtml_(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
