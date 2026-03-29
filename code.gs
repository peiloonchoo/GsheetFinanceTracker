// ============================================================
//  PERSONAL FINANCE TRACKER v2
//  Menu: Accounts | Transactions | Stocks | Mutual Funds | Dashboard
// ============================================================


// ── MENU ─────────────────────────────────────────────────────
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Finance')
    .addItem('📊 Open Dashboard',               'renderDashboardSheet')
    .addSeparator()
    .addItem('Create Account',                  'createAccount')
    .addItem('Add Transaction',                 'showTransactionDialog')
    .addSeparator()
    .addItem('Create Stock Portfolio Sheet',    'createStockPortfolioSheet')
    .addItem('Add Stock',                       'showAddStockDialog')
    .addItem('💸 Sell Stock',                  'showSellStockDialog')
    .addSeparator()
    .addItem('🏦 Create Mutual Fund Sheet',     'createMutualFundSheet')
    .addItem('➕ Add Fund',                     'showAddFundDialog')
    .addItem('💸 Sell Fund',                    'showSellFundDialog')
    .addSeparator()
    .addItem('🪙 Create Crypto Portfolio Sheet','createCryptoSheet')
    .addItem('➕ Add Crypto',                   'showAddCryptoDialog')
    .addItem('💸 Sell Crypto',                  'showSellCryptoDialog')
    .addSeparator()
    .addItem('🥇 Create Gold Portfolio Sheet',  'createGoldSheet')
    .addItem('➕ Add Gold',                      'showAddGoldDialog')
    .addItem('💸 Sell Gold',                     'showSellGoldDialog')
    .addSeparator()
    .addItem('🔄 Refresh All Prices Now',       'refreshAllStockPrices')
    .addItem('💰 Refresh Dividend Info',          'refreshDividendInfo')
    .addItem('⏱️ Enable Auto-Refresh (5 min)',  'enableAutoRefresh')
    .addItem('⏹️ Disable Auto-Refresh',         'disableAutoRefresh')
    .addSeparator()
    .addItem('📅 Enable Daily NAV Update (7 PM)', 'enableDailyNavRefresh')
    .addItem('⏹️ Disable Daily NAV Update',       'disableDailyNavRefresh')
    .addSeparator()
    .addItem('📋 Refresh FIMM NAV Cache Now',       'refreshFimmNavCache')
    .addItem('📅 Enable Daily FIMM Update (7 AM)',  'enableDailyFimmRefresh')
    .addItem('⏹️ Disable Daily FIMM Update',        'disableDailyFimmRefresh')
    .addSeparator()
    .addItem('💹 Set Account Interest Rate',      'showSetInterestDialog')
    .addItem('💹 Apply Interest Now',             'applyInterestNow')
    .addItem('📅 Enable Daily Interest Posting',  'enableDailyInterest')
    .addItem('⏹️ Disable Daily Interest Posting', 'disableDailyInterest')
    .addSeparator()
    .addItem('🏖️ Configure Retirement Portfolio', 'showRetirementConfigDialog')
    .addItem('🏖️ Refresh Retirement Sheet',      'refreshRetirementSheet')
    .addSeparator()
    .addItem('🤝 Create Loans & Debts Sheet',    'createLoansSheet')
    .addItem('➕ Add Lent / Borrowed',           'showAddLoanDialog')
    .addItem('💰 Record Repayment',              'showRepaymentDialog')
    .addToUi();
}


// ── CONSTANTS ─────────────────────────────────────────────────
const DASH_NAME          = '📊 Total Account Balance';
const MF_SHEET_NAME      = '🏦 Mutual Funds';
const LOANS_SHEET_NAME   = '🤝 Loans & Debts';
const FSM_BASE           = 'https://www.fsmone.com.my';
const FIMM_PDF_URL       = 'https://www.fimm.com.my/report/dailypricefms2.php';
const FIMM_NAV_SHEET     = '📋 FIMM NAV Cache';
const FIMM_NAV_TRIGGER_KEY = 'fimmNavTriggerId';
const FSM_SEARCH_URL     = FSM_BASE + '/rest/product/search-products-and-underlying-by-keyword/v2/read?product=UT&keyword=';
const FSM_FACTSHEET_URL  = FSM_BASE + '/rest/fund/get-factsheet?paramSedolnumber=';
const AUTO_REFRESH_TRIGGER_KEY  = 'autoRefreshTriggerId';
const DAILY_NAV_TRIGGER_KEY     = 'dailyNavTriggerId';
const DAILY_INTEREST_TRIGGER_KEY = 'dailyInterestTriggerId';

const MF_COLS = {
  NAME: 1, CODE: 2, UNITS: 3, BUY_NAV: 4, CUR_NAV: 5,
  MKT_VAL: 6, GAIN: 7, GAIN_PCT: 8, CCY: 9,
  ACCOUNT: 10, UPDATED: 11, NOTES: 12
};


// ── CREATE ACCOUNT ───────────────────────────────────────────
function createAccount() {
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#1a73e8}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#1a73e8;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>Create New Account</h2>' +
    '<label>Account Name</label>' +
    '<input type="text" id="name" placeholder="e.g. Maybank Savings" />' +
    '<label>Currency</label>' +
    '<select id="currency"><option value="MYR">MYR – Malaysian Ringgit</option>' +
    '<option value="USD">USD – US Dollar</option><option value="SGD">SGD – Singapore Dollar</option>' +
    '<option value="HKD">HKD – Hong Kong Dollar</option><option value="RMB">RMB – Chinese Yuan</option></select>' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="save()">Create Account</button></div>' +
    '<script>function save(){' +
    'var n=document.getElementById("name").value.trim();' +
    'var c=document.getElementById("currency").value;' +
    'if(!n){alert("Enter account name.");return;}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message);})' +
    '.createAccountFromDialog(n,c);}' +
    '</script></body></html>'
  ).setWidth(400).setHeight(260).setTitle('Create Account');
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Account');
}

function createAccountFromDialog(name, currency) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(name)) throw new Error('Sheet "' + name + '" already exists.');

  const fmtMap = {
    MYR: '"RM "#,##0.00', USD: '"$"#,##0.00',
    SGD: '"S$"#,##0.00', HKD: '"HK$"#,##0.00',
    RMB: '"¥"#,##0.00',  CNY: '"¥"#,##0.00'
  };
  const symMap = { MYR:'RM', USD:'$', SGD:'S$', HKD:'HK$', RMB:'¥', CNY:'¥' };
  const fmt    = fmtMap[currency] || fmtMap['MYR'];
  const sym    = symMap[currency] || 'RM';

  const sheet = ss.insertSheet(name, ss.getSheets().length);
  compactSheet_(sheet, 50, 9);  // cols A-I (G=currency, H=interest rate, I=frequency)

  // Banner row
  sheet.setRowHeight(1, 52);
  sheet.getRange(1, 1, 1, 6).merge()
    .setValue('📒  ' + name.toUpperCase() + '  (' + currency + ')')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Headers row
  sheet.setRowHeight(2, 30);
  const headers = ['Date','Category','Description','Amount ('+sym+')','Type','Balance ('+sym+')'];
  headers.forEach((h, i) => {
    sheet.getRange(2, i + 1)
      .setValue(h).setBackground('#e8f0fe').setFontColor('#1a73e8')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // F2 = "Balance" marker so getAccountNames() can detect it
  sheet.getRange('F2').setValue('Balance');

  // Column widths
  [90, 120, 200, 110, 70, 110].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Number formats for amount and balance columns
  sheet.getRange(3, 4, 48, 1).setNumberFormat(fmt);
  sheet.getRange(3, 6, 48, 1).setNumberFormat(fmt);

  // Hidden col G stores currency code
  sheet.getRange('G2').setValue(currency);
  sheet.hideColumns(7, 3);  // hide G (currency), H (interest rate), I (interest frequency)

  // Opening balance row — RM 0.00 so the account shows 0 on the dashboard
  // and getAccountBalances() has a valid last row to read from.
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  sheet.setRowHeight(3, 32);
  sheet.getRange(3, 1).setValue(today).setFontColor('#9aa0a6').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(3, 2).setValue('Opening').setFontColor('#9aa0a6').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(3, 3).setValue('Opening balance').setFontColor('#9aa0a6').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(3, 4).setValue(0).setNumberFormat(fmt).setFontColor('#9aa0a6').setFontSize(10).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(3, 5).setValue('IN').setFontColor('#9aa0a6').setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(3, 6).setValue(0).setNumberFormat(fmt).setFontColor('#9aa0a6').setFontSize(11).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(3, 1, 1, 6).setBackground('#f8f9fa');

  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
  sheet.activate();
}



// ── ACCOUNT INTEREST ─────────────────────────────────────────
// Interest config stored in hidden cols on each account sheet:
//   H2 = annual interest rate (%, e.g. 3.5)
//   I2 = frequency: "daily" or "monthly"
// A daily trigger posts interest entries automatically.

function showSetInterestDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheets = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'));
  if (!accSheets.length) {
    SpreadsheetApp.getUi().alert('No accounts found. Create an account first.');
    return;
  }

  // Read existing configs
  const accounts = accSheets.map(s => {
    const rate = s.getRange('H2').getValue();
    const freq = s.getRange('I2').getValue();
    return {
      name:  s.getName(),
      rate:  typeof rate === 'number' && rate > 0 ? rate : 0,
      freq:  freq || 'none'
    };
  });
  const acctJson = JSON.stringify(accounts);

  const htmlStr = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:0;background:#f8f9fa;color:#202124;font-size:13px}'
    + '.hdr{background:#1a73e8;color:#fff;padding:14px 18px;font-size:15px;font-weight:700}'
    + '.body{padding:16px}'
    + 'label{display:block;font-weight:600;margin:12px 0 4px;color:#3c4043;font-size:12px}'
    + 'select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}'
    + '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px}'
    + '.field{display:flex;flex-direction:column}'
    + 'input{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}'
    + 'input:focus,select:focus{outline:none;border-color:#1a73e8}'
    + '.infobox{background:#e8f0fe;border-radius:8px;padding:10px 14px;margin-top:10px;font-size:12px;color:#1a73e8;display:none}'
    + '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:8px 11px;font-size:12px;color:#856404;margin-top:8px;display:none}'
    + '.footer{display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;border-top:1px solid #e8eaed;background:#fff}'
    + '.btn{padding:8px 22px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}'
    + '.bsave{background:#1a73e8;color:#fff}.bsave:hover{background:#1557b0}'
    + '.bcancel{background:#f1f3f4;color:#3c4043}'
    + '.bclear{background:#fce8e6;color:#c62828}'
    + '</style></head><body>'
    + '<div class="hdr">💹 Set Account Interest Rate</div>'
    + '<div class="body">'
    + '<label>Account</label>'
    + '<select id="account"></select>'
    + '<div class="row2">'
    + '<div class="field"><label>Annual Interest Rate (%)</label>'
    + '<input type="number" id="rate" placeholder="e.g. 3.5" step="0.01" min="0" max="100" /></div>'
    + '<div class="field"><label>Interest Posted</label>'
    + '<select id="freq">'
    + '<option value="none">— None (disable) —</option>'
    + '<option value="daily">Daily</option>'
    + '<option value="monthly">Monthly (1st of month)</option>'
    + '</select></div>'
    + '</div>'
    + '<div class="infobox" id="preview"></div>'
    + '<div class="warn" id="warn">Setting rate to 0 or frequency to "None" will disable interest for this account.</div>'
    + '</div>'
    + '<div class="footer">'
    + '<button class="btn bcancel" id="cancelBtn">Cancel</button>'
    + '<button class="btn bclear" id="clearBtn">Remove Interest</button>'
    + '<button class="btn bsave" id="saveBtn">Save</button>'
    + '</div>'
    + '<script>'
    + 'var ACCOUNTS=' + acctJson + ';'
    + 'var sel=document.getElementById("account");'
    + 'ACCOUNTS.forEach(function(a){'
    + '  var opt=document.createElement("option");opt.value=a.name;opt.textContent=a.name;sel.appendChild(opt);'
    + '});'
    + 'document.getElementById("cancelBtn").onclick=function(){google.script.host.close();};'
    + 'document.getElementById("clearBtn").onclick=function(){saveInterest(0,"none");};'
    + 'document.getElementById("saveBtn").onclick=doSave;'
    + 'document.getElementById("account").onchange=loadAccount;'
    + 'document.getElementById("rate").oninput=updatePreview;'
    + 'document.getElementById("freq").onchange=updatePreview;'
    + 'loadAccount();'
    + 'function loadAccount(){'
    + '  var name=sel.value;'
    + '  var a=ACCOUNTS.find(function(x){return x.name===name;});'
    + '  if(!a)return;'
    + '  document.getElementById("rate").value=a.rate>0?a.rate:"";'
    + '  document.getElementById("freq").value=a.freq||"none";'
    + '  updatePreview();'
    + '}'
    + 'function updatePreview(){'
    + '  var r=parseFloat(document.getElementById("rate").value)||0;'
    + '  var f=document.getElementById("freq").value;'
    + '  var box=document.getElementById("preview");'
    + '  var warn=document.getElementById("warn");'
    + '  if(r>0&&f!=="none"){'
    + '    var daily=(r/100/365);'
    + '    var monthly=(r/100/12);'
    + '    var perEntry=f==="daily"?daily:monthly;'
    + '    box.innerHTML="<strong>On a RM 10,000 balance:</strong><br>"'
    + '      +(f==="daily"?'
    + '        "Daily entry: ~RM "+(10000*daily).toFixed(4)+" &nbsp;·&nbsp; Monthly: ~RM "+(10000*daily*30).toFixed(2)+" &nbsp;·&nbsp; Annual: ~RM "+(10000*r/100).toFixed(2)'
    + '        :"Monthly entry: ~RM "+(10000*monthly).toFixed(2)+" &nbsp;·&nbsp; Annual: ~RM "+(10000*r/100).toFixed(2));'
    + '    box.style.display="block";'
    + '    warn.style.display="none";'
    + '  }else{'
    + '    box.style.display="none";'
    + '    if(r===0||f==="none")warn.style.display="block"; else warn.style.display="none";'
    + '  }'
    + '}'
    + 'function doSave(){'
    + '  var r=parseFloat(document.getElementById("rate").value);'
    + '  var f=document.getElementById("freq").value;'
    + '  if(isNaN(r)||r<0){alert("Enter a valid interest rate (0 or above).");return;}'
    + '  saveInterest(r,f);'
    + '}'
    + 'function saveInterest(rate,freq){'
    + '  var acct=document.getElementById("account").value;'
    + '  document.getElementById("saveBtn").disabled=true;'
    + '  document.getElementById("saveBtn").textContent="Saving...";'
    + '  google.script.run'
    + '    .withSuccessHandler(function(msg){alert(msg);google.script.host.close();})'
    + '    .withFailureHandler(function(e){alert("Error: "+e.message);document.getElementById("saveBtn").disabled=false;document.getElementById("saveBtn").textContent="Save";})'
    + '    .saveAccountInterest(acct,rate,freq);'
    + '}'
    + '</script></body></html>';

  const html = HtmlService.createHtmlOutput(htmlStr).setWidth(460).setHeight(440).setTitle('Set Account Interest Rate');
  SpreadsheetApp.getUi().showModalDialog(html, 'Set Account Interest Rate');
}

function saveAccountInterest(accountName, rate, freq) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(accountName);
  if (!sheet) throw new Error('Account not found: ' + accountName);

  // Unhide col H/I temporarily, write, re-hide
  sheet.showColumns(8, 2);
  if (rate > 0 && freq !== 'none') {
    sheet.getRange('H2').setValue(rate);
    sheet.getRange('I2').setValue(freq);
  } else {
    sheet.getRange('H2').clearContent();
    sheet.getRange('I2').clearContent();
  }
  sheet.hideColumns(8, 2);

  if (rate > 0 && freq !== 'none') {
    return accountName + ': Interest set to ' + rate + '% per annum, posted ' + freq + '.\n\nEnable "Daily Interest Posting" from the menu to apply automatically.';
  } else {
    return accountName + ': Interest disabled.';
  }
}

/**
 * Post interest for one account.
 * Calculates interest on the current balance and posts as an IN transaction.
 * Skips if: no balance, no rate, wrong day (for monthly).
 */
function postInterestForAccount_(sheet) {
  const rate = sheet.getRange('H2').getValue();
  const freq = sheet.getRange('I2').getValue();
  if (!rate || rate <= 0 || !freq || freq === 'none') return false;

  // For monthly: only post on the 1st of the month
  const today = new Date();
  if (freq === 'monthly' && today.getDate() !== 1) return false;

  // Get current balance (last value in col F from row 3 down)
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return false;
  const balance = sheet.getRange(lastRow, 6).getValue();
  if (typeof balance !== 'number' || balance <= 0) return false;

  // Calculate interest
  const annualRate = rate / 100;
  let interest;
  if (freq === 'daily') {
    interest = balance * annualRate / 365;
  } else {
    // Monthly: divide by 12
    interest = balance * annualRate / 12;
  }
  interest = Math.round(interest * 10000) / 10000; // 4 decimal places
  if (interest <= 0) return false;

  // Post the interest as an IN transaction
  const currency = sheet.getRange('G2').getValue() || 'MYR';
  const fmtMap   = { MYR:'"RM "#,##0.0000', USD:'"$"#,##0.0000', SGD:'"S$"#,##0.0000', HKD:'"HK$"#,##0.0000', RMB:'"¥"#,##0.0000', CNY:'"¥"#,##0.0000' };
  const fmt      = fmtMap[currency] || fmtMap['MYR'];
  const dateStr  = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const newRow   = lastRow + 1;
  const prevBal  = 'F' + lastRow;
  const desc     = rate + '% p.a. ' + (freq === 'daily' ? 'daily' : 'monthly') + ' interest';

  sheet.setRowHeight(newRow, 30);
  const bg = newRow % 2 === 0 ? '#f8f9fa' : '#ffffff';
  sheet.getRange(newRow, 1, 1, 6).setBackground(bg);
  sheet.getRange(newRow, 1).setValue(dateStr);
  sheet.getRange(newRow, 2).setValue('Interest');
  sheet.getRange(newRow, 3).setValue(desc);
  sheet.getRange(newRow, 4).setValue(interest).setNumberFormat(fmt);
  sheet.getRange(newRow, 5).setValue('IN').setFontColor('#0f9d58').setFontWeight('bold');
  sheet.getRange(newRow, 6).setFormula('=' + prevBal + '+D' + newRow).setNumberFormat(fmt);

  return true;
}

/** Menu: apply interest to all eligible accounts right now */
function applyInterestNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accSheets = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'));

  let posted = [], skipped = [];
  accSheets.forEach(s => {
    const rate = s.getRange('H2').getValue();
    const freq = s.getRange('I2').getValue();
    if (rate > 0 && freq && freq !== 'none') {
      const done = postInterestForAccount_(s);
      if (done) posted.push(s.getName()); else skipped.push(s.getName() + ' (monthly: not 1st, or zero balance)');
    }
  });

  if (posted.length === 0 && skipped.length === 0) {
    SpreadsheetApp.getUi().alert('No interest-bearing accounts found.\n\nSet up interest via 💹 Set Account Interest Rate first.');
  } else {
    let msg = '';
    if (posted.length)  msg += '✅ Interest posted to:\n' + posted.map(n => '  • ' + n).join('\n');
    if (skipped.length) msg += (msg ? '\n\n' : '') + '⏭️ Skipped:\n' + skipped.map(n => '  • ' + n).join('\n');
    SpreadsheetApp.getUi().alert(msg);
  }
}

/** Internal: run by daily trigger */
function dailyInterestJob_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheets()
      .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
      .forEach(s => postInterestForAccount_(s));
  } catch(e) {
    console.error('dailyInterestJob_ error: ' + e.message);
  }
}

function enableDailyInterest() {
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(DAILY_INTEREST_TRIGGER_KEY);
  if (existing) {
    SpreadsheetApp.getUi().alert('Daily interest posting is already enabled.\n\nIt runs at midnight every night and posts to all interest-bearing accounts.');
    return;
  }
  const trigger = ScriptApp.newTrigger('dailyInterestJob_')
    .timeBased().atHour(0).everyDays(1).create();
  props.setProperty(DAILY_INTEREST_TRIGGER_KEY, trigger.getUniqueId());
  SpreadsheetApp.getUi().alert('✅ Daily Interest Posting enabled!\n\nInterest will be posted every night at midnight to all accounts with an interest rate configured.\n\nUse "Set Account Interest Rate" to configure each account.');
}

function disableDailyInterest() {
  const props    = PropertiesService.getScriptProperties();
  const savedId  = props.getProperty(DAILY_INTEREST_TRIGGER_KEY);
  if (!savedId) {
    SpreadsheetApp.getUi().alert('Daily interest posting is not currently enabled.');
    return;
  }
  ScriptApp.getProjectTriggers()
    .filter(t => t.getUniqueId() === savedId)
    .forEach(t => ScriptApp.deleteTrigger(t));
  props.deleteProperty(DAILY_INTEREST_TRIGGER_KEY);
  SpreadsheetApp.getUi().alert('Daily interest posting has been disabled.');
}

// ── ADD TRANSACTION ──────────────────────────────────────────
function showTransactionDialog() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const accSheets = ss.getSheets().filter(s => s.getRange('F2').getValue().toString().includes('Balance'));
  if (!accSheets.length) {
    SpreadsheetApp.getUi().alert('No accounts found. Create an account first.');
    return;
  }
  const accountData = accSheets.map(s => ({
    name:     s.getName(),
    currency: s.getRange('G2').getValue() || 'MYR'
  }));
  const accountsJson = JSON.stringify(accountData);

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#1a73e8}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.type-row{display:flex;gap:6px;margin:8px 0}' +
    '.type-btn{flex:1;padding:9px 4px;border:2px solid #dadce0;border-radius:8px;background:#fff;cursor:pointer;font-size:12px;font-weight:600;text-align:center}' +
    '.type-btn.active-in{border-color:#0f9d58;background:#e6f4ea;color:#0f9d58}' +
    '.type-btn.active-out{border-color:#d93025;background:#fce8e6;color:#d93025}' +
    '.type-btn.active-transfer{border-color:#1a73e8;background:#e8f0fe;color:#1a73e8}' +
    '.type-btn.active-adjust{border-color:#e65100;background:#fff3e0;color:#e65100}' +
    '.adjust-hint{font-size:11px;color:#e65100;background:#fff3e0;border:1px solid #ffcc80;border-radius:6px;padding:7px 10px;margin:6px 0 0;display:none}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#1a73e8;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '#transferRow{display:none}#categoryRow{display:block}' +
    '</style></head><body>' +
    '<h2>Add Transaction</h2>' +
    '<label>Account</label>' +
    '<select id="account" onchange="onAccountChange()">' +
    accountData.map(a => '<option value="' + a.name + '">' + a.name + ' (' + a.currency + ')</option>').join('') +
    '</select>' +
    '<div id="categoryRow"><label>Category</label>' +
    '<input type="text" id="category" placeholder="e.g. Food, Salary, Utilities" /></div>' +
    '<label>Description</label>' +
    '<input type="text" id="description" placeholder="e.g. Lunch at Restoran ABC" />' +
    '<label id="amtLabel">Amount</label>' +
    '<input type="number" id="amount" placeholder="0.00" step="0.01" />' +
    '<div class="adjust-hint" id="adjustHint">⚖️ <strong>Adjust</strong> sets the account balance directly to the value you enter. The difference is recorded as an adjustment entry.</div>' +
    '<label>Type</label>' +
    '<div class="type-row">' +
    '<div class="type-btn active-in" id="btn-IN" onclick="setType(\'IN\')">💚 IN</div>' +
    '<div class="type-btn" id="btn-OUT" onclick="setType(\'OUT\')">🔴 OUT</div>' +
    '<div class="type-btn" id="btn-Transfer" onclick="setType(\'Transfer\')">🔁 Transfer</div>' +
    '<div class="type-btn" id="btn-Adjust" onclick="setType(\'Adjust\')">⚖️ Adjust</div>' +
    '</div>' +
    '<div id="transferRow"><label>Transfer To</label>' +
    '<select id="toAccount">' +
    accountData.map(a => '<option value="' + a.name + '">' + a.name + ' (' + a.currency + ')</option>').join('') +
    '</select></div>' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="save()">Save</button></div>' +
    '<script>' +
    'var ACCOUNTS=' + accountsJson + ';' +
    'var selectedType="IN";' +
    'function setType(t){' +
    'selectedType=t;' +
    '["IN","OUT","Transfer","Adjust"].forEach(function(x){' +
    'document.getElementById("btn-"+x).className="type-btn"+(x===t?" active-"+x.toLowerCase():"");});' +
    'document.getElementById("transferRow").style.display=t==="Transfer"?"block":"none";' +
    'document.getElementById("categoryRow").style.display=(t==="Transfer"||t==="Adjust")?"none":"block";' +
    'document.getElementById("adjustHint").style.display=t==="Adjust"?"block":"none";' +
    'var isAdj=t==="Adjust";' +
    'var sym={MYR:"RM",USD:"$",SGD:"S$",HKD:"HK$",RMB:"\u00a5",CNY:"\u00a5"};' +
    'var acc=ACCOUNTS.find(function(a){return a.name===document.getElementById("account").value;});' +
    'var s=sym[acc?acc.currency:"MYR"]||"RM";' +
    'document.getElementById("amtLabel").textContent=isAdj?"New Balance ("+s+")":"Amount ("+s+")";' +
    'document.getElementById("amount").min=isAdj?"0":"0.01";' +
    'document.getElementById("amount").placeholder=isAdj?"Enter the new balance":"0.00";}' +
    'function onAccountChange(){' +
    'var sel=document.getElementById("account").value;' +
    'var acc=ACCOUNTS.find(function(a){return a.name===sel;});' +
    'var sym={MYR:"RM",USD:"$",SGD:"S$",HKD:"HK$",RMB:"\u00a5",CNY:"\u00a5"};' +
    'var s=sym[acc.currency]||"RM";' +
    'var isAdj=selectedType==="Adjust";' +
    'document.getElementById("amtLabel").textContent=isAdj?"New Balance ("+s+")":"Amount ("+s+")";' +
    'var toSel=document.getElementById("toAccount");' +
    'Array.from(toSel.options).forEach(function(o){o.disabled=o.value===sel;});' +
    'toSel.value=(ACCOUNTS.filter(function(a){return a.name!==sel;})[0]||{}).name||"";}' +
    'function save(){' +
    'var acct=document.getElementById("account").value;' +
    'var cat=document.getElementById("category").value.trim();' +
    'var desc=document.getElementById("description").value.trim();' +
    'var amt=parseFloat(document.getElementById("amount").value);' +
    'if(isNaN(amt)||amt<0){alert("Enter a valid amount.");return;}' +
    'if(selectedType==="Adjust"){' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message);}).saveAdjustment(acct,amt,desc);' +
    '}else if(selectedType==="Transfer"){' +
    'if(amt<=0){alert("Enter a valid amount.");return;}' +
    'var to=document.getElementById("toAccount").value;' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message);}).saveTransfer(acct,to,amt,desc);' +
    '}else{' +
    'if(amt<=0){alert("Enter a valid amount.");return;}' +
    'if(!cat){alert("Enter a category.");return;}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message);}).saveTransaction(acct,cat,desc,amt,selectedType);}' +
    '}</script></body></html>'
  ).setWidth(440).setHeight(520).setTitle('Add Transaction');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Transaction');
}

function saveTransaction(accountName, category, desc, amount, type) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(accountName);
  if (!sheet) throw new Error('Account not found: ' + accountName);

  const fmtMap = {
    MYR: '"RM "#,##0.00', USD: '"$"#,##0.00',
    SGD: '"S$"#,##0.00',  HKD: '"HK$"#,##0.00',
    RMB: '"¥"#,##0.00',   CNY: '"¥"#,##0.00'
  };
  const currency = sheet.getRange('G2').getValue() || 'MYR';
  const fmt      = fmtMap[currency] || fmtMap['MYR'];
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const lastRow  = Math.max(sheet.getLastRow(), 2) + 1;
  const prevBal  = lastRow === 3 ? 0 : 'F' + (lastRow - 1);
  const balFormula = lastRow === 3
    ? (type === 'IN' ? '=D3' : '=-D3')
    : (type === 'IN' ? '=' + prevBal + '+D' + lastRow : '=' + prevBal + '-D' + lastRow);

  sheet.setRowHeight(lastRow, 30);
  const bg = lastRow % 2 === 0 ? '#f8f9fa' : '#ffffff';
  sheet.getRange(lastRow, 1, 1, 6).setBackground(bg);
  sheet.getRange(lastRow, 1).setValue(today);
  sheet.getRange(lastRow, 2).setValue(category);
  sheet.getRange(lastRow, 3).setValue(desc);
  sheet.getRange(lastRow, 4).setValue(amount).setNumberFormat(fmt);
  sheet.getRange(lastRow, 5).setValue(type)
    .setFontColor(type === 'IN' ? '#0f9d58' : '#d93025').setFontWeight('bold');
  sheet.getRange(lastRow, 6).setFormula(balFormula).setNumberFormat(fmt);
}

function saveTransfer(fromName, toName, amount, desc) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const fromSheet = ss.getSheetByName(fromName);
  const toSheet   = ss.getSheetByName(toName);
  if (!fromSheet) throw new Error('From account not found: ' + fromName);
  if (!toSheet)   throw new Error('To account not found: '   + toName);

  const fromCcy   = fromSheet.getRange('G2').getValue() || 'MYR';
  const toCcy     = toSheet.getRange('G2').getValue()   || 'MYR';
  let toAmount    = amount;
  let fxNote      = '';

  if (normCurrency_(fromCcy) !== normCurrency_(toCcy)) {
    // Need FX rate
    try {
      const pair  = 'CURRENCY:' + normCurrency_(fromCcy) + normCurrency_(toCcy);
      const temp  = ss.insertSheet('_FX_TEMP_');
      compactSheet_(temp, 10, 1); // 1 col × 10 rows
      temp.getRange('A1').setFormula('=GOOGLEFINANCE("' + pair + '")');
      SpreadsheetApp.flush();
      Utilities.sleep(2000);
      const rate = parseFloat(temp.getRange('A1').getValue());
      ss.deleteSheet(temp);
      if (!isNaN(rate) && rate > 0) {
        toAmount = amount * rate;
        fxNote   = ' [1 ' + fromCcy + ' = ' + rate.toFixed(4) + ' ' + toCcy + ']';
      } else {
        fxNote = ' [FX rate unavailable]';
      }
    } catch(e) {
      try { ss.deleteSheet(ss.getSheetByName('_FX_TEMP_')); } catch(_) {}
      fxNote = ' [FX error]';
    }
  }

  const note = (desc ? desc : 'Transfer: ' + fromName + ' → ' + toName) + fxNote;
  saveTransaction(fromName, 'Transfer', note, amount,   'OUT');
  saveTransaction(toName,   'Transfer', note, toAmount, 'IN');
}


// ── ACCOUNT HELPERS ──────────────────────────────────────────

/**
 * Adjust: sets the account balance to `newBalance` by writing a single
 * adjustment row. The row records the difference vs. the current balance
 * so the running balance formula resolves to exactly `newBalance`.
 *
 * The row is styled distinctly (orange) so it's easy to spot in history.
 */
function saveAdjustment(accountName, newBalance, desc) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(accountName);
  if (!sheet) throw new Error('Account not found: ' + accountName);

  const fmtMap = {
    MYR: '"RM "#,##0.00', USD: '"$"#,##0.00',
    SGD: '"S$"#,##0.00',  HKD: '"HK$"#,##0.00',
    RMB: '"¥"#,##0.00',   CNY: '"¥"#,##0.00'
  };
  const currency = sheet.getRange('G2').getValue() || 'MYR';
  const fmt      = fmtMap[currency] || fmtMap['MYR'];
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // Get current balance (last row col F)
  const lastRow   = sheet.getLastRow();
  const curBal    = lastRow >= 3 ? (sheet.getRange(lastRow, 6).getValue() || 0) : 0;
  const diff      = newBalance - curBal;
  const adjType   = diff >= 0 ? 'IN' : 'OUT';
  const adjAmount = Math.abs(diff);

  const note       = desc || 'Balance adjustment';
  const newRow     = Math.max(lastRow, 2) + 1;
  const prevBalRef = newRow === 3 ? null : 'F' + (newRow - 1);

  // Balance formula resolves to exactly newBalance
  // = prevBal + adjAmount  (IN)  or  = prevBal - adjAmount  (OUT)
  const balFormula = prevBalRef === null
    ? (adjType === 'IN' ? '=D' + newRow : '=-D' + newRow)
    : (adjType === 'IN'
        ? '=' + prevBalRef + '+D' + newRow
        : '=' + prevBalRef + '-D' + newRow);

  sheet.setRowHeight(newRow, 30);
  // Orange tint so adjustments stand out in history
  sheet.getRange(newRow, 1, 1, 6).setBackground('#fff3e0');
  sheet.getRange(newRow, 1).setValue(today);
  sheet.getRange(newRow, 2).setValue('Adjustment');
  sheet.getRange(newRow, 3).setValue(note);
  // If diff is 0 write 0 amount; formula still resolves correctly
  sheet.getRange(newRow, 4).setValue(adjAmount === 0 ? 0 : adjAmount).setNumberFormat(fmt);
  sheet.getRange(newRow, 5).setValue('⚖️ ' + adjType)
    .setFontColor('#e65100').setFontWeight('bold');
  sheet.getRange(newRow, 6).setFormula(balFormula).setNumberFormat(fmt);
}

function getAccountNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => s.getName());
}

function getAccountBalances() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => {
      const lastRow = s.getLastRow();
      const balance = lastRow >= 3 ? (s.getRange(lastRow, 6).getValue() || 0) : 0;
      return { name: s.getName(), balance: balance, currency: s.getRange('G2').getValue() || 'MYR' };
    });
}

function compactSheet_(sheet, maxRows, maxCols) {
  if (sheet.getMaxRows()    > maxRows)  sheet.deleteRows(maxRows + 1,    sheet.getMaxRows()    - maxRows);
  if (sheet.getMaxColumns() > maxCols)  sheet.deleteColumns(maxCols + 1, sheet.getMaxColumns() - maxCols);
}

function normCurrency_(c) {
  if (!c) return 'MYR';
  return c.toString().trim().toUpperCase() === 'RMB' ? 'CNY' : c.toString().trim().toUpperCase();
}

// ============================================================
//  MUTUAL FUND / UNIT TRUST TRACKER
//  Sheet: 🏦 Mutual Funds
// ============================================================

function searchFSMFunds_(keyword) {
  // Primary: search FIMM NAV Cache sheet
  try {
    const kw = keyword.toString().trim().toUpperCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheet = ss.getSheetByName(FIMM_NAV_SHEET);
    if (cacheSheet) {
      const lastRow = cacheSheet.getLastRow();
      if (lastRow >= 3) {
        const data = cacheSheet.getRange(3, 1, lastRow - 2, 2).getValues();
        const results = [];
        for (var i = 0; i < data.length; i++) {
          const code = data[i][0].toString().toUpperCase();
          const name = data[i][1].toString();
          if (code.includes(kw) || name.toUpperCase().includes(kw)) {
            results.push({ code: code, name: name, source: 'FIMM' });
            if (results.length >= 15) break;
          }
        }
        if (results.length > 0) return results;
      }
    }
  } catch(e2) { /* fall through */ }

  // Fallback: FSMOne search API
  try {
    const url = FSM_SEARCH_URL + encodeURIComponent(keyword) + '&page=0&limit=30';
    const res = UrlFetchApp.fetch(url, {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Referer': FSM_BASE }
    });
    if (res.getResponseCode() !== 200) return [];
    const json = JSON.parse(res.getContentText());
    if (json.status !== 'SUCCESS') return [];
    return Object.values(json.data || {})
      .filter(i => i.productType === 'UT')
      .map(i => ({ code: i.productCode, name: i.productName, source: 'FSMOne' }));
  } catch(e) { return []; }
}

function fetchFSMFundData_(fundCode) {
  const code = fundCode.toString().trim().toUpperCase();

  // Primary: look up FIMM NAV Cache sheet
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheet = ss.getSheetByName(FIMM_NAV_SHEET);
    if (cacheSheet) {
      const lastRow = cacheSheet.getLastRow();
      if (lastRow >= 3) {
        const data = cacheSheet.getRange(3, 1, lastRow - 2, 4).getValues();
        for (var i = 0; i < data.length; i++) {
          if (data[i][0].toString().toUpperCase() === code) {
            const nav = parseFloat(data[i][2]);
            if (!isNaN(nav) && nav > 0) {
              return { nav: nav, currency: 'MYR', date: new Date(), source: 'FIMM' };
            }
          }
        }
      }
    }
  } catch(e2) { /* fall through */ }

  // Fallback: FSMOne factsheet API
  try {
    const res = UrlFetchApp.fetch(FSM_FACTSHEET_URL + encodeURIComponent(fundCode), {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Referer': FSM_BASE }
    });
    if (res.getResponseCode() !== 200) return null;
    const json = JSON.parse(res.getContentText());
    if (json.status !== 'SUCCESS') return null;
    const d = json.data;
    return {
      nav:      parseFloat(d.latestNavPrice.navPrice),
      currency: d.fundCurrencyCode || 'MYR',
      date:     new Date(d.latestNavPrice.dailyPricePk.showDate),
      source:   'FSMOne'
    };
  } catch(e) { return null; }
}

function searchFundsForDialog(keyword) {
  return JSON.stringify(searchFSMFunds_(keyword).slice(0, 10));
}

function getFundDataForDialog(fundCode) {
  const data = fetchFSMFundData_(fundCode);
  // No data — cache empty and FSMOne blocked
  if (!data) return JSON.stringify({ navUnavailable: true, currency: 'MYR', date: '', source: '' });
  var dateStr = '';
  try { dateStr = Utilities.formatDate(data.date, Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch(_) {}
  return JSON.stringify({
    nav:           (data.nav && data.nav > 0) ? data.nav : null,
    navUnavailable: (!data.nav || data.nav <= 0),
    currency:      data.currency || 'MYR',
    date:          dateStr,
    source:        data.source || 'FIMM'
  });
}

function getFimmCacheStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheet = ss.getSheetByName(FIMM_NAV_SHEET);
    if (!cacheSheet) return JSON.stringify({ count: 0, updated: '' });
    const lastRow = cacheSheet.getLastRow();
    if (lastRow < 3) return JSON.stringify({ count: 0, updated: '' });
    const subtitle = cacheSheet.getRange(1, 1).getValue().toString();
    const parts = subtitle.split('Last updated:');
    const updated = parts.length > 1 ? parts[1].trim() : '';
    return JSON.stringify({ count: lastRow - 2, updated: updated });
  } catch(e) { return JSON.stringify({ count: 0, updated: '' }); }
}

function ensureFimmNavCacheSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(FIMM_NAV_SHEET)) return false;
  const sheet = ss.insertSheet(FIMM_NAV_SHEET);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(ss.getNumSheets());
  sheet.setRowHeight(1, 36);
  sheet.getRange(1, 1, 1, 5).setValues([['Fund Code', 'Fund Name', 'NAV', 'Updated', 'Source']])
    .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(11).setVerticalAlignment('middle');
  sheet.setColumnWidth(1, 120); sheet.setColumnWidth(2, 380);
  sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 140); sheet.setColumnWidth(5, 80);
  sheet.getRange(2, 1).setValue('Cache is empty — use "Refresh FIMM NAV Cache Now" from the Finance menu to populate.');
  return true;
}

function createMutualFundSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Ensure FIMM NAV Cache sheet exists
  const cacheCreated = ensureFimmNavCacheSheet_();

  // 2. Create or reuse MF sheet
  let sheet = ss.getSheetByName(MF_SHEET_NAME);
  if (sheet) {
    ui.alert('The Mutual Fund sheet already exists.\n\nClick the "🏦 Mutual Funds" tab to view it.');
    sheet.activate();
  } else {
    sheet = ss.insertSheet(MF_SHEET_NAME, 1);
    buildMutualFundSheet_(sheet);
    sheet.activate();
    if (cacheCreated) {
      ui.alert('✅ Mutual Fund & FIMM NAV Cache sheets created!\n\nRun "Refresh FIMM NAV Cache Now" from the Finance menu to populate fund data before searching.\n\nOpening Add Fund dialog…');
    } else {
      ui.alert('✅ Mutual Fund sheet created!\n\nOpening Add Fund dialog…');
    }
  }

  // 3. Show Add Fund dialog
  showAddFundDialog();
}

function buildMutualFundSheet_(sheet) {
  const colWidths = [240,110,110,110,110,120,120,90,60,160,120,180];
  colWidths.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  sheet.setRowHeight(1, 52);
  sheet.getRange(1,1,1,12).merge()
    .setValue('🏦  MUTUAL FUND / UNIT TRUST PORTFOLIO')
    .setBackground('#0d47a1').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(2, 26);
  sheet.getRange(2,1,1,12).merge()
    .setValue('NAV prices via FSMOne  ·  All values in fund currency')
    .setBackground('#e8f0fe').setFontColor('#5f6368')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(3, 32);
  const headers = ['Fund Name','Code','Units','Avg Buy NAV','Current NAV',
    'Market Value','Gain / Loss','Gain %','CCY','Linked Account','Last Updated','Notes'];
  headers.forEach((h,i) => {
    sheet.getRange(3,i+1)
      .setValue(h).setBackground('#1a73e8').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sheet.setFrozenRows(3);
  sheet.setHiddenGridlines(true);
}

function getLastFundRow_(sheet) {
  // Valid fund row has CODE in col 2 AND numeric UNITS in col 3
  // Summary rows have merged text in cols 1-3, so col 3 won't be a positive number
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 3) return 3;
  const data = sheet.getRange(4, MF_COLS.CODE, maxRows-3, 2).getValues();
  let last = 3;
  for (let i = 0; i < data.length; i++) {
    const code  = data[i][0]; // MF_COLS.CODE
    const units = data[i][1]; // MF_COLS.UNITS
    if (code !== '' && typeof units === 'number' && units > 0) last = i + 4;
  }
  return last;
}

function showAddFundDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(MF_SHEET_NAME)) {
    const resp = ui.alert('No Mutual Fund sheet found.', 'Create it now?', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.YES) createMutualFundSheet(); else return;
    return; // createMutualFundSheet() will call showAddFundDialog() itself
  }
  const allAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }));
  const accountsJson = JSON.stringify(allAccounts);

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#1a73e8}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.search-row{display:flex;gap:6px}' +
    '.search-row input{flex:1}' +
    '.search-btn{padding:7px 14px;background:#1a73e8;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;white-space:nowrap}' +
    '.results-list{border:1px solid #dadce0;border-radius:6px;max-height:140px;overflow-y:auto;background:#fff;margin-top:4px;display:none}' +
    '.result-item{padding:8px 10px;cursor:pointer;border-bottom:1px solid #f1f3f4;font-size:12px}' +
    '.result-item:hover{background:#e8f0fe}' +
    '.result-item .code{color:#1a73e8;font-weight:bold;margin-right:6px}' +
    '.enter-manual-link{color:#1a73e8;cursor:pointer;text-decoration:underline;margin-left:6px;font-size:12px}' +
    '.manual-form{background:#fff;border:1px solid #dadce0;border-radius:6px;padding:12px;margin-top:6px;display:none}' +
    '.manual-form-title{font-size:12px;font-weight:600;color:#3c4043;margin-bottom:8px}' +
    '.manual-confirm-btn{margin-top:10px;padding:6px 16px;background:#1a73e8;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px;font-weight:600}' +
    '.selected-fund{background:#e8f4fd;border:1px solid #1a73e8;border-radius:6px;padding:8px 10px;margin-top:4px;font-size:12px;display:none}' +
    '.nav-badge{display:inline-block;background:#e6f4ea;color:#137333;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:bold;margin-left:6px}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.cost-preview{background:#e8f4fd;border-radius:6px;padding:8px 10px;margin-top:8px;font-size:12px;color:#1a73e8;display:none}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#1a73e8;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '.no-acct-warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px;display:none}' +
    '.spinner{display:none;color:#5f6368;font-size:12px;margin-top:4px}' +
    '</style></head><body>' +
    '<h2>Add Mutual Fund / Unit Trust</h2>' +
    '<div id="cacheStatus" style="background:#e8f5e9;border:1px solid #a5d6a7;border-radius:6px;padding:6px 10px;font-size:11px;color:#2e7d32;margin-bottom:10px">&#x231B; Checking FIMM NAV cache...</div>' +
    '<label>Search Fund by Name or Code</label>' +
    '<div class="search-row">' +
    '<input type="text" id="searchInput" placeholder="e.g. RHB Cash Management or MANUAG" />' +
    '<button class="search-btn" onclick="doSearch()">Search</button></div>' +
    '<div class="spinner" id="spinner">Searching FIMM cache...</div>' +
    '<div class="results-list" id="resultsList"></div>' +
    '<div class="manual-form" id="manualForm">' +
    '<div class="manual-form-title">Enter Fund Details Manually</div>' +
    '<label>Fund Name</label>' +
    '<input type="text" id="manualName" placeholder="e.g. RHB Cash Management Fund" />' +
    '<label>Fund Code</label>' +
    '<input type="text" id="manualCode" placeholder="e.g. RHBCASH" />' +
    '<label>Currency</label>' +
    '<select id="manualCurrency">' +
    '<option value="MYR">MYR – Malaysian Ringgit</option>' +
    '<option value="USD">USD – US Dollar</option>' +
    '<option value="SGD">SGD – Singapore Dollar</option>' +
    '<option value="HKD">HKD – Hong Kong Dollar</option>' +
    '</select>' +
    '<div style="text-align:right"><button class="manual-confirm-btn" onclick="confirmManual()">Confirm</button></div>' +
    '</div>' +
    '<div class="selected-fund" id="selectedFund"></div>' +
    '<div id="mainForm" style="display:none">' +
    '<div class="row2">' +
    '<div><label>Units Purchased</label>' +
    '<input type="number" id="units" placeholder="e.g. 1000.00" step="0.01" min="0.01" oninput="updateCost()" /></div>' +
    '<div><label>Buy NAV (per unit)</label>' +
    '<input type="number" id="buyNav" placeholder="auto-filled" step="0.0001" oninput="updateCost()" /></div></div>' +
    '<label>Linked Account (deduct purchase cost)</label>' +
    '<select id="linkedAccount"><option value="">— None —</option></select>' +
    '<div class="no-acct-warn" id="noAcctWarn"></div>' +
    '<div class="cost-preview" id="costPreview"></div>' +
    '<label>Notes (optional)</label>' +
    '<input type="text" id="notes" placeholder="e.g. Monthly RSP" /></div>' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="saveBtn" onclick="save()" style="display:none">Add Fund</button></div>' +
    '<script>' +
    'var ALL_ACCOUNTS=' + accountsJson + ';' +
    'google.script.run.withSuccessHandler(function(j){' +
    '  var d=JSON.parse(j);var el=document.getElementById("cacheStatus");' +
    '  if(d.count>0){el.textContent="\u{1F4CB} FIMM NAV Cache: "+d.count+" funds loaded"+(d.updated?" \u00b7 Last updated: "+d.updated:"");}' +
    '  else{el.style.background="#fff3e0";el.style.color="#e65100";el.style.borderColor="#ffcc80";' +
    '  el.textContent="\u26A0\uFE0F FIMM NAV cache is empty \u2014 run Refresh FIMM NAV Cache Now first.";}' +
    '}).getFimmCacheStatus();' +
    'var selectedCode="",selectedName="",selectedCurrency="",selectedNav=0;' +
    'function doSearch(){' +
    '  var kw=document.getElementById("searchInput").value.trim();if(!kw)return;' +
    '  document.getElementById("spinner").style.display="block";' +
    '  document.getElementById("spinner").textContent="Searching FIMM cache...";' +
    '  document.getElementById("resultsList").style.display="none";' +
    '  document.getElementById("manualForm").style.display="none";' +
    '  google.script.run.withSuccessHandler(showResults)' +
    '  .withFailureHandler(function(e){document.getElementById("spinner").style.display="none";alert("Search error: "+e.message);})' +
    '  .searchFundsForDialog(kw);}' +
    'function showResults(json){' +
    '  document.getElementById("spinner").style.display="none";' +
    '  var results=JSON.parse(json);' +
    '  var list=document.getElementById("resultsList");' +
    '  if(!results.length){' +
    '    list.innerHTML=\'<div class="result-item" style="cursor:default">No funds found in cache. \' +' +
    '    \'<span class="enter-manual-link" onclick="enterManually()">Enter manually</span></div>\';' +
    '    list.style.display="block";return;}' +
    '  list.innerHTML=results.map(function(r){' +
    '    return "<div class=\\"result-item\\" data-code=\\""+r.code+"\\" data-name=\\""+r.name.replace(/"/g,"&quot;")+"\\" onclick=\\"pickResult(this)\\"><span class=\\"code\\">"+r.code+"</span>"+r.name+"</div>";' +
    '  }).join("");' +
    '  list.style.display="block";}' +
    'function enterManually(){' +
    '  document.getElementById("resultsList").style.display="none";' +
    '  document.getElementById("manualForm").style.display="block";' +
    '  document.getElementById("manualName").focus();}' +
    'function confirmManual(){' +
    '  var name=document.getElementById("manualName").value.trim();' +
    '  var code=document.getElementById("manualCode").value.trim().toUpperCase();' +
    '  var currency=document.getElementById("manualCurrency").value;' +
    '  if(!name){alert("Enter a fund name.");return;}' +
    '  if(!code){alert("Enter a fund code.");return;}' +
    '  selectedCode=code;selectedName=name;selectedCurrency=currency;selectedNav=0;' +
    '  document.getElementById("manualForm").style.display="none";' +
    '  var sf=document.getElementById("selectedFund");' +
    '  sf.innerHTML="<strong>"+name+"</strong>  <span style=\'color:#5f6368\'>"+code+"</span>" +' +
    '    "  <span style=\'background:#fff3e0;color:#e65100;border:1px solid #ffcc80;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:600;margin-left:6px\'>Manual entry</span>";' +
    '  sf.style.display="block";' +
    '  document.getElementById("buyNav").value="";' +
    '  document.getElementById("buyNav").placeholder="Enter buy NAV manually";' +
    '  document.getElementById("mainForm").style.display="block";' +
    '  document.getElementById("saveBtn").style.display="inline-block";' +
    '  buildAccountDropdown(currency);updateCost();}' +
    'function pickResult(el){selectFund(el.getAttribute("data-code"),el.getAttribute("data-name"));}' +
    'function selectFund(code,name){' +
    '  document.getElementById("resultsList").style.display="none";' +
    '  document.getElementById("spinner").style.display="block";' +
    '  document.getElementById("spinner").textContent="Fetching NAV from FIMM cache...";' +
    '  selectedCode=code;selectedName=name;' +
    '  google.script.run.withSuccessHandler(onFundData)' +
    '  .withFailureHandler(function(e){document.getElementById("spinner").style.display="none";alert("Could not fetch fund data: "+e.message);})' +
    '  .getFundDataForDialog(code);}' +
    'function onFundData(json){' +
    '  document.getElementById("spinner").style.display="none";' +
    '  var fd=JSON.parse(json);' +
    '  selectedCurrency=fd.currency||"MYR";selectedNav=fd.nav||0;' +
    '  var sf=document.getElementById("selectedFund");' +
    '  var navBadge;' +
    '  if(fd.navUnavailable||!fd.nav){' +
    '    navBadge="<span style=\'background:#fff3e0;color:#e65100;border:1px solid #ffcc80;border-radius:4px;padding:2px 8px;font-size:11px;font-weight:600;margin-left:8px\'>\u26a0\ufe0f NAV unavailable \u2014 enter manually</span>";' +
    '    document.getElementById("buyNav").value="";' +
    '    document.getElementById("buyNav").placeholder="Enter buy NAV manually";' +
    '  }else{' +
    '    navBadge="<span class=\'nav-badge\'>NAV: "+selectedCurrency+" "+fd.nav.toFixed(4)+"</span>"+(fd.date?"  <span style=\'color:#9aa0a6;font-size:11px\'>as of "+fd.date+"</span>":"");' +
    '    document.getElementById("buyNav").value=fd.nav.toFixed(4);' +
    '  }' +
    '  sf.innerHTML="<strong>"+selectedName+"</strong>  <span style=\'color:#5f6368\'>"+selectedCode+"</span>"+navBadge;' +
    '  sf.style.display="block";' +
    '  document.getElementById("mainForm").style.display="block";' +
    '  document.getElementById("saveBtn").style.display="inline-block";' +
    '  buildAccountDropdown(selectedCurrency);updateCost();}' +
    'function buildAccountDropdown(currency){' +
    '  var sel=document.getElementById("linkedAccount");' +
    '  var warn=document.getElementById("noAcctWarn");' +
    '  var matching=ALL_ACCOUNTS.filter(function(a){return a.currency===currency||(currency==="MYR"&&a.currency==="MYR");});' +
    '  sel.innerHTML=\'<option value="">— None (no deduction) —</option>\';' +
    '  matching.forEach(function(a){sel.innerHTML+=\'<option value="\'+a.name+\'">\'+a.name+" ("+a.currency+")</option>";});' +
    '  if(matching.length===0){warn.textContent="No "+currency+" accounts found. Purchase cost won\'t be deducted.";warn.style.display="block";}' +
    '  else{warn.style.display="none";}}' +
    'function updateCost(){' +
    '  var units=parseFloat(document.getElementById("units").value)||0;' +
    '  var nav=parseFloat(document.getElementById("buyNav").value)||0;' +
    '  var prev=document.getElementById("costPreview");' +
    '  if(units>0&&nav>0){prev.textContent="Total cost: "+(selectedCurrency||"MYR")+" "+(units*nav).toFixed(2);prev.style.display="block";}' +
    '  else{prev.style.display="none";}}' +
    'function save(){' +
    '  var units=parseFloat(document.getElementById("units").value);' +
    '  if(!selectedCode){alert("Please search and select a fund first.");return;}' +
    '  var buyNav=parseFloat(document.getElementById("buyNav").value);' +
    '  if(!units||units<=0){alert("Enter a valid number of units.");return;}' +
    '  if(!buyNav||buyNav<=0){alert("Enter a valid buy NAV.");return;}' +
    '  var account=document.getElementById("linkedAccount").value;' +
    '  var notes=document.getElementById("notes").value.trim();' +
    '  document.getElementById("saveBtn").disabled=true;' +
    '  document.getElementById("saveBtn").textContent="Saving...";' +
    '  google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '  .withFailureHandler(function(e){document.getElementById("saveBtn").disabled=false;document.getElementById("saveBtn").textContent="Add Fund";alert("Error: "+e.message);})' +
    '  .saveFund(selectedCode,selectedName,units,buyNav,selectedCurrency,account,notes);}' +
    'document.getElementById("searchInput").addEventListener("keydown",function(e){if(e.key==="Enter")doSearch();});' +
    '</script></body></html>'
  ).setWidth(520).setHeight(600).setTitle('Add Mutual Fund');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Mutual Fund');
}

function saveFund(fundCode, fundName, units, buyNav, currency, linkedAccount, notes) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let sheet  = ss.getSheetByName(MF_SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(MF_SHEET_NAME, 1); buildMutualFundSheet_(sheet); }

  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = getLastFundRow_(sheet);
  const row     = lastRow + 1;

  const fmtMap = { MYR:'"RM "#,##0.0000', USD:'"$"#,##0.0000', SGD:'"S$"#,##0.0000', HKD:'"HK$"#,##0.0000', CNY:'"¥"#,##0.0000', RMB:'"¥"#,##0.0000' };
  const mktFmt = { MYR:'"RM "#,##0.00',   USD:'"$"#,##0.00',   SGD:'"S$"#,##0.00',   HKD:'"HK$"#,##0.00',   CNY:'"¥"#,##0.00',   RMB:'"¥"#,##0.00'   };
  const navFmt  = fmtMap[currency] || fmtMap['MYR'];
  const valFmt  = mktFmt[currency] || mktFmt['MYR'];

  sheet.setRowHeight(row, 34);
  const bg = row % 2 === 0 ? '#f8f9fa' : '#ffffff';
  sheet.getRange(row, 1, 1, 13).setBackground(bg);

  sheet.getRange(row, MF_COLS.NAME   ).setValue(fundName).setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.CODE   ).setValue(fundCode).setFontSize(10).setFontColor('#1a73e8').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.UNITS  ).setValue(units).setNumberFormat('#,##0.0000').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.BUY_NAV).setValue(buyNav).setNumberFormat(navFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.CUR_NAV).setFormula('=IFERROR(FSMFund("' + fundCode + '"),D' + row + ')').setNumberFormat(navFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.MKT_VAL).setFormula('=C' + row + '*E' + row).setNumberFormat(valFmt).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.GAIN   ).setFormula('=F' + row + '-C' + row + '*D' + row).setNumberFormat(valFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.GAIN_PCT).setFormula('=IF(D' + row + '*C' + row + '<>0,(E' + row + '-D' + row + ')/D' + row + '*100,0)').setNumberFormat('0.00"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.CCY    ).setValue(currency).setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.ACCOUNT).setValue(linkedAccount || '').setFontSize(10).setFontColor('#5f6368').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.UPDATED).setValue(today).setFontSize(9).setFontColor('#9aa0a6').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, MF_COLS.NOTES  ).setValue(notes || '').setFontSize(10).setFontColor('#9aa0a6').setVerticalAlignment('middle');

  SpreadsheetApp.flush();
  const gainVal = sheet.getRange(row, MF_COLS.GAIN).getValue();
  const glColor = (typeof gainVal === 'number' && gainVal < 0) ? '#d93025' : '#0f9d58';
  sheet.getRange(row, MF_COLS.GAIN    ).setFontColor(glColor);
  sheet.getRange(row, MF_COLS.GAIN_PCT).setFontColor(glColor);

  if (linkedAccount) {
    const accSheet   = ss.getSheetByName(linkedAccount);
    if (!accSheet) throw new Error('Account "' + linkedAccount + '" not found.');
    const accCurrency = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCurrency !== normCurrency_(currency)) throw new Error('Currency mismatch: account is ' + accCurrency + ' but fund is ' + currency + '.');
    const cost       = units * buyNav;
    const accFmt     = mktFmt[accCurrency] || mktFmt['MYR'];
    const lastAcc    = accSheet.getLastRow() + 1;
    const balFormula = lastAcc === 2 ? '=D' + lastAcc : '=F' + (lastAcc-1) + '-D' + lastAcc;
    accSheet.getRange(lastAcc, 1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc, 2).setValue('Unit Trust');
    accSheet.getRange(lastAcc, 3).setValue('Buy ' + fundCode + ' x' + units.toFixed(2) + ' @ ' + currency + ' ' + buyNav.toFixed(4));
    accSheet.getRange(lastAcc, 4).setValue(cost).setNumberFormat(accFmt);
    accSheet.getRange(lastAcc, 5).setValue('OUT').setFontColor('#d93025').setFontWeight('bold');
    accSheet.getRange(lastAcc, 6).setFormula(balFormula).setNumberFormat(accFmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc, 1, 1, 6).setBackground('#f8f9fa');
  }

  refreshMutualFundSummary_(sheet);
  sheet.activate();
}

function showSellFundDialog() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MF_SHEET_NAME);
  if (!sheet || getLastFundRow_(sheet) < 4) { ui.alert('No funds found. Add a fund first via "Add Fund".'); return; }

  const lastRow  = getLastFundRow_(sheet);
  const holdings = sheet.getRange(4, 1, lastRow-3, 12).getValues()
    .filter(r => r[MF_COLS.CODE-1] !== '')
    .map(r => ({
      name:     r[MF_COLS.NAME-1],    code:    r[MF_COLS.CODE-1],
      units:    r[MF_COLS.UNITS-1],   buyNav:  r[MF_COLS.BUY_NAV-1],
      curNav:   r[MF_COLS.CUR_NAV-1], currency:r[MF_COLS.CCY-1],
      account:  r[MF_COLS.ACCOUNT-1]
    }));
  if (!holdings.length) { ui.alert('No active fund holdings found.'); return; }

  const allAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }));

  const holdingsJson = JSON.stringify(holdings);
  const accountsJson = JSON.stringify(allAccounts);
  const optionsHtml  = holdings.map((h,i) =>
    '<option value="' + i + '">' + h.name + ' (' + h.units.toFixed(4) + ' units)</option>'
  ).join('');

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#c62828}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.holding-info{background:#fff8e1;border:1px solid #fbc02d;border-radius:6px;padding:8px 10px;font-size:12px;margin-top:4px;display:none}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.proceeds-preview{background:#e8f5e9;border-radius:6px;padding:8px 10px;margin-top:8px;font-size:12px;color:#2e7d32;display:none}' +
    '.no-acct-warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px;display:none}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#c62828;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>Sell Mutual Fund</h2>' +
    '<label>Select Fund to Sell</label>' +
    '<select id="fundSelect" onchange="onFundChange()">' +
    '<option value="">— Select a fund —</option>' + optionsHtml + '</select>' +
    '<div class="holding-info" id="holdingInfo"></div>' +
    '<div id="sellForm" style="display:none">' +
    '<div class="row2">' +
    '<div><label>Units to Sell</label>' +
    '<input type="number" id="sellUnits" placeholder="Max: 0" step="0.0001" min="0.0001" oninput="updateProceeds()" /></div>' +
    '<div><label>Sell NAV (per unit)</label>' +
    '<input type="number" id="sellNav" placeholder="Current NAV" step="0.0001" oninput="updateProceeds()" /></div></div>' +
    '<label>Return Proceeds to Account</label>' +
    '<select id="returnAccount"><option value="">— None —</option></select>' +
    '<div class="no-acct-warn" id="noAcctWarn"></div>' +
    '<div class="proceeds-preview" id="proceedsPreview"></div>' +
    '<label>Notes (optional)</label>' +
    '<input type="text" id="notes" placeholder="e.g. Profit taking" /></div>' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="saveBtn" onclick="save()" style="display:none">Confirm Sell</button></div>' +
    '<script>' +
    'var HOLDINGS=' + holdingsJson + ';' +
    'var ALL_ACCOUNTS=' + accountsJson + ';' +
    'var selIdx=-1;' +
    'function onFundChange(){' +
    'selIdx=parseInt(document.getElementById("fundSelect").value);' +
    'if(isNaN(selIdx)){document.getElementById("sellForm").style.display="none";document.getElementById("holdingInfo").style.display="none";return;}' +
    'var h=HOLDINGS[selIdx];' +
    'var info=document.getElementById("holdingInfo");' +
    'info.innerHTML="<b>"+h.name+"</b> | Held: <b>"+h.units.toFixed(4)+" units</b> | Avg Buy NAV: <b>"+h.currency+" "+h.buyNav.toFixed(4)+"</b> | Current NAV: <b>"+h.currency+" "+h.curNav.toFixed(4)+"</b>";' +
    'info.style.display="block";' +
    'document.getElementById("sellUnits").max=h.units;' +
    'document.getElementById("sellUnits").placeholder="Max: "+h.units.toFixed(4);' +
    'document.getElementById("sellNav").value=h.curNav.toFixed(4);' +
    'buildAccountDropdown(h.currency,h.account);' +
    'document.getElementById("sellForm").style.display="block";' +
    'document.getElementById("saveBtn").style.display="inline-block";updateProceeds();}' +
    'function buildAccountDropdown(currency,preferred){' +
    'var sel=document.getElementById("returnAccount");' +
    'var warn=document.getElementById("noAcctWarn");' +
    'var matching=ALL_ACCOUNTS.filter(function(a){return a.currency===currency;});' +
    'sel.innerHTML=\'<option value="">— None (no credit) —</option>\';' +
    'matching.forEach(function(a){sel.innerHTML+=\'<option value="\'+a.name+\'"\'+(a.name===preferred?" selected":"")+\'>\'+a.name+" ("+a.currency+")</option>";});' +
    'if(matching.length===0){warn.textContent="No "+currency+" accounts found.";warn.style.display="block";}' +
    'else{warn.style.display="none";}}' +
    'function updateProceeds(){' +
    'if(selIdx<0)return;' +
    'var h=HOLDINGS[selIdx];' +
    'var units=parseFloat(document.getElementById("sellUnits").value)||0;' +
    'var nav=parseFloat(document.getElementById("sellNav").value)||0;' +
    'var prev=document.getElementById("proceedsPreview");' +
    'if(units>0&&nav>0){' +
    'var proceeds=(units*nav).toFixed(2);' +
    'var gain=((nav-h.buyNav)*units).toFixed(2);' +
    'var gs=parseFloat(gain)>=0?"+":"";' +
    'prev.innerHTML="Proceeds: <b>"+h.currency+" "+proceeds+"</b> | Gain/Loss vs buy: <b>"+gs+h.currency+" "+gain+"</b>";' +
    'prev.style.display="block";}else{prev.style.display="none";}}' +
    'function save(){' +
    'if(selIdx<0){alert("Select a fund.");return;}' +
    'var h=HOLDINGS[selIdx];' +
    'var sellUnits=parseFloat(document.getElementById("sellUnits").value);' +
    'var sellNav=parseFloat(document.getElementById("sellNav").value);' +
    'if(!sellUnits||sellUnits<=0){alert("Enter units to sell.");return;}' +
    'if(sellUnits>h.units+0.00001){alert("Cannot sell more than you hold ("+h.units.toFixed(4)+" units).");return;}' +
    'if(!sellNav||sellNav<=0){alert("Enter a valid sell NAV.");return;}' +
    'var returnAcct=document.getElementById("returnAccount").value;' +
    'var notes=document.getElementById("notes").value.trim();' +
    'document.getElementById("saveBtn").disabled=true;' +
    'document.getElementById("saveBtn").textContent="Processing...";' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){document.getElementById("saveBtn").disabled=false;document.getElementById("saveBtn").textContent="Confirm Sell";alert("Error: "+e.message);})' +
    '.saveSellFund(selIdx,sellUnits,sellNav,returnAcct,notes);}' +
    '</script></body></html>'
  ).setWidth(520).setHeight(480).setTitle('Sell Fund');
  SpreadsheetApp.getUi().showModalDialog(html, 'Sell Fund');
}

function saveSellFund(holdingIdx, sellUnits, sellNav, returnAccount, notes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MF_SHEET_NAME);
  if (!sheet) throw new Error('Mutual Fund sheet not found.');

  const lastRow  = getLastFundRow_(sheet);
  const fundRows = sheet.getRange(4, 1, lastRow-3, 12).getValues()
    .map((r,i) => ({ rowNum: i+4, r }))
    .filter(x => x.r[MF_COLS.CODE-1] !== '');
  if (holdingIdx >= fundRows.length) throw new Error('Fund index out of range.');

  const { rowNum, r } = fundRows[holdingIdx];
  const fundCode   = r[MF_COLS.CODE-1];
  const fundName   = r[MF_COLS.NAME-1];
  const heldUnits  = r[MF_COLS.UNITS-1];
  const currency   = r[MF_COLS.CCY-1];
  const buyNav     = r[MF_COLS.BUY_NAV-1];
  const today      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  if (sellUnits > heldUnits + 0.00001) throw new Error('Cannot sell more units than held.');
  const remaining = heldUnits - sellUnits;

  if (remaining < 0.0001) {
    sheet.deleteRow(rowNum);
  } else {
    sheet.getRange(rowNum, MF_COLS.UNITS  ).setValue(remaining);
    sheet.getRange(rowNum, MF_COLS.NOTES  ).setValue((r[MF_COLS.NOTES-1] ? r[MF_COLS.NOTES-1]+' | ' : '') + 'Sold '+sellUnits.toFixed(4)+' @ '+currency+' '+sellNav.toFixed(4)+' on '+today.split(' ')[0]);
    sheet.getRange(rowNum, MF_COLS.UPDATED).setValue(today);
  }

  if (returnAccount) {
    const accSheet    = ss.getSheetByName(returnAccount);
    if (!accSheet) throw new Error('Account "' + returnAccount + '" not found.');
    const accCurrency = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCurrency !== normCurrency_(currency)) throw new Error('Currency mismatch.');
    const proceeds  = sellUnits * sellNav;
    const mktFmt    = { MYR:'"RM "#,##0.00', USD:'"$"#,##0.00', SGD:'"S$"#,##0.00', HKD:'"HK$"#,##0.00', CNY:'"¥"#,##0.00', RMB:'"¥"#,##0.00' };
    const fmt       = mktFmt[accCurrency] || mktFmt['MYR'];
    const lastAcc   = accSheet.getLastRow() + 1;
    const balFormula = lastAcc === 2 ? '=D' + lastAcc : '=F' + (lastAcc-1) + '+D' + lastAcc;
    accSheet.getRange(lastAcc,1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc,2).setValue('Unit Trust');
    accSheet.getRange(lastAcc,3).setValue('Sell '+fundCode+' x'+sellUnits.toFixed(4)+' @ '+currency+' '+sellNav.toFixed(4));
    accSheet.getRange(lastAcc,4).setValue(proceeds).setNumberFormat(fmt);
    accSheet.getRange(lastAcc,5).setValue('IN').setFontColor('#0f9d58').setFontWeight('bold');
    accSheet.getRange(lastAcc,6).setFormula(balFormula).setNumberFormat(fmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc, 1, 1, 6).setBackground('#f8f9fa');
  }

  refreshMutualFundSummary_(sheet);
}

function refreshMutualFundSummary_(sheet) {
  const lastFundRow = getLastFundRow_(sheet);
  const maxRows     = sheet.getMaxRows();
  const clearStart  = lastFundRow + 1;
  if (clearStart <= maxRows) sheet.getRange(clearStart, 1, maxRows - clearStart + 1, 12).clearContent().clearFormat();
  if (lastFundRow < 4) return;

  const data = sheet.getRange(4, 1, lastFundRow-3, 12).getValues()
    .filter(r => r[MF_COLS.CODE-1] !== '');
  const totalCost  = data.reduce((s,r) => s + (r[MF_COLS.UNITS-1] * r[MF_COLS.BUY_NAV-1]), 0);
  const totalValue = data.reduce((s,r) => s + (typeof r[MF_COLS.MKT_VAL-1] === 'number' ? r[MF_COLS.MKT_VAL-1] : 0), 0);
  const totalGain  = totalValue - totalCost;
  const gainPct    = totalCost > 0 ? (totalGain / totalCost * 100) : 0;
  const glColor    = totalGain >= 0 ? '#0f9d58' : '#d93025';

  const summaryRow = lastFundRow + 2;
  sheet.setRowHeight(summaryRow, 30);
  sheet.getRange(summaryRow, 1, 1, 12).merge()
    .setValue('📊  PORTFOLIO SUMMARY')
    .setBackground('#0d47a1').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  [
    ['Total Holdings',        data.length + ' funds', null,            '#202124'],
    ['Total Cost',            totalCost,               '"RM "#,##0.00', '#202124'],
    ['Total Market Value',    totalValue,              '"RM "#,##0.00', '#202124'],
    ['Total Gain / Loss',     totalGain,               '"RM "#,##0.00', glColor  ],
    ['Total Gain / Loss (%)', gainPct,                 '0.00"%"',       glColor  ]
  ].forEach(([label, val, fmt, color], i) => {
    const r = summaryRow + 1 + i;
    sheet.setRowHeight(r, 28);
    sheet.getRange(r, 1, 1, 3).merge()
      .setValue(label).setFontWeight('bold').setFontSize(10)
      .setBackground(i % 2 === 0 ? '#e8eaf6' : '#f3f4f9').setVerticalAlignment('middle');
    const vc = sheet.getRange(r, 4);
    vc.setValue(val).setFontColor(color).setFontWeight('bold').setVerticalAlignment('middle');
    if (fmt) vc.setNumberFormat(fmt);
  });
}

function refreshMutualFundNavs_(sheet) {
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = getLastFundRow_(sheet);
  if (lastRow < 4) return;
  for (let row = 4; row <= lastRow; row++) {
    const code = sheet.getRange(row, MF_COLS.CODE).getValue();
    if (!code) continue;
    const cell    = sheet.getRange(row, MF_COLS.CUR_NAV);
    const formula = cell.getFormula();
    if (formula) { cell.clearContent(); SpreadsheetApp.flush(); cell.setFormula(formula); }
    sheet.getRange(row, MF_COLS.UPDATED).setValue(today);
  }
  SpreadsheetApp.flush();
  for (let row = 4; row <= lastRow; row++) {
    const gain = sheet.getRange(row, MF_COLS.GAIN).getValue();
    if (typeof gain !== 'number') continue;
    const c = gain < 0 ? '#d93025' : '#0f9d58';
    sheet.getRange(row, MF_COLS.GAIN    ).setFontColor(c);
    sheet.getRange(row, MF_COLS.GAIN_PCT).setFontColor(c);
  }
  refreshMutualFundSummary_(sheet);
}

// ============================================================
//  STOCK PORTFOLIO
// ============================================================

const MARKET_CONFIG = {
  MY: {
    label:'Bursa Malaysia', flag:'🇲🇾', currency:'MYR',
    numFmt:'"RM "#,##0.00',   divFmt:'"RM "#,##0.00',
    sumFmt:'"RM "#,##0.00',   subtitle:'Bursa Malaysia  ·  Prices via KLSEScreener',
    placeholder:'e.g. 5347',  hint:'Enter KLSE stock code (numbers only)',
    sectors:['Banking','Technology','Consumer','Industrial','Healthcare','Property','Plantation','Energy','Telco','REITs','Others']
  },
  US: {
    label:'US Markets', flag:'🇺🇸', currency:'USD',
    numFmt:'"$"#,##0.00',     divFmt:'"$"#,##0.0000',
    sumFmt:'"$"#,##0.00',     subtitle:'NYSE / NASDAQ  ·  Prices via Google Finance',
    placeholder:'e.g. AAPL',  hint:'Enter ticker symbol (e.g. AAPL, MSFT, GOOGL)',
    sectors:['Technology','Healthcare','Finance','Consumer Discretionary','Communication','Industrials','Energy','Utilities','Materials','Real Estate','Others']
  },
  SG: {
    label:'SGX Singapore', flag:'🇸🇬', currency:'SGD',
    numFmt:'"S$"#,##0.00',    divFmt:'"S$"#,##0.0000',
    sumFmt:'"S$"#,##0.00',    subtitle:'SGX Singapore  ·  Prices via GrowBeanSprout',
    placeholder:'e.g. D05',   hint:'Enter SGX stock code (e.g. D05 for DBS)',
    sectors:['Banking','REITs','Technology','Industrials','Healthcare','Consumer','Telco','Others']
  },
  HK: {
    label:'HKEX Hong Kong', flag:'🇭🇰', currency:'HKD',
    numFmt:'"HK$"#,##0.00',   divFmt:'"HK$"#,##0.0000',
    sumFmt:'"HK$"#,##0.00',   subtitle:'HKEX  ·  Prices via Google Finance (HKG:)',
    placeholder:'e.g. 0700',  hint:'Enter HKEX code (e.g. 0700 for Tencent)',
    sectors:['Technology','Finance','Property','Consumer','Healthcare','Industrials','Energy','Utilities','Others']
  },
  CN: {
    label:'China A-Shares', flag:'🇨🇳', currency:'CNY',
    numFmt:'"¥"#,##0.00',     divFmt:'"¥"#,##0.0000',
    sumFmt:'"¥"#,##0.00',     subtitle:'Shanghai / Shenzhen  ·  Prices via Sina Finance',
    placeholder:'e.g. 600036', hint:'Enter A-share code (e.g. 600036 for CMB)',
    sectors:['Banking','Technology','Consumer','Healthcare','Industrials','Energy','Materials','Property','Others']
  }
};

function getStockSheetName_(market) { return '📈 Stock Portfolio (' + market + ')'; }

function getActivePortfolioSheet_() {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name   = sheet.getName();
  const valid  = ['MY','US','SG','HK','CN'].some(m => name === getStockSheetName_(m));
  if (!valid) {
    SpreadsheetApp.getUi().alert('Please activate a stock portfolio sheet first.\n\nNavigate to one of the "📈 Stock Portfolio" tabs.');
    return null;
  }
  return sheet;
}

function getSheetMarket_(sheet) {
  const v = sheet.getRange('P1').getValue();
  return v || 'MY';
}

function getLastStockRow_(sheet) {
  // Read cols A, B, C — a valid stock row has a code in A AND a company name in B
  // Summary rows have text in A but empty B, so we stop when B is empty for non-blank A
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 3) return 3;
  const data = sheet.getRange(4, 1, maxRows-3, 3).getValues();
  let last   = 3;
  for (let i = 0; i < data.length; i++) {
    const colA = data[i][0];
    const colB = data[i][1];
    const colC = data[i][2];
    // Valid stock row: has a code (A), a company name (B), and numeric shares (C)
    if (colA !== '' && colB !== '' && typeof colC === 'number' && colC > 0) {
      last = i + 4;
    }
  }
  return last;
}

function getCurrencySymbol_(market) {
  const map = { MY:'MYR', US:'USD', SG:'SGD', HK:'HKD', CN:'CNY' };
  return map[market] || 'MYR';
}

function navigateToSheet(name) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (s) s.activate();
}

// ── CREATE STOCK PORTFOLIO SHEET ─────────────────────────────
function createStockPortfolioSheet() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt('Create Stock Portfolio', 'Enter market code:\nMY = Bursa Malaysia\nUS = NYSE/NASDAQ\nSG = SGX Singapore\nHK = HKEX Hong Kong\nCN = China A-Shares', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const market = res.getResponseText().trim().toUpperCase();
  if (!MARKET_CONFIG[market]) { ui.alert('Invalid market code. Use MY, US, SG, HK, or CN.'); return; }
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = getStockSheetName_(market);
  if (ss.getSheetByName(name)) { ui.alert('"' + name + '" already exists.'); return; }
  const sheet = ss.insertSheet(name, ss.getSheets().length);
  buildPortfolioSheet_(sheet, market);
  sheet.activate();
  ui.alert('Stock Portfolio sheet for ' + MARKET_CONFIG[market].flag + ' ' + MARKET_CONFIG[market].label + ' created!');
}

function buildPortfolioSheet_(sheet, market) {
  const cfg = MARKET_CONFIG[market];
  compactSheet_(sheet, 100, 17);

  // Col widths: A-O + hidden P
  const widths = [80,200,80,110,110,120,110,90,110,80,70,90,90,120,160];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1, w));
  sheet.hideColumns(16); // P = market code

  // Row 1: Banner
  sheet.setRowHeight(1, 52);
  sheet.getRange(1,1,1,15).merge()
    .setValue(cfg.flag + '  ' + cfg.label.toUpperCase() + '  STOCK PORTFOLIO')
    .setBackground('#0d47a1').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Row 2: subtitle
  sheet.setRowHeight(2, 26);
  sheet.getRange(2,1,1,15).merge()
    .setValue(cfg.subtitle)
    .setBackground('#e8f0fe').setFontColor('#5f6368')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Row 3: Headers
  sheet.setRowHeight(3, 32);
  const headers = ['Code','Company','Shares','Avg Buy Price','Current Price','Market Value',
    'Gain/Loss','Gain %','Sector','Dividend','DIV %','Ex-Date','Pay Date','Last Updated','Notes'];
  headers.forEach((h,i) => {
    sheet.getRange(3,i+1)
      .setValue(h).setBackground('#1a73e8').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  // Store market code in P1
  sheet.getRange('P1').setValue(market);
  sheet.setFrozenRows(3);
  sheet.setHiddenGridlines(true);
}


// ── CUSTOM PRICE FUNCTIONS ────────────────────────────────────

/**
 * Gets KLSE stock price from klsescreener.com
 * @param {string} stockCode - KLSE stock code (e.g. "5347")
 * @customfunction
 */
function KLSE(stockCode) {
  try {
    const code    = stockCode.toString().trim();
    const url     = 'https://www.klsescreener.com/v2/stocks/' + code;
    const cookie  = Array.from({length:26}, () => 'abcdefghijklmnopqrstuvwxyz0123456789'[Math.floor(Math.random()*36)]).join('');
    const res     = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Cookie': 'session=' + cookie }
    });
    if (res.getResponseCode() !== 200) return 'N/A';
    const text    = res.getContentText();
    const match   = text.match(/data-value="([\d.]+)"/);
    if (match) {
      const price = parseFloat(match[1]);
      if (!isNaN(price)) return price;
    }
    // Fallback: Yahoo Finance v8 API (code must be numeric KLSE code like 1155)
    const yfUrl = 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.KL?interval=1d&range=1d';
    const yfRes = UrlFetchApp.fetch(yfUrl, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
    if (yfRes.getResponseCode() === 200) {
      const yfJson  = JSON.parse(yfRes.getContentText());
      const yfPrice = yfJson?.chart?.result?.[0]?.meta?.regularMarketPrice;
      if (yfPrice && !isNaN(yfPrice)) return yfPrice;
    }
    return 'N/A';
  } catch(e) { return 'Error: ' + e.message; }
}

/**
 * Gets SGX stock price from growbeansprout.com
 * @param {string} stockCode - SGX stock code (e.g. "D05")
 * @customfunction
 */
function SGX(stockCode) {
  // Uses Yahoo Finance v8 JSON API — growbeansprout.com is defunct
  try {
    const code   = stockCode.toString().trim().toUpperCase();
    const symbol = code.includes('.SI') ? code : code + '.SI';
    const url    = 'https://query1.finance.yahoo.com/v8/finance/chart/' + symbol +
                   '?interval=1d&range=1d';
    const res    = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    if (res.getResponseCode() !== 200) return 'N/A';
    const json  = JSON.parse(res.getContentText());
    const price = json?.chart?.result?.[0]?.meta?.regularMarketPrice;
    if (!price || isNaN(price)) return 'N/A';
    return price;
  } catch(e) { return 'Error: ' + e.message; }
}

/**
 * Gets China A-share price via Sina Finance real-time API
 * @param {string} stockCode - A-share code (e.g. "600036" or "SHA:600036")
 * @customfunction
 */
function CNStock(stockCode) {
  try {
    let code = stockCode.toString().trim().replace(/^(SHA:|SHE:)/i,'');
    const prefix = (code.startsWith('6') ? 'sh' : 'sz');
    const url  = 'https://hq.sinajs.cn/list=' + prefix + code;
    const res  = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'Referer': 'https://finance.sina.com.cn' }
    });
    if (res.getResponseCode() !== 200) return 'N/A';
    const text   = res.getContentText();
    const match  = text.match(/="([^"]+)"/);
    if (!match) return 'N/A';
    const fields = match[1].split(',');
    const price  = parseFloat(fields[3]);
    return (isNaN(price) || price === 0) ? 'No Data' : price;
  } catch(e) { return 'Error: ' + e.message; }
}

// ── DIVIDEND CUSTOM FUNCTIONS (StockAnalysis.com __data.json) ─
/**
 * Fetches dividend data from StockAnalysis.com for a given stock.
 * Returns a 3-element array [amount, exDate, payDate] or null on failure.
 * Internal helper used by StockDividend, StockExDate, StockPayDate.
 */
function fetchStockDividendData_(code, market) {
  try {
    const c = code.toString().trim();
    let url;
    if      (market === 'MY') url = 'https://stockanalysis.com/quote/klse/' + encodeURIComponent(c) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else if (market === 'US') url = 'https://stockanalysis.com/stocks/'     + encodeURIComponent(c.toLowerCase()) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else if (market === 'SG') url = 'https://stockanalysis.com/quote/sgx/'  + encodeURIComponent(c) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else if (market === 'HK') url = 'https://stockanalysis.com/quote/hkg/'  + encodeURIComponent(c) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else return null;

    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }
    });
    if (res.getResponseCode() !== 200) return null;

    const json    = JSON.parse(res.getContentText());
    const nodes   = json.nodes || [];
    let dataArr   = null;
    let root      = null;

    // SvelteKit deduplication: find the node whose data[] has {infoTable, history}
    for (let n = nodes.length - 1; n >= 0; n--) {
      const nd = nodes[n];
      if (!nd || nd.type !== 'data' || !Array.isArray(nd.data)) continue;
      for (let i = 0; i < nd.data.length; i++) {
        const item = nd.data[i];
        if (item && typeof item === 'object' && 'infoTable' in item && 'history' in item) {
          dataArr = nd.data; root = item; break;
        }
      }
      if (root) break;
    }
    if (!root) return null;

    function res_(ref) { return typeof ref === 'number' ? dataArr[ref] : ref; }

    // Most-recent dividend row
    const histIdxs = res_(root.history);
    if (!Array.isArray(histIdxs) || histIdxs.length === 0) return null;
    const row = res_(histIdxs[0]);
    if (!row) return null;

    // Parse amount (e.g. "0.330 MYR" → 0.33)
    const amtRaw = res_(row.amt) || '';
    const amtM   = amtRaw.toString().match(/([\d.]+)/);
    const amount = amtM ? parseFloat(amtM[1]) : null;

    // ISO dates → dd/MM/yyyy
    function isoToDisplay(iso) {
      if (!iso) return 'N/A';
      const p = iso.toString().split('-');
      return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : iso;
    }
    const exDate  = isoToDisplay(res_(row.dt));
    const payDate = isoToDisplay(res_(row.pay));

    return [amount || 'N/A', exDate, payDate];
  } catch(e) {
    Logger.log('fetchStockDividendData_ error: ' + e.message);
    return null;
  }
}

/**
 * Returns the latest dividend amount for a stock from StockAnalysis.com.
 * @param {string} code   Stock code (e.g. "MAYBANK", "AAPL", "D05", "0700")
 * @param {string} market Market: "MY", "US", "SG", "HK"
 * @return Dividend amount or "N/A"
 * @customfunction
 */
function StockDividend(code, market) {
  const d = fetchStockDividendData_(code, market);
  return d ? d[0] : 'N/A';
}

/**
 * Returns the latest ex-dividend date for a stock from StockAnalysis.com.
 * @param {string} code   Stock code
 * @param {string} market Market: "MY", "US", "SG", "HK"
 * @return Ex-date string (dd/MM/yyyy) or "N/A"
 * @customfunction
 */
function StockExDate(code, market) {
  const d = fetchStockDividendData_(code, market);
  return d ? d[1] : 'N/A';
}

/**
 * Returns the latest pay date for a stock from StockAnalysis.com.
 * @param {string} code   Stock code
 * @param {string} market Market: "MY", "US", "SG", "HK"
 * @return Pay-date string (dd/MM/yyyy) or "N/A"
 * @customfunction
 */
function StockPayDate(code, market) {
  const d = fetchStockDividendData_(code, market);
  return d ? d[2] : 'N/A';
}


/**
 * Gets mutual fund NAV from FSMOne Malaysia
 * @param {string} fundCode - FSMOne fund sedol code (e.g. "MYOSKCMF")
 * @customfunction
 */
function FSMFund(fundCode) {
  // Primary: look up FIMM NAV Cache sheet (populated daily at 7 AM)
  try {
    const code = fundCode.toString().trim().toUpperCase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cacheSheet = ss.getSheetByName(FIMM_NAV_SHEET);
    if (cacheSheet) {
      const lastRow = cacheSheet.getLastRow();
      if (lastRow >= 2) {
        const data = cacheSheet.getRange(2, 1, lastRow - 1, 3).getValues(); // code, name, nav
        for (var i = 0; i < data.length; i++) {
          if (data[i][0].toString().toUpperCase() === code) {
            var nav = parseFloat(data[i][2]);
            if (!isNaN(nav) && nav > 0) return nav;
          }
        }
      }
    }
  } catch(e2) { /* fall through to FSMOne */ }

  // Fallback: try FSMOne factsheet API (may be blocked)
  try {
    const res = UrlFetchApp.fetch(FSM_FACTSHEET_URL + encodeURIComponent(fundCode.toString().trim()), {
      method: 'post', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Referer': FSM_BASE }
    });
    if (res.getResponseCode() !== 200) return 'N/A';
    const json = JSON.parse(res.getContentText());
    if (json.status !== 'SUCCESS') return 'N/A';
    const nav = parseFloat(json.data.latestNavPrice.navPrice);
    return isNaN(nav) ? 'N/A' : nav;
  } catch(e) { return 'N/A'; }
}


// ── ADD STOCK DIALOG ─────────────────────────────────────────
function showAddStockDialog() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = getActivePortfolioSheet_();
  if (!sheet) return;

  const market      = getSheetMarket_(sheet);
  const cfg         = MARKET_CONFIG[market];
  const sectorsJson = JSON.stringify(cfg.sectors);
  const portCurrency = getCurrencySymbol_(market);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matchingAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }))
    .filter(a => a.currency === normCurrency_(portCurrency));

  const accountsHtml = matchingAccounts.map(a =>
    '<option value="' + a.name + '">' + a.name + ' (' + a.currency + ')</option>'
  ).join('');
  const noMatch      = matchingAccounts.length === 0;

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#1a73e8}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.hint{color:#9aa0a6;font-size:11px;margin-top:2px}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#1a73e8;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>' + cfg.flag + ' Add Stock – ' + cfg.label + '</h2>' +
    '<label>Stock Code</label>' +
    '<input type="text" id="code" placeholder="' + cfg.placeholder + '" />' +
    '<div class="hint">' + cfg.hint + '</div>' +
    '<label>Company Name</label>' +
    '<input type="text" id="company" placeholder="e.g. Maybank" />' +
    '<div class="row2">' +
    '<div><label>Shares</label><input type="number" id="shares" placeholder="100" step="1" min="1" /></div>' +
    '<div><label>Avg Buy Price (' + portCurrency + ')</label><input type="number" id="buyPrice" placeholder="0.00" step="0.001" min="0.001" /></div>' +
    '</div>' +
    '<label>Sector</label>' +
    '<select id="sector">' + cfg.sectors.map(s => '<option>' + s + '</option>').join('') + '</select>' +
    '<label>Linked Account (deduct purchase)</label>' +
    (noMatch
      ? '<div class="warn">No ' + portCurrency + ' accounts found. Cost won\'t be deducted.</div>'
      : '<select id="linkedAccount"><option value="">— None —</option>' + accountsHtml + '</select>') +
    '<label>Notes (optional)</label>' +
    '<input type="text" id="notes" placeholder="e.g. Long term hold" />' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="save()">Add Stock</button></div>' +
    '<script>function save(){' +
    'var code=document.getElementById("code").value.trim();' +
    'var company=document.getElementById("company").value.trim();' +
    'var shares=parseFloat(document.getElementById("shares").value);' +
    'var buyPrice=parseFloat(document.getElementById("buyPrice").value);' +
    'if(!code){alert("Enter stock code.");return;}' +
    'if(!shares||shares<=0){alert("Enter valid shares.");return;}' +
    'if(!buyPrice||buyPrice<=0){alert("Enter valid buy price.");return;}' +
    'var sector=document.getElementById("sector").value;' +
    'var notes=document.getElementById("notes").value.trim();' +
    'var acctEl=document.getElementById("linkedAccount");' +
    'var account=acctEl?acctEl.value:"";' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){alert("Error: "+e.message);})' +
    '.saveStock(code,company,shares,buyPrice,sector,notes,account);}' +
    '</script></body></html>'
  ).setWidth(460).setHeight(520).setTitle('Add Stock');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Stock');
}

function saveStock(code, company, shares, buyPrice, sector, notes, linkedAccount) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getActivePortfolioSheet_();
  if (!sheet) return;

  const market  = getSheetMarket_(sheet);
  const cfg     = MARKET_CONFIG[market];
  const portCcy = getCurrencySymbol_(market);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const row     = getLastStockRow_(sheet) + 1;

  // Server-side currency guard
  if (linkedAccount) {
    const accSheet = ss.getSheetByName(linkedAccount);
    if (!accSheet) throw new Error('Account "' + linkedAccount + '" not found.');
    const accCcy   = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCcy !== normCurrency_(portCcy)) throw new Error('Currency mismatch: account is ' + accCcy + ' but portfolio is ' + portCcy + '.');
  }

  // Price formula
  let priceFormula;
  if (market === 'MY') priceFormula = '=IFERROR(KLSE("' + code + '"),"N/A")';
  else if (market === 'US') priceFormula = '=IFERROR(GOOGLEFINANCE("' + code + '","price"),"N/A")';
  else if (market === 'SG') priceFormula = '=IFERROR(SGX("' + code + '"),"N/A")';
  else if (market === 'HK') priceFormula = '=IFERROR(GOOGLEFINANCE("HKG:' + code + '","price"),"N/A")';
  else                       priceFormula = '=IFERROR(CNStock("' + code + '"),"N/A")';

  sheet.setRowHeight(row, 34);
  const bg = row % 2 === 0 ? '#f8f9fa' : '#ffffff';
  sheet.getRange(row, 1, 1, 15).setBackground(bg);

  sheet.getRange(row, 1).setValue(code.toUpperCase()).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.getRange(row, 2).setValue(company).setVerticalAlignment('middle');
  sheet.getRange(row, 3).setValue(shares).setNumberFormat('#,##0').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, 4).setValue(buyPrice).setNumberFormat(cfg.numFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, 5).setFormula(priceFormula).setNumberFormat(cfg.numFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, 6).setFormula('=IF(ISNUMBER(E' + row + '),C' + row + '*E' + row + ',0)').setNumberFormat(cfg.numFmt).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, 7).setFormula('=IF(ISNUMBER(E' + row + '),F' + row + '-C' + row + '*D' + row + ',0)').setNumberFormat(cfg.numFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, 8).setFormula('=IF(D' + row + '*C' + row + '<>0,(E' + row + '-D' + row + ')/D' + row + '*100,0)').setNumberFormat('0.00"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, 9).setValue(sector).setVerticalAlignment('middle');
  // Dividend formulas using native IMPORTHTML — no GAS fetching needed.
  // References B{row} (ticker, e.g. MAYBANK) — col A has numeric code, col B has ticker.
  // Table 1 on the dividend page: col1=Ex-Date, col2=Amount, col4=Pay Date
  const aRef = 'B' + row;
  if (market !== 'CN') {
    // stockanalysis.com table structure (single table per page):
    // col1=Ex-Date, col2=Amount ("$0.26" for US, "0.06 MYR" for MY etc.), col4=Pay Date
    // Dates are in "Feb 9, 2026" format; CLEAN+TRIM normalises whitespace before DATEVALUE.
    var divAmtFml, exDateFml, payDateFml, divYldFml;
    const xpYld = '/html/body/div/div[1]/div[2]/main/div[2]/div/div[2]/div[1]/div';
    if (market === 'US') {
      // US stocks use /stocks/{ticker}/ and ETFs use /etf/{ticker}/ — try stocks first, ETF as fallback
      const stockUrl = 'CONCATENATE("https://stockanalysis.com/stocks/",LOWER(' + aRef + '),"/dividend/")';
      const etfUrl   = 'CONCATENATE("https://stockanalysis.com/etf/",LOWER(' + aRef + '),"/dividend/")';
      divAmtFml = '=IFERROR('
        + 'VALUE(SUBSTITUTE(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,2),"$","")),'
        + 'IFERROR('
        + 'VALUE(SUBSTITUTE(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,2),"$","")),'
        + '"No Data"))';
      exDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,1)))),'
        + 'IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,1)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,1))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,1))),'
        + '"No Data"))))';
      payDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,4)))),'
        + 'IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,4)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,4))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,4))),'
        + '"No Data"))))';
      divYldFml = '=IFERROR('
        + 'IMPORTXML(' + stockUrl + ',"' + xpYld + '")*100,'
        + 'IFERROR('
        + 'IMPORTXML(' + etfUrl + ',"' + xpYld + '")*100,'
        + '"N/A"))';
    } else {
      var divUrl;
      if      (market === 'MY') divUrl = 'CONCATENATE("https://stockanalysis.com/quote/klse/",' + aRef + ',"/dividend/")';
      else if (market === 'SG') divUrl = 'CONCATENATE("https://stockanalysis.com/quote/sgx/",' + aRef + ',"/dividend/")';
      else if (market === 'HK') divUrl = 'CONCATENATE("https://stockanalysis.com/quote/hkg/",' + aRef + ',"/dividend/")';
      var ccy = market === 'MY' ? ' MYR' : market === 'SG' ? ' SGD' : ' HKD';
      divAmtFml = '=IFERROR(VALUE(SUBSTITUTE(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,2),"' + ccy + '","")),"No Data")';
      exDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,1)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,1))),'
        + '"No Data"))';
      payDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,4)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,4))),'
        + '"No Data"))';
      divYldFml = '=IFERROR(IMPORTXML(' + divUrl + ',"' + xpYld + '")*100,"N/A")';
    }
    sheet.getRange(row,10).setFormula(divAmtFml)
      .setNumberFormat('0.0000').setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(row,11).setFormula(divYldFml)
      .setNumberFormat('0.0000"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(row,12).setFormula(exDateFml)
      .setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(row,13).setFormula(payDateFml)
      .setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('center').setVerticalAlignment('middle');
  } else {
    sheet.getRange(row,10).setValue('N/A').setVerticalAlignment('middle');
    sheet.getRange(row,11).setValue('N/A').setVerticalAlignment('middle');
    sheet.getRange(row,12).setValue('N/A').setVerticalAlignment('middle');
    sheet.getRange(row,13).setValue('N/A').setVerticalAlignment('middle');
  }
  sheet.getRange(row,14).setValue(today).setFontSize(9).setFontColor('#9aa0a6').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row,15).setValue(notes||'').setFontSize(10).setFontColor('#9aa0a6').setVerticalAlignment('middle');

  SpreadsheetApp.flush();
  const gainVal = sheet.getRange(row, 7).getValue();
  const glColor = (typeof gainVal === 'number' && gainVal < 0) ? '#d93025' : '#0f9d58';
  sheet.getRange(row, 7).setFontColor(glColor);
  sheet.getRange(row, 8).setFontColor(glColor);

  if (linkedAccount) {
    const cost     = shares * buyPrice;
    const accSheet = ss.getSheetByName(linkedAccount);
    const accFmt   = { MYR:'"RM "#,##0.00', USD:'"$"#,##0.00', SGD:'"S$"#,##0.00', HKD:'"HK$"#,##0.00', CNY:'"¥"#,##0.00', RMB:'"¥"#,##0.00' };
    const fmt      = accFmt[portCcy] || accFmt['MYR'];
    const lastAcc  = accSheet.getLastRow() + 1;
    const balFormula = lastAcc === 2 ? '=D' + lastAcc : '=F' + (lastAcc-1) + '-D' + lastAcc;
    accSheet.getRange(lastAcc,1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc,2).setValue('Stock');
    accSheet.getRange(lastAcc,3).setValue('Buy ' + code + ' x' + shares + ' @ ' + portCcy + ' ' + buyPrice.toFixed(3));
    accSheet.getRange(lastAcc,4).setValue(cost).setNumberFormat(fmt);
    accSheet.getRange(lastAcc,5).setValue('OUT').setFontColor('#d93025').setFontWeight('bold');
    accSheet.getRange(lastAcc,6).setFormula(balFormula).setNumberFormat(fmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc, 1, 1, 6).setBackground('#f8f9fa');
  }

  refreshStockSummary_(sheet, cfg);
}


// ── SELL STOCK DIALOG ─────────────────────────────────────────
function showSellStockDialog() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = getActivePortfolioSheet_();
  if (!sheet) return;

  const market    = getSheetMarket_(sheet);
  const cfg       = MARKET_CONFIG[market];
  const portCcy   = getCurrencySymbol_(market);
  const lastRow   = getLastStockRow_(sheet);
  if (lastRow < 4) { ui.alert('No stocks found. Add a stock first.'); return; }

  const stocks = sheet.getRange(4, 1, lastRow-3, 8).getValues()
    .filter(r => r[0] !== '' && typeof r[2] === 'number' && r[2] > 0)
    .map(r => ({ code: r[0], company: r[1], shares: r[2], buyPrice: r[3], curPrice: r[4] }));

  if (!stocks.length) { ui.alert('No active stock holdings found.'); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const matchingAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }))
    .filter(a => a.currency === normCurrency_(portCcy));

  const stocksJson   = JSON.stringify(stocks);
  const accountsJson = JSON.stringify(matchingAccounts);
  const stockOptions = stocks.map((s,i) =>
    '<option value="' + i + '">' + s.code + ' – ' + s.company + ' (' + s.shares + ' shares)</option>'
  ).join('');

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#c62828}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.info{background:#fff8e1;border:1px solid #fbc02d;border-radius:6px;padding:8px 10px;font-size:12px;margin-top:4px;display:none}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.preview{background:#e8f5e9;border-radius:6px;padding:8px 10px;margin-top:8px;font-size:12px;color:#2e7d32;display:none}' +
    '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px;display:none}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#c62828;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>' + cfg.flag + ' Sell Stock</h2>' +
    '<label>Select Stock</label>' +
    '<select id="stockSel" onchange="onStockChange()"><option value="">— Select —</option>' + stockOptions + '</select>' +
    '<div class="info" id="stockInfo"></div>' +
    '<div id="sellForm" style="display:none">' +
    '<div class="row2">' +
    '<div><label>Shares to Sell</label><input type="number" id="sellShares" step="1" min="1" oninput="updatePreview()" /></div>' +
    '<div><label>Sell Price (' + portCcy + ')</label><input type="number" id="sellPrice" step="0.001" oninput="updatePreview()" /></div>' +
    '</div>' +
    '<label>Return Proceeds to Account</label>' +
    '<select id="returnAcct"><option value="">— None —</option>' +
    matchingAccounts.map(a => '<option value="' + a.name + '">' + a.name + ' (' + a.currency + ')</option>').join('') +
    '</select>' +
    '<div class="warn" id="noAcctWarn">' + (matchingAccounts.length === 0 ? 'No ' + portCcy + ' accounts found.' : '') + '</div>' +
    '<div class="preview" id="preview"></div>' +
    '<label>Notes (optional)</label><input type="text" id="notes" /></div>' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" id="saveBtn" onclick="save()" style="display:none">Confirm Sell</button></div>' +
    '<script>' +
    'var STOCKS=' + stocksJson + ';' +
    'var ACCOUNTS=' + accountsJson + ';' +
    'var selIdx=-1;' +
    'function onStockChange(){' +
    'selIdx=parseInt(document.getElementById("stockSel").value);' +
    'if(isNaN(selIdx)){document.getElementById("sellForm").style.display="none";document.getElementById("stockInfo").style.display="none";return;}' +
    'var s=STOCKS[selIdx];' +
    'var info=document.getElementById("stockInfo");' +
    'info.innerHTML="<b>"+s.code+"</b> – "+s.company+" | Held: <b>"+s.shares+" shares</b> | Avg Buy: <b>"+s.buyPrice.toFixed(3)+"</b> | Current: <b>"+s.curPrice+"</b>";' +
    'info.style.display="block";' +
    'document.getElementById("sellShares").max=s.shares;' +
    'document.getElementById("sellShares").placeholder="Max: "+s.shares;' +
    'document.getElementById("sellPrice").value=s.curPrice||"";' +
    'document.getElementById("sellForm").style.display="block";' +
    'document.getElementById("saveBtn").style.display="inline-block";' +
    'if(ACCOUNTS.length===0)document.getElementById("noAcctWarn").style.display="block";}' +
    'function updatePreview(){' +
    'if(selIdx<0)return;var s=STOCKS[selIdx];' +
    'var sh=parseFloat(document.getElementById("sellShares").value)||0;' +
    'var pr=parseFloat(document.getElementById("sellPrice").value)||0;' +
    'var prev=document.getElementById("preview");' +
    'if(sh>0&&pr>0){' +
    'var proceeds=(sh*pr).toFixed(2);var gain=((pr-s.buyPrice)*sh).toFixed(2);' +
    'var gs=parseFloat(gain)>=0?"+":"";' +
    'prev.innerHTML="Proceeds: <b>"+proceeds+"</b> | Gain/Loss: <b>"+gs+gain+"</b>";' +
    'prev.style.display="block";}else prev.style.display="none";}' +
    'function save(){' +
    'if(selIdx<0){alert("Select a stock.");return;}' +
    'var s=STOCKS[selIdx];' +
    'var sh=parseFloat(document.getElementById("sellShares").value);' +
    'var pr=parseFloat(document.getElementById("sellPrice").value);' +
    'if(!sh||sh<=0){alert("Enter shares.");return;}' +
    'if(sh>s.shares+0.001){alert("Cannot sell more than held ("+s.shares+").");return;}' +
    'if(!pr||pr<=0){alert("Enter sell price.");return;}' +
    'var acct=document.getElementById("returnAcct").value;' +
    'var notes=document.getElementById("notes").value.trim();' +
    'document.getElementById("saveBtn").disabled=true;' +
    'document.getElementById("saveBtn").textContent="Processing...";' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){document.getElementById("saveBtn").disabled=false;document.getElementById("saveBtn").textContent="Confirm Sell";alert("Error: "+e.message);})' +
    '.saveSellStock(selIdx,sh,pr,acct,notes);}' +
    '</script></body></html>'
  ).setWidth(460).setHeight(520).setTitle('Sell Stock');
  SpreadsheetApp.getUi().showModalDialog(html, 'Sell Stock');
}

function saveSellStock(stockIdx, sellShares, sellPrice, returnAccount, notes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getActivePortfolioSheet_();
  if (!sheet) return;

  const market  = getSheetMarket_(sheet);
  const cfg     = MARKET_CONFIG[market];
  const portCcy = getCurrencySymbol_(market);
  const lastRow = getLastStockRow_(sheet);
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  const stockRows = sheet.getRange(4, 1, lastRow-3, 8).getValues()
    .map((r,i) => ({ rowNum: i+4, r }))
    .filter(x => x.r[0] !== '' && typeof x.r[2] === 'number' && x.r[2] > 0);

  if (stockIdx >= stockRows.length) throw new Error('Stock index out of range.');
  const { rowNum, r } = stockRows[stockIdx];
  const code      = r[0];
  const heldShares= r[2];
  const buyPrice  = r[3];
  const remaining = heldShares - sellShares;

  if (returnAccount) {
    const accSheet = ss.getSheetByName(returnAccount);
    if (!accSheet) throw new Error('Account "' + returnAccount + '" not found.');
    const accCcy   = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCcy !== normCurrency_(portCcy)) throw new Error('Currency mismatch.');
  }

  if (remaining < 0.5) {
    sheet.deleteRow(rowNum);
  } else {
    sheet.getRange(rowNum, 3).setValue(Math.round(remaining));
    sheet.getRange(rowNum,13).setValue(today);
    const existNotes = sheet.getRange(rowNum,14).getValue();
    sheet.getRange(rowNum,14).setValue((existNotes ? existNotes + ' | ' : '') + 'Sold ' + sellShares + ' @ ' + portCcy + ' ' + sellPrice.toFixed(3) + ' on ' + today.split(' ')[0]);
    SpreadsheetApp.flush();
    const gainVal = sheet.getRange(rowNum, 7).getValue();
    const glColor = (typeof gainVal === 'number' && gainVal < 0) ? '#d93025' : '#0f9d58';
    sheet.getRange(rowNum, 7).setFontColor(glColor);
    sheet.getRange(rowNum, 8).setFontColor(glColor);
  }

  if (returnAccount) {
    const proceeds = sellShares * sellPrice;
    const accSheet = ss.getSheetByName(returnAccount);
    const accFmt   = { MYR:'"RM "#,##0.00', USD:'"$"#,##0.00', SGD:'"S$"#,##0.00', HKD:'"HK$"#,##0.00', CNY:'"¥"#,##0.00', RMB:'"¥"#,##0.00' };
    const fmt      = accFmt[portCcy] || accFmt['MYR'];
    const lastAcc  = accSheet.getLastRow() + 1;
    const balFormula = lastAcc === 2 ? '=D' + lastAcc : '=F' + (lastAcc-1) + '+D' + lastAcc;
    accSheet.getRange(lastAcc,1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc,2).setValue('Stock');
    accSheet.getRange(lastAcc,3).setValue('Sell ' + code + ' x' + sellShares + ' @ ' + portCcy + ' ' + sellPrice.toFixed(3));
    accSheet.getRange(lastAcc,4).setValue(proceeds).setNumberFormat(fmt);
    accSheet.getRange(lastAcc,5).setValue('IN').setFontColor('#0f9d58').setFontWeight('bold');
    accSheet.getRange(lastAcc,6).setFormula(balFormula).setNumberFormat(fmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc, 1, 1, 6).setBackground('#f8f9fa');
  }

  refreshStockSummary_(sheet, cfg);
}

// ── DASHBOARD ─────────────────────────────────────────────────
function renderDashboardSheet() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const accounts = getAccountBalances();
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  // ── FX Rate fetcher ───────────────────────────────────────────
  const BASE = 'MYR';
  const fmtMap = {
    MYR:'"RM "#,##0.00', USD:'"$"#,##0.00', SGD:'"S$"#,##0.00',
    HKD:'"HK$"#,##0.00', CNY:'"¥"#,##0.00',  RMB:'"¥"#,##0.00'
  };
  const myrFmt = fmtMap['MYR'];

  // Collect all foreign currencies needed
  const allCurrencies = new Set();
  accounts.forEach(a => { if (normCurrency_(a.currency) !== BASE) allCurrencies.add(normCurrency_(a.currency)); });

  // Collect stock portfolios
  const portfolios = [];
  ['MY','US','SG','HK','CN'].forEach(mkt => {
    const sheet = ss.getSheetByName(getStockSheetName_(mkt));
    if (!sheet) return;
    const lastRow = getLastStockRow_(sheet);
    if (lastRow < 4) return;
    const cfg    = MARKET_CONFIG[mkt];
    const data   = sheet.getRange(4, 1, lastRow-3, 7).getValues()
      .filter(r => r[0] !== '' && typeof r[2] === 'number' && r[2] > 0);
    if (!data.length) return;
    const value  = data.reduce((s,r) => s + (typeof r[5] === 'number' ? r[5] : 0), 0);
    const cost   = data.reduce((s,r) => s + (r[2] * r[3]), 0);
    const portCcy = getCurrencySymbol_(mkt);
    if (normCurrency_(portCcy) !== BASE) allCurrencies.add(normCurrency_(portCcy));
    portfolios.push({ market: mkt, flag: cfg.flag, label: cfg.label, currency: portCcy, value, gainLoss: value - cost, count: data.length, stocksFmt: cfg.numFmt });
  });

  // Collect mutual fund data
  const mfSheet = ss.getSheetByName(MF_SHEET_NAME);
  const mfFunds = [];
  if (mfSheet) {
    const mfLast = getLastFundRow_(mfSheet);
    if (mfLast >= 4) {
      const mfData = mfSheet.getRange(4, 1, mfLast-3, 12).getValues()
        .filter(r => r[MF_COLS.CODE-1] !== '');
      const byCcy = {};
      mfData.forEach(r => {
        const ccy = r[MF_COLS.CCY-1] || 'MYR';
        if (!byCcy[ccy]) byCcy[ccy] = { currency: ccy, value: 0, cost: 0, count: 0 };
        byCcy[ccy].value += typeof r[MF_COLS.MKT_VAL-1] === 'number' ? r[MF_COLS.MKT_VAL-1] : 0;
        byCcy[ccy].cost  += r[MF_COLS.UNITS-1] * r[MF_COLS.BUY_NAV-1];
        byCcy[ccy].count++;
      });
      Object.values(byCcy).forEach(m => {
        if (normCurrency_(m.currency) !== BASE) allCurrencies.add(normCurrency_(m.currency));
        mfFunds.push(m);
      });
    }
  }

  const cryptoSheetDash = ss.getSheetByName(CRYPTO_SHEET_NAME);
  let cryptoHoldings = [];
  if (cryptoSheetDash) {
    const lastCRow = getLastCryptoRow_(cryptoSheetDash);
    if (lastCRow >= 4) {
      cryptoHoldings = cryptoSheetDash.getRange(4, 1, lastCRow - 3, 6).getValues()
        .filter(r => r[0] !== '' && typeof r[2] === 'number' && r[2] > 0);
      // All crypto is priced in USD — ensure USD rate is always fetched
      allCurrencies.add('USD');
    }
  }

  // Collect gold holdings (all in MYR — no FX needed)
  const goldSheetDash = ss.getSheetByName(GOLD_SHEET_NAME);
  let goldHoldings = [];
  if (goldSheetDash) {
    const lastGRow = getLastGoldRow_(goldSheetDash);
    if (lastGRow >= 4) {
      goldHoldings = goldSheetDash.getRange(4, 1, lastGRow - 3, 9).getValues()
        .filter(r => r[GC.TYPE-1] !== '' && r[GC.DESC-1] !== '' && typeof r[GC.WEIGHT-1] === 'number' && r[GC.WEIGHT-1] > 0);
    }
  }

  const foreignCurrencies = [...allCurrencies];

  // FX rates: dashboard now uses live GOOGLEFINANCE formulas in cells.
  // toMYR() is only used here for section-total approximations at render time.
  // We still snapshot rates for totals — but via a compact 1-col × 10-row temp sheet.
  const fxRates = {};
  if (foreignCurrencies.length > 0) {
    try {
      let temp = ss.getSheetByName('_FX_TEMP_');
      if (temp) ss.deleteSheet(temp);
      temp = ss.insertSheet('_FX_TEMP_');
      compactSheet_(temp, 10, 1); // 1 column, 10 rows only
      foreignCurrencies.forEach((cur, i) => {
        temp.getRange(i + 1, 1).setFormula('=GOOGLEFINANCE("CURRENCY:' + cur + BASE + '")');
      });
      SpreadsheetApp.flush();
      Utilities.sleep(3000);
      foreignCurrencies.forEach((cur, i) => {
        const val = temp.getRange(i + 1, 1).getValue();
        fxRates[cur] = (typeof val === 'number' && val > 0) ? val : null;
      });
      ss.deleteSheet(temp);
    } catch(e) {
      try { const t = ss.getSheetByName('_FX_TEMP_'); if (t) ss.deleteSheet(t); } catch(_) {}
      foreignCurrencies.forEach(cur => { fxRates[cur] = null; });
    }
  }

  function toMYR(amount, currency) {
    const cur = normCurrency_(currency);
    if (cur === BASE) return { ok: true, myr: amount };
    const rate = fxRates[cur];
    if (rate === null || rate === undefined) return { ok: false, myr: 0 };
    return { ok: true, myr: amount * rate };
  }

  // ── Get or create dashboard sheet ──────────────────────────
  let dash = ss.getSheetByName(DASH_NAME);
  if (!dash) dash = ss.insertSheet(DASH_NAME, 0);
  else { dash.clearContents(); dash.clearFormats(); }

  // ── Layout ─────────────────────────────────────────────────
  const n = accounts.length;
  const p = portfolios.length;
  const m = mfFunds.length;
  const f = foreignCurrencies.length;

  const accStart  = 6;
  const accTotal  = accStart + n;
  const stkHeader = accTotal + 2;
  const stkColHdr = stkHeader + 1;
  const stkStart  = stkColHdr + 1;
  const stkTotal  = stkStart + p;
  const mfHeader  = stkTotal + 2;
  const mfColHdr  = mfHeader + 1;
  const mfStart   = mfColHdr + 1;
  const mfTotal   = mfStart + m;
  // Crypto section
  const cr = cryptoHoldings.length;
  const cryptoHeader = mfTotal + 2;
  const cryptoColHdr = cryptoHeader + 1;
  const cryptoStart  = cryptoColHdr + 1;
  const cryptoTotal  = cryptoStart + Math.max(cr, 1);
  // Gold section (MYR, no FX)
  const gr = goldHoldings.length;
  const goldHeader = cryptoTotal + 2;
  const goldColHdr = goldHeader + 1;
  const goldStart  = goldColHdr + 1;
  const goldTotal  = goldStart + Math.max(gr, 1);
  const grandRow  = goldTotal + 2;
  const fxHeader  = grandRow + 2;
  const loanHeader = fxHeader + f + 2;
  const totalRows = loanHeader + 10; // enough for loan summary rows

  compactSheet_(dash, totalRows + 4, 6);

  // ── Column widths ───────────────────────────────────────────
  dash.setColumnWidth(1, 30);
  dash.setColumnWidth(2, 220);
  dash.setColumnWidth(3, 160);
  dash.setColumnWidth(4, 60);
  dash.setColumnWidth(5, 10);
  dash.setColumnWidth(6, 160);

  // ── Helpers ─────────────────────────────────────────────────
  function writeColHeaders_(row, col3Label) {
    dash.setRowHeight(row, 30);
    [1,2,3,4,5,6].forEach(c => dash.getRange(row,c).setBackground('#e8eaf6'));
    dash.getRange(row,2).setValue('Name / Market').setFontWeight('bold').setFontSize(10).setFontColor('#3949ab').setVerticalAlignment('middle');
    dash.getRange(row,3).setValue(col3Label).setFontWeight('bold').setFontSize(10).setFontColor('#3949ab').setHorizontalAlignment('right').setVerticalAlignment('middle');
    dash.getRange(row,4).setValue('CCY').setFontWeight('bold').setFontSize(10).setFontColor('#3949ab').setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange(row,6).setValue('≈ RM Equivalent').setFontWeight('bold').setFontSize(10).setFontColor('#3949ab').setHorizontalAlignment('right').setVerticalAlignment('middle');
  }

  function writeDataRow_(row, bg, label, amount, currency, nativeFmt, myrResult, extraNote) {
    dash.setRowHeight(row, 36);
    dash.getRange(row,1,1,6).setBackground(bg);
    dash.getRange(row,2).setValue(extraNote ? label + '  ' + extraNote : label).setFontSize(11).setFontColor('#202124').setVerticalAlignment('middle');
    const balColor = amount < 0 ? '#d93025' : '#202124';
    dash.getRange(row,3).setValue(amount).setNumberFormat(nativeFmt).setFontSize(12).setFontWeight('bold').setFontColor(balColor).setHorizontalAlignment('right').setVerticalAlignment('middle');
    dash.getRange(row,4).setValue(currency).setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.getRange(row,5).setBackground('#e8f0fe');
    if (myrResult.ok) {
      dash.getRange(row,6).setValue(myrResult.myr).setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor(myrResult.myr < 0 ? '#d93025' : '#1a73e8').setHorizontalAlignment('right').setVerticalAlignment('middle');
    } else {
      dash.getRange(row,6).setValue('N/A').setFontSize(10).setFontColor('#d93025').setFontStyle('italic').setHorizontalAlignment('right').setVerticalAlignment('middle');
    }
  }

  function writeTotalRow_(row, label, myrTotal, bgColor, fontSize) {
    dash.setRowHeight(row, 44);
    dash.getRange(row,1,1,6).setBackground(bgColor);
    dash.getRange(row,2).setValue(label).setFontSize(fontSize||12).setFontWeight('bold').setFontColor('#ffffff').setVerticalAlignment('middle');
    dash.getRange(row,5).setBackground(bgColor);
    if (myrTotal !== null) {
      dash.getRange(row,6).setValue(myrTotal).setNumberFormat(myrFmt).setFontSize(fontSize||13).setFontWeight('bold').setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
    } else {
      dash.getRange(row,6).setValue('Partial (FX unavailable)').setFontSize(10).setFontWeight('bold').setFontColor('#fff9c4').setHorizontalAlignment('right').setVerticalAlignment('middle');
    }
  }

  // ── 1. Banner ───────────────────────────────────────────────
  dash.setRowHeight(1, 56);
  dash.getRange(1,1,1,6).merge().setValue('💰  FINANCE DASHBOARD').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#1a73e8').setFontColor('#ffffff');

  // ── 2. Sub-banner ───────────────────────────────────────────
  dash.setRowHeight(2, 26);
  const fxNote = foreignCurrencies.length > 0 ? '  ·  FX rates live-snapshotted at refresh' : '';
  dash.getRange(2,1,1,6).merge().setValue('Last updated: ' + today + fxNote).setFontSize(10).setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f0fe');

  // ── 3. Spacer ───────────────────────────────────────────────
  dash.setRowHeight(3, 16);

  // ── 4. ACCOUNTS section ─────────────────────────────────────
  dash.setRowHeight(4, 26);
  dash.getRange(4,1,1,6).merge().setValue('  🏦  ACCOUNTS').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1565c0').setVerticalAlignment('middle');
  writeColHeaders_(5, 'Balance');

  const ACC_COLORS = ['#e8f0fe','#f3f4f9','#eaf4fb','#e8f5e9','#fefce8'];
  let accMyrTotal  = 0, accFxMissing = false;
  accounts.forEach((acc, i) => {
    const nativeFmt = fmtMap[acc.currency] || myrFmt;
    const myrResult = toMYR(acc.balance, acc.currency);
    if (myrResult.ok) accMyrTotal += myrResult.myr; else accFxMissing = true;
    const row = accStart + i;
    // Use writeDataRow_ for formatting/layout, then overwrite col 3 with a live formula
    // so the balance updates automatically when transactions are added — no menu refresh needed.
    writeDataRow_(row, ACC_COLORS[i%ACC_COLORS.length], acc.name, acc.balance, acc.currency, nativeFmt, myrResult, null);
    // Overwrite static balance with live formula: last value in col F of account sheet
    const safeSheet = acc.name.replace(/'/g, "''");
    // LOOKUP(2,1/(F:F<>""),F:F) finds the LAST non-empty value in col F
    // regardless of gaps — INDEX/COUNTA was returning the first transaction row
    // because COUNTA counts the header too.
    // F3:F1000 skips header rows; ISNUMBER ensures we only match numeric balance cells
    // Use MAX(ARRAYFORMULA(IF(ISNUMBER(...)))) to get the ROW of the last numeric cell
    // then INDEX to retrieve that row's value — works on unsorted data
    const liveFormula = '=IFERROR(INDEX(\'' + safeSheet + '\'!F:F,MAX(ARRAYFORMULA(IF(ISNUMBER(\'' + safeSheet + '\'!F3:F1000),ROW(\'' + safeSheet + '\'!F3:F1000),0)))),0)';
    dash.getRange(row, 3)
      .setFormula(liveFormula)
      .setNumberFormat(nativeFmt)
      .setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    // MYR equivalent: if same currency, reference the live cell; else multiply by FX rate
    const cur = normCurrency_(acc.currency);
    const accFxFormula = cur === 'MYR'
      ? '=' + dash.getRange(row, 3).getA1Notation()
      : '=IFERROR(' + dash.getRange(row, 3).getA1Notation() + '*GOOGLEFINANCE("CURRENCY:' + cur + 'MYR"),0)';
    dash.getRange(row, 6).setFormula(accFxFormula)
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#1a73e8')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    // Also overwrite the Accounts Total row with a live SUM formula
  });
  // Live SUM for total MYR column in accounts section
  if (accounts.length > 0) {
    const sumRange = dash.getRange(accStart, 6, accounts.length, 1).getA1Notation();
    dash.getRange(accTotal, 6)
      .setFormula('=IFERROR(SUM(' + sumRange + '),0)')
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold')
      .setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
  }
  if (!accounts.length) {
    dash.setRowHeight(accStart, 32);
    dash.getRange(accStart,2,1,4).merge().setValue('No accounts found. Create one via the menu.').setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(accTotal, 'Total Cash & Savings', accFxMissing ? null : accMyrTotal, '#1565c0', 12);

  // ── 5. STOCK PORTFOLIOS section ─────────────────────────────
  dash.setRowHeight(stkHeader-1, 16);
  dash.setRowHeight(stkHeader, 26);
  dash.getRange(stkHeader,1,1,6).merge().setValue('  📈  STOCK PORTFOLIOS').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#0d47a1').setVerticalAlignment('middle');
  writeColHeaders_(stkColHdr, 'Portfolio Value');

  const STK_COLORS = ['#e8f5e9','#f1f8e9','#e0f7fa','#fce4ec','#fff3e0'];
  let stkMyrTotal  = 0, stkFxMissing = false;
  portfolios.forEach((port, i) => {
    const nativeFmt = port.stocksFmt || myrFmt;
    const myrResult = toMYR(port.value, port.currency);
    if (myrResult.ok) stkMyrTotal += myrResult.myr; else stkFxMissing = true;
    const glSign = port.gainLoss >= 0 ? '+' : '';
    const glNote = '(' + glSign + port.gainLoss.toFixed(2) + ' ' + port.currency + '  ·  ' + port.count + ' stocks)';
    const row = stkStart + i;
    writeDataRow_(row, STK_COLORS[i%STK_COLORS.length], port.flag + '  ' + port.label, port.value, port.currency, nativeFmt, myrResult, glNote);
    // Overwrite col 3 with live SUM of market value column from portfolio sheet
    const safeStkSheet = getStockSheetName_(port.market).replace(/'/g, "''");
    dash.getRange(row, 3)
      .setFormula('=IFERROR(SUMIF(\'' + safeStkSheet + '\'!A4:A1000,"<>",\'' + safeStkSheet + '\'!F4:F1000),0)')
      .setNumberFormat(nativeFmt).setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    const stkCur = normCurrency_(port.currency);
    const stkFxFormula = stkCur === 'MYR'
      ? '=' + dash.getRange(row, 3).getA1Notation()
      : '=IFERROR(' + dash.getRange(row, 3).getA1Notation() + '*GOOGLEFINANCE("CURRENCY:' + stkCur + 'MYR"),0)';
    dash.getRange(row, 6).setFormula(stkFxFormula)
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#1a73e8')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });
  // Live SUM total for stocks
  if (portfolios.length > 0) {
    const stkSumRange = dash.getRange(stkStart, 6, portfolios.length, 1).getA1Notation();
    dash.getRange(stkTotal, 6).setFormula('=IFERROR(SUM(' + stkSumRange + '),0)')
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold')
      .setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
  }
  if (!portfolios.length) {
    dash.setRowHeight(stkStart, 32);
    dash.getRange(stkStart,2,1,4).merge().setValue('No stock portfolios found. Create one via the menu.').setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(stkTotal, 'Total Stock Portfolio', stkFxMissing ? null : stkMyrTotal, '#0d47a1', 12);

  // ── 6. MUTUAL FUNDS section ──────────────────────────────────
  dash.setRowHeight(mfHeader-1, 16);
  dash.setRowHeight(mfHeader, 26);
  dash.getRange(mfHeader,1,1,6).merge().setValue('  🏦  MUTUAL FUNDS / UNIT TRUST').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#00695c').setVerticalAlignment('middle');
  writeColHeaders_(mfColHdr, 'Market Value');

  const MF_COLORS = ['#e8f5e9','#e0f2f1','#f1f8e9','#e8f5e9','#e0f2f1'];
  let mfMyrTotal   = 0, mfFxMissing = false;
  mfFunds.forEach((mf, i) => {
    const nativeFmt = fmtMap[mf.currency] || myrFmt;
    const myrResult = toMYR(mf.value, mf.currency);
    if (myrResult.ok) mfMyrTotal += myrResult.myr; else mfFxMissing = true;
    const glVal  = mf.value - mf.cost;
    const glSign = glVal >= 0 ? '+' : '';
    const glNote = '(' + glSign + glVal.toFixed(2) + ' ' + mf.currency + '  ·  ' + mf.count + ' fund' + (mf.count !== 1 ? 's' : '') + ')';
    const row = mfStart + i;
    writeDataRow_(row, MF_COLORS[i%MF_COLORS.length], mf.currency + ' Funds', mf.value, mf.currency, nativeFmt, myrResult, glNote);
    // Live SUMIF on MF sheet col F (market value), filtering by currency in col I
    const safeMfSheet = MF_SHEET_NAME.replace(/'/g, "''");
    const mfCcyFilter = mf.currency;
    dash.getRange(row, 3)
      .setFormula('=IFERROR(SUMIF(\'' + safeMfSheet + '\'!I4:I1000,"' + mfCcyFilter + '",\'' + safeMfSheet + '\'!F4:F1000),0)')
      .setNumberFormat(nativeFmt).setFontSize(12).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    const mfCurN = normCurrency_(mf.currency);
    const mfFxFormula = mfCurN === 'MYR'
      ? '=' + dash.getRange(row, 3).getA1Notation()
      : '=IFERROR(' + dash.getRange(row, 3).getA1Notation() + '*GOOGLEFINANCE("CURRENCY:' + mfCurN + 'MYR"),0)';
    dash.getRange(row, 6).setFormula(mfFxFormula)
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#1a73e8')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  });
  // Live SUM total for MF
  if (mfFunds.length > 0) {
    const mfSumRange = dash.getRange(mfStart, 6, mfFunds.length, 1).getA1Notation();
    dash.getRange(mfTotal, 6).setFormula('=IFERROR(SUM(' + mfSumRange + '),0)')
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold')
      .setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
  }
  if (!mfFunds.length) {
    dash.setRowHeight(mfStart, 32);
    dash.getRange(mfStart,2,1,4).merge().setValue('No mutual funds found. Create one via the menu.').setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(mfTotal, 'Total Mutual Funds', mfFxMissing ? null : mfMyrTotal, '#00695c', 12);

  // ── CRYPTO section ────────────────────────────────────────────
  dash.setRowHeight(cryptoHeader - 1, 16);
  dash.setRowHeight(cryptoHeader, 26);
  dash.getRange(cryptoHeader,1,1,6).merge().setValue('  🪙  CRYPTO PORTFOLIO').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#6d1b7b').setVerticalAlignment('middle');
  writeColHeaders_(cryptoColHdr, 'Holdings Value');

  const CRYPTO_COLORS = ['#f3e5f5','#ede7f6','#e8eaf6','#f3e5f5','#ede7f6'];
  let cryptoMyrTotal = 0, cryptoFxMissing = false;

  if (cr > 0) {
    // Group crypto by display currency
    const byCcy = {};
    cryptoHoldings.forEach(r => {
      const qty     = r[2];          // C = quantity
      const buyPrice= r[3];          // D = avg buy price (USD)
      const curPrice= typeof r[4] === 'number' ? r[4] : 0; // E = current price (USD)
      const mktVal  = qty * curPrice;
      const cost    = qty * buyPrice;
      const ccy     = 'USD';         // crypto always priced in USD
      if (!byCcy[ccy]) byCcy[ccy] = { currency: ccy, value: 0, cost: 0, count: 0 };
      byCcy[ccy].value += mktVal;
      byCcy[ccy].cost  += cost;
      byCcy[ccy].count++;
    });
    const cryptoGroups = Object.values(byCcy);
    cryptoGroups.forEach((g, i) => {
      const nativeFmt = '"$"#,##0.00';
      const myrResult = toMYR(g.value, g.currency);
      if (myrResult.ok) cryptoMyrTotal += myrResult.myr; else cryptoFxMissing = true;
      const glVal  = g.value - g.cost;
      const glSign = glVal >= 0 ? '+' : '';
      const glNote = '(' + glSign + glVal.toFixed(2) + ' USD  ·  ' + g.count + ' coin' + (g.count !== 1 ? 's' : '') + ')';
      const row = cryptoStart + i;
      writeDataRow_(row, CRYPTO_COLORS[i % CRYPTO_COLORS.length], '🪙 Crypto Holdings', g.value, g.currency, nativeFmt, myrResult, glNote);
      // Live SUMIF on crypto sheet col F (market value USD), col A non-empty = has a holding
      const safeCryptoSheet = CRYPTO_SHEET_NAME.replace(/'/g, "''");
      dash.getRange(row, 3)
        .setFormula('=IFERROR(SUMIF(\'' + safeCryptoSheet + '\'!A4:A1000,"<>",\'' + safeCryptoSheet + '\'!F4:F1000),0)')
        .setNumberFormat(nativeFmt).setFontSize(12).setFontWeight('bold')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      dash.getRange(row, 6)
        .setFormula('=IFERROR(' + dash.getRange(row, 3).getA1Notation() + '*GOOGLEFINANCE("CURRENCY:USDMYR"),0)')
        .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#1a73e8')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    });
    // Live SUM total for crypto
    const cryptoSumRange = dash.getRange(cryptoStart, 6, cryptoGroups.length, 1).getA1Notation();
    dash.getRange(cryptoTotal, 6).setFormula('=IFERROR(SUM(' + cryptoSumRange + '),0)')
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold')
      .setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
  } else {
    dash.setRowHeight(cryptoStart, 32);
    dash.getRange(cryptoStart,2,1,4).merge().setValue('No crypto holdings found. Create one via the menu.').setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(cryptoTotal, 'Total Crypto Portfolio', cryptoFxMissing ? null : cryptoMyrTotal, '#6d1b7b', 12);

  // ── 7b. GOLD SECTION ─────────────────────────────────────────
  let goldMyrTotal = 0;
  dash.setRowHeight(goldHeader - 1, 16);
  dash.setRowHeight(goldHeader, 26);
  dash.getRange(goldHeader,1,1,6).merge().setValue('  🥇  GOLD PORTFOLIO').setFontSize(11).setFontWeight('bold').setFontColor('#ffffff').setBackground('#e65100').setVerticalAlignment('middle');
  writeColHeaders_(goldColHdr, 'Current Value');
  if (gr > 0) {
    // Group by type (916 / 999)
    const goldByType = {};
    goldHoldings.forEach(r => {
      const type  = r[GC.TYPE-1];
      const wt    = r[GC.WEIGHT-1];
      const cost  = typeof r[GC.BUY_TOTAL-1]  === 'number' ? r[GC.BUY_TOTAL-1]  : 0;
      const val   = typeof r[GC.CUR_VALUE-1]  === 'number' ? r[GC.CUR_VALUE-1]  : (wt * (typeof r[GC.CUR_PRICE-1] === 'number' ? r[GC.CUR_PRICE-1] : r[GC.BUY_PRICE-1]));
      if (!goldByType[type]) goldByType[type] = { type, value: 0, cost: 0, weight: 0, count: 0 };
      goldByType[type].value  += val;
      goldByType[type].cost   += cost;
      goldByType[type].weight += wt;
      goldByType[type].count++;
    });
    const GOLD_COLORS = ['#fff8e1','#fff3e0'];
    Object.values(goldByType).forEach((g, i) => {
      const glVal  = g.value - g.cost;
      const glSign = glVal >= 0 ? '+' : '';
      const glNote = '(' + glSign + 'RM ' + glVal.toFixed(2) + '  ·  ' + g.weight.toFixed(3) + 'g  ·  ' + g.count + ' item' + (g.count !== 1 ? 's' : '') + ')';
      const myrResult = { ok: true, myr: g.value };
      goldMyrTotal += g.value;
      dash.setRowHeight(goldStart + i, 36);
      dash.getRange(goldStart + i,1,1,6).setBackground(GOLD_COLORS[i % GOLD_COLORS.length]);
      dash.getRange(goldStart + i,2).setValue('🥇 ' + g.type + ' Gold  ' + (glNote)).setFontSize(10).setFontColor('#202124').setVerticalAlignment('middle');
      // Live SUMIF on gold sheet col G (current value MYR), filtered by type in col A
      const safeGoldSheet = GOLD_SHEET_NAME.replace(/'/g, "''");
      dash.getRange(goldStart + i,3)
        .setFormula('=IFERROR(SUMIF(\'' + safeGoldSheet + '\'!A4:A1000,"' + g.type + '",\'' + safeGoldSheet + '\'!G4:G1000),0)')
        .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#202124')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      dash.getRange(goldStart + i,4).setValue('MYR').setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
      dash.getRange(goldStart + i,5).setBackground('#e8f0fe');
      dash.getRange(goldStart + i,6)
        .setFormula('=' + dash.getRange(goldStart + i, 3).getA1Notation())
        .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#1a73e8')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    });
  } else {
    dash.setRowHeight(goldStart, 32);
    dash.getRange(goldStart,2,1,4).merge().setValue('No gold holdings found. Create one via the menu.').setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(goldTotal, 'Total Gold Portfolio', goldMyrTotal, '#e65100', 12);
  // Overwrite gold total with live SUM
  if (gr > 0) {
    const goldSumRange = dash.getRange(goldStart, 6, Object.keys(goldByType).length, 1).getA1Notation();
    dash.getRange(goldTotal, 6).setFormula('=IFERROR(SUM(' + goldSumRange + '),0)')
      .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold')
      .setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
  }
  dash.setRowHeight(goldHeader - 1, 16);
  dash.setRowHeight(goldHeader, 26);

  // ── GRAND TOTAL ───────────────────────────────────────────────
  dash.setRowHeight(grandRow-1, 16);
  const anyMissing  = accFxMissing || stkFxMissing || mfFxMissing || cryptoFxMissing;
  const grandTotal  = anyMissing ? null : accMyrTotal + stkMyrTotal + mfMyrTotal + cryptoMyrTotal + goldMyrTotal;
  writeTotalRow_(grandRow, '🏆  TOTAL NET WORTH  (in MYR)', grandTotal, '#0a2d6e', 14);
  // Overwrite grand total with live SUM of all section total rows
  {
    const totalRefs = [accTotal, stkTotal, mfTotal, cryptoTotal, goldTotal]
      .map(r => dash.getRange(r, 6).getA1Notation()).join('+');
    dash.getRange(grandRow, 6).setFormula('=IFERROR(' + totalRefs + ',0)')
      .setNumberFormat(myrFmt).setFontSize(16).setFontWeight('bold')
      .setFontColor('#ffffff').setHorizontalAlignment('right').setVerticalAlignment('middle');
  }
  dash.getRange(grandRow,1,1,6).setBackground('#0a2d6e');
  dash.setRowHeight(grandRow, 52);
  dash.getRange(grandRow,2).setFontSize(14).setFontWeight('bold').setFontColor('#ffffff').setVerticalAlignment('middle');
  if (grandTotal !== null) dash.getRange(grandRow,6).setFontSize(16).setFontWeight('bold');

  // ── 8. FX Rates snapshot ─────────────────────────────────────
  if (foreignCurrencies.length > 0) {
    dash.setRowHeight(fxHeader-1, 16);
    dash.setRowHeight(fxHeader, 26);
    dash.getRange(fxHeader,2,1,5).merge().setValue('📌  FX RATES SNAPSHOT  (MYR base, at time of refresh)').setFontSize(10).setFontWeight('bold').setFontColor('#1a73e8').setBackground('#e8f0fe').setVerticalAlignment('middle');
    foreignCurrencies.forEach((cur, i) => {
      const r    = fxHeader + 1 + i;
      const rate = fxRates[cur];
      dash.setRowHeight(r, 28);
      dash.getRange(r,1,1,6).setBackground(i%2===0 ? '#f8fbff' : '#ffffff');
      dash.getRange(r,2).setValue('1 ' + cur + '  →').setFontSize(11).setFontColor('#5f6368').setVerticalAlignment('middle');
      if (rate !== null && rate !== undefined) {
        dash.getRange(r,3).setValue(rate).setNumberFormat(myrFmt).setFontSize(11).setFontWeight('bold').setFontColor('#0f9d58').setHorizontalAlignment('right').setVerticalAlignment('middle');
        dash.getRange(r,4).setValue('MYR').setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
      } else {
        dash.getRange(r,3,1,3).merge().setValue('Rate unavailable').setFontSize(10).setFontColor('#d93025').setFontStyle('italic').setHorizontalAlignment('left').setVerticalAlignment('middle');
      }
    });
  }

  // ── 9. Loans & Debts summary (info only — not in net worth) ──
  const loansSheet = ss.getSheetByName(LOANS_SHEET_NAME);
  if (loansSheet) {
    const lastLoanRow = loansSheet.getLastRow();
    let totalLent = 0, totalBorrowed = 0, openLent = 0, openBorrowed = 0;
    let lentCount = 0, borrowedCount = 0;

    if (lastLoanRow >= 3) {
      const loanData = loansSheet.getRange(3, 1, lastLoanRow - 2, 8).getValues();
      loanData.forEach(r => {
        const type      = r[0] ? r[0].toString().toUpperCase() : '';
        const amount    = typeof r[3] === 'number' ? r[3] : 0;
        const repaid    = typeof r[5] === 'number' ? r[5] : 0;
        const status    = r[6] ? r[6].toString().toUpperCase() : '';
        const outstanding = Math.max(0, amount - repaid);
        if (type === 'LENT') {
          totalLent += amount;
          if (status !== 'REPAID') { openLent += outstanding; lentCount++; }
        } else if (type === 'BORROWED') {
          totalBorrowed += amount;
          if (status !== 'REPAID') { openBorrowed += outstanding; borrowedCount++; }
        }
      });
    }

    dash.setRowHeight(loanHeader - 1, 16);
    dash.setRowHeight(loanHeader, 28);
    dash.getRange(loanHeader, 2, 1, 5).merge()
      .setValue('🤝  LOANS & DEBTS  (not included in net worth)')
      .setFontSize(10).setFontWeight('bold').setFontColor('#4a148c')
      .setBackground('#ede7f6').setVerticalAlignment('middle');

    // Lent row
    const lentRow = loanHeader + 1;
    dash.setRowHeight(lentRow, 32);
    dash.getRange(lentRow, 1, 1, 6).setBackground('#f3e5f5');
    dash.getRange(lentRow, 2).setValue('💸  Money you LENT  (' + lentCount + ' open)')
      .setFontSize(11).setFontWeight('bold').setFontColor('#6a1b9a').setVerticalAlignment('middle');
    dash.getRange(lentRow, 5).setValue('Outstanding:').setFontSize(10).setFontColor('#6a1b9a').setHorizontalAlignment('right').setVerticalAlignment('middle');
    dash.getRange(lentRow, 6).setValue(openLent).setNumberFormat(myrFmt)
      .setFontSize(12).setFontWeight('bold').setFontColor('#6a1b9a').setHorizontalAlignment('right').setVerticalAlignment('middle');

    // Borrowed row
    const borrowedRow = loanHeader + 2;
    dash.setRowHeight(borrowedRow, 32);
    dash.getRange(borrowedRow, 1, 1, 6).setBackground('#fce4ec');
    dash.getRange(borrowedRow, 2).setValue('🏦  Money you BORROWED  (' + borrowedCount + ' open)')
      .setFontSize(11).setFontWeight('bold').setFontColor('#880e4f').setVerticalAlignment('middle');
    dash.getRange(borrowedRow, 5).setValue('Outstanding:').setFontSize(10).setFontColor('#880e4f').setHorizontalAlignment('right').setVerticalAlignment('middle');
    dash.getRange(borrowedRow, 6).setValue(openBorrowed).setNumberFormat(myrFmt)
      .setFontSize(12).setFontWeight('bold').setFontColor('#880e4f').setHorizontalAlignment('right').setVerticalAlignment('middle');

    // Net exposure row
    const netRow = loanHeader + 3;
    const netExposure = openLent - openBorrowed;
    dash.setRowHeight(netRow, 32);
    dash.getRange(netRow, 1, 1, 6).setBackground('#ede7f6');
    dash.getRange(netRow, 2).setValue('Net Loan Exposure  (Lent − Borrowed)')
      .setFontSize(10).setFontWeight('bold').setFontColor('#4a148c').setVerticalAlignment('middle');
    dash.getRange(netRow, 5).setValue('Net:').setFontSize(10).setFontColor('#4a148c').setHorizontalAlignment('right').setVerticalAlignment('middle');
    dash.getRange(netRow, 6).setValue(netExposure).setNumberFormat(myrFmt)
      .setFontSize(11).setFontWeight('bold')
      .setFontColor(netExposure >= 0 ? '#6a1b9a' : '#880e4f')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  }

  dash.setHiddenGridlines(true);
  dash.setFrozenRows(2);
  dash.activate();
}


// ── REFRESH STOCK PRICES ─────────────────────────────────────
function refreshAllStockPrices() {
  const sheet = getActivePortfolioSheet_();
  if (!sheet) return;
  refreshPortfolioSheet_(sheet);
  SpreadsheetApp.getUi().alert('Prices refreshed!\n\nTip: Enable Auto-Refresh (5 min) from the menu to keep prices updated automatically.');
}

// ── DIVIDEND INFO (stockanalysis.com) ────────────────────────
/**
 * Fetches dividend amount, ex-date and pay date from stockanalysis.com
 * for the currently active stock portfolio sheet.
 * Writes into cols 10 (Dividend), 11 (Ex-Date), 12 (Pay Date).
 *
 * URL patterns:
 *   MY  → https://stockanalysis.com/quote/klse/{CODE}/dividend/
 *   US  → https://stockanalysis.com/stocks/{code}/dividend/
 *   SG  → https://stockanalysis.com/quote/sgx/{CODE}/dividend/
 *   HK  → https://stockanalysis.com/quote/hkg/{CODE}/dividend/
 *   CN  → not supported (N/A)
 */
function refreshDividendInfo() {
  // Writes IMPORTHTML formulas to Dividend/Ex-Date/Pay-Date columns for all rows.
  // Run once to migrate existing rows; new rows get formulas automatically on add.
  const ui    = SpreadsheetApp.getUi();
  const sheet = getActivePortfolioSheet_();
  if (!sheet) { ui.alert('Please open a Stock Portfolio sheet first.'); return; }

  const market  = getSheetMarket_(sheet);
  const lastRow = getLastStockRow_(sheet);
  if (lastRow < 4) { ui.alert('No stocks found in this portfolio.'); return; }

  if (market === 'CN') {
    ui.alert('Dividend info is not available for China A-Shares on StockAnalysis.com.');
    return;
  }

  const cfg = MARKET_CONFIG[market];
  let updated = 0;

  for (let row = 4; row <= lastRow; row++) {
    const code = sheet.getRange(row, 1).getValue();
    if (!code) continue;

    const aRef = 'B' + row;  // Col B = ticker (e.g. MAYBANK), not numeric code
    // stockanalysis.com table structure (single table per page):
    // col1=Ex-Date, col2=Amount ("$0.26" for US, "0.06 MYR" for MY etc.), col4=Pay Date
    // Dates are in "Feb 9, 2026" format; CLEAN+TRIM normalises whitespace before DATEVALUE.
    var divAmtFml, exDateFml, payDateFml, divYldFml;
    const xpYld = '/html/body/div/div[1]/div[2]/main/div[2]/div/div[2]/div[1]/div';
    if (market === 'US') {
      // US stocks use /stocks/{ticker}/ and ETFs use /etf/{ticker}/ — try stocks first, ETF as fallback
      const stockUrl = 'CONCATENATE("https://stockanalysis.com/stocks/",LOWER(' + aRef + '),"/dividend/")';
      const etfUrl   = 'CONCATENATE("https://stockanalysis.com/etf/",LOWER(' + aRef + '),"/dividend/")';
      divAmtFml = '=IFERROR('
        + 'VALUE(SUBSTITUTE(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,2),"$","")),'
        + 'IFERROR('
        + 'VALUE(SUBSTITUTE(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,2),"$","")),'
        + '"No Data"))';
      exDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,1)))),'
        + 'IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,1)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,1))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,1))),'
        + '"No Data"))))';
      payDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,4)))),'
        + 'IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,4)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + stockUrl + ',"table",1),2,4))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + etfUrl + ',"table",1),2,4))),'
        + '"No Data"))))';
      divYldFml = '=IFERROR('
        + 'IMPORTXML(' + stockUrl + ',"' + xpYld + '")*100,'
        + 'IFERROR('
        + 'IMPORTXML(' + etfUrl + ',"' + xpYld + '")*100,'
        + '"N/A"))';
    } else {
      var divUrl;
      if      (market === 'MY') divUrl = 'CONCATENATE("https://stockanalysis.com/quote/klse/",' + aRef + ',"/dividend/")';
      else if (market === 'SG') divUrl = 'CONCATENATE("https://stockanalysis.com/quote/sgx/",' + aRef + ',"/dividend/")';
      else if (market === 'HK') divUrl = 'CONCATENATE("https://stockanalysis.com/quote/hkg/",' + aRef + ',"/dividend/")';
      else continue;
      var ccy = market === 'MY' ? ' MYR' : market === 'SG' ? ' SGD' : ' HKD';
      divAmtFml = '=IFERROR(VALUE(SUBSTITUTE(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,2),"' + ccy + '","")),"No Data")';
      exDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,1)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,1))),'
        + '"No Data"))';
      payDateFml = '=IFERROR('
        + 'DATEVALUE(CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,4)))),'
        + 'IFERROR('
        + 'CLEAN(TRIM(INDEX(IMPORTHTML(' + divUrl + ',"table",1),2,4))),'
        + '"No Data"))';
      divYldFml = '=IFERROR(IMPORTXML(' + divUrl + ',"' + xpYld + '")*100,"N/A")';
    }

    sheet.getRange(row, 10).setFormula(divAmtFml)
      .setNumberFormat('0.0000').setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(row, 11).setFormula(divYldFml)
      .setNumberFormat('0.0000"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(row, 12).setFormula(exDateFml)
      .setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(row, 13).setFormula(payDateFml)
      .setNumberFormat('dd-mmm-yyyy').setHorizontalAlignment('center').setVerticalAlignment('middle');
    updated++;
  }

  ui.alert('✅ Dividend Formulas Set',
    updated + ' row(s) updated.\n\n' +
    'Dividend, Ex-Date and Pay-Date will now load automatically\n' +
    'via IMPORTHTML from StockAnalysis.com.\n\n' +
    'Values may take 10-30 seconds to appear as Google Sheets\n' +
    'fetches the data in the background.',
    ui.ButtonSet.OK);
}

/**
 * Fetches dividend info for one stock from stockanalysis.com.
 * Returns { amount, exDate, payDate } or null if unavailable.
 */
function fetchDividendInfo_(code, market) {
  // StockAnalysis.com is a Svelte/SvelteKit app — the HTML page contains no dividend data.
  // Instead we call the SvelteKit __data.json endpoint which returns the page data as JSON.
  //
  // JSON structure (nodes[2].data):
  //   infoTable: { exdiv, annual, yield, frequency, payoutRatio, growth }
  //   history:   [ { dt, amt, record, pay }, ... ]  (most recent first)
  try {
    const c = code.toString().trim();
    let dataUrl;
    if      (market === 'MY') dataUrl = 'https://stockanalysis.com/quote/klse/' + encodeURIComponent(c) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else if (market === 'US') dataUrl = 'https://stockanalysis.com/stocks/'     + encodeURIComponent(c.toLowerCase()) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else if (market === 'SG') dataUrl = 'https://stockanalysis.com/quote/sgx/'  + encodeURIComponent(c) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else if (market === 'HK') dataUrl = 'https://stockanalysis.com/quote/hkg/'  + encodeURIComponent(c) + '/dividend/__data.json?x-sveltekit-trailing-slash=1';
    else return null;

    const res = UrlFetchApp.fetch(dataUrl, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'application/json'
      }
    });
    if (res.getResponseCode() !== 200) {
      Logger.log('fetchDividendInfo_ HTTP ' + res.getResponseCode() + ' for ' + code);
      return null;
    }

    const json = JSON.parse(res.getContentText());

    // SvelteKit data format: nodes array, dividend page is always the last node
    // Each node uses a deduplication scheme: data[] array + uses index references
    // We need to resolve the actual values from the data array
    const nodes = json.nodes || [];
    // Find the node containing dividend data (has 'infoTable' or 'history' key)
    let divNode = null;
    for (var n = nodes.length - 1; n >= 0; n--) {
      const nd = nodes[n];
      if (!nd || nd.type !== 'data' || !nd.data) continue;
      // The data array is flat; keys are objects with index references
      // Find the root object which has infoTable and history keys
      const dataArr = nd.data;
      for (var i = 0; i < dataArr.length; i++) {
        const item = dataArr[i];
        if (item && typeof item === 'object' && 'infoTable' in item && 'history' in item) {
          divNode = { dataArr: dataArr, root: item };
          break;
        }
      }
      if (divNode) break;
    }

    if (!divNode) {
      Logger.log('fetchDividendInfo_ no dividend node found for ' + code);
      return null;
    }

    const dataArr = divNode.dataArr;
    const root    = divNode.root;

    // Helper: resolve an index reference to actual value
    function resolve(ref) {
      if (typeof ref === 'number') return dataArr[ref];
      return ref;
    }

    // Get infoTable (has exdiv field)
    const infoTableRef = root.infoTable;
    const infoTable    = resolve(infoTableRef);
    const exDateStr    = infoTable ? resolve(infoTable.exdiv) : null;

    // Get history array (array of row index references)
    const historyRef  = root.history;
    const historyIdxs = resolve(historyRef); // array of indices into dataArr
    let amount  = null;
    let exDate  = exDateStr || null;
    let payDate = null;

    if (Array.isArray(historyIdxs) && historyIdxs.length > 0) {
      // Most recent dividend is first
      const firstRowRef = historyIdxs[0];
      const firstRow    = resolve(firstRowRef);
      if (firstRow && typeof firstRow === 'object') {
        const dtRaw  = resolve(firstRow.dt);
        const amtRaw = resolve(firstRow.amt);
        const payRaw = resolve(firstRow.pay);

        // dt is ISO format "2026-03-12" — convert to dd/MM/yyyy
        if (dtRaw && !exDate) {
          const parts = dtRaw.toString().split('-');
          if (parts.length === 3) exDate = parts[2] + '/' + parts[1] + '/' + parts[0];
        }

        // amt is like "0.330 MYR" or "$0.330"
        if (amtRaw) {
          const amtMatch = amtRaw.toString().match(/([\d.]+)/);
          if (amtMatch) amount = parseFloat(amtMatch[1]);
        }

        // pay is ISO format "2026-03-26"
        if (payRaw) {
          const parts = payRaw.toString().split('-');
          if (parts.length === 3) payDate = parts[2] + '/' + parts[1] + '/' + parts[0];
        }
      }
    }

    // Also convert exDate from "Mar 12, 2026" format to dd/MM/yyyy if needed
    if (exDate && exDate.includes(',')) {
      try {
        const d = new Date(exDate);
        if (!isNaN(d)) {
          const dd = String(d.getDate()).padStart(2, '0');
          const mm = String(d.getMonth() + 1).padStart(2, '0');
          exDate = dd + '/' + mm + '/' + d.getFullYear();
        }
      } catch(_) {}
    }
    if (payDate && payDate.includes(',')) {
      try {
        const d = new Date(payDate);
        if (!isNaN(d)) {
          const dd = String(d.getDate()).padStart(2, '0');
          const mm = String(d.getMonth() + 1).padStart(2, '0');
          payDate = dd + '/' + mm + '/' + d.getFullYear();
        }
      } catch(_) {}
    }

    if (!exDate && !amount) return null;

    return {
      amount:  amount  || null,
      exDate:  exDate  || 'N/A',
      payDate: payDate || 'N/A'
    };
  } catch(e) {
    Logger.log('fetchDividendInfo_ error for ' + code + ': ' + e.message);
    return null;
  }
}


function autoRefreshAllPortfolios_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['MY','US','SG','HK','CN'].forEach(mkt => {
    const sheet = ss.getSheetByName(getStockSheetName_(mkt));
    if (!sheet) return;
    if (getLastStockRow_(sheet) < 4) return;
    refreshPortfolioSheet_(sheet);
  });
  const mfSheet = ss.getSheetByName(MF_SHEET_NAME);
  if (mfSheet && getLastFundRow_(mfSheet) >= 4) refreshMutualFundNavs_(mfSheet);
  // Refresh crypto portfolio
  const cryptoSheet = ss.getSheetByName(CRYPTO_SHEET_NAME);
  if (cryptoSheet && getLastCryptoRow_(cryptoSheet) >= 4) refreshCryptoPrices_(cryptoSheet);
  // Refresh gold portfolio
  const goldSheet = ss.getSheetByName(GOLD_SHEET_NAME);
  if (goldSheet && getLastGoldRow_(goldSheet) >= 4) refreshGoldPrices_(goldSheet);
}

function refreshPortfolioSheet_(sheet) {
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;
  const market  = getSheetMarket_(sheet);
  const cfg     = MARKET_CONFIG[market];

  for (let row = 4; row <= getLastStockRow_(sheet); row++) {
    const code = sheet.getRange(row, 1).getValue();
    if (!code) continue;
    const cell    = sheet.getRange(row, 5);
    const formula = cell.getFormula();
    if (formula) { cell.clearContent(); SpreadsheetApp.flush(); cell.setFormula(formula); }
    sheet.getRange(row, 14).setValue(today);
  }
  SpreadsheetApp.flush();
  for (let row = 4; row <= getLastStockRow_(sheet); row++) {
    const gain = sheet.getRange(row, 7).getValue();
    if (typeof gain !== 'number') continue;
    const c = gain < 0 ? '#d93025' : '#0f9d58';
    sheet.getRange(row, 7).setFontColor(c);
    sheet.getRange(row, 8).setFontColor(c);
  }
  refreshStockSummary_(sheet, cfg);
}


// ── STOCK SUMMARY ────────────────────────────────────────────
function refreshStockSummary_(sheet, cfg) {
  if (!cfg) { const market = getSheetMarket_(sheet); cfg = MARKET_CONFIG[market] || MARKET_CONFIG['MY']; }
  const lastStockRow = getLastStockRow_(sheet);
  if (lastStockRow < 4) return;

  const data = sheet.getRange(4, 1, lastStockRow-3, 15).getValues()
    .filter(r => r[0] !== '' && typeof r[2] === 'number' && r[2] > 0);

  const totalCost  = data.reduce((s,r) => s + (r[2]*r[3]), 0);
  const totalValue = data.reduce((s,r) => s + (typeof r[5] === 'number' ? r[5] : 0), 0);
  const totalGain  = totalValue - totalCost;
  const gainPct    = totalCost > 0 ? (totalGain / totalCost * 100) : 0;
  const totalDiv   = data.reduce((s,r) => s + (typeof r[9] === 'number' ? r[2]*r[9] : 0), 0);
  const glColor    = totalGain >= 0 ? '#0f9d58' : '#d93025';

  // Clear below last stock row
  const maxRows    = sheet.getMaxRows();
  const clearStart = lastStockRow + 1;
  if (clearStart <= maxRows) sheet.getRange(clearStart, 1, maxRows-clearStart+1, 15).clearContent().clearFormat();

  const summaryRow = lastStockRow + 2;
  sheet.setRowHeight(summaryRow, 30);
  sheet.getRange(summaryRow, 1, 1, 15).merge()
    .setValue('📊  PORTFOLIO SUMMARY').setBackground('#0d47a1').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');

  [
    ['Total Holdings',       data.length + ' stocks', null,         '#202124'],
    ['Total Cost',           totalCost,               cfg.sumFmt,   '#202124'],
    ['Portfolio Value',      totalValue,              cfg.sumFmt,   '#202124'],
    ['Gain / Loss',          totalGain,               cfg.sumFmt,   glColor  ],
    ['Gain / Loss (%)',      gainPct,                 '0.00"%"',    glColor  ],
    ['Est. Annual Dividends',totalDiv,                cfg.divFmt,   '#1a73e8'],
  ].forEach(([label, val, fmt, color], i) => {
    const r = summaryRow + 1 + i;
    sheet.setRowHeight(r, 28);
    sheet.getRange(r, 1, 1, 2).merge().setValue(label).setFontWeight('bold').setFontSize(10).setBackground(i%2===0 ? '#e8eaf6' : '#f3f4f9').setVerticalAlignment('middle');
    const vc = sheet.getRange(r, 3);
    vc.setValue(val).setFontColor(color).setFontWeight('bold').setVerticalAlignment('middle');
    if (fmt) vc.setNumberFormat(fmt);
  });
}


// ── AUTO-REFRESH (5-min stocks) ───────────────────────────────
function enableAutoRefresh() {
  const ui    = SpreadsheetApp.getUi();
  const existing = getAutoRefreshTrigger_();
  if (existing) { ui.alert('Auto-refresh is already enabled.\n\nPrices refresh every 5 minutes automatically.'); return; }
  const trigger = ScriptApp.newTrigger('autoRefreshAllPortfolios_').timeBased().everyMinutes(5).create();
  PropertiesService.getScriptProperties().setProperty(AUTO_REFRESH_TRIGGER_KEY, trigger.getUniqueId());
  ui.alert('Auto-Refresh Enabled!\n\nAll stock portfolio prices will now refresh every 5 minutes.\n\nTo stop: use "Disable Auto-Refresh" from the menu.');
}

function disableAutoRefresh() {
  const ui      = SpreadsheetApp.getUi();
  const trigger = getAutoRefreshTrigger_();
  if (!trigger) { ui.alert('Auto-refresh is not currently enabled.'); return; }
  ScriptApp.deleteTrigger(trigger);
  PropertiesService.getScriptProperties().deleteProperty(AUTO_REFRESH_TRIGGER_KEY);
  ui.alert('Auto-Refresh Disabled.\n\nPrices will no longer update automatically.');
}

function getAutoRefreshTrigger_() {
  const savedId = PropertiesService.getScriptProperties().getProperty(AUTO_REFRESH_TRIGGER_KEY);
  if (!savedId) return null;
  return ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === savedId) || null;
}


// ── DAILY NAV REFRESH (Mutual Funds, ~7 PM) ───────────────────
function enableDailyNavRefresh() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(DAILY_NAV_TRIGGER_KEY);
  if (existingId) {
    const exists = ScriptApp.getProjectTriggers().some(t => t.getUniqueId() === existingId);
    if (exists) { ui.alert('Daily NAV update is already enabled.\n\nMutual fund prices refresh automatically at ~7 PM every day.'); return; }
  }
  const trigger = ScriptApp.newTrigger('dailyNavRefreshJob_').timeBased().everyDays(1).atHour(19).create();
  props.setProperty(DAILY_NAV_TRIGGER_KEY, trigger.getUniqueId());
  ui.alert('Daily NAV Update Enabled!\n\nMutual fund NAV prices will refresh every day at ~7 PM.\n\nFSMOne NAV prices update once daily after market close.\n\nTo stop: use "Disable Daily NAV Update" from the menu.');
}

function disableDailyNavRefresh() {
  const ui      = SpreadsheetApp.getUi();
  const props   = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(DAILY_NAV_TRIGGER_KEY);
  if (!savedId) { ui.alert('Daily NAV update is not currently enabled.'); return; }
  const trigger = ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === savedId);
  if (trigger) ScriptApp.deleteTrigger(trigger);
  props.deleteProperty(DAILY_NAV_TRIGGER_KEY);
  ui.alert('Daily NAV Update Disabled.\n\nMutual fund prices will no longer auto-update.\n\nYou can still refresh manually via "Refresh All Prices Now".');
}

function dailyNavRefreshJob_() {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const mfSheet = ss.getSheetByName(MF_SHEET_NAME);
    if (!mfSheet || getLastFundRow_(mfSheet) < 4) return;
    refreshMutualFundNavs_(mfSheet);
    const dash = ss.getSheetByName(DASH_NAME);
    if (dash) renderDashboardSheet();
  } catch(e) {
    console.error('dailyNavRefreshJob_ error: ' + e.message);
  }
}



// ============================================================

// ============================================================
//  CRYPTO PORTFOLIO TRACKER  v3
//  Sheet: 🪙 Crypto Portfolio
//  Price source: cryptoprices.cc  (no API key, 10,000+ tokens)
//  Usage: GET https://cryptoprices.cc/BTC  → returns plain USD price
//  404 = token not found
//  Prices written as VALUES by refreshCryptoPrices_() — no formula caching issues
// ============================================================

const CRYPTO_SHEET_NAME = '🪙 Crypto Portfolio';

const CC = {
  SYMBOL:      1,   // A
  NAME:        2,   // B
  QTY:         3,   // C
  BUY:         4,   // D – avg buy price USD
  PRICE:       5,   // E – current price USD (value, not formula)
  MKT_VAL:     6,   // F – market value USD (=C*E)
  GAIN:        7,   // G – gain/loss USD
  GAIN_PCT:    8,   // H – gain %
  CHG24H:      9,   // I – 24h change %
  MKT_VAL_MYR: 10,  // J – market value in RM (written on refresh)
  ACCOUNT:     11,  // K
  UPDATED:     12,  // L
  NOTES:       13,  // M
};

const POPULAR_CRYPTO = [
  {symbol:'BTC',  name:'Bitcoin'},
  {symbol:'ETH',  name:'Ethereum'},
  {symbol:'BNB',  name:'BNB'},
  {symbol:'SOL',  name:'Solana'},
  {symbol:'XRP',  name:'XRP'},
  {symbol:'USDT', name:'Tether'},
  {symbol:'USDC', name:'USD Coin'},
  {symbol:'ADA',  name:'Cardano'},
  {symbol:'DOGE', name:'Dogecoin'},
  {symbol:'TRX',  name:'TRON'},
  {symbol:'TON',  name:'Toncoin'},
  {symbol:'AVAX', name:'Avalanche'},
  {symbol:'LINK', name:'Chainlink'},
  {symbol:'DOT',  name:'Polkadot'},
  {symbol:'MATIC',name:'Polygon'},
  {symbol:'UNI',  name:'Uniswap'},
  {symbol:'NEAR', name:'NEAR Protocol'},
  {symbol:'ATOM', name:'Cosmos'},
  {symbol:'LTC',  name:'Litecoin'},
  {symbol:'XLM',  name:'Stellar'},
  {symbol:'SUI',  name:'Sui'},
  {symbol:'NEO',  name:'NEO'},
  {symbol:'INJ',  name:'Injective'},
  {symbol:'XAUT', name:'Tether Gold'},
  {symbol:'PAXG', name:'PAX Gold'},
  {symbol:'RENDER',name:'Render'},
  {symbol:'PEPE', name:'Pepe'},
  {symbol:'WIF',  name:'dogwifhat'},
  {symbol:'JUP',  name:'Jupiter'},
  {symbol:'FET',  name:'Fetch.ai'},
];


// ── PRICE FETCH (single) ──────────────────────────────────────
/**
 * Fetches the USD price for one symbol from cryptoprices.cc.
 * Returns { price } on success, or { error: true, message } on failure.
 */
function fetchCryptoPrice_(symbol) {
  symbol = symbol.toString().trim().toUpperCase();
  const url = 'https://cryptoprices.cc/' + encodeURIComponent(symbol);
  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 404) return { error: true, message: '"' + symbol + '" not found. Check the token symbol.' };
    if (code !== 200) return { error: true, message: 'HTTP ' + code + ' for ' + symbol };
    const text  = res.getContentText().trim();
    const price = parseFloat(text);
    if (isNaN(price)) return { error: true, message: 'Unexpected response: ' + text.substring(0, 40) };
    return { price: price };
  } catch(e) {
    return { error: true, message: e.message };
  }
}

/**
 * Batch-fetches prices for multiple symbols.
 * Uses parallel UrlFetchApp.fetchAll() for speed.
 * Returns map: { BTC: 69913, ETH: 2051, ... }  (missing = not found)
 */
function fetchCryptoBatch_(symbols) {
  if (!symbols || symbols.length === 0) return {};
  const upper    = symbols.map(s => s.toString().trim().toUpperCase());
  const requests = upper.map(sym => ({
    url:               'https://cryptoprices.cc/' + encodeURIComponent(sym),
    muteHttpExceptions: true,
  }));

  const result = {};
  try {
    const responses = UrlFetchApp.fetchAll(requests);
    responses.forEach((res, i) => {
      if (res.getResponseCode() !== 200) return;
      const price = parseFloat(res.getContentText().trim());
      if (!isNaN(price)) result[upper[i]] = price;
    });
  } catch(e) {
    // fetchAll failed — fall back to sequential
    upper.forEach(sym => {
      const d = fetchCryptoPrice_(sym);
      if (!d.error) result[sym] = d.price;
    });
  }
  return result;
}

/**
 * Fetches live USD->MYR rate using a temporary sheet + GOOGLEFINANCE.
 * Returns the rate (number) or null on failure.
 */
function fetchUsdMyrRate_() {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = '_FX_TEMP_';
  let usdMyr = null;
  try {
    let temp = ss.getSheetByName(tempName);
    if (temp) ss.deleteSheet(temp);
    temp = ss.insertSheet(tempName);
    compactSheet_(temp, 10, 1); // 1 col × 10 rows
    temp.getRange(1, 1).setFormula('=GOOGLEFINANCE("CURRENCY:USDMYR")');
    SpreadsheetApp.flush();
    Utilities.sleep(3000);
    const val = temp.getRange(1, 1).getValue();
    if (typeof val === 'number' && val > 0) usdMyr = val;
    ss.deleteSheet(temp);
  } catch(e) {
    try { const t = ss.getSheetByName(tempName); if (t) ss.deleteSheet(t); } catch(_) {}
  }
  return usdMyr;
}

/** Called from Add Crypto dialog */
function getCryptoPriceForDialog(symbol) {
  const d = fetchCryptoPrice_(symbol);
  return JSON.stringify(d);
}


// ── CREATE SHEET ──────────────────────────────────────────────
function createCryptoSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CRYPTO_SHEET_NAME);
  if (sheet) {
    ui.alert('The Crypto Portfolio sheet already exists.\n\nClick the "🪙 Crypto Portfolio" tab to view it.');
    sheet.activate();
    return;
  }
  sheet = ss.insertSheet(CRYPTO_SHEET_NAME, 1);
  buildCryptoSheet_(sheet);
  sheet.activate();
  ui.alert(
    'Crypto Portfolio sheet created!\n\n' +
    'Use "Add Crypto" from the menu to add your first coin.\n\n' +
    'Prices via cryptoprices.cc — 10,000+ tokens supported.\n' +
    'Use the exact CoinMarketCap ticker symbol (e.g. BTC, XAUT, SUI).\n\n' +
    'Tip: "Refresh All Prices Now" updates all prices in one batch call.'
  );
}

function buildCryptoSheet_(sheet) {
  // 13 cols: Symbol|Name|Qty|BuyPx|CurPx|MktValUSD|Gain|Gain%|24h|ValRM|Account|Updated|Notes
  const widths = [80,170,110,125,125,125,125,85,85,130,140,125,170];
  widths.forEach((w,i) => sheet.setColumnWidth(i+1, w));

  sheet.setRowHeight(1, 52);
  sheet.getRange(1,1,1,13).merge()
    .setValue('🧹  CRYPTO PORTFOLIO')
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(2, 26);
  sheet.getRange(2,1,1,13).merge()
    .setValue('Live prices via cryptoprices.cc  ·  USD & RM values updated on each refresh')
    .setBackground('#f3e5f5').setFontColor('#6d1b7b')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(3, 38);
  const headers = ['Symbol','Name','Quantity',
    'Avg Buy\nPrice (USD)','Current\nPrice (USD)',
    'Market Value\n(USD)','Gain/Loss\n(USD)','Gain/Loss %','24h\nChange %',
    'Value (RM)','Linked Account','Last Updated','Notes'];
  sheet.getRange(3,1,1,13).setValues([headers])
    .setBackground('#6d1b7b').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  // Highlight RM column in gold to make it stand out
  sheet.getRange(3, CC.MKT_VAL_MYR)
    .setBackground('#1a237e').setFontColor('#ffe082');

  sheet.getRange('D4:G53').setNumberFormat('"$"#,##0.00######');
  sheet.getRange('H4:I53').setNumberFormat('0.00"%"');
  sheet.getRange('C4:C53').setNumberFormat('#,##0.########');
  sheet.getRange('J4:J53').setNumberFormat('"RM "#,##0.00');
  sheet.setFrozenRows(3);
  sheet.setHiddenGridlines(true);
}
function getLastCryptoRow_(sheet) {
  // Valid crypto row has SYMBOL in col A AND NAME in col B
  // Summary rows only have text in col A (merged), col B is empty
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 3) return 3;
  const data = sheet.getRange(4, CC.SYMBOL, maxRows-3, 2).getValues();
  let last = 3;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][1] !== '') last = i + 4;
  }
  return last;
}


// ── ADD CRYPTO DIALOG ─────────────────────────────────────────
function showAddCryptoDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(CRYPTO_SHEET_NAME)) {
    const resp = ui.alert('No Crypto Portfolio sheet found.', 'Create it now?', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.YES) createCryptoSheet(); else return;
  }

  const allAccounts  = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }));
  const accountsJson = JSON.stringify(allAccounts);
  const popularJson  = JSON.stringify(POPULAR_CRYPTO);

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 12px;font-size:15px;color:#6d1b7b}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    'input:focus,select:focus{outline:none;border-color:#6d1b7b}' +
    '.sr{display:flex;gap:6px}.sr input{flex:1}' +
    '.lbtn{padding:8px 14px;background:#6d1b7b;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;white-space:nowrap}' +
    '.lbtn:hover{background:#4a148c}' +
    '.grid{display:flex;flex-wrap:wrap;gap:5px;margin-top:6px}' +
    '.chip{padding:3px 10px;background:#f3e5f5;color:#6d1b7b;border:1px solid #ce93d8;border-radius:20px;cursor:pointer;font-size:11px;font-weight:600}' +
    '.chip:hover{background:#e1bee7}' +
    '.pbox{background:#f3e5f5;border:1px solid #ce93d8;border-radius:8px;padding:10px 12px;margin-top:8px;display:none}' +
    '.pval{font-size:16px;font-weight:bold;color:#4a148c}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.cprev{background:#ede7f6;border-radius:6px;padding:7px 10px;margin-top:6px;font-size:12px;color:#4a148c;display:none}' +
    '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px}' +
    '.err{background:#fce8e6;border:1px solid #f28b82;border-radius:6px;padding:8px 10px;font-size:12px;color:#c62828;margin-top:6px;display:none}' +
    '.spin{color:#9aa0a6;font-size:12px;margin-top:4px;display:none}' +
    '.hint{font-size:11px;color:#9aa0a6;margin-top:3px}' +
    '.br{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.bok{background:#6d1b7b;color:#fff}.bok:hover{background:#4a148c}' +
    '.bca{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>🪙 Add Crypto Holding</h2>' +
    '<label>Token Symbol</label>' +
    '<div class="sr">' +
    '<input type="text" id="sym" placeholder="e.g. BTC, ETH, XAUT, SUI, NEO..." />' +
    '<button class="lbtn" onclick="lookup()">Get Price</button>' +
    '</div>' +
    '<p class="hint">Use the exact CoinMarketCap ticker. Powered by cryptoprices.cc (10,000+ tokens).</p>' +
    '<div class="spin" id="spin">⏳ Fetching price from cryptoprices.cc...</div>' +
    '<div class="err" id="err"></div>' +
    '<label style="margin-top:10px">Quick pick</label>' +
    '<div class="grid" id="grid"></div>' +
    '<div class="pbox" id="pbox"><span class="pval" id="pval"></span></div>' +
    '<div id="form" style="display:none;margin-top:10px">' +
    '<label>Coin / Token Name</label>' +
    '<input type="text" id="cname" placeholder="Full name (e.g. Bitcoin)" />' +
    '<div class="row2" style="margin-top:8px">' +
    '<div><label>Quantity</label><input type="number" id="qty" placeholder="e.g. 0.5" step="any" min="0" oninput="calcCost()"/></div>' +
    '<div><label>Avg Buy Price (USD)</label><input type="number" id="buypx" placeholder="auto-filled" step="any" min="0" oninput="calcCost()"/></div>' +
    '</div>' +
    '<div class="cprev" id="cprev"></div>' +
    '<label>Deduct from Account <span style="font-weight:400">(USD or RM, optional — FX applied automatically)</span></label>' +
    '<select id="acct"><option value="">— None —</option></select>' +
    '<div class="warn" id="nousd" style="display:none">No accounts found. Create an account first.</div>' +
    '<label>Notes (optional)</label>' +
    '<input type="text" id="notes" placeholder="e.g. DCA, hardware wallet" />' +
    '</div>' +
    '<div class="br">' +
    '<button class="btn bca" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn bok" id="saveBtn" onclick="save()" style="display:none">Add Holding</button>' +
    '</div>' +
    '<script>' +
    'var AA=' + accountsJson + ';' +
    'var POP=' + popularJson + ';' +
    'var selSym="",selName="",selPrice=0;' +
    'var g=document.getElementById("grid");' +
    'POP.forEach(function(c){' +
    '  var b=document.createElement("div");b.className="chip";b.textContent=c.symbol;b.title=c.name;' +
    '  b.onclick=function(){document.getElementById("sym").value=c.symbol;selName=c.name;lookup();};' +
    '  g.appendChild(b);' +
    '});' +
    'function lookup(){' +
    '  var s=document.getElementById("sym").value.trim().toUpperCase();if(!s)return;' +
    '  selSym=s;' +
    '  document.getElementById("spin").style.display="block";' +
    '  document.getElementById("pbox").style.display="none";' +
    '  document.getElementById("err").style.display="none";' +
    '  google.script.run.withSuccessHandler(onPrice).withFailureHandler(function(e){showErr(e.message);}).getCryptoPriceForDialog(s);' +
    '}' +
    'function onPrice(json){' +
    '  document.getElementById("spin").style.display="none";' +
    '  var d=JSON.parse(json);' +
    '  if(d.error){showErr(d.message);return;}' +
    '  selPrice=d.price;' +
    '  document.getElementById("pval").textContent=selSym+" = $"+selPrice.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:8});' +
    '  document.getElementById("pbox").style.display="block";' +
    '  document.getElementById("buypx").value=selPrice;' +
    '  if(!document.getElementById("cname").value&&selName)document.getElementById("cname").value=selName;' +
    '  document.getElementById("form").style.display="block";' +
    '  document.getElementById("saveBtn").style.display="inline-block";' +
    '  buildAccts();calcCost();' +
    '}' +
    'function buildAccts(){var sel=document.getElementById(\'acct\');sel.innerHTML=\'<option value=\"\">— None (no deduction) —</option>\';AA.forEach(function(a){sel.innerHTML+=\'<option value=\"\'+ a.name +\'\">\'+ a.name +\' (\' + a.currency + \')</option>\';});document.getElementById(\'nousd\').style.display=AA.length===0?\'block\':\'none\';}' +
    'function calcCost(){' +
    '  var q=parseFloat(document.getElementById("qty").value)||0;' +
    '  var p=parseFloat(document.getElementById("buypx").value)||0;' +
    '  var el=document.getElementById("cprev");' +
    '  if(q>0&&p>0){el.textContent="Total cost: $"+(q*p).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});el.style.display="block";}' +
    '  else el.style.display="none";' +
    '}' +
    'function showErr(msg){var e=document.getElementById("err");e.textContent="\u26A0\uFE0F "+msg;e.style.display="block";}' +
    'function save(){' +
    '  if(!selSym){alert("Search for a token first.");return;}' +
    '  var qty=parseFloat(document.getElementById("qty").value);' +
    '  var buy=parseFloat(document.getElementById("buypx").value);' +
    '  var name=document.getElementById("cname").value.trim()||selSym;' +
    '  if(!qty||qty<=0){alert("Enter a valid quantity.");return;}' +
    '  if(!buy||buy<=0){alert("Enter a valid buy price.");return;}' +
    '  var acct=document.getElementById("acct").value;' +
    '  var notes=document.getElementById("notes").value.trim();' +
    '  document.getElementById("saveBtn").disabled=true;' +
    '  document.getElementById("saveBtn").textContent="Saving...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close();})' +
    '    .withFailureHandler(function(e){document.getElementById("saveBtn").disabled=false;document.getElementById("saveBtn").textContent="Add Holding";showErr(e.message);})' +
    '    .saveCrypto(selSym,name,qty,buy,acct,notes);' +
    '}' +
    'document.getElementById("sym").addEventListener("keydown",function(e){if(e.key==="Enter")lookup();});' +
    '</script></body></html>'
  ).setWidth(500).setHeight(580).setTitle('Add Crypto');

  SpreadsheetApp.getUi().showModalDialog(html, 'Add Crypto');
}


// ── SAVE CRYPTO ───────────────────────────────────────────────
function saveCrypto(symbol, name, qty, buyPrice, linkedAccount, notes) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CRYPTO_SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(CRYPTO_SHEET_NAME, 1); buildCryptoSheet_(sheet); }

  symbol = symbol.toString().trim().toUpperCase();
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = getLastCryptoRow_(sheet);
  const row     = lastRow + 1;


  const mktValFormula  = '=C' + row + '*E' + row;
  const gainFormula    = '=F' + row + '-C' + row + '*D' + row;
  const gainPctFormula = '=IF(D' + row + '*C' + row + '<>0,(E' + row + '-D' + row + ')/D' + row + '*100,0)';

  const usdFmt4 = '"$"#,##0.00######';
  const usdFmt2 = '"$"#,##0.00';
  const bg      = row % 2 === 0 ? '#f8f9fa' : '#ffffff';

  sheet.setRowHeight(row, 36);
  sheet.getRange(row, 1, 1, 12).setBackground(bg);
  sheet.getRange(row, CC.SYMBOL)  .setValue(symbol)   .setFontWeight('bold').setFontColor('#6d1b7b').setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(row, CC.NAME)    .setValue(name)      .setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(row, CC.QTY)     .setValue(qty)       .setNumberFormat('#,##0.########').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.BUY)     .setValue(buyPrice)  .setNumberFormat(usdFmt4).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.PRICE)   .setFormula('=IFERROR(IMPORTDATA(CONCATENATE("https://cryptoprices.cc/",A' + row + ')),0)').setNumberFormat(usdFmt4).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.MKT_VAL) .setFormula(mktValFormula) .setNumberFormat(usdFmt2).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.GAIN)    .setFormula(gainFormula)   .setNumberFormat(usdFmt2).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.GAIN_PCT).setFormula(gainPctFormula).setNumberFormat('0.00"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.CHG24H)     .setValue('')           .setNumberFormat('0.00"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.MKT_VAL_MYR).setFormula('=IFERROR(C' + row + '*E' + row + '*GOOGLEFINANCE("CURRENCY:USDMYR"),0)').setNumberFormat('"RM "#,##0.00').setFontWeight('bold').setFontColor('#1a237e').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, CC.ACCOUNT)    .setValue(linkedAccount || '').setFontSize(10).setFontColor('#5f6368').setVerticalAlignment('middle');
  sheet.getRange(row, CC.UPDATED) .setValue(today)     .setFontSize(9).setFontColor('#9aa0a6').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, CC.NOTES)   .setValue(notes || '').setFontSize(10).setFontColor('#9aa0a6').setVerticalAlignment('middle');

  SpreadsheetApp.flush();
  const gainVal = sheet.getRange(row, CC.GAIN).getValue();
  const glColor = (typeof gainVal === 'number' && gainVal < 0) ? '#d93025' : '#0f9d58';
  sheet.getRange(row, CC.GAIN).setFontColor(glColor);
  sheet.getRange(row, CC.GAIN_PCT).setFontColor(glColor);

  if (linkedAccount) {
    const accSheet = ss.getSheetByName(linkedAccount);
    if (!accSheet) throw new Error('Account "' + linkedAccount + '" not found.');
    const accCcy = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCcy !== 'USD' && accCcy !== 'MYR') throw new Error('Crypto purchases can only be linked to USD or MYR accounts. "' + linkedAccount + '" is ' + accCcy + '.');

    let costInAccCcy = qty * buyPrice;    // default: USD cost
    let fxNote = '';
    if (accCcy === 'MYR') {
      const rate = fetchUsdMyrRate_();
      if (!rate) throw new Error('Could not fetch USD/MYR exchange rate. Try again or link a USD account.');
      costInAccCcy = qty * buyPrice * rate;
      fxNote = ' [FX: 1 USD = RM ' + rate.toFixed(4) + ']';
    }
    const fmtMap = { USD: '"$"#,##0.00', MYR: '"RM "#,##0.00' };
    const fmt    = fmtMap[accCcy] || '"$"#,##0.00';
    const lastAcc = accSheet.getLastRow() + 1;
    const balFml  = lastAcc <= 2 ? '=D' + lastAcc : '=F' + (lastAcc-1) + '-D' + lastAcc;
    accSheet.getRange(lastAcc,1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc,2).setValue('Crypto');
    accSheet.getRange(lastAcc,3).setValue('Buy ' + symbol + ' x' + qty + ' @ $' + buyPrice + fxNote);
    accSheet.getRange(lastAcc,4).setValue(costInAccCcy).setNumberFormat(fmt);
    accSheet.getRange(lastAcc,5).setValue('OUT').setFontColor('#d93025').setFontWeight('bold');
    accSheet.getRange(lastAcc,6).setFormula(balFml).setNumberFormat(fmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc,1,1,6).setBackground('#f8f9fa');
  }

  refreshCryptoSummary_(sheet, null);
  sheet.activate();
}


// ── SELL CRYPTO DIALOG ────────────────────────────────────────
function showSellCryptoDialog() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CRYPTO_SHEET_NAME);
  if (!sheet || getLastCryptoRow_(sheet) < 4) {
    ui.alert('No crypto holdings found. Add a holding first via "Add Crypto".');
    return;
  }

  const lastRow  = getLastCryptoRow_(sheet);
  const raw      = sheet.getRange(4, 1, lastRow-3, 13).getValues();
  const holdings = raw.filter(r => r[CC.SYMBOL-1] !== '').map(r => ({
    symbol:   r[CC.SYMBOL-1],
    name:     r[CC.NAME-1],
    qty:      r[CC.QTY-1],
    buyPrice: r[CC.BUY-1],
    curPrice: typeof r[CC.PRICE-1] === 'number' ? r[CC.PRICE-1] : 0,
    account:  r[CC.ACCOUNT-1],
  }));

  const allSellAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }))
    .filter(a => a.currency === 'USD' || a.currency === 'MYR');

  const holdingsJson = JSON.stringify(holdings);
  const acctOptions  = allSellAccounts.map(a => '<option value="' + a.name + '">' + a.name + ' (' + a.currency + ')</option>').join('');

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 12px;font-size:15px;color:#c62828}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.info{background:#fff8e1;border:1px solid #fbc02d;border-radius:6px;padding:8px 10px;font-size:12px;margin-top:4px;display:none}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.prev{background:#e8f5e9;border-radius:6px;padding:8px 10px;margin-top:6px;font-size:12px;color:#2e7d32;display:none}' +
    '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px}' +
    '.br{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.bsell{background:#c62828;color:#fff}.bsell:hover{background:#b71c1c}' +
    '.bca{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>💸 Sell Crypto</h2>' +
    '<label>Select Holding</label>' +
    '<select id="sel" onchange="onChange()">' +
    '<option value="">— Select a token —</option>' +
    holdings.map((h,i) => '<option value="'+i+'">'+h.symbol+' — '+h.name+' ('+h.qty+' units)</option>').join('') +
    '</select>' +
    '<div class="info" id="info"></div>' +
    '<div id="form" style="display:none">' +
    '<div class="row2" style="margin-top:10px">' +
    '<div><label>Quantity to Sell</label><input type="number" id="sq" step="any" min="0" oninput="calc()"/></div>' +
    '<div><label>Sell Price (USD)</label><input type="number" id="sp" step="any" min="0" oninput="calc()"/></div>' +
    '</div>' +
    '<div class="prev" id="prev"></div>' +
    '<label>Return Proceeds to Account <span style="font-weight:400;color:#5f6368">(USD or RM — FX applied automatically)</span></label>' +
    '<select id="ret"><option value="">— None —</option>' + acctOptions + '</select>' +
    (allSellAccounts.length === 0 ? '<div class="warn">No USD or MYR accounts found.</div>' : '') +
    '<label>Notes (optional)</label>' +
    '<input type="text" id="notes" placeholder="e.g. Profit taking" />' +
    '</div>' +
    '<div class="br">' +
    '<button class="btn bca" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn bsell" id="sb" onclick="save()" style="display:none">Confirm Sell</button>' +
    '</div>' +
    '<script>' +
    'var H='+holdingsJson+';var idx=-1;' +
    'function onChange(){' +
    '  idx=parseInt(document.getElementById("sel").value);' +
    '  if(isNaN(idx)){document.getElementById("form").style.display="none";document.getElementById("info").style.display="none";return;}' +
    '  var h=H[idx];' +
    '  var info=document.getElementById("info");' +
    '  info.innerHTML="<b>"+h.symbol+"</b>  Held: <b>"+h.qty+"</b>  Buy: <b>$"+h.buyPrice.toFixed(6)+"</b>  Current: <b>$"+h.curPrice.toFixed(6)+"</b>";' +
    '  info.style.display="block";' +
    '  document.getElementById("sq").placeholder="Max: "+h.qty;' +
    '  document.getElementById("sp").value=h.curPrice>0?h.curPrice:"";' +
    '  var rs=document.getElementById("ret");' +
    '  if(h.account){for(var i=0;i<rs.options.length;i++){if(rs.options[i].value===h.account){rs.selectedIndex=i;break;}}}' +
    '  document.getElementById("form").style.display="block";' +
    '  document.getElementById("sb").style.display="inline-block";' +
    '  calc();' +
    '}' +
    'function calc(){' +
    '  if(idx<0)return;var h=H[idx];' +
    '  var q=parseFloat(document.getElementById("sq").value)||0;' +
    '  var p=parseFloat(document.getElementById("sp").value)||0;' +
    '  var el=document.getElementById("prev");' +
    '  if(q>0&&p>0){' +
    '    var pr=q*p;var g=(p-h.buyPrice)*q;var gs=g>=0?"+":"";' +
    '    el.innerHTML="Proceeds: <b>$"+pr.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})+"</b>  GL vs buy: <b>"+gs+"$"+g.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2})+"</b>";' +
    '    el.style.display="block";' +
    '  }else el.style.display="none";' +
    '}' +
    'function save(){' +
    '  if(idx<0){alert("Select a holding.");return;}' +
    '  var h=H[idx];' +
    '  var qty=parseFloat(document.getElementById("sq").value);' +
    '  var px=parseFloat(document.getElementById("sp").value);' +
    '  if(!qty||qty<=0){alert("Enter a valid quantity.");return;}' +
    '  if(qty>h.qty+1e-10){alert("Cannot sell more than held ("+h.qty+").");return;}' +
    '  if(!px||px<=0){alert("Enter a valid sell price.");return;}' +
    '  var ret=document.getElementById("ret").value;' +
    '  var notes=document.getElementById("notes").value.trim();' +
    '  document.getElementById("sb").disabled=true;document.getElementById("sb").textContent="Processing...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close();})' +
    '    .withFailureHandler(function(e){document.getElementById("sb").disabled=false;document.getElementById("sb").textContent="Confirm Sell";alert("Error: "+e.message);})' +
    '    .saveSellCrypto(idx,qty,px,ret,notes);' +
    '}' +
    '</script></body></html>'
  ).setWidth(480).setHeight(490).setTitle('Sell Crypto');

  SpreadsheetApp.getUi().showModalDialog(html, 'Sell Crypto');
}


// ── SAVE SELL CRYPTO ──────────────────────────────────────────
function saveSellCrypto(holdingIdx, sellQty, sellPrice, returnAccount, notes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CRYPTO_SHEET_NAME);
  if (!sheet) throw new Error('Crypto Portfolio sheet not found.');

  const lastRow  = getLastCryptoRow_(sheet);
  const allRows  = sheet.getRange(4, 1, lastRow-3, 13).getValues();
  const nonEmpty = [];
  allRows.forEach((r, i) => { if (r[CC.SYMBOL-1] !== '') nonEmpty.push({ r, rowNum: i+4 }); });
  if (holdingIdx >= nonEmpty.length) throw new Error('Holding index out of range.');

  const { r, rowNum } = nonEmpty[holdingIdx];
  const symbol   = r[CC.SYMBOL-1];
  const heldQty  = r[CC.QTY-1];
  const today    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  if (sellQty > heldQty + 1e-10) throw new Error('Cannot sell more than held.');
  const remaining = heldQty - sellQty;

  if (remaining < 1e-10) {
    sheet.deleteRow(rowNum);
  } else {
    sheet.getRange(rowNum, CC.QTY).setValue(remaining);
    sheet.getRange(rowNum, CC.UPDATED).setValue(today);
    const en = sheet.getRange(rowNum, CC.NOTES).getValue();
    const sn = 'Sold ' + sellQty + ' @ $' + sellPrice + ' on ' + today.split(' ')[0];
    sheet.getRange(rowNum, CC.NOTES).setValue(en ? en + ' | ' + sn : sn);
  }

  if (returnAccount) {
    const accSheet = ss.getSheetByName(returnAccount);
    if (!accSheet) throw new Error('Account "' + returnAccount + '" not found.');
    const accCcy = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCcy !== 'USD' && accCcy !== 'MYR') throw new Error('Crypto proceeds can only be credited to a USD or MYR account. "' + returnAccount + '" is ' + accCcy + '.');

    let proceedsInAccCcy = sellQty * sellPrice;   // USD by default
    let fxNote = '';
    if (accCcy === 'MYR') {
      const rate = fetchUsdMyrRate_();
      if (!rate) throw new Error('Could not fetch USD/MYR exchange rate. Try again or select a USD account.');
      proceedsInAccCcy = sellQty * sellPrice * rate;
      fxNote = ' [FX: 1 USD = RM ' + rate.toFixed(4) + ']';
    }
    const fmtMap = { USD: '"$"#,##0.00', MYR: '"RM "#,##0.00' };
    const fmt    = fmtMap[accCcy] || '"$"#,##0.00';
    const lastAcc = accSheet.getLastRow() + 1;
    const balFml  = lastAcc <= 2 ? '=D' + lastAcc : '=F' + (lastAcc-1) + '+D' + lastAcc;
    accSheet.getRange(lastAcc,1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc,2).setValue('Crypto');
    accSheet.getRange(lastAcc,3).setValue('Sell ' + symbol + ' x' + sellQty + ' @ $' + sellPrice + (notes ? ' ('+notes+')' : '') + fxNote);
    accSheet.getRange(lastAcc,4).setValue(proceedsInAccCcy).setNumberFormat(fmt);
    accSheet.getRange(lastAcc,5).setValue('IN').setFontColor('#0f9d58').setFontWeight('bold');
    accSheet.getRange(lastAcc,6).setFormula(balFml).setNumberFormat(fmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc,1,1,6).setBackground('#f8f9fa');
  }

  refreshCryptoSummary_(sheet, null);
}


// ── REFRESH CRYPTO PRICES ─────────────────────────────────────
/**
 * 1. Batch-fetches all crypto prices via UrlFetchApp.fetchAll() (parallel).
 * 2. Fetches live USD→MYR rate via GOOGLEFINANCE (_FX_TEMP_ sheet).
 * 3. Writes USD price to col E and RM market value to col J as plain values.
 */
function refreshCryptoPrices_(sheet) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = getLastCryptoRow_(sheet);
  if (lastRow < 4) return;

  // Ensure all rows use IMPORTDATA formula for live price (migrates old setValue rows)
  for (let row = 4; row <= lastRow; row++) {
    const sym = sheet.getRange(row, CC.SYMBOL).getValue();
    if (!sym) continue;
    const priceCell = sheet.getRange(row, CC.PRICE);
    if (!priceCell.getFormula()) {
      priceCell.setFormula('=IFERROR(IMPORTDATA(CONCATENATE("https://cryptoprices.cc/",A' + row + ')),0)');
    }
  }
  SpreadsheetApp.flush();
  Utilities.sleep(2000); // allow IMPORTDATA to populate

  // Gather symbols and qtys
  const rowMeta = [];
  for (let row = 4; row <= lastRow; row++) {
    const sym = sheet.getRange(row, CC.SYMBOL).getValue();
    const qty = sheet.getRange(row, CC.QTY).getValue();
    if (sym) rowMeta.push({ row, symbol: sym.toString().toUpperCase(), qty: qty || 0 });
  }
  if (!rowMeta.length) return;

  // ── Batch fetch USD prices ───────────────────────────────────
  const priceMap = fetchCryptoBatch_(rowMeta.map(m => m.symbol));

  // ── Write prices — MKT_VAL_MYR uses live GOOGLEFINANCE formula, no FX fetch needed ──
  rowMeta.forEach(m => {
    const price = priceMap[m.symbol];
    if (price !== undefined) {
      // Re-set IMPORTDATA formula to force a fresh fetch
      sheet.getRange(m.row, CC.PRICE)
        .setFormula('=IFERROR(IMPORTDATA(CONCATENATE("https://cryptoprices.cc/",A' + m.row + ')),0)');
      // Ensure live GOOGLEFINANCE formula for RM value
      sheet.getRange(m.row, CC.MKT_VAL_MYR)
        .setFormula('=IFERROR(C' + m.row + '*E' + m.row + '*GOOGLEFINANCE("CURRENCY:USDMYR"),0)')
        .setNumberFormat('"RM "#,##0.00').setFontWeight('bold')
        .setFontColor('#1a237e').setHorizontalAlignment('right');
      sheet.getRange(m.row, CC.UPDATED).setValue(today);
    }
  });

  // Update subtitle row
  sheet.getRange(2, 1, 1, 13).merge()
    .setValue('Live prices via cryptoprices.cc  ·  RM values via GOOGLEFINANCE(USDMYR)  ·  Updated: ' + today)
    .setBackground('#f3e5f5').setFontColor('#6d1b7b')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  SpreadsheetApp.flush();

  // Recolour gain/loss
  for (let row = 4; row <= lastRow; row++) {
    const gainVal = sheet.getRange(row, CC.GAIN).getValue();
    if (typeof gainVal === 'number') {
      const c = gainVal < 0 ? '#d93025' : '#0f9d58';
      sheet.getRange(row, CC.GAIN).setFontColor(c);
      sheet.getRange(row, CC.GAIN_PCT).setFontColor(c);
    }
  }

  refreshCryptoSummary_(sheet, usdMyr);
}


// ── CRYPTO SUMMARY ────────────────────────────────────────────
function refreshCryptoSummary_(sheet, usdMyr) {
  const lastRow = getLastCryptoRow_(sheet);
  const maxRows = sheet.getMaxRows();
  if (lastRow + 1 <= maxRows) {
    sheet.getRange(lastRow+1, 1, maxRows-lastRow, 12).clearContent().clearFormat();
  }
  if (lastRow < 4) return;

  const data = sheet.getRange(4, 1, lastRow-3, 12).getValues()
    .filter(r => r[CC.SYMBOL-1] !== '');

  const totalCost  = data.reduce((s,r) => s + (r[CC.QTY-1] * r[CC.BUY-1]), 0);
  const totalValue = data.reduce((s,r) => s + (typeof r[CC.MKT_VAL-1] === 'number' ? r[CC.MKT_VAL-1] : 0), 0);
  const totalGain  = totalValue - totalCost;
  const gainPct    = totalCost > 0 ? (totalGain / totalCost * 100) : 0;
  const glColor    = totalGain >= 0 ? '#0f9d58' : '#d93025';
  const usdFmt     = '"$"#,##0.00';

  const sr = lastRow + 2;
  sheet.setRowHeight(sr, 30);
  sheet.getRange(sr, 1, 1, 12).merge()
    .setValue('📊  PORTFOLIO SUMMARY  (USD)')
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const rmFmt  = '"RM "#,##0.00';
  const rmRate = (usdMyr && typeof usdMyr === 'number' && usdMyr > 0) ? usdMyr : null;
  const rmLabel = rmRate ? ' (RM ' + rmRate.toFixed(4) + '/USD)' : '';

  [
    ['Total Holdings',          data.length + ' tokens',                 null,    '#202124'],
    ['Total Cost (USD)',        totalCost,                               usdFmt,  '#202124'],
    ['Current Value (USD)',     totalValue,                              usdFmt,  '#202124'],
    ['Total Gain / Loss',       totalGain,                               usdFmt,  glColor ],
    ['Total Gain / Loss (%)',   gainPct,                                 '0.00"%"',glColor],
    ['Current Value (RM)' + rmLabel, rmRate ? totalValue * rmRate : 'N/A', rmFmt, '#1a237e'],
    ['Total Gain / Loss (RM)',  rmRate ? totalGain * rmRate : 'N/A',    rmFmt,   glColor ],
  ].forEach(([label, val, fmt, color], i) => {
    const r = sr + 1 + i;
    sheet.setRowHeight(r, 28);
    sheet.getRange(r,1,1,3).merge()
      .setValue(label).setFontWeight('bold').setFontSize(10)
      .setBackground(i%2===0?'#f3e5f5':'#fce4ec').setVerticalAlignment('middle');
    const vc = sheet.getRange(r, 4);
    vc.setValue(val).setFontColor(color).setFontWeight('bold').setVerticalAlignment('middle');
    if (fmt) vc.setNumberFormat(fmt);
  });
}

// ============================================================
//  GOLD PORTFOLIO TRACKER
//  Sheet: 🥇 Gold Portfolio
//  Price source: GOOGLEFINANCE("CURRENCY:XAUMYR") = XAU price per troy oz in MYR
//  Supports: 916 (91.6% purity) and 999 (99.9% purity) gold
//  All values in MYR. Buy/sell linked to MYR accounts.
//  1 troy oz = 31.1035 grams (fixed constant)
// ============================================================

const GOLD_SHEET_NAME  = '🥇 Gold Portfolio';
const TROY_OZ_TO_GRAM  = 31.1035;

// Column indices (1-based)
const GC = {
  TYPE:      1,   // A – 916 or 999
  DESC:      2,   // B – description e.g. "Public Gold bar 100g"
  WEIGHT:    3,   // C – weight in grams
  BUY_PRICE: 4,   // D – buy price per gram (RM) at purchase
  BUY_TOTAL: 5,   // E – total cost (RM) = C * D
  CUR_PRICE: 6,   // F – current price per gram (RM) — written on refresh
  CUR_VALUE: 7,   // G – current value (RM) = C * F  (formula)
  GAIN:      8,   // H – gain/loss (RM) = G - E  (formula)
  GAIN_PCT:  9,   // I – gain % (formula)
  ACCOUNT:   10,  // J – linked MYR account
  UPDATED:   11,  // K – last price update
  NOTES:     12,  // L – free text
};

// ── FETCH GOLD PRICE PER GRAM IN MYR ─────────────────────────
/**
 * Fetches XAU/MYR from GOOGLEFINANCE via a temp sheet.
 * Returns { price916, price999, xauMyr } where prices are per gram in RM.
 * price916 = (xauMyr / 31.1035) * 0.916
 * price999 = (xauMyr / 31.1035) * 0.999
 */
/**
 * Fetch live MYR gold prices per gram by scraping livepriceofgold.com.
 * Returns { price916, price999, source } or null on failure.
 * Falls back to GOOGLEFINANCE (XAU/USD × USD/MYR) if scraping fails.
 *
 * Results cached in CacheService for 3 minutes to avoid hammering the site.
 */
function fetchGoldPrices_() {
  const CACHE_KEY = 'gold_prices_myr_v1';
  const cache = CacheService.getScriptCache();

  // ── 1. Try cache first ───────────────────────────────────────
  try {
    const cached = cache.get(CACHE_KEY);
    if (cached) {
      const d = JSON.parse(cached);
      if (d && d.price916 > 0 && d.price999 > 0) return d;
    }
  } catch(_) {}

  // ── 2. Scrape livepriceofgold.com ────────────────────────────
  try {
    const url = 'https://www.livepriceofgold.com/malaysia-gold-price.html';
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() === 200) {
      const html = res.getContentText();

      // Each gold row looks like:
      //   <tr>...<td>24K Gold/gram ... </td><td>MID</td><td>BUY</td><td>SELL</td></tr>
      // td:nth-child(3) = mid price (index 2), td:nth-child(4) = buy (index 3)
      // We use the mid price (td index 2) for a neutral market reference.

      const parseRow = (label) => {
        // Find the row containing the label
        const idx = html.indexOf(label);
        if (idx < 0) return null;
        // Walk back to find opening <tr>
        const trStart = html.lastIndexOf('<tr', idx);
        if (trStart < 0) return null;
        // Walk forward to find closing </tr>
        const trEnd = html.indexOf('</tr>', idx);
        if (trEnd < 0) return null;
        const row = html.substring(trStart, trEnd);
        // Extract all <td>...</td> text values
        const tdRx = /<td[^>]*>([\s\S]*?)<\/td>/gi;
        const vals = [];
        let m;
        while ((m = tdRx.exec(row)) !== null) {
          // Strip inner tags and whitespace
          const txt = m[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
          vals.push(txt);
        }
        // vals[0]=empty, vals[1]=label+change, vals[2]=mid, vals[3]=buy, vals[4]=sell
        // Use mid price (vals[2]) — most neutral, closest to spot
        const price = parseFloat((vals[2] || '').replace(/,/g, ''));
        return isNaN(price) || price <= 0 ? null : price;
      };

      const price999 = parseRow('24K Gold/gram');  // 999 = 24K
      const price916 = parseRow('22K Gold/gram');  // 916 = 22K

      if (price999 && price916) {
        const result = {
          price916: price916,
          price999: price999,
          source: 'livepriceofgold.com'
        };
        try { cache.put(CACHE_KEY, JSON.stringify(result), 180); } catch(_) {}
        return result;
      }
    }
  } catch(scrapeErr) {
    Logger.log('Gold scrape failed: ' + scrapeErr.message);
  }

  // ── 3. Fallback: GOOGLEFINANCE XAU/USD × USD/MYR ─────────────
  Logger.log('Gold scrape failed, falling back to GOOGLEFINANCE');
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = '_FX_TEMP_';
  try {
    let temp = ss.getSheetByName(tempName);
    if (temp) ss.deleteSheet(temp);
    temp = ss.insertSheet(tempName);
    compactSheet_(temp, 10, 1); // 1 col × 10 rows
    temp.getRange('A1').setFormula('=GOOGLEFINANCE("CURRENCY:XAUUSD")');
    temp.getRange('A2').setFormula('=GOOGLEFINANCE("CURRENCY:USDMYR")');
    SpreadsheetApp.flush();
    Utilities.sleep(4000);
    const xauUsd = temp.getRange('A1').getValue();
    const usdMyr = temp.getRange('A2').getValue();
    ss.deleteSheet(temp);
    if (typeof xauUsd !== 'number' || xauUsd <= 0) return null;
    if (typeof usdMyr !== 'number' || usdMyr <= 0) return null;
    const perGram = (xauUsd * usdMyr) / TROY_OZ_TO_GRAM;
    const result = {
      price916: perGram * (0.916 / 0.999),
      price999: perGram,
      xauMyr:   xauUsd * usdMyr,
      source: 'GOOGLEFINANCE (fallback)'
    };
    try { cache.put(CACHE_KEY, JSON.stringify(result), 180); } catch(_) {}
    return result;
  } catch(e) {
    try { const t = ss.getSheetByName(tempName); if (t) ss.deleteSheet(t); } catch(_) {}
    return null;
  }
}

/** Called from Add Gold dialog — returns JSON string */
function getGoldPricesForDialog() {
  const d = fetchGoldPrices_();
  if (!d) return JSON.stringify({ error: true, message: 'Could not fetch gold price. Please check your internet connection and try again.' });
  return JSON.stringify(d);
}


// ── CREATE GOLD SHEET ─────────────────────────────────────────
function createGoldSheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GOLD_SHEET_NAME);
  if (sheet) {
    ui.alert('The Gold Portfolio sheet already exists.\n\nClick the "🥇 Gold Portfolio" tab to view it.');
    sheet.activate();
    return;
  }
  sheet = ss.insertSheet(GOLD_SHEET_NAME, 1);
  buildGoldSheet_(sheet);
  sheet.activate();
  ui.alert(
    '🥇 Gold Portfolio sheet created!\n\n' +
    'Use "Add Gold" from the menu to record a purchase.\n\n' +
    'Supports 916 (91.6%) and 999 (99.9%) gold.\n' +
    'Prices via livepriceofgold.com — all values in Ringgit.\n\n' +
    'Tip: Use "Refresh All Prices Now" to update gold prices.'
  );
}

function buildGoldSheet_(sheet) {
  const widths = [60, 220, 100, 130, 130, 130, 130, 120, 90, 160, 130, 180];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Row 1: Banner
  sheet.setRowHeight(1, 52);
  sheet.getRange(1, 1, 1, 12).merge()
    .setValue('🥇  GOLD PORTFOLIO')
    .setBackground('#f57f17').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Row 2: Subtitle
  sheet.setRowHeight(2, 26);
  sheet.getRange(2, 1, 1, 12).merge()
    .setValue('Prices via livepriceofgold.com  ·  All values in RM  ·  Use "Refresh All Prices Now" to update')
    .setBackground('#fff8e1').setFontColor('#f57f17')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Row 3: Headers
  sheet.setRowHeight(3, 38);
  const headers = [
    'Type', 'Description', 'Weight\n(grams)',
    'Buy Price\n(RM/gram)', 'Total Cost\n(RM)',
    'Current\nPrice (RM/g)', 'Current\nValue (RM)',
    'Gain/Loss\n(RM)', 'Gain/Loss\n%',
    'Linked Account', 'Last Updated', 'Notes'
  ];
  sheet.getRange(3, 1, 1, 12).setValues([headers])
    .setBackground('#e65100').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);

  // Pre-format data range
  sheet.getRange('D4:H53').setNumberFormat('"RM "#,##0.00');
  sheet.getRange('I4:I53').setNumberFormat('0.00"%"');
  sheet.getRange('C4:C53').setNumberFormat('#,##0.###');
  sheet.setFrozenRows(3);
  sheet.setHiddenGridlines(true);
}

function getLastGoldRow_(sheet) {
  // Valid gold row: TYPE in A AND DESC in B AND numeric WEIGHT in C
  const maxRows = sheet.getMaxRows();
  if (maxRows <= 3) return 3; // only header rows, no data
  const data = sheet.getRange(4, GC.TYPE, maxRows - 3, 3).getValues();
  let last = 3;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== '' && data[i][1] !== '' && typeof data[i][2] === 'number' && data[i][2] > 0) {
      last = i + 4;
    }
  }
  return last;
}


// ── ADD GOLD DIALOG ───────────────────────────────────────────
function showAddGoldDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss.getSheetByName(GOLD_SHEET_NAME)) {
    const resp = ui.alert('No Gold Portfolio sheet found.', 'Create it now?', ui.ButtonSet.YES_NO);
    if (resp === ui.Button.YES) createGoldSheet(); else return;
  }

  const myrAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }))
    .filter(a => a.currency === 'MYR');

  const acctOptions = ['<option value="">— None (no account deduction) —</option>']
    .concat(myrAccounts.map(a => '<option value="' + a.name + '">' + a.name + '</option>'))
    .join('');

  const noAcctWarn = myrAccounts.length === 0
    ? '<div class="warn">No MYR accounts found. Create a MYR account to link purchases.</div>'
    : '';

  // All JS event handlers attached in <script> block — zero inline onclick quote nesting
  const htmlStr = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:0;background:#f8f9fa;color:#202124;font-size:13px}'
    + '.hdr{background:#e65100;color:#fff;padding:14px 18px;font-size:15px;font-weight:700}'
    + '.body{padding:16px}'
    + 'label{display:block;font-weight:600;margin:12px 0 4px;color:#3c4043;font-size:12px}'
    + 'input,select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}'
    + 'input:focus,select:focus{outline:none;border-color:#e65100}'
    + '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}'
    + '.pricebox{background:#fff8e1;border:1px solid #ffcc02;border-radius:8px;padding:10px 14px;margin-top:4px;display:none}'
    + '.priceval{font-size:15px;font-weight:700;color:#e65100}'
    + '.pricesub{font-size:11px;color:#5f6368;margin-top:2px}'
    + '.costprev{background:#e8f5e9;border-radius:6px;padding:7px 11px;margin-top:6px;font-size:12px;font-weight:600;color:#2e7d32;display:none}'
    + '.spin{color:#e65100;font-size:12px;font-weight:600;padding:8px 12px;background:#fff8e1;border-radius:6px;border:1px solid #ffcc02;margin-top:6px;display:none}'
    + '.err{background:#fce8e6;border:1px solid #f28b82;border-radius:6px;padding:8px 12px;font-size:12px;color:#c62828;margin-top:6px;display:none}'
    + '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 11px;font-size:12px;color:#856404;margin-top:4px}'
    + '.footer{display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;border-top:1px solid #e8eaed;background:#fff}'
    + '.btn{padding:8px 22px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}'
    + '.bsave{background:#e65100;color:#fff}.bsave:hover{background:#bf360c}'
    + '.bcancel{background:#f1f3f4;color:#3c4043}'
    + '</style></head><body>'
    + '<div class="hdr">🥇 Add Gold Transaction</div>'
    + '<div class="body">'
    + '<label>Gold Type</label>'
    + '<select id="goldType">'
    + '<option value="">— Select gold type —</option>'
    + '<option value="916">916 Gold (91.6% purity)</option>'
    + '<option value="999">999 Gold (99.9% purity)</option>'
    + '</select>'
    + '<div class="spin" id="spin">⏳ Fetching live gold price… (~5 sec)</div>'
    + '<div class="err" id="err"></div>'
    + '<div class="pricebox" id="pricebox">'
    + '<div class="priceval" id="priceval"></div>'
    + '<div class="pricesub" id="pricesub"></div>'
    + '</div>'
    + '<label>Description</label>'
    + '<input type="text" id="desc" placeholder="e.g. Public Gold bar 50g, gold ring" />'
    + '<div class="row2">'
    + '<div><label>Weight (grams)</label><input type="number" id="wt" placeholder="e.g. 10" step="any" min="0" /></div>'
    + '<div><label>Buy Price (RM / gram)</label><input type="number" id="buypx" placeholder="auto-filled" step="any" min="0" /></div>'
    + '</div>'
    + '<div class="costprev" id="costprev"></div>'
    + '<label>Deduct from Account <span style="font-weight:400;color:#5f6368">(MYR only, optional)</span></label>'
    + '<select id="acct">' + acctOptions + '</select>'
    + noAcctWarn
    + '<label>Notes <span style="font-weight:400;color:#5f6368">(optional)</span></label>'
    + '<input type="text" id="notes" placeholder="e.g. Kedai Emas Taka, Public Gold" />'
    + '</div>'
    + '<div class="footer">'
    + '<button class="btn bcancel" id="cancelBtn">Cancel</button>'
    + '<button class="btn bsave" id="saveBtn" disabled>Add Transaction</button>'
    + '</div>'
    + '<script>'
    + 'var goldPrices=null,loading=false;'
    + 'document.getElementById("cancelBtn").onclick=function(){google.script.host.close();};'
    + 'document.getElementById("saveBtn").onclick=save;'
    + 'document.getElementById("goldType").onchange=onTypeChange;'
    + 'document.getElementById("wt").oninput=calcCost;'
    + 'document.getElementById("buypx").oninput=function(){this.dataset.edited="1";calcCost();};'
    + 'function onTypeChange(){'
    + '  var t=document.getElementById("goldType").value;'
    + '  document.getElementById("err").style.display="none";'
    + '  if(!t){document.getElementById("pricebox").style.display="none";return;}'
    + '  if(goldPrices){updatePrice(t);return;}'
    + '  if(loading)return;'
    + '  loading=true;'
    + '  document.getElementById("spin").style.display="block";'
    + '  document.getElementById("pricebox").style.display="none";'
    + '  google.script.run'
    + '    .withSuccessHandler(onPrices)'
    + '    .withFailureHandler(onPriceFail)'
    + '    .getGoldPricesForDialog();'
    + '}'
    + 'function onPrices(json){'
    + '  loading=false;'
    + '  document.getElementById("spin").style.display="none";'
    + '  var d=JSON.parse(json);'
    + '  if(d.error){showErr(d.message);return;}'
    + '  goldPrices=d;'
    + '  updatePrice(document.getElementById("goldType").value);'
    + '}'
    + 'function onPriceFail(e){loading=false;document.getElementById("spin").style.display="none";showErr(e.message);}'
    + 'function updatePrice(t){'
    + '  if(!t||!goldPrices)return;'
    + '  var px=t==="916"?goldPrices.price916:goldPrices.price999;'
    + '  document.getElementById("priceval").textContent=t+" Gold — RM "+px.toFixed(2)+" / gram";'
    + '  var sub=goldPrices.xauMyr?"XAU/MYR = RM "+goldPrices.xauMyr.toLocaleString("en-MY",{minimumFractionDigits:2,maximumFractionDigits:2})+" / troy oz":"";if(goldPrices.source)sub+=(sub?" · ":"")+"Source: "+goldPrices.source;document.getElementById("pricesub").textContent=sub;'
    + '  document.getElementById("pricebox").style.display="block";'
    + '  if(!document.getElementById("buypx").dataset.edited){document.getElementById("buypx").value=px.toFixed(2);}'
    + '  document.getElementById("saveBtn").disabled=false;'
    + '  calcCost();'
    + '}'
    + 'function showErr(msg){'
    + '  var el=document.getElementById("err");'
    + '  el.textContent="\u26a0\ufe0f "+msg+" — you can still enter buy price manually.";'
    + '  el.style.display="block";'
    + '  document.getElementById("saveBtn").disabled=false;'
    + '}'
    + 'function calcCost(){'
    + '  var w=parseFloat(document.getElementById("wt").value)||0;'
    + '  var p=parseFloat(document.getElementById("buypx").value)||0;'
    + '  var el=document.getElementById("costprev");'
    + '  if(w>0&&p>0){el.textContent="Total cost: RM "+(w*p).toLocaleString("en-MY",{minimumFractionDigits:2,maximumFractionDigits:2});el.style.display="block";}else el.style.display="none";'
    + '}'
    + 'function save(){'
    + '  var t=document.getElementById("goldType").value;'
    + '  var desc=document.getElementById("desc").value.trim();'
    + '  var wt=parseFloat(document.getElementById("wt").value);'
    + '  var buy=parseFloat(document.getElementById("buypx").value);'
    + '  if(!t){alert("Please select a gold type.");return;}'
    + '  if(!desc){alert("Please enter a description.");return;}'
    + '  if(!wt||wt<=0){alert("Please enter a valid weight.");return;}'
    + '  if(!buy||buy<=0){alert("Please enter a buy price.");return;}'
    + '  var acct=document.getElementById("acct").value;'
    + '  var notes=document.getElementById("notes").value.trim();'
    + '  document.getElementById("saveBtn").disabled=true;'
    + '  document.getElementById("saveBtn").textContent="Saving...";'
    + '  document.getElementById("cancelBtn").disabled=true;'
    + '  google.script.run'
    + '    .withSuccessHandler(function(){google.script.host.close();})'
    + '    .withFailureHandler(function(e){'
    + '      document.getElementById("saveBtn").disabled=false;'
    + '      document.getElementById("saveBtn").textContent="Add Transaction";'
    + '      document.getElementById("cancelBtn").disabled=false;'
    + '      document.getElementById("err").textContent="\u26a0\ufe0f "+e.message;'
    + '      document.getElementById("err").style.display="block";'
    + '    })'
    + '    .saveGold(t,desc,wt,buy,acct,notes);'
    + '}'
    + '</script></body></html>';

  const html = HtmlService.createHtmlOutput(htmlStr).setWidth(460).setHeight(580).setTitle('Add Gold Transaction');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Gold Transaction');
}


// ── SAVE GOLD ─────────────────────────────────────────────────
function saveGold(goldType, desc, weight, buyPricePerGram, linkedAccount, notes) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(GOLD_SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(GOLD_SHEET_NAME, 1); buildGoldSheet_(sheet); }

  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = getLastGoldRow_(sheet);
  const row     = lastRow + 1;
  const totalCost = weight * buyPricePerGram;

  // Fetch live current price
  const prices    = fetchGoldPrices_();
  const curPrice  = prices ? (goldType === '916' ? prices.price916 : prices.price999) : buyPricePerGram;

  const rmFmt   = '"RM "#,##0.00';
  const rmFmt4  = '"RM "#,##0.0000';
  const bg      = row % 2 === 0 ? '#f8f9fa' : '#ffffff';

  const curValFormula  = '=C' + row + '*F' + row;
  const gainFormula    = '=G' + row + '-E' + row;
  const gainPctFormula = '=IF(E' + row + '<>0,(G' + row + '-E' + row + ')/E' + row + '*100,0)';

  // Type badge color
  const typeBg    = goldType === '916' ? '#ffd54f' : '#fff9c4';
  const typeColor = goldType === '916' ? '#e65100' : '#f57f17';

  sheet.setRowHeight(row, 36);
  sheet.getRange(row, 1, 1, 12).setBackground(bg);
  sheet.getRange(row, GC.TYPE)      .setValue(goldType).setFontWeight('bold').setFontColor(typeColor).setBackground(typeBg).setHorizontalAlignment('center').setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(row, GC.DESC)      .setValue(desc).setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(row, GC.WEIGHT)    .setValue(weight).setNumberFormat('#,##0.###').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.BUY_PRICE) .setValue(buyPricePerGram).setNumberFormat(rmFmt4).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.BUY_TOTAL) .setValue(totalCost).setNumberFormat(rmFmt).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.CUR_PRICE) .setValue(curPrice).setNumberFormat(rmFmt4).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.CUR_VALUE) .setFormula(curValFormula).setNumberFormat(rmFmt).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.GAIN)      .setFormula(gainFormula).setNumberFormat(rmFmt).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.GAIN_PCT)  .setFormula(gainPctFormula).setNumberFormat('0.00"%"').setHorizontalAlignment('right').setVerticalAlignment('middle');
  sheet.getRange(row, GC.ACCOUNT)   .setValue(linkedAccount || '').setFontSize(10).setFontColor('#5f6368').setVerticalAlignment('middle');
  sheet.getRange(row, GC.UPDATED)   .setValue(today).setFontSize(9).setFontColor('#9aa0a6').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange(row, GC.NOTES)     .setValue(notes || '').setFontSize(10).setFontColor('#9aa0a6').setVerticalAlignment('middle');

  SpreadsheetApp.flush();

  // Colour gain/loss
  const gainVal = sheet.getRange(row, GC.GAIN).getValue();
  const glColor = (typeof gainVal === 'number' && gainVal < 0) ? '#d93025' : '#0f9d58';
  sheet.getRange(row, GC.GAIN).setFontColor(glColor);
  sheet.getRange(row, GC.GAIN_PCT).setFontColor(glColor);

  // Deduct from linked MYR account
  if (linkedAccount) {
    const accSheet = ss.getSheetByName(linkedAccount);
    if (!accSheet) throw new Error('Account "' + linkedAccount + '" not found.');
    const accCcy = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCcy !== 'MYR') throw new Error('Gold purchases can only be linked to MYR accounts. "' + linkedAccount + '" is ' + accCcy + '.');
    const lastAcc = accSheet.getLastRow() + 1;
    const balFml  = lastAcc <= 2 ? '=D' + lastAcc : '=F' + (lastAcc - 1) + '-D' + lastAcc;
    accSheet.getRange(lastAcc, 1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc, 2).setValue('Gold');
    accSheet.getRange(lastAcc, 3).setValue('Buy ' + goldType + ' gold ' + weight + 'g @ RM ' + buyPricePerGram.toFixed(2) + '/g — ' + desc);
    accSheet.getRange(lastAcc, 4).setValue(totalCost).setNumberFormat(rmFmt);
    accSheet.getRange(lastAcc, 5).setValue('OUT').setFontColor('#d93025').setFontWeight('bold');
    accSheet.getRange(lastAcc, 6).setFormula(balFml).setNumberFormat(rmFmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc, 1, 1, 6).setBackground('#f8f9fa');
  }

  refreshGoldSummary_(sheet);
  sheet.activate();
}


// ── SELL GOLD DIALOG ──────────────────────────────────────────
function showSellGoldDialog() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GOLD_SHEET_NAME);
  if (!sheet || getLastGoldRow_(sheet) < 4) {
    ui.alert('No gold holdings found. Add a holding first via "Add Gold".');
    return;
  }

  const lastRow  = getLastGoldRow_(sheet);
  const raw      = sheet.getRange(4, 1, lastRow - 3, 12).getValues();
  const holdings = raw
    .filter(r => r[GC.TYPE-1] !== '' && r[GC.DESC-1] !== '')
    .map(r => ({
      type:      r[GC.TYPE-1],
      desc:      r[GC.DESC-1],
      weight:    r[GC.WEIGHT-1],
      buyPrice:  r[GC.BUY_PRICE-1],
      curPrice:  typeof r[GC.CUR_PRICE-1] === 'number' ? r[GC.CUR_PRICE-1] : 0,
      account:   r[GC.ACCOUNT-1],
    }));

  const myrAccounts = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .map(s => ({ name: s.getName(), currency: normCurrency_(s.getRange('G2').getValue() || 'MYR') }))
    .filter(a => a.currency === 'MYR');

  const holdingsJson = JSON.stringify(holdings);
  const acctOptions  = myrAccounts.map(a => '<option value="' + a.name + '">' + a.name + '</option>').join('');

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 12px;font-size:15px;color:#c62828}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.info{background:#fff8e1;border:1px solid #ffcc02;border-radius:6px;padding:8px 10px;font-size:12px;margin-top:4px;display:none}' +
    '.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px}' +
    '.prev{background:#e8f5e9;border-radius:6px;padding:8px 10px;margin-top:6px;font-size:12px;color:#2e7d32;display:none}' +
    '.warn{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:7px 10px;font-size:12px;color:#856404;margin-top:4px}' +
    '.br{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.bsell{background:#c62828;color:#fff}.bsell:hover{background:#b71c1c}' +
    '.bca{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>💰 Sell Gold</h2>' +
    '<label>Select Holding</label>' +
    '<select id="sel" onchange="onChange()">' +
    '<option value="">— Select a holding —</option>' +
    holdings.map((h, i) => '<option value="' + i + '">[' + h.type + '] ' + h.desc + ' (' + h.weight + 'g)</option>').join('') +
    '</select>' +
    '<div class="info" id="info"></div>' +
    '<div id="form" style="display:none">' +
    '<div class="row2" style="margin-top:10px">' +
    '<div><label>Weight to Sell (grams)</label><input type="number" id="sw" step="any" min="0" oninput="calc()"/></div>' +
    '<div><label>Sell Price (RM/gram)</label><input type="number" id="sp" step="any" min="0" oninput="calc()"/></div>' +
    '</div>' +
    '<div class="prev" id="prev"></div>' +
    '<label>Deposit Proceeds to Account <span style="font-weight:400">(MYR accounts only)</span></label>' +
    '<select id="ret"><option value="">— None —</option>' + acctOptions + '</select>' +
    (myrAccounts.length === 0 ? '<div class="warn">No MYR accounts found.</div>' : '') +
    '<label>Notes (optional)</label>' +
    '<input type="text" id="notes" placeholder="e.g. Sold at bank, buyback" />' +
    '</div>' +
    '<div class="br">' +
    '<button class="btn bca" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn bsell" id="sb" onclick="save()" style="display:none">Confirm Sell</button>' +
    '</div>' +
    '<script>' +
    'var H=' + holdingsJson + ';var idx=-1;' +
    'function onChange(){' +
    '  idx=parseInt(document.getElementById("sel").value);' +
    '  if(isNaN(idx)){document.getElementById("form").style.display="none";document.getElementById("info").style.display="none";return;}' +
    '  var h=H[idx];' +
    '  var info=document.getElementById("info");' +
    '  info.innerHTML="<b>["+h.type+"] "+h.desc+"</b>  Weight: <b>"+h.weight+"g</b>  Buy: <b>RM "+h.buyPrice.toFixed(2)+"/g</b>  Current: <b>RM "+h.curPrice.toFixed(2)+"/g</b>";' +
    '  info.style.display="block";' +
    '  document.getElementById("sw").placeholder="Max: "+h.weight+"g";' +
    '  document.getElementById("sp").value=h.curPrice>0?h.curPrice.toFixed(2):"";' +
    '  var rs=document.getElementById("ret");' +
    '  if(h.account){for(var i=0;i<rs.options.length;i++){if(rs.options[i].value===h.account){rs.selectedIndex=i;break;}}}' +
    '  document.getElementById("form").style.display="block";' +
    '  document.getElementById("sb").style.display="inline-block";' +
    '  calc();' +
    '}' +
    'function calc(){' +
    '  if(idx<0)return;var h=H[idx];' +
    '  var w=parseFloat(document.getElementById("sw").value)||0;' +
    '  var p=parseFloat(document.getElementById("sp").value)||0;' +
    '  var el=document.getElementById("prev");' +
    '  if(w>0&&p>0){' +
    '    var pr=w*p;var g=(p-h.buyPrice)*w;var gs=g>=0?"+":"";' +
    '    el.innerHTML="Proceeds: <b>RM "+pr.toLocaleString("en-MY",{minimumFractionDigits:2,maximumFractionDigits:2})+"</b>  GL vs buy: <b>"+gs+"RM "+g.toLocaleString("en-MY",{minimumFractionDigits:2,maximumFractionDigits:2})+"</b>";' +
    '    el.style.display="block";' +
    '  }else el.style.display="none";' +
    '}' +
    'function save(){' +
    '  if(idx<0){alert("Select a holding.");return;}' +
    '  var h=H[idx];' +
    '  var wt=parseFloat(document.getElementById("sw").value);' +
    '  var px=parseFloat(document.getElementById("sp").value);' +
    '  if(!wt||wt<=0){alert("Enter a valid weight to sell.");return;}' +
    '  if(wt>h.weight+1e-9){alert("Cannot sell more than held ("+h.weight+"g).");return;}' +
    '  if(!px||px<=0){alert("Enter a valid sell price.");return;}' +
    '  var ret=document.getElementById("ret").value;' +
    '  var notes=document.getElementById("notes").value.trim();' +
    '  document.getElementById("sb").disabled=true;document.getElementById("sb").textContent="Processing...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close();})' +
    '    .withFailureHandler(function(e){document.getElementById("sb").disabled=false;document.getElementById("sb").textContent="Confirm Sell";alert("Error: "+e.message);})' +
    '    .saveSellGold(idx,wt,px,ret,notes);' +
    '}' +
    '</script></body></html>'
  ).setWidth(490).setHeight(500).setTitle('Sell Gold');

  SpreadsheetApp.getUi().showModalDialog(html, 'Sell Gold');
}


// ── SAVE SELL GOLD ────────────────────────────────────────────
function saveSellGold(holdingIdx, sellWeight, sellPrice, returnAccount, notes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GOLD_SHEET_NAME);
  if (!sheet) throw new Error('Gold Portfolio sheet not found.');

  const lastRow  = getLastGoldRow_(sheet);
  const allRows  = sheet.getRange(4, 1, lastRow - 3, 12).getValues();
  const nonEmpty = [];
  allRows.forEach((r, i) => {
    if (r[GC.TYPE-1] !== '' && r[GC.DESC-1] !== '') nonEmpty.push({ r, rowNum: i + 4 });
  });
  if (holdingIdx >= nonEmpty.length) throw new Error('Holding index out of range.');

  const { r, rowNum } = nonEmpty[holdingIdx];
  const goldType   = r[GC.TYPE-1];
  const desc       = r[GC.DESC-1];
  const heldWeight = r[GC.WEIGHT-1];
  const today      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const rmFmt      = '"RM "#,##0.00';

  if (sellWeight > heldWeight + 1e-9) throw new Error('Cannot sell more than held (' + heldWeight + 'g).');
  const remaining = heldWeight - sellWeight;

  if (remaining < 1e-9) {
    sheet.deleteRow(rowNum);
  } else {
    sheet.getRange(rowNum, GC.WEIGHT).setValue(remaining);
    sheet.getRange(rowNum, GC.BUY_TOTAL).setValue(remaining * r[GC.BUY_PRICE-1]).setNumberFormat(rmFmt);
    sheet.getRange(rowNum, GC.UPDATED).setValue(today);
    const en = sheet.getRange(rowNum, GC.NOTES).getValue();
    const sn = 'Sold ' + sellWeight + 'g @ RM ' + sellPrice.toFixed(2) + '/g on ' + today.split(' ')[0];
    sheet.getRange(rowNum, GC.NOTES).setValue(en ? en + ' | ' + sn : sn);
  }

  // Credit to MYR account
  if (returnAccount) {
    const accSheet = ss.getSheetByName(returnAccount);
    if (!accSheet) throw new Error('Account "' + returnAccount + '" not found.');
    const accCcy = normCurrency_(accSheet.getRange('G2').getValue() || 'MYR');
    if (accCcy !== 'MYR') throw new Error('Gold proceeds can only be credited to a MYR account. "' + returnAccount + '" is ' + accCcy + '.');
    const proceeds = sellWeight * sellPrice;
    const lastAcc  = accSheet.getLastRow() + 1;
    const balFml   = lastAcc <= 2 ? '=D' + lastAcc : '=F' + (lastAcc - 1) + '+D' + lastAcc;
    accSheet.getRange(lastAcc, 1).setValue(today.split(' ')[0]);
    accSheet.getRange(lastAcc, 2).setValue('Gold');
    accSheet.getRange(lastAcc, 3).setValue('Sell ' + goldType + ' gold ' + sellWeight + 'g @ RM ' + sellPrice.toFixed(2) + '/g — ' + desc + (notes ? ' (' + notes + ')' : ''));
    accSheet.getRange(lastAcc, 4).setValue(proceeds).setNumberFormat(rmFmt);
    accSheet.getRange(lastAcc, 5).setValue('IN').setFontColor('#0f9d58').setFontWeight('bold');
    accSheet.getRange(lastAcc, 6).setFormula(balFml).setNumberFormat(rmFmt);
    if (lastAcc % 2 === 0) accSheet.getRange(lastAcc, 1, 1, 6).setBackground('#f8f9fa');
  }

  refreshGoldSummary_(sheet);
}


// ── REFRESH GOLD PRICES ───────────────────────────────────────
/**
 * Fetches XAU/MYR once, calculates 916 and 999 per-gram prices,
 * then writes the correct price into col F of each holding row.
 */
function refreshGoldPrices_(sheet) {
  const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const lastRow = getLastGoldRow_(sheet);
  if (lastRow < 4) return;

  const prices = fetchGoldPrices_();
  if (!prices) return; // leave existing values if fetch fails

  for (let row = 4; row <= lastRow; row++) {
    const goldType = sheet.getRange(row, GC.TYPE).getValue();
    const desc     = sheet.getRange(row, GC.DESC).getValue();
    if (!goldType || !desc) continue;
    const curPrice = goldType === '999' ? prices.price999 : prices.price916;
    sheet.getRange(row, GC.CUR_PRICE).setValue(curPrice).setNumberFormat('"RM "#,##0.0000');
    sheet.getRange(row, GC.UPDATED).setValue(today);
  }

  // Update subtitle with XAU rate
  sheet.getRange(2, 1, 1, 12).merge()
    .setValue('916: RM ' + prices.price916.toFixed(2) + '/g  ·  999: RM ' + prices.price999.toFixed(2) + '/g' + (prices.xauMyr ? '  ·  XAU/MYR = RM ' + prices.xauMyr.toFixed(2) + '/troy oz' : '') + '  ·  Source: ' + (prices.source || 'live') + '  ·  Updated: ' + today)
    .setBackground('#fff8e1').setFontColor('#e65100')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  SpreadsheetApp.flush();

  // Recolour gain/loss
  for (let row = 4; row <= lastRow; row++) {
    const gainVal = sheet.getRange(row, GC.GAIN).getValue();
    if (typeof gainVal === 'number') {
      const c = gainVal < 0 ? '#d93025' : '#0f9d58';
      sheet.getRange(row, GC.GAIN).setFontColor(c);
      sheet.getRange(row, GC.GAIN_PCT).setFontColor(c);
    }
  }

  refreshGoldSummary_(sheet);
}


// ── GOLD SUMMARY ──────────────────────────────────────────────
function refreshGoldSummary_(sheet) {
  const lastRow = getLastGoldRow_(sheet);
  const maxRows = sheet.getMaxRows();

  // Clear any existing summary rows below the data (guard: only if there are rows to clear)
  const clearStart = lastRow + 1;
  const clearCount = maxRows - lastRow;
  if (clearStart <= maxRows && clearCount > 0) {
    sheet.getRange(clearStart, 1, clearCount, 12).clearContent().clearFormat();
  }
  if (lastRow < 4) return;

  const data = sheet.getRange(4, 1, lastRow - 3, 9).getValues()
    .filter(r => r[GC.TYPE-1] !== '' && r[GC.DESC-1] !== '');

  const totalCost  = data.reduce((s, r) => s + (typeof r[GC.BUY_TOTAL-1] === 'number' ? r[GC.BUY_TOTAL-1] : 0), 0);
  const totalValue = data.reduce((s, r) => s + (typeof r[GC.CUR_VALUE-1] === 'number' ? r[GC.CUR_VALUE-1] : 0), 0);
  const totalGain  = totalValue - totalCost;
  const gainPct    = totalCost > 0 ? (totalGain / totalCost * 100) : 0;
  const totalWeight916 = data.filter(r => r[GC.TYPE-1] === '916').reduce((s, r) => s + r[GC.WEIGHT-1], 0);
  const totalWeight999 = data.filter(r => r[GC.TYPE-1] === '999').reduce((s, r) => s + r[GC.WEIGHT-1], 0);
  const glColor    = totalGain >= 0 ? '#0f9d58' : '#d93025';
  const rmFmt      = '"RM "#,##0.00';

  const sr = lastRow + 2;
  sheet.setRowHeight(sr, 30);
  sheet.getRange(sr, 1, 1, 12).merge()
    .setValue('📊  PORTFOLIO SUMMARY')
    .setBackground('#e65100').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const rows = [
    ['Total Holdings',       data.length + ' item(s)',  null,    '#202124'],
    ['Total Weight (916)',   totalWeight916.toFixed(3) + ' g', null, '#e65100'],
    ['Total Weight (999)',   totalWeight999.toFixed(3) + ' g', null, '#f57f17'],
    ['Total Cost',           totalCost,                 rmFmt,   '#202124'],
    ['Current Value',        totalValue,                rmFmt,   '#202124'],
    ['Total Gain / Loss',    totalGain,                 rmFmt,   glColor],
    ['Total Gain / Loss (%)', gainPct,                  '0.00"%"', glColor],
  ];

  rows.forEach(([label, val, fmt, color], i) => {
    const r = sr + 1 + i;
    sheet.setRowHeight(r, 28);
    sheet.getRange(r, 1, 1, 3).merge()
      .setValue(label).setFontWeight('bold').setFontSize(10)
      .setBackground(i % 2 === 0 ? '#fff8e1' : '#fff3e0').setVerticalAlignment('middle');
    const vc = sheet.getRange(r, 4);
    vc.setValue(val).setFontColor(color).setFontWeight('bold').setVerticalAlignment('middle');
    if (fmt) vc.setNumberFormat(fmt);
  });
}


// ============================================================
//  RETIREMENT PORTFOLIO TRACKER
//  Sheet: 🏖️ Retirement Portfolio
//  Lets the user tag any accounts + mutual funds as retirement
//  assets and renders a dedicated summary sheet.
// ============================================================

const RETIRE_SHEET_NAME  = '🏖️ Retirement Portfolio';
const RETIRE_CONFIG_KEY  = 'retirementConfig';   // JSON in ScriptProperties

// ── READ / WRITE CONFIG ───────────────────────────────────────

function getRetirementConfig_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(RETIRE_CONFIG_KEY);
    if (!raw) return { accounts: [], funds: [], goldItems: [] };
    const cfg = JSON.parse(raw);
    if (!cfg.goldItems) cfg.goldItems = [];   // migrate older configs
    return cfg;
  } catch(_) { return { accounts: [], funds: [], goldItems: [] }; }
}

function saveRetirementConfig_(cfg) {
  PropertiesService.getScriptProperties().setProperty(RETIRE_CONFIG_KEY, JSON.stringify(cfg));
}

// ── CONFIG DIALOG ─────────────────────────────────────────────

function showRetirementConfigDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Collect all accounts
  const accSheets = ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'));
  const accounts = accSheets.map(s => ({
    name:     s.getName(),
    currency: normCurrency_(s.getRange('G2').getValue() || 'MYR')
  }));

  // Collect all mutual funds
  const mfSheet = ss.getSheetByName(MF_SHEET_NAME);
  const funds = [];
  if (mfSheet) {
    const lastMF = getLastFundRow_(mfSheet);
    if (lastMF >= 4) {
      mfSheet.getRange(4, 1, lastMF - 3, MF_COLS.ACCOUNT).getValues()
        .filter(r => r[MF_COLS.CODE - 1] !== '')
        .forEach(r => {
          const name = r[MF_COLS.NAME - 1];
          const code = r[MF_COLS.CODE - 1];
          if (code) funds.push({ name: name || code, code: code, currency: r[MF_COLS.CCY - 1] || 'MYR' });
        });
    }
  }

  // Collect all gold items
  const goldSheet = ss.getSheetByName(GOLD_SHEET_NAME);
  const goldItems = [];
  if (goldSheet) {
    const lastGold = getLastGoldRow_(goldSheet);
    if (lastGold >= 4) {
      goldSheet.getRange(4, 1, lastGold - 3, GC.GAIN_PCT).getValues()
        .filter(r => r[GC.TYPE - 1] !== '' && r[GC.DESC - 1] !== '')
        .forEach(function(r, i) {
          const type   = r[GC.TYPE - 1];
          const desc   = r[GC.DESC - 1];
          const weight = r[GC.WEIGHT - 1] || 0;
          const key    = 'gold_row_' + (i + 4);
          goldItems.push({ key: key, label: type + ' \u2013 ' + desc, weight: weight, type: type });
        });
    }
  }

  const cfg  = getRetirementConfig_();
  const data = JSON.stringify({ accounts: accounts, funds: funds, goldItems: goldItems, cfg: cfg });

  const htmlStr = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
    + 'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:0;background:#f8f9fa;color:#202124;font-size:13px}'
    + '.hdr{background:#1b5e20;color:#fff;padding:14px 18px;font-size:15px;font-weight:700}'
    + '.body{padding:16px;max-height:480px;overflow-y:auto}'
    + '.section{margin-bottom:16px}'
    + '.section-title{font-weight:700;font-size:12px;color:#1b5e20;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px;padding-bottom:4px;border-bottom:2px solid #c8e6c9}'
    + '.item{display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:6px;margin-bottom:4px;cursor:pointer;transition:background .1s}'
    + '.item:hover{background:#f1f8e9}'
    + '.item.checked{background:#e8f5e9}'
    + '.gold-item.checked{background:#fff8e1!important}'
    + '.gold-item:hover{background:#fffde7!important}'
    + 'input[type=checkbox]{width:16px;height:16px;accent-color:#2e7d32;cursor:pointer;flex-shrink:0}'
    + '.item-name{font-weight:600;font-size:13px}'
    + '.item-sub{font-size:11px;color:#5f6368;margin-left:auto}'
    + '.empty{color:#9aa0a6;font-style:italic;font-size:12px;padding:8px}'
    + '.footer{display:flex;justify-content:flex-end;gap:8px;padding:12px 16px;border-top:1px solid #e8eaed;background:#fff}'
    + '.btn{padding:8px 22px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}'
    + '.bsave{background:#2e7d32;color:#fff}.bsave:hover{background:#1b5e20}'
    + '.bcancel{background:#f1f3f4;color:#3c4043}'
    + '.selectall{font-size:11px;color:#2e7d32;cursor:pointer;text-decoration:underline;float:right}'
    + '</style></head><body>'
    + '<div class="hdr">🏖️ Configure Retirement Portfolio</div>'
    + '<div class="body">'
    + '<div class="section">'
    + '<div class="section-title">Cash / Savings Accounts <span class="selectall" id="selAllAcc">Select all</span></div>'
    + '<div id="accList"></div>'
    + '</div>'
    + '<div class="section">'
    + '<div class="section-title">Mutual Funds / Unit Trusts <span class="selectall" id="selAllFunds">Select all</span></div>'
    + '<div id="fundList"></div>'
    + '</div>'
    + '<div class="section">'
    + '<div class="section-title">🥇 Gold Investments <span class="selectall" id="selAllGold">Select all</span></div>'
    + '<div id="goldList"></div>'
    + '</div>'
    + '</div>'
    + '<div class="footer">'
    + '<button class="btn bcancel" id="cancelBtn">Cancel</button>'
    + '<button class="btn bsave" id="saveBtn">Save & Refresh Sheet</button>'
    + '</div>'
    + '<script>'
    + 'var D=' + data + ';'
    + 'var selAccs=new Set(D.cfg.accounts||[]);'
    + 'var selFunds=new Set(D.cfg.funds||[]);'
    + 'var selGold=new Set(D.cfg.goldItems||[]);'
    + 'function buildList(cid,items,selSet,keyFn,labelFn,subFn,xCls){'
    + '  var el=document.getElementById(cid);'
    + '  if(!items.length){el.innerHTML="<div class=\\"empty\\">None found.</div>";return;}'
    + '  items.forEach(function(item){'
    + '    var key=keyFn(item);'
    + '    var div=document.createElement("div");'
    + '    div.className="item"+(xCls?" "+xCls:"")+(selSet.has(key)?" checked":"");'
    + '    var cb=document.createElement("input");cb.type="checkbox";cb.checked=selSet.has(key);cb.dataset.key=key;'
    + '    cb.onchange=function(){if(this.checked){selSet.add(key);div.classList.add("checked");}else{selSet.delete(key);div.classList.remove("checked");}};'
    + '    div.onclick=function(e){if(e.target!==cb)cb.click();};'
    + '    var nm=document.createElement("span");nm.className="item-name";nm.textContent=labelFn(item);'
    + '    var sb=document.createElement("span");sb.className="item-sub";sb.textContent=subFn(item);'
    + '    div.append(cb,nm,sb);el.appendChild(div);'
    + '  });'
    + '}'
    + 'buildList("accList",D.accounts,selAccs,function(a){return a.name;},function(a){return a.name;},function(a){return a.currency;},null);'
    + 'buildList("fundList",D.funds,selFunds,function(f){return f.code;},function(f){return f.name;},function(f){return f.code+" \u00b7 "+f.currency;},null);'
    + 'buildList("goldList",D.goldItems,selGold,function(g){return g.key;},function(g){return g.label;},function(g){return g.weight.toFixed(2)+"g"},"gold-item");'
    + 'document.getElementById("selAllAcc").onclick=function(){'
    + '  D.accounts.forEach(function(a){selAccs.add(a.name);});'
    + '  document.querySelectorAll("#accList input").forEach(function(c){c.checked=true;c.closest(".item").classList.add("checked");});'
    + '};'
    + 'document.getElementById("selAllFunds").onclick=function(){'
    + '  D.funds.forEach(function(f){selFunds.add(f.code);});'
    + '  document.querySelectorAll("#fundList input").forEach(function(c){c.checked=true;c.closest(".item").classList.add("checked");});'
    + '};'
    + 'document.getElementById("selAllGold").onclick=function(){'
    + '  D.goldItems.forEach(function(g){selGold.add(g.key);});'
    + '  document.querySelectorAll("#goldList input").forEach(function(c){c.checked=true;c.closest(".item").classList.add("checked");});'
    + '};'
    + 'document.getElementById("cancelBtn").onclick=function(){google.script.host.close();};'
    + 'document.getElementById("saveBtn").onclick=function(){'
    + '  document.getElementById("saveBtn").disabled=true;'
    + '  document.getElementById("saveBtn").textContent="Saving...";'
    + '  google.script.run'
    + '    .withSuccessHandler(function(){google.script.host.close();})'
    + '    .withFailureHandler(function(e){alert("Error: "+e.message);document.getElementById("saveBtn").disabled=false;document.getElementById("saveBtn").textContent="Save & Refresh Sheet";})'
    + '    .saveRetirementConfigAndRender(Array.from(selAccs),Array.from(selFunds),Array.from(selGold));'
    + '};'
    + '</script></body></html>';

  const html = HtmlService.createHtmlOutput(htmlStr)
    .setWidth(500).setHeight(640).setTitle('Configure Retirement Portfolio');
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Retirement Portfolio');
}

function saveRetirementConfigAndRender(selectedAccounts, selectedFunds, selectedGoldItems) {
  saveRetirementConfig_({ accounts: selectedAccounts, funds: selectedFunds, goldItems: selectedGoldItems || [] });
  renderRetirementSheet();
}

// ── RENDER RETIREMENT SHEET ───────────────────────────────────

function renderRetirementSheet() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const cfg  = getRetirementConfig_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  // ── Collect accounts ────────────────────────────────────────
  const accRows = [];
  let   needFX  = new Set();
  ss.getSheets()
    .filter(s => s.getRange('F2').getValue().toString().includes('Balance'))
    .filter(s => cfg.accounts.includes(s.getName()))
    .forEach(s => {
      const lastRow = s.getLastRow();
      const balance = lastRow >= 3 ? (s.getRange(lastRow, 6).getValue() || 0) : 0;
      const ccy     = normCurrency_(s.getRange('G2').getValue() || 'MYR');
      if (ccy !== 'MYR') needFX.add(ccy);
      accRows.push({ name: s.getName(), balance: balance, ccy: ccy });
    });

  // ── Collect mutual funds ─────────────────────────────────────
  const fundRows = [];
  const mfSheet  = ss.getSheetByName(MF_SHEET_NAME);
  if (mfSheet && cfg.funds && cfg.funds.length > 0) {
    const lastMF = getLastFundRow_(mfSheet);
    if (lastMF >= 4) {
      mfSheet.getRange(4, 1, lastMF - 3, MF_COLS.MKT_VAL).getValues()
        .filter(r => r[MF_COLS.CODE - 1] !== '' && cfg.funds.includes(r[MF_COLS.CODE - 1]))
        .forEach(r => {
          const ccy    = normCurrency_(r[MF_COLS.CCY - 1] || 'MYR');
          const mktVal = typeof r[MF_COLS.MKT_VAL - 1] === 'number' ? r[MF_COLS.MKT_VAL - 1] : 0;
          const cost   = (r[MF_COLS.UNITS - 1] || 0) * (r[MF_COLS.BUY_NAV - 1] || 0);
          const gain   = mktVal - cost;
          if (ccy !== 'MYR') needFX.add(ccy);
          fundRows.push({
            name: r[MF_COLS.NAME - 1] || r[MF_COLS.CODE - 1],
            code: r[MF_COLS.CODE - 1],
            units: r[MF_COLS.UNITS - 1] || 0,
            curNav: r[MF_COLS.CUR_NAV - 1] || 0,
            mktVal: mktVal, cost: cost, gain: gain, ccy: ccy
          });
        });
    }
  }

  // ── Collect gold items ────────────────────────────────────────
  const goldRows = [];
  const goldSheet = ss.getSheetByName(GOLD_SHEET_NAME);
  if (goldSheet && cfg.goldItems && cfg.goldItems.length > 0) {
    const lastGold = getLastGoldRow_(goldSheet);
    if (lastGold >= 4) {
      goldSheet.getRange(4, 1, lastGold - 3, GC.GAIN).getValues()
        .forEach(function(r, i) {
          const key = 'gold_row_' + (i + 4);
          if (!cfg.goldItems.includes(key)) return;
          if (r[GC.TYPE - 1] === '' || r[GC.DESC - 1] === '') return;
          const type    = r[GC.TYPE - 1];
          const desc    = r[GC.DESC - 1];
          const weight  = r[GC.WEIGHT - 1] || 0;
          const buyTot  = typeof r[GC.BUY_TOTAL - 1] === 'number' ? r[GC.BUY_TOTAL - 1] : 0;
          const curVal  = typeof r[GC.CUR_VALUE - 1] === 'number' ? r[GC.CUR_VALUE - 1] : buyTot;
          const gain    = typeof r[GC.GAIN - 1] === 'number' ? r[GC.GAIN - 1] : (curVal - buyTot);
          goldRows.push({ type: type, desc: desc, weight: weight, buyTot: buyTot, curVal: curVal, gain: gain });
        });
    }
  }

  // ── FX rates (gold is always MYR, so only needed for acc/funds) ──
  const fxRates = {};
  const fxList  = [...needFX];
  if (fxList.length > 0) {
    try {
      let temp = ss.getSheetByName('_FX_TEMP_');
      if (temp) ss.deleteSheet(temp);
      temp = ss.insertSheet('_FX_TEMP_');
      compactSheet_(temp, 10, 1); // 1 col × 10 rows
      fxList.forEach(function(cur, i) { temp.getRange(i + 1, 1).setFormula('=GOOGLEFINANCE("CURRENCY:' + cur + 'MYR")'); });
      SpreadsheetApp.flush();
      Utilities.sleep(3000);
      fxList.forEach(function(cur, i) {
        const v = temp.getRange(i + 1, 1).getValue();
        fxRates[cur] = (typeof v === 'number' && v > 0) ? v : null;
      });
      ss.deleteSheet(temp);
    } catch(e) {
      try { const t = ss.getSheetByName('_FX_TEMP_'); if (t) ss.deleteSheet(t); } catch(_) {}
      fxList.forEach(function(cur) { fxRates[cur] = null; });
    }
  }

  function toMYR(amount, ccy) {
    if (ccy === 'MYR') return { ok: true, myr: amount };
    const rate = fxRates[ccy];
    if (!rate) return { ok: false, myr: 0 };
    return { ok: true, myr: amount * rate };
  }

  // ── Get or create sheet ──────────────────────────────────────
  let sheet = ss.getSheetByName(RETIRE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RETIRE_SHEET_NAME, 1);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  // ── Layout constants ─────────────────────────────────────────
  const COL_ICON  = 1;
  const COL_NAME  = 2;
  const COL_VALUE = 3;
  const COL_CCY   = 4;
  const COL_MYR   = 5;
  const COL_EXTRA = 6;
  const NCOLS     = 6;
  const myrFmt    = '"RM "#,##0.00';
  const GREEN     = '#1b5e20';
  const LGREEN    = '#2e7d32';
  const MIDG      = '#43a047';
  const WHITE     = '#ffffff';
  const GOLD_HDR_BG   = '#e65100';
  const GOLD_TOT_BG   = '#bf360c';
  const ACC_COLORS    = ['#e8f5e9','#f1f8e9'];
  const FUND_COLORS   = ['#e3f2fd','#e8eaf6'];
  const GOLD_COLORS   = ['#fff8e1','#fff3e0'];

  // ── Column widths ────────────────────────────────────────────
  sheet.setColumnWidth(COL_ICON,  36);
  sheet.setColumnWidth(COL_NAME,  240);
  sheet.setColumnWidth(COL_VALUE, 160);
  sheet.setColumnWidth(COL_CCY,   60);
  sheet.setColumnWidth(COL_MYR,   10);
  sheet.setColumnWidth(COL_EXTRA, 160);

  // ── Row layout ───────────────────────────────────────────────
  const BANNER_ROW   = 1;
  const SUBTITLE_ROW = 2;
  const SPACER1      = 3;
  const ACC_HDR      = 4;
  const ACC_COL_HDR  = 5;
  const ACC_START    = 6;
  const ACC_TOTAL    = ACC_START + Math.max(accRows.length, 1);
  const SPACER2      = ACC_TOTAL + 1;
  const FUND_HDR     = ACC_TOTAL + 2;
  const FUND_COL_HDR = FUND_HDR + 1;
  const FUND_START   = FUND_COL_HDR + 1;
  const FUND_TOTAL   = FUND_START + Math.max(fundRows.length, 1);
  const SPACER3      = FUND_TOTAL + 1;
  const GOLD_HDR     = FUND_TOTAL + 2;
  const GOLD_COL_HDR = GOLD_HDR + 1;
  const GOLD_START   = GOLD_COL_HDR + 1;
  const GOLD_TOTAL   = GOLD_START + Math.max(goldRows.length, 1);
  const SPACER4      = GOLD_TOTAL + 1;
  const GRAND_ROW    = GOLD_TOTAL + 2;
  const TOTAL_ROWS   = GRAND_ROW + 5;

  if (sheet.getMaxRows() < TOTAL_ROWS) {
    sheet.insertRowsAfter(sheet.getMaxRows(), TOTAL_ROWS - sheet.getMaxRows());
  }
  compactSheet_(sheet, TOTAL_ROWS, NCOLS);

  // ── Helpers ──────────────────────────────────────────────────
  function writeColHdr_(row, valLabel, extraLabel) {
    sheet.setRowHeight(row, 28);
    [1,2,3,4,5,6].forEach(function(c) { sheet.getRange(row, c).setBackground('#c8e6c9'); });
    sheet.getRange(row, COL_NAME) .setValue('Name').setFontWeight('bold').setFontSize(10).setFontColor(GREEN).setVerticalAlignment('middle');
    sheet.getRange(row, COL_VALUE).setValue(valLabel || 'Value').setFontWeight('bold').setFontSize(10).setFontColor(GREEN).setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(row, COL_CCY)  .setValue('CCY').setFontWeight('bold').setFontSize(10).setFontColor(GREEN).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange(row, COL_EXTRA).setValue(extraLabel || '≈ RM Equivalent').setFontWeight('bold').setFontSize(10).setFontColor(GREEN).setHorizontalAlignment('right').setVerticalAlignment('middle');
  }

  function writeSectionHdr_(row, icon, label, bg) {
    sheet.setRowHeight(row - 1, 14);
    sheet.setRowHeight(row, 28);
    sheet.getRange(row, 1, 1, NCOLS).merge()
      .setValue(icon + '  ' + label)
      .setBackground(bg).setFontColor(WHITE)
      .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  }

  function writeTotalRow_(row, label, myrTotal, bg) {
    sheet.setRowHeight(row, 40);
    sheet.getRange(row, 1, 1, NCOLS).setBackground(bg);
    sheet.getRange(row, COL_NAME).setValue(label)
      .setFontSize(11).setFontWeight('bold').setFontColor(WHITE).setVerticalAlignment('middle');
    if (myrTotal !== null) {
      sheet.getRange(row, COL_EXTRA).setValue(myrTotal)
        .setNumberFormat(myrFmt).setFontSize(13).setFontWeight('bold')
        .setFontColor(WHITE).setHorizontalAlignment('right').setVerticalAlignment('middle');
    } else {
      sheet.getRange(row, COL_EXTRA).setValue('Partial (FX unavailable)')
        .setFontSize(10).setFontWeight('bold').setFontColor('#ffccbc')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    }
  }

  // ── 1. Banner ────────────────────────────────────────────────
  sheet.setRowHeight(BANNER_ROW, 56);
  sheet.getRange(BANNER_ROW, 1, 1, NCOLS).merge()
    .setValue('🏖️  RETIREMENT PORTFOLIO')
    .setBackground(GREEN).setFontColor(WHITE)
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(SUBTITLE_ROW, 26);
  const fxNote = fxList.length > 0 ? '  ·  FX rates live-snapshotted at refresh' : '';
  sheet.getRange(SUBTITLE_ROW, 1, 1, NCOLS).merge()
    .setValue('Last updated: ' + today + fxNote + '  ·  Use 🏖️ Refresh Retirement Sheet to update')
    .setBackground('#e8f5e9').setFontColor('#388e3c')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setRowHeight(SPACER1, 14);

  // ── 2. Accounts section ──────────────────────────────────────
  writeSectionHdr_(ACC_HDR, '🏦', 'CASH & SAVINGS ACCOUNTS', LGREEN);
  writeColHdr_(ACC_COL_HDR, 'Balance', '≈ RM Equivalent');

  let accMYR = 0, accFxMissing = false;
  if (accRows.length > 0) {
    accRows.forEach(function(acc, i) {
      const row = ACC_START + i;
      const r   = toMYR(acc.balance, acc.ccy);
      if (r.ok) accMYR += r.myr; else accFxMissing = true;
      const bg  = ACC_COLORS[i % ACC_COLORS.length];
      const fmt = acc.ccy === 'MYR' ? myrFmt : '#,##0.00';
      sheet.setRowHeight(row, 34);
      sheet.getRange(row, 1, 1, NCOLS).setBackground(bg);
      sheet.getRange(row, COL_ICON) .setValue('🏦').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(row, COL_NAME) .setValue(acc.name).setFontSize(11).setFontWeight('bold').setFontColor('#202124').setVerticalAlignment('middle');
      // Live formula — always reads latest balance from account sheet
      const safeAcc = acc.name.replace(/'/g, "''");
      sheet.getRange(row, COL_VALUE)
        .setFormula('=IFERROR(INDEX(\'' + safeAcc + '\'!F:F,MAX(ARRAYFORMULA(IF(ISNUMBER(\'' + safeAcc + '\'!F3:F1000),ROW(\'' + safeAcc + '\'!F3:F1000),0)))),0)')
        .setNumberFormat(fmt).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(row, COL_CCY)  .setValue(acc.ccy).setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(row, COL_MYR)  .setBackground('#e8f5e9');
      // Live MYR equivalent using GOOGLEFINANCE
      const retAccCur = acc.ccy === 'MYR' ? 'MYR' : acc.ccy;
      const retFxFormula = retAccCur === 'MYR'
        ? '=' + sheet.getRange(row, COL_VALUE).getA1Notation()
        : '=IFERROR(' + sheet.getRange(row, COL_VALUE).getA1Notation() + '*GOOGLEFINANCE("CURRENCY:' + retAccCur + 'MYR"),0)';
      sheet.getRange(row, COL_EXTRA).setFormula(retFxFormula)
        .setNumberFormat(myrFmt).setFontSize(12).setFontWeight('bold').setFontColor('#1b5e20')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    });
  } else {
    sheet.setRowHeight(ACC_START, 30);
    sheet.getRange(ACC_START, COL_NAME, 1, 4).merge()
      .setValue('No accounts selected. Use 🏖️ Configure Retirement Portfolio to add accounts.')
      .setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(ACC_TOTAL, 'Total Cash & Savings', accFxMissing ? null : accMYR, MIDG);
  sheet.setRowHeight(SPACER2, 14);

  // ── 3. Mutual Funds section ──────────────────────────────────
  writeSectionHdr_(FUND_HDR, '📈', 'MUTUAL FUNDS / UNIT TRUSTS', '#1565c0');
  writeColHdr_(FUND_COL_HDR, 'Market Value', '≈ RM Equivalent  (Gain/Loss)');

  let fundMYR = 0, fundFxMissing = false;
  if (fundRows.length > 0) {
    fundRows.forEach(function(f, i) {
      const row = FUND_START + i;
      const r   = toMYR(f.mktVal, f.ccy);
      const rg  = toMYR(f.gain,   f.ccy);
      if (r.ok) fundMYR += r.myr; else fundFxMissing = true;
      const bg     = FUND_COLORS[i % FUND_COLORS.length];
      const fmt    = f.ccy === 'MYR' ? myrFmt : '"' + f.ccy + ' "#,##0.00';
      const gSign  = f.gain >= 0 ? '+' : '';
      sheet.setRowHeight(row, 38);
      sheet.getRange(row, 1, 1, NCOLS).setBackground(bg);
      sheet.getRange(row, COL_ICON).setValue('📈').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(row, COL_NAME).setValue(f.name + '\n' + f.code)
        .setFontSize(10).setFontWeight('bold').setFontColor('#202124').setVerticalAlignment('middle').setWrap(true);
      sheet.getRange(row, COL_VALUE).setValue(f.mktVal).setNumberFormat(fmt)
        .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(row, COL_CCY).setValue(f.ccy)
        .setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(row, COL_MYR).setBackground(r.ok ? '#e3f2fd' : '#fff3e0');
      if (r.ok) {
        const gainStr = gSign + 'RM ' + Math.abs(rg.myr).toLocaleString('en-MY', {minimumFractionDigits:2, maximumFractionDigits:2});
        sheet.getRange(row, COL_EXTRA).setValue('RM ' + r.myr.toFixed(2) + '  (' + gainStr + ')')
          .setFontSize(10).setFontWeight('bold').setFontColor('#1565c0').setHorizontalAlignment('right').setVerticalAlignment('middle');
      } else {
        sheet.getRange(row, COL_EXTRA).setValue('FX unavailable').setFontSize(10).setFontColor('#e65100').setHorizontalAlignment('right').setVerticalAlignment('middle');
      }
    });
  } else {
    sheet.setRowHeight(FUND_START, 30);
    sheet.getRange(FUND_START, COL_NAME, 1, 4).merge()
      .setValue('No funds selected. Use 🏖️ Configure Retirement Portfolio to add funds.')
      .setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(FUND_TOTAL, 'Total Mutual Funds', fundFxMissing ? null : fundMYR, '#1565c0');
  sheet.setRowHeight(SPACER3, 14);

  // ── 4. Gold section ──────────────────────────────────────────
  writeSectionHdr_(GOLD_HDR, '🥇', 'GOLD INVESTMENTS', GOLD_HDR_BG);
  writeColHdr_(GOLD_COL_HDR, 'Current Value (RM)', 'RM Value  (Gain/Loss)');

  let goldMYR = 0;
  if (goldRows.length > 0) {
    goldRows.forEach(function(g, i) {
      const row    = GOLD_START + i;
      const bg     = GOLD_COLORS[i % GOLD_COLORS.length];
      const gSign  = g.gain >= 0 ? '+' : '';
      const gainStr = gSign + 'RM ' + Math.abs(g.gain).toLocaleString('en-MY', {minimumFractionDigits:2, maximumFractionDigits:2});
      goldMYR += g.curVal;
      sheet.setRowHeight(row, 38);
      sheet.getRange(row, 1, 1, NCOLS).setBackground(bg);
      sheet.getRange(row, COL_ICON).setValue('🥇').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(row, COL_NAME).setValue(g.type + ' – ' + g.desc + '\n' + g.weight.toFixed(3) + 'g')
        .setFontSize(10).setFontWeight('bold').setFontColor('#202124').setVerticalAlignment('middle').setWrap(true);
      sheet.getRange(row, COL_VALUE).setValue(g.curVal).setNumberFormat(myrFmt)
        .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(row, COL_CCY).setValue('MYR')
        .setFontSize(9).setFontWeight('bold').setFontColor('#5f6368').setHorizontalAlignment('center').setVerticalAlignment('middle');
      sheet.getRange(row, COL_MYR).setBackground('#fff8e1');
      sheet.getRange(row, COL_EXTRA).setValue('RM ' + g.curVal.toFixed(2) + '  (' + gainStr + ')')
        .setFontSize(10).setFontWeight('bold').setFontColor(g.gain >= 0 ? '#bf360c' : '#d93025')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
    });
  } else {
    sheet.setRowHeight(GOLD_START, 30);
    sheet.getRange(GOLD_START, COL_NAME, 1, 4).merge()
      .setValue('No gold items selected. Use 🏖️ Configure Retirement Portfolio to add gold.')
      .setFontColor('#9aa0a6').setFontStyle('italic').setHorizontalAlignment('center');
  }
  writeTotalRow_(GOLD_TOTAL, 'Total Gold', goldMYR, GOLD_TOT_BG);
  sheet.setRowHeight(SPACER4, 14);

  // ── 5. Grand Total ────────────────────────────────────────────
  const anyMissing = accFxMissing || fundFxMissing;
  const grandTotal = anyMissing ? null : accMYR + fundMYR + goldMYR;

  sheet.setRowHeight(GRAND_ROW, 60);
  sheet.getRange(GRAND_ROW, 1, 1, NCOLS).setBackground(GREEN);
  sheet.getRange(GRAND_ROW, COL_NAME).setValue('🏆  TOTAL RETIREMENT SAVINGS')
    .setFontSize(14).setFontWeight('bold').setFontColor(WHITE).setVerticalAlignment('middle');
  if (grandTotal !== null) {
    sheet.getRange(GRAND_ROW, COL_EXTRA).setValue(grandTotal)
      .setNumberFormat(myrFmt).setFontSize(16).setFontWeight('bold')
      .setFontColor(WHITE).setHorizontalAlignment('right').setVerticalAlignment('middle');
  } else {
    sheet.getRange(GRAND_ROW, COL_EXTRA).setValue('Partial (FX unavailable)')
      .setFontSize(11).setFontWeight('bold').setFontColor('#c8e6c9')
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
  }

  // ── 6. Breakdown mini-summary ──────────────────────────────
  if (grandTotal !== null && grandTotal > 0) {
    const accPct  = Math.round(accMYR  / grandTotal * 100);
    const fundPct = Math.round(fundMYR / grandTotal * 100);
    const goldPct = Math.round(goldMYR / grandTotal * 100);
    const summaryRow = GRAND_ROW + 2;
    sheet.setRowHeight(GRAND_ROW + 1, 10);
    sheet.setRowHeight(summaryRow, 30);
    sheet.getRange(summaryRow, 1, 1, NCOLS).setBackground('#f1f8e9');
    sheet.getRange(summaryRow, COL_NAME).setValue(
      '🏦 Cash: RM ' + accMYR.toFixed(2) + ' (' + accPct + '%)   '
      + '📈 Funds: RM ' + fundMYR.toFixed(2) + ' (' + fundPct + '%)   '
      + '🥇 Gold: RM ' + goldMYR.toFixed(2) + ' (' + goldPct + '%)'
    ).setFontSize(10).setFontColor('#2e7d32').setVerticalAlignment('middle');
  }

  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(2);
  sheet.activate();
}

function refreshRetirementSheet() {
  renderRetirementSheet();
}


// ============================================================
//  FIMM NAV CACHE
//  Downloads the latest fund price PDF from fimm.com.my,
//  OCR-converts it via Google Drive, parses fund code + NAV,
//  and writes results to sheet "📋 FIMM NAV Cache".
//  FSMFund() looks here first before calling FSMOne.
//  Daily trigger: 7 AM.
// ============================================================

// ── MAIN ENTRY POINTS ─────────────────────────────────────────

/** Menu: refresh now */
function refreshFimmNavCache() {
  try {
    Logger.log('[FIMM] Starting refresh...');

    // 1. Fetch PDF bytes server-side
    const pdfRes = UrlFetchApp.fetch(FIMM_PDF_URL, {
      method: 'get', muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript/1.0)' }
    });
    Logger.log('[FIMM] PDF HTTP ' + pdfRes.getResponseCode());
    if (pdfRes.getResponseCode() !== 200) {
      throw new Error('FIMM PDF fetch failed (HTTP ' + pdfRes.getResponseCode() + ')');
    }
    const pdfBase64 = Utilities.base64Encode(pdfRes.getBlob().getBytes());
    Logger.log('[FIMM] PDF base64 length: ' + pdfBase64.length);

    // 2. Fetch pako inflate (only 6.9 KB) to inline in dialog for PDF decompression
    const pakoRes = UrlFetchApp.fetch(
      'https://cdnjs.cloudflare.com/ajax/libs/pako/2.1.0/pako_inflate.min.js',
      { muteHttpExceptions: true }
    );
    Logger.log('[FIMM] pako HTTP ' + pakoRes.getResponseCode());
    if (pakoRes.getResponseCode() !== 200) {
      throw new Error('Could not fetch pako library (HTTP ' + pakoRes.getResponseCode() + ')');
    }
    const pakoCode = pakoRes.getContentText();
    Logger.log('[FIMM] pako size: ' + pakoCode.length + ' bytes');

    // 3. Build and show dialog — pako+parser run in browser, result sent back to GAS
    const html = buildFimmDialog_(pdfBase64, pakoCode);
    Logger.log('[FIMM] Showing dialog...');
    SpreadsheetApp.getUi().showModalDialog(html, '📋 Refreshing FIMM NAV Cache...');
    Logger.log('[FIMM] Dialog shown.');
  } catch(e) {
    Logger.log('[FIMM] ERROR: ' + e.message);
    SpreadsheetApp.getUi().alert('Error updating FIMM NAV Cache:\n' + e.message);
  }
}

/** Called by dialog with parsed records JSON */
function saveFimmRecordsFromDialog(recordsJson) {
  try {
    Logger.log('[FIMM] saveFimmRecordsFromDialog: parsing...');
    const records = JSON.parse(recordsJson);
    Logger.log('[FIMM] Records received: ' + records.length);
    writeFimmCache_(records);
    Logger.log('[FIMM] Cache written OK');
    return { success: true, count: records.length };
  } catch(e) {
    Logger.log('[FIMM] Save error: ' + e.message);
    return { success: false, error: e.message };
  }
}

/**
 * Dialog uses pako (6.9KB, inlined) to decompress PDF FlateDecode streams,
 * then parses PDF content stream operators (BT/ET/Td/Tj) to extract text,
 * then runs the flat-text state-machine to pair fund names with NAVs.
 * No PDF.js needed — no worker, no CDN, no CSP issues.
 */
function buildFimmDialog_(pdfBase64, pakoCode) {
  const parserJs =
    'var PDF_B64=' + JSON.stringify(pdfBase64) + ';' +

    'function b64ToArr(b){var bin=atob(b),a=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)a[i]=bin.charCodeAt(i);return a;}' +

    // Extract and decompress all FlateDecode streams from raw PDF bytes
    'function extractStreams(bytes){' +
    '  var raw="";for(var i=0;i<bytes.length;i++)raw+=String.fromCharCode(bytes[i]);' +
    '  var re=/\\d+ 0 obj[\\s\\S]*?\\/Length (\\d+)[\\s\\S]*?stream\\r?\\n([\\s\\S]*?)endstream/g;' +
    '  var streams=[],m;' +
    '  while((m=re.exec(raw))!==null){' +
    '    if(!m[0].includes("FlateDecode"))continue;' +
    '    var len=parseInt(m[1]);' +
    '    var off=m.index+m[0].indexOf("stream\\n")+7;' +
    '    var sb=bytes.slice(off,off+len);' +
    '    try{var inf=pako.inflate(sb);streams.push(new TextDecoder("latin1").decode(inf));}catch(e){}' +
    '  }' +
    '  return streams;' +
    '}' +

    // Parse PDF content stream: BT/ET blocks with Td/Tj operators -> text strings in order
    'function parseStream(s){' +
    '  var items=[],blocks=s.split(/\\bBT\\b/);' +
    '  for(var b=1;b<blocks.length;b++){' +
    '    var block=blocks[b].split(/\\bET\\b/)[0];' +
    '    var ls=block.split("\\n");' +
    '    for(var i=0;i<ls.length;i++){' +
    '      var l=ls[i].trim();' +
    '      var tj=l.match(/^\\(([^)]*)\\)Tj$/);' +
    '      if(tj){var s2=tj[1].replace(/\\\\\\(/g,"(").replace(/\\\\\\)/g,")").trim();if(s2)items.push(s2);}' +
    '    }' +
    '  }' +
    '  return items;' +
    '}' +

    // State-machine: accumulate name parts until NAV number
    'function parseRecords(lines){' +
    '  var NR=/^\\d+\\.\\d{3,6}$/;' +
    '  var DR=/^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i;' +
    '  var HR=/\\b(SDN BHD|BERHAD|ASSET MANAGEMENT|FORMERLY KNOWN|PTE LTD|CORPORATION)/i;' +
    '  var JR=/(\\d+\\.\\d+%|p\\.a\\.|preceding business|Daily Dividend|\\bSUSP\\b)/i;' +
    '  var AR=/\\s*\\([a-uw-z]\\)\\s*/gi;' +
    '  var records=[],seen={},parts=[];' +
    '  for(var i=0;i<lines.length;i++){' +
    '    var ln=lines[i].trim();if(!ln)continue;' +
    '    if(DR.test(ln)||/^(NAV|RM|-|\\d{1,3})$/.test(ln)){parts=[];continue;}' +
    '    if(/^(NA|N\\/A)$/i.test(ln)){parts=[];continue;}' +
    '    if(NR.test(ln)){' +
    '      if(parts.length>0){' +
    '        var nm=parts.join(" ").replace(AR," ").replace(/\\s{2,}/g," ").trim().replace(/\\s*\\([a-z]\\)$/,"").trim();' +
    '        var nav=parseFloat(ln);' +
    '        if(nm.length>=4&&nm.length<=80&&/^[A-Z]/.test(nm)&&!HR.test(nm)&&!JR.test(nm)&&nav>0&&!seen[nm]){' +
    '          seen[nm]=true;records.push([nm,nav]);' +
    '        }' +
    '      }' +
    '      parts=[];continue;' +
    '    }' +
    '    if(HR.test(ln)||JR.test(ln)){parts=[];continue;}' +
    '    var st=ln.replace(AR,"").trim();' +
    '    if(!st||/^\\([a-z]\\)$/.test(st))continue;' +
    '    if(/^[A-Z]/.test(st)||parts.length>0)parts.push(st);' +
    '  }' +
    '  return records;' +
    '}' +

    'function ss(msg,pct,det){' +
    '  document.getElementById("st").textContent=msg;' +
    '  if(pct!==undefined)document.getElementById("bar").style.width=pct+"%";' +
    '  if(det!==undefined)document.getElementById("dt").textContent=det||"\\u00a0";' +
    '}' +

    'async function run(){' +
    '  try{' +
    '    ss("Decoding PDF...",5);' +
    '    var pdfBytes=b64ToArr(PDF_B64);' +
    '    ss("Decompressing streams...",15);' +
    '    var streams=extractStreams(pdfBytes);' +
    '    ss("Parsing "+streams.length+" pages...",25,"Decompressed "+streams.length+" streams");' +
    '    var allLines=[];' +
    '    for(var i=0;i<streams.length;i++){' +
    '      var items=parseStream(streams[i]);' +
    '      allLines=allLines.concat(items);' +
    '      ss("Parsed page "+(i+1)+" of "+streams.length+"...",25+Math.round((i+1)/streams.length*55),"Found "+allLines.length+" text items");' +
    '    }' +
    '    ss("Extracting fund records...",82);' +
    '    var records=parseRecords(allLines);' +
    '    ss("Saving "+records.length+" records...",90,"Please wait...");' +
    '    google.script.run' +
    '      .withSuccessHandler(function(r){' +
    '        if(r&&r.success){ss("\\u2705 Done! "+r.count+" NAVs cached.",100,"Closing...");setTimeout(function(){google.script.host.close();},2000);}' +
    '        else{ss("\\u274c Save error: "+(r?r.error:"null"),100);}' +
    '      })' +
    '      .withFailureHandler(function(e){ss("\\u274c GAS error: "+e.message,100);})' +
    '      .saveFimmRecordsFromDialog(JSON.stringify(records));' +
    '  }catch(e){ss("\\u274c "+e.message,100);}' +
    '}' +
    'window.onload=function(){ss("Starting...",2);setTimeout(run,100);};';

  const html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:Arial,sans-serif;padding:18px 20px;margin:0;color:#333;}' +
    'h3{margin:0 0 10px;color:#1565c0;font-size:15px;}' +
    '#st{font-size:13px;margin:6px 0;min-height:18px;}' +
    '#prog{width:100%;height:7px;background:#e0e0e0;border-radius:4px;overflow:hidden;margin:10px 0;}' +
    '#bar{height:100%;width:2%;background:#1565c0;transition:width 0.3s;}' +
    '#dt{font-size:11px;color:#888;}</style>' +
    // Inline pako (6.9KB) — no external script loading
    '<script>' + pakoCode + '</script>' +
    '</head><body>' +
    '<h3>📋 Refreshing FIMM NAV Cache</h3>' +
    '<div id="st">Loading...</div>' +
    '<div id="prog"><div id="bar"></div></div>' +
    '<div id="dt">&nbsp;</div>' +
    '<script>' + parserJs + '</script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html).setWidth(440).setHeight(180);
}

function dailyFimmNavJob_() {
  // For the daily trigger we still need a server-side approach.
  // Use Drive OCR as a best-effort fallback when no UI is available.
  try {
    const count = fetchAndParseFimmPdf_server_();
    Logger.log('FIMM daily NAV update: ' + count + ' records');
  } catch(e) {
    Logger.log('FIMM NAV daily job error: ' + e.message);
  }
}

/**
 * Server-side fallback for scheduled/trigger-based refresh (no browser available).
 * Uses Drive OCR PDF→Doc→text extraction.
 */
function fetchAndParseFimmPdf_server_() {
  const pdfRes = UrlFetchApp.fetch(FIMM_PDF_URL, {
    method: 'get', muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript/1.0)' }
  });
  if (pdfRes.getResponseCode() !== 200) {
    throw new Error('FIMM PDF fetch failed (HTTP ' + pdfRes.getResponseCode() + ')');
  }
  const pdfBlob = pdfRes.getBlob().setContentType('application/pdf').setName('_fimm_nav_ocr_.pdf');

  var tempDocId = null;
  const tempFile = DriveApp.createFile(pdfBlob);
  const tempId   = tempFile.getId();
  const token    = ScriptApp.getOAuthToken();

  try {
    const copyRes = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + tempId + '/copy',
      {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({ name: '_fimm_nav_doc_', mimeType: 'application/vnd.google-apps.document' }),
        muteHttpExceptions: true
      }
    );
    tempFile.setTrashed(true);

    if (copyRes.getResponseCode() !== 200) {
      throw new Error('Drive OCR copy failed: ' + copyRes.getContentText().substring(0, 200));
    }
    tempDocId = JSON.parse(copyRes.getContentText()).id;
    Utilities.sleep(5000);

    const txtRes = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + tempDocId + '/export?mimeType=text/plain',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );

    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + tempDocId,
      { method: 'delete', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    tempDocId = null;

    if (txtRes.getResponseCode() !== 200) throw new Error('Doc export failed: HTTP ' + txtRes.getResponseCode());

    const records = parseFimmOcrText_(txtRes.getContentText());
    if (records.length > 0) writeFimmCache_(records);
    return records.length;

  } finally {
    try { if (tempDocId) UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + tempDocId,
      { method: 'delete', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }); } catch(_) {}
    try { if (tempId) DriveApp.getFileById(tempId).setTrashed(true); } catch(_) {}
  }
}

/**
 * Parses flat OCR text from a Drive-exported PDF→Doc→text conversion.
 * State machine: accumulates name fragments until hitting a NAV number.
 */
function parseFimmOcrText_(text) {
  const NAV_RE         = /^\d+\.\d{3,6}$/;
  const DATE_RE        = /^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i;
  const COMPANY_HDR_RE = /\b(SDN BHD|BERHAD|ASSET MANAGEMENT|FORMERLY KNOWN|PTE LTD|INC\.|CORPORATION|INVESTMENTS? LTD)\b/i;
  const JUNK_RE        = /(\d+\.\d+%|p\.a\.|preceding business|Daily Dividend|Manager.s Price|\bSUSP\b|BOS \d)/i;
  const ANNOT_RE       = /\s*\([a-uw-z]\)\s*/gi;

  const lines   = text.split(/\r?\n/).map(function(l) { return l.trim(); });
  const records = [];
  const seen    = {};
  var nameParts = [];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line) continue;
    if (DATE_RE.test(line) || /^(NAV|RM|-|\d{1,3}|Page \d)$/.test(line)) { nameParts = []; continue; }
    if (/^(NA|N\/A)$/i.test(line)) { nameParts = []; continue; }

    if (NAV_RE.test(line)) {
      if (nameParts.length > 0) {
        var name = nameParts.join(' ').replace(ANNOT_RE, ' ').replace(/\s{2,}/g, ' ').trim().replace(/\s*\([a-z]\)$/, '').trim();
        var nav = parseFloat(line);
        if (name.length >= 4 && name.length <= 70 && /^[A-Z]/.test(name) &&
            !COMPANY_HDR_RE.test(name) && !JUNK_RE.test(name) && nav > 0 && !seen[name]) {
          seen[name] = true;
          records.push([name, nav]);
        }
      }
      nameParts = [];
      continue;
    }

    if (COMPANY_HDR_RE.test(line) || JUNK_RE.test(line)) { nameParts = []; continue; }
    var stripped = line.replace(ANNOT_RE, '').trim();
    if (!stripped || /^\([a-z]\)$/.test(stripped)) continue;
    if (/^[A-Z]/.test(stripped) || nameParts.length > 0) nameParts.push(stripped);
  }

  return records;
}

function writeFimmCache_(records) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  // Get or create cache sheet
  var sheet = ss.getSheetByName(FIMM_NAV_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FIMM_NAV_SHEET);
    // Move to near end (don't clutter the front)
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(ss.getNumSheets());
  }

  sheet.clearContents();
  sheet.clearFormats();

  // ── Header row ───────────────────────────────────────────────
  sheet.setRowHeight(1, 36);
  sheet.getRange(1, 1, 1, 5).setValues([['Fund Code', 'Fund Name', 'NAV', 'Updated', 'Source']])
    .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold')
    .setFontSize(11).setVerticalAlignment('middle');

  // ── Data ─────────────────────────────────────────────────────
  if (records.length === 0) {
    sheet.getRange(2, 1).setValue('No data parsed — check PDF format or try refreshing again.');
    return;
  }

  // Batch write for performance
  const rows = records.map(r => [r[0], r[1], r[2], today, 'FIMM']);
  sheet.getRange(2, 1, rows.length, 5).setValues(rows);

  // Format NAV column
  sheet.getRange(2, 3, rows.length, 1).setNumberFormat('0.0000').setHorizontalAlignment('right');
  sheet.getRange(2, 1, rows.length, 1).setFontWeight('bold').setFontColor('#1565c0');

  // Alternating row colors
  for (var i = 0; i < rows.length; i++) {
    sheet.getRange(i + 2, 1, 1, 5).setBackground(i % 2 === 0 ? '#e3f2fd' : '#ffffff');
    sheet.setRowHeight(i + 2, 22);
  }

  // ── Column widths ────────────────────────────────────────────
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 380);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 80);

  // ── Subtitle row at top showing count + date ─────────────────
  sheet.insertRowBefore(1);
  sheet.setRowHeight(1, 30);
  sheet.getRange(1, 1, 1, 5).merge()
    .setValue('📋 FIMM NAV Cache  ·  ' + records.length + ' funds  ·  Last updated: ' + today)
    .setBackground('#0d47a1').setFontColor('#ffffff').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
}


// ============================================================
//  LOANS & DEBTS TRACKER
//  Sheet: 🤝 Loans & Debts
//  Tracks money lent out and money borrowed.
//  NOT included in net worth dashboard — shown as a separate
//  info panel below the Grand Total.
// ============================================================

// ── COLUMN MAP ────────────────────────────────────────────────
// A=1  Type (LENT / BORROWED)
// B=2  Person / Institution
// C=3  Date
// D=4  Amount (original)
// E=5  Due Date
// F=6  Repaid (running total of repayments)
// G=7  Status  (Open / Partial / Repaid)
// H=8  Notes

// ── CREATE SHEET ──────────────────────────────────────────────
function createLoansSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOANS_SHEET_NAME);
  if (sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + LOANS_SHEET_NAME + '" already exists.');
    sheet.activate();
    return;
  }
  sheet = ss.insertSheet(LOANS_SHEET_NAME, 1);
  buildLoansSheet_(sheet);
  sheet.activate();
}

function buildLoansSheet_(sheet) {
  compactSheet_(sheet, 200, 8);

  // Column widths
  sheet.setColumnWidth(1, 90);   // Type
  sheet.setColumnWidth(2, 160);  // Person
  sheet.setColumnWidth(3, 90);   // Date
  sheet.setColumnWidth(4, 130);  // Amount
  sheet.setColumnWidth(5, 90);   // Due Date
  sheet.setColumnWidth(6, 130);  // Repaid
  sheet.setColumnWidth(7, 80);   // Status
  sheet.setColumnWidth(8, 200);  // Notes

  const myrFmt = '"RM "#,##0.00';

  // Row 1 — banner
  sheet.setRowHeight(1, 48);
  sheet.getRange(1, 1, 1, 8).merge()
    .setValue('🤝  LOANS & DEBTS')
    .setBackground('#4a148c').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Row 2 — column headers
  sheet.setRowHeight(2, 30);
  const headers = ['Type', 'Person / Institution', 'Date', 'Amount (RM)', 'Due Date', 'Repaid (RM)', 'Status', 'Notes'];
  const hdrBg   = '#7b1fa2';
  headers.forEach((h, i) => {
    sheet.getRange(2, i + 1)
      .setValue(h).setBackground(hdrBg).setFontColor('#fff')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment(i >= 3 && i <= 5 ? 'right' : 'center')
      .setVerticalAlignment('middle');
  });

  sheet.setFrozenRows(2);
  sheet.setHiddenGridlines(true);
}

// ── ADD LOAN DIALOG ───────────────────────────────────────────
function showAddLoanDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(LOANS_SHEET_NAME)) {
    createLoansSheet();
  }

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#4a148c}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select,textarea{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    'textarea{resize:vertical;height:56px}' +
    '.type-row{display:flex;gap:10px;margin:8px 0}' +
    '.type-btn{flex:1;padding:12px;border:2px solid #dadce0;border-radius:8px;background:#fff;cursor:pointer;font-size:13px;font-weight:600;text-align:center}' +
    '.type-btn.active-lent{border-color:#6a1b9a;background:#f3e5f5;color:#6a1b9a}' +
    '.type-btn.active-borrowed{border-color:#880e4f;background:#fce4ec;color:#880e4f}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#4a148c;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>🤝 Add Lent / Borrowed</h2>' +
    '<label>Type</label>' +
    '<div class="type-row">' +
    '<div class="type-btn active-lent" id="btn-LENT" onclick="setType(\'LENT\')">💸 I LENT money</div>' +
    '<div class="type-btn" id="btn-BORROWED" onclick="setType(\'BORROWED\')">🏦 I BORROWED money</div>' +
    '</div>' +
    '<label id="personLabel">Person / Who you lent to</label>' +
    '<input type="text" id="person" placeholder="e.g. Ali, CIMB Bank" />' +
    '<label>Amount (RM)</label>' +
    '<input type="number" id="amount" placeholder="0.00" step="0.01" min="0.01" />' +
    '<label>Date</label>' +
    '<input type="date" id="date" />' +
    '<label>Due Date <span style="font-weight:400;color:#9aa0a6">(optional)</span></label>' +
    '<input type="date" id="dueDate" />' +
    '<label>Notes <span style="font-weight:400;color:#9aa0a6">(optional)</span></label>' +
    '<textarea id="notes" placeholder="e.g. For house renovation, Personal loan ref #1234"></textarea>' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="save()">Save</button>' +
    '</div>' +
    '<script>' +
    'var selType="LENT";' +
    'var today=new Date().toISOString().split("T")[0];' +
    'document.getElementById("date").value=today;' +
    'function setType(t){' +
    '  selType=t;' +
    '  ["LENT","BORROWED"].forEach(function(x){' +
    '    document.getElementById("btn-"+x).className="type-btn"+(x===t?" active-"+x.toLowerCase():"");' +
    '  });' +
    '  document.getElementById("personLabel").textContent=t==="LENT"?"Person / Who you lent to":"Person / Institution you borrowed from";' +
    '}' +
    'function save(){' +
    '  var person=document.getElementById("person").value.trim();' +
    '  var amt=parseFloat(document.getElementById("amount").value);' +
    '  var date=document.getElementById("date").value;' +
    '  var due=document.getElementById("dueDate").value;' +
    '  var notes=document.getElementById("notes").value.trim();' +
    '  if(!person){alert("Enter a person or institution.");return;}' +
    '  if(!amt||amt<=0){alert("Enter a valid amount.");return;}' +
    '  if(!date){alert("Enter a date.");return;}' +
    '  document.querySelector(".btn-primary").disabled=true;' +
    '  document.querySelector(".btn-primary").textContent="Saving...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close();})' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);document.querySelector(".btn-primary").disabled=false;document.querySelector(".btn-primary").textContent="Save";})' +
    '    .saveLoan(selType,person,amt,date,due,notes);' +
    '}' +
    '</script></body></html>'
  ).setWidth(460).setHeight(540).setTitle('Add Lent / Borrowed');
  SpreadsheetApp.getUi().showModalDialog(html, 'Add Lent / Borrowed');
}

// ── SAVE LOAN ────────────────────────────────────────────────
function saveLoan(type, person, amount, dateStr, dueDateStr, notes) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(LOANS_SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(LOANS_SHEET_NAME, 1); buildLoansSheet_(sheet); }

  const myrFmt  = '"RM "#,##0.00';
  const dateFmt = 'dd/MM/yyyy';

  const lastRow = Math.max(sheet.getLastRow(), 2) + 1;
  const isLent  = type === 'LENT';
  const bg      = isLent ? '#f3e5f5' : '#fce4ec';
  const altBg   = isLent ? '#fce4ff' : '#ffe4ec';
  const rowBg   = lastRow % 2 === 0 ? bg : altBg;

  sheet.setRowHeight(lastRow, 32);
  sheet.getRange(lastRow, 1, 1, 8).setBackground(rowBg);

  // Col A – Type
  sheet.getRange(lastRow, 1).setValue(type)
    .setFontWeight('bold').setFontColor(isLent ? '#6a1b9a' : '#880e4f')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Col B – Person
  sheet.getRange(lastRow, 2).setValue(person)
    .setFontSize(11).setVerticalAlignment('middle');

  // Col C – Date
  const dateParts = dateStr.split('-');
  const dateObj   = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
  sheet.getRange(lastRow, 3).setValue(dateObj).setNumberFormat(dateFmt)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Col D – Amount
  sheet.getRange(lastRow, 4).setValue(amount).setNumberFormat(myrFmt)
    .setFontWeight('bold').setHorizontalAlignment('right').setVerticalAlignment('middle');

  // Col E – Due Date
  if (dueDateStr) {
    const dp  = dueDateStr.split('-');
    const due = new Date(parseInt(dp[0]), parseInt(dp[1]) - 1, parseInt(dp[2]));
    sheet.getRange(lastRow, 5).setValue(due).setNumberFormat(dateFmt)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  } else {
    sheet.getRange(lastRow, 5).setValue('—').setHorizontalAlignment('center').setVerticalAlignment('middle').setFontColor('#9aa0a6');
  }

  // Col F – Repaid (starts at 0)
  sheet.getRange(lastRow, 6).setValue(0).setNumberFormat(myrFmt)
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setFontColor('#9aa0a6');

  // Col G – Status
  sheet.getRange(lastRow, 7).setValue('Open')
    .setFontColor('#e65100').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Col H – Notes
  sheet.getRange(lastRow, 8).setValue(notes || '').setVerticalAlignment('middle');
}

// ── REPAYMENT DIALOG ─────────────────────────────────────────
function showRepaymentDialog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOANS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 3) {
    SpreadsheetApp.getUi().alert('No loans/debts found. Add one first via "Add Lent / Borrowed".');
    return;
  }

  // Read open loans
  const lastRow  = sheet.getLastRow();
  const data     = sheet.getRange(3, 1, lastRow - 2, 8).getValues();
  const openLoans = [];
  data.forEach(function(r, i) {
    const type      = r[0] ? r[0].toString() : '';
    const person    = r[1] ? r[1].toString() : '';
    const amount    = typeof r[3] === 'number' ? r[3] : 0;
    const repaid    = typeof r[5] === 'number' ? r[5] : 0;
    const status    = r[6] ? r[6].toString().toUpperCase() : '';
    const outstanding = Math.max(0, amount - repaid);
    if (status !== 'REPAID' && type && person) {
      openLoans.push({
        row:         i + 3,  // 1-indexed sheet row
        type:        type,
        person:      person,
        amount:      amount,
        repaid:      repaid,
        outstanding: outstanding,
        label:       type + ' — ' + person + '  (outstanding: RM ' + outstanding.toFixed(2) + ')'
      });
    }
  });

  if (!openLoans.length) {
    SpreadsheetApp.getUi().alert('All loans/debts are already fully repaid!');
    return;
  }

  const loansJson = JSON.stringify(openLoans);

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><style>' +
    'body{font-family:Google Sans,Arial,sans-serif;margin:0;padding:16px;background:#f8f9fa;color:#202124;font-size:13px}' +
    'h2{margin:0 0 14px;font-size:15px;color:#4a148c}' +
    'label{display:block;font-weight:600;margin:10px 0 3px;color:#3c4043;font-size:12px}' +
    'input,select{width:100%;box-sizing:border-box;padding:7px 10px;border:1px solid #dadce0;border-radius:6px;font-size:13px;background:#fff}' +
    '.info{background:#ede7f6;border-radius:6px;padding:8px 12px;font-size:12px;color:#4a148c;margin:6px 0}' +
    '.btn-row{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}' +
    '.btn{padding:8px 20px;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600}' +
    '.btn-primary{background:#4a148c;color:#fff}.btn-cancel{background:#f1f3f4;color:#3c4043}' +
    '</style></head><body>' +
    '<h2>💰 Record Repayment</h2>' +
    '<label>Loan / Debt</label>' +
    '<select id="loanSel" onchange="onLoanChange()">' +
    openLoans.map(function(l) { return '<option value="' + l.row + '">' + l.label + '</option>'; }).join('') +
    '</select>' +
    '<div class="info" id="loanInfo"></div>' +
    '<label>Repayment Amount (RM)</label>' +
    '<input type="number" id="repayAmt" placeholder="0.00" step="0.01" min="0.01" />' +
    '<label>Date</label>' +
    '<input type="date" id="repayDate" />' +
    '<div class="btn-row">' +
    '<button class="btn btn-cancel" onclick="google.script.host.close()">Cancel</button>' +
    '<button class="btn btn-primary" onclick="save()">Save Repayment</button>' +
    '</div>' +
    '<script>' +
    'var LOANS=' + loansJson + ';' +
    'document.getElementById("repayDate").value=new Date().toISOString().split("T")[0];' +
    'function getLoan(){return LOANS.find(function(l){return l.row==document.getElementById("loanSel").value;});}' +
    'function onLoanChange(){' +
    '  var l=getLoan();if(!l)return;' +
    '  document.getElementById("loanInfo").innerHTML=' +
    '    "<strong>"+l.type+"</strong> · "+l.person+"<br>"' +
    '    +"Original: <strong>RM "+l.amount.toFixed(2)+"</strong> · Already repaid: <strong>RM "+l.repaid.toFixed(2)+"</strong> · Outstanding: <strong>RM "+l.outstanding.toFixed(2)+"</strong>";' +
    '  document.getElementById("repayAmt").max=l.outstanding;' +
    '  document.getElementById("repayAmt").value="";' +
    '}' +
    'onLoanChange();' +
    'function save(){' +
    '  var l=getLoan();' +
    '  var amt=parseFloat(document.getElementById("repayAmt").value);' +
    '  var date=document.getElementById("repayDate").value;' +
    '  if(!amt||amt<=0){alert("Enter a valid repayment amount.");return;}' +
    '  if(amt>l.outstanding+0.001){alert("Repayment cannot exceed outstanding balance (RM "+l.outstanding.toFixed(2)+").");return;}' +
    '  if(!date){alert("Enter a date.");return;}' +
    '  document.querySelector(".btn-primary").disabled=true;' +
    '  document.querySelector(".btn-primary").textContent="Saving...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close();})' +
    '    .withFailureHandler(function(e){alert("Error: "+e.message);document.querySelector(".btn-primary").disabled=false;document.querySelector(".btn-primary").textContent="Save Repayment";})' +
    '    .saveRepayment(l.row,amt,date);' +
    '}' +
    '</script></body></html>'
  ).setWidth(460).setHeight(420).setTitle('Record Repayment');
  SpreadsheetApp.getUi().showModalDialog(html, 'Record Repayment');
}

// ── SAVE REPAYMENT ───────────────────────────────────────────
function saveRepayment(loanRow, repayAmount, dateStr) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LOANS_SHEET_NAME);
  if (!sheet) throw new Error('Loans sheet not found.');

  const myrFmt = '"RM "#,##0.00';
  const amount  = sheet.getRange(loanRow, 4).getValue();
  const repaid  = sheet.getRange(loanRow, 6).getValue() || 0;
  const newRepaid = repaid + repayAmount;

  sheet.getRange(loanRow, 6).setValue(newRepaid).setNumberFormat(myrFmt)
    .setFontColor(newRepaid >= amount ? '#0f9d58' : '#e65100')
    .setFontWeight('bold');

  // Update status
  let status = 'Partial';
  if (newRepaid >= amount - 0.001) {
    status = 'Repaid';
    // Green out the row
    const rowBg = '#e8f5e9';
    sheet.getRange(loanRow, 1, 1, 8).setBackground(rowBg);
  }
  sheet.getRange(loanRow, 7).setValue(status)
    .setFontColor(status === 'Repaid' ? '#0f9d58' : '#f57c00')
    .setFontWeight('bold');

  // Append a repayment note to Notes column
  const dateParts = dateStr.split('-');
  const dateObj   = new Date(parseInt(dateParts[0]), parseInt(dateParts[1]) - 1, parseInt(dateParts[2]));
  const dateFormatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const existingNote  = sheet.getRange(loanRow, 8).getValue() || '';
  const repayNote     = '[Repaid RM ' + repayAmount.toFixed(2) + ' on ' + dateFormatted + ']';
  sheet.getRange(loanRow, 8).setValue(existingNote ? existingNote + '  ' + repayNote : repayNote);
}