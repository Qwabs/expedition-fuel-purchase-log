/**
 * Fuel Purchase Log — year logs, Add Purchase sidebar, mobile Form and audit.
 *
 * Install once after importing the workbook into Google Sheets:
 *   1. Extensions > Apps Script
 *   2. Paste Code.gs; add an HTML file named Sidebar with Sidebar.html
 *   3. Run setupAudit() and setupMobileForm() from the designated owner Google account
 */

var AUDIT_SHEET = 'Audit Log';
var LOG_RE = /^\d{4} Log$/;
var LOG_ROWS = 5000;
var LOG_LAST_ROW = 4 + LOG_ROWS;
var ENTRY_FIRST_ROW = 5;

var COL_DATE = 1;        // A
var COL_PURCHASER = 3;   // C
var COL_AMOUNT = 4;      // D
var COL_CAR = 5;         // E
var COL_ENTERED_BY = 6;  // F
var COL_ENTERED_AT = 7;  // G

var AUDIT_HEADER_ROW = 4;
var AUDIT_WIDTH = 12;
var ENTRY_COLUMNS = [COL_DATE, COL_PURCHASER, COL_AMOUNT, COL_CAR];
var FIELD_NAMES = {
  1: 'Date',
  3: 'Purchaser',
  4: 'Amount',
  5: 'Car Number'
};

var MOBILE_FORM_PROPERTY = 'MOBILE_FORM_ID';
var MASTER_SHEET_PROPERTY = 'MASTER_SHEET_ID';
var Q_DATE = 'Purchase date';
var Q_PURCHASER = 'Purchaser';
var Q_CAR = 'Car Number';
var Q_AMOUNT = 'Amount (GH₵)';


function alert_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    console.log(message);
  }
}


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fuel Log')
    .addItem('Add Purchase', 'showAddPurchase')
    .addItem('Open current year log', 'openCurrentLog')
    .addSeparator()
    .addItem('Mobile form link', 'showMobileFormLink')
    .addItem('Create / refresh mobile form', 'setupMobileForm')
    .addItem('Sync mobile form purchasers', 'syncMobileFormPurchasers')
    .addSeparator()
    .addItem('Open Audit Log', 'openAuditLog')
    .addItem('Install audit trigger', 'setupAudit')
    .addToUi();
}


function showAddPurchase() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Add Purchase')
    .setWidth(330);
  SpreadsheetApp.getUi().showSidebar(html);
}


function getSidebarData() {
  var ss = SpreadsheetApp.getActive();
  return {
    purchasers: getPurchasers_(ss),
    years: getLogYears_(ss),
    today: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd')
  };
}


function appendPurchase(form) {
  try {
    if (!form || !form.date) {
      return { ok: false, message: 'Choose a purchase date.' };
    }
    var date = parseDate_(form.date);
    if (!date) {
      return { ok: false, message: 'Enter a valid date.' };
    }

    var ss = SpreadsheetApp.getActive();
    var purchaser = String(form.purchaser || '').trim();
    if (!purchaser || getPurchasers_(ss).indexOf(purchaser) === -1) {
      return { ok: false, message: 'Choose a purchaser from the Settings list.' };
    }

    var amount = Number(form.amount);
    if (!isFinite(amount) || amount <= 0) {
      return { ok: false, message: 'Enter an amount greater than zero.' };
    }

    var car = normalizeCar_(form.carNumber);
    if (!car) {
      return { ok: false, message: 'Car Number is required.' };
    }

    var email = Session.getActiveUser().getEmail() || 'unknown';
    var row = appendPurchaseRecord_(
      ss, date, purchaser, amount, car, email, 'Added via sidebar', new Date()
    );

    return {
      ok: true,
      row: row,
      message: 'Purchase added to ' + yearLogName_(date) + ', row ' + row + '.'
    };
  } catch (err) {
    return { ok: false, message: err.message || String(err) };
  }
}


/**
 * Audit-first save shared by the sidebar and the mobile Form.
 * The audit row is appended first; if the purchase write fails, it is removed.
 */
function appendPurchaseRecord_(ss, date, purchaser, amount, car, email, action, enteredAt) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Another entry is saving. Please try again.');
  }

  var audit = null;
  var auditStart = null;
  try {
    var log = ss.getSheetByName(yearLogName_(date));
    if (!log) {
      throw new Error('There is no ' + date.getFullYear() + ' Log.');
    }
    var row = findAppendRow_(log);
    audit = ensureAuditSheet_(ss);

    auditStart = appendAuditRows_(audit, [[
      new Date(), email, log.getName(), row, date, purchaser, amount, car,
      'Entry', '', purchaser + ' · ' + amount + ' · ' + car, action
    ]]);

    try {
      log.getRange(row, COL_DATE).setValue(date);
      log.getRange(row, COL_PURCHASER).setValue(purchaser);
      log.getRange(row, COL_AMOUNT).setValue(amount);
      log.getRange(row, COL_CAR).setValue(car);
      log.getRange(row, COL_ENTERED_BY).setValue(email);
      log.getRange(row, COL_ENTERED_AT).setValue(enteredAt);

      if (log.getRange(row, 2).getFormula() === '') {
        log.getRange(row, 2).setFormula(
          '=IF(A' + row + '="","",TEXT(A' + row + ',"dddd"))'
        );
      }
      log.getRange(row, COL_DATE).setNumberFormat('dd mmm yyyy');
      log.getRange(row, COL_AMOUNT).setNumberFormat('"GH₵"#,##0.00');
      log.getRange(row, COL_ENTERED_AT).setNumberFormat('dd mmm yyyy hh:mm');
    } catch (writeErr) {
      var cleared = false;
      try {
        log.getRange(row, 1, 1, 7).clearContent();
        log.getRange(row, 2).setFormula(
          '=IF(A' + row + '="","",TEXT(A' + row + ',"dddd"))'
        );
        cleared = true;
      } catch (clearErr) {
        // leave the audit row in place rather than an unlogged partial row
      }

      if (cleared) {
        try {
          audit.deleteRow(auditStart);
        } catch (rollbackErr) {
          throw new Error(
            'Purchase write failed and the audit row could not be rolled back: ' +
            writeErr.message
          );
        }
      }

      throw new Error(
        writeErr.message + (cleared ? '' : ' (partial row retained with audit record)')
      );
    }

    SpreadsheetApp.flush();
    return row;
  } finally {
    lock.releaseLock();
  }
}


function auditOnEdit(e) {
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  if (!LOG_RE.test(sheetName)) return;
  if (range.getLastRow() < ENTRY_FIRST_ROW || range.getRow() > LOG_LAST_ROW) return;

  var firstCol = range.getColumn();
  var lastCol = range.getLastColumn();
  if (lastCol < COL_DATE || firstCol > COL_CAR) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    var ss = sheet.getParent();
    var email = getEditorEmail_(e);
    var now = new Date();
    var audit = ensureAuditSheet_(ss);
    var values = range.getValues();
    var singleCell = range.getNumCells() === 1;
    var oldValue = singleCell ? e.oldValue : '';
    var auditRows = [];

    var firstRow = Math.max(range.getRow(), ENTRY_FIRST_ROW);
    var lastRow = Math.min(range.getLastRow(), LOG_LAST_ROW);

    for (var row = firstRow; row <= lastRow; row++) {
      var rowOffset = row - range.getRow();
      var rowValues = values[rowOffset];
      var changed = false;

      for (var c = 0; c < ENTRY_COLUMNS.length; c++) {
        var col = ENTRY_COLUMNS[c];
        var colOffset = col - firstCol;
        if (colOffset < 0 || colOffset >= range.getNumColumns()) continue;

        var newValue = rowValues[colOffset];
        if (singleCell && String(oldValue) === String(newValue)) continue;

        auditRows.push([
          now,
          email,
          sheetName,
          row,
          sheet.getRange(row, COL_DATE).getValue(),
          sheet.getRange(row, COL_PURCHASER).getValue(),
          sheet.getRange(row, COL_AMOUNT).getValue(),
          sheet.getRange(row, COL_CAR).getValue(),
          FIELD_NAMES[col],
          stringify_(oldValue),
          stringify_(newValue),
          singleCell ? action_(oldValue, newValue) : 'Pasted'
        ]);
        changed = true;
      }

      if (changed) {
        updateStamp_(sheet, row, email, now);
      }
    }

    if (auditRows.length) {
      appendAuditRows_(audit, auditRows);
    }
  } finally {
    lock.releaseLock();
  }
}


function updateStamp_(sheet, row, email, now) {
  var purchaser = sheet.getRange(row, COL_PURCHASER).getValue();
  var amount = sheet.getRange(row, COL_AMOUNT).getValue();
  var stamp = sheet.getRange(row, COL_ENTERED_BY, 1, 2);
  var current = stamp.getValues()[0];
  var hasEntry = !isBlank_(purchaser) || !isBlank_(amount);

  if (hasEntry && isBlank_(current[0])) {
    stamp.setValues([[email, now]]);
    stamp.setNumberFormats([['@', 'dd mmm yyyy hh:mm']]);
  } else if (!hasEntry && !isBlank_(current[0])) {
    stamp.clearContent();
  }
}


function findAppendRow_(log) {
  var values = log.getRange(ENTRY_FIRST_ROW, COL_DATE, LOG_ROWS, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (!isBlank_(values[i][0])) {
      var next = ENTRY_FIRST_ROW + i + 1;
      if (next > LOG_LAST_ROW) {
        throw new Error('The ' + log.getName() + ' is full (' + LOG_ROWS + ' purchases).');
      }
      return next;
    }
  }
  return ENTRY_FIRST_ROW;
}


function yearLogName_(date) {
  return date.getFullYear() + ' Log';
}


function normalizeCar_(value) {
  return String(value === null || value === undefined ? '' : value)
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}


function getPurchasers_(ss) {
  var settings = ss.getSheetByName('Settings');
  if (!settings) return [];
  return settings.getRange(15, 1, 50, 1).getValues()
    .map(function (row) { return String(row[0]).trim(); })
    .filter(function (value) { return value !== ''; });
}


function getLogYears_(ss) {
  var years = [];
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var match = sheets[i].getName().match(/^(\d{4}) Log$/);
    if (match) years.push(Number(match[1]));
  }
  return years.sort();
}


function parseDate_(value) {
  var parts = String(value).split('-');
  if (parts.length !== 3) return null;
  var year = Number(parts[0]);
  var month = Number(parts[1]);
  var day = Number(parts[2]);
  if (!year || !month || !day) return null;
  var date = new Date(year, month - 1, day, 12, 0, 0);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 ||
      date.getDate() !== day) {
    return null;
  }
  return date;
}


function getEditorEmail_(e) {
  try {
    if (e && e.user && typeof e.user.getEmail === 'function') {
      var email = e.user.getEmail();
      if (email) return email;
    }
  } catch (err) {
    // fall through to Session
  }
  try {
    var active = Session.getActiveUser().getEmail();
    if (active) return active;
  } catch (err2) {
    // fall through
  }
  return 'unknown';
}


function ensureAuditSheet_(ss) {
  var sheet = ss.getSheetByName(AUDIT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(AUDIT_SHEET);
  }

  var headers = [
    'Timestamp', 'User', 'Sheet', 'Row', 'Purchase date', 'Purchaser',
    'Amount (GH₵)', 'Car Number', 'Field', 'Old value',
    'New value', 'Action'
  ];
  var current = sheet.getRange(AUDIT_HEADER_ROW, 1, 1, AUDIT_WIDTH).getValues()[0];
  if (current.join('') === '') {
    sheet.getRange(AUDIT_HEADER_ROW, 1, 1, AUDIT_WIDTH).setValues([headers]);
    sheet.getRange(AUDIT_HEADER_ROW, 1, 1, AUDIT_WIDTH)
      .setBackground('#1E4D3D')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(10);
    sheet.setFrozenRows(AUDIT_HEADER_ROW);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange('E:E').setNumberFormat('d mmm yyyy');
    sheet.getRange('G:G').setNumberFormat('"GH₵"#,##0.00');
  }
  return sheet;
}


function appendAuditRows_(audit, rows) {
  var start = Math.max(audit.getLastRow() + 1, AUDIT_HEADER_ROW + 1);
  var block = audit.getRange(start, 1, rows.length, AUDIT_WIDTH);
  block.setValues(rows);
  block.setFontSize(9);
  block.setBorder(true, true, true, true, true, true, '#D6DDD8',
                   SpreadsheetApp.BorderStyle.SOLID);
  audit.getRange(start, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  audit.getRange(start, 5, rows.length, 1).setNumberFormat('d mmm yyyy');
  audit.getRange(start, 7, rows.length, 1).setNumberFormat('"GH₵"#,##0.00');
  return start;
}


function action_(oldValue, newValue) {
  var wasBlank = isBlank_(oldValue);
  var nowBlank = isBlank_(newValue);
  if (wasBlank && !nowBlank) return 'Added';
  if (!wasBlank && nowBlank) return 'Cleared';
  return 'Edited';
}


function isBlank_(value) {
  if (value === '' || value === null || value === undefined) return true;
  if (typeof value === 'string' && value.trim() === '') return true;
  return false;
}


function stringify_(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(),
                                'yyyy-MM-dd HH:mm:ss');
  }
  return String(value);
}


// ------------------------------------------------------------------ mobile form
function setupMobileForm() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties()
    .setProperty(MASTER_SHEET_PROPERTY, ss.getId());
  var form = getOrCreateMobileForm_();
  var purchasers = getPurchasers_(ss);
  console.log('setupMobileForm: form ready; purchasers=' + purchasers.length);

  form.setTitle('Fuel Purchase — ' + ss.getName());
  form.setDescription(
    'One purchase per submission. Multiple purchases on the same day are fine. ' +
    'Car Number is required. Your Google account and submission time are recorded.'
  );
  form.setConfirmationMessage(
    'Submission received. Valid entries appear in the purchase log.'
  );
  try { form.setAcceptingResponses(true); } catch (err) {}
  try { form.setCollectEmail(true); } catch (err) {}
  try { form.setRequireLogin(true); } catch (err) {}

  removeReceiptItems_(form);

  var dateItem = getOrAddItem_(form, Q_DATE, 'date');
  dateItem.setTitle(Q_DATE);
  dateItem.setRequired(true);
  setDateIncludesYear_(dateItem);

  var purchaserItem = getOrAddItem_(form, Q_PURCHASER, 'list');
  purchaserItem.setTitle(Q_PURCHASER);
  purchaserItem.setRequired(true);
  if (purchasers.length) {
    setPurchaserChoices_(purchaserItem, purchasers);
  }

  var carItem = getOrAddItem_(form, Q_CAR, 'text');
  carItem.setTitle(Q_CAR);
  carItem.setRequired(true);
  try { carItem.setHelpText('Vehicle plate, for example GT 1234-24.'); } catch (err) {}

  var amountItem = getOrAddItem_(form, Q_AMOUNT, 'text');
  amountItem.setTitle(Q_AMOUNT);
  amountItem.setRequired(true);
  try { amountItem.setHelpText('Enter digits only, for example 420.50.'); } catch (err) {}

  // Responses stay in the Form's Responses tab. The form-submit trigger writes
  // valid submissions into the year log; linking the Form to the 73-sheet
  // workbook is unnecessary and can time out.
  console.log('setupMobileForm: questions configured');

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processFormResponse') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processFormResponse').forForm(form).onFormSubmit().create();
  console.log('setupMobileForm: trigger installed');

  PropertiesService.getScriptProperties()
    .setProperty(MOBILE_FORM_PROPERTY, form.getId());

  var url = form.getPublishedUrl();
  writeMobileFormLink_(ss, url);
  console.log('setupMobileForm: link written');

  alert_(
    'Mobile form created / refreshed.\n\nResponder link:\n' + url +
    '\n\nShare this link with field officers. Use the owner account that installed this script.'
  );
}


function getOrCreateMobileForm_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(MOBILE_FORM_PROPERTY);
  if (id) {
    try {
      return FormApp.openById(id);
    } catch (err) {
      // recreate below
    }
  }
  var form = FormApp.create('Fuel Purchase');
  props.setProperty(MOBILE_FORM_PROPERTY, form.getId());
  return form;
}


function getOrAddItem_(form, title, type) {
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === title) {
      return asTypedItem_(items[i], type);
    }
  }
  if (type === 'date') {
    var dateItem = form.addDateItem();
    dateItem.setTitle(title);
    return dateItem;
  }
  if (type === 'list') {
    var listItem = form.addListItem();
    listItem.setTitle(title);
    return listItem;
  }
  var textItem = form.addTextItem();
  textItem.setTitle(title);
  return textItem;
}


function asTypedItem_(item, type) {
  try {
    if (type === 'date' && typeof item.asDateItem === 'function') {
      return item.asDateItem();
    }
    if (type === 'list' && typeof item.asListItem === 'function') {
      return item.asListItem();
    }
    if (type === 'text' && typeof item.asTextItem === 'function') {
      return item.asTextItem();
    }
  } catch (err) {
    // fall through to the original item
  }
  return item;
}


function removeReceiptItems_(form) {
  var items = form.getItems();
  for (var i = items.length - 1; i >= 0; i--) {
    if (/^receipt/i.test(items[i].getTitle())) {
      form.deleteItem(items[i]);
    }
  }
}


function syncPurchaserChoices_(form, ss) {
  var purchasers = getPurchasers_(ss);
  if (!purchasers.length) {
    throw new Error('Add purchaser names in Settings first.');
  }
  var items = form.getItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === Q_PURCHASER) {
      setPurchaserChoices_(items[i], purchasers);
      return;
    }
  }
  var item = form.addListItem();
  item.setTitle(Q_PURCHASER);
  item.setRequired(true);
  setPurchaserChoices_(item, purchasers);
}


function setPurchaserChoices_(item, purchasers) {
  if (typeof item.setChoiceValues === 'function') {
    item.setChoiceValues(purchasers);
    return;
  }
  if (typeof item.asListItem === 'function') {
    item.asListItem().setChoiceValues(purchasers);
    return;
  }
  throw new Error(
    'The "Purchaser" question in the Google Form is not a dropdown. ' +
    'Delete that question in the Form editor and rerun setupMobileForm.'
  );
}


function setDateIncludesYear_(item) {
  try {
    if (typeof item.setIncludesYear === 'function') {
      item.setIncludesYear(true);
      return;
    }
    if (typeof item.asDateItem === 'function') {
      item.asDateItem().setIncludesYear(true);
    }
  } catch (err) {
    // The date question still works without this optional setting.
  }
}


function syncMobileFormPurchasers() {
  var ss = SpreadsheetApp.getActive();
  var form = getMobileForm_();
  if (!form) {
    alert_('Create the mobile form first.');
    return;
  }
  syncPurchaserChoices_(form, ss);
  writeMobileFormLink_(ss, form.getPublishedUrl());
  alert_('Purchaser choices synced from Settings.');
}


function getMobileForm_() {
  var id = PropertiesService.getScriptProperties().getProperty(MOBILE_FORM_PROPERTY);
  if (!id) return null;
  try {
    return FormApp.openById(id);
  } catch (err) {
    return null;
  }
}


function showMobileFormLink() {
  var form = getMobileForm_();
  if (!form) {
    alert_(
      'No mobile form yet. Run Fuel Log > Create / refresh mobile form.'
    );
    return;
  }
  writeMobileFormLink_(SpreadsheetApp.getActive(), form.getPublishedUrl());
  alert_(
    'Mobile form responder link:\n\n' + form.getPublishedUrl()
  );
}


function writeMobileFormLink_(ss, url) {
  var settings = ss.getSheetByName('Settings');
  if (!settings) return;
  var cell = settings.getRange('B67');
  cell.setFormula('=HYPERLINK("' + String(url).replace(/"/g, '""') +
                  '","Open mobile form")');
  cell.setFontColor('#1E4D3D');
  cell.setFontWeight('bold');
  cell.setHorizontalAlignment('left');
}


function processFormResponse(e) {
  if (!e || !e.response) return;

  var props = PropertiesService.getScriptProperties();
  var masterId = props.getProperty(MASTER_SHEET_PROPERTY);
  var ss = masterId
    ? SpreadsheetApp.openById(masterId)
    : (SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive());
  if (!ss) return;
  var response = e.response;
  var email = response.getRespondentEmail() || 'unknown';
  var submittedAt = response.getTimestamp() || new Date();
  var answers = responseItemMap_(response);

  var date = toDate_(answers[Q_DATE]);
  var purchaser = String(answers[Q_PURCHASER] || '').trim();
  var car = normalizeCar_(answers[Q_CAR]);
  var amount = Number(String(answers[Q_AMOUNT] || '').replace(/,/g, ''));

  try {
    if (!date) throw new Error('Invalid purchase date.');
    if (!ss.getSheetByName(yearLogName_(date))) {
      throw new Error('No log for ' + date.getFullYear() + '.');
    }
    if (!purchaser || getPurchasers_(ss).indexOf(purchaser) === -1) {
      throw new Error('Purchaser is not in Settings.');
    }
    if (!isFinite(amount) || amount <= 0) {
      throw new Error('Amount must be greater than zero.');
    }
    if (!car) {
      throw new Error('Car Number is required.');
    }
    if (!isAuthorizedRespondent_(ss, email)) {
      throw new Error('Respondent is not an editor of this Sheet.');
    }

    appendPurchaseRecord_(
      ss, date, purchaser, amount, car, email, 'Added via form', submittedAt
    );
  } catch (err) {
    rejectForm_(ss, email, date, purchaser, amount, car, err.message || String(err));
  }
}


function responseItemMap_(response) {
  var map = {};
  var items = response.getItemResponses();
  for (var i = 0; i < items.length; i++) {
    map[items[i].getItem().getTitle()] = items[i].getResponse();
  }
  return map;
}


function toDate_(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'string') {
    var parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}


function isAuthorizedRespondent_(ss, email) {
  if (!email || email === 'unknown') return false;
  var allowed = [];
  try {
    var file = DriveApp.getFileById(ss.getId());
    allowed.push(String(file.getOwner().getEmail() || '').toLowerCase());
    var editors = file.getEditors();
    for (var i = 0; i < editors.length; i++) {
      allowed.push(String(editors[i].getEmail() || '').toLowerCase());
    }
  } catch (err) {
    return false;
  }
  return allowed.indexOf(String(email).toLowerCase()) !== -1;
}


function rejectForm_(ss, email, date, purchaser, amount, car, reason) {
  try {
    var audit = ensureAuditSheet_(ss);
    appendAuditRows_(audit, [[
      new Date(), email, 'Form', '', date || '', purchaser || '',
      isFinite(amount) ? amount : '', car || '', 'Validation', '',
      reason, 'Rejected via form'
    ]]);
  } catch (err) {
    // Nothing else can be done if the audit log is unavailable.
  }
}


function hideFormResponses_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (/^Form Responses/.test(sheets[i].getName())) {
      try { sheets[i].hideSheet(); } catch (err) {}
    }
  }
}


function openCurrentLog() {
  var ss = SpreadsheetApp.getActive();
  var year = new Date().getFullYear();
  var log = ss.getSheetByName(year + ' Log');
  if (log) {
    ss.setActiveSheet(log);
  } else {
    alert_('There is no ' + year + ' Log in this workbook.');
  }
}


function setupAudit() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties()
    .setProperty(MASTER_SHEET_PROPERTY, ss.getId());
  ensureAuditSheet_(ss);

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'onEdit' || handler === 'auditOnEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('auditOnEdit').forSpreadsheet(ss).onEdit().create();

  alert_(
    'Audit trigger installed.\n\nUse Fuel Log > Add Purchase or the mobile Form. ' +
    'Every change is written to the Audit Log.'
  );
}


function openAuditLog() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ensureAuditSheet_(ss);
  ss.setActiveSheet(sheet);
}
