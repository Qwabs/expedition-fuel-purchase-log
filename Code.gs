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
var STRUCTURAL_CHANGE_TYPES = [
  'INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN'
];
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

var REPORT_SHEET = 'Weekly Report';
var REPORT_RECIPIENTS_CELL = 'B71';
var REPORT_LAST_SENT_CELL = 'B72';


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
    .addItem('Weekly report', 'showWeeklyReport')
    .addItem('Email weekly report', 'emailWeeklyReport')
    .addItem('Set weekly report recipients', 'setWeeklyReportRecipients')
    .addSeparator()
    .addItem('Audit logs', 'showLogHealth')
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
    today: formatDate_(new Date(), 'yyyy-MM-dd')
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


function auditOnChange(e) {
  if (!e || !e.changeType) return;
  if (STRUCTURAL_CHANGE_TYPES.indexOf(e.changeType) === -1) return;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    var ss = e.source || SpreadsheetApp.getActive();
    if (!ss) return;

    // The change event exposes the change type and user, but not the affected
    // sheet or row. Version history is still required to recover row contents.
    var audit = ensureAuditSheet_(ss);
    appendAuditRows_(audit, [[
      new Date(),
      getEditorEmail_(e),
      'Workbook',
      '',
      '',
      '',
      '',
      '',
      e.changeType,
      '',
      '',
      'Structure change'
    ]]);
  } catch (err) {
    console.log('auditOnChange failed: ' + err.message);
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
    return Utilities.formatDate(value, timeZone_(),
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
    if (handler === 'onEdit' || handler === 'auditOnEdit' ||
        handler === 'auditOnChange') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('auditOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('auditOnChange').forSpreadsheet(ss).onChange().create();

  alert_(
    'Audit triggers installed.\n\nUse Fuel Log > Add Purchase or the mobile Form. ' +
    'Cell edits and row/column structure changes are written to the Audit Log.'
  );
}


function openAuditLog() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ensureAuditSheet_(ss);
  ss.setActiveSheet(sheet);
}


// --------------------------------------------------------------- log health
function showLogHealth() {
  var ss = SpreadsheetApp.getActive();
  var findings = scanLogHealth_(ss);
  var ui = SpreadsheetApp.getUi();

  if (!findings.length) {
    ui.alert(
      'Log health check',
      'No gaps, hidden rows, possible duplicates or missing fields were found in the year logs.',
      ui.ButtonSet.OK
    );
    return;
  }

  var issueOrder = [
    'Row gap', 'Hidden rows', 'Possible duplicate', 'Missing date',
    'Invalid date', 'Missing car number', 'Missing purchaser',
    'Purchaser not in Settings', 'Missing amount', 'Date outside log year'
  ];
  var counts = {};
  for (var i = 0; i < findings.length; i++) {
    counts[findings[i].issue] = (counts[findings[i].issue] || 0) + 1;
  }

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#20342b;margin:14px;}' +
    'h2{margin:0 0 4px;color:#1E4D3D;font-size:20px;}' +
    '.muted{color:#62746b;font-size:12px;margin:0 0 14px;}' +
    '.summary{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 14px;}' +
    '.pill{border:1px solid #cfded6;background:#f3f8f5;border-radius:14px;' +
    'padding:5px 10px;font-size:12px;}' +
    '.pill strong{color:#1E4D3D;}' +
    'table{border-collapse:collapse;width:100%;font-size:12px;}' +
    'th{position:sticky;top:0;background:#1E4D3D;color:#fff;text-align:left;' +
    'padding:7px 8px;}' +
    'td{border-bottom:1px solid #e1e9e4;padding:7px 8px;vertical-align:top;}' +
    'tr:nth-child(even) td{background:#f8fbf9;}' +
    '.sheet{white-space:nowrap;font-weight:bold;}' +
    '.rows{white-space:nowrap;}' +
    '.issue{white-space:nowrap;color:#8a3b12;font-weight:bold;}' +
    '.note{margin-top:10px;color:#62746b;font-size:12px;}' +
    '</style></head><body>';
  html += '<h2>Log health check</h2>';
  html += '<p class="muted">Read-only scan · ' +
    escapeHtml_(formatDate_(new Date(), 'd MMM yyyy HH:mm')) + '</p>';
  html += '<div class="summary">';
  for (var j = 0; j < issueOrder.length; j++) {
    var issue = issueOrder[j];
    if (counts[issue]) {
      html += '<span class="pill">' + escapeHtml_(issue) + ': <strong>' +
        counts[issue] + '</strong></span>';
    }
  }
  html += '</div>';

  var limit = 300;
  html += '<table><thead><tr><th>Sheet</th><th>Rows</th><th>Issue</th>' +
    '<th>Details</th></tr></thead><tbody>';
  for (var k = 0; k < findings.length && k < limit; k++) {
    var finding = findings[k];
    html += '<tr><td class="sheet">' + escapeHtml_(finding.sheet) +
      '</td><td class="rows">' + escapeHtml_(finding.rows) +
      '</td><td class="issue">' + escapeHtml_(finding.issue) +
      '</td><td>' + escapeHtml_(finding.detail) + '</td></tr>';
  }
  html += '</tbody></table>';
  if (findings.length > limit) {
    html += '<p class="note">Showing the first ' + limit + ' of ' +
      findings.length + ' findings. Fix these and run the check again.</p>';
  }
  html += '<p class="note">This check never edits the workbook. ' +
    'Use File → Version history to recover deleted rows.</p>';
  html += '</body></html>';

  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(960).setHeight(680),
    'Log health check'
  );
}


function scanLogHealth_(ss) {
  var findings = [];
  var sheets = ss.getSheets();
  var purchaserList = getPurchasers_(ss);
  var purchaserLookup = {};
  for (var p = 0; p < purchaserList.length; p++) {
    purchaserLookup[purchaserList[p]] = true;
  }

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var match = sheet.getName().match(/^(\d{4}) Log$/);
    if (!match) continue;

    var year = Number(match[1]);
    var height = Math.min(LOG_ROWS, sheet.getMaxRows() - ENTRY_FIRST_ROW + 1);
    if (height <= 0) continue;

    var values = sheet.getRange(
      ENTRY_FIRST_ROW, COL_DATE, height, COL_ENTERED_AT
    ).getValues();
    var firstRow = -1;
    var lastRow = -1;
    var row;

    for (var i = 0; i < values.length; i++) {
      if (healthRowHasData_(values[i])) {
        if (firstRow === -1) firstRow = ENTRY_FIRST_ROW + i;
        lastRow = ENTRY_FIRST_ROW + i;
      }
    }
    if (firstRow === -1) continue;

    // Gaps and field checks within the populated part of the log.
    var gapStart = -1;
    var duplicateRows = {};

    for (row = firstRow; row <= lastRow; row++) {
      var rowValues = values[row - ENTRY_FIRST_ROW];
      if (!healthRowHasData_(rowValues)) {
        if (gapStart === -1) gapStart = row;
        continue;
      }
      if (gapStart !== -1) {
        findings.push({
          sheet: sheet.getName(),
          rows: gapStart + '–' + (row - 1),
          issue: 'Row gap',
          detail: 'No purchase data between the rows above and below.'
        });
        gapStart = -1;
      }

      var date = rowValues[COL_DATE - 1];
      var purchaser = rowValues[COL_PURCHASER - 1];
      var amount = rowValues[COL_AMOUNT - 1];
      var car = rowValues[COL_CAR - 1];
      var hasCore = !isBlank_(purchaser) || !isBlank_(amount) || !isBlank_(car);

      if (isBlank_(date)) {
        if (hasCore) {
          findings.push({
            sheet: sheet.getName(),
            rows: String(row),
            issue: 'Missing date',
            detail: 'Purchaser, amount or car number is present without a purchase date.'
          });
        }
        continue;
      }

      if (!(date instanceof Date)) {
        findings.push({
          sheet: sheet.getName(),
          rows: String(row),
          issue: 'Invalid date',
          detail: 'The purchase date is not a valid date value.'
        });
        continue;
      }

      if (date.getFullYear() !== year) {
        findings.push({
          sheet: sheet.getName(),
          rows: String(row),
          issue: 'Date outside log year',
          detail: 'The date is ' + formatDate_(date, 'd MMM yyyy') +
            ', not in ' + year + '.'
        });
      }

      if (isBlank_(purchaser)) {
        findings.push({
          sheet: sheet.getName(),
          rows: String(row),
          issue: 'Missing purchaser',
          detail: 'The purchase has no purchaser.'
        });
      } else if (purchaserList.length) {
        var rawPurchaser = String(purchaser);
        var purchaserName = rawPurchaser.trim();
        if (purchaserName && purchaserLookup[purchaserName] !== true) {
          findings.push({
            sheet: sheet.getName(),
            rows: String(row),
            issue: 'Purchaser not in Settings',
            detail: 'Add "' + purchaserName + '" to Settings!A15:A64 so ' +
              'the dropdown, reports and Form include it.'
          });
        } else if (rawPurchaser !== purchaserName) {
          findings.push({
            sheet: sheet.getName(),
            rows: String(row),
            issue: 'Purchaser not in Settings',
            detail: 'The purchaser name has leading or trailing spaces; ' +
              'trim it to match Settings.'
          });
        }
      }
      if (isBlank_(amount) || !isFinite(Number(amount)) || Number(amount) === 0) {
        findings.push({
          sheet: sheet.getName(),
          rows: String(row),
          issue: 'Missing amount',
          detail: 'The purchase has no valid amount.'
        });
      }
      if (isBlank_(car)) {
        findings.push({
          sheet: sheet.getName(),
          rows: String(row),
          issue: 'Missing car number',
          detail: 'The purchase has no car number.'
        });
      }

      if (!isBlank_(purchaser) && !isBlank_(amount) && !isBlank_(car)) {
        var key = formatDate_(date, 'yyyy-MM-dd') + '|' +
          String(purchaser).trim().toUpperCase() + '|' +
          Number(amount).toFixed(2) + '|' + normalizeCar_(car);
        if (!duplicateRows[key]) duplicateRows[key] = [];
        duplicateRows[key].push(row);
      }
    }

    if (gapStart !== -1) {
      findings.push({
        sheet: sheet.getName(),
        rows: gapStart + '–' + lastRow,
        issue: 'Row gap',
        detail: 'No purchase data before the last populated row.'
      });
    }

    // Hidden rows are the easiest way for purchases to look missing while
    // remaining present. Report them separately from blank gaps.
    var hiddenStart = -1;
    var hiddenEnd = -1;
    var hiddenKind = '';

    for (row = firstRow; row <= lastRow; row++) {
      // Only populated rows can represent missing purchases. Skipping blank
      // rows also keeps the hidden-row check inexpensive on a large log.
      if (!healthRowHasData_(values[row - ENTRY_FIRST_ROW])) {
        if (hiddenStart !== -1) {
          addHiddenFinding_(findings, sheet, values, firstRow,
                            hiddenStart, hiddenEnd, hiddenKind);
          hiddenStart = -1;
        }
        continue;
      }

      var kind = '';
      try {
        if (sheet.isRowHiddenByFilter(row)) {
          kind = 'filter';
        }
      } catch (errFilter) {
        // Older or unusual sheet objects may not expose this method.
      }
      if (!kind) {
        try {
          if (sheet.isRowHiddenByUser(row)) kind = 'user';
        } catch (errUser) {
          // Fall through with no hidden state.
        }
      }
      if (!kind && typeof sheet.getRowGroup === 'function') {
        for (var depth = 1; depth <= 8; depth++) {
          try {
            var group = sheet.getRowGroup(row, depth);
            if (group && group.isCollapsed()) {
              kind = 'group';
              break;
            }
          } catch (errGroup) {
            // No row group at this depth.
          }
        }
      }

      if (kind) {
        if (hiddenStart !== -1 &&
            (hiddenKind !== kind || row !== hiddenEnd + 1)) {
          addHiddenFinding_(findings, sheet, values, firstRow,
                            hiddenStart, hiddenEnd, hiddenKind);
          hiddenStart = -1;
        }
        if (hiddenStart === -1) {
          hiddenStart = row;
          hiddenKind = kind;
        }
        hiddenEnd = row;
      } else if (hiddenStart !== -1) {
        addHiddenFinding_(findings, sheet, values, firstRow,
                          hiddenStart, hiddenEnd, hiddenKind);
        hiddenStart = -1;
      }
    }
    if (hiddenStart !== -1) {
      addHiddenFinding_(findings, sheet, values, firstRow,
                        hiddenStart, hiddenEnd, hiddenKind);
    }

    for (var key in duplicateRows) {
      if (!duplicateRows.hasOwnProperty(key)) continue;
      var rows = duplicateRows[key];
      if (rows.length > 1) {
        findings.push({
          sheet: sheet.getName(),
          rows: rows.join(', '),
          issue: 'Possible duplicate',
          detail: 'Same date, purchaser, amount and car number on ' +
            rows.length + ' rows. Confirm this is not a double entry.'
        });
      }
    }
  }

  findings.sort(function (a, b) {
    if (a.sheet !== b.sheet) return a.sheet < b.sheet ? -1 : 1;
    var aRow = parseInt(String(a.rows), 10) || 0;
    var bRow = parseInt(String(b.rows), 10) || 0;
    return aRow - bRow;
  });
  return findings;
}


function healthRowHasData_(rowValues) {
  return !isBlank_(rowValues[COL_DATE - 1]) ||
    !isBlank_(rowValues[COL_PURCHASER - 1]) ||
    !isBlank_(rowValues[COL_AMOUNT - 1]) ||
    !isBlank_(rowValues[COL_CAR - 1]);
}


function addHiddenFinding_(findings, sheet, values, firstRow,
                           start, end, kind) {
  var entries = 0;
  for (var row = start; row <= end; row++) {
    if (healthRowHasData_(values[row - firstRow])) entries++;
  }
  var how = kind === 'filter' ? 'a filter' :
    (kind === 'group' ? 'a collapsed row group' : 'a hidden-row setting');
  findings.push({
    sheet: sheet.getName(),
    rows: start + '–' + end,
    issue: 'Hidden rows',
    detail: 'Rows are hidden by ' + how + '; ' + entries +
      ' purchase row(s) are not visible.'
  });
}

// ---------------------------------------------------------------- weekly report
function showWeeklyReport() {
  var ss = SpreadsheetApp.getActive();
  ensureWeeklyReportSettings_(ss);
  var weekStart = promptWeekStart_();
  if (!weekStart) return;

  var data = buildWeeklyReportData_(weekStart, ss);
  refreshWeeklyReport_(data);
  ss.setActiveSheet(ss.getSheetByName(REPORT_SHEET));
}


function emailWeeklyReport() {
  var ss = SpreadsheetApp.getActive();
  ensureWeeklyReportSettings_(ss);

  var weekStart = promptWeekStart_();
  if (!weekStart) return;

  var data = buildWeeklyReportData_(weekStart, ss);
  refreshWeeklyReport_(data);

  var saved = getWeeklyRecipients_(ss);
  var prompt = 'Recipients (comma-separated). Leave blank to use the saved list.';
  if (saved.length) {
    prompt += '\n\nSaved: ' + saved.join(', ');
  }
  var response = SpreadsheetApp.getUi().prompt(
    'Email weekly report', prompt, SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;

  var entered = String(response.getResponseText() || '').trim();
  var recipients = entered ? parseRecipients_(entered) : { valid: saved, invalid: [] };
  if (!recipients.valid.length) {
    alert_('Enter at least one valid recipient email address.');
    return;
  }
  if (recipients.invalid.length) {
    alert_('These addresses are invalid:\n' + recipients.invalid.join('\n'));
    return;
  }

  var subject = 'Fuel Purchase Weekly Report — ' + weekRangeLabel_(data);
  var confirm = SpreadsheetApp.getUi().alert(
    'Send weekly report?',
    subject + '\n\nTo: ' + recipients.valid.join(', ') +
    '\nPurchases: ' + data.count + '\nTotal: ' + formatMoney_(data.total),
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  var pdf = buildWeeklyReportPdf_(data);
  sendWeeklyReportEmail_(data, recipients.valid, pdf);
  writeLastSent_(ss);
  ss.setActiveSheet(ss.getSheetByName(REPORT_SHEET));

  alert_(
    'Weekly report emailed to:\n' + recipients.valid.join('\n') +
    (pdf ? '\n\nPDF attached.' : '\n\nPDF generation was unavailable; the HTML report was sent.')
  );
}


function setWeeklyReportRecipients() {
  var ss = SpreadsheetApp.getActive();
  ensureWeeklyReportSettings_(ss);
  var saved = getWeeklyRecipients_(ss);
  var response = SpreadsheetApp.getUi().prompt(
    'Weekly report recipients',
    'Enter one or more email addresses, comma-separated.' +
    (saved.length ? '\n\nCurrent: ' + saved.join(', ') : ''),
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return;

  var parsed = parseRecipients_(response.getResponseText());
  if (parsed.invalid.length) {
    alert_('These addresses are invalid:\n' + parsed.invalid.join('\n'));
    return;
  }
  saveWeeklyRecipients_(ss, parsed.valid);
  alert_(parsed.valid.length ? 'Weekly report recipients saved.' : 'Recipient list cleared.');
}


function promptWeekStart_() {
  var ss = SpreadsheetApp.getActive();
  var currentMonday = normalizeMonday_(new Date());
  var defaultWeek = addDays_(currentMonday, -7);
  var defaultText = formatDate_(defaultWeek, 'yyyy-MM-dd');
  var response = SpreadsheetApp.getUi().prompt(
    'Weekly report',
    'Week starting date (yyyy-mm-dd). Leave blank for ' + defaultText + '.',
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== SpreadsheetApp.getUi().Button.OK) return null;

  var text = String(response.getResponseText() || '').trim();
  var date = text ? parseDate_(text) : defaultWeek;
  if (!date) {
    alert_('Enter a valid date in the format yyyy-mm-dd.');
    return null;
  }
  return normalizeMonday_(date);
}


function buildWeeklyReportData_(weekStart, ss) {
  var start = dateOnly_(normalizeMonday_(weekStart));
  var end = dateOnly_(addDays_(start, 6));
  var entries = [];
  var years = [start.getFullYear(), end.getFullYear()];
  if (years[0] === years[1]) years.pop();

  for (var y = 0; y < years.length; y++) {
    var log = ss.getSheetByName(years[y] + ' Log');
    if (!log) continue;
    var values = log.getRange(ENTRY_FIRST_ROW, 1, LOG_ROWS, 7).getValues();
    for (var i = 0; i < values.length; i++) {
      var rawDate = values[i][COL_DATE - 1];
      if (!(rawDate instanceof Date)) continue;
      var day = dateOnly_(rawDate);
      if (day < start || day > end) continue;

      entries.push({
        date: day,
        purchaser: String(values[i][COL_PURCHASER - 1] || '').trim() || '(blank)',
        amount: isFinite(Number(values[i][COL_AMOUNT - 1]))
          ? Number(values[i][COL_AMOUNT - 1]) : 0,
        car: normalizeCar_(values[i][COL_CAR - 1]) || '(missing)',
        enteredBy: String(values[i][COL_ENTERED_BY - 1] || '').trim(),
        enteredAt: values[i][COL_ENTERED_AT - 1] instanceof Date
          ? values[i][COL_ENTERED_AT - 1] : null
      });
    }
  }

  entries.sort(function (a, b) {
    var delta = a.date.getTime() - b.date.getTime();
    if (delta) return delta;
    var aTime = a.enteredAt ? a.enteredAt.getTime() : 0;
    var bTime = b.enteredAt ? b.enteredAt.getTime() : 0;
    return aTime - bTime;
  });

  var total = 0;
  var count = 0;
  for (var e = 0; e < entries.length; e++) {
    total += entries[e].amount;
    if (entries[e].amount > 0) count++;
  }

  return {
    weekStart: start,
    weekEnd: end,
    entries: entries,
    total: total,
    count: count,
    average: count ? total / count : 0,
    byPurchaser: summarizeReport_(entries, 'purchaser'),
    byCar: summarizeReport_(entries, 'car')
  };
}


function summarizeReport_(entries, key) {
  var map = {};
  for (var i = 0; i < entries.length; i++) {
    var name = entries[i][key];
    if (!map[name]) map[name] = { name: name, count: 0, total: 0 };
    map[name].count += entries[i].amount > 0 ? 1 : 0;
    map[name].total += entries[i].amount;
  }
  var out = [];
  for (var k in map) if (Object.prototype.hasOwnProperty.call(map, k)) out.push(map[k]);
  out.sort(function (a, b) {
    if (b.total !== a.total) return b.total - a.total;
    return a.name.localeCompare(b.name);
  });
  return out;
}


function refreshWeeklyReport_(data) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(REPORT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REPORT_SHEET, 1);
  } else {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    sheet.clear();
  }

  try { sheet.setHiddenGridlines(true); } catch (err) {}
  try { sheet.setTabColor('#B08D57'); } catch (err) {}
  try { sheet.setFrozenRows(7); } catch (err) {}

  var widths = [13, 12, 24, 22, 16, 30];
  for (var c = 0; c < widths.length; c++) sheet.setColumnWidth(c + 1, widths[c] * 8);

  var title = sheet.getRange(1, 1, 1, 6);
  title.merge();
  title.setValue('WEEKLY FUEL REPORT · ' + weekRangeLabel_(data))
    .setBackground('#1E4D3D').setFontColor('#FFFFFF')
    .setFontSize(16).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  var subtitle = sheet.getRange(2, 1, 1, 6);
  subtitle.merge();
  subtitle.setValue('Monday–Sunday · Generated from the live purchase logs')
    .setFontColor('#6B7280').setFontStyle('italic').setFontSize(9);
  sheet.setRowHeight(2, 20);

  sheet.getRange('A3:B3').setValues([['Week starting', data.weekStart]]);
  sheet.getRange('A4:B4').setValues([['Week ending', data.weekEnd]]);
  sheet.getRange('A5:B5').setValues([['Generated', new Date()]]);
  sheet.getRange('B3:B4').setNumberFormat('dd mmm yyyy');
  sheet.getRange('B5').setNumberFormat('dd mmm yyyy hh:mm');
  sheet.getRange('A3:A5').setFontWeight('bold').setFontColor('#1E4D3D');
  sheet.getRange('A3:B5').setBackground('#FBF7EC');
  sheet.getRange('A3:B5').setBorder(true, true, true, true, true, true,
                                    '#D6DDD8', SpreadsheetApp.BorderStyle.SOLID);

  var band = sheet.getRange(6, 1, 1, 6);
  band.merge();
  band.setValue('ENTRIES').setBackground('#1E4D3D').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(6, 24);

  var headers = ['Date', 'Day', 'Purchaser', 'Car Number', 'Amount (GH₵)', 'Entered by'];
  var headerRange = sheet.getRange(7, 1, 1, 6);
  headerRange.setValues([headers]).setBackground('#EFE4CF')
    .setFontColor('#6B4F2A').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#D6DDD8',
                SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(7, 24);

  var row = 8;
  for (var i = 0; i < data.entries.length; i++) {
    var entry = data.entries[i];
    var values = [[
      entry.date, reportDay_(entry.date), entry.purchaser, entry.car,
      entry.amount, entry.enteredBy
    ]];
    var range = sheet.getRange(row, 1, 1, 6);
    range.setValues(values)
      .setFontSize(10)
      .setBackground(i % 2 === 0 ? '#FFFFFF' : '#FBF9F2')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#D6DDD8',
                  SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, 1).setNumberFormat('dd mmm yyyy');
    sheet.getRange(row, 5).setNumberFormat('"GH₵"#,##0.00');
    sheet.getRange(row, 5).setHorizontalAlignment('right');
    sheet.getRange(row, 2).setHorizontalAlignment('center');
    sheet.setRowHeight(row, 20);
    row++;
  }

  if (!data.entries.length) {
    var empty = sheet.getRange(8, 1, 1, 6);
    empty.merge();
    empty.setValue('No purchases recorded for this week.')
      .setFontColor('#6B7280').setFontStyle('italic')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBackground('#FBF7EC');
    sheet.setRowHeight(8, 26);
    row = 9;
  }

  row += 1;
  sheet.getRange(row, 1).setValue('Purchases').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(data.count).setHorizontalAlignment('center');
  sheet.getRange(row, 4).setValue('Total').setFontWeight('bold');
  sheet.getRange(row, 5).setValue(data.total).setNumberFormat('"GH₵"#,##0.00')
    .setHorizontalAlignment('right');
  sheet.getRange(row, 1, 1, 5).setBackground('#FBF7EC')
    .setBorder(true, true, true, true, true, true, '#D6DDD8',
                SpreadsheetApp.BorderStyle.SOLID);
  row++;
  sheet.getRange(row, 1).setValue('Average').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(data.average).setNumberFormat('"GH₵"#,##0.00')
    .setHorizontalAlignment('center');
  sheet.getRange(row, 1, 1, 5).setBackground('#FBF7EC')
    .setBorder(true, true, true, true, true, true, '#D6DDD8',
                SpreadsheetApp.BorderStyle.SOLID);

  row += 3;
  var purchaserBand = sheet.getRange(row, 1, 1, 6);
  purchaserBand.merge();
  purchaserBand.setValue('BY PURCHASER')
    .setBackground('#1E4D3D').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('left');
  sheet.setRowHeight(row, 22);
  row++;
  sheet.getRange(row, 1, 1, 3).setValues([['Purchaser', 'Purchases', 'Total']])
    .setBackground('#EFE4CF').setFontColor('#6B4F2A').setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  for (var p = 0; p < data.byPurchaser.length; p++) {
    var item = data.byPurchaser[p];
    sheet.getRange(row, 1, 1, 3).setValues([[item.name, item.count, item.total]])
      .setBackground(p % 2 === 0 ? '#FFFFFF' : '#FBF9F2')
      .setBorder(true, true, true, true, true, true, '#D6DDD8',
                  SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, 3).setNumberFormat('"GH₵"#,##0.00')
      .setHorizontalAlignment('right');
    sheet.getRange(row, 2).setHorizontalAlignment('center');
    row++;
  }
  if (!data.byPurchaser.length) {
    sheet.getRange(row, 1).setValue('No entries').setFontColor('#6B7280');
    row++;
  }

  row += 2;
  var carBand = sheet.getRange(row, 1, 1, 6);
  carBand.merge();
  carBand.setValue('BY CAR')
    .setBackground('#1E4D3D').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('left');
  sheet.setRowHeight(row, 22);
  row++;
  sheet.getRange(row, 1, 1, 3).setValues([['Car Number', 'Purchases', 'Total']])
    .setBackground('#EFE4CF').setFontColor('#6B4F2A').setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  for (var c2 = 0; c2 < data.byCar.length; c2++) {
    var carItem = data.byCar[c2];
    sheet.getRange(row, 1, 1, 3).setValues([[carItem.name, carItem.count, carItem.total]])
      .setBackground(c2 % 2 === 0 ? '#FFFFFF' : '#FBF9F2')
      .setBorder(true, true, true, true, true, true, '#D6DDD8',
                  SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(row, 3).setNumberFormat('"GH₵"#,##0.00')
      .setHorizontalAlignment('right');
    sheet.getRange(row, 2).setHorizontalAlignment('center');
    row++;
  }
  if (!data.byCar.length) {
    sheet.getRange(row, 1).setValue('No entries').setFontColor('#6B7280');
    row++;
  }

  row += 2;
  var footer = sheet.getRange(row, 1, 1, 6);
  footer.merge();
  footer.setValue('Generated from live purchase logs. Use Fuel Log → Weekly report to refresh.')
    .setFontSize(8).setFontStyle('italic').setFontColor('#6B7280');
  sheet.setRowHeight(row, 20);

  try { sheet.setPageOrientation('landscape'); } catch (err) {}
  try { sheet.setFrozenRows(7); } catch (err) {}
  try { sheet.autoResizeColumn(1); } catch (err) {}
  sheet.getRange(1, 1, row, 6).setFontFamily('Arial');
}


function buildWeeklyReportPdf_(data) {
  try {
    var html = buildWeeklyReportHtml_(data);
    var blob = Utilities.newBlob(html, MimeType.HTML, 'weekly-report.html');
    var pdf = blob.getAs(MimeType.PDF);
    pdf.setName('Fuel-Weekly-Report-' + formatDate_(data.weekStart, 'yyyy-MM-dd') + '.pdf');
    return pdf;
  } catch (err) {
    console.log('Weekly PDF generation failed: ' + err);
    return null;
  }
}


function sendWeeklyReportEmail_(data, recipients, pdf) {
  var sender = Session.getActiveUser().getEmail();
  var bcc = [];
  for (var i = 0; i < recipients.length; i++) {
    if (!sender || recipients[i].toLowerCase() !== sender.toLowerCase()) bcc.push(recipients[i]);
  }

  var params = {
    subject: 'Fuel Purchase Weekly Report — ' + weekRangeLabel_(data),
    htmlBody: buildWeeklyReportHtml_(data),
    name: 'Fuel Purchase Log'
  };
  if (sender) {
    params.to = sender;
    params.replyTo = sender;
    if (bcc.length) params.bcc = bcc.join(',');
  } else {
    params.to = recipients[0];
    if (recipients.length > 1) params.bcc = recipients.slice(1).join(',');
  }
  if (pdf) params.attachments = [pdf];
  MailApp.sendEmail(params);
}


function buildWeeklyReportHtml_(data) {
  var html = [];
  html.push('<div style="font-family:Arial,Helvetica,sans-serif;color:#374151;max-width:900px">');
  html.push('<h1 style="color:#1E4D3D;font-size:20px;margin:0 0 4px">FUEL PURCHASE WEEKLY REPORT</h1>');
  html.push('<p style="color:#6B7280;font-size:12px;margin:0 0 18px">' +
            escapeHtml_(weekRangeLabel_(data)) + ' · Ghana Cedi (GH₵)</p>');

  html.push('<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px">');
  html.push('<tr><td style="background:#FBF7EC;padding:8px;border:1px solid #D6DDD8"><b>Purchases</b></td>' +
            '<td style="padding:8px;border:1px solid #D6DDD8" align="right">' + data.count + '</td>' +
            '<td style="background:#FBF7EC;padding:8px;border:1px solid #D6DDD8"><b>Total</b></td>' +
            '<td style="padding:8px;border:1px solid #D6DDD8" align="right">' + formatMoney_(data.total) + '</td>' +
            '<td style="background:#FBF7EC;padding:8px;border:1px solid #D6DDD8"><b>Average</b></td>' +
            '<td style="padding:8px;border:1px solid #D6DDD8" align="right">' + formatMoney_(data.average) + '</td></tr>');
  html.push('</table>');

  html.push('<h2 style="color:#1E4D3D;font-size:14px;margin:18px 0 6px">Entries</h2>');
  html.push('<table style="width:100%;border-collapse:collapse;font-size:11px">');
  html.push('<thead><tr>' +
            '<th style="background:#1E4D3D;color:#fff;padding:6px;text-align:left">Date</th>' +
            '<th style="background:#1E4D3D;color:#fff;padding:6px;text-align:left">Purchaser</th>' +
            '<th style="background:#1E4D3D;color:#fff;padding:6px;text-align:left">Car Number</th>' +
            '<th style="background:#1E4D3D;color:#fff;padding:6px;text-align:right">Amount</th>' +
            '<th style="background:#1E4D3D;color:#fff;padding:6px;text-align:left">Entered by</th>' +
            '</tr></thead><tbody>');
  if (!data.entries.length) {
    html.push('<tr><td colspan="5" style="padding:10px;text-align:center;color:#6B7280">No purchases recorded for this week.</td></tr>');
  } else {
    for (var i = 0; i < data.entries.length; i++) {
      var e = data.entries[i];
      html.push('<tr style="background:' + (i % 2 === 0 ? '#FFFFFF' : '#FBF9F2') + '">' +
                '<td style="padding:6px;border-bottom:1px solid #E5E7EB">' + formatDate_(e.date, 'dd MMM yyyy') + '</td>' +
                '<td style="padding:6px;border-bottom:1px solid #E5E7EB">' + escapeHtml_(e.purchaser) + '</td>' +
                '<td style="padding:6px;border-bottom:1px solid #E5E7EB">' + escapeHtml_(e.car) + '</td>' +
                '<td style="padding:6px;border-bottom:1px solid #E5E7EB;text-align:right">' + formatMoney_(e.amount) + '</td>' +
                '<td style="padding:6px;border-bottom:1px solid #E5E7EB">' + escapeHtml_(e.enteredBy) + '</td>' +
                '</tr>');
    }
  }
  html.push('</tbody></table>');

  html.push(reportSummaryTable_('By Purchaser', data.byPurchaser, 'Purchaser'));
  html.push(reportSummaryTable_('By Car', data.byCar, 'Car Number'));
  html.push('<p style="color:#6B7280;font-size:10px;margin-top:18px">Generated from the live purchase logs. ' +
            'Currency: Ghana Cedi (GH₵).</p>');
  html.push('</div>');
  return html.join('');
}


function reportSummaryTable_(title, items, label) {
  var html = [];
  html.push('<h2 style="color:#1E4D3D;font-size:14px;margin:18px 0 6px">' +
            escapeHtml_(title) + '</h2>');
  html.push('<table style="width:100%;border-collapse:collapse;font-size:11px">');
  html.push('<thead><tr><th style="background:#EFE4CF;color:#6B4F2A;padding:6px;text-align:left">' +
            escapeHtml_(label) + '</th><th style="background:#EFE4CF;color:#6B4F2A;padding:6px">Purchases</th>' +
            '<th style="background:#EFE4CF;color:#6B4F2A;padding:6px;text-align:right">Total</th></tr></thead><tbody>');
  if (!items.length) {
    html.push('<tr><td colspan="3" style="padding:8px;color:#6B7280">No entries</td></tr>');
  } else {
    for (var i = 0; i < items.length; i++) {
      html.push('<tr><td style="padding:5px;border-bottom:1px solid #E5E7EB">' +
                escapeHtml_(items[i].name) + '</td><td style="padding:5px;border-bottom:1px solid #E5E7EB;text-align:center">' +
                items[i].count + '</td><td style="padding:5px;border-bottom:1px solid #E5E7EB;text-align:right">' +
                formatMoney_(items[i].total) + '</td></tr>');
    }
  }
  html.push('</tbody></table>');
  return html.join('');
}


function ensureWeeklyReportSettings_(ss) {
  try {
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return;
    if (isBlank_(sheet.getRange('A70').getValue())) {
      sheet.getRange('A70').setValue('WEEKLY REPORT')
        .setBackground('#FBF7EC').setFontWeight('bold').setFontColor('#8A6D3B');
      sheet.getRange('A71').setValue('Recipients').setFontWeight('bold');
      sheet.getRange('A72').setValue('Last sent').setFontWeight('bold');
      sheet.getRange('A73').setValue('Comma-separated emails. Manual sends can override.')
        .setFontSize(8).setFontStyle('italic').setFontColor('#6B7280');
      sheet.getRange('A70:B70').setBorder(false, false, true, false, false, false,
                                            '#D6DDD8', SpreadsheetApp.BorderStyle.SOLID);
    }
  } catch (err) {
    console.log('Weekly settings setup skipped: ' + err);
  }
}


function getWeeklyRecipients_(ss) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return [];
  return parseRecipients_(sheet.getRange(REPORT_RECIPIENTS_CELL).getValue()).valid;
}


function saveWeeklyRecipients_(ss, recipients) {
  var sheet = ss.getSheetByName('Settings');
  if (!sheet) return;
  sheet.getRange(REPORT_RECIPIENTS_CELL).setValue(recipients.join(', '));
}


function writeLastSent_(ss) {
  try {
    var sheet = ss.getSheetByName('Settings');
    if (!sheet) return;
    sheet.getRange(REPORT_LAST_SENT_CELL)
      .setValue(new Date())
      .setNumberFormat('dd mmm yyyy hh:mm');
  } catch (err) {
    console.log('Could not write last-sent time: ' + err);
  }
}


function parseRecipients_(raw) {
  var parts = String(raw === null || raw === undefined ? '' : raw)
    .split(/[\s,;]+/)
    .filter(function (value) { return value !== ''; });
  var valid = [];
  var invalid = [];
  var seen = {};
  for (var i = 0; i < parts.length; i++) {
    var email = parts[i];
    var key = email.toLowerCase();
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      invalid.push(email);
    } else if (!seen[key]) {
      seen[key] = true;
      valid.push(email);
    }
  }
  return { valid: valid, invalid: invalid };
}


function normalizeMonday_(date) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
  var offset = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - offset);
  return d;
}


function addDays_(date, days) {
  var d = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
  d.setDate(d.getDate() + days);
  return d;
}


function dateOnly_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}


function timeZone_() {
  var tz = '';
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) tz = ss.getSpreadsheetTimeZone();
  } catch (err) {
    // fall through to script timezone
  }
  if (typeof tz !== 'string' || tz.trim() === '') {
    try {
      tz = Session.getScriptTimeZone();
    } catch (err2) {
      tz = '';
    }
  }
  if (typeof tz !== 'string' || tz.trim() === '') {
    tz = 'Africa/Accra';
  }
  return tz;
}


function formatDate_(date, pattern) {
  return Utilities.formatDate(date, timeZone_(), pattern);
}


function reportDay_(date) {
  return formatDate_(date, 'EEEE');
}


function weekRangeLabel_(data) {
  var start = formatDate_(data.weekStart, 'd MMM');
  var end = formatDate_(data.weekEnd, 'd MMM yyyy');
  return start + ' – ' + end;
}


function formatMoney_(value) {
  var number = isFinite(Number(value)) ? Number(value) : 0;
  return 'GH₵' + number.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}


function escapeHtml_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
