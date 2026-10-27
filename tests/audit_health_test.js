const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

function pad(value) {
  return String(value).padStart(2, '0');
}

function mockFormatDate(date, timeZone, pattern) {
  if (pattern === 'yyyy-MM-dd') {
    return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' +
      pad(date.getDate());
  }
  if (pattern === 'd MMM yyyy') {
    return date.getDate() + ' Jan ' + date.getFullYear();
  }
  return date.toString();
}

function makeRange(sheet, row, col, numRows, numCols) {
  return {
    getValues: function () {
      if (sheet.headerRange &&
          row === 4 && col === 1 && numRows === 1 && numCols === 12) {
        return [sheet.headers.slice()];
      }
      if (sheet.logValues && row === 5 && col === 1 &&
          numRows === 5000 && numCols === 7) {
        return sheet.logValues;
      }
      if (sheet.data && row === 5 && col === 1 &&
          numRows === 1 && numCols === 12) {
        return [sheet.data[4] || []];
      }
      return [[].fill('')];
    },
    setValues: function (values) {
      if (sheet.data) {
        for (var i = 0; i < values.length; i++) {
          sheet.data[row - 1 + i] = values[i].slice();
        }
        sheet.lastRow = Math.max(sheet.lastRow || 0, row + values.length - 1);
      }
      return this;
    },
    setFontSize: function () { return this; },
    setBorder: function () { return this; },
    setNumberFormat: function () { return this; }
  };
}

function makeLogSheet(values, hiddenFilter, hiddenUser) {
  var sheet = {
    getName: function () { return '2026 Log'; },
    getMaxRows: function () { return 5004; },
    getRange: function (row, col, numRows, numCols) {
      return makeRange(sheet, row, col, numRows, numCols);
    },
    logValues: values,
    isRowHiddenByFilter: function (row) {
      return !!(hiddenFilter && hiddenFilter[row]);
    },
    isRowHiddenByUser: function (row) {
      return !!(hiddenUser && hiddenUser[row]);
    }
  };
  return sheet;
}

function makeAuditSheet() {
  var headers = [
    'Timestamp', 'User', 'Sheet', 'Row', 'Purchase date', 'Purchaser',
    'Amount (GH₵)', 'Car Number', 'Field', 'Old value',
    'New value', 'Action'
  ];
  var sheet = {
    headers: headers,
    headerRange: true,
    data: [],
    lastRow: 4,
    getName: function () { return 'Audit Log'; },
    getLastRow: function () { return sheet.lastRow; },
    getRange: function (row, col, numRows, numCols) {
      return makeRange(sheet, row, col, numRows, numCols);
    }
  };
  sheet.data[3] = headers.slice();
  return sheet;
}

function makeSettingsSheet(names) {
  var values = names.map(function (name) { return [name]; });
  while (values.length < 50) values.push(['']);
  return {
    getName: function () { return 'Settings'; },
    getRange: function (row, col, numRows, numCols) {
      if (row === 15 && col === 1 && numRows === 50 && numCols === 1) {
        return { getValues: function () { return values; } };
      }
      return { getValues: function () { return []; } };
    }
  };
}

const logValues = [];
function logRow(date, purchaser, amount, car) {
  return [date, '', purchaser, amount, car, '', ''];
}
logValues[0] = logRow(new Date(2026, 0, 1), 'A', 100, 'CAR1');
logValues[1] = logRow(new Date(2026, 0, 1), 'A', 100, 'CAR1');
logValues[2] = logRow(new Date(2026, 0, 2), '', 50, 'CAR2');
logValues[3] = logRow(new Date(2026, 0, 3), 'B', 50, '');
logValues[4] = logRow('', 'C', 30, 'CAR3');
logValues[5] = logRow(new Date(2027, 0, 1), 'B', 20, 'CAR2');
logValues[6] = ['', '', '', '', '', '', ''];
logValues[7] = logRow(new Date(2026, 0, 4), 'D', 0, 'CAR4');
logValues[8] = logRow(new Date(2026, 0, 6), 'E', 25, 'CAR5');
while (logValues.length < 5000) {
  logValues.push(['', '', '', '', '', '', '']);
}

const logSheet = makeLogSheet(logValues, { 8: true }, { 9: true });
const auditSheet = makeAuditSheet();
const settingsSheet = makeSettingsSheet(['A', 'B', 'C', 'D']);
const healthSpreadsheet = {
  getSheets: function () { return [logSheet]; },
  getSpreadsheetTimeZone: function () { return 'Africa/Accra'; },
  getSheetByName: function (name) {
    if (name === 'Audit Log') return auditSheet;
    if (name === 'Settings') return settingsSheet;
    return null;
  }
};
const auditSpreadsheet = {
  getSheets: function () { return [auditSheet]; },
  getSpreadsheetTimeZone: function () { return 'Africa/Accra'; },
  getSheetByName: function (name) {
    return name === 'Audit Log' ? auditSheet : null;
  }
};

const sandbox = {
  console: console,
  Date: Date,
  Math: Math,
  Number: Number,
  String: String,
  isFinite: isFinite,
  SpreadsheetApp: {
    getActive: function () { return auditSpreadsheet; },
    getActiveSpreadsheet: function () { return healthSpreadsheet; },
    BorderStyle: { SOLID: 'SOLID' }
  },
  Session: {
    getActiveUser: function () { return { getEmail: function () { return ''; } }; },
    getScriptTimeZone: function () { return 'Africa/Accra'; }
  },
  Utilities: { formatDate: mockFormatDate },
  LockService: {
    getScriptLock: function () {
      return {
        tryLock: function () { return true; },
        releaseLock: function () {}
      };
    }
  },
  HtmlService: {},
  PropertiesService: {}
};
sandbox.SpreadsheetApp.getUi = function () {
  return {
    alert: function () {},
    showModalDialog: function () {}
  };
};
sandbox.global = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Code.gs', 'utf8'), sandbox);

const findings = sandbox.scanLogHealth_(healthSpreadsheet);
function count(issue) {
  return findings.filter(function (finding) {
    return finding.issue === issue;
  }).length;
}

assert.strictEqual(count('Row gap'), 1, 'one row gap');
assert.strictEqual(count('Hidden rows'), 2, 'two hidden ranges');
assert.strictEqual(count('Possible duplicate'), 1, 'one duplicate group');
assert.strictEqual(count('Missing date'), 1, 'one missing date');
assert.strictEqual(count('Missing car number'), 1, 'one missing car');
assert.strictEqual(count('Missing purchaser'), 1, 'one missing purchaser');
assert.strictEqual(count('Missing amount'), 1, 'one missing amount');
assert.strictEqual(count('Date outside log year'), 1, 'one out-of-year date');
assert.strictEqual(count('Purchaser not in Settings'), 1, 'one unlisted purchaser');
assert.strictEqual(findings.length, 10, 'total findings');

const unlisted = findings.filter(function (finding) {
  return finding.issue === 'Purchaser not in Settings';
});
assert.strictEqual(unlisted[0].rows, '13');
assert.ok(unlisted[0].detail.indexOf('E') !== -1);

const duplicate = findings.find(function (finding) {
  return finding.issue === 'Possible duplicate';
});
assert.strictEqual(duplicate.rows, '5, 6');

const hidden = findings.filter(function (finding) {
  return finding.issue === 'Hidden rows';
});
assert.strictEqual(hidden.map(function (finding) { return finding.rows; }).join(','),
                   '8–8,9–9');

sandbox.auditOnChange({
  changeType: 'REMOVE_ROW',
  source: auditSpreadsheet,
  user: { getEmail: function () { return 'editor@example.com'; } }
});
const structural = auditSheet.data[4];
assert.ok(structural, 'structural audit row written');
assert.strictEqual(structural[2], 'Workbook');
assert.strictEqual(structural[8], 'REMOVE_ROW');
assert.strictEqual(structural[11], 'Structure change');

const rowsAfterRemove = auditSheet.data.length;
sandbox.auditOnChange({
  changeType: 'EDIT',
  source: auditSpreadsheet,
  user: { getEmail: function () { return 'editor@example.com'; } }
});
assert.strictEqual(auditSheet.data.length, rowsAfterRemove,
                   'ordinary edit is not logged by auditOnChange');

sandbox.auditOnChange({
  changeType: 'INSERT_COLUMN',
  source: auditSpreadsheet,
  user: { getEmail: function () { return 'editor@example.com'; } }
});
assert.strictEqual(auditSheet.data[5][8], 'INSERT_COLUMN');
assert.strictEqual(auditSheet.data[5][11], 'Structure change');

console.log('Audit health tests passed');
console.log('Findings: ' + findings.length);
