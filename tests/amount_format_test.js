const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const sandbox = {
  console: console,
  Date: Date,
  Math: Math,
  Number: Number,
  String: String,
  isFinite: isFinite,
  isNaN: isNaN
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Code.gs', 'utf8'), sandbox);

function makeSheet(name, entries) {
  const data = {};
  const formulas = {};
  const sheet = {
    formatCalls: [],
    bulkWrites: 0,
    singleWrites: [],
    getName: function () { return name; },
    getMaxRows: function () { return 5004; },
    getRange: function (row, col, numRows, numCols) {
      assert.strictEqual(col, 4, 'amount formatting only touches column D');
      const height = numRows === undefined ? 1 : numRows;
      const width = numCols === undefined ? 1 : numCols;
      assert.strictEqual(width, 1, 'amount formatting uses a single column');

      return {
        getValues: function () {
          const out = [];
          for (let i = 0; i < height; i++) {
            out.push([data[row + i] === undefined ? '' : data[row + i]]);
          }
          return out;
        },
        getFormulas: function () {
          const out = [];
          for (let i = 0; i < height; i++) {
            out.push([formulas[row + i] || '']);
          }
          return out;
        },
        setValues: function (values) {
          sheet.bulkWrites++;
          for (let i = 0; i < values.length; i++) {
            data[row + i] = values[i][0];
          }
          return this;
        },
        setValue: function (value) {
          sheet.singleWrites.push({ row: row, value: value });
          data[row] = value;
          return this;
        },
        setNumberFormat: function (format) {
          sheet.formatCalls.push({ row: row, numRows: height, format: format });
          return this;
        }
      };
    }
  };

  entries.forEach(function (entry) {
    data[entry[0]] = entry[1];
    if (entry[2]) formulas[entry[0]] = entry[2];
  });

  return sheet;
}

assert.strictEqual(sandbox.normalizeAmountValue_('GH₵1,200.50'), 1200.5);
assert.strictEqual(sandbox.normalizeAmountValue_('GHS 400'), 400);
assert.strictEqual(sandbox.normalizeAmountValue_('420.00'), 420);
assert.strictEqual(sandbox.normalizeAmountValue_('GH¢ 75.25'), 75.25);
assert.strictEqual(sandbox.normalizeAmountValue_('fuel note'), 'fuel note');
assert.strictEqual(sandbox.normalizeAmountValue_(500), 500);
assert.strictEqual(sandbox.normalizeAmountValue_(''), '');

const sheet = makeSheet('2026 Log', [
  [5, 'GH₵1,200.50'],
  [6, 500],
  [7, 'fuel note'],
  [9, '420.5']
]);

assert.strictEqual(sandbox.formatAmountCells_(sheet, 5, 5), 5);
assert.deepStrictEqual(sheet.getRange(5, 4, 5, 1).getValues(), [
  [1200.5], [500], ['fuel note'], [''], [420.5]
]);
assert.strictEqual(sheet.bulkWrites, 1, 'numeric text is written back once');
assert.strictEqual(sheet.formatCalls.length, 1, 'one format call');
assert.strictEqual(sheet.formatCalls[0].row, 5);
assert.strictEqual(sheet.formatCalls[0].numRows, 5);
assert.strictEqual(sheet.formatCalls[0].format, sandbox.CURRENCY_FORMAT);

const formulaSheet = makeSheet('2026 Log', [
  [5, 100, '=50+50'],
  [6, 'GHS 75.00']
]);
assert.strictEqual(sandbox.formatAmountCells_(formulaSheet, 5, 2), 2);
assert.strictEqual(
  formulaSheet.getRange(5, 4, 2, 1).getFormulas()[0][0],
  '=50+50',
  'formula cell is not replaced'
);
assert.strictEqual(formulaSheet.bulkWrites, 0,
                   'bulk write skipped when formulas are present');
assert.deepStrictEqual(formulaSheet.singleWrites,
                       [{ row: 6, value: 75 }],
                       'only the changed non-formula cell is written');
assert.strictEqual(formulaSheet.formatCalls.length, 1,
                   'formatted range still includes the formula cell');

const bounded = makeSheet('2026 Log', []);
assert.strictEqual(sandbox.formatAmountCells_(bounded, 5003, 10), 2);
assert.strictEqual(bounded.formatCalls[0].numRows, 2,
                   'formatting is capped at the last log row');
assert.strictEqual(sandbox.formatAmountCells_(bounded, 5005, 10), 0);
assert.strictEqual(bounded.formatCalls.length, 1,
                   'rows below the log are not formatted');

const all = {
  getSheets: function () {
    return [
      makeSheet('2026 Log', []),
      makeSheet('2027 Log', []),
      makeSheet('Settings', [])
    ];
  }
};
assert.strictEqual(sandbox.formatAllAmountColumns_(all), 2,
                   'both year logs are formatted');

// The installed edit trigger must format the edited amount cell before the
// audit row is written.
const DAY_FORMULA = '=IF($A5="","",TEXT($A5,"dddd"))';
const triggerValues = {
  5: [new Date(2026, 9, 29), DAY_FORMULA, 'A', 'GH₵1,200.50', '',
      'editor@example.com', new Date(2026, 9, 29, 10, 0, 0)]
};
const triggerFormatCalls = [];
const triggerSheet = {
  getName: function () { return '2026 Log'; },
  getParent: function () { return spreadsheet; },
  getMaxRows: function () { return 5004; },
  getRange: function (row, col, numRows, numCols) {
    if (col === 4 && numRows !== undefined) {
      const height = numRows;
      return {
        getFormulas: function () {
          return Array.from({ length: height }, function () { return ['']; });
        },
        getValues: function () {
          const out = [];
          for (let i = 0; i < height; i++) out.push([triggerValues[row + i][3]]);
          return out;
        },
        setValues: function (values) {
          for (let i = 0; i < values.length; i++) {
            triggerValues[row + i][3] = values[i][0];
          }
          return this;
        },
        setNumberFormat: function (format) {
          triggerFormatCalls.push({ row: row, numRows: height, format: format });
          return this;
        }
      };
    }

    const height = numRows === undefined ? 1 : numRows;
    const width = numCols === undefined ? 1 : numCols;
    return {
      getValue: function () {
        return triggerValues[row] ? triggerValues[row][col - 1] : '';
      },
      setValue: function (value) {
        if (!triggerValues[row]) triggerValues[row] = [];
        triggerValues[row][col - 1] = value;
        return this;
      },
      getFormula: function () {
        return triggerValues[row] ? triggerValues[row][col - 1] : '';
      },
      setFormula: function (formula) {
        triggerValues[row][col - 1] = formula;
        return this;
      },
      getValues: function () {
        const out = [];
        for (let i = 0; i < height; i++) {
          const source = triggerValues[row + i] || [];
          out.push(source.slice(col - 1, col - 1 + width));
        }
        return out;
      },
      setValues: function (values) {
        for (let i = 0; i < values.length; i++) {
          if (!triggerValues[row + i]) triggerValues[row + i] = [];
          for (let c = 0; c < values[i].length; c++) {
            triggerValues[row + i][col - 1 + c] = values[i][c];
          }
        }
        return this;
      },
      setNumberFormats: function () { return this; }
    };
  }
};

const triggerAuditRows = [];
const triggerAuditSheet = {
  getLastRow: function () { return 4 + triggerAuditRows.length; },
  getRange: function (row) {
    return {
      getValues: function () {
        return [['Timestamp', 'User', 'Sheet', 'Row', 'Purchase date',
                 'Purchaser', 'Amount (GH₵)', 'Car Number', 'Field',
                 'Old value', 'New value', 'Action']];
      },
      setValues: function (rows) {
        for (let i = 0; i < rows.length; i++) triggerAuditRows[row - 5 + i] = rows[i];
        return this;
      },
      setFontSize: function () { return this; },
      setBorder: function () { return this; },
      setNumberFormat: function () { return this; }
    };
  }
};
const spreadsheet = {
  getSheetByName: function (name) {
    return name === 'Audit Log' ? triggerAuditSheet : null;
  }
};

sandbox.SpreadsheetApp = { BorderStyle: { SOLID: 'SOLID' } };
sandbox.LockService = {
  getScriptLock: function () {
    return {
      tryLock: function () { return true; },
      releaseLock: function () {}
    };
  }
};

sandbox.auditOnEdit({
  range: {
    getSheet: function () { return triggerSheet; },
    getRow: function () { return 5; },
    getLastRow: function () { return 5; },
    getColumn: function () { return 4; },
    getLastColumn: function () { return 4; },
    getNumRows: function () { return 1; },
    getNumColumns: function () { return 1; },
    getNumCells: function () { return 1; },
    getValues: function () {
      return [[triggerSheet.getRange(5, 4).getValue()]];
    }
  },
  oldValue: '',
  user: { getEmail: function () { return 'editor@example.com'; } }
});

assert.strictEqual(triggerFormatCalls.length, 1,
                   'edit trigger formats the amount cell');
assert.strictEqual(triggerFormatCalls[0].format, sandbox.CURRENCY_FORMAT);
assert.strictEqual(triggerSheet.getRange(5, 4).getValue(), 1200.5,
                   'currency text is stored as a number');
assert.strictEqual(triggerAuditRows.length, 1, 'edit is still audited');
assert.strictEqual(triggerAuditRows[0][6], 1200.5, 'audit captures the value');
assert.strictEqual(triggerAuditRows[0][8], 'Amount', 'audit names the field');
assert.strictEqual(triggerAuditRows[0][11], 'Added', 'audit action is Added');

console.log('Amount format tests passed');
