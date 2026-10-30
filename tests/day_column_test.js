const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const LOG_LAST_ROW = 5004;
const ENTRY_FIRST_ROW = 5;

function formulaFor(row) {
  return '=IF($A' + row + '="","",TEXT($A' + row + ',"dddd"))';
}

function makeLogSheet() {
  const formulas = [];
  for (let row = ENTRY_FIRST_ROW; row <= LOG_LAST_ROW; row++) {
    formulas.push([formulaFor(row)]);
  }

  const sheet = {
    formulas: formulas,
    getName: function () {
      return '2026 Log';
    },
    getMaxRows: function () {
      return LOG_LAST_ROW;
    },
    getRange: function (row, col, numRows, numCols) {
      if (col === 2 && row === ENTRY_FIRST_ROW &&
          numRows === LOG_LAST_ROW - ENTRY_FIRST_ROW + 1 && numCols === 1) {
        return {
          getFormulas: function () {
            return formulas.map(function (value) { return value.slice(); });
          },
          setFormulas: function (values) {
            for (let i = 0; i < values.length; i++) {
              formulas[i] = values[i].slice();
            }
          }
        };
      }
      if (col === 2 && (numRows === undefined || numRows === 1) &&
          (numCols === undefined || numCols === 1)) {
        return {
          getFormula: function () {
            return formulas[row - ENTRY_FIRST_ROW][0];
          },
          setFormula: function (formula) {
            formulas[row - ENTRY_FIRST_ROW][0] = formula;
          }
        };
      }
      throw new Error('Unexpected range ' + row + ',' + col + ',' +
                      numRows + ',' + numCols);
    }
  };
  return sheet;
}

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

assert.strictEqual(
  sandbox.dayFormula_(5),
  formulaFor(5),
  'day formula uses the absolute date reference'
);

const sheet = makeLogSheet();
sheet.formulas[0] = [formulaFor(5)];          // row 5: already correct
sheet.formulas[1] = [''];                     // row 6: blank
sheet.formulas[2] = [''];                     // row 7: static value
sheet.formulas[3] = ['=TEXT($A8,"dddd")'];    // row 8: wrong formula
sheet.formulas[4] = [formulaFor(9)];          // row 9: already correct

const repaired = sandbox.fillDayFormulas_(sheet);
assert.strictEqual(repaired, 3, 'three Day cells repaired');
assert.strictEqual(sheet.formulas[1][0], formulaFor(6), 'blank formula filled');
assert.strictEqual(sheet.formulas[2][0], formulaFor(7), 'static value replaced');
assert.strictEqual(sheet.formulas[3][0], formulaFor(8), 'wrong formula replaced');
assert.strictEqual(sheet.formulas[4][0], formulaFor(9), 'correct formula kept');

const second = sandbox.fillDayFormulas_(sheet);
assert.strictEqual(second, 0, 'complete Day column needs no repair');

sandbox.ensureDayFormula_(sheet, 10);
assert.strictEqual(sheet.formulas[5][0], formulaFor(10),
                   'single row Day formula written');

sandbox.ensureDayFormula_(sheet, 5);
assert.strictEqual(sheet.formulas[0][0], formulaFor(5),
                   'correct single row formula kept');

console.log('Day column tests passed');
