const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const EMPTY_ROW = ['', '', '', '', '', '', ''];

function makeRows() {
  const rows = [];
  for (let i = 0; i < 5000; i++) rows.push(EMPTY_ROW.slice());
  return rows;
}

function makeLogSheet(rows) {
  return {
    rows: rows,
    getName: function () {
      return '2026 Log';
    },
    getMaxRows: function () {
      return 5004;
    },
    getRange: function (row, col, numRows, numCols) {
      return {
        getValues: function () {
          const out = [];
          for (let r = 0; r < numRows; r++) {
            const source = rows[row - 5 + r] || EMPTY_ROW;
            out.push(source.slice(col - 1, col - 1 + numCols));
          }
          return out;
        },
        setValues: function (values) {
          for (let r = 0; r < values.length; r++) {
            const index = row - 5 + r;
            if (!rows[index]) rows[index] = EMPTY_ROW.slice();
            for (let c = 0; c < values[r].length; c++) {
              rows[index][col - 1 + c] = values[r][c];
            }
          }
        }
      };
    }
  };
}

function entry(date, purchaser, amount, car, email, enteredAt) {
  return [date, '', purchaser, amount, car, email, enteredAt];
}

function october(day) {
  return new Date(2026, 9, day, 12, 0, 0);
}

const sandbox = {
  console: console,
  Date: Date,
  Math: Math,
  Number: Number,
  String: String,
  isFinite: isFinite,
  isNaN: isNaN,
  SpreadsheetApp: {},
  Session: {},
  Utilities: {},
  LockService: {},
  HtmlService: {},
  PropertiesService: {}
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Code.gs', 'utf8'), sandbox);

const rows = makeRows();
rows[0] = entry(october(29), 'A', 100, 'CAR-A', 'a@example.com', october(29));
// Both 27 October rows have different times. The day column decides order, so
// the row entered first must stay first even though its time is later.
rows[1] = entry(new Date(2026, 9, 27, 18, 0, 0), 'B', 200, 'CAR-B',
                'b@example.com', october(27));
rows[2] = entry(october(28), 'C', 300, 'CAR-C', 'c@example.com', october(28));
// rows[3] stays blank: a row gap must be preserved.
rows[4] = entry(new Date(2026, 9, 27, 8, 0, 0), 'D', 400, 'CAR-D',
                'd@example.com', october(27));
rows[5] = ['', '', 'E', 500, 'CAR-E', '', ''];
rows[6] = entry(october(30), 'F', 600, 'CAR-F', 'f@example.com', october(30));
for (let row = 5; row <= 11; row++) rows[row - 5][1] = 'B' + row;

const sheet = makeLogSheet(rows);
const result = sandbox.sortLogByDate_(sheet);

assert.strictEqual(result.count, 5, 'five dated purchases');
assert.strictEqual(result.moved, 3, 'three rows moved');
assert.strictEqual(result.rowMap[5], 9, 'row 5 moved to row 9');
assert.strictEqual(result.rowMap[6], 5, 'row 6 moved to row 5');
assert.strictEqual(result.rowMap[7], 7, 'row 7 stayed');
assert.strictEqual(result.rowMap[9], 6, 'row 9 moved to row 6');
assert.strictEqual(result.rowMap[11], 11, 'row 11 stayed');

function rowAt(row) {
  return rows[row - 5];
}

assert.strictEqual(rowAt(5)[0].getDate(), 27, 'row 5 date');
assert.strictEqual(rowAt(5)[0].getHours(), 18, 'row 5 time travels with entry');
assert.strictEqual(rowAt(5)[2], 'B', 'row 5 purchaser');
assert.strictEqual(rowAt(5)[3], 200, 'row 5 amount');
assert.strictEqual(rowAt(5)[4], 'CAR-B', 'row 5 car');
assert.strictEqual(rowAt(5)[5], 'b@example.com', 'row 5 stamp');

assert.strictEqual(rowAt(6)[0].getDate(), 27, 'row 6 date');
assert.strictEqual(rowAt(6)[0].getHours(), 8, 'row 6 time travels with entry');
assert.strictEqual(rowAt(6)[2], 'D', 'same-day order is stable');
assert.strictEqual(rowAt(7)[0].getDate(), 28, 'row 7 date');
assert.strictEqual(rowAt(9)[0].getDate(), 29, 'row 9 date');
assert.strictEqual(rowAt(9)[2], 'A', 'row 9 purchaser');
assert.strictEqual(rowAt(11)[0].getDate(), 30, 'row 11 date');

assert.strictEqual(
  [rowAt(8)[0], rowAt(8)[2], rowAt(8)[3], rowAt(8)[4],
   rowAt(8)[5], rowAt(8)[6]].join(''),
  '',
  'row gap preserved'
);
assert.strictEqual(rowAt(10)[0], '', 'missing date stays in place');
assert.strictEqual(rowAt(10)[2], 'E', 'missing-date row data preserved');
assert.strictEqual(rowAt(10)[3], 500, 'missing-date amount preserved');
assert.strictEqual(rowAt(10)[4], 'CAR-E', 'missing-date car preserved');

for (let row = 5; row <= 11; row++) {
  assert.strictEqual(rowAt(row)[1], 'B' + row, 'day column untouched on row ' + row);
}

const second = sandbox.sortLogByDate_(sheet);
assert.strictEqual(second.count, 5, 'five dated purchases after sort');
assert.strictEqual(second.moved, 0, 'already-sorted log is not rewritten');
assert.strictEqual(rowAt(5)[2], 'B', 'already-sorted data unchanged');

const emptySheet = makeLogSheet(makeRows());
const empty = sandbox.sortLogByDate_(emptySheet);
assert.strictEqual(empty.count, 0, 'empty log has no dated purchases');
assert.strictEqual(empty.moved, 0, 'empty log is not rewritten');

const appendRows = makeRows();
appendRows[0] = entry(october(20), 'A', 100, 'CAR-A', '', '');
appendRows[1] = ['', '', 'B', 200, 'CAR-B', '', ''];
const appendSheet = makeLogSheet(appendRows);
assert.strictEqual(
  sandbox.findAppendRow_(appendSheet),
  7,
  'append row follows a row with data but no date'
);

assert.strictEqual(
  sandbox.isCompletePurchaseRow_(october(27), 'A', 100, 'CAR-A'),
  true,
  'complete purchase row'
);
assert.strictEqual(
  sandbox.isCompletePurchaseRow_(october(27), 'A', 100, ''),
  false,
  'row without a car is not complete'
);
assert.strictEqual(
  sandbox.isCompletePurchaseRow_(october(27), 'A', 0, 'CAR-A'),
  false,
  'zero amount is not complete'
);
assert.strictEqual(
  sandbox.isCompletePurchaseRow_('2026-10-27', 'A', 100, 'CAR-A'),
  false,
  'text date is not complete'
);

console.log('Sort log tests passed');
