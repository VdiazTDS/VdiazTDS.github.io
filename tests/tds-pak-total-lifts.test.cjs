const assert = require('node:assert/strict');
const { test } = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../app-tds-pak.js'), 'utf8');
function loadFunctions(overrides = {}) {
  const context = vm.createContext({ window: {}, ...overrides });
  for (const name of [
    'normalizeDayToken', 'getMappedRowRoute', 'getMappedRowDayRaw',
    'normalizeSummaryHeaderValue', 'findSummaryHeaderByAliases',
    'parseSummaryRouteDayCombinedValue', 'getSummaryRouteDayFields',
    'getSummaryRouteDayFromRow', 'isSummaryAggregateLabel', 'isSummaryAggregateRow',
    'addSummaryTotalLifts', 'getSummarySourceRouteRows'
  ]) {
    const start = source.search(new RegExp(`^(?:async )?function ${name}\\(`, 'm'));
    assert.notEqual(start, -1);
    const end = source.indexOf('\n}', start) + 2;
    vm.runInContext(source.slice(start, end), context);
  }
  return context;
}

test('sums all records by route and normalized day, including records without coordinates', () => {
  const { addSummaryTotalLifts } = loadFunctions();
  const rows = [{ Route: ' A ', Day: 'Monday' }, { Route: 'A', Day: 'Tuesday' }, { Route: 'B', Day: 1 }];
  const result = addSummaryTotalLifts(rows, ['Route', 'Day'], [
    { ROUTE: 'a', DAY: 1, QTY: 2 },
    { ROUTE: 'A', DAY: 'Mon', QTY: '1,200.5' },
    { ROUTE: 'A', DAY: 2, QTY: 7 },
    { NEWROUTE: 'B', NEWDAY: 'Monday', QTY: 4 },
    { ROUTE: 'A', DAY: 1, QTY: '' },
    { ROUTE: 'A', DAY: 1, QTY: 'invalid' }
  ]);
  assert.deepEqual(Array.from(result.rows, row => row['Total Lifts']), [1202.5, 7, 4]);
  assert.equal(result.headers.at(-1), 'Total Lifts');
  assert.equal(rows[0]['Total Lifts'], undefined);
});

test('supports combined route/day headers and replaces existing Total Lifts', () => {
  const { addSummaryTotalLifts } = loadFunctions();
  const result = addSummaryTotalLifts(
    [{ 'Route + Day': 'A | Monday', 'Total Lifts': 999 }],
    ['Route + Day', 'Total Lifts'], [{ ROUTE: 'A', DAY: 1, QTY: 3 }]
  );
  assert.equal(result.rows[0]['Total Lifts'], 3);
  assert.equal(result.headers.filter(header => header === 'Total Lifts').length, 1);
});

test('distinguishes zero from missing data and leaves aggregate rows blank', () => {
  const { addSummaryTotalLifts } = loadFunctions();
  const result = addSummaryTotalLifts(
    ['A', 'B', 'C', 'Grand Total'].map(Route => ({ Route, Day: 1 })),
    ['Route', 'Day'], [{ ROUTE: 'A', DAY: 1, QTY: 0 }, { ROUTE: 'B', DAY: 1 }]
  );
  assert.deepEqual(Array.from(result.rows, row => row['Total Lifts']), [0, '', '', '']);
});

test('uses loaded rows only when they belong to the summary route file', async () => {
  const currentRows = [{ ROUTE: 'A', DAY: 1, QTY: 10 }];
  const context = loadFunctions({ window: { _currentFilePath: 'correct.xlsx', _currentRows: currentRows } });
  assert.equal(await context.getSummarySourceRouteRows('summary.xlsx', 'correct.xlsx'), currentRows);
});

test('standalone summary fetches its associated workbook instead of another open map', async () => {
  const fetched = [];
  const sourceRows = [{ CustomQuantity: 8 }];
  const context = loadFunctions({
    window: { _currentFilePath: 'other.xlsx', _currentRows: [{ QTY: 999 }] },
    BUCKET: 'routes',
    ensureSummaryAttachmentsInitialized: async () => {},
    ensureFileColumnMappingsInitialized: async () => {},
    isRouteSummaryFileName: name => name === 'summary.xlsx',
    isSystemCloudFileName: () => false,
    resolveSummaryForRoute: name => name === 'correct.xlsx' ? 'summary.xlsx' : null,
    sb: { storage: { from: () => ({
      list: async () => ({ data: [{ name: 'other.xlsx' }, { name: 'correct.xlsx' }, { name: 'summary.xlsx' }] }),
      getPublicUrl: name => ({ data: { publicUrl: `https://example.test/${name}` } })
    }) } },
    fetch: async url => { fetched.push(url); return { ok: true, arrayBuffer: async () => new ArrayBuffer(0) }; },
    XLSX: { read: () => ({ SheetNames: ['Sheet1'], Sheets: { Sheet1: {} } }), utils: { sheet_to_json: () => sourceRows } },
    collectColumnMappingHeaders: () => ['CustomQuantity'],
    buildInitialColumnMapping: (headers, name) => { assert.equal(name, 'correct.xlsx'); return { QTY: 'CustomQuantity' }; },
    applyColumnAliasesToRows: (rows, mapping) => { rows[0].QTY = rows[0][mapping.QTY]; }
  });
  const result = await context.getSummarySourceRouteRows('summary.xlsx');
  assert.equal(result[0].QTY, 8);
  assert.match(fetched[0], /\/correct\.xlsx\?v=/);
  assert.equal(context.window._currentFilePath, 'other.xlsx');
});
