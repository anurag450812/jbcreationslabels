/**
 * Standalone checks for STOCK COUNT matching.
 * Run: node test/stock-sku-match.test.js
 *
 * The old deductStock matcher lowercased and trimmed, then used
 * cell.includes(label) and kept the first row. That is reproduced below
 * only to show the bug. Production matching is exact after the same
 * normalization on both sides: trim, lowercase, strip one leading "jb-".
 */

const assert = require('assert');
const {
    normalizeStockSku,
    findExactStockRowIndexes,
    planStockDeductions
} = require('../google-apps-script.js');

let passed = 0;

function test(name, fn) {
    fn();
    passed += 1;
    console.log('ok - ' + name);
}

function sheet(rows) {
    return [['SKU', 'Qty']].concat(rows);
}

function legacyFirstIncludesMatch(rows, label) {
    const searchLabel = String(label).toLowerCase().trim();
    for (let i = 1; i < rows.length; i++) {
        const cellLabel = String(rows[i][0]).toLowerCase().trim();
        if (cellLabel.includes(searchLabel)) return i;
    }
    return -1;
}

function quantitiesAfter(rows, labelCounts) {
    const snapshot = rows.map((row) => row.slice());
    const plan = planStockDeductions(rows, labelCounts);
    assert.deepStrictEqual(rows, snapshot, 'planner must not mutate the sheet values');
    const next = snapshot.map((row) => row.slice());
    plan.updates.forEach((update) => {
        next[update.rowNumber - 1][1] = update.newStock;
    });
    return { plan, rows: next };
}

test('normalization trims, lowercases, and strips one leading jb- prefix', () => {
    assert.strictEqual(normalizeStockSku('  JB-BD10  '), 'bd10');
    assert.strictEqual(normalizeStockSku('sv28'), 'sv28');
    assert.strictEqual(normalizeStockSku('jb-jb-bd10'), 'jb-bd10');
    assert.strictEqual(normalizeStockSku('0baby_1,2,11,12'), '0baby_1,2,11,12');
    assert.strictEqual(normalizeStockSku(null), '');
});

test('old includes() match deducts bd10 from jb-bd100 when that row is first', () => {
    const rows = sheet([
        ['jb-bd100', 40],
        ['jb-bd10', 8]
    ]);
    assert.strictEqual(legacyFirstIncludesMatch(rows, 'bd10'), 1);
    assert.strictEqual(String(rows[1][0]), 'jb-bd100');
});

test('bd10 matches only jb-bd10, never jb-bd100, even when bd100 is first', () => {
    const rows = sheet([
        ['jb-bd100', 40],
        ['jb-bd10', 8],
        ['jb-bd101', 3]
    ]);
    const { plan, rows: next } = quantitiesAfter(rows, { bd10: 2 });
    assert.deepStrictEqual(findExactStockRowIndexes(rows, 'bd10'), [2]);
    assert.strictEqual(plan.errors.length, 0);
    assert.strictEqual(plan.updates.length, 1);
    assert.strictEqual(plan.updates[0].matchedCell, 'jb-bd10');
    assert.strictEqual(plan.updates[0].previousStock, 8);
    assert.strictEqual(plan.updates[0].newStock, 6);
    assert.strictEqual(plan.updates[0].rowNumber, 3);
    assert.strictEqual(next[1][1], 40);
    assert.strictEqual(next[2][1], 6);
    assert.strictEqual(next[3][1], 3);
});

test('bd100 does not match jb-bd10, and nearby prefixes do not match either', () => {
    const rows = sheet([
        ['jb-bd10', 8],
        ['jb-bd100', 40],
        ['jb-fr-bd105', 2],
        ['jb-ch-sv28', 6],
        ['jb-sv28', 9]
    ]);
    const bd100 = quantitiesAfter(rows, { bd100: 1 });
    assert.strictEqual(bd100.plan.updates[0].matchedCell, 'jb-bd100');
    assert.strictEqual(bd100.rows[1][1], 8);
    assert.strictEqual(bd100.rows[2][1], 39);

    const sv28 = quantitiesAfter(rows, { sv28: 1 });
    assert.deepStrictEqual(sv28.plan.updates.map((update) => update.matchedCell), ['jb-sv28']);
    assert.strictEqual(sv28.rows[4][1], 6);
    assert.strictEqual(sv28.rows[5][1], 8);

    const ch = quantitiesAfter(rows, { 'ch-sv28': 4 });
    assert.strictEqual(ch.plan.updates[0].matchedCell, 'jb-ch-sv28');
    assert.strictEqual(ch.rows[4][1], 2);
    assert.strictEqual(ch.rows[5][1], 9);
});

test('case and surrounding whitespace still match the exact SKU', () => {
    const rows = sheet([
        ['  JB-Sv28 ', 12]
    ]);
    const { plan } = quantitiesAfter(rows, { '  sv28  ': 2 });
    assert.strictEqual(plan.errors.length, 0);
    assert.strictEqual(plan.updates[0].matchedCell, '  JB-Sv28 ');
    assert.strictEqual(plan.updates[0].newStock, 10);
    assert.strictEqual(plan.updates[0].rowNumber, 2);
});

test('combo SKUs match only the full normalized string', () => {
    const combo = '0baby_1,2,11,12';
    const rows = sheet([
        ['jb-0baby_1', 5],
        ['jb-0baby_1,2,11', 6],
        ['jb-0baby_1,2,11,12', 7],
        ['jb-0baby_1,2,11,12,13', 8]
    ]);
    const { plan, rows: next } = quantitiesAfter(rows, { [combo]: 1 });
    assert.strictEqual(plan.updates.length, 1);
    assert.strictEqual(plan.updates[0].matchedCell, 'jb-0baby_1,2,11,12');
    assert.strictEqual(next[1][1], 5);
    assert.strictEqual(next[2][1], 6);
    assert.strictEqual(next[3][1], 6);
    assert.strictEqual(next[4][1], 8);

    const shorter = quantitiesAfter(rows, { '0baby_1': 1 });
    assert.strictEqual(shorter.plan.updates[0].matchedCell, 'jb-0baby_1');
    assert.strictEqual(shorter.rows[3][1], 7);

    const mixedCase = quantitiesAfter(rows, { 'JB-0Baby_1,2,11,12': 2 });
    assert.strictEqual(mixedCase.plan.updates[0].newStock, 5);
});

test('a missing SKU deducts nothing and is reported as not found', () => {
    const rows = sheet([
        ['jb-bd10', 8],
        ['', 4]
    ]);
    const { plan, rows: next } = quantitiesAfter(rows, { zz99: 3 });
    assert.strictEqual(plan.updates.length, 0);
    assert.strictEqual(plan.errors.length, 1);
    assert.strictEqual(plan.errors[0].code, 'not_found');
    assert.strictEqual(plan.errors[0].label, 'zz99');
    assert.strictEqual(next[1][1], 8);
    assert.strictEqual(next[2][1], 4);
});

test('duplicate exact rows deduct nothing and name every matching row', () => {
    const rows = sheet([
        ['jb-bd10', 8],
        ['jb-bd100', 40],
        ['  BD10 ', 2]
    ]);
    const { plan, rows: next } = quantitiesAfter(rows, { bd10: 5 });
    assert.strictEqual(plan.updates.length, 0);
    assert.strictEqual(plan.errors.length, 1);
    assert.strictEqual(plan.errors[0].code, 'duplicate');
    assert.ok(plan.errors[0].message.indexOf('rows 2, 4') !== -1);
    assert.strictEqual(next[1][1], 8);
    assert.strictEqual(next[2][1], 40);
    assert.strictEqual(next[3][1], 2);
});

test('stock never goes below 0, and a second alias of the same row deducts from the new quantity', () => {
    const floored = quantitiesAfter(sheet([['jb-bd10', 1]]), { bd10: 5 });
    assert.strictEqual(floored.plan.updates[0].previousStock, 1);
    assert.strictEqual(floored.plan.updates[0].newStock, 0);

    const blankQty = quantitiesAfter(sheet([['jb-sv28', '']]), { sv28: 1 });
    assert.strictEqual(blankQty.plan.updates[0].newStock, 0);

    const shared = quantitiesAfter(sheet([['jb-bd10', 10]]), { bd10: 2, 'JB-bd10': 1 });
    assert.strictEqual(shared.plan.errors.length, 0);
    assert.strictEqual(shared.plan.updates[0].newStock, 8);
    assert.strictEqual(shared.plan.updates[1].previousStock, 8);
    assert.strictEqual(shared.plan.updates[1].newStock, 7);
    assert.strictEqual(shared.rows[1][1], 7);
});

test('a mixed batch updates only the unique exact match', () => {
    const rows = sheet([
        ['jb-bd100', 40],
        ['jb-bd10', 8],
        ['jb-bd10', 3],
        ['jb-0baby_1,2,11,12', 15]
    ]);
    const { plan, rows: next } = quantitiesAfter(rows, {
        bd10: 1,
        bd100: 4,
        '0baby_1,2,11,12': 5,
        nope: 9
    });
    const byLabel = {};
    plan.updates.forEach((update) => {
        byLabel[update.label] = update;
    });
    const errorsByLabel = {};
    plan.errors.forEach((error) => {
        errorsByLabel[error.label] = error;
    });

    assert.strictEqual(byLabel.bd10, undefined);
    assert.strictEqual(errorsByLabel.bd10.code, 'duplicate');
    assert.strictEqual(byLabel.bd100.newStock, 36);
    assert.strictEqual(byLabel['0baby_1,2,11,12'].matchedCell, 'jb-0baby_1,2,11,12');
    assert.strictEqual(byLabel['0baby_1,2,11,12'].newStock, 10);
    assert.strictEqual(errorsByLabel.nope.code, 'not_found');
    assert.strictEqual(next[1][1], 36);
    assert.strictEqual(next[2][1], 8);
    assert.strictEqual(next[3][1], 3);
    assert.strictEqual(next[4][1], 10);
});

test('the header row is not a stock match', () => {
    const rows = sheet([['jb-bd18', 4]]);
    rows[0] = ['jb-bd10', 'Qty'];
    const { plan } = quantitiesAfter(rows, { bd10: 1 });
    assert.strictEqual(plan.updates.length, 0);
    assert.strictEqual(plan.errors[0].code, 'not_found');
});

console.log(passed + ' tests passed');
