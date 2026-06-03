// Self-contained API test: boots the server on a fresh port, exercises every
// endpoint via fetch, prints a report, then exits. No shell sleep / no lingering process.
process.env.PORT = '4173';
require('./server.js');

const BASE = 'http://localhost:4173';
const delay = (ms) => new Promise(r => setTimeout(r, ms));
const j = async (res) => ({ status: res.status, body: await res.json().catch(() => null) });
let pass = 0, fail = 0;
const ok = (cond, label, extra = '') => { (cond ? pass++ : fail++); console.log(`${cond ? 'PASS' : 'FAIL'}  ${label}${extra ? '  ' + extra : ''}`); };

(async () => {
    // wait for listen
    for (let i = 0; i < 40; i++) {
        try { const r = await fetch(BASE + '/api/categories'); if (r.ok) break; } catch (_) {}
        await delay(100);
    }

    // categories
    let { body: cats } = await j(await fetch(BASE + '/api/categories'));
    ok(Array.isArray(cats.categories) && cats.categories.length === 9, 'GET /api/categories returns 9 categories');

    // vehicles
    let { body: vehicles } = await j(await fetch(BASE + '/api/vehicles'));
    ok(Array.isArray(vehicles) && vehicles.length > 100, 'GET /api/vehicles', `count=${vehicles.length}`);

    // paginated items
    let t = Date.now();
    let { body: page } = await j(await fetch(BASE + '/api/items?page=1&limit=50'));
    ok(page.items.length === 50 && page.total === 2954 && page.totalPages === 60, 'GET /api/items paginated', `total=${page.total} in ${Date.now() - t}ms`);

    // category filter
    ({ body: page } = await j(await fetch(BASE + '/api/items?page=1&limit=10&category=Filters')));
    ok(page.items.every(i => i.category === 'Filters'), 'category filter = Filters', `total=${page.total}`);

    // vehicle + date range
    ({ body: page } = await j(await fetch(BASE + '/api/items?page=1&limit=50&vehicle=SL-11&startDate=2025-01-01&endDate=2026-12-31')));
    ok(page.items.every(i => i.vehicleMachinery === 'SL-11'), 'vehicle + date-range filter', `total=${page.total}`);

    // search across receipts
    ({ body: page } = await j(await fetch(BASE + '/api/items?page=1&limit=5&search=battery')));
    ok(page.total > 0, 'free-text search "battery"', `total=${page.total}`);

    // CREATE item -> auto category
    let { body: created } = await j(await fetch(BASE + '/api/items', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ mrnNum: 'TEST-001', reqDate: '2026-06-03', vehicleMachinery: 'TEST-VH', itemName: '150 Amp Battery', itemDesc: 'x', reqQty: 5 }) }));
    ok(created.success && created.id && created.category === 'Battery', 'POST /api/items auto-classifies Battery', `id=${created.id}`);
    const itemId = created.id;

    // UPDATE item with manual category override
    let { body: upd } = await j(await fetch(BASE + '/api/items/' + itemId, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ mrnNum: 'TEST-001', reqDate: '2026-06-03', vehicleMachinery: 'TEST-VH', itemName: '150 Amp Battery', itemDesc: 'x', reqQty: 8, category: 'Electrical' }) }));
    ok(upd.success && upd.category === 'Electrical', 'PUT /api/items manual category override');

    // add receipt (receive)
    let { body: rec } = await j(await fetch(BASE + `/api/items/${itemId}/receipts`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ qty: 3, transactionType: 'Receive', deliveryDate: '2026-06-04', purchaseSource: 'Local Store' }) }));
    ok(rec.success && rec.id, 'POST receipt (update receive)', `recId=${rec.id}`);
    const recId = rec.id;

    // update GRN/pricing on receipt
    let { body: grn } = await j(await fetch(BASE + '/api/receipts/' + recId, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ grnNumber: 'GRN-99', invoiceNumber: 'INV-77', supplierName: 'ACME', unitPrice: 1200 }) }));
    ok(grn.success, 'PUT receipt (update GRN/pricing)');

    // verify recQty + receipt attached
    ({ body: page } = await j(await fetch(BASE + '/api/items?page=1&limit=1&search=TEST-001')));
    let it = page.items[0];
    ok(it.recQty === 3 && it.receipts.length === 1 && it.receipts[0].grnNumber === 'GRN-99', 'recQty computed + GRN saved', `recQty=${it.recQty}`);

    // ISSUES CRUD
    let { body: iss } = await j(await fetch(BASE + '/api/issues', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ issueDate: '2026-06-05', vehicleMachinery: 'TEST-VH', itemName: 'Air Filter', qty: 2, issuedTo: 'Site A', issuedBy: 'Store' }) }));
    ok(iss.success && iss.id && iss.category === 'Filters', 'POST /api/issues auto-classifies', `id=${iss.id}`);
    const issId = iss.id;
    let { body: issList } = await j(await fetch(BASE + '/api/issues?vehicle=TEST-VH'));
    ok(Array.isArray(issList) && issList.length === 1, 'GET /api/issues vehicle filter');
    await j(await fetch(BASE + '/api/issues/' + issId, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ issueDate: '2026-06-05', vehicleMachinery: 'TEST-VH', itemName: 'Air Filter', qty: 9, issuedTo: 'Site B', issuedBy: 'Store' }) }));
    ({ body: issList } = await j(await fetch(BASE + '/api/issues?vehicle=TEST-VH')));
    ok(issList[0].qty === 9 && issList[0].issuedTo === 'Site B', 'PUT /api/issues updates');

    // DELETE everything we created
    let { body: dIss } = await j(await fetch(BASE + '/api/issues/' + issId, { method: 'DELETE' }));
    let { body: dRec } = await j(await fetch(BASE + '/api/receipts/' + recId, { method: 'DELETE' }));
    let { body: dItem } = await j(await fetch(BASE + '/api/items/' + itemId, { method: 'DELETE' }));
    ok(dIss.success && dRec.success && dItem.success, 'DELETE issue/receipt/item');

    // confirm cleanup
    ({ body: page } = await j(await fetch(BASE + '/api/items?page=1&limit=1&search=TEST-001')));
    ok(page.total === 0, 'cleanup verified (item gone)');

    console.log(`\n${pass} passed, ${fail} failed`);
    process.exit(fail ? 1 : 0);
})().catch(e => { console.error('TEST ERROR:', e); process.exit(1); });
