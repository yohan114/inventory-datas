const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const db = require('./db');

const app = express();
const PORT = 4000;

app.use(express.json({ limit: '100mb' }));
app.use(express.static(__dirname));

const DB_PATH = path.join(__dirname, 'inventory.db');

const classifyItem = (name) => {
    if (!name) return 'General Items';
    const lower = name.toLowerCase();
    
    if (lower.includes('filter') || lower.includes('cleaner') || lower.includes('element')) {
        return 'Filters';
    }
    if (lower.includes('battery') || lower.includes('batt') || lower.includes('accumulator') || lower.includes('battey')) {
        return 'Battery';
    }
    if (lower.includes('tyre') || lower.includes('tire') || lower.includes('tube') || lower.includes('flap')) {
        return 'Tyre';
    }
    if (lower.includes('oil') || lower.includes('grease') || lower.includes('lubricant') || lower.includes('coolant') || lower.includes('fluid') || lower.includes('petrol') || lower.includes('diesel')) {
        return 'Oil & Lubricants';
    }
    if (lower.includes('bearing') || lower.includes('seal') || lower.includes('gasket') || lower.includes('o-ring') || lower.includes('o ring') || lower.includes('bush') || lower.includes('spacer') || lower.includes('shim')) {
        return 'Bearings & Seals';
    }
    if (lower.includes('hydraulic') || lower.includes('hose') || lower.includes('coupling') || lower.includes('cylinder') || lower.includes('fittings') || lower.includes('adapter') || lower.includes('valve') || lower.includes('pump')) {
        return 'Hydraulics';
    }
    if (lower.includes('cable') || lower.includes('switch') || lower.includes('light') || lower.includes('bulb') || lower.includes('wire') || lower.includes('harness') || lower.includes('solenoid') || lower.includes('starter') || lower.includes('alternator') || lower.includes('dynamo') || lower.includes('sensor') || lower.includes('relay') || lower.includes('fuse') || lower.includes('horn') || lower.includes('meter')) {
        return 'Electrical';
    }
    if (lower.includes('belt') || lower.includes('v-belt') || lower.includes('v belt')) {
        return 'Belts';
    }
    return 'General Items';
};

// Heuristic PDF Parser logic
function parsePdfTextHeuristically(text) {
    const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
    
    let mrnNum = "";
    let reqDate = new Date().toISOString().split('T')[0]; // Default to today
    let vehicleMachinery = "";
    let itemName = "";
    let itemDesc = "";
    let reqQty = 1;
    let supplierName = "";
    let invoiceNumber = "";
    let invoiceDate = "";
    let unitPrice = null;
    let grnNumber = "";

    // 1. Extract MRN Number (e.g. MRN-001, Requisition 123)
    const mrnRegex = /(?:mrn|requisition|req)(?:\s*number|\s*no\.?)?[\s:-]*([a-z0-9-]+)/i;
    const mrnMatch = text.match(mrnRegex);
    if (mrnMatch) {
        mrnNum = mrnMatch[1].trim().toUpperCase();
    } else {
        const standaloneMrn = text.match(/\b(mrn-[0-9a-z-]+)\b/i);
        if (standaloneMrn) {
            mrnNum = standaloneMrn[1].toUpperCase();
        }
    }

    // 2. Extract Date
    const dateRegex = /\b(\d{4}[-/]\d{1,2}[-/]\d{1,2}|d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b/;
    const dateMatch = text.match(dateRegex);
    if (dateMatch) {
        const rawDate = dateMatch[1];
        try {
            const d = new Date(rawDate);
            if (!isNaN(d.getTime())) {
                reqDate = d.toISOString().split('T')[0];
            }
        } catch(e) {}
    }

    // 3. Extract Vehicle/Machinery
    const vehicleRegex = /(?:vehicle|machinery|unit|fleet|eq|equip)(?:\s*number|\s*no\.?)?[\s:-]+([^\n,]+)/i;
    const vehicleMatch = text.match(vehicleRegex);
    if (vehicleMatch) {
        vehicleMachinery = vehicleMatch[1].trim();
    } else {
        const keywords = ['excavator', 'truck', 'car', 'dumper', 'digger', 'loader', 'tractor', 'generator', 'roller', 'forklift'];
        for (let line of lines) {
            const lineLower = line.toLowerCase();
            if (keywords.some(k => lineLower.includes(k))) {
                vehicleMachinery = line;
                break;
            }
        }
    }

    // 4. Extract Supplier
    const supplierRegex = /(?:supplier|vendor|invoice\s+from|billed\s+by)(?:\s*name)?[\s:-]+([^\n,]+)/i;
    const supplierMatch = text.match(supplierRegex);
    if (supplierMatch) {
        supplierName = supplierMatch[1].trim();
    }

    // 5. Extract Invoice Number
    const invoiceRegex = /(?:invoice|inv)(?:\s*number|\s*no\.?)?[\s:-]+([a-z0-9-]+)/i;
    const invoiceMatch = text.match(invoiceRegex);
    if (invoiceMatch) {
        invoiceNumber = invoiceMatch[1].trim().toUpperCase();
    }

    // 6. Extract GRN
    const grnRegex = /(?:grn|goods\s+received\s+note|receipt)(?:\s*number|\s*no\.?)?[\s:-]+([a-z0-9-]+)/i;
    const grnMatch = text.match(grnRegex);
    if (grnMatch) {
        grnNumber = grnMatch[1].trim().toUpperCase();
    }

    // 7. Extract Quantity and Item Name
    let itemCandidates = [];
    for (let line of lines) {
        const lineLower = line.toLowerCase();
        if (lineLower.includes('monitor') || lineLower.includes('tracker') || lineLower.includes('requisition') || lineLower.includes('report') || lineLower.includes('invoice')) {
            continue;
        }
        
        // Match: [Quantity] [Unit] [Item Name]
        const qtyItemRegex = /\b(\d+(?:\.\d+)?)\s*(?:x|pcs|units|qty|qty:)?\s+([a-zA-Z\s\-]{3,40})\b/i;
        const match = line.match(qtyItemRegex);
        if (match && !lineLower.includes('date') && !lineLower.includes('phone') && !lineLower.includes('total') && !lineLower.includes('no')) {
            const qtyVal = parseFloat(match[1]);
            const nameVal = match[2].trim();
            if (qtyVal > 0 && nameVal.length > 3) {
                itemCandidates.push({ name: nameVal, qty: qtyVal, desc: line });
            }
        }
    }

    if (itemCandidates.length > 0) {
        itemName = itemCandidates[0].name;
        reqQty = itemCandidates[0].qty;
        itemDesc = itemCandidates[0].desc;
    } else {
        const doubleMatch = text.match(/\b(\d+(?:\.\d+)?)\b/);
        if (doubleMatch) {
            reqQty = parseFloat(doubleMatch[1]);
        }
        itemName = "Unparsed Item";
        itemDesc = text.substring(0, 120).replace(/\r?\n/g, ' ') + "...";
    }

    // 8. Extract Unit Price
    const priceRegex = /(?:unit\s*price|rate|price|cost|amount)[\s:-]+(?:rs\.?|usd\.?)?\s*(\d+(?:\.\d+)?)/i;
    const priceMatch = text.match(priceRegex);
    if (priceMatch) {
        unitPrice = parseFloat(priceMatch[1]);
    }

    return {
        mrnNum,
        reqDate,
        vehicleMachinery,
        itemName,
        itemDesc,
        reqQty,
        supplierName,
        invoiceNumber,
        unitPrice,
        grnNumber
    };
}

// 1. Get all items with receipts (with SQL pagination, search, status, and category filtering)
app.get('/api/items', (req, res) => {
    try {
        const page = parseInt(req.query.page) || null;
        const limit = parseInt(req.query.limit) || null;
        const skip = page && limit ? (page - 1) * limit : 0;
        
        const search = req.query.search || null;
        const filter = req.query.filter || null;
        const startDate = req.query.startDate || null;
        const endDate = req.query.endDate || null;
        const vehicle = req.query.vehicle || null;
        const category = req.query.category || null;
        const sort = req.query.sort || 'reqDate';
        const order = req.query.order || 'desc';

        // Base query
        let baseQuery = `
            SELECT 
                i.id,
                i.mrnNum,
                i.reqDate,
                i.vehicleMachinery,
                i.itemName,
                i.itemDesc,
                i.reqQty,
                i.category,
                COALESCE(SUM(r.qty), 0.0) as recQty,
                CASE WHEN COUNT(r.id) > 0 THEN 1 ELSE 0 END as hasReceipts,
                CASE WHEN COUNT(r.id) > 0 AND SUM(CASE WHEN r.unitPrice = 0 OR r.unitPrice IS NULL OR r.invoiceNumber = '' OR r.invoiceNumber IS NULL THEN 1 ELSE 0 END) > 0 THEN 1 ELSE 0 END as isPendingPricing
            FROM items i
            LEFT JOIN receipts r ON i.id = r.itemId
        `;

        const whereClauses = [];
        const queryParams = [];

        if (search) {
            whereClauses.push(`(i.mrnNum LIKE ? OR i.itemName LIKE ? OR i.vehicleMachinery LIKE ? OR i.itemDesc LIKE ? OR i.category LIKE ? OR r.grnNumber LIKE ? OR r.invoiceNumber LIKE ? OR r.supplierName LIKE ?)`);
            const s = `%${search}%`;
            queryParams.push(s, s, s, s, s, s, s, s);
        }

        if (category && category !== 'all') {
            whereClauses.push(`i.category = ?`);
            queryParams.push(category);
        }

        if (vehicle && vehicle !== 'all') {
            whereClauses.push(`LOWER(TRIM(i.vehicleMachinery)) = LOWER(TRIM(?))`);
            queryParams.push(vehicle);
        }

        if (startDate) {
            whereClauses.push(`i.reqDate >= ?`);
            queryParams.push(startDate);
        }

        if (endDate) {
            whereClauses.push(`i.reqDate <= ?`);
            queryParams.push(endDate);
        }

        if (whereClauses.length > 0) {
            baseQuery += ` WHERE ` + whereClauses.join(' AND ');
        }

        baseQuery += ` GROUP BY i.id `;

        const havingClauses = [];
        if (filter === 'pending-delivery') {
            havingClauses.push(`i.reqQty > COALESCE(SUM(r.qty), 0.0)`);
        } else if (filter === 'pending-pricing') {
            havingClauses.push(`i.reqQty <= COALESCE(SUM(r.qty), 0.0)`);
            havingClauses.push(`SUM(CASE WHEN r.unitPrice = 0 OR r.unitPrice IS NULL OR r.invoiceNumber = '' OR r.invoiceNumber IS NULL THEN 1 ELSE 0 END) > 0`);
        } else if (filter === 'completed') {
            havingClauses.push(`i.reqQty <= COALESCE(SUM(r.qty), 0.0)`);
            havingClauses.push(`(COUNT(r.id) = 0 OR SUM(CASE WHEN r.unitPrice = 0 OR r.unitPrice IS NULL OR r.invoiceNumber = '' OR r.invoiceNumber IS NULL THEN 1 ELSE 0 END) = 0)`);
        }

        if (havingClauses.length > 0) {
            baseQuery += ` HAVING ` + havingClauses.join(' AND ');
        }

        // Sorting
        const allowedSortCols = {
            mrnNum: 'i.mrnNum',
            itemName: 'i.itemName',
            vehicleMachinery: 'i.vehicleMachinery',
            reqQty: 'i.reqQty',
            recQty: 'recQty',
            gap: '(i.reqQty - COALESCE(SUM(r.qty), 0.0))',
            reqDate: 'i.reqDate',
            category: 'i.category'
        };

        const sortCol = allowedSortCols[sort] || 'i.reqDate';
        const sortOrder = order === 'asc' ? 'ASC' : 'DESC';
        baseQuery += ` ORDER BY ${sortCol} ${sortOrder} `;

        // Count Query for Pagination
        const countQuery = `SELECT COUNT(*) as count FROM (${baseQuery})`;
        const totalResult = db.queryGet(countQuery, queryParams);
        const total = totalResult ? totalResult.count : 0;

        // Apply pagination
        let itemsQuery = baseQuery;
        const itemsParams = [...queryParams];
        if (page && limit) {
            itemsQuery += ` LIMIT ? OFFSET ? `;
            itemsParams.push(limit, skip);
        }

        const items = db.queryAll(itemsQuery, itemsParams);

        // Fetch receipts for the returned items
        if (items.length > 0) {
            const itemIds = items.map(item => item.id);
            const placeholders = itemIds.map(() => '?').join(',');
            const receipts = db.queryAll(`SELECT * FROM receipts WHERE itemId IN (${placeholders})`, itemIds);
            
            const receiptsByItem = {};
            for (let r of receipts) {
                if (!receiptsByItem[r.itemId]) receiptsByItem[r.itemId] = [];
                receiptsByItem[r.itemId].push(r);
            }

            for (let item of items) {
                item.receipts = receiptsByItem[item.id] || [];
                item.name = item.itemName; // Frontend mapping compat
            }
        }

        if (page && limit) {
            res.json({
                items: items,
                total: total,
                page,
                limit,
                totalPages: Math.ceil(total / limit) || 1
            });
        } else {
            res.json(items);
        }
    } catch (e) {
        console.error('API /api/items error:', e);
        res.status(500).json({ error: e.message });
    }
});

// 2. Add new item
app.post('/api/items', (req, res) => {
    try {
        const { mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty, category } = req.body;
        const cat = category || classifyItem(itemName);
        
        const result = db.run(
            'INSERT INTO items (mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty, category) VALUES (?, ?, ?, ?, ?, ?, ?)',
            [mrnNum || '', reqDate || '', vehicleMachinery || '', itemName || '', itemDesc || '', reqQty || 0, cat]
        );
        res.json({ success: true, id: result.lastInsertRowid });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 3. Delete item
app.delete('/api/items/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id);
        db.run('DELETE FROM items WHERE id = ?', [id]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 4. Add receipt
app.post('/api/items/:id/receipts', (req, res) => {
    try {
        const itemId = parseInt(req.params.id);
        const { qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice } = req.body;
        
        const result = db.run(
            'INSERT INTO receipts (itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
            [itemId, qty || 0, transactionType || 'Receive', deliveryDate || '', purchaseSource || '', grnNumber || '', invoiceNumber || '', invoiceDate || '', supplierName || '', unitPrice ? parseFloat(unitPrice) : null]
        );
        res.json({ success: true, id: result.lastInsertRowid });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5. Delete receipt
app.delete('/api/receipts/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id);
        db.run('DELETE FROM receipts WHERE id = ?', [id]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5.5 Update item
app.put('/api/items/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id);
        const { mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty, category } = req.body;
        const cat = category || classifyItem(itemName);
        
        db.run(
            'UPDATE items SET mrnNum = ?, reqDate = ?, vehicleMachinery = ?, itemName = ?, itemDesc = ?, reqQty = ?, category = ? WHERE id = ?',
            [mrnNum || '', reqDate || '', vehicleMachinery || '', itemName || '', itemDesc || '', reqQty || 0, cat, id]
        );
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5.6 Update receipt details & pricing
app.put('/api/receipts/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id);
        const { qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice } = req.body;
        
        db.run(
            'UPDATE receipts SET qty = ?, transactionType = ?, deliveryDate = ?, purchaseSource = ?, grnNumber = ?, invoiceNumber = ?, invoiceDate = ?, supplierName = ?, unitPrice = ? WHERE id = ?',
            [qty || 0, transactionType || 'Receive', deliveryDate || '', purchaseSource || '', grnNumber || '', invoiceNumber || '', invoiceDate || '', supplierName || '', unitPrice ? parseFloat(unitPrice) : null, id]
        );
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 6. Import Data (bulk importer)
app.post('/api/import', (req, res) => {
    const data = req.body;
    if (!Array.isArray(data)) {
        return res.status(400).json({ error: "Data must be an array of items" });
    }
    
    try {
        db.transaction(() => {
            for (let item of data) {
                const cat = item.category || classifyItem(item.itemName || item.name);
                const result = db.run(
                    'INSERT INTO items (mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty, category) VALUES (?, ?, ?, ?, ?, ?, ?)',
                    [item.mrnNum || '', item.reqDate || '', item.vehicleMachinery || '', item.itemName || item.name || '', item.itemDesc || '', item.reqQty || 0, cat]
                );
                const itemId = result.lastInsertRowid;
                
                if (item.receipts && Array.isArray(item.receipts)) {
                    for (let r of item.receipts) {
                        db.run(
                            'INSERT INTO receipts (itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                            [itemId, r.qty || 0, r.transactionType || r.type || 'Receive', r.deliveryDate || r.date || '', r.purchaseSource || r.source || '', r.grnNumber || '', r.invoiceNumber || '', r.invoiceDate || '', r.supplierName || '', r.unitPrice ? parseFloat(r.unitPrice) : null]
                        );
                    }
                }
            }
        });
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 7. Premium multi-sheet Excel Export Endpoint
app.get('/api/export/excel', (req, res) => {
    try {
        const items = db.queryAll('SELECT * FROM items');
        const receipts = db.queryAll('SELECT * FROM receipts');
        const issues = db.queryAll('SELECT * FROM issues');
        
        const receiptsByItem = {};
        for (let r of receipts) {
            if (!receiptsByItem[r.itemId]) receiptsByItem[r.itemId] = [];
            receiptsByItem[r.itemId].push(r);
        }

        const wb = XLSX.utils.book_new();

        // 1. Build main sheet data
        const itemsSheetData = [];
        itemsSheetData.push([
            "MRN Number", "Request Date", "Vehicle/Machinery", "Item Name", "Item Description", "Category",
            "Requested Qty", "Received Qty", "Receive Date", "Purchase Source", "Qty Gap", 
            "Date Gap (Days)", "Status", "GRN Number", "Invoice Number", "Invoice Date", 
            "Supplier Name", "Unit Price (Rs.)", "Total Price (Rs.)"
        ]);

        const supplierSpend = {};
        let totalSpend = 0;
        let activeSuppliers = new Set();
        let pricedCount = 0;
        let unpricedCount = 0;

        for (let item of items) {
            const itemRecs = receiptsByItem[item.id] || [];
            const recQty = itemRecs.reduce((sum, r) => sum + r.qty, 0);
            const recQtyRounded = Math.round(recQty * 100) / 100;
            
            let recDate = "";
            if (itemRecs.length > 0) {
                const sorted = [...itemRecs].sort((a, b) => new Date(b.deliveryDate) - new Date(a.deliveryDate));
                recDate = sorted[0].deliveryDate;
            }
            
            const uniqueSources = [...new Set(itemRecs.map(r => r.purchaseSource).filter(Boolean))].join(' & ');
            const qtyGap = Math.round((item.reqQty - recQtyRounded) * 100) / 100;
            
            let dateGapDays = "";
            if (recDate && item.reqDate) {
                const d1 = new Date(item.reqDate);
                const d2 = new Date(recDate);
                dateGapDays = Math.ceil((d2 - d1) / (1000 * 60 * 60 * 24));
            }

            let status = "Pending";
            if (recQtyRounded > 0) {
                if (recQtyRounded < item.reqQty) status = "Partial";
                else if (recQtyRounded === item.reqQty) status = "Complete";
                else status = "Over-received";
            }

            const grns = [...new Set(itemRecs.map(r => r.grnNumber).filter(Boolean))].join('; ');
            const invoices = [...new Set(itemRecs.map(r => r.invoiceNumber).filter(Boolean))].join('; ');
            const invoiceDates = [...new Set(itemRecs.map(r => r.invoiceDate).filter(Boolean))].filter(d => d && d !== '1899-12-30').join('; ');
            const suppliers = [...new Set(itemRecs.map(r => r.supplierName).filter(Boolean))].join('; ');
            
            let totalUnitPrice = "";
            let totalPrice = 0;
            let hasPricing = false;

            const pricedReceipts = itemRecs.filter(r => r.unitPrice);
            if (pricedReceipts.length > 0) {
                totalUnitPrice = pricedReceipts.map(r => r.unitPrice).join('; ');
                totalPrice = itemRecs.reduce((sum, r) => {
                    if (r.unitPrice && r.qty > 0) {
                        const cost = Math.abs(r.qty) * r.unitPrice;
                        totalSpend += cost;
                        const sup = r.supplierName || 'Unknown Supplier';
                        supplierSpend[sup] = (supplierSpend[sup] || 0) + cost;
                        activeSuppliers.add(sup);
                        return sum + cost;
                    }
                    return sum;
                }, 0);
                totalPrice = Math.round(totalPrice * 100) / 100;
                hasPricing = true;
            }

            if (recQtyRounded > 0) {
                if (hasPricing) pricedCount++;
                else unpricedCount++;
            }

            itemsSheetData.push([
                item.mrnNum || "",
                item.reqDate || "",
                item.vehicleMachinery || "",
                item.itemName || "",
                item.itemDesc || "",
                item.category || "General Items",
                item.reqQty || 0,
                recQtyRounded,
                recDate,
                uniqueSources,
                qtyGap,
                dateGapDays,
                status,
                grns,
                invoices,
                invoiceDates,
                suppliers,
                totalUnitPrice,
                totalPrice || ""
            ]);
        }

        const wsItems = XLSX.utils.aoa_to_sheet(itemsSheetData);
        XLSX.utils.book_append_sheet(wb, wsItems, "Requests & Deliveries");

        // 2. Build summary sheet data
        const summarySheetData = [];
        summarySheetData.push(["Supplier Name", "Total Spend (Rs.)", "Spend Share (%)"]);
        
        const sortedSuppliers = Object.entries(supplierSpend).sort((a, b) => b[1] - a[1]);
        for (let [name, amount] of sortedSuppliers) {
            const pct = totalSpend > 0 ? ((amount / totalSpend) * 100).toFixed(1) : 0;
            summarySheetData.push([name, amount, `${pct}%`]);
        }
        
        if (sortedSuppliers.length > 0) {
            summarySheetData.push([""]);
            summarySheetData.push(["TOTAL SPEND", totalSpend, "100.0%"]);
            summarySheetData.push(["ACTIVE SUPPLIERS", activeSuppliers.size, ""]);
            summarySheetData.push(["PRICED DELIVERIES", pricedCount, ""]);
            summarySheetData.push(["UNPRICED DELIVERIES", unpricedCount, ""]);
        }

        const wsSummary = XLSX.utils.aoa_to_sheet(summarySheetData);
        XLSX.utils.book_append_sheet(wb, wsSummary, "Financial Summary");

        // 3. Build Outbound Issues sheet data
        const issuesSheetData = [];
        issuesSheetData.push(["Issue Date", "MRN Ref", "Item Name", "Qty Issued", "Issued to Vehicle", "Issued By", "Notes"]);
        for (let issue of issues) {
            issuesSheetData.push([
                issue.date || "",
                issue.mrnNum || "",
                issue.itemName || "",
                issue.qty || 0,
                issue.vehicleMachinery || "",
                issue.issuedBy || "",
                issue.notes || ""
            ]);
        }
        const wsIssues = XLSX.utils.aoa_to_sheet(issuesSheetData);
        XLSX.utils.book_append_sheet(wb, wsIssues, "Outbound Issues");

        const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
        
        res.setHeader('Content-Disposition', 'attachment; filename="delivery_monitor_report.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buf);

    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 8. Heuristic PDF Text Parser Endpoint
app.post('/api/import/pdf', async (req, res) => {
    const { pdfBase64 } = req.body;
    if (!pdfBase64) {
        return res.status(400).json({ error: "Missing pdfBase64 content" });
    }

    try {
        const buffer = Buffer.from(pdfBase64, 'base64');
        const data = await pdfParse(buffer);
        const text = data.text;
        
        // Match heuristic values
        const parsedData = parsePdfTextHeuristically(text);
        res.json({ success: true, text, data: parsedData });
    } catch (e) {
        console.error("PDF Parsing Error:", e);
        res.status(500).json({ error: "Failed to parse PDF file: " + e.message });
    }
});

// 9. Issues CRUD Endpoints
app.get('/api/issues', (req, res) => {
    try {
        let sql = 'SELECT * FROM issues';
        const params = [];
        const where = [];

        if (req.query.vehicle && req.query.vehicle !== 'all') {
            where.push('LOWER(TRIM(vehicleMachinery)) = LOWER(TRIM(?))');
            params.push(req.query.vehicle);
        }
        if (req.query.startDate) {
            where.push('date >= ?');
            params.push(req.query.startDate);
        }
        if (req.query.endDate) {
            where.push('date <= ?');
            params.push(req.query.endDate);
        }
        
        if (where.length > 0) {
            sql += ' WHERE ' + where.join(' AND ');
        }
        sql += ' ORDER BY date DESC, id DESC';

        const issues = db.queryAll(sql, params);
        res.json(issues);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/issues', (req, res) => {
    try {
        const { date, itemName, qty, vehicleMachinery, issuedBy, mrnNum, notes, itemId } = req.body;
        const result = db.run(
            'INSERT INTO issues (date, itemName, qty, vehicleMachinery, issuedBy, mrnNum, notes, itemId) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            [date || '', itemName || '', qty || 0, vehicleMachinery || '', issuedBy || '', mrnNum || null, notes || '', itemId || null]
        );
        res.json({ success: true, id: result.lastInsertRowid });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.put('/api/issues/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id);
        const { date, itemName, qty, vehicleMachinery, issuedBy, mrnNum, notes, itemId } = req.body;
        db.run(
            'UPDATE issues SET date = ?, itemName = ?, qty = ?, vehicleMachinery = ?, issuedBy = ?, mrnNum = ?, notes = ?, itemId = ? WHERE id = ?',
            [date || '', itemName || '', qty || 0, vehicleMachinery || '', issuedBy || '', mrnNum || null, notes || '', itemId || null, id]
        );
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.delete('/api/issues/:id', (req, res) => {
    try {
        const id = parseInt(req.params.id);
        db.run('DELETE FROM issues WHERE id = ?', [id]);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 10. Vehicle History Timeline Endpoint
app.get('/api/vehicle-history', (req, res) => {
    try {
        const vehicle = req.query.vehicle;
        if (!vehicle) {
            return res.status(400).json({ error: "Missing vehicle parameter" });
        }
        const startDate = req.query.startDate || null;
        const endDate = req.query.endDate || null;

        const dateFilter = (col) => {
            let filter = '';
            const params = [];
            if (startDate) {
                filter += ` AND ${col} >= ? `;
                params.push(startDate);
            }
            if (endDate) {
                filter += ` AND ${col} <= ? `;
                params.push(endDate);
            }
            return { filter, params };
        };

        // 1. Fetch Requisitions (Requests)
        const reqFilter = dateFilter('reqDate');
        const reqs = db.queryAll(
            `SELECT id, mrnNum, reqDate as date, itemName, reqQty as qty, 'Request' as type, itemDesc as details FROM items WHERE LOWER(TRIM(vehicleMachinery)) = LOWER(TRIM(?)) ${reqFilter.filter}`,
            [vehicle, ...reqFilter.params]
        );

        // 2. Fetch Receipts (Receives)
        const recFilter = dateFilter('r.deliveryDate');
        const recs = db.queryAll(
            `SELECT r.id, i.mrnNum, r.deliveryDate as date, i.itemName, r.qty, 'Receive' as type, r.supplierName || ' (GRN: ' || r.grnNumber || ')' as details FROM receipts r JOIN items i ON r.itemId = i.id WHERE LOWER(TRIM(i.vehicleMachinery)) = LOWER(TRIM(?)) ${recFilter.filter}`,
            [vehicle, ...recFilter.params]
        );

        // 3. Fetch Issues (Outbounds)
        const issFilter = dateFilter('date');
        const iss = db.queryAll(
            `SELECT id, mrnNum, date, itemName, qty, 'Issue' as type, issuedBy || ' (Notes: ' || notes || ')' as details FROM issues WHERE LOWER(TRIM(vehicleMachinery)) = LOWER(TRIM(?)) ${issFilter.filter}`,
            [vehicle, ...issFilter.params]
        );

        // Combine and Sort by Date Descending
        const timeline = [...reqs, ...recs, ...iss].sort((a, b) => {
            return new Date(b.date) - new Date(a.date);
        });

        res.json(timeline);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// Dashboard aggregates endpoint
app.get('/api/dashboard/stats', (req, res) => {
    try {
        // Total Spend
        const spendResult = db.queryGet('SELECT SUM(ABS(qty) * unitPrice) as total FROM receipts WHERE unitPrice IS NOT NULL AND qty > 0');
        const totalSpend = spendResult ? spendResult.total || 0 : 0;

        // Active Suppliers
        const supplierResult = db.queryGet('SELECT COUNT(DISTINCT supplierName) as count FROM receipts WHERE supplierName IS NOT NULL AND supplierName != \'\'');
        const supplierCount = supplierResult ? supplierResult.count : 0;

        // Priced Deliveries
        const pricedResult = db.queryGet('SELECT COUNT(*) as count FROM receipts WHERE unitPrice IS NOT NULL AND unitPrice > 0 AND invoiceNumber IS NOT NULL AND invoiceNumber != \'\'');
        const pricedCount = pricedResult ? pricedResult.count : 0;

        // Unpriced / Pending
        const unpricedResult = db.queryGet(`
            SELECT COUNT(*) as count FROM items i
            JOIN receipts r ON i.id = r.itemId
            WHERE r.unitPrice IS NULL OR r.unitPrice = 0 OR r.invoiceNumber IS NULL OR r.invoiceNumber = ''
        `);
        const unpricedCount = unpricedResult ? unpricedResult.count : 0;

        // Category breakdown
        const categoryBreakdown = db.queryAll(`
            SELECT category, COUNT(*) as count, SUM(reqQty) as totalReqQty 
            FROM items 
            GROUP BY category
        `);

        // Spend Trend by Month
        const spendTrend = db.queryAll(`
            SELECT strftime('%Y-%m', deliveryDate) as month, SUM(qty * unitPrice) as spend
            FROM receipts
            WHERE deliveryDate IS NOT NULL AND deliveryDate != '' AND unitPrice IS NOT NULL
            GROUP BY month
            ORDER BY month ASC
        `);

        // Supplier Distribution
        const supplierShare = db.queryAll(`
            SELECT supplierName, SUM(qty * unitPrice) as spend
            FROM receipts
            WHERE supplierName IS NOT NULL AND supplierName != '' AND unitPrice IS NOT NULL
            GROUP BY supplierName
            ORDER BY spend DESC
            LIMIT 5
        `);

        res.json({
            totalSpend,
            supplierCount,
            pricedCount,
            unpricedCount,
            categoryBreakdown,
            spendTrend,
            supplierShare
        });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// Fetch active vehicles list
app.get('/api/vehicles', (req, res) => {
    try {
        const vehicles = db.queryAll('SELECT DISTINCT vehicleMachinery FROM items WHERE vehicleMachinery IS NOT NULL AND vehicleMachinery != \'\' ORDER BY vehicleMachinery ASC');
        res.json(vehicles.map(v => v.vehicleMachinery));
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Delivery Monitor Server running at http://localhost:${PORT}`);
    console.log(`Database: ${DB_PATH}`);
});

// Automatic Backup System (Every 30 Minutes)
const BACKUP_DIR = path.join(__dirname, 'backups');
const BACKUP_INTERVAL = 30 * 60 * 1000;

function runAutomaticBackup() {
    try {
        if (!fs.existsSync(BACKUP_DIR)) {
            fs.mkdirSync(BACKUP_DIR, { recursive: true });
        }
        
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        
        const timestamp = `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;
        const backupPath = path.join(BACKUP_DIR, `inventory_backup_${timestamp}.db`);
        
        if (fs.existsSync(DB_PATH)) {
            fs.copyFileSync(DB_PATH, backupPath);
            console.log(`[BACKUP] Automatic backup created successfully: ${backupPath}`);
            cleanOldBackups();
        } else {
            console.error(`[BACKUP] DB_PATH file not found, backup failed!`);
        }
    } catch (e) {
        console.error(`[BACKUP] Failed to create automatic backup:`, e.message);
    }
}

function cleanOldBackups() {
    try {
        const files = fs.readdirSync(BACKUP_DIR)
            .filter(file => file.startsWith('inventory_backup_') && file.endsWith('.db'))
            .map(file => ({
                name: file,
                path: path.join(BACKUP_DIR, file),
                mtime: fs.statSync(path.join(BACKUP_DIR, file)).mtime
            }))
            .sort((a, b) => b.mtime - a.mtime);
            
        if (files.length > 48) {
            const extraFiles = files.slice(48);
            extraFiles.forEach(file => {
                try {
                    fs.unlinkSync(file.path);
                    console.log(`[BACKUP] Deleted old backup file: ${file.name}`);
                } catch (err) {
                    console.error(`[BACKUP] Failed to delete old backup file ${file.name}:`, err.message);
                }
            });
        }
    } catch (err) {
        console.error(`[BACKUP] Error cleaning up old backups:`, err.message);
    }
}

setInterval(runAutomaticBackup, BACKUP_INTERVAL);
setTimeout(runAutomaticBackup, 5000);
