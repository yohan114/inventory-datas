const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const { run, transaction, queryGet } = require('./db');

const DB_PATH = path.join(__dirname, 'inventory.accdb');
const JSON_PATH = path.join(__dirname, 'tracker_data.json');

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

const parseDateToISO = (dateStr) => {
    if (!dateStr) return '';
    const str = String(dateStr).trim();
    if (!str || str === '1899-12-30') return '';
    
    // Handle YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
        return str.substring(0, 10);
    }
    
    // Handle MM/DD/YYYY
    if (str.includes('/')) {
        const parts = str.split('/');
        if (parts.length === 3) {
            if (parts[0].length === 4) {
                const year = parts[0];
                const month = String(parts[1]).padStart(2, '0');
                const day = String(parts[2]).padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
            if (parts[2].length === 4) {
                const year = parts[2];
                const month = String(parts[0]).padStart(2, '0');
                const day = String(parts[1]).padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
            if (parts[2].length === 2) {
                const year = '20' + parts[2];
                const month = String(parts[0]).padStart(2, '0');
                const day = String(parts[1]).padStart(2, '0');
                return `${year}-${month}-${day}`;
            }
        }
    }
    
    const parsed = new Date(str);
    if (isNaN(parsed.getTime())) return str;
    return parsed.toISOString().split('T')[0];
};

const loadDataFromAccess = () => {
    console.log('Spawning PowerShell to read from Access database (inventory.accdb)...');
    const psScript = `
        $ErrorActionPreference = 'Stop'
        $dbPath = '${DB_PATH.replace(/'/g, "''")}'
        $conn = New-Object System.Data.OleDb.OleDbConnection
        $conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
        $conn.Open()
        
        $cmd = $conn.CreateCommand()
        $cmd.CommandText = "SELECT * FROM items"
        $reader = $cmd.ExecuteReader()
        $items = @()
        while ($reader.Read()) {
            $items += @{
                id = $reader["id"]
                mrnNum = $reader["mrnNum"]
                reqDate = $reader["reqDate"]
                vehicleMachinery = $reader["vehicleMachinery"]
                itemName = $reader["itemName"]
                itemDesc = $reader["itemDesc"]
                reqQty = $reader["reqQty"]
            }
        }
        $reader.Close()
        
        $cmd.CommandText = "SELECT * FROM receipts"
        $reader = $cmd.ExecuteReader()
        $receipts = @()
        while ($reader.Read()) {
            $receipts += @{
                id = $reader["id"]
                itemId = $reader["itemId"]
                qty = $reader["qty"]
                transactionType = $reader["transactionType"]
                deliveryDate = $reader["deliveryDate"]
                purchaseSource = $reader["purchaseSource"]
                grnNumber = $reader["grnNumber"]
                invoiceNumber = $reader["invoiceNumber"]
                invoiceDate = $reader["invoiceDate"]
                supplierName = $reader["supplierName"]
                unitPrice = $reader["unitPrice"]
            }
        }
        $reader.Close()
        $conn.Close()
        
        @{ items = $items; receipts = $receipts } | ConvertTo-Json -Depth 5 -Compress
    `;

    const tempPath = path.join(__dirname, '_migrate_dump.ps1');
    fs.writeFileSync(tempPath, psScript, 'utf8');

    try {
        const output = execSync(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "${tempPath}"`, {
            encoding: 'utf8',
            maxBuffer: 1024 * 1024 * 100
        });
        return JSON.parse(output.trim());
    } catch (e) {
        console.warn('[WARN] Failed to read from MS Access database. Falling back to JSON backup.', e.message);
        return null;
    } finally {
        try { fs.unlinkSync(tempPath); } catch (err) {}
    }
};

const loadDataFromJson = () => {
    if (!fs.existsSync(JSON_PATH)) {
        console.error('[FAIL] No JSON backup or Access database found.');
        process.exit(1);
    }
    console.log('Reading from tracker_data.json backup...');
    const raw = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
    
    const items = [];
    const receipts = [];
    let itemIdCounter = 1;
    let receiptIdCounter = 1;

    for (const rawItem of raw) {
        const itemId = itemIdCounter++;
        items.push({
            id: itemId,
            mrnNum: rawItem.mrnNum,
            reqDate: rawItem.reqDate,
            vehicleMachinery: rawItem.vehicleMachinery,
            itemName: rawItem.name || rawItem.itemName || '',
            itemDesc: rawItem.itemDesc || '',
            reqQty: rawItem.reqQty || 0
        });

        if (rawItem.receipts && Array.isArray(rawItem.receipts)) {
            for (const rawRec of rawItem.receipts) {
                receipts.push({
                    id: receiptIdCounter++,
                    itemId: itemId,
                    qty: rawRec.qty || 0,
                    transactionType: rawRec.type || 'Receive',
                    deliveryDate: rawRec.date || '',
                    purchaseSource: rawRec.source || '',
                    grnNumber: rawRec.grnNumber || '',
                    invoiceNumber: rawRec.invoiceNumber || '',
                    invoiceDate: rawRec.invoiceDate || '',
                    supplierName: rawRec.supplierName || '',
                    unitPrice: rawRec.unitPrice || 0
                });
            }
        }
    }
    return { items, receipts };
};

const runMigration = () => {
    let data = loadDataFromAccess();
    if (!data) {
        data = loadDataFromJson();
    }

    const { items, receipts } = data;
    console.log(`Loaded ${items.length} items and ${receipts.length} receipts.`);

    transaction(() => {
        // Clear tables
        run('DELETE FROM items');
        run('DELETE FROM receipts');

        // Insert items
        const itemStmt = 'INSERT INTO items (id, mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty, category) VALUES (?, ?, ?, ?, ?, ?, ?, ?)';
        let itemsCount = 0;
        const itemIds = new Set();
        for (const item of items) {
            const cat = classifyItem(item.itemName);
            const dateISO = parseDateToISO(item.reqDate);
            run(itemStmt, [
                item.id,
                item.mrnNum || '',
                dateISO,
                item.vehicleMachinery || '',
                item.itemName || '',
                item.itemDesc || '',
                item.reqQty || 0,
                cat
            ]);
            itemIds.add(item.id);
            itemsCount++;
        }

        // Insert receipts
        const recStmt = 'INSERT INTO receipts (id, itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)';
        let recsCount = 0;
        let skippedOrphans = 0;
        for (const rec of receipts) {
            if (!itemIds.has(rec.itemId)) {
                skippedOrphans++;
                continue;
            }
            const delDateISO = parseDateToISO(rec.deliveryDate);
            const invDateISO = parseDateToISO(rec.invoiceDate);
            run(recStmt, [
                rec.id,
                rec.itemId,
                rec.qty || 0,
                rec.transactionType || 'Receive',
                delDateISO,
                rec.purchaseSource || '',
                rec.grnNumber || '',
                rec.invoiceNumber || '',
                invDateISO,
                rec.supplierName || '',
                rec.unitPrice || 0
            ]);
            recsCount++;
        }

        console.log(`[OK] Successfully migrated ${itemsCount} items and ${recsCount} receipts into SQLite! (Skipped ${skippedOrphans} orphan receipts)`);
    });
};

runMigration();
