const { DatabaseSync } = require('node:sqlite');
const path = require('path');
const fs = require('fs');

const DB_FILE = path.join(__dirname, 'inventory.db');

const db = new DatabaseSync(DB_FILE);

// Enable foreign key support
db.exec('PRAGMA foreign_keys = ON;');

// Create tables
db.exec(`
    CREATE TABLE IF NOT EXISTS items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mrnNum TEXT,
        reqDate TEXT,
        vehicleMachinery TEXT,
        itemName TEXT,
        itemDesc TEXT,
        reqQty REAL,
        category TEXT
    );

    CREATE TABLE IF NOT EXISTS receipts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        itemId INTEGER,
        qty REAL,
        transactionType TEXT,
        deliveryDate TEXT,
        purchaseSource TEXT,
        grnNumber TEXT,
        invoiceNumber TEXT,
        invoiceDate TEXT,
        supplierName TEXT,
        unitPrice REAL,
        FOREIGN KEY (itemId) REFERENCES items (id) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS issues (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        itemId INTEGER,
        date TEXT,
        itemName TEXT,
        qty REAL,
        vehicleMachinery TEXT,
        issuedBy TEXT,
        mrnNum TEXT,
        notes TEXT,
        FOREIGN KEY (itemId) REFERENCES items (id) ON DELETE SET NULL
    );
`);

// Create indexes for fast filtering, searching, and joins
db.exec(`
    CREATE INDEX IF NOT EXISTS idx_items_mrnNum ON items (mrnNum);
    CREATE INDEX IF NOT EXISTS idx_items_reqDate ON items (reqDate);
    CREATE INDEX IF NOT EXISTS idx_items_vehicle ON items (vehicleMachinery);
    CREATE INDEX IF NOT EXISTS idx_items_category ON items (category);

    CREATE INDEX IF NOT EXISTS idx_receipts_itemId ON receipts (itemId);
    CREATE INDEX IF NOT EXISTS idx_receipts_deliveryDate ON receipts (deliveryDate);
    CREATE INDEX IF NOT EXISTS idx_receipts_grnNumber ON receipts (grnNumber);

    CREATE INDEX IF NOT EXISTS idx_issues_date ON issues (date);
    CREATE INDEX IF NOT EXISTS idx_issues_vehicle ON issues (vehicleMachinery);
    CREATE INDEX IF NOT EXISTS idx_issues_itemId ON issues (itemId);
`);

// Query wrappers for cleaner code
function queryAll(sql, params = []) {
    const stmt = db.prepare(sql);
    return stmt.all(...params);
}

function queryGet(sql, params = []) {
    const stmt = db.prepare(sql);
    return stmt.get(...params);
}

function run(sql, params = []) {
    const stmt = db.prepare(sql);
    return stmt.run(...params);
}

function transaction(fn) {
    db.exec('BEGIN TRANSACTION;');
    try {
        const result = fn();
        db.exec('COMMIT;');
        return result;
    } catch (e) {
        db.exec('ROLLBACK;');
        throw e;
    }
}

module.exports = {
    db,
    queryAll,
    queryGet,
    run,
    transaction
};
