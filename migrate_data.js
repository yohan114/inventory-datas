/**
 * migrate_data.js - Migrates existing data from tracker_data.json to MS Access database
 */
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const DB_PATH = path.join(__dirname, 'inventory.accdb');
const JSON_PATH = path.join(__dirname, 'tracker_data.json');

if (!fs.existsSync(DB_PATH)) {
    console.error("[FAIL] Database doesn't exist. Run 'node setup_db.js' first.");
    process.exit(1);
}

if (!fs.existsSync(JSON_PATH)) {
    console.log("[WARN] No tracker_data.json found. Nothing to migrate.");
    process.exit(0);
}

console.log('Reading existing data from tracker_data.json...');
const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
console.log(`Found ${data.length} items to migrate.`);

let psScript = `
$ErrorActionPreference = 'Stop'
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

Write-Host "Starting migration..."
`;

const sanitize = (str) => {
    if (str === null || str === undefined) return '';
    return String(str).replace(/'/g, "''").replace(/\n/g, " ");
}

for (let item of data) {
    const { mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty, receipts } = item;
    
    psScript += `
$cmd = $conn.CreateCommand()
$cmd.CommandText = "INSERT INTO items (mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty) VALUES ('${sanitize(mrnNum)}', '${sanitize(reqDate)}', '${sanitize(vehicleMachinery)}', '${sanitize(itemName)}', '${sanitize(itemDesc)}', ${reqQty || 0})"
$cmd.ExecuteNonQuery() | Out-Null
$cmd.CommandText = "SELECT @@IDENTITY"
$itemId = $cmd.ExecuteScalar()
`;

    if (receipts && receipts.length > 0) {
        for (let receipt of receipts) {
            const { qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice } = receipt;
            psScript += `
$cmd = $conn.CreateCommand()
$cmd.CommandText = "INSERT INTO receipts (itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice) VALUES ($itemId, ${qty || 0}, '${sanitize(transactionType)}', '${sanitize(deliveryDate)}', '${sanitize(purchaseSource)}', '${sanitize(grnNumber)}', '${sanitize(invoiceNumber)}', '${sanitize(invoiceDate)}', '${sanitize(supplierName)}', ${unitPrice || 0})"
$cmd.ExecuteNonQuery() | Out-Null
`;
        }
    }
}

psScript += `
$conn.Close()
Write-Host "Migration complete!"
`;

const tempPath = path.join(__dirname, '_migrate_temp.ps1');
fs.writeFileSync(tempPath, psScript, 'utf8');

try {
    console.log('Executing migration into MS Access (this might take a moment)...');
    execSync(`powershell -ExecutionPolicy Bypass -File "${tempPath}"`, { stdio: 'inherit' });
    console.log('[OK] Migration successful!');
} catch (err) {
    console.error('[FAIL] Migration failed:', err.message);
} finally {
    try { fs.unlinkSync(tempPath); } catch (e) {}
}
