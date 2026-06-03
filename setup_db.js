const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const DB_PATH = path.join(__dirname, 'inventory.accdb');

if (fs.existsSync(DB_PATH)) {
    console.log('Database already exists:', DB_PATH);
    process.exit(0);
}

console.log('Creating MS Access database...');

const psScript = `
$ErrorActionPreference = 'Stop'
$dbPath = '${DB_PATH.replace(/'/g, "''")}'

Write-Host "  Creating database file..."
try {
    $cat = New-Object -ComObject ADOX.Catalog
    $connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;Jet OLEDB:Engine Type=6"
    $cat.Create($connStr)
    $cat.ActiveConnection.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cat) | Out-Null
    Write-Host "  [OK] Database file created"
} catch {
    Write-Host "  [FAIL] Failed to create database: $($_.Exception.Message)"
    exit 1
}

Write-Host "  Creating tables..."
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd1 = $conn.CreateCommand()
$cmd1.CommandText = "CREATE TABLE items (id AUTOINCREMENT PRIMARY KEY, mrnNum TEXT(50), reqDate TEXT(30), vehicleMachinery TEXT(100), itemName TEXT(255), itemDesc TEXT(255), reqQty DOUBLE)"
$cmd1.ExecuteNonQuery() | Out-Null
Write-Host "  [OK] Table 'items' created"

$cmd2 = $conn.CreateCommand()
$cmd2.CommandText = "CREATE TABLE receipts (id AUTOINCREMENT PRIMARY KEY, itemId LONG, qty DOUBLE, transactionType TEXT(20), deliveryDate TEXT(30), purchaseSource TEXT(50), grnNumber TEXT(50), invoiceNumber TEXT(50), invoiceDate TEXT(30), supplierName TEXT(100), unitPrice DOUBLE)"
$cmd2.ExecuteNonQuery() | Out-Null
Write-Host "  [OK] Table 'receipts' created"

$conn.Close()
Write-Host "  Database location: $dbPath"
`;

const tempPath = path.join(__dirname, '_setup_temp.ps1');
fs.writeFileSync(tempPath, psScript, 'utf8');

try {
    execSync(`powershell -ExecutionPolicy Bypass -File "${tempPath}"`, { stdio: 'inherit' });
    console.log('Database setup complete!');
} catch (err) {
    console.error('Setup failed.');
    process.exit(1);
} finally {
    try { fs.unlinkSync(tempPath); } catch(e) {}
}
