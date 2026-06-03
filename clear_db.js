const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const DB_PATH = path.join(__dirname, 'inventory.accdb');

const psScript = `
$ErrorActionPreference = 'Stop'
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = "DELETE FROM receipts"
$cmd.ExecuteNonQuery() | Out-Null

$cmd.CommandText = "DELETE FROM items"
$cmd.ExecuteNonQuery() | Out-Null

$conn.Close()
Write-Host 'Database cleared successfully.'
`;

const tempPath = path.join(__dirname, '_clear_db.ps1');
fs.writeFileSync(tempPath, psScript, 'utf8');

try {
    const output = execSync(`powershell -ExecutionPolicy Bypass -File "${tempPath}"`, { encoding: 'utf8' });
    console.log(output);
} catch (e) {
    console.error(e.message);
} finally {
    fs.unlinkSync(tempPath);
}
