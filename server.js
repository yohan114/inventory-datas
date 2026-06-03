const express = require('express');
const { execSync } = require('child_process');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const os = require('os');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');

const app = express();
const PORT = 4000;

app.use(express.json({ limit: '100mb' }));
app.use(express.static(__dirname));

const DB_PATH = path.join(__dirname, 'inventory.accdb');

// Helper to run a PowerShell DB script with robust Base64 parameters and fast non-profile startup
function runDbQuery(psScript, inputData = null) {
    const uniqueId = crypto.randomBytes(6).toString('hex');
    const tempPath = path.join(__dirname, `_query_${uniqueId}.ps1`);
    
    let fullScript = `$ErrorActionPreference = 'Stop'\n`;
    
    if (inputData !== null) {
        const base64Str = Buffer.from(JSON.stringify(inputData)).toString('base64');
        fullScript += `$base64Input = "${base64Str}"\n`;
        fullScript += `$decodedBytes = [System.Convert]::FromBase64String($base64Input)\n`;
        fullScript += `$jsonText = [System.Text.Encoding]::UTF8.GetString($decodedBytes)\n`;
        fullScript += `$params = $jsonText | ConvertFrom-Json\n`;
    }
    
    fullScript += psScript;
    fs.writeFileSync(tempPath, fullScript, 'utf8');
    
    let retries = 5;
    let delay = 100;
    
    try {
        while (retries > 0) {
            try {
                const output = execSync(`powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "${tempPath}"`, { 
                    encoding: 'utf8', 
                    stdio: ['pipe', 'pipe', 'pipe'], 
                    maxBuffer: 1024 * 1024 * 100 
                });
                return JSON.parse(output.trim() || '[]');
            } catch (e) {
                const errStr = (e.message || '') + ' ' + (e.stderr || '');
                const isLockError = errStr.toLowerCase().includes("lock") || errStr.toLowerCase().includes("sharing") || errStr.toLowerCase().includes("in use");
                if (isLockError && retries > 1) {
                    retries--;
                    console.warn(`Database locked or busy, retrying in ${delay}ms... (${retries} retries left)`);
                    execSync(`powershell -Command "Start-Sleep -Milliseconds ${delay}"`);
                    delay *= 2;
                    continue;
                }
                console.error("DB Query Error:", e.message);
                if (e.stderr) {
                    console.error("PowerShell Stderr:", e.stderr);
                }
                throw e;
            }
        }
    } finally {
        try { fs.unlinkSync(tempPath); } catch (err) {}
    }
}

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
    const dateRegex = /\b(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b/;
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

// 1. Get all items with receipts (with optional pagination, search, and status tab filtering support)
app.get('/api/items', (req, res) => {
    const page = parseInt(req.query.page) || null;
    const limit = parseInt(req.query.limit) || null;
    const skip = page && limit ? (page - 1) * limit : 0;
    const search = req.query.search || null;
    const filter = req.query.filter || null;
    const startDate = req.query.startDate || null;
    const endDate = req.query.endDate || null;
    const vehicle = req.query.vehicle || null;
    const sort = req.query.sort || 'reqDate';
    const order = req.query.order || 'desc';

    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

# Get Items
$cmdItems = $conn.CreateCommand()
$cmdItems.CommandText = "SELECT * FROM items"
$readerItems = $cmdItems.ExecuteReader()
$items = @()
while ($readerItems.Read()) {
    $item = @{
        id = $readerItems["id"]
        mrnNum = $readerItems["mrnNum"]
        reqDate = $readerItems["reqDate"]
        vehicleMachinery = $readerItems["vehicleMachinery"]
        itemName = $readerItems["itemName"]
        itemDesc = $readerItems["itemDesc"]
        reqQty = $readerItems["reqQty"]
    }
    $items += $item
}
$readerItems.Close()

# Get Receipts
$cmdReceipts = $conn.CreateCommand()
$cmdReceipts.CommandText = "SELECT * FROM receipts"
$readerReceipts = $cmdReceipts.ExecuteReader()
$receipts = @()
while ($readerReceipts.Read()) {
    $receipt = @{
        id = $readerReceipts["id"]
        itemId = $readerReceipts["itemId"]
        qty = $readerReceipts["qty"]
        transactionType = $readerReceipts["transactionType"]
        deliveryDate = $readerReceipts["deliveryDate"]
        purchaseSource = $readerReceipts["purchaseSource"]
        grnNumber = $readerReceipts["grnNumber"]
        invoiceNumber = $readerReceipts["invoiceNumber"]
        invoiceDate = $readerReceipts["invoiceDate"]
        supplierName = $readerReceipts["supplierName"]
        unitPrice = $readerReceipts["unitPrice"]
    }
    $receipts += $receipt
}
$readerReceipts.Close()
$conn.Close()

# Map receipts to items temporarily in PowerShell to calculate recQty
foreach ($item in $items) {
    $itemRecs = @()
    if ($null -ne $receipts) {
        foreach ($r in $receipts) {
            if ($r.itemId -eq $item.id) {
                $itemRecs += $r
            }
        }
    }
    
    # Calculate recQty
    $recQty = 0.0
    foreach ($r in $itemRecs) {
        $recQty += $r.qty
    }
    $item.recQty = $recQty
    $item.hasReceipts = $itemRecs.Count -gt 0
    
    # Check if pending pricing
    $isPendingPricing = $false
    if ($itemRecs.Count -gt 0) {
        foreach ($r in $itemRecs) {
            if ($null -eq $r.unitPrice -or $r.unitPrice -eq 0 -or $null -eq $r.invoiceNumber -or $r.invoiceNumber -eq "") {
                $isPendingPricing = $true
            }
        }
    }
    $item.isPendingPricing = $isPendingPricing
}

# Filter by search if provided
if ($null -ne $params.search -and $params.search -ne "") {
    $searchLower = $params.search.ToLower()
    $filtered = @()
    foreach ($item in $items) {
        $match = $false
        if ($null -ne $item.mrnNum -and $item.mrnNum.ToLower().Contains($searchLower)) { $match = $true }
        elseif ($null -ne $item.itemName -and $item.itemName.ToLower().Contains($searchLower)) { $match = $true }
        elseif ($null -ne $item.vehicleMachinery -and $item.vehicleMachinery.ToLower().Contains($searchLower)) { $match = $true }
        elseif ($null -ne $item.itemDesc -and $item.itemDesc.ToLower().Contains($searchLower)) { $match = $true }
        else {
            # Check receipts fields
            if ($null -ne $receipts) {
                foreach ($r in $receipts) {
                    if ($r.itemId -eq $item.id) {
                        if ($null -ne $r.grnNumber -and $r.grnNumber.ToLower().Contains($searchLower)) { $match = $true; break }
                        if ($null -ne $r.invoiceNumber -and $r.invoiceNumber.ToLower().Contains($searchLower)) { $match = $true; break }
                        if ($null -ne $r.supplierName -and $r.supplierName.ToLower().Contains($searchLower)) { $match = $true; break }
                    }
                }
            }
        }
        if ($match) { $filtered += $item }
    }
    $items = $filtered
}

# Filter by tab if provided
if ($null -ne $params.filter -and $params.filter -ne "" -and $params.filter -ne "all") {
    $filtered = @()
    foreach ($item in $items) {
        $isPendingDelivery = $item.reqQty -gt $item.recQty
        $isPendingPricing = $item.isPendingPricing
        $isCompleted = -not $isPendingDelivery -and -not $isPendingPricing
        
        $keep = $false
        if ($params.filter -eq "pending-delivery" -and $isPendingDelivery) { $keep = $true }
        elseif ($params.filter -eq "pending-pricing" -and (-not $isPendingDelivery) -and $isPendingPricing) { $keep = $true }
        elseif ($params.filter -eq "completed" -and $isCompleted) { $keep = $true }
        
        if ($keep) { $filtered += $item }
    }
    $items = $filtered
}

# Date parsing helper inside PowerShell
function Parse-MyDate($dateStr) {
    if ([string]::IsNullOrEmpty($dateStr)) { return [DateTime]::MinValue }
    $parsed = [DateTime]::MinValue
    if ($dateStr.Contains("-")) {
        $parts = $dateStr.Split("-")
        if ($parts.Length -eq 3 -and $parts[0].Length -eq 4) {
            try { return New-Object DateTime ([int]$parts[0]), ([int]$parts[1]), ([int]$parts[2]) } catch {}
        }
    }
    if ($dateStr.Contains("/")) {
        $parts = $dateStr.Split("/")
        if ($parts.Length -eq 3) {
            try { return New-Object DateTime ([int]$parts[2]), ([int]$parts[0]), ([int]$parts[1]) } catch {}
            try { return New-Object DateTime ([int]$parts[2]), ([int]$parts[1]), ([int]$parts[0]) } catch {}
        }
    }
    if ([DateTime]::TryParse($dateStr, [ref]$parsed)) {
        return $parsed
    }
    return [DateTime]::MinValue
}

# Filter by date range if provided
if ($null -ne $params.startDate -and $params.startDate -ne "") {
    $startD = Parse-MyDate $params.startDate
    if ($startD -ne [DateTime]::MinValue) {
        $filtered = @()
        foreach ($item in $items) {
            $itemD = Parse-MyDate $item.reqDate
            if ($itemD -ne [DateTime]::MinValue -and $itemD -ge $startD) {
                $filtered += $item
            }
        }
        $items = $filtered
    }
}

if ($null -ne $params.endDate -and $params.endDate -ne "") {
    $endD = Parse-MyDate $params.endDate
    if ($endD -ne [DateTime]::MinValue) {
        $filtered = @()
        foreach ($item in $items) {
            $itemD = Parse-MyDate $item.reqDate
            if ($itemD -ne [DateTime]::MinValue -and $itemD -le $endD) {
                $filtered += $item
            }
        }
        $items = $filtered
    }
}

# Filter by vehicle number if provided
if ($null -ne $params.vehicle -and $params.vehicle -ne "" -and $params.vehicle -ne "all") {
    $vehicleLower = $params.vehicle.ToLower().Trim()
    $filtered = @()
    foreach ($item in $items) {
        if ($null -ne $item.vehicleMachinery -and $item.vehicleMachinery.ToLower().Trim() -eq $vehicleLower) {
            $filtered += $item
        }
    }
    $items = $filtered
}

# Sort items before slicing
$desc = $params.order -eq "desc"
if ($params.sort -eq "mrnNum") {
    if ($desc) { $items = $items | Sort-Object @{Expression={$_.mrnNum}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={$_.mrnNum}} }
} elseif ($params.sort -eq "itemName") {
    if ($desc) { $items = $items | Sort-Object @{Expression={$_.itemName}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={$_.itemName}} }
} elseif ($params.sort -eq "vehicleMachinery") {
    if ($desc) { $items = $items | Sort-Object @{Expression={$_.vehicleMachinery}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={$_.vehicleMachinery}} }
} elseif ($params.sort -eq "reqQty") {
    if ($desc) { $items = $items | Sort-Object @{Expression={[double]$_.reqQty}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={[double]$_.reqQty}} }
} elseif ($params.sort -eq "recQty") {
    if ($desc) { $items = $items | Sort-Object @{Expression={[double]$_.recQty}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={[double]$_.recQty}} }
} elseif ($params.sort -eq "gap") {
    if ($desc) { $items = $items | Sort-Object @{Expression={[double]($_.reqQty - $_.recQty)}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={[double]($_.reqQty - $_.recQty)}} }
} else {
    # Default is sorted by reqDate
    if ($desc) { $items = $items | Sort-Object @{Expression={Parse-MyDate $_.reqDate}; Descending=$true} }
    else { $items = $items | Sort-Object @{Expression={Parse-MyDate $_.reqDate}} }
}

$total = $items.Count

# Slice in PowerShell if pagination is requested
if ($null -ne $params.limit -and $params.limit -gt 0) {
    $items = $items | Select-Object -Skip $params.skip -First $params.limit
}

@{ items = $items; receipts = $receipts; total = $total } | ConvertTo-Json -Depth 5 -Compress
    `;

    try {
        const inputData = { skip, limit, search, filter, startDate, endDate, vehicle, sort, order };
        const data = runDbQuery(psScript, inputData);
        const { items, receipts, total } = data;
        
        // Group receipts by itemId in O(N) using JS
        const receiptsByItem = {};
        if (receipts && Array.isArray(receipts)) {
            for (let r of receipts) {
                if (!receiptsByItem[r.itemId]) receiptsByItem[r.itemId] = [];
                receiptsByItem[r.itemId].push(r);
            }
        }
        
        function parseDate(dateStr) {
            if (!dateStr) return new Date(0);
            const str = String(dateStr);
            if (str.includes('-')) {
                const parts = str.split('-');
                if (parts.length === 3 && parts[0].length === 4) {
                    return new Date(parts[0], parts[1] - 1, parts[2]);
                }
            }
            if (str.includes('/')) {
                const parts = str.split('/');
                if (parts.length === 3) {
                    const month = parseInt(parts[0], 10);
                    const day = parseInt(parts[1], 10);
                    const year = parseInt(parts[2], 10);
                    return new Date(year, month - 1, day);
                }
            }
            const parsed = new Date(str);
            return isNaN(parsed.getTime()) ? new Date(0) : parsed;
        }

        if (items && Array.isArray(items)) {
            for (let item of items) {
                item.receipts = receiptsByItem[item.id] || [];
            }
        }
        
        if (page && limit) {
            res.json({
                items: items || [],
                total: total || 0,
                page,
                limit,
                totalPages: total && limit ? Math.ceil(total / limit) : 1
            });
        } else {
            res.json(items || []);
        }
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 2. Add new item (fully parameterized)
app.post('/api/items', (req, res) => {
    const { mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty } = req.body;
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = "INSERT INTO items (mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty) VALUES (?, ?, ?, ?, ?, ?)"
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.mrnNum) { [string]$params.mrnNum } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.reqDate) { [string]$params.reqDate } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.vehicleMachinery) { [string]$params.vehicleMachinery } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.itemName) { [string]$params.itemName } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.itemDesc) { [string]$params.itemDesc } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.reqQty) { [double]$params.reqQty } else { 0.0 })) | Out-Null
$cmd.ExecuteNonQuery() | Out-Null

$cmd.CommandText = "SELECT @@IDENTITY"
$newId = $cmd.ExecuteScalar()

$conn.Close()
@{ id = $newId } | ConvertTo-Json -Compress
    `;

    try {
        const inputData = { mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty };
        const result = runDbQuery(psScript, inputData);
        res.json({ success: true, id: result.id });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 3. Delete item (fully parameterized)
app.delete('/api/items/:id', (req, res) => {
    const id = parseInt(req.params.id);
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

# Delete receipts first (FK)
$cmd = $conn.CreateCommand()
$cmd.CommandText = "DELETE FROM receipts WHERE itemId = ?"
$cmd.Parameters.AddWithValue("?", [int]$params.id) | Out-Null
$cmd.ExecuteNonQuery() | Out-Null

# Delete item
$cmd.CommandText = "DELETE FROM items WHERE id = ?"
$cmd.Parameters.AddWithValue("?", [int]$params.id) | Out-Null
$cmd.ExecuteNonQuery() | Out-Null

$conn.Close()
@{ success = $true } | ConvertTo-Json -Compress
    `;

    try {
        runDbQuery(psScript, { id });
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 4. Add receipt (fully parameterized)
app.post('/api/items/:id/receipts', (req, res) => {
    const itemId = parseInt(req.params.id);
    const { qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice } = req.body;
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = "INSERT INTO receipts (itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
$cmd.Parameters.AddWithValue("?", [int]$params.itemId) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.qty) { [double]$params.qty } else { 0.0 })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.transactionType) { [string]$params.transactionType } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.deliveryDate) { [string]$params.deliveryDate } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.purchaseSource) { [string]$params.purchaseSource } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.grnNumber) { [string]$params.grnNumber } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.invoiceNumber) { [string]$params.invoiceNumber } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.invoiceDate) { [string]$params.invoiceDate } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.supplierName) { [string]$params.supplierName } else { "" })) | Out-Null

if ($null -ne $params.unitPrice -and "" -ne $params.unitPrice) {
    $cmd.Parameters.AddWithValue("?", [double]$params.unitPrice) | Out-Null
} else {
    $cmd.Parameters.AddWithValue("?", [System.DBNull]::Value) | Out-Null
}

$cmd.ExecuteNonQuery() | Out-Null

$cmd.CommandText = "SELECT @@IDENTITY"
$newId = $cmd.ExecuteScalar()

$conn.Close()
@{ id = $newId } | ConvertTo-Json -Compress
    `;

    try {
        const inputData = { itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice };
        const result = runDbQuery(psScript, inputData);
        res.json({ success: true, id: result.id });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5. Delete receipt (fully parameterized)
app.delete('/api/receipts/:id', (req, res) => {
    const id = parseInt(req.params.id);
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = "DELETE FROM receipts WHERE id = ?"
$cmd.Parameters.AddWithValue("?", [int]$params.id) | Out-Null
$cmd.ExecuteNonQuery() | Out-Null

$conn.Close()
@{ success = $true } | ConvertTo-Json -Compress
    `;

    try {
        runDbQuery(psScript, { id });
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5.5 Update item (fully parameterized)
app.put('/api/items/:id', (req, res) => {
    const id = parseInt(req.params.id);
    const { mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty } = req.body;
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = "UPDATE items SET mrnNum = ?, reqDate = ?, vehicleMachinery = ?, itemName = ?, itemDesc = ?, reqQty = ? WHERE id = ?"
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.mrnNum) { [string]$params.mrnNum } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.reqDate) { [string]$params.reqDate } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.vehicleMachinery) { [string]$params.vehicleMachinery } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.itemName) { [string]$params.itemName } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.itemDesc) { [string]$params.itemDesc } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.reqQty) { [double]$params.reqQty } else { 0.0 })) | Out-Null
$cmd.Parameters.AddWithValue("?", [int]$params.id) | Out-Null
$cmd.ExecuteNonQuery() | Out-Null

$conn.Close()
@{ success = $true } | ConvertTo-Json -Compress
    `;

    try {
        const inputData = { id, mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty };
        runDbQuery(psScript, inputData);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 5.6 Update receipt details & pricing (fully parameterized)
app.put('/api/receipts/:id', (req, res) => {
    const id = parseInt(req.params.id);
    const { qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice } = req.body;
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = "UPDATE receipts SET qty = ?, transactionType = ?, deliveryDate = ?, purchaseSource = ?, grnNumber = ?, invoiceNumber = ?, invoiceDate = ?, supplierName = ?, unitPrice = ? WHERE id = ?"
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.qty) { [double]$params.qty } else { 0.0 })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.transactionType) { [string]$params.transactionType } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.deliveryDate) { [string]$params.deliveryDate } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.purchaseSource) { [string]$params.purchaseSource } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.grnNumber) { [string]$params.grnNumber } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.invoiceNumber) { [string]$params.invoiceNumber } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.invoiceDate) { [string]$params.invoiceDate } else { "" })) | Out-Null
$cmd.Parameters.AddWithValue("?", $(if ($null -ne $params.supplierName) { [string]$params.supplierName } else { "" })) | Out-Null

if ($null -ne $params.unitPrice -and "" -ne $params.unitPrice) {
    $cmd.Parameters.AddWithValue("?", [double]$params.unitPrice) | Out-Null
} else {
    $cmd.Parameters.AddWithValue("?", [System.DBNull]::Value) | Out-Null
}

$cmd.Parameters.AddWithValue("?", [int]$params.id) | Out-Null
$cmd.ExecuteNonQuery() | Out-Null

$conn.Close()
@{ success = $true } | ConvertTo-Json -Compress
    `;

    try {
        const inputData = { id, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice };
        runDbQuery(psScript, inputData);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 6. Import Data (fully parameterized bulk importer)
app.post('/api/import', (req, res) => {
    const data = req.body;
    if (!Array.isArray(data)) {
        return res.status(400).json({ error: "Data must be an array of items" });
    }
    
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

foreach ($item in $params) {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "INSERT INTO items (mrnNum, reqDate, vehicleMachinery, itemName, itemDesc, reqQty) VALUES (?, ?, ?, ?, ?, ?)"
    $cmd.Parameters.AddWithValue("?", $(if ($null -ne $item.mrnNum) { [string]$item.mrnNum } else { "" })) | Out-Null
    $cmd.Parameters.AddWithValue("?", $(if ($null -ne $item.reqDate) { [string]$item.reqDate } else { "" })) | Out-Null
    $cmd.Parameters.AddWithValue("?", $(if ($null -ne $item.vehicleMachinery) { [string]$item.vehicleMachinery } else { "" })) | Out-Null
    $cmd.Parameters.AddWithValue("?", $(if ($null -ne $item.itemName) { [string]$item.itemName } else { "" })) | Out-Null
    $cmd.Parameters.AddWithValue("?", $(if ($null -ne $item.itemDesc) { [string]$item.itemDesc } else { "" })) | Out-Null
    $cmd.Parameters.AddWithValue("?", $(if ($null -ne $item.reqQty) { [double]$item.reqQty } else { 0.0 })) | Out-Null
    $cmd.ExecuteNonQuery() | Out-Null
    
    $cmd.CommandText = "SELECT @@IDENTITY"
    $itemId = $cmd.ExecuteScalar()
    
    if ($null -ne $item.receipts) {
        foreach ($r in $item.receipts) {
            $cmdR = $conn.CreateCommand()
            $cmdR.CommandText = "INSERT INTO receipts (itemId, qty, transactionType, deliveryDate, purchaseSource, grnNumber, invoiceNumber, invoiceDate, supplierName, unitPrice) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
            $cmdR.Parameters.AddWithValue("?", [int]$itemId) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.qty) { [double]$r.qty } else { 0.0 })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.transactionType) { [string]$r.transactionType } else { "" })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.deliveryDate) { [string]$r.deliveryDate } else { "" })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.purchaseSource) { [string]$r.purchaseSource } else { "" })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.grnNumber) { [string]$r.grnNumber } else { "" })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.invoiceNumber) { [string]$r.invoiceNumber } else { "" })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.invoiceDate) { [string]$r.invoiceDate } else { "" })) | Out-Null
            $cmdR.Parameters.AddWithValue("?", $(if ($null -ne $r.supplierName) { [string]$r.supplierName } else { "" })) | Out-Null
            
            if ($null -ne $r.unitPrice -and "" -ne $r.unitPrice) {
                $cmdR.Parameters.AddWithValue("?", [double]$r.unitPrice) | Out-Null
            } else {
                $cmdR.Parameters.AddWithValue("?", [System.DBNull]::Value) | Out-Null
            }
            $cmdR.ExecuteNonQuery() | Out-Null
        }
    }
}

$conn.Close()
@{ success = $true } | ConvertTo-Json -Compress
    `;

    try {
        runDbQuery(psScript, data);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// 7. Premium multi-sheet Excel Export Endpoint
app.get('/api/export/excel', (req, res) => {
    const psScript = `
$dbPath = '${DB_PATH.replace(/'/g, "''")}'
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$conn.Open()

# Get Items
$cmdItems = $conn.CreateCommand()
$cmdItems.CommandText = "SELECT * FROM items"
$readerItems = $cmdItems.ExecuteReader()
$items = @()
while ($readerItems.Read()) {
    $item = @{
        id = $readerItems["id"]
        mrnNum = $readerItems["mrnNum"]
        reqDate = $readerItems["reqDate"]
        vehicleMachinery = $readerItems["vehicleMachinery"]
        itemName = $readerItems["itemName"]
        itemDesc = $readerItems["itemDesc"]
        reqQty = $readerItems["reqQty"]
    }
    $items += $item
}
$readerItems.Close()

# Get Receipts
$cmdReceipts = $conn.CreateCommand()
$cmdReceipts.CommandText = "SELECT * FROM receipts"
$readerReceipts = $cmdReceipts.ExecuteReader()
$receipts = @()
while ($readerReceipts.Read()) {
    $receipt = @{
        id = $readerReceipts["id"]
        itemId = $readerReceipts["itemId"]
        qty = $readerReceipts["qty"]
        transactionType = $readerReceipts["transactionType"]
        deliveryDate = $readerReceipts["deliveryDate"]
        purchaseSource = $readerReceipts["purchaseSource"]
        grnNumber = $readerReceipts["grnNumber"]
        invoiceNumber = $readerReceipts["invoiceNumber"]
        invoiceDate = $readerReceipts["invoiceDate"]
        supplierName = $readerReceipts["supplierName"]
        unitPrice = $readerReceipts["unitPrice"]
    }
    $receipts += $receipt
}
$readerReceipts.Close()
$conn.Close()

@{ items = $items; receipts = $receipts } | ConvertTo-Json -Depth 5 -Compress
    `;

    try {
        const data = runDbQuery(psScript);
        const { items, receipts } = data;
        
        const receiptsByItem = {};
        if (receipts && Array.isArray(receipts)) {
            for (let r of receipts) {
                if (!receiptsByItem[r.itemId]) receiptsByItem[r.itemId] = [];
                receiptsByItem[r.itemId].push(r);
            }
        }

        const wb = XLSX.utils.book_new();

        // Build main sheet data
        const itemsSheetData = [];
        itemsSheetData.push([
            "MRN Number", "Request Date", "Vehicle/Machinery", "Item Name", "Item Description", 
            "Requested Qty", "Received Qty", "Receive Date", "Purchase Source", "Qty Gap", 
            "Date Gap (Days)", "Status", "GRN Number", "Invoice Number", "Invoice Date", 
            "Supplier Name", "Unit Price (Rs.)", "Total Price (Rs.)"
        ]);

        const supplierSpend = {};
        let totalSpend = 0;
        let activeSuppliers = new Set();
        let pricedCount = 0;
        let unpricedCount = 0;

        if (items && Array.isArray(items)) {
            for (let item of items) {
                item.receipts = receiptsByItem[item.id] || [];
                const recQty = item.receipts.reduce((sum, r) => sum + r.qty, 0);
                const recQtyRounded = Math.round(recQty * 100) / 100;
                
                let recDate = "";
                if (item.receipts.length > 0) {
                    const sorted = [...item.receipts].sort((a, b) => new Date(b.deliveryDate) - new Date(a.deliveryDate));
                    recDate = sorted[0].deliveryDate;
                }
                
                const uniqueSources = [...new Set(item.receipts.map(r => r.purchaseSource).filter(Boolean))].join(' & ');
                const qtyGap = Math.round((item.reqQty - recQtyRounded) * 100) / 100;
                
                let dateGapDays = "";
                if (recDate) {
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

                const grns = [...new Set(item.receipts.map(r => r.grnNumber).filter(Boolean))].join('; ');
                const invoices = [...new Set(item.receipts.map(r => r.invoiceNumber).filter(Boolean))].join('; ');
                const invoiceDates = [...new Set(item.receipts.map(r => r.invoiceDate).filter(Boolean))].filter(d => d && d !== '1899-12-30').join('; ');
                const suppliers = [...new Set(item.receipts.map(r => r.supplierName).filter(Boolean))].join('; ');
                
                let totalUnitPrice = "";
                let totalPrice = 0;
                let hasPricing = false;

                const pricedReceipts = item.receipts.filter(r => r.unitPrice);
                if (pricedReceipts.length > 0) {
                    totalUnitPrice = pricedReceipts.map(r => r.unitPrice).join('; ');
                    totalPrice = item.receipts.reduce((sum, r) => {
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
        }

        const wsItems = XLSX.utils.aoa_to_sheet(itemsSheetData);
        XLSX.utils.book_append_sheet(wb, wsItems, "Requests & Deliveries");

        // Build summary sheet data
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

app.listen(PORT, '0.0.0.0', () => {
    console.log(`Delivery Monitor Server running at http://localhost:${PORT}`);
    console.log(`Database: ${DB_PATH}`);
    
    const networkInterfaces = os.networkInterfaces();
    const localIPs = [];
    for (const interfaceName in networkInterfaces) {
        const interfaces = networkInterfaces[interfaceName];
        for (const iface of interfaces) {
            if (iface.family === 'IPv4' && !iface.internal) {
                localIPs.push(iface.address);
            }
        }
    }
    
    if (localIPs.length > 0) {
        console.log(`\nTo access this server from other computers on your network:`);
        localIPs.forEach(ip => {
            console.log(`  http://${ip}:${PORT}`);
        });
        console.log();
    }
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
        const backupPath = path.join(BACKUP_DIR, `inventory_backup_${timestamp}.accdb`);
        
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
            .filter(file => file.startsWith('inventory_backup_') && file.endsWith('.accdb'))
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
