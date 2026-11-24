Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ApiKey = $env:TWELVEDATA_API_KEY
if (-not $ApiKey) {throw 'Environment variable TWELVEDATA_API_KEY is not set'}

$symbols = [ordered]@{
    'EMR' = 'Emerson Electric'
    'ROK' = 'Rockwell Automation'
    'HON' = 'Honeywell'
    'ETN' = 'Eaton'
    'PTC' = 'PTC Inc'
    'RAL' = 'Ralliant'
}

$culture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')

# --- API call, one batched call for all symbols ---
$symbolList = $symbols.Keys -join ','
$uri = "https://api.twelvedata.com/quote?symbol=$symbolList&apikey=$ApiKey"

try {
    $resp = Invoke-RestMethod -Uri $uri -Method Get
} catch {
    throw "API call failed: $($_.Exception.Message)"
}

if ($resp.status -eq 'error') {throw "TwelveData error: $($resp.message)"}

$rows = foreach ($sym in $symbols.Keys) {
    $q = $resp.$sym
    if(-not $q -or $q.status -eq 'error') {
        [PSCustomObject]@{
            Symbol = $sym; Name = $symbols[$sym]; Open = $null; Close = $null
            Change = $null; PctChange = $null; AsOf = "Error: $($q.message)" 
        }
        continue 
    }
    $open = [double]$q.open
    $close = [double]$q.close
    [PSCustomObject]@{
        Symbol = $sym
        Name = $symbols[$sym]
        Open = $open
        Close = $close
        Change = $close - $open
        PctChange = if ($open -ne 0) {(($close - $open) / $open) * 100} else {$null}
        AsOf = [string]$q.datetime
    }
}

# --- Formatting helpers ---
function Format-Money($v) {
    if($null -eq $v) { 'na'} else {$v.ToString('C2', $culture)}
}

function Format-Pct($v) {
    if ($null -eq $v) {'n/a'} else {$v.ToString('+0.00;-0.00;0.00', $culture) + '%'}
}


# --- Session date from the EMR row, shown in the footer ---
$emrRow = $rows | Where-Object {$_.Symbol -eq 'EMR'} | Select-Object -First 1
$asOf = if ($emrRow) {$emrRow.AsOf} else {'n/a'}

# --- Window ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "EMR & Peers - Close $(Get-Date -Format 'MMM d, yyyy')"
$form.Size = New-Object System.Drawing.Size(900, 330)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToResizeRows = $false
$grid.RowHeadersVisible = $false
$grid.MultiSelect = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'None'
$grid.CellBorderStyle = 'SingleHorizontal'
$grid.GridColor = [System.Drawing.Color]::FromArgb(217, 217, 217)
$grid.AutoSizeColumnsMode = 'Fill'
$grid.RowTemplate.Height = 30
$grid.EnableHeadersVisualStyles = $false

# --- Header styling ---
$grid.ColumnHeadersBorderStyle = 'None'
$grid.ColumnHeadersHeight = 34
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(31, 56, 100)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.Alignment = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter

# --- Body styling ---
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(242, 242, 242)
$grid.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(8, 0, 8, 0)
$grid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(33, 33, 33)

# --- Colulmns added manually, no data binding needed
[void]$grid.Columns.Add('Symbol', 'Symbol')
[void]$grid.Columns.Add('Company', 'Company')
[void]$grid.Columns.Add('Open', 'Open')
[void]$grid.Columns.Add('Close', 'Close')
[void]$grid.Columns.Add('$ Change', '$ Change')
[void]$grid.Columns.Add('% Change', '% Change')
[void]$grid.Columns.Add('As Of', 'As Of')

$right = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleRight
$center = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleCenter
$left = [System.Windows.Forms.DataGridViewContentAlignment]::MiddleLeft

$grid.Columns[2].DefaultCellStyle.Alignment = $right
$grid.Columns[3].DefaultCellStyle.Alignment = $right
$grid.Columns[4].DefaultCellStyle.Alignment = $right
$grid.Columns[5].DefaultCellStyle.Alignment = $right
$grid.Columns[6].DefaultCellStyle.Alignment = $center
$grid.Columns[0].DefaultCellStyle.Alignment = $left
$grid.Columns[1].DefaultCellStyle.Alignment = $left

# --- Add rows then color gains / losses and highlight EMR
$green = [System.Drawing.Color]::FromArgb(0, 128, 0)
$red = [System.Drawing.Color]::FromArgb(192, 0, 0)

foreach ($r in $rows) {
    $idx = $grid.Rows.Add(
        $r.Symbol,
        $r.Name,
        (Format-Money $r.Open),
        (Format-Money $r.Close),
        (Format-Money $r.Change),
        (Format-Pct $r.PctChange),
        $r.AsOf
    )
    $row = $grid.Rows[$idx]

    if($null -ne $r.Change) {
        $row.Cells[4].Style.ForeColor = if ($r.Change -ge 0) {$green} else {$red}
        $row.Cells[5].Style.ForeColor = if ($r.PctChange -ge 0) {$green} else {$red}
    }

    if ($r.Symbol -eq 'EMR') {
        $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 242, 205)
        $row.DefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    }
}


# --- Bottom bar: footer text on the left ---
$bottomBar = New-Object System.Windows.Forms.Panel
$bottomBar.Dock = 'Bottom'
$bottomBar.Height = 44

$footer = New-Object System.Windows.Forms.Label
$footer.Dock = 'Fill'
$footer.TextAlign = 'MiddleLeft'
$footer.Padding = New-Object System.Windows.Forms.Padding(10, 0, 0, 0)
$footer.ForeColor = [System.Drawing.Color]::FromArgb(89, 89, 89)
$footer.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$footer.Text = "Session date: $asOf Pulled: $(Get-Date -Format 'h:mm tt')"

$bottomBar.Controls.Add($footer)

$form.Controls.Add($grid)
$form.Controls.Add($bottomBar)
[void]$form.ShowDialog()

