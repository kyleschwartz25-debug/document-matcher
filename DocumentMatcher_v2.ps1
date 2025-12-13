# Document Matcher Tool v2 - PowerShell GUI with Manual Data Entry
# Compares Sales Orders (SO) with Purchase Orders (PO) for matching

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Compare Ship To addresses
function Compare-ShipTo {
    param(
        $SOAddress,
        $POAddress
    )
    
    $issues = @()
    
    if ($SOAddress.Name -ne $POAddress.Name) {
        $issues += "Ship To Name mismatch: SO='$($SOAddress.Name)' vs PO='$($POAddress.Name)'"
    }
    if ($SOAddress.Address -ne $POAddress.Address) {
        $issues += "Ship To Address mismatch: SO='$($SOAddress.Address)' vs PO='$($POAddress.Address)'"
    }
    if ($SOAddress.City -ne $POAddress.City) {
        $issues += "Ship To City mismatch: SO='$($SOAddress.City)' vs PO='$($POAddress.City)'"
    }
    if ($SOAddress.State -ne $POAddress.State) {
        $issues += "Ship To State mismatch: SO='$($SOAddress.State)' vs PO='$($POAddress.State)'"
    }
    if ($SOAddress.Zip -ne $POAddress.Zip) {
        $issues += "Ship To Zip mismatch: SO='$($SOAddress.Zip)' vs PO='$($POAddress.Zip)'"
    }
    
    return $issues
}

# Compare line items
function Compare-LineItems {
    param(
        $SOItems,
        $POItems
    )
    
    $issues = @()
    
    if ($SOItems.Count -ne $POItems.Count) {
        $issues += "Line count mismatch: SO has $($SOItems.Count) items, PO has $($POItems.Count) items"
    }
    
    $count = [Math]::Min($SOItems.Count, $POItems.Count)
    for ($i = 0; $i -lt $count; $i++) {
        $so = $SOItems[$i]
        $po = $POItems[$i]
        
        if ($so.SKU -ne $po.SKU) {
            $issues += "Line $($i+1) SKU mismatch: SO='$($so.SKU)' vs PO='$($po.SKU)'"
        }
        if ($so.Desc -ne $po.Desc) {
            $issues += "Line $($i+1) Description mismatch: SO='$($so.Desc)' vs PO='$($po.Desc)'"
        }
        if ($so.Qty -ne $po.Qty) {
            $issues += "Line $($i+1) Qty mismatch: SO=$($so.Qty) vs PO=$($po.Qty)"
        }
    }
    
    return $issues
}

# === GUI SETUP ===

$form = New-Object System.Windows.Forms.Form
$form.Text = "SO vs PO Document Matcher"
$form.Size = New-Object System.Drawing.Size(1200, 900)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Sales Order (SO) vs Purchase Order (PO) Matcher"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)
$titleLabel.Size = New-Object System.Drawing.Size(1160, 30)
$form.Controls.Add($titleLabel)

# === LEFT PANEL: SO DATA ===
$soLabel = New-Object System.Windows.Forms.Label
$soLabel.Text = "SALES ORDER (SO)"
$soLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$soLabel.Location = New-Object System.Drawing.Point(20, 60)
$soLabel.Size = New-Object System.Drawing.Size(500, 25)
$soLabel.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($soLabel)

# SO Ship To
$soShipToGroupBox = New-Object System.Windows.Forms.GroupBox
$soShipToGroupBox.Text = "Ship To Address"
$soShipToGroupBox.Location = New-Object System.Drawing.Point(20, 90)
$soShipToGroupBox.Size = New-Object System.Drawing.Size(550, 130)
$form.Controls.Add($soShipToGroupBox)

# SO Ship To fields
$controls = @()
$y = 20
@("Name", "Address", "City", "State", "Zip") | ForEach-Object {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "$_`:"
    $label.Location = New-Object System.Drawing.Point(10, $y)
    $label.Size = New-Object System.Drawing.Size(60, 20)
    $soShipToGroupBox.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(75, $y)
    $textBox.Size = New-Object System.Drawing.Size(460, 20)
    $textBox.Name = "SO_ShipTo_$_"
    $soShipToGroupBox.Controls.Add($textBox)
    
    $y += 22
}

# SO Line Items
$soItemsLabel = New-Object System.Windows.Forms.Label
$soItemsLabel.Text = "Line Items (Dropship Only)"
$soItemsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$soItemsLabel.Location = New-Object System.Drawing.Point(20, 230)
$soItemsLabel.Size = New-Object System.Drawing.Size(550, 20)
$form.Controls.Add($soItemsLabel)

$soItemsGrid = New-Object System.Windows.Forms.DataGridView
$soItemsGrid.Location = New-Object System.Drawing.Point(20, 255)
$soItemsGrid.Size = New-Object System.Drawing.Size(550, 320)
$soItemsGrid.AllowUserToAddRows = $true
$soItemsGrid.AllowUserToDeleteRows = $true
$soItemsGrid.ColumnCount = 3
$soItemsGrid.Columns[0].Name = "SKU"
$soItemsGrid.Columns[0].Width = 150
$soItemsGrid.Columns[1].Name = "Description"
$soItemsGrid.Columns[1].Width = 250
$soItemsGrid.Columns[2].Name = "Qty"
$soItemsGrid.Columns[2].Width = 80
$form.Controls.Add($soItemsGrid)

# === RIGHT PANEL: PO DATA ===
$poLabel = New-Object System.Windows.Forms.Label
$poLabel.Text = "PURCHASE ORDER (PO)"
$poLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$poLabel.Location = New-Object System.Drawing.Point(630, 60)
$poLabel.Size = New-Object System.Drawing.Size(550, 25)
$poLabel.ForeColor = [System.Drawing.Color]::DarkGreen
$form.Controls.Add($poLabel)

# PO Ship To
$poShipToGroupBox = New-Object System.Windows.Forms.GroupBox
$poShipToGroupBox.Text = "Ship To Address"
$poShipToGroupBox.Location = New-Object System.Drawing.Point(630, 90)
$poShipToGroupBox.Size = New-Object System.Drawing.Size(550, 130)
$form.Controls.Add($poShipToGroupBox)

# PO Ship To fields
$y = 20
@("Name", "Address", "City", "State", "Zip") | ForEach-Object {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "$_`:"
    $label.Location = New-Object System.Drawing.Point(10, $y)
    $label.Size = New-Object System.Drawing.Size(60, 20)
    $poShipToGroupBox.Controls.Add($label)
    
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(75, $y)
    $textBox.Size = New-Object System.Drawing.Size(460, 20)
    $textBox.Name = "PO_ShipTo_$_"
    $poShipToGroupBox.Controls.Add($textBox)
    
    $y += 22
}

# PO Line Items
$poItemsLabel = New-Object System.Windows.Forms.Label
$poItemsLabel.Text = "Line Items"
$poItemsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$poItemsLabel.Location = New-Object System.Drawing.Point(630, 230)
$poItemsLabel.Size = New-Object System.Drawing.Size(550, 20)
$form.Controls.Add($poItemsLabel)

$poItemsGrid = New-Object System.Windows.Forms.DataGridView
$poItemsGrid.Location = New-Object System.Drawing.Point(630, 255)
$poItemsGrid.Size = New-Object System.Drawing.Size(550, 320)
$poItemsGrid.AllowUserToAddRows = $true
$poItemsGrid.AllowUserToDeleteRows = $true
$poItemsGrid.ColumnCount = 3
$poItemsGrid.Columns[0].Name = "SKU"
$poItemsGrid.Columns[0].Width = 150
$poItemsGrid.Columns[1].Name = "Description"
$poItemsGrid.Columns[1].Width = 250
$poItemsGrid.Columns[2].Name = "Qty"
$poItemsGrid.Columns[2].Width = 80
$form.Controls.Add($poItemsGrid)

# === BOTTOM PANEL: RESULTS ===
$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Text = "Comparison Results:"
$resultsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$resultsLabel.Location = New-Object System.Drawing.Point(20, 590)
$resultsLabel.Size = New-Object System.Drawing.Size(1160, 20)
$form.Controls.Add($resultsLabel)

$resultsText = New-Object System.Windows.Forms.RichTextBox
$resultsText.Location = New-Object System.Drawing.Point(20, 615)
$resultsText.Size = New-Object System.Drawing.Size(1160, 200)
$resultsText.ReadOnly = $true
$resultsText.BackColor = [System.Drawing.Color]::LightGray
$resultsText.Font = New-Object System.Drawing.Font("Courier New", 9)
$form.Controls.Add($resultsText)

# === BUTTONS ===
$compareBtn = New-Object System.Windows.Forms.Button
$compareBtn.Text = "COMPARE"
$compareBtn.Location = New-Object System.Drawing.Point(20, 825)
$compareBtn.Size = New-Object System.Drawing.Size(140, 40)
$compareBtn.BackColor = [System.Drawing.Color]::LightGreen
$compareBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($compareBtn)

$copyBtn = New-Object System.Windows.Forms.Button
$copyBtn.Text = "COPY TO CLIPBOARD"
$copyBtn.Location = New-Object System.Drawing.Point(180, 825)
$copyBtn.Size = New-Object System.Drawing.Size(140, 40)
$copyBtn.BackColor = [System.Drawing.Color]::LightBlue
$copyBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($copyBtn)

$clearBtn = New-Object System.Windows.Forms.Button
$clearBtn.Text = "CLEAR ALL"
$clearBtn.Location = New-Object System.Drawing.Point(340, 825)
$clearBtn.Size = New-Object System.Drawing.Size(140, 40)
$clearBtn.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($clearBtn)

$exitBtn = New-Object System.Windows.Forms.Button
$exitBtn.Text = "EXIT"
$exitBtn.Location = New-Object System.Drawing.Point(1040, 825)
$exitBtn.Size = New-Object System.Drawing.Size(140, 40)
$form.Controls.Add($exitBtn)

# === EVENT HANDLERS ===

$compareBtn.Add_Click({
    $resultsText.Clear()
    
    # Get SO Ship To
    $soAddress = @{
        Name = $soShipToGroupBox.Controls | Where-Object { $_.Name -eq "SO_ShipTo_Name" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        Address = $soShipToGroupBox.Controls | Where-Object { $_.Name -eq "SO_ShipTo_Address" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        City = $soShipToGroupBox.Controls | Where-Object { $_.Name -eq "SO_ShipTo_City" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        State = $soShipToGroupBox.Controls | Where-Object { $_.Name -eq "SO_ShipTo_State" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        Zip = $soShipToGroupBox.Controls | Where-Object { $_.Name -eq "SO_ShipTo_Zip" } | Select-Object -First 1 | ForEach-Object { $_.Text }
    }
    
    # Get PO Ship To
    $poAddress = @{
        Name = $poShipToGroupBox.Controls | Where-Object { $_.Name -eq "PO_ShipTo_Name" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        Address = $poShipToGroupBox.Controls | Where-Object { $_.Name -eq "PO_ShipTo_Address" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        City = $poShipToGroupBox.Controls | Where-Object { $_.Name -eq "PO_ShipTo_City" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        State = $poShipToGroupBox.Controls | Where-Object { $_.Name -eq "PO_ShipTo_State" } | Select-Object -First 1 | ForEach-Object { $_.Text }
        Zip = $poShipToGroupBox.Controls | Where-Object { $_.Name -eq "PO_ShipTo_Zip" } | Select-Object -First 1 | ForEach-Object { $_.Text }
    }
    
    # Get SO Line Items
    $soItems = @()
    foreach ($row in $soItemsGrid.Rows) {
        if ($row.Cells[0].Value -ne $null -and $row.Cells[0].Value -ne "") {
            $soItems += @{
                SKU = $row.Cells[0].Value
                Desc = $row.Cells[1].Value
                Qty = [int]($row.Cells[2].Value -as [int])
            }
        }
    }
    
    # Get PO Line Items
    $poItems = @()
    foreach ($row in $poItemsGrid.Rows) {
        if ($row.Cells[0].Value -ne $null -and $row.Cells[0].Value -ne "") {
            $poItems += @{
                SKU = $row.Cells[0].Value
                Desc = $row.Cells[1].Value
                Qty = [int]($row.Cells[2].Value -as [int])
            }
        }
    }
    
    # Compare
    $shipToIssues = Compare-ShipTo -SOAddress $soAddress -POAddress $poAddress
    $itemIssues = Compare-LineItems -SOItems $soItems -POItems $poItems
    $allIssues = $shipToIssues + $itemIssues
    
    # Display results
    $resultsText.Clear()
    
    if ($allIssues.Count -eq 0) {
        $resultsText.SelectionColor = [System.Drawing.Color]::Green
        $resultsText.AppendText("STATUS: COMPLETE MATCH`n`n")
    } else {
        $resultsText.SelectionColor = [System.Drawing.Color]::Red
        $resultsText.AppendText("STATUS: MISMATCH - MANUAL REVIEW REQUIRED`n`n")
        $resultsText.SelectionColor = [System.Drawing.Color]::Black
        $resultsText.AppendText("Issues Found:`n")
        $resultsText.AppendText(("-" * 100) + "`n")
        foreach ($issue in $allIssues) {
            $resultsText.AppendText("  - $issue`n")
        }
    }
    
    $resultsText.SelectionColor = [System.Drawing.Color]::Black
})

$copyBtn.Add_Click({
    if ($resultsText.Text -eq "") {
        [System.Windows.Forms.MessageBox]::Show("No results to copy. Run a comparison first.", "No Results", "OK") | Out-Null
        return
    }
    [System.Windows.Forms.Clipboard]::SetText($resultsText.Text)
    [System.Windows.Forms.MessageBox]::Show("Results copied to clipboard!", "Success", "OK") | Out-Null
})

$clearBtn.Add_Click({
    $soShipToGroupBox.Controls | ForEach-Object {
        if ($_ -is [System.Windows.Forms.TextBox]) { $_.Text = "" }
    }
    $poShipToGroupBox.Controls | ForEach-Object {
        if ($_ -is [System.Windows.Forms.TextBox]) { $_.Text = "" }
    }
    $soItemsGrid.Rows.Clear()
    $poItemsGrid.Rows.Clear()
    $resultsText.Clear()
})

$exitBtn.Add_Click({
    $form.Close()
})

$form.ShowDialog() | Out-Null
