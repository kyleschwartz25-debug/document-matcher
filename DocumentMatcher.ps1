# Document Matcher Tool - PowerShell GUI
# Compares Sales Orders (SO) with Purchase Orders (PO) for matching

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# PDF text extraction using Word COM object (Windows native)
function Extract-TextFromPDF {
    param([string]$PdfPath)
    
    try {
        # Try using Word to open and extract text
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        $doc = $word.Documents.Open($PdfPath, $false, $true)
        $text = $doc.Content.Text
        
        $doc.Close()
        $word.Quit()
        
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        
        return $text
    } catch {
        # Fallback: read raw bytes and extract text strings
        try {
            $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
            $text = [System.Text.Encoding]::UTF8.GetString($bytes)
            # Extract readable strings
            $text = $text -replace '[^\x20-\x7E\n\r]', ''
            return $text
        } catch {
            return "ERROR: Could not extract PDF content from $PdfPath. Try pasting text manually."
        }
    }
}

# Parse Ship To information
function Parse-ShipTo {
    param([string]$Text)
    
    $result = @{
        Name = ""
        Address = ""
        City = ""
        State = ""
        Zip = ""
    }
    
    # Find Ship To section (more flexible pattern)
    if ($Text -match "(?:Ship To|Ship to):?[^\n]*\n([\s\S]{0,800}?)(?:\n\n|Notes:|Vendor:|Bill To:)") {
        $shipToSection = $matches[1]
        
        # Remove extra whitespace and split into lines
        $lines = $shipToSection -split "`n" | 
                 ForEach-Object { $_.Trim() } | 
                 Where-Object { $_ -ne "" -and $_ -notmatch "^UNITED STATES$|^Phone:|^Email:" }
        
        if ($lines.Count -ge 2) {
            $result.Name = $lines[0]
            $result.Address = if ($lines.Count -gt 1) { $lines[1] } else { "" }
            
            # Find City, State, Zip in remaining lines
            $remainingText = ($lines | Select-Object -Skip 2) -join " "
            
            if ($remainingText -match "([^,]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)") {
                $result.City = $matches[1].Trim()
                $result.State = $matches[2].Trim()
                $result.Zip = $matches[3].Trim()
            }
        }
    }
    
    return $result
}

# Parse line items from PDF text
function Parse-LineItems {
    param([string]$Text, [bool]$IsInvoice = $false)
    
    $items = @()
    $lines = $Text -split "`n"
    $inItemsSection = $false
    $lineNum = 0
    
    foreach ($line in $lines) {
        $lineNum++
        
        # Detect start of items section - look for table headers
        if ($line -match "Item.*Number.*Description.*Qty" -or 
            ($line -match "^Item" -and $lineNum -lt 50)) {
            $inItemsSection = $true
            continue
        }
        
        # Detect end of items section
        if ($inItemsSection) {
            if ($line -match "Subtotal|Shipping|Tariff|Total|Notes|approval" -or 
                ($line -match "^\s*$" -and $items.Count -gt 0)) {
                break
            }
        }
        
        if ($inItemsSection -and $line.Trim() -ne "") {
            # For invoice (SO): Extract dropship items only
            if ($IsInvoice) {
                if ($line -match "Drop.?Ship") {
                    # Try to extract: line# | type | sku | description | price | qty | total
                    $patterns = @(
                        "Drop.?Ship\s+([\w\-\.]+)\s+(.*?)\s+\d+\.\d{2}\s+(\d+)\s+",
                        "Drop.?Ship\s+([\w\-\.]+)\s+(.*?)\s+(\d+)\s+ea"
                    )
                    
                    foreach ($pattern in $patterns) {
                        if ($line -match $pattern) {
                            $sku = $matches[1].Trim()
                            $desc = ($matches[2].Trim() -replace "\s+", " ")
                            $qty = [int]$matches[3]
                            
                            $items += @{
                                SKU = $sku
                                Description = $desc
                                Qty = $qty
                            }
                            break
                        }
                    }
                }
            } else {
                # For PO: Extract items
                # Pattern: number | sku | description | qty | uom
                $patterns = @(
                    "^\s*(\d+)\s+([\w\-\.]+)\s+(.*?)\s+(\d+)\s+ea\s*$",
                    "^\s*(\d+)\s+([\w\-\.]+)\s+(.*?)\s+(\d+)\s*$"
                )
                
                foreach ($pattern in $patterns) {
                    if ($line -match $pattern) {
                        $sku = $matches[2].Trim()
                        $desc = ($matches[3].Trim() -replace "\s+", " ")
                        $qty = [int]$matches[4]
                        
                        # Skip non-dropship items
                        if (-not ($desc -match "Subtotal|Shipping|Freight|Tariff|Tax|President|Fax")) {
                            $items += @{
                                SKU = $sku
                                Description = $desc
                                Qty = $qty
                            }
                        }
                        break
                    }
                }
            }
        }
    }
    
    return $items
}

# Main comparison function
function Compare-Documents {
    param([string]$SOPath, [string]$POPath)
    
    $result = @{
        Match = $false
        Issues = @()
        SOShipTo = @{}
        POShipTo = @{}
        SOItems = @()
        POItems = @()
    }
    
    # Extract text
    $result.SOText = Extract-TextFromPDF -PdfPath $SOPath
    $result.POText = Extract-TextFromPDF -PdfPath $POPath
    
    if ($result.SOText -match "ERROR:" -or $result.POText -match "ERROR:") {
        $result.Issues += $result.SOText
        $result.Issues += $result.POText
        return $result
    }
    
    # Parse Ship To
    $result.SOShipTo = Parse-ShipTo -Text $result.SOText
    $result.POShipTo = Parse-ShipTo -Text $result.POText
    
    # Parse line items
    $result.SOItems = Parse-LineItems -Text $result.SOText -IsInvoice $true
    $result.POItems = Parse-LineItems -Text $result.POText -IsInvoice $false
    
    # Compare Ship To
    if ($result.SOShipTo.Name -ne $result.POShipTo.Name) {
        $result.Issues += "Ship To Name: SO='$($result.SOShipTo.Name)' / PO='$($result.POShipTo.Name)'"
    }
    
    if ($result.SOShipTo.Address -ne $result.POShipTo.Address) {
        $result.Issues += "Ship To Address: SO='$($result.SOShipTo.Address)' / PO='$($result.POShipTo.Address)'"
    }
    
    if ($result.SOShipTo.City -ne $result.POShipTo.City) {
        $result.Issues += "Ship To City: SO='$($result.SOShipTo.City)' / PO='$($result.POShipTo.City)'"
    }
    
    if ($result.SOShipTo.State -ne $result.POShipTo.State) {
        $result.Issues += "Ship To State: SO='$($result.SOShipTo.State)' / PO='$($result.POShipTo.State)'"
    }
    
    if ($result.SOShipTo.Zip -ne $result.POShipTo.Zip) {
        $result.Issues += "Ship To Zip: SO='$($result.SOShipTo.Zip)' / PO='$($result.POShipTo.Zip)'"
    }
    
    # Compare line items
    if ($result.SOItems.Count -ne $result.POItems.Count) {
        $result.Issues += "Line item count: SO=$($result.SOItems.Count) / PO=$($result.POItems.Count)"
    } else {
        for ($i = 0; $i -lt $result.SOItems.Count; $i++) {
            $soItem = $result.SOItems[$i]
            $poItem = $result.POItems[$i]
            
            if ($soItem.SKU -ne $poItem.SKU) {
                $result.Issues += "Line $($i+1) SKU: SO='$($soItem.SKU)' / PO='$($poItem.SKU)'"
            }
            if ($soItem.Description -ne $poItem.Description) {
                $result.Issues += "Line $($i+1) Desc: SO='$($soItem.Description)' / PO='$($poItem.Description)'"
            }
            if ($soItem.Qty -ne $poItem.Qty) {
                $result.Issues += "Line $($i+1) Qty: SO=$($soItem.Qty) / PO=$($poItem.Qty)"
            }
        }
    }
    
    $result.Match = ($result.Issues.Count -eq 0)
    
    return $result
}

# === GUI SETUP ===

$form = New-Object System.Windows.Forms.Form
$form.Text = "Document Matcher - SO vs PO"
$form.Size = New-Object System.Drawing.Size(950, 750)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::White

# Title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Sales Order & Purchase Order Matcher"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.Location = New-Object System.Drawing.Point(20, 20)
$titleLabel.Size = New-Object System.Drawing.Size(910, 30)
$form.Controls.Add($titleLabel)

# SO Selection
$soLabel = New-Object System.Windows.Forms.Label
$soLabel.Text = "Sales Order (SO):"
$soLabel.Location = New-Object System.Drawing.Point(20, 70)
$soLabel.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($soLabel)

$soTextBox = New-Object System.Windows.Forms.TextBox
$soTextBox.Location = New-Object System.Drawing.Point(170, 70)
$soTextBox.Size = New-Object System.Drawing.Size(710, 20)
$soTextBox.ReadOnly = $true
$form.Controls.Add($soTextBox)

$soBrowseBtn = New-Object System.Windows.Forms.Button
$soBrowseBtn.Text = "Browse"
$soBrowseBtn.Location = New-Object System.Drawing.Point(890, 70)
$soBrowseBtn.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($soBrowseBtn)

# PO Selection
$poLabel = New-Object System.Windows.Forms.Label
$poLabel.Text = "Purchase Order (PO):"
$poLabel.Location = New-Object System.Drawing.Point(20, 110)
$poLabel.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($poLabel)

$poTextBox = New-Object System.Windows.Forms.TextBox
$poTextBox.Location = New-Object System.Drawing.Point(170, 110)
$poTextBox.Size = New-Object System.Drawing.Size(710, 20)
$poTextBox.ReadOnly = $true
$form.Controls.Add($poTextBox)

$poBrowseBtn = New-Object System.Windows.Forms.Button
$poBrowseBtn.Text = "Browse"
$poBrowseBtn.Location = New-Object System.Drawing.Point(890, 110)
$poBrowseBtn.Size = New-Object System.Drawing.Size(60, 20)
$form.Controls.Add($poBrowseBtn)

# Results Display
$resultsLabel = New-Object System.Windows.Forms.Label
$resultsLabel.Text = "Results:"
$resultsLabel.Location = New-Object System.Drawing.Point(20, 150)
$resultsLabel.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($resultsLabel)

$resultsText = New-Object System.Windows.Forms.RichTextBox
$resultsText.Location = New-Object System.Drawing.Point(20, 180)
$resultsText.Size = New-Object System.Drawing.Size(930, 420)
$resultsText.ReadOnly = $true
$resultsText.BackColor = [System.Drawing.Color]::LightGray
$resultsText.Font = New-Object System.Drawing.Font("Courier New", 9)
$form.Controls.Add($resultsText)

# Compare Button
$compareBtn = New-Object System.Windows.Forms.Button
$compareBtn.Text = "Compare Documents"
$compareBtn.Location = New-Object System.Drawing.Point(20, 620)
$compareBtn.Size = New-Object System.Drawing.Size(150, 40)
$compareBtn.BackColor = [System.Drawing.Color]::LightGreen
$compareBtn.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($compareBtn)

# Copy Button
$copyBtn = New-Object System.Windows.Forms.Button
$copyBtn.Text = "Copy to Clipboard"
$copyBtn.Location = New-Object System.Drawing.Point(190, 620)
$copyBtn.Size = New-Object System.Drawing.Size(150, 40)
$copyBtn.BackColor = [System.Drawing.Color]::LightBlue
$copyBtn.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($copyBtn)

# Clear Button
$clearBtn = New-Object System.Windows.Forms.Button
$clearBtn.Text = "Clear"
$clearBtn.Location = New-Object System.Drawing.Point(360, 620)
$clearBtn.Size = New-Object System.Drawing.Size(150, 40)
$form.Controls.Add($clearBtn)

# Exit Button
$exitBtn = New-Object System.Windows.Forms.Button
$exitBtn.Text = "Exit"
$exitBtn.Location = New-Object System.Drawing.Point(800, 620)
$exitBtn.Size = New-Object System.Drawing.Size(150, 40)
$form.Controls.Add($exitBtn)

# === EVENT HANDLERS ===

$soBrowseBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
    $dialog.InitialDirectory = "c:\Scripts\Document Checker\Example_Pairs"
    
    if ($dialog.ShowDialog() -eq "OK") {
        $soTextBox.Text = $dialog.FileName
    }
})

$poBrowseBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
    $dialog.InitialDirectory = "c:\Scripts\Document Checker\Example_Pairs"
    
    if ($dialog.ShowDialog() -eq "OK") {
        $poTextBox.Text = $dialog.FileName
    }
})

$compareBtn.Add_Click({
    if ([string]::IsNullOrWhiteSpace($soTextBox.Text) -or [string]::IsNullOrWhiteSpace($poTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please select both SO and PO files.", "Missing Files", "OK", [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        return
    }
    
    $resultsText.Clear()
    $comparison = Compare-Documents -SOPath $soTextBox.Text -POPath $poTextBox.Text
    
    $output = "SO: $([System.IO.Path]::GetFileName($soTextBox.Text))`n"
    $output += "PO: $([System.IO.Path]::GetFileName($poTextBox.Text))`n`n"
    $resultsText.AppendText($output)
    
    if ($comparison.Match) {
        $resultsText.SelectionColor = [System.Drawing.Color]::Green
        $resultsText.AppendText("RESULT: COMPLETE MATCH`n`n")
    } else {
        $resultsText.SelectionColor = [System.Drawing.Color]::Red
        $resultsText.AppendText("RESULT: MISMATCH - MANUAL REVIEW REQUIRED`n`n")
        $resultsText.SelectionColor = [System.Drawing.Color]::Black
        foreach ($issue in $comparison.Issues) {
            $resultsText.AppendText("  - $issue`n")
        }
    }
    
    $resultsText.SelectionColor = [System.Drawing.Color]::Black
    $resultsText.AppendText("`n`nShip To Addresses:`n")
    $resultsText.AppendText("  SO: $($comparison.SOShipTo.Name), $($comparison.SOShipTo.Address), $($comparison.SOShipTo.City), $($comparison.SOShipTo.State) $($comparison.SOShipTo.Zip)`n")
    $resultsText.AppendText("  PO: $($comparison.POShipTo.Name), $($comparison.POShipTo.Address), $($comparison.POShipTo.City), $($comparison.POShipTo.State) $($comparison.POShipTo.Zip)`n")
})

$copyBtn.Add_Click({
    if ($resultsText.Text.Length -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to copy.", "No Results", "OK", [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        return
    }
    
    [System.Windows.Forms.Clipboard]::SetText($resultsText.Text)
    [System.Windows.Forms.MessageBox]::Show("Results copied to clipboard!", "Success", "OK", [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
})

$clearBtn.Add_Click({
    $soTextBox.Text = ""
    $poTextBox.Text = ""
    $resultsText.Clear()
})

$exitBtn.Add_Click({
    $form.Close()
})

$form.ShowDialog() | Out-Null
