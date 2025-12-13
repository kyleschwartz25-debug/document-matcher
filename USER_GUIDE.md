# Document Matcher Tool - User Guide

## Overview
The Document Matcher compares Sales Orders (SO) with Purchase Orders (PO) to verify that all details match correctly.

## How to Launch
1. **Double-click** `LaunchMatcher.bat` in the `c:\Scripts\Document Checker` folder
   - Or run: `powershell -ExecutionPolicy Bypass -File DocumentMatcher_v2.ps1`

## How to Use

### Step 1: Open Your Documents
- Open the SO PDF and PO PDF side-by-side (or in separate windows)
- Have them ready to reference

### Step 2: Enter Sales Order (SO) Data
**Left side of the application:**

1. **Ship To Address**
   - Copy the "Ship To:" box information from your SO
   - Fill in: Name, Address, City, State, Zip
   - Example:
     ```
     Name: Tom Mention
     Address: 10181 LAUREL DR
     City: EDEN PRAIRIE
     State: MN
     Zip: 55347-3048
     ```

2. **Line Items** (Dropship items ONLY)
   - For each row in the SO that is marked "Drop Ship":
     - **SKU**: Copy the Number column (e.g., "350027-M")
     - **Description**: Copy from Description column (e.g., "Custom - Storm Training Group Lightweight Shorts Black 350027-M")
     - **Qty**: Copy the Qty Ordered value (e.g., "6")
   - Ignore shipping, tariff, and tax rows

### Step 3: Enter Purchase Order (PO) Data
**Right side of the application:**

1. **Ship To Address**
   - Copy the "Ship To:" box information from your PO
   - Fill in exactly the same fields

2. **Line Items**
   - Enter all line items (same as SO dropship items)
   - Use same format: SKU, Description, Qty

### Step 4: Compare
1. Click the green **COMPARE** button
2. Results will appear at the bottom:
   - **COMPLETE MATCH**: All fields are identical (good!)
   - **MISMATCH**: Fields differ - review the issues listed

### Step 5: Copy Results to CRM
1. Click the blue **COPY TO CLIPBOARD** button
2. Paste the results directly into your CRM system

### Step 6: Continue or Clear
- **Clear All**: Start a new comparison
- **Exit**: Close the application

## Example Workflow

For the sample pair (SO-L029638 + PO-KS2013442):

**SO Data:**
```
Name: Tom Mention
Address: 10181 LAUREL DR
City: EDEN PRAIRIE
State: MN
Zip: 55347-3048

Line Items:
350027-M | Custom - Storm Training Group Lightweight Shorts Black 350027-M | 6
350027-L | Custom - Storm Training Group Lightweight Shorts Black 350027-L | 14
350027-XL | Custom - Storm Training Group Lightweight Shorts Black 350027-XL | 14
350027-2XL | Custom - Storm Training Group Lightweight Shorts Black 350027-2XL | 10
350027-3XL | Custom - Storm Training Group Lightweight Shorts Black 350027-3XL | 6
350027-YM | Custom - Storm Training Group Lightweight Shorts Black 350027-YM | 2
350027-YXL | Custom - Storm Training Group Lightweight Shorts Black 350027-YXL | 2
```

**PO Data:**
(Same as SO)

**Result:** COMPLETE MATCH ✓

## Tips

1. **Copy text carefully** - Use CTRL+C to copy exact text from PDFs
2. **Watch for extra spaces** - Leading/trailing spaces will cause mismatches
3. **SKU numbers** - Must match exactly (case-sensitive)
4. **Quantities** - Must be numbers only (no "ea", "pcs", etc.)
5. **City/State/Zip** - Format must match exactly between SO and PO

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Ship To Name mismatch" | Check for extra spaces, typos, or abbreviations |
| "Line count mismatch" | Make sure you're only including dropship items in SO |
| "SKU mismatch" | Verify exact SKU format - may have dashes or hyphens |
| "Qty mismatch" | Check for "ea" or other units - enter numbers only |
| Results won't copy | Click COMPARE first, then try COPY again |

## For Recurring Use

Every time you place a new SO/PO pair:
1. Run LaunchMatcher.bat
2. Quickly copy data from each PDF
3. Compare and copy results
4. Estimated time per pair: 2-3 minutes

## Data Privacy

- All data is processed locally on your computer
- Nothing is sent to external servers
- Results can be safely copied to your CRM

---

**Version**: 1.0  
**Last Updated**: December 13, 2025
