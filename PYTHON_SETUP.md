# Document Matcher - Python Setup Guide

## IMPORTANT: First Time Setup Required

This tool uses Python with PDF extraction capabilities. You need to install Python first.

---

## Step 1: Install Python 3.11+

### Option A: Automated Setup (Recommended)
1. **Double-click** `SETUP.bat` in the `c:\Scripts\Document Checker` folder
2. The script will check for Python and install required packages
3. Follow any prompts

### Option B: Manual Setup
1. Download Python 3.11+ from https://www.python.org/downloads/
2. **IMPORTANT:** During installation:
   - ✓ Check **"Add Python to PATH"** (bottom of installer)
   - ✓ Check **"pip"** in custom installation
   - Click **Install Now**
3. Close the installer when done

### Verify Installation
Open Command Prompt and run:
```
python --version
```

You should see: `Python 3.11.x` or higher

---

## Step 2: Install Dependencies

Once Python is installed, run:

```
SETUP.bat
```

Or manually from Command Prompt:
```
python -m pip install PyPDF2
```

---

## Step 3: Launch the Tool

### Option A: Double-click
```
document_matcher.py
```

### Option B: Command Prompt
```
cd c:\Scripts\Document Checker
python document_matcher.py
```

---

## How It Works

1. **SELECT SO PDF** - Click to choose your Sales Order PDF
2. **SELECT PO PDF** - Click to choose your Purchase Order PDF
3. **COMPARE** - Automatically extracts data and compares
4. **View Results** - Shows MATCH or MISMATCH with details
5. **COPY TO CLIPBOARD** - Paste results into your CRM

---

## Features

✓ Automatic PDF text extraction  
✓ Smart parsing of Ship To addresses  
✓ Line-by-line item comparison  
✓ Detailed mismatch reporting  
✓ Copy results to clipboard  
✓ No manual data entry  

---

## Troubleshooting

### "Python not found"
- Install Python from https://www.python.org/downloads/
- Make sure to check "Add Python to PATH" during installation
- Restart your computer after installing Python

### "PyPDF2 not found"
- Run `SETUP.bat`
- Or manually: `python -m pip install PyPDF2`

### "PDF extraction failed"
- The PDF may be scanned/image-based rather than text-based
- Text-based PDFs (like your examples) work best
- If it fails, try opening the PDF in Adobe Reader and re-saving it

### Application won't open
- Make sure Python is installed (`python --version`)
- Make sure PyPDF2 is installed (`python -m pip list | find PyPDF2`)
- Try running from Command Prompt to see error messages

---

## System Requirements

- Windows 7 or later
- Python 3.11+ (will be installed)
- ~200 MB disk space for Python
- Internet connection (one-time for downloads)

---

## First Run Checklist

- [ ] Python installed (run `python --version`)
- [ ] Dependencies installed (run `SETUP.bat`)
- [ ] Can double-click `document_matcher.py`
- [ ] GUI window opens
- [ ] Can select SO PDF
- [ ] Can select PO PDF
- [ ] COMPARE button shows results

---

## Getting Help

If something doesn't work:

1. Open Command Prompt in `c:\Scripts\Document Checker`
2. Run: `python document_matcher.py`
3. Copy any error messages
4. Share the error message for troubleshooting

---

## Next Steps

Once setup is complete:

1. **Test with sample pair:**
   - SO: `SO-L029638-Higashi-Kerwyn Tokeshi-Rev B-Receipt.pdf`
   - PO: `PO-KS2013442AIR-Storm Training Group-Chaz Lemon-Rev A.pdf`

2. **Expected result:** COMPLETE MATCH

3. **If it works:** You're all set! Use it on all future SO/PO pairs

4. **If it doesn't work:** Let me know the error message

---

## One-Time Setup Time

- Install Python: 2-5 minutes
- Run SETUP.bat: 1-2 minutes
- Test the tool: 2-3 minutes

**Total: ~10 minutes, one time**

Then each SO/PO comparison takes: **30 seconds**

---

**Version**: 1.0 Python Edition  
**Last Updated**: December 13, 2025
