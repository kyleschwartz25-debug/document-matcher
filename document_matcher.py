#!/usr/bin/env python3
"""
Document Matcher Tool - Drag and Drop PDF Comparison
Compares Sales Orders (SO) with Purchase Orders (PO)
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Frame, Label
import PyPDF2
import re
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple
import difflib

@dataclass
class ShipToAddress:
    name: str = ""
    address: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""

@dataclass
class LineItem:
    sku: str = ""
    description: str = ""
    qty: int = 0

class PDFExtractor:
    """Extract text and data from PDF files"""
    
    @staticmethod
    def extract_text(pdf_path: str) -> str:
        """Extract all text from PDF"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()
                # DEBUG: Print the raw extracted text for inspection
                print(f"\n--- RAW TEXT FROM: {pdf_path} ---\n{text}\n--- END RAW TEXT ---\n")
                return text
        except Exception as e:
            return f"ERROR: Could not extract text from PDF: {e}"
    
    @staticmethod
    def parse_ship_to(text: str) -> ShipToAddress:
        """Parse Ship To address from text, handling both SO and PO formats, and skipping vendor blocks"""
        address = ShipToAddress()
        lines = text.split('\n')
        vendor_keywords = ['fuji', 'pakistan', 'muhabbat', 'khan', 'industrial', 'estate', 'sialkot', 'daska']
        # 1. Try to find the standard 'Ship To:' block (for SO/PO)
        for i, line in enumerate(lines):
            # Print each line for debugging header detection
            if i > 0:
                print(f"[DEBUG] Line {i-1}: {lines[i-1]}")
            print(f"[DEBUG] Line {i}: {line}")
            if i < len(lines) - 1:
                print(f"[DEBUG] Line {i+1}: {lines[i+1]}")
            if 'Ship To' in line and ':' in line:
                # Try to extract next 3-4 lines as address
                collected = []
                for j in range(i + 1, len(lines)):
                    next_line = lines[j].strip()
                    if not next_line:
                        if collected:
                            break
                        continue
                    if any(h in next_line.upper() for h in ['ITEM', 'TYPE', 'NUMBER', 'DESCRIPTION', 'BILL TO', 'NOTES', 'BUYER', 'PAYMENT', 'FOB', 'SHIPPING', 'CREATED', 'VENDOR']):
                        break
                    if next_line.upper() in ['UNITED STATES', 'USA', 'PAKISTAN']:
                        continue
                    if len(collected) >= 4:
                        break
                    collected.append(next_line)
                # Check for vendor keywords in the collected address
                is_vendor = any(any(kw in (collected[k].lower() if k < len(collected) else '') for kw in vendor_keywords) for k in range(len(collected)))
                if not is_vendor:
                    if len(collected) >= 1:
                        address.name = collected[0]
                    if len(collected) >= 2:
                        address.address = collected[1]
                    if len(collected) >= 3:
                        remaining = " ".join(collected[2:])
                        match = re.search(r'([A-Z][A-Z\s,]*?)\s+([A-Z]{2})\s+(\d{5}(?:-\d{4})?)', remaining)
                        if match:
                            address.city = match.group(1).replace(',', '').strip()
                            address.state = match.group(2)
                            address.zip_code = match.group(3)
                    if address.name and address.address:
                        return address
        # 2. Try PO format: look for 'Fax:' line with customer name
        for i, line in enumerate(lines):
            if 'Fax:' in line and len(line.strip()) > 5:
                # Customer name is after 'Fax:'
                parts = line.split('Fax:')
                if len(parts) > 1 and parts[1].strip():
                    address.name = parts[1].strip()
                # Next lines: address, city/state/zip
                if i+1 < len(lines):
                    address.address = lines[i+1].strip()
                if i+2 < len(lines):
                    city_line = lines[i+2].strip()
                    match = re.search(r'([A-Z][A-Z\s,]*?)\s*,?\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)', city_line)
                    if match:
                        address.city = match.group(1).replace(',', '').strip()
                        address.state = match.group(2)
                        address.zip_code = match.group(3)
                return address
        return address
    
    @staticmethod
    def parse_line_items(text: str, is_invoice: bool = False) -> List[LineItem]:
        """Parse line items from text for SO and PO tables, extracting only Number, Description, and Qty Ordered. Handles PO two-line items."""
        items = []
        lines = text.split('\n')
        in_items = False
        buffer = []
        i = 0
        while i < len(lines):
            # Flexible header detection for SO and PO
            if i+2 < len(lines):
                if (lines[i].strip().lower() == 'item'):
                    second = lines[i+1].replace(' ', '').lower()
                    third = lines[i+2].replace(' ', '').lower()
                    # SO: '#typenumberdescriptionunitpriceqty' and 'orderedtotalprice'
                    if ('number' in second and 'qty' in second and 'ordered' in third):
                        in_items = True
                        i += 3
                        continue
                    # PO: '#numberdescriptionqty' and 'ordereduom'
                    if ('number' in second and 'qty' in second and 'uom' in third):
                        in_items = True
                        i += 3
                        continue
            # Detect end of items section
            if in_items and re.search(r"Subtotal|Shipping|Tariff|Notes|approval|Total", lines[i], re.IGNORECASE):
                break
            if in_items and lines[i].strip():
                buffer.append(lines[i].strip())
            i += 1
        # Debug: print buffer contents
        print("[DEBUG] Buffer after header detection:", buffer)
        # Improved SO buffered lines (join only valid pairs)
        if is_invoice and buffer:
            i = 0
            while i < len(buffer) - 1:
                first = buffer[i]
                second = buffer[i+1]
                # First line: starts with number, Drop Ship, SKU; second line ends with 'ea'
                if re.match(r"^\d+\s+Drop Ship\s+[\w\-]+", first) and re.search(r"\d+ea", second):
                    # Example first: '1 Drop Ship 350027-M Custom - Storm Training Group Lightweight Shorts Black 350027-M$30.98'
                    # Example second: '6ea $ 185.88'
                    match1 = re.match(r"^\d+\s+Drop Ship\s+([\w\-]+)\s+(.*)", first)
                    match2 = re.match(r"(.*?)(\d+)ea", second)
                    if match1 and match2:
                        sku = match1.group(1).strip()
                        # Remove trailing 'SKU$price' from desc_part
                        desc_part = match1.group(2)
                        # Remove the exact pattern 'sku+$price' from the end
                        desc_part = re.sub(rf"{re.escape(sku)}\$\d+(\.\d{{2}})?$", "", desc_part)
                        # Remove any $price left
                        desc_part = re.sub(r"\$\d+(\.\d{2})?", "", desc_part)
                        # Remove trailing qty/ea if present
                        desc_part = re.sub(r"\s*\d+ea.*", "", desc_part)
                        desc = (desc_part + ' ' + match2.group(1)).replace('\n', ' ').strip()
                        # Remove any dollar sign and everything after it
                        desc = re.sub(r'\$.*', '', desc).strip()
                        desc = re.sub(r'\s+', ' ', desc)
                        qty = int(match2.group(2))
                        items.append(LineItem(sku=sku, description=desc, qty=qty))
                        i += 2
                    else:
                        i += 1
                else:
                    i += 1
        # Improved PO buffered lines (join only valid pairs)
        if not is_invoice and buffer:
            i = 0
            while i < len(buffer) - 1:
                first = buffer[i]
                second = buffer[i+1]
                # First line: starts with number and SKU, second line ends with 'ea'
                if re.match(r"^\d+\s+[\w\-]+", first) and re.search(r"\d+\s+ea$", second):
                    match1 = re.match(r"^\d+\s+([\w\-]+)\s+(.*)", first)
                    match2 = re.match(r"(.*?)(\d+)\s+ea$", second)
                    if match1 and match2:
                        sku = match1.group(1).strip()
                        desc = (match1.group(2) + ' ' + match2.group(1)).replace('\n', ' ').strip()
                        desc = re.sub(r'\s+', ' ', desc)
                        qty = int(match2.group(2))
                        items.append(LineItem(sku=sku, description=desc, qty=qty))
                        i += 2
                    else:
                        i += 1
                else:
                    i += 1
        print("[DEBUG] Extracted items:", items)
        return items

class DocumentMatcher:
    """Compare SO and PO documents"""
    
    def __init__(self):
        self.so_path = None
        self.po_path = None
        self.so_address = None
        self.po_address = None
        self.so_items = []
        self.po_items = []
    
    def load_so(self, path: str):
        """Load and parse Sales Order"""
        text = PDFExtractor.extract_text(path)
        if text.startswith("ERROR"):
            raise Exception(text)
        
        self.so_path = path
        self.so_address = PDFExtractor.parse_ship_to(text)
        self.so_items = PDFExtractor.parse_line_items(text, is_invoice=True)
    
    def load_po(self, path: str):
        """Load and parse Purchase Order"""
        text = PDFExtractor.extract_text(path)
        if text.startswith("ERROR"):
            raise Exception(text)
        
        self.po_path = path
        self.po_address = PDFExtractor.parse_ship_to(text)
        self.po_items = PDFExtractor.parse_line_items(text, is_invoice=False)
    
    def compare(self) -> Tuple[bool, List[str], dict, list]:
        """Compare SO and PO, return (match: bool, issues: List[str], field_status: dict, lineitem_status: list)"""
        if not self.so_address or not self.po_address:
            return False, ["Missing address data"], {}

        issues = []
        field_status = {}
        lineitem_status = []  # List of dicts for each line item comparison

        def fuzzy_status(a, b):
            if a == b:
                return 'green'
            if a and b:
                ratio = difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()
                if ratio > 0.85:
                    return 'yellow'
            return 'red'

        # Compare addresses
        fields = [
            ('name', 'Ship To Name', self.so_address.name, self.po_address.name),
            ('address', 'Ship To Address', self.so_address.address, self.po_address.address),
            ('city', 'Ship To City', self.so_address.city, self.po_address.city),
            ('state', 'Ship To State', self.so_address.state, self.po_address.state),
            ('zip_code', 'Ship To Zip', self.so_address.zip_code, self.po_address.zip_code),
        ]
        for key, label, so_val, po_val in fields:
            status = fuzzy_status(so_val, po_val)
            field_status[key] = status
            if status == 'red':
                issues.append(f"{label}: SO='{so_val}' vs PO='{po_val}'")
            elif status == 'yellow':
                issues.append(f"{label} (close): SO='{so_val}' vs PO='{po_val}'")

        # Build dicts for fast SKU lookup
        so_dict = {item.sku: item for item in self.so_items if item.sku}
        po_dict = {item.sku: item for item in self.po_items if item.sku}
        all_skus = sorted(set(so_dict.keys()) | set(po_dict.keys()))
        for sku in all_skus:
            so_item = so_dict.get(sku, LineItem(sku=sku))
            po_item = po_dict.get(sku, LineItem(sku=sku))
            sku_status = fuzzy_status(so_item.sku, po_item.sku)
            desc_status = fuzzy_status(so_item.description, po_item.description)
            qty_status = 'green' if so_item.qty == po_item.qty and so_item.qty != 0 else 'red'
            lineitem_status.append({
                'so': so_item,
                'po': po_item,
                'sku_status': sku_status,
                'desc_status': desc_status,
                'qty_status': qty_status
            })
            if sku_status == 'red':
                issues.append(f"SKU: SO='{so_item.sku}' vs PO='{po_item.sku}'")
            if desc_status == 'red':
                issues.append(f"Desc: SO='{so_item.description}' vs PO='{po_item.description}'")
            if qty_status == 'red':
                issues.append(f"Qty: SO={so_item.qty} vs PO={po_item.qty}")
        return len(issues) == 0, issues, field_status, lineitem_status

class DocumentMatcherGUI:
    """GUI for Document Matcher"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Document Matcher - SO vs PO Comparison")
        self.root.geometry("1000x700")
        self.root.configure(bg='white')
        
        self.matcher = DocumentMatcher()
        
        # Title
        title = tk.Label(root, text="Sales Order & Purchase Order Matcher", 
                        font=("Arial", 14, "bold"), bg='white')
        title.pack(pady=10)
        
        # SO Frame
        so_frame = tk.LabelFrame(root, text="SALES ORDER (SO)", 
                                font=("Arial", 11, "bold"), bg='white', padx=10, pady=10)
        so_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.so_label = tk.Label(so_frame, text="No file selected", fg="gray")
        self.so_label.pack(side=tk.LEFT)
        
        so_btn = tk.Button(so_frame, text="SELECT SO PDF", 
                          command=self.select_so, bg='lightblue', padx=20)
        so_btn.pack(side=tk.RIGHT)
        
        # PO Frame
        po_frame = tk.LabelFrame(root, text="PURCHASE ORDER (PO)", 
                                font=("Arial", 11, "bold"), bg='white', padx=10, pady=10)
        po_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.po_label = tk.Label(po_frame, text="No file selected", fg="gray")
        self.po_label.pack(side=tk.LEFT)
        
        po_btn = tk.Button(po_frame, text="SELECT PO PDF", 
                          command=self.select_po, bg='lightgreen', padx=20)
        po_btn.pack(side=tk.RIGHT)
        
        # Summary Grid
        summary_label = tk.Label(root, text="Field Match Summary:", font=("Arial", 10, "bold"), bg='white')
        summary_label.pack(anchor=tk.W, padx=10, pady=(10, 0))
        self.summary_frame = Frame(root, bg='white')
        self.summary_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Address fields to show
        self.summary_fields = [
            ('name', 'Name'),
            ('address', 'Address'),
            ('city', 'City'),
            ('state', 'State'),
            ('zip_code', 'Zip'),
        ]
        self.summary_labels = {}
        for idx, (key, label) in enumerate(self.summary_fields):
            lbl = Label(self.summary_frame, text=label, width=12, relief='groove', bg='lightgray', font=("Arial", 10, "bold"))
            lbl.grid(row=0, column=idx, padx=2, pady=2)
            val = Label(self.summary_frame, text="", width=20, relief='ridge', bg='white', font=("Arial", 10))
            val.grid(row=1, column=idx, padx=2, pady=2)
            self.summary_labels[key] = val

        # Line Item Grid
        lineitem_label = tk.Label(root, text="Line Item Comparison:", font=("Arial", 10, "bold"), bg='white')
        lineitem_label.pack(anchor=tk.W, padx=10, pady=(0, 0))
        self.lineitem_frame = Frame(root, bg='white')
        self.lineitem_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        # Results
        results_label = tk.Label(root, text="Comparison Results:", 
                    font=("Arial", 10, "bold"), bg='white')
        results_label.pack(anchor=tk.W, padx=10, pady=(0, 0))
        self.results_text = scrolledtext.ScrolledText(root, height=20, width=120, 
                                 font=("Courier", 9))
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Buttons
        button_frame = tk.Frame(root, bg='white')
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.compare_btn = tk.Button(button_frame, text="COMPARE", 
                                    command=self.compare, bg='#90EE90', 
                                    font=("Arial", 10, "bold"), padx=20, pady=10)
        self.compare_btn.pack(side=tk.LEFT, padx=5)
        
        copy_btn = tk.Button(button_frame, text="COPY TO CLIPBOARD", 
                            command=self.copy_to_clipboard, bg='#87CEEB', 
                            font=("Arial", 10, "bold"), padx=20, pady=10)
        copy_btn.pack(side=tk.LEFT, padx=5)
        
        clear_btn = tk.Button(button_frame, text="CLEAR", 
                             command=self.clear, font=("Arial", 10, "bold"), 
                             padx=20, pady=10)
        clear_btn.pack(side=tk.LEFT, padx=5)
        
        exit_btn = tk.Button(button_frame, text="EXIT", 
                            command=root.quit, font=("Arial", 10, "bold"), 
                            padx=20, pady=10)
        exit_btn.pack(side=tk.RIGHT, padx=5)
    
    def select_so(self):
        """Select SO PDF file"""
        path = filedialog.askopenfilename(
            title="Select Sales Order PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=str(Path(__file__).parent / "Example_Pairs")
        )
        
        if path:
            try:
                self.matcher.load_so(path)
                print("[DEBUG] SO items loaded:", len(self.matcher.so_items))
                for item in self.matcher.so_items:
                    print("[DEBUG] SO item:", item)
                self.so_label.config(text=f"✓ {Path(path).name}", fg="green")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load SO: {e}")
                self.so_label.config(text="Error loading file", fg="red")
    
    def select_po(self):
        """Select PO PDF file"""
        path = filedialog.askopenfilename(
            title="Select Purchase Order PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=str(Path(__file__).parent / "Example_Pairs")
        )
        
        if path:
            try:
                self.matcher.load_po(path)
                print("[DEBUG] PO items loaded:", len(self.matcher.po_items))
                for item in self.matcher.po_items:
                    print("[DEBUG] PO item:", item)
                self.po_label.config(text=f"✓ {Path(path).name}", fg="green")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PO: {e}")
                self.po_label.config(text="Error loading file", fg="red")
    
    def compare(self):
        """Run comparison"""
        if not self.matcher.so_path or not self.matcher.po_path:
            messagebox.showwarning("Missing Files", "Please select both SO and PO files")
            return
        
        try:
            match, issues, field_status, lineitem_status = self.matcher.compare()
            print("[DEBUG] lineitem_status count:", len(lineitem_status))
            for item in lineitem_status:
                print("[DEBUG] lineitem_status entry:", item)
            # Update line item grid
            for widget in self.lineitem_frame.winfo_children():
                widget.destroy()
            # Header (no line number column)
            headers = ["SO SKU", "PO SKU", "SO Desc", "PO Desc", "SO Qty", "PO Qty"]
            for idx, h in enumerate(headers):
                # Make description columns wider
                if h in ("SO Desc", "PO Desc"):
                    col_width = 40
                elif h in ("SO Qty", "PO Qty"):
                    col_width = 8
                else:
                    col_width = 16
                lbl = Label(self.lineitem_frame, text=h, width=col_width, relief='groove', bg='lightgray', font=("Arial", 9, "bold"))
                lbl.grid(row=0, column=idx, padx=1, pady=1)
            # Rows
            for i, item in enumerate(lineitem_status):
                so = item['so']
                po = item['po']
                sku_color = {'green': '#90EE90', 'yellow': '#FFFF99', 'red': '#FF7F7F'}[item['sku_status']]
                desc_color = {'green': '#90EE90', 'yellow': '#FFFF99', 'red': '#FF7F7F'}[item['desc_status']]
                qty_color = {'green': '#90EE90', 'red': '#FF7F7F'}[item['qty_status']]
                Label(self.lineitem_frame, text=so.sku, width=16, relief='ridge', bg=sku_color).grid(row=i+1, column=0, padx=1, pady=1)
                Label(self.lineitem_frame, text=po.sku, width=16, relief='ridge', bg=sku_color).grid(row=i+1, column=1, padx=1, pady=1)
                Label(self.lineitem_frame, text=so.description, width=60, wraplength=600, relief='ridge', bg=desc_color, anchor='w', justify='left').grid(row=i+1, column=2, padx=1, pady=1)
                Label(self.lineitem_frame, text=po.description, width=60, wraplength=600, relief='ridge', bg=desc_color, anchor='w', justify='left').grid(row=i+1, column=3, padx=1, pady=1)
                Label(self.lineitem_frame, text=str(so.qty), width=8, relief='ridge', bg=qty_color).grid(row=i+1, column=4, padx=1, pady=1)
                Label(self.lineitem_frame, text=str(po.qty), width=8, relief='ridge', bg=qty_color).grid(row=i+1, column=5, padx=1, pady=1)
            # Update summary grid
            so_addr = self.matcher.so_address
            po_addr = self.matcher.po_address
            for key, label in self.summary_fields:
                so_val = getattr(so_addr, key, "")
                po_val = getattr(po_addr, key, "")
                status = field_status.get(key, 'white')
                display = f"SO: {so_val}\nPO: {po_val}"
                color = {'green': '#90EE90', 'yellow': '#FFFF99', 'red': '#FF7F7F', 'white': 'white'}[status]
                self.summary_labels[key].config(text=display, bg=color)
            self.results_text.config(state=tk.NORMAL)
            self.results_text.delete(1.0, tk.END)
            so_file = Path(self.matcher.so_path).name
            po_file = Path(self.matcher.po_path).name
            self.results_text.insert(tk.END, f"SO: {so_file}\n")
            self.results_text.insert(tk.END, f"PO: {po_file}\n")
            self.results_text.insert(tk.END, "\n" + "="*100 + "\n\n")
            if match:
                self.results_text.insert(tk.END, "STATUS: COMPLETE MATCH\n\n", "success")
            else:
                self.results_text.insert(tk.END, "STATUS: MISMATCH - MANUAL REVIEW REQUIRED\n\n", "error")
                self.results_text.insert(tk.END, "Issues Found:\n" + "-"*100 + "\n")
                for issue in issues:
                    self.results_text.insert(tk.END, f"  - {issue}\n")
            self.results_text.config(state=tk.DISABLED)
            # Configure tags for colors
            self.results_text.tag_config("success", foreground="green")
            self.results_text.tag_config("error", foreground="red")
        except Exception as e:
            messagebox.showerror("Error", f"Comparison failed: {e}")
    
    def copy_to_clipboard(self):
        """Copy results to clipboard"""
        text = self.results_text.get(1.0, tk.END)
        if text.strip():
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("Success", "Results copied to clipboard!")
        else:
            messagebox.showwarning("No Results", "Run a comparison first")
    
    def clear(self):
        """Clear all data"""
        self.matcher = DocumentMatcher()
        self.so_label.config(text="No file selected", fg="gray")
        self.po_label.config(text="No file selected", fg="gray")
        self.results_text.config(state=tk.NORMAL)
        self.results_text.delete(1.0, tk.END)
        self.results_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentMatcherGUI(root)
    root.mainloop()
