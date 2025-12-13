#!/usr/bin/env python3
"""
Debug script to see what text is being extracted from PDFs
Run this to diagnose extraction issues
"""

import PyPDF2
from pathlib import Path

def extract_and_display_text(pdf_path):
    """Extract and display raw text from PDF"""
    print(f"\n{'='*80}")
    print(f"PDF: {Path(pdf_path).name}")
    print(f"{'='*80}\n")
    
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            print(f"Total Pages: {len(pdf_reader.pages)}\n")
            
            # Extract from first page only (most relevant)
            page = pdf_reader.pages[0]
            text = page.extract_text()
            
            print("RAW EXTRACTED TEXT:")
            print("-" * 80)
            print(text)
            print("-" * 80)
            
            # Look for Ship To section
            if "Ship To" in text:
                idx = text.find("Ship To")
                print(f"\nFound 'Ship To' at position {idx}")
                print("Context (500 chars after 'Ship To'):")
                print(text[idx:idx+500])
            else:
                print("\nWARNING: 'Ship To' not found in text!")
                
    except Exception as e:
        print(f"ERROR: {e}")

if __name__ == "__main__":
    # Get the SO file
    so_file = Path("c:/Scripts/Document Checker/Example_Pairs").glob("SO-L030812*")
    so_path = list(so_file)
    
    if so_path:
        extract_and_display_text(str(so_path[0]))
    else:
        print("SO file not found in Example_Pairs folder")
        print("Please specify the path manually:")
        path = input("Enter SO PDF path: ")
        extract_and_display_text(path)
