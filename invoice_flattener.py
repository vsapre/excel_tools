# ===============================================================================
# Invoice Flattening Tool
# Started as a chat with Claude.ai, where it created the main structure of the
# tool based on a single invoice example.
# Vishal Sapre added missing and generic rules, which would apply to multiple invoices
# Vishal Sapre added further text patterns seen across multiple invoices...WIP
#
# LICENCE: BSD
# Date: 15/10/2025
# ===============================================================================
import pandas as pd
import re
import os
from pathlib import Path
from datetime import datetime

part_patterns = ['CAT NO:',
"RIGHILL PART NO. :",
"RIGHILL PART NO.:",
"RIGHILL PART NO.",
"RIGHILL PART NO:",
"RIGHILL PART NO",
"RIGHILL PT/NO.",
"RIGHILL PT/.NO.",
"RIGHILL PT NO.",
"RIGHILL PT NO:",
"RIGHILL PT NO :",
"RIGHILL PART #",
"RIGHILL P/N"]

service_names = (
"commissioning","repair","troubleshooting","trouble shooting","days","timesheet",
"time sheet","installation","servicing","service","re-commissioning","health",
"check up","calibration","tuning","diagnosis","survey","deputation","training",
"testing", "sitc")

# From Google Gemini: "python re remove month names from strings"
# store all month names in a pattern list
month_names = ("January", "February", "March", "April", "May", "June", "July",
"August", "September", "October", "November", "December", "Jan", "Feb", "Mar",
"Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

invoice_number_patterns = ('Invoice Sl.No', 'Invoice Sl.', 'Invoice Serial No', 'Invoice Serial Number')

po_number_patterns = ("PO/WO.No & Dt", 'PONo & Dt', 'PONo', 'PO No', 'P.O. No', 'WO.No & Dt', 'WO No', 'SONo & Dt', 'Contract No & Dt')

# Create a regex pattern to match any of the month names, case-insensitive
# The '|' acts as an OR operator, and re.IGNORECASE makes it case-insensitive
month_pattern = r'\b(?:' + '|'.join(re.escape(name) for name in month_names) + r')\b'

class GSTInvoiceFlattener:
    def __init__(self, output_file='flattened_invoices.xlsx'):
        self.output_file = output_file
        self.processed_invoices = set()
##        self.load_existing_data()
        self.folders = []

    def load_existing_data(self):
        """Load existing flattened data if file exists"""
        if os.path.exists(self.output_file):
            self.df_existing = pd.read_excel(self.output_file)

            # Track already processed invoices from the path based file names
            if 'Invoice File' in self.df_existing.columns:
                self.processed_invoices = set(self.df_existing['Invoice File'].unique())

##            if 'Invoice Number' in self.df_existing.columns:
##                self.processed_invoices = set(self.df_existing['Invoice Number'].unique())

            print(f"Loaded existing file with {len(self.df_existing)} rows")
        else:
            self.df_existing = pd.DataFrame()
            print("No existing file found. Will create new one.")

    def extract_material_code(self, description):
        """Extract material code from description"""
        if pd.isna(description):
            return None

        # Pattern: MATERIAL CODE: followed by alphanumeric
        match = re.search(r'\b(?:Matl[\.]*|MAT|MATERIAL)\s+CODE\b[\.\s]*[:\s]+([A-Z0-9]+)', str(description), re.IGNORECASE)
##        if(('MAT CODE' in str(description)) or ('MATERIAL CODE' in str(description)) or
        if match:
            return match.group(1).strip()
        return None

    def extract_righill_part(self, description):
        """Extract Righill part number from description"""
        if pd.isna(description):
            return None

        descUpper = str(description).upper()

        splitBase = ""
        for pattern in part_patterns:
            if(pattern in descUpper):
                splitBase = pattern
                break
        else:
            return None

        parts = descUpper.split(splitBase)
        match = re.search(r'([A-Z0-9]+[\s–\-\s]*[0-9]+[\s–\-\s]*[A-Z0-9]*)', parts[-1], re.IGNORECASE)
##        # Pattern: RIGHILL PART NO: followed by alphanumeric having either normal dash or unicode dash_punctuation or space
##        match = re.search(r'(?:RIGHILL|Righill)\s+PART NO[\.\s]*[:–\-\s]*([A-Z0-9]+[\s–\-\s]*[0-9]+[\s–\-\s]*[A-Z0-9]*)', str(description), re.IGNORECASE)
##
        if match:
            partNumber = match.group(1).strip()
            # remove unecessary suffixes that have crept in
            partNumber = partNumber.removesuffix("FOR")
            partNumber = partNumber.removesuffix("MATERIAL")
            partNumber = partNumber.removesuffix("MAT")
            partNumber = partNumber.removesuffix("SKU")
            partNumber = partNumber.removesuffix("MATL")
            partNumber = partNumber.removesuffix("BHEL")
            partNumber = partNumber.removesuffix("OIL")
            partNumber = partNumber.removesuffix("ONGC")
            partNumber = partNumber.removesuffix("SCHNEIDER")
            pre = partNumber[:3]   # part number prefix identifying part type e.g. RCP-450 or RCB-3504
            # remove all un-needed characters from prefix
            pre = pre.replace(" ", "").replace("-", "").replace("–", "").replace(".", "")
            post = partNumber[2:]  # suffix identifying that exact part number
            # remove all un-needed characters from suffix
            post = post.replace(" ", "").replace("-", "").replace("–", "").replace(".", "")
            # for a 3 letter type indentifier, the last letter of pre will showup in post also.
            post = post.removeprefix(pre[-1])
            partNumber = "-".join((pre, post))
            return partNumber

        return None

    def parse_invoice_excel(self, file_path):
        """Parse a single Excel invoice file"""
        try:
            # Read the Excel file
            df = pd.read_excel(file_path, header=None)

            # take last 2 items from the path to uniquely identify this file
            invoice_fname = os.path.join(file_path.parts[-2], file_path.parts[-1])

            # if this invoice file name appears already, skip and move to next file
            if invoice_fname in self.processed_invoices:
                print(f"  ⊗ Skipped (already processed): {invoice_fname}")
                return None

            # find the customer name
            fname = file_path.name.removesuffix(file_path.suffix)
            custList = fname.split(" ")[1:-1]
            customer = ""
            if(not custList):   # list with length zero
                customer = fname.split(" ")[1]
            else:
                customer = " ".join(custList)

            # some file names have 'Export' at the end which creeps in here
            customer = customer.replace("Export", "")

            # some file names have month names at the end which creeps in here
            # Use re.sub to replace all occurrences of the pattern with an empty string
            cleaned_text = re.sub(month_pattern, '', customer, flags=re.IGNORECASE)

            customer = cleaned_text.strip()

##            # Optionally, remove extra spaces that might result from removal
##            cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()

            buyer_name = ""
            # Initialize invoice data dictionary
            invoice_data = {
                'invoice_fname' : invoice_fname,
                'invoice_number': None,
                'invoice_date': None,
                'invoice_type':None,
                'invoice_currency': None,
                'po_number': None,
                'po_date': None,
                'customer_name': customer,
                'buyer_name': None,
                'items': [],
                'cgst_rate': None,
                'sgst_rate': None,
                'igst_rate': None,
                'total_cgst': 0,
                'total_sgst': 0,
                'total_igst': 0,
                'total_gst': 0
            }

            # Extract header information by searching through cells
            for idx, row in df.iterrows():
                row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])

                # Invoice Number
                if(not invoice_data['invoice_number']):
                    invoiceNumMatchList = [(pattern in row_str) for pattern in invoice_number_patterns]
                    if(any(invoiceNumMatchList)):
                        pattern = invoice_number_patterns[invoiceNumMatchList.index(True)]
                        inter = row_str.split(pattern)[1]
                        invoice_num = inter.split()[0]
                        invoice_num = invoice_num.strip(":. ")
                        invoice_data['invoice_number'] = invoice_num
                        if('G/' in invoice_num):
                            invoice_data['invoice_type'] = "GOODS"
                        if('S/' in invoice_num):
                            invoice_data['invoice_type'] = "SERVICE"
                        continue

##                if 'Invoice Sl.' in row_str or 'Invoice No' in row_str or 'Invoice Serial No' in row_str or 'Invoice Serial Number' in row_str:
##                    for cell in row:
##                        if pd.notna(cell) and ('REPL/' in str(cell) or 'G/' in str(cell) or 'S/' in str(cell)):
##                            if('G/' in str(cell)):
##                                invoice_data['invoice_type'] = "GOODS"
##                            if('S/' in str(cell)):
##                                invoice_data['invoice_type'] = "SERVICE"
##                            invoice_data['invoice_number'] = str(cell).strip()
##                            break

                # Invoice Date
                if(not invoice_data['invoice_date']):
                    if 'Invoice Date' in row_str:
                        dating = row_str.split("Place of Delivery")[0].split('Invoice Date')[1]
                        dating = dating.strip(": ")
                        if(not pd.isna(pd.to_datetime(dating, errors='coerce'))):
                            invoice_data['invoice_date'] = dating
                            continue

##                    if 'Invoice Date' in row_str:
##                    invoice_data['invoice_date'] = ""
##                    for i, cell in enumerate(row):
##                        cellVal = pd.to_datetime(cell, errors='coerce')
##                        if(not pd.isna(cellVal)):
##                            invoice_data['invoice_date'] = str(cell).strip().split(" ")[0]
##                            break

##                        if pd.notna(cell) and 'Invoice Date' in str(cell):
##                            if i + 1 < len(row) and pd.notna(row.iloc[i + 1]): #
##                                invoice_data['invoice_date'] = str(row.iloc[i + 1]).strip()
##                                break

                # PO Number and Date
                if(not invoice_data['po_number']):
                    poNumMatchList = [(pattern in row_str) for pattern in po_number_patterns]
                    if(any(poNumMatchList)):
                        if("dt." in row_str):
                            datePattern = "dt."
                        elif("Date" in row_str):
                            datePattern = "Date"
                        elif("  & dtd. " in row_str):
                            datePattern = "  & dtd. "
                        else:
                            datePattern = " & "
                        poNumAndDate = row_str.split("Place of Delivery")[0].split(datePattern)
                        poDate = poNumAndDate[1].strip(": ")
                        pattern = po_number_patterns[poNumMatchList.index(True)]
                        poNum = poNumAndDate[0].split(pattern)[1]
                        poNum = poNum.strip(": ")
                        invoice_data['po_number'] = poNum
                        if(not pd.isna(pd.to_datetime(poDate, errors='coerce'))):
                            invoice_data['po_date'] = poDate
                        continue

##                if 'PO No' in row_str or 'P.O. No' in row_str:
##                    #remove all nan values
##                    row = row.dropna()
##                    for i, cell in enumerate(row):
##                        if pd.notna(cell) and ('PO No' in str(cell) or 'P.O. No' in str(cell)):
##                            # Look for PO number in next cells
##                            if (i + 1 < len(row)) and pd.notna(row.iloc[i + 1]):
##                                invoice_data['po_number'] = str(row.iloc[i + 1]).strip()
##                            # Look for date
##                            for j in range(i + 1, min(i + 5, len(row))):
##                                if pd.notna(row.iloc[j]) and 'Date' in str(row.iloc[j]):
##                                    if j + 1 < len(row) and pd.notna(row.iloc[j + 1]):
##                                        invoice_data['po_date'] = str(row.iloc[j + 1]).strip().split(" ")[0]
##                                        break
##                            break

                # Buyer Name
                if(not invoice_data['buyer_name']):
                    if 'Details of Receiver' in row_str or 'Billed to' in row_str:
                        # Buyer name is usually in next row after "Name:"
                        if idx + 1 < len(df):
                            next_row = df.iloc[idx + 1]
                            newRowStr = ' '.join([str(cell) for cell in next_row if pd.notna(cell)])
                            newRowStr = newRowStr.split("Name")[1] # 0th will be empty, 1st will be buyer name
                            newRowStr = newRowStr.strip(": ")
                            # some invoices will have 'Consignee' on the 'shipped to' side.
                            if('Consignee' in newRowStr):
                                newRowStr = newRowStr.split('Consignee')[0]
                            invoice_data['buyer_name'] = buyer_name = newRowStr.strip(": ")
                            continue

##                if 'Details of Receiver' in row_str or 'Billed to' in row_str:
##                    # Buyer name is usually in next row after "Name:"
##                    if idx + 1 < len(df):
##                        next_row = df.iloc[idx + 1]
##                        for i, cell in enumerate(next_row):
##                            if pd.notna(cell) and 'Name' in str(cell):
##                                newRowStr = ' '.join([str(cell) for cell in next_row if pd.notna(cell)])
##                                newRowStr = newRowStr.split("Name")[1] # 0th will be empty, 1st will be buyer name
##                                newRowStr = newRowStr.strip(": ")
##                                # some invoices will have 'Consignee' on the 'shipped to' side.
##                                if('Consignee' in newRowStr):
##                                    newRowStr = newRowStr.split('Consignee')[0]
##                                invoice_data['buyer_name'] = buyer_name = newRowStr.strip(": ")
####                                if i + 1 < len(next_row) and pd.notna(next_row.iloc[i + 1]):
####                                    invoice_data['buyer_name'] = buyer_name = str(next_row.iloc[i + 1]).strip()
####                                    if(customer not in buyer_name.upper()):
####                                        invoice_data['buyer_name'] = " ".join((buyer_name, customer))
##                                break

                # Tax rates
                if 'CGST @' in row_str:
                    match = re.search(r'CGST @\s*(\d+)', row_str)
                    if match: invoice_data['cgst_rate'] = f"{match.group(1)}%"

                if 'SGST @' in row_str:
                    match = re.search(r'SGST @\s*(\d+)', row_str)
                    if match: invoice_data['sgst_rate'] = f"{match.group(1)}%"

                if 'IGST @' in row_str:
                    match = re.search(r'IGST @\s*(\d+)', row_str)
                    if match: invoice_data['igst_rate'] = f"{match.group(1)}%"

                # Tax amounts
                if 'CGST @' in row_str:
                    for cell in row:
                        if pd.notna(cell) and isinstance(cell, (int, float)):
                            invoice_data['total_cgst'] = float(cell)
                            break

                if 'SGST @' in row_str:
                    for cell in row:
                        if pd.notna(cell) and isinstance(cell, (int, float)):
                            invoice_data['total_sgst'] = float(cell)
                            break

                if 'IGST @' in row_str:
                    for cell in row:
                        if pd.notna(cell) and isinstance(cell, (int, float)):
                            invoice_data['total_igst'] = float(cell)
                            break

                # some invoices will have only IGST and so those may not show Total GST separately.
                if 'Total GST' in row_str or 'IGST @' in row_str:
                    for cell in row:
                        if pd.notna(cell) and isinstance(cell, (int, float)):
                            invoice_data['total_gst'] = float(cell)
                            break

            # Extract line items
            # Find the row with column headers (S.no, Description, HSN code, etc.)
            header_row_idx = None
            for idx, row in df.iterrows():
                row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                if 'S.no' in row_str and ('Description' in row_str or 'Discription' in row_str) and ('HSN' in row_str or "SAC" in row_str):
                    if('Rate' in row_str and 'USD' in row_str):
                        invoice_data['invoice_currency'] = 'USD'
                    else:
                        invoice_data['invoice_currency'] = "INR"

                    header_row_idx = idx
                    break

            if header_row_idx is not None:
                # Process items starting from next to next row because the 21st
                # row has been left alone and merged with 20th from column 2 onwards.
                for idx in range(header_row_idx + 2, len(df)):
                    row = df.iloc[idx]
                    # lets remove the NaNs
                    row = row.dropna()
                    row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                    if row.empty: continue
                    qVal = 0
                    # Check if this is an item row (starts with a number)
                    first_cell = row.iloc[0]
                    # if this row has good data then there will be exactly 7 columns in this row.
                    # check if there are 7 individual items in this row, else move on.
                    if pd.notna(first_cell) and str(first_cell).strip().isdigit() and len(row) >= 7:
##                        item = {
##                            's_no': str(first_cell).strip(),
##                            'description': '',
##                            'hsn_code': '',
##                            'unit': '',
##                            'quantity': 0,
##                            'rate': 0,
##                            'total_value': 0
##                        }
                        # sometimes quantity is expressed as "1   Set", thus we need
                        # to gather the numeric part of this string to use as
                        # quantity
                        qVal = row.iloc[4]
                        if(isinstance(qVal, str)):
                            qVal = qVal.split(" ")[0]

                        item = {
                            'invoice_item_no': str(row.iloc[0]).strip(),
                            'description': str(row.iloc[1]).strip(),
                            'hsn_code': str(row.iloc[2]).strip(),
                            'unit': str(row.iloc[3]).strip(),
                            'quantity': float(qVal),
                            'rate': float(row.iloc[5]),
                            'total_value': float(row.iloc[6])
                        }

                        # some older invoices donot specify goods or servies in their name
                        # for those, we try to infer the invoice type from the description here
                        # if the invoice_type has not yet been finalized.
                        if(not invoice_data['invoice_type']):
                            descrLower = item['description'].lower()
                            if(any([service in descrLower for service in service_names])):
                                invoice_data['invoice_type'] = "SERVICE"
                            else:
                                invoice_data['invoice_type'] = "GOODS"

##                        # Extract data from row
##                        for i, cell in enumerate(row):
##                            if pd.notna(cell):
##                                cell_str = str(cell).strip()
##
##                                # Description (usually column 1)
##                                if i == 1:
##                                    item['description'] = cell_str
##
##                                # HSN code (usually column 2)
##                                elif i == 2 and cell_str.isdigit():
##                                    item['hsn_code'] = cell_str
##
##                                # Unit (No, Nos, etc.)
##                                elif 'No' in cell_str and len(cell_str) < 5:
##                                    item['unit'] = cell_str
##
##                                # Numeric values (Qty, Rate, Total)
##                                elif isinstance(cell, (int, float)):
##                                    if item['quantity'] == 0:
##                                        item['quantity'] = float(cell)
##                                    elif item['rate'] == 0:
##                                        item['rate'] = float(cell)
##                                    elif item['total_value'] == 0:
##                                        item['total_value'] = float(cell)

##                        # Look for material code in description rows below
##                        desc_lines = []
##                        for desc_idx in range(idx, min(idx + 10, len(df))):
##                            desc_row = df.iloc[desc_idx]
##                            for cell in desc_row:
##                                if pd.notna(cell):
##                                    desc_lines.append(str(cell))
##
##                        full_description = ' '.join(desc_lines)
##                        item['description'] = full_description

                        if item['total_value'] > 0:  # Valid item
                            invoice_data['items'].append(item)

                    # Stop at subtotal or taxable value
                    row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                    if 'Sub Total' in row_str or 'Taxable Value' in row_str:
                        break

            return invoice_data

        except Exception as e:
            print(f"Error parsing {file_path}: {str(e)}")
            return None

    def flatten_invoice(self, invoice_data):
        """Convert invoice data to flattened rows"""
        rows = []

        if not invoice_data or not invoice_data['items']:
            return rows

        # Calculate total taxable value
        total_taxable = sum(item['total_value'] for item in invoice_data['items'])

        for item in invoice_data['items']:
            # Calculate proportional tax for this item
            if total_taxable > 0:
                proportion = item['total_value'] / total_taxable
                item_cgst = round(invoice_data['total_cgst'] * proportion, 2)
                item_sgst = round(invoice_data['total_sgst'] * proportion, 2)
                item_igst = round(invoice_data['total_igst'] * proportion, 2)
                if(item_igst == 0):
                    item_total_gst = round(item_cgst + item_sgst, 2)
                else:
                    item_total_gst = round(item_igst, 2)

                item_grand_total = round(item['total_value'] + item_total_gst, 2)
            else:
                item_cgst = item_sgst = item_igst = item_total_gst = item_grand_total = 0

            row = {
                'Invoice File' : invoice_data['invoice_fname'],
                'Invoice Number': invoice_data['invoice_number'],
                'Invoice Date': invoice_data['invoice_date'],
                'Invoice Type': invoice_data['invoice_type'],
                'Invoice Currency': invoice_data['invoice_currency'],
                'PO Number': invoice_data['po_number'],
                'PO Date': invoice_data['po_date'],
                'Customer Name': invoice_data['customer_name'],
                'Buyer Name': invoice_data['buyer_name'],
                'Invoice Item No.': item['invoice_item_no'],
                'Description': item['description'],
                'Material Code': self.extract_material_code(item['description']),
                'Righill Part No': self.extract_righill_part(item['description']),
                'HSN Code': item['hsn_code'],
                'Unit': item['unit'],
                'Quantity': item['quantity'],
                'Rate': item['rate'],
                'Total Value': item['total_value'],
                'CGST %': invoice_data['cgst_rate'],
                'CGST Amount': item_cgst,
                'SGST %': invoice_data['sgst_rate'],
                'SGST Amount': item_sgst,
                'IGST %': invoice_data['igst_rate'],
                'IGST Amount': item_igst,
                'Total GST': item_total_gst,
                'Grand Total': item_grand_total
            }
            rows.append(row)

        return rows

    def gather_folders(self, root_path):
        """
        Vishal.Sapre
        function to gather all folders with invoice files.
        root_path: input path which contains multiple sub folders with invoice files.
        """
        dirsToRemove = []
        # collect all path values from the root folder
        for path, subdirs, files in os.walk(root_path):
            # collect all paths
            self.folders.append(path)

        # now gather those paths which have any subdirectories
        for folder in self.folders:
            path = Path(folder)
            # if the path is not a directory, move on.
            if(not path.is_dir()):
                continue

            # if the path is a directory, see if there are any subdirectories
            for item in path.iterdir():
                if(item.is_dir()):
                    # if there is subdirectory, add this path to the removal list.
                    dirsToRemove.append(folder)
##                    self.folders.remove(folder)
                    break

        # remove those paths that had subdirectories
        for directory in dirsToRemove:
            self.folders.remove(directory)

        print("="*60)
        print("Folders with Files: ")
        for folder in self.folders:
            print(folder)
        print("="*60)

    def process_folder(self, folder_path):
        """Process all Excel files in a folder"""

        # load existing data from output file if it exists.
        self.load_existing_data()

        # now consume this given folder
        folder = Path(folder_path)
        excel_files = list(folder.glob('*.xlsx')) + list(folder.glob('*.xls'))

        if not excel_files:
            print(f"No Excel files found in {folder_path}")
            return

        print("="*60)
        print(folder)
        print(f"Found {len(excel_files)} Excel files")
        print("="*60)

        all_rows = []
        skipped = 0
        processed = 0

        for file_path in excel_files:
            print(f"\nProcessing: {file_path.name}")

            invoice_data = self.parse_invoice_excel(file_path)

            # Flag if already processed
            if not invoice_data:
                skipped += 1
                continue

            if invoice_data and invoice_data['invoice_number']:
##                # Check if already processed
##                if invoice_data['invoice_number'] in self.processed_invoices:
##                    print(f"  ⊗ Skipped (already processed): {invoice_data['invoice_number']}")
##                    skipped += 1
##                    continue

                rows = self.flatten_invoice(invoice_data)
                if rows:
                    all_rows.extend(rows)
                    self.processed_invoices.add(invoice_data['invoice_number'])
                    processed += 1
                    print(f"  ✓ Extracted {len(rows)} items from invoice {invoice_data['invoice_number']}")
                else:
                    print(f"  ✗ No items found")
            else:
                print(f"  ✗ Failed to extract invoice data")

        # Append to existing data
        if all_rows:
            df_new = pd.DataFrame(all_rows)

            if not self.df_existing.empty:
                df_combined = pd.concat([self.df_existing, df_new], ignore_index=True)
            else:
                df_combined = df_new

            # Save to Excel
            df_combined.to_excel(self.output_file, index=False)
            print(f"\n{'='*60}")
            print(f"✓ Successfully processed {processed} new invoices")
            print(f"⊗ Skipped {skipped} duplicate invoices")
            print(f"✓ Total rows in output file: {len(df_combined)}")
            print(f"✓ Saved to: {self.output_file}")
            print(f"{'='*60}")
        else:
            print("\nNo new data to add.")


def main():
    """Main function to run the script"""
    print("="*60)
    print("GST Invoice Flattener")
    print("="*60)

    # Get folder path from user
##    folder_path = input("\nEnter the folder path containing Excel invoices: ").strip()

    folder_path = input("\nEnter the root folder path containing folders with Excel invoices: ").strip()

    if not os.path.exists(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist!")
        return

    # Optional: custom output file name
    output_file = input("Enter output file name (press Enter for 'flattened_invoices.xlsx'): ").strip()
    if not output_file:
        output_file = 'flattened_invoices.xlsx'

    # Process invoices
    flattener = GSTInvoiceFlattener(output_file)

    # gather all folders that are leafs
    flattener.gather_folders(folder_path)
##    print("="*60)
##    print("Following Folders Found: ")
##    for folder in flattener.folders:
##        print(folder)
##    print("="*60)

    input("\nPress Enter to go ahead")

    # now go process each folder.
    for folder in flattener.folders:
        flattener.process_folder(folder)
##    flattener.process_folder(folder_path)

    print("\nDone!")


if __name__ == "__main__":
    main()
