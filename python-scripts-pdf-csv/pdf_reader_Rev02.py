import os
import re
import pytesseract
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
from tkinterdnd2 import DND_FILES, TkinterDnD
import pdfplumber
import csv
from dotenv import load_dotenv
from openai import OpenAI
from datetime import datetime
import calendar

# Load environment variables from a .env file
load_dotenv()
client = OpenAI()

client.api_key = os.getenv("OPENAI_API_KEY")

# Define the path to Poppler binaries
# Path to Poppler on macOS
POPPLER_PATH = "/opt/homebrew/bin"

class InvoiceProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Processor with Drag-and-Drop")
        self.root.geometry("500x400")

        # List to store file paths
        self.files = []

        # Set up Drag-and-Drop
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

        # File selection button
        self.select_button = tk.Button(root, text="Select PDF Files", command=self.select_files)
        self.select_button.pack(pady=10)

        # Listbox to show selected files
        self.file_listbox = Listbox(root, width=50, height=10)
        self.file_listbox.pack(pady=10)

        # Process button
        self.process_button = tk.Button(root, text="Process Files", command=self.process_files)
        self.process_button.pack(pady=10)

    def select_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file_path in file_paths:
            self.add_file(file_path)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            if file_path.endswith('.pdf'):
                self.add_file(file_path)
            else:
                messagebox.showwarning("Invalid File", f"{file_path} is not a PDF file and was ignored.")

    def add_file(self, file_path):
        if file_path not in self.files:
            self.files.append(file_path)
            self.file_listbox.insert(tk.END, os.path.basename(file_path))

    def process_files(self):
        if not self.files:
            messagebox.showwarning("No Files Selected", "Please select or drop PDF files to process.")
            return

        csv_data = []

        for file in self.files:
            extracted_data = self.extract_data_from_pdf(file)
            if extracted_data:
                csv_data.append(extracted_data)

        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if save_path:
            self.save_to_csv(csv_data, save_path)
            messagebox.showinfo("Process Complete", f"Data extracted and saved to {save_path}.")

    def extract_data_from_pdf(self, file_path):
        data = {
            "*ContactName": "",
            "EmailAddress": "",
            "POAddressLine1": "", "POAddressLine2": "", "POAddressLine3": "", "POAddressLine4": "",
            "POCity": "", "PORegion": "", "POPostalCode": "", "POCountry": "",
            "*InvoiceNumber": "",
            "*InvoiceDate": "",
            "*DueDate": "",
            "Total": "",
            "InventoryItemCode": "",
            "Description": "",
            "*Quantity": "1",
            "*UnitAmount": "",
            "*AccountCode": "540",
            "*TaxType": "20% (VAT on Expenses)",
            "TaxAmount": "",
            "TrackingName1": "Website",
            "TrackingOption1": "",
            "TrackingName2": "",
            "TrackingOption2": "",
            "Currency": "GBP"
        }

        with pdfplumber.open(file_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""

        # Extract text from images using OCR
        images = convert_from_path(file_path, dpi=300)
               # Extract text from images using OCR
        image_text = ""
        try:
            images = convert_from_path(file_path, dpi=300, poppler_path=POPPLER_PATH)
            for image in images:
                image_text += pytesseract.image_to_string(image)
        except Exception as e:  # Catch all exceptions for missing Poppler or other issues
            print(f"Error converting PDF to images or using OCR: {e}")
            messagebox.showerror("PDF Conversion Error", f"Error processing the file: {e}")
            return None

        combined_text = text + "\n" + image_text

        # Debug: Log combined text for troubleshooting
        print("Extracted Text from PDF and Images:\n", combined_text)

        prompt = f"""
        You are an assistant designed to extract structured invoice details. Below is the invoice text:

        {combined_text}

        Your task:
        1. Extract all details from the invoice explicitly, ensuring the "ContactName" is not confused with the "Ship To" or "Billing Address." If a logo or branding is present, prioritize that as the company name.
        2. Ensure all lines of the shipping address and invoice details are extracted accurately.
        3. Extract any email addresses found and assign them to "EmailAddress."
        4. For ambiguous fields (e.g., multiple names), use context (e.g., addresses, invoice number location) to infer the most likely value.

        Output format:
        *ContactName: [Company Name]
        EmailAddress: [Email Address]
        POAddressLine1: [Address Line 1]
        POAddressLine2: [Address Line 2]
        POAddressLine3: [Address Line 3]
        POAddressLine4: [Address Line 4]
        POCity: [City]
        PORegion: [Region]
        POPostalCode: [Postal Code]
        POCountry: [Country]
        *InvoiceNumber: [Invoice Number]
        *InvoiceDate: [Invoice Date]
        *DueDate: [Due Date]
        Total: [Invoice Total]
        InventoryItemCode: [Code]
        Description: [Description]
        *Quantity: [Quantity]
        *UnitAmount: [Unit Amount]
        *AccountCode: [Account Code]
        *TaxType: [Tax Type]
        TaxAmount: [Tax Amount]
        TrackingName1: [Tracking Name 1]
        TrackingOption1: [Tracking Option 1]
        TrackingName2: [Tracking Name 2]
        TrackingOption2: [Tracking Option 2]
        Currency: [Currency]
        """

        try:
            response = client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant for processing invoices."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1000
            )
            
            ai_response = response.choices[0].message.content.strip()
            ai_extracted_data = ai_response.split("\n")

            print("\nParsed AI Extracted Data:\n", ai_extracted_data)
            
            for line in ai_extracted_data:
                line = line.strip()
                if "*ContactName" in line:
                    extracted_name = line.split(":", 1)[1].strip()
                    # Avoid incorrect names like 'Catercall Ltd'
                    if "Catercall" not in extracted_name and "Ship To" not in extracted_name:
                        data["*ContactName"] = extracted_name
                    else:
                        data["*ContactName"] = "Unknown (Check Invoice)"
                elif "EmailAddress" in line:
                    data["EmailAddress"] = line.split(":", 1)[1].strip()
                elif "POAddressLine1" in line:
                    data["POAddressLine1"] = line.split(":", 1)[1].strip()
                elif "POAddressLine2" in line:
                    data["POAddressLine2"] = line.split(":", 1)[1].strip()
                elif "POAddressLine3" in line:
                    data["POAddressLine3"] = line.split(":", 1)[1].strip()
                elif "POAddressLine4" in line:
                    data["POAddressLine4"] = line.split(":", 1)[1].strip()
                elif "POCity" in line:
                    data["POCity"] = line.split(":", 1)[1].strip()
                elif "PORegion" in line:
                    data["PORegion"] = line.split(":", 1)[1].strip()
                elif "POPostalCode" in line:
                    data["POPostalCode"] = line.split(":", 1)[1].strip()
                elif "POCountry" in line:
                    data["POCountry"] = line.split(":", 1)[1].strip()
                elif "*InvoiceNumber" in line:
                    data["*InvoiceNumber"] = line.split(":", 1)[1].strip()
                elif "*InvoiceDate" in line:
                    date_str = line.split(":", 1)[1].strip()
                    data["*InvoiceDate"] = self.format_date(date_str)
                elif "*DueDate" in line:
                    data["*DueDate"] = line.split(":", 1)[1].strip()
                elif "Total" in line:
                    raw_total = line.split(":", 1)[1].strip()
                    clean_total = re.sub(r"[^\d.]", "", raw_total)
                    data["Total"] = f"{float(clean_total):.2f}"

                    # Calculate *UnitAmount and TaxAmount assuming VAT rate is 20%
                    data["*UnitAmount"] = f"{float(clean_total) / 1.2:.2f}"
                    data["TaxAmount"] = f"{float(clean_total) - float(data['*UnitAmount']):.2f}"
                elif "InventoryItemCode" in line:
                    data["InventoryItemCode"] = line.split(":", 1)[1].strip()
                elif "Description" in line:
                    data["Description"] = line.split(":", 1)[1].strip()
                elif "*Quantity" in line:
                    data["*Quantity"] = line.split(":", 1)[1].strip()
                elif "*AccountCode" in line:
                    data["*AccountCode"] = line.split(":", 1)[1].strip()
                elif "*TaxType" in line:
                    data["*TaxType"] = line.split(":", 1)[1].strip()
                elif "TaxAmount" in line:
                    tax_raw = line.split(":", 1)[1].strip()
                    data["TaxAmount"] = re.sub(r"[^\d.]", "", tax_raw)
                elif "TrackingName1" in line:
                    data["TrackingName1"] = line.split(":", 1)[1].strip()
                elif "TrackingOption1" in line:
                    data["TrackingOption1"] = line.split(":", 1)[1].strip()
                elif "TrackingName2" in line:
                    data["TrackingName2"] = line.split(":", 1)[1].strip()
                elif "TrackingOption2" in line:
                    data["TrackingOption2"] = line.split(":", 1)[1].strip()
                elif "Currency" in line:
                    data["Currency"] = line.split(":", 1)[1].strip()

        except Exception as e:
            print("Error with OpenAI API:", e)
            messagebox.showerror("AI Error", f"Could not process the invoice data using AI. Error: {e}")
            return None

         # Fallbacks for missing data
        if not data["*ContactName"]:
            supplier_match = re.search(r"(Supplier|From):\s*([\w\s]+)", text, re.IGNORECASE)
            if supplier_match:
                data["*ContactName"] = supplier_match.group(2).strip()
        if not data["EmailAddress"]:
            email_match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", combined_text)
            if email_match:
                data["EmailAddress"] = email_match.group(0).strip()
        if not data["*InvoiceNumber"]:
            data["*InvoiceNumber"] = "Unknown Invoice Number"
        if not data["*InvoiceDate"]:
            data["*InvoiceDate"] = "Invalid Date"
            data["*DueDate"] = "Invalid Due Date"
        if not data["Total"]:
            data["Total"] = "0.00"
            data["*UnitAmount"] = "0.00"
            data["TaxAmount"] = "0.00"

        return data

    def format_date(self, date_str):
        date_formats = ["%d-%b-%y", "%d/%m/%Y", "%Y-%m-%d"]
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt).strftime("%d/%m/%Y")
            except ValueError:
                continue
        print(f"Warning: Unable to format date '{date_str}'")
        return "Invalid Date"

    def calculate_due_date(self, invoice_date_str):
        try:
            invoice_date = datetime.strptime(invoice_date_str, "%d/%m/%Y")
            next_month = invoice_date.month % 12 + 1
            year = invoice_date.year + (1 if next_month == 1 else 0)
            return datetime(year, next_month, calendar.monthrange(year, next_month)[1]).strftime("%d/%m/%Y")
        except ValueError:
            return ""

    def match_tracking_option(self, po_number):
        if "C" in po_number:
            return "Caterspeed"
        elif "H" in po_number:
            return "Hotel Buyer"
        elif "R" in po_number:
            return "Restaurant Supply Store"
        elif "T" in po_number:
            return "The Restaurant Store"
        else:
            return "ERROR"

    def save_to_csv(self, data, save_path):
        columns = [
            "*ContactName", "EmailAddress", "POAddressLine1", "POAddressLine2", "POAddressLine3", 
            "POAddressLine4", "POCity", "PORegion", "POPostalCode", "POCountry", "*InvoiceNumber", 
            "*InvoiceDate", "*DueDate", "Total", "InventoryItemCode", "Description", "*Quantity", 
            "*UnitAmount", "*AccountCode", "*TaxType", "TaxAmount", "TrackingName1", 
            "TrackingOption1", "TrackingName2", "TrackingOption2", "Currency"
        ]

        with open(save_path, mode="w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for row in data:
                writer.writerow(row)

if __name__ == "__main__":
    root = TkinterDnD.Tk()  # Use TkinterDnD for drag-and-drop support
    app = InvoiceProcessorApp(root)
    root.mainloop()
