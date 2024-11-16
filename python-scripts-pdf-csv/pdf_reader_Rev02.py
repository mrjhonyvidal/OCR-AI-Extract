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
        You are a helpful assistant for extracting structured details from invoices. Below is the invoice text:

        {combined_text}

        Your task is to extract the following details:
        1. *ContactName: Extract the company name based on branding text in the invoice (e.g., DUCK ISLAND). Avoid using supplier names like "Catercall Ltd". Ignore logos, URLs, or IP addresses.
        2. EmailAddress: Extract the email address (e.g., custserv@nisbets.co.uk).
        3. POAddressLine1-4: Extract up to 4 lines of the address. Use "Ship To:" or similar contextual clues.
        4. POCity: Extract the city from the address.
        5. PORegion: Extract the region (e.g., county/state) from the address.
        6. POPostalCode: Extract the postal code (e.g., SM4 4LU).
        7. POCountry: Extract the country (if present).
        8. *InvoiceNumber: Extract the invoice number (e.g., 30114156).
        9. *InvoiceDate: Extract the invoice date and format it as DD/MM/YYYY.
        10. *DueDate: Calculate the due date based on the payment terms (e.g., net 30 days from the invoice date).
        11. Total: Extract the total invoice value (e.g., 11.38 GBP).
        12. InventoryItemCode, Description, and *Quantity: Extract these fields from the product description table, if present.
        13. *UnitAmount: Extract the unit price of items, if present.
        14. *AccountCode: Set to "540" by default unless another account code is explicitly mentioned.
        15. *TaxType: Extract or default to "20% (VAT on Expenses)".
        16. TaxAmount: Extract the tax value (e.g., 1.89 GBP).
        17. TrackingName1: Extract any tracking names or descriptions (e.g., Despatch No).
        18. TrackingOption1: Extract tracking options, such as order references (e.g., 32596160).
        19. TrackingName2, TrackingOption2: Extract any additional tracking details, if present.
        20. Currency: Extract or default to "GBP".

        Provide the results in this exact format:
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
        InventoryItemCode: [Item Code]
        Description: [Product Description]
        *Quantity: [Quantity]
        *UnitAmount: [Unit Price]
        *AccountCode: [Account Code]
        *TaxType: [Tax Type]
        TaxAmount: [Tax Amount]
        TrackingName1: [Tracking Name]
        TrackingOption1: [Tracking Option]
        TrackingName2: [Additional Tracking Name]
        TrackingOption2: [Additional Tracking Option]
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
                    data["*ContactName"] = line.split(":", 1)[1].strip()
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
                elif "TrackingOption1" in line:
                    data["TrackingOption1"] = line.split(":", 1)[1].strip()

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
