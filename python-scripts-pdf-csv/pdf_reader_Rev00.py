import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
import pdfplumber
import csv
from dotenv import load_dotenv
import openai

# Load environment variables from a .env file
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

class InvoiceProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Processor")
        self.root.geometry("400x300")

        # List to store file paths
        self.files = []

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
        # Open file dialog to select PDF files
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file_path in file_paths:
            self.files.append(file_path)
            self.file_listbox.insert(tk.END, os.path.basename(file_path))

    def process_files(self):
        if not self.files:
            messagebox.showwarning("No Files Selected", "Please select PDF files to process.")
            return

        # Placeholder for CSV data
        csv_data = []

        for file in self.files:
            # Extract data from each PDF
            extracted_data = self.extract_data_from_pdf(file)
            if extracted_data:
                csv_data.append(extracted_data)

        # Save data to CSV
        self.save_to_csv(csv_data)

        messagebox.showinfo("Process Complete", "Data extracted and saved to invoices.csv.")

    def extract_data_from_pdf(self, file_path):
        # Placeholder dictionary to match template structure with empty address fields
        data = {
            "*ContactName": "",
            "EmailAddress": "",  # Blank
            "POAddressLine1": "", "POAddressLine2": "", "POAddressLine3": "", "POAddressLine4": "",
            "POCity": "", "PORegion": "", "POPostalCode": "", "POCountry": "",  # Blank address fields
            "*InvoiceNumber": "",
            "*InvoiceDate": "",
            "*DueDate": "",
            "Total": "",
            "InventoryItemCode": "",
            "Description": "",
            "*Quantity": "",
            "*UnitAmount": "",
            "*AccountCode": "",
            "*TaxType": "",
            "TaxAmount": "",
            "TrackingName1": "Website",  # Fixed to "Website" as per requirement
            "TrackingOption1": "",  # Will set based on order type
            "TrackingName2": "",
            "TrackingOption2": "",
            "Currency": "GBP"  # Example currency
        }

        # Open PDF and extract text
        with pdfplumber.open(file_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""

        # Use OpenAI API for field extraction
        prompt = f"Extract the following details from this invoice text:\n\n{text}\n\n"
        prompt += "Extract: Supplier name, Invoice number, Purchase order number, Value, Invoice date, and Due date."

        try:
            response = openai.Completion.create(
                engine="text-davinci-003",
                prompt=prompt,
                max_tokens=100,
                temperature=0
            )

            # Process the AI response
            ai_extracted_data = response.choices[0].text.strip().split("\n")
            for line in ai_extracted_data:
                if "Supplier" in line:
                    data["*ContactName"] = line.split(":")[1].strip()
                elif "Invoice number" in line:
                    data["*InvoiceNumber"] = line.split(":")[1].strip()
                elif "Purchase order number" in line:
                    po_number = line.split(":")[1].strip()
                    
                    # Logic to determine TrackingOption1 based on PO number
                    if "C" in po_number:
                        data["TrackingOption1"] = "Caterspeed"
                    elif "H" in po_number:
                        data["TrackingOption1"] = "Hotel Buyer"
                    elif "R" in po_number:
                        data["TrackingOption1"] = "Restaurant Supply Store"
                    else:
                        data["TrackingOption1"] = "The Restaurant Store"
                        
                elif "Value" in line:
                    data["Total"] = line.split(":")[1].strip()
                    data["*UnitAmount"] = line.split(":")[1].strip()  # Assuming single-item quantity
                elif "Invoice date" in line:
                    data["*InvoiceDate"] = line.split(":")[1].strip()
                elif "Due date" in line:
                    data["*DueDate"] = line.split(":")[1].strip()

            data["*Quantity"] = "1"  # Default quantity to 1
            data["*AccountCode"] = "540"  # Example account code
            data["*TaxType"] = "20% VAT"  # Example tax type

        except Exception as e:
            print("Error with OpenAI API:", e)
            messagebox.showerror("AI Error", "Could not process the invoice data using AI.")

        return data

    def save_to_csv(self, data):
        # Define template columns
        columns = [
            "*ContactName", "EmailAddress", "POAddressLine1", "POAddressLine2", "POAddressLine3", 
            "POAddressLine4", "POCity", "PORegion", "POPostalCode", "POCountry", "*InvoiceNumber", 
            "*InvoiceDate", "*DueDate", "Total", "InventoryItemCode", "Description", "*Quantity", 
            "*UnitAmount", "*AccountCode", "*TaxType", "TaxAmount", "TrackingName1", 
            "TrackingOption1", "TrackingName2", "TrackingOption2", "Currency"
        ]

        with open("invoices.csv", mode="w", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=columns)
            writer.writeheader()
            for row in data:
                writer.writerow(row)

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceProcessorApp(root)
    root.mainloop()
