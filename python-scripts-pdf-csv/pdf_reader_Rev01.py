import os
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox
from tkinterdnd2 import DND_FILES, TkinterDnD
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
        # Open file dialog to select PDF files
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        for file_path in file_paths:
            self.add_file(file_path)

    def on_drop(self, event):
        # Handle drag-and-drop event
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            if file_path.endswith('.pdf'):
                self.add_file(file_path)
            else:
                messagebox.showwarning("Invalid File", f"{file_path} is not a PDF file and was ignored.")

    def add_file(self, file_path):
        # Add file to the list if not already present
        if file_path not in self.files:
            self.files.append(file_path)
            self.file_listbox.insert(tk.END, os.path.basename(file_path))

    def process_files(self):
        if not self.files:
            messagebox.showwarning("No Files Selected", "Please select or drop PDF files to process.")
            return

        # Placeholder for CSV data
        csv_data = []

        for file in self.files:
            # Extract data from each PDF
            extracted_data = self.extract_data_from_pdf(file)
            if extracted_data:
                csv_data.append(extracted_data)

        # Ask user for save location and save the CSV file
        save_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if save_path:
            self.save_to_csv(csv_data, save_path)
            messagebox.showinfo("Process Complete", f"Data extracted and saved to {save_path}.")

    def extract_data_from_pdf(self, file_path):
        # Define data dictionary with required headers
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
            "*Quantity": "1",
            "*UnitAmount": "",
            "*AccountCode": "540",
            "*TaxType": "20% (VAT on Expenses)",
            "TaxAmount": "",
            "TrackingName1": "Website",
            "TrackingOption1": "",  # Filled based on purchase order number
            "TrackingName2": "",  # Blank
            "TrackingOption2": "",  # Blank
            "Currency": "GBP"
        }

        with pdfplumber.open(file_path) as pdf:
            text = ""
            for page in pdf.pages:
                text += page.extract_text() or ""

        prompt = f"Extract the following details from this invoice text:\n\n{text}\n\n"
        prompt += "Extract: Supplier name, Invoice number, Purchase order number, Value, Invoice date, and Due date."

        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant that extracts invoice details."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=100
            )
            
            # Extract the AI's response text
            ai_response = response.choices[0].message['content'].strip()
            print("AI Response:", ai_response)  # Debug output

            ai_extracted_data = ai_response.split("\n")
            for line in ai_extracted_data:
                if "Supplier" in line:
                    data["*ContactName"] = line.split(":")[1].strip()
                elif "Invoice number" in line:
                    data["*InvoiceNumber"] = line.split(":")[1].strip()
                elif "Purchase order number" in line:
                    po_number = line.split(":")[1].strip()
                    
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
                    data["*UnitAmount"] = line.split(":")[1].strip()
                elif "Invoice date" in line:
                    data["*InvoiceDate"] = line.split(":")[1].strip()
                elif "Due date" in line:
                    data["*DueDate"] = line.split(":")[1].strip()

        except Exception as e:
            print("Error with OpenAI API:", e)
            messagebox.showerror("AI Error", f"Could not process the invoice data using AI. Error: {e}")
            return None

        return data

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
