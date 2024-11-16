import os
import re
import json
import requests
import pytesseract
from pdf2image import convert_from_path
import tkinter as tk
from tkinter import filedialog, messagebox, Listbox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import pdfplumber
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

WEBHOOK_URL = os.getenv("MAKE_WEBHOOK_URL")

class InvoiceProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Processor with Drag-and-Drop")
        self.root.geometry("700x500")

        self.files = []

        # Drag-and-Drop Setup
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)

        # Buttons and Listboxes
        self.select_button = tk.Button(root, text="Select PDF Files", command=self.select_files)
        self.select_button.pack(pady=5)

        self.file_listbox = Listbox(root, width=50, height=10, selectmode=tk.BROWSE)
        self.file_listbox.pack(pady=5)

        self.progress_label = tk.Label(root, text="Progress:")
        self.progress_label.pack()

        self.progress_bar = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=400, mode="determinate")
        self.progress_bar.pack(pady=5)

        self.processed_label = tk.Label(root, text="Processed Files:")
        self.processed_label.pack()

        self.processed_listbox = Listbox(root, width=50, height=5)
        self.processed_listbox.pack(pady=5)

        self.failed_label = tk.Label(root, text="Failed Files:")
        self.failed_label.pack()

        self.failed_listbox = Listbox(root, width=50, height=5)
        self.failed_listbox.pack(pady=5)

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

        total_files = len(self.files)
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0

        for file in list(self.files):  # Iterate over a copy of the file list
            extracted_data = self.extract_data_from_pdf(file)
            if extracted_data:
                success = self.send_to_webhook(extracted_data)
                if success:
                    self.processed_listbox.insert(tk.END, os.path.basename(file))
                else:
                    self.failed_listbox.insert(tk.END, os.path.basename(file))
            else:
                self.failed_listbox.insert(tk.END, os.path.basename(file))

            # Update progress bar and remove processed file from the list
            self.files.remove(file)
            self.file_listbox.delete(0)  # Remove the first item in the listbox
            self.progress_bar["value"] += 1
            self.root.update_idletasks()  # Refresh the UI dynamically

        if not self.failed_listbox.size():
            messagebox.showinfo("Success", "All files processed successfully!")
        else:
            messagebox.showwarning("Partial Success", "Some files failed to process. Check the failed files list.")

        # To store names of files that failed to send
        failed_files = []

        for file in self.files:
            extracted_data = self.extract_data_from_pdf(file)
            if extracted_data:
                success = self.send_to_webhook(extracted_data)
                if not success:
                    # Add the file name to the failure list
                    failed_files.append(os.path.basename(file))

        # Show a single message after processing all files
        if not failed_files:
            messagebox.showinfo("Success", "All extracted data were sent successfully!")
            self.clear_files()
        else:
            failed_files_str = "\n".join(failed_files)
            messagebox.showerror(
                "Error",
                f"Some files failed to send:\n{failed_files_str}\nCheck the logs for more details."
            )

    def clear_files(self):
        """Clear the selected files and reset the listbox."""
        # Clear the list of files
        self.files = []
         # Reset the listbox
        self.file_listbox.delete(0, tk.END)

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
        image_text = ""
        try:
            images = convert_from_path(file_path, dpi=300, poppler_path=POPPLER_PATH)
            for image in images:
                image_text += pytesseract.image_to_string(image)
        except Exception as e:
            print(f"Error converting PDF to images or using OCR: {e}")
            return None

        combined_text = text + "\n" + image_text

        # Debug: Log combined text for troubleshooting
        print("Extracted Text from PDF and Images:\n", combined_text)

        prompt = f"""
        You are an intelligent assistant designed to extract structured data from invoices. Below is the invoice text:

        {combined_text}

        Your task:
        - Extract key details from the invoice text and return the data in a **valid JSON** format.
        - Use context from the invoice (e.g., headings, labels, and patterns) to identify each field correctly.
        - Follow these instructions for each field:

        1. **ContactName**: Extract the **company name** based on prominent branding, header, or logo text (e.g., DUCK ISLAND). Avoid supplier names in "Ship To" or "Billing Address" unless they are clearly the company.
        2. **EmailAddress**: Extract the first valid email address (e.g., custserv@nisbets.co.uk). If no email is present, leave it as an empty string.
        3. **POAddressLine1-4**: Extract up to 4 address lines under the "Ship To" or "Delivery Address" section. Ensure the lines are in the correct order. If there are fewer than 4 lines, leave the remaining lines as empty strings.
        4. **POCity**: Extract the city from the shipping address.
        5. **PORegion**: Extract the region, county, or state from the shipping address, if provided. Leave blank if missing.
        6. **POPostalCode**: Extract the postal code from the shipping address. Ensure correct formatting (e.g., SM4 4LU).
        7. **POCountry**: Extract the country from the shipping address, if explicitly mentioned. Leave blank if missing.
        8. **InvoiceNumber**: Extract the invoice number (e.g., 30114156) from headings like "Invoice No" or "Invoice Number."
        9. **InvoiceDate**: Extract the invoice date (e.g., 13/11/2024) and ensure it is in DD/MM/YYYY format.
        10. **DueDate**: Calculate the due date based on payment terms (e.g., "30 days from the invoice date") and display it in DD/MM/YYYY format. If payment terms are missing, assume a default of 30 days.
        11. **Total**: Extract the total invoice amount (e.g., 55.82 GBP). If the currency is explicitly mentioned, include it; otherwise, default to "GBP."
        12. **InventoryItemCode**: Extract all item codes (e.g., "C/HW5000(2)") listed in the product table.
        13. **Description**: Extract all product descriptions (e.g., "Classic Hand Wash 5L packed in 2") listed in the product table.
        14. **Quantity**: Extract all quantities (e.g., "1") from the product table.
        15. **UnitAmount**: Extract all unit prices (e.g., "33.57") for items in the product table.
        16. **AccountCode**: Default to "540" unless another account code is explicitly mentioned.
        17. **TaxType**: Extract the tax type (e.g., "20% (VAT on Expenses)"). Default to "20% (VAT on Expenses)" if not specified.
        18. **TaxAmount**: Extract the total tax amount (e.g., 9.30 GBP) from the invoice.
        19. **TrackingName1**: Extract any tracking names or labels (e.g., "Order Reference").
        20. **TrackingOption1**: Extract any tracking option values (e.g., "H150690").
        21. **TrackingName2** and **TrackingOption2**: Extract any additional tracking details, if available. Leave blank if none exist.
        22. **Currency**: Extract the currency (e.g., GBP). Default to "GBP" if not explicitly mentioned.

        ### Important Notes:
        - Ensure all extracted data matches the context and structure of the invoice.
        - Format your response as valid JSON with proper key-value pairs for all fields. Missing or unavailable fields should have an empty string ("") as their value.
        - If data for certain fields exists in multiple places (e.g., addresses), prioritize the most relevant section (e.g., "Ship To" for shipping details).

        ### Example JSON Output:
        {{
            "*ContactName": "Duck Island Limited",
            "EmailAddress": "sales@duckisland.co.uk",
            "POAddressLine1": "The Townhouse",
            "POAddressLine2": "High Street",
            "POAddressLine3": "Sutton Coldfield",
            "POAddressLine4": "Suburban Inns Operations Ltd",
            "POCity": "Sutton Coldfield",
            "PORegion": "",
            "POPostalCode": "B72 1UD",
            "POCountry": "",
            "*InvoiceNumber": "0000027558",
            "*InvoiceDate": "12/11/2024",
            "*DueDate": "12/12/2024",
            "Total": "55.82",
            "InventoryItemCode": ["C/HW5000(2)", "Car"],
            "Description": ["Classic Hand Wash 5L packed in 2", "Carriage as"],
            "*Quantity": ["1", "1"],
            "*UnitAmount": ["33.57", "12.95"],
            "*AccountCode": "540",
            "*TaxType": "20% (VAT on Expenses)",
            "TaxAmount": "9.30",
            "TrackingName1": "Order Reference",
            "TrackingOption1": "H150690",
            "TrackingName2": "",
            "TrackingOption2": "",
            "Currency": "GBP"
        }}

        Please ensure your response is in valid JSON format with no additional explanations or text.
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
            print("AI Response JSON:\n", ai_response)
            extracted_data = json.loads(ai_response)
            return extracted_data
        except Exception as e:
            print("Error with OpenAI API:", e)
            return None

    def send_to_webhook(self, data):
        try:
            response = requests.post(WEBHOOK_URL, json=data)
            if response.status_code == 200:
                print("Data sent successfully to webhook!")
                return True
            else:
                print(f"Failed to send data to webhook. Status code: {response.status_code}, Response: {response.text}")
                return False
        except Exception as e:
            print(f"Error sending data to webhook: {e}")
            return False


if __name__ == "__main__":
    root = TkinterDnD.Tk()  # Use TkinterDnD for drag-and-drop support
    app = InvoiceProcessorApp(root)
    root.mainloop()