import os
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import PyPDF2
from PIL import Image
import pytesseract
import io
import re
import shutil
import sys
import subprocess
from pathlib import Path

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Check if running on Windows
is_windows = sys.platform.startswith('win')
    
class DriverPDFSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Driver PDF Sorter")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)
        
        # Data storage
        self.data_file = 'data/driver_data.json'
        os.makedirs('data', exist_ok=True)
        os.makedirs('uploads', exist_ok=True)
        
        # Create notebook with tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Excel Tab
        self.excel_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.excel_tab, text="Excel Processor")
        
        # PDF Tab
        self.pdf_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.pdf_tab, text="Route Splitter")
        
        # Setup UI for both tabs
        self.setup_excel_ui()
        self.setup_pdf_ui()
        
    def setup_excel_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.excel_tab, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Driver PDF Sorter", font=("Segoe UI", 18, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions frame
        instructions_frame = ttk.LabelFrame(main_frame, text="Instructions", padding=10)
        instructions_frame.pack(fill=tk.X, pady=(0, 20))
        
        instructions_label = ttk.Label(
            instructions_frame, 
            text="Upload an Excel file (.xlsx or .xls) to extract and store values from columns C and J.\n"
                 "The application will create a mapping between these columns and store them for later use in PDF sorting.",
            wraplength=700
        )
        instructions_label.pack(fill=tk.X)
        
        # Upload frame
        upload_frame = ttk.LabelFrame(main_frame, text="Upload Excel File", padding=10)
        upload_frame.pack(fill=tk.X, pady=(0, 20))
        
        upload_btn = ttk.Button(upload_frame, text="Select Excel File", command=self.select_file)
        upload_btn.pack(pady=10)
        
        self.file_label = ttk.Label(upload_frame, text="No file selected")
        self.file_label.pack(pady=5)
        
        process_btn = ttk.Button(upload_frame, text="Process File", command=self.process_file)
        process_btn.pack(pady=10)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding=10)
        results_frame.pack(fill=tk.BOTH, expand=True)
        
        # Search box
        search_frame = ttk.Frame(results_frame)
        search_frame.pack(fill=tk.X, pady=(0, 10))
        
        search_label = ttk.Label(search_frame, text="Search:")
        search_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.filter_treeview)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Treeview for data display
        columns = ("column_c", "column_j")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings")
        
        # Define headings
        self.tree.heading("column_c", text="Column C Value")
        self.tree.heading("column_j", text="Column J Value")
        
        # Define columns
        self.tree.column("column_c", width=100, anchor=tk.W)
        self.tree.column("column_j", width=100, anchor=tk.W)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Try to load data if it exists
        self.load_existing_data()
        
    def setup_pdf_ui(self):
        # Main frame for PDF tab
        main_frame = ttk.Frame(self.pdf_tab, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Route Splitter", font=("Segoe UI", 18, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions frame
        instructions_frame = ttk.LabelFrame(main_frame, text="Instructions", padding=10)
        instructions_frame.pack(fill=tk.X, pady=(0, 20))
        
        instructions_label = ttk.Label(
            instructions_frame, 
            text="Upload a PDF file to extract customer references and split into separate PDFs based on route information.\n"
                 "The application will use OCR to identify customer references and group pages accordingly.",
            wraplength=700
        )
        instructions_label.pack(fill=tk.X)
        
        # PDF Upload frame
        upload_frame = ttk.LabelFrame(main_frame, text="Upload PDF File", padding=10)
        upload_frame.pack(fill=tk.X, pady=(0, 20))
        
        upload_pdf_btn = ttk.Button(upload_frame, text="Select PDF File", command=self.select_pdf_file)
        upload_pdf_btn.pack(pady=10)
        
        self.pdf_file_label = ttk.Label(upload_frame, text="No file selected")
        self.pdf_file_label.pack(pady=5)
        
        # Output directory selection
        output_dir_frame = ttk.Frame(upload_frame)
        output_dir_frame.pack(fill=tk.X, pady=10)
        
        output_dir_label = ttk.Label(output_dir_frame, text="Output Directory:")
        output_dir_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.output_dir_var = tk.StringVar()
        output_dir_entry = ttk.Entry(output_dir_frame, textvariable=self.output_dir_var, width=40)
        output_dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        browse_dir_btn = ttk.Button(output_dir_frame, text="Browse", command=self.select_output_directory)
        browse_dir_btn.pack(side=tk.RIGHT)
        
        # Process button
        process_pdf_btn = ttk.Button(upload_frame, text="Process PDF", command=self.process_pdf_file)
        process_pdf_btn.pack(pady=10)
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding=10)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        # Progress indicator
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=10)
        
        # Log area
        self.log_area = ScrolledText(progress_frame, height=10)
        self.log_area.pack(fill=tk.BOTH, expand=True)
        self.log_area.config(state=tk.DISABLED)
        
        # PDF Status bar
        self.pdf_status_var = tk.StringVar()
        self.pdf_status_var.set("Ready")
        pdf_status_bar = ttk.Label(progress_frame, textvariable=self.pdf_status_var, relief=tk.SUNKEN, anchor=tk.W)
        pdf_status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def select_file(self):
        filetypes = [
            ("Excel files", "*.xlsx;*.xls"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file = filename
            self.file_label.config(text=os.path.basename(filename))
            self.status_var.set(f"File selected: {os.path.basename(filename)}")
        else:
            self.selected_file = None
    
    def process_file(self):
        if not hasattr(self, 'selected_file') or not self.selected_file:
            messagebox.showerror("Error", "No file selected")
            return
        
        try:
            self.status_var.set("Processing file...")
            self.root.update()
            
            # Read Excel file
            df = pd.read_excel(self.selected_file)
            
            # Extract values from columns C and J (indices 2 and 9)
            # Note: pandas uses 0-based indexing
            column_c = df.iloc[:, 2].astype(str)
            column_j = df.iloc[:, 9].astype(str)
            
            # Create a dictionary mapping C values to J values
            data_dict = {c_val: j_val for c_val, j_val in zip(column_c, column_j)}
            
            # Save the dictionary to a JSON file
            os.makedirs(os.path.dirname(self.data_file), exist_ok=True)
            with open(self.data_file, 'w') as f:
                json.dump(data_dict, f, indent=4)
            
            # Update the treeview
            self.update_treeview(data_dict)
            
            self.status_var.set(f"Processed {len(data_dict)} entries from {os.path.basename(self.selected_file)}")
            messagebox.showinfo("Success", f"Successfully processed {len(data_dict)} entries from the Excel file.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing file: {str(e)}")
            self.status_var.set("Error processing file")
    
    def update_treeview(self, data_dict):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Insert data
        for c_val, j_val in data_dict.items():
            self.tree.insert("", tk.END, values=(c_val, j_val))
    
    def filter_treeview(self, *args):
        search_term = self.search_var.get().lower()
        
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Try to load data
        try:
            with open(self.data_file, 'r') as f:
                data_dict = json.load(f)
            
            # Filter and insert data
            for c_val, j_val in data_dict.items():
                if search_term in c_val.lower() or search_term in j_val.lower():
                    self.tree.insert("", tk.END, values=(c_val, j_val))
        except FileNotFoundError:
            pass
    
    def load_existing_data(self):
        try:
            with open(self.data_file, 'r') as f:
                data_dict = json.load(f)
            self.update_treeview(data_dict)
            self.status_var.set(f"Loaded {len(data_dict)} entries from existing data")
        except FileNotFoundError:
            self.status_var.set("No existing data found")
        except Exception as e:
            self.status_var.set(f"Error loading data: {str(e)}")
    
    # PDF Processing Methods
    def select_pdf_file(self):
        filetypes = [
            ("PDF files", "*.pdf"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_pdf_file = filename
            self.pdf_file_label.config(text=os.path.basename(filename))
            self.pdf_status_var.set(f"PDF file selected: {os.path.basename(filename)}")
            self.log("PDF file selected: " + os.path.basename(filename))
        else:
            self.selected_pdf_file = None
    
    def select_output_directory(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        
        if directory:
            self.output_dir_var.set(directory)
            self.log(f"Output directory set to: {directory}")
    
    def log(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)
        self.root.update()
    
    def extract_text_from_page(self, pdf_path, page_num):
        try:
            # For Windows, we'll use PyPDF2 directly since poppler is harder to set up
            if is_windows:
                # Extract text directly from PDF
                with open(pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    if page_num < len(pdf_reader.pages):
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()
                        
                        # Log a sample of the extracted text for debugging
                        sample_text = page_text[:200] + "..." if len(page_text) > 200 else page_text
                        self.log(f"Sample text from page {page_num+1}:\n{sample_text}")
                        
                        return page_text
                    else:
                        return ""
            else:
                # Try to use pdf2image and pytesseract (requires poppler)
                from pdf2image import convert_from_path
                images = convert_from_path(pdf_path, first_page=page_num+1, last_page=page_num+1)
                if not images:
                    return ""
                
                # Perform OCR on the image with improved settings
                page_text = pytesseract.image_to_string(images[0], config='--psm 6')
                sample_text = page_text[:200] + "..." if len(page_text) > 200 else page_text
                self.log(f"Sample text from page {page_num+1}:\n{sample_text}")
                
                return page_text
                
        except Exception as e:
            self.log(f"Error extracting text from page {page_num}: {str(e)}")
            self.log("Falling back to direct PDF text extraction...")
            
            # Fallback to direct PDF text extraction
            try:
                with open(pdf_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    if page_num < len(pdf_reader.pages):
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()
                        
                        # Log a sample of the extracted text for debugging
                        sample_text = page_text[:200] + "..." if len(page_text) > 200 else page_text
                        self.log(f"Sample text (fallback) from page {page_num+1}:\n{sample_text}")
                        
                        return page_text
            except Exception as inner_e:
                self.log(f"Fallback also failed: {str(inner_e)}")
            
            return ""
    
    def find_customer_ref(self, text):
        if not text:
            return None
            
        # Look for exact "Customer Ref." followed by value format
        # Pattern specifically for the Around Noon format seen in the screenshots
        customer_ref_pattern = re.compile(r'Customer\s+Ref\.?\s*[:.]?\s*([A-Z0-9]{3,10})\b', re.IGNORECASE)
        match = customer_ref_pattern.search(text)
        if match:
            ref = match.group(1).strip()
            self.log(f"Found exact Customer Ref: {ref}")
            return ref
            
        # Try another pattern - look for Customer Ref followed by uppercase alphanumeric
        alt_pattern = re.compile(r'(?:Customer|Cust)[\s\.]+Ref[\s\.]*[:.]?\s*([A-Z][A-Z0-9]{2,9})\b', re.IGNORECASE)
        match = alt_pattern.search(text)
        if match:
            ref = match.group(1).strip()
            self.log(f"Found Customer Ref with alt pattern: {ref}")
            return ref
            
        # Special case for ARAM pattern as seen in screenshot
        aram_pattern = re.compile(r'\bARAM\d{3}\b')
        match = aram_pattern.search(text)
        if match:
            ref = match.group(0).strip()
            self.log(f"Found ARAM reference: {ref}")
            return ref
            
        # Special case for KSG pattern as seen in previous screenshot
        ksg_pattern = re.compile(r'\bKSG[A-Z]?\d{2,4}\b')
        match = ksg_pattern.search(text)
        if match:
            ref = match.group(0).strip()
            self.log(f"Found KSG reference: {ref}")
            return ref
        
        # Only as a last resort, try to find patterns directly from the database
        try:
            with open(self.data_file, 'r') as f:
                customer_data = json.load(f)
                
            # Look for any customer ref from our database directly in the text
            for db_ref in customer_data.keys():
                if db_ref and len(db_ref) >= 3 and db_ref in text:
                    self.log(f"Found direct database match: {db_ref}")
                    return db_ref
        except:
            pass
            
        return None
    
    def process_pdf_file(self):
        if not hasattr(self, 'selected_pdf_file') or not self.selected_pdf_file:
            messagebox.showerror("Error", "No PDF file selected")
            return
        
        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showerror("Error", "No output directory selected")
            return
        
        try:
            # Load customer data (maps customer refs to routes)
            try:
                with open(self.data_file, 'r') as f:
                    customer_data = json.load(f)
                    
                self.log(f"Loaded customer data with {len(customer_data)} entries")
                self.log("Sample customer data entries:")
                sample_count = 0
                for ref, route in customer_data.items():
                    if sample_count < 5:  # Show first 5 entries
                        self.log(f"  {ref} -> {route}")
                        sample_count += 1
                    else:
                        break
                        
            except FileNotFoundError:
                messagebox.showerror("Error", "No customer data found. Please process an Excel file first.")
                return
            
            self.pdf_status_var.set("Processing PDF...")
            self.log("\nStarting PDF processing...")
            
            # Dictionary to store pages by route
            pages_by_route = {}
            customer_routes = {}  # Maps customer refs to their routes
            unassigned_pages = []
            
            # Read the PDF
            with open(self.selected_pdf_file, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                total_pages = len(pdf_reader.pages)
                
                self.log(f"PDF has {total_pages} pages")
                
                # Process each page
                for i in range(total_pages):
                    self.progress_var.set((i / total_pages) * 100)
                    self.pdf_status_var.set(f"Processing page {i+1} of {total_pages}")
                    self.log(f"\nProcessing page {i+1}...")
                    
                    # Extract text from the page
                    page_text = self.extract_text_from_page(self.selected_pdf_file, i)
                    
                    # Find customer reference in the text
                    customer_ref = self.find_customer_ref(page_text)
                    
                    if customer_ref:
                        self.log(f"Found customer reference: {customer_ref} on page {i+1}")
                        
                        # Check if customer reference exists in our database
                        if customer_ref in customer_data:
                            route = customer_data[customer_ref]
                            self.log(f"Customer {customer_ref} belongs to route: {route}")
                            
                            # Store customer and route mapping
                            customer_routes[customer_ref] = route
                            
                            # Add page to the route
                            if route not in pages_by_route:
                                pages_by_route[route] = []
                            
                            pages_by_route[route].append(i)
                        else:
                            self.log(f"Customer {customer_ref} not found in database, checking for exact matches...")
                            
                            # Try to find exact matches only - no more partial matching
                            found_match = False
                            for db_ref, route in customer_data.items():
                                # Only exact match after normalization
                                clean_customer_ref = customer_ref.replace(" ", "").upper()
                                clean_db_ref = db_ref.replace(" ", "").upper()
                                
                                if clean_customer_ref == clean_db_ref:
                                    self.log(f"Exact match found: {customer_ref} = {db_ref} -> {route}")
                                    
                                    # Store customer and route mapping
                                    customer_routes[customer_ref] = route
                                    
                                    # Add page to the route
                                    if route not in pages_by_route:
                                        pages_by_route[route] = []
                                    
                                    pages_by_route[route].append(i)
                                    found_match = True
                                    break
                            
                            if not found_match:
                                self.log(f"No exact match found for customer {customer_ref}")
                                unassigned_pages.append(i)
                    else:
                        self.log(f"No customer reference found on page {i+1}")
                        unassigned_pages.append(i)
                
                # Create PDF for each route (with route name as filename)
                self.log("\nCreating output PDFs by route:")
                
                for route, pages in pages_by_route.items():
                    # Create valid filename from route name - preserve exact route name
                    safe_route_name = route
                    
                    # Only replace characters that are invalid in filenames
                    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
                    for char in invalid_chars:
                        safe_route_name = safe_route_name.replace(char, '_')
                    
                    # Ensure we have a valid filename
                    if not safe_route_name or safe_route_name.isspace():
                        safe_route_name = "Unknown_Route"
                    
                    output_path = os.path.join(output_dir, f"{safe_route_name}.pdf")
                    
                    self.log(f"Creating PDF for route '{route}' with {len(pages)} pages: {output_path}")
                    
                    # Create a new PDF
                    pdf_writer = PyPDF2.PdfWriter()
                    
                    # Add pages from the original PDF
                    with open(self.selected_pdf_file, 'rb') as input_file:
                        pdf_reader = PyPDF2.PdfReader(input_file)
                        
                        for page_num in pages:
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                    # Write the output file
                    with open(output_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    
                    self.log(f"Created: {output_path}")
                
                # Create PDF for unassigned pages
                if unassigned_pages:
                    output_path = os.path.join(output_dir, "Unassigned_Pages.pdf")
                    
                    self.log(f"\nCreating PDF for {len(unassigned_pages)} unassigned pages")
                    
                    # Create a new PDF
                    pdf_writer = PyPDF2.PdfWriter()
                    
                    # Add pages from the original PDF
                    with open(self.selected_pdf_file, 'rb') as input_file:
                        pdf_reader = PyPDF2.PdfReader(input_file)
                        
                        for page_num in unassigned_pages:
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                    
                    # Write the output file
                    with open(output_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                    
                    self.log(f"Created: {output_path}")
                
                # Create a summary
                self.log("\nProcessing Summary:")
                self.log(f"Total pages: {total_pages}")
                self.log(f"Routes created: {len(pages_by_route)}")
                self.log(f"Customer references found and matched: {len(customer_routes)}")
                self.log(f"Pages assigned to routes: {sum(len(pages) for pages in pages_by_route.values())}")
                self.log(f"Unassigned pages: {len(unassigned_pages)}")
                
                if pages_by_route:
                    self.log("\nRoute details:")
                    for route, pages in pages_by_route.items():
                        customer_list = [ref for ref, r in customer_routes.items() if r == route]
                        self.log(f"  Route '{route}': {len(pages)} pages, {len(customer_list)} customers")
                        for customer in customer_list[:5]:  # Show max 5 customers per route
                            self.log(f"    - {customer}")
                        if len(customer_list) > 5:
                            self.log(f"    - ... and {len(customer_list)-5} more customers")
                
                # Show success message
                self.progress_var.set(100)
                self.pdf_status_var.set("PDF processing complete")
                messagebox.showinfo("Success", "PDF processing complete")
        
        except Exception as e:
            self.log(f"Error processing PDF: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            self.pdf_status_var.set("Error processing PDF")
            messagebox.showerror("Error", f"Error processing PDF: {str(e)}")

def main():
    root = tk.Tk()
    app = DriverPDFSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 