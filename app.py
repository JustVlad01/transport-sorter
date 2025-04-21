import os
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

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
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=20)
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

def main():
    root = tk.Tk()
    app = DriverPDFSorterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 