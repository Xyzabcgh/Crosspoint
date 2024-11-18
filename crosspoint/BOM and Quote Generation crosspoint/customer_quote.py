#5th option customer quote generator based on combined source list as suggested by John anna
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
from datetime import datetime
import logging
import threading

# Function to clean and convert values to numeric (handling large values or non-numeric data)
def safe_convert_to_numeric(value):
    try:
        value = str(value).replace(',', '').strip()
        return pd.to_numeric(value, errors='coerce')  # Coerce errors to NaN
    except Exception:
        return float('nan')

# Main processing function for the BOM
def process_bom(input_base_dir, output_base_dir, source_mapping_file, column_mapping_file, log_widget, update_log_callback):
    try:
        # Load the shared Source Mapping sheet
        source_mapping_df = pd.read_csv(source_mapping_file)
        source_mapping = source_mapping_df.set_index("Manufacturer Part Number")["Source"].to_dict()

        # Load the column mapping sheet for vendor-specific column positions
        column_mapping_df = pd.read_excel(column_mapping_file, index_col=0)

        # Configure logging
        logging.basicConfig(filename='process_log.txt', level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')
        logger = logging.getLogger()

        os.makedirs(output_base_dir, exist_ok=True)

        # Process each customer's folder in the input directory
        for customer_name in os.listdir(input_base_dir):
            customer_dir = os.path.join(input_base_dir, customer_name)
            
            # Ensure it's a directory before proceeding
            if not os.path.isdir(customer_dir):
                update_log_callback(f"Skipping {customer_name}, not a directory.\n", log_widget)
                continue
            
            # Load the BOM sheet for the customer
            bom_file = os.path.join(customer_dir, f"{customer_name}_BOM.xlsx")
            bom_file = os.path.normpath(bom_file) 
            
            if not os.path.isfile(bom_file):
                update_log_callback(f"Missing BOM file for {customer_name}. Skipping...\n", log_widget)
                continue
            
            try:
                bom_df = pd.read_excel(bom_file)
                update_log_callback(f"Loaded BOM sheet for {customer_name}\n", log_widget)
            except Exception as e:
                update_log_callback(f"Error reading BOM file for {customer_name}: {e}\n", log_widget)
                continue

            # Load vendor quote sheets and add their price columns to the BOM
            vendor_quotes = {}
            for vendor_file in os.listdir(customer_dir):
                if "_Quote.xlsx" in vendor_file:
                    vendor_name = vendor_file.split("_")[0]
                    vendor_path = os.path.join(customer_dir, vendor_file)
                    try:
                        vendor_df = pd.read_excel(vendor_path)
                        update_log_callback(f"Loaded quote file for vendor: {vendor_name}\n", log_widget)
                        
                        if vendor_name in column_mapping_df.columns:
                            vendor_col_mapping = column_mapping_df[vendor_name]
                            
                            # Map column indices from the template (zero-based indexing)
                            part_num_col = int(vendor_col_mapping["Manufacturer Part Number"])
                            price_col = int(vendor_col_mapping["Unit Price"])
                            availability_col = int(vendor_col_mapping["Availability"])
                            order_qty_col = int(vendor_col_mapping["Order Quantity"])
                            
                            # Extract and rename columns based on these indices
                            vendor_df = vendor_df.iloc[:, [part_num_col, price_col, availability_col, order_qty_col]]
                            vendor_df.columns = ["Manufacturer Part Number", "Price", "Availability", "Order Quantity"]
                            
                            # Clean 'Price' column and convert to numeric
                            vendor_df["Price"] = vendor_df["Price"].replace({'\$': '', ',': ''}, regex=True)
                            vendor_df["Price"] = pd.to_numeric(vendor_df["Price"], errors='coerce')
                            
                            # Convert 'Availability' and 'Order Quantity' columns using the safe function
                            vendor_df["Availability"] = vendor_df["Availability"].apply(safe_convert_to_numeric)
                            vendor_df["Order Quantity"] = vendor_df["Order Quantity"].apply(safe_convert_to_numeric)
                            
                            # Calculate the computed price only if Availability >= Order Quantity
                            vendor_df["Computed Price"] = vendor_df.apply(lambda row: row["Price"] * row["Order Quantity"]
                                                                        if pd.notna(row["Availability"]) and row["Availability"] >= row["Order Quantity"]
                                                                        else None, axis=1)
                            
                            # Save the filtered vendor quote for later processing
                            vendor_quotes[vendor_name] = vendor_df
                        else:
                            update_log_callback(f"Vendor {vendor_name} not found in column mapping file.\n", log_widget)
                    except Exception as e:
                        update_log_callback(f"Error reading vendor file {vendor_file} for {customer_name}: {e}\n", log_widget)
                        continue

            # Remove duplicate part numbers in BOM and vendor data before merging
            bom_df = bom_df.drop_duplicates(subset="Manufacturer Part Number")
            
            # Merge vendor price data into the BOM
            for vendor, df in vendor_quotes.items():
                df = df.drop_duplicates(subset="Manufacturer Part Number")
                price_column_name = f"{vendor}_Price"
                price_data = df[["Manufacturer Part Number", "Computed Price"]].rename(columns={"Computed Price": price_column_name})
                bom_df = bom_df.merge(price_data, on="Manufacturer Part Number", how="left")

            # Determine final price and source
            final_prices = []
            sources = []
            for _, row in bom_df.iterrows():
                part_number = row["Manufacturer Part Number"]
                preferred_vendor = source_mapping.get(part_number)
                price = row.get(f"{preferred_vendor}_Price") if preferred_vendor else None
                final_prices.append(price)
                sources.append(preferred_vendor)
            
            bom_df["Price"] = final_prices
            bom_df["Source"] = sources

            # Create output directory for the customer
            customer_output_dir = os.path.join(output_base_dir, customer_name)
            os.makedirs(customer_output_dir, exist_ok=True)
            
            # Generate a unique filename with date and time
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"{customer_name}_Updated_BOM_Sheet_{timestamp}.xlsx"
            output_path = os.path.join(customer_output_dir, output_filename)
            
            # Save the updated BOM sheet
            try:
                bom_df.to_excel(output_path, index=False)
                update_log_callback(f"Updated BOM sheet saved for {customer_name} at {output_path}\n", log_widget)
            except Exception as e:
                update_log_callback(f"Error saving output file for {customer_name}: {e}\n", log_widget)

    except Exception as e:
        update_log_callback(f"Error processing BOM files: {e}\n", log_widget)

# GUI Class
class BOMProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BOM Processor")

        # Input folder selection
        tk.Label(root, text="Select Customer Input Folder:").grid(row=0, column=0, padx=10, pady=5)
        self.input_dir_entry = tk.Entry(root, width=40)
        self.input_dir_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.select_input_folder).grid(row=0, column=2, padx=10, pady=5)

        # Output folder selection
        tk.Label(root, text="Select Output Folder:").grid(row=1, column=0, padx=10, pady=5)
        self.output_dir_entry = tk.Entry(root, width=40)
        self.output_dir_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.select_output_folder).grid(row=1, column=2, padx=10, pady=5)

        # Source mapping file selection
        tk.Label(root, text="Select Source Mapping CSV:").grid(row=2, column=0, padx=10, pady=5)
        self.source_mapping_entry = tk.Entry(root, width=40)
        self.source_mapping_entry.grid(row=2, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.select_source_mapping_file).grid(row=2, column=2, padx=10, pady=5)

        # Column mapping file selection
        tk.Label(root, text="Select Column Mapping Excel:").grid(row=3, column=0, padx=10, pady=5)
        self.column_mapping_entry = tk.Entry(root, width=40)
        self.column_mapping_entry.grid(row=3, column=1, padx=10, pady=5)
        tk.Button(root, text="Browse", command=self.select_column_mapping_file).grid(row=3, column=2, padx=10, pady=5)

        # Process button
        tk.Button(root, text="Process BOM", command=self.start_processing).grid(row=4, column=1, padx=10, pady=10)

        # Log output
        self.log_text = tk.Text(root, width=60, height=15)
        self.log_text.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

    def select_input_folder(self):
        folder_selected = filedialog.askdirectory()
        self.input_dir_entry.delete(0, tk.END)
        self.input_dir_entry.insert(0, folder_selected)

    def select_output_folder(self):
        folder_selected = filedialog.askdirectory()
        self.output_dir_entry.delete(0, tk.END)
        self.output_dir_entry.insert(0, folder_selected)

    def select_source_mapping_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.source_mapping_entry.delete(0, tk.END)
        self.source_mapping_entry.insert(0, file_selected)

    def select_column_mapping_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.column_mapping_entry.delete(0, tk.END)
        self.column_mapping_entry.insert(0, file_selected)

    def start_processing(self):
        input_dir = self.input_dir_entry.get()
        output_dir = self.output_dir_entry.get()
        source_mapping_file = self.source_mapping_entry.get()
        column_mapping_file = self.column_mapping_entry.get()

        # Ensure all fields are filled
        if not all([input_dir, output_dir, source_mapping_file, column_mapping_file]):
            messagebox.showerror("Error", "Please fill in all fields.")
            return

        # Start processing in a separate thread to keep the GUI responsive
        threading.Thread(target=process_bom, args=(input_dir, output_dir, source_mapping_file, column_mapping_file, self.log_text, self.update_log)).start()

    def update_log(self, message, log_widget):
        log_widget.insert(tk.END, message)
        log_widget.yview(tk.END)

# Function to launch the GUI
def run_bom_processor_gui():
    root = tk.Tk()
    gui = BOMProcessorGUI(root)
    root.mainloop()

