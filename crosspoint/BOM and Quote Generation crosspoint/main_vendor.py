
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Checkbutton, IntVar
from tkinter.scrolledtext import ScrolledText
import numpy as np

def convert_xls_to_xlsx(xls_file_path):
    # Define the path for the new .xlsx file
    xlsx_file_path = xls_file_path.replace('.xls', '.xlsx')
    
    try:
        # Read the .xls file with pandas
        df = pd.read_excel(xls_file_path, engine='openpyxl')  # Use openpyxl engine for both
        # Save as .xlsx using pandas (which uses openpyxl for .xlsx format)
        df.to_excel(xlsx_file_path, index=False, engine='openpyxl')
        print(f"Converted {xls_file_path} to {xlsx_file_path}")
    except Exception as e:
        print(f"Error converting {xls_file_path}: {e}")

def process_vendor(df, column_indices, vendor_name, text_widget):
    # Extract column indices from the template mapping
    part_desc_col = column_indices['Part Description']
    part_num_col = column_indices["Manufacturer's Part Number"]
    req_qty_col = column_indices['Required Quantity']
    avail_qty_col = column_indices['Availability']
    unit_price_col = column_indices['Unit Price']
    order_qty_col = column_indices['Order Quantity']

    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]

    # Ensure numeric columns are properly converted
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    # Add vendor-specific columns if they don't exist
    if f"{vendor_name} Availability" not in df.columns:
        df[f"{vendor_name} Availability"] = None
    if f"{vendor_name} Unit Price" not in df.columns:
        df[f"{vendor_name} Unit Price"] = None
    if f"{vendor_name} Ordered Quantity" not in df.columns:
        df[f"{vendor_name} Ordered Quantity"] = None
    if f"{vendor_name} Total Price" not in df.columns:
        df[f"{vendor_name} Total Price"] = None

    vendor_data = []
    for _, row in df.iterrows():
        part_num = row.iloc[part_num_col]
        part_desc = row.iloc[part_desc_col]
        required_qty = row.iloc[req_qty_col]
        available_qty = row.iloc[avail_qty_col]
        unit_price = row.iloc[unit_price_col]
        order_qty = row.iloc[order_qty_col] if available_qty >= required_qty else None
        total_price = unit_price * order_qty if order_qty else None

        vendor_data.append({
            "Manufacturer's Part Number": part_num,
            "Part Description": part_desc,
            "Required Quantity": required_qty,
            f"{vendor_name} Availability": available_qty,
            f"{vendor_name} Unit Price": unit_price,
            f"{vendor_name} Ordered Quantity": order_qty,
            f"{vendor_name} Total Price": total_price
        })

    vendor_df = pd.DataFrame(vendor_data)

    # Merge the vendor-specific data back into the main dataframe
    for col in vendor_df.columns:
        if col not in df.columns:
            df[col] = vendor_df[col]
        else:
            df[col] = df[col].combine_first(vendor_df[col])  # Combine only non-null values

    return df

def main_gui():
    root = tk.Tk()
    root.title("Customer Quote Creator")

    text_widget = ScrolledText(root, width=100, height=25)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    input_folder_var = tk.StringVar()
    output_file_var = tk.StringVar()
    template_file_var = tk.StringVar()
    vendor_var_dict = {}
    conversion_factor_var = tk.DoubleVar()

    def select_input_folder():
        folder_path = filedialog.askdirectory()
        input_folder_var.set(folder_path)
        text_widget.insert(tk.END, f"\nSelected Input Folder: {folder_path}\n")

    def select_output_file():
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        output_file_var.set(file_path)
        text_widget.insert(tk.END, f"\nOutput will be saved to: {file_path}\n")

    def select_template_file():
        template_file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        template_file_var.set(template_file_path)
        text_widget.insert(tk.END, f"\nSelected Template File: {template_file_path}\n")

    def load_vendor_checkboxes():
        template_file = template_file_var.get()
        if not template_file:
            messagebox.showerror("Error", "Please select a template file first.")
            return

        try:
            column_mapping_df = pd.read_excel(template_file, index_col=0)
        except Exception as e:
            messagebox.showerror("Error", f"Error reading template file: {e}")
            return

        vendor_names = column_mapping_df.columns

        for i, vendor in enumerate(vendor_names):
            var = IntVar(root)
            vendor_var_dict[vendor] = var
            Checkbutton(root, text=vendor, variable=vendor_var_dict[vendor]).grid(row=5 + i, column=0, sticky='w')

    def preview_data(combined_df):
        # Show a preview of the combined data in the text widget
        text_widget.insert(tk.END, "\n--- Data Preview ---\n")
        text_widget.insert(tk.END, combined_df.head().to_string(index=False))

    def process_files():
        input_folder = input_folder_var.get()
        output_file = output_file_var.get()
        template_file = template_file_var.get()
        conversion_factor = conversion_factor_var.get()

        try:
            column_mapping_df = pd.read_excel(template_file, index_col=0)
        except Exception as e:
            messagebox.showerror("Error", f"Error reading template file: {e}")
            return

        # Initialize an empty dataframe to store results
        combined_df = pd.DataFrame()

        for filename in os.listdir(input_folder):
            if filename.startswith('~$') or not filename.lower().endswith(('.xls', '.xlsx')):
                continue
            file_path = os.path.join(input_folder, filename)
            vendor_name = filename.split('_')[0]
            text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

            # Convert .xls to .xlsx if needed
            if filename.lower().endswith('.xls'):
                convert_xls_to_xlsx(file_path)
                file_path = file_path.replace('.xls', '.xlsx')  # Update the file path to the new .xlsx file

            try:
                columns = column_mapping_df[vendor_name].to_dict()
            except KeyError as e:
                messagebox.showerror("Error", f"Vendor '{vendor_name}' not found in the template: {e}")
                continue

            try:
                df = pd.read_excel(file_path)  # Support for both .xls and .xlsx files
            except Exception as e:
                messagebox.showerror("Error", f"Error reading file {filename}: {e}")
                continue

            # Process the file using the vendor's data
            processed_df = process_vendor(df, columns, vendor_name, text_widget)

            # Use Mouser's part number and description for the combined file
            if vendor_name.lower() == 'mouser':
                combined_df["Manufacturer's Part Number"] = processed_df["Manufacturer's Part Number"]
                combined_df["Part Description"] = processed_df["Part Description"]

            # Add vendor-specific columns to the combined dataframe
            combined_df = pd.concat([combined_df, processed_df[[
                f"{vendor_name} Availability", f"{vendor_name} Unit Price",
                f"{vendor_name} Ordered Quantity", f"{vendor_name} Total Price"]]], axis=1)

        combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]

        # Filter only selected vendors
        selected_vendors = [vendor for vendor, var in vendor_var_dict.items() if var.get() == 1]

        combined_df['Final Price'] = np.nan
        combined_df['Selected Vendor'] = ''

        # Calculate final price and selected vendor
        for index, row in combined_df.iterrows():
            min_price = np.inf
            selected_vendor = ''
            for vendor in selected_vendors:
                price_col = f"{vendor} Total Price"
                if price_col in row and pd.notna(row[price_col]):
                    if row[price_col] < min_price:
                        min_price = row[price_col]
                        selected_vendor = vendor

            combined_df.at[index, 'Final Price'] = min_price
            combined_df.at[index, 'Selected Vendor'] = selected_vendor

        combined_df['Final Price (INR)'] = combined_df['Final Price'] * conversion_factor

        # Rearrange columns: ensure Part Number and Description are the first two columns
        cols = ['Manufacturer\'s Part Number', 'Part Description'] + [col for col in combined_df.columns if col not in ['Manufacturer\'s Part Number', 'Part Description']]
        combined_df = combined_df[cols]

        # Save the final combined data
        try:
            combined_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Data successfully saved to {output_file}")
            preview_data(combined_df)
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the combined file: {e}")

    def validate_and_run():
        if not input_folder_var.get() or not output_file_var.get() or not template_file_var.get():
            messagebox.showerror("Error", "Please select all required files and folder.")
            return
        process_files()

    tk.Label(root, text="Input Folder:").grid(row=1, column=0, sticky='e')
    tk.Entry(root, textvariable=input_folder_var, width=60).grid(row=1, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_input_folder).grid(row=1, column=2)

    tk.Label(root, text="Output File:").grid(row=2, column=0, sticky='e')
    tk.Entry(root, textvariable=output_file_var, width=60).grid(row=2, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_output_file).grid(row=2, column=2)

    tk.Label(root, text="Template File:").grid(row=3, column=0, sticky='e')
    tk.Entry(root, textvariable=template_file_var, width=60).grid(row=3, column=1, padx=10)
    tk.Button(root, text="Browse", command=select_template_file).grid(row=3, column=2)

    tk.Label(root, text="Conversion Factor (for INR):").grid(row=4, column=0, sticky='e')
    tk.Entry(root, textvariable=conversion_factor_var, width=60).grid(row=4, column=1, padx=10)

    tk.Button(root, text="Load Vendors", command=load_vendor_checkboxes).grid(row=5, column=1)

    tk.Button(root, text="Process Files", command=validate_and_run).grid(row=6, column=1, pady=10)

    root.mainloop()

