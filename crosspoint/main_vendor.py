"""import pandas as pd
import os

def get_vendor_columns(vendor_name):
    
    #Prompt the user for column numbers of necessary fields for each vendor file.
    
    print(f"\nPlease provide the column numbers for the file: {vendor_name}")
    part_desc_col = int(input("Enter the column number for Part Description: "))
    part_num_col = int(input("Enter the column number for Manufacturer's Part Number: "))
    req_qty_col = int(input("Enter the column number for Required Quantity: "))
    avail_qty_col = int(input("Enter the column number for Available Quantity: "))
    unit_price_col = int(input("Enter the column number for Unit Price: "))
    order_qty_col = int(input("Enter the column number for Order Quantity: "))
    
    return {
        'part_desc_col': part_desc_col,
        'part_num_col': part_num_col,
        'req_qty_col': req_qty_col,
        'avail_qty_col': avail_qty_col,
        'unit_price_col': unit_price_col,
        'order_qty_col': order_qty_col
    }

def process_vendor(df, columns, vendor_name):
    
    #Process a single vendor file to compute the final cost for each part.
    
    #for val in df['avail_qty_col']:
     #   print(val)
    with pd.option_context('display.max_rows', None, 'display.max_columns', None):
        print(df)
    final_list = []
    part_desc_col = columns['part_desc_col']
    part_num_col = columns['part_num_col']
    req_qty_col = columns['req_qty_col']
    avail_qty_col = columns['avail_qty_col']
    unit_price_col = columns['unit_price_col']
    order_qty_col = columns['order_qty_col']
    #print("Requirement qty")
    #print(columns['req_qty_col'])
    #print("Unit price")
    #print(columns['unit_price_col'])
    #print("order qty col")
    #print(columns['order_qty_col'])
    # Ensure all specified columns are numeric, replace non-numeric with 0
    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)
        print(f"Column {col} (after conversion):")
        print(df.iloc[:, col])  # Debug statement

    # Create a list to hold parts' final costs
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        print(f"\nProcessing Part Number: {part_num}")
        print(f"Required Quantity: {required_qty}")  # Debug statement

        available_vendors = group[group.iloc[:, avail_qty_col] >= required_qty]
        print(f"Available Vendors for {part_num} (after availability check): {len(available_vendors)}")
        print(available_vendors.head())  # Debug statement to see if any rows pass the filter

        if not available_vendors.empty:
            valid_vendors = []
            for idx, vendor in available_vendors.iterrows():
                final_order_qty = vendor.iloc[order_qty_col]
                unit_price = vendor.iloc[unit_price_col]
                
                # Calculate final price
                final_price = final_order_qty * unit_price
                print(f"Vendor: {vendor_name}, Final Order Qty: {final_order_qty}, Unit Price: {unit_price}, Final Price: {final_price}")  # Debugging statement
                
                vendor['Final Ordered Quantity'] = final_order_qty
                vendor['Final Price'] = final_price
                vendor['Source'] = vendor_name

                valid_vendors.append(vendor)

            if valid_vendors:
                selected_vendor = min(valid_vendors, key=lambda x: x['Final Price'])
                final_list.append({
                    "Manufacturer's Part Number": selected_vendor.iloc[part_num_col],
                    "Part Description": selected_vendor.iloc[part_desc_col],
                    "Required Quantity": required_qty,
                    "Final Ordered Quantity": selected_vendor['Final Ordered Quantity'],
                    "Unit Price": selected_vendor.iloc[unit_price_col],
                    "Final Price": selected_vendor['Final Price'],
                    "Source": selected_vendor['Source']
                })

    # Convert to DataFrame for final output
    result_df = pd.DataFrame(final_list)
    print("Processed DataFrame:", result_df.head())  # Final Debugging Statement
    return result_df


def main():
    # Get input folder path
    input_folder = input("Enter the address of the input folder containing all vendor Excel sheets: ")
    
    # Process each Excel file in the input folder
    data_frames = []
    for filename in os.listdir(input_folder):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_folder, filename)
            print(f"\n--- Processing Vendor File: {filename} ---")
            
            # Load the Excel file
            df = pd.read_excel(file_path)
            
            # Display the column structure for reference
            print("\nColumn No | Column Name | First Example")
            print("----------------------------------------")
            for col_no, col_name in enumerate(df.columns):
                example_value = df.iloc[0, col_no] if not df.empty else "N/A"
                print(f"{col_no:<10} | {col_name:<15} | {example_value}")
            
            # Get column mappings for this vendor
            columns = get_vendor_columns(filename)
            print(columns)
            # Process this vendor's data and calculate costs
            processed_df = process_vendor(df, columns, filename)
            data_frames.append(processed_df)
    
    # Combine data from all vendor files into a single DataFrame
    if data_frames:
        final_df = pd.concat(data_frames, ignore_index=True)
        print(final_df.head())
    
        # Group by part number using the column index, and select the lowest price across vendors
        grouped = final_df.groupby(final_df.columns[columns['part_num_col']]).apply(
            lambda x: x.loc[x['Final Price'].idxmin()]
        )

        print("\n--- Final Compiled Data ---")
        print(grouped)
        
        # Save to CSV
        output_csv = input("Enter the path to save the output as a CSV file (e.g., output.csv): ")
        grouped.to_csv(output_csv, index=False)
        print(f"Data saved to {output_csv}")
        
        # Save to Excel
        output_excel = input("Enter the path to save the output as an Excel file (e.g., output.xlsx): ")
        grouped.to_excel(output_excel, index=False)
        print(f"Data saved to {output_excel}")
    else:
        print("No vendor data to process.")

main()
###### only cheapest vendor details


import pandas as pd
import os

def get_vendor_columns(vendor_name):
    
    #Prompt the user for column numbers of necessary fields for each vendor file.
    
    print(f"\nPlease provide the column numbers for the file: {vendor_name}")
    part_desc_col = int(input("Enter the column number for Part Description: "))
    part_num_col = int(input("Enter the column number for Manufacturer's Part Number: "))
    req_qty_col = int(input("Enter the column number for Required Quantity: "))
    avail_qty_col = int(input("Enter the column number for Available Quantity: "))
    unit_price_col = int(input("Enter the column number for Unit Price: "))
    order_qty_col = int(input("Enter the column number for Order Quantity: "))
    
    return {
        'part_desc_col': part_desc_col,
        'part_num_col': part_num_col,
        'req_qty_col': req_qty_col,
        'avail_qty_col': avail_qty_col,
        'unit_price_col': unit_price_col,
        'order_qty_col': order_qty_col
    }

def process_vendor(df, columns, vendor_name):

    #Process a single vendor file to compute the final cost for each part and add vendor-specific columns.
    
    final_list = []
    part_desc_col = columns['part_desc_col']
    part_num_col = columns['part_num_col']
    req_qty_col = columns['req_qty_col']
    avail_qty_col = columns['avail_qty_col']
    unit_price_col = columns['unit_price_col']
    order_qty_col = columns['order_qty_col']

    # Ensure all specified columns are numeric, replace non-numeric with 0
    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    # Create a list to hold parts' final costs
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]

        available_vendors = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendors.empty:
            vendor_data = {"Manufacturer's Part Number": part_num, "Required Quantity": required_qty}

            for idx, vendor in available_vendors.iterrows():
                final_order_qty = vendor.iloc[order_qty_col]
                unit_price = vendor.iloc[unit_price_col]
                final_price = final_order_qty * unit_price
                
                vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
                vendor_data[f"{vendor_name} Price"] = final_price

            final_list.append(vendor_data)

    # Convert to DataFrame for final output
    result_df = pd.DataFrame(final_list)
    return result_df

def main():
    input_folder = input("Enter the address of the input folder containing all vendor Excel sheets: ")
    
    data_frames = []
    vendor_columns = {}

    # Process each Excel file in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_folder, filename)
            print(f"\n--- Processing Vendor File: {filename} ---")
            
            # Load the Excel file
            df = pd.read_excel(file_path)
            
            # Display the column structure for reference
            print("\nColumn No | Column Name | First Example")
            print("----------------------------------------")
            for col_no, col_name in enumerate(df.columns):
                example_value = df.iloc[0, col_no] if not df.empty else "N/A"
                print(f"{col_no:<10} | {col_name:<15} | {example_value}")
            
            # Get column mappings for this vendor
            columns = get_vendor_columns(filename)
            vendor_columns[filename] = columns

            # Process this vendor's data and calculate costs
            processed_df = process_vendor(df, columns, filename)
            data_frames.append(processed_df)
    
    # Combine data from all vendor files into a single DataFrame
    if data_frames:
        final_df = pd.concat(data_frames, axis=1)
        print(final_df.head())
        
        # Save to CSV
        output_csv = input("Enter the path to save the output as a CSV file (e.g., output.csv): ")
        final_df.to_csv(output_csv, index=False)
        print(f"Data saved to {output_csv}")
        
        # Save to Excel
        output_excel = input("Enter the path to save the output as an Excel file (e.g., output.xlsx): ")
        final_df.to_excel(output_excel, index=False)
        print(f"Data saved to {output_excel}")
    else:
        print("No vendor data to process.")

main()



# The below code will find which vendor will offer the least price for each component in bulk 

import pandas as pd
import os

def get_vendor_columns(vendor_name):
    print(f"\nPlease provide the column numbers for the file: {vendor_name}")
    part_desc_col = int(input("Enter the column number for Part Description: "))
    part_num_col = int(input("Enter the column number for Manufacturer's Part Number: "))
    req_qty_col = int(input("Enter the column number for Required Quantity: "))
    avail_qty_col = int(input("Enter the column number for Available Quantity: "))
    unit_price_col = int(input("Enter the column number for Unit Price: "))
    order_qty_col = int(input("Enter the column number for Order Quantity: "))
    
    return {
        'part_desc_col': part_desc_col,
        'part_num_col': part_num_col,
        'req_qty_col': req_qty_col,
        'avail_qty_col': avail_qty_col,
        'unit_price_col': unit_price_col,
        'order_qty_col': order_qty_col
    }

def process_vendor(df, columns, vendor_name):
    part_desc_col = columns['part_desc_col']
    part_num_col = columns['part_num_col']
    req_qty_col = columns['req_qty_col']
    avail_qty_col = columns['avail_qty_col']
    unit_price_col = columns['unit_price_col']
    order_qty_col = columns['order_qty_col']

    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    vendor_data_list = []
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        vendor_data = {
            "Manufacturer's Part Number": part_num,
            "Required Quantity": required_qty
        }
        
        available_vendor = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendor.empty:
            vendor = available_vendor.iloc[0]  # First available row if multiple rows exist
            final_order_qty = vendor.iloc[order_qty_col]
            unit_price = vendor.iloc[unit_price_col]
            final_price = final_order_qty * unit_price
            
            vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
            vendor_data[f"{vendor_name} Price"] = final_price
        else:
            vendor_data[f"{vendor_name} Ordered Quantity"] = None
            vendor_data[f"{vendor_name} Price"] = None

        vendor_data_list.append(vendor_data)

    result_df = pd.DataFrame(vendor_data_list)
    return result_df

def main():
    input_folder = input("Enter the address of the input folder containing all vendor Excel sheets: ")
    vendor_dfs = []
    vendor_columns = {}

    for filename in os.listdir(input_folder):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(input_folder, filename)
            print(f"\n--- Processing Vendor File: {filename} ---")
            
            df = pd.read_excel(file_path)
            print("\nColumn No | Column Name | First Example")
            print("----------------------------------------")
            for col_no, col_name in enumerate(df.columns):
                example_value = df.iloc[0, col_no] if not df.empty else "N/A"
                print(f"{col_no:<10} | {col_name:<15} | {example_value}")
            
            columns = get_vendor_columns(filename)
            vendor_columns[filename] = columns

            processed_df = process_vendor(df, columns, filename)
            vendor_dfs.append(processed_df)
    
    # Initialize combined_df with the first vendor DataFrame
    combined_df = vendor_dfs[0]
    
    # Merge each subsequent vendor DataFrame
    for vendor_df in vendor_dfs[1:]:
        combined_df = pd.merge(combined_df, vendor_df, on="Manufacturer's Part Number", how="outer")

    # Compute final ordered quantity, price, and sour
    # ce based on the minimum price
    final_order_list = []
    for _, row in combined_df.iterrows():
        row_data = row.to_dict()
        required_qty = row_data.get("Required Quantity")
        min_price = float('inf')
        final_vendor = None
        final_order_qty = None

        for vendor in vendor_columns.keys():
            vendor_price_col = f"{vendor} Price"
            vendor_qty_col = f"{vendor} Ordered Quantity"
            if vendor_price_col in row_data and not pd.isna(row_data[vendor_price_col]):
                if row_data[vendor_price_col] < min_price:
                    min_price = row_data[vendor_price_col]
                    final_order_qty = row_data[vendor_qty_col]
                    final_vendor = vendor
        
        row_data["Final Ordered Quantity"] = final_order_qty
        row_data["Final Price"] = min_price if final_vendor else None
        row_data["Source"] = final_vendor

        final_order_list.append(row_data)

    final_df = pd.DataFrame(final_order_list)
    
    output_csv = input("Enter the path to save the output as a CSV file (e.g., output.csv): ")
    final_df.to_csv(output_csv, index=False)
    print(f"Data saved to {output_csv}")
    
    output_excel = input("Enter the path to save the output as an Excel file (e.g., output.xlsx): ")
    final_df.to_excel(output_excel, index=False)
    print(f"Data saved to {output_excel}")

main()

######working code that considers the entire filename as source 

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


def process_vendor(df, columns, vendor_name, text_widget):
    part_desc_col = columns['part_desc_col']
    part_num_col = columns['part_num_col']
    req_qty_col = columns['req_qty_col']
    avail_qty_col = columns['avail_qty_col']
    unit_price_col = columns['unit_price_col']
    order_qty_col = columns['order_qty_col']

    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    vendor_data_list = []
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        vendor_data = {
            "Manufacturer's Part Number": part_num,
            "Required Quantity": required_qty
        }

        available_vendor = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendor.empty:
            vendor = available_vendor.iloc[0]  # First available row if multiple rows exist
            final_order_qty = vendor.iloc[order_qty_col]
            unit_price = vendor.iloc[unit_price_col]
            final_price = final_order_qty * unit_price

            vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
            vendor_data[f"{vendor_name} Price"] = final_price
        else:
            vendor_data[f"{vendor_name} Ordered Quantity"] = None
            vendor_data[f"{vendor_name} Price"] = None

        vendor_data_list.append(vendor_data)

    result_df = pd.DataFrame(vendor_data_list)
    return result_df


def get_vendor_columns(vendor_name, df, root, text_widget):
    def confirm_columns():
        nonlocal columns_selected
        columns_selected = {
            'part_desc_col': int(part_desc_col.get()),
            'part_num_col': int(part_num_col.get()),
            'req_qty_col': int(req_qty_col.get()),
            'avail_qty_col': int(avail_qty_col.get()),
            'unit_price_col': int(unit_price_col.get()),
            'order_qty_col': int(order_qty_col.get())
        }
        col_dialog.destroy()

    col_dialog = tk.Toplevel(root)
    col_dialog.title(f"Select Columns for {vendor_name}")

    columns_selected = {}

    # Display the first few rows of the dataframe
    text_widget.insert(tk.END, f"\nPlease provide the column numbers for {vendor_name}\n")
    for col_no, col_name in enumerate(df.columns):
        example_value = df.iloc[0, col_no] if not df.empty else "N/A"
        text_widget.insert(tk.END, f"{col_no}: {col_name} (Example: {example_value})\n")

    tk.Label(col_dialog, text="Part Description Column:").grid(row=0, column=0)
    part_desc_col = tk.Spinbox(col_dialog, from_=0, to=len(df.columns) - 1, width=5)
    part_desc_col.grid(row=0, column=1)

    tk.Label(col_dialog, text="Manufacturer's Part Number Column:").grid(row=1, column=0)
    part_num_col = tk.Spinbox(col_dialog, from_=0, to=len(df.columns) - 1, width=5)
    part_num_col.grid(row=1, column=1)

    tk.Label(col_dialog, text="Required Quantity Column:").grid(row=2, column=0)
    req_qty_col = tk.Spinbox(col_dialog, from_=0, to=len(df.columns) - 1, width=5)
    req_qty_col.grid(row=2, column=1)

    tk.Label(col_dialog, text="Available Quantity Column:").grid(row=3, column=0)
    avail_qty_col = tk.Spinbox(col_dialog, from_=0, to=len(df.columns) - 1, width=5)
    avail_qty_col.grid(row=3, column=1)

    tk.Label(col_dialog, text="Unit Price Column:").grid(row=4, column=0)
    unit_price_col = tk.Spinbox(col_dialog, from_=0, to=len(df.columns) - 1, width=5)
    unit_price_col.grid(row=4, column=1)

    tk.Label(col_dialog, text="Order Quantity Column:").grid(row=5, column=0)
    order_qty_col = tk.Spinbox(col_dialog, from_=0, to=len(df.columns) - 1, width=5)
    order_qty_col.grid(row=5, column=1)

    confirm_button = tk.Button(col_dialog, text="Confirm", command=confirm_columns)
    confirm_button.grid(row=6, columnspan=2)

    root.wait_window(col_dialog)

    return columns_selected


def main_gui():
    global root
    root = tk.Tk()
    root.title("Vendor Data Processor")

    # Scrolled text for displaying messages to the user
    text_widget = ScrolledText(root, width=100, height=25)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    # Function for selecting the input folder
    def select_input_folder():
        folder_path = filedialog.askdirectory()
        input_folder_var.set(folder_path)
        text_widget.insert(tk.END, f"\nSelected Input Folder: {folder_path}\n")

    # Function for selecting the output file
    def select_output_file():
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        output_file_var.set(file_path)
        text_widget.insert(tk.END, f"\nOutput will be saved to: {file_path}\n")

    # Function to run the processing
    def process_files():
        input_folder = input_folder_var.get()
        output_file = output_file_var.get()
        vendor_dfs = []
        vendor_columns = {}

        for filename in os.listdir(input_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

                df = pd.read_excel(file_path)
                columns = get_vendor_columns(filename, df, root, text_widget)
                vendor_columns[filename] = columns

                processed_df = process_vendor(df, columns, filename, text_widget)
                vendor_dfs.append(processed_df)

        # Merge each vendor DataFrame
        combined_df = vendor_dfs[0]
        for vendor_df in vendor_dfs[1:]:
            combined_df = pd.merge(combined_df, vendor_df, on="Manufacturer's Part Number", how="outer")

        # Compute final ordered quantity, price, and source based on the minimum price
        final_order_list = []
        for _, row in combined_df.iterrows():
            row_data = row.to_dict()
            required_qty = row_data.get("Required Quantity")
            min_price = float('inf')
            final_vendor = None
            final_order_qty = None

            for vendor in vendor_columns.keys():
                vendor_price_col = f"{vendor} Price"
                vendor_qty_col = f"{vendor} Ordered Quantity"
                if vendor_price_col in row_data and not pd.isna(row_data[vendor_price_col]):
                    if row_data[vendor_price_col] < min_price:
                        min_price = row_data[vendor_price_col]
                        final_order_qty = row_data[vendor_qty_col]
                        final_vendor = vendor

            row_data["Final Ordered Quantity"] = final_order_qty
            row_data["Final Price"] = min_price if final_vendor else None
            row_data["Source"] = final_vendor

            final_order_list.append(row_data)

        final_df = pd.DataFrame(final_order_list)

        # Save to file
        if output_file.endswith('.csv'):
            final_df.to_csv(output_file, index=False)
        else:
            final_df.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"Data saved to {output_file}")

    # GUI Widgets
    input_folder_var = tk.StringVar()
    output_file_var = tk.StringVar()

    tk.Button(root, text="Select Input Folder", command=select_input_folder).grid(row=1, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=1, column=1, padx=10)

    tk.Button(root, text="Select Output File", command=select_output_file).grid(row=2, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=output_file_var, width=50).grid(row=2, column=1, padx=10)

    tk.Button(root, text="Process Files", command=process_files).grid(row=3, column=0, columnspan=2, padx=10, pady=20)

    root.mainloop()


#main_gui()


############################### working wothout default directory
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


def process_vendor(df, columns, vendor_name, text_widget):
    # Extract the column indices dynamically based on the mapping
    part_desc_col = columns['Part Description']
    part_num_col = columns['Manufacturer Part Number']
    req_qty_col = columns['Required Quantity']
    avail_qty_col = columns['Availability']
    unit_price_col = columns['Unit Price']
    order_qty_col = columns['Order Quantity']

    # Convert the relevant columns to numeric values where needed
    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    vendor_data_list = []
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        vendor_data = {
            "Manufacturer's Part Number": part_num,
            "Required Quantity": required_qty
        }

        # Find the vendor data with availability >= required quantity
        available_vendor = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendor.empty:
            vendor = available_vendor.iloc[0]  # First available row if multiple rows exist
            final_order_qty = vendor.iloc[order_qty_col]
            unit_price = vendor.iloc[unit_price_col]
            final_price = final_order_qty * unit_price

            vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
            vendor_data[f"{vendor_name} Price"] = final_price
        else:
            vendor_data[f"{vendor_name} Ordered Quantity"] = None
            vendor_data[f"{vendor_name} Price"] = None

        vendor_data_list.append(vendor_data)

    result_df = pd.DataFrame(vendor_data_list)
    return result_df


def get_vendor_columns(vendor_name, column_mapping_df):

    Retrieve the column numbers for a given vendor using the template Excel sheet.
    
    # Fetch the column numbers for the current vendor from the template sheet
    vendor_columns = column_mapping_df[vendor_name].to_dict()
    
    # Return the column mapping for this vendor
    return vendor_columns


def main_gui():
    global root
    root = tk.Tk()
    root.title("Vendor Data Processor")

    # Scrolled text for displaying messages to the user
    text_widget = ScrolledText(root, width=100, height=25)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    # Function for selecting the input folder
    def select_input_folder():
        folder_path = filedialog.askdirectory()
        input_folder_var.set(folder_path)
        text_widget.insert(tk.END, f"\nSelected Input Folder: {folder_path}\n")

    # Function for selecting the output file
    def select_output_file():
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
        output_file_var.set(file_path)
        text_widget.insert(tk.END, f"\nOutput will be saved to: {file_path}\n")

    # Function for selecting the template file
    def select_template_file():
        template_file_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        template_file_var.set(template_file_path)
        text_widget.insert(tk.END, f"\nSelected Template File: {template_file_path}\n")

    # Function to run the processing
    def process_files():
        input_folder = input_folder_var.get()
        output_file = output_file_var.get()
        template_file = template_file_var.get()

        # Load the template file that contains column mappings
        column_mapping_df = pd.read_excel(template_file, index_col=0)

        vendor_dfs = []

        for filename in os.listdir(input_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                vendor_name = filename.split('_')[0]  # Extract the vendor name from the filename (before first '_')
                text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

                # Get the column mapping for the current vendor
                columns = get_vendor_columns(vendor_name, column_mapping_df)

                # Read vendor file
                df = pd.read_excel(file_path)

                # Process the vendor data
                processed_df = process_vendor(df, columns, vendor_name, text_widget)
                vendor_dfs.append(processed_df)

        # Merge each vendor DataFrame
        combined_df = vendor_dfs[0]
        for vendor_df in vendor_dfs[1:]:
            combined_df = pd.merge(combined_df, vendor_df, on="Manufacturer's Part Number", how="outer")

        # Compute final ordered quantity, price, and source based on the minimum price
        final_order_list = []
        for _, row in combined_df.iterrows():
            row_data = row.to_dict()
            required_qty = row_data.get("Required Quantity")
            min_price = float('inf')
            final_vendor = None
            final_order_qty = None

            for vendor in column_mapping_df.columns:
                vendor_price_col = f"{vendor} Price"
                vendor_qty_col = f"{vendor} Ordered Quantity"
                if vendor_price_col in row_data and not pd.isna(row_data[vendor_price_col]):
                    if row_data[vendor_price_col] < min_price:
                        min_price = row_data[vendor_price_col]
                        final_order_qty = row_data[vendor_qty_col]
                        final_vendor = vendor

            # Extract vendor source (filename up to the first underscore)
            if final_vendor:
                source_name = final_vendor.split('_')[0]  # Get part before the first underscore
                row_data["Final Ordered Quantity"] = final_order_qty
                row_data["Final Price"] = min_price if final_vendor else None
                row_data["Source"] = source_name  # Use the part before the first underscore in the vendor filename

            final_order_list.append(row_data)

        final_df = pd.DataFrame(final_order_list)

        # Save to file
        if output_file.endswith('.csv'):
            final_df.to_csv(output_file, index=False)
        else:
            final_df.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"Data saved to {output_file}")

    # GUI Widgets
    input_folder_var = tk.StringVar()
    output_file_var = tk.StringVar()
    template_file_var = tk.StringVar()

    tk.Button(root, text="Select Input Folder", command=select_input_folder).grid(row=1, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=1, column=1, padx=10)

    tk.Button(root, text="Select Output File", command=select_output_file).grid(row=2, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=output_file_var, width=50).grid(row=2, column=1, padx=10)

    tk.Button(root, text="Select Template File", command=select_template_file).grid(row=3, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=template_file_var, width=50).grid(row=3, column=1, padx=10)

    tk.Button(root, text="Process Files", command=process_files).grid(row=4, column=0, columnspan=2, padx=10, pady=20)

    root.mainloop()

    
    #working alone but not integrating 

#main_gui()
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Checkbutton, IntVar
from tkinter.scrolledtext import ScrolledText

def process_vendor(df, columns, vendor_name, text_widget):
    part_desc_col = columns['Part Description']
    part_num_col = columns['Manufacturer Part Number']
    req_qty_col = columns['Required Quantity']
    avail_qty_col = columns['Availability']
    unit_price_col = columns['Unit Price']
    order_qty_col = columns['Order Quantity']

    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    vendor_data_list = []
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        vendor_data = {
            "Manufacturer's Part Number": part_num,
            "Required Quantity": required_qty
        }

        available_vendor = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendor.empty:
            vendor = available_vendor.iloc[0]
            final_order_qty = vendor.iloc[order_qty_col]
            unit_price = vendor.iloc[unit_price_col]
            final_price = final_order_qty * unit_price

            vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
            vendor_data[f"{vendor_name} Price"] = final_price
        else:
            vendor_data[f"{vendor_name} Ordered Quantity"] = None
            vendor_data[f"{vendor_name} Price"] = None

        vendor_data_list.append(vendor_data)

    result_df = pd.DataFrame(vendor_data_list)
    return result_df


def get_vendor_columns(vendor_name, column_mapping_df):
    vendor_columns = column_mapping_df[vendor_name].to_dict()
    return vendor_columns


def main_gui():
    global root
    root = tk.Tk()
    root.title("Vendor Data Processor")

    current_directory = os.path.dirname(os.path.abspath(__file__))
    print(current_directory)
    text_widget = ScrolledText(root, width=100, height=25)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    def select_input_folder():
        folder_path = filedialog.askdirectory(initialdir=current_directory)
        input_folder_var.set(folder_path)
        text_widget.insert(tk.END, f"\nSelected Input Folder: {folder_path}\n")

    def select_output_file():
        file_path = filedialog.asksaveasfilename(
            initialdir=current_directory,
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        )
        output_file_var.set(file_path)
        text_widget.insert(tk.END, f"\nOutput will be saved to: {file_path}\n")

    def select_template_file():
        template_file_path = filedialog.askopenfilename(
            initialdir=current_directory,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        template_file_var.set(template_file_path)
        text_widget.insert(tk.END, f"\nSelected Template File: {template_file_path}\n")

    def load_vendor_checkboxes():
        template_file = template_file_var.get()
        column_mapping_df = pd.read_excel(template_file, index_col=0)
        vendor_names = column_mapping_df.columns

        for i, vendor in enumerate(vendor_names):
            var = IntVar()
            vendor_var_dict[vendor] = var
            Checkbutton(root, text=vendor, variable=vendor_var_dict[vendor]).grid(row=5 + i, column=0, sticky='w')

    def process_files():
        input_folder = input_folder_var.get()
        output_file = output_file_var.get()
        template_file = template_file_var.get()

        column_mapping_df = pd.read_excel(template_file, index_col=0)
        vendor_dfs = []

        for filename in os.listdir(input_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                vendor_name = filename.split('_')[0]
                text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

                columns = get_vendor_columns(vendor_name, column_mapping_df)
                df = pd.read_excel(file_path)
                processed_df = process_vendor(df, columns, vendor_name, text_widget)
                vendor_dfs.append(processed_df)

        combined_df = vendor_dfs[0]
        for vendor_df in vendor_dfs[1:]:
            combined_df = pd.merge(combined_df, vendor_df, on="Manufacturer's Part Number", how="outer")

        selected_vendors = [vendor for vendor, var in vendor_var_dict.items() if var.get() == 1]
        final_order_list = []

        for _, row in combined_df.iterrows():
            row_data = row.to_dict()
            required_qty = row_data.get("Required Quantity")
            min_price = float('inf')
            final_vendor = None
            final_order_qty = None

            for vendor in selected_vendors:
                vendor_price_col = f"{vendor} Price"
                vendor_qty_col = f"{vendor} Ordered Quantity"
                if vendor_price_col in row_data and not pd.isna(row_data[vendor_price_col]):
                    if row_data[vendor_price_col] < min_price:
                        min_price = row_data[vendor_price_col]
                        final_order_qty = row_data[vendor_qty_col]
                        final_vendor = vendor

            if final_vendor:
                source_name = final_vendor.split('_')[0]
                row_data["Final Ordered Quantity"] = final_order_qty
                row_data["Final Price"] = min_price if final_vendor else None
                row_data["Source"] = source_name

            final_order_list.append(row_data)

        final_df = pd.DataFrame(final_order_list)

        if output_file.endswith('.csv'):
            final_df.to_csv(output_file, index=False)
        else:
            final_df.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"Data saved to {output_file}")

    input_folder_var = tk.StringVar()
    output_file_var = tk.StringVar()
    template_file_var = tk.StringVar()
    vendor_var_dict = {}

    tk.Button(root, text="Select Input Folder", command=select_input_folder).grid(row=1, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=1, column=1, padx=10)

    tk.Button(root, text="Select Output File", command=select_output_file).grid(row=2, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=output_file_var, width=50).grid(row=2, column=1, padx=10)

    tk.Button(root, text="Select Template File", command=select_template_file).grid(row=3, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=template_file_var, width=50).grid(row=3, column=1, padx=10)

    tk.Button(root, text="Load Vendor Checkboxes", command=load_vendor_checkboxes).grid(row=4, column=0, columnspan=2, padx=10, pady=5)
    tk.Button(root, text="Process Files", command=process_files).grid(row=4, column=2, columnspan=2, padx=10, pady=20)

    root.mainloop()

# Uncomment the following line to run the GUI
#main_gui()

#################################################### code with options for selecting vendor and working
"""


"""import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Checkbutton, IntVar
from tkinter.scrolledtext import ScrolledText

def process_vendor(df, columns, vendor_name, text_widget):
    part_desc_col = columns['Part Description']
    part_num_col = columns['Manufacturer Part Number']
    req_qty_col = columns['Required Quantity']
    avail_qty_col = columns['Availability']
    unit_price_col = columns['Unit Price']
    order_qty_col = columns['Order Quantity']

    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    vendor_data_list = []
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        vendor_data = {
            "Manufacturer's Part Number": part_num,
            "Required Quantity": required_qty
        }

        available_vendor = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendor.empty:
            vendor = available_vendor.iloc[0]
            final_order_qty = vendor.iloc[order_qty_col]
            unit_price = vendor.iloc[unit_price_col]
            final_price = final_order_qty * unit_price

            vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
            vendor_data[f"{vendor_name} Price"] = final_price
        else:
            vendor_data[f"{vendor_name} Ordered Quantity"] = None
            vendor_data[f"{vendor_name} Price"] = None

        vendor_data_list.append(vendor_data)

    result_df = pd.DataFrame(vendor_data_list)
    return result_df


def get_vendor_columns(vendor_name, column_mapping_df):
    vendor_columns = column_mapping_df[vendor_name].to_dict()
    return vendor_columns


def main_gui():
    global root
    root = tk.Tk()
    root.title("Customer Quote Creater")

    text_widget = ScrolledText(root, width=100, height=25)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    input_folder_var = tk.StringVar()
    output_file_var = tk.StringVar()
    template_file_var = tk.StringVar()
    vendor_var_dict = {}

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

        column_mapping_df = pd.read_excel(template_file, index_col=0)
        vendor_names = column_mapping_df.columns

        for i, vendor in enumerate(vendor_names):
            var = IntVar(root)  # Associate IntVar with the root window
            vendor_var_dict[vendor] = var
            Checkbutton(root, text=vendor, variable=vendor_var_dict[vendor]).grid(row=5 + i, column=0, sticky='w')


    def process_files():
        input_folder = input_folder_var.get()
        output_file = output_file_var.get()
        template_file = template_file_var.get()

        column_mapping_df = pd.read_excel(template_file, index_col=0)
        vendor_dfs = []

        # Processing vendor files
        for filename in os.listdir(input_folder):
            if filename.startswith('~$') or not filename.endswith('.xlsx'):
                continue
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                vendor_name = filename.split('_')[0]
                text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

                columns = get_vendor_columns(vendor_name, column_mapping_df)
                df = pd.read_excel(file_path)
                processed_df = process_vendor(df, columns, vendor_name, text_widget)
                vendor_dfs.append(processed_df)

        # Start merging vendor data frames
        combined_df = vendor_dfs[0]
        for vendor_df in vendor_dfs[1:]:
            combined_df = pd.merge(
                combined_df,
                vendor_df,
                on="Manufacturer's Part Number",
                how="outer",
                suffixes=('', '_duplicate')  # Prevent appending suffixes for common columns
            )

        # Remove duplicate "Required Quantity" columns
        if 'Required Quantity_duplicate' in combined_df.columns:
            combined_df.drop(columns=['Required Quantity_duplicate'], inplace=True)

        # Ensure selected vendors are populated
        selected_vendors = [vendor for vendor, var in vendor_var_dict.items() if var.get() == 1]
        text_widget.insert(tk.END, f"\nSelected vendors: {selected_vendors}\n")

        final_order_list = []
        for _, row in combined_df.iterrows():
            row_data = row.to_dict()
            required_qty = row_data.get("Required Quantity")
            min_price = float('inf')
            final_vendor = None
            final_order_qty = None

            # Process selected vendor columns
            for vendor in selected_vendors:
                vendor_price_col = f"{vendor} Price"
                vendor_qty_col = f"{vendor} Ordered Quantity"
                if vendor_price_col in row_data and not pd.isna(row_data[vendor_price_col]):
                    if row_data[vendor_price_col] < min_price:
                        min_price = row_data[vendor_price_col]
                        final_order_qty = row_data[vendor_qty_col]
                        final_vendor = vendor

            # Add final price and source
            if final_vendor:
                source_name = final_vendor.split('_')[0]
                row_data["Final Ordered Quantity"] = final_order_qty
                row_data["Final Price"] = min_price
                row_data["Source"] = source_name

            final_order_list.append(row_data)

        final_df = pd.DataFrame(final_order_list)

        # Ensure the output has the required columns
        if 'Final Price' not in final_df.columns:
            text_widget.insert(tk.END, f"\n'Final Price' column is missing from the final dataframe.\n")
        if 'Source' not in final_df.columns:
            text_widget.insert(tk.END, f"\n'Source' column is missing from the final dataframe.\n")

        # Save the result to file
        if output_file.endswith('.xlsx'):
            final_df.to_csv(output_file, index=False)
        else:
            final_df.to_excel(output_file, index=False)

        messagebox.showinfo("Success", f"Data saved to {output_file}")

   

    tk.Button(root, text="Select Input Folder", command=select_input_folder).grid(row=1, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=1, column=1, padx=10)

    tk.Button(root, text="Select Output File", command=select_output_file).grid(row=2, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=output_file_var, width=50).grid(row=2, column=1, padx=10)

    tk.Button(root, text="Select Template File", command=select_template_file).grid(row=3, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=template_file_var, width=50).grid(row=3, column=1, padx=10)

    tk.Button(root, text="Load Vendor Checkboxes", command=load_vendor_checkboxes).grid(row=4, column=0, columnspan=2, padx=10, pady=5)
    tk.Button(root, text="Process Files", command=process_files).grid(row=4, column=2, columnspan=2, padx=10, pady=20)

    root.mainloop()
"""

# Uncomment the following line to run the GUI
#main_gui()


"""

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Checkbutton, IntVar
from tkinter.scrolledtext import ScrolledText
from datetime import datetime

# Function to get the current directory of the script
def get_current_dir():
    return os.path.dirname(os.path.abspath(__file__))

# Initialize input and output folders within the script directory
script_dir = get_current_dir()
input_folder_path = os.path.join(script_dir, 'input')
output_folder_path = os.path.join(script_dir, 'output')

# Create input and output directories if they don't exist
os.makedirs(input_folder_path, exist_ok=True)
os.makedirs(output_folder_path, exist_ok=True)

def process_vendor(df, columns, vendor_name, text_widget):
    part_desc_col = columns['Part Description']
    part_num_col = columns['Manufacturer Part Number']
    req_qty_col = columns['Required Quantity']
    avail_qty_col = columns['Availability']
    unit_price_col = columns['Unit Price']
    order_qty_col = columns['Order Quantity']

    numeric_columns = [req_qty_col, avail_qty_col, unit_price_col, order_qty_col]
    for col in numeric_columns:
        df.iloc[:, col] = pd.to_numeric(df.iloc[:, col].replace({'\$': '', ',': ''}, regex=True), errors='coerce').fillna(0)

    vendor_data_list = []
    for part_num, group in df.groupby(df.columns[part_num_col]):
        required_qty = group.iloc[0, req_qty_col]
        vendor_data = {
            "Manufacturer's Part Number": part_num,
            "Required Quantity": required_qty
        }

        available_vendor = group[group.iloc[:, avail_qty_col] >= required_qty]

        if not available_vendor.empty:
            vendor = available_vendor.iloc[0]
            final_order_qty = vendor.iloc[order_qty_col]
            unit_price = vendor.iloc[unit_price_col]
            final_price = final_order_qty * unit_price

            vendor_data[f"{vendor_name} Ordered Quantity"] = final_order_qty
            vendor_data[f"{vendor_name} Price"] = final_price
        else:
            vendor_data[f"{vendor_name} Ordered Quantity"] = None
            vendor_data[f"{vendor_name} Price"] = None

        vendor_data_list.append(vendor_data)

    result_df = pd.DataFrame(vendor_data_list)
    return result_df

def get_vendor_columns(vendor_name, column_mapping_df):
    vendor_columns = column_mapping_df[vendor_name].to_dict()
    return vendor_columns

def main_gui():
    global root
    root = tk.Tk()
    root.title("Vendor Data Processor")

    text_widget = ScrolledText(root, width=100, height=25)
    text_widget.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    input_folder_var = tk.StringVar(value=input_folder_path)  # Default input folder
    output_file_var = tk.StringVar(value=output_folder_path)  # Default output folder
    template_file_var = tk.StringVar()
    vendor_var_dict = {}

    def select_input_folder():
        folder_path = filedialog.askdirectory(initialdir=input_folder_var.get())
        if folder_path:
            input_folder_var.set(folder_path)
            text_widget.insert(tk.END, f"\nSelected Input Folder: {folder_path}\n")

    def select_output_file():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"Quote_{timestamp}.xlsx"
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=default_filename,
            initialdir=output_file_var.get()
        )
        if file_path:
            output_file_var.set(file_path)
            text_widget.insert(tk.END, f"\nOutput will be saved to: {file_path}\n")

    def select_template_file():
        file_path = filedialog.askopenfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialdir=template_file_var.get() or script_dir
        )
        if file_path:
            template_file_var.set(file_path)
            text_widget.insert(tk.END, f"\nSelected Template File: {file_path}\n")

    def load_vendor_checkboxes():
        template_file = template_file_var.get()
        if not template_file:
            messagebox.showerror("Error", "Please select a template file first.")
            return

        column_mapping_df = pd.read_excel(template_file, index_col=0)
        vendor_names = column_mapping_df.columns

        for i, vendor in enumerate(vendor_names):
            var = IntVar(root)
            vendor_var_dict[vendor] = var
            Checkbutton(root, text=vendor, variable=vendor_var_dict[vendor]).grid(row=5 + i, column=0, sticky='w')

    def process_files():
        input_folder = input_folder_var.get()
        output_file = output_file_var.get().strip()
        template_file = template_file_var.get()

        # Set default output filename if none is specified
        if not output_file:
            output_file = os.path.join(output_folder_path, "output.xlsx")
            output_file_var.set(output_file)

        # Ensure the output file has a .xlsx extension
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'
            output_file_var.set(output_file)

        if not template_file:
            messagebox.showerror("Error", "Please select a template file.")
            return

        column_mapping_df = pd.read_excel(template_file, index_col=0)
        vendor_dfs = []

        for filename in os.listdir(input_folder):
            if filename.endswith('.xlsx'):
                file_path = os.path.join(input_folder, filename)
                vendor_name = filename.split('_')[0]
                text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

                columns = get_vendor_columns(vendor_name, column_mapping_df)
                df = pd.read_excel(file_path)
                processed_df = process_vendor(df, columns, vendor_name, text_widget)
                vendor_dfs.append(processed_df)

        combined_df = vendor_dfs[0]
        for vendor_df in vendor_dfs[1:]:
            combined_df = pd.merge(combined_df, vendor_df, on="Manufacturer's Part Number", how="outer")

        selected_vendors = [vendor for vendor, var in vendor_var_dict.items() if var.get() == 1]
        text_widget.insert(tk.END, f"\nSelected vendors: {selected_vendors}\n")

        final_order_list = []
        for _, row in combined_df.iterrows():
            row_data = row.to_dict()
            required_qty = row_data.get("Required Quantity")
            min_price = float('inf')
            final_vendor = None
            final_order_qty = None

            for vendor in selected_vendors:
                vendor_price_col = f"{vendor} Price"
                vendor_qty_col = f"{vendor} Ordered Quantity"
                if vendor_price_col in row_data and not pd.isna(row_data[vendor_price_col]):
                    if row_data[vendor_price_col] < min_price:
                        min_price = row_data[vendor_price_col]
                        final_order_qty = row_data[vendor_qty_col]
                        final_vendor = vendor

            if final_vendor:
                source_name = final_vendor.split('_')[0]
                row_data["Final Ordered Quantity"] = final_order_qty
                row_data["Final Price"] = min_price
                row_data["Source"] = source_name

            final_order_list.append(row_data)

        final_df = pd.DataFrame(final_order_list)

        try:
            final_df.to_excel(output_file, index=False)  # Save as an Excel file
            messagebox.showinfo("Success", f"Data saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {e}")



    tk.Button(root, text="Select Input Folder", command=select_input_folder).grid(row=1, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=1, column=1, padx=10)

    tk.Button(root, text="Select Output File", command=select_output_file).grid(row=2, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=output_file_var, width=50).grid(row=2, column=1, padx=10)

    tk.Button(root, text="Select Template File", command=select_template_file).grid(row=3, column=0, padx=10, pady=5)
    tk.Entry(root, textvariable=template_file_var, width=50).grid(row=3, column=1, padx=10)

    tk.Button(root, text="Load Vendor Checkboxes", command=load_vendor_checkboxes).grid(row=4, column=0, columnspan=2, padx=10, pady=5)
    tk.Button(root, text="Process Files", command=process_files).grid(row=4, column=2, columnspan=2, padx=10, pady=20)

    root.mainloop()

# Uncomment the following line to run the GUI
#main_gui()
"""
"""import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Checkbutton, IntVar
from tkinter.scrolledtext import ScrolledText
import numpy as np

def process_vendor(df, column_indices, vendor_name, text_widget):
    #Process the vendor's data using the column mappings from the template.
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

    # Check if the vendor columns already exist, and if not, add them
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
        
        # Calculate total price for the vendor
        total_price = unit_price * order_qty if order_qty else None

        # For Digikey, calculate the total price separately if needed
        if vendor_name.lower() == "digikey" and order_qty is not None:
            total_price = unit_price * order_qty

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

    # Merge the vendor-specific data back into the main dataframe only if the columns don't exist
    for col in vendor_df.columns:
        if col not in df.columns:
            df[col] = vendor_df[col]
        else:
            df[col] = df[col].combine_first(vendor_df[col])  # Combine only non-null values

    # Ensure that part number and description are always included in the final output
    df["Manufacturer's Part Number"] = vendor_df["Manufacturer's Part Number"]
    df["Part Description"] = vendor_df["Part Description"]

    return df

def update_column_mapping(column_mapping_df, vendor_name, columns):
    
    column_mapping_df[vendor_name] = columns
    return column_mapping_df

def main_gui():
    global root
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
            if filename.startswith('~$') or not filename.endswith('.xlsx'):
                continue
            file_path = os.path.join(input_folder, filename)
            vendor_name = filename.split('_')[0]
            text_widget.insert(tk.END, f"\n--- Processing Vendor File: {filename} ---\n")

            try:
                # Get the column indices for the vendor from the template
                columns = column_mapping_df[vendor_name].to_dict()
            except KeyError as e:
                messagebox.showerror("Error", f"Vendor '{vendor_name}' not found in the template: {e}")
                continue

            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                messagebox.showerror("Error", f"Error reading file {filename}: {e}")
                continue
            
            processed_df = process_vendor(df, columns, vendor_name, text_widget)

            # Ensure we align based on row index (keeping the same number of rows for each vendor)
            combined_df = pd.concat([combined_df, processed_df[[f"{vendor_name} Availability", f"{vendor_name} Unit Price", f"{vendor_name} Ordered Quantity", f"{vendor_name} Total Price", "Manufacturer's Part Number", "Part Description"]]], axis=1)

        # Remove duplicate columns (except for the part number and description)
        combined_df = combined_df.loc[:, ~combined_df.columns.duplicated()]

        # Filter only the selected vendors
        selected_vendors = [vendor for vendor, var in vendor_var_dict.items() if var.get() == 1]

        # Debugging output
        print("Selected Vendors:", selected_vendors)
        
        # Compare prices for each row and select the lowest price from selected vendors
        combined_df['Final Price'] = np.nan
        combined_df['Selected Vendor'] = ''
        
        # Iterate over the rows and find the vendor with the minimum price
        for index, row in combined_df.iterrows():
            min_price = np.inf
            selected_vendor = ''
            
            for vendor in selected_vendors:
                price_col = f"{vendor} Total Price"
                
                # Check if the column exists before trying to access it
                if price_col in row and pd.notna(row[price_col]):
                    if row[price_col] < min_price:
                        min_price = row[price_col]
                        selected_vendor = vendor
                else:
                    print(f"Warning: Column '{price_col}' not found for {vendor} at row {index}")

            combined_df.at[index, 'Final Price'] = min_price
            combined_df.at[index, 'Selected Vendor'] = selected_vendor

        # Calculate the total final price in INR using the conversion factor
        combined_df['Final Price (INR)'] = combined_df['Final Price'] * conversion_factor

        # Reorder columns to move 'Manufacturer's Part Number' and 'Part Description' to the front
        final_columns = ["Manufacturer's Part Number", "Part Description", "Final Price (INR)"] + [col for col in combined_df.columns if col not in ["Manufacturer's Part Number", "Part Description", "Final Price (INR)"]]
        combined_df = combined_df[final_columns]

        # Preview the data
        preview_data(combined_df)

        # Save to output file
        try:
            combined_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Combined data saved to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error saving the output file: {e}")

    # Create UI elements for the conversion factor input
    tk.Label(root, text="Enter Conversion Factor (USD to INR):").grid(row=3, column=0, padx=10, pady=10)
    tk.Entry(root, textvariable=conversion_factor_var).grid(row=3, column=1, padx=10, pady=10)

    # Create buttons and UI elements
    tk.Button(root, text="Select Input Folder", command=select_input_folder).grid(row=1, column=0, padx=10, pady=10)
    tk.Button(root, text="Select Output File", command=select_output_file).grid(row=1, column=1, padx=10, pady=10)
    tk.Button(root, text="Select Template File", command=select_template_file).grid(row=1, column=2, padx=10, pady=10)
    tk.Button(root, text="Load Vendor Checkboxes", command=load_vendor_checkboxes).grid(row=1, column=3, padx=10, pady=10)
    tk.Button(root, text="Process Files", command=process_files).grid(row=2, column=0, columnspan=4, padx=10, pady=10)

    root.mainloop()


"""
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

