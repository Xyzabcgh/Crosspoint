#Developed By Mohammed Afzal (18/11/2024)


import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, scrolledtext
import subprocess
import assembly_sheet
import main_vendor
import sys
import customer_quote
from datetime import datetime
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
# Function for Combined File Generator (Placeholder)
def call_script(script_name):
    if getattr(sys, 'frozen', False):  # Check if running in PyInstaller bundle
        script_path = os.path.join(sys._MEIPASS, script_name)
        subprocess.run([sys.executable, script_path])
    else:
        subprocess.run([sys.executable, script_name])
def ask_columns_dialog():
    # Create a Toplevel window for the column inputs
    col_dialog = tk.Toplevel()
    col_dialog.title("Enter Column Numbers")

    # Create labels and spinboxes for the columns
    labels = ["Part Description Column:", "Manufacturer's Part Number Column:", "Final Quantity Column:","All value:","Source:"]
    entries = []

    for i, label in enumerate(labels):
        tk.Label(col_dialog, text=label).grid(row=i, column=0)
        entry = tk.Spinbox(col_dialog, from_=0, to=100, width=5)
        entry.grid(row=i, column=1)
        entries.append(entry)

    # Variable to store the column numbers
    column_numbers = []

    # Function to confirm and retrieve the selected column numbers
    def confirm_columns():
        nonlocal column_numbers  # Allow modifying the outer variable
        column_numbers = [int(entry.get()) for entry in entries]
        col_dialog.destroy()

    # Confirm button to close the dialog and return the column values
    confirm_button = tk.Button(col_dialog, text="Confirm", command=confirm_columns)
    confirm_button.grid(row=len(labels), columnspan=2, pady=10)

    col_dialog.wait_window()  # Wait for the dialog to be closed

    return column_numbers  # Return the column 
# Combined BOM generator with all value and source MODIFY this code for that 
def combined_file_generator(input_folder, output_folder, gui):
    
    gui.show_content("Combined Quote Generator ", "Processing BOM files for combined mode...")
    combined_df = pd.DataFrame(columns=["Part Description", "Manufacturer's Part Number", "Final QTY", "Projects", "All value", "Source"])

    input_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

    for file_name in input_files:
        file_path = os.path.join(input_folder, file_name)
        project_name = os.path.splitext(file_name)[0]
        try:
            df_bom = pd.read_excel(file_path, sheet_name="BOM Quote")
            df_readme = pd.read_excel(file_path, sheet_name="Readme")

            gui.show_content(f"File: {file_name} (Readme)", df_readme.to_string(index=False))
            row_num = simpledialog.askinteger("Input", f"Enter row number for 'BOM Procurement QTY' in {file_name}:", minvalue=0)
            bom_procurement_qty_value = float(df_readme.iloc[row_num, 1])

            columns_info = "\n".join([f"Column {idx}: {col} - Sample: {df_bom.iloc[0, idx]}" for idx, col in enumerate(df_bom.columns)])
            gui.show_content(f"File: {file_name} (BOM Quote Columns)", columns_info)
            messagebox.showinfo("Columns Info", f"Columns in BOM Quote:\n{columns_info}")

            # Collect column numbers
            part_desc_col, mfg_part_num_col, qty_col, all_value_col, source_col = ask_columns_dialog()

            output_df = pd.DataFrame()
            output_df['Part Description'] = df_bom.iloc[:, part_desc_col]
            output_df["Manufacturer's Part Number"] = df_bom.iloc[:, mfg_part_num_col].fillna("Unknown")
            output_df["Quantity"] = df_bom.iloc[:, qty_col].astype(float)
            output_df["All value"] = df_bom.iloc[:, all_value_col]
            output_df["Source"] = df_bom.iloc[:, source_col]

            

            #output_df['Process Loss Code Factor'] = output_df['Part Description'].apply(get_process_loss_factor)
            output_df["Final QTY"] = output_df["Quantity"]
            output_df["Projects"] = project_name

           # output_df = output_df.drop(columns=['Process Loss Code Factor'])

            # Group by with "Unknown" handling for missing Manufacturer's Part Number
            output_df = output_df.groupby(['Part Description', "Manufacturer's Part Number"], as_index=False).apply(
                lambda group: pd.Series({
                    'Final QTY': group['Final QTY'].sum(),
                    'Projects': ', '.join(set(group['Projects'].astype(str))),
                    'All value': group['All value'].min() if group['All value'].notna().any() else None,
                    'Source': group.loc[group['All value'].idxmin(), 'Source'] if group['All value'].notna().any() else None
                })
            ).reset_index(drop=True)

            combined_df = pd.concat([combined_df, output_df], ignore_index=True)
            output_df = output_df.drop(columns=['All value', 'Source', 'Projects'])
        except Exception as e:
            messagebox.showerror("Error", f"Error processing {file_name}: {e}")

    # Final aggregation for combined DataFrame
    combined_df = combined_df.groupby(['Part Description', "Manufacturer's Part Number"], as_index=False).apply(
        lambda group: pd.Series({
            'Final QTY': group['Final QTY'].sum(),
            'Projects': ', '.join(set(group['Projects'].astype(str))),
            'All value': group['All value'].min() if group['All value'].notna().any() else None,
            'Source': group.loc[group['All value'].idxmin(), 'Source'] if group['All value'].notna().any() else None
        })
    ).reset_index(drop=True)

    combined_df.to_excel(f"{output_folder}/combined_BOM_{timestamp}.xlsx", index=False)
    messagebox.showinfo("Success", "Combined file generated successfully!")

# Function Customer BOM generator

def customer_file_generator(input_folder, output_folder, gui, process_loss_factors):
    # Add your BOM customer file generation logic here
    gui.show_content("Customer BOM Generator ", "Processing BOM files for customer mode...")
    input_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

    # Load external file for mapping (if needed)
    # external_df = pd.read_excel(external_file)
    # mapping_df = external_df[["Manufacturer's Part Number", "Type", "Pin Count"]]

    for file_name in input_files:
        file_path = os.path.join(input_folder, file_name)
        project_name = os.path.splitext(file_name)[0]
        try:
            # Read necessary sheets from the file
            df_bom = pd.read_excel(file_path, sheet_name="Orginal BOM")
            df_assembly = pd.read_excel(file_path, sheet_name="Assembly Quote")
            df_readme = pd.read_excel(file_path, sheet_name="Readme")
            
            # Display Readme content to the GUI
            gui.show_content(f"File: {file_name} (Readme)", df_readme.to_string(index=False))
            
            # Prompt user to input BOM Procurement QTY value
            row_num = simpledialog.askinteger("Input", f"Enter row number for 'BOM Procurement QTY' in {file_name}:", minvalue=0)
            bom_procurement_qty_value = float(df_readme.iloc[row_num, 1])

            # Show columns in BOM Quote for reference
            columns_info = "\n".join([f"Column {idx}: {col} - Sample: {df_bom.iloc[0, idx]}" for idx, col in enumerate(df_bom.columns)])
            gui.show_content(f"File: {file_name} (BOM Quote Columns)", columns_info)
            messagebox.showinfo("Columns Info", f"Columns in BOM Quote:\n{columns_info}")

            # User inputs for relevant column numbers
            part_desc_col = simpledialog.askinteger("Input", "Enter column number for 'Part Description':")
            mfg_part_num_col = simpledialog.askinteger("Input", "Enter column number for 'Manufacturer's Part Number' in BOM:")
            qty_col = simpledialog.askinteger("Input", "Enter column number for 'Quantity':")

            # Create the output dataframe with relevant columns
            output_df = pd.DataFrame()
            output_df['Part Description'] = df_bom.iloc[:, part_desc_col]
            output_df["Manufacturer's Part Number"] = df_bom.iloc[:, mfg_part_num_col]
            output_df["Quantity"] = df_bom.iloc[:, qty_col].astype(float)

            # Process loss factors
            def get_process_loss_factor(description):
                for code, factor in process_loss_factors.items():
                    if code in str(description):
                        return factor
                return 0

            # Add process loss code factor and calculate final quantity
            output_df['Process Loss Code Factor'] = output_df['Part Description'].apply(get_process_loss_factor)
            output_df["Final QTY"] = (output_df["Quantity"] * bom_procurement_qty_value) + output_df["Process Loss Code Factor"]
            output_df = output_df.drop(columns=['Process Loss Code Factor', 'Quantity'])

            # Ensure missing parts are included without filtering out those with missing values in Part Description or Manufacturer's Part Number
            output_df['Part Description'].fillna('Unknown', inplace=True)
            output_df["Manufacturer's Part Number"].fillna('Unknown', inplace=True)
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            output_df.to_excel(os.path.join(output_folder, f"{project_name} Customer BOM without grouping components {timestamp}.xlsx"), index=False)
            # Group by Part Description and Manufacturer's Part Number for final output
            output_df = output_df.groupby(['Part Description', "Manufacturer's Part Number"], as_index=False).apply(
                lambda group: pd.Series({'Final QTY': group['Final QTY'].sum()})).reset_index(drop=True)

            # Save the final output to an Excel file
              # Timestamp for file naming
            output_df.to_excel(os.path.join(output_folder, f"{project_name} Customer BOM after grouping components{timestamp}.xlsx"), index=False)

            # Save the updated Assembly sheet to an Excel file (if needed)
            # df_assembly.to_excel(os.path.join(output_folder, f"{project_name}_updated_Assembly_sheet.xlsx"), index=False)

        except Exception as e:
            messagebox.showerror("Error", f"Error processing {file_name}: {e}")

    messagebox.showinfo("Success", "Customer BOM files updated successfully!")
    gui.root.quit()


# Function to load process loss factors (Placeholder)
def load_process_loss_factors(file):
    df = pd.read_excel(file)
    return dict(zip(df['Code'], df['Loss Factor']))
# Function to get vendor-specific columns (Placeholder)



class BOMProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("BOM Processor")
        self.input_folder = "./input"
        self.output_folder = "./output"
        self.process_loss_file = ""
        self.external_file = ""  # Store path to the external file

        # Processing mode selection
        self.mode = tk.StringVar(value="customer")
        tk.Label(root, text="Select Processing Mode:", pady=5).pack()
        tk.Radiobutton(root, text="Customer BOM Generator", variable=self.mode, value="customer", command=self.update_mode).pack(anchor="w")
        tk.Radiobutton(root, text="Customer Quote Generator(Source List)", variable=self.mode, value="vendor", command=self.update_mode).pack(anchor="w")
        tk.Radiobutton(root, text="Update Assembly Sheet", variable=self.mode, value="Assembly", command=self.update_mode).pack(anchor="w")
        tk.Radiobutton(root, text="Combined BOM Generator with all value and source", variable=self.mode, value="combined", command=self.update_mode).pack(anchor="w")
        tk.Radiobutton(root, text="Customer Quote Generator based from combined source list", variable=self.mode, value="customer_quote", command=self.update_mode).pack(anchor="w")
        os.makedirs(self.input_folder, exist_ok=True)
        os.makedirs(self.output_folder, exist_ok=True)

        # Display default input and output directories to the user
        #tk.Label(root, text="Default Directories:", pady=5).pack()
        
        # Labels to display the selected input and output folders
      
        # Button to allow the user to change input and output folders
        
        # Input folder selection (initially disabled)
        tk.Label(root, text="Select Input Folder:", pady=5).pack()
        self.input_folder_btn = tk.Button(root, text="Choose Input Folder", command=self.select_input_folder)
        self.input_folder_btn.pack()

        # Output folder selection (initially disabled)
        tk.Label(root, text="Select Output Folder:", pady=5).pack()
        self.output_folder_btn = tk.Button(root, text="Choose Output Folder", command=self.select_output_folder)
        self.output_folder_btn.pack()

        # Process loss file selection
        tk.Label(root, text="Select Process Loss Code File:", pady=5).pack()
        self.process_loss_file_btn = tk.Button(root, text="Choose Process Loss File", command=self.select_process_loss_file)
        self.process_loss_file_btn.pack()

        # External file selection for mapping (Initially hidden)
        self.external_file_label = tk.Label(root, text="Select External File for Mapping:", pady=5)
        self.external_file_btn = tk.Button(root, text="Choose External File", command=self.select_external_file)

        # Display the Readme content in a scrolled text area
        tk.Label(root, text="Content Preview:", pady=5).pack()
        self.content_text = scrolledtext.ScrolledText(root, width=80, height=15)
        self.content_text.pack()

        # Button to start processing files
        self.process_button = tk.Button(root, text="Process BOM Files", command=self.run_process, state="disabled")
        self.process_button.pack(pady=10)

        # Initially update the mode (hide external file selection if combined mode is selected)
        self.update_mode()
    def update_mode(self):
        """Update the GUI based on the selected processing mode"""
        #if self.mode.get() == "customer":
            #self.external_file_label.pack()
            #self.external_file_btn.pack()
            #self.selected_external_file_label.pack()

        if self.mode.get() == "Assembly":
           self.start_Assembly()
        elif self.mode.get()=="combined":
            self.process_loss_file_btn.config(state="disabled")
        elif self.mode.get() == "vendor":
            # Hide the BOM-specific settings and show vendor processing settings
            #self.hide_bom_inputs()
            self.start_vendor_processing()
        elif self.mode.get()=="customer_quote":
            self.start_customer_quote()
        else:
            self.show_bom_inputs()

    def show_bom_inputs(self):
        """Show BOM-specific inputs for combined and customer modes."""
        self.input_folder_btn.pack()
        self.output_folder_btn.pack()
        self.process_loss_file_btn.pack()

    def hide_bom_inputs(self):
        """Hide BOM-specific inputs when switching to vendor mode."""
        self.input_folder_btn.pack_forget()
        self.output_folder_btn.pack_forget()
        self.process_loss_file_btn.pack_forget()

        # Call vendor processing setup here
        #VendorDataProcessor(self.root)
    def start_Assembly(self):
        #subprocess.Popen(['python', 'assembly_sheet.py'])
        #call_script("assembly_sheet.py")
        self.root.destroy() 
        assembly_sheet.main()
    def start_vendor_processing(self):
        """Start the vendor data processing GUI."""
        #call_script("main_vendor.py")
        self.root.destroy() 
        main_vendor.main_gui()
        #subprocess.Popen(['python', 'main_vendor.py'])
    def start_customer_quote(self):
        self.root.destroy() 
        customer_quote.run_bom_processor_gui()

    def select_input_folder(self):
        self.input_folder = filedialog.askdirectory(title="Select Input Folder")
        #self.selected_input_folder_label.config(text=f"Selected Input Folder: {self.input_folder}")
        #self.selected_input_folder_label.pack()
        self.update_process_button_state()
        

    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory(title="Select Output Folder")
        self.update_process_button_state()
        

    def select_process_loss_file(self):
        self.process_loss_file = filedialog.askopenfilename(title="Select Process Loss Code File", filetypes=[("Excel files", "*.xlsx")])
        self.update_process_button_state()

    def select_external_file(self):
        self.external_file = filedialog.askopenfilename(title="Select External File", filetypes=[("Excel files", "*.xlsx")])
        self.selected_external_file_label.config(text=f"Selected External File: {self.external_file}")
        self.selected_external_file_label.pack()
        self.update_process_button_state()
    
    def update_process_button_state(self):
        if (self.input_folder and self.output_folder and self.process_loss_file and  self.mode.get()=="customer") or ( self.input_folder and self.output_folder and self.mode.get()=="combined"):
            self.process_button.config(state="normal")
        else:
            self.process_button.config(state="disabled")

    def show_content(self, title, content):
        """Displays specified content in the text area with a title."""
        self.content_text.delete(1.0, tk.END)
        self.content_text.insert(tk.END, f"{title}\n")
        self.content_text.insert(tk.END, content)

    def run_process(self):
        if self.input_folder and self.output_folder and self.process_loss_file and  self.mode.get()=="customer":
            process_loss_factors = load_process_loss_factors(self.process_loss_file)
            print(process_loss_factors)
            customer_file_generator(self.input_folder, self.output_folder, self, process_loss_factors)
        elif  self.input_folder and self.output_folder and self.mode.get()=="combined":
            combined_file_generator(self.input_folder,self.output_folder,self)

            
        else:
            messagebox.showerror("Error", "Please select input folder, output folder, process loss code file, and external file (for customer mode).")

# Vendor Data Processing GUI

# Main application
if __name__ == "__main__":
    root = tk.Tk()
    app = BOMProcessorGUI(root)
    root.mainloop()
