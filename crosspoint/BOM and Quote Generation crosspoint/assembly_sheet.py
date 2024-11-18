#Developed By Mohammed Afzal (18/11/2024)
#3rd option for Assembly sheet
import os
import pandas as pd
from tkinter import Tk, Label, Button, Entry, filedialog, messagebox
import tkinter as tk

# Function to open file dialog and select folder
def select_folder(entry):
    folder = filedialog.askdirectory()
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)

# Function to open file dialog and select a file
def select_file(entry):
    file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file:
        entry.delete(0, tk.END)
        entry.insert(0, file)

# Function to run the assembly sheet update process
def run_update(gui):
    input_folder = gui.input_folder_entry.get()
    output_folder = gui.output_folder_entry.get()
    external_file = gui.external_file_entry.get()

    if not input_folder or not output_folder or not external_file:
        messagebox.showerror("Input Error", "Please fill in all fields.")
        return

    # Call the update_assembly_sheet function
    try:
        update_assembly_sheet(input_folder, output_folder, gui, external_file)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while updating assembly sheets: {e}")

# Create the GUI class
class AssemblySheetGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Assembly Sheet Update")

        # Define default input and output folders as relative paths
        self.default_input_folder = "./input"
        self.default_output_folder = "./output"

        # Input Folder
        self.input_folder_label = Label(root, text="Input Folder")
        self.input_folder_label.grid(row=0, column=0, padx=10, pady=10)
        self.input_folder_entry = Entry(root, width=50)
        self.input_folder_entry.grid(row=0, column=1, padx=10, pady=10)
        self.input_folder_button = Button(root, text="Browse", command=lambda: select_folder(self.input_folder_entry))
        self.input_folder_button.grid(row=0, column=2, padx=10, pady=10)

        # Set default value for input folder
        self.input_folder_entry.insert(0, self.default_input_folder)

        # Output Folder
        self.output_folder_label = Label(root, text="Output Folder")
        self.output_folder_label.grid(row=1, column=0, padx=10, pady=10)
        self.output_folder_entry = Entry(root, width=50)
        self.output_folder_entry.grid(row=1, column=1, padx=10, pady=10)
        self.output_folder_button = Button(root, text="Browse", command=lambda: select_folder(self.output_folder_entry))
        self.output_folder_button.grid(row=1, column=2, padx=10, pady=10)

        # Set default value for output folder
        self.output_folder_entry.insert(0, self.default_output_folder)

        # External File (Components file)
        self.external_file_label = Label(root, text="External Components File")
        self.external_file_label.grid(row=2, column=0, padx=10, pady=10)
        self.external_file_entry = Entry(root, width=50)
        self.external_file_entry.grid(row=2, column=1, padx=10, pady=10)
        self.external_file_button = Button(root, text="Browse", command=lambda: select_file(self.external_file_entry))
        self.external_file_button.grid(row=2, column=2, padx=10, pady=10)

        # Status Label
        self.status_label = Label(root, text="", fg="blue")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=10)

        # Update Button
        self.update_button = Button(root, text="Update Assembly Sheets", command=lambda: run_update(self))
        self.update_button.grid(row=4, column=0, columnspan=3, pady=20)

    def update_status(self, text):
        self.status_label.config(text=text)

# Function to update assembly sheet
def update_assembly_sheet(input_folder, output_folder, gui, external_file):
    gui.update_status("Updating Assembly sheets...")
    input_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

    # Load external file for mapping
    external_df = pd.read_excel(external_file)
    mapping_df = external_df[["Manufacturer's Part Number", "Type", "Pin Count"]]

    # List to hold entries with unknown Manufacturer's Part Numbers
    unknown_part_numbers = []

    for file_name in input_files:
        file_path = os.path.join(input_folder, file_name)
        project_name = os.path.splitext(file_name)[0]
        try:
            df_assembly = pd.read_excel(file_path, sheet_name="Assembly Quote")

            # Replace empty Manufacturer's Part Number cells with "Unknown"
            df_assembly["Manufacturer's Part Number"].fillna("Unknown", inplace=True)

            # Flag entries with "Unknown" Manufacturer's Part Number and add to unknown list
            df_assembly['Status'] = df_assembly["Manufacturer's Part Number"].apply(lambda x: "Unknown" if x == "Unknown" else "Known")
            unknown_entries = df_assembly[df_assembly["Status"] == "Unknown"]
            unknown_part_numbers.extend(unknown_entries["Part Description"].tolist())

            # Merge 'df_assembly' with 'mapping_df' to get "Type" and "Pin Count"
            df_assembly = pd.merge(df_assembly, mapping_df, on="Manufacturer's Part Number", how="left")

            # Calculate "Per Board Count" as the product of "Quantity" and "Pin Count"
            df_assembly['Per Board Count'] = df_assembly['Quantity'] * df_assembly['Pin Count'].fillna(0)

            # Save the updated Assembly sheet
            df_assembly.to_excel(os.path.join(output_folder, f"{project_name}_updated_Assembly_sheet.xlsx"), index=False)

        except Exception as e:
            messagebox.showerror("Error", f"Error processing {file_name}: {e}")

    # Show message with the list of unknown part numbers, if any
    if unknown_part_numbers:
        messagebox.showinfo("Unknown Part Numbers", f"Entries with unknown Manufacturer's Part Numbers:\n" + "\n".join(unknown_part_numbers))
    
    messagebox.showinfo("Success", "Assembly sheets updated successfully!")
    gui.update_status("Update complete!")
    gui.root.quit()

# Main function to run the GUI
def main():
    root = Tk()
    gui = AssemblySheetGUI(root)
    root.mainloop()

"""if __name__ == "__main__":
    main()"""
#main()
