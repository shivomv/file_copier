import os
import shutil
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox, StringVar, OptionMenu
import openpyxl

# Function to copy files based on the selected file extension and log missing folders
def copy_files_and_log(source_folder, destination_folder, file_type):
    extensions = {
        "All": "*",
        "PDF": [".pdf"],
        "Images": [".jpg", ".jpeg", ".png", ".gif"],
        "InDesign": [".indd"],  # Added InDesign file extension
        "Other": []
    }

    # Get the list of extensions based on the selected file type
    selected_extensions = extensions[file_type]
    
    # Create a new Excel workbook and sheet
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "File Copy Log"
    ws.append(["Copied Files", "Folders Without Files"])

    non_matching_folders = []
    copied_files = []

    # Traverse the directory structure
    for root, dirs, files in os.walk(source_folder):
        # Skip the parent folder and only evaluate subfolders
        if root == source_folder:
            continue
        
        folder_has_files = False  # Track if the folder contains files of the selected extension
        for file in files:
            if file_type == "All" or file.lower().endswith(tuple(selected_extensions)):
                folder_has_files = True
                source_file = os.path.join(root, file)
                shutil.copy(source_file, destination_folder)
                copied_files.append(file)  # Log copied file

        if not folder_has_files:  # If folder has no matching files, log it
            non_matching_folders.append(root)

    # Write copied files and non-matching folders to Excel
    max_len = max(len(copied_files), len(non_matching_folders))
    for i in range(max_len):
        copied_file = copied_files[i] if i < len(copied_files) else ""
        non_matching_folder = non_matching_folders[i] if i < len(non_matching_folders) else ""
        ws.append([copied_file, non_matching_folder])

    # Save the Excel workbook
    wb.save(os.path.join(destination_folder, "file_copy_log.xlsx"))

    messagebox.showinfo("Success", f"All {file_type.lower()} files have been copied and log generated in Excel!")

# Function to open folder dialog for source folder
def browse_source_folder():
    folder_selected = filedialog.askdirectory()
    source_entry.delete(0, 'end')
    source_entry.insert(0, folder_selected)

# Function to open folder dialog for destination folder
def browse_destination_folder():
    folder_selected = filedialog.askdirectory()
    destination_entry.delete(0, 'end')
    destination_entry.insert(0, folder_selected)

# Function to handle the copy process when the button is clicked
def on_submit():
    source_folder = source_entry.get()
    destination_folder = destination_entry.get()
    file_type = file_type_var.get()

    if not source_folder or not destination_folder:
        messagebox.showerror("Error", "Both folder paths are required!")
        return

    copy_files_and_log(source_folder, destination_folder, file_type)

# Create the main window
root = Tk()
root.title("File Copier")

# Add label and input for source folder
Label(root, text="Source Folder:").grid(row=0, column=0, padx=10, pady=10)
source_entry = Entry(root, width=50)
source_entry.grid(row=0, column=1, padx=10, pady=10)
Button(root, text="Browse", command=browse_source_folder).grid(row=0, column=2, padx=10, pady=10)

# Add label and input for destination folder
Label(root, text="Destination Folder:").grid(row=1, column=0, padx=10, pady=10)
destination_entry = Entry(root, width=50)
destination_entry.grid(row=1, column=1, padx=10, pady=10)
Button(root, text="Browse", command=browse_destination_folder).grid(row=1, column=2, padx=10, pady=10)

# Add dropdown for file type selection
Label(root, text="File Type:").grid(row=2, column=0, padx=10, pady=10)
file_type_var = StringVar(root)
file_type_var.set("All")  # Default value

file_type_options = ["All", "PDF", "Images", "InDesign", "Other"]
OptionMenu(root, file_type_var, *file_type_options).grid(row=2, column=1, padx=10, pady=10)

# Add a submit button
Button(root, text="Submit", command=on_submit).grid(row=3, columnspan=3, pady=20)

# Run the Tkinter event loop
root.mainloop()
