import zipfile
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Create a Tkinter root window (it won't be shown)
root = Tk()
root.withdraw()

# Ask the user to select an Excel file using Windows Explorer
excel_file_path = askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])

# Check if the user selected a file
if excel_file_path:
    # Replace the extension with '.zip'
    zip_file_path = os.path.splitext(excel_file_path)[0] + '.zip'

    # Create a new zip file
    with zipfile.ZipFile(zip_file_path, 'w') as zip_file:
        # Add the Excel file to the zip file with the same name
        zip_file.write(excel_file_path, arcname=os.path.basename(excel_file_path))

    print(f"Excel file '{excel_file_path}' has been converted to '{zip_file_path}' successfully.")
else:
    print("No Excel file selected.")
