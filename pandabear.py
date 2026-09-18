import openpyxl
import pandas as pd

# Load workbook and worksheet
wb = openpyxl.load_workbook("C:/Users/beldredge246949/Downloads/Book1.xlsx")
ws = wb.active

data = []

#import tkinter as tk
# from tkinter import filedialog

# Hide the main Tkinter root window
# root = tk.Tk()
# root.withdraw()

# Open the file dialog and grab the selected file path
# file_path = filedialog.askopenfilename(
#     title="Select an Excel (xlsx) File",
#     filetypes=[("Excel Files", "*.xlsx)]
# )

# if file_path:
#     print(f"User selected: {file_path}")
# else:
#     print("Selection canceled.")


# file explorer opens and allows user to navigate to find the sheet which they wish to convert
#1 get amount of rows and columns from list based on fixed input as a starting
#2 double iterate through each name/row and every coloumn on the row before moving to the next one while storing 
#   pertinent info in accurate temp dictionary
#3 convert dictionary into adaquate excel sheet format for importation into the downloads folder


#notes - use regex on last name to connect email. Store a small database of 
# use live calendar to determine which sheets to convert based off the massive main excel sheet
# access to a local dictionary database of emails and first and last name values for easy referencing

# Loop through rows (skip header if row 1 is headers)
for row in ws.iter_rows(min_row=2, values_only=False):
  cell_value = row[0].value  # Assuming data is in the first column
  # Get fill color (hex code or theme)
  fill_color = row[3].fill.fgColor.value if row[0].fill else None

  data.append({"Value": cell_value, "Color": fill_color})

# Convert to a pandas DataFrame
df = pd.DataFrame(data)
print(df)
