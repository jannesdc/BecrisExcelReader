import os
import xlwings as xw
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Text, Button

def select_file():
    # Open a file selection dialog
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select Excel File",
                                               filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
    file_entry.insert(tk.END, root.filename)
    print(root.filename)


def fetch_data():
    # Get the selected file path
    file_path = file_entry.get()

    # Check if the file is already open
    is_file_open = False
    for app in xw.apps:
        for wb in app.books:
            if wb.name in file_path:
                is_file_open = True
                wb_to_use = wb
                break
        if is_file_open:
            break

    # If the file is not open, open it
    if not is_file_open:
        wb_to_use = xw.books.open(file_path)

    # Get the "05-2022" sheet of the workbook
    source_sheet = wb_to_use.sheets['05-2022']

    # Find the last row with data in column A of the source sheet
    last_row = source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row

    # Read the data from the source sheet into a DataFrame
    header_range = source_sheet.range("A2:AK2")
    data_range = source_sheet.range("A3").expand("down").resize(last_row - 1, header_range.columns.count)

    columns = header_range.value
    data = data_range.value

    # Create an empty list to store the filtered data rows
    filtered_data = []

    # Iterate over the data rows and filter
    for row in data:
        if str(row[columns.index("sch A acc")]).startswith(('11', '12')) and row[columns.index("ACCT NO.")] != {None,0} and str(row[columns.index("CN")]) != {None,"ALL"}:
            filtered_data.append(row)

    # Create a DataFrame from the filtered data
    df = pd.DataFrame(filtered_data, columns=columns)

    # Convert the values in the 'sch A acc' column to strings
    df['sch A acc'] = df['sch A acc'].astype(str)

    # Display the data in the GUI
    data_text.delete("1.0", "end")
    data_text.insert("1.0", df.to_string())

    print(df)

# Create the GUI
root = tk.Tk()

# File selection button and entry
select_button = Button(root, text="Select File", command=select_file)
select_button.pack()

file_entry = tk.Entry(root)
file_entry.pack()

# Text widget for displaying data
data_text = Text(root)
data_text.pack()

# Button to fetch data
fetch_button = Button(root, text="Fetch Data", command=fetch_data)
fetch_button.pack()

# Run the GUI
root.mainloop()
