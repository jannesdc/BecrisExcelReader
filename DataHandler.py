import xlwings as xw
import pandas as pd
from tkinter import Tk, Label, Text, Button


def fetch_data():
    # Open the Excel file to fetch data from
    source_file_path = "BECRIS.xlsx"  # Specify the source file name
    wb = xw.Book(source_file_path)

    # Get the "04-2022" sheet of the workbook
    source_sheet = wb.sheets['04-2022']

    # Find the last row with data in column A of the source sheet
    last_row = source_sheet.range("A" + str(source_sheet.cells.last_cell.row)).end("up").row

    # Read the data from the source sheet into a DataFrame
    header_range = source_sheet.range("A1").expand("right")
    data_range = source_sheet.range("A2").expand("down").resize(last_row - 1, header_range.columns.count)

    columns = header_range.value
    data = data_range.value

    # Create an empty list to store the filtered data rows
    filtered_data = []

    # Iterate over the data rows and filter based on the condition for 'schAcode' column
    for row in data:
        if str(row[1]).startswith(('11', '12')):
            filtered_data.append(row)

    # Create a DataFrame from the filtered data
    df = pd.DataFrame(filtered_data, columns=columns)

    # Convert the values in the 'schAcode' column to strings
    df['schAcode'] = df['schAcode'].astype(str)

    # Get the "Data" sheet of the workbook
    target_sheet = wb.sheets['Data']

    # Find the last row with data in column A of the target sheet
    last_row = target_sheet.range("A" + str(target_sheet.cells.last_cell.row)).end("up").row

    # Write the header row to the "Data" sheet
    header_range = target_sheet.range("A1").expand("right")
    header_range.value = columns

    # Write the data to the "Data" sheet starting from the first empty row
    row_offset = last_row + 1
    data_range = target_sheet.range((row_offset, 1), (row_offset + len(df), len(columns) - 1))
    data_range.value = df.values

    # Display the data in the GUI
    data_text.delete("1.0", "end")
    data_text.insert("1.0", df.to_string())

# Create the GUI
root = Tk()

# Label for data display
data_label = Label(root, text="Data:")
data_label.pack()

# Text widget for displaying data
data_text = Text(root)
data_text.pack()

# Button to fetch data
fetch_button = Button(root, text="Fetch Data", command=fetch_data)
fetch_button.pack()

# Run the GUI
root.mainloop()
