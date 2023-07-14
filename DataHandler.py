import xlwings as xw
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Text, Button
import tabulate

global dataframe
global wb_to_use
dataframe = None
wb_to_use = None

def select_file():
    # Open a file selection dialog
    root.filename = filedialog.askopenfilename(initialdir="/", title="Select Excel File",
                                               filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
    file_entry.insert(tk.END, root.filename)
    print(root.filename)


def fetch_data():
    global dataframe
    global wb_to_use
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
        acct_no = row[columns.index("ACCT NO.")]
        sch_a_code = row[columns.index("sch A acc")]
        cn = str(row[columns.index("CN")])

        if acct_no and sch_a_code and cn != "ALL" \
                and str(sch_a_code).startswith(("11", "12")):
            filtered_data.append(row)

    # Create a DataFrame from the filtered data
    df = pd.DataFrame(filtered_data, columns=columns)
    dataframe = df

    # Convert the values in the 'sch A acc' column to strings
    df['sch A acc'] = df['sch A acc'].astype(str)

    # Calculate summary information
    summary_info = calculate_summary()

    # Display the data summary in the GUI
    display_summary(summary_info)

    # Update the log widget
    log_text.insert(tk.END, "Data fetched successfully.\n")

    print(df)

    paste_button.config(state=tk.NORMAL)

def calculate_summary():
    global dataframe

    # Calculate the number of rows
    num_rows = len(dataframe)

    # Calculate the sums of "CONV AMT" based on "sch A acc" code
    sums_by_sch_a_code = dataframe.groupby("sch A acc")["CONV  AMT"].sum().reset_index()

    # Store the summary information
    summary_info = {
        "Number of Rows": num_rows,
        "Sums by sch A acc": sums_by_sch_a_code
    }

    return summary_info

def display_summary(summary_info):
    # Display the data summary in the GUI
    data_text.delete("1.0", "end")
    data_text.insert("1.0", "Summary Information:\n\n")

    # Display the number of rows
    num_rows = summary_info["Number of Rows"]
    data_text.insert(tk.END, f"Number of Rows: {num_rows}\n\n")

    # Display the sums by sch A acc
    sums_by_sch_a_acc = summary_info["Sums by sch A acc"]
    sums_by_group = sums_by_sch_a_acc.groupby(sums_by_sch_a_acc["sch A acc"].str[:4])
    data_text.insert(tk.END, "Sums by Group:\n")
    for group, group_data in sums_by_group:
        conv_amt_sum = abs(group_data["CONV  AMT"].sum())
        data_text.insert(tk.END, f"Group: {group}\tCONV AMT: {conv_amt_sum:,.2f}\n")

    # Display the sums by sch A acc (2 digits)
    sums_by_group_2 = sums_by_sch_a_acc.groupby(sums_by_sch_a_acc["sch A acc"].str[:2])
    data_text.insert(tk.END, "\nSums by Group:\n")
    for group_2, group_data_2 in sums_by_group_2:
        conv_amt_sum_2 = abs(group_data_2["CONV  AMT"].sum())
        data_text.insert(tk.END, f"Group: {group_2}\tCONV AMT: {conv_amt_sum_2:,.2f}\n")

    # Display the sum of all data
    conv_amt_sum_total = abs(sums_by_sch_a_acc["CONV  AMT"].sum())
    data_text.insert(tk.END, f"\nTotal CONV AMT: {conv_amt_sum_total:,.2f}\n")

def paste_data():
    global dataframe
    global wb_to_use

    # Get the "ALL CP" sheet from the workbook
    cp_sheet = wb_to_use.sheets["DataExtraction"]

    # Find the last row with data in column A and clear the content from the sheet
    last_row_cp = cp_sheet.range("A2:C" + str(cp_sheet.cells.last_cell.row))
    last_row_cp.clear_contents()

    # Determine range where data should be pasted, running from A2 to column C with the rows being determined
    # by the amount of rows in the dataframe
    dest_range = cp_sheet.range("A2:C{}".format(len(dataframe) + 1))

    dest_range.value = dataframe[["ACCT NO.", "sch A acc","CONV  AMT"]].values

    # Update the log widget
    log_text.insert(tk.END, "Data pasted successfully.\n")

    print("Data pasted successfully.")

# Create the GUI
root = tk.Tk()

# File selection button and entry
select_button = Button(root, text="Select File", command=select_file)
select_button.pack()

file_entry = tk.Entry(root)
file_entry.pack()

# Text widget for displaying data
data_text = Text(root, width=80,height=20)
data_text.pack()

# Button to fetch data
fetch_button = Button(root, text="Fetch Data", command=fetch_data)
fetch_button.pack()

# Button to paste data
paste_button = Button(root, text="Paste Data in ALL CP tab", command=paste_data, state=tk.DISABLED)
paste_button.pack()

# Log widget
log_text = Text(root, height=5)
log_text.pack()

# Run the GUI
root.mainloop()
