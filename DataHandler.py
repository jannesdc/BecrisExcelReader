import xlwings as xw
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Text, Button

global dataframe
global wb_to_use
dataframe = None
wb_to_use = None

def select_file():
    try:
        # Open a file selection dialog
        root.filename = filedialog.askopenfilename(initialdir="/", title="Select Excel File",
                                                   filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        file_entry.insert(tk.END, root.filename)
        print(root.filename)
        fetch_button.config(state=tk.NORMAL)
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

def fetch_data():
    global dataframe
    global wb_to_use
    try:
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
            wb_to_use = xw.Book(file_path, update_links=False)

        # Get the "05-2022" sheet of the workbook
        source_sheet = wb_to_use.sheets['CurrentMonth']

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

        # Make the paste data and check for new instruments buttons available
        paste_button.config(state=tk.NORMAL)
        update_new_button.config(state=tk.NORMAL)
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

def calculate_summary():
    global dataframe

    try:
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
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

def display_summary(summary_info):
    try:
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
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

def paste_data():
    global dataframe
    global wb_to_use

    try:
        # Get the "DataExtraction" sheet from the workbook
        extraction_sheet = wb_to_use.sheets["DataExtraction"]

        # Find the last row with data in column A and clear the content from the sheet
        last_row_cp = extraction_sheet.range("A2:C" + str(extraction_sheet.cells.last_cell.row))
        last_row_cp.clear_contents()

        # Determine range where data should be pasted, running from A2 to column C with the rows being determined
        # by the amount of rows in the dataframe
        dest_range = extraction_sheet.range("A2:C{}".format(len(dataframe) + 1))

        dest_range.value = dataframe[["ACCT NO.", "sch A acc","CONV  AMT"]].values

        # Update the log widget
        log_text.insert(tk.END, "Data pasted successfully.\n")
        wb_to_use.save()
        print("Data pasted successfully.")
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

def check_new():
    global wb_to_use
    global dataframe

    try:
        # Get the "ALL CP" sheet
        cp_sheet = wb_to_use.sheets("ALL CP")

        # create new_instruments list
        new_instruments = []

        # Check the "ALL CP" sheet for the existing instruments
        existing_instr_range = cp_sheet.range("B2:B" + str(cp_sheet.range("A" + str(cp_sheet.cells.last_cell.row)).end("up").row))
        existing_instruments = existing_instr_range.value

        # Check for any new instruments comparing the existing instruments listed in the "ALL CP" sheet
        # comparing it with the instruments in the dataframe
        for index, row in dataframe.iterrows():
            acct_no = str(row["ACCT NO."])
            if acct_no not in map(str,existing_instruments):
                new_instruments.append(row)

        # Paste the new instruments data at the end of the existing list in the "ALL CP" sheet
        if new_instruments:
            # Get the range to paste the new instruments data
            last_row = cp_sheet.range("A" + str(cp_sheet.cells.last_cell.row)).end("up").row
            if cp_sheet.range("A" + str(last_row)).value is None:
                paste_range = cp_sheet.range("A" + str(last_row))
            else:
                paste_range = cp_sheet.range("A" + str(last_row + 1))

            # Iterate over the new instruments and paste the data
            for i, instrument in enumerate(new_instruments):
                # Calculate the row index for pasting
                paste_row = paste_range.row + i

                # Paste the data in the respective columns
                cp_sheet.range("A" + str(paste_row)).value = "New"
                cp_sheet.range("B" + str(paste_row)).value = instrument[2]
                cp_sheet.range("C" + str(paste_row)).value = instrument[3]
                cp_sheet.range("E" + str(paste_row)).value = instrument[5]
                cp_sheet.range("F" + str(paste_row)).value = instrument[23]
                if int(instrument[12]) == 18:
                    cp_sheet.range("G" + str(paste_row)).value = 20
                elif int(instrument[12]) == 12:
                    cp_sheet.range("G" + str(paste_row)).value = 1000
                elif int(instrument[12]) in {20,21,22}:
                    cp_sheet.range("G" + str(paste_row)).value = 1004
                else:
                    cp_sheet.range("G" + str(paste_row)).value = "NOT FOUND"
                    cp_sheet.range("G" + str(paste_row)).color = 6



        # Update the log widget
        if new_instruments:
            log_text.insert(tk.END, f"{len(new_instruments)} new instruments pasted successfully.\n")
        else:
            log_text.insert(tk.END, "No new instruments found.\n")
        log_text.insert(tk.END, "New instruments checked successfully.\n")
        for instrument in new_instruments:
            print(instrument[3])
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")

# Create the GUI
root = tk.Tk()
root.title("Data Handler")

# File selection button and entry
select_button = Button(root, text="Select file", command=select_file)
select_button.pack()

file_entry = tk.Entry(root, width=50)
file_entry.pack()

# Text widget for displaying data
data_text = Text(root, width=80, height=20)
data_text.pack()

# Button to fetch data
fetch_button = Button(root, text="Fetch data", command=fetch_data, state=tk.DISABLED)
fetch_button.pack()

# Button to paste data
paste_button = Button(root, text="Paste data in ""DataExtraction"" tab", command=paste_data, state=tk.DISABLED)
paste_button.pack()

update_new_button = Button(root, text="Check for new instruments", command=check_new, state=tk.DISABLED)
update_new_button.pack()

# Log widget
log_text = Text(root, height=5)
log_text.pack()

# Run the GUI
root.mainloop()
