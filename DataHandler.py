import xlwings as xw
import pandas as pd
import numpy as np
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, Text, Button, Frame

global status_list_dataframe
global wb_to_use
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
    global status_list_dataframe
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

        # Get the "CurrentMonth" sheet of the workbook
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
        status_list_dataframe = df

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
        validate_button.config(state=tk.NORMAL)
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")


def calculate_summary():
    global status_list_dataframe

    try:
        # Calculate the number of rows
        num_rows = len(status_list_dataframe)

        # Calculate the sums of "CONV AMT" based on "sch A acc" code
        sums_by_sch_a_code = status_list_dataframe.groupby("sch A acc")["CONV  AMT"].sum().reset_index()

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
    global status_list_dataframe
    global wb_to_use
    global update_progress_bar

    try:
        # Get the "DataExtraction" sheet from the workbook
        update_progress_bar.config(mode="indeterminate")
        update_progress_bar.start()
        extraction_sheet = wb_to_use.sheets["DataExtraction"]

        # Find the last row with data in column A and clear the content from the sheet
        last_row_cp = extraction_sheet.range("A2:L" + str(extraction_sheet.cells.last_cell.row))
        last_row_cp.clear_contents()

        dest_range = extraction_sheet.range("A2:G{}".format(len(status_list_dataframe) + 1))
        interest_rate_type_range = extraction_sheet.range("H2:H{}".format(len(status_list_dataframe) + 1))
        mat_date_range = extraction_sheet.range("I2:I{}".format(len(status_list_dataframe) + 1))
        interest_rate_range = extraction_sheet.range("J2:J{}".format(len(status_list_dataframe) + 1))
        interest_reset_date_range = extraction_sheet.range("K2:K{}".format(len(status_list_dataframe) + 1))
        undrawn_range = extraction_sheet.range("L2:L{}".format(len(status_list_dataframe) + 1))

        dest_range.value = status_list_dataframe[
            ["ACCT NO.", "sch A acc", "CONV  AMT", "CON RT", "AMT", "ACCD INTT", "CUR"]].values

        # Get the values from the dataframe
        mat_date_values = status_list_dataframe["MAT DT"].values
        interest_type_values = status_list_dataframe["FLOATING/ FIXED"].values
        interest_rate_values = status_list_dataframe["ROI"].values
        interest_reset_date_values = status_list_dataframe["LAST RESET"].values
        undrawn_values = status_list_dataframe["UNDRAWN"].values

        # Prepare the values for insertion
        mat_date_values_array = [[value if not (pd.isna(value) or np.isnat(value)) else "NotApplicable"] for value in
                                 mat_date_values]
        interest_type_values_array = [[value] if value else ["NotApplicable"] for value in interest_type_values]
        interest_rate_array = []
        undrawn_array = []
        for value, cif_value in zip(interest_rate_values, status_list_dataframe["CIF"]):
            if cif_value == 'Office Account':
                interest_rate_array.append(["NotApplicable"])
            elif value:
                interest_rate_array.append([abs(value / 100)])
            else:
                interest_rate_array.append([""])
        interest_reset_date_array = [[value] if value else ["NotApplicable"] for value in interest_reset_date_values]
        for value, cif_value in zip(undrawn_values, status_list_dataframe["CIF"]):
            if cif_value == "Office Account":
                undrawn_array.append(["NotApplicable"])
            elif value:
                undrawn_array.append([abs(value)])
            else:
                undrawn_array.append([0])

        # Insert the values into the range
        mat_date_range.value = mat_date_values_array
        interest_rate_type_range.value = interest_type_values_array
        interest_rate_range.value = interest_rate_array
        interest_reset_date_range.value = interest_reset_date_array
        undrawn_range.value = undrawn_array

        # Update the log widget
        update_progress_bar.stop()
        update_progress_bar.config(mode="determinate")
        log_text.insert(tk.END, "Data pasted successfully.\n")
        wb_to_use.save()
        print("Data pasted successfully.")

    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")


def check_new():
    global wb_to_use
    global status_list_dataframe
    global update_progress_bar

    try:
        # Get the "ALL CP" sheet
        cp_sheet = wb_to_use.sheets("ALL CP")

        # create new_instruments list
        new_instruments = []

        # Check the "ALL CP" sheet for the existing instruments
        existing_instr_range = cp_sheet.range(
            "B2:B" + str(cp_sheet.range("A" + str(cp_sheet.cells.last_cell.row)).end("up").row))
        existing_instruments = existing_instr_range.value

        # Check for any new instruments comparing the existing instruments listed in the "ALL CP" sheet
        # comparing it with the instruments in the dataframe
        for index, row in status_list_dataframe.iterrows():
            acct_no = str(row["ACCT NO."])
            if acct_no not in map(str, existing_instruments):
                new_instruments.append(row)

        # Start the progress bar
        total_iterations = len(new_instruments) * 2
        update_progress_bar.config(maximum=total_iterations)
        update_progress_bar.start()

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

                # Update progress bar
                update_progress_bar.step()
                root.update()

                # Paste the data in the respective columns
                cp_sheet.range("A" + str(paste_row)).value = "New"
                cp_sheet.range("B" + str(paste_row)).value = instrument[2]
                cp_sheet.range("C" + str(paste_row)).value = instrument[3]
                cp_sheet.range("E" + str(paste_row)).value = -instrument[5]
                cp_sheet.range("F" + str(paste_row)).value = instrument[23]
                if int(instrument[12]) == 18:
                    cp_sheet.range("G" + str(paste_row)).value = 20
                elif int(instrument[12]) == 12:
                    cp_sheet.range("G" + str(paste_row)).value = 1000
                elif int(instrument[12]) == 22:
                    cp_sheet.range("G" + str(paste_row)).value = 71
                elif int(instrument[12]) in {20, 21}:
                    cp_sheet.range("G" + str(paste_row)).value = 1004
                else:
                    cp_sheet.range("G" + str(paste_row)).value = "NOT FOUND"
                    cp_sheet.range("G" + str(paste_row)).color = 6

        # Stop the progress bar
        update_progress_bar.stop()

        # Update the log widget
        if new_instruments:
            log_text.insert(tk.END, f"{len(new_instruments)} new instruments pasted successfully.\n")
        else:
            log_text.insert(tk.END, "No new instruments found.\n")
        log_text.insert(tk.END, "New instruments checked successfully.\n")
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")


def validate_data():
    """
    Checks per column of the "Data" tab the if the data is possible
    :return:
    """
    global status_list_dataframe
    global wb_to_use

    # First we fill the becris and counterparty dataframes,
    # so they can be used later for validation.
    counterparty_sheet = wb_to_use.sheets['Counterparties references']
    becris_sheet = wb_to_use.sheets['Data']

    # Populating the counterparty dataframe
    counterparty_last_row = counterparty_sheet.range("A" + str(counterparty_sheet.cells.last_cell.row)).end("up").row
    counterparty_header_range = counterparty_sheet.range("A1").expand("right")
    counterparty_data_range = counterparty_sheet.range("A2").expand("down").resize(counterparty_last_row - 1,
                                                                                   counterparty_header_range.columns.count)

    counterparty_columns = counterparty_header_range.value
    counterparty_data = counterparty_data_range.value

    counterparty_dataframe = pd.DataFrame(counterparty_data, columns=counterparty_columns)

    # Populating the becris dataframe
    becris_last_row = becris_sheet.range("A" + str(becris_sheet.cells.last_cell.row)).end("up").row
    becris_header_range = becris_sheet.range("A1").expand("right")
    becris_data_range = becris_sheet.range("A2").expand("down").resize(becris_last_row - 1,
                                                                       becris_header_range.columns.count)

    becris_columns = becris_header_range.value
    becris_data = becris_data_range.value
    filtered_becris_data = []
    for row in becris_data:
        if row[0] is not None:
            filtered_becris_data.append(row)

    becris_dataframe = pd.DataFrame(filtered_becris_data, columns=becris_columns)

    # Clear and update the data_text box
    data_text.delete("1.0", tk.END)
    data_text.insert("1.0", "Becris and Counterparty data summary:\n\n")
    data_text.insert(tk.END, f"Counterparties found: {len(counterparty_dataframe)}\n")
    data_text.insert(tk.END, f"Instruments found in becris data: {len(becris_dataframe)}\n\n")

    def check_counterparty_identifier_uniqueness(counterparty_dataframe):
        """
        Checks if Counterparty identifiers (ENI, LEI, RACI) in the Counterparty reference dataset are unique.
        :param counterparty_dataframe: The DataFrame containing the Counterparty reference data.
        :return: A boolean indicating whether the identifiers are unique or not.
        """

        # Create a new DataFrame to store the validity check results
        columns = ["Identifier", "Duplicate Count"]
        validity_result = []

        # Check ENI uniqueness, ignoring "NotRequired" values
        eni_unique = (counterparty_dataframe["ENI"] != "NotRequired") \
                     & counterparty_dataframe.duplicated(subset=["ENI"], keep=False)
        eni_duplicates = eni_unique.sum()
        validity_result.append(["ENI", eni_duplicates])

        # Check LEI uniqueness, ignoring "NotApplicable" values
        lei_unique = (counterparty_dataframe["LEI"] != "NotApplicable") \
                     & counterparty_dataframe.duplicated(subset=["LEI"], keep=False)
        lei_duplicates = lei_unique.sum()
        validity_result.append(["LEI", lei_duplicates])

        # Check RACI for empty values (should not be empty)
        raci_unique = counterparty_dataframe.duplicated(subset=["RACI"], keep=False)
        raci_duplicates = raci_unique.sum()
        raci_empty = counterparty_dataframe["RACI"].isnull().sum()
        validity_result.append(["RACI", raci_duplicates])

        validity_result_df = pd.DataFrame(validity_result, columns=columns)
        is_unique = pd.to_numeric(validity_result_df["Duplicate Count"], errors="coerce").sum() == 0

        if is_unique and raci_empty == 0:
            data_text.insert(tk.END, "Counterparty identifiers are UNIQUE.\n")
        else:
            data_text.insert(tk.END, "Counterparty identifiers are NOT CORRECT.\n")
            if not is_unique and raci_empty == 0:
                for idx, row in validity_result_df.iterrows():
                    identifier = row["Identifier"]
                    duplicate_count = row["Duplicate Count"]
                    data_text.insert(tk.END, f"{identifier} - Duplicate Count: {duplicate_count}\n")
            elif raci_empty != 0 and is_unique:
                data_text.insert(tk.END, f"RACI - Missing RACI values: {raci_empty}\n")
            else:
                for idx, row in validity_result_df.iterrows():
                    identifier = row["Identifier"]
                    duplicate_count = row["Duplicate Count"]
                    data_text.insert(tk.END, f"{identifier} - Duplicate Count: {duplicate_count}\n")
                data_text.insert(tk.END, f"There are {raci_empty} missing RACI values.\n")

        # Return True if all identifiers are unique, otherwise False
        return pd.to_numeric(validity_result_df["Duplicate Count"], errors="coerce").sum() == 0 and raci_empty == 0

    def check_accumulated_write_offs(becris_dataframe):
        """
        Checks the validity of the "Accumulated write-offs" column in the becris_dataframe.
        :param becris_dataframe: The DataFrame containing the becris data.
        :return: A boolean indicating whether the "Accumulated write-offs" column is valid or not.
        """
        # Create a new DataFrame to store the validity check results
        columns = ["Check", "Result"]
        validity_result = []

        # Check 1: Check [Financial.Outstanding nominal amount] + [Accounting.Accumulated write-offs]
        #                                                   >= [Financial.Arrears for the instrument]
        condition_1_valid = True
        if "NotRequired" not in becris_dataframe["Accumulated write-offs"].values:
            condition_1_valid = (
                    (pd.to_numeric(becris_dataframe["Outstanding nominal amount"]) +
                     pd.to_numeric(becris_dataframe["Accumulated write-offs"]))
                    >= pd.to_numeric(becris_dataframe["Arrears for the instrument"])
            )
        validity_result.append(["ER_DTS_CS_FIN_048", condition_1_valid])

        # ADD MORE CHECKS HERE IF NEEDED

        # Display the validity check result in the data_text widget
        data_text.insert(tk.END, "Accumulated Write-Offs Check:\n")

        if len(validity_result) == 1:
            result = "Passed" if all(validity_result[0][1]) else "Failed"
            data_text.insert(tk.END, f"{validity_result[0][0]} - Result: {result}\n")
            all_passed = validity_result[0][0]
        else:
            all_passed = True  # Assume all checks pass initially
            for idx, row in validity_result:
                check = row["Check"]
                result = "Passed" if row["Result"] else "Failed"
                data_text.insert(tk.END, f"{check} - Result: {result}\n")
                if not row["Result"]:  # If any check fails, set all_passed to False
                    all_passed = False

        if all_passed:
            data_text.insert(tk.END, "All checks passed for the 'Accumulated write-offs' column.\n")
        else:
            data_text.insert(tk.END, "One or more checks failed for the 'Accumulated write-offs' column.\n")

        # Return True if all checks passed, otherwise False
        return all_passed

    data_text.insert(tk.END, "Performing validity checks on counterparty references data...\n")
    is_counterparty_unique = check_counterparty_identifier_uniqueness(counterparty_dataframe)

    data_text.insert(tk.END, "\nPerforming validity checks on becris data...\n")
    is_accumulated_write_offs_valid = check_accumulated_write_offs(becris_dataframe)


# Create the GUI
root = tk.Tk()
root.title("Data Handler")

# Create top frame
top_frame = Frame(root)
top_frame.pack(side=tk.TOP)

# File selection button and entry
select_button = Button(top_frame, text="Select file", command=select_file)
select_button.grid(row=1, column=0, padx=(0, 10))  # Add some padding on the right

file_entry = tk.Entry(top_frame, width=90)
file_entry.grid(row=1, column=1)  # Place the entry next to the button

# Text widget for displaying data
data_text = Text(root, width=80, height=20)
data_text.pack()

# Progress bar
update_progress_bar = ttk.Progressbar(root, mode="determinate", length=400)
update_progress_bar.pack()
update_progress_bar.stop()

# Button to fetch data
fetch_button = Button(root, text="Fetch data", command=fetch_data, state=tk.DISABLED)
fetch_button.pack()

# Frame to hold the buttons
button_frame = Frame(root)
button_frame.pack()

# Button to check for new instruments
update_new_button = Button(button_frame, text="Check for new instruments", command=check_new, state=tk.DISABLED)
update_new_button.pack(side=tk.LEFT, padx=5)

# Button to paste data
paste_button = Button(button_frame, text="Paste data in ""DataExtraction"" tab", command=paste_data, state=tk.DISABLED)
paste_button.pack(side=tk.LEFT)

# Button to validate data
validate_button = Button(button_frame, text="Validate Becris Data", command=validate_data, state=tk.DISABLED)
validate_button.pack(side=tk.LEFT, padx=5)

# Log widget
log_text = Text(root, height=5)
log_text.pack()

# Run the GUI
root.mainloop()
