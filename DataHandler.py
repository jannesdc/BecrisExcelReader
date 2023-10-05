import xlwings as xw
import pandas as pd
import numpy as np
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, Text, Button, Frame
import configparser
import os

global status_list_dataframe
global wb_to_use
wb_to_use = None


def select_file():
    try:
        # Read config to remember last selected file directory
        config = configparser.ConfigParser()
        config.read("config.ini")
        # Check if the config.ini file exists
        if not os.path.exists("config.ini"):
            # Create the config.ini file with default values if it doesn't exist
            config['Settings'] = {'last_directory': '/'}
            with open("config.ini", "w") as configfile:
                config.write(configfile)
        if 'Settings' not in config:
            config['Settings'] = {'last_directory': '/'}
        last_directory = config.get("Settings", "last_directory", fallback="/")

        # Open a file selection dialog
        root.filename = filedialog.askopenfilename(initialdir=last_directory, title="Select Excel File",
                                                   filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))

        if root.filename:
            config.set("Settings", "last_directory", "/".join(root.filename.split("/")[:-1]))
            with open("config.ini", "w") as configfile:
                config.write(configfile)
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
        last_row = source_sheet.range("C" + str(source_sheet.cells.last_cell.row)).end("up").row

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
            gl_hd = str(row[columns.index("GL Hd")])

            if acct_no and sch_a_code and cn != "ALL" \
                    and str(sch_a_code).startswith(("34321", "34311", "34220", "36300", "36400")):
                first_five_numbers = str(sch_a_code)[:5]
                if isinstance(acct_no, str):
                    acct_no = str(acct_no)
                elif isinstance(acct_no, float):
                    acct_no = str(int(acct_no))
                else:
                    acct_no = str(row["ACCT NO."])
                new_acct_no = f"{acct_no}_{first_five_numbers}"
                row.append("Off-balance")
                row.append(new_acct_no)
                filtered_data.append(row)

            elif acct_no and sch_a_code and cn != "ALL":
                if gl_hd in {'12.0', '18.0'}:
                    row.append("On-balance")
                    row.append(acct_no)
                    filtered_data.append(row)
                elif str(sch_a_code).startswith(("11", "12")):
                    row.append("On-balance")
                    row.append(acct_no)
                    filtered_data.append(row)

        # Create a DataFrame from the filtered data
        columns.append("Balance Type")
        columns.append("Modified ACCT NO.")
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
        num_rows_onb = len(status_list_dataframe[status_list_dataframe["Balance Type"] == "On-balance"])
        num_rows_ofb = len(status_list_dataframe[status_list_dataframe["Balance Type"] == "Off-balance"])

        # Calculate the sums of "CONV AMT" based on "sch A acc" code
        sums_by_sch_a_code = status_list_dataframe[status_list_dataframe["Balance Type"] ==
                                                   "On-balance"].groupby("sch A acc")["CONV  AMT"].sum().reset_index()
        tot_sum_amt_ofb = status_list_dataframe[status_list_dataframe["Balance Type"] ==
                                                "Off-balance"].groupby("sch A acc")["AMT"].sum().reset_index()
        tot_sum_conv_amt_ofb = status_list_dataframe[status_list_dataframe["Balance Type"] ==
                                                     "Off-balance"].groupby("sch A acc")[
            "CONV  AMT"].sum().reset_index()

        # Store the summary information
        summary_info = {
            "Number of Rows ONB": num_rows_onb,
            "Number of Rows OFB": num_rows_ofb,
            "Sums by sch A acc": sums_by_sch_a_code,
            "Tot sum OFB": tot_sum_amt_ofb,
            "Tot conv sum OFB": tot_sum_conv_amt_ofb
        }

        return summary_info
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")


def display_summary(summary_info):
    try:
        # Display the data summary in the GUI
        data_text.delete("1.0", "end")
        data_text.insert("1.0", "Summary Information:\n\n")

        # Display the number of on balance rows
        num_rows_onb = summary_info["Number of Rows ONB"]
        data_text.insert(tk.END, f"Number of On-balance Rows: {num_rows_onb}\n\n")

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

        # Display the sum of all on balance data
        conv_amt_sum_total = abs(sums_by_sch_a_acc["CONV  AMT"].sum())
        data_text.insert(tk.END, f"\nTotal CONV AMT: {conv_amt_sum_total:,.2f}\n")

        # Display the number of off balance rows
        num_rows_ofb = summary_info["Number of Rows OFB"]
        data_text.insert(tk.END, f"\nNumber of Off-balance Rows: {num_rows_ofb}\n\n")

        # Display the sum of all off balance data
        tot_sum_ofb = summary_info["Tot sum OFB"]
        amt_sum_total = abs(tot_sum_ofb["AMT"].sum())
        data_text.insert(tk.END, f"Total AMT: {amt_sum_total:,.2f}\n")
        tot_sum_conv_ofb = summary_info["Tot conv sum OFB"]
        ofb_conv_amt_sum_total = abs(tot_sum_conv_ofb["CONV  AMT"].sum())
        data_text.insert(tk.END, f"Total CONV AMT: {ofb_conv_amt_sum_total:,.2f}\n")
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
        last_row_cp = extraction_sheet.range("A2:M" + str(extraction_sheet.cells.last_cell.row))
        last_row_cp.clear_contents()

        dest_range = extraction_sheet.range("A2:G{}".format(len(status_list_dataframe) + 1))
        interest_rate_type_range = extraction_sheet.range("H2:H{}".format(len(status_list_dataframe) + 1))
        mat_date_range = extraction_sheet.range("I2:I{}".format(len(status_list_dataframe) + 1))
        interest_rate_range = extraction_sheet.range("J2:J{}".format(len(status_list_dataframe) + 1))
        interest_reset_date_range = extraction_sheet.range("K2:K{}".format(len(status_list_dataframe) + 1))
        undrawn_range = extraction_sheet.range("L2:L{}".format(len(status_list_dataframe) + 1))
        modified_acct_no_range = extraction_sheet.range("M2:M{}".format(len(status_list_dataframe) + 1))

        dest_range.value = status_list_dataframe[
            ["ACCT NO.", "sch A acc", "CONV  AMT", "CON RT", "AMT", "ACCD INTT", "CUR"]].values
        modified_acct_no_range.value = status_list_dataframe[["Modified ACCT NO."]].values

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


def check_new_ended():
    global wb_to_use
    global status_list_dataframe
    global update_progress_bar

    try:
        # Get the "ALL CP" sheet
        cp_sheet = wb_to_use.sheets("ALL CP")
        prev_month_sheet = wb_to_use.sheets("PreviousMonth")

        # Populate previous month's dataframe
        prev_month_last_row = prev_month_sheet.range("A" + str(prev_month_sheet.cells.last_cell.row)).end("up").row
        prev_month_header_range = prev_month_sheet.range("A2:AK2")
        prev_month_data_range = prev_month_sheet.range("A3").expand("down").resize(prev_month_last_row - 1,
                                                                                   prev_month_header_range.columns.count)
        columns = prev_month_header_range.value
        data = prev_month_data_range.value

        prev_month_filtered_data = []
        for row in data:
            acct_no = row[columns.index("ACCT NO.")]
            sch_a_code = row[columns.index("sch A acc")]
            cn = str(row[columns.index("CN")])

            if acct_no and sch_a_code and cn != "ALL" and str(sch_a_code).startswith(("11", "12")):
                prev_month_filtered_data.append(row)

            elif acct_no and sch_a_code and cn != "ALL" \
                    and str(sch_a_code).startswith(("34321", "34311", "34220", "36300", "36400")):
                if str(sch_a_code).startswith("36300"):
                    type_acct = row[columns.index("TYPE OF ACCOUNT")]
                    if type_acct == "BANK GUARANTEES":
                        prev_month_filtered_data.append(row)
                elif str(sch_a_code).startswith(("34321", "34311", "34220", "36400")):
                    prev_month_filtered_data.append(row)

        prev_month_dataframe = pd.DataFrame(prev_month_filtered_data, columns=columns)

        # create new_instruments and ended_instruments list
        new_instruments = []
        ended_instruments = []

        # Check the "ALL CP" sheet for the existing instruments
        existing_instr_range = cp_sheet.range(
            "B2:B" + str(cp_sheet.range("A" + str(cp_sheet.cells.last_cell.row)).end("up").row))
        existing_instruments = existing_instr_range.value

        # Before we add new instruments we check for any ended instruments
        for index, row in enumerate(existing_instruments):
            if row not in status_list_dataframe[status_list_dataframe["Balance Type"] == "On-balance"][
                "ACCT NO."].values:
                if row not in status_list_dataframe[status_list_dataframe["Balance Type"] == "Off-balance"][
                    "Modified ACCT NO."].values:
                    ended_instruments.append(index)

        if ended_instruments:
            data_text.delete("1.0", tk.END)
            data_text.insert("1.0",
                             f"\n\n{len(ended_instruments)} instrument(s) marked as ""ENDED"" in ALL CP sheet:\n")
            for index in ended_instruments:
                acct_no = existing_instruments[index]
                acct_no = int(acct_no) if isinstance(acct_no, float) else acct_no
                cp_sheet.range(f"A{index + 2}").value = "ENDED"  # +2 to account for header row and 0 based indexing
                data_text.insert(tk.END, f"Instrument with ID {str(acct_no)} is ended.\n")
        else:
            log_text.insert(tk.END, "No instruments have ended this month.\n")

        # Compare the current and previous month dataframes to check if we missed any instruments
        merged_df = pd.merge(prev_month_dataframe, status_list_dataframe, on="ACCT NO.", how="outer",
                             indicator=True)
        instruments_not_in_current_month = merged_df[merged_df["_merge"] == "left_only"]
        instruments_not_in_cp = instruments_not_in_current_month[
            ~instruments_not_in_current_month["ACCT NO."].astype(str).isin(map(str,
                                                                               existing_instruments))]

        if not instruments_not_in_cp.empty:
            data_text.insert(tk.END, f"\n{len(instruments_not_in_cp)} instrument(s) found in the previous month that "
                                     f"are not found in the current month or ALL CP, please investigate these "
                                     f"manually:\n")
            data_text.insert(tk.END, instruments_not_in_cp["ACCT NO."].apply(
                lambda x: "{:.0f}".format(x) if isinstance(x, float) else x).to_string(index=False))
        else:
            log_text.insert(tk.END,
                            "No instruments in previous month are not found in ALL CP or Current month.\n")

        # Check for any new instruments comparing the existing instruments listed in the "ALL CP" sheet
        # comparing it with the instruments in the dataframe
        for index, row in status_list_dataframe[status_list_dataframe["Balance Type"] == "On-balance"].iterrows():
            acct_no = str(row["ACCT NO."])
            if acct_no not in map(str, existing_instruments):
                new_instruments.append(row)
        for index, row in status_list_dataframe[status_list_dataframe["Balance Type"] == "Off-balance"].iterrows():
            acct_no = str(row["Modified ACCT NO."])
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
                if instrument["Balance Type"] == "On-balance":
                    cp_sheet.range("A" + str(paste_row)).value = "New"
                    cp_sheet.range("B" + str(paste_row)).value = instrument.iloc[2]
                    cp_sheet.range("C" + str(paste_row)).value = instrument.iloc[3]
                    cp_sheet.range("E" + str(paste_row)).value = -instrument.iloc[5]
                    cp_sheet.range("F" + str(paste_row)).value = instrument.iloc[23]
                    if int(instrument.iloc[12]) == 18:
                        cp_sheet.range("G" + str(paste_row)).value = 20
                    elif int(instrument.iloc[12]) == 12:
                        cp_sheet.range("G" + str(paste_row)).value = 1000
                    elif int(instrument.iloc[12]) == 22:
                        cp_sheet.range("G" + str(paste_row)).value = 71
                    elif int(instrument.iloc[12]) in {20, 21}:
                        cp_sheet.range("G" + str(paste_row)).value = 1004
                    else:
                        cp_sheet.range("G" + str(paste_row)).value = "NOT FOUND"
                        cp_sheet.range("G" + str(paste_row)).color = 6
                elif instrument["Balance Type"] == "Off-balance":
                    cp_sheet.range("A" + str(paste_row)).value = "New"
                    cp_sheet.range("B" + str(paste_row)).value = instrument["Modified ACCT NO."]
                    cp_sheet.range("C" + str(paste_row)).value = instrument.iloc[3]
                    cp_sheet.range("E" + str(paste_row)).value = instrument.iloc[5]
                    if str(instrument["sch A acc"]).startswith(("34220", "36300", "36400")):
                        cp_sheet.range("G" + str(paste_row)).value = 9000
                    elif str(instrument["sch A acc"]).startswith(("34311", "34321")):
                        cp_sheet.range("G" + str(paste_row)).value = 9002

        # Stop the progress bar
        update_progress_bar.stop()

        # Update the log widget
        if new_instruments:
            log_text.insert(tk.END, f"{len(new_instruments)} new instruments pasted successfully.\n")
        else:
            log_text.insert(tk.END, "No new instruments found.\n")
        log_text.insert(tk.END, "ALL CP checked successfully.\n")
    except Exception as e:
        log_text.insert(tk.END, f"Error: {str(e)}\n")


def validate_data():
    """
    Checks per column of the "Data" tab the if the data is possible
    :return:
    """
    try:
        global wb_to_use
        global status_list_dataframe

        # First we fill the becris and counterparty dataframes,
        # so they can be used later for validation.
        counterparty_sheet = wb_to_use.sheets['Counterparties references']
        becris_sheet = wb_to_use.sheets['Data']

        # Populating the counterparty dataframe
        counterparty_last_row = counterparty_sheet.range("A" + str(counterparty_sheet.cells.last_cell.row)).end(
            "up").row
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
    except Exception as e:
        log_text.insert(tk.END, f'Error: {str(e)}\n')

    def check_counterparty_identifier_uniqueness(cp_dataframe):
        """
        Checks if Counterparty identifiers (ENI, LEI, RACI) in the Counterparty reference dataset are unique.
        :param cp_dataframe: The DataFrame containing the Counterparty reference data.
        :return: A boolean indicating whether the identifiers are unique or not.
        """

        # Create a new DataFrame to store the validity check results
        columns = ["Identifier", "Duplicate Count"]
        validity_result = []

        # Check ENI uniqueness, ignoring "NotRequired" values
        eni_unique = (cp_dataframe["ENI"] != "NotRequired") \
                     & cp_dataframe.duplicated(subset=["ENI"], keep=False)
        eni_duplicates = eni_unique.sum()
        validity_result.append(["ENI", eni_duplicates])

        # Check LEI uniqueness, ignoring "NotApplicable" values
        lei_unique = (cp_dataframe["LEI"] != "NotApplicable") \
                     & cp_dataframe.duplicated(subset=["LEI"], keep=False)
        lei_duplicates = lei_unique.sum()
        validity_result.append(["LEI", lei_duplicates])

        # Check RACI for empty values (should not be empty)
        raci_unique = cp_dataframe.duplicated(subset=["RACI"], keep=False)
        raci_duplicates = raci_unique.sum()
        raci_empty = cp_dataframe["RACI"].isnull().sum()
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

    """def check_accumulated_write_offs(becris_dataframe):
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

        # Check 2: Accumulated write-offs should always be a number never anything else
        try:
            pd.to_numeric(becris_dataframe["Accumulated write-offs"])
            condition_2_valid = True
        except:
            condition_2_valid = False

        validity_result.append(["Accumulated write-offs should be numbers", condition_2_valid])

        # Display the validity check result in the data_text widget
        data_text.insert(tk.END, "Accumulated Write-Offs Check:\n")

        if len(validity_result) == 1:
            result = "Passed" if all(validity_result[0][1]) else "Failed"
            data_text.insert(tk.END, f"{validity_result[0][0]} - Result: {result}\n")
            all_passed = all(validity_result[0][1])
        else:
            all_passed = True  # Assume all checks pass initially
            for row in validity_result:
                check = row[0]
                if type(row[1]) is bool:
                    result = "Passed" if row[1] else "Failed"
                    if not row[1]:  # If any check fails, set all_passed to False
                        all_passed = False
                else:
                    result = "Passed" if all(row[1]) else "Failed"
                    if not all(row[1]):  # If any check fails, set all_passed to False
                        all_passed = False
                data_text.insert(tk.END, f"{check} - Result: {result}\n")


        if all_passed:
            data_text.insert(tk.END, "All checks passed for the 'Accumulated write-offs' column.\n")
        else:
            data_text.insert(tk.END, "One or more checks failed for the 'Accumulated write-offs' column.\n")

        # Return True if all checks passed, otherwise False
        return all_passed
"""

    data_text.insert(tk.END, "Performing validity checks on counterparty references data...\n")
    is_counterparty_unique = check_counterparty_identifier_uniqueness(counterparty_dataframe)


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
data_text = Text(root, width=85, height=25)
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
update_new_button = Button(button_frame, text="Check for new/ended instruments", command=check_new_ended,
                           state=tk.DISABLED)
update_new_button.pack(side=tk.LEFT, padx=5)

# Button to paste data
paste_button = Button(button_frame, text="Paste data in ""DataExtraction"" tab", command=paste_data, state=tk.DISABLED)
paste_button.pack(side=tk.LEFT)

# Button to validate data
validate_button = Button(button_frame, text="Validate CP data", command=validate_data, state=tk.DISABLED)
validate_button.pack(side=tk.LEFT, padx=5)

# Log widget
log_text = Text(root, height=5)
log_text.pack()

# Run the GUI
root.mainloop()
