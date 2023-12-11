import os
import xlwings as xw
import pandas as pd
import numpy as np
import math

from src.GUI import *

global status_list_dataframe
global wb_to_use
wb_to_use = None


def fetch_data(app_instance):
    global status_list_dataframe
    global wb_to_use

    try:
        # Get the selected file path
        file_path = app_instance.filename

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

                if str(sch_a_code).startswith("36300"):
                    type_acct = str(row[columns.index("TYPE OF ACCOUNT")])
                    if type_acct == "BANK GUARANTEES":
                        filtered_data.append(row)
                elif str(sch_a_code).startswith(("34321", "34311", "34220", "36400")):
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

        # Change amounts to 0 when needed
        condition = ((status_list_dataframe["GL Hd"].isin([12, 18]))
                     & (status_list_dataframe["AMT"] > 0)
                     & (status_list_dataframe["CONV  AMT"] > 0))

        status_list_dataframe.loc[condition, ["AMT", "CONV  AMT"]] = 0

        # Convert the values in the 'sch A acc' column to strings
        df['sch A acc'] = df['sch A acc'].astype(str)

        # Calculate summary information
        summary_info = calculate_summary()

        # Display the data summary in the GUI
        app_instance.display_summary(summary_info)

        app_instance.menu.extraction_frame.fetch_button.configure(fg_color="green")
        app_instance.menu.extraction_frame.paste_button.configure(state=ctk.NORMAL)
        app_instance.menu.extraction_frame.check_new_ended_button.configure(state=ctk.NORMAL)

        # Update the log widget
        # log_text.insert(tk.END, "Data fetched successfully.\n")

        print(df)

        # Make the paste data and check for new instruments buttons available
        # paste_button.config(state=tk.NORMAL)
        # update_new_button.config(state=tk.NORMAL)
        # validate_button.config(state=tk.NORMAL)
    except Exception as e:
        app_instance.menu.extraction_frame.fetch_button.configure(fg_color="darkred")
        print(f"Error: {str(e)}\n")


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
        print(f"Error: {str(e)}\n")


def paste_data(app_instance):
    global status_list_dataframe
    global wb_to_use

    try:
        # Get the "DataExtraction" sheet from the workbook
        app_instance.progress_bar.start_indeterminate()
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
        interest_reset_date_array = [
            [value] if not pd.isnull(value) else ["NotApplicable"]
            for value in interest_reset_date_values
        ]
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
        app_instance.progress_bar.stop_progress()
        wb_to_use.save()
        app_instance.menu.extraction_frame.paste_button.configure(fg_color="green")
        print("Data pasted successfully.")

    except Exception as e:
        app_instance.menu.extraction_frame.paste_button.configure(fg_color="darkred")
        print(f"Error: {str(e)}\n")


def check_new_ended(app_instance):
    global wb_to_use
    global status_list_dataframe

    try:
        # Start progress bar
        app_instance.progress_bar.start_indeterminate()

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
            app_instance.text_output_log.delete("1.0", ctk.END)
            app_instance.text_output_log.insert("1.0", f"\n\n{len(ended_instruments)} instrument(s) "
                                                       f"marked as ""ENDED"" in ALL CP sheet:\n")
            for index in ended_instruments:
                acct_no = existing_instruments[index]
                acct_no = int(acct_no) if isinstance(acct_no, float) else acct_no
                cp_sheet.range(f"A{index + 2}").value = "ENDED"  # +2 to account for header row and 0 based indexing
                app_instance.text_output_log.insert(tk.END, f"Instrument with ID {str(acct_no)} is ended.\n")
        else:
            print("No instruments have ended this month.\n")

        # Compare the current and previous month dataframes to check if we missed any instruments
        merged_df = pd.merge(prev_month_dataframe, status_list_dataframe, on="ACCT NO.", how="outer",
                             indicator=True)
        instruments_not_in_current_month = merged_df[merged_df["_merge"] == "left_only"]
        instruments_not_in_cp = instruments_not_in_current_month[
            ~instruments_not_in_current_month["ACCT NO."].astype(str).isin(map(str,
                                                                               existing_instruments))]

        if not instruments_not_in_cp.empty:
            app_instance.text_output_log.insert(tk.END,
                                                f"\n{len(instruments_not_in_cp)} instrument(s) found in the previous month that "
                                                f"are not found in the current month or ALL CP, please investigate these "
                                                f"manually:\n")
            app_instance.text_output_log.insert(tk.END, instruments_not_in_cp["ACCT NO."].apply(
                lambda x: "{:.0f}".format(x) if isinstance(x, float) else x).to_string(index=False))
        else:
            print("No instruments in previous month are not found in ALL CP or Current month.\n")

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
        app_instance.progress_bar.stop_progress()
        app_instance.menu.extraction_frame.check_new_ended_button.configure(fg_color="green")

        # Update the log widget
        if new_instruments:
            print(f"{len(new_instruments)} new instruments pasted successfully.\n")
        else:
            print("No new instruments found.\n")
        print("ALL CP checked successfully.\n")
    except Exception as e:
        app_instance.menu.extraction_frame.check_new_ended_button.configure(fg_color="darkred")
        print(f"Error: {str(e)}\n")
