import os
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk
import configparser
import pandas as pd

from tkinter import filedialog
from src.UI import *


class App(ctk.CTk):

    def __init__(self):

        # Initial window setup
        super().__init__()
        self.menu = None
        self.progress_bar = None
        self.text_output_log = None
        self.text_frame = None
        self.menu_frame = None
        self.filename = None
        self.title("Becris Excel Reader")
        self.geometry("800x500")
        self.minsize(800, 500)

        # Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # Read config file
        self.read_config()
        print(self.config.get("Settings", "last_directory"))

        self.select_file_screen = SelectFile(self, self.select_file, self.config.get("Settings", "last_directory",
                                                                                     fallback="/"))

    def select_file(self, path):
        self.filename = path
        self.save_last_directory()
        self.select_file_screen.grid_forget()

        # Show the next interface
        self.start_interface()

    def start_interface(self):

        # Create 2 frames to fit the widgets
        self.menu_frame = ctk.CTkFrame(self, width=200, fg_color="transparent")
        self.menu_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.text_frame = ctk.CTkFrame(self)
        self.text_frame.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        self.text_frame.grid_rowconfigure(0, weight=1)
        self.text_frame.grid_columnconfigure(0, weight=1)

        # Create text box on the right side of the app
        self.text_output_log = FileOutput(self.text_frame)
        self.progress_bar = ProgressBar(self.text_frame, self)

        # Create menu on the left side of the app
        self.menu = Menu(self.menu_frame, self.select_file, self.filename, self)

    def read_config(self):

        self.config = configparser.ConfigParser()
        if not os.path.exists("config.ini"):
            # Create config.ini file with default values
            self.config['Settings'] = {'last_directory': '/'}
            with open("config.ini", "w") as configfile:
                self.config.write(configfile)

        # Read config file
        self.config.read("config.ini")
        if 'Settings' not in self.config:
            self.__dict__config['Settings'] = {'last_directory': '/'}

    def save_last_directory(self):
        if self.filename:
            print(f"Selected file: {self.filename}")
            self.config.set("Settings", "last_directory", "/".join(self.filename.split("/")[:-1]))
            with open("config.ini", "w") as configfile:
                self.config.write(configfile)

    def display_summary(self, summary_info):
        try:
            # Display the data summary in the GUI
            self.text_output_log.configure(state="normal")
            self.text_output_log.delete("1.0", "end")
            self.text_output_log.configure(state="disabled")
            self.text_output_log.insert_text("1.0", "Summary Information:\n\n")

            # Display the number of on balance rows
            num_rows_onb = summary_info["Number of Rows ONB"]
            self.text_output_log.insert_text(tk.END, f"Number of On-balance Rows: {num_rows_onb}\n\n")

            # Display the sums by sch A acc
            sums_by_sch_a_acc = summary_info["Sums by sch A acc"]
            sums_by_group = sums_by_sch_a_acc.groupby(sums_by_sch_a_acc["sch A acc"].str[:4])
            self.text_output_log.insert_text(tk.END, "Sums by Group:\n")
            for group, group_data in sums_by_group:
                conv_amt_sum = abs(group_data["CONV  AMT"].sum())
                self.text_output_log.insert_text(tk.END, f"Group: {group}\tCONV AMT: {conv_amt_sum:,.2f}\n")

            # Display the sums by sch A acc (2 digits)
            sums_by_group_2 = sums_by_sch_a_acc.groupby(sums_by_sch_a_acc["sch A acc"].str[:2])
            self.text_output_log.insert_text(tk.END, "\nSums by Group:\n")
            for group_2, group_data_2 in sums_by_group_2:
                conv_amt_sum_2 = abs(group_data_2["CONV  AMT"].sum())
                self.text_output_log.insert_text(tk.END, f"Group: {group_2}\tCONV AMT: {conv_amt_sum_2:,.2f}\n")

            # Display the sum of all on balance data
            conv_amt_sum_total = abs(sums_by_sch_a_acc["CONV  AMT"].sum())
            self.text_output_log.insert_text(tk.END, f"\nTotal CONV AMT: {conv_amt_sum_total:,.2f}\n")

            # Display the number of off balance rows
            num_rows_ofb = summary_info["Number of Rows OFB"]
            self.text_output_log.insert_text(tk.END, f"\nNumber of Off-balance Rows: {num_rows_ofb}\n\n")

            # Display the sum of all off balance data
            tot_sum_ofb = summary_info["Tot sum OFB"]
            amt_sum_total = abs(tot_sum_ofb["AMT"].sum())
            self.text_output_log.insert_text(tk.END, f"Total AMT: {amt_sum_total:,.2f}\n")
            tot_sum_conv_ofb = summary_info["Tot conv sum OFB"]
            ofb_conv_amt_sum_total = abs(tot_sum_conv_ofb["CONV  AMT"].sum())
            self.text_output_log.insert_text(tk.END, f"Total CONV AMT: {ofb_conv_amt_sum_total:,.2f}\n")
        except Exception as e:
            print(f"Error: {str(e)}\n")

