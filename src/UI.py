import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog


class SelectFile(ctk.CTkFrame):
    def __init__(self, parent, select_file_func, last_directory):
        super().__init__(master=parent)
        self.grid(column=0, columnspan=2, row=0, sticky="nsew")
        self.select_file_func = select_file_func
        self.last_directory = last_directory

        ctk.CTkButton(self, text="Select File",
                      command=self.open_dialog).pack(expand=True)

    def open_dialog(self):
        path = filedialog.askopenfilename(initialdir=self.last_directory, title="Select Excel File",
                                          filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        self.select_file_func(path)


class Menu(ctk.CTkTabview):
    def __init__(self, parent, select_file_func, path):
        super().__init__(master=parent)
        self.pack(expand=True, fill="both")

        # Tabs
        self.add("File")
        self.add("Extraction")
        self.add("Verification")
        self.add("Settings")
        self.set("Extraction")

        # Widgets
        FileFrame(self.tab("File"), select_file_func, path)
        ExtractionFrame(self.tab("Extraction"))
        VerificationFrame(self.tab("Verification"))
        SettingsFrame(self.tab("Settings"))


class FileFrame(ctk.CTkFrame):
    def __init__(self, parent, select_file_func, path):
        super().__init__(master=parent, fg_color="transparent")
        self.grid_rowconfigure(0, weight=1)
        self.pack(expand=True, fill="both")
        self.select_file_func = select_file_func
        self.last_directory = path

        # Path entry box and frame to contain it
        self.top_frame = ctk.CTkFrame(master=self, fg_color="transparent")
        self.top_frame.pack(fill="x", ipady=5)
        stringvar_path = tk.StringVar(value=self.last_directory)
        self.path_entry = ctk.CTkEntry(
            self.top_frame, textvariable=stringvar_path)
        self.path_entry.pack(side=ctk.LEFT)
        self.path_entry.configure(state="disabled")

        # Select file button
        self.select_file_button = ctk.CTkButton(
            self.top_frame, text="Select File", command=self.open_dialog)
        self.select_file_button.pack(side=ctk.LEFT)

        # Textbox where the summary information will be kept even if it is removed in the main text box
        self.summary_textbox = ctk.CTkTextbox(
            self, state="disabled", wrap="none")
        self.summary_textbox.pack(fill="both", expand=True )

    def open_dialog(self):
        path = filedialog.askopenfilename(initialdir="/", title="Select Excel File",
                                          filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        self.select_file_func(path)


class ExtractionFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(master=parent, fg_color="transparent")
        self.pack(expand=True, fill="both")


class VerificationFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(master=parent, fg_color="transparent")
        self.pack(expand=True, fill="both")


class SettingsFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(master=parent, fg_color="transparent")
        self.pack(expand=True, fill="both")


class FileOutput(ctk.CTkTextbox):
    def __init__(self, parent):
        super().__init__(master=parent, state="disabled", wrap="none")
        self.grid(sticky="nsew", padx=6, pady=6)

    def insert_text(self, index, text):
        """Inserts text at the specified location but this method works even if the textbox is read-only unlike the standard insert method

        Args:
            location (ANY): ctk.END: inserts at the end of the text box
            text (str): string that to print in the textbox
        """
        self.configure(state="normal")
        self.insert(index, text)
        self.configure(state="disabled")


class ProgressBar(ctk.CTkProgressBar):
    def __init__(self, parent):
        super().__init__(master=parent, mode="determinate")
        self.grid(row=1, sticky="nsew")
        self.set(0)

    def start_indeterminate(self):
        self.configure(mode="indeterminate")
        self.start()
