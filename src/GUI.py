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
        self.title("Becris Excel Reader")
        self.geometry("800x500")
        self.minsize(800, 500)
        
        # Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        
        # Read config file
        self.read_config()
        print(self.config.get("Settings", "last_directory"))

        self.select_file_screen = SelectFile(self, self.select_file, self.config.get("Settings", "last_directory", fallback="/"))

    def select_file(self, path):
        self.filename = path
        self.save_last_directory()
        self.select_file_screen.grid_forget()
        
        # Show the next interface
        self.start_interface()

    def start_interface(self):
        
        # Create 2 frames to fit the widgets
        self.menu_frame = ctk.CTkFrame(self, width=200, fg_color= "transparent")
        self.menu_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.text_frame = ctk.CTkFrame(self)
        self.text_frame.grid(row=0, column=1,sticky="nsew", padx=8, pady=8)
        self.text_frame.grid_rowconfigure(0, weight=1)
        self.text_frame.grid_columnconfigure(0, weight=1)
        
        # Create text box on the right side of the app
        self.text_output_log = FileOutput(self.text_frame)
        self.progress_bar = ProgressBar(self.text_frame)
        
        # Create menu on the left side of the app
        self.menu = Menu(self.menu_frame,self.select_file, self.filename)
        
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
        
        
