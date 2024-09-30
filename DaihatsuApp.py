import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import zipfile
import os
import shutil
import io
from PIL import Image as PILimg
from PIL import ImageTk as PILImageTk
import gc
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image as OpenPyXLImage
import openpyxl  # Import openpyxl
import sys  # Import sys to exit the program
import matplotlib.dates as mdates  # Import for date formatting
import platform
import subprocess
import matplotlib.ticker as ticker
from data_processor import data_process
from excel_export import Exporter



class DaihatsuApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # Set window properties
        self.title("DAIHATSU DIESEL MFG.CO.,LTD")
        self.geometry("530x600")
        self.resizable(False, False)

        # Initialize variables
        self.zip_file_path = None
        self.destination = 'temp'
        self.extracted_files = []

        # Call methods to set up the GUI layout
        self.setup_logo()
        self.setup_browse_button()
        self.setup_browseFolder_button()
        self.setup_selected_file_label()
        self.status_label()
        self.setup_progress_bar()
        self.setup_input_fields()
        self.setup_file_listbox()
        self.setup_action_buttons()




    def setup_logo(self):
        # Load and display logo
        self.image = PILimg.open("assets/img_daihatsu_logo.png")
        self.image = self.image.resize((200, 50), PILimg.LANCZOS)
        self.logo = PILImageTk.PhotoImage(self.image)
        self.logo_label = tk.Label(self, image=self.logo)
        self.logo_label.grid(row=0, column=0)

    def setup_browse_button(self):
        # Browse button
        self.browse_button = ttk.Button(self, text="Browse ZIP File", command=self.browse_zip)
        self.browse_button.grid(row=0, column=1, sticky="ew")
    def setup_browseFolder_button(self):
        # Browse Folder button
        self.browse_button = ttk.Button(self, text="Browse folder", command=self.browse_folder)
        self.browse_button.grid(row=0, column=2, sticky="ew")
    def setup_selected_file_label(self):
        # Label to show selected file
        self.selected_file_label = ttk.Label(self, text="No file selected")
        self.selected_file_label.grid(row=1, column=0, columnspan=2, padx=20, pady=10, sticky="ew")
    
    def status_label(self):
        self.status_label=ttk.Label(self,text='')
        self.status_label.grid(row=5,column=0, columnspan=2, padx=20, pady=10, sticky="w")

    def setup_progress_bar(self):
        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=2, column=0, columnspan=2, padx=20, pady=10, sticky="ew")

    def setup_input_fields(self):
        # Upper and Lower limit input fields
        self.upper_label = tk.Label(self, text="Upper Limit:")
        self.upper_label.grid(row=3, column=0, padx=10, pady=10, sticky='w')
        self.upper_limit = tk.Entry(self)
        self.upper_limit.grid(row=3, column=0, padx=10, pady=10, sticky='e')
        self.upper_limit.insert(0, "1200")  # Default value

        self.lower_label = tk.Label(self, text="Lower Limit:")
        self.lower_label.grid(row=4, column=0, padx=10, pady=10, sticky='w')
        self.lower_limit = tk.Entry(self)
        self.lower_limit.grid(row=4, column=0, padx=10, pady=10, sticky='e')
        self.lower_limit.insert(0, "0")  # Default value

        # Drop-down menu
        self.selection_label = tk.Label(self, text="Please select X or O")
        self.selection_label.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        self.selected_value = tk.StringVar(value="X")
        self.selection_menu = tk.OptionMenu(self, self.selected_value, "X", "O")
        self.selection_menu.grid(row=4, column=1, padx=5, pady=5, sticky='ew')

    def setup_file_listbox(self):
        # Listbox to display extracted files
        self.file_listbox = tk.Listbox(self, border=3, borderwidth=3, activestyle="dotbox", width=60, height=10)
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.file_listbox.yview)
        self.file_listbox['yscrollcommand'] = scrollbar.set
        self.file_listbox.grid(row=6, column=0, columnspan=2, padx=20, pady=10, sticky="ns")

    def setup_action_buttons(self):
        # View chart button
        self.view_chart_button = ttk.Button(self, text="View Chart", command=self.process_file)
        self.view_chart_button.grid(row=7, column=0, padx=20, pady=10, sticky="ew")

        # Export to Excel button
        self.export_button = ttk.Button(self, text="Export Chart to Excel", command=self.export_excel)
        self.export_button.grid(row=7, column=1, padx=20, pady=10, sticky="ew")

        # Exit button
        self.exit_button = ttk.Button(self, text="Exit", command=self.confirm_exit)
        self.exit_button.grid(row=8, column=0, columnspan=2, padx=20, pady=30, sticky="ew")

    def browse_zip(self):

        ###Browse *zip
        self.zip_file_path = filedialog.askopenfilename(
            title="Select ZIP file", 
            filetypes=[("ZIP files", "*.zip")]
        )
        # Nếu người dùng đã chọn tệp ZIP
        if self.zip_file_path:
            # Cập nhật nhãn tệp được chọn
            self.selected_file_label.config(text=f"Selected: {os.path.basename(self.zip_file_path)}")

            # Xóa nội dung của status_label
            self.status_label.config(text="")

            # Xóa tất cả các mục trong file_listbox
            self.file_listbox.delete(0, tk.END)
            
            # Xóa thư mục đích nếu đã tồn tại (để giải nén mới)
            path = self.destination
            if os.path.exists(path):
                shutil.rmtree(path)

            # Tiến hành giải nén tệp
            self.unzip_file()
        else:
            # Nếu người dùng không chọn tệp ZIP
            self.status_label.config(text="No file selected.")

        pass
    def browse_folder(self):
        self.folder_path = filedialog.askdirectory(
        title="Select Folder"
        )
        if self.folder_path:
            self.status_label.config(text=f"Selected Folder: {self.folder_path}")
        csv_files = [f for f in os.listdir(self.folder_path) if f.endswith('.csv')]
        # Sao chép các tệp .csv vào thư mục temp
        i=0
        self.progress["value"] = 0
        self.progress["maximum"] = len(csv_files)
        for file in csv_files:
            self.destination = "temp"
            if not self.destination:
                return
            source_file = os.path.join(self.folder_path, file)
            destination_file = os.path.join(self.destination, file)
            shutil.copy(source_file, destination_file)
            self.file_listbox.insert(tk.END, file)
            i+=1
            self.progress["value"] += 1
            self.status_label.config(text=f"Copied {i}/{len(csv_files)} to {self.destination}")
            self.update_idletasks()
        messagebox.showinfo(f"Success file :", "copied complete!")
        self.status_label.config(text="copied complete!")
    def copy_csv_files(source_folder, temp_folder):
        # Tạo thư mục temp nếu chưa tồn tại
        if not os.path.exists(temp_folder):
            os.makedirs(temp_folder)
    def unzip_file(self):
        if not self.zip_file_path:
            return
        try:
            with zipfile.ZipFile(self.zip_file_path, 'r') as zip_ref:
                file_list = zip_ref.namelist()
                total_files = len(file_list)

                self.progress["value"] = 0
                self.progress["maximum"] = total_files

                self.destination = "temp"
                if not self.destination:
                    return

                self.extracted_files = []
                for idx, file in enumerate(file_list):
                    zip_ref.extract(file, self.destination)
                    self.extracted_files.append(file)
                    self.file_listbox.insert(tk.END, file)
                    self.progress["value"] += 1
                    self.status_label.config(text=f"Extracting {file} ({idx + 1}/{total_files})")
                    self.update_idletasks()

                messagebox.showinfo(f"Success file :", "Extraction complete!")
                self.status_label.config(text="Extraction complete!")

        except zipfile.BadZipFile:
            messagebox.showerror("Error", "Invalid ZIP file")
            self.status_label.config(text="Failed to extract files")
        pass
    def browse_file(self):
        #Logic for browing folder
        pass
    def process_file(self):
        # Logic for processing files
        pass

    def export_excel(self):
        # Tạo đối tượng Exporter và truyền self (DaihatsuApp)
        exporter = Exporter(self)

        # Gọi hàm export_excel của Exporter để bắt đầu quá trình xử lý
        exporter.export_excel()

    def confirm_exit(self):
        shutil.rmtree('temp')
        shutil.rmtree('image_temp')
        self.quit()    