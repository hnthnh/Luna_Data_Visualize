import os
import sys
import openpyxl
import gc
from openpyxl.drawing.image import Image as OpenPyXLImage
from tkinter import messagebox
from data_processor import data_process
import matplotlib.pyplot as plt

class Exporter:
    def __init__(self, app):
        self.app = app  # Lưu đối tượng DaihatsuApp
   
    def export_excel(self):
        # Khởi tạo một workbook mới hoặc mở file Excel hiện tại
        directory = 'temp'
        if not os.path.exists(directory):
            raise FileNotFoundError(f"The directory {directory} does not exist.")
        
        processor = data_process()

        # Lấy giá trị từ các widget Tkinter trong DaihatsuApp
        toplimit = int(self.app.upper_limit.get())
        bottomlimit = int(self.app.lower_limit.get())
        flag = str(self.app.selected_value.get())
        
        excel_file = 'plots.xlsx'
        processor.checkExist_ExcelFile(excel_file)

        # Lấy danh sách file CSV
        csv_files = list(processor.load_csv(directory))
        total_files = len(csv_files)
        
        if total_files == 0:
            messagebox.showinfo("No CSV files found in the directory. Stopping the program.")
            sys.exit()

        for i, csv_file in enumerate(csv_files):
            self.process_file(processor, csv_file, i, total_files, toplimit, bottomlimit, flag)

        # Lưu workbook và giải phóng bộ nhớ
        self.save_workbook(total_files)

    def process_file(self, processor, csv_file, file_index, total_files, toplimit, bottomlimit, flag):
        # Load và xử lý dữ liệu từ file CSV
        processor = data_process()
        df = processor.cut_data(csv_file)
        state_data = processor.verify_data(df)

        if state_data == 0:
            messagebox.showerror("Phát hiện dữ liệu không phải là số trong cột 'Engine Speed'. Dừng chương trình.")
            return
        elif state_data == 1:
            messagebox.showerror("Phát hiện dữ liệu không phải là số trong cột 'LEL Detector'. Dừng chương trình.")
            return
        
        # Điều chỉnh dữ liệu theo điều kiện
        if flag == "O":
            df['Engine Speed'] *= 0.1

        invalid_engine_speed = df[(df['Engine Speed'] < 0) | (df['Engine Speed'] > (200 if flag == "O" else 2000)) | df['Engine Speed'].isna()]
        if not invalid_engine_speed.empty:
            print("Invalid data in the 'Engine Speed' column:")
            print(invalid_engine_speed)
            raise ValueError("Invalid data detected in the 'Engine Speed' column. Stopping the program.")
        
        print(f"Processing file: {csv_file} ({file_index + 1}/{total_files})")

        # Gọi hàm tạo plot và lưu hình ảnh
        save_dir = 'image_temp'
        file_name = f'plot_{file_index + 1}.jpeg'
        plot_path = processor.create_plot(df, bottomlimit, toplimit, save_dir, file_name)

        # Cập nhật tiến độ và giải phóng bộ nhớ sau khi xử lý xong file
        self.update_progress(file_index, total_files)
        self.clean_up_memory(df, plot_path)

    def update_progress(self, file_index, total_files):
        """
        Cập nhật thanh tiến độ sau mỗi file xử lý
        """
        progress_value = (file_index + 1) / total_files * 100
        self.app.progress["value"] = progress_value  # Cập nhật tiến độ trong thanh progress của DaihatsuApp
        self.app.progress.update_idletasks()  # Đảm bảo giao diện người dùng được cập nhật


    def save_workbook(self, processed_files):
        excel_file = 'plots.xlsx'
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        for i in range(processed_files):
            current_row = 2 + i
            img_path = os.path.join('image_temp', f'plot_{i + 1}.jpeg')
            img_openpyxl = OpenPyXLImage(img_path)
            ws.add_image(img_openpyxl, f'E{current_row}')
            ws.row_dimensions[current_row].height = img_openpyxl.height * 0.75 + 20
            ws.column_dimensions['E'].width = img_openpyxl.width / 7.0

        wb.save(excel_file)
        wb.close()

    def clean_up_memory(self, df, plot_path):
        del df
        if os.path.exists(plot_path):
            os.remove(plot_path)
        plt.close()
        gc.collect()
