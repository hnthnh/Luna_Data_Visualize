import sys
import os
import pandas as pd
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog,QTableWidgetItem

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # Tải giao diện từ file .ui
        uic.loadUi('assets/giaodien.ui', self)
            # Ngắt kết nối nếu có
        self.btn_browsefolder.clicked.disconnect()
        self.btn_add.clicked.disconnect()
        self.btn_delete.clicked.disconnect()
        self.btn_start.clicked.disconnect()
        # Kết nối nút với hàm xử lý chỉ một lần
        self.btn_browsefolder.clicked.connect(self.on_btn_browsefolder_clicked)
        self.btn_add.clicked.connect(self.on_btn_add_clicked)
        self.btn_delete.clicked.connect(self.on_btn_delete_clicked)
        self.btn_start.clicked.connect(self.on_btn_start_clicked)

        print("Kết nối tín hiệu thành công.")

    def on_btn_browsefolder_clicked(self):
        # Mở hộp thoại để chọn thư mục
        folder_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục")
        if folder_path:
            # Lấy danh sách file CSV trong thư mục
            csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
            if csv_files:
                # Cập nhật đường dẫn vào textbox
                self.text_browsefolder.setText(folder_path)

                # Đọc file CSV đầu tiên
                first_csv_file = os.path.join(folder_path, csv_files[0])
                df = pd.read_csv(first_csv_file)

                # Lấy danh sách cột và thêm vào combobox
                self.comboBox_graphitems.clear()  # Xóa các mục cũ
                self.comboBox_graphitems.addItems(df.columns.tolist())  # Thêm danh sách cột vào combobox
                QMessageBox.information(self, "Thông báo", f"Đã lấy {len(df.columns)} cột từ file: {first_csv_file}")
            else:
                # Hiển thị thông báo lỗi nếu không có file CSV nào
                QMessageBox.critical(self, "Lỗi", "Thư mục không chứa file CSV nào!")
                
        

    def on_btn_add_clicked(self):
        # Lấy giá trị đã chọn từ combobox
        selected_name = self.comboBox_graphitems.currentText()
        
        if selected_name:
            # Lấy số lượng hàng hiện tại để làm ID
            row_position = self.table_data.rowCount()
            item_id = row_position + 1  # Tăng ID từ 1

            # Sử dụng các giá trị mặc định cho các cột khác
            upper_limit = 1200
            lower_limit = 0
            expan_number = 1.0
            
            # Thêm hàng vào bảng
            self.table_data.insertRow(row_position)
            
            # Đặt giá trị cho từng cột
            self.table_data.setItem(row_position, 0, QTableWidgetItem(selected_name))  # Cột NAME
            self.table_data.setItem(row_position, 1, QTableWidgetItem(str(upper_limit)))  # Cột UPPER LIMIT
            self.table_data.setItem(row_position, 2, QTableWidgetItem(str(lower_limit)))  # Cột LOWER LIMIT
            self.table_data.setItem(row_position, 3, QTableWidgetItem(str(expan_number)))  # Cột EXPAN NUMBER
        else:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn một tên trước khi thêm.")

    def on_btn_delete_clicked(self):
        # Ngắt kết nối trước khi thực hiện hành động xóa
        self.btn_delete.clicked.disconnect()

        # Thực hiện hành động xóa
        # Giả sử bạn đang sử dụng QTableWidget
        current_row = self.table_data.currentRow()
        if current_row >= 0:  # Kiểm tra xem có hàng nào được chọn không
            self.table_data.removeRow(current_row)
        else:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn hàng để xóa.")

        # Kết nối lại nút Delete
        self.btn_delete.clicked.connect(self.on_btn_delete_clicked)

    def on_btn_start_clicked(self):
        print("btn_start clicked")  # In ra console
        QMessageBox.information(self, "Thông báo", "Nút btn_start đã được nhấn!")
    def filter_combobox(self, text):
        # Lọc các mục trong combobox dựa trên văn bản tìm kiếm
        for index in range(self.comboBox_graphitems.count()):
            item_text = self.comboBox_graphitems.itemText(index)
            self.comboBox_graphitems.setItemData(index, item_text.lower().startswith(text.lower()), role=0)
        self.comboBox_graphitems.showPopup()    
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
