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
        
        self.set_screen()
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
    def set_screen(self):
        screen = QApplication.primaryScreen()
        screen_size = screen.size()

        # Tính toán kích thước 70% của màn hình
        new_width = int(screen_size.width() * 0.625)
        new_height = int(screen_size.height() * 0.9)
        # Đặt kích thước cố định 70% màn hình
        self.setFixedSize(new_width, new_height)
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
    # Lấy dữ liệu từ bảng
        table_data = self.get_table_data()

        # Xóa sạch dữ liệu trong QListWidget trước khi thêm mới
        self.list_process.clear()

        # Duyệt qua dữ liệu của bảng và thêm từng hàng vào QListWidget
        for row_data in table_data:
            # Chuyển đổi dữ liệu hàng thành chuỗi để hiển thị trong QListWidget
            row_text = ', '.join(row_data)
            self.list_process.addItem(row_text)
    def get_table_data(self):
        # Tạo danh sách để lưu dữ liệu từ table
        table_data = []

        # Duyệt qua từng hàng trong QTableWidget
        row_count = self.table_data.rowCount()
        column_count = self.table_data.columnCount()

        for row in range(row_count):
            row_data = []
            for column in range(column_count):
                # Lấy dữ liệu từ mỗi ô trong hàng
                item = self.table_data.item(row, column)
                if item is not None:
                    row_data.append(item.text())  # Lưu giá trị text của ô vào danh sách
                else:
                    row_data.append('')  # Nếu ô trống, thêm chuỗi rỗng

            # Thêm dữ liệu hàng này vào table_data
            table_data.append(row_data)

        return table_data

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
