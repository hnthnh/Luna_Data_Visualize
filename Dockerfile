# Bước 1: Chọn image Python cho Windows
FROM python:3.8

# Bước 2: Cài đặt các thư viện cần thiết
RUN pip install --upgrade pip
COPY requirements.txt ./
RUN pip install -r requirements.txt

# Bước 3: Cài đặt PyInstaller
RUN pip install pyinstaller

LABEL version="0.0.1"
LABEL description="DAIHATSU DIESEL MFG.CO.,LTD"

# Bước 4: Tạo thư mục cho ứng dụng
WORKDIR /app

# Bước 5: Copy mã nguồn vào container
COPY . /app

# Bước 6: Build ứng dụng với PyInstaller
CMD ["pyinstaller", "--onefile", "--name=DAIHATSU_DIESEL_MFG_CO_LTD.exe", "main.py"]
