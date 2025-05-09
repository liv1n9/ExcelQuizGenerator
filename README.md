# Ứng Dụng Tạo Đề Thi Trắc Nghiệm

Ứng dụng web này cho phép người dùng tạo các đề thi trắc nghiệm ngẫu nhiên từ một ngân hàng câu hỏi được lưu trong file Excel.

## Tính Năng

- Tải lên file Excel chứa câu hỏi trắc nghiệm và đáp án
- Chọn ngẫu nhiên số lượng câu hỏi mong muốn
- Tạo nhiều phiên bản đề thi khác nhau
- Tải xuống file Word định dạng đẹp
- Hỗ trợ hai phiên bản: đề thi thường và đề thi có bôi đậm đáp án đúng (dành cho giáo viên)

## Yêu Cầu Hệ Thống

- Python 3.11 trở lên
- Các thư viện Python cần thiết (xem file requirements.txt)

## Cài Đặt

### 1. Clone Repository

```bash
git clone https://github.com/your-username/quiz-generator.git
cd quiz-generator
```

### 2. Tạo Môi Trường Ảo (Khuyến nghị)

```bash
python -m venv venv
```

Kích hoạt môi trường ảo:

- Windows:
```bash
venv\Scripts\activate
```

- macOS/Linux:
```bash
source venv/bin/activate
```

### 3. Cài Đặt Các Thư Viện Cần Thiết

Cài đặt các thư viện cần thiết:

```bash
pip install flask flask-sqlalchemy pandas numpy python-docx werkzeug gunicorn email-validator psycopg2-binary zipfile36
```

## Chạy Ứng Dụng

```bash
python main.py
```

Ứng dụng sẽ chạy tại địa chỉ http://localhost:5000

Với Gunicorn (khuyến nghị cho môi trường sản xuất):

```bash
gunicorn --bind 0.0.0.0:5000 --reuse-port --reload main:app
```

## Cách Sử Dụng

1. Chuẩn bị file Excel với các cột sau:
   - `Câu hỏi`: Nội dung câu hỏi
   - `A`: Nội dung đáp án A
   - `B`: Nội dung đáp án B
   - `C`: Nội dung đáp án C
   - `D`: Nội dung đáp án D
   - `đáp án`: Đáp án đúng (A, B, C, hoặc D)

2. Mở ứng dụng trong trình duyệt và tải lên file Excel

3. Nhập số lượng câu hỏi (n) và số phiên bản đề thi (m)

4. Nhấn "Tạo Đề Thi" và tải về các file ZIP được tạo

## Cấu Trúc Dự Án

```
.
├── main.py               # File khởi động ứng dụng
├── app.py                # Ứng dụng Flask chính
├── requirements.txt      # Các thư viện cần thiết
├── static/               # CSS, JavaScript, và các tài nguyên tĩnh
├── templates/            # HTML templates
└── utils/                # Các module hỗ trợ
    ├── excel_processor.py  # Xử lý file Excel
    └── document_generator.py  # Tạo tài liệu Word
```

## Đóng Góp

Mọi đóng góp đều được hoan nghênh. Vui lòng tạo issue hoặc pull request nếu bạn muốn đóng góp cho dự án.

## Giấy Phép

[MIT License](LICENSE)