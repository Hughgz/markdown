# Backend DOCX/DOC to Markdown

Backend Flask cho ứng dụng hợp nhất DOCX/DOC và chuyển đổi sang Markdown.

## Cài đặt

1. Tạo môi trường ảo (virtual environment):

```bash
python -m venv venv
```

2. Kích hoạt môi trường ảo:

- Windows:
```bash
venv\Scripts\activate
```

- macOS/Linux:
```bash
source venv/bin/activate
```

3. Cài đặt các dependency:

```bash
pip install -r requirements.txt
```

## Yêu cầu hệ thống

- **Windows**: Không cần cài đặt thêm (sử dụng Microsoft Word thông qua Win32 API)
- **Linux/macOS**: Cần cài đặt LibreOffice để chuyển đổi từ DOC sang DOCX

## Chạy ứng dụng

```bash
flask run
```

Hoặc sử dụng Python trực tiếp:

```bash
python app.py
```

Mặc định, server sẽ chạy ở địa chỉ `http://localhost:5000`.

## API Endpoints

### Ghép file và chuyển đổi

- **URL**: `/api/merge-and-convert`
- **Method**: POST
- **Content-Type**: multipart/form-data
- **Params**:
  - `files[]`: Danh sách các file DOCX hoặc DOC (Bắt buộc)
  - `format`: Định dạng đầu ra, "markdown" hoặc "text" (Mặc định: "markdown")

- **Response**: File ghép được chuyển đổi sang định dạng đã chọn