# Quản lý Điểm Sinh viên (Bài tập Excel)

## 📌 Giới thiệu
Đây là dự án/bài tập thực hành xử lý bảng điểm thi học phần **CT428** (Năm học 2020-2021). File Excel này thực hiện việc tổng hợp điểm số từ các thành phần (thực hành, bài tập lớn, lý thuyết), tự động quy đổi sang thang điểm 10, xếp loại điểm chữ và thống kê chi tiết kết quả học tập.

## 📁 Cấu trúc Dữ liệu
Dữ liệu trong `BAI1.xlsx` được chia thành 7 sheets để phân loại và xử lý:
* **Tổng hợp điểm:** Bảng tính chính chứa dữ liệu điểm chi tiết, tổng điểm, điểm quy đổi và điểm chữ của sinh viên.
* **Điểm thực hành các buổi:** Bảng ghi nhận điểm thực hành qua từng buổi.
* **Điểm bài tập lớn:** Điểm đánh giá phần chuyên đề/bài tập lớn.
* **Điểm lý thuyết:** Điểm thi lý thuyết cuối kỳ.
* **Các bảng tham chiếu & Thống kê:** Bảng phụ trợ dùng để thiết lập quy tắc quy đổi điểm số và biểu diễn thống kê tỷ lệ kết quả thi.
* **SinhVienDiem < 7:** Danh sách trích xuất tự động những sinh viên có kết quả dưới 7.
* **SinhVienDiem >= 7:** Danh sách trích xuất tự động những sinh viên đạt từ 7 điểm trở lên.

## 🛠 Công cụ & Kỹ năng
* **Công cụ:** Microsoft Excel.
* **Kỹ năng ứng dụng:**
  * Thao tác và quản lý dữ liệu bảng tính.
  * Ứng dụng các hàm tham chiếu (`VLOOKUP`) và tính toán logic/thống kê cơ bản (`IF`, `SUM`, `AVERAGE`).
  * Sử dụng tính năng Lọc nâng cao (Advanced Filter) để trích xuất danh sách sinh viên theo điều kiện.
# 📊 Dataset Kết quả Học tập Sinh viên (BAI2)

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Data Format](https://img.shields.io/badge/format-CSV%2FExcel-green.svg)
![Language](https://img.shields.io/badge/language-Vietnamese-orange.svg)

## 📖 Giới thiệu
Đây là kho lưu trữ (repository) chứa tập dữ liệu tổng hợp kết quả học tập của sinh viên. Dữ liệu này được trích xuất từ bảng tính **BAI2** (Sheet `Tong_hop`) và bao gồm thông tin chi tiết về chuyên ngành, môn học cũng như điểm số. Dataset này rất phù hợp cho các mục đích phân tích dữ liệu giáo dục (Educational Data Mining), thống kê, biểu đồ hóa, hoặc xây dựng các bài tập thực hành về cơ sở dữ liệu.

## 📂 Cấu trúc Dữ liệu (Schema)
Tập dữ liệu bao gồm các thông tin chi tiết với cấu trúc như sau:

| Tên Cột | Mô tả chi tiết | Kiểu dữ liệu | Ví dụ |
| :--- | :--- | :--- | :--- |
| `MSSV` | Mã số sinh viên (Định danh duy nhất) | `String` | S1800005, B1706575 |
| `Ho` | Họ và tên đệm của sinh viên | `String` | Le Duy, Pham Hai |
| `Ten` | Tên của sinh viên | `String` | Anh, Do |
| `Ten lop` | Mã lớp sinh hoạt / lớp học phần | `String` | DI18V7A5, DI16Z6A2 |
| `Ma Nganh` | Mã số chuyên ngành đào tạo | `String` | 7480201, 7480101 |
| `Ten Nganh`| Tên ngành đào tạo (VD: CNTT, Khoa học Máy tính) | `String` | Information Technology |
| `MaMH` | Mã số học phần / môn học | `String` | CT179, CT200 |
| `TenMH` | Tên môn học | `String` | System Administration |
| `Tin Chi` | Trọng số tín chỉ của học phần | `Integer` | 2, 3, 4 |
| `Diem` | Điểm số đánh giá hệ chữ (A, B, C, D, F) | `String` | B+, C+, F |

*Lưu ý: Dữ liệu hiện tại chứa thông tin của sinh viên thuộc các khối ngành như Information Technology, Computer Science, Software Engineering và Computer Networks And Communication.*

## 🚀 Hướng dẫn Sử dụng (Usage)

### 1. Phân tích Dữ liệu (Data Analysis bằng Python)
Bạn có thể dễ dàng sử dụng dữ liệu này với thư viện **Pandas** trong Python để thực hiện các thao tác xử lý:

```python
import pandas as pd

# Đọc dữ liệu từ file CSV hoặc Excel
df = pd.read_csv('data/tong_hop.csv') # Thay đổi đường dẫn tới file của bạn

# Hiển thị 5 dòng đầu tiên
print(df.head())

# Thống kê phân bố điểm số theo môn học
diem_theo_mon = df.groupby('TenMH')['Diem'].value_counts().unstack()
print(diem_theo_mon)
```

### 2. Các hướng tiếp cận và khai thác
*   **Thống kê mô tả (Descriptive Statistics):** Phân tích phổ điểm của các môn học nền tảng (như *Principles of Operating Systems* hoặc *System Administration*).
