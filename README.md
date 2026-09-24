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
