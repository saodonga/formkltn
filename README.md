# Công cụ Kiểm tra Định dạng KLTN

Kiểm tra tự động các file `.docx` của sinh viên xem có đáp ứng yêu cầu trình bày Khóa Luận Tốt Nghiệp (KLTN/ĐATN) của Trường Đại học Thủy Lợi hay không.

## Yêu cầu cài đặt

```bash
pip install python-docx openpyxl
```

(Công cụ sẽ tự cài nếu thiếu khi chạy lần đầu)

## Cách dùng

### Kiểm tra một file duy nhất
```bash
python check_format_kltn.py "path/to/kltn_sinh_vien.docx"
```

### Quét toàn bộ thư mục
```bash
python check_format_kltn.py "path/to/thu_muc_nop_bai/"
```

### Chạy không có đối số → Hiện hộp thoại chọn file/thư mục
```bash
python check_format_kltn.py
```

## Tiêu chuẩn kiểm tra

| Hạng mục | Chuẩn yêu cầu |
|---|---|
| Khổ giấy | A4 (21 × 29.7 cm) |
| Lề trái | 3.0 cm |
| Lề phải | 2.0 cm |
| Lề trên / Dưới | 2.5 cm |
| Font chữ | Times New Roman |
| Cỡ chữ Heading 1 (Chương) | 14pt, **đậm**, IN HOA |
| Cỡ chữ Heading 2 | 13pt, **đậm** |
| Cỡ chữ Heading 3 | 13pt, **đậm + nghiêng** |
| Cỡ chữ Heading 4 | 13pt, *nghiêng* |
| Cỡ chữ nội dung (Body) | 13pt |
| Cỡ chữ Caption | 12pt, *nghiêng* |
| Giãn dòng nội dung | 1.5 lines |
| Canh lề nội dung | Justify (đều hai bên) |
| Cấu trúc bắt buộc | Lời cam đoan, Mục lục, Danh mục, TLTK |
| Trích dẫn | Kiểu IEEE ([1], [2], …) |

## Kết quả đầu ra

Công cụ xuất file Excel (`KIEM_TRA_DINH_DANG_KLTN_YYYYMMDD_HHMM.xlsx`) gồm 3 sheet:

1. **Tổng hợp** — Danh sách tất cả file với điểm số (0–100) và đánh giá Đạt / Không đạt
2. **Chi tiết lỗi** — Liệt kê từng lỗi, mức độ (ERROR / WARNING), vị trí và gợi ý sửa
3. **Thống kê lỗi** — Nhóm lỗi phổ biến nhất để giảng viên nắm tổng quan

### Thang điểm
| Điểm | Đánh giá |
|---|---|
| 90–100 | ✅ Đạt tốt |
| 70–89  | ✔ Đạt |
| 50–69  | ⚠ Cần sửa |
| 0–49   | ❌ Không đạt |

Mỗi lỗi nghiêm trọng (ERROR) trừ 10 điểm, cảnh báo (WARNING) trừ 3 điểm.

## File trong thư mục

| File | Mô tả |
|---|---|
| `check_format_kltn.py` | **Công cụ kiểm tra** (file chính) |
| `scan_kltn.py` | Quét & trích xuất danh sách sinh viên ra Excel |
| `1. Huong dan trinh bay DATN.docx` | Hướng dẫn trình bày KLTN |
| `2026_Mau_viet_chuan_KLTN.docx` | File mẫu chuẩn cho sinh viên |
