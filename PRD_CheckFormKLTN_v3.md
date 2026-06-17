# 🎓 CHECKFORM KLTN — Tài liệu PRD & Hướng dẫn sử dụng

> **Phiên bản:** 3.1 — Cập nhật tháng 06/2026  
> **Tổ chức:** Khoa Kinh tế & QTKD, Trường Đại học Thủy Lợi  
> **Repository:** https://github.com/saodonga/formkltn

---

## 📑 PHẦN 1: PRODUCT REQUIREMENTS DOCUMENT (PRD)

### 1. Tổng quan

| Mục | Nội dung |
|---|---|
| **Tên sản phẩm** | CheckForm KLTN — Công cụ kiểm tra định dạng Khóa Luận Tốt Nghiệp |
| **Mục tiêu** | Tự động hóa quá trình chấm điểm và phát hiện lỗi định dạng KLTN/ĐATN theo chuẩn của Trường ĐH Thủy Lợi |
| **Đối tượng** | Giảng viên hướng dẫn, Giáo vụ khoa, Sinh viên tự kiểm tra trước khi nộp bảo vệ |
| **Môi trường** | Web App (Flask + Docker) / Desktop App (Python tkinter) / CLI |
| **Công nghệ** | Python 3, python-docx, openpyxl, Flask, Docker |

---

### 2. Tiêu chuẩn kiểm tra (Dựa trên quy định ĐH Thủy Lợi)

| Tiêu chí | Chuẩn yêu cầu |
|---|---|
| Khổ giấy | A4 (21.0 × 29.7 cm), sai số ±5.5mm |
| Lề trái | 3.0 cm |
| Lề phải | 2.0 cm |
| Lề trên / Dưới | 2.5 cm |
| Font chữ | Times New Roman (toàn bộ văn bản) |
| Cỡ chữ Heading 1 | 14pt, Bold, ALL CAPS |
| Cỡ chữ Heading 2 | 13pt, Bold |
| Cỡ chữ Heading 3 | 13pt, Bold + Italic |
| Cỡ chữ Heading 4 | 13pt, Italic |
| Cỡ chữ nội dung | 13pt |
| Cỡ chữ Caption | 12pt, Italic |
| Giãn dòng nội dung | 1.5 lines |
| Giãn dòng Heading | Single |
| Spacing nội dung | Before 10pt, After 0pt |
| Spacing Heading 1 | Before 24pt, After 24pt |
| Spacing Heading 2-3 | Before 6pt, After 12pt |
| Canh lề nội dung | Justify (đều hai bên) |
| Số trang | Footer, canh giữa; phần mở đầu (i,ii,iii), phần nội dung (1,2,3) |

---

### 3. Các phân hệ tính năng lõi

#### 3.1. Phân tích Cấu trúc & Layout

- Kiểm tra **khổ giấy A4** (21×29.7 cm) với dung sai ±5.5mm
- Kiểm tra **4 lề** chính xác (Trái 3cm, Phải 2cm, Trên/Dưới 2.5cm)
- Kiểm tra **cấu trúc bắt buộc**: Lời cam đoan, Mục lục, Danh mục hình/bảng, Tài liệu tham khảo
- Kiểm tra **đề mục Chương** (Heading 1): phải có ít nhất 2 chương chính
- Kiểm tra **tên chương**: phải có các từ khóa chuẩn (Tổng quan / Thực trạng / Giải pháp)
- Kiểm tra **Section Break** để đánh số trang khác nhau giữa phần mở đầu và nội dung

#### 3.2. Nhận diện Thông tin Trang bìa

- Trích xuất và kiểm tra **tên đề tài** (≥50 ký tự, cỡ chữ 16pt, Bold, IN HOA)
- Phát hiện **placeholder chưa được điền** ("TÊN ĐỀ TÀI KLTN", "HỌ VÀ TÊN")
- Đọc và xác thực **tên Sinh viên & MSSV**
- Nhận diện **Giảng viên hướng dẫn** thông qua Regex, đối chiếu với danh sách GVHD được cấu hình
- Kiểm tra cụm từ cơ quan chủ quản (phân biệt "Bộ Nông nghiệp và Môi trường" với phiên bản cũ)

#### 3.3. Kiểm tra Heading (Tiêu đề)

Duyệt toàn bộ Heading 1→4, kiểm tra từng tiêu chí:
- **Cỡ chữ** (14pt/13pt theo cấp độ)
- **In đậm / Nghiêng** theo đúng quy định
- **ALL CAPS** (chỉ bắt buộc với Heading 1)
- **Canh lề trái** (tất cả heading phải Left)
- **Khoảng cách đoạn** (Before/After đúng từng cấp)
- **Giãn dòng Single** (bắt buộc với Heading)
- Báo lỗi kèm **tên tiêu đề cụ thể** (7 từ đầu) để sinh viên Ctrl+F tìm nhanh

#### 3.4. Kiểm tra Nội dung (Body Text) — Thông minh, bỏ qua trang bìa

> **Thiết kế quan trọng:** Tool **tự động bỏ qua toàn bộ đoạn văn trước Heading 1 đầu tiên** (tức trang bìa và trang bìa phụ) để tránh báo lỗi oan cho các dòng như "BỘ GIÁO DỤC VÀ ĐÀO TẠO", "TRƯỜNG ĐẠI HỌC THỦY LỢI" vốn có định dạng khác quy chuẩn nội dung.

Từ Heading 1 đầu tiên trở đi, tool kiểm tra 4 tiêu chí với style `Body LA / Content / Normal / Body Text`:

| Tiêu chí | Chuẩn | Mức lỗi |
|---|---|---|
| Khoảng cách (Spacing) | Before 10pt, After 0pt, 1.5 lines | ❌ ERROR |
| Canh lề | Justify (đều 2 bên) | ❌ ERROR |
| Thụt đầu dòng | First-line indent = 0 | ⚠️ WARNING |
| Font & cỡ chữ | Times New Roman, 13pt | ⚠️ WARNING |

**Báo cáo chi tiết dòng lỗi:** Với mỗi loại lỗi, tool liệt kê **tối đa 10 đoạn văn bị lỗi đầu tiên** (7 từ đầu của mỗi đoạn) giúp sinh viên Ctrl+F trong Word để tìm và sửa trực tiếp.

**Các đoạn bị bỏ qua hoàn toàn** (không chấm lỗi nội dung):
- Đoạn bắt đầu bằng "Ghi chú" — chú thích nhỏ bên dưới bảng
- Đoạn bắt đầu bằng "Nguồn" / "nguồn" — nguồn số liệu dưới bảng/hình
- **Đoạn trong ô bảng** (`w:tc`) — nội dung cell bảng thường có định dạng riêng, không theo chuẩn nội dung thân bài
- **Đoạn trong hộp text box / box vẽ** (`w:txbxContent`) — nội dung khung vẽ, sơ đồ, hộp tóm tắt có định dạng riêng

#### 3.5. Kiểm tra Caption (Chú thích Hình/Bảng)

- Phát hiện caption qua style `Caption` hoặc pattern text (`Hình x.y`, `Bảng x.y`)
- Kiểm tra **canh giữa**, **12pt**, **Italic**
- Cảnh báo nếu không có caption nào trong toàn tài liệu

#### 3.6. Kiểm tra Trích dẫn & Tài liệu tham khảo

- Đếm số lượng **Tài liệu tham khảo** (tối thiểu 5-10 tài liệu)
- Phát hiện **chuẩn trích dẫn đang dùng**: IEEE ([1], [2]) hoặc APA (Nguyen, 2023)
- Cảnh báo nếu **lẫn lộn 2 chuẩn** (>3 lần dùng chuẩn phụ)
- Phát hiện **trích dẫn nguyên văn** (ngoặc kép) không ghi nguồn
- Cảnh báo **trích dẫn nguyên văn quá dài** (>2 câu hoặc >60 từ)

#### 3.7. Kiểm tra Học thuật & Phát hiện AI Copy

| Dấu hiệu | Mô tả |
|---|---|
| Markdown bold `**text**` | Dấu hiệu copy từ AI (ChatGPT/Claude giữ lại ký tự markdown) |
| ALL CAPS giữa đoạn nội dung | Đoạn viết HOA toàn bộ lạc lõng trong nội dung thường |
| Bold rải rác giữa câu | AI thường bold các từ khóa quan trọng |
| Đại từ nhân xưng "bạn"/"tôi" | Văn phong học thuật không dùng; phổ biến trong bài AI dịch/tạo |
| Lạm dụng viết tắt không khai báo | >10 từ viết tắt mà không có Bảng danh mục từ viết tắt |

#### 3.8. Kiểm tra Đánh số trang

- Phát hiện **field Page Number** trong Footer
- Kiểm tra **canh giữa** của số trang
- Kiểm tra **Section Break** trước Chương 1

---

### 4. Hệ thống Chấm điểm

**Điểm khởi đầu: 100 điểm** (tối thiểu: 0)

| Loại | Trừ điểm |
|---|---|
| ❌ **ERROR** (Lỗi nghiêm trọng) | **-10 điểm** mỗi lỗi |
| ⚠️ **WARNING** (Cảnh báo) | **-3 điểm** mỗi cảnh báo |
| ℹ️ **INFO** (Thông tin) | Không trừ điểm |

| Điểm | Điểm chữ | Đánh giá |
|---|---|---|
| ≥ 90 | A+ | ✅ Đạt tốt |
| ≥ 85 | A | ✅ Đạt tốt |
| ≥ 80 | B+ | ✅ Đạt tốt |
| ≥ 70 | B | ✔ Đạt |
| ≥ 65 | C+ | ⚠ Cần sửa |
| ≥ 55 | C | ⚠ Cần sửa |
| ≥ 50 | D+ | ⚠ Cần sửa |
| ≥ 40 | D | ❌ Không đạt |
| ≥ 20 | E+ | ❌ Không đạt |
| < 20 | E | ❌ Không đạt |

---

### 5. Báo cáo Kết quả — File Excel (.xlsx)

File Excel xuất ra gồm **4 sheet**:

| Sheet | Nội dung |
|---|---|
| **Ghi chú hạn chế** | Các hạn chế của tool + **Hướng dẫn cách tính điểm** (ERROR -10đ, WARNING -3đ) |
| **Tổng hợp** | Bảng tóm tắt từng file: Điểm hệ 100, Điểm chữ, Số lỗi/cảnh báo, Đánh giá |
| **Chi tiết lỗi** | Từng lỗi của từng file: Mức độ, Nhóm lỗi, Mô tả + danh sách 10 dòng trích dẫn, Gợi ý sửa |
| **Thống kê lỗi** | Biểu đồ nhóm lỗi phổ biến nhất |

> **Tự động giãn dòng Excel:** Ô "Mô tả lỗi" tự động phình to theo số dòng danh sách — không bao giờ bị cắt chữ.

---

### 6. Kiến trúc hệ thống

```
┌─────────────────────────────────────────────────────────┐
│                    GIAO DIỆN NGƯỜI DÙNG                  │
│   Web App (Flask)          Desktop App (tkinter)         │
│   http://server:5000       CheckFormKLTN_GUI.exe         │
└────────────────────┬──────────────────┬─────────────────┘
                     │                  │
                     ▼                  ▼
          ┌──────────────────────────────────────┐
          │          check_format_kltn.py         │
          │        (Engine kiểm tra lõi)          │
          │                                       │
          │  KLTNChecker                          │
          │  ├── _check_page_setup()              │
          │  ├── _check_cover_page()              │
          │  ├── _check_structure()               │
          │  ├── _check_font_and_styles()         │
          │  ├── _check_body_text()  ← bỏ trang bìa │
          │  ├── _check_captions()                │
          │  ├── _check_references()              │
          │  ├── _check_citations()               │
          │  ├── _check_abbreviations()           │
          │  ├── _check_ai_copy_anomalies()       │
          │  ├── _check_page_numbers()            │
          │  └── _compute_score()                 │
          └──────────────────┬───────────────────┘
                             │
                             ▼
                   export_excel() → .xlsx
```

---

### 7. Web App — Tính năng đặc biệt

- **CAPTCHA** tích hợp (phép toán ngẫu nhiên) chống spam
- **Background worker + SSE**: Thanh tiến độ real-time không cần refresh trang
- **Multi-file upload**: Nhiều file .docx cùng lúc, xử lý tuần tự với tiến độ tổng
- **JSON Logging**: Ghi log mỗi lần upload (IP, tên file, thời gian), tự xóa log >30 ngày
- **Thống kê**: Hiển thị tổng số file đã kiểm tra
- **Rate Limiting**: **200 lần/giờ · 30 lần check/phút** mỗi IP
- **Giới hạn upload**: Tối đa 500MB/request

---

### 8. Xử lý trường hợp đặc biệt

| Tình huống | Xử lý |
|---|---|
| File .docx chứa ảnh bị lỗi CRC-32 | Patch zipfile để bỏ qua CRC, không crash |
| Đoạn văn bị mất Style (`para.style = None`) | Skip đoạn đó, tiếp tục kiểm tra bình thường |
| Trang bìa có định dạng khác nội dung | Tự động phát hiện Heading 1 đầu tiên, chỉ kiểm tra từ đó trở đi |
| Đoạn "Ghi chú" / "Nguồn" dưới bảng/hình | Bỏ qua không chấm lỗi spacing/font |
| **Đoạn trong ô bảng (Table Cell)** | **Bỏ qua hoàn toàn** — nội dung bảng không theo chuẩn thân bài |
| **Đoạn trong hộp text box / box vẽ** | **Bỏ qua hoàn toàn** — nội dung box vẽ không theo chuẩn thân bài |
| File biểu mẫu (BanCBHD, BanPBIEN) | Vẫn chạy được, điểm thấp do thiếu cấu trúc KLTN |

---

## 🛠 PHẦN 2: HƯỚNG DẪN SỬ DỤNG

### A. Sử dụng qua Web App

1. Truy cập địa chỉ server (VD: `http://localhost:5000`)
2. Giải **CAPTCHA** (phép toán ngẫu nhiên) — nút "Bắt đầu kiểm tra" sáng lên khi đúng
3. **Upload file** `.docx` (có thể kéo thả, hỗ trợ nhiều file cùng lúc)
4. **Xem kết quả** real-time với thanh tiến độ — bao gồm điểm, danh sách lỗi chi tiết, gợi ý sửa
5. Nhấn **"Xuất Excel"** để tải báo cáo `.xlsx`

### B. Sử dụng qua CLI

```bash
# Kiểm tra 1 file
python check_format_kltn.py "path/to/kltn.docx"

# Quét cả thư mục
python check_format_kltn.py "path/to/folder/"

# Chọn file qua hộp thoại
python check_format_kltn.py
```

Kết quả in ra console và xuất file Excel tự động cùng thư mục chứa file `.docx`.

### C. Cấu hình Giảng viên hướng dẫn

Chỉnh sửa `config_kltn.json`:
```json
{
  "advisors": ["TS Nguyễn Văn A", "ThS Trần Thị B", "PGS.TS Lê Văn C"]
}
```

### D. Deploy bằng Docker

```bash
# Lần đầu
git clone https://github.com/saodonga/formkltn.git
cd formkltn
docker compose up -d --build

# Cập nhật code mới
git pull && docker compose up -d --build
```

---

### E. Những gì tool KHÔNG kiểm tra được

> Giảng viên và Sinh viên cần **tự kiểm tra mắt** các nội dung sau:

1. **Ngắt trang giữa hình/bảng** — Không phát hiện được nếu bảng/hình bị chia cắt sang trang khác
2. **Vị trí Caption** — Chú thích Bảng phải ở PHÍA TRÊN, chú thích Hình phải ở PHÍA DƯỚI
3. **Chỉ mục IEEE chính xác** — Khớp [1], [2] trong văn bản với danh mục tài liệu tham khảo

---

*Tài liệu cập nhật tháng 06/2026. Liên hệ nhóm phát triển qua GitHub Issues.*
