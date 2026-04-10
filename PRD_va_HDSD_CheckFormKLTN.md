# 🎓 TÀI LIỆU CÔNG CỤ CHECKFORM KLTN
*(Bao gồm PRD & Hướng dẫn sử dụng)*

---

## 📑 PHẦN 1: PRODUCT REQUIREMENTS DOCUMENT (PRD)

### 1. Tổng quan dự án (Overview)
* **Tên sản phẩm:** CheckForm KLTN (Graduation Thesis Format Checker).
* **Mục tiêu:** Công cụ Desktop App tích hợp thuật toán xử lý dữ liệu tự động, giúp chấm điểm và bắt lỗi định dạng Khóa Luận / Đồ Án Tốt Nghiệp (ĐATN/KLTN) dựa trên các quy chuẩn văn bản học thuật (ĐH Thủy Lợi).
* **Đối tượng sử dụng:** Giảng viên duyệt khóa luận, Giáo vụ khoa, hoặc Sinh viên tự kiểm tra trước khi nộp bảo vệ.
* **Môi trường:** Python 3 (tkinter, python-docx, openpyxl).

### 2. Các phân hệ tính năng lõi (Core Modules)

#### 2.1. Phân tích Cấu trúc & Layout (Page Setup)
* Ép chuẩn khổ giấy A4 (Rộng: 21cm, Cao: 29.7cm).
* Lề giấy chuẩn xác: Trái 3.0cm | Phải 2.0cm | Trên 2.5cm | Dưới 2.5cm.
* Cấu trúc khung bài: Đòi hỏi bắt buộc sự có mặt của 4 phần: *Lời cam đoan, Mục lục, Danh mục, Tài liệu tham khảo*.

#### 2.2. Nhận diện Thông tin Trang bìa (Metadata Extraction)
* Kiểm tra tên Cơ quan chủ quản: App kiểm tra và sẽ cảnh báo, yêu cầu chuyển cụm từ "Bộ nông nghiệp và ptnt" thành "Bộ nông nghiệp và môi trường" theo đúng quy định.
* Đọc và trích xuất Tên đề tài (yêu cầu $\ge$ 50 kí tự, không dùng ngoặc chữ). Phát hiện và chặn lỗi chưa thay thế cụm từ "TÊN ĐỀ TÀI KLTN" (placeholder).
* Đọc tên Sinh viên ($<$ 35 kí tự). Tự động báo lỗi nếu chưa điền tên thật mà vẫn để cụm từ placeholder "HỌ VÀ TÊN".
* Kiểm tra định dạng trang bìa: Tên đề tài phải là 16pt, IN HOA, In đậm; Họ tên SV phải là 14pt, IN HOA.
* Nhận dạng Giảng viên hướng dẫn: Đối soát chuỗi bằng Regex dựa trên Danh sách GV được nạp trong Database cấu hình. Trả về Lỗi nếu tên GV sai hoặc thiếu chức danh.

#### 2.3. Quét hệ thống Tiêu đề và Văn bản (Formatting Parser)
* Thuật toán bóc tách Node Heading 1 đến 4:
  * Cảnh báo nếu bài viết thiếu các chương tiêu chuẩn (Tổng quan, Thực trạng, Giải pháp) hoặc có dưới 2 chương.
  * **Heading 1:** Bold, ALL CAPS, 14pt, Before 24pt, After 24pt, Single, Left.
  * **Heading 2:** Bold, 13pt, Before 6pt, After 12pt, Single, Left.
  * **Heading 3:** Bold + Italic, 13pt, Before 6, After 12.
  * **Heading 4:** Italic, 13pt, Before 6, After 12.
* Thuật toán quy chuẩn Body Text:
  * Hoàn toàn không thụt lề dòng đầu. Font Times New Roman 13pt (bắt buộc cả bài).
  * Spacing Before 10pt, After 0pt, Giãn dòng 1.5 lines, canh đều Justify. Cảnh báo lỗi NGHIÊM TRỌNG (ERROR) nếu văn bản căn lề trái hoặc lộn xộn.
* Caption Hình/Bảng: Canh giữa, 12pt Italic, Spacing 6pt/6pt. Phải bắt đầu bằng chữ "Hình" hoặc "Bảng".

#### 2.4. Công cụ học thuật & Nhận diện AI (Academic Integrity)
* Đếm số File trích dẫn thực tế, tối thiểu 3-5 tài liệu. Cảnh báo lỗi trích dẫn lậu (có ngoặc kép copy nguyên văn nhưng không có cite dạng IEEE / APA).
* Dò vi phạm dùng từ Viết tắt nếu không có Bảng danh mục từ viết tắt đính kèm.
* **Thuật toán quét đạo văn AI:**
  * Dò ký tự đánh dấu thừa `**` đặc trưng thường có khi sinh viên Copy từ AI text (ChatGPT / Claude).
  * Khóa lạm dụng ALL CAPS cả câu (nghi vấn lỗi format copy).
  * Khóa văn phong tóm tắt rập khuôn: Nhiều đoạn văn bắt đầu bằng cụm text in đậm (Bold) kết thúc với dấu `:` liệt kê.
  * Nhiều đoạn có nhiều cụm từ được bôi đậm nằm giữa trong đoạn văn nhằm nhấn mạnh nhận định nào đó >> có khả năng sử dụng AI.
  * Nhận diện văn phong phi học thuật / dịch máy (AI), kể từ Chương 1 trở đi, Không tính trong phần lời cam đoan: Theo dõi và cảnh báo số lượng từ nhân xưng "bạn", "tôi" nằm trong nội dung thân bài (kể từ Chương 1 trở đi).

#### 2.5. Cơ chế Chấm điểm & Đánh giá (Scoring)
* Chấm điểm KPI cơ sở hệ 100 với thang phạt phân loại (ERROR = -10đ, WARNING = -3đ).
* Tự động quy đổi và cấp "Điểm chữ" (tương ứng với thang điểm đánh giá đồ án), hỗ trợ tích cực cho giảng viên dựa vào để xếp loại. Hệ thống chỉ ghi nhận đơn thuần giá trị điểm chữ trên cột (A+, A, B+, B,...). Cơ chế quy đổi tự động từ Hệ 100 theo 10 bậc:
  * Điểm $\ge$ 90 $\rightarrow$ **A+**
  * Điểm $\ge$ 85 $\rightarrow$ **A**
  * Điểm $\ge$ 80 $\rightarrow$ **B+**
  * Điểm $\ge$ 70 $\rightarrow$ **B**
  * Điểm $\ge$ 65 $\rightarrow$ **C+**
  * Điểm $\ge$ 55 $\rightarrow$ **C**
  * Điểm $\ge$ 50 $\rightarrow$ **D+**
  * Điểm $\ge$ 40 $\rightarrow$ **D**
  * Điểm $\ge$ 20 $\rightarrow$ **E+**
  * Dưới 20 $\rightarrow$ **E**
* Tự động sắp xếp mức biểu giá đánh giá: Đạt tốt (✅), Đạt (✔), Cần sửa (⚠) và Không đạt (❌).

#### 2.6. UI/UX Interface (Giao diện)
* Multi-threading scanning (Tránh lag/treo app). Có nút Hủy để chặn đứng luồng CPU ngay lập tức khi đang quét dở.
* Xuất báo cáo Log cho người dùng theo chuẩn Excel (.xlsx). Thích hợp chạy thống kê cho cả trường (bao gồm cả cột Hệ 100 và Điểm chữ).
* Quản lý động danh sách Giảng viên Hướng dẫn ngay trên GUI mà không cần chọc vào file code.
* UI thiết kế theo dải màu Light Mode (Modern Web Color) với độ tương phản tốt, đẹp và tươi sáng, hỗ trợ trải nghiệm thị giác. Bảng list kết quả tích hợp cột "Điểm chữ" để review nhanh mức độ.

---

## 🛠 PHẦN 2: HƯỚNG DẪN SỬ DỤNG GIAO DIỆN SẢN PHẨM (USER MANUAL)

### Bước 1: Khởi động phần mềm
* Mở terminal hoặc CMD tại thư mục chứa mã nguồn dự án. Chạy lệnh: `python gui_check_kltn.py`.
* Giao diện CheckForm sẽ hiển thị cùng các thanh phím chức năng.

### Bước 2: Thiết lập Giảng viên hướng dẫn (GVHD)
*(Lưu ý: Bạn nên làm việc này ở ngày đầu tiên ứng dụng tool, để hệ thống không chấm báo lỗi oan Giảng viên)*.
1. Từ Excel danh sách Cán bộ giáo viên mà Khoa biên soạn gửi xuống. Hãy Copy cột Tên giáo viên (Vd: `TS Triệu Đình Phương`, `ThS Đỗ Nguyệt Minh`).
2. Trên màn hình App Tool, ấn vô nút **"👥 Cấu hình GVHD"** (Góc trên bên phải thanh Toolbar).
3. Hộp thoại cấu hình mở ra, Paste toàn bộ danh sách vừa copy vào ô màn hình nội dung (Lưu ý Format 1 người / 1 dòng).
4. Nhấn **"💾 Lưu cấu hình"**. Phần mềm sẽ tự động cất vào Database chuẩn đễ tự động dò danh sách ở tất cả đồ án sinh viên.

### Bước 3: Đưa KLTN / ĐATN vào máy quét
1. Cick nút **"📄 Chọn File .docx"** nếu chỉ muốn kiểm tra 1, 2 bài làm ngẫu nhiên.
2. Click **"📂 Chọn Thư mục"** nếu muốn nạp hàng loạt (Rất hữu ích khi có 1 Folder ZIP chứa vài chục đồ án cuối kỳ). Máy sẽ thống kê sơ bộ số học số lượng docx chuẩn bị được đẩy vào xử lý.

### Bước 4: Chạy kiểm toán định dạng & Hủy khẩn cấp
* Nhấn **"▶ Bắt đầu kiểm tra"**. Phần mềm sẽ tự chia luồng để đọc node XML của từng bản word rồi hiển thị thanh progress bar (tiến trình). Quá trình diễn ra tương đối gọn nhẹ.
* Mỗi khi được 1 file, bảng cây danh mục phía cột trái sẽ nhảy dấu tick `✔` báo cáo kết quả điểm.
* **Tính năng STOP đột ngột**: Cảm thấy lỡ chọn nhầm thư mục chứa quá nhiều file lỗi, bấm ngay vào nút đỏ **"⏹ Hủy / Dừng"**, thuật toán sẽ khóa luồng đọc và nhả giao diện ngay lập tức mà không gây đứng hình.

### Bước 5: Đọc lỗi & Sửa bài
Click vào từng file đồ án ở danh sách cột trái. Nửa menu bên phải giao diện sẽ trích xuất thành bảng mô tả chi tiết list lỗi. Được phân phối theo cường độ ở 3 ngăn Tab:
* ❌ **Lỗi (ERROR):** Cực kì nghiêm trọng (vd. Sai khổ giấy nộp, Thiếu trang Mục lục, Giãn dòng sai số trầm trọng).
* ⚠️ **Cảnh báo (WARNING):** Lỗi hình thức, lạm dụng copy AI, không canh Justify cho đoạn, Heading sai màu... Hãy click vào thông báo để xem **"Khuyến nghị Hướng dẫn sửa lỗi bằng Phím tắt Word"** dưới mỗi thẻ card.
* ℹ **Thông tin (INFO):** Thống kê số lượng trang, GV, số file...

### Bước 6: Tập hợp Báo cáo Hồ sơ
* Nhấn nút màu cam **"💾 Xuất Excel"**.
* Chỉ định nơi Save để hệ thống In báo cáo. Bạn có thể gửi file Excel này trực tiếp vào Group Lớp / Khoa để các sinh viên tự đối chiếu kết quả điểm bị trừ cùng list vi phạm và nộp lại File đã được fix.
