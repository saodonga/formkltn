# PRD-12: Hệ thống Chấm Điểm Quá trình KLTN

**EduFlow · v1.8 · Cập nhật 06/04/2026**

> **Phụ thuộc:** PRD-00, PRD-02 (KLTN), PRD-04 (Bộ môn cấu hình)  
> **Sprint:** Sprint 7 — sau khi hoàn thành Sprint 4 (KLTN cơ bản)  
> **Roles liên quan:** LANH_DAO_BO_MON, GIANG_VIEN, SINH_VIEN

---

## Changelog

| Phiên bản | Ngày | Thay đổi |
|-----------|------|----------|
| v1.0 | 20/03/2026 | Phát hành ban đầu |
| v1.1 | 20/03/2026 | Clarification mốc thời gian A+/A |
| v1.2 | 26/03/2026 | Thêm hệ thống Lộ trình nộp bài (BM → GV → SV) |
| **v1.3** | 26/03/2026 | Quy tắc tên file mới · Phân quyền GV · Tích hợp lộ trình vào form SV |
| **v1.4** | 28/03/2026 | Admin/LĐBM Review Dashboard · Sắp xếp & Xuất Excel · Fix Ý thức SV chưa vào nhóm |
| **v1.5** | 29/03/2026 | Chuẩn hóa Múi giờ GMT+7 · Nâng cấp Excel chuyên nghiệp (ExcelJS - Times New Roman 12) |
| **v1.6** | 30/03/2026 | Tách file ChamDiemTuanGv · Cột LV editable (popover kiểu YT) · History dropdowns NX T-1/T-2 · API lĩnh vực với fallback boMon |
| **v1.7** | 02/04/2026 | Tìm kiếm MSSV/Họ tên · Tối ưu Toolbar (TH, icon-only) · Stats Bar hợp nhất 1 dòng · Fix bug overwrite null NX khi chấm CM |
| **v1.8** | **06/04/2026** | **Fix logic auto-select tuần mặc định (so sánh timestamp đầy đủ giờ:phút, không chỉ thứ) · Fix nút "Làm mới" header hoạt động đúng** |

---

## 1. Tổng quan

Module bổ sung hệ thống **chấm điểm quá trình KLTN theo tuần** (14 tuần), bao gồm:

1. **Thang điểm chữ cái** — mỗi Bộ môn cấu hình riêng (tối đa 10 bậc, dynamic)
2. **Lộ trình nộp bài** — LĐBM thiết lập → GV xác nhận/tùy chỉnh → SV xem
3. **Điểm Ý thức (30%)** — tự động gán theo thời gian nộp so với mốc lộ trình
4. **Điểm Chuyên môn (70%)** — GV chấm thủ công
5. **Tuần 14 đặc biệt** — thu quyển báo cáo cứng (3 hạng mục)
6. **Sự kiện đánh giá bổ sung** — LĐBM thêm sự kiện ngoài 14 tuần
7. **Hệ thống Giám sát Admin/LĐBM** — Dashboard xem toàn bộ SV/Nhóm trong đợt, lọc theo Nhóm.
8. **Công cụ Bảng điểm** — Sắp xếp đa tiêu chí, Xuất Excel (.xlsx), Thanh thống kê tuần (Stats Bar).

---

## 2. Thang điểm Chữ cái (Dynamic Per Bộ môn)

### 2.1 Cấu hình

Cấu hình trong **Cài đặt Bộ môn** (`/bo-mon/[id]/cau-hinh` → Tab **⭐ Thang điểm**).

**Bảng default gợi ý (10 bậc):**

| Bậc | Ký hiệu | Điểm số | Ghi chú |
|-----|---------|---------|---------|
| 1 | A+ | 9.0 | Xuất sắc |
| 2 | A | 8.5 | Giỏi |
| 3 | B+ | 8.0 | Khá+ |
| 4 | B | 7.0 | Khá |
| 5 | C+ | 6.5 | Trung bình+ |
| 6 | C | 5.5 | Trung bình |
| 7 | D+ | 5.0 | Trung bình yếu+ |
| 8 | D | 4.0 | Yếu |
| 9 | E+ | 2.0 | Kém+ |
| 10 | E | 0.0 | Kém/Không nộp |

### 2.2 Schema DB: `thang_diem_bo_mon`

```sql
CREATE TABLE thang_diem_bo_mon (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  bo_mon_id    UUID NOT NULL REFERENCES bo_mon(id) ON DELETE CASCADE,
  ky_hieu      VARCHAR(5) NOT NULL,     -- 'A+', 'A', 'B+', ...
  diem_so      DECIMAL(4,2) NOT NULL,  -- 9.0, 8.5, 8.0, ...
  mo_ta        TEXT,
  thu_tu       INTEGER NOT NULL,       -- thứ tự giảm dần (1=cao nhất)
  is_active    BOOLEAN DEFAULT true,
  created_at   TIMESTAMP DEFAULT NOW(),
  UNIQUE(bo_mon_id, ky_hieu),
  UNIQUE(bo_mon_id, thu_tu)
);
```

---

## 3. Hệ thống Lộ trình Nộp bài ✅ TRIỂN KHAI XONG

### 3.1 Luồng nghiệp vụ

| STT | Tính năng | Mô tả |
|-----|----------|-------|
| 1 | Lộ trình | LĐBM thiết lập lộ trình 14 tuần (mỗi tuần: thứ hạn, giờ, tên mốc, file mẫu, nội dung yêu cầu) |
| 2 | Xác nhận | GV xem lộ trình BM → Xác nhận áp dụng hoặc Tùy chỉnh từng tuần cho nhóm mình |
| 3 | Xem | SV xem lộ trình đã xác nhận: ngày cụ thể (dd/mm/yyyy), nội dung cần đạt, tên file cần nộp |
| 4 | Theo dõi làm việc | GV/SV ghi nhận thời gian, hình thức (Online/Offline) và tình trạng tham dự (Muộn, Vắng, Đúng giờ...) |

### 3.2 Schema DB

#### `kltn_lo_trinh_bo_mon` — Template mặc định của Bộ môn

```sql
CREATE TABLE kltn_lo_trinh_bo_mon (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    dot_kltn_id UUID NOT NULL REFERENCES dot_kltn(id) ON DELETE CASCADE,
    so_tuan INTEGER NOT NULL,           -- 1..14
    thu_han_thu INTEGER NOT NULL,       -- 0=CN..6=T7 (0-indexed theo JS getDay())
    thu_han_gio INTEGER NOT NULL,
    thu_han_phut INTEGER NOT NULL DEFAULT 0,
    ten_moc TEXT NOT NULL,              -- VD: "Hoàn thiện Đề cương KLTN"
    ten_file_mau TEXT,                  -- VD: "KLTN_HK2_2025_<MSSV>_<HoTen>_Đề cương.docx"
    yeu_cau_noi_dung TEXT,              -- Nội dung cần đạt chi tiết
    ghi_chu TEXT,
    created_at TIMESTAMP DEFAULT NOW() NOT NULL,
    updated_at TIMESTAMP DEFAULT NOW() NOT NULL,
    UNIQUE(dot_kltn_id, so_tuan)
);
```

#### `kltn_lo_trinh_nhom` — Tùy chỉnh của Nhóm / GV

```sql
CREATE TABLE kltn_lo_trinh_nhom (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    nhom_kltn_id UUID NOT NULL REFERENCES nhom_kltn(id) ON DELETE CASCADE,
    so_tuan INTEGER NOT NULL,   -- 0 = row xác nhận tổng thể; 1..14 = tùy chỉnh từng tuần
    thu_han_thu INTEGER,        -- null = kế thừa từ BM
    thu_han_gio INTEGER,
    thu_han_phut INTEGER,
    ten_moc TEXT,
    ten_file_mau TEXT,
    yeu_cau_noi_dung TEXT,
    ghi_chu TEXT,
    da_xac_nhan BOOLEAN NOT NULL DEFAULT FALSE,
    xac_nhan_boi UUID REFERENCES users(id),
    xac_nhan_luc TIMESTAMP,
    created_at TIMESTAMP DEFAULT NOW() NOT NULL,
    updated_at TIMESTAMP DEFAULT NOW() NOT NULL,
    UNIQUE(nhom_kltn_id, so_tuan)
);
```

> **Row `so_tuan = 0`**: Trạng thái xác nhận tổng thể của nhóm (GV đã confirm apply lộ trình BM).

### 3.3 API Endpoints Lộ trình

#### Bộ môn (LĐBM)

| Method | Path | Mô tả |
|--------|------|--------|
| GET | `/api/kltn/lo-trinh?dotKltnId=<id>` | Lấy lộ trình BM |
| POST | `/api/kltn/lo-trinh` | Upsert toàn bộ lộ trình (batch) |
| PATCH | `/api/kltn/lo-trinh?dotKltnId=<id>&soTuan=<n>` | Sửa 1 mốc tuần |
| POST | `/api/kltn/lo-trinh/seed-default` | Seed 14 tuần mặc định |

#### Nhóm / GV

| Method | Path | Mô tả |
|--------|------|--------|
| GET | `/api/kltn/lo-trinh-nhom?nhomKltnId=<id>` | Lộ trình merge (BM + custom), kèm `ngayBatDau`, `soTuanThucHien` |
| POST | `/api/kltn/lo-trinh-nhom` | Xác nhận (`action: 'xac-nhan'`) hoặc tùy chỉnh |
| PATCH | `/api/kltn/lo-trinh-nhom?nhomKltnId=<id>&soTuan=<n>` | Tùy chỉnh 1 tuần |
| DELETE | `/api/kltn/lo-trinh-nhom?nhomKltnId=<id>&soTuan=<n>` | Reset tuần về BM |

### 3.4 UI Components

| Component | Vai trò | Vị trí |
|-----------|---------|--------|
| `LoTrinhBoMonConfig.tsx` | LĐBM thiết lập, nút **"Xác nhận & Áp dụng"** trigger tính lại điểm hàng loạt | `/bo-mon/[id]/cau-hinh` → Tab Lộ trình |
| `LoTrinhNhomPanel.tsx` | GV xem, xác nhận, tùy chỉnh từng tuần | Modal nhóm KLTN (nút 📅) |
| `LoTrinhSvView.tsx` | SV xem lộ trình read-only với ngày cụ thể, countdown, badge trạng thái | Tab `lo-trinh` trong `SinhVienKLTNView` |

### 3.5 Phân quyền Lộ trình (quan trọng)

| Role | API `/api/nhom-kltn` | Nút 📅 Lộ trình |
|------|----------------------|-----------------|
| LANH_DAO_BO_MON / SYSTEM_ADMIN | Thấy **tất cả nhóm** | Thông qua `LoTrinhBoMonConfig` |
| GIANG_VIEN | Chỉ thấy **nhóm mình làm GVHD** | ✅ Hiện trên mỗi nhóm |
| SINH_VIEN | Chỉ thấy **nhóm của mình** | Xem qua tab lộ trình trong trang SV |

---

## 4. Logic Tính Điểm Ý thức (Tuần 1–14)

### 4.1 Mốc thời gian — DYNAMIC từ lộ trình nhóm

> **Thay đổi v1.2**: Mốc không còn hardcode "8:00 Thứ 5" mà lấy từ `kltn_lo_trinh_nhom` (hoặc `kltn_lo_trinh_bo_mon` nếu nhóm chưa tùy chỉnh).

Mỗi tuần có mốc riêng: `thuHanThu`, `thuHanGio`, `thuHanPhut`.

**Tính ngày cụ thể** cho tuần N:
```typescript
// ngayBatDau = dot_kltn.gd3Tu (ngày bắt đầu giai đoạn 3)
function tinhNgayDeadline(soTuan, ngayBatDau, thuHanThu, thuHanGio, thuHanPhut): Date {
    const tuanBatDau = new Date(ngayBatDau)
    tuanBatDau.setDate(tuanBatDau.getDate() + (soTuan - 1) * 7)
    let delta = thuHanThu - tuanBatDau.getDay()
    if (delta < 0) delta += 7
    const deadline = new Date(tuanBatDau)
    deadline.setDate(deadline.getDate() + delta)
    deadline.setHours(thuHanGio, thuHanPhut, 0, 0)
    return deadline
}
```

### 4.2 Thuật toán tính điểm Ý thức ✅ CONFIRMED

```typescript
// File: lib/kltn/diem-y-thuc.ts → hàm tinhDiemYThuc()
// Chuẩn hóa GMT+7: Toàn bộ so sánh thực hiện trên hệ UTC để đồng bộ Server/Client

const diffMs = ngayNop.getTime() - moc.getTime()
const diffHours = diffMs / (1000 * 60 * 60)

// TRƯỚC MỐC
if (diffMs < 0) {
    // A+: nộp sớm ≥ 8h trước mốc  (diffHours ≤ -8)
    // A : nộp trong vòng 8h trước mốc (-8 < diffHours < 0)
    const idx = (-diffHours >= 8) ? 0 : 1  // 0=A+, 1=A
    return { kyHieu: bacs[idx].kyHieu, ... }
}

// SAU MỐC: hạ 1 bậc mỗi 4h (bắt đầu từ bậc B+ = index 2)
const bacIndex = Math.min(2 + Math.floor(diffHours / intervalHours), bacs.length - 1)
```

**Bảng tham chiếu (thang điểm 10 bậc mặc định):**

| Thời điểm nộp | Bậc | Điểm | Logic |
|--------------|-----|------|-------|
| Nộp sớm ≥ 8h trước mốc | **A+** | 9.0 | `diffHours ≤ -8` |
| Nộp trong vòng 8h trước mốc | **A** | 8.5 | `-8 < diffHours < 0` |
| Trễ 0–4h | B+ | 8.0 | |
| Trễ 4–8h | B | 7.0 | |
| Trễ 8–12h | C+ | 6.5 | |
| Trễ 12–16h | C | 5.5 | |
| Trễ 16–20h | D+ | 5.0 | |
| Trễ 20–24h | D | 4.0 | |
| Trễ 24–28h | E+ | 2.0 | |
| Trễ ≥ 28h hoặc không nộp | E | 0.0 | |

> GV có thể **ghi đè toàn bộ điểm ý thức** nếu chất lượng chuyên môn không đạt.

---

## 5. Quy tắc Tên File Báo cáo ✅ CẬP NHẬT v1.3

### 5.1 Format chuẩn

| Tuần | Format tên file |
|------|----------------|
| Tuần 1–13 | `KLTN_{TenDot}_{MSSV}_{Họ và tên SV}_Tuan {N}.docx` |
| Tuần cuối (14) | `KLTN_{TenDot}_{MSSV}_{Họ và tên SV}_Bao cao KLTN.docx` |

**Ví dụ:**
```
KLTN_HK2_2025_2254105227_Nguyễn Phương Anh_Tuan 1.docx
KLTN_HK2_2025_2254105227_Nguyễn Phương Anh_Bao cao KLTN.docx
```

> **Lưu ý quan trọng:**
> - `{TenDot}` = `dot_kltn.tenDot` (ví dụ: `HK2_2025`, `KTS.KLTN.25-26`)
> - `{Họ và tên SV}` **giữ nguyên dấu tiếng Việt** (chỉ bỏ ký tự đặc biệt như `/\:*?"<>|`)
> - Nếu BM thiết lập `ten_file_mau` trong lộ trình → **ưu tiên hoàn toàn** (override quy tắc generic)
> - Validate tuần cuối: chấp nhận cả `_Bao cao KLTN.docx` lẫn `_Tuan 14.docx`

### 5.2 Luồng validation phía client

```
1. Kiểm tra extension: chỉ .docx → lỗi nếu sai
2. Kiểm tra tên file theo format chuẩn (có tenFileMauLoTrinh nếu BM đã cấu hình)
3. Nếu tên sai → Dialog cảnh báo:
   "Tên chuẩn: KLTN_HK2_2025_2254105227_Nguyễn Phương Anh_Tuan 3.docx"
   → [Hủy - đặt lại tên]  [Đồng ý & nộp → hạ 1 bậc điểm]
```

### 5.3 Hiển thị gợi ý tên file trong Form

Form `BaoCaoTuanKlForm` hiển thị theo thứ tự ưu tiên:
1. **`tenFileMauLoTrinh`** (BM/GV cấu hình trong lộ trình) — ưu tiên cao nhất
2. **`buildStandardFileName(mssv, hoTen, soTuan, isLastTuan, tenDot)`** — fallback generic

---

## 6. Màn hình SV: Báo cáo Tuần ✅ TRIỂN KHAI XONG

### 6.1 Component: `BaoCaoTuanKlTab`

Route: `/reports?type=KLTN`

### 6.2 Thông tin hiển thị từ lộ trình

Mỗi card tuần (khi mở rộng) hiển thị:

```
┌─ Tuần 1 — Hoàn thiện Đề cương KLTN ─────────────────┐
│  📅 Hạn nộp: Thứ 5, 27/03/2026 lúc 08:00            │
│  ⏰ Còn 2 ngày 6h                                     │
│                                                        │
│  [⚡ Sắp đến hạn]                                      │
│                                                        │
│  ─── Thông tin nộp bài ───                            │
│  📋 Nội dung cần đạt tuần này                        │
│  "Hoàn thiện đề cương KLTN theo form chuẩn,           │
│   có đủ mục 1-5..."                                   │
│                                                        │
│  📄 Tên file cần nộp                                  │
│  KLTN_HK2_2025_22541xxxxx_NguyễnPhươngAnh_             │
│  Đề cương.docx                                       │
│                                                        │
│  📝 Ghi chú: ...                                     │
│                                                        │
│  ─── Form nộp ───                                     │
│  Tên chuẩn: KLTN_HK2_2025_..._Tuan 1.docx (fallback) │
│  [📁 Chọn file]  [📤 Nộp báo cáo]                    │
└───────────────────────────────────────────────────────┘
```

### 6.3 Màu sắc badge tuần

| Trạng thái | Màu badge tuần | Badge text |
|------------|----------------|------------|
| Tuần tương lai (>7 ngày) | Mờ, không click được | — |
| Sắp đến hạn (≤24h) | 🟡 Vàng animate | ⚡ Sắp đến hạn |
| Đã qua hạn | ⚫ Xám | ✓ Đã qua |
| Tuần cuối | 🔴 Đỏ | — |
| Bình thường | 🟣 Tím | — |

---

## 7. Màn hình GV: Quản lý Lộ trình Nhóm ✅ TRIỂN KHAI XONG

GV truy cập từ nút **📅** trên NhomCard (trang `/groups?type=KLTN`).

### 7.1 Quyền hiển thị

- GV chỉ thấy **nhóm mình làm GVHD** (filter server-side tại `/api/nhom-kltn`)
- Nút 📅 hiện với `isGiangVien = hasRole('GIANG_VIEN') && loai === 'KLTN'`

### 7.2 Chức năng `LoTrinhNhomPanel`

| Chức năng | Mô tả |
|-----------|-------|
| Xem lộ trình BM | Hiển thị 14 mốc với merge BM + custom |
| Xác nhận lộ trình | GV confirm áp dụng lộ trình BM cho nhóm |
| Tùy chỉnh tuần | GV thay đổi thứ hạn/giờ/tên mốc cho từng tuần |
| Reset tuần | Quay về mặc định BM |

---

## 8. Schema DB Chính: `diem_qua_trinh_kl`

```sql
CREATE TABLE diem_qua_trinh_kl (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  sv_nhom_kltn_id     UUID NOT NULL REFERENCES sv_nhom_kltn(id) ON DELETE CASCADE,
  dot_kltn_id         UUID NOT NULL,
  so_tuan             INTEGER NOT NULL,       -- 1..soTuanThucHien
  loai_tuan           loai_tuan_kltn_enum,    -- 'BAO_CAO' | 'THU_QUYEN'

  -- Điểm Ý thức — tự động, GV có thể ghi đè
  y_thuc_ky_hieu      VARCHAR(5),
  y_thuc_diem_so      DECIMAL(4,2),
  y_thuc_tu_dong      BOOLEAN DEFAULT true,   -- false = GV đã ghi đè
  y_thuc_ghi_chu      TEXT,
  y_thuc_ghi_de_boi   UUID REFERENCES users(id),
  y_thuc_ghi_de_luc   TIMESTAMP,

  -- Điểm Chuyên môn (70%) — chỉ tuần 1-14
  chuyen_mon_ky_hieu  VARCHAR(5),
  chuyen_mon_diem_so  DECIMAL(4,2),
  chuyen_mon_cham_boi UUID REFERENCES users(id),
  chuyen_mon_cham_luc TIMESTAMP,
  chuyen_mon_nhan_xet TEXT,

  -- Tuần 1-13: y_thuc*0.3 + chuyen_mon*0.7
  -- Tuần 14  : y_thuc*1.0
  diem_tong_tuan      DECIMAL(4,2),

  trang_thai          trang_thai_diem_qt_enum DEFAULT 'CHUA_CHAM',

  -- Theo dõi làm việc/hướng dẫn (Metric bổ sung v1.7)
  working_at          TIMESTAMP,      -- Thời gian làm việc trong tuần
  working_format      VARCHAR(20),    -- 'ONLINE' | 'OFFLINE'
  working_status      VARCHAR(40),    -- 'DU_GIO', 'MUON_10P', 'XIN_PHEP', ...

  created_at          TIMESTAMP DEFAULT NOW(),
  updated_at          TIMESTAMP DEFAULT NOW(),

  UNIQUE(sv_nhom_kltn_id, so_tuan)
);
```

---

## 9. API Endpoints Chính

### 9.1 Điểm quá trình

| Method | Path | Roles | Mô tả |
|--------|------|-------|-------|
| GET | `/api/kltn/diem-qt?svNhomKltnId={id}` | GV, SV | Điểm từng tuần + tổng kết |
| PATCH | `/api/kltn/diem-qt/[id]/chuyen-mon` | GV, LĐBM | Chấm điểm chuyên môn |
| PATCH | `/api/kltn/diem-qt/[id]/y-thuc-ghi-de` | GV, LĐBM | Ghi đè điểm ý thức |
| POST | `/api/kltn/diem-qt/recalculate` | SYSTEM | Tính lại điểm một đợt |
| POST | `/api/reports/repair-diem-y-thuc` | ADMIN | Bảo trì: Tạo/Repair record điểm QT thiếu |

### 9.2 Báo cáo tuần

| Method | Path | Roles | Mô tả |
|--------|------|-------|-------|
| GET | `/api/bao-cao-tuan/kltn?svNhomKltnId={id}` | GV, SV | DS báo cáo đã nộp |
| POST | `/api/bao-cao-tuan/kltn` | SV | Nộp báo cáo tuần |

---

## 10. Điểm Tổng Cả Đợt (2 loại)

**Loại 1 — Điểm TB cập nhật** (`diem_tb_cap_nhat`):
```typescript
diem_tb_cap_nhat = sum(diem_tong_tuan of completed weeks) / count(completed weeks)
```

**Loại 2 — Điểm TB tích lũy** (`diem_tb_tich_luy`):
```typescript
// Tuần chưa có điểm = 0
diem_tb_tich_luy = sum(diem_tong_tuan ?? 0, for all tuans) / totalItems
```

---

## 11. Trạng thái Triển khai ✅

| # | Module | Trạng thái | Ghi chú |
|---|--------|------------|---------|
| 1 | Schema `thang_diem_bo_mon` | ✅ Xong | Migration đã chạy |
| 2 | Schema `diem_qua_trinh_kl` | ✅ Xong | |
| 3 | Schema `kltn_lo_trinh_bo_mon` + `kltn_lo_trinh_nhom` | ✅ Xong | |
| 4 | API lộ trình BM + Nhóm | ✅ Xong | Kèm `ngayBatDau`, `soTuanThucHien` |
| 5 | UI LĐBM: `LoTrinhBoMonConfig` | ✅ Xong | Nút Xác nhận → trigger tính lại |
| 6 | UI GV: `LoTrinhNhomPanel` | ✅ Xong | Modal từ nút 📅 trên NhomCard |
| 7 | Phân quyền GV: chỉ thấy nhóm mình | ✅ Xong | Filter server-side `/api/nhom-kltn` |
| 8 | UI SV: `LoTrinhSvView` | ✅ Xong | Ngày cụ thể, countdown, badge |
| 9 | UI SV: `BaoCaoTuanKlTab` tích hợp lộ trình | ✅ Xong | Hiển thị nội dung cần đạt, tên file |
| 10 | Logic tính điểm Ý thức (A+/A dynamic) | ✅ Xong | Dựa trên mốc lộ trình, không hardcode |
| 11 | Quy tắc tên file v1.3 | ✅ Xong | `KLTN_{TenDot}_{MSSV}_{HoTen}_...docx` |
| 12 | GV: Chấm điểm Chuyên môn | ✅ Xong | `ChamDiemTuanGv` |
| 13 | **Dashboard Giám sát Admin/LĐBM** | ✅ Xong | `/reports?type=KLTN` (Admin View) |
| 14 | **Sắp xếp & Xuất Excel bảng điểm** | ✅ Xong | Tích hợp SheetJS + Sort logic |
| 15 | **Thống kê Tuần (Stats Bar)** | ✅ Xong | Real-time counts & Progress Bar |
| 16 | **Tách file ChamDiemTuanGv (v1.6)** | ✅ Xong | 4 file riêng: `.tsx`, `.cells.tsx`, `.actions.ts`, `.types.ts` |
| 17 | **Cột LV editable — Popover lĩnh vực** | ✅ Xong | Click → popover danh sách lĩnh vực nhóm theo Cap1/Cap2 |
| 18 | **History Dropdowns NX T-1/T-2** | ✅ Xong | Toggle độc lập, inline trong ô nhận xét, không popup nổi |
| 19 | **API `/api/kltn/linh-vuc`** | ✅ Xong | Fallback sang `boMonId` khi `chuyenNganhId = null` |
| 20 | **API `/api/kltn/de-tai/linh-vuc` (PATCH)** | ✅ Xong | GV cập nhật lĩnh vực qua `svNhomKltnId`, ghi `deTaiLog` |
| 21 | **Fix auto-select tuần (timestamp đầy đủ)** | ✅ Xong (v1.8) |
| 22 | **Fix nút "Làm mới" header kết nối `kltn-refresh`** | ✅ Xong (v1.8) |
| 23 | Thu quyển tuần 14 | 🔜 Pending | |
| 24 | Sự kiện đánh giá bổ sung | 🔜 Pending | |
| 25 | Word Export (Báo cáo tổng hợp) | 🔜 Pending | |

---

## 12. Acceptance Criteria

| ID | Tình huống | Kết quả mong đợi |
|----|-----------|-----------------|
| AC-01 | LĐBM tạo thang điểm 8 bậc | Lưu vào `thang_diem_bo_mon`, hiển thị trong cài đặt |
| AC-02 | SV nộp sớm ≥ 8h trước mốc | Điểm ý thức = A+, kèm ghi chú GV duyệt |
| AC-03 | SV nộp trong vòng 8h trước mốc | Điểm ý thức = A |
| AC-04 | SV nộp trễ 5h sau mốc | Điểm ý thức = B (index 3) |
| AC-05 | SV nộp file tên sai chuẩn | Dialog cảnh báo → chấp nhận → hạ 1 bậc |
| AC-06 | LĐBM cập nhật lộ trình, nhấn "Xác nhận & Áp dụng" | Trigger tính lại điểm ý thức cho toàn bộ SV trong đợt |
| AC-07 | GV xem trang `/groups?type=KLTN` | Chỉ thấy nhóm mình làm GVHD |
| AC-08 | GV nhấn nút 📅 trên nhóm | Mở `LoTrinhNhomPanel`, xem/sửa/xác nhận lộ trình |
| AC-09 | SV xem trang báo cáo tuần | Thấy ngày cụ thể (dd/mm/yyyy), nội dung cần đạt, tên file cần nộp |
| AC-10 | SV xem tuần 1 đã cấu hình "Đề cương" | Thấy tên file `KLTN_HK2_2025_MSSV_HoTen_Đề cương.docx` |
| AC-11 | GV chấm chuyên môn B+ | Cập nhật `chuyen_mon_ky_hieu`, tính lại `diem_tong_tuan` |
| AC-12 | GV ghi đè ý thức từ A+ → E | `y_thuc_tu_dong = false`, ghi log |

---

---

## 13. Hệ thống Giám sát & Báo cáo Admin (MỚI v1.4)

### 13.1 Phân quyền Giám sát (Admin Oversight)

Trang `/reports?type=KLTN` tự động chuyển đổi View dựa trên Role:
- **Giảng viên**: Thấy danh sách sinh viên thuộc các nhóm mình hướng dẫn.
- **Quản lý (Admin/LĐBM/Trưởng ĐV)**: Thấy Dashboard Giám sát cho phép xem toàn bộ sinh viên trong phạm vi quản lý.

| Role | Phạm vi xem | Component |
|------|-------------|-----------|
| `SYSTEM_ADMIN` | Toàn trường (Tất cả Đơn vị/Bộ môn/Đợt) | `KltnAdminView` |
| `DON_VI_TRUONG` | Toàn bộ Bộ môn thuộc Đơn vị mình | `KltnAdminView` |
| `LANH_DAO_BO_MON` | Toàn bộ Nhóm/SV thuộc Bộ môn mình | `KltnAdminView` |

### 13.2 Tính năng Bảng điểm Nâng cao

Component `ChamDiemTuanGv.tsx` được nâng cấp với các công cụ quản lý mạnh mẽ:

1. **Sắp xếp (Multivariate Sorting)**:
   - Click vào Header để Sort: Họ tên, MSSV, Điểm Ý thức, Điểm Chuyên môn, Điểm Tổng tuần.
   - Cycle: Tăng dần → Giảm dần → Trạng thái gốc (theo thứ tự DB).
2. **Xuất Excel Chuyên nghiệp (v1.5 Upgrade)**:
   - Chuyển đổi từ `xlsx` (SheetJS) sang **`exceljs`** để hỗ trợ định dạng nâng cao.
   - **Font chuẩn**: **Times New Roman**, cỡ chữ **12** toàn bộ file.
   - **Header**: Cao 30px, nền xanh đậm `#1E3A5F`, chữ trắng đậm, căn giữa.
   - **Layout**: Khổ A4, Nằm ngang (Landscape), Fit to Width, Freeze dòng đầu.
   - **Zebra Stripes**: Màu nền xen kẽ xanh nhạt `#FFF0F4FF` cho hàng chẵn.
   - **Báo cáo đồng bộ**: Áp dụng chung cho `ChamDiemTuanGv`, `ThesesTable`, `GroupsTable`.
3. **Thanh Thống kê Tuần (Stats Bar)**:
   - Hiển thị ngay trên bảng: Tổng số SV, Đúng hạn, Trễ hạn, Đã chấm, Chưa nộp.
   - Tỷ lệ nộp (%) và Progress Bar màu sắc (Xanh/Vàng/Tím).

### 13.3 Logic Nộp lại & Fix Bug Ý thức

- **Fix Ý thức cho SV tự do**: Trước v1.4, SV chưa vào nhóm bị skip tính điểm auto. Hiện tại hệ thống tự động fallback về lộ trình mặc định của Bộ môn giúp tính điểm ngay lập tức cho SV tự do.
- **Resubmission (Nộp lại)**: 
  - Khi SV nộp bản mới (v2, v3...), hệ thống ghi nhận `ngayNop` mới nhất.
  - Điểm Ý thức sẽ được tính lại theo thời điểm nộp mới (có thể bị downgrade nếu nộp lại sau deadline).
  - **Exception**: Nếu GV đã sửa điểm ý thức thủ công (`yThucTuDong = false`), hệ thống sẽ KHÔNG nộp lại tự động để bảo toàn quyết định của GV.

---

*Tài liệu này là phần mở rộng của PRD-02. Xem thêm `docs/PRD-02_Module_KLTN.md` cho context tổng thể.*
*File lib tính điểm: `lib/kltn/diem-y-thuc.ts`*
*Components: `components/kltn/ChamDiemTuanGv.tsx`, `ChamDiemTuanGv.cells.tsx`, `ChamDiemTuanGv.actions.ts`, `ChamDiemTuanGv.types.ts`*
*APIs mới: `app/api/kltn/linh-vuc/route.ts`, `app/api/kltn/de-tai/linh-vuc/route.ts`*

---

## 15. Fix Logic Auto-select Tuần & Nút Refresh (v1.8 — 06/04/2026)

### 15.1 Bug: Auto-select tuần chỉ so sánh thứ, bỏ qua giờ

**Root cause:** Logic cũ so sánh `vnDay < deadlineDay` (số thứ VN) — khi cùng ngày Thứ 5 nhưng trước/sau 08:00 cho kết quả giống nhau, gây UX không nhất quán.

**Ví dụ lỗi:** Thứ 5 07:59 (chưa đến mốc 08:00) → hiển thị tuần hiện tại ❌ (đúng phải là tuần trước).

### 15.2 Business Rule

> **Sau khi timestamp deadline của tuần N đã qua** → trang mặc định hiển thị **Tuần N**.
> **Trước timestamp deadline** → hiển thị **Tuần N-1**.

Deadline được tính đầy đủ: `(ngày trong tuần) + (giờ:phút)` từ `mocBaoCaoTuanThu`, `mocBaoCaoTuanGio`, `mocBaoCaoTuanPhut`.

### 15.3 Thuật toán mới (đã triển khai)

```typescript
// Tính ngày bắt đầu tuần đang diễn ra
const tuanStart = new Date(gd3Start)
tuanStart.setDate(tuanStart.getDate() + (tuanDangDienRa - 1) * 7)

// Mốc deadline: thứ (VN 2–8) + giờ + phút, mặc định Thứ 5 08:00
const mocThu  = dotInfo?.mocBaoCaoTuanThu  ?? 5
const mocGio  = dotInfo?.mocBaoCaoTuanGio  ?? 8
const mocPhut = dotInfo?.mocBaoCaoTuanPhut ?? 0

// Chuyển mocThu VN (2–8) → JS getDay() (0–6)
const mocThuJs = mocThu === 8 ? 0 : mocThu - 1

let delta = mocThuJs - tuanStart.getDay()
if (delta < 0) delta += 7
const deadline = new Date(tuanStart)
deadline.setDate(deadline.getDate() + delta)
deadline.setHours(mocGio, mocPhut, 0, 0)

// So sánh timestamp đầy đủ
const tuanCanXem = now >= deadline ? tuanDangDienRa : tuanDangDienRa - 1
```

### 15.4 Fix: Nút "Làm mới" header kết nối component

**Root cause:** `page.tsx` dispatch `CustomEvent('kltn-refresh')` nhưng `ChamDiemTuanGv` không có listener.

**Giải pháp:** Thêm `useEffect` subscribe event:

```typescript
useEffect(() => {
    const handler = () => fetchData(true)  // silent=true: không reset tuần đang xem
    window.addEventListener('kltn-refresh', handler)
    return () => window.removeEventListener('kltn-refresh', handler)
}, [fetchData])
```

### 15.5 Acceptance Criteria

| ID | Tình huống | Kết quả |
|----|-----------|--------|
| AC-13 | Thứ 5 trước 08:00, deadline Thứ 5 08:00 | Auto-select **Tuần N-1** ✅ |
| AC-14 | Thứ 5 từ 08:00 trở đi | Auto-select **Tuần N** ✅ |
| AC-15 | Bấm "Làm mới" ngoài header | Dữ liệu reload, tuần giữ nguyên ✅ |

---

## 14. Nâng cấp Bảng Chấm điểm GV (v1.6 — 30/03/2026)

### 14.1 Tách file `ChamDiemTuanGv`

Component monolith được refactor thành 4 file độc lập:

| File | Nội dung |
|------|----------|
| `ChamDiemTuanGv.tsx` | Component chính — state, fetch, render bảng |
| `ChamDiemTuanGv.cells.tsx` | `CommentEditCell` — ô nhận xét inline có history dropdowns |
| `ChamDiemTuanGv.actions.ts` | Server actions / API calls (quickSave, bulkSet...) |
| `ChamDiemTuanGv.types.ts` | Types: `SvRow`, `DiemTuan`, `BaoCaoRow`, `ThangDiemItem` |

### 14.2 History Dropdowns NX T-1/T-2

Giảng viên có thể xem nhận xét của 2 tuần trước ngay trong ô nhập liệu:

- **Hiển thị**: `NX T{w} ▼` — 2 dòng xếp theo thứ tự thời gian (tuần nhỏ trên)
- **Toggle**: Nhấp để xổ/thu gọn từng dòng riêng biệt (state `Set<number>`)
- **Inline**: Panel hiển thị trong luồng nội dung (không popup nổi)
- **Trạng thái cố định**: Giữ trạng thái mở khi click sang ô khác
- **Chèn nhanh**: Nút 📋 copy nội dung tuần cũ vào ô soạn thảo hiện tại
- **Màu sắc**: Mã cyan `#22d3ee` cho nhãn NX Tx, phân biệt với nội dung hiện tại

### 14.3 Cột LV — Editable Field của Lĩnh vực

Giảng viên có thể cập nhật lĩnh vực đề tài trực tiếp từ bảng chấm điểm:

**UI (giống cột YT):**
- Hiển thị mã lĩnh vực (VD: `A.1`) với màu tím + badge
- Hover: icon ✏️ mờ hiện rõ dần
- Click: mở popover `absolute` tại cell
- Popover: danh sách lĩnh vực nhóm theo **maCap1** → **maCap2** (nhận từ `linhVucDeTai` table)
- Highlight mục đang chọn (ring tím)
- Nút "— Bỏ chọn" để xoá lĩnh vực
- Lưu ngay khi chọn, đóng popover, cập nhật local state không reload

**API flow:**
```
GET /api/kltn/linh-vuc?dotKltnId=...
  → Lấy chuyenNganhId từ dotKltn
  → Nếu null: fallback toBộ môn → lấy tất cả chuyenNganh thuộc boMonId
  → Query linhVucDeTai WHERE chuyenNganhId IN [...] AND isActive = true
  → Trả về list sorted by thuTu, maCap2

PATCH /api/kltn/de-tai/linh-vuc
  Body: { svNhomKltnId, linhVucKltn }
  → Lookup deTai by svNhomKltnId
  → Update linhVucKltn
  → Insert deTaiLog (loai: SUA_NOI_DUNG, nguon: GV_CHAM_DIEM_BANG)
```

**Dữ liệu lĩnh vực**: Lấy từ bảng `linh_vuc_de_tai` — cùng nguồn với trang cấu hình Bộ môn (`/bo-mon/[id]/cau-hinh`). Chỉ hiển thị lĩnh vực có `isActive = true`.

- Thêm `'linhVuc'` vào `SortKey` union type
- Click header cột LV cycle: `asc` → `desc` → reset
- Sort theo `row.linhVucKltn` (string, null-last)

### 14.5 Fix Bug Overwrite Null Nhận xét (v1.7)

Sửa lỗi nghiêm trọng khi GV chấm điểm Chuyên môn (CM) làm mất nội dung nhận xét ở 2 cột "Cấu trúc" và "Nội dung":

- **Nguyên nhân**: API patch CM trước đây overwrite `chuyênMonCauTruc` và `chuyenMonNoiDung` thành `null` nếu body request không gửi kèm.
- **Giải pháp**: Sử dụng cơ chế conditional spread trong Drizzle `set()` object. Chỉ cập nhật 2 trường nhận xét nếu chúng được định nghĩa tường minh trong body (`!== undefined`).
- **Phạm vi fix**: API route `/api/kltn/diem-qt/[id]/chuyen-mon/route.ts`.
