# CheckForm KLTN — Kiểm tra Định dạng Khóa Luận Tốt Nghiệp

> Công cụ kiểm tra tự động định dạng file `.docx` theo tiêu chuẩn trình bày KLTN/ĐATN của Trường Đại học Thủy Lợi — Khoa KT & QTKD.

[![Build & Push Docker image to GHCR](https://github.com/saodonga/formkltn/actions/workflows/docker-publish.yml/badge.svg)](https://github.com/saodonga/formkltn/actions/workflows/docker-publish.yml)

---

## 🐳 Cài đặt bằng Docker (Khuyến nghị)

### Yêu cầu

- Ubuntu 20.04+ / Debian 11+ (hoặc bất kỳ Linux nào có Docker)
- Docker Engine ≥ 24.x
- Docker Compose ≥ 2.x

---

### Bước 1 — Cài Docker trên Ubuntu (bỏ qua nếu đã có)

```bash
# Cập nhật package list
sudo apt update && sudo apt upgrade -y

# Cài các package cần thiết
sudo apt install -y ca-certificates curl gnupg lsb-release

# Thêm GPG key của Docker
sudo install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg \
  | sudo gpg --dearmor -o /etc/apt/keyrings/docker.gpg
sudo chmod a+r /etc/apt/keyrings/docker.gpg

# Thêm repo Docker
echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.gpg] \
  https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" \
  | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null

# Cài Docker
sudo apt update
sudo apt install -y docker-ce docker-ce-cli containerd.io docker-buildx-plugin docker-compose-plugin

# Cho phép chạy Docker không cần sudo
sudo usermod -aG docker $USER
newgrp docker

# Kiểm tra
docker --version
docker compose version
```

---

### Bước 2 — Tạo thư mục và lấy file cấu hình

```bash
mkdir -p ~/checkform-kltn && cd ~/checkform-kltn
```

Tạo file `docker-compose.yml`:

```bash
cat > docker-compose.yml << 'EOF'
services:
  checkform-kltn:
    image: ghcr.io/saodonga/formkltn:latest
    container_name: checkform-kltn
    restart: unless-stopped
    ports:
      - "8080:8080"
    volumes:
      - ./config_kltn.json:/app/config_kltn.json
    environment:
      - FLASK_ENV=production
      - PORT=8080
    healthcheck:
      test: ["CMD", "python3", "-c", "import urllib.request; urllib.request.urlopen('http://localhost:8080/')"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 10s
EOF
```

Tạo file `config_kltn.json` ban đầu (danh sách GVHD có thể sửa trên web):

```bash
cat > config_kltn.json << 'EOF'
{
  "advisors": [],
  "_title_min_length": 50
}
EOF
```

---

### Bước 3 — Pull image và chạy

```bash
# Pull image mới nhất từ GHCR
docker compose pull

# Chạy nền (detached)
docker compose up -d

# Kiểm tra trạng thái
docker compose ps

# Xem log
docker compose logs -f
```

✅ Truy cập tại: **`http://localhost:8080`** (hoặc `http://IP_SERVER:8080`)

---

### Bước 4 — Cập nhật khi có phiên bản mới

```bash
cd ~/checkform-kltn

# Pull image mới nhất và restart
docker compose pull && docker compose up -d
```

---

## ⚙️ Cấu hình

### Mở port trên firewall (nếu dùng VPS)

```bash
# Ubuntu UFW
sudo ufw allow 8080/tcp

# Hoặc dùng port 80 (HTTP mặc định): đổi "8080:8080" → "80:8080" trong docker-compose.yml
sudo ufw allow 80/tcp
```

### Đổi port

Sửa dòng `ports` trong `docker-compose.yml`:

```yaml
ports:
  - "80:8080"   # Truy cập qua port 80 (không cần gõ :8080)
```

Sau đó restart:

```bash
docker compose up -d
```

---

## 🖥️ Chạy trực tiếp bằng Python (không cần Docker)

### Yêu cầu

- Python 3.10+
- pip

### Cài đặt

```bash
git clone https://github.com/saodonga/formkltn.git
cd formkltn
pip install -r requirements.txt
```

### Chạy

```bash
python3 web_app.py
# Mở http://localhost:5000
```

---

## 📋 Các lệnh Docker hữu ích

```bash
# Xem log real-time
docker compose logs -f

# Dừng app
docker compose stop

# Khởi động lại
docker compose restart

# Xem tài nguyên đang dùng (CPU/RAM)
docker stats checkform-kltn

# Vào shell bên trong container (debug)
docker exec -it checkform-kltn bash

# Xóa và chạy lại sạch
docker compose down && docker compose up -d
```

---

## 📌 Tiêu chuẩn kiểm tra

| Hạng mục | Chuẩn yêu cầu |
|---|---|
| Khổ giấy | A4 (21 × 29.7 cm) |
| Lề trái | 3.0 cm |
| Lề phải | 2.0 cm |
| Lề trên / Dưới | 2.5 cm |
| Font chữ | Times New Roman |
| Cỡ chữ Heading 1 | 14pt, **đậm**, IN HOA, căn giữa |
| Cỡ chữ Heading 2 | 13pt, **đậm**, canh trái |
| Cỡ chữ Heading 3 | 13pt, **đậm + nghiêng**, canh trái |
| Cỡ chữ Heading 4 | 13pt, *nghiêng*, canh trái |
| Cỡ chữ nội dung | 13pt |
| Giãn dòng nội dung | 1.5 lines |
| Canh lề nội dung | Justify (đều hai bên) |
| Cấu trúc bắt buộc | Lời cam đoan, Mục lục, Danh mục, TLTK |

---

## 📂 Cấu trúc thư mục

```
formkltn/
├── web_app.py              ← Flask server (entry point)
├── check_format_kltn.py    ← Engine kiểm tra định dạng
├── config_kltn.json        ← Danh sách GVHD hướng dẫn
├── web_static/
│   ├── index.html          ← Giao diện web
│   ├── style.css
│   └── app.js
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```
