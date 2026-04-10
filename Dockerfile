# ── Stage: Build ────────────────────────────────────────────────
FROM python:3.12-slim

# Metadata
LABEL org.opencontainers.image.title="CheckForm KLTN"
LABEL org.opencontainers.image.description="Kiểm tra định dạng Khóa Luận Tốt Nghiệp — Trường ĐH Thủy Lợi"
LABEL org.opencontainers.image.source="https://github.com/saodonga/formkltn"

# Môi trường
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    FLASK_ENV=production \
    PORT=8080

WORKDIR /app

# Cài dependencies trước (layer cache hiệu quả)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ source
COPY check_format_kltn.py .
COPY web_app.py .
COPY config_kltn.json .
COPY web_static/ ./web_static/

# Tạo thư mục cần thiết
RUN mkdir -p /tmp/kltn_uploads

# Port expose
EXPOSE 8080

# Healthcheck
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/')" || exit 1

# Chạy bằng gunicorn (production WSGI server)
CMD ["gunicorn", "web_app:app", \
     "--workers", "2", \
     "--timeout", "120", \
     "--bind", "0.0.0.0:8080", \
     "--access-logfile", "-", \
     "--error-logfile", "-"]
