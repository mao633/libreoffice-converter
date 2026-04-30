FROM ubuntu:22.04
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-impress \
    libreoffice-writer \
    python3 python3-pip \
    fonts-dejavu fonts-liberation fonts-noto \
    curl ca-certificates \
    && rm -rf /var/lib/apt/lists/*
RUN pip3 install --no-cache-dir flask gunicorn pillow
WORKDIR /app
COPY server.py /app/server.py
ENV PORT=10000
ENV PYTHONUNBUFFERED=1
EXPOSE 10000
# 1 worker, 1 thread, max-requests 1 -> worker ricreato dopo ogni request
# (zero memory leak accumulation, importante con LibreOffice + Pillow su 512MB).
CMD ["gunicorn", "-w", "1", "--threads", "1", "--max-requests", "1", \
     "-b", "0.0.0.0:10000", "--timeout", "240", \
     "--access-logfile", "-", "--error-logfile", "-", \
     "server:app"]
