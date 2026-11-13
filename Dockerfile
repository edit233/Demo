# PDF → 可编辑 PPTX 转换服务
#
# 构建：docker build -t pdf2pptx-service:1.0.0 .
# 运行：docker run -d -p 8000:8000 -v pdf2pptx-data:/data pdf2pptx-service:1.0.0

FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PDF2PPTX_UPLOAD_DIR=/data/uploads \
    PDF2PPTX_OUTPUT_DIR=/data/outputs \
    PDF2PPTX_ENABLE_OCR=1

WORKDIR /app

# opencv（rapidocr 依赖）需要 libGL / libglib
RUN apt-get update \
    && apt-get install -y --no-install-recommends libgl1 libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY scripts ./scripts
COPY main.py ./

RUN mkdir -p /data/uploads /data/outputs

EXPOSE 8000

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8000/api/health',timeout=4)"

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]
