# ===== 1단계: 빌드용 =====
FROM node:20-slim AS builder

WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
RUN npx vite build

# ===== 2단계: 실행용 (경량 - Tesseract OCR) =====
FROM python:3.11-slim

# Node.js + Tesseract OCR 설치 (PyTorch 불필요!)
RUN apt-get update && apt-get install -y \
    curl \
    tesseract-ocr \
    tesseract-ocr-kor \
    tesseract-ocr-eng \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python 의존성 (경량 - PyTorch/EasyOCR 제거!)
RUN pip install --no-cache-dir \
    opencv-python-headless \
    numpy \
    Pillow \
    openpyxl \
    pytesseract

# Node.js 의존성 (프로덕션만)
COPY package*.json ./
RUN npm install --omit=dev

# 빌드된 프론트엔드 복사
COPY --from=builder /app/dist ./dist

# 소스 코드 복사
COPY server.js ./
COPY ocr_engine_v3.py ./
COPY excel_converter.py ./
COPY basic_excel_generator.py ./
COPY json_to_floor_unit.py ./
COPY .env* ./

RUN mkdir -p uploads

EXPOSE ${PORT:-3500}
CMD ["node", "server.js"]
