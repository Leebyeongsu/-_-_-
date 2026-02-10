# Node.js + Python 멀티 런타임 이미지
FROM node:20-slim

# Python 설치
RUN apt-get update && apt-get install -y \
    python3 \
    python3-pip \
    python3-venv \
    libgl1-mesa-glx \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

# python3를 python으로 심볼릭 링크
RUN ln -sf /usr/bin/python3 /usr/bin/python

WORKDIR /app

# Node.js 의존성 설치
COPY package*.json ./
RUN npm install

# Python 의존성 설치
COPY requirements.txt ./
RUN pip3 install --no-cache-dir --break-system-packages -r requirements.txt

# 소스 코드 복사
COPY . .

# Vite 빌드 (프론트엔드 정적 파일 생성)
RUN npm run build

# uploads 폴더 생성
RUN mkdir -p uploads

# 포트 노출
EXPOSE ${PORT:-3500}

# 서버 실행
CMD ["node", "server.js"]
