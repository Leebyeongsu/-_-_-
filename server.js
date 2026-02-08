require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const app = express();
const port = 3500;

app.use(cors());
app.use(express.json());

if (!fs.existsSync('uploads')) {
    fs.mkdirSync('uploads');
}
const upload = multer({ dest: 'uploads/' });

// Python 스크립트 실행 함수 (v3 - 자동 그리드 감지 + 텍스트)
function runPythonOCR(imagePath) {
    return new Promise((resolve, reject) => {
        const pythonScript = path.join(__dirname, 'ocr_engine_v3.py');

        // Python 실행 (Windows에서는 python, Unix에서는 python3)
        const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

        const pythonProcess = spawn(pythonCmd, [pythonScript, imagePath], {
            encoding: 'utf-8',
            env: {
                ...process.env,
                PYTHONIOENCODING: 'utf-8',
                GEMINI_API_KEY: process.env.GEMINI_API_KEY || '',
                ANTHROPIC_API_KEY: process.env.ANTHROPIC_API_KEY || '',
                AI_PROVIDER: process.env.AI_PROVIDER || ''
            }
        });

        let stdout = '';
        let stderr = '';

        pythonProcess.stdout.on('data', (data) => {
            stdout += data.toString('utf-8');
        });

        pythonProcess.stderr.on('data', (data) => {
            stderr += data.toString('utf-8');
            console.log('[Python]', data.toString('utf-8'));
        });

        pythonProcess.on('close', (code) => {
            if (code === 0) {
                try {
                    const result = JSON.parse(stdout);
                    resolve(result);
                } catch (e) {
                    reject(new Error(`JSON 파싱 오류: ${e.message}\n출력: ${stdout}`));
                }
            } else {
                reject(new Error(`Python 오류 (code ${code}): ${stderr}`));
            }
        });

        pythonProcess.on('error', (err) => {
            reject(new Error(`Python 실행 실패: ${err.message}`));
        });
    });
}

app.post('/api/ocr', upload.single('image'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: '이미지가 없습니다.' });

    const startTime = Date.now();
    console.log(`[${new Date().toLocaleTimeString()}] 🚀 OCR 분석 시작 (OpenCV + EasyOCR)...`);

    try {
        // Python OCR 엔진 실행
        const jsonData = await runPythonOCR(req.file.path);

        const elapsed = Date.now() - startTime;
        // 새 포맷: { header: {...}, data: [...] } 또는 기존 배열 호환
        const floorCount = jsonData.data ? jsonData.data.length : jsonData.length;
        console.log(`[${elapsed}ms] ✅ 분석 완료 (${floorCount}층)`);

        res.json(jsonData);

    } catch (error) {
        console.error('❌ 오류:', error.message);
        res.status(500).json({ error: error.message });
    } finally {
        // 임시 파일 정리
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
    }
});

app.listen(port, () => {
    console.log(`🚀 Server at http://localhost:${port}`);
    console.log(`📦 OCR Engine: v3 (자동 그리드 + EasyOCR)`);
    console.log(`📊 자동 크기 감지, 색상 + 텍스트 인식`);
});
