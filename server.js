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

// JSON 데이터 → 기본 엑셀 생성 API
app.post('/api/download-basic-excel', async (req, res) => {
    const startTime = Date.now();
    console.log(`[${new Date().toLocaleTimeString()}] 📥 기본 엑셀 생성 시작...`);

    const jsonData = req.body;
    if (!jsonData || !jsonData.data) {
        return res.status(400).json({ error: '데이터가 없습니다.' });
    }

    const outputPath = path.join('uploads', `basic_excel_${Date.now()}.xlsx`);

    try {
        // Python 스크립트 실행
        const pythonScript = path.join(__dirname, 'basic_excel_generator.py');
        const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

        await new Promise((resolve, reject) => {
            const pythonProcess = spawn(pythonCmd, [
                pythonScript,
                JSON.stringify(jsonData),
                outputPath
            ], {
                encoding: 'utf-8',
                env: {
                    ...process.env,
                    PYTHONIOENCODING: 'utf-8'
                }
            });

            let stderr = '';

            pythonProcess.stderr.on('data', (data) => {
                stderr += data.toString('utf-8');
                console.log('[Basic Excel Generator]', data.toString('utf-8'));
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    resolve();
                } else {
                    reject(new Error(`생성 실패 (code ${code}): ${stderr}`));
                }
            });

            pythonProcess.on('error', (err) => {
                reject(new Error(`Python 실행 실패: ${err.message}`));
            });
        });

        const elapsed = Date.now() - startTime;
        console.log(`[${elapsed}ms] ✅ 기본 엑셀 생성 완료`);

        // 생성된 파일 전송
        res.download(outputPath, `현황표_${Date.now()}.xlsx`, (err) => {
            // 전송 완료 후 임시 파일 정리
            if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

            if (err) {
                console.error('파일 전송 오류:', err);
            }
        });

    } catch (error) {
        console.error('❌ 생성 오류:', error.message);
        res.status(500).json({ error: error.message });

        // 오류 발생 시 임시 파일 정리
        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
    }
});

// 엑셀 변환 API (개선된 버전 - 파일명 인코딩 + 에러 로깅 강화)
app.post('/api/convert-excel', upload.single('excel'), async (req, res) => {
    if (!req.file) return res.status(400).json({ error: '엑셀 파일이 없습니다.' });

    const startTime = Date.now();
    console.log(`\n${'='.repeat(60)}`);
    console.log(`[${new Date().toLocaleTimeString()}] 📊 엑셀 변환 시작...`);
    console.log(`📄 원본 파일명: ${req.file.originalname}`);
    console.log(`📦 MIME 타입: ${req.file.mimetype}`);
    console.log(`📏 파일 크기: ${(req.file.size / 1024).toFixed(2)} KB`);
    console.log(`💾 저장 경로: ${req.file.path}`);

    // multer는 확장자 없이 저장하므로, .xlsx 확장자를 붙여서 rename
    const inputPathRaw = req.file.path;
    const inputPath = inputPathRaw + '.xlsx';
    fs.renameSync(inputPathRaw, inputPath);
    console.log(`📁 확장자 추가: ${inputPathRaw} → ${inputPath}`);

    // ASCII 안전한 출력 파일명 생성 (한글 제거)
    const safeOutputName = `converted_${Date.now()}.xlsx`;
    const outputPath = path.join('uploads', safeOutputName);
    console.log(`📤 출력 경로: ${outputPath}`);

    try {
        // 파일 검증: ZIP 헤더 확인 (.xlsx는 ZIP 형식)
        const fileBuffer = fs.readFileSync(inputPath);
        const zipHeader = fileBuffer.slice(0, 4).toString('hex');
        const fileSize = fileBuffer.length;
        console.log(`🔍 파일 헤더: ${zipHeader} (올바른 ZIP: 504b0304)`);
        console.log(`🔍 실제 파일 크기: ${fileSize} bytes`);

        if (zipHeader !== '504b0304') {
            console.error('❌ 유효하지 않은 엑셀 파일 (ZIP 헤더 없음)');
            if (fs.existsSync(inputPath)) fs.unlinkSync(inputPath);
            return res.status(400).json({
                error: '유효하지 않은 엑셀 파일입니다. ZIP 헤더가 없습니다.'
            });
        }

        // Python 엑셀 변환 스크립트 실행
        const pythonScript = path.join(__dirname, 'excel_converter.py');
        const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

        console.log(`🐍 Python 실행: ${pythonCmd} ${pythonScript}`);
        console.log(`📥 입력: ${inputPath}`);
        console.log(`📤 출력: ${outputPath}`);

        await new Promise((resolve, reject) => {
            const pythonProcess = spawn(pythonCmd, [pythonScript, inputPath, outputPath], {
                encoding: 'utf-8',
                env: {
                    ...process.env,
                    PYTHONIOENCODING: 'utf-8',
                    PYTHONUNBUFFERED: '1'
                }
            });

            let stdout = '';
            let stderr = '';

            pythonProcess.stdout.on('data', (data) => {
                const output = data.toString('utf-8');
                stdout += output;
                console.log('[Python STDOUT]', output.trim());
            });

            pythonProcess.stderr.on('data', (data) => {
                const output = data.toString('utf-8');
                stderr += output;
                console.log('[Python STDERR]', output.trim());
            });

            pythonProcess.on('close', (code) => {
                console.log(`🔚 Python 프로세스 종료 (code: ${code})`);
                if (code === 0) {
                    // 출력 파일 확인
                    if (fs.existsSync(outputPath)) {
                        const outputSize = fs.statSync(outputPath).size;
                        console.log(`✅ 출력 파일 생성됨: ${outputSize} bytes`);
                        resolve();
                    } else {
                        reject(new Error('출력 파일이 생성되지 않았습니다.'));
                    }
                } else {
                    reject(new Error(`변환 실패 (exit code ${code})\n${stderr}`));
                }
            });

            pythonProcess.on('error', (err) => {
                console.error('❌ Python 프로세스 오류:', err);
                reject(new Error(`Python 실행 실패: ${err.message}`));
            });
        });

        const elapsed = Date.now() - startTime;
        console.log(`[${elapsed}ms] ✅ 엑셀 변환 완료`);

        // 원본 파일명 기반으로 다운로드 파일명 생성
        const originalNameWithoutExt = req.file.originalname.replace(/\.(xlsx|xls)$/i, '');
        const downloadName = `${originalNameWithoutExt}_변환.xlsx`;
        console.log(`📥 다운로드 파일명: ${downloadName}`);

        // 변환된 파일 전송
        res.download(outputPath, downloadName, (err) => {
            console.log(`🧹 임시 파일 정리 중...`);
            // 전송 완료 후 임시 파일 정리
            if (fs.existsSync(inputPath)) {
                fs.unlinkSync(inputPath);
                console.log(`  ✓ 입력 파일 삭제: ${inputPath}`);
            }
            if (fs.existsSync(outputPath)) {
                fs.unlinkSync(outputPath);
                console.log(`  ✓ 출력 파일 삭제: ${outputPath}`);
            }

            if (err) {
                console.error('❌ 파일 전송 오류:', err);
            } else {
                console.log(`✅ 파일 전송 완료`);
            }
            console.log(`${'='.repeat(60)}\n`);
        });

    } catch (error) {
        console.error('❌ 변환 오류:', error.message);
        console.error('상세 에러:', error.stack);
        res.status(500).json({
            error: error.message,
            details: '서버 로그를 확인하세요.'
        });

        // 오류 발생 시 임시 파일 정리
        console.log(`🧹 오류 발생 - 임시 파일 정리 중...`);
        if (fs.existsSync(inputPath)) {
            fs.unlinkSync(inputPath);
            console.log(`  ✓ 입력 파일 삭제: ${inputPath}`);
        }
        if (fs.existsSync(outputPath)) {
            fs.unlinkSync(outputPath);
            console.log(`  ✓ 출력 파일 삭제: ${outputPath}`);
        }
        console.log(`${'='.repeat(60)}\n`);
    }
});

// JSON 데이터 → 층호수 형태 엑셀 변환 API
app.post('/api/convert-data-to-floor-unit', async (req, res) => {
    const startTime = Date.now();
    console.log(`[${new Date().toLocaleTimeString()}] 📊 층호수 형태 변환 시작...`);

    const jsonData = req.body;
    if (!jsonData || !jsonData.data) {
        return res.status(400).json({ error: '데이터가 없습니다.' });
    }

    const outputPath = path.join('uploads', `floor_unit_${Date.now()}.xlsx`);

    try {
        // Python 스크립트 실행
        const pythonScript = path.join(__dirname, 'json_to_floor_unit.py');
        const pythonCmd = process.platform === 'win32' ? 'python' : 'python3';

        await new Promise((resolve, reject) => {
            const pythonProcess = spawn(pythonCmd, [
                pythonScript,
                JSON.stringify(jsonData),
                outputPath
            ], {
                encoding: 'utf-8',
                env: {
                    ...process.env,
                    PYTHONIOENCODING: 'utf-8'
                }
            });

            let stderr = '';

            pythonProcess.stderr.on('data', (data) => {
                stderr += data.toString('utf-8');
                console.log('[Floor Unit Converter]', data.toString('utf-8'));
            });

            pythonProcess.on('close', (code) => {
                if (code === 0) {
                    resolve();
                } else {
                    reject(new Error(`변환 실패 (code ${code}): ${stderr}`));
                }
            });

            pythonProcess.on('error', (err) => {
                reject(new Error(`Python 실행 실패: ${err.message}`));
            });
        });

        const elapsed = Date.now() - startTime;
        console.log(`[${elapsed}ms] ✅ 층호수 형태 변환 완료`);

        // 변환된 파일 전송
        res.download(outputPath, `층호수형태_${Date.now()}.xlsx`, (err) => {
            // 전송 완료 후 임시 파일 정리
            if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);

            if (err) {
                console.error('파일 전송 오류:', err);
            }
        });

    } catch (error) {
        console.error('❌ 변환 오류:', error.message);
        res.status(500).json({ error: error.message });

        // 오류 발생 시 임시 파일 정리
        if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
    }
});

app.listen(port, () => {
    console.log(`🚀 Server at http://localhost:${port}`);
    console.log(`📦 OCR Engine: v3 (자동 그리드 + EasyOCR)`);
    console.log(`📦 Excel Converter: openpyxl 기반 변환`);
    console.log(`📦 Floor-Unit Converter: JSON → 층호수 형태 엑셀`);
    console.log(`📊 자동 크기 감지, 색상 + 텍스트 인식`);
});
