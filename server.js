require('dotenv').config({ quiet: true });
const express = require('express');
const multer = require('multer');
const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { PublicError, assertDeploymentIsolation, securityHeaders, createWorkspace, runPython } = require('./request-isolation');

function createApp({ env = process.env, workspaceRoot = path.join(__dirname, 'uploads'), executePython = runPython, uploadLimit = 50 * 1024 * 1024 } = {}) {
    assertDeploymentIsolation(env);
    const app = express();
    app.disable('x-powered-by');
    app.disable('etag');
    app.use(securityHeaders(env));

    const storage = multer.diskStorage({
        destination: (req, file, cb) => cb(null, req.workspace.directory),
        filename: (req, file, cb) => cb(null, `${randomUUID()}${file.fieldname === 'image' ? '.image' : '.xlsx'}`),
    });
    const upload = multer({ storage, limits: { fileSize: uploadLimit, files: 1, fields: 0, parts: 2, fieldNameSize: 32, fieldArrayIndexLimit: 0, fieldNestingDepthLimit: 0 } });
    const jsonParser = express.json({ limit: '5mb' });

    function parse(parser, req, res) {
        return new Promise((resolve, reject) => {
            const signal = req.workspace.controller.signal;
            const aborted = () => reject(new PublicError(499, '요청이 취소되었습니다.'));
            if (signal.aborted) return aborted();
            signal.addEventListener('abort', aborted, { once: true });
            parser(req, res, (error) => {
                signal.removeEventListener('abort', aborted);
                if (error) {
                    reject(new PublicError(error.code === 'LIMIT_FILE_SIZE' || error.type === 'entity.too.large' ? 413 : 400,
                        error.code === 'LIMIT_FILE_SIZE' || error.type === 'entity.too.large' ? '파일 또는 요청 데이터가 너무 큽니다.' : '업로드 형식이 올바르지 않습니다.'));
                } else resolve();
            });
        });
    }

    function isolatedRoute(route, parser, handler) {
        app.post(route, async (req, res) => {
            const workspace = createWorkspace(workspaceRoot);
            req.workspace = workspace;
            const aborted = () => workspace.controller.abort();
            const closed = () => { if (!res.writableFinished) aborted(); };
            req.once('aborted', aborted);
            res.once('close', closed);
            try {
                await parse(parser, req, res);
                if (workspace.controller.signal.aborted) throw new PublicError(499, '요청이 취소되었습니다.');
                await handler(req, res, workspace);
            } catch (error) {
                // Keep filenames, paths, OCR text, and converter output out of shared logs/errors.
                if (!res.headersSent && !res.destroyed) {
                    res.status(error instanceof PublicError ? error.status : 500).json({
                        error: error instanceof PublicError ? error.message : '변환 중 오류가 발생했습니다. 파일을 확인해 다시 시도하세요.',
                    });
                }
            } finally {
                req.removeListener('aborted', aborted);
                res.removeListener('close', closed);
                try { await workspace.cleanup(); } catch { console.error('Temporary workspace cleanup failed:', workspace.id); }
            }
        });
    }

    async function execute(script, args, workspace) {
        return executePython(path.join(__dirname, script), args, workspace, { env });
    }

    async function download(res, workspace, filename) {
        const output = path.join(workspace.directory, 'result.xlsx');
        if (!fs.existsSync(output)) throw new PublicError(422, '변환 결과를 생성할 수 없습니다. 파일 내용을 확인하세요.');
        await new Promise((resolve, reject) => res.download(output, filename,
            { etag: false, lastModified: false, cacheControl: false },
            (error) => error ? reject(error) : resolve()));
    }

    function data(req) {
        if (!req.body || !Array.isArray(req.body.data) || req.body.data.length === 0) {
            throw new PublicError(400, '변환할 데이터가 없습니다.');
        }
        return JSON.stringify(req.body);
    }

    function excelInput(req) {
        if (!req.file) throw new PublicError(400, '엑셀 파일이 없습니다.');
        const handle = fs.openSync(req.file.path, 'r');
        const signature = Buffer.alloc(4);
        try { fs.readSync(handle, signature, 0, 4, 0); } finally { fs.closeSync(handle); }
        if (signature.toString('hex') !== '504b0304') throw new PublicError(400, '유효하지 않은 엑셀 파일입니다. .xlsx 파일을 선택하세요.');
        return req.file.path;
    }

    function originalDownloadName(req, suffix) {
        const base = path.basename(req.file.originalname.replace(/\\/g, '/')).replace(/\.(xlsx|xls)$/i, '')
            .replace(/[\x00-\x1f\x7f]/g, '').slice(0, 100) || '엑셀';
        return `${base}_${suffix}.xlsx`;
    }

    isolatedRoute('/api/ocr', upload.single('image'), async (req, res, workspace) => {
        if (!req.file) throw new PublicError(400, '이미지가 없습니다.');
        const result = JSON.parse(await execute('ocr_engine_v3.py', [req.file.path], workspace));
        if (result.error) throw new PublicError(422, '이미지에서 표를 인식할 수 없습니다. 선명한 표 이미지를 선택하세요.');
        res.json(result);
    });

    isolatedRoute('/api/download-basic-excel', jsonParser, async (req, res, workspace) => {
        await execute('basic_excel_generator.py', [data(req), path.join(workspace.directory, 'result.xlsx')], workspace);
        await download(res, workspace, '현황표.xlsx');
    });

    isolatedRoute('/api/convert-data-to-floor-unit', jsonParser, async (req, res, workspace) => {
        await execute('json_to_floor_unit.py', [data(req), path.join(workspace.directory, 'result.xlsx')], workspace);
        await download(res, workspace, '층호수형태.xlsx');
    });

    isolatedRoute('/api/convert-excel', upload.single('excel'), async (req, res, workspace) => {
        await execute('excel_converter.py', [excelInput(req), path.join(workspace.directory, 'result.xlsx')], workspace);
        await download(res, workspace, originalDownloadName(req, '변환'));
    });

    for (const [route, script] of [['/api/convert-map', 'map_converter.py'], ['/api/convert-map-seo', 'seo_map_converter.py']]) {
        isolatedRoute(route, upload.single('mapExcel'), async (req, res, workspace) => {
            const result = JSON.parse(await execute(script, [excelInput(req), path.join(workspace.directory, 'result.xlsx')], workspace));
            if (!result.success) throw new PublicError(422, '맵자료를 변환할 수 없습니다. 엑셀 내용을 확인하세요.');
            await download(res, workspace, originalDownloadName(req, '동별시트'));
        });
    }

    app.get('/api/guide', (req, res) => {
        res.type('text/plain; charset=utf-8').sendFile(path.join(__dirname, '사무실_사용_가이드.md'));
    });

    // Never expose workspaces, environment files, or arbitrary server paths through the SPA fallback.
    app.use((req, res, next) => {
        if (/^\/(?:uploads|\.env|\.git|request-isolation\.js|server\.js)(?:[./]|$)/i.test(req.path)) {
            return res.status(404).end();
        }
        next();
    });
    const dist = path.join(__dirname, 'dist');
    if (fs.existsSync(dist)) {
        app.use(express.static(dist, { dotfiles: 'deny' }));
        app.use((req, res, next) => {
            if (req.method === 'GET' && !req.path.startsWith('/api/')) res.sendFile(path.join(dist, 'index.html'));
            else next();
        });
    }
    app.use((error, req, res, next) => {
        if (res.headersSent) return next(error);
        res.status(500).json({ error: '요청을 처리할 수 없습니다.' });
    });
    return app;
}

if (require.main === module) {
    createApp().listen(process.env.PORT || 3500, '0.0.0.0', () => console.log('VisionFlow image converter started.'));
}
module.exports = { createApp };
