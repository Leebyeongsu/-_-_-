const fs = require('node:fs');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { spawn } = require('node:child_process');

class PublicError extends Error {
    constructor(status, message) {
        super(message);
        this.status = status;
    }
}

function assertDeploymentIsolation(env) {
    if (env.NODE_ENV !== 'production') return;
    // This service must never inherit an ERP database connection.
    for (const key of ['DATABASE_URL', 'DATABASE_AUTH_TOKEN', 'DATABASE_READ_AUTH_TOKEN', 'DB_PATH']) {
        if (env[key]) throw new Error(`Image service must not receive ${key}.`);
    }
}

function isImageHost(hostname, env) {
    const host = hostname.toLowerCase();
    if (['dsone.shop', 'www.dsone.shop', 'image-to-excel-dsone.vercel.app'].includes(host)) return true;
    if (/^image-to-excel-dsone-[a-z0-9-]+\.vercel\.app$/.test(host)) return true;
    if (host === 'trustworthy-enchantment-production-40ea.up.railway.app') return true;
    return env.NODE_ENV !== 'production' && ['localhost', '127.0.0.1', '[::1]'].includes(host);
}

function securityHeaders(env) {
    return (req, res, next) => {
        res.set({
            'X-Content-Type-Options': 'nosniff',
            'X-Frame-Options': 'DENY',
            'Referrer-Policy': 'no-referrer',
            'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
        });
        if (env.NODE_ENV === 'production') {
            res.set('Content-Security-Policy', [
                "default-src 'self'",
                "script-src 'self' https://cdn.jsdelivr.net/npm/marked/marked.min.js https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js",
                "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
                "font-src 'self' https://fonts.gstatic.com data:",
                "img-src 'self' data: blob:",
                "connect-src 'self'",
                "frame-src 'self' blob:",
                "frame-ancestors 'none'",
                "base-uri 'none'",
                "form-action 'self'",
            ].join('; '));
        }
        let requestUrl;
        try { requestUrl = new URL(`http://${req.get('host')}`); } catch { /* rejected below */ }
        if (!requestUrl || requestUrl.username || requestUrl.password || !isImageHost(requestUrl.hostname, env)) {
            return res.status(421).json({ error: '이미지 변환 서비스 주소로 접속하세요.' });
        }
        if (req.path.startsWith('/api/')) {
            res.set('Cache-Control', 'private, no-store, max-age=0');
            res.set('CDN-Cache-Control', 'no-store');
            res.set('Vercel-CDN-Cache-Control', 'no-store');
            if (!['GET', 'HEAD'].includes(req.method)) {
                const origin = req.get('origin');
                let originUrl;
                try { if (origin) originUrl = new URL(origin); } catch { /* rejected below */ }
                const localDev = env.NODE_ENV !== 'production' && originUrl
                    && ['localhost', '127.0.0.1', '[::1]'].includes(originUrl.hostname)
                    && ['localhost', '127.0.0.1', '[::1]'].includes(requestUrl.hostname);
                if ((origin && (!originUrl || !['http:', 'https:'].includes(originUrl.protocol)
                    || originUrl.username || originUrl.password || (!localDev && originUrl.host !== requestUrl.host)))
                    || req.get('sec-fetch-site') === 'cross-site') {
                    return res.status(403).json({ error: '다른 사이트에서 보낸 변환 요청은 허용하지 않습니다.' });
                }
            }
        }
        next();
    };
}

function createWorkspace(root) {
    const base = path.resolve(root);
    fs.mkdirSync(base, { recursive: true, mode: 0o700 });
    const directory = fs.mkdtempSync(path.join(base, 'job-'));
    fs.chmodSync(directory, 0o700);
    return {
        id: randomUUID(),
        directory,
        controller: new AbortController(),
        async cleanup() {
            // Only remove the generated direct child of this service's private root.
            if (path.dirname(directory) !== base || !path.basename(directory).startsWith('job-')) {
                throw new Error('Invalid workspace cleanup target.');
            }
            await fs.promises.rm(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
        },
    };
}

function pythonEnvironment(workspace, env = process.env) {
    const result = {};
    for (const key of ['PATH', 'Path', 'SystemRoot', 'SYSTEMROOT', 'WINDIR', 'LANG', 'LC_ALL', 'TESSERACT_CMD', 'TESSDATA_PREFIX']) {
        if (env[key]) result[key] = env[key];
    }
    return {
        ...result,
        HOME: workspace.directory,
        TMPDIR: workspace.directory,
        TEMP: workspace.directory,
        TMP: workspace.directory,
        PYTHONIOENCODING: 'utf-8',
        PYTHONUNBUFFERED: '1',
        PYTHONDONTWRITEBYTECODE: '1',
    };
}

function runPython(script, args, workspace, { env = process.env, timeoutMs = 180000 } = {}) {
    if (workspace.controller.signal.aborted) return Promise.reject(new PublicError(499, '요청이 취소되었습니다.'));
    return new Promise((resolve, reject) => {
        const child = spawn(process.platform === 'win32' ? 'python' : 'python3', [script, ...args], {
            cwd: workspace.directory,
            env: pythonEnvironment(workspace, env),
            windowsHide: true,
        });
        const chunks = [];
        let length = 0;
        let failure;
        const stop = (error) => { failure ||= error; child.kill('SIGKILL'); };
        const aborted = () => stop(new PublicError(499, '요청이 취소되었습니다.'));
        const timer = setTimeout(() => stop(new PublicError(504, '처리 시간이 초과되었습니다. 파일 크기를 줄여 다시 시도하세요.')), timeoutMs);
        workspace.controller.signal.addEventListener('abort', aborted, { once: true });
        child.stdout.on('data', (chunk) => {
            length += chunk.length;
            if (length > 4 * 1024 * 1024) stop(new PublicError(413, '변환 결과가 너무 큽니다. 파일을 나누어 처리하세요.'));
            else chunks.push(chunk);
        });
        // OCR stderr can contain customer names. Drain it without logging or returning it.
        child.stderr.resume();
        child.on('error', () => { failure ||= new PublicError(500, '변환 프로그램을 실행할 수 없습니다.'); });
        child.on('close', (code) => {
            clearTimeout(timer);
            workspace.controller.signal.removeEventListener('abort', aborted);
            if (failure) reject(failure);
            else if (code !== 0) reject(new PublicError(422, '파일을 처리할 수 없습니다. 파일 형식과 내용을 확인하세요.'));
            else resolve(Buffer.concat(chunks).toString('utf8'));
        });
    });
}

module.exports = { PublicError, assertDeploymentIsolation, securityHeaders, createWorkspace, pythonEnvironment, runPython };
