const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const http = require('node:http');
const { once } = require('node:events');
const { createApp } = require('../server');
const { pythonEnvironment, createWorkspace, runPython } = require('../request-isolation');

const production = { NODE_ENV: 'production' };
const root = path.resolve(__dirname, '../uploads/security-tests');

// Native HTTP preserves Host; Node fetch deliberately replaces it with the URL host.
async function request(url, options = {}) {
    const input = new Request(url, options);
    const body = input.body ? Buffer.from(await input.arrayBuffer()) : undefined;
    return new Promise((resolve, reject) => {
        const outgoing = http.request(url, { method: input.method, headers: Object.fromEntries(input.headers) }, incoming => {
            const chunks = [];
            incoming.on('data', chunk => chunks.push(chunk));
            incoming.on('error', reject);
            incoming.on('end', () => resolve(new Response(Buffer.concat(chunks), { status: incoming.statusCode, headers: incoming.headers })));
        });
        outgoing.on('error', reject);
        outgoing.end(body);
    });
}

async function harness(executePython, options = {}) {
    const directory = await fs.mkdtemp(path.join(await fs.mkdir(root, { recursive: true }).then(() => root), 'case-'));
    const server = createApp({ env: production, workspaceRoot: directory, executePython, ...options }).listen(0, '127.0.0.1');
    await once(server, 'listening');
    return {
        directory,
        url: `http://127.0.0.1:${server.address().port}`,
        async close() {
            server.closeAllConnections();
            await new Promise(resolve => server.close(resolve));
            await empty(directory);
            assert.equal(path.dirname(directory), root);
            await fs.rm(directory, { recursive: true, force: true });
        },
    };
}

async function empty(directory) {
    for (let i = 0; i < 100; i++) {
        if ((await fs.readdir(directory)).length === 0) return;
        await new Promise(resolve => setTimeout(resolve, 10));
    }
    assert.deepEqual(await fs.readdir(directory), [], 'request workspaces must be removed');
}

function post(h, route, marker, headers = {}) {
    return request(h.url + route, {
        method: 'POST',
        headers: { host: 'dsone.shop', origin: 'https://dsone.shop', 'content-type': 'application/json', ...headers },
        body: JSON.stringify({ data: [{ floor: marker, units: {} }] }),
    });
}

function upload(h, route, field, marker, headers = {}) {
    const form = new FormData();
    form.append(field, new Blob([field === 'image' ? marker : `PK\x03\x04${marker}`]), '../../customer.xlsx');
    return request(h.url + route, { method: 'POST', headers: { host: 'dsone.shop', origin: 'https://dsone.shop', ...headers }, body: form });
}

test('overlapping requests in the same millisecond only receive their own input, output and scratch files', async () => {
    const directories = new Set();
    let arrived = 0;
    let release;
    const barrier = new Promise(resolve => { release = resolve; });
    const routes = [
        ['/api/download-basic-excel', null], ['/api/convert-data-to-floor-unit', null],
        ['/api/convert-excel', 'excel'], ['/api/convert-map', 'mapExcel'],
        ['/api/convert-map-seo', 'mapExcel'], ['/api/ocr', 'image'],
    ];
    const h = await harness(async (script, args, workspace) => {
        assert(!directories.has(workspace.directory));
        directories.add(workspace.directory);
        const marker = path.basename(script).includes('generator') || path.basename(script).includes('json_to')
            ? JSON.parse(args[0]).data[0].floor
            : (await fs.readFile(args[0], 'utf8')).replace(/^PK\x03\x04/, '');
        await fs.writeFile(path.join(workspace.directory, 'scratch.txt'), marker);
        if (++arrived === routes.length) release();
        await barrier;
        assert.equal(await fs.readFile(path.join(workspace.directory, 'scratch.txt'), 'utf8'), marker);
        if (path.basename(script) === 'ocr_engine_v3.py') return JSON.stringify({ data: [{ floor: marker }] });
        assert.equal(path.dirname(args[1]), workspace.directory);
        await fs.writeFile(args[1], marker);
        return JSON.stringify({ success: true });
    });
    const originalNow = Date.now;
    try {
        Date.now = () => 1700000000000;
        const responses = await Promise.all(routes.map(async ([route, field], i) => {
            const marker = `AGENCY_${i}_ONLY`;
            const response = await (field ? upload(h, route, field, marker) : post(h, route, marker));
            assert.equal(response.status, 200);
            assert.match(response.headers.get('cache-control'), /no-store/);
            assert.equal(response.headers.get('access-control-allow-origin'), null);
            assert.equal(response.headers.get('etag'), null);
            assert.equal(response.headers.get('last-modified'), null);
            const result = route === '/api/ocr' ? (await response.json()).data[0].floor : await response.text();
            assert.equal(result, marker);
            return result;
        }));
        assert.equal(new Set(responses).size, routes.length);
        assert.equal(directories.size, routes.length);
    } finally { Date.now = originalNow; await h.close(); }
});

test('cross-site and ERP-host requests are rejected before any file or Python job is created', async () => {
    let executions = 0;
    const h = await harness(async () => { executions++; });
    try {
        for (const headers of [{ origin: 'https://megatel.store' }, { origin: 'null' }, { origin: 'https://evil.example' }, { 'sec-fetch-site': 'cross-site' }]) {
            assert.equal((await post(h, '/api/download-basic-excel', 'PRIVATE', headers)).status, 403);
        }
        assert.equal((await post(h, '/api/download-basic-excel', 'PRIVATE', { host: 'megatel.store' })).status, 421);
        assert.equal((await post(h, '/api/download-basic-excel', 'PRIVATE', { host: 'dsone-erp.vercel.app' })).status, 421);
        assert.equal(executions, 0);
        assert.deepEqual(await fs.readdir(h.directory), []);
    } finally { await h.close(); }
});

test('public access, same-origin requests and Vercel project aliases remain available', async () => {
    const h = await harness(async () => { throw new Error('UNUSED'); });
    try {
        for (const host of ['image.savetelecom.shop', 'dsone.shop', 'image-to-excel-dsone.vercel.app', 'image-to-excel-dsone-build-byeongsulee1-2203s-projects.vercel.app']) {
            const response = await request(h.url + '/api/guide', { headers: { host } });
            assert.equal(response.status, 200);
            assert.match(response.headers.get('content-security-policy'), /frame-ancestors 'none'/);
            assert.equal(response.headers.get('set-cookie'), null);
        }
        assert.equal((await post(h, '/api/download-basic-excel', 'PUBLIC')).status, 500);
    } finally { await h.close(); }
});

test('invalid, oversized and failed requests leave no files and never return internal details', async () => {
    const h = await harness(async () => { throw new Error('SECRET_CUSTOMER /app/uploads/other-job DATABASE_AUTH_TOKEN'); }, { uploadLimit: 64 });
    try {
        const failed = await post(h, '/api/download-basic-excel', 'X');
        assert.equal(failed.status, 500);
        assert.doesNotMatch(await failed.text(), /SECRET|other-job|DATABASE_AUTH_TOKEN/);
        assert.equal((await upload(h, '/api/ocr', 'image', 'X'.repeat(100))).status, 413);
        assert.equal((await upload(h, '/api/ocr', 'unexpected', 'X')).status, 400);
        const invalid = await request(h.url + '/api/convert-excel', {
            method: 'POST', headers: { host: 'dsone.shop', 'content-type': 'application/json' }, body: '{}',
        });
        assert.equal(invalid.status, 400);
        const malformed = await request(h.url + '/api/download-basic-excel', {
            method: 'POST', headers: { host: 'dsone.shop', 'content-type': 'application/json' }, body: '{',
        });
        assert.equal(malformed.status, 400);
        await empty(h.directory);
        for (const route of ['/uploads/job-guess/result.xlsx', '/.env', '/server.js', '/request-isolation.js']) {
            assert.equal((await request(h.url + route, { headers: { host: 'dsone.shop' } })).status, 404);
        }
    } finally { await h.close(); }
});

test('disconnecting a running conversion cancels its work and cleans only its own folder', async () => {
    let entered;
    const started = new Promise(resolve => { entered = resolve; });
    const h = await harness(async (script, args, workspace) => {
        await fs.writeFile(path.join(workspace.directory, 'private.txt'), 'ABORTED_REQUEST_ONLY');
        entered();
        await new Promise((resolve, reject) => workspace.controller.signal.addEventListener('abort', () => reject(new Error('aborted')), { once: true }));
    });
    try {
        const request = http.request(h.url + '/api/download-basic-excel', {
            method: 'POST', headers: { host: 'dsone.shop', 'content-type': 'application/json' },
        });
        request.on('error', () => {});
        request.end(JSON.stringify({ data: [{ floor: 'ABORT' }] }));
        await started;
        request.destroy();
        await empty(h.directory);
    } finally { await h.close(); }
});

test('ERP configuration fails closed and child processes receive no database, session or AI credentials', () => {
    for (const key of ['DATABASE_URL', 'DATABASE_AUTH_TOKEN', 'DATABASE_READ_AUTH_TOKEN', 'DB_PATH']) {
        assert.throws(() => createApp({ env: { NODE_ENV: 'production', [key]: 'SECRET' } }), /must not receive/);
    }
    const child = pythonEnvironment({ directory: 'REQUEST_PRIVATE_DIR' }, {
        PATH: 'python-path', DATABASE_URL: 'SECRET', DATABASE_AUTH_TOKEN: 'SECRET', SESSION_SECRET: 'SECRET',
        ANTHROPIC_API_KEY: 'SECRET', GEMINI_API_KEY: 'SECRET', VERCEL_TOKEN: 'SECRET',
    });
    assert.equal(child.PATH, 'python-path');
    for (const key of ['DATABASE_URL', 'DATABASE_AUTH_TOKEN', 'SESSION_SECRET', 'ANTHROPIC_API_KEY', 'GEMINI_API_KEY', 'VERCEL_TOKEN']) assert.equal(child[key], undefined);
    assert.equal(child.TEMP, 'REQUEST_PRIVATE_DIR');
    assert.equal(child.TMPDIR, 'REQUEST_PRIVATE_DIR');
});

test('aborted multipart uploads close their file streams and remove partial files', async () => {
    let executions = 0;
    const h = await harness(async () => { executions++; });
    try {
        const request = http.request(h.url + '/api/ocr', {
            method: 'POST', headers: { host: 'dsone.shop', 'content-type': 'multipart/form-data; boundary=partial' },
        });
        request.on('error', () => {});
        request.write('--partial\r\nContent-Disposition: form-data; name="image"; filename="private.png"\r\nContent-Type: image/png\r\n\r\n');
        request.write(Buffer.alloc(256 * 1024, 65));
        let written = false;
        for (let i = 0; i < 100 && !written; i++) {
            const directories = await fs.readdir(h.directory);
            for (const directory of directories) {
                const files = await fs.readdir(path.join(h.directory, directory));
                if (files.length) written = true;
            }
            if (!written) await new Promise(resolve => setTimeout(resolve, 10));
        }
        assert(written, 'partial upload must reach disk before cancellation');
        request.destroy();
        await empty(h.directory);
        assert.equal(executions, 0);
    } finally { await h.close(); }
});

test('actual Python processes use private working/temp folders, suppress stderr, and exit before cleanup on timeout', async () => {
    await fs.mkdir(root, { recursive: true });
    const job = createWorkspace(root);
    try {
        const script = path.join(job.directory, 'environment-check.py');
        await fs.writeFile(script, 'import os, json, sys\nprint("SECRET_CUSTOMER", file=sys.stderr)\nprint(json.dumps({"cwd":os.getcwd(),"temp":os.environ["TEMP"],"db":os.environ.get("DATABASE_AUTH_TOKEN")}))\n');
        const result = JSON.parse(await runPython(script, [], job, { env: { ...process.env, DATABASE_AUTH_TOKEN: 'DO_NOT_INHERIT' } }));
        assert.equal(result.cwd, job.directory);
        assert.equal(result.temp, job.directory);
        assert.equal(result.db, null);
        const slow = path.join(job.directory, 'slow.py');
        await fs.writeFile(slow, 'import time\ntime.sleep(30)\n');
        await assert.rejects(runPython(slow, [], job, { timeoutMs: 250 }), error => error.status === 504);
    } finally { await job.cleanup(); }
});
