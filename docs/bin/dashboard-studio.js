#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const readline = require('readline');
const { spawn } = require('child_process');
const http = require('http');
const https = require('https');
const os = require('os');
const Database = require('better-sqlite3');
const { chromium } = require('playwright');
const { PDFDocument } = require('pdf-lib');
const { pathToFileURL } = require('url');

const colors = {
    reset: '\x1b[0m',
    bright: '\x1b[1m',
    dim: '\x1b[2m',
    red: '\x1b[31m',
    green: '\x1b[32m',
    yellow: '\x1b[33m',
    blue: '\x1b[34m',
    magenta: '\x1b[35m',
    cyan: '\x1b[36m',
    white: '\x1b[37m'
};

function log(message, color = colors.reset) {
    console.log(`${color}${message}${colors.reset}`);
}

function logSuccess(message) {
    log(`[OK] ${message}`, colors.green);
}

function logError(message) {
    log(`[ERR] ${message}`, colors.red);
}

function logWarning(message) {
    log(`[WARN] ${message}`, colors.yellow);
}

function clearScreen() {
    console.clear();
}

function printHeader(title) {
    log(`\n${'='.repeat(60)}`, colors.cyan);
    log(`  ${title}`, colors.bright);
    log(`${'='.repeat(60)}`, colors.cyan);
}

function printSubHeader(title) {
    log(`\n${'-'.repeat(50)}`, colors.dim);
    log(`  ${title}`, colors.bright);
    log(`${'-'.repeat(50)}`, colors.dim);
}

class Spinner {
    constructor(message) {
        this.message = message;
        this.frames = ['-', '\\\\', '|', '/'];
        this.frameIndex = 0;
        this.interval = null;
    }

    start() {
        process.stdout.write(`${colors.cyan}${this.frames[0]}${colors.reset} ${this.message}`);
        this.interval = setInterval(() => {
            this.frameIndex = (this.frameIndex + 1) % this.frames.length;
            process.stdout.clearLine(0);
            process.stdout.cursorTo(0);
            process.stdout.write(`${colors.cyan}${this.frames[this.frameIndex]}${colors.reset} ${this.message}`);
        }, 80);
    }

    stop(success = true) {
        if (this.interval) {
            clearInterval(this.interval);
            this.interval = null;
        }
        process.stdout.clearLine(0);
        process.stdout.cursorTo(0);
        if (success) {
            console.log(`${colors.green}[OK]${colors.reset} ${this.message}`);
        } else {
            console.log(`${colors.yellow}[WARN]${colors.reset} ${this.message}`);
        }
    }
}

function prompt(question) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
    return new Promise((resolve) => {
        rl.question(`${colors.yellow}${question}${colors.reset}`, (answer) => {
            rl.close();
            resolve(answer.trim());
        });
    });
}

function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForHealth(url, attempts = 20, delayMs = 300) {
    for (let i = 0; i < attempts; i += 1) {
        try {
            const res = await new Promise((resolve, reject) => {
                const client = url.startsWith('https') ? https : http;
                const req = client.get(url, (response) => {
                    response.resume();
                    resolve(response.statusCode);
                });
                req.on('error', reject);
            });
            if (res && res >= 200 && res < 500) {
                return true;
            }
        } catch {
            // ignore and retry
        }
        await sleep(delayMs);
    }
    return false;
}

function openBrowser(url) {
    const platform = process.platform;
    if (platform === 'darwin') {
        spawn('open', [url], { stdio: 'ignore', detached: true });
        return;
    }
    if (platform === 'win32') {
        spawn('cmd', ['/c', 'start', '', url], { stdio: 'ignore', detached: true });
        return;
    }
    spawn('xdg-open', [url], { stdio: 'ignore', detached: true });
}

function getLanUrls(port) {
    const networks = os.networkInterfaces();
    const urls = new Set();
    Object.values(networks).forEach((entries) => {
        (entries || []).forEach((entry) => {
            if (!entry || entry.internal) return;
            if (entry.family !== 'IPv4') return;
            urls.add(`http://${entry.address}:${port}`);
        });
    });
    return Array.from(urls).sort();
}

async function confirm(question) {
    const answer = await prompt(`${question} (y/N): `);
    return answer.toLowerCase() === 'y' || answer.toLowerCase() === 'yes';
}

function parseArgs(args) {
    const parsed = { _: [] };
    for (let i = 0; i < args.length; i += 1) {
        const arg = args[i];
        if (arg.startsWith('--')) {
            const key = arg.replace(/^--/, '');
            const next = args[i + 1];
            if (next && !next.startsWith('--')) {
                parsed[key] = next;
                i += 1;
            } else {
                parsed[key] = true;
            }
        } else {
            parsed._.push(arg);
        }
    }
    return parsed;
}

function parseData(payload) {
    if (!payload) return null;
    try {
        return JSON.parse(payload);
    } catch (err) {
        return null;
    }
}

function openDb() {
    const dataDir = path.join(__dirname, '..', 'data');
    fs.mkdirSync(dataDir, { recursive: true });
    const dbPath = process.env.DB_PATH || path.join(dataDir, 'app.db');
    const db = new Database(dbPath);
    db.pragma('journal_mode = WAL');
    db.pragma('foreign_keys = ON');
    db.exec(`
      CREATE TABLE IF NOT EXISTS dashboards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        type TEXT NOT NULL UNIQUE,
        data TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS dashboard_versions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        dashboard_id INTEGER NOT NULL,
        data TEXT NOT NULL,
        created_at TEXT NOT NULL,
        FOREIGN KEY (dashboard_id) REFERENCES dashboards(id) ON DELETE CASCADE
      );
    `);
    return db;
}

function nowIso() {
    return new Date().toISOString();
}

function ensureDashboard(db, type) {
    let row = db.prepare('SELECT * FROM dashboards WHERE type = ?').get(type);
    if (!row) {
        const timestamp = nowIso();
        const result = db
            .prepare('INSERT INTO dashboards (type, data, created_at, updated_at) VALUES (?, ?, ?, ?)')
            .run(type, null, timestamp, timestamp);
        row = db.prepare('SELECT * FROM dashboards WHERE id = ?').get(result.lastInsertRowid);
    }
    return {
        id: row.id,
        type: row.type,
        data: parseData(row.data),
        createdAt: row.created_at,
        updatedAt: row.updated_at
    };
}

function updateDashboard(db, id, data) {
    const updatedAt = nowIso();
    db.prepare('UPDATE dashboards SET data = ?, updated_at = ? WHERE id = ?')
        .run(JSON.stringify(data), updatedAt, id);
    return updatedAt;
}

function saveVersion(db, dashboardId, data) {
    const timestamp = nowIso();
    const result = db
        .prepare('INSERT INTO dashboard_versions (dashboard_id, data, created_at) VALUES (?, ?, ?)')
        .run(dashboardId, JSON.stringify(data), timestamp);

    db.prepare(`
      DELETE FROM dashboard_versions
      WHERE dashboard_id = ?
      AND id NOT IN (
        SELECT id FROM dashboard_versions
        WHERE dashboard_id = ?
        ORDER BY created_at DESC
        LIMIT 20
      )
    `).run(dashboardId, dashboardId);

    return { id: result.lastInsertRowid, createdAt: timestamp };
}

function listVersions(db, dashboardId) {
    return db
        .prepare('SELECT id, created_at FROM dashboard_versions WHERE dashboard_id = ? ORDER BY created_at DESC')
        .all(dashboardId)
        .map((row) => ({ id: row.id, createdAt: row.created_at }));
}

function getVersion(db, dashboardId, versionId) {
    const row = db
        .prepare('SELECT data FROM dashboard_versions WHERE id = ? AND dashboard_id = ?')
        .get(versionId, dashboardId);
    return row ? parseData(row.data) : null;
}

function listDashboards(db) {
    return db.prepare('SELECT id, type, created_at, updated_at FROM dashboards ORDER BY type').all();
}

function ensureDirForFile(filePath) {
    const dir = path.dirname(filePath);
    fs.mkdirSync(dir, { recursive: true });
}

async function exportDashboard({ mode, data, format, outPath }) {
    const browser = await chromium.launch({ headless: true });
    const page = await browser.newPage({ viewport: { width: 1500, height: 1000, deviceScaleFactor: 2 } });
    try {
        await page.addInitScript((payload) => {
            window.__EXPORT_DATA__ = payload.data;
            window.__EXPORT_MODE__ = payload.mode;
        }, { data, mode });

        const indexPath = path.join(__dirname, '..', 'public', 'index.html');
        const fileUrl = pathToFileURL(indexPath).toString();
        await page.goto(`${fileUrl}?mode=${mode}`, { waitUntil: 'networkidle' });
        await page.waitForSelector('#dashboard-preview', { timeout: 15000 });
        await page.waitForFunction(() => window.__EXPORT_READY__ === true, null, { timeout: 15000 });

        const preview = page.locator('#dashboard-preview');
        const box = await preview.boundingBox();
        if (!box) throw new Error('Preview bounding box not found');

        const clip = {
            x: Math.max(box.x, 0),
            y: Math.max(box.y, 0),
            width: Math.ceil(box.width),
            height: Math.ceil(box.height)
        };

        const pngBuffer = await page.screenshot({ type: 'png', clip });

        if (format === 'png') {
            ensureDirForFile(outPath);
            fs.writeFileSync(outPath, pngBuffer);
            return;
        }

        const pdfDoc = await PDFDocument.create();
        const pngImage = await pdfDoc.embedPng(pngBuffer);
        const { width, height } = pngImage.scale(1);
        const pdfPage = pdfDoc.addPage([width, height]);
        pdfPage.drawImage(pngImage, { x: 0, y: 0, width, height });
        const pdfBytes = await pdfDoc.save();

        ensureDirForFile(outPath);
        fs.writeFileSync(outPath, Buffer.from(pdfBytes));
    } finally {
        await page.close();
        await browser.close();
    }
}

async function runInteractive(db) {
    let running = true;
    while (running) {
        clearScreen();
        printHeader('Dashboard Studio CLI');
        log('1. Build exports (PNG + PDF)');
        log('2. List dashboards');
        log('3. View saved versions');
        log('4. Export JSON');
        log('5. Import JSON');
        log('6. Export PNG/PDF');
        log('7. Launch site (local)');
        log('8. Quick launch (local network)');
        log('9. Exit');

        const choice = await prompt('\nSelect an option: ');
        switch (choice.trim()) {
            case '1':
                await handleBuildExports(db);
                break;
            case '2':
                await handleListDashboards(db);
                break;
            case '3':
                await handleViewVersions(db);
                break;
            case '4':
                await handleExportJson(db);
                break;
            case '5':
                await handleImportJson(db);
                break;
            case '6':
                await handleExportMedia(db);
                break;
            case '7':
                await handleLaunchSite({ open: true });
                break;
            case '8':
                await handleLaunchSite({ open: true, lan: true });
                break;
            case '9':
                running = false;
                break;
            default:
                logWarning('Invalid option.');
                await pause();
                break;
        }
    }
}

async function pause() {
    await prompt('Press Enter to continue...');
}

async function promptForMode() {
    const mode = await prompt('Mode (monthly/yearly): ');
    return mode.toLowerCase().startsWith('y') ? 'yearly' : 'monthly';
}

async function handleListDashboards(db, shouldPause = true) {
    printSubHeader('Dashboards');
    const dashboards = listDashboards(db);
    if (!dashboards.length) {
        logWarning('No dashboards found.');
    } else {
        dashboards.forEach((row) => {
            log(`- ${row.type} (id ${row.id}) updated ${row.updated_at}`, colors.cyan);
        });
    }
    if (shouldPause) {
        await pause();
    }
}

async function handleViewVersions(db) {
    const mode = await promptForMode();
    const dashboard = ensureDashboard(db, mode);
    printSubHeader(`Saved Versions (${mode})`);
    const versions = listVersions(db, dashboard.id);
    if (!versions.length) {
        logWarning('No saved versions.');
    } else {
        versions.forEach((v) => {
            log(`- Version ${v.id} at ${v.createdAt}`, colors.cyan);
        });
    }
    await pause();
}

async function handleExportJson(db) {
    const mode = await promptForMode();
    const dashboard = ensureDashboard(db, mode);
    const versionInput = await prompt('Version id (blank = latest): ');
    const versionId = versionInput ? Number(versionInput) : null;
    let data = dashboard.data;
    if (versionId) {
        const versionData = getVersion(db, dashboard.id, versionId);
        if (!versionData) {
            logError('Version not found.');
            await pause();
            return;
        }
        data = versionData;
    }

    const outPath = await prompt('Output file (default: exports/' + mode + '.json): ');
    const finalPath = outPath || path.join(__dirname, '..', 'exports', `${mode}.json`);

    ensureDirForFile(finalPath);
    fs.writeFileSync(finalPath, JSON.stringify({ mode, data }, null, 2));
    logSuccess(`JSON exported to ${finalPath}`);
    await pause();
}

async function handleImportJson(db) {
    const mode = await promptForMode();
    const filePath = await prompt('JSON file path: ');
    if (!filePath) {
        logWarning('No file provided.');
        await pause();
        return;
    }
    if (!fs.existsSync(filePath)) {
        logError('File not found.');
        await pause();
        return;
    }

    let raw;
    try {
        raw = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    } catch (err) {
        logError('Invalid JSON file.');
        await pause();
        return;
    }

    const payload = raw && raw.data ? raw.data : raw;
    const dashboard = ensureDashboard(db, mode);
    updateDashboard(db, dashboard.id, payload);
    if (await confirm('Save as a version snapshot too?')) {
        saveVersion(db, dashboard.id, payload);
    }
    logSuccess('Import complete.');
    await pause();
}

async function handleExportMedia(db) {
    const mode = await promptForMode();
    const format = (await prompt('Format (png/pdf): ')).toLowerCase() === 'pdf' ? 'pdf' : 'png';
    const dashboard = ensureDashboard(db, mode);
    const versionInput = await prompt('Version id (blank = latest): ');
    const versionId = versionInput ? Number(versionInput) : null;
    let data = dashboard.data;
    if (versionId) {
        const versionData = getVersion(db, dashboard.id, versionId);
        if (!versionData) {
            logError('Version not found.');
            await pause();
            return;
        }
        data = versionData;
    }

    const defaultName = `${mode}-dashboard.${format}`;
    const outPath = await prompt(`Output file (default: exports/${defaultName}): `);
    const finalPath = outPath || path.join(__dirname, '..', 'exports', defaultName);

    const spinner = new Spinner(`Building ${format.toUpperCase()} export...`);
    spinner.start();
    try {
        await exportDashboard({ mode, data, format, outPath: finalPath });
        spinner.stop(true);
        logSuccess(`Exported to ${finalPath}`);
    } catch (err) {
        spinner.stop(false);
        logError(err.message || 'Export failed');
    }
    await pause();
}

async function handleBuildExports(db, options = {}) {
    const shouldPause = options.shouldPause !== false;
    let outDir = options.outDir;
    if (!outDir) {
        const outDirInput = await prompt('Output folder (default: exports): ');
        outDir = outDirInput || path.join(__dirname, '..', 'exports');
    }
    const spinner = new Spinner('Building monthly + yearly exports...');
    spinner.start();
    try {
        const monthly = ensureDashboard(db, 'monthly');
        const yearly = ensureDashboard(db, 'yearly');

        await exportDashboard({
            mode: 'monthly',
            data: monthly.data,
            format: 'png',
            outPath: path.join(outDir, 'monthly-dashboard.png')
        });
        await exportDashboard({
            mode: 'monthly',
            data: monthly.data,
            format: 'pdf',
            outPath: path.join(outDir, 'monthly-dashboard.pdf')
        });
        await exportDashboard({
            mode: 'yearly',
            data: yearly.data,
            format: 'png',
            outPath: path.join(outDir, 'yearly-dashboard.png')
        });
        await exportDashboard({
            mode: 'yearly',
            data: yearly.data,
            format: 'pdf',
            outPath: path.join(outDir, 'yearly-dashboard.pdf')
        });
        spinner.stop(true);
        logSuccess(`Exports saved in ${outDir}`);
    } catch (err) {
        spinner.stop(false);
        logError(err.message || 'Export failed');
    }
    if (shouldPause) {
        await pause();
    }
}

async function handleLaunchSite(options = {}) {
    const rootDir = path.join(__dirname, '..');
    const parsedPort = Number(options.port || process.env.PORT || 3000);
    const port = Number.isFinite(parsedPort) ? parsedPort : 3000;
    const configuredHost = (options.host || process.env.HOST || '127.0.0.1').toString();
    const host = options.lan ? '0.0.0.0' : configuredHost;
    const openAfter = options.open === true;
    const localUrl = `http://localhost:${port}`;
    const primaryUrl = host === '0.0.0.0' ? localUrl : `http://${host}:${port}`;
    const lanUrls = host === '0.0.0.0' ? getLanUrls(port) : [];

    printSubHeader(options.lan ? 'Quick Launch (LAN)' : 'Launching Dashboard Studio');
    log(`Server: ${primaryUrl}`, colors.cyan);
    if (lanUrls.length) {
        log('Share on your local network:', colors.cyan);
        lanUrls.forEach((url) => log(`  - ${url}`, colors.green));
    } else if (options.lan) {
        logWarning('No non-internal IPv4 interfaces found. LAN URL unavailable.');
    }
    log('Press Ctrl+C to stop the server.', colors.dim);

    const env = { ...process.env, PORT: String(port), HOST: host };
    const child = spawn('node', [path.join(rootDir, 'src', 'server.js')], {
        stdio: 'inherit',
        env
    });

    if (openAfter) {
        const healthUrl = `http://127.0.0.1:${port}/api/health`;
        const ready = await waitForHealth(healthUrl);
        if (ready) {
            openBrowser(localUrl);
        } else {
            logWarning('Server did not respond yet. Open the URL manually.');
        }
    }

    await new Promise((resolve) => {
        child.on('exit', resolve);
    });
}

function showHelp() {
    printHeader('Dashboard Studio CLI');
    log('Usage: dashboard-studio [command] [options]\n');
    log('Commands:');
    log('  menu                     Interactive mode (default)');
    log('  list                     List dashboards');
    log('  history --type <mode>    List saved versions');
    log('  export-json --type <mode> [--version <id>] [--out <file>]');
    log('  import-json --type <mode> --file <path> [--save-version]');
    log('  export-png --type <mode> [--version <id>] --out <file>');
    log('  export-pdf --type <mode> [--version <id>] --out <file>');
    log('  launch [--port <port>] [--host <host>] [--open] [--lan]');
    log('  quick-launch [--port <port>] [--open]  Start and share on local network');
    log('  launch-lan [--port <port>] [--open]    Alias for quick-launch');
    log('  build                    Export monthly+yearly PNG/PDF to exports/');
}

async function runCommand(args) {
    const parsed = parseArgs(args);
    const command = parsed._[0] || 'menu';
    const db = openDb();

    switch (command) {
        case 'menu':
            await runInteractive(db);
            break;
        case 'list':
            await handleListDashboards(db, false);
            break;
        case 'history': {
            const mode = parsed.type || parsed.mode || 'monthly';
            const dashboard = ensureDashboard(db, mode);
            printSubHeader(`Saved Versions (${mode})`);
            const versions = listVersions(db, dashboard.id);
            if (!versions.length) {
                logWarning('No saved versions.');
            } else {
                versions.forEach((v) => {
                    log(`- Version ${v.id} at ${v.createdAt}`, colors.cyan);
                });
            }
            break;
        }
        case 'export-json': {
            const mode = parsed.type || parsed.mode || 'monthly';
            const dashboard = ensureDashboard(db, mode);
            const versionId = parsed.version ? Number(parsed.version) : null;
            const data = versionId ? getVersion(db, dashboard.id, versionId) : dashboard.data;
            if (!data) {
                logError('No data found for export.');
                break;
            }
            const outPath = parsed.out || path.join(__dirname, '..', 'exports', `${mode}.json`);
            ensureDirForFile(outPath);
            fs.writeFileSync(outPath, JSON.stringify({ mode, data }, null, 2));
            logSuccess(`JSON exported to ${outPath}`);
            break;
        }
        case 'import-json': {
            const mode = parsed.type || parsed.mode || 'monthly';
            const filePath = parsed.file;
            if (!filePath) {
                logError('Missing --file <path>.');
                break;
            }
            if (!fs.existsSync(filePath)) {
                logError('File not found.');
                break;
            }
            const raw = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
            const payload = raw && raw.data ? raw.data : raw;
            const dashboard = ensureDashboard(db, mode);
            updateDashboard(db, dashboard.id, payload);
            if (parsed['save-version']) {
                saveVersion(db, dashboard.id, payload);
            }
            logSuccess('Import complete.');
            break;
        }
        case 'export-png':
        case 'export-pdf': {
            const format = command.endsWith('pdf') ? 'pdf' : 'png';
            const mode = parsed.type || parsed.mode || 'monthly';
            const dashboard = ensureDashboard(db, mode);
            const versionId = parsed.version ? Number(parsed.version) : null;
            const data = versionId ? getVersion(db, dashboard.id, versionId) : dashboard.data;
            if (!data) {
                logError('No data found for export.');
                break;
            }
            const outPath = parsed.out || path.join(__dirname, '..', 'exports', `${mode}-dashboard.${format}`);
            const spinner = new Spinner(`Building ${format.toUpperCase()} export...`);
            spinner.start();
            try {
                await exportDashboard({ mode, data, format, outPath });
                spinner.stop(true);
                logSuccess(`Exported to ${outPath}`);
            } catch (err) {
                spinner.stop(false);
                logError(err.message || 'Export failed');
            }
            break;
        }
        case 'launch': {
            const port = parsed.port ? Number(parsed.port) : (process.env.PORT ? Number(process.env.PORT) : 3000);
            await handleLaunchSite({
                open: Boolean(parsed.open),
                port: Number.isFinite(port) ? port : 3000,
                host: parsed.host || process.env.HOST || '127.0.0.1',
                lan: Boolean(parsed.lan)
            });
            break;
        }
        case 'quick-launch':
        case 'launch-lan': {
            const port = parsed.port ? Number(parsed.port) : (process.env.PORT ? Number(process.env.PORT) : 3000);
            await handleLaunchSite({
                open: parsed.open === undefined ? true : Boolean(parsed.open),
                port: Number.isFinite(port) ? port : 3000,
                lan: true
            });
            break;
        }
        case 'build':
            await handleBuildExports(db, {
                outDir: parsed.out || parsed.dir || path.join(__dirname, '..', 'exports'),
                shouldPause: false
            });
            break;
        case '-h':
        case '--help':
        case 'help':
            showHelp();
            break;
        default:
            logError('Unknown command.');
            showHelp();
            break;
    }
}

(async () => {
    const args = process.argv.slice(2);
    if (args.length === 0) {
        const db = openDb();
        await runInteractive(db);
        return;
    }
    await runCommand(args);
})();
