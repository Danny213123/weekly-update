const path = require('path');
const fs = require('fs');
const express = require('express');
const Database = require('better-sqlite3');
const { chromium } = require('playwright');
const { PDFDocument } = require('pdf-lib');

require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '2mb' }));

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
    name TEXT,
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

  CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    data TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  );
`);

const ensureColumn = (table, column, type) => {
  const columns = db.prepare(`PRAGMA table_info(${table})`).all();
  if (!columns.find((col) => col.name === column)) {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`);
  }
};

ensureColumn('dashboards', 'name', 'TEXT');

const nowIso = () => new Date().toISOString();

let browserPromise = null;

async function getBrowser() {
  if (!browserPromise) {
    browserPromise = chromium.launch({ headless: true }).catch((err) => {
      browserPromise = null;
      throw err;
    });
  }
  return browserPromise;
}

function parseData(payload) {
  if (!payload) return null;
  try {
    return JSON.parse(payload);
  } catch (err) {
    return null;
  }
}

function requireDashboardType(req, res, next) {
  const type = (req.query.type || '').toString().trim();
  if (!type) {
    return res.status(400).json({ error: 'Query param "type" is required.' });
  }
  req.dashboardType = type;
  return next();
}

app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: nowIso() });
});

app.post('/api/save', (req, res) => {
  const { id, name, data } = req.body || {};
  if (!data || typeof data !== 'object') {
    return res.status(400).json({ error: 'Body must include a "data" object.' });
  }
  const projectName = (name || 'Dashboard Studio').toString();
  const timestamp = nowIso();

  if (id) {
    const result = db
      .prepare('UPDATE projects SET name = ?, data = ?, updated_at = ? WHERE id = ?')
      .run(projectName, JSON.stringify(data), timestamp, id);
    if (result.changes > 0) {
      return res.json({ ok: true, id, updatedAt: timestamp });
    }
  }

  const insert = db
    .prepare('INSERT INTO projects (name, data, created_at, updated_at) VALUES (?, ?, ?, ?)')
    .run(projectName, JSON.stringify(data), timestamp, timestamp);
  return res.json({ ok: true, id: insert.lastInsertRowid, createdAt: timestamp, updatedAt: timestamp });
});

app.get('/api/load/:id', (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isFinite(id)) {
    return res.status(400).json({ error: 'Invalid project id.' });
  }
  const row = db.prepare('SELECT * FROM projects WHERE id = ?').get(id);
  if (!row) {
    return res.status(404).json({ error: 'Project not found.' });
  }
  return res.json({
    project: {
      id: row.id,
      name: row.name,
      data: parseData(row.data),
      createdAt: row.created_at,
      updatedAt: row.updated_at
    }
  });
});

app.get('/api/projects', (req, res) => {
  const rows = db
    .prepare('SELECT id, name, updated_at FROM projects ORDER BY updated_at DESC LIMIT 20')
    .all();
  res.json({
    projects: rows.map((row) => ({
      id: row.id,
      name: row.name,
      updatedAt: row.updated_at
    }))
  });
});

app.get('/api/dashboards', requireDashboardType, (req, res) => {
  const type = req.dashboardType;
  let row = db.prepare('SELECT * FROM dashboards WHERE type = ?').get(type);

  if (!row) {
    const timestamp = nowIso();
    const result = db
      .prepare('INSERT INTO dashboards (type, data, created_at, updated_at) VALUES (?, ?, ?, ?)')
      .run(type, null, timestamp, timestamp);
    row = db.prepare('SELECT * FROM dashboards WHERE id = ?').get(result.lastInsertRowid);
  }

  res.json({
    dashboard: {
      id: row.id,
      type: row.type,
      data: parseData(row.data),
      createdAt: row.created_at,
      updatedAt: row.updated_at
    }
  });
});

app.put('/api/dashboards/:id', (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isFinite(id)) {
    return res.status(400).json({ error: 'Invalid dashboard id.' });
  }

  const { data } = req.body || {};
  if (!data || typeof data !== 'object') {
    return res.status(400).json({ error: 'Body must include a "data" object.' });
  }

  const updatedAt = nowIso();
  const result = db
    .prepare('UPDATE dashboards SET data = ?, updated_at = ? WHERE id = ?')
    .run(JSON.stringify(data), updatedAt, id);

  if (result.changes === 0) {
    return res.status(404).json({ error: 'Dashboard not found.' });
  }

  return res.json({ ok: true, updatedAt });
});

app.get('/api/dashboards/:id/versions', (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isFinite(id)) {
    return res.status(400).json({ error: 'Invalid dashboard id.' });
  }

  const rows = db
    .prepare(
      'SELECT id, created_at FROM dashboard_versions WHERE dashboard_id = ? ORDER BY created_at DESC LIMIT 20'
    )
    .all(id);

  res.json({
    versions: rows.map(row => ({
      id: row.id,
      createdAt: row.created_at
    }))
  });
});

app.post('/api/dashboards/:id/versions', (req, res) => {
  const id = Number(req.params.id);
  if (!Number.isFinite(id)) {
    return res.status(400).json({ error: 'Invalid dashboard id.' });
  }

  const { data } = req.body || {};
  if (!data || typeof data !== 'object') {
    return res.status(400).json({ error: 'Body must include a "data" object.' });
  }

  const timestamp = nowIso();
  const result = db
    .prepare('INSERT INTO dashboard_versions (dashboard_id, data, created_at) VALUES (?, ?, ?)')
    .run(id, JSON.stringify(data), timestamp);

  db.prepare(
    `DELETE FROM dashboard_versions
     WHERE dashboard_id = ?
     AND id NOT IN (
       SELECT id FROM dashboard_versions
       WHERE dashboard_id = ?
       ORDER BY created_at DESC
       LIMIT 20
     )`
  ).run(id, id);

  res.json({
    ok: true,
    version: {
      id: result.lastInsertRowid,
      createdAt: timestamp
    }
  });
});

app.get('/api/dashboards/:id/versions/:versionId', (req, res) => {
  const dashboardId = Number(req.params.id);
  const versionId = Number(req.params.versionId);

  if (!Number.isFinite(dashboardId) || !Number.isFinite(versionId)) {
    return res.status(400).json({ error: 'Invalid dashboard or version id.' });
  }

  const row = db
    .prepare('SELECT data FROM dashboard_versions WHERE id = ? AND dashboard_id = ?')
    .get(versionId, dashboardId);

  if (!row) {
    return res.status(404).json({ error: 'Version not found.' });
  }

  return res.json({ data: parseData(row.data) });
});

app.delete('/api/dashboards/:id/versions/:versionId', (req, res) => {
  const dashboardId = Number(req.params.id);
  const versionId = Number(req.params.versionId);

  if (!Number.isFinite(dashboardId) || !Number.isFinite(versionId)) {
    return res.status(400).json({ error: 'Invalid dashboard or version id.' });
  }

  const result = db
    .prepare('DELETE FROM dashboard_versions WHERE id = ? AND dashboard_id = ?')
    .run(versionId, dashboardId);

  if (result.changes === 0) {
    return res.status(404).json({ error: 'Version not found.' });
  }

  return res.json({ ok: true });
});

app.delete('/api/dashboards/:id/versions', (req, res) => {
  const dashboardId = Number(req.params.id);
  if (!Number.isFinite(dashboardId)) {
    return res.status(400).json({ error: 'Invalid dashboard id.' });
  }

  db.prepare('DELETE FROM dashboard_versions WHERE dashboard_id = ?').run(dashboardId);
  return res.json({ ok: true });
});

app.post('/api/export', async (req, res) => {
  const { mode, data, format } = req.body || {};
  const exportMode = mode === 'yearly' ? 'yearly' : 'monthly';
  const exportFormat = format === 'pdf' ? 'pdf' : 'png';

  if (!data || typeof data !== 'object') {
    return res.status(400).json({ error: 'Body must include a "data" object.' });
  }

  const browser = await getBrowser();
  const page = await browser.newPage({
    viewport: { width: 1500, height: 1000, deviceScaleFactor: 3 }
  });

  try {
    await page.addInitScript((payload) => {
      window.__EXPORT_DATA__ = payload.data;
      window.__EXPORT_MODE__ = payload.mode;
    }, { data, mode: exportMode });

    const baseUrl = `${req.protocol}://${req.get('host')}`;
    await page.goto(`${baseUrl}/?mode=${exportMode}`, { waitUntil: 'networkidle' });
    await page.waitForSelector('#dashboard-preview', { timeout: 15000 });
    await page.waitForFunction(() => window.__EXPORT_READY__ === true, null, { timeout: 15000 });

    const exportSize = await page.evaluate(() => {
      if (typeof window.__prepareExport__ === 'function') {
        return window.__prepareExport__();
      }
      return null;
    });

    if (exportSize && exportSize.width && exportSize.height) {
      await page.setViewportSize({
        width: Math.max(1200, Math.ceil(exportSize.width)),
        height: Math.max(800, Math.ceil(exportSize.height))
      });
      await page.waitForTimeout(50);
    }

    const preview = page.locator('#dashboard-preview');
    let box = await preview.boundingBox();
    if (!box) {
      throw new Error('Preview bounding box not found.');
    }

    const viewport = page.viewportSize();
    const requiredWidth = Math.ceil(box.x + box.width);
    const requiredHeight = Math.ceil(box.y + box.height);
    if (viewport && (requiredWidth > viewport.width || requiredHeight > viewport.height)) {
      await page.setViewportSize({
        width: Math.max(viewport.width, requiredWidth),
        height: Math.max(viewport.height, requiredHeight)
      });
      await page.waitForTimeout(50);
      box = await preview.boundingBox();
      if (!box) {
        throw new Error('Preview bounding box not found.');
      }
    }

    const clip = {
      x: Math.max(box.x, 0),
      y: Math.max(box.y, 0),
      width: Math.ceil(box.width),
      height: Math.ceil(box.height)
    };

    const pngBuffer = await page.screenshot({ type: 'png', clip, omitBackground: false });

    if (exportFormat === 'png') {
      res.setHeader('Content-Type', 'image/png');
      res.setHeader('Content-Disposition', 'attachment; filename="dashboard.png"');
      return res.send(pngBuffer);
    }

    const pdfDoc = await PDFDocument.create();
    const pngImage = await pdfDoc.embedPng(pngBuffer);
    const { width, height } = pngImage.scale(1);
    const pdfPage = pdfDoc.addPage([width, height]);
    pdfPage.drawImage(pngImage, { x: 0, y: 0, width, height });
    const pdfBytes = await pdfDoc.save();

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="dashboard.pdf"');
    return res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error('Export failed:', err);
    if (!res.headersSent) {
      return res.status(500).json({ error: 'Export failed.' });
    }
  } finally {
    try {
      await page.close();
    } catch (err) {
      console.warn('Failed to close export page:', err);
    }
  }
});

app.post('/api/export-v2', async (req, res) => {
  const { data, format } = req.body || {};
  const exportFormat = format === 'pdf' ? 'pdf' : format === 'png4k' ? 'png4k' : 'png';

  if (!data || typeof data !== 'object') {
    return res.status(400).json({ error: 'Body must include a "data" object.' });
  }

  const browser = await getBrowser();

  const openExportPage = async (deviceScaleFactor, exportScale = 1) => {
    const page = await browser.newPage({
      viewport: { width: 1600, height: 1000, deviceScaleFactor }
    });

    await page.addInitScript((payload) => {
      window.__EXPORT_DATA__ = payload;
    }, data);

    const baseUrl = `${req.protocol}://${req.get('host')}`;
    await page.goto(`${baseUrl}/v2`, { waitUntil: 'networkidle' });
    await page.waitForFunction(() => window.__EXPORT_READY__ === true, null, { timeout: 15000 });

    const exportSize = await page.evaluate((scale) => {
      if (typeof window.__prepareExport__ === 'function') {
        return window.__prepareExport__(scale);
      }
      return null;
    }, exportScale);

    if (exportSize && exportSize.width && exportSize.height) {
      const scaledWidth = Math.ceil(exportSize.width * exportScale);
      const scaledHeight = Math.ceil(exportSize.height * exportScale);
      await page.setViewportSize({
        width: Math.max(1200, scaledWidth),
        height: Math.max(800, scaledHeight)
      });
      await page.waitForTimeout(50);
    }

    return { page, exportSize };
  };

  let page = null;

  try {
    if (exportFormat === 'png4k') {
      const base = await openExportPage(1, 1);
      const baseWidth = base.exportSize?.width || 1600;
      const baseHeight = base.exportSize?.height || 1000;
      const scale = Math.max(3840 / baseWidth, 2160 / baseHeight, 1);
      await base.page.close();
      page = (await openExportPage(1, scale)).page;
    } else {
      page = (await openExportPage(3, 1)).page;
    }

    const frame = page.locator('.canvas-frame');
    let box = await frame.boundingBox();
    if (!box) {
      const surface = page.locator('.canvas-surface');
      box = await surface.boundingBox();
    }
    if (!box) {
      throw new Error('Canvas frame not found.');
    }

    const clip = {
      x: Math.max(box.x, 0),
      y: Math.max(box.y, 0),
      width: Math.ceil(box.width),
      height: Math.ceil(box.height)
    };

    const pngBuffer = await page.screenshot({ type: 'png', clip, omitBackground: false });

    if (exportFormat === 'png' || exportFormat === 'png4k') {
      res.setHeader('Content-Type', 'image/png');
      res.setHeader(
        'Content-Disposition',
        exportFormat === 'png4k'
          ? 'attachment; filename="dashboard-v2-4k.png"'
          : 'attachment; filename="dashboard-v2.png"'
      );
      return res.send(pngBuffer);
    }

    const pdfDoc = await PDFDocument.create();
    const pngImage = await pdfDoc.embedPng(pngBuffer);
    const { width, height } = pngImage.scale(1);
    const pdfPage = pdfDoc.addPage([width, height]);
    pdfPage.drawImage(pngImage, { x: 0, y: 0, width, height });
    const pdfBytes = await pdfDoc.save();

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename="dashboard-v2.pdf"');
    return res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error('Export v2 failed:', err);
    if (!res.headersSent) {
      return res.status(500).json({ error: 'Export failed.' });
    }
  } finally {
    try {
      if (page) await page.close();
    } catch (err) {
      console.warn('Failed to close export page:', err);
    }
  }
});

const publicDir = path.join(__dirname, '..', 'public');
app.use(express.static(publicDir));

app.get('/yearly', (req, res) => {
  res.redirect('/?mode=yearly');
});

app.get('/monthly', (req, res) => {
  res.redirect('/?mode=monthly');
});

app.listen(PORT, () => {
  console.log(`Dashboard Builder running on http://localhost:${PORT}`);
});
