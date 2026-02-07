const body = document.body;
const leftToggleBtn = document.getElementById('toggle-left');
const rightToggleBtn = document.getElementById('toggle-right');
const leftPanel = document.querySelector('.panel-left');
const canvasSurface = document.getElementById('canvas-surface');
const canvasContent = document.getElementById('canvas-content');
const tileLibrary = document.getElementById('tile-library');
const jsonInput = document.getElementById('json-input');
const jsonImportBtn = document.getElementById('json-import');
const jsonExportBtn = document.getElementById('json-export');
const canvasBgInput = document.getElementById('canvas-bg');
const canvasGridInput = document.getElementById('canvas-grid');
const tileStyleFillInput = document.getElementById('tile-style-fill');
const tileStyleFillHexInput = document.getElementById('tile-style-fill-hex');
const tileStyleBorderInput = document.getElementById('tile-style-border');
const tileStyleBorderHexInput = document.getElementById('tile-style-border-hex');
const tileStyleBorderWidthInput = document.getElementById('tile-style-border-width');
const tileStyleRadiusInput = document.getElementById('tile-style-radius');
const tileStyleShapeInput = document.getElementById('tile-style-shape');
const tileStyleOpacityInput = document.getElementById('tile-style-opacity');
const tileStyleOpacityValueInput = document.getElementById('tile-style-opacity-value');
const tileStylePaddingInput = document.getElementById('tile-style-padding');
const tileStyleShadowInput = document.getElementById('tile-style-shadow');
const tileTextFontInput = document.getElementById('tile-text-font');
const tileTextSizeInput = document.getElementById('tile-text-size');
const tileTextWeightInput = document.getElementById('tile-text-weight');
const tileTextColorInput = document.getElementById('tile-text-color');
const tileTextAlignInput = document.getElementById('tile-text-align');
const tileTextLineHeightInput = document.getElementById('tile-text-line-height');
const tileTextLetterSpacingInput = document.getElementById('tile-text-letter-spacing');
const kpiCellIndexInput = document.getElementById('kpi-cell-index');
const kpiCellFillInput = document.getElementById('kpi-cell-fill');
const kpiCellBorderInput = document.getElementById('kpi-cell-border');
const kpiCellBorderWidthInput = document.getElementById('kpi-cell-border-width');
const kpiCellRadiusInput = document.getElementById('kpi-cell-radius');
const kpiCellShapeInput = document.getElementById('kpi-cell-shape');
const kpiCellTextInput = document.getElementById('kpi-cell-text');
const kpiTitleColorInput = document.getElementById('kpi-title-color');
const kpiValueColorInput = document.getElementById('kpi-value-color');
const kpiDeltaColorInput = document.getElementById('kpi-delta-color');
const kpiTitleSizeInput = document.getElementById('kpi-title-size');
const kpiValueSizeInput = document.getElementById('kpi-value-size');
const kpiDeltaSizeInput = document.getElementById('kpi-delta-size');
const kpiTitleWeightInput = document.getElementById('kpi-title-weight');
const kpiValueWeightInput = document.getElementById('kpi-value-weight');
const kpiDeltaWeightInput = document.getElementById('kpi-delta-weight');
const tileArrXInput = document.getElementById('tile-arr-x');
const tileArrYInput = document.getElementById('tile-arr-y');
const tileArrWInput = document.getElementById('tile-arr-w');
const tileArrHInput = document.getElementById('tile-arr-h');
const graphJsonInput = document.getElementById('graph-json');
const graphApplyBtn = document.getElementById('graph-apply');
const tableJsonInput = document.getElementById('table-json');
const tableApplyBtn = document.getElementById('table-apply');
const graphSection = document.getElementById('graph-section');
const tableSection = document.getElementById('table-section');
const bigStatSection = document.getElementById('bigstat-section');
const kpiSection = document.getElementById('kpi-section');
const demographicsSection = document.getElementById('demographics-section');
const tileArrFrontBtn = document.getElementById('tile-arr-front');
const tileArrForwardBtn = document.getElementById('tile-arr-forward');
const tileArrBackwardBtn = document.getElementById('tile-arr-backward');
const tileArrBackBtn = document.getElementById('tile-arr-back');
const tileArrDupBtn = document.getElementById('tile-arr-dup');
const tileArrDelBtn = document.getElementById('tile-arr-del');
const selectionLabel = document.getElementById('selection-label');
const canvasGuides = document.getElementById('canvas-guides');
const rightPanel = document.querySelector('.panel-right');
const resizeHandle = document.getElementById('resize-handle');
const undoBtn = document.getElementById('btn-undo');
const redoBtn = document.getElementById('btn-redo');
const saveBtn = document.getElementById('save-btn');
const exportTriggerBtn = document.getElementById('export-trigger');
const exportMenu = document.getElementById('export-menu');
const historyToggleBtn = document.getElementById('history-toggle');
const historyMenu = document.getElementById('history-menu');
const projectNameInput = document.getElementById('project-name');
const projectList = document.getElementById('project-list');
const newProjectBtn = document.getElementById('action-new-project');
const importJsonNavBtn = document.getElementById('action-import-json');
const projectSubtitle = document.getElementById('project-subtitle');
const sidebarProjectTitle = document.getElementById('sidebar-project-title');
const canvasGridSizeInput = document.getElementById('canvas-grid-size');
const bigStatHeaderInput = document.getElementById('bigstat-header');
const bigStatLabelInput = document.getElementById('bigstat-label');
const bigStatValueInput = document.getElementById('bigstat-value');
const bigStatValueSizeInput = document.getElementById('bigstat-value-size');
const bigStatLabelSizeInput = document.getElementById('bigstat-label-size');
const bigStatValueAlignInput = document.getElementById('bigstat-value-align');
const bigStatLabelAlignInput = document.getElementById('bigstat-label-align');
const kpiAddBtn = document.getElementById('kpi-add');
const kpiRemoveBtn = document.getElementById('kpi-remove');
const demoAddBtn = document.getElementById('demo-add');
const demoRemoveBtn = document.getElementById('demo-remove');

const exportPayload = window.__EXPORT_DATA__ ?? null;
const isExport = Boolean(exportPayload && typeof exportPayload === 'object');

let setActiveTab = null;
let lastInspectorTileId = null;

const HISTORY_LIMIT = 50;
let history = [];
let historyIndex = -1;
let isApplyingHistory = false;
let historyTimer = null;
let pendingHistoryName = null;

const normalizeHex = (value) => {
  if (!value) return null;
  let hex = value.trim();
  if (!hex) return null;
  if (!hex.startsWith('#')) hex = `#${hex}`;
  if (/^#([0-9a-fA-F]{3})$/.test(hex)) {
    const short = hex.slice(1);
    hex = `#${short
      .split('')
      .map((ch) => ch + ch)
      .join('')}`;
  }
  if (/^#([0-9a-fA-F]{6})$/.test(hex)) return hex.toLowerCase();
  return null;
};

const deepClone = (value) => JSON.parse(JSON.stringify(value));

const updateUndoRedoButtons = () => {
  if (undoBtn) undoBtn.disabled = historyIndex <= 0;
  if (redoBtn) redoBtn.disabled = historyIndex >= history.length - 1;
};

const formatDate = (value) => {
  if (!value) return '';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
};

const renderProjectList = (projects = []) => {
  if (!projectList) return;
  projectList.innerHTML = '';
  projects.forEach((project) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'project-item';
    if (String(project.id) === String(state.project?.id)) {
      button.classList.add('active');
    }
    button.dataset.projectId = project.id;
    button.innerHTML = `
      ${project.name || 'Untitled'}
      <span class="meta">${formatDate(project.updatedAt)}</span>
    `;
    projectList.appendChild(button);
  });
};

const loadProjectList = async () => {
  if (isExport || window.location.protocol === 'file:') return;
  if (!projectList) return;
  try {
    const res = await fetch('/api/projects');
    if (!res.ok) return;
    const payload = await res.json();
    renderProjectList(payload.projects || []);
  } catch (err) {
    console.warn('Failed to load project list', err);
  }
};

const updateHistoryMenu = () => {
  if (!historyMenu) return;
  historyMenu.innerHTML = '';
  const items = history.slice().reverse().slice(0, 10);
  items.forEach((entry, idx) => {
    const actualIndex = history.length - 1 - idx;
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'history-item';
    if (actualIndex === historyIndex) button.classList.add('active');
    button.dataset.index = actualIndex.toString();
    button.textContent = entry.name || `Step ${actualIndex + 1}`;
    historyMenu.appendChild(button);
  });
};

const pushHistory = (name = 'Update') => {
  if (isApplyingHistory) return;
  const snapshot = deepClone(state);
  const serialized = JSON.stringify(snapshot);
  const currentSerialized = history[historyIndex]?.serialized;
  if (serialized === currentSerialized) return;
  history = history.slice(0, historyIndex + 1);
  history.push({ name, state: snapshot, serialized, at: Date.now() });
  if (history.length > HISTORY_LIMIT) {
    history.shift();
  }
  historyIndex = history.length - 1;
  updateUndoRedoButtons();
  updateHistoryMenu();
};

const scheduleHistory = (name = 'Update') => {
  if (isApplyingHistory) return;
  pendingHistoryName = name;
  if (historyTimer) clearTimeout(historyTimer);
  historyTimer = setTimeout(() => {
    pushHistory(pendingHistoryName);
    pendingHistoryName = null;
  }, 200);
};

const applyHistorySnapshot = (snapshot) => {
  isApplyingHistory = true;
  replaceState(snapshot);
  render();
  isApplyingHistory = false;
  updateUndoRedoButtons();
  updateHistoryMenu();
  updateProjectNameDisplay();
};

const STORAGE_KEYS = {
  hideLeft: 'dashboardStudioV2.hideLeft',
  hideRight: 'dashboardStudioV2.hideRight',
  rightWidth: 'dashboardStudioV2.rightWidth',
  activeTab: 'dashboardStudioV2.activeTab',
  projectId: 'dashboardStudioV2.projectId',
  gridSize: 'dashboardStudioV2.gridSize'
};

const MIN_TILE_WIDTH = 200;
const MIN_TILE_HEIGHT = 120;
const MAX_KPI_CELLS = 6;
const MAX_DEMO_ITEMS = 10;
const DEMO_COLUMNS = 5;
let lastRoundedRadius = 12;
const kpiCellRadiusMemory = {};

const DEFAULT_PROJECT_NAME = 'Dashboard Studio V2';

const DEFAULT_STATE = {
  project: {
    id: null,
    name: DEFAULT_PROJECT_NAME
  },
  canvas: {
    width: 1600,
    height: 1000,
    background: '#0b0f14',
    gridOpacity: 0.08,
    gridSize: 12,
    zoom: 1
  },
  theme: {
    fontFamily:
      'SF Pro Text, SF Pro Display, -apple-system, BlinkMacSystemFont, Segoe UI, Helvetica Neue, Arial, sans-serif',
    tileBackground: '#0f172a',
    tileBorder: '#2b3445',
    textColor: '#e5e7f0',
    accent: '#3b82f6'
  },
  tiles: [],
  selectedTileId: null
};

const state = deepClone(DEFAULT_STATE);

const normalizeState = (next = {}) => {
  const merged = deepClone(DEFAULT_STATE);
  if (next && typeof next === 'object') {
    Object.assign(merged, next);
    merged.project = { ...DEFAULT_STATE.project, ...(next.project || {}) };
    merged.canvas = { ...DEFAULT_STATE.canvas, ...(next.canvas || {}) };
    merged.theme = { ...DEFAULT_STATE.theme, ...(next.theme || {}) };
    if (!Array.isArray(merged.tiles)) merged.tiles = [];
    if (typeof merged.selectedTileId === 'undefined') merged.selectedTileId = null;
  }
  return merged;
};

const replaceState = (next) => {
  const normalized = normalizeState(next);
  Object.keys(state).forEach((key) => {
    delete state[key];
  });
  Object.assign(state, normalized);
};

const getProjectName = (useFallback = true) => {
  const name = state.project?.name;
  if (!useFallback) return name ?? '';
  const trimmed = typeof name === 'string' ? name.trim() : '';
  return trimmed || DEFAULT_PROJECT_NAME;
};

const updateActiveProjectItemName = (name) => {
  if (!projectList) return;
  const activeItem = projectList.querySelector('.project-item.active');
  if (!activeItem) return;
  const meta = activeItem.querySelector('.meta');
  activeItem.textContent = name;
  if (meta) activeItem.appendChild(meta);
};

const updateProjectListSelection = () => {
  if (!projectList) return;
  const currentId = state.project?.id;
  projectList.querySelectorAll('.project-item').forEach((item) => {
    const isActive = currentId && String(item.dataset.projectId) === String(currentId);
    item.classList.toggle('active', Boolean(isActive));
  });
};

const updateProjectNameLabel = () => {
  const name = getProjectName(true);
  if (projectSubtitle) projectSubtitle.textContent = name;
  if (sidebarProjectTitle) sidebarProjectTitle.textContent = name;
  if (name) document.title = name;
};

const updateProjectNameDisplay = () => {
  const name = getProjectName(true);
  if (projectNameInput && projectNameInput.value !== name) {
    projectNameInput.value = name;
  }
  updateProjectNameLabel();
  updateProjectListSelection();
  updateActiveProjectItemName(name);
};

const clampGridSize = (value) => {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return DEFAULT_STATE.canvas.gridSize;
  const rounded = Math.round(parsed);
  return Math.min(100, Math.max(1, rounded));
};

const getGridSize = () => clampGridSize(state.canvas.gridSize ?? DEFAULT_STATE.canvas.gridSize);

const updateGridSizeDisplay = () => {
  const gridSize = getGridSize();
  if (canvasGridSizeInput && Number(canvasGridSizeInput.value) !== gridSize) {
    canvasGridSizeInput.value = String(gridSize);
  }
  document.documentElement.style.setProperty('--grid-size', `${gridSize}px`);
};

const clampFontSize = (value, min, max, fallback) => {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, Math.round(parsed)));
};

const createDefaultKpiItem = (index = 0) => {
  const base = DEFAULT_TILES['kpi-row']?.data?.items?.[0] || {};
  const next = JSON.parse(JSON.stringify(base));
  next.title = `KPI ${index + 1}`;
  next.value = next.value || '0';
  next.delta = next.delta || '+0%';
  return next;
};

const createDefaultDemoItem = (index = 0) => ({
  label: `Country ${index + 1}`,
  value: '0'
});

const DEFAULT_TILES = {
  'kpi-row': {
    size: { w: 880, h: 140 },
    data: {
      items: [
        {
          title: 'Active Users',
          value: '4,200',
          delta: '+10.5%',
          fill: '#101723',
          border: '#1f2937',
          borderWidth: 1,
          radius: 10,
          textColor: '#e6e7eb',
          titleColor: '#8e94a3',
          valueColor: '#e6e7eb',
          deltaColor: '#6ea3ff',
          titleSize: 10,
          valueSize: 20,
          deltaSize: 11,
          titleWeight: 500,
          valueWeight: 600,
          deltaWeight: 500
        },
        {
          title: 'New Users',
          value: '2,400',
          delta: '-20.0%',
          fill: '#101723',
          border: '#1f2937',
          borderWidth: 1,
          radius: 10,
          textColor: '#e6e7eb',
          titleColor: '#8e94a3',
          valueColor: '#e6e7eb',
          deltaColor: '#6ea3ff',
          titleSize: 10,
          valueSize: 20,
          deltaSize: 11,
          titleWeight: 500,
          valueWeight: 600,
          deltaWeight: 500
        },
        {
          title: 'Search Views',
          value: '13,600',
          delta: '+13.3%',
          fill: '#101723',
          border: '#1f2937',
          borderWidth: 1,
          radius: 10,
          textColor: '#e6e7eb',
          titleColor: '#8e94a3',
          valueColor: '#e6e7eb',
          deltaColor: '#6ea3ff',
          titleSize: 10,
          valueSize: 20,
          deltaSize: 11,
          titleWeight: 500,
          valueWeight: 600,
          deltaWeight: 500
        },
        {
          title: 'Social Views',
          value: '2,200',
          delta: '+10.0%',
          fill: '#101723',
          border: '#1f2937',
          borderWidth: 1,
          radius: 10,
          textColor: '#e6e7eb',
          titleColor: '#8e94a3',
          valueColor: '#e6e7eb',
          deltaColor: '#6ea3ff',
          titleSize: 10,
          valueSize: 20,
          deltaSize: 11,
          titleWeight: 500,
          valueWeight: 600,
          deltaWeight: 500
        }
      ]
    }
  },
  'big-stat': {
    size: { w: 520, h: 200 },
    data: { header: 'Big Stat', label: 'Total Blogs Released', value: '13' }
  },
  highlights: {
    size: { w: 520, h: 260 },
    data: {
      title: 'Monthly Highlights',
      items: [
        'Month-over-month user activity grew through organic search.',
        'Returning users increased by 8.5% after onboarding refresh.',
        'Community outreach programs boost engagement in Q3.'
      ]
    }
  },
  blogs: {
    size: { w: 520, h: 240 },
    data: {
      title: 'New Monthly Blogs Traffic',
      subtitle: 'Sorted by Views',
      items: [
        { title: 'Power Up Qwen 3 with AMD', views: '6.5K' },
        { title: 'Llama 4 Developer Quickstart', views: '1.5K' },
        { title: 'Supercharge DeepSeek-R1', views: '1.3K' }
      ]
    }
  },
  graph: {
    size: { w: 520, h: 240 },
    data: {
      title: 'Traffic Trend',
      subtitle: 'Last 8 Weeks',
      points: [120, 160, 140, 210, 180, 240, 220, 260],
      xAxis: ['W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8'],
      yAxis: { show: true, min: 0, max: 300, ticks: 4, showGrid: true },
      style: {
        strokeWidth: 3,
        lineColor: '#60a5fa',
        fillColor: 'rgba(96, 165, 250, 0.2)',
        gridColor: 'rgba(148, 163, 184, 0.2)'
      }
    }
  },
  demographics: {
    size: { w: 520, h: 180 },
    data: {
      title: 'Top Demographics',
      items: [
        { label: 'USA', value: '3.8K' },
        { label: 'China', value: '1.1K' },
        { label: 'Russia', value: '0.9K' }
      ]
    }
  },
  table: {
    size: { w: 620, h: 260 },
    data: {
      title: 'Table',
      columns: [
        { label: 'Name', align: 'left', width: '2fr' },
        { label: 'Value', align: 'right', width: '1fr' },
        { label: 'Delta', align: 'right', width: '1fr' }
      ],
      rows: [
        ['Active Users', '4,200', '+10.5%'],
        ['New Users', '2,400', '-20.0%'],
        ['Search Views', '13,600', '+13.3%']
      ],
      settings: { showHeader: true, rowHeight: 36, striped: true }
    }
  }
};

const SHADOW_PRESETS = {
  none: 'none',
  soft: '0 12px 24px rgba(0, 0, 0, 0.35)',
  strong: '0 20px 40px rgba(0, 0, 0, 0.45)'
};

const getTileStyle = (tile) => ({
  background: tile.style?.background ?? state.theme.tileBackground,
  border: tile.style?.border ?? state.theme.tileBorder,
  borderWidth: tile.style?.borderWidth ?? 1,
  borderRadius: tile.style?.borderRadius ?? 12,
  opacity: tile.style?.opacity ?? 1,
  padding: tile.style?.padding ?? 16,
  shadow: tile.style?.shadow ?? SHADOW_PRESETS.soft,
  textColor: tile.style?.textColor ?? state.theme.textColor,
  fontFamily: tile.style?.fontFamily ?? state.theme.fontFamily,
  fontSize: tile.style?.fontSize ?? 13,
  fontWeight: tile.style?.fontWeight ?? 500,
  textAlign: tile.style?.textAlign ?? 'left',
  lineHeight: tile.style?.lineHeight ?? 1.5,
  letterSpacing: tile.style?.letterSpacing ?? 0
});

if (isExport) {
  replaceState(exportPayload);
  document.body.classList.add('export-mode');
}

const updateToggleLabels = () => {
  if (leftToggleBtn) {
    const collapsed = leftPanel?.classList.contains('collapsed');
    leftToggleBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    leftToggleBtn.setAttribute('aria-label', collapsed ? 'Show sidebar' : 'Hide sidebar');
    leftToggleBtn.setAttribute('title', collapsed ? 'Show sidebar' : 'Hide sidebar');
    leftToggleBtn.classList.toggle('is-collapsed', collapsed);
  }
  if (rightToggleBtn) {
    const collapsed = body.classList.contains('hide-right');
    rightToggleBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    rightToggleBtn.setAttribute('aria-label', collapsed ? 'Show inspector' : 'Hide inspector');
    rightToggleBtn.setAttribute('title', collapsed ? 'Show inspector' : 'Hide inspector');
    rightToggleBtn.classList.toggle('is-collapsed', collapsed);
  }
};

const applyStoredPrefs = () => {
  const hideLeft = localStorage.getItem(STORAGE_KEYS.hideLeft) === '1';
  const hideRight = localStorage.getItem(STORAGE_KEYS.hideRight) === '1';
  const storedGrid = Number(localStorage.getItem(STORAGE_KEYS.gridSize));
  if (Number.isFinite(storedGrid)) {
    state.canvas.gridSize = clampGridSize(storedGrid);
  }
  if (leftPanel) {
    leftPanel.classList.toggle('collapsed', hideLeft);
  }
  body.classList.toggle('hide-right', hideRight);
  applySidebarWidth();
  updateToggleLabels();
  requestAnimationFrame(() => {
    updateCanvasFrame();
  });
};

const getProjectIdFromUrl = () => {
  const params = new URLSearchParams(window.location.search);
  return params.get('project');
};

const loadProjectFromServer = async () => {
  if (isExport || window.location.protocol === 'file:') return;
  const queryId = getProjectIdFromUrl();
  const storedId = localStorage.getItem(STORAGE_KEYS.projectId);
  const id = queryId || storedId;
  if (!id) return;
  try {
    const res = await fetch(`/api/load/${id}`);
    if (!res.ok) return;
    const payload = await res.json();
    if (payload && payload.project && payload.project.data) {
      replaceState(payload.project.data);
      state.project = state.project || {};
      state.project.id = payload.project.id;
      state.project.name = payload.project.name || state.project.name || DEFAULT_PROJECT_NAME;
      localStorage.setItem(STORAGE_KEYS.projectId, payload.project.id);
    }
  } catch (err) {
    console.warn('Failed to load project', err);
  }
};

const loadProjectById = async (id) => {
  if (!id || isExport || window.location.protocol === 'file:') return;
  try {
    const res = await fetch(`/api/load/${id}`);
    if (!res.ok) return;
    const payload = await res.json();
    if (payload && payload.project && payload.project.data) {
      replaceState(payload.project.data);
      state.project = state.project || {};
      state.project.id = payload.project.id;
      state.project.name = payload.project.name || state.project.name || DEFAULT_PROJECT_NAME;
      localStorage.setItem(STORAGE_KEYS.projectId, payload.project.id);
      render();
      history = [];
      historyIndex = -1;
      pushHistory('Load Project');
      updateUndoRedoButtons();
      await loadProjectList();
    }
  } catch (err) {
    console.warn('Failed to load project', err);
  }
};

const createNewProject = () => {
  replaceState(DEFAULT_STATE);
  state.project = state.project || {};
  state.project.id = null;
  state.project.name = DEFAULT_PROJECT_NAME;
  localStorage.removeItem(STORAGE_KEYS.projectId);
  localStorage.setItem(STORAGE_KEYS.gridSize, String(DEFAULT_STATE.canvas.gridSize));
  render();
  history = [];
  historyIndex = -1;
  pushHistory('New Project');
  updateUndoRedoButtons();
  updateHistoryMenu();
  updateProjectNameDisplay();
};

const snap = (value) => {
  const grid = getGridSize();
  return Math.round(value / grid) * grid;
};

const generateId = () => `tile-${Date.now()}-${Math.floor(Math.random() * 10000)}`;

const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

const getCanvasPointer = (event) => {
  const rect = canvasContent.getBoundingClientRect();
  const zoom = state.canvas.zoom || 1;
  const x = (event.clientX - rect.left) / zoom;
  const y = (event.clientY - rect.top) / zoom;
  return { x, y };
};

const getSnapTargets = (excludeId) => {
  const targets = state.tiles
    .filter((tile) => tile.id !== excludeId)
    .map((tile) => ({
      left: tile.x,
      right: tile.x + tile.width,
      top: tile.y + 0,
      bottom: tile.y + tile.height,
      centerX: tile.x + tile.width / 2,
      centerY: tile.y + tile.height / 2
    }));

  const canvasWidth = canvasContent?.offsetWidth || 0;
  const canvasHeight = canvasContent?.offsetHeight || 0;
  if (canvasWidth > 0 && canvasHeight > 0) {
    targets.push({
      left: 0,
      right: canvasWidth,
      top: 0,
      bottom: canvasHeight,
      centerX: canvasWidth / 2,
      centerY: canvasHeight / 2
    });
  }

  return targets;
};

const applyGuides = (guides) => {
  if (!canvasGuides) return;
  canvasGuides.innerHTML = '';
  guides.forEach((guide) => {
    const el = document.createElement('div');
    el.className = `guide ${guide.axis === 'x' ? 'guide-vertical' : 'guide-horizontal'}`;
    if (guide.axis === 'x') {
      el.style.left = `${guide.pos}px`;
    } else {
      el.style.top = `${guide.pos}px`;
    }
    canvasGuides.appendChild(el);
  });
};

const clearGuides = () => {
  if (canvasGuides) canvasGuides.innerHTML = '';
};

const snapAxis = ({ position, size, axis, targets, threshold = 8 }) => {
  let snapped = position;
  let minDiff = threshold;
  const guides = new Set();
  const edges = [];
  const gridSnap = snap(position);
  const gridDiff = Math.abs(position - gridSnap);

  targets.forEach((target) => {
    if (axis === 'x') {
      edges.push(target.left, target.right, target.centerX);
    } else {
      edges.push(target.top, target.bottom, target.centerY);
    }
  });

  edges.forEach((edge) => {
    const diffStart = Math.abs(position - edge);
    if (diffStart <= threshold) {
      guides.add(edge);
    }
    if (diffStart < minDiff) {
      minDiff = diffStart;
      snapped = edge;
    }
    const diffEnd = Math.abs(position + size - edge);
    if (diffEnd <= threshold) {
      guides.add(edge);
    }
    if (diffEnd < minDiff) {
      minDiff = diffEnd;
      snapped = edge - size;
    }
    const center = position + size / 2;
    const diffCenter = Math.abs(center - edge);
    if (diffCenter <= threshold) {
      guides.add(edge);
    }
    if (diffCenter < minDiff) {
      minDiff = diffCenter;
      snapped = edge - size / 2;
    }
  });

  const guideList = Array.from(guides).map((pos) => ({ axis, pos }));

  if (minDiff < threshold && minDiff < gridDiff) {
    return { value: snapped, guides: guideList };
  }
  return { value: gridSnap, guides: [] };
};

const snapSizeAxis = ({ start, size, axis, targets, threshold = 8 }) => {
  let snappedSize = size;
  let minDiff = threshold;
  const guides = new Set();
  const edges = [];
  const gridSnap = snap(size);
  const gridDiff = Math.abs(size - gridSnap);

  targets.forEach((target) => {
    if (axis === 'x') {
      edges.push(target.left, target.right);
    } else {
      edges.push(target.top, target.bottom);
    }
  });

  edges.forEach((edge) => {
    const end = start + size;
    const diff = Math.abs(end - edge);
    if (diff <= threshold) {
      guides.add(edge);
    }
    if (diff < minDiff) {
      minDiff = diff;
      snappedSize = edge - start;
    }
  });

  const guideList = Array.from(guides).map((pos) => ({ axis, pos }));
  if (minDiff < threshold && minDiff < gridDiff) {
    return { value: snappedSize, guides: guideList };
  }
  return { value: gridSnap, guides: [] };
};

const setByPath = (obj, path, value) => {
  const parts = path.split('.');
  let ref = obj;
  for (let i = 0; i < parts.length - 1; i += 1) {
    const key = parts[i];
    const idx = Number.isFinite(Number(key)) ? Number(key) : key;
    if (ref[idx] === undefined) ref[idx] = {};
    ref = ref[idx];
  }
  const lastKey = parts[parts.length - 1];
  const lastIdx = Number.isFinite(Number(lastKey)) ? Number(lastKey) : lastKey;
  ref[lastIdx] = value;
};

const applyCanvasStyles = () => {
  const gridOpacity = Number(state.canvas.gridOpacity ?? 0.08);
  const gridColor = `rgba(255, 255, 255, ${gridOpacity})`;
  document.documentElement.style.setProperty('--grid', gridColor);
  updateGridSizeDisplay();
  if (canvasSurface) {
    canvasSurface.style.background = state.canvas.background;
  }
  if (canvasBgInput) canvasBgInput.value = state.canvas.background;
  if (canvasGridInput) canvasGridInput.value = gridOpacity;
  updateProjectNameDisplay();
};

const applySidebarWidth = () => {
  const saved = Number(localStorage.getItem(STORAGE_KEYS.rightWidth));
  if (!Number.isFinite(saved) || saved <= 0) return;
  document.documentElement.style.setProperty('--right-width', `${saved}px`);
};

const updateCanvasFrame = () => {
  if (isExport) return;
  const frame = document.querySelector('.canvas-frame');
  const stage = document.querySelector('.canvas-stage');
  if (!frame || !canvasContent || !stage) return;
  const stageWidth = stage.clientWidth;
  const stageHeight = stage.clientHeight;
  const contentWidth = Math.max(canvasContent.scrollWidth, canvasContent.offsetWidth);
  const contentHeight = Math.max(canvasContent.scrollHeight, canvasContent.offsetHeight);
  const zoom = state.canvas.zoom || 1;
  const width = Math.max(contentWidth, stageWidth / zoom, 1400);
  const height = Math.max(contentHeight, stageHeight / zoom, 900);
  const grid = getGridSize();
  const alignedWidth = Math.ceil(width / grid) * grid;
  const alignedHeight = Math.ceil(height / grid) * grid;
  frame.style.width = `${alignedWidth}px`;
  frame.style.height = `${alignedHeight}px`;
  frame.style.zoom = zoom;
  canvasContent.style.width = `${alignedWidth}px`;
  canvasContent.style.height = `${alignedHeight}px`;
};

const applyInspectorValues = () => {
  const tile = state.tiles.find((item) => item.id === state.selectedTileId);
  const hasTile = Boolean(tile);
  const style = tile ? getTileStyle(tile) : null;
  const disable = (el) => {
    if (el) el.disabled = !hasTile;
  };

  [
    tileStyleFillInput,
    tileStyleFillHexInput,
    tileStyleBorderInput,
    tileStyleBorderHexInput,
    tileStyleBorderWidthInput,
    tileStyleRadiusInput,
    tileStyleShapeInput,
    tileStyleOpacityInput,
    tileStyleOpacityValueInput,
    tileStylePaddingInput,
    tileStyleShadowInput,
    tileTextFontInput,
    tileTextSizeInput,
    tileTextWeightInput,
    tileTextColorInput,
    tileTextAlignInput,
    tileTextLineHeightInput,
    tileTextLetterSpacingInput,
    tileArrXInput,
    tileArrYInput,
    tileArrWInput,
    tileArrHInput,
    tileArrFrontBtn,
    tileArrForwardBtn,
    tileArrBackwardBtn,
    tileArrBackBtn,
    tileArrDupBtn,
    tileArrDelBtn,
    bigStatHeaderInput,
    bigStatLabelInput,
    bigStatValueInput,
    bigStatValueSizeInput,
    bigStatLabelSizeInput,
    bigStatValueAlignInput,
    bigStatLabelAlignInput,
    kpiAddBtn,
    kpiRemoveBtn,
    demoAddBtn,
    demoRemoveBtn,
    kpiCellIndexInput,
    kpiCellFillInput,
    kpiCellBorderInput,
    kpiCellBorderWidthInput,
    kpiCellRadiusInput,
    kpiCellShapeInput,
    kpiCellTextInput,
    kpiTitleColorInput,
    kpiValueColorInput,
    kpiDeltaColorInput,
    kpiTitleSizeInput,
    kpiValueSizeInput,
    kpiDeltaSizeInput,
    kpiTitleWeightInput,
    kpiValueWeightInput,
    kpiDeltaWeightInput,
    graphJsonInput,
    graphApplyBtn,
    tableJsonInput,
    tableApplyBtn
  ].forEach(disable);

  if (selectionLabel) {
    selectionLabel.textContent = tile ? tile.type.replace(/-/g, ' ') : 'None';
  }
  if (bigStatSection) bigStatSection.classList.toggle('hidden', tile?.type !== 'big-stat');
  if (kpiSection) kpiSection.classList.toggle('hidden', tile?.type !== 'kpi-row');
  if (demographicsSection) demographicsSection.classList.toggle('hidden', tile?.type !== 'demographics');
  if (graphSection) graphSection.classList.toggle('hidden', tile?.type !== 'graph');
  if (tableSection) tableSection.classList.toggle('hidden', tile?.type !== 'table');
  if (!tile || !style) {
    lastInspectorTileId = null;
    return;
  }

  if (tileStyleFillInput) tileStyleFillInput.value = style.background;
  if (tileStyleFillHexInput) tileStyleFillHexInput.value = style.background;
  if (tileStyleBorderInput) tileStyleBorderInput.value = style.border;
  if (tileStyleBorderHexInput) tileStyleBorderHexInput.value = style.border;
  if (tileStyleBorderWidthInput) tileStyleBorderWidthInput.value = style.borderWidth;
  if (tileStyleRadiusInput) tileStyleRadiusInput.value = style.borderRadius;
  if (style.borderRadius > 0) lastRoundedRadius = style.borderRadius;
  if (tileStyleShapeInput) tileStyleShapeInput.value = style.borderRadius > 0 ? 'rounded' : 'square';
  if (tileStyleOpacityInput) tileStyleOpacityInput.value = style.opacity;
  if (tileStyleOpacityValueInput) tileStyleOpacityValueInput.value = Math.round(style.opacity * 100);
  if (tileStylePaddingInput) tileStylePaddingInput.value = style.padding;
  if (tileStyleShadowInput) {
    const preset = Object.entries(SHADOW_PRESETS).find(([, val]) => val === style.shadow);
    tileStyleShadowInput.value = preset ? preset[0] : 'soft';
  }

  if (tileTextFontInput) tileTextFontInput.value = style.fontFamily;
  if (tileTextSizeInput) tileTextSizeInput.value = style.fontSize;
  if (tileTextWeightInput) tileTextWeightInput.value = style.fontWeight;
  if (tileTextColorInput) tileTextColorInput.value = style.textColor;
  if (tileTextAlignInput) tileTextAlignInput.value = style.textAlign;
  if (tileTextLineHeightInput) tileTextLineHeightInput.value = style.lineHeight;
  if (tileTextLetterSpacingInput) tileTextLetterSpacingInput.value = style.letterSpacing;

  if (tileArrXInput) tileArrXInput.value = Math.round(tile.x);
  if (tileArrYInput) tileArrYInput.value = Math.round(tile.y);
  if (tileArrWInput) tileArrWInput.value = Math.round(tile.width);
  if (tileArrHInput) tileArrHInput.value = Math.round(tile.height);

  if (tile.type === 'big-stat') {
    const valueSize = clampFontSize(tile.data?.valueSize ?? style.fontSize ?? 20, 12, 120, 20);
    const labelSize = clampFontSize(
      tile.data?.labelSize ?? Math.round(valueSize * 0.55),
      8,
      60,
      12
    );
    if (bigStatHeaderInput) bigStatHeaderInput.value = tile.data?.header || 'Big Stat';
    if (bigStatLabelInput) bigStatLabelInput.value = tile.data?.label || '';
    if (bigStatValueInput) bigStatValueInput.value = tile.data?.value || '';
    if (bigStatValueSizeInput) bigStatValueSizeInput.value = String(valueSize);
    if (bigStatLabelSizeInput) bigStatLabelSizeInput.value = String(labelSize);
    if (bigStatValueAlignInput) bigStatValueAlignInput.value = tile.data?.valueAlign || style.textAlign || 'left';
    if (bigStatLabelAlignInput) bigStatLabelAlignInput.value = tile.data?.labelAlign || style.textAlign || 'left';
  }

  if (tile.type === 'graph' && graphJsonInput) {
    if (document.activeElement !== graphJsonInput || tile.id !== lastInspectorTileId) {
      graphJsonInput.value = JSON.stringify(tile.data, null, 2);
    }
  }
  if (tile.type === 'table' && tableJsonInput) {
    if (document.activeElement !== tableJsonInput || tile.id !== lastInspectorTileId) {
      tableJsonInput.value = JSON.stringify(tile.data, null, 2);
    }
  }

  if (tile.type === 'kpi-row') {
    const total = Array.isArray(tile.data?.items) ? tile.data.items.length : 0;
    if (kpiAddBtn) kpiAddBtn.disabled = total >= MAX_KPI_CELLS;
    if (kpiRemoveBtn) kpiRemoveBtn.disabled = total <= 1;
    if (kpiCellIndexInput) {
      kpiCellIndexInput.max = String(Math.max(1, Math.min(total, MAX_KPI_CELLS)));
      if (Number(kpiCellIndexInput.value) > total && total > 0) {
        kpiCellIndexInput.value = String(total);
      }
    }
    const index = Math.max(0, Number(kpiCellIndexInput?.value || 1) - 1);
    const item = tile.data?.items?.[index];
    if (item) {
      if (kpiCellFillInput) kpiCellFillInput.value = item.fill || '#101723';
      if (kpiCellBorderInput) kpiCellBorderInput.value = item.border || '#1f2937';
      if (kpiCellBorderWidthInput) kpiCellBorderWidthInput.value = item.borderWidth ?? 1;
      if (kpiCellRadiusInput) kpiCellRadiusInput.value = item.radius ?? 10;
      if (item.radius > 0) kpiCellRadiusMemory[index] = item.radius;
      if (kpiCellShapeInput) kpiCellShapeInput.value = (item.radius ?? 0) > 0 ? 'rounded' : 'square';
      if (kpiCellTextInput) kpiCellTextInput.value = item.textColor || '#e6e7eb';
      if (kpiTitleColorInput) kpiTitleColorInput.value = item.titleColor || '#8e94a3';
      if (kpiValueColorInput) kpiValueColorInput.value = item.valueColor || '#e6e7eb';
      if (kpiDeltaColorInput) kpiDeltaColorInput.value = item.deltaColor || '#6ea3ff';
      if (kpiTitleSizeInput) kpiTitleSizeInput.value = item.titleSize ?? 10;
      if (kpiValueSizeInput) kpiValueSizeInput.value = item.valueSize ?? 20;
      if (kpiDeltaSizeInput) kpiDeltaSizeInput.value = item.deltaSize ?? 11;
      if (kpiTitleWeightInput) kpiTitleWeightInput.value = item.titleWeight ?? 500;
      if (kpiValueWeightInput) kpiValueWeightInput.value = item.valueWeight ?? 600;
      if (kpiDeltaWeightInput) kpiDeltaWeightInput.value = item.deltaWeight ?? 500;
    }
  }
  if (tile.type === 'demographics') {
    const total = Array.isArray(tile.data?.items) ? tile.data.items.length : 0;
    if (demoAddBtn) demoAddBtn.disabled = total >= MAX_DEMO_ITEMS;
    if (demoRemoveBtn) demoRemoveBtn.disabled = total <= 1;
  }
  lastInspectorTileId = tile.id;
};

const createTile = (type) => {
  const blueprint = DEFAULT_TILES[type];
  if (!blueprint) return null;
  const offset = state.tiles.length * 18;
  return {
    id: generateId(),
    type,
    x: 80 + offset,
    y: 80 + offset,
    width: blueprint.size.w,
    height: blueprint.size.h,
    data: JSON.parse(JSON.stringify(blueprint.data)),
    style: {}
  };
};

const renderGraph = (data = {}) => {
  const points = Array.isArray(data.points) && data.points.length > 1 ? data.points : [0, 0];
  const xAxis = Array.isArray(data.xAxis) && data.xAxis.length ? data.xAxis : points.map((_, i) => `P${i + 1}`);
  const yAxis = data.yAxis || {};
  const style = data.style || {};
  const min = Number.isFinite(yAxis.min) ? yAxis.min : Math.min(...points);
  const max = Number.isFinite(yAxis.max) ? yAxis.max : Math.max(...points);
  const ticks = Math.max(2, Number(yAxis.ticks || 4));
  const showAxis = yAxis.show !== false;
  const showGrid = yAxis.showGrid !== false;

  const width = 520;
  const height = 160;
  const margin = { left: 40, right: 12, top: 12, bottom: 26 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const range = max - min || 1;
  const step = points.length > 1 ? plotWidth / (points.length - 1) : plotWidth;

  const lineColor = style.lineColor || '#60a5fa';
  const fillColor = style.fillColor || 'rgba(96, 165, 250, 0.15)';
  const gridColor = style.gridColor || 'rgba(148, 163, 184, 0.2)';
  const strokeWidth = style.strokeWidth || 3;

  const coords = points.map((value, index) => {
    const x = margin.left + index * step;
    const y = margin.top + (1 - (value - min) / range) * plotHeight;
    return { x, y };
  });

  const path = coords
    .map((point, index) => `${index === 0 ? 'M' : 'L'}${point.x},${point.y}`)
    .join(' ');

  const areaPath = `${path} L${margin.left + plotWidth},${margin.top + plotHeight} L${margin.left},${
    margin.top + plotHeight
  } Z`;

  const yLabels = Array.from({ length: ticks + 1 }, (_, i) => {
    const value = min + (range / ticks) * i;
    const y = margin.top + plotHeight - (plotHeight / ticks) * i;
    return { value: Math.round(value), y };
  });

  const xLabels = xAxis.slice(0, points.length);

  return `
    <div class="graph-wrap">
      <svg class="graph-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none">
        ${showGrid ? yLabels.map((tick) => `<line x1="${margin.left}" x2="${margin.left + plotWidth}" y1="${tick.y}" y2="${tick.y}" stroke="${gridColor}" stroke-width="1" />`).join('') : ''}
        ${showAxis ? yLabels.map((tick) => `<text x="${margin.left - 8}" y="${tick.y + 4}" text-anchor="end" class="graph-label">${tick.value}</text>`).join('') : ''}
        ${showAxis ? xLabels.map((label, i) => `<text x="${margin.left + i * step}" y="${height - 6}" text-anchor="middle" class="graph-label">${label}</text>`).join('') : ''}
        <path d="${areaPath}" fill="${fillColor}" opacity="0.6"></path>
        <path d="${path}" fill="none" stroke="${lineColor}" stroke-width="${strokeWidth}" />
        ${coords.map((point) => `<circle cx="${point.x}" cy="${point.y}" r="3" fill="${lineColor}" />`).join('')}
      </svg>
    </div>
  `;
};

const renderTileContent = (tile) => {
  switch (tile.type) {
    case 'kpi-row':
      {
        const columns = Math.max(1, Math.min(tile.data.items.length, MAX_KPI_CELLS));
        return `
        <div class="kpi-row" style="--kpi-columns:${columns};">
          ${tile.data.items
            .map(
              (item, idx) => `
              <div class="kpi-card" style="
                background:${item.fill || 'rgba(255,255,255,0.02)'};
                border-color:${item.border || '#1f2937'};
                border-width:${item.borderWidth ?? 1}px;
                border-radius:${item.radius ?? 10}px;
                color:${item.textColor || 'inherit'};
              ">
                <div class="kpi-title editable" contenteditable data-field="items.${idx}.title" style="
                  color:${item.titleColor || '#8e94a3'};
                  font-size:${item.titleSize ?? 10}px;
                  font-weight:${item.titleWeight ?? 500};
                ">${item.title}</div>
                <div class="kpi-value editable" contenteditable data-field="items.${idx}.value" style="
                  color:${item.valueColor || '#e6e7eb'};
                  font-size:${item.valueSize ?? 20}px;
                  font-weight:${item.valueWeight ?? 600};
                ">${item.value}</div>
                <div class="kpi-delta editable" contenteditable data-field="items.${idx}.delta" style="
                  color:${item.deltaColor || '#6ea3ff'};
                  font-size:${item.deltaSize ?? 11}px;
                  font-weight:${item.deltaWeight ?? 500};
                ">${item.delta}</div>
              </div>
            `
            )
            .join('')}
        </div>
      `;
      }
    case 'big-stat':
      {
        const style = getTileStyle(tile);
        const valueSize = clampFontSize(tile.data?.valueSize ?? style.fontSize ?? 20, 12, 120, 20);
        const labelSize = clampFontSize(tile.data?.labelSize ?? Math.round(valueSize * 0.55), 8, 60, 12);
        const headerText = tile.data?.header || 'Big Stat';
        const labelText = tile.data?.label || '';
        const valueAlign = tile.data?.valueAlign || style.textAlign || 'left';
        const labelAlign = tile.data?.labelAlign || style.textAlign || 'left';
        const headerAlign = tile.data?.headerAlign || labelAlign || style.textAlign || 'left';
        return `
          <div class="tile-header editable" contenteditable data-field="header" style="
            display:block;
            text-align:${headerAlign};
          ">${headerText}</div>
          <div class="tile-title editable" contenteditable data-field="value" style="
            font-size:${valueSize}px;
            text-align:${valueAlign};
          ">${tile.data.value}</div>
          <div class="tile-subtitle editable" contenteditable data-field="label" style="
            font-size:${labelSize}px;
            text-align:${labelAlign};
          ">${labelText}</div>
        `;
      }
    case 'highlights':
      return `
        <div class="tile-header editable" contenteditable data-field="title">${tile.data.title}</div>
        <ul>
          ${tile.data.items
            .map(
              (item, idx) => `
            <li class="editable" contenteditable data-field="items.${idx}">${item}</li>
          `
            )
            .join('')}
        </ul>
      `;
    case 'blogs':
      return `
        <div class="tile-header">
          <div>
            <div class="editable" contenteditable data-field="title">${tile.data.title}</div>
            <div class="tile-subtitle editable" contenteditable data-field="subtitle">${tile.data.subtitle}</div>
          </div>
        </div>
        <div>
          ${tile.data.items
            .map(
              (item, idx) => `
            <div class="tile-subtitle editable" contenteditable data-field="items.${idx}.title">${item.title}</div>
          `
            )
            .join('')}
        </div>
      `;
    case 'graph':
      return `
        <div class="tile-header">
          <div>
            <div class="editable" contenteditable data-field="title">${tile.data.title}</div>
            <div class="tile-subtitle editable" contenteditable data-field="subtitle">${tile.data.subtitle}</div>
          </div>
        </div>
        <div>${renderGraph(tile.data)}</div>
      `;
    case 'demographics':
      {
        const columns = Math.max(1, Math.min(tile.data.items.length, DEMO_COLUMNS));
        return `
        <div class="tile-header editable" contenteditable data-field="title">${tile.data.title}</div>
        <div class="kpi-row" style="--kpi-columns:${columns};">
          ${tile.data.items
            .map(
              (item, idx) => `
            <div class="kpi-card">
              <div class="kpi-title editable" contenteditable data-field="items.${idx}.label">${item.label}</div>
              <div class="kpi-value editable" contenteditable data-field="items.${idx}.value">${item.value}</div>
            </div>
          `
            )
            .join('')}
        </div>
      `;
      }
    case 'table':
      {
        const rawColumns = Array.isArray(tile.data.columns) ? tile.data.columns : [];
        const columns = rawColumns.map((col) =>
          col && typeof col === 'object'
            ? col
            : { label: String(col ?? ''), align: 'left', width: '1fr' }
        );
        if (columns !== rawColumns) tile.data.columns = columns;
        const rows = Array.isArray(tile.data.rows) ? tile.data.rows : [];
        const settings = tile.data.settings || {};
        const showHeader = settings.showHeader !== false;
        const rowHeight = settings.rowHeight || 36;
        const striped = settings.striped !== false;
        const templateColumns = columns
          .map((col) => (col && typeof col === 'object' ? col.width || '1fr' : '1fr'))
          .join(' ');
        return `
          <div class="tile-header editable" contenteditable data-field="title">${tile.data.title}</div>
          <div class="table-grid" style="--table-columns:${templateColumns}; --row-height:${rowHeight}px;">
            ${
              showHeader
                ? `<div class="table-row table-header">
                    ${columns
                      .map((col, idx) => {
                        const label = typeof col === 'object' ? col.label : col;
                        const align = typeof col === 'object' ? col.align || 'left' : 'left';
                        return `<div class="table-cell align-${align} editable" contenteditable data-field="columns.${idx}.label">${label}</div>`;
                      })
                      .join('')}
                  </div>`
                : ''
            }
            ${rows
              .map(
                (row, rowIdx) => `
              <div class="table-row ${striped && rowIdx % 2 === 1 ? 'striped' : ''}">
                ${row
                  .map((cell, cellIdx) => {
                    const align = columns[cellIdx] && typeof columns[cellIdx] === 'object'
                      ? columns[cellIdx].align || 'left'
                      : 'left';
                    return `<div class="table-cell align-${align} editable" contenteditable data-field="rows.${rowIdx}.${cellIdx}">${cell}</div>`;
                  })
                  .join('')}
              </div>
            `
              )
              .join('')}
          </div>
        `;
      }
    default:
      return `<div class="tile-header">${tile.type}</div>`;
  }
};

const applyTileElementPosition = (element, tile) => {
  if (!element || !tile) return;
  element.style.left = `${tile.x}px`;
  element.style.top = `${tile.y}px`;
  element.style.width = `${tile.width}px`;
  element.style.height = `${tile.height}px`;
};

const updateTileSelectionClasses = () => {
  if (!canvasContent) return;
  canvasContent.querySelectorAll('.tile').forEach((node) => {
    const isSelected = node.dataset.tileId === state.selectedTileId;
    node.classList.toggle('selected', isSelected);
  });
};

const renderTiles = () => {
  if (!canvasContent) return;
  canvasContent.querySelectorAll('.tile').forEach((node) => node.remove());
  const emptyState = canvasContent.querySelector('.empty-state');
  if (emptyState) {
    emptyState.style.display = state.tiles.length ? 'none' : 'block';
  }

  state.tiles.forEach((tile) => {
    const node = document.createElement('div');
    node.className = `tile tile-${tile.type}`;
    node.dataset.tileId = tile.id;
    if (state.selectedTileId === tile.id) node.classList.add('selected');
    applyTileElementPosition(node, tile);
    const style = getTileStyle(tile);
    node.style.background = style.background;
    node.style.borderColor = style.border;
    node.style.borderWidth = `${style.borderWidth}px`;
    node.style.borderRadius = `${style.borderRadius}px`;
    node.style.opacity = style.opacity;
    node.style.padding = `${style.padding}px`;
    node.style.boxShadow = style.shadow;
    node.style.color = style.textColor;
    node.style.fontFamily = style.fontFamily;
    node.style.fontSize = `${style.fontSize}px`;
    node.style.fontWeight = style.fontWeight;
    node.style.textAlign = style.textAlign;
    node.style.lineHeight = style.lineHeight;
    node.style.letterSpacing = `${style.letterSpacing}px`;
    node.style.zIndex = tile.zIndex ?? 3;
    node.innerHTML = `
      <div class="tile-toolbar">
        <button type="button" data-action="duplicate">Duplicate</button>
        <button type="button" data-action="remove">Remove</button>
      </div>
      ${renderTileContent(tile)}
      <div class="tile-resize" data-action="resize"></div>
    `;
    canvasContent.appendChild(node);
  });
};

const render = () => {
  applyCanvasStyles();
  renderTiles();
  applyInspectorValues();
  updateCanvasFrame();
};

const selectTile = (tileId, { renderFull = true } = {}) => {
  state.selectedTileId = tileId;
  if (tileId && typeof setActiveTab === 'function') {
    setActiveTab('design');
  }
  if (renderFull) {
    render();
  } else {
    updateTileSelectionClasses();
    applyInspectorValues();
  }
};


const attachCanvasHandlers = () => {
  if (!canvasContent) return;

  let dragState = null;

  const getTileFromEvent = (event) => {
    const tileEl = event.target.closest('.tile');
    if (!tileEl) return null;
    const tileId = tileEl.dataset.tileId;
    return state.tiles.find((t) => t.id === tileId) || null;
  };

  const onPointerDown = (event) => {
    const tile = getTileFromEvent(event);
    if (!tile) {
      if (state.selectedTileId) {
        selectTile(null, { renderFull: false });
      }
      clearGuides();
      return;
    }
    const action = event.target.dataset.action || event.target.closest('[data-action]')?.dataset.action;
    selectTile(tile.id, { renderFull: false });

    if (action === 'duplicate') {
      const newTile = { ...tile, id: generateId(), x: tile.x + 24, y: tile.y + 24 };
      state.tiles.push(newTile);
      render();
      pushHistory('Duplicate Tile');
      return;
    }
    if (action === 'remove') {
      state.tiles = state.tiles.filter((t) => t.id !== tile.id);
      if (state.selectedTileId === tile.id) state.selectedTileId = null;
      render();
      pushHistory('Delete Tile');
      return;
    }

    if (event.target.closest('.editable')) {
      return;
    }

    const pointer = getCanvasPointer(event);
    dragState = {
      tile,
      action: action === 'resize' ? 'resize' : 'move',
      element: event.target.closest('.tile'),
      pointerId: event.pointerId,
      offsetX: pointer.x - tile.x,
      offsetY: pointer.y - tile.y,
      startWidth: tile.width,
      startHeight: tile.height,
      startX: tile.x,
      startY: tile.y,
      changed: false,
      snapTargets: getSnapTargets(tile.id)
    };
    try {
      dragState.element?.setPointerCapture(event.pointerId);
    } catch (err) {
      // ignore
    }
    event.preventDefault();
  };

  const onPointerMove = (event) => {
    if (!dragState || (dragState.pointerId && dragState.pointerId !== event.pointerId)) return;
    const pointer = getCanvasPointer(event);

    if (dragState.action === 'move') {
      const proposedX = pointer.x - dragState.offsetX;
      const proposedY = pointer.y - dragState.offsetY;
      const snapX = snapAxis({
        position: proposedX,
        size: dragState.tile.width,
        axis: 'x',
        targets: dragState.snapTargets
      });
      const snapY = snapAxis({
        position: proposedY,
        size: dragState.tile.height,
        axis: 'y',
        targets: dragState.snapTargets
      });
      dragState.tile.x = Math.max(0, snapX.value);
      dragState.tile.y = Math.max(0, snapY.value);
      const guides = [...snapX.guides, ...snapY.guides];
      applyGuides(guides);
    } else {
      const rawWidth = Math.max(MIN_TILE_WIDTH, pointer.x - dragState.tile.x);
      const rawHeight = Math.max(MIN_TILE_HEIGHT, pointer.y - dragState.tile.y);
      const snapX = snapSizeAxis({
        start: dragState.tile.x,
        size: rawWidth,
        axis: 'x',
        targets: dragState.snapTargets
      });
      const snapY = snapSizeAxis({
        start: dragState.tile.y,
        size: rawHeight,
        axis: 'y',
        targets: dragState.snapTargets
      });
      dragState.tile.width = Math.max(MIN_TILE_WIDTH, snapX.value);
      dragState.tile.height = Math.max(MIN_TILE_HEIGHT, snapY.value);
      const guides = [...snapX.guides, ...snapY.guides];
      applyGuides(guides);
    }
    applyTileElementPosition(dragState.element, dragState.tile);
    dragState.changed = true;
  };

  const onPointerUp = () => {
    if (!dragState) return;
    try {
      dragState.element?.releasePointerCapture(dragState.pointerId);
    } catch (err) {
      // ignore
    }
    if (dragState.changed) {
      const actionName = dragState.action === 'resize' ? 'Resize Tile' : 'Move Tile';
      pushHistory(actionName);
    }
    dragState = null;
    clearGuides();
    renderTiles();
    updateCanvasFrame();
  };

  canvasContent.addEventListener('pointerdown', onPointerDown);
  window.addEventListener('pointermove', onPointerMove);
  window.addEventListener('pointerup', onPointerUp);

  canvasContent.addEventListener('focusout', (event) => {
    const target = event.target;
    if (!target || !target.dataset.field) return;
    const tileEl = target.closest('.tile');
    if (!tileEl) return;
    const tileId = tileEl.dataset.tileId;
    const tile = state.tiles.find((t) => t.id === tileId);
    if (!tile) return;
    setByPath(tile.data, target.dataset.field, target.textContent.trim());
    scheduleHistory('Edit Text');
  });
};

const attachInspectorHandlers = () => {
  if (canvasBgInput) {
    canvasBgInput.addEventListener('input', () => {
      state.canvas.background = canvasBgInput.value;
      applyCanvasStyles();
      scheduleHistory('Canvas Background');
    });
  }
  if (canvasGridInput) {
    canvasGridInput.addEventListener('input', () => {
      state.canvas.gridOpacity = Number(canvasGridInput.value);
      applyCanvasStyles();
      scheduleHistory('Grid Opacity');
    });
  }
  if (canvasGridSizeInput) {
    canvasGridSizeInput.addEventListener('input', () => {
      const next = clampGridSize(canvasGridSizeInput.value);
      state.canvas.gridSize = next;
      localStorage.setItem(STORAGE_KEYS.gridSize, String(next));
      updateGridSizeDisplay();
      updateCanvasFrame();
      scheduleHistory('Grid Size');
    });
  }
  const updateTileStyle = (patch, { rerender = true } = {}) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    tile.style = tile.style || {};
    Object.assign(tile.style, patch);
    if (rerender) renderTiles();
    scheduleHistory('Style Change');
  };

  const updateTileGeometry = (patch) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    Object.assign(tile, patch);
    renderTiles();
    updateCanvasFrame();
    scheduleHistory('Arrange');
  };

  const updateBigStatData = (patch) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'big-stat') return;
    tile.data = tile.data || {};
    Object.assign(tile.data, patch);
    renderTiles();
    scheduleHistory('Big Stat');
  };

  const updateKpiItems = (mutate) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'kpi-row') return;
    tile.data = tile.data || {};
    tile.data.items = Array.isArray(tile.data.items) ? tile.data.items : [];
    mutate(tile.data.items);
    renderTiles();
    applyInspectorValues();
    scheduleHistory('KPI Cells');
  };

  const updateDemoItems = (mutate) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'demographics') return;
    tile.data = tile.data || {};
    tile.data.items = Array.isArray(tile.data.items) ? tile.data.items : [];
    mutate(tile.data.items);
    renderTiles();
    applyInspectorValues();
    scheduleHistory('Demographics');
  };

  tileStyleFillInput?.addEventListener('input', () => {
    updateTileStyle({ background: tileStyleFillInput.value });
    if (tileStyleFillHexInput) tileStyleFillHexInput.value = tileStyleFillInput.value;
  });
  tileStyleFillHexInput?.addEventListener('input', () => {
    const hex = normalizeHex(tileStyleFillHexInput.value);
    if (!hex) return;
    tileStyleFillHexInput.value = hex;
    if (tileStyleFillInput) tileStyleFillInput.value = hex;
    updateTileStyle({ background: hex });
  });
  tileStyleBorderInput?.addEventListener('input', () => {
    updateTileStyle({ border: tileStyleBorderInput.value });
    if (tileStyleBorderHexInput) tileStyleBorderHexInput.value = tileStyleBorderInput.value;
  });
  tileStyleBorderHexInput?.addEventListener('input', () => {
    const hex = normalizeHex(tileStyleBorderHexInput.value);
    if (!hex) return;
    tileStyleBorderHexInput.value = hex;
    if (tileStyleBorderInput) tileStyleBorderInput.value = hex;
    updateTileStyle({ border: hex });
  });
  tileStyleBorderWidthInput?.addEventListener('input', () =>
    updateTileStyle({ borderWidth: Number(tileStyleBorderWidthInput.value || 1) })
  );
  tileStyleRadiusInput?.addEventListener('input', () => {
    const radius = Number(tileStyleRadiusInput.value || 0);
    if (radius > 0) lastRoundedRadius = radius;
    if (tileStyleShapeInput) tileStyleShapeInput.value = radius > 0 ? 'rounded' : 'square';
    updateTileStyle({ borderRadius: radius });
  });
  tileStyleShapeInput?.addEventListener('change', () => {
    const next = tileStyleShapeInput.value === 'square' ? 0 : lastRoundedRadius || 12;
    if (tileStyleRadiusInput) tileStyleRadiusInput.value = next;
    updateTileStyle({ borderRadius: Number(next) });
  });
  tileStyleOpacityInput?.addEventListener('input', () => {
    const value = Number(tileStyleOpacityInput.value || 1);
    updateTileStyle({ opacity: value });
    if (tileStyleOpacityValueInput) tileStyleOpacityValueInput.value = Math.round(value * 100);
  });
  tileStyleOpacityValueInput?.addEventListener('input', () => {
    const raw = Number(tileStyleOpacityValueInput.value || 100);
    const clamped = Math.max(0, Math.min(100, raw));
    const value = clamped / 100;
    if (tileStyleOpacityInput) tileStyleOpacityInput.value = value.toFixed(2);
    updateTileStyle({ opacity: value });
  });
  tileStylePaddingInput?.addEventListener('input', () =>
    updateTileStyle({ padding: Number(tileStylePaddingInput.value || 0) })
  );
  tileStyleShadowInput?.addEventListener('change', () =>
    updateTileStyle({ shadow: SHADOW_PRESETS[tileStyleShadowInput.value] || SHADOW_PRESETS.soft })
  );

  tileTextFontInput?.addEventListener('change', () => updateTileStyle({ fontFamily: tileTextFontInput.value }));
  tileTextSizeInput?.addEventListener('input', () =>
    updateTileStyle({ fontSize: Number(tileTextSizeInput.value || 13) })
  );
  tileTextWeightInput?.addEventListener('change', () => updateTileStyle({ fontWeight: tileTextWeightInput.value }));
  tileTextColorInput?.addEventListener('input', () => updateTileStyle({ textColor: tileTextColorInput.value }));
  tileTextAlignInput?.addEventListener('change', () => updateTileStyle({ textAlign: tileTextAlignInput.value }));
  tileTextLineHeightInput?.addEventListener('input', () =>
    updateTileStyle({ lineHeight: Number(tileTextLineHeightInput.value || 1.5) })
  );
  tileTextLetterSpacingInput?.addEventListener('input', () =>
    updateTileStyle({ letterSpacing: Number(tileTextLetterSpacingInput.value || 0) })
  );

  bigStatHeaderInput?.addEventListener('input', () => updateBigStatData({ header: bigStatHeaderInput.value }));
  bigStatLabelInput?.addEventListener('input', () => updateBigStatData({ label: bigStatLabelInput.value }));
  bigStatValueInput?.addEventListener('input', () => updateBigStatData({ value: bigStatValueInput.value }));
  bigStatValueSizeInput?.addEventListener('input', () =>
    updateBigStatData({
      valueSize: clampFontSize(bigStatValueSizeInput.value, 12, 120, 20)
    })
  );
  bigStatLabelSizeInput?.addEventListener('input', () =>
    updateBigStatData({
      labelSize: clampFontSize(bigStatLabelSizeInput.value, 8, 60, 12)
    })
  );
  bigStatValueAlignInput?.addEventListener('change', () =>
    updateBigStatData({ valueAlign: bigStatValueAlignInput.value })
  );
  bigStatLabelAlignInput?.addEventListener('change', () =>
    updateBigStatData({ labelAlign: bigStatLabelAlignInput.value })
  );

  kpiAddBtn?.addEventListener('click', () => {
    updateKpiItems((items) => {
      if (items.length >= MAX_KPI_CELLS) return;
      items.push(createDefaultKpiItem(items.length));
    });
  });

  kpiRemoveBtn?.addEventListener('click', () => {
    updateKpiItems((items) => {
      if (items.length <= 1) return;
      const index = Math.max(0, Math.min(items.length - 1, Number(kpiCellIndexInput?.value || items.length) - 1));
      items.splice(index, 1);
    });
  });

  demoAddBtn?.addEventListener('click', () => {
    updateDemoItems((items) => {
      if (items.length >= MAX_DEMO_ITEMS) return;
      items.push(createDefaultDemoItem(items.length));
    });
  });

  demoRemoveBtn?.addEventListener('click', () => {
    updateDemoItems((items) => {
      if (items.length <= 1) return;
      items.pop();
    });
  });

  const updateKpiCell = (patch) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'kpi-row') return;
    const index = Math.max(0, Number(kpiCellIndexInput?.value || 1) - 1);
    if (!tile.data?.items?.[index]) return;
    Object.assign(tile.data.items[index], patch);
    renderTiles();
    scheduleHistory('KPI Update');
  };

  kpiCellIndexInput?.addEventListener('input', () => applyInspectorValues());
  kpiCellFillInput?.addEventListener('input', () =>
    updateKpiCell({ fill: kpiCellFillInput.value })
  );
  kpiCellBorderInput?.addEventListener('input', () =>
    updateKpiCell({ border: kpiCellBorderInput.value })
  );
  kpiCellBorderWidthInput?.addEventListener('input', () =>
    updateKpiCell({ borderWidth: Number(kpiCellBorderWidthInput.value || 1) })
  );
  kpiCellRadiusInput?.addEventListener('input', () => {
    const radius = Number(kpiCellRadiusInput.value || 10);
    const index = Math.max(0, Number(kpiCellIndexInput?.value || 1) - 1);
    if (radius > 0) kpiCellRadiusMemory[index] = radius;
    if (kpiCellShapeInput) kpiCellShapeInput.value = radius > 0 ? 'rounded' : 'square';
    updateKpiCell({ radius });
  });
  kpiCellShapeInput?.addEventListener('change', () => {
    const index = Math.max(0, Number(kpiCellIndexInput?.value || 1) - 1);
    const next = kpiCellShapeInput.value === 'square' ? 0 : kpiCellRadiusMemory[index] || 10;
    if (kpiCellRadiusInput) kpiCellRadiusInput.value = next;
    updateKpiCell({ radius: Number(next) });
  });
  kpiCellTextInput?.addEventListener('input', () =>
    updateKpiCell({ textColor: kpiCellTextInput.value })
  );
  kpiTitleColorInput?.addEventListener('input', () =>
    updateKpiCell({ titleColor: kpiTitleColorInput.value })
  );
  kpiValueColorInput?.addEventListener('input', () =>
    updateKpiCell({ valueColor: kpiValueColorInput.value })
  );
  kpiDeltaColorInput?.addEventListener('input', () =>
    updateKpiCell({ deltaColor: kpiDeltaColorInput.value })
  );
  kpiTitleSizeInput?.addEventListener('input', () =>
    updateKpiCell({ titleSize: Number(kpiTitleSizeInput.value || 10) })
  );
  kpiValueSizeInput?.addEventListener('input', () =>
    updateKpiCell({ valueSize: Number(kpiValueSizeInput.value || 20) })
  );
  kpiDeltaSizeInput?.addEventListener('input', () =>
    updateKpiCell({ deltaSize: Number(kpiDeltaSizeInput.value || 11) })
  );
  kpiTitleWeightInput?.addEventListener('change', () =>
    updateKpiCell({ titleWeight: Number(kpiTitleWeightInput.value || 500) })
  );
  kpiValueWeightInput?.addEventListener('change', () =>
    updateKpiCell({ valueWeight: Number(kpiValueWeightInput.value || 600) })
  );
  kpiDeltaWeightInput?.addEventListener('change', () =>
    updateKpiCell({ deltaWeight: Number(kpiDeltaWeightInput.value || 500) })
  );

  tileArrXInput?.addEventListener('input', () => updateTileGeometry({ x: Number(tileArrXInput.value || 0) }));
  tileArrYInput?.addEventListener('input', () => updateTileGeometry({ y: Number(tileArrYInput.value || 0) }));
  tileArrWInput?.addEventListener('input', () =>
    updateTileGeometry({ width: Math.max(MIN_TILE_WIDTH, Number(tileArrWInput.value || 0)) })
  );
  tileArrHInput?.addEventListener('input', () =>
    updateTileGeometry({ height: Math.max(MIN_TILE_HEIGHT, Number(tileArrHInput.value || 0)) })
  );

  tileArrFrontBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    const maxZ = Math.max(...state.tiles.map((t) => t.zIndex || 1), 1);
    tile.zIndex = maxZ + 1;
    renderTiles();
  });
  tileArrForwardBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    tile.zIndex = (tile.zIndex || 1) + 1;
    renderTiles();
  });
  tileArrBackwardBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    tile.zIndex = Math.max(1, (tile.zIndex || 1) - 1);
    renderTiles();
  });
  tileArrBackBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    tile.zIndex = 1;
    renderTiles();
  });
  tileArrDupBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    const clone = JSON.parse(JSON.stringify(tile));
    clone.id = generateId();
    clone.x += 24;
    clone.y += 24;
    state.tiles.push(clone);
    selectTile(clone.id);
  });
  tileArrDelBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    state.tiles = state.tiles.filter((t) => t.id !== tile.id);
    selectTile(null);
  });

  graphApplyBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'graph' || !graphJsonInput) return;
    try {
      const next = JSON.parse(graphJsonInput.value);
      if (next && typeof next === 'object') {
        tile.data = next;
        render();
        pushHistory('Update Graph');
      }
    } catch (err) {
      alert('Invalid graph JSON. Please fix and try again.');
    }
  });

  tableApplyBtn?.addEventListener('click', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'table' || !tableJsonInput) return;
    try {
      const next = JSON.parse(tableJsonInput.value);
      if (next && typeof next === 'object') {
        tile.data = next;
        render();
        pushHistory('Update Table');
      }
    } catch (err) {
      alert('Invalid table JSON. Please fix and try again.');
    }
  });
};

const attachLibraryHandlers = () => {
  if (!tileLibrary) return;
  tileLibrary.addEventListener('click', (event) => {
    const button = event.target.closest('[data-tile-type]');
    if (!button) return;
    const type = button.dataset.tileType;
    const tile = createTile(type);
    if (!tile) return;
    state.tiles.push(tile);
    selectTile(tile.id);
    pushHistory('Add Tile');
  });
  projectList?.addEventListener('click', (event) => {
    const button = event.target.closest('[data-project-id]');
    if (!button) return;
    const id = button.dataset.projectId;
    loadProjectById(id);
  });
};

const attachJsonHandlers = () => {
  if (jsonExportBtn) {
    jsonExportBtn.addEventListener('click', () => {
      const payload = JSON.stringify(state, null, 2);
      if (jsonInput) jsonInput.value = payload;
      downloadJson(payload);
    });
  }
  if (jsonImportBtn) {
    jsonImportBtn.addEventListener('click', () => {
      if (!jsonInput || !jsonInput.value.trim()) {
        alert('Paste JSON into the field before importing.');
        return;
      }
      try {
        const next = JSON.parse(jsonInput.value);
        if (next && typeof next === 'object') {
          replaceState(next);
          if (state.project?.id) {
            localStorage.setItem(STORAGE_KEYS.projectId, String(state.project.id));
          } else {
            localStorage.removeItem(STORAGE_KEYS.projectId);
          }
          localStorage.setItem(STORAGE_KEYS.gridSize, String(getGridSize()));
          render();
          history = [];
          historyIndex = -1;
          pushHistory('Import JSON');
          updateUndoRedoButtons();
          updateHistoryMenu();
        }
      } catch (err) {
        console.error('Invalid JSON', err);
        alert('Invalid JSON. Please fix the JSON and try again.');
      }
    });
  }
};

const attachTabHandlers = () => {
  const tabs = document.querySelectorAll('.figma-tab[data-tab]');
  const panels = document.querySelectorAll('.tab-content');
  if (!tabs.length || !panels.length) return;
  setActiveTab = (target) => {
    tabs.forEach((btn) => btn.classList.toggle('active', btn.dataset.tab === target));
    panels.forEach((panel) => {
      panel.classList.toggle('active', panel.id === `tab-${target}`);
    });
    localStorage.setItem(STORAGE_KEYS.activeTab, target);
  };

  tabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      const target = tab.dataset.tab;
      setActiveTab?.(target);
    });
  });

  const saved = localStorage.getItem(STORAGE_KEYS.activeTab);
  const initial = saved || (state.selectedTileId ? 'design' : 'assets');
  setActiveTab?.(initial);
};

const attachResizeHandler = () => {
  if (!resizeHandle || !rightPanel) return;
  let startX = 0;
  let startWidth = 0;
  const minWidth = 280;
  const maxWidth = 520;

  const onMove = (event) => {
    const delta = startX - event.clientX;
    const nextWidth = Math.min(maxWidth, Math.max(minWidth, startWidth + delta));
    document.documentElement.style.setProperty('--right-width', `${nextWidth}px`);
  };

  const onUp = () => {
    resizeHandle.classList.remove('dragging');
    const width = rightPanel.getBoundingClientRect().width;
    localStorage.setItem(STORAGE_KEYS.rightWidth, String(Math.round(width)));
    window.removeEventListener('mousemove', onMove);
    window.removeEventListener('mouseup', onUp);
  };

  resizeHandle.addEventListener('mousedown', (event) => {
    startX = event.clientX;
    startWidth = rightPanel.getBoundingClientRect().width;
    resizeHandle.classList.add('dragging');
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  });
};

const exportFromServer = async (format = 'png') => {
  if (window.location.protocol === 'file:') {
    alert('Exports require the app to run on http://localhost:3000. Please start the server and reload.');
    return;
  }
  try {
    const res = await fetch('/api/export-v2', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        format,
        data: state
      })
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(text || `Export failed (${res.status})`);
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.download =
      format === 'pdf'
        ? 'dashboard-v2.pdf'
        : format === 'png4k'
          ? 'dashboard-v2-4k.png'
          : 'dashboard-v2.png';
    link.href = url;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (err) {
    console.error('Export failed:', err);
    alert('Export failed. Check console for details.');
  }
};

const saveProject = async () => {
  if (window.location.protocol === 'file:') {
    alert('Saving requires the app to run on http://localhost:3000. Please start the server and reload.');
    return;
  }
  try {
    const res = await fetch('/api/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        id: state.project?.id ?? null,
        name: state.project?.name || DEFAULT_PROJECT_NAME,
        data: state
      })
    });
    const payload = await res.json();
    if (!res.ok) {
      throw new Error(payload?.error || 'Save failed');
    }
    if (payload && payload.id) {
      state.project = state.project || {};
      state.project.id = payload.id;
      localStorage.setItem(STORAGE_KEYS.projectId, payload.id);
    }
    await loadProjectList();
  } catch (err) {
    alert(`Save failed: ${err.message}`);
  }
};

const sanitizeFilename = (value) =>
  (value || 'dashboard')
    .toLowerCase()
    .replace(/[^a-z0-9-_]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'dashboard';

const downloadJson = (payload) => {
  const name = sanitizeFilename(state.project?.name || 'dashboard');
  const blob = new Blob([payload], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${name}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
};

const initPanning = () => {
  if (!canvasSurface) return;
  const stage = canvasSurface.querySelector('.canvas-stage');
  if (!stage) return;

  let isPanning = false;
  let startX = 0;
  let startY = 0;
  let startScrollLeft = 0;
  let startScrollTop = 0;

  const onDown = (event) => {
    if (event.button !== 0) return;
    if (event.target.closest('.tile')) return;
    isPanning = true;
    startX = event.clientX;
    startY = event.clientY;
    startScrollLeft = stage.scrollLeft;
    startScrollTop = stage.scrollTop;
    stage.setPointerCapture(event.pointerId);
  };

  const onMove = (event) => {
    if (!isPanning) return;
    const dx = event.clientX - startX;
    const dy = event.clientY - startY;
    stage.scrollLeft = startScrollLeft - dx;
    stage.scrollTop = startScrollTop - dy;
  };

  const onUp = (event) => {
    if (!isPanning) return;
    isPanning = false;
    try {
      stage.releasePointerCapture(event.pointerId);
    } catch (err) {
      // ignore
    }
  };

  stage.addEventListener('pointerdown', onDown);
  stage.addEventListener('pointermove', onMove);
  stage.addEventListener('pointerup', onUp);
  stage.addEventListener('pointerleave', onUp);

  stage.addEventListener(
    'wheel',
    (event) => {
      if (!(event.ctrlKey || event.metaKey)) return;
      event.preventDefault();
      const delta = event.deltaY > 0 ? -0.1 : 0.1;
      const nextZoom = clamp((state.canvas.zoom || 1) + delta, 0.5, 2);
      state.canvas.zoom = Number(nextZoom.toFixed(2));
      updateCanvasFrame();
    },
    { passive: false }
  );
};

const initResizeObserver = () => {
  const stage = canvasSurface?.querySelector('.canvas-stage');
  if (!stage || !canvasContent || !('ResizeObserver' in window)) return;
  const observer = new ResizeObserver(() => {
    updateCanvasFrame();
  });
  observer.observe(stage);
  observer.observe(canvasContent);
};

window.addEventListener('resize', () => {
  requestAnimationFrame(() => updateCanvasFrame());
});

if (leftPanel) {
  leftPanel.addEventListener('transitionend', (event) => {
    if (event.propertyName === 'width' || event.propertyName === 'transform' || event.propertyName === 'margin-left') {
      updateCanvasFrame();
    }
  });
}

if (leftToggleBtn && leftPanel) {
  leftToggleBtn.addEventListener('click', () => {
    const next = !leftPanel.classList.contains('collapsed');
    leftPanel.classList.toggle('collapsed', next);
    localStorage.setItem(STORAGE_KEYS.hideLeft, next ? '1' : '0');
    updateToggleLabels();
  });
}

if (rightToggleBtn) {
  rightToggleBtn.addEventListener('click', () => {
    const next = !body.classList.contains('hide-right');
    body.classList.toggle('hide-right', next);
    localStorage.setItem(STORAGE_KEYS.hideRight, next ? '1' : '0');
    updateToggleLabels();
    requestAnimationFrame(() => updateCanvasFrame());
  });
}

const boot = async () => {
  applyStoredPrefs();
  attachLibraryHandlers();
  attachCanvasHandlers();
  attachInspectorHandlers();
  attachJsonHandlers();
  attachTabHandlers();
  attachResizeHandler();
  initPanning();
  initResizeObserver();
  await loadProjectFromServer();
  await loadProjectList();
  render();
  if (!isExport) {
    pushHistory('Init');
  }
  updateUndoRedoButtons();
  updateHistoryMenu();

  let exportPrepared = false;
  let exportSize = null;

  window.__prepareExport__ = (requestedScale = 1) => {
    const scale = Number.isFinite(requestedScale) && requestedScale > 0 ? requestedScale : 1;
    if (!state.tiles.length || !canvasContent) {
      if (!exportPrepared) {
        exportSize = {
          width: canvasContent?.scrollWidth || 1400,
          height: canvasContent?.scrollHeight || 900
        };
        exportPrepared = true;
      }
      const frame = document.querySelector('.canvas-frame');
      if (frame) frame.style.zoom = scale;
      return { ...exportSize, scale };
    }
    if (!exportPrepared) {
      const stage = canvasSurface?.querySelector('.canvas-stage');
      if (stage) {
        stage.scrollTop = 0;
        stage.scrollLeft = 0;
      }
      const minX = Math.min(...state.tiles.map((t) => t.x));
      const minY = Math.min(...state.tiles.map((t) => t.y));
      const maxX = Math.max(...state.tiles.map((t) => t.x + t.width));
      const maxY = Math.max(...state.tiles.map((t) => t.y + t.height));
      const padding = 16;
      const offsetX = padding - minX;
      const offsetY = padding - minY;
      state.tiles.forEach((tile) => {
        tile.x += offsetX;
        tile.y += offsetY;
      });
      renderTiles();
      canvasContent.style.padding = '0px';
      const width = maxX - minX + padding * 2;
      const height = maxY - minY + padding * 2;
      exportSize = { width, height };
      exportPrepared = true;
      const frame = document.querySelector('.canvas-frame');
      if (frame) {
        frame.style.width = `${width}px`;
        frame.style.height = `${height}px`;
        frame.style.minWidth = '0px';
        frame.style.minHeight = '0px';
      }
      canvasContent.style.width = `${width}px`;
      canvasContent.style.height = `${height}px`;
      canvasContent.style.minWidth = '0px';
      canvasContent.style.minHeight = '0px';
    }
    const frame = document.querySelector('.canvas-frame');
    if (frame) frame.style.zoom = scale;
    return { ...exportSize, scale };
  };

  if (isExport) {
    window.__EXPORT_READY__ = true;
  }

  if (exportTriggerBtn && exportMenu) {
    exportTriggerBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      exportMenu.classList.toggle('open');
    });
    exportMenu.addEventListener('click', (event) => {
      const button = event.target.closest('button[data-export]');
      if (!button) return;
      const format = button.dataset.export;
      exportMenu.classList.remove('open');
      if (format === 'json') {
        const payload = JSON.stringify(state, null, 2);
        downloadJson(payload);
        return;
      }
      const exportFormat = format === 'pdf' ? 'pdf' : format === 'png4k' ? 'png4k' : 'png';
      exportFromServer(exportFormat);
    });
    window.addEventListener('click', () => {
      exportMenu.classList.remove('open');
    });
  }

  if (undoBtn) {
    undoBtn.addEventListener('click', () => {
      if (historyIndex <= 0) return;
      historyIndex -= 1;
      applyHistorySnapshot(history[historyIndex].state);
    });
  }

  if (redoBtn) {
    redoBtn.addEventListener('click', () => {
      if (historyIndex >= history.length - 1) return;
      historyIndex += 1;
      applyHistorySnapshot(history[historyIndex].state);
    });
  }

  if (historyToggleBtn && historyMenu) {
    historyToggleBtn.addEventListener('click', (event) => {
      event.stopPropagation();
      historyMenu.classList.toggle('open');
    });
    historyMenu.addEventListener('click', (event) => {
      const button = event.target.closest('.history-item');
      if (!button) return;
      const idx = Number(button.dataset.index);
      if (!Number.isFinite(idx)) return;
      historyIndex = idx;
      applyHistorySnapshot(history[historyIndex].state);
      historyMenu.classList.remove('open');
    });
    window.addEventListener('click', () => {
      historyMenu.classList.remove('open');
    });
  }

  if (saveBtn) {
    saveBtn.addEventListener('click', () => {
      saveProject();
    });
  }

  if (newProjectBtn) {
    newProjectBtn.addEventListener('click', () => {
      createNewProject();
    });
  }

  if (importJsonNavBtn) {
    importJsonNavBtn.addEventListener('click', () => {
      setActiveTab?.('settings');
      if (jsonInput) jsonInput.focus();
    });
  }

  if (projectNameInput) {
    projectNameInput.addEventListener('input', () => {
      const raw = projectNameInput.value;
      state.project = state.project || {};
      state.project.name = raw;
      updateProjectNameLabel();
    });
    projectNameInput.addEventListener('change', () => {
      const next = projectNameInput.value.trim() || DEFAULT_PROJECT_NAME;
      state.project = state.project || {};
      state.project.name = next;
      updateProjectNameDisplay();
      pushHistory('Rename Project');
    });
  }

  window.addEventListener('keydown', (event) => {
    if (event.metaKey || event.ctrlKey) {
      const isZ = event.key.toLowerCase() === 'z';
      if (!isZ) return;
      const target = event.target;
      const isEditable =
        target instanceof HTMLInputElement ||
        target instanceof HTMLTextAreaElement ||
        target?.isContentEditable;
      if (isEditable) return;
      event.preventDefault();
      if (event.shiftKey) {
        redoBtn?.click();
      } else {
        undoBtn?.click();
      }
    }
  });
};

boot();
