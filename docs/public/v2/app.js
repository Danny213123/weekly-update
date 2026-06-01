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
const xmlInput = document.getElementById('xml-input');
const xmlImportBtn = document.getElementById('xml-import');
const xmlExportBtn = document.getElementById('xml-export');
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
const tileStyleHighlightInput = document.getElementById('tile-style-highlight');
const tileStyleHighlightIdleInput = document.getElementById('tile-style-highlight-idle');
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
const highlightsSection = document.getElementById('highlights-section');
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
const copyProjectBtn = document.getElementById('copy-project-btn');
const deleteProjectBtn = document.getElementById('delete-project-btn');
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
const highlightsAddBtn = document.getElementById('highlights-add');
const highlightsRemoveBtn = document.getElementById('highlights-remove');
const highlightsLineIndexInput = document.getElementById('highlights-line-index');
const highlightsLineTextInput = document.getElementById('highlights-line-text');
const tileGraphicIconInput = document.getElementById('tile-graphic-icon');
const tileGraphicImageUrlInput = document.getElementById('tile-graphic-image-url');
const tileGraphicImageFileInput = document.getElementById('tile-graphic-image-file');
const tileGraphicImageFitInput = document.getElementById('tile-graphic-image-fit');
const tileGraphicImageOpacityInput = document.getElementById('tile-graphic-image-opacity');
const tileGraphicImageClearBtn = document.getElementById('tile-graphic-image-clear');
const jsonImportDataBtn = document.getElementById('json-import-data');
const jsonImportMergeBtn = document.getElementById('json-import-merge');

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
let projectCache = [];

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

const getFirstElementChild = (node) =>
  Array.from(node?.childNodes || []).find((child) => child.nodeType === Node.ELEMENT_NODE) || null;

const encodeXmlValue = (doc, value) => {
  const node = doc.createElement('value');
  if (value === null) {
    node.setAttribute('type', 'null');
    return node;
  }
  if (Array.isArray(value)) {
    node.setAttribute('type', 'array');
    value.forEach((item) => {
      const itemNode = doc.createElement('item');
      itemNode.appendChild(encodeXmlValue(doc, item));
      node.appendChild(itemNode);
    });
    return node;
  }
  const valueType = typeof value;
  if (valueType === 'object') {
    node.setAttribute('type', 'object');
    Object.entries(value).forEach(([key, item]) => {
      const fieldNode = doc.createElement('field');
      fieldNode.setAttribute('name', key);
      fieldNode.appendChild(encodeXmlValue(doc, item));
      node.appendChild(fieldNode);
    });
    return node;
  }
  if (valueType === 'number') {
    node.setAttribute('type', 'number');
    node.textContent = Number.isFinite(value) ? String(value) : '0';
    return node;
  }
  if (valueType === 'boolean') {
    node.setAttribute('type', 'boolean');
    node.textContent = value ? 'true' : 'false';
    return node;
  }
  node.setAttribute('type', 'string');
  node.textContent = String(value ?? '');
  return node;
};

const decodeXmlValue = (node) => {
  if (!node) return null;
  if (node.nodeName === 'state') {
    return decodeXmlValue(getFirstElementChild(node));
  }

  const type = node.getAttribute?.('type') || '';
  if (type === 'null') return null;
  if (type === 'number') {
    const parsed = Number(node.textContent || '0');
    return Number.isFinite(parsed) ? parsed : 0;
  }
  if (type === 'boolean') return (node.textContent || '').trim().toLowerCase() === 'true';
  if (type === 'string') return node.textContent || '';
  if (type === 'array') {
    return Array.from(node.children)
      .filter((child) => child.nodeName === 'item')
      .map((child) => decodeXmlValue(getFirstElementChild(child)));
  }
  if (type === 'object') {
    const out = {};
    Array.from(node.children)
      .filter((child) => child.nodeName === 'field')
      .forEach((field) => {
        const key = field.getAttribute('name') || '';
        if (!key) return;
        out[key] = decodeXmlValue(getFirstElementChild(field));
      });
    return out;
  }

  // Fallback for XML authored manually without type attributes.
  const objectFields = Array.from(node.children).filter(
    (child) => child.nodeName === 'field' && child.hasAttribute('name')
  );
  if (objectFields.length) {
    const out = {};
    objectFields.forEach((field) => {
      const key = field.getAttribute('name') || '';
      if (!key) return;
      out[key] = decodeXmlValue(getFirstElementChild(field));
    });
    return out;
  }

  const arrayItems = Array.from(node.children).filter((child) => child.nodeName === 'item');
  if (arrayItems.length) {
    return arrayItems.map((child) => decodeXmlValue(getFirstElementChild(child)));
  }

  return node.textContent || '';
};

const xmlNumber = (value, fallback = 0) => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
};

const xmlClamp = (value, min, max) => Math.min(max, Math.max(min, value));

const sanitizeXmlString = (value) =>
  String(value ?? '')
    // XML 1.0 legal chars: TAB, LF, CR, U+0020..U+D7FF, U+E000..U+FFFD
    .replace(/[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD]/g, '');

const sanitizeDrawioStyleValue = (value) =>
  sanitizeXmlString(value)
    .replaceAll(';', ',')
    .replaceAll('=', ':')
    .replace(/[\r\n\t]+/g, ' ')
    .trim();

const encodeMetadataJson = (value) => {
  try {
    const serialized = JSON.stringify(value ?? {});
    return encodeBase64Utf8(sanitizeXmlString(serialized));
  } catch (err) {
    return '';
  }
};

const encodeBase64Utf8 = (value) => {
  try {
    const input = String(value ?? '');
    const bytes = new TextEncoder().encode(input);
    let binary = '';
    bytes.forEach((byte) => {
      binary += String.fromCharCode(byte);
    });
    return btoa(binary);
  } catch (err) {
    return '';
  }
};

const decodeBase64Utf8 = (value) => {
  try {
    const binary = atob(String(value ?? ''));
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder().decode(bytes);
  } catch (err) {
    return '';
  }
};

const parseDrawioStyle = (styleText) => {
  const output = {};
  String(styleText || '')
    .split(';')
    .forEach((entry) => {
      const trimmed = entry.trim();
      if (!trimmed) return;
      const equalsAt = trimmed.indexOf('=');
      if (equalsAt <= 0) return;
      const key = trimmed.slice(0, equalsAt).trim();
      const value = trimmed.slice(equalsAt + 1).trim();
      if (!key) return;
      output[key] = value;
    });
  return output;
};

const toDrawioStyleString = (styleMap) => {
  const entries = Object.entries(styleMap || {}).filter(
    ([key, value]) => key && value !== undefined && value !== null && value !== ''
  );
  if (!entries.length) return '';
  return `${entries
    .map(([key, value]) => `${sanitizeXmlString(key).replace(/[;\s=]+/g, '')}=${sanitizeDrawioStyleValue(value)}`)
    .join(';')};`;
};

const normalizeDrawioColor = (value, fallback) => {
  const raw = typeof value === 'string' ? value.trim() : '';
  if (!raw || raw === 'none' || raw === 'default') return fallback;
  return normalizeHex(raw) || fallback;
};

const drawioKnownTileTypes = new Set([
  'kpi-row',
  'big-stat',
  'highlights',
  'blogs',
  'graph',
  'demographics',
  'table'
]);

const getTileLabelForDrawio = (tile) => {
  const type = String(tile?.type || '');
  const data = tile?.data || {};
  if (type === 'kpi-row') {
    const icon = typeof data.icon === 'string' ? data.icon.trim() : '';
    return icon || 'KPI';
  }
  if (type === 'demographics' || type === 'table') {
    const icon = typeof data.icon === 'string' ? data.icon.trim() : '';
    const title = typeof data.title === 'string' ? data.title.trim() : '';
    return [icon, title].filter(Boolean).join('\n');
  }

  const lines = [];
  const pushLine = (value) => {
    const text = typeof value === 'string' ? value.trim() : '';
    if (text) lines.push(text);
  };
  pushLine(data.icon);
  pushLine(data.header);
  pushLine(data.title);
  pushLine(data.subtitle);
  pushLine(data.value);
  pushLine(data.label);

  if (Array.isArray(data.items)) {
    data.items.slice(0, 6).forEach((item) => {
      if (typeof item === 'string') {
        pushLine(`- ${item}`);
        return;
      }
      if (!item || typeof item !== 'object') return;
      const left = item.title || item.label || '';
      const right = item.value || item.views || item.delta || '';
      const text = [left, right].filter(Boolean).join(': ');
      pushLine(text);
    });
  }

  if (Array.isArray(data.rows)) {
    data.rows.slice(0, 4).forEach((row) => {
      if (!Array.isArray(row)) return;
      const line = row
        .map((cell) => String(cell ?? '').trim())
        .filter(Boolean)
        .join(' | ');
      pushLine(line);
    });
  }

  if (!lines.length) {
    pushLine((tile?.type || 'tile').replaceAll('-', ' '));
  }
  return lines.slice(0, 14).join('\n');
};

const getDrawioPartStyle = ({
  fillColor,
  strokeColor,
  strokeWidth,
  borderRadius,
  width,
  height,
  fontColor,
  fontFamily,
  fontSize,
  bold = false,
  align = 'left',
  verticalAlign = 'top',
  spacing = 8
}) =>
  toDrawioStyleString({
    shape: 'rectangle',
    rounded: xmlNumber(borderRadius, 0) > 0 ? 1 : 0,
    arcSize:
      xmlNumber(borderRadius, 0) > 0 ? getDrawioArcSize(borderRadius, width, height) : null,
    whiteSpace: 'wrap',
    html: 1,
    fillColor,
    strokeColor,
    strokeWidth: Math.max(0, xmlNumber(strokeWidth, 1)),
    fontColor,
    fontFamily: fontFamily || 'Arial',
    fontSize: Math.max(8, Math.round(xmlNumber(fontSize, 12))),
    fontStyle: bold ? 1 : 0,
    align: align === 'center' || align === 'right' ? align : 'left',
    verticalAlign,
    spacing: Math.max(0, Math.round(xmlNumber(spacing, 8)))
  });

const getDrawioArcSize = (radius, width, height) => {
  const parsedRadius = Math.max(0, xmlNumber(radius, 0));
  if (parsedRadius <= 0) return 0;
  const minSize = Math.max(1, Math.min(xmlNumber(width, 1), xmlNumber(height, 1)));
  const percent = Math.round((parsedRadius / minSize) * 100);
  return xmlClamp(percent, 1, 50);
};

const getDrawioTileStyle = (tile) => {
  const style = getTileStyle(tile);
  const rounded = xmlNumber(style.borderRadius, 0) > 0;
  const fontWeight = xmlNumber(style.fontWeight, 500);
  const fillColor = normalizeDrawioColor(style.background, '#0f172a');
  const borderColor = normalizeDrawioColor(style.border, '#2b3445');
  const textColor = normalizeDrawioColor(style.textColor, '#e5e7f0');
  const drawioStyle = {
    shape: 'rectangle',
    rounded: rounded ? 1 : 0,
    arcSize: rounded ? getDrawioArcSize(style.borderRadius, tile?.width, tile?.height) : null,
    whiteSpace: 'wrap',
    html: 1,
    fillColor,
    strokeColor: borderColor,
    strokeWidth: Math.max(0, xmlNumber(style.borderWidth, 1)),
    opacity: Math.round(xmlClamp(xmlNumber(style.opacity, 1), 0, 1) * 100),
    shadow: style.shadow && style.shadow !== 'none' ? 1 : 0,
    fontColor: textColor,
    fontFamily: style.fontFamily || 'Arial',
    fontSize: Math.max(8, Math.round(xmlNumber(style.fontSize, 13))),
    fontStyle: fontWeight >= 600 ? 1 : 0,
    align:
      style.textAlign === 'center' || style.textAlign === 'right' ? style.textAlign : 'left',
    verticalAlign: 'top',
    spacing: Math.max(0, Math.round(xmlNumber(style.padding, 16)))
  };
  return toDrawioStyleString(drawioStyle);
};

const getDrawioGraphModelFromDiagramNode = (diagramNode) => {
  if (!diagramNode) return null;
  const directModel = Array.from(diagramNode.childNodes).find(
    (child) => child.nodeType === Node.ELEMENT_NODE && child.nodeName === 'mxGraphModel'
  );
  if (directModel) return directModel;

  const raw = (diagramNode.textContent || '').trim();
  if (!raw) return null;
  // Defensive fallback: treat XML-looking payload as uncompressed draw.io data
  // regardless of compressed attribute state.
  if (!raw.startsWith('<')) return null;
  const nested = new DOMParser().parseFromString(raw, 'application/xml');
  if (nested.querySelector('parsererror')) return null;
  return nested.documentElement?.nodeName === 'mxGraphModel' ? nested.documentElement : null;
};

const getDrawioGraphModel = (doc) => {
  const root = doc.documentElement;
  if (!root) return { graphModel: null, diagramNode: null };
  if (root.nodeName === 'mxGraphModel') {
    return { graphModel: root, diagramNode: null };
  }
  if (root.nodeName === 'diagram') {
    return { graphModel: getDrawioGraphModelFromDiagramNode(root), diagramNode: root };
  }
  if (root.nodeName === 'mxfile') {
    const diagramNode = root.querySelector('diagram');
    return {
      graphModel: getDrawioGraphModelFromDiagramNode(diagramNode),
      diagramNode
    };
  }
  return { graphModel: null, diagramNode: null };
};

const decodeDrawioCellValue = (value) =>
  String(value || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .trim();

const buildFallbackTileFromDrawioCell = ({ cell, geometry, styleMap, index }) => {
  const width = Math.max(120, xmlNumber(geometry?.getAttribute('width'), 320));
  const height = Math.max(80, xmlNumber(geometry?.getAttribute('height'), 180));
  const x = Math.max(0, xmlNumber(geometry?.getAttribute('x'), 0));
  const y = Math.max(0, xmlNumber(geometry?.getAttribute('y'), 0));
  const rounded = styleMap.rounded === '1' || styleMap.rounded === 'true';
  const arcSize = xmlNumber(styleMap.arcSize, 12);
  const fallbackText = decodeDrawioCellValue(cell.getAttribute('value') || '');
  const lines = fallbackText.split('\n').map((line) => line.trim()).filter(Boolean);
  const title = lines[0] || 'Imported shape';
  const value = lines[1] || lines[0] || 'Value';
  const label = lines.slice(2).join(' ');
  const borderRadius = rounded ? Math.max(2, Math.round((Math.min(width, height) * arcSize) / 100)) : 0;

  return {
    id: `tile-import-${Date.now()}-${index}`,
    type: 'big-stat',
    x,
    y,
    width,
    height,
    zIndex: index + 1,
    style: {
      background: normalizeDrawioColor(styleMap.fillColor, '#0f172a'),
      border: normalizeDrawioColor(styleMap.strokeColor, '#2b3445'),
      borderWidth: Math.max(0, xmlNumber(styleMap.strokeWidth, 1)),
      borderRadius,
      opacity: xmlClamp(xmlNumber(styleMap.opacity, 100) / 100, 0, 1),
      padding: Math.max(0, Math.round(xmlNumber(styleMap.spacing, 16))),
      shadow: styleMap.shadow === '1' ? '0 12px 24px rgba(0, 0, 0, 0.35)' : 'none',
      textColor: normalizeDrawioColor(styleMap.fontColor, '#e5e7f0'),
      fontFamily: styleMap.fontFamily || 'Arial',
      fontSize: Math.max(8, Math.round(xmlNumber(styleMap.fontSize, 13))),
      fontWeight: (xmlNumber(styleMap.fontStyle, 0) & 1) === 1 ? 700 : 500,
      textAlign:
        styleMap.align === 'center' || styleMap.align === 'right' ? styleMap.align : 'left',
      lineHeight: 1.5,
      letterSpacing: 0
    },
    data: {
      header: title,
      value,
      label
    }
  };
};

const parseDrawioXmlToState = (doc) => {
  const { graphModel, diagramNode } = getDrawioGraphModel(doc);
  if (!graphModel) return null;

  const vertexCells = Array.from(graphModel.querySelectorAll('root > mxCell[vertex="1"]'));
  const taggedRootCells = vertexCells.filter((cell) => {
    const role = cell.getAttribute('dashboardRole');
    if (role === 'tile-root') return true;
    if (role === 'tile-part') return false;
    return (
      cell.hasAttribute('dashboardTile') ||
      cell.hasAttribute('tileType') ||
      cell.hasAttribute('tileId')
    );
  });
  const importCells = taggedRootCells.length
    ? taggedRootCells
    : vertexCells.filter((cell) => cell.getAttribute('dashboardRole') !== 'tile-part');
  if (!importCells.length) {
    const encodedWorkspace =
      diagramNode?.getAttribute('dashboardState') || graphModel.getAttribute('dashboardState') || '';
    if (encodedWorkspace) {
      const decoded = decodeBase64Utf8(encodedWorkspace);
      if (decoded) {
        try {
          const parsed = JSON.parse(decoded);
          if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
            return parsed;
          }
        } catch (err) {
          // ignore and fall back to default empty import
        }
      }
    }
  }

  const tiles = [];
  importCells.forEach((cell, index) => {
    const geometry = cell.querySelector('mxGeometry');
    if (!geometry) return;

    const styleMap = parseDrawioStyle(cell.getAttribute('style') || '');
    const encodedTile = cell.getAttribute('dashboardTile') || '';
    let tile = null;

    if (encodedTile) {
      const decodedTile = decodeBase64Utf8(encodedTile);
      if (decodedTile) {
        try {
          const parsedTile = JSON.parse(decodedTile);
          if (parsedTile && typeof parsedTile === 'object' && !Array.isArray(parsedTile)) {
            tile = parsedTile;
          }
        } catch (err) {
          tile = null;
        }
      }
    }

    if (!tile) {
      tile = buildFallbackTileFromDrawioCell({ cell, geometry, styleMap, index });
    }

    tile.id = String(cell.getAttribute('tileId') || tile.id || `tile-import-${index + 1}`);
    const rawType = String(cell.getAttribute('tileType') || tile.type || '').trim();
    tile.type = drawioKnownTileTypes.has(rawType) ? rawType : tile.type || 'big-stat';
    tile.x = Math.max(0, xmlNumber(geometry.getAttribute('x'), tile.x || 0));
    tile.y = Math.max(0, xmlNumber(geometry.getAttribute('y'), tile.y || 0));
    tile.width = Math.max(120, xmlNumber(geometry.getAttribute('width'), tile.width || 320));
    tile.height = Math.max(80, xmlNumber(geometry.getAttribute('height'), tile.height || 180));
    const parsedZ = xmlNumber(cell.getAttribute('tileZIndex'), Number.NaN);
    tile.zIndex = Number.isFinite(parsedZ) ? parsedZ : xmlNumber(tile.zIndex, index + 1);
    tile.style = tile.style && typeof tile.style === 'object' ? tile.style : {};
    tile.style.background = normalizeDrawioColor(
      styleMap.fillColor,
      tile.style.background || '#0f172a'
    );
    tile.style.border = normalizeDrawioColor(styleMap.strokeColor, tile.style.border || '#2b3445');
    tile.style.borderWidth = Math.max(0, xmlNumber(styleMap.strokeWidth, tile.style.borderWidth || 1));
    const styleOpacityFallback =
      typeof tile.style.opacity === 'number' ? tile.style.opacity * 100 : 100;
    tile.style.opacity = xmlClamp(
      xmlNumber(styleMap.opacity, styleOpacityFallback) / 100,
      0,
      1
    );
    tile.style.textColor = normalizeDrawioColor(styleMap.fontColor, tile.style.textColor || '#e5e7f0');
    tile.style.fontSize = Math.max(8, Math.round(xmlNumber(styleMap.fontSize, tile.style.fontSize || 13)));
    tile.style.padding = Math.max(0, Math.round(xmlNumber(styleMap.spacing, tile.style.padding || 16)));
    tile.style.textAlign =
      styleMap.align === 'center' || styleMap.align === 'right'
        ? styleMap.align
        : tile.style.textAlign || 'left';
    if (styleMap.rounded === '1' || styleMap.rounded === 'true') {
      const arcSize = xmlNumber(styleMap.arcSize, 12);
      tile.style.borderRadius = Math.max(2, Math.round((Math.min(tile.width, tile.height) * arcSize) / 100));
    } else if (styleMap.rounded === '0' || styleMap.rounded === 'false') {
      tile.style.borderRadius = 0;
    }

    tiles.push(tile);
  });

  const computedWidth =
    Math.max(
      xmlNumber(graphModel.getAttribute('pageWidth'), 0),
      ...tiles.map((tile) => tile.x + tile.width + 40),
      1600
    ) || 1600;
  const computedHeight =
    Math.max(
      xmlNumber(graphModel.getAttribute('pageHeight'), 0),
      ...tiles.map((tile) => tile.y + tile.height + 40),
      1000
    ) || 1000;
  const title = diagramNode?.getAttribute('name') || 'Imported Draw.io';
  return {
    project: {
      id: null,
      name: title
    },
    canvas: {
      width: Math.round(computedWidth),
      height: Math.round(computedHeight),
      background: '#0b0f14',
      gridOpacity: 0.08,
      gridSize: 12,
      zoom: 1
    },
    tiles,
    selectedTileId: tiles[0]?.id || null
  };
};

const serializeStateToXml = (payload) => {
  const doc = document.implementation.createDocument('', '', null);
  const root = doc.createElement('mxfile');
  root.setAttribute('host', 'app.diagrams.net');
  root.setAttribute('modified', new Date().toISOString());
  root.setAttribute('agent', 'Dashboard Studio V2');
  root.setAttribute('version', '25.0.0');
  root.setAttribute('type', 'device');
  root.setAttribute('compressed', 'false');

  const diagram = doc.createElement('diagram');
  diagram.setAttribute('id', 'dashboard-v2');
  diagram.setAttribute('name', sanitizeXmlString(payload?.project?.name || 'Dashboard'));
  diagram.setAttribute('compressed', 'false');
  diagram.setAttribute('dashboardState', encodeMetadataJson(payload || {}));

  const graphModel = doc.createElement('mxGraphModel');
  graphModel.setAttribute('dx', '1200');
  graphModel.setAttribute('dy', '800');
  graphModel.setAttribute('grid', '1');
  graphModel.setAttribute('gridSize', '10');
  graphModel.setAttribute('guides', '1');
  graphModel.setAttribute('tooltips', '1');
  graphModel.setAttribute('connect', '1');
  graphModel.setAttribute('arrows', '1');
  graphModel.setAttribute('fold', '1');
  graphModel.setAttribute('page', '1');
  graphModel.setAttribute('pageScale', '1');
  graphModel.setAttribute('pageWidth', String(Math.round(xmlNumber(payload?.canvas?.width, 1600))));
  graphModel.setAttribute('pageHeight', String(Math.round(xmlNumber(payload?.canvas?.height, 1000))));
  graphModel.setAttribute('math', '0');
  graphModel.setAttribute('shadow', '0');

  const modelRoot = doc.createElement('root');
  const cell0 = doc.createElement('mxCell');
  cell0.setAttribute('id', '0');
  modelRoot.appendChild(cell0);
  const cell1 = doc.createElement('mxCell');
  cell1.setAttribute('id', '1');
  cell1.setAttribute('parent', '0');
  modelRoot.appendChild(cell1);

  const orderedTiles = Array.isArray(payload?.tiles)
    ? payload.tiles
        .slice()
        .sort((left, right) => xmlNumber(left?.zIndex, 1) - xmlNumber(right?.zIndex, 1))
    : [];

  let nextCellId = 2;
  const allocCellId = () => String(nextCellId++);

  const appendVertexCell = ({
    x,
    y,
    width,
    height,
    value = '',
    style = '',
    attrs = {},
    parent = '1'
  }) => {
    const cell = doc.createElement('mxCell');
    const id = allocCellId();
    cell.setAttribute('id', id);
    cell.setAttribute('vertex', '1');
    cell.setAttribute('parent', parent);
    cell.setAttribute('value', sanitizeXmlString(value));
    cell.setAttribute('style', sanitizeXmlString(style));
    Object.entries(attrs).forEach(([key, rawValue]) => {
      if (rawValue === undefined || rawValue === null || rawValue === '') return;
      cell.setAttribute(sanitizeXmlString(key), sanitizeXmlString(rawValue));
    });

    const geometry = doc.createElement('mxGeometry');
    geometry.setAttribute('as', 'geometry');
    geometry.setAttribute('x', String(Math.round(Math.max(0, xmlNumber(x, 0)))));
    geometry.setAttribute('y', String(Math.round(Math.max(0, xmlNumber(y, 0)))));
    geometry.setAttribute('width', String(Math.round(Math.max(1, xmlNumber(width, 1)))));
    geometry.setAttribute('height', String(Math.round(Math.max(1, xmlNumber(height, 1)))));
    cell.appendChild(geometry);
    modelRoot.appendChild(cell);
    return id;
  };

  const appendKpiParts = (tile, baseStyle) => {
    const items = Array.isArray(tile?.data?.items) ? tile.data.items : [];
    if (!items.length) return;
    const padding = Math.max(8, Math.round(xmlNumber(baseStyle.padding, 16)));
    const gap = 12;
    const availableWidth = Math.max(1, xmlNumber(tile.width, 320) - padding * 2);
    const cardWidth = Math.max(80, (availableWidth - gap * (items.length - 1)) / items.length);
    const cardHeight = Math.max(48, xmlNumber(tile.height, 180) - padding * 2);
    const defaultFill = normalizeDrawioColor(baseStyle.background, '#101723');
    const defaultBorder = normalizeDrawioColor(baseStyle.border, '#1f2937');
    const defaultText = normalizeDrawioColor(baseStyle.textColor, '#e6e7eb');

    items.forEach((item, index) => {
      const x = xmlNumber(tile.x, 0) + padding + index * (cardWidth + gap);
      const y = xmlNumber(tile.y, 0) + padding;
      const fillColor = normalizeDrawioColor(item?.fill, defaultFill);
      const strokeColor = normalizeDrawioColor(item?.border, defaultBorder);
      const fontColor = normalizeDrawioColor(
        item?.textColor || item?.valueColor || item?.titleColor,
        defaultText
      );
      const fontSize = Math.max(
        9,
        Math.round(xmlNumber(item?.valueSize, xmlNumber(baseStyle.fontSize, 13)))
      );
      const borderRadius = Math.max(0, xmlNumber(item?.radius, 10));
      const value = [item?.title, item?.value, item?.delta]
        .map((line) => (typeof line === 'string' ? line.trim() : ''))
        .filter(Boolean)
        .join('\n');
      appendVertexCell({
        x,
        y,
        width: cardWidth,
        height: cardHeight,
        value,
        style: getDrawioPartStyle({
          fillColor,
          strokeColor,
          strokeWidth: xmlNumber(item?.borderWidth, 1),
          borderRadius,
          width: cardWidth,
          height: cardHeight,
          fontColor,
          fontFamily: baseStyle.fontFamily,
          fontSize,
          bold: xmlNumber(item?.valueWeight, 600) >= 600,
          align: 'left',
          verticalAlign: 'top',
          spacing: 8
        }),
        attrs: {
          dashboardRole: 'tile-part',
          dashboardPartType: 'kpi-card',
          dashboardParentTileId: tile?.id || ''
        }
      });
    });
  };

  const appendDemographicParts = (tile, baseStyle) => {
    const items = Array.isArray(tile?.data?.items) ? tile.data.items : [];
    if (!items.length) return;
    const padding = Math.max(8, Math.round(xmlNumber(baseStyle.padding, 16)));
    const title = typeof tile?.data?.title === 'string' ? tile.data.title.trim() : '';
    const titleHeight = title ? Math.max(24, Math.round(xmlNumber(baseStyle.fontSize, 13) + 14)) : 0;
    const columns = Math.max(1, Math.min(5, items.length));
    const rows = Math.max(1, Math.ceil(items.length / columns));
    const gap = 12;
    const availableWidth = Math.max(1, xmlNumber(tile.width, 320) - padding * 2);
    const cardWidth = Math.max(70, (availableWidth - gap * (columns - 1)) / columns);
    const availableHeight = Math.max(44, xmlNumber(tile.height, 180) - padding * 2 - titleHeight);
    const cardHeight = Math.max(34, (availableHeight - gap * (rows - 1)) / rows);
    const startY = xmlNumber(tile.y, 0) + padding + titleHeight;
    const fillColor = normalizeDrawioColor(baseStyle.background, '#101723');
    const strokeColor = normalizeDrawioColor(baseStyle.border, '#1f2937');
    const fontColor = normalizeDrawioColor(baseStyle.textColor, '#e6e7eb');
    const borderRadius = Math.max(0, xmlNumber(baseStyle.borderRadius, 10));

    items.forEach((item, index) => {
      const col = index % columns;
      const row = Math.floor(index / columns);
      const x = xmlNumber(tile.x, 0) + padding + col * (cardWidth + gap);
      const y = startY + row * (cardHeight + gap);
      const value = [item?.label, item?.value]
        .map((line) => (typeof line === 'string' ? line.trim() : ''))
        .filter(Boolean)
        .join('\n');
      appendVertexCell({
        x,
        y,
        width: cardWidth,
        height: cardHeight,
        value,
        style: getDrawioPartStyle({
          fillColor,
          strokeColor,
          strokeWidth: Math.max(0, xmlNumber(baseStyle.borderWidth, 1)),
          borderRadius,
          width: cardWidth,
          height: cardHeight,
          fontColor,
          fontFamily: baseStyle.fontFamily,
          fontSize: Math.max(9, Math.round(xmlNumber(baseStyle.fontSize, 13))),
          bold: true,
          align: 'left',
          verticalAlign: 'middle',
          spacing: 8
        }),
        attrs: {
          dashboardRole: 'tile-part',
          dashboardPartType: 'demographic-card',
          dashboardParentTileId: tile?.id || ''
        }
      });
    });
  };

  const parseTableColumnWeight = (widthValue) => {
    const text = String(widthValue || '1fr').trim();
    const fr = text.match(/^([0-9]*\.?[0-9]+)\s*fr$/i);
    if (fr) {
      const parsed = Number(fr[1]);
      return Number.isFinite(parsed) && parsed > 0 ? parsed : 1;
    }
    const numeric = Number(text);
    return Number.isFinite(numeric) && numeric > 0 ? numeric : 1;
  };

  const appendTableParts = (tile, baseStyle) => {
    const data = tile?.data || {};
    const rawColumns = Array.isArray(data.columns) ? data.columns : [];
    const columns = rawColumns.length
      ? rawColumns.map((column) =>
          column && typeof column === 'object'
            ? column
            : { label: String(column ?? ''), align: 'left', width: '1fr' }
        )
      : [{ label: 'Value', align: 'left', width: '1fr' }];
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const settings = data.settings && typeof data.settings === 'object' ? data.settings : {};
    const showHeader = settings.showHeader !== false;
    const striped = settings.striped !== false;
    const padding = Math.max(8, Math.round(xmlNumber(baseStyle.padding, 16)));
    const title = typeof data.title === 'string' ? data.title.trim() : '';
    const titleHeight = title ? Math.max(24, Math.round(xmlNumber(baseStyle.fontSize, 13) + 16)) : 0;
    const rowHeight = xmlClamp(Math.round(xmlNumber(settings.rowHeight, 36)), 20, 90);
    const tableX = xmlNumber(tile.x, 0) + padding;
    const tableY = xmlNumber(tile.y, 0) + padding + titleHeight;
    const tableWidth = Math.max(60, xmlNumber(tile.width, 320) - padding * 2);
    const tableHeight = Math.max(40, xmlNumber(tile.height, 180) - padding * 2 - titleHeight);
    const headerRows = showHeader ? 1 : 0;
    const capacity = Math.max(1, Math.floor(tableHeight / rowHeight) - headerRows);
    const renderRows = rows.slice(0, capacity);
    const weights = columns.map((column) => parseTableColumnWeight(column.width));
    const totalWeight = weights.reduce((sum, value) => sum + value, 0) || 1;

    const xOffsets = [];
    let cursor = tableX;
    weights.forEach((weight, index) => {
      xOffsets.push(cursor);
      const remaining = tableX + tableWidth - cursor;
      if (index === weights.length - 1) {
        cursor += remaining;
      } else {
        cursor += (tableWidth * weight) / totalWeight;
      }
    });

    const drawRowCells = (cells, rowIndex, options = {}) => {
      const y = tableY + rowIndex * rowHeight;
      columns.forEach((column, colIndex) => {
        const x = xOffsets[colIndex];
        const nextX =
          colIndex === columns.length - 1 ? tableX + tableWidth : xOffsets[colIndex + 1];
        const width = Math.max(24, nextX - x);
        const raw = cells[colIndex];
        const value = String(raw ?? '').trim();
        const align =
          column.align === 'center' || column.align === 'right' ? column.align : 'left';
        const baseFill = options.fillColor || normalizeDrawioColor(baseStyle.background, '#0f172a');
        appendVertexCell({
          x,
          y,
          width,
          height: rowHeight,
          value,
          style: getDrawioPartStyle({
            fillColor: baseFill,
            strokeColor: normalizeDrawioColor(baseStyle.border, '#2b3445'),
            strokeWidth: Math.max(0, xmlNumber(baseStyle.borderWidth, 1)),
            borderRadius: 0,
            width,
            height: rowHeight,
            fontColor: normalizeDrawioColor(baseStyle.textColor, '#e5e7f0'),
            fontFamily: baseStyle.fontFamily,
            fontSize: Math.max(8, Math.round(xmlNumber(baseStyle.fontSize, 12))),
            bold: Boolean(options.bold),
            align,
            verticalAlign: 'middle',
            spacing: 6
          }),
          attrs: {
            dashboardRole: 'tile-part',
            dashboardPartType: options.partType || 'table-cell',
            dashboardParentTileId: tile?.id || ''
          }
        });
      });
    };

    let rowCursor = 0;
    if (showHeader) {
      drawRowCells(
        columns.map((column) => String(column.label || '').trim()),
        rowCursor,
        {
          bold: true,
          partType: 'table-header-cell',
          fillColor: normalizeDrawioColor(baseStyle.border, '#1f2937')
        }
      );
      rowCursor += 1;
    }

    renderRows.forEach((row, rowIndex) => {
      const normalizedRow = Array.isArray(row) ? row : [];
      const fillColor =
        striped && rowIndex % 2 === 1
          ? normalizeDrawioColor(baseStyle.background, '#0f172a')
          : normalizeDrawioColor(baseStyle.background, '#111827');
      drawRowCells(normalizedRow, rowCursor + rowIndex, {
        partType: 'table-cell',
        fillColor
      });
    });
  };

  orderedTiles.forEach((tile, index) => {
    if (!tile || typeof tile !== 'object') return;
    const rootId = appendVertexCell({
      x: xmlNumber(tile?.x, 0),
      y: xmlNumber(tile?.y, 0),
      width: Math.max(40, xmlNumber(tile?.width, 320)),
      height: Math.max(40, xmlNumber(tile?.height, 180)),
      value: getTileLabelForDrawio(tile),
      style: getDrawioTileStyle(tile),
      attrs: {
        tileId: String(tile?.id || `tile-${index + 1}`),
        tileType: String(tile?.type || 'big-stat'),
        tileZIndex: String(xmlNumber(tile?.zIndex, index + 1)),
        dashboardTile: encodeMetadataJson(tile || {}),
        dashboardRole: 'tile-root'
      }
    });

    const baseStyle = getTileStyle(tile);
    if (tile.type === 'kpi-row') {
      appendKpiParts(tile, baseStyle);
    } else if (tile.type === 'demographics') {
      appendDemographicParts(tile, baseStyle);
    } else if (tile.type === 'table') {
      appendTableParts(tile, baseStyle);
    }

    void rootId;
  });

  graphModel.appendChild(modelRoot);
  diagram.appendChild(graphModel);
  root.appendChild(diagram);
  doc.appendChild(root);
  const serializer = new XMLSerializer();
  return `<?xml version="1.0" encoding="UTF-8"?>\n${serializer.serializeToString(doc)}`;
};

const parseXmlToState = (xmlText) => {
  const source = (xmlText || '').trim();
  if (!source) {
    throw new Error('Paste XML into the field before importing.');
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(source, 'application/xml');
  const parserError = doc.querySelector('parsererror');
  if (parserError) {
    throw new Error('Invalid XML. Please fix the XML and try again.');
  }

  const root = doc.documentElement;
  if (!root) {
    throw new Error('Invalid XML. Root element is missing.');
  }

  const drawioState = parseDrawioXmlToState(doc);
  if (drawioState) return drawioState;

  let valueNode = null;
  if (root.nodeName === 'dashboard-studio' || root.nodeName === 'dashboardStudio') {
    const stateNode = Array.from(root.children).find((child) => child.nodeName === 'state');
    valueNode = stateNode ? getFirstElementChild(stateNode) : getFirstElementChild(root);
  } else if (root.nodeName === 'state' || root.nodeName === 'value') {
    valueNode = root.nodeName === 'state' ? getFirstElementChild(root) : root;
  } else {
    valueNode = root;
  }

  const decoded = decodeXmlValue(valueNode || root);
  if (!decoded || typeof decoded !== 'object' || Array.isArray(decoded)) {
    throw new Error('Imported XML must decode to a workspace object.');
  }
  return decoded;
};

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
  projectCache = Array.isArray(projects) ? projects.slice() : [];
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
const MAX_HIGHLIGHT_ITEMS = 20;
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

const getUniqueProjectCopyName = (baseName) => {
  const normalizedBase = (baseName || DEFAULT_PROJECT_NAME).trim() || DEFAULT_PROJECT_NAME;
  const copyBase = normalizedBase.toLowerCase().endsWith(' copy') ? normalizedBase : `${normalizedBase} Copy`;
  const existingNames = new Set(
    projectCache
      .map((project) => (project?.name || '').trim().toLowerCase())
      .filter(Boolean)
  );
  let candidate = copyBase;
  let suffix = 2;
  while (existingNames.has(candidate.toLowerCase())) {
    candidate = `${copyBase} ${suffix}`;
    suffix += 1;
  }
  return candidate;
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

const createDefaultHighlightLine = (index = 0) => `New highlight ${index + 1}`;

const DEFAULT_TILES = {
  'kpi-row': {
    size: { w: 880, h: 140 },
    data: {
      icon: '📊',
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
    data: { header: 'Big Stat', label: 'Total Blogs Released', value: '13', icon: '✨' }
  },
  highlights: {
    size: { w: 520, h: 260 },
    data: {
      icon: '🔥',
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
      icon: '📰',
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
      icon: '📈',
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
      icon: '🌍',
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
      icon: '🧾',
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
  highlightColor: tile.style?.highlightColor ?? state.theme.accent,
  highlightIdle: tile.style?.highlightIdle !== false,
  textColor: tile.style?.textColor ?? state.theme.textColor,
  fontFamily: tile.style?.fontFamily ?? state.theme.fontFamily,
  fontSize: tile.style?.fontSize ?? 13,
  fontWeight: tile.style?.fontWeight ?? 500,
  textAlign: tile.style?.textAlign ?? 'left',
  lineHeight: tile.style?.lineHeight ?? 1.5,
  letterSpacing: tile.style?.letterSpacing ?? 0
});

const clampImageOpacity = (value) => {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) return 0.2;
  return Math.min(1, Math.max(0, parsed));
};

const toHighlightRgba = (color, alpha = 0.28) => {
  const hex = normalizeHex(color);
  if (!hex) return `rgba(91, 130, 246, ${alpha})`;
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
};

const getTileGraphics = (tile) => {
  const icon = typeof tile?.data?.icon === 'string' ? tile.data.icon.trim() : '';
  const imageUrl = typeof tile?.data?.imageUrl === 'string' ? tile.data.imageUrl.trim() : '';
  const imageFit = tile?.data?.imageFit === 'contain' ? 'contain' : 'cover';
  const imageOpacity = clampImageOpacity(tile?.data?.imageOpacity);
  return { icon, imageUrl, imageFit, imageOpacity };
};

const renderTileGraphics = (tile) => {
  const graphics = getTileGraphics(tile);
  const iconHtml = graphics.icon ? `<div class="tile-media-icon editable" contenteditable data-field="icon">${graphics.icon}</div>` : '';
  const imageHtml = graphics.imageUrl
    ? `<div class="tile-media-image"><img src="${graphics.imageUrl}" alt="" loading="lazy" style="object-fit:${graphics.imageFit};opacity:${graphics.imageOpacity};" /></div>`
    : '';
  return `${iconHtml}${imageHtml}`;
};

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

const getTileBounds = (tiles = []) => {
  const validTiles = tiles.filter((tile) => tile && typeof tile === 'object');
  if (!validTiles.length) return null;
  return validTiles.reduce(
    (bounds, tile) => ({
      minX: Math.min(bounds.minX, Number(tile.x) || 0),
      minY: Math.min(bounds.minY, Number(tile.y) || 0),
      maxX: Math.max(bounds.maxX, (Number(tile.x) || 0) + (Number(tile.width) || MIN_TILE_WIDTH)),
      maxY: Math.max(bounds.maxY, (Number(tile.y) || 0) + (Number(tile.height) || MIN_TILE_HEIGHT))
    }),
    {
      minX: Number.POSITIVE_INFINITY,
      minY: Number.POSITIVE_INFINITY,
      maxX: Number.NEGATIVE_INFINITY,
      maxY: Number.NEGATIVE_INFINITY
    }
  );
};

const cloneTilesForMerge = (tiles = []) =>
  tiles
    .filter((tile) => tile && typeof tile === 'object' && typeof tile.type === 'string')
    .map((tile, index) => {
      const clone = deepClone(tile);
      clone.id = generateId();
      clone.x = Number.isFinite(Number(clone.x)) ? Number(clone.x) : 80 + index * 24;
      clone.y = Number.isFinite(Number(clone.y)) ? Number(clone.y) : 80 + index * 24;
      clone.width = Number.isFinite(Number(clone.width)) ? Number(clone.width) : MIN_TILE_WIDTH;
      clone.height = Number.isFinite(Number(clone.height)) ? Number(clone.height) : MIN_TILE_HEIGHT;
      if (!clone.data || typeof clone.data !== 'object' || Array.isArray(clone.data)) {
        clone.data = {};
      }
      if (!clone.style || typeof clone.style !== 'object' || Array.isArray(clone.style)) {
        clone.style = {};
      }
      return clone;
    });

const hasOwn = (value, key) => Object.prototype.hasOwnProperty.call(value || {}, key);

const assignOwnFields = (target, source, keys = []) => {
  keys.forEach((key) => {
    if (hasOwn(source, key)) {
      target[key] = deepClone(source[key]);
    }
  });
};

const mergeGraphicFields = (target, source) => {
  assignOwnFields(target, source, ['icon', 'imageUrl', 'imageFit', 'imageOpacity']);
};

const normalizeTableColumn = (column, fallback = {}) => {
  const base = column && typeof column === 'object' && !Array.isArray(column) ? column : {};
  return {
    label: String(base.label ?? (typeof column === 'string' ? column : fallback.label ?? '')),
    align: base.align || fallback.align || 'left',
    width: base.width || fallback.width || '1fr'
  };
};

const mergeTileDataContent = (currentTile, importedTile) => {
  const currentData =
    currentTile?.data && typeof currentTile.data === 'object' && !Array.isArray(currentTile.data)
      ? currentTile.data
      : {};
  const incomingData =
    importedTile?.data && typeof importedTile.data === 'object' && !Array.isArray(importedTile.data)
      ? importedTile.data
      : {};
  const nextData = deepClone(currentData);
  mergeGraphicFields(nextData, incomingData);

  switch (currentTile.type) {
    case 'kpi-row': {
      if (!hasOwn(incomingData, 'items')) break;
      const currentItems = Array.isArray(currentData.items) ? currentData.items : [];
      const fallbackTemplate = deepClone(
        currentItems[currentItems.length - 1] || DEFAULT_TILES['kpi-row']?.data?.items?.[0] || {}
      );
      const incomingItems = Array.isArray(incomingData.items) ? incomingData.items : [];
      nextData.items = incomingItems.map((incomingItem, index) => {
        const base = deepClone(currentItems[index] || currentItems[currentItems.length - 1] || fallbackTemplate);
        const safeItem =
          incomingItem && typeof incomingItem === 'object' && !Array.isArray(incomingItem) ? incomingItem : {};
        if (hasOwn(safeItem, 'title')) base.title = String(safeItem.title ?? '');
        if (hasOwn(safeItem, 'value')) base.value = String(safeItem.value ?? '');
        if (hasOwn(safeItem, 'delta')) base.delta = String(safeItem.delta ?? '');
        return base;
      });
      break;
    }
    case 'big-stat': {
      assignOwnFields(nextData, incomingData, ['header', 'label', 'value']);
      break;
    }
    case 'highlights': {
      assignOwnFields(nextData, incomingData, ['title']);
      if (hasOwn(incomingData, 'items')) {
        nextData.items = Array.isArray(incomingData.items)
          ? incomingData.items.map((item) => String(item ?? ''))
          : [];
      }
      break;
    }
    case 'blogs': {
      assignOwnFields(nextData, incomingData, ['title', 'subtitle']);
      if (hasOwn(incomingData, 'items')) {
        const currentItems = Array.isArray(currentData.items) ? currentData.items : [];
        const incomingItems = Array.isArray(incomingData.items) ? incomingData.items : [];
        nextData.items = incomingItems.map((incomingItem, index) => {
          const base = deepClone(currentItems[index] || {});
          const safeItem =
            incomingItem && typeof incomingItem === 'object' && !Array.isArray(incomingItem) ? incomingItem : {};
          if (hasOwn(safeItem, 'title')) base.title = String(safeItem.title ?? '');
          if (hasOwn(safeItem, 'views')) base.views = String(safeItem.views ?? '');
          return base;
        });
      }
      break;
    }
    case 'graph': {
      assignOwnFields(nextData, incomingData, ['title', 'subtitle', 'points', 'xAxis']);
      if (hasOwn(incomingData, 'yAxis')) {
        nextData.yAxis = {
          ...(currentData.yAxis || {}),
          ...deepClone(incomingData.yAxis || {})
        };
      }
      if (hasOwn(currentData, 'style')) {
        nextData.style = deepClone(currentData.style);
      }
      break;
    }
    case 'demographics': {
      assignOwnFields(nextData, incomingData, ['title']);
      if (hasOwn(incomingData, 'items')) {
        const currentItems = Array.isArray(currentData.items) ? currentData.items : [];
        const incomingItems = Array.isArray(incomingData.items) ? incomingData.items : [];
        nextData.items = incomingItems.map((incomingItem, index) => {
          const base = deepClone(currentItems[index] || {});
          const safeItem =
            incomingItem && typeof incomingItem === 'object' && !Array.isArray(incomingItem) ? incomingItem : {};
          if (hasOwn(safeItem, 'label')) base.label = String(safeItem.label ?? '');
          if (hasOwn(safeItem, 'value')) base.value = String(safeItem.value ?? '');
          return base;
        });
      }
      break;
    }
    case 'table': {
      assignOwnFields(nextData, incomingData, ['title']);
      if (hasOwn(incomingData, 'columns')) {
        const currentColumns = Array.isArray(currentData.columns) ? currentData.columns.map((column) => normalizeTableColumn(column)) : [];
        const fallbackColumn = currentColumns[currentColumns.length - 1] || normalizeTableColumn({});
        const incomingColumns = Array.isArray(incomingData.columns) ? incomingData.columns : [];
        nextData.columns = incomingColumns.map((column, index) => {
          const base = normalizeTableColumn(currentColumns[index], fallbackColumn);
          if (column && typeof column === 'object' && !Array.isArray(column)) {
            return {
              ...base,
              label: hasOwn(column, 'label') ? String(column.label ?? '') : base.label
            };
          }
          return {
            ...base,
            label: String(column ?? '')
          };
        });
      }
      if (hasOwn(incomingData, 'rows')) {
        nextData.rows = Array.isArray(incomingData.rows) ? deepClone(incomingData.rows) : [];
      }
      if (hasOwn(currentData, 'settings')) {
        nextData.settings = deepClone(currentData.settings);
      }
      break;
    }
    default: {
      Object.entries(incomingData).forEach(([key, value]) => {
        if (key === 'style') return;
        nextData[key] = deepClone(value);
      });
      if (hasOwn(currentData, 'style')) {
        nextData.style = deepClone(currentData.style);
      }
      break;
    }
  }

  return nextData;
};

const findMatchingTileForDataReplace = (importedTile, currentTiles, usedTileIds, typeQueues) => {
  if (!importedTile || typeof importedTile !== 'object') return null;
  if (importedTile.id) {
    const exactMatch = currentTiles.find(
      (tile) =>
        tile.id === importedTile.id &&
        !usedTileIds.has(tile.id) &&
        (!importedTile.type || tile.type === importedTile.type)
    );
    if (exactMatch) return exactMatch;
  }

  if (!importedTile.type) return null;
  const queue = typeQueues.get(importedTile.type) || [];
  while (queue.length) {
    const candidate = queue.shift();
    if (!usedTileIds.has(candidate.id)) return candidate;
  }
  return null;
};

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
    tileStyleHighlightInput,
    tileStyleHighlightIdleInput,
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
    tileGraphicIconInput,
    tileGraphicImageUrlInput,
    tileGraphicImageFileInput,
    tileGraphicImageFitInput,
    tileGraphicImageOpacityInput,
    tileGraphicImageClearBtn,
    kpiAddBtn,
    kpiRemoveBtn,
    demoAddBtn,
    demoRemoveBtn,
    highlightsAddBtn,
    highlightsRemoveBtn,
    highlightsLineIndexInput,
    highlightsLineTextInput,
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
    selectionLabel.textContent = tile ? tile.type.replace(/-/g, ' ') : 'Select a tile';
  }
  if (bigStatSection) bigStatSection.classList.toggle('hidden', tile?.type !== 'big-stat');
  if (highlightsSection) highlightsSection.classList.toggle('hidden', tile?.type !== 'highlights');
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
  if (tileStyleHighlightInput) tileStyleHighlightInput.value = style.highlightColor;
  if (tileStyleHighlightIdleInput) tileStyleHighlightIdleInput.checked = style.highlightIdle;

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

  const graphics = getTileGraphics(tile);
  if (tileGraphicIconInput && document.activeElement !== tileGraphicIconInput) {
    tileGraphicIconInput.value = graphics.icon;
  }
  if (tileGraphicImageUrlInput && document.activeElement !== tileGraphicImageUrlInput) {
    tileGraphicImageUrlInput.value = graphics.imageUrl;
  }
  if (tileGraphicImageFitInput) {
    tileGraphicImageFitInput.value = graphics.imageFit;
  }
  if (tileGraphicImageOpacityInput && document.activeElement !== tileGraphicImageOpacityInput) {
    tileGraphicImageOpacityInput.value = String(Number(graphics.imageOpacity.toFixed(2)));
  }

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
  if (tile.type === 'highlights') {
    const lines = Array.isArray(tile.data?.items) ? tile.data.items : [];
    if (highlightsAddBtn) highlightsAddBtn.disabled = lines.length >= MAX_HIGHLIGHT_ITEMS;
    if (highlightsRemoveBtn) highlightsRemoveBtn.disabled = lines.length <= 1;
    if (highlightsLineIndexInput) {
      const max = Math.max(1, lines.length);
      highlightsLineIndexInput.max = String(max);
      const current = Number(highlightsLineIndexInput.value || 1);
      const clamped = Math.min(max, Math.max(1, Number.isFinite(current) ? current : 1));
      if (current !== clamped) highlightsLineIndexInput.value = String(clamped);
    }
    const index = Math.max(
      0,
      Math.min(lines.length - 1, Number(highlightsLineIndexInput?.value || 1) - 1)
    );
    if (highlightsLineTextInput && document.activeElement !== highlightsLineTextInput) {
      highlightsLineTextInput.value = lines[index] || '';
    }
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
        ${renderTileGraphics(tile)}
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
          ${renderTileGraphics(tile)}
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
        ${renderTileGraphics(tile)}
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
        ${renderTileGraphics(tile)}
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
        ${renderTileGraphics(tile)}
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
        ${renderTileGraphics(tile)}
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
          ${renderTileGraphics(tile)}
          <div class="tile-header editable" contenteditable data-field="title">${tile.data.title}</div>
          <div class="table-grid" style="--table-columns:${templateColumns}; --row-height:${rowHeight}px;">
            ${
              showHeader
                ? `<div class="table-row table-header">
                    ${columns
                      .map((col, idx) => {
                        const label = typeof col === 'object' ? col.label : col;
                        const align = typeof col === 'object' ? col.align || 'left' : 'left';
                        const wrapClass = align === 'left' && idx === 0 ? 'wrap' : 'nowrap';
                        return `<div class="table-cell ${wrapClass} align-${align} editable" contenteditable data-field="columns.${idx}.label">${label}</div>`;
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
                    const wrapClass = align === 'left' && cellIdx === 0 ? 'wrap' : 'nowrap';
                    return `<div class="table-cell ${wrapClass} align-${align} editable" contenteditable data-field="rows.${rowIdx}.${cellIdx}">${cell}</div>`;
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
      return `${renderTileGraphics(tile)}<div class="tile-header">${tile.type}</div>`;
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
    const highlightRing = style.highlightIdle
      ? `0 0 0 1px ${toHighlightRgba(style.highlightColor, 0.28)}`
      : 'none';
    if (style.shadow && style.shadow !== 'none') {
      node.style.boxShadow = style.highlightIdle ? `${style.shadow}, ${highlightRing}` : style.shadow;
    } else {
      node.style.boxShadow = style.highlightIdle ? highlightRing : 'none';
    }
    node.style.setProperty('--tile-highlight', style.highlightColor);
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

  const updateTileGraphics = (patch) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile) return;
    tile.data = tile.data || {};
    Object.assign(tile.data, patch);
    renderTiles();
    scheduleHistory('Graphics');
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

  const updateHighlightsItems = (mutate) => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'highlights') return;
    tile.data = tile.data || {};
    tile.data.items = Array.isArray(tile.data.items) ? tile.data.items : [];
    mutate(tile.data.items);
    renderTiles();
    applyInspectorValues();
    scheduleHistory('Highlights');
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
  tileStyleHighlightInput?.addEventListener('input', () =>
    updateTileStyle({ highlightColor: tileStyleHighlightInput.value })
  );
  tileStyleHighlightIdleInput?.addEventListener('change', () =>
    updateTileStyle({ highlightIdle: Boolean(tileStyleHighlightIdleInput.checked) })
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

  tileGraphicIconInput?.addEventListener('input', () =>
    updateTileGraphics({ icon: tileGraphicIconInput.value })
  );
  tileGraphicImageUrlInput?.addEventListener('input', () =>
    updateTileGraphics({ imageUrl: tileGraphicImageUrlInput.value.trim() })
  );
  tileGraphicImageFitInput?.addEventListener('change', () =>
    updateTileGraphics({ imageFit: tileGraphicImageFitInput.value === 'contain' ? 'contain' : 'cover' })
  );
  tileGraphicImageOpacityInput?.addEventListener('input', () =>
    updateTileGraphics({ imageOpacity: clampImageOpacity(tileGraphicImageOpacityInput.value) })
  );
  tileGraphicImageClearBtn?.addEventListener('click', () => {
    if (tileGraphicImageUrlInput) tileGraphicImageUrlInput.value = '';
    if (tileGraphicImageFileInput) tileGraphicImageFileInput.value = '';
    updateTileGraphics({ imageUrl: '' });
  });
  tileGraphicImageFileInput?.addEventListener('change', () => {
    const file = tileGraphicImageFileInput.files && tileGraphicImageFileInput.files[0];
    if (!file) return;
    if (!file.type || !file.type.startsWith('image/')) {
      alert('Please choose an image file.');
      tileGraphicImageFileInput.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = typeof reader.result === 'string' ? reader.result : '';
      if (!dataUrl) return;
      if (tileGraphicImageUrlInput) tileGraphicImageUrlInput.value = dataUrl;
      updateTileGraphics({ imageUrl: dataUrl });
      tileGraphicImageFileInput.value = '';
    };
    reader.onerror = () => {
      alert('Could not read that file. Try another image.');
      tileGraphicImageFileInput.value = '';
    };
    reader.readAsDataURL(file);
  });

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

  highlightsAddBtn?.addEventListener('click', () => {
    updateHighlightsItems((items) => {
      if (items.length >= MAX_HIGHLIGHT_ITEMS) return;
      items.push(createDefaultHighlightLine(items.length));
    });
  });

  highlightsRemoveBtn?.addEventListener('click', () => {
    updateHighlightsItems((items) => {
      if (items.length <= 1) return;
      const rawIndex = Number(highlightsLineIndexInput?.value || items.length);
      const index = Math.max(0, Math.min(items.length - 1, Number.isFinite(rawIndex) ? rawIndex - 1 : items.length - 1));
      items.splice(index, 1);
    });
  });

  highlightsLineIndexInput?.addEventListener('input', () => applyInspectorValues());
  highlightsLineTextInput?.addEventListener('input', () => {
    const tile = state.tiles.find((t) => t.id === state.selectedTileId);
    if (!tile || tile.type !== 'highlights') return;
    tile.data = tile.data || {};
    tile.data.items = Array.isArray(tile.data.items) ? tile.data.items : [];
    if (!tile.data.items.length) return;
    const rawIndex = Number(highlightsLineIndexInput?.value || 1);
    const index = Math.max(0, Math.min(tile.data.items.length - 1, Number.isFinite(rawIndex) ? rawIndex - 1 : 0));
    tile.data.items[index] = highlightsLineTextInput.value;
    renderTiles();
    scheduleHistory('Highlights');
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

const applyImportedWorkspace = (next, historyLabel = 'Import') => {
  if (!next || typeof next !== 'object' || Array.isArray(next)) {
    throw new Error('Imported payload must be an object.');
  }
  replaceState(next);
  if (state.project?.id) {
    localStorage.setItem(STORAGE_KEYS.projectId, String(state.project.id));
  } else {
    localStorage.removeItem(STORAGE_KEYS.projectId);
  }
  localStorage.setItem(STORAGE_KEYS.gridSize, String(getGridSize()));
  render();
  updateProjectNameDisplay();
  history = [];
  historyIndex = -1;
  pushHistory(historyLabel);
  updateUndoRedoButtons();
  updateHistoryMenu();
};

const replaceImportedTileData = (next, historyLabel = 'Replace Data') => {
  if (!next || typeof next !== 'object' || Array.isArray(next)) {
    throw new Error('Imported payload must be an object.');
  }
  const importedTiles = Array.isArray(next.tiles) ? next.tiles : [];
  if (!importedTiles.length) {
    throw new Error('Replace Data requires a JSON object with a tiles array.');
  }

  const typeQueues = new Map();
  state.tiles.forEach((tile) => {
    if (!typeQueues.has(tile.type)) {
      typeQueues.set(tile.type, []);
    }
    typeQueues.get(tile.type).push(tile);
  });

  const usedTileIds = new Set();
  let matchedCount = 0;

  importedTiles.forEach((importedTile) => {
    const targetTile = findMatchingTileForDataReplace(importedTile, state.tiles, usedTileIds, typeQueues);
    if (!targetTile) return;
    targetTile.data = mergeTileDataContent(targetTile, importedTile);
    usedTileIds.add(targetTile.id);
    matchedCount += 1;
  });

  if (!matchedCount) {
    throw new Error('No imported tiles matched the current project. Use existing tile ids or matching tile types.');
  }

  render();
  updateProjectNameDisplay();
  pushHistory(historyLabel);

  if (matchedCount < importedTiles.length) {
    console.warn(`Replace Data skipped ${importedTiles.length - matchedCount} unmatched tile(s).`);
  }
};

const mergeImportedWorkspace = (next, historyLabel = 'Merge Import') => {
  if (!next || typeof next !== 'object' || Array.isArray(next)) {
    throw new Error('Imported payload must be an object.');
  }
  const imported = normalizeState(next);
  const incomingTiles = cloneTilesForMerge(imported.tiles);
  if (!incomingTiles.length) {
    throw new Error('Imported workspace does not contain any tiles to add.');
  }

  const currentBounds = getTileBounds(state.tiles);
  const incomingBounds = getTileBounds(incomingTiles);
  if (currentBounds && incomingBounds) {
    const gap = Math.max(32, getGridSize() * 3);
    const offsetX = currentBounds.maxX - incomingBounds.minX + gap;
    const offsetY = currentBounds.minY - incomingBounds.minY;
    incomingTiles.forEach((tile) => {
      tile.x += offsetX;
      tile.y += offsetY;
    });
  }

  state.tiles = state.tiles.concat(incomingTiles);
  state.selectedTileId = incomingTiles[incomingTiles.length - 1].id;
  render();
  updateProjectNameDisplay();
  pushHistory(historyLabel);
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
        applyImportedWorkspace(next, 'Import JSON');
      } catch (err) {
        console.error('Invalid JSON', err);
        alert('Invalid JSON. Please fix the JSON and try again.');
      }
    });
  }
  if (jsonImportDataBtn) {
    jsonImportDataBtn.addEventListener('click', () => {
      if (!jsonInput || !jsonInput.value.trim()) {
        alert('Paste JSON into the field before importing.');
        return;
      }
      try {
        const next = JSON.parse(jsonInput.value);
        replaceImportedTileData(next, 'Replace Data');
      } catch (err) {
        console.error('Invalid JSON', err);
        alert(err?.message || 'Invalid JSON. Please fix the JSON and try again.');
      }
    });
  }
  if (jsonImportMergeBtn) {
    jsonImportMergeBtn.addEventListener('click', () => {
      if (!jsonInput || !jsonInput.value.trim()) {
        alert('Paste JSON into the field before importing.');
        return;
      }
      try {
        const next = JSON.parse(jsonInput.value);
        mergeImportedWorkspace(next, 'Merge JSON');
      } catch (err) {
        console.error('Invalid JSON', err);
        alert(err?.message || 'Invalid JSON. Please fix the JSON and try again.');
      }
    });
  }

  if (xmlExportBtn) {
    xmlExportBtn.addEventListener('click', () => {
      try {
        const payload = serializeStateToXml(state);
        if (xmlInput) xmlInput.value = payload;
        downloadXml(payload);
      } catch (err) {
        console.error('XML export failed', err);
        alert('XML export failed. Check console for details.');
      }
    });
  }

  if (xmlImportBtn) {
    xmlImportBtn.addEventListener('click', () => {
      if (!xmlInput || !xmlInput.value.trim()) {
        alert('Paste XML into the field before importing.');
        return;
      }
      try {
        const next = parseXmlToState(xmlInput.value);
        applyImportedWorkspace(next, 'Import XML');
      } catch (err) {
        console.error('Invalid XML', err);
        alert(err?.message || 'Invalid XML. Please fix the XML and try again.');
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

const duplicateCurrentProject = async () => {
  const copyName = getUniqueProjectCopyName(getProjectName(true));
  const snapshot = deepClone(state);
  snapshot.project = {
    ...DEFAULT_STATE.project,
    ...(snapshot.project || {}),
    id: null,
    name: copyName
  };

  if (window.location.protocol === 'file:') {
    replaceState(snapshot);
    localStorage.removeItem(STORAGE_KEYS.projectId);
    render();
    history = [];
    historyIndex = -1;
    pushHistory('Copy Project');
    updateUndoRedoButtons();
    updateHistoryMenu();
    updateProjectNameDisplay();
    return;
  }

  try {
    const res = await fetch('/api/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        id: null,
        name: copyName,
        data: snapshot
      })
    });
    const payload = await res.json();
    if (!res.ok) {
      throw new Error(payload?.error || 'Copy failed');
    }
    replaceState(snapshot);
    state.project = state.project || {};
    state.project.id = payload.id;
    state.project.name = copyName;
    localStorage.setItem(STORAGE_KEYS.projectId, String(payload.id));
    render();
    history = [];
    historyIndex = -1;
    pushHistory('Copy Project');
    updateUndoRedoButtons();
    updateHistoryMenu();
    updateProjectNameDisplay();
    await loadProjectList();
  } catch (err) {
    alert(`Copy failed: ${err.message}`);
  }
};

const deleteCurrentProject = async () => {
  const projectId = state.project?.id;
  const projectName = getProjectName(true);
  const confirmed = window.confirm(
    projectId
      ? `Delete "${projectName}"? This removes it from saved projects.`
      : `Discard "${projectName}" and start a blank project?`
  );
  if (!confirmed) return;

  if (!projectId || window.location.protocol === 'file:') {
    createNewProject();
    return;
  }

  try {
    const res = await fetch(`/api/projects/${projectId}`, { method: 'DELETE' });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) {
      throw new Error(payload?.error || 'Delete failed');
    }
    createNewProject();
    await loadProjectList();
  } catch (err) {
    alert(`Delete failed: ${err.message}`);
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

const downloadXml = (payload) => {
  const name = sanitizeFilename(state.project?.name || 'dashboard');
  const blob = new Blob([payload], { type: 'application/xml' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${name}.xml`;
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
      if (format === 'xml') {
        try {
          const payload = serializeStateToXml(state);
          downloadXml(payload);
        } catch (err) {
          console.error('XML export failed', err);
          alert('XML export failed. Check console for details.');
        }
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

  if (copyProjectBtn) {
    copyProjectBtn.addEventListener('click', () => {
      duplicateCurrentProject();
    });
  }

  if (deleteProjectBtn) {
    deleteProjectBtn.addEventListener('click', () => {
      deleteCurrentProject();
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
