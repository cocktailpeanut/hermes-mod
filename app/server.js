const express = require('express');
const fs = require('fs');
const fsp = require('fs/promises');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const figlet = require('figlet');
const Jimp = require('jimp');
const yaml = require('js-yaml');

const app = express();
const PORT = Number(process.env.PORT || 3210);
const HOST = '127.0.0.1';
const publicDir = path.join(__dirname, 'public');

const HERMES_HOME = process.env.HERMES_HOME || path.join(os.homedir(), '.hermes');
const SKINS_DIR = path.join(HERMES_HOME, 'skins');
const CONFIG_PATH = path.join(HERMES_HOME, 'config.yaml');

function resolveHermesAppRoot() {
  const candidates = [
    process.env.HERMES_AGENT_ROOT,
    path.join(os.homedir(), '.hermes', 'hermes-agent', 'app'),
    path.join('C:', 'pinokio', 'api', 'hermes-agent.pinokio.git', 'app')
  ].filter(Boolean);

  return candidates.find((candidate) => fs.existsSync(candidate)) || candidates[0] || '';
}

function resolveHermesPython(appRoot) {
  if (process.env.HERMES_PYTHON) return process.env.HERMES_PYTHON;

  const candidates = [
    path.join(appRoot, 'env', 'bin', 'python3'),
    path.join(appRoot, 'env', 'bin', 'python'),
    path.join(appRoot, 'env', 'Scripts', 'python.exe')
  ];

  return candidates.find((candidate) => fs.existsSync(candidate)) || candidates[candidates.length - 1];
}

const HERMES_APP_ROOT = resolveHermesAppRoot();
const SKIN_ENGINE_PATH = path.join(HERMES_APP_ROOT, 'hermes_cli', 'skin_engine.py');
const HERMES_PYTHON = resolveHermesPython(HERMES_APP_ROOT);
const HERO_ASCII_DEFAULTS = {
  style: 'braille',
  width: 40
};
const HERO_ASCII_WIDTH_LIMITS = { min: 16, max: 60 };
const DATA_URL_IMAGE_PATTERN = /^data:(image\/(?:png|jpe?g|gif|webp));base64,([a-z0-9+/=]+)$/i;
const IMAGE_EXTENSION_MIME_MAP = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.webp': 'image/webp'
};
const MAX_HERO_IMAGE_BYTES = 6 * 1024 * 1024;
const BRAILLE_BLANK = '\u2800';
const BRAILLE_BIT_GRID = [
  [0x01, 0x08],
  [0x02, 0x10],
  [0x04, 0x20],
  [0x40, 0x80]
];
const BAYER_4X4 = [
  [0, 8, 2, 10],
  [12, 4, 14, 6],
  [3, 11, 1, 9],
  [15, 7, 13, 5]
];
const HERO_AUTOCROP_OPTIONS = {
  tolerance: 0.1,
  cropOnlyFrames: false,
  leaveBorder: 1
};
const HERO_SHARPEN_KERNEL = [
  [0, -1, 0],
  [-1, 5, -1],
  [0, -1, 0]
];
const HERO_STYLE_MAP = {
  braille: {
    id: 'braille',
    label: 'Braille',
    description: 'High-detail unicode braille shading',
    renderer: 'braille',
    widthScale: 2,
    heightScale: 4,
    contrast: 0.38,
    sharpen: true,
    autocrop: true
  },
  ascii: {
    id: 'ascii',
    label: 'ASCII Ramp',
    description: 'Classic dense ASCII shading',
    renderer: 'ramp',
    chars: ' .\'`^",:;Il!i~+_-?][}{1)(|\\\\/tfjrxnuvczXYUJCLQ0OZmwqpdbkhao*#MW&8%B@$',
    widthScale: 1,
    heightScale: 2,
    contrast: 0.32,
    sharpen: true,
    autocrop: true,
    dither: true
  },
  blocks: {
    id: 'blocks',
    label: 'Blocks',
    description: 'Chunky unicode block shading',
    renderer: 'ramp',
    chars: ' ░▒▓█',
    widthScale: 1,
    heightScale: 2,
    contrast: 0.35,
    sharpen: true,
    autocrop: true,
    dither: false
  },
  dots: {
    id: 'dots',
    label: 'Dots',
    description: 'Soft dotted stipple look',
    renderer: 'ramp',
    chars: ' .·•●',
    widthScale: 1,
    heightScale: 2,
    contrast: 0.2,
    autocrop: true,
    dither: false
  }
};

const FIGLET_STYLE_MAP = {
  minimal: 'Standard',
  slant: 'Slant',
  small: 'Small',
  heavy: 'Doom',
  block: 'Big',
  shadow: 'ANSI Shadow',
  wide: 'Banner',
  compact: 'Lean'
};

const BUILTIN_SKIN_TEMPLATES = {
  default: {
    name: 'default',
    description: 'Classic Hermes — gold and kawaii',
    colors: {
      banner_border: '#CD7F32',
      banner_title: '#FFD700',
      banner_accent: '#FFBF00',
      banner_dim: '#B8860B',
      banner_text: '#FFF8DC',
      ui_accent: '#FFBF00',
      ui_label: '#4dd0e1',
      ui_ok: '#4caf50',
      ui_error: '#ef5350',
      ui_warn: '#ffa726',
      prompt: '#FFF8DC',
      input_rule: '#CD7F32',
      response_border: '#FFD700',
      session_label: '#DAA520',
      session_border: '#8B8682'
    },
    spinner: {},
    branding: {
      agent_name: 'Hermes Agent',
      welcome: 'Welcome to Hermes Agent! Type your message or /help for commands.',
      goodbye: 'Goodbye! ⚕',
      response_label: ' ⚕ Hermes ',
      prompt_symbol: '❯ ',
      help_header: '(^_^)? Available Commands'
    },
    tool_prefix: '┊',
    tool_emojis: {},
    banner_logo: '',
    banner_hero: ''
  },
  ares: {
    name: 'ares',
    description: 'War-god theme — crimson and bronze',
    colors: {
      banner_border: '#9F1C1C',
      banner_title: '#C7A96B',
      banner_accent: '#DD4A3A',
      banner_dim: '#6B1717',
      banner_text: '#F1E6CF',
      ui_accent: '#DD4A3A',
      ui_label: '#C7A96B',
      ui_ok: '#4caf50',
      ui_error: '#ef5350',
      ui_warn: '#ffa726',
      prompt: '#F1E6CF',
      input_rule: '#9F1C1C',
      response_border: '#C7A96B',
      session_label: '#C7A96B',
      session_border: '#6E584B'
    },
    spinner: {
      waiting_faces: ['(⚔)', '(⛨)', '(▲)', '(<>)', '(/)'],
      thinking_faces: ['(⚔)', '(⛨)', '(▲)', '(⌁)', '(<>)'],
      thinking_verbs: ['forging', 'marching', 'sizing the field', 'holding the line'],
      wings: [['⟪⚔', '⚔⟫'], ['⟪▲', '▲⟫']]
    },
    branding: {
      agent_name: 'Ares Agent',
      welcome: 'Welcome to Ares Agent! Type your message or /help for commands.',
      goodbye: 'Farewell, warrior! ⚔',
      response_label: ' ⚔ Ares ',
      prompt_symbol: '⚔ ❯ ',
      help_header: '(⚔) Available Commands'
    },
    tool_prefix: '╎',
    banner_logo: `[bold #A3261F] █████╗ ██████╗ ███████╗███████╗       █████╗  ██████╗ ███████╗███╗   ██╗████████╗[/]\n[bold #B73122]██╔══██╗██╔══██╗██╔════╝██╔════╝      ██╔══██╗██╔════╝ ██╔════╝████╗  ██║╚══██╔══╝[/]\n[#C93C24]███████║██████╔╝█████╗  ███████╗█████╗███████║██║  ███╗█████╗  ██╔██╗ ██║   ██║[/]\n[#D84A28]██╔══██║██╔══██╗██╔══╝  ╚════██║╚════╝██╔══██║██║   ██║██╔══╝  ██║╚██╗██║   ██║[/]\n[#E15A2D]██║  ██║██║  ██║███████╗███████║      ██║  ██║╚██████╔╝███████╗██║ ╚████║   ██║[/]\n[#EB6C32]╚═╝  ╚═╝╚═╝  ╚═╝╚══════╝╚══════╝      ╚═╝  ╚═╝ ╚═════╝ ╚══════╝╚═╝  ╚═══╝   ╚═╝[/]`,
    banner_hero: `[#9F1C1C]⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣤⣤⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀[/]\n[#9F1C1C]⠀⠀⠀⠀⠀⠀⠀⠀⠀⢀⣴⣿⠟⠻⣿⣦⡀⠀⠀⠀⠀⠀⠀⠀⠀⠀[/]\n[#C7A96B]⠀⠀⠀⠀⠀⠀⠀⣠⣾⡿⠋⠀⠀⠀⠙⢿⣷⣄⠀⠀⠀⠀⠀⠀⠀[/]\n[#C7A96B]⠀⠀⠀⠀⠀⢀⣾⡿⠋⠀⠀⢠⡄⠀⠀⠙⢿⣷⡀⠀⠀⠀⠀⠀[/]\n[#DD4A3A]⠀⠀⠀⠀⣰⣿⠟⠀⠀⠀⣰⣿⣿⣆⠀⠀⠀⠻⣿⣆⠀⠀⠀⠀[/]\n[#DD4A3A]⠀⠀⠀⢰⣿⠏⠀⠀⢀⣾⡿⠉⢿⣷⡀⠀⠀⠹⣿⡆⠀⠀⠀[/]\n[#9F1C1C]⠀⠀⠀⣿⡟⠀⠀⣠⣿⠟⠀⠀⠀⠻⣿⣄⠀⠀⢻⣿⠀⠀⠀[/]\n[#9F1C1C]⠀⠀⠀⣿⡇⠀⠀⠙⠋⠀⠀⚔⠀⠀⠙⠋⠀⠀⢸⣿⠀⠀⠀[/]\n[#6B1717]⠀⠀⠀⢿⣧⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⣼⡿⠀⠀⠀[/]\n[#6B1717]⠀⠀⠀⠘⢿⣷⣄⠀⠀⠀⠀⠀⠀⠀⠀⠀⣠⣾⡿⠃⠀⠀⠀[/]\n[#C7A96B]⠀⠀⠀⠀⠈⠻⣿⣷⣦⣤⣀⣀⣤⣤⣶⣿⠿⠋⠀⠀⠀⠀[/]\n[#C7A96B]⠀⠀⠀⠀⠀⠀⠀⠉⠛⠿⠿⠿⠿⠛⠉⠀⠀⠀⠀⠀⠀⠀[/]\n[#DD4A3A]⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⚔⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀[/]\n[dim #6B1717]⠀⠀⠀⠀⠀⠀⠀⠀war god online⠀⠀⠀⠀⠀⠀⠀⠀[/]`
  },
  mono: {
    name: 'mono',
    description: 'Monochrome — clean grayscale',
    colors: {
      banner_border: '#555555',
      banner_title: '#e6edf3',
      banner_accent: '#aaaaaa',
      banner_dim: '#444444',
      banner_text: '#c9d1d9',
      ui_accent: '#aaaaaa',
      ui_label: '#888888',
      ui_ok: '#888888',
      ui_error: '#cccccc',
      ui_warn: '#bbbbbb',
      prompt: '#c9d1d9',
      input_rule: '#555555',
      response_border: '#888888',
      session_label: '#aaaaaa',
      session_border: '#555555'
    },
    spinner: {
      waiting_faces: ['[·]', '[•]', '[·]'],
      thinking_faces: ['[·]', '[•]', '[·]'],
      thinking_verbs: ['thinking', 'processing', 'writing'],
      wings: [['[', ']']]
    },
    branding: {
      agent_name: 'Hermes Mono',
      welcome: 'Minimal mode engaged.',
      goodbye: 'Goodbye.',
      response_label: ' Mono ',
      prompt_symbol: '› ',
      help_header: 'Commands'
    },
    tool_prefix: '┆',
    tool_emojis: {},
    banner_logo: '',
    banner_hero: ''
  },
  slate: {
    name: 'slate',
    description: 'Cool blue developer-focused theme',
    colors: {
      banner_border: '#4C6FFF',
      banner_title: '#DCE6FF',
      banner_accent: '#7DD3FC',
      banner_dim: '#41557B',
      banner_text: '#CBD5E1',
      ui_accent: '#7DD3FC',
      ui_label: '#93C5FD',
      ui_ok: '#22C55E',
      ui_error: '#F87171',
      ui_warn: '#FBBF24',
      prompt: '#E2E8F0',
      input_rule: '#4C6FFF',
      response_border: '#60A5FA',
      session_label: '#93C5FD',
      session_border: '#475569'
    },
    spinner: {
      waiting_faces: ['◐', '◓', '◑', '◒'],
      thinking_faces: ['◐', '◓', '◑', '◒'],
      thinking_verbs: ['thinking', 'routing', 'compiling context'],
      wings: [['‹', '›']]
    },
    branding: {
      agent_name: 'Hermes Slate',
      welcome: 'Developer mode ready.',
      goodbye: 'Session closed.',
      response_label: ' Slate ',
      prompt_symbol: '› ',
      help_header: 'Developer Commands'
    },
    tool_prefix: '▏',
    tool_emojis: {},
    banner_logo: '',
    banner_hero: ''
  }
};

const BUILTIN_SKIN_CACHE_TTL_MS = Math.max(1000, Number.parseInt(process.env.BUILTIN_SKIN_CACHE_TTL_MS || '30000', 10) || 30000);
const builtinSkinCache = {
  loadedAt: 0,
  templates: BUILTIN_SKIN_TEMPLATES,
  source: 'fallback'
};

function normalizeBuiltinSkinTemplates(rawTemplates) {
  const source = rawTemplates && typeof rawTemplates === 'object' ? rawTemplates : {};
  const entries = Object.entries(source);
  const normalizedEntries = entries
    .map(([key, value]) => {
      if (!value || typeof value !== 'object') return null;
      const preferredName = value.name || key;
      let safeName;
      try {
        safeName = sanitizeSkinName(preferredName);
      } catch {
        safeName = sanitizeSkinName(key);
      }
      return [safeName, { ...value, name: safeName }];
    })
    .filter(Boolean);

  return normalizedEntries.length ? Object.fromEntries(normalizedEntries) : {};
}

function loadBuiltinSkinsFromHermes() {
  try {
    if (!HERMES_APP_ROOT || !fs.existsSync(HERMES_PYTHON) || !fs.existsSync(SKIN_ENGINE_PATH)) return {};
    const script = [
      'import json, sys',
      `sys.path.insert(0, ${JSON.stringify(HERMES_APP_ROOT)})`,
      'from hermes_cli.skin_engine import _BUILTIN_SKINS',
      'print(json.dumps(_BUILTIN_SKINS, ensure_ascii=False))'
    ].join('\n');
    const output = execFileSync(HERMES_PYTHON, ['-c', script], {
      encoding: 'utf8',
      windowsHide: true,
      cwd: HERMES_APP_ROOT,
      maxBuffer: 10 * 1024 * 1024
    });
    const parsed = JSON.parse(output);
    return normalizeBuiltinSkinTemplates(parsed);
  } catch (error) {
    console.warn('Failed to load built-in skins from Hermes skin engine:', error.message);
    return {};
  }
}

function getBuiltinSkinTemplates({ forceRefresh = false } = {}) {
  const now = Date.now();
  const cacheFresh = (now - builtinSkinCache.loadedAt) < BUILTIN_SKIN_CACHE_TTL_MS;
  if (!forceRefresh && cacheFresh) {
    return { templates: builtinSkinCache.templates, source: builtinSkinCache.source };
  }

  const liveTemplates = loadBuiltinSkinsFromHermes();
  if (Object.keys(liveTemplates).length) {
    builtinSkinCache.templates = liveTemplates;
    builtinSkinCache.source = 'hermes-skin-engine';
  } else {
    builtinSkinCache.templates = BUILTIN_SKIN_TEMPLATES;
    builtinSkinCache.source = 'fallback';
  }
  builtinSkinCache.loadedAt = now;

  return { templates: builtinSkinCache.templates, source: builtinSkinCache.source };
}

const COLOR_KEYS = [
  'banner_border', 'banner_title', 'banner_accent', 'banner_dim', 'banner_text',
  'ui_accent', 'ui_label', 'ui_ok', 'ui_error', 'ui_warn', 'prompt', 'input_rule',
  'response_border', 'session_label', 'session_border'
];

const BRANDING_KEYS = ['agent_name', 'welcome', 'goodbye', 'response_label', 'prompt_symbol', 'help_header'];
const TOOL_EMOJI_KEYS = ['terminal', 'web_search', 'browser_navigate', 'file', 'todo'];

app.use(express.json({ limit: '8mb' }));
app.use(express.static(publicDir));

function sanitizeSkinName(name) {
  const normalized = String(name || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '');
  if (!normalized) throw new Error('Skin name is required');
  return normalized;
}

function ensureObject(value) {
  return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
}

function clampInteger(value, fallback, min, max) {
  const parsed = Number.parseInt(String(value ?? ''), 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, parsed));
}

function normalizeAsciiBlock(input) {
  const lines = String(input || '')
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((line) => line.replace(/[\u2800\s]+$/g, ''));

  while (lines.length && /^[\u2800\s]*$/.test(lines[0])) lines.shift();
  while (lines.length && /^[\u2800\s]*$/.test(lines[lines.length - 1])) lines.pop();

  return lines.join('\n');
}

function decodeImagePayload(imageData) {
  const match = String(imageData || '').match(DATA_URL_IMAGE_PATTERN);
  if (!match) throw new Error('Provide a PNG, JPG, GIF, or WEBP image as a data URL');

  const buffer = Buffer.from(match[2], 'base64');
  if (!buffer.length) throw new Error('Image data is empty');
  if (buffer.length > MAX_HERO_IMAGE_BYTES) {
    throw new Error(`Image is too large. Maximum size is ${Math.round(MAX_HERO_IMAGE_BYTES / (1024 * 1024))}MB`);
  }

  return {
    mime: String(match[1] || '').toLowerCase(),
    buffer
  };
}

function normalizeImageBufferForJimp(imageBuffer, mime) {
  if (mime !== 'image/webp') return imageBuffer;

  try {
    return execFileSync('ffmpeg', [
      '-hide_banner',
      '-loglevel', 'error',
      '-i', 'pipe:0',
      '-frames:v', '1',
      '-f', 'image2pipe',
      '-vcodec', 'png',
      'pipe:1'
    ], {
      input: imageBuffer,
      encoding: 'buffer',
      maxBuffer: 32 * 1024 * 1024
    });
  } catch (error) {
    throw new Error('Failed to convert the WEBP image for hero generation');
  }
}

function getImageMimeFromPath(filePath) {
  const extension = path.extname(String(filePath || '')).toLowerCase();
  const mime = IMAGE_EXTENSION_MIME_MAP[extension];
  if (!mime) throw new Error('Choose a PNG, JPG, GIF, or WEBP image');
  return mime;
}

function openImageFilePickerWindows() {
  const script = [
    '[Console]::OutputEncoding = [System.Text.Encoding]::UTF8',
    'Add-Type -AssemblyName System.Windows.Forms',
    '$dialog = New-Object System.Windows.Forms.OpenFileDialog',
    '$dialog.Filter = "Images|*.png;*.jpg;*.jpeg;*.gif;*.webp"',
    '$dialog.Title = "Choose image"',
    '$dialog.Multiselect = $false',
    '$dialog.CheckFileExists = $true',
    '$dialog.CheckPathExists = $true',
    '$result = $dialog.ShowDialog()',
    'if ($result -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $dialog.FileName }'
  ].join('; ');

  const output = execFileSync('powershell.exe', ['-NoProfile', '-STA', '-Command', script], {
    encoding: 'utf8'
  }).trim();

  return output || null;
}

function openImageFilePickerMac() {
  try {
    const output = execFileSync('osascript', ['-e', 'POSIX path of (choose file with prompt "Choose image")'], {
      encoding: 'utf8'
    }).trim();
    return output || null;
  } catch (error) {
    if (error.status === 1) return null;
    throw error;
  }
}

function openImageFilePickerLinux() {
  try {
    const output = execFileSync('zenity', [
      '--file-selection',
      '--title=Choose image',
      '--file-filter=Images | *.png *.jpg *.jpeg *.gif *.webp'
    ], {
      encoding: 'utf8'
    }).trim();
    return output || null;
  } catch (error) {
    if (error.status === 1) return null;
    throw new Error('System image picker is unavailable on this machine');
  }
}

function openImageFilePicker() {
  switch (process.platform) {
    case 'win32':
      return openImageFilePickerWindows();
    case 'darwin':
      return openImageFilePickerMac();
    case 'linux':
      return openImageFilePickerLinux();
    default:
      throw new Error(`System image picker is not supported on ${process.platform}`);
  }
}

function getHeroStyle(styleId) {
  const normalized = String(styleId || '').trim().toLowerCase();
  return HERO_STYLE_MAP[normalized] || HERO_STYLE_MAP[HERO_ASCII_DEFAULTS.style];
}

function calculateAutoHeight(sourceWidth, sourceHeight, heroStyle, outputWidth) {
  if (sourceWidth <= 0 || sourceHeight <= 0) return outputWidth;
  const aspectFactor = heroStyle.widthScale / heroStyle.heightScale;
  return Math.max(1, Math.round((sourceHeight / sourceWidth) * outputWidth * aspectFactor));
}

function clampUnit(value) {
  return Math.max(0, Math.min(1, value));
}

function getPixelDarkness(image, x, y) {
  const { r, g, b, a } = Jimp.intToRGBA(image.getPixelColor(x, y));
  const alpha = a / 255;
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255;
  return clampUnit((1 - luminance) * alpha);
}

function buildDarknessGrid(image, sampleWidth = 1, sampleHeight = 1) {
  const grid = [];

  for (let y = 0; y < image.bitmap.height; y += sampleHeight) {
    const row = [];
    for (let x = 0; x < image.bitmap.width; x += sampleWidth) {
      let totalDarkness = 0;
      let totalSamples = 0;

      for (let dy = 0; dy < sampleHeight; dy += 1) {
        for (let dx = 0; dx < sampleWidth; dx += 1) {
          const sampleX = x + dx;
          const sampleY = y + dy;
          if (sampleX >= image.bitmap.width || sampleY >= image.bitmap.height) continue;
          totalDarkness += getPixelDarkness(image, sampleX, sampleY);
          totalSamples += 1;
        }
      }

      row.push(totalSamples ? totalDarkness / totalSamples : 0);
    }
    grid.push(row);
  }

  return grid;
}

function diffuseDarknessError(grid, x, y, error) {
  const height = grid.length;
  const width = grid[0]?.length || 0;
  if (!width || !height) return;

  if (x + 1 < width) grid[y][x + 1] = clampUnit(grid[y][x + 1] + (error * 7 / 16));
  if (y + 1 >= height) return;
  if (x > 0) grid[y + 1][x - 1] = clampUnit(grid[y + 1][x - 1] + (error * 3 / 16));
  grid[y + 1][x] = clampUnit(grid[y + 1][x] + (error * 5 / 16));
  if (x + 1 < width) grid[y + 1][x + 1] = clampUnit(grid[y + 1][x + 1] + (error * 1 / 16));
}

function getAdaptiveBrailleThreshold(cellAverageDarkness) {
  if (cellAverageDarkness >= 0.72) return 0.34;
  if (cellAverageDarkness >= 0.48) return 0.42;
  return 0.5;
}

function renderBrailleArt(image) {
  const lines = [];
  for (let y = 0; y < image.bitmap.height; y += 4) {
    let line = '';
    for (let x = 0; x < image.bitmap.width; x += 2) {
      let bits = 0;
      let totalDarkness = 0;
      let totalSamples = 0;

      for (let dy = 0; dy < 4; dy += 1) {
        for (let dx = 0; dx < 2; dx += 1) {
          const sampleX = x + dx;
          const sampleY = y + dy;
          if (sampleX >= image.bitmap.width || sampleY >= image.bitmap.height) continue;
          totalDarkness += getPixelDarkness(image, sampleX, sampleY);
          totalSamples += 1;
        }
      }

      const cellAverageDarkness = totalSamples ? totalDarkness / totalSamples : 0;
      if (cellAverageDarkness < 0.04) {
        line += BRAILLE_BLANK;
        continue;
      }

      const threshold = getAdaptiveBrailleThreshold(cellAverageDarkness);
      for (let dy = 0; dy < 4; dy += 1) {
        for (let dx = 0; dx < 2; dx += 1) {
          const sampleX = x + dx;
          const sampleY = y + dy;
          if (sampleX >= image.bitmap.width || sampleY >= image.bitmap.height) continue;
          if (getPixelDarkness(image, sampleX, sampleY) >= threshold) {
            bits |= BRAILLE_BIT_GRID[dy][dx];
          }
        }
      }
      line += bits ? String.fromCharCode(0x2800 + bits) : BRAILLE_BLANK;
    }
    lines.push(line);
  }
  return normalizeAsciiBlock(lines.join('\n'));
}

function renderRampArt(image, chars, useDither = false, sampleWidth = 1, sampleHeight = 1) {
  const palette = String(chars || ' .:-=+*#%@');
  const maxIndex = Math.max(0, palette.length - 1);
  const darknessGrid = buildDarknessGrid(image, sampleWidth, sampleHeight);
  const lines = [];

  for (let y = 0; y < darknessGrid.length; y += 1) {
    let line = '';
    for (let x = 0; x < darknessGrid[y].length; x += 1) {
      const darkness = darknessGrid[y][x];
      const index = Math.max(0, Math.min(maxIndex, Math.round(darkness * maxIndex)));
      if (useDither && maxIndex > 0) {
        diffuseDarknessError(darknessGrid, x, y, darkness - (index / maxIndex));
      }
      line += palette[index];
    }
    lines.push(line);
  }

  return normalizeAsciiBlock(lines.join('\n'));
}

function preprocessHeroImage(sourceImage, heroStyle, outputWidth) {
  const image = sourceImage.clone();

  if (heroStyle.autocrop) {
    image.autocrop(HERO_AUTOCROP_OPTIONS);
  }

  const outputHeight = calculateAutoHeight(image.bitmap.width, image.bitmap.height, heroStyle, outputWidth);
  const pixelWidth = Math.max(1, outputWidth * heroStyle.widthScale);
  const pixelHeight = Math.max(1, outputHeight * heroStyle.heightScale);

  image
    .greyscale()
    .resize(pixelWidth, pixelHeight, heroStyle.resizeMode || Jimp.RESIZE_BEZIER)
    .normalize()
    .contrast(heroStyle.contrast ?? 0.3);

  if (heroStyle.sharpen) {
    image.convolute(HERO_SHARPEN_KERNEL);
  }

  return { image, outputHeight };
}

async function generateHeroArt(imageBuffer, { style, width }) {
  const sourceImage = await Jimp.read(imageBuffer);
  const heroStyle = getHeroStyle(style);
  const outputWidth = clampInteger(width, HERO_ASCII_DEFAULTS.width, HERO_ASCII_WIDTH_LIMITS.min, HERO_ASCII_WIDTH_LIMITS.max);
  const { image, outputHeight } = preprocessHeroImage(sourceImage, heroStyle, outputWidth);

  if (heroStyle.renderer === 'ramp') {
    return {
      ascii: renderRampArt(image, heroStyle.chars, heroStyle.dither, heroStyle.widthScale, heroStyle.heightScale),
      width: outputWidth,
      height: outputHeight
    };
  }

  return {
    ascii: renderBrailleArt(image),
    width: outputWidth,
    height: outputHeight
  };
}

function normalizeSkin(input = {}) {
  const source = ensureObject(input);
  const colors = ensureObject(source.colors);
  const branding = ensureObject(source.branding);
  const spinner = ensureObject(source.spinner);
  const toolEmojis = ensureObject(source.tool_emojis);

  const normalized = {
    name: sanitizeSkinName(source.name || 'custom-skin'),
    description: String(source.description || ''),
    colors: {},
    spinner: {},
    branding: {},
    tool_prefix: String(source.tool_prefix || '┊'),
    tool_emojis: {},
    banner_logo: String(source.banner_logo || ''),
    banner_hero: String(source.banner_hero || '')
  };

  for (const key of COLOR_KEYS) {
    if (colors[key]) normalized.colors[key] = String(colors[key]);
  }
  for (const key of BRANDING_KEYS) {
    if (branding[key]) normalized.branding[key] = String(branding[key]);
  }
  for (const key of TOOL_EMOJI_KEYS) {
    if (toolEmojis[key]) normalized.tool_emojis[key] = String(toolEmojis[key]);
  }

  const waitingFaces = Array.isArray(spinner.waiting_faces) ? spinner.waiting_faces.map(String).filter(Boolean) : [];
  const thinkingFaces = Array.isArray(spinner.thinking_faces) ? spinner.thinking_faces.map(String).filter(Boolean) : [];
  const thinkingVerbs = Array.isArray(spinner.thinking_verbs) ? spinner.thinking_verbs.map(String).filter(Boolean) : [];
  const wings = Array.isArray(spinner.wings)
    ? spinner.wings
        .filter((pair) => Array.isArray(pair) && pair.length === 2)
        .map((pair) => [String(pair[0] || ''), String(pair[1] || '')])
    : [];

  if (waitingFaces.length) normalized.spinner.waiting_faces = waitingFaces;
  if (thinkingFaces.length) normalized.spinner.thinking_faces = thinkingFaces;
  if (thinkingVerbs.length) normalized.spinner.thinking_verbs = thinkingVerbs;
  if (wings.length) normalized.spinner.wings = wings;

  return normalized;
}

function dumpSkinYaml(skin) {
  return yaml.dump(normalizeSkin(skin), {
    lineWidth: 120,
    noRefs: true,
    quotingType: '"',
    forceQuotes: false
  });
}

async function ensureHermesDirs() {
  await fsp.mkdir(SKINS_DIR, { recursive: true });
}

async function readYamlFile(filePath) {
  const content = await fsp.readFile(filePath, 'utf8');
  const parsed = yaml.load(content) || {};
  return normalizeSkin(parsed);
}

function getBuiltinSkins() {
  const { templates } = getBuiltinSkinTemplates();
  return Object.entries(templates).map(([key, skin]) => ({
    name: skin?.name || key,
    description: skin?.description || '',
    source: 'builtin'
  }));
}

async function getUserSkins() {
  await ensureHermesDirs();
  const entries = await fsp.readdir(SKINS_DIR, { withFileTypes: true });
  const skins = [];
  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith('.yaml')) continue;
    const fullPath = path.join(SKINS_DIR, entry.name);
    try {
      const parsed = await readYamlFile(fullPath);
      const stats = await fsp.stat(fullPath);
      skins.push({
        name: parsed.name || path.basename(entry.name, '.yaml'),
        description: parsed.description || '',
        source: 'user',
        file: fullPath,
        modified_at: stats.mtime.toISOString()
      });
    } catch (error) {
      skins.push({
        name: path.basename(entry.name, '.yaml'),
        description: 'Unreadable skin file',
        source: 'user',
        file: fullPath,
        invalid: true,
        error: error.message
      });
    }
  }
  skins.sort((a, b) => a.name.localeCompare(b.name));
  return skins;
}

async function getActiveSkin() {
  try {
    const configContent = await fsp.readFile(CONFIG_PATH, 'utf8');
    const parsed = yaml.load(configContent) || {};
    return parsed?.display?.skin || 'default';
  } catch {
    return 'default';
  }
}

function updateDisplaySkinInConfig(content, skinName) {
  const lines = content.split(/\r?\n/);
  let displayIndex = lines.findIndex((line) => /^display:\s*$/.test(line));

  if (displayIndex === -1) {
    const suffix = content.endsWith('\n') || content.length === 0 ? '' : '\n';
    return `${content}${suffix}display:\n  skin: ${skinName}\n`;
  }

  let blockEnd = lines.length;
  for (let i = displayIndex + 1; i < lines.length; i += 1) {
    const line = lines[i];
    if (/^[^\s#][^:]*:\s*/.test(line)) {
      blockEnd = i;
      break;
    }
  }

  for (let i = displayIndex + 1; i < blockEnd; i += 1) {
    if (/^\s{2}skin:\s*/.test(lines[i])) {
      lines[i] = `  skin: ${skinName}`;
      return `${lines.join('\n')}\n`;
    }
  }

  lines.splice(displayIndex + 1, 0, `  skin: ${skinName}`);
  return `${lines.join('\n')}\n`;
}

async function activateSkin(skinName) {
  await ensureHermesDirs();
  let content = '';
  try {
    content = await fsp.readFile(CONFIG_PATH, 'utf8');
  } catch {
    content = '';
  }
  const updated = updateDisplaySkinInConfig(content, sanitizeSkinName(skinName));
  await fsp.writeFile(CONFIG_PATH, updated, 'utf8');
}

function makeSkinFilePath(name) {
  return path.join(SKINS_DIR, `${sanitizeSkinName(name)}.yaml`);
}

app.get('/api/status', async (req, res) => {
  try {
    const [active_skin, user_skins] = await Promise.all([getActiveSkin(), getUserSkins()]);
    const preset_skins = getBuiltinSkins();
    res.json({
      active_skin,
      hermes_home: HERMES_HOME,
      skins_dir: SKINS_DIR,
      config_path: CONFIG_PATH,
      preset_skins,
      builtin_skins: preset_skins,
      user_skins
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/presets', async (req, res) => {
  try {
    res.json({ presets: getBuiltinSkins() });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/skins/:name', async (req, res) => {
  try {
    const source = req.query.source === 'builtin' ? 'builtin' : 'user';
    const name = sanitizeSkinName(req.params.name);
    if (source === 'builtin') {
      const { templates } = getBuiltinSkinTemplates();
      const skin = templates[name];
      if (!skin) return res.status(404).json({ error: 'Built-in skin not found' });
      return res.json({ source, skin: normalizeSkin(skin) });
    }
    const filePath = makeSkinFilePath(name);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'Skin not found' });
    }
    const skin = await readYamlFile(filePath);
    return res.json({ source, skin });
  } catch (error) {
    return res.status(400).json({ error: error.message });
  }
});

app.post('/api/skins', async (req, res) => {
  try {
    await ensureHermesDirs();
    const skin = normalizeSkin(req.body);
    const filePath = makeSkinFilePath(skin.name);
    await fsp.writeFile(filePath, dumpSkinYaml(skin), 'utf8');
    res.json({ ok: true, skin, file: filePath });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.put('/api/skins/:name', async (req, res) => {
  try {
    await ensureHermesDirs();
    const requestedName = sanitizeSkinName(req.params.name);
    const skin = normalizeSkin({ ...req.body, name: requestedName });
    const filePath = makeSkinFilePath(requestedName);
    await fsp.writeFile(filePath, dumpSkinYaml(skin), 'utf8');
    res.json({ ok: true, skin, file: filePath });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.delete('/api/skins/:name', async (req, res) => {
  try {
    const name = sanitizeSkinName(req.params.name);
    const filePath = makeSkinFilePath(name);
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'Skin not found' });
    await fsp.unlink(filePath);
    res.json({ ok: true });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post('/api/activate/:name', async (req, res) => {
  try {
    const name = sanitizeSkinName(req.params.name);
    await activateSkin(name);
    res.json({ ok: true, active_skin: name });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.get('/api/meta', async (req, res) => {
  const { source: builtin_skin_source } = getBuiltinSkinTemplates();

  let skinEngineFound = false;
  let hermesPythonFound = false;
  try {
    await fsp.access(SKIN_ENGINE_PATH);
    skinEngineFound = true;
  } catch {
    skinEngineFound = false;
  }
  try {
    await fsp.access(HERMES_PYTHON);
    hermesPythonFound = true;
  } catch {
    hermesPythonFound = false;
  }

  res.json({
    hermes_home: HERMES_HOME,
    skins_dir: SKINS_DIR,
    config_path: CONFIG_PATH,
    hermes_app_root: HERMES_APP_ROOT,
    hermes_python: HERMES_PYTHON,
    hermes_python_found: hermesPythonFound,
    skin_engine_path: SKIN_ENGINE_PATH,
    skin_engine_found: skinEngineFound,
    builtin_skin_source
  });
});

app.get('/api/logo-styles', (req, res) => {
  res.json({
    styles: Object.entries(FIGLET_STYLE_MAP).map(([id, font]) => ({ id, font }))
  });
});

app.get('/api/hero-styles', (req, res) => {
  res.json({
    styles: Object.values(HERO_STYLE_MAP).map(({ id, label, description }) => ({ id, label, description }))
  });
});

app.post('/api/generate-logo', async (req, res) => {
  try {
    const title = String(req.body?.title || '').trim();
    const styleId = String(req.body?.style || 'minimal').trim().toLowerCase();
    if (!title) return res.status(400).json({ error: 'Title is required' });
    const font = FIGLET_STYLE_MAP[styleId] || FIGLET_STYLE_MAP.minimal;
    const ascii = figlet.textSync(title, {
      font,
      horizontalLayout: 'default',
      verticalLayout: 'default',
      width: 120,
      whitespaceBreak: true
    });
    res.json({ ok: true, title, style: styleId, font, ascii });
  } catch (error) {
    res.status(400).json({ error: error.message });
  }
});

app.post('/api/generate-hero', async (req, res) => {
  try {
    const { mime, buffer } = decodeImagePayload(req.body?.image_data || req.body?.imageData);
    const imageBuffer = normalizeImageBufferForJimp(buffer, mime);
    const heroStyle = getHeroStyle(req.body?.style);
    const generated = await generateHeroArt(imageBuffer, { style: heroStyle.id, width: req.body?.width });

    res.json({
      ok: true,
      ascii: generated.ascii,
      options: { style: heroStyle.id, width: generated.width, height: generated.height }
    });
  } catch (error) {
    res.status(400).json({ error: error.message || 'Failed to generate hero art' });
  }
});

app.post('/api/pick-hero-image', async (req, res) => {
  try {
    const filePath = openImageFilePicker();
    if (!filePath) {
      return res.json({ ok: false, canceled: true });
    }

    const mime = getImageMimeFromPath(filePath);
    const imageBuffer = await fsp.readFile(filePath);
    if (!imageBuffer.length) {
      return res.status(400).json({ error: 'Selected image is empty' });
    }
    if (imageBuffer.length > MAX_HERO_IMAGE_BYTES) {
      return res.status(400).json({ error: `Image is too large. Maximum size is ${Math.round(MAX_HERO_IMAGE_BYTES / (1024 * 1024))}MB` });
    }

    res.json({
      ok: true,
      file_name: path.basename(filePath),
      image_data: `data:${mime};base64,${imageBuffer.toString('base64')}`
    });
  } catch (error) {
    res.status(400).json({ error: error.message || 'Failed to choose image' });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(publicDir, 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`Hermes Skin Studio running at http://${HOST}:${PORT}`);
});
