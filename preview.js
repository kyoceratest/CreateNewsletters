// preview.js
// Extracted from preview.html to keep JS separate without changing UI/behavior.
// On-screen text remains French; code and comments are in English.

function buildCategoryBarElement() {
  const bar = document.createElement('div');
  bar.id = 'categoryBar';
  bar.style.display = 'flex';
  bar.style.flexWrap = 'wrap';
  bar.style.gap = '8px';
  bar.style.marginTop = '12px';

  const config = [
    { cat: 'edito', label: 'Édito' },
    { cat: 'com_actus', label: 'Com & Actus' },
    { cat: 'animation', label: 'Animation' },
    { cat: 'zoom_materiel', label: 'Zoom Matériel' },
    { cat: 'solutions_insights', label: 'Solutions & Insights' },
    { cat: 'outils_astuces', label: 'Outils & Astuces' },
    { cat: 'branding', label: 'Branding' },
    { cat: 'formations', label: 'Formations' }
  ];

  config.forEach(item => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'kyo-btn';
    btn.dataset.cat = item.cat;
    btn.textContent = item.label;
    btn.style.padding = '8px 12px';
    btn.style.border = '1px solid #ccc';
    btn.style.borderRadius = '6px';
    btn.style.background = '#fff';
    btn.style.cursor = 'pointer';
    bar.appendChild(btn);
  });

  return bar;
}

function ensureCategoryBarBeforeTarget(target) {
  if (!target || !target.parentNode) return null;
  const parent = target.parentNode;
  const existing = parent.querySelector('#categoryBar');
  if (existing && existing.parentNode === parent) return existing;
  const bar = buildCategoryBarElement();
  parent.insertBefore(bar, target);
  return bar;
}

// 1) Preview Content Handler: receives postMessage from history modal
function handlePreviewContent(data) {
  const { content, title } = data || {};
  const target = document.getElementById('loadedContent');
  const historyId = getUrlParameter('historyId');

  const renderHtml = (html) => {
    if (!target) return;
    const wrapper = document.createElement('div');
    wrapper.innerHTML = html || '';
    try {
      normalizeAssetsInWrapper(wrapper);
      // Sanitize editing affordances
      wrapper.querySelectorAll('[contenteditable]').forEach(el => {
        el.removeAttribute('contenteditable');
        el.contentEditable = 'false';
      });
      ['input','textarea','select','button'].forEach(sel => {
        wrapper.querySelectorAll(sel).forEach(ctrl => {
          ctrl.setAttribute('disabled', '');
          ctrl.setAttribute('tabindex', '-1');
        });
      });
    } catch (_) {}
    ensureCategoryBarBeforeTarget(target);
    target.innerHTML = '';
    while (wrapper.firstChild) target.appendChild(wrapper.firstChild);
    if (title) { try { document.title = title + ' - Aperçu'; } catch (_) {} }
  };

  // Prefer full content from IndexedDB when a historyId is present
  if (historyId) {
    try {
      getFullContentFromIDB(historyId)
        .then(full => { renderHtml(full || content || ''); })
        .catch(() => { renderHtml(content || ''); });
    } catch (_) {
      renderHtml(content || '');
    }
  } else {
    renderHtml(content || '');
  }
}

// Resolve bare asset names relative to a base directory (e.g., TypeNews/animation)
function normalizeAssetsRelative(wrapper, baseDir) {
  if (!baseDir) return;
  const prefix = baseDir.endsWith('/') ? baseDir : baseDir + '/';
  const isBare = (p) => !!p && !/^https?:\/\//i.test(p) && !/^data:/i.test(p) && !/^blob:/i.test(p) && !/^file:\/\//i.test(p) && !/^\//.test(p) && !p.includes('/');
  const winAbsRe = /^[a-zA-Z]:\\|^\\\\/; // C:\... or \\server\share
  const fileProtoRe = /^file:\/\//i; // file:///C:/...

  // imgs
  Array.from(wrapper.querySelectorAll('img[src]')).forEach(img => {
    let raw = img.getAttribute('src') || '';
    if (winAbsRe.test(raw) || fileProtoRe.test(raw)) return; // leave absolute OS paths to other normalizers
    if (isBare(raw)) img.setAttribute('src', prefix + raw);
  });
  // video/audio elements with src
  Array.from(wrapper.querySelectorAll('video[src],audio[src]')).forEach(el => {
    let raw = el.getAttribute('src') || '';
    if (winAbsRe.test(raw) || fileProtoRe.test(raw)) return;
    if (isBare(raw)) el.setAttribute('src', prefix + raw);
  });
  // <source src> inside media
  Array.from(wrapper.querySelectorAll('video source[src], audio source[src]')).forEach(srcEl => {
    let raw = srcEl.getAttribute('src') || '';
    if (winAbsRe.test(raw) || fileProtoRe.test(raw)) return;
    if (isBare(raw)) srcEl.setAttribute('src', prefix + raw);
  });
}

// Listen for messages from the parent window (history modal)
window.addEventListener('message', function (event) {
  if (event.data && event.data.type === 'previewContent') {
    handlePreviewContent(event.data);
  }
});

// 2) Shared helpers
function getUrlParameter(name) {
  name = name.replace(/[\[]/, '\\[').replace(/[\]]/, '\\]');
  const regex = new RegExp('[\\?&]' + name + '=([^&#]*)');
  const results = regex.exec(location.search);
  return results === null ? '' : decodeURIComponent(results[1].replace(/\+/g, ' '));
}

// 3) Initialize base behaviors depending on preview mode
(function initBase() {
  document.addEventListener('DOMContentLoaded', function () {
    const historyId = getUrlParameter('historyId');
    if (historyId) {
      // In preview mode: keep header and footer visible; just clear the content area
      const target = document.getElementById('loadedContent');
      if (target) target.innerHTML = '';
      // Fallback load by ID (in case postMessage cannot access due to file:// or cross-origin)
      try { loadHistoryPreviewById(historyId); } catch (_) {}
    } else {
      // Normal page: hamburger toggle
      const hamburger = document.getElementById('hamburger');
      const headerNav = document.querySelector('.header-nav');
      if (hamburger && headerNav) {
        hamburger.addEventListener('click', function () {
          headerNav.classList.toggle('active');
          const spans = hamburger.querySelectorAll('span');
          if (headerNav.classList.contains('active')) {
            spans[0].style.transform = 'rotate(45deg) translate(5px, 5px)';
            spans[1].style.opacity = '0';
            spans[2].style.transform = 'rotate(-45deg) translate(7px, -6px)';
          } else {
            spans[0].style.transform = 'none';
            spans[1].style.opacity = '1';
            spans[2].style.transform = 'none';
          }
        });
      }
    }
  });
})();

// 6) Fallback: load preview content directly from storage by historyId
async function loadHistoryPreviewById(historyId) {
  const id = String(historyId);
  const target = document.getElementById('loadedContent');
  if (!target) return;
  let item = null;
  try {
    const raw = localStorage.getItem('newsletterHistory');
    const list = raw ? JSON.parse(raw) : [];
    if (Array.isArray(list)) item = list.find(it => String(it.id) === id) || null;
  } catch (_) {}
  if (!item) {
    try {
      const raw2 = sessionStorage.getItem('newsletterHistoryFallback');
      const list2 = raw2 ? JSON.parse(raw2) : [];
      if (Array.isArray(list2)) item = list2.find(it => String(it.id) === id) || null;
    } catch (_) {}
  }
  if (!item) return;

  // Sanitize: remove editing affordances, disable form controls
  const wrapper = document.createElement('div');
  let fullHtml = '';
  try {
    // Prefer full snapshot from IndexedDB if available
    fullHtml = await getFullContentFromIDB(id);
  } catch (_) { fullHtml = ''; }
  wrapper.innerHTML = fullHtml || (item.content || '');
  try {
    normalizeAssetsInWrapper(wrapper);
    wrapper.querySelectorAll('[contenteditable]').forEach(el => {
      el.removeAttribute('contenteditable');
      el.contentEditable = 'false';
    });
    ['input','textarea','select','button'].forEach(sel => {
      wrapper.querySelectorAll(sel).forEach(ctrl => {
        ctrl.setAttribute('disabled', '');
        ctrl.setAttribute('tabindex', '-1');
      });
    });
  } catch (_) {}
  ensureCategoryBarBeforeTarget(target);
  target.innerHTML = '';
  while (wrapper.firstChild) target.appendChild(wrapper.firstChild);
  try { if (item.name) document.title = item.name + ' - Aperçu'; } catch (_) {}
}

// Local IndexedDB helpers (preview page may not include editor.js)
function openHistoryDB() {
  return new Promise((resolve, reject) => {
    try {
      const req = indexedDB.open('NewsletterDB', 1);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains('historyFull')) {
          db.createObjectStore('historyFull', { keyPath: 'id' });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    } catch (e) { reject(e); }
  });
}

function getFullContentFromIDB(id) {
  return openHistoryDB().then(db => new Promise(resolve => {
    const tx = db.transaction('historyFull', 'readonly');
    const store = tx.objectStore('historyFull');
    const req = store.get(String(id));
    req.onsuccess = () => { try { db.close(); } catch (_) {}; resolve((req.result && req.result.content) || ''); };
    req.onerror = () => { try { db.close(); } catch (_) {}; resolve(''); };
  })).catch(() => '');
}

// Helper: normalize bare asset paths (e.g., "image.jpg" -> "Image/image.jpg")
function normalizeAssetsInWrapper(wrapper) {
  const isBare = (p) => !!p && !/^https?:\/\//i.test(p) && !/^data:/i.test(p) && !/^blob:/i.test(p) && !/^file:\/\//i.test(p) && !/^\//.test(p) && !p.includes('/');
  const clean = (p) => (p || '').replace(/^\.\//, '');
  const toFolder = (p) => {
    const name = clean(p);
    const lower = name.toLowerCase();
    const isImg = /(\.png|\.jpg|\.jpeg|\.gif|\.svg|\.webp|\.bmp|\.ico)$/i.test(lower);
    const isMedia = /(\.mp4|\.webm|\.ogg|\.mp3|\.wav)$/i.test(lower);
    if (isImg) return 'Image/' + name;
    if (isMedia) return 'media/' + name;
    return name;
  };
  const winAbsRe = /^[a-zA-Z]:\\|^\\\\/; // C:\... or \\server\share
  const fileProtoRe = /^file:\/\//i; // file:///C:/...
  const basename = (p) => {
    const norm = (p || '').replace(/^[a-zA-Z]+:\/\//, '') // strip proto if any
                          .replace(/^file:\/\//i, '')
                          .replace(/\\/g, '/');
    const parts = norm.split('/');
    return parts[parts.length - 1] || '';
  };
  // imgs
  Array.from(wrapper.querySelectorAll('img[src]')).forEach(img => {
    let raw = img.getAttribute('src') || '';
    if (winAbsRe.test(raw) || fileProtoRe.test(raw)) {
      const name = basename(raw);
      if (name) raw = toFolder(name);
    }
    if (isBare(raw)) raw = toFolder(raw);
    img.setAttribute('src', raw);
  });
  // video/audio elements with src
  Array.from(wrapper.querySelectorAll('video[src],audio[src]')).forEach(el => {
    let raw = el.getAttribute('src') || '';
    if (winAbsRe.test(raw) || fileProtoRe.test(raw)) {
      const name = basename(raw);
      if (name) raw = toFolder(name);
    }
    // If a blob URL is present, try resolving via data-local-filename hint
    if (/^blob:\/\//i.test(raw)) {
      const hint = el.getAttribute('data-local-filename') || '';
      if (hint) raw = toFolder(hint);
    }
    if (isBare(raw)) raw = toFolder(raw);
    el.setAttribute('src', raw);
  });
  // <source src> inside media
  Array.from(wrapper.querySelectorAll('video source[src], audio source[src]')).forEach(srcEl => {
    let raw = srcEl.getAttribute('src') || '';
    if (winAbsRe.test(raw) || fileProtoRe.test(raw)) {
      const name = basename(raw);
      if (name) raw = toFolder(name);
    }
    if (/^blob:\/\//i.test(raw)) {
      // Prefer own hint; else look at parent <video>
      let hint = srcEl.getAttribute('data-local-filename') || '';
      if (!hint) {
        try { const parent = srcEl.closest('video,audio'); hint = parent ? (parent.getAttribute('data-local-filename') || '') : ''; } catch (_) {}
      }
      if (hint) raw = toFolder(hint);
    }
    if (isBare(raw)) raw = toFolder(raw);
    srcEl.setAttribute('src', raw);
  });
}
// 4) Content Loader (only when not in preview mode)
(function initContentLoader() {
  const historyId = getUrlParameter('historyId');
  if (historyId) return; // skip in preview mode

  document.addEventListener('DOMContentLoaded', function () {
    const selectEl = document.getElementById('contentSelect');
    const statusEl = document.getElementById('contentStatus');
    const targetEl = document.getElementById('loadedContent');

    function extractTitleFromDoc(doc) {
      let el = doc.querySelector('span[style*="font-size: 52px"], span[style*="font-size:52px"], [style*="font-size: 52px"], [style*="font-size:52px"]');
      if (el && (el.innerText || el.textContent)) return (el.innerText || el.textContent).trim();
      el = doc.querySelector('font[size="7"]');
      if (el && (el.innerText || el.textContent)) return (el.innerText || el.textContent).trim();
      el = doc.querySelector('h1, h2, h3');
      if (el && (el.innerText || el.textContent)) return (el.innerText || el.textContent).trim();
      return '';
    }

    async function syncDropdownLabels() {
      if (!selectEl) return;
      const opts = Array.from(selectEl.options).filter(o => o.value && o.value.endsWith('.html'));
      for (const opt of opts) {
        try {
          const syncFile = opt.value;
          if (syncFile === 'teteSuperieure.html') {
            const res = await fetch(`${syncFile}?t=${Date.now()}`);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const html = await res.text();
            const parser = new DOMParser();
            const doc = parser.parseFromString(html, 'text/html');
            const title = extractTitleFromDoc(doc);
            if (title) opt.textContent = title;
          }
        } catch (e) {
          console.warn('Label sync failed for', opt.value, e);
        }
      }
    }

    function mapPageToFile(page) {
      switch (page) {
        case '1': return 'teteSuperieure.html';
        default: return '';
      }
    }

    async function loadSelected(path) {
      if (!path) { targetEl.innerHTML = ''; statusEl.textContent = ''; return; }
      statusEl.textContent = 'Chargement...';
      // Hidden fallback mapping to existing in-repo files (UI unchanged)
      const fallbackMap = {
        'teteSuperieure.html': 'TypeNews/edito/articleedito.html'
      };
      const candidates = [path, fallbackMap[path]].filter(Boolean);
      try {
        // Try primary, then mapped fallback if needed; remember which path succeeded
        let usedPath = candidates[0];
        let res = await fetch(`${candidates[0]}?t=${Date.now()}`);
        if (!res.ok && candidates[1]) {
          const res2 = await fetch(`${candidates[1]}?t=${Date.now()}`);
          if (res2.ok) { res = res2; usedPath = candidates[1]; }
        }
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const html = await res.text();
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const content = doc.querySelector('.newsletter-content') || doc.body;
        const wrapper = document.createElement('div');
        wrapper.innerHTML = content ? content.innerHTML : html;

        // Normalization:
        // - Known synced files: map bare names to Image/ or media/
        // - Other files: resolve bare names relative to the fetched file's directory (GitHub-like behavior)
        try {
          const normalizeFor = ['teteSuperieure.html'];
          if (normalizeFor.includes(path)) {
            normalizeAssetsInWrapper(wrapper);
          } else {
            const baseDir = (usedPath && usedPath.includes('/')) ? usedPath.substring(0, usedPath.lastIndexOf('/')) : '';
            normalizeAssetsRelative(wrapper, baseDir);
          }
        } catch (_) { /* ignore */ }
        wrapper.querySelectorAll('[contenteditable]').forEach(el => {
          el.removeAttribute('contenteditable');
          el.contentEditable = 'false';
        });
        ['input','textarea','select','button'].forEach(sel => {
          wrapper.querySelectorAll(sel).forEach(ctrl => {
            ctrl.setAttribute('disabled', '');
            ctrl.setAttribute('tabindex', '-1');
          });
        });
        targetEl.innerHTML = '';
        while (wrapper.firstChild) targetEl.appendChild(wrapper.firstChild);
        // If nothing rendered, fall back to autosaved editor content
        if (!targetEl.innerHTML || targetEl.innerHTML.trim() === '') {
          try {
            let saved = localStorage.getItem('currentNewsletterContent');
            if (!saved) saved = sessionStorage.getItem('currentNewsletterContentSession');
            if (saved) {
              targetEl.innerHTML = saved;
            }
          } catch (_) { /* ignore */ }
        }
        statusEl.textContent = '';
      } catch (e) {
        console.error('Load failed:', e);
        // Fallback to autosaved editor content so preview isn't empty
        try {
          let saved = localStorage.getItem('currentNewsletterContent');
          if (!saved) saved = sessionStorage.getItem('currentNewsletterContentSession');
          if (saved) {
            targetEl.innerHTML = saved;
            statusEl.textContent = '';
            return;
          }
        } catch (_) { /* ignore */ }
        statusEl.textContent = 'Erreur de chargement';
      }
    }

    // Bootstrap
    syncDropdownLabels().then(() => {
      if (selectEl) {
        const page = getUrlParameter('page');
        const contentFile = mapPageToFile(page);
        if (contentFile) {
          selectEl.value = contentFile;
          loadSelected(contentFile);
        }
      }
    });

    if (selectEl) {
      selectEl.addEventListener('change', (e) => loadSelected(e.target.value));
    }
  });
})();

// 5) Dynamic sync (tiles + hero) when not in preview mode
(function initDynamicSync() {
  document.addEventListener('DOMContentLoaded', function () {
    try {
      const isPreview = new URLSearchParams(window.location.search).has('historyId');
      if (isPreview) return;
      // When opened via file://, browser fetch of local files often fails with TypeError: Failed to fetch.
      // To avoid noisy console errors, disable dynamic sync under file protocol.
      if (window.location && window.location.protocol === 'file:') return;
    } catch (e) {}

    const contentConfig = [
      { source: 'teteSuperieure.html', titleSelector: 'h1, h2, h3', imageSelector: 'img', targetTitle: '.title-accent', targetImage: '.newsletter-hero img', errorMsg: 'teteSuperieure (Main Title)' },
      { source: 'contenuDeGauche.html', titleSelector: 'h2, h3', imageSelector: 'img', targetTitle: '.cards-grid.kyo-tiles .kyo-tile:nth-child(1) h3', targetImage: '.cards-grid.kyo-tiles .kyo-tile:nth-child(1) img', errorMsg: 'contenuDeGauche (First Tile)' },
      { source: 'contenuCentral.html', titleSelector: 'h2, h3', imageSelector: 'img', targetTitle: '.cards-grid.kyo-tiles .kyo-tile:nth-child(2) h3', targetImage: '.cards-grid.kyo-tiles .kyo-tile:nth-child(2) img', errorMsg: 'contenuCentral (Second Tile)' },
      { source: 'contenuDeDroite.html', titleSelector: 'h2, h3', imageSelector: 'img', targetTitle: '.cards-grid.kyo-tiles .kyo-tile:nth-child(3) h3', targetImage: '.cards-grid.kyo-tiles .kyo-tile:nth-child(3) img', errorMsg: 'contenuDeDroite (Third Tile)' }
    ];

    async function updateContent(config) {
      try {
        const response = await fetch(`${config.source}?t=${Date.now()}`);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        const html = await response.text();
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const titleElement = doc.querySelector(config.titleSelector);
        const targetTitle = document.querySelector(config.targetTitle);
        if (titleElement && targetTitle) {
          const text = (titleElement.innerText || titleElement.textContent || '').trim();
          if (text) targetTitle.textContent = text;
        }
        const imageElement = doc.querySelector(config.imageSelector);
        const targetImage = document.querySelector(config.targetImage);
        if (imageElement && targetImage) {
          const imageSrc = imageElement.src || imageElement.getAttribute('src');
          if (imageSrc) {
            targetImage.src = imageSrc;
            targetImage.alt = imageElement.alt || '';
          }
        }
        return true;
      } catch (error) {
        console.error(`Error syncing ${config.errorMsg}:`, error);
        return false;
      }
    }

    async function syncContentWithRetry(config, retries = 3, delay = 1000) {
      for (let i = 0; i < retries; i++) {
        const ok = await updateContent(config);
        if (ok) return true;
        if (i < retries - 1) await new Promise(r => setTimeout(r, delay));
      }
      return false;
    }

    function initRealTimeSync() {
      contentConfig.forEach(config => { syncContentWithRetry(config); });
      const syncInterval = setInterval(() => { contentConfig.forEach(config => { syncContentWithRetry(config); }); }, 5000);
      window.addEventListener('beforeunload', () => { clearInterval(syncInterval); });
    }

    initRealTimeSync();
  });
})();

