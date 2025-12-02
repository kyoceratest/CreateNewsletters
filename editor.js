class NewsletterEditor {
    constructor() {
        this.history = [];
        this.currentHistoryIndex = -1;
        this.maxHistorySize = 50;
        this.currentEditingImage = null;
        this.currentImageWrapper = null;
        this.cropOverlay = null;
        this.cropEventHandlers = null;
        this.resizeEventHandlers = null;
        this.imageToolbarInitialized = false;
        this.originalImageWidth = 0;
        this.originalImageHeight = 0;
        this.originalImageSrc = null;
        this.currentTargetContainer = null;
        this.savedSelection = null;
        this.currentEditingSection = null;
        this.currentEditingVideo = null;
        this.currentEditingTable = null;
        this.lastMousePosition = { x: 0, y: 0 };
        this.lastHoveredTableCell = null;
        this.lastClickedTableCell = null;
        this.tableContextCell = null;
        this._selectedRowForAction = null;
        this._inputSaveTimer = null;
        // Increment this when autosave schema/behavior changes to avoid restoring stale content
        this.storageVersion = '2';
        // Track last user action for history subtitle
        this.lastAction = 'Contenu modifié';
        // Initialize editor (wire buttons, toolbars, handlers)
        this.init();
    }

    init() {
        try { console.log('NewsletterEditor initializing...'); } catch (_) {}
        try { this.setupEventListeners(); } catch (_) {}
        try {
            const legacy = document.getElementById('tableContextMenu');
            if (legacy && legacy.parentNode) legacy.parentNode.removeChild(legacy);
        } catch (_) {}
        try { this.restoreLastFromLocalStorage(); } catch (_) {}
        try { this.saveState(); } catch (_) {}
        try { this.updateLastModified && this.updateLastModified(); } catch (_) {}
        try { console.log('NewsletterEditor initialized successfully'); } catch (_) {}
    }

    // Keep tables performant for typing: fixed layout, equal column widths, wrap long words
    normalizeTableForTyping(table) {
        if (!table) return;
        try {
            table.style.tableLayout = 'fixed';
            table.style.width = table.style.width || '100%';
            // Ensure an even-width colgroup exists
            const rows = Array.from(table.rows || []);
            const cols = rows[0] ? rows[0].cells.length : 0;
            if (cols > 0) {
                let cg = table.querySelector('colgroup[data-equalized="1"]');
                if (!cg) {
                    // Remove non-managed colgroups to avoid conflicts
                    const existing = table.querySelectorAll('colgroup');
                    existing.forEach(g => g.parentNode && g.parentNode.removeChild(g));
                    cg = document.createElement('colgroup');
                    cg.setAttribute('data-equalized', '1');
                    for (let i = 0; i < cols; i++) {
                        const col = document.createElement('col');
                        col.style.width = (100 / cols) + '%';
                        cg.appendChild(col);
                    }
                    table.insertBefore(cg, table.firstChild);
                }
            }
            // Make all cells wrap safely to avoid width growth
            table.querySelectorAll('th, td').forEach(cell => {
                cell.style.wordBreak = 'break-word';
                cell.style.overflowWrap = 'anywhere';
                cell.style.whiteSpace = 'normal';
                cell.contentEditable = 'true';
            });
        } catch (_) { /* no-op */ }
    }

    // Allow dragging an image wrapper up/down to reposition in text flow (not absolute)
    makeImageFlowReorderable(wrapper) {
        if (!wrapper) return;
        // Avoid double-binding
        if (wrapper.__flowDragBound) return;
        wrapper.__flowDragBound = true;

        const isAbsolute = () => wrapper.classList.contains('position-absolute') || wrapper.style.position === 'absolute';

        let dragging = false;
        let down = false;
        let startX = 0, startY = 0;
        const DRAG_THRESHOLD = 5; // px before we treat as a drag
        let indicator = null;
        const editableEl = document.getElementById('editableContent');

        const cleanup = () => {
            dragging = false;
            document.removeEventListener('mousemove', onMove, true);
            document.removeEventListener('mouseup', onUp, true);
            if (indicator && indicator.parentNode) indicator.parentNode.removeChild(indicator);
            indicator = null;
            try { wrapper.classList.remove('dragging'); } catch (_) {}
        };

        const ensureIndicator = () => {
            if (indicator) return indicator;
            indicator = document.createElement('div');
            indicator.style.position = 'fixed';
            indicator.style.height = '6px';
            indicator.style.background = 'linear-gradient(90deg, rgba(10,155,205,0.9), rgba(10,155,205,0.6))';
            indicator.style.boxShadow = '0 0 0 1px rgba(10,155,205,0.35), 0 2px 6px rgba(10,155,205,0.25)';
            indicator.style.borderRadius = '3px';
            indicator.style.zIndex = '99999';
            indicator.style.pointerEvents = 'none';
            document.body.appendChild(indicator);
            return indicator;
        };

        const onMove = (e) => {
            if (!editableEl) return;
            if (!dragging) {
                if (!down) return;
                const dx = Math.abs(e.clientX - startX);
                const dy = Math.abs(e.clientY - startY);
                if (dx < DRAG_THRESHOLD && dy < DRAG_THRESHOLD) return;
                // Begin dragging now
                dragging = true;
                try { wrapper.classList.add('dragging'); } catch (_) {}
            }
            // Find block boundary at pointer
            let r = null;
            try { r = this.computeRangeFromPoint(e.clientX, e.clientY); } catch (_) {}
            if (!r) return;
            const containerRect = editableEl.getBoundingClientRect();
            // Determine target block's top or bottom edge
            let block = r.startContainer && (r.startContainer.nodeType === Node.ELEMENT_NODE ? r.startContainer : r.startContainer.parentElement);
            while (block && block.parentElement && block.parentElement !== editableEl) {
                block = block.parentElement;
            }
            if (!block || block === editableEl) return;
            const br = block.getBoundingClientRect();
            const after = e.clientY > (br.top + br.height / 2);
            const y = after ? br.bottom : br.top;
            const ind = ensureIndicator();
            ind.style.left = (containerRect.left + 8) + 'px';
            ind.style.width = Math.max(0, containerRect.width - 16) + 'px';
            ind.style.top = (y - 3) + 'px'; // center the 6px bar on the boundary
        };

        const onUp = (e) => {
            if (!down) { cleanup(); return; }
            if (!dragging) { cleanup(); return; }
            // Compute final range and move wrapper there
            let r = null;
            try { r = this.computeRangeFromPoint(e.clientX, e.clientY); } catch (_) {}
            if (r) {
                try {
                    r.collapse(true);
                    r.insertNode(wrapper);
                    // Place caret after moved wrapper
                    const sel = window.getSelection();
                    const spacer = document.createElement('p');
                    spacer.innerHTML = '<br>';
                    wrapper.after(spacer);
                    const newRange = document.createRange();
                    newRange.selectNodeContents(spacer);
                    newRange.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(newRange);
                    this.saveState();
                    this.lastAction = 'Image déplacée';
                } catch (_) {}
            }
            cleanup();
        };

        wrapper.addEventListener('mousedown', (e) => {
            // Ignore when absolute drag/resize/rotate handles are engaged
            if (isAbsolute()) return; // free-position images use absolute drag handlers
            const t = e.target;
            if (t && (t.classList && (t.classList.contains('resize-handle') || t.classList.contains('rotation-handle')))) return;
            // Start flow drag
            down = true;
            startX = e.clientX;
            startY = e.clientY;
            document.addEventListener('mousemove', onMove, true);
            document.addEventListener('mouseup', onUp, true);
        });

        // Show grab cursor when this wrapper can be flow-dragged
        const updateCursor = () => {
            try {
                if (!isAbsolute()) {
                    wrapper.style.cursor = 'grab';
                } else {
                    wrapper.style.cursor = '';
                }
            } catch (_) {}
        };
        updateCursor();
        wrapper.addEventListener('mouseenter', updateCursor);
        wrapper.addEventListener('transitionend', updateCursor);
    }

    // Remove inline font and color styling from pasted HTML to enforce defaults
    sanitizePastedHtml(rawHtml) {
        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(rawHtml, 'text/html');
            const walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_ELEMENT, null, false);
            const toRemove = [];
            const stripStyleProps = (el) => {
                try {
                    if (el.style) {
                        el.style.removeProperty('font-family');
                        el.style.removeProperty('font');
                        el.style.removeProperty('font-size');
                        el.style.removeProperty('color');
                        el.style.removeProperty('background');
                        el.style.removeProperty('background-color');
                        el.style.removeProperty('text-decoration-color');
                    }
                    // Also remove presentational attributes
                    if (el.hasAttribute && el.hasAttribute('color')) el.removeAttribute('color');
                    if (el.hasAttribute && el.hasAttribute('face')) el.removeAttribute('face');
                    if (el.hasAttribute && el.hasAttribute('size')) el.removeAttribute('size');
                } catch (_) {}
            };
            // Convert <font> to <span> and strip attributes
            doc.body.querySelectorAll('font').forEach(f => {
                const span = doc.createElement('span');
                span.innerHTML = f.innerHTML;
                f.parentNode && f.parentNode.replaceChild(span, f);
            });
            // Walk and strip unwanted inline styles
            while (walker.nextNode()) {
                const el = walker.currentNode;
                stripStyleProps(el);
            }
            return doc.body.innerHTML;
        } catch (e) {
            return rawHtml; // fallback
        }
    }

    // Normalize any blob-based video sources to media/<filename> on page load (so refresh keeps videos playable)
    static normalizeLocalVideoSourcesOnLoad() {
        try {
            const host = document.getElementById('editableContent') || document;
            // For <video src> directly
            Array.from(host.querySelectorAll('video[src]')).forEach(v => {
                const raw = v.getAttribute('src') || '';
                if (/^blob:/i.test(raw)) {
                    const hint = v.getAttribute('data-local-filename') || '';
                    if (hint) {
                        const safe = hint.replace(/^[\\\/]+/, '');
                        v.setAttribute('src', 'media/' + encodeURIComponent(safe));
                    }
                }
            });
            // For <source src> inside media
            Array.from(host.querySelectorAll('video source[src], audio source[src]')).forEach(s => {
                const raw = s.getAttribute('src') || '';
                if (/^blob:/i.test(raw)) {
                    let hint = s.getAttribute('data-local-filename') || '';
                    if (!hint) {
                        try { const parent = s.closest('video,audio'); hint = parent ? (parent.getAttribute('data-local-filename') || '') : ''; } catch (_) {}
                    }
                    if (hint) {
                        const safe = hint.replace(/^[\\\/]+/, '');
                        s.setAttribute('src', 'media/' + encodeURIComponent(safe));
                    }
                }
            });
        } catch (_) { /* ignore */ }
    }

    // Open a sanitized, read-only preview in a new tab without altering the editor UI
    preview() {
        try {
            this.previewFull();
        } catch (e) {
            console.error('Preview failed:', e);
        }
    }
    // Full-page preview reusing preview.html (header/footer, category bar, styles)
    previewFull() {
        try {
            const editableContent = document.getElementById('editableContent');
            if (!editableContent) return;

            // Clone current content
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = editableContent.innerHTML;

            // Sanitize editor-only controls
            try {
                const removeSelectors = [
                    '.add-image-placeholder',
                    '.add-more-btn',
                    '.image-toolbar',
                    '.resize-handle',
                    '.rotation-handle',
                    '.gallery-upload',
                    '.gallery-controls',
                    '.gallery-editor-only',
                    '.gallery-actions',
                    'input[type="file"]'
                ];
                removeSelectors.forEach(sel => tempDiv.querySelectorAll(sel).forEach(el => el.remove()));

                const buttonTextsToRemove = [
                    'Ajouter une image',
                    'Modifier la galerie',
                    "Ajouter plus d'images"
                ];
                tempDiv.querySelectorAll('button, a').forEach(el => {
                    const txt = (el.innerText || el.textContent || '').trim().toLowerCase();
                    if (buttonTextsToRemove.some(t => txt.includes(t.toLowerCase()))) {
                        el.remove();
                    }
                });

                tempDiv.querySelectorAll('[class*="placeholder"]').forEach(el => el.remove());
                tempDiv.querySelectorAll('[data-editor-only], [data-placeholder], [data-editable-control]').forEach(el => el.remove());
                tempDiv.querySelectorAll('[contenteditable]').forEach(el => el.removeAttribute('contenteditable'));
            } catch (_) { /* best effort */ }

            const pageTitle = (document.querySelector('h1, h2, h3')?.textContent || document.title || 'Aperçu').trim();

            const previewWindow = window.open('preview.html', 'previewWindowFromEditor');
            if (!previewWindow) return;

            const payload = {
                type: 'previewContent',
                content: tempDiv.innerHTML,
                title: pageTitle
            };

            const interval = setInterval(() => {
                try {
                    if (!previewWindow || previewWindow.closed) {
                        clearInterval(interval);
                        return;
                    }
                    if (previewWindow.document && previewWindow.document.readyState === 'complete') {
                        clearInterval(interval);
                        previewWindow.postMessage(payload, '*');
                    }
                } catch (_) {
                    // Cross-origin or timing issue; keep trying until window is ready or closed
                }
            }, 100);
        } catch (e) {
            console.error('Preview (full) failed:', e);
        }
    }

    // Safely delete an image wrapper while preserving adjacent text content
    deleteImageWrapperSafe(wrapper) {
        if (!wrapper || !wrapper.parentNode) return;
        const parent = wrapper.parentNode;
        const prev = wrapper.previousSibling;
        const next = wrapper.nextSibling;

    // If two text nodes would become adjacent without a space, insert one
    let spacer = null;
    const prevText = prev && prev.nodeType === Node.TEXT_NODE ? prev.nodeValue : null;
    const nextText = next && next.nodeType === Node.TEXT_NODE ? next.nodeValue : null;
    const needSpace = !!(prevText !== null && nextText !== null && !/\s$/.test(prevText) && !/^\s/.test(nextText));
    if (needSpace) {
        spacer = document.createTextNode(' ');
        parent.insertBefore(spacer, wrapper);
    }

    // Remove wrapper (image + any handles)
    parent.removeChild(wrapper);

    try {
        const columnEl = parent && parent.closest ? parent.closest('.column') : null;
        const inTwoColumn = columnEl && columnEl.closest && !!columnEl.closest('.two-column-layout');
        if (inTwoColumn) {
            const hasImg = columnEl.querySelector('img');
            const hasPlaceholder = columnEl.querySelector('.image-placeholder');
            if (!hasImg && !hasPlaceholder) {
                const placeholder = document.createElement('div');
                placeholder.className = 'image-placeholder';
                placeholder.setAttribute('contenteditable', 'false');
                placeholder.style.border = '2px dashed #ccc';
                placeholder.style.padding = '40px';
                placeholder.style.textAlign = 'center';
                placeholder.style.cursor = 'pointer';
                placeholder.style.borderRadius = '8px';
                placeholder.innerHTML = '<i class="fas fa-image" style="font-size: 24px; color: #6c757d; margin-bottom: 10px;"></i>\n<p style="color: #6c757d; margin: 0;">Cliquez pour ajouter une image</p>';
                placeholder.addEventListener('click', () => {
                    const fileInput = document.createElement('input');
                    fileInput.type = 'file';
                    fileInput.accept = 'image/*';
                    fileInput.style.display = 'none';
                    
                    fileInput.onchange = (e) => {
                        const file = e.target.files && e.target.files[0];
                        if (!file) return;
                        const reader = new FileReader();
                        reader.onload = (event) => {
                            const img = document.createElement('img');
                            img.src = event.target.result;
                            img.alt = 'Image';
                            img.setAttribute('contenteditable', 'false');
                            img.style.cssText = 'width:100%;height:auto;border-radius:8px;display:block;';
                            img.addEventListener('click', (ev) => {
                                ev.stopPropagation();
                                try { this.selectImage(img); } catch (_) {}
                            });
                            try {
                                let sib = placeholder.nextElementSibling;
                                while (sib) {
                                    const nextSib = sib.nextElementSibling;
                                    sib.remove();
                                    sib = nextSib;
                                }
                            } catch (_) {}
                            placeholder.replaceWith(img);
                            try {
                                const col = img.closest('.column');
                                if (col) {
                                    const children = Array.from(col.children);
                                    let seen = false;
                                    for (const child of children) {
                                        if (child === img || (child.querySelector && child.querySelector('img') === img)) { seen = true; continue; }
                                        if (seen) { try { child.remove(); } catch (_) {} }
                                    }
                                }
                            } catch (_) {}
                            this.saveState();
                            this.lastAction = 'Image ajoutée';
                        };
                        reader.readAsDataURL(file);
                    };
                    fileInput.click();
                });
                columnEl.appendChild(placeholder);
            }
        }
    } catch (_) {}

    // Place caret after spacer or before next sibling to avoid deleting nearby text
    try {
        const sel = window.getSelection();
        if (sel) {
            const range = document.createRange();
            if (spacer && spacer.parentNode) {
                range.setStartAfter(spacer);
            } else if (next && next.parentNode === parent) {
                range.setStart(parent, Array.prototype.indexOf.call(parent.childNodes, next));
            } else if (prev && prev.parentNode === parent) {
                // After previous node
                const idx = Array.prototype.indexOf.call(parent.childNodes, prev);
                range.setStart(parent, idx + 1);
            } else {
                range.selectNodeContents(parent);
                range.collapse(false);
            }
            range.collapse(true);
            sel.removeAllRanges();
            sel.addRange(range);
        }
    } catch (_) { /* no-op */ }

    // Persist editor state
    try {
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = "Image supprimée";
    } catch (_) { /* no-op */ }
}

    // Restore the most recent saved content from localStorage history into the editor
    restoreLastFromLocalStorage() {
        try {
            const editable = document.getElementById('editableContent');
            if (!editable) return;
            let history = [];
            try {
                const raw = localStorage.getItem('newsletterHistory');
                history = raw ? JSON.parse(raw) : [];
                if (!Array.isArray(history)) history = [];
            } catch (_) { history = []; }
            if (history.length === 0) return;
            const latest = history[0];
            const slimHtml = (latest && typeof latest.content === 'string') ? latest.content : '';
            const id = latest && latest.id ? String(latest.id) : '';
            // Prefer the full HTML from IndexedDB (keeps <img src> intact); fallback to slim localStorage copy
            if (id && typeof getFullContentFromIDB === 'function') {
                try {
                    getFullContentFromIDB(id).then((full) => {
                        const html = (full && typeof full === 'string' && full.length > 0) ? full : slimHtml;
                        if (html) {
                            editable.innerHTML = html;
                            try { this.saveState(); this.updateLastModified(); this.lastAction = 'Restauration automatique'; } catch (_) {}
                        }
                    }).catch(() => {
                        if (slimHtml) {
                            editable.innerHTML = slimHtml;
                            try { this.saveState(); this.updateLastModified(); this.lastAction = 'Restauration automatique'; } catch (_) {}
                        }
                    });
                    return; // async path handles update
                } catch (_) { /* fall through to slim */ }
            }
            if (slimHtml) {
                editable.innerHTML = slimHtml;
                try { this.saveState(); this.updateLastModified(); this.lastAction = 'Restauration automatique'; } catch (_) {}
            }
        } catch (_) { /* ignore */ }
    }

    setupEventListeners() {
        console.log('Setting up event listeners...');
        
        // Track caret/range within the editor so insertions follow the user's pointer/caret
        const editableHost = document.getElementById('editableContent');
        if (editableHost) {
            // Sanitize pasted content and force default paste size to 26px
            editableHost.addEventListener('paste', (e) => {
                try {
                    const cd = e.clipboardData || window.clipboardData;
                    if (!cd) return; // let default happen
                    const html = cd.getData('text/html');
                    const text = cd.getData('text/plain');
                    if (!html && !text) return;
                    e.preventDefault();
                    if (html) {
                        const clean = this.sanitizePastedHtml(html);
                        const wrapped = `<span style="font-size:26px;">${clean}</span>`;
                        document.execCommand('insertHTML', false, wrapped);
                    } else {
                        const safe = (text || '').replace(/[&<>\n]/g, (ch) => ({'&':'&amp;','<':'&lt;','>':'&gt;','\n':'<br>'}[ch]));
                        const wrapped = `<span style="font-size:26px;">${safe}</span>`;
                        document.execCommand('insertHTML', false, wrapped);
                    }
                    try { this.saveState(); } catch (_) {}
                } catch (_) { /* fall back to default paste */ }
            });
            const updateRangeFromSelection = () => {
                const sel = window.getSelection();
                if (sel && sel.rangeCount > 0) {
                    const r = sel.getRangeAt(0);
                    // Only persist ranges inside the editor
                    const node = r.startContainer;
                    const host = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
                    if (host && host.closest && host.closest('#editableContent')) {
                        this.lastMouseRange = r.cloneRange();
                    }
                }
            };
            // Update on mouse/key interactions
            editableHost.addEventListener('mouseup', updateRangeFromSelection);
            editableHost.addEventListener('keyup', updateRangeFromSelection);
            editableHost.addEventListener('click', updateRangeFromSelection);

            // Event delegation to re-enable image selection after restore/history load
            editableHost.addEventListener('click', (e) => {
                // Case 1: wrapped image
                const wrapper = e.target && (e.target.closest && e.target.closest('.image-wrapper'));
                if (wrapper) {
                    const img = wrapper.querySelector('img') || wrapper;
                    try { this.selectImage(img); } catch (_) {}
                    return;
                }
                // Case 2: plain <img> without wrapper (e.g., from history or paste)
                const el = e.target;
                if (el && el.tagName === 'IMG') {
                    try { this.selectImage(el); } catch (_) {}
                }
            });
        }

        // Column image placeholder click handler - using gallery section logic
        const columnImagePlaceholder = document.getElementById('columnImagePlaceholder');
        if (columnImagePlaceholder) {
            console.log('Column image placeholder found, adding click handler');
            columnImagePlaceholder.addEventListener('click', (event) => {
                // Prevent parent click handlers from interpreting this as a section click
                try { event.stopPropagation(); } catch (_) {}
                console.log('Column image placeholder clicked');
                // Create hidden file input like gallery section
                const fileInput = document.createElement('input');
                fileInput.type = 'file';
                fileInput.accept = 'image/*';
                fileInput.style.display = 'none';
                
                // Handle file selection like gallery
                fileInput.addEventListener('change', (e) => {
                    const file = e.target.files[0];
                    if (file && file.type.startsWith('image/')) {
                        this.addImageToColumn(file, columnImagePlaceholder);
                    }
                    fileInput.value = '';
                });
                
                fileInput.click();
            });
        } else {
            // Optional placeholder; not an error if missing on this page
            console.debug('columnImagePlaceholder not present on this page');
        }
        
        // Insert Image button
        const insertImageBtn = document.getElementById('insertImageBtn');
        if (insertImageBtn) {
            insertImageBtn.addEventListener('click', () => {
                const options = document.getElementById('imageOptions');
                options.style.display = options.style.display === 'none' ? 'block' : 'none';
            });
        } else {
            console.error('insertImageBtn not found');
        }

        // New unified resize menu
        const resizeMainBtn = document.getElementById('resizeMainBtn');
        if (resizeMainBtn) {
            resizeMainBtn.addEventListener('click', () => {
                const opts = document.getElementById('resizeOptions');
                if (opts) opts.style.display = opts.style.display === 'none' ? 'block' : 'none';
            });
        }

        // Resize image (server-side) button
        const resizeImageBtn = document.getElementById('resizeImageBtn');
        if (resizeImageBtn) {
            resizeImageBtn.addEventListener('click', () => {
                try {
                    const input = document.createElement('input');
                    input.type = 'file';
                    input.accept = 'image/*';
                    input.onchange = async (e) => {
                        const file = e.target.files && e.target.files[0];
                        if (!file) return;
                        const desired = prompt('Nom du fichier (sans extension):');
                        if (desired === null) return;
                        const fd = new FormData();
                        fd.append('image', file);
                        fd.append('filename', desired);
                        try {
                            const resp = await fetch('upload_resize.php', { method: 'POST', body: fd });
                            const data = await resp.json();
                            if (data && data.success) {
                                alert('Image enregistrée: ' + data.url);
                            } else {
                                alert('Erreur: ' + (data && data.error ? data.error : 'inconnue'));
                            }
                        } catch (err) {
                            alert('Erreur réseau lors du redimensionnement. Assurez-vous d\'ouvrir la page via un serveur PHP (ex: http://localhost/...) et non via un serveur statique.');
                        }
                    };
                    input.click();
                } catch (_) {}
            });
        }

        // Batch resize (server-side) button
        const batchResizeBtn = document.getElementById('batchResizeBtn');
        if (batchResizeBtn) {
            batchResizeBtn.addEventListener('click', async () => {
                // Prefer local folder processing when available
                if (window.showDirectoryPicker && confirm('Utiliser le dossier local (sans serveur) pour redimensionner ?')) {
                    try {
                        await localBatchResizeFolders();
                    } catch (err) {
                        alert('Erreur lors du traitement local: ' + (err && err.message ? err.message : String(err)));
                    }
                    return;
                }

                // Fallback to server endpoint
                if (!confirm('Redimensionner via le serveur toutes les images du dossier Image/ vers subImage/?')) return;
                try {
                    const resp = await fetch('batch_resize.php', { method: 'POST' });
                    const data = await resp.json();
                    if (data && data.success) {
                        const msg = `Traitement terminé\nTotal: ${data.total}\nRedimensionnées: ${data.resized}\nIgnorées: ${data.skipped}\nErreurs: ${data.errors?.length || 0}`;
                        alert(msg);
                    } else {
                        alert('Erreur: ' + (data && data.error ? data.error : 'inconnue'));
                    }
                } catch (e) {
                    alert('Erreur réseau. Ouvrez la page via un serveur PHP (ex: http://localhost/...).');
                }
            });
        }
        const resizeFolderUnderInsertBtn = document.getElementById('resizeFolderUnderInsertBtn');
        if (resizeFolderUnderInsertBtn) {
            resizeFolderUnderInsertBtn.addEventListener('click', async () => {
                try {
                    if (window.showDirectoryPicker && confirm('Utiliser le dossier local (sans serveur) pour redimensionner ?')) {
                        try { await localBatchResizeFolders(); } catch (err) {
                            alert('Erreur lors du traitement local: ' + (err && err.message ? err.message : String(err)));
                        }
                        try { document.getElementById('imageOptions').style.display = 'none'; } catch (_) {}
                        return;
                    }
                    if (!confirm('Redimensionner via le serveur toutes les images du dossier Image/ vers subImage/?')) return;
                    try {
                        const resp = await fetch('batch_resize.php', { method: 'POST' });
                        const data = await resp.json();
                        if (data && data.success) {
                            const msg = `Traitement terminé\nTotal: ${data.total}\nRedimensionnées: ${data.resized}\nIgnorées: ${data.skipped}\nErreurs: ${data.errors?.length || 0}`;
                            alert(msg);
                        } else {
                            alert('Erreur: ' + (data && data.error ? data.error : 'inconnue'));
                        }
                    } catch (e) {
                        alert('Erreur réseau. Ouvrez la page via un serveur PHP (ex: http://localhost/...).');
                    }
                } finally {
                    try { document.getElementById('imageOptions').style.display = 'none'; } catch (_) {}
                }
            });
        }

        // Local folder batch resize using File System Access API
        async function localBatchResizeFolders() {
            if (!window.showDirectoryPicker) throw new Error('File System Access API non supportée par ce navigateur');

            const srcHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
            try { await srcHandle.requestPermission({ mode: 'readwrite' }); } catch (_) {}
            const dstHandle = srcHandle;

            const maxW = 1600, maxH = 1600;
            let total = 0, resized = 0, skipped = 0, errors = 0;
            let origBytes = 0, newBytes = 0;
            const reportRows = [['path','original_bytes','new_bytes','saved_bytes','saved_percent']];

            async function ensureSubdir(baseHandle, pathParts) {
                let cur = baseHandle;
                for (const part of pathParts) {
                    try {
                        cur = await cur.getDirectoryHandle(part, { create: true });
                    } catch (_) {
                        cur = await cur.getDirectoryHandle(part, { create: true });
                    }
                }
                return cur;
            }

            async function processFile(fileHandle, relParts) {
                const file = await fileHandle.getFile();
                if (!file.type || !file.type.startsWith('image/')) return; // ignore non-images
                total++;
                try {
                    const { blob: outBlob, changed, name } = await resizeImageFile(file, maxW, maxH);
                    const parent = await ensureSubdir(dstHandle, relParts);
                    const outHandle = await parent.getFileHandle(name, { create: true });
                    const writable = await outHandle.createWritable();
                    await writable.write(outBlob);
                    await writable.close();
                    const o = file.size || 0;
                    const n = outBlob.size || 0;
                    origBytes += o;
                    newBytes += n;
                    const saved = Math.max(0, o - n);
                    const pct = o > 0 ? Math.round((saved / o) * 100) : 0;
                    const relPath = (relParts.length ? relParts.join('/') + '/' : '') + name;
                    reportRows.push([relPath, String(o), String(n), String(saved), String(pct)]);
                    if (changed) resized++; else skipped++;
                } catch (_) {
                    errors++;
                }
            }

            async function walk(dirHandle, relParts = []) {
                for await (const entry of dirHandle.values()) {
                    if (entry.kind === 'directory') {
                        await walk(entry, relParts.concat([entry.name]));
                    } else if (entry.kind === 'file') {
                        await processFile(entry, relParts);
                    }
                }
            }

            await walk(srcHandle, []);

            // Write CSV report at destination root
            try {
                const csv = reportRows.map(r => r.map(v => /[",\n]/.test(v) ? '"' + v.replace(/"/g,'""') + '"' : v).join(',')).join('\n');
                const reportHandle = await dstHandle.getFileHandle('resize_report.csv', { create: true });
                const writable = await reportHandle.createWritable();
                await writable.write(new Blob([csv], { type: 'text/csv' }));
                await writable.close();
            } catch (_) {}

            const savedBytes = Math.max(0, origBytes - newBytes);
            const savedKB = Math.round(savedBytes / 102.4) / 10; // 1 decimal
            const newKB = Math.round(newBytes / 102.4) / 10;
            const origKB = Math.round(origBytes / 102.4) / 10;
            const pctTotal = origBytes > 0 ? Math.round((savedBytes / origBytes) * 100) : 0;
            alert(`Traitement local terminé\nTotal: ${total}\nRedimensionnées: ${resized}\nIgnorées: ${skipped}\nErreurs: ${errors}\n\nTaille totale avant: ${origKB} KB\nTaille totale après: ${newKB} KB\nGain: ${savedKB} KB (${pctTotal}%)\n\nRapport: resize_report.csv dans le dossier de destination`);
        }

        async function resizeImageFile(file, maxW, maxH) {
            const img = await createImageBitmap(file);
            let w = img.width, h = img.height;
            const ratio = Math.min(maxW / w, maxH / h, 1);
            const changed = ratio < 1;
            const targetW = Math.round(w * ratio);
            const targetH = Math.round(h * ratio);

            const canvas = document.createElement('canvas');
            canvas.width = targetW; canvas.height = targetH;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(img, 0, 0, targetW, targetH);

            const type = file.type || 'image/jpeg';
            const quality = type === 'image/jpeg' || type === 'image/webp' ? 0.82 : undefined;
            const blob = await new Promise((res) => canvas.toBlob((b) => res(b || file), type, quality));

            // Keep original name and extension
            const name = file.name;
            return { blob, changed, name };
        }

        // Local image button
        document.getElementById('localImageBtn').addEventListener('click', () => {
            // Persist current caret range so insertion follows the caret after file dialog
            try {
                const sel = window.getSelection();
                if (sel && sel.rangeCount > 0) this.lastMouseRange = sel.getRangeAt(0).cloneRange();
            } catch (_) {}
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'image/*';
            input.onchange = (e) => {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = (ev) => this.insertImage(ev.target.result, file.name);
                    reader.readAsDataURL(file);
                }
            };
            input.click();
            document.getElementById('imageOptions').style.display = 'none';
        });

        // URL image button
        document.getElementById('urlImageBtn').addEventListener('click', () => {
            // Persist current caret range so insertion follows the caret after prompt
            try {
                const sel = window.getSelection();
                if (sel && sel.rangeCount > 0) this.lastMouseRange = sel.getRangeAt(0).cloneRange();
            } catch (_) {}
            const url = prompt('Entrez l\'URL de l\'image:');
            if (url) {
                this.insertImage(url, 'Image URL');
            }
            document.getElementById('imageOptions').style.display = 'none';
        });

        // Insert Video button
        document.getElementById('insertVideoBtn').addEventListener('click', () => {
            const options = document.getElementById('videoOptions');
            options.style.display = options.style.display === 'none' ? 'block' : 'none';
        });

        // URL video button
        document.getElementById('urlVideoBtn').addEventListener('click', () => {
            // Persist current caret range so insertion follows the caret after prompt
            try {
                const sel = window.getSelection();
                if (sel && sel.rangeCount > 0) this.lastMouseRange = sel.getRangeAt(0).cloneRange();
            } catch (_) {}
            const url = prompt('Entrez l\'URL de la vidéo (YouTube, Vimeo, etc.):');
            if (url) {
                this.insertVideo(url);
            }
            document.getElementById('videoOptions').style.display = 'none';
        });

        // Local video button
        document.getElementById('localVideoBtn').addEventListener('click', () => {
            // Persist current caret range so insertion follows the caret after file dialog
            try {
                const sel = window.getSelection();
                if (sel && sel.rangeCount > 0) this.lastMouseRange = sel.getRangeAt(0).cloneRange();
            } catch (_) {}
            const input = document.createElement('input');
            input.type = 'file';
            input.accept = 'video/*';
            input.onchange = (e) => {
                const file = e.target.files[0];
                if (file) {
                    // Use a blob URL to avoid creating massive data URLs that can freeze the page
                    const objectUrl = URL.createObjectURL(file);
                    this.insertLocalVideo(objectUrl, file.name);
                }
            };
            input.click();
            document.getElementById('videoOptions').style.display = 'none';
        });

        // Standardize videos button (apply 70% centered sizing on demand)
        const standardizeBtn = document.getElementById('standardizeVideosBtn');
        if (standardizeBtn) {
            standardizeBtn.addEventListener('click', () => {
                try { this.normalizeVideoStyles(); } catch (_) {}
                const opts = document.getElementById('videoOptions');
                if (opts) opts.style.display = 'none';
            });
        }

        // Insert Table button
        const insertTableBtn = document.getElementById('insertTableBtn');
        if (insertTableBtn) {
            insertTableBtn.addEventListener('click', () => {
                return;
            });
        }

        // Insert Section button
        document.getElementById('insertSectionBtn').addEventListener('click', () => {
            const options = document.getElementById('sectionOptions');
            options.style.display = options.style.display === 'none' ? 'block' : 'none';
        });

        // Section options
        document.getElementById('articleSectionBtn').addEventListener('click', () => {
            this.insertArticleSection();
        });

        document.getElementById('gallerySectionBtn').addEventListener('click', () => {
            this.insertGallerySection();
        });

        const multiBtn = document.getElementById('multiImagesSectionBtn');
        if (multiBtn) {
            multiBtn.addEventListener('click', () => {
                this.insertMultiImagesSection();
            });
        }

        document.getElementById('quoteSectionBtn').addEventListener('click', () => {
            this.insertQuoteSection();
        });

        document.getElementById('ctaSectionBtn').addEventListener('click', () => {
            this.insertCTASection();
        });

        document.getElementById('contactSectionBtn').addEventListener('click', () => {
            this.insertContactSection();
        });

        document.getElementById('twoColumnSectionBtn').addEventListener('click', () => {
            this.insertTwoColumnSection();
        });

        // Import buttons (HTML / Excel)
        const importBtn = document.getElementById('importBtn');
        if (importBtn) {
            importBtn.addEventListener('click', () => {
                const options = document.getElementById('importOptions');
                if (options) options.style.display = options.style.display === 'none' ? 'block' : 'none';
            });
        }

        const importHtmlBtn = document.getElementById('importHtmlBtn');
        if (importHtmlBtn) {
            importHtmlBtn.addEventListener('click', () => {
                try {
                    const sel = window.getSelection();
                    if (sel && sel.rangeCount > 0) this.lastMouseRange = sel.getRangeAt(0).cloneRange();
                } catch (_) {}
                const input = document.createElement('input');
                input.type = 'file';
                input.accept = '.html,.htm,text/html';
                input.onchange = (e) => {
                    const file = e.target.files && e.target.files[0];
                    if (!file) return;
                    const reader = new FileReader();
                    reader.onload = () => {
                        try {
                            const htmlText = String(reader.result || '');
                            const parser = new DOMParser();
                            const doc = parser.parseFromString(htmlText, 'text/html');
                            const container = document.createElement('div');
                            container.className = 'imported-html';
                            const source = doc.body || doc;
                            // Minimal sanitization: remove scripts
                            source.querySelectorAll('script').forEach(n => n.remove());
                            // Move children into container
                            Array.from(source.childNodes).forEach(node => {
                                container.appendChild(node.cloneNode(true));
                            });
                            this.insertElementAtCursor(container);
                            this.saveState();
                            this.updateLastModified();
                            this.autoSaveToLocalStorage();
                            this.lastAction = 'Contenu HTML importé';
                        } catch (err) {
                            console.error('Import HTML failed:', err);
                            alert("Échec de l'import HTML");
                        }
                    };
                    reader.readAsText(file, 'utf-8');
                };
                input.click();
                const options = document.getElementById('importOptions');
                if (options) options.style.display = 'none';
            });
        }

        const importExcelBtn = document.getElementById('importExcelBtn');
        if (importExcelBtn) {
            importExcelBtn.addEventListener('click', () => {
                try {
                    const sel = window.getSelection();
                    if (sel && sel.rangeCount > 0) this.lastMouseRange = sel.getRangeAt(0).cloneRange();
                } catch (_) {}
                const input = document.createElement('input');
                input.type = 'file';
                input.accept = '.pdf,application/pdf';
                input.onchange = async (e) => {
                    const file = e.target.files && e.target.files[0];
                    if (!file) return;
                    try {
                        const blobUrl = URL.createObjectURL(file);
                        // Load PDF.js on demand if not already present
                        const ensurePdfJs = () => new Promise((resolve, reject) => {
                            if (window.pdfjsLib) return resolve();
                            const script = document.createElement('script');
                            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.10.111/pdf.min.js';
                            script.onload = () => {
                                try {
                                    if (window.pdfjsLib && window.pdfjsLib.GlobalWorkerOptions) {
                                        window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.10.111/pdf.worker.min.js';
                                    }
                                } catch (_) {}
                                resolve();
                            };
                            script.onerror = () => reject(new Error('PDF.js load failed'));
                            document.head.appendChild(script);
                        });
                        await ensurePdfJs();
                        const pdf = await window.pdfjsLib.getDocument({ url: blobUrl }).promise;
                        const page = await pdf.getPage(1);
                        const viewport = page.getViewport({ scale: 1.5 });
                        const canvas = document.createElement('canvas');
                        const ctx = canvas.getContext('2d');
                        canvas.width = viewport.width;
                        canvas.height = viewport.height;
                        await page.render({ canvasContext: ctx, viewport }).promise;
                        const dataUrl = canvas.toDataURL('image/png');
                        URL.revokeObjectURL(blobUrl);

                        const img = document.createElement('img');
                        img.src = dataUrl;
                        img.alt = file.name || 'PDF';
                        img.style.maxWidth = '100%';
                        img.style.height = 'auto';

                        const wrapper = document.createElement('div');
                        wrapper.className = 'imported-pdf-image';
                        wrapper.appendChild(img);

                        this.insertElementAtCursor(wrapper);
                        this.saveState();
                        this.updateLastModified();
                        this.autoSaveToLocalStorage();
                        this.lastAction = 'Image PDF importée';
                    } catch (err) {
                        console.error('Import PDF failed:', err);
                        alert("Échec de l'import PDF");
                    }
                };
                input.click();
                const options = document.getElementById('importOptions');
                if (options) options.style.display = 'none';
            });
        }

        // Action buttons
        document.getElementById('undoBtn').addEventListener('click', () => {
            this.undo();
        });

        document.getElementById('refreshBtn').addEventListener('click', () => {
            this.redo();
        });

        document.getElementById('clearBtn').addEventListener('click', () => {
            this.clear();
        });

        // Save button
        document.getElementById('saveBtn').addEventListener('click', () => {
            this.save();
        });

        // History button
        const historyBtn = document.getElementById('historyBtn');
        if (historyBtn) {
            if (window.__useInlineHistoryModal) {
                console.log('Using inline History modal on this page; editor.js will not wire its own history handler.');
            } else {
                console.log('History button found, adding event listener');
                historyBtn.addEventListener('click', () => {
                    console.log('History button clicked');
                    try {
                        this.showHistory();
                    } catch (error) {
                        console.error('Error in showHistory:', error);
                        alert('Erreur lors de l\'affichage de l\'historique: ' + error.message);
                    }
                });
            }
        } else {
            console.error('historyBtn not found');
        }

        // Rich text toolbar events
        this.setupRichTextToolbar();

        // Table toolbar events
        this.setupTableToolbar();

        // Content change detection (debounced to improve typing performance)
        document.getElementById('editableContent').addEventListener('input', (e) => {
            // Update last action for history subtitle immediately
            this.lastAction = 'Texte modifié';
            // If typing inside a table cell, disarm row deletion selection
            try {
                const inCell = e && e.target && e.target.closest ? e.target.closest('td,th') : null;
                if (inCell) { this._selectedRowForAction = null; const t = inCell.closest('table'); this.normalizeTableForTyping(t); }
            } catch (_) { /* no-op */ }
            // Debounce heavy saves while typing
            if (this._inputSaveTimer) clearTimeout(this._inputSaveTimer);
            this._inputSaveTimer = setTimeout(() => {
                try {
                    this.saveState();
                    this.updateLastModified();
                    this.autoSaveToLocalStorage();
                    // Re-apply any persisted section background colors to withstand inner HTML rewrites
                    try {
                        const sectionsWithBg = document.querySelectorAll('.newsletter-section[data-section-bg]');
                        sectionsWithBg.forEach(sec => this.reapplySectionBackground(sec));
                    } catch (_) { /* no-op */ }
                } catch (_) { /* no-op */ }
            }, 300);
        });

        // Track selection inside editable content to keep toolbar actions working
        const editable = document.getElementById('editableContent');
        const edge=8;
        const clrSel=(t)=>{if(!t)return;t.querySelectorAll('td,th').forEach(c=>{c.style.outline='';c.style.outlineOffset='';c.style.backgroundColor='';});};
        const mark=(c,on)=>{if(!c)return;if(on){c.style.outline='2px solid rgba(10,155,205,0.6)';c.style.outlineOffset='-2px';c.style.backgroundColor='rgba(10,155,205,0.07)';}else{c.style.outline='';c.style.outlineOffset='';c.style.backgroundColor='';}};
        const rectSel=(t,a,b)=>{if(!t||!a||!b)return;const rs=Array.from(t.rows||[]);const ap={ri:rs.indexOf(a.parentElement),ci:a.cellIndex};const bp={ri:rs.indexOf(b.parentElement),ci:b.cellIndex};const r0=Math.min(ap.ri,bp.ri),r1=Math.max(ap.ri,bp.ri),c0=Math.min(ap.ci,bp.ci),c1=Math.max(ap.ci,bp.ci);rs.forEach((r,ri)=>{if(ri<r0||ri>r1)return;Array.from(r.cells).forEach((cc,ci)=>{if(ci>=c0&&ci<=c1)mark(cc,true);});});};
        const atR=(r,x)=>r.right-x<=edge;const atB=(r,y)=>r.bottom-y<=edge;
        const ensureFixed=(t)=>{if(!t)return;const r=t.getBoundingClientRect();if(!t.style.width||t.style.width.indexOf('%')!==-1)t.style.width=Math.max(20,Math.floor(r.width))+'px';t.style.tableLayout='fixed';};
        const setColW=(t,ci,w)=>{w=Math.max(20,Math.floor(w));Array.from(t.rows||[]).forEach(r=>{const c=r.cells[ci];if(c)c.style.width=w+'px';});};
        const autoFitCol=(t,ci)=>{let m=20;Array.from(t.rows||[]).forEach(r=>{const c=r.cells[ci];if(!c)return;const cs=window.getComputedStyle(c);const pad=parseFloat(cs.paddingLeft)+parseFloat(cs.paddingRight)+parseFloat(cs.borderLeftWidth)+parseFloat(cs.borderRightWidth);m=Math.max(m,Math.ceil(c.scrollWidth+pad));});ensureFixed(t);setColW(t,ci,m);};
        const autoFitRow=(row)=>{let m=40;Array.from(row.cells||[]).forEach(c=>{const cs=window.getComputedStyle(c);const pad=parseFloat(cs.paddingTop)+parseFloat(cs.paddingBottom)+parseFloat(cs.borderTopWidth)+parseFloat(cs.borderBottomWidth);m=Math.max(m,Math.ceil(c.scrollHeight+pad),40);});row.style.height=m+'px';};
        const autoFitTableWindow=(t)=>{t.style.width='100%';};
        const distributeCols=(t)=>{const rs=Array.from(t.rows||[]);if(rs.length===0)return;const n=rs[0].cells.length;if(n===0)return;const tw=t.getBoundingClientRect().width;const w=Math.max(20,Math.floor(tw/n));ensureFixed(t);for(let i=0;i<n;i++)setColW(t,i,w);};
        const distributeRows=(t)=>{const rs=Array.from(t.rows||[]);if(rs.length===0)return;let total=0;rs.forEach(r=>total+=r.getBoundingClientRect().height);const h=Math.max(20,Math.floor(total/rs.length));rs.forEach(r=>r.style.height=h+'px');};
        let drag=null;let anchorCell=null;let activeTable=null;
        let moveHandle=null;let moving=false;let movePlaceholder=null;let ctxMenu=null;let ctxRef={};
        let selecting=false;
        editable.addEventListener('mousedown',(e)=>{
            const cell=e.target&&e.target.closest?e.target.closest('td,th'):null;const tbl=cell?cell.closest('table'):(e.target.closest?e.target.closest('table'):null);if(!tbl)return;activeTable=tbl;try{tbl.querySelectorAll('td,th').forEach(c=>c.contentEditable='true');}catch(_){}const tr=tbl.getBoundingClientRect();if(atR(tr,e.clientX)){drag={m:'tw',sx:e.clientX,w:tr.width,t:tbl};document.body.style.userSelect='none';e.preventDefault();return;}if(atB(tr,e.clientY)){drag={m:'th',sy:e.clientY,h:tr.height,t:tbl};document.body.style.userSelect='none';e.preventDefault();return;}if(!cell)return;const cr=cell.getBoundingClientRect();if(e.detail===2&&atR(cr,e.clientX)){autoFitCol(tbl,cell.cellIndex);return;}if(e.detail===2&&atB(cr,e.clientY)){autoFitRow(cell.parentElement);return;}if(atR(cr,e.clientX)){ensureFixed(tbl);drag={m:'col',sx:e.clientX,w:cr.width,ci:cell.cellIndex,t:tbl};document.body.style.userSelect='none';e.preventDefault();return;}if(atB(cr,e.clientY)){drag={m:'row',sy:e.clientY,h:cell.parentElement.getBoundingClientRect().height,row:cell.parentElement,t:tbl};document.body.style.userSelect='none';e.preventDefault();return;}if(e.shiftKey){if(!anchorCell||anchorCell.closest('table')!==tbl)anchorCell=cell;clrSel(tbl);rectSel(tbl,anchorCell,cell);return;}if(e.ctrlKey||e.metaKey){clrSel(tbl);const ci=cell.cellIndex;Array.from(tbl.rows||[]).forEach(r=>{const c=r.cells[ci];if(c)mark(c,true);});try{const sel=window.getSelection();if(sel)sel.removeAllRanges();}catch(_){}return;}if(e.altKey){clrSel(tbl);Array.from(cell.parentElement.cells||[]).forEach(c=>mark(c,true));try{const sel=window.getSelection();if(sel)sel.removeAllRanges();}catch(_){}return;}clrSel(tbl);mark(cell,true);anchorCell=cell;
            // Start drag selection with left mouse button and no modifiers
            if (e.button===0 && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey) {
                selecting=true; document.body.style.userSelect='none';
            }
        });
        const ensureMoveHandle=()=>{if(moveHandle)return;moveHandle=document.createElement('div');moveHandle.style.position='absolute';moveHandle.style.width='12px';moveHandle.style.height='12px';moveHandle.style.border='1px solid #0a9bcd';moveHandle.style.background='#fff';moveHandle.style.cursor='move';moveHandle.style.zIndex='5000';document.body.appendChild(moveHandle);moveHandle.addEventListener('mousedown',(e)=>{if(!activeTable)return;moving=true;e.preventDefault();document.body.style.userSelect='none';if(!movePlaceholder){movePlaceholder=document.createElement('div');movePlaceholder.style.height=activeTable.getBoundingClientRect().height+'px';movePlaceholder.style.border='1px dashed #0a9bcd';}activeTable.parentNode.insertBefore(movePlaceholder,activeTable);});};
        const positionMoveHandle=()=>{if(!activeTable)return;if(!moveHandle)ensureMoveHandle();const r=activeTable.getBoundingClientRect();const sx=window.pageXOffset||document.documentElement.scrollLeft;const sy=window.pageYOffset||document.documentElement.scrollTop;moveHandle.style.left=Math.max(0,Math.floor(r.left+sx-14))+'px';moveHandle.style.top=Math.max(0,Math.floor(r.top+sy-14))+'px';moveHandle.style.display='block';};
        editable.addEventListener('click',(e)=>{const t=e.target.closest?e.target.closest('table'):null;if(t){activeTable=t;this.normalizeTableForTyping(t);positionMoveHandle();} const inCell=e.target&&e.target.closest?e.target.closest('td,th'):null; if(inCell){ this._selectedRowForAction=null; }});
        const updateDropPreview=(x,y)=>{const el=document.elementFromPoint(x,y);if(!el)return;const host=document.getElementById('editableContent');if(!host)return;const blk=el.closest?el.closest('#editableContent > *'):null; if(blk&&movePlaceholder&&blk!==movePlaceholder){host.insertBefore(movePlaceholder,blk);} else if(!blk&&movePlaceholder){host.appendChild(movePlaceholder);}};
        document.addEventListener('mousemove',(e)=>{if(moving){updateDropPreview(e.clientX,e.clientY);e.preventDefault();return;}if(selecting&&activeTable){const el=document.elementFromPoint(e.clientX,e.clientY);const over=el&&el.closest?el.closest('td,th'):null; if(over&&over.closest('table')===activeTable&&anchorCell){clrSel(activeTable);rectSel(activeTable,anchorCell,over); const rs=Array.from(activeTable.rows||[]); const ap={ri:rs.indexOf(anchorCell.parentElement),ci:anchorCell.cellIndex}; const bp={ri:rs.indexOf(over.parentElement),ci:over.cellIndex}; const r0=Math.min(ap.ri,bp.ri), r1=Math.max(ap.ri,bp.ri); if (r0===r1 && r0>=0) { this._selectedRowForAction = rs[r0]; } else { this._selectedRowForAction = null; } try{const sel=window.getSelection();if(sel)sel.removeAllRanges();}catch(_){}} e.preventDefault();return;} if(activeTable&&moveHandle)positionMoveHandle();});
        document.addEventListener('mouseup',()=>{if(selecting){selecting=false;document.body.style.userSelect='';} if(moving){moving=false;document.body.style.userSelect='';if(movePlaceholder&&movePlaceholder.parentNode){movePlaceholder.parentNode.insertBefore(activeTable,movePlaceholder);movePlaceholder.remove();}try{this.saveState();this.updateLastModified();this.autoSaveToLocalStorage();this.lastAction='Table déplacée';}catch(_){}}});
        const ensureCtxMenu=()=>{ return; };
        editable.addEventListener('contextmenu',(e)=>{ return; });
        const mm=(e)=>{if(!drag)return;if(drag.m==='col'){const dx=e.clientX-drag.sx;setColW(drag.t,drag.ci,drag.w+dx);e.preventDefault();return;}if(drag.m==='row'){const dy=e.clientY-drag.sy;drag.row.style.height=Math.max(20,Math.floor(drag.h+dy))+'px';e.preventDefault();return;}if(drag.m==='tw'){const dx=e.clientX-drag.sx;drag.t.style.width=Math.max(50,Math.floor(drag.w+dx))+'px';e.preventDefault();return;}if(drag.m==='th'){const dy=e.clientY-drag.sy;drag.t.style.height=Math.max(20,Math.floor(drag.h+dy))+'px';e.preventDefault();return;}};
        const mu=()=>{if(!drag)return;drag=null;document.body.style.userSelect='';try{this.saveState();this.updateLastModified();this.autoSaveToLocalStorage();this.lastAction='Taille du tableau modifiée';}catch(_){}};
        document.addEventListener('mousemove',mm);
        document.addEventListener('mouseup',mu);
        document.addEventListener('keydown',(e)=>{ return; });
        // Disarm row deletion selection on any non-delete key press
        document.addEventListener('keydown',(e)=>{ if (e.key !== 'Delete' && e.key !== 'Backspace') { this._selectedRowForAction = null; } });
        // Delete selected row with keyboard if a row was selected via context action
        document.addEventListener('keydown',(e)=>{
            if ((e.key === 'Delete' || e.key === 'Backspace') && this._selectedRowForAction && this.currentEditingTable) {
                const row = this._selectedRowForAction;
                // If the caret/selection is inside this row, do NOT delete the row (let text delete normally)
                let selectionInsideRow = false;
                try {
                    const sel = window.getSelection();
                    if (sel && sel.rangeCount > 0) {
                        const container = sel.getRangeAt(0).startContainer;
                        const host = container && (container.nodeType === Node.ELEMENT_NODE ? container : container.parentElement);
                        if (host && row.contains(host)) selectionInsideRow = true;
                    }
                } catch (_) {}
                if (selectionInsideRow) return; // allow normal text deletion

                const allRows = this.currentEditingTable.querySelectorAll('tr');
                if (allRows.length > 1) {
                    row.remove();
                    this._selectedRowForAction = null;
                    try { this.saveState(); this.updateLastModified(); this.autoSaveToLocalStorage(); this.lastAction = 'Ligne de tableau supprimée'; } catch(_) {}
                    e.preventDefault();
                }
            }
        });
        editable.addEventListener('mouseup', () => this.saveSelection());
        editable.addEventListener('keyup', () => this.saveSelection());

        const hoverCursor=(e)=>{
            if (drag||moving) return;
            const tbl=e.target&&e.target.closest?e.target.closest('table'):null;
            if (!tbl) { editable.style.cursor=''; return; }
            const tr=tbl.getBoundingClientRect();
            if (tr.right - e.clientX <= edge) { editable.style.cursor='col-resize'; return; }
            if (tr.bottom - e.clientY <= edge) { editable.style.cursor='row-resize'; return; }
            const cell=e.target.closest?e.target.closest('td,th'):null;
            if (!cell) { editable.style.cursor=''; return; }
            const cr=cell.getBoundingClientRect();
            if (cr.right - e.clientX <= edge) { editable.style.cursor='col-resize'; return; }
            if (cr.bottom - e.clientY <= edge) { editable.style.cursor='row-resize'; return; }
            editable.style.cursor='';
        };
        editable.addEventListener('mousemove', hoverCursor);

        // Show Video/Table toolbars on click; Section toolbar shows on Ctrl+Click
        editable.addEventListener('click', (e) => {
            // If clicking the dedicated column image placeholder, do not show the section toolbar.
            // Let the placeholder's own click handler open the file picker.
            try {
                if (e.target && e.target.closest && e.target.closest('#columnImagePlaceholder')) {
                    this.hideSectionToolbar();
                    if (!e.target.closest('#videoToolbar')) this.hideVideoToolbar();
                    if (!e.target.closest('#tableToolbar')) this.hideTableToolbar();
                    return;
                }
            } catch (_) { /* no-op */ }

            // Ctrl+Click to show Section toolbar directly on the clicked section
            // This is an additive shortcut; existing behaviors remain unchanged.
            if (e.ctrlKey) {
                const sectionElCtrl = e.target.closest('.newsletter-section, .gallery-section, .two-column-layout, .syc-item, .cta-section');
                if (sectionElCtrl) {
                    e.preventDefault();
                    this.showSectionToolbar(sectionElCtrl);
                    return;
                }
            }

            // If there is an active text selection (non-collapsed), do NOT show the section toolbar.
            // This avoids the section floating tools appearing while selecting text.
            try {
                const sel = window.getSelection && window.getSelection();
                const hasSelection = sel && sel.rangeCount > 0 && !sel.getRangeAt(0).collapsed;
                if (hasSelection) {
                    this.hideSectionToolbar();
                    // Also hide other non-text toolbars to reduce interference during selection
                    if (!e.target.closest('#videoToolbar')) this.hideVideoToolbar();
                    if (!e.target.closest('#tableToolbar')) this.hideTableToolbar();
                    return;
                }
            } catch (_) { /* no-op */ }

            const sectionEl = e.target.closest('.newsletter-section, .gallery-section, .two-column-layout, .syc-item, .cta-section');
            const videoEl = e.target.closest('video, iframe');
            const tableEl = e.target.closest('table');

            // Do not show section toolbar on single click anymore; hide if clicking outside any section
            if (!sectionEl) {
                this.hideSectionToolbar();
            }

            if (videoEl) {
                this.showVideoToolbar(videoEl);
            } else if (!e.target.closest('#videoToolbar')) {
                this.hideVideoToolbar();
            }

            if (tableEl) {
                this.showTableToolbar(tableEl);
            } else if (!e.target.closest('#tableToolbar')) {
                this.hideTableToolbar();
            }
        });

        // Removed: double-click to show Section toolbar (replaced by Ctrl+Click shortcut)

        // Track mouse position to insert sections at the pointer
        this.lastMouseRange = null;
        const getRangeFromPoint = (x, y) => {
            const editableEl = document.getElementById('editableContent');
            if (!editableEl) return null;
            const containerRect = editableEl.getBoundingClientRect();
            if (x < containerRect.left || x > containerRect.right || y < containerRect.top || y > containerRect.bottom) {
                return null;
            }

            // Helper to snap a generic (possibly inside child) range to before/after the nearest direct child
            const snapRangeToBlockBoundary = (node) => {
                if (!node) return null;
                let child = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
                while (child && child.parentElement !== editableEl) {
                    child = child.parentElement;
                }
                const range = document.createRange();
                if (!child || child === editableEl) {
                    const last = editableEl.lastChild;
                    if (last) range.setStartAfter(last); else range.setStart(editableEl, editableEl.childNodes.length);
                    range.collapse(true);
                    return range;
                }
                const childRect = child.getBoundingClientRect();
                const placeAfter = y > (childRect.top + childRect.height / 2);
                if (placeAfter) range.setStartAfter(child); else range.setStartBefore(child);
                range.collapse(true);
                return range;
            };

            // Try standard caret APIs first, but snap to top-level section boundary
            if (document.caretRangeFromPoint) {
                const r = document.caretRangeFromPoint(x, y);
                if (r) return snapRangeToBlockBoundary(r.startContainer);
            }
            if (document.caretPositionFromPoint) {
                const pos = document.caretPositionFromPoint(x, y);
                if (pos && pos.offsetNode != null) {
                    return snapRangeToBlockBoundary(pos.offsetNode);
                }
            }

            // Fallback using elementFromPoint
            const target = document.elementFromPoint(x, y);
            if (!target) return null;
            const container = target.closest('#editableContent');
            if (!container) return null;
            return snapRangeToBlockBoundary(target);
        };

        const updateMouseRange = (e) => {
            // Track pointer inside editor for precise insertion
            this.lastMousePosition = { x: e.clientX, y: e.clientY };
            const r = this.computeRangeFromPoint(e.clientX, e.clientY);
            if (r) {
                this.lastMouseRange = r.cloneRange ? r.cloneRange() : r;
                return;
            }
            // Fallback: caret at end
            const editableEl = document.getElementById('editableContent');
            const endRange = document.createRange();
            if (editableEl.lastChild) {
                endRange.setStartAfter(editableEl.lastChild);
            } else {
                endRange.setStart(editableEl, editableEl.childNodes.length);
            }
            endRange.collapse(true);
            this.lastMouseRange = endRange;
        };

        editable.addEventListener('mousemove', updateMouseRange);
        editable.addEventListener('click', updateMouseRange);
        
        // Track mouse position for table color picker
        document.addEventListener('mousemove', (e) => {
            this.lastMousePosition = { x: e.clientX, y: e.clientY };
            
            // Store the current table cell if hovering over one
            const elementUnderMouse = document.elementFromPoint(e.clientX, e.clientY);
            if (elementUnderMouse) {
                const cell = elementUnderMouse.closest('td, th');
                if (cell && cell.closest('table')) {
                    this.lastHoveredTableCell = cell;
                }
            }
        });
        
        // Also capture cell on click for more precision
        document.addEventListener('click', (e) => {
            const elementUnderClick = document.elementFromPoint(e.clientX, e.clientY);
            if (elementUnderClick) {
                const cell = elementUnderClick.closest('td, th');
                if (cell && cell.closest('table')) {
                    this.lastClickedTableCell = cell;
                }
            }
        });
        // While starting a new drag selection inside the editor, hide the toolbar to avoid interference
        editable.addEventListener('mousedown', () => {
            const toolbar = document.getElementById('richTextToolbar');
            if (toolbar) toolbar.style.display = 'none';
            this.hideSectionToolbar();
            this.hideVideoToolbar();
            this.hideTableToolbar();
        });

        // Add keyboard support for deleting images, even when the caret is just next to them
        document.getElementById('editableContent').addEventListener('keydown', (e) => {
            if (e.key === 'Delete' || e.key === 'Backspace') {
                const selection = window.getSelection();
                if (selection.rangeCount > 0) {
                    const range = selection.getRangeAt(0);
                    const selectedElement = range.commonAncestorContainer;

                    // 1) Detect if an image (or its wrapper) is actually selected
                    let imageWrapper = null;
                    if (selectedElement.nodeType === Node.ELEMENT_NODE) {
                        const el = selectedElement;
                        if (el.classList && el.classList.contains('image-wrapper')) {
                            imageWrapper = el;
                        } else if (el.tagName === 'IMG') {
                            imageWrapper = el.closest('.image-wrapper');
                        }
                    } else if (selectedElement.parentElement) {
                        const parent = selectedElement.parentElement;
                        if (parent.classList && parent.classList.contains('image-wrapper')) {
                            imageWrapper = parent;
                        } else if (parent.tagName === 'IMG') {
                            imageWrapper = parent.closest('.image-wrapper');
                        }
                    }

                    // 2) If nothing is selected (collapsed caret), check adjacency to an image so text next to image isn't deleted silently
                    if (!imageWrapper && selection.isCollapsed) {
                        // Helper to resolve node at caret boundary and find adjacent sibling in the right direction
                        const getAdjacentNode = (rng, direction) => {
                            let container = rng.startContainer;
                            let offset = rng.startOffset;

                            // If we're inside a text node, adjacent sibling depends on offset
                            if (container.nodeType === Node.TEXT_NODE) {
                                const parent = container.parentNode;
                                if (!parent) return null;
                                // Backspace removes previous content at start of text node
                                if (direction === 'backward' && offset === 0) {
                                    return parent.childNodes[parent.childNodes.length ? Array.prototype.indexOf.call(parent.childNodes, container) - 1 : -1] || container.previousSibling;
                                }
                                // Delete removes next content at end of text node
                                if (direction === 'forward' && offset === container.nodeValue.length) {
                                    return container.nextSibling;
                                }
                                return null;
                            }

                            // If we're in an element node, use offset to get child or neighbor
                            if (container.nodeType === Node.ELEMENT_NODE) {
                                const child = container.childNodes[offset] || null;
                                if (direction === 'forward') {
                                    // If the child is an element, it's the candidate; otherwise the next sibling
                                    return child || container.childNodes[offset] || null;
                                } else {
                                    // Backward: previous sibling of the position
                                    return container.childNodes[offset - 1] || null;
                                }
                            }
                            return null;
                        };

                        const dir = (e.key === 'Backspace') ? 'backward' : 'forward';
                        let neighbor = null;
                        // Only consider TEXT_NODE boundaries to avoid false positives while editing inside elements
                        const container = range.startContainer;
                        const offset = range.startOffset;
                        if (container && container.nodeType === Node.TEXT_NODE) {
                            if (dir === 'backward' && offset === 0) {
                                neighbor = container.previousSibling;
                            } else if (dir === 'forward' && offset === container.nodeValue.length) {
                                neighbor = container.nextSibling;
                            }
                        }

                        if (neighbor && neighbor.nodeType === Node.ELEMENT_NODE) {
                            // Immediate sibling must be an image wrapper or IMG node
                            if (neighbor.classList && neighbor.classList.contains('image-wrapper')) {
                                imageWrapper = neighbor;
                            } else if (neighbor.tagName === 'IMG') {
                                imageWrapper = neighbor;
                            }
                        }
                    }

                    // 3) If we determined an image wrapper to delete, intercept and confirm
                    if (imageWrapper) {
                        e.preventDefault();
                        e.stopPropagation();
                        (async () => { if (await this.confirmWithCancel("Êtes-vous sûr de vouloir supprimer cette image ?")) { this.deleteImageWrapperSafe(imageWrapper); } })();
                    }
                }
            }
        });


        // Close modal events
        const closeHistoryModal = document.getElementById('closeHistoryModal');
        if (closeHistoryModal) {
            closeHistoryModal.addEventListener('click', () => {
                const modal = document.getElementById('historyModal');
                if (modal) {
                    modal.style.display = 'none';
                }
            });
        } else {
            console.error('closeHistoryModal not found');
        }

        // Sidebar (mobile) overlay toggle — non-intrusive wiring
        const sidebarHamburger = document.getElementById('sidebarHamburger');
        const sidebarOverlay = document.getElementById('sidebarOverlay');
        const editorSidebar = document.querySelector('.editor-sidebar');
        if (sidebarHamburger && sidebarOverlay && editorSidebar) {
            const openSidebar = () => {
                editorSidebar.classList.add('active');
                sidebarOverlay.classList.add('active');
                document.body.classList.add('sidebar-open');
            };
            const closeSidebar = () => {
                editorSidebar.classList.remove('active');
                sidebarOverlay.classList.remove('active');
                document.body.classList.remove('sidebar-open');
            };

            sidebarHamburger.addEventListener('click', () => {
                // Toggle to allow open/close
                if (editorSidebar.classList.contains('active')) {
                    closeSidebar();
                } else {
                    openSidebar();
                }
            });
            sidebarOverlay.addEventListener('click', () => {
                closeSidebar();
            });
            document.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') closeSidebar();
            });
            // On resize to desktop, ensure closed state
            window.addEventListener('resize', () => {
                if (window.innerWidth > 768) {
                    closeSidebar();
                }
            });
        }

        window.addEventListener('click', (e) => {
            if (e.target.id === 'historyModal') {
                document.getElementById('historyModal').style.display = 'none';
            }
            
            // Hide image toolbar when clicking outside of an image or the toolbar
            if (!e.target.closest('#imageToolbar') && 
                e.target.tagName !== 'IMG' && 
                !e.target.closest('.crop-overlay')) {
                this.hideImageEditingTools();
            }

            // Hide section toolbar when clicking outside
            if (!e.target.closest('#sectionToolbar')) {
                if (!e.target.closest('.newsletter-section') && !e.target.closest('.gallery-section') && !e.target.closest('.two-column-layout') && !e.target.closest('.syc-item') && !e.target.closest('.cta-section')) {
                    this.hideSectionToolbar();
                }
            }

            // Webinar toolbar visibility
            const webinarSection = e.target.closest('.cta-section');
            if (webinarSection) {
                this.showWebinarToolbar(webinarSection);
            } else if (!e.target.closest('#webinarToolbar')) {
                this.hideWebinarToolbar();
            }

            // Hide video toolbar when clicking outside
            if (!e.target.closest('#videoToolbar')) {
                if (!e.target.closest('video') && !e.target.closest('iframe')) {
                    this.hideVideoToolbar();
                }
            }

            // Hide table toolbar when clicking outside
            if (!e.target.closest('#tableToolbar')) {
                if (!e.target.closest('table')) {
                    this.hideTableToolbar();
                }
            }

            // Close dropdowns when clicking outside
            const videoDD = document.getElementById('videoSizeOptions');
            const videoBtn = document.getElementById('videoSizeBtn');
            if (videoDD && videoDD.style.display === 'block' && !videoDD.contains(e.target) && !videoBtn.contains(e.target)) {
                videoDD.style.display = 'none';
            }
            const sectionDD = document.getElementById('sectionWidthOptions');
            const sectionBtn = document.getElementById('sectionWidthBtn');
            if (sectionDD && sectionDD.style.display === 'block' && !sectionDD.contains(e.target) && !sectionBtn.contains(e.target)) {
                sectionDD.style.display = 'none';
            }
            const tableBgDD = document.getElementById('tableBgColorDropdownContent');
            const tableBgBtn = document.getElementById('tableBgColorDropdownBtn');
            if (tableBgDD && tableBgDD.style.display === 'block' && !tableBgDD.contains(e.target) && !tableBgBtn.contains(e.target)) {
                tableBgDD.style.display = 'none';
            }
            // Close section background dropdown when clicking outside
            const sectionBgDD = document.getElementById('sectionBgColorDropdownContent');
            const sectionBgBtn = document.getElementById('sectionBgColorDropdownBtn');
            if (sectionBgDD && sectionBgDD.style.display === 'block' && !sectionBgDD.contains(e.target) && !sectionBgBtn.contains(e.target)) {
                sectionBgDD.style.display = 'none';
            }
        });
    }
    
    // Apply image positioning modes from the image toolbar
    setImagePosition(mode) {
        if (!this.currentEditingImage) return;
        // Ensure we have a wrapper around the image, consistent with other tools
        let wrapper = this.currentEditingImage.closest('.image-wrapper');
        if (!wrapper) {
            // Reuse the logic from addResizeHandlesToImage to create a wrapper if missing
            wrapper = document.createElement('div');
            wrapper.className = 'image-wrapper';
            wrapper.style.cssText = 'position: relative; display: inline-block; margin: 10px 0;';
            const img = this.currentEditingImage;
            if (img.parentNode) {
                img.parentNode.insertBefore(wrapper, img);
                wrapper.appendChild(img);
            } else {
                // Fallback: append to editable area
                const host = document.getElementById('editableContent');
                if (host) { host.appendChild(wrapper); wrapper.appendChild(img); }
            }
        }
        
        // Remove previous positioning classes
        ['position-absolute', 'float-left', 'float-right', 'position-inline'].forEach(cls => {
            wrapper.classList.remove(cls);
        });
        
        // Reset inline positioning styles when switching modes
        wrapper.style.position = wrapper.style.position || 'relative';
        wrapper.style.left = '';
        wrapper.style.top = '';
        wrapper.style.right = '';
        wrapper.style.bottom = '';
        wrapper.style.margin = wrapper.style.margin || '10px 0';
        
        const img = this.currentEditingImage;

        if (mode === 'float-left') {
            wrapper.classList.add('float-left');
            // Ensure normal flow
            wrapper.style.position = 'relative';
            // Reset gallery absolute image styles so float takes effect
            img.style.display = '';
            img.style.position = '';
            img.style.width = '';
            img.style.height = '';
            img.style.objectFit = '';
        } else if (mode === 'float-right') {
            wrapper.classList.add('float-right');
            wrapper.style.position = 'relative';
            img.style.display = '';
            img.style.position = '';
            img.style.width = '';
            img.style.height = '';
            img.style.objectFit = '';
        } else {
            // Default to inline (centered block)
            wrapper.classList.add('position-inline');
            wrapper.style.position = 'relative';
            img.style.display = 'block';
            img.style.position = '';
            img.style.width = '';
            img.style.height = '';
            img.style.objectFit = '';
        }
        
        // Update state and keep handles in sync
        this.removeResizeHandles();
        this.addResizeHandlesToImage(this.currentEditingImage);
        this.saveState();
        this.lastAction = 'Position de l\'image modifiée';
    }

    // Make an absolutely positioned image wrapper draggable within its parent container bounds
    enableAbsoluteDrag(wrapper) {
        try {
            const container = wrapper.parentElement && (wrapper.parentElement.closest('.gallery-image-container') || wrapper.parentElement.closest('#editableContent') || wrapper.parentElement);
            if (!container) return;
            const onMouseDown = (e) => {
                if (!wrapper.classList.contains('position-absolute')) return;
                e.preventDefault();
                e.stopPropagation();
                const startX = e.clientX;
                const startY = e.clientY;
                const rect = wrapper.getBoundingClientRect();
                const contRect = container.getBoundingClientRect();
                const offsetLeft = rect.left - contRect.left;
                const offsetTop = rect.top - contRect.top;
                const onMove = (ev) => {
                    const dx = ev.clientX - startX;
                    const dy = ev.clientY - startY;
                    let newLeft = offsetLeft + dx;
                    let newTop = offsetTop + dy;
                    // Constrain within container
                    newLeft = Math.max(0, Math.min(newLeft, contRect.width - rect.width));
                    newTop = Math.max(0, Math.min(newTop, contRect.height - rect.height));
                    wrapper.style.left = newLeft + 'px';
                    wrapper.style.top = newTop + 'px';
                };
                const onUp = () => {
                    document.removeEventListener('mousemove', onMove);
                    document.removeEventListener('mouseup', onUp);
                    // Persist move
                    try { this.saveState(); this.lastAction = 'Image déplacée'; } catch (_) {}
                };
                document.addEventListener('mousemove', onMove);
                document.addEventListener('mouseup', onUp);
            };
            // Avoid stacking multiple listeners
            wrapper.removeEventListener('mousedown', wrapper.__absDragHandler);
            wrapper.__absDragHandler = onMouseDown;
            wrapper.addEventListener('mousedown', onMouseDown);
        } catch (_) { /* no-op */ }
    }

    setupRichTextToolbar() {
        const toolbar = document.getElementById('richTextToolbar');
        const editable = document.getElementById('editableContent');

        const isSelectionInEditable = () => {
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) return false;
            const container = sel.getRangeAt(0).commonAncestorContainer;
            const node = container.nodeType === Node.ELEMENT_NODE ? container : container.parentNode;
            return editable.contains(node);
        };
        const isActiveInToolbar = () => {
            const active = document.activeElement;
            return active && toolbar.contains(active);
        };
        let isInteractingWithToolbar = false;

        const execWithRestore = (command, showUi = false, value = null) => {
            // Ensure the editable retains focus so execCommand targets it
            if (editable) {
                editable.focus();
            }
            this.restoreSelection();
            document.execCommand(command, showUi, value);
            this.saveSelection();
        };

        // Ensure list markers follow text alignment by toggling list-style-position
        const setListMarkerPositionForSelection = (mode) => {
            try {
                const sel = window.getSelection();
                if (!sel || sel.rangeCount === 0) return;
                let node = sel.anchorNode;
                if (node && node.nodeType === 3) node = node.parentNode;
                if (!node || !node.closest) return;
                const list = node.closest('ul, ol');
                if (!list) return;
                if (mode === 'center' || mode === 'right') {
                    list.style.listStylePosition = 'inside';
                    list.style.textAlign = mode; // ensure alignment applies to the list block
                } else {
                    list.style.listStylePosition = 'outside';
                    list.style.textAlign = 'left';
                }
            } catch (_) { /* ignore */ }
        };

        // Prefer styling with CSS spans instead of deprecated <font>
        try {
            document.execCommand('styleWithCSS', false, true);
        } catch (_) {}
        
        // Keep selection when interacting with toolbar UI, but allow native controls (select/input)
        toolbar.addEventListener('mousedown', (e) => {
            // Guard window selectionchange race when clicking toolbar controls
            isInteractingWithToolbar = true;
            setTimeout(() => { isInteractingWithToolbar = false; }, 300);
            const target = e.target;
            const tag = target.tagName;
            const isFormControl = tag === 'SELECT' || tag === 'INPUT' || tag === 'TEXTAREA' ||
                target.closest('select') || target.closest('input') || target.closest('textarea');
            if (isFormControl) {
                // Keep focus on the select so the dropdown stays open
                target.focus();
                return; // allow native dropdowns like font size to open
            }
            // Prevent focus shift for non-form controls to preserve selection
            e.preventDefault();
        });

        // Don’t steal focus on toolbar click; just restore selection before executing actions in handlers
        toolbar.addEventListener('click', (e) => {
            // If clicking within toolbar but not a form control, keep editor selection
            const target = e.target;
            if (!(target.tagName === 'SELECT' || target.closest('select'))) {
                this.restoreSelection();
            }
        });
        // Preview shortcut: Ctrl+Shift+P opens a sanitized preview in a new tab
        document.addEventListener('keydown', (ev) => {
            try {
                const key = ev.key || ev.code;
                if (ev.ctrlKey && ev.shiftKey && (key === 'P' || key === 'KeyP')) {
                    ev.preventDefault();
                    this.previewFull();
                }
            } catch (_) { /* no-op */ }
        });

        // Font family
        document.getElementById('fontFamily').addEventListener('change', (e) => {
            execWithRestore('fontName', false, e.target.value);
        });

        // Helpers to enforce single-use for Titre_P (52px) and Sous_titre_P (22px)
        const getEditableRoot = () => document.getElementById('editableContent');
        const countInlineFontSize = (px) => {
            try {
                const root = getEditableRoot();
                if (!root) return 0;
                let n = 0;
                root.querySelectorAll('*').forEach(el => {
                    if (el && el.style && el.style.fontSize === px) n++;
                });
                return n;
            } catch (_) { return 0; }
        };
        const selectionWithinInlineFontSize = (px) => {
            try {
                const sel = window.getSelection();
                if (!sel || sel.rangeCount === 0) return false;
                let node = sel.getRangeAt(0).commonAncestorContainer;
                if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
                const root = getEditableRoot();
                while (node && node !== root) {
                    if (node && node.style && node.style.fontSize === px) return true;
                    node = node.parentElement;
                }
                return false;
            } catch (_) { return false; }
        };

        // Helper to unlock special title classes (e.g., imported .sujette-title)
        // and clear inline 52px / 22px font sizes inside the current context
        // when changing font-size away from the predefined 52px / 22px titles.
        const unlockTitleClassesOnPath = (startNode) => {
            try {
                const root = getEditableRoot();
                if (!root || !startNode) return;
                let node = startNode.nodeType === Node.TEXT_NODE ? startNode.parentElement : startNode;
                while (node && node !== root) {
                    if (node.classList && node.classList.contains('sujette-title')) {
                        node.classList.remove('sujette-title');
                    }
                    node = node.parentElement;
                }
            } catch (_) { /* no-op */ }
        };

        const clearInlineTitleSizesInRange = (range) => {
            try {
                const root = getEditableRoot();
                if (!root || !range) return;
                const walker = document.createTreeWalker(
                    root,
                    NodeFilter.SHOW_ELEMENT,
                    {
                        acceptNode(node) {
                            if (!range.intersectsNode(node)) return NodeFilter.FILTER_REJECT;
                            const s = node.style && node.style.fontSize;
                            if (s === '52px' || s === '22px') return NodeFilter.FILTER_ACCEPT;
                            return NodeFilter.FILTER_SKIP;
                        }
                    }
                );
                const toClear = [];
                while (walker.nextNode()) {
                    toClear.push(walker.currentNode);
                }
                toClear.forEach(node => {
                    try { node.style.fontSize = ''; } catch (_) {}
                });
            } catch (_) { /* no-op */ }
        };

        // Font size handlers — apply on both change and click (when value may not change)
        let fontSizeAppliedRecently = false; // guard to avoid double-apply
        const applyFontSize = (fontSize) => {
            fontSizeAppliedRecently = true;
            // Ignore placeholder/no-op so opening the dropdown doesn't steal focus and close it
            if (!fontSize) return;
            this.restoreSelection();
            const selection = window.getSelection();
            if (!selection || selection.rangeCount === 0) return;
            const range = selection.getRangeAt(0);

            // If we are changing to a size other than 52 (Titre_P) or 22 (Sous_titre_P),
            // unlock any imported title classes on the path and clear existing
            // inline 52px / 22px sizes inside the selection so the new size can
            // take visual precedence.
            if (fontSize !== '52' && fontSize !== '22') {
                unlockTitleClassesOnPath(range.commonAncestorContainer);
                clearInlineTitleSizesInRange(range);
            }

            // Enforce uniqueness for 52 (Titre_P) and 22 (Sous_titre_P)
            // Soft rule: warn the user, but allow override via confirmation.
            if (fontSize === '52' || fontSize === '22') {
                const px = fontSize + 'px';
                const existsCount = countInlineFontSize(px);
                const alreadyInside = selectionWithinInlineFontSize(px);
                if (existsCount >= 1 && !alreadyInside) {
                    const baseMsg = fontSize === '52'
                        ? "Titre_P ne peut être utilisé qu'une seule fois dans ce document."
                        : "Sous_titre_P ne peut être utilisé qu'une seule fois dans ce document.";
                    const proceed = window.confirm(baseMsg + "\n\nCliquez sur OK pour l'appliquer quand même, ou sur Annuler pour choisir une autre taille.");
                    if (!proceed) {
                        try { document.getElementById('fontSize').value = ''; } catch (_) {}
                        fontSizeAppliedRecently = false;
                        return; // user chose to respect the single-use rule
                    }
                }
            }

            // If collapsed selection, apply to the nearest block so one click works
            if (selection.isCollapsed) {
                let node = range.startContainer;
                if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
                const targetEl = node && node.closest && node.closest('h1,h2,h3,h4,h5,h6,p,div,span,a');
                if (targetEl) {
                    // If target is a link (e.g., CTA button), wrap its contents in a span
                    // so the font-size applies to the inner text and is not impacted by
                    // the anchor's own CSS (which may use !important).
                    const applyToAnchorContents = (px) => {
                        const span = document.createElement('span');
                        span.style.fontSize = px;
                        // Move existing children into the span
                        while (targetEl.firstChild) {
                            span.appendChild(targetEl.firstChild);
                        }
                        targetEl.appendChild(span);
                    };

                    if (fontSize === '52') {
                        if (targetEl.tagName === 'A') {
                            applyToAnchorContents('52px');
                            const inner = targetEl.querySelector('span');
                            if (inner) {
                                inner.style.lineHeight = '1.2';
                                inner.style.fontWeight = 'bold';
                            }
                        } else {
                            targetEl.style.fontSize = '52px';
                            targetEl.style.lineHeight = '1.2';
                            targetEl.style.fontWeight = 'bold';
                        }
                    } else {
                        if (targetEl.tagName === 'A') {
                            applyToAnchorContents(fontSize + 'px');
                        } else {
                            targetEl.style.fontSize = fontSize + 'px';
                        }
                    }
                    this.saveSelection();
                    this.saveState();
                    try { document.getElementById('fontSize').value = ''; } catch (_) {}
                    setTimeout(() => { fontSizeAppliedRecently = false; }, 0);
                    return;
                }
            }

            if (fontSize === '52') {
                try {
                    const span = document.createElement('span');
                    span.style.fontSize = '52px';
                    span.style.lineHeight = '1.2';
                    span.style.fontWeight = 'bold';
                    span.appendChild(range.extractContents());
                    range.insertNode(span);
                    const newRange = document.createRange();
                    newRange.selectNodeContents(span);
                    selection.removeAllRanges();
                    selection.addRange(newRange);
                    this.saveSelection();
                    this.saveState();
                    try { document.getElementById('fontSize').value = ''; } catch (_) {}
                    setTimeout(() => { fontSizeAppliedRecently = false; }, 0);
                } catch (ex) {
                    console.error('Error applying Titre_P style:', ex);
                    fontSizeAppliedRecently = false;
                }
            } else {
                try {
                    const span = document.createElement('span');
                    span.style.fontSize = fontSize + 'px';
                    span.appendChild(range.extractContents());
                    range.insertNode(span);
                    const newRange = document.createRange();
                    newRange.selectNodeContents(span);
                    selection.removeAllRanges();
                    selection.addRange(newRange);
                    this.saveSelection();
                    this.saveState();
                    try { document.getElementById('fontSize').value = ''; } catch (_) {}
                    setTimeout(() => { fontSizeAppliedRecently = false; }, 0);
                } catch (ex) {
                    console.error('Error applying font size:', ex);
                    fontSizeAppliedRecently = false;
                }
            }
        };

        const fontSizeSelect = document.getElementById('fontSize');
        // Preserve selection right before opening the dropdown
        fontSizeSelect.addEventListener('mousedown', () => this.saveSelection());
        fontSizeSelect.addEventListener('focus', () => this.saveSelection());
        // Apply on input (fires immediately when value changes) and on change (fallback)
        fontSizeSelect.addEventListener('input', (e) => applyFontSize(e.target.value));
        fontSizeSelect.addEventListener('change', (e) => applyFontSize(e.target.value));
        // Some browsers won't fire change when re-selecting the same option.
        // Fallback: on pointerup (option chosen), if no recent apply happened, apply current value.
        fontSizeSelect.addEventListener('pointerup', () => {
            setTimeout(() => {
                if (!fontSizeAppliedRecently) {
                    const v = fontSizeSelect.value;
                    if (v) applyFontSize(v);
                }
            }, 0);
        });

        // Reflect actual selection font-size in the dropdown
        const optionValues = new Set(['52','48','36','26','22']);
        const mapPxToOptionValue = (px) => {
            const n = parseInt(px, 10);
            if (!Number.isFinite(n)) return '';
            if (n === 52) return '52';
            if (n === 48) return '48';
            if (n === 36) return '36';
            if (n === 26) return '26';
            if (n === 22) return '22';
            return '';
        };
        const isInsideEditable = (node) => {
            try {
                const root = document.getElementById('editableContent');
                if (!root || !node) return false;
                if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
                return !!(node && root.contains(node));
            } catch (_) { return false; }
        };
        const getExplicitInlineSizePx = (startNode) => {
            try {
                const root = document.getElementById('editableContent');
                let node = startNode && (startNode.nodeType === Node.TEXT_NODE ? startNode.parentElement : startNode);
                while (node && node !== root) {
                    if (node.style && node.style.fontSize) return node.style.fontSize; // e.g., '52px'
                    node = node.parentElement;
                }
                return '';
            } catch (_) { return ''; }
        };
        let selectionReflectRaf = null;
        const reflectSelectionFontSize = () => {
            if (typeof isInteractingWithToolbar !== 'undefined' && isInteractingWithToolbar) return;
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) return;
            const range = sel.getRangeAt(0);
            const root = document.getElementById('editableContent');
            if (!isInsideEditable(range.commonAncestorContainer)) return;

            // If collapsed, use innermost explicit inline size first; fallback to nearest computed
            if (sel.isCollapsed) {
                let node = range.commonAncestorContainer;
                const explicit = getExplicitInlineSizePx(node);
                if (explicit) {
                    const val = mapPxToOptionValue(explicit);
                    if (fontSizeSelect && fontSizeSelect.value !== val) fontSizeSelect.value = val;
                    return;
                }
                let el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
                if (el && el.closest) {
                    const nearest = el.closest('span, a, p, h1, h2, h3, h4, h5, h6, div');
                    if (nearest) el = nearest;
                }
                const fs = window.getComputedStyle(el).fontSize;
                const val = mapPxToOptionValue(fs);
                if (fontSizeSelect && fontSizeSelect.value !== val) fontSizeSelect.value = val;
                return;
            }

            // Non-collapsed: gather sizes of text nodes intersecting the selection
            const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
            const sizes = new Set();
            while (walker.nextNode()) {
                const tn = walker.currentNode;
                if (!tn.nodeValue || !tn.nodeValue.trim()) continue;
                if (!range.intersectsNode(tn)) continue;
                const explicit = getExplicitInlineSizePx(tn);
                const fs = explicit || window.getComputedStyle(tn.parentElement || tn).fontSize;
                const mapped = mapPxToOptionValue(fs);
                if (mapped) sizes.add(mapped);
                else sizes.add('other');
                if (sizes.size > 1) break; // mixed sizes
            }
            let val = '';
            if (sizes.size === 1) {
                const only = sizes.values().next().value;
                val = optionValues.has(only) ? only : '';
            } else {
                val = '';
            }
            if (fontSizeSelect && fontSizeSelect.value !== val) fontSizeSelect.value = val;
        };
        document.addEventListener('selectionchange', () => {
            // Throttle with rAF
            if (selectionReflectRaf) cancelAnimationFrame(selectionReflectRaf);
            selectionReflectRaf = requestAnimationFrame(reflectSelectionFontSize);
        });

        // Line height handlers — similar behavior to font size
        const applyLineHeight = (lh) => {
            if (!lh) return; // ignore placeholder option
            this.restoreSelection();
            const selection = window.getSelection();
            if (!selection || selection.rangeCount === 0) return;
            const range = selection.getRangeAt(0);

            // If collapsed, apply to nearest block element (avoid inline spans)
            if (selection.isCollapsed) {
                let node = range.startContainer;
                if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
                const targetEl = node && node.closest && node.closest('h1,h2,h3,h4,h5,h6,p,div,li');
                if (targetEl) {
                    // Apply to the target
                    targetEl.style.lineHeight = lh;
                    // Also apply to indent wrapper, if present, so wrapper doesn't block effect
                    const wrap = targetEl.querySelector && targetEl.querySelector('span[data-indent-wrap="1"]');
                    if (wrap) wrap.style.lineHeight = lh;
                    this.saveSelection();
                    this.saveState();
                    return;
                }
            }

            // Non-collapsed: apply to all intersecting block elements to affect full lines
            try {
                let hostNode = range.commonAncestorContainer;
                if (hostNode.nodeType === Node.TEXT_NODE) hostNode = hostNode.parentElement;
                const root = document.getElementById('editableContent');
                const blocks = root ? root.querySelectorAll('h1,h2,h3,h4,h5,h6,p,div,li') : [];
                let applied = 0;
                blocks.forEach(b => {
                    try {
                        if (range.intersectsNode(b)) {
                            b.style.lineHeight = lh;
                            const wrap = b.querySelector && b.querySelector('span[data-indent-wrap="1"]');
                            if (wrap) wrap.style.lineHeight = lh;
                            applied++;
                        }
                    } catch (_) {}
                });
                if (!applied) {
                    // Fallback: wrap selection if no block found
                    const span = document.createElement('span');
                    span.style.lineHeight = lh;
                    span.appendChild(range.extractContents());
                    range.insertNode(span);
                    const newRange = document.createRange();
                    newRange.selectNodeContents(span);
                    selection.removeAllRanges();
                    selection.addRange(newRange);
                }
                this.saveSelection();
                this.saveState();
            } catch (ex) {
                console.error('Error applying line height:', ex);
            }
        };

        const lineHeightSelect = document.getElementById('lineHeight');
        if (lineHeightSelect) {
            // Preserve selection right before opening the dropdown
            lineHeightSelect.addEventListener('mousedown', () => this.saveSelection());
            lineHeightSelect.addEventListener('focus', () => this.saveSelection());
            lineHeightSelect.addEventListener('input', (e) => applyLineHeight(e.target.value));
            lineHeightSelect.addEventListener('change', (e) => applyLineHeight(e.target.value));
        }

        // Text formatting buttons
        document.getElementById('boldBtn').addEventListener('click', () => {
            execWithRestore('bold');
        });

        document.getElementById('italicBtn').addEventListener('click', () => {
            execWithRestore('italic');
        });

        document.getElementById('underlineBtn').addEventListener('click', () => {
            execWithRestore('underline');
        });

        document.getElementById('strikeBtn').addEventListener('click', () => {
            execWithRestore('strikeThrough');
        });

        // Color pickers
        const textColorPickerEl = document.getElementById('textColorPicker');
        if (textColorPickerEl) {
            textColorPickerEl.addEventListener('input', (e) => {
                execWithRestore('foreColor', false, e.target.value);
                const ic = document.querySelector('#textColorDropdownBtn i');
                if (ic) ic.style.color = e.target.value;
            });
        }

        const bgColorPickerEl = document.getElementById('bgColorPicker');
        if (bgColorPickerEl) {
            bgColorPickerEl.addEventListener('input', (e) => {
                const bgCmd = document.queryCommandSupported && document.queryCommandSupported('hiliteColor') ? 'hiliteColor' : 'backColor';
                execWithRestore(bgCmd, false, e.target.value);
                const ic = document.querySelector('#bgColorDropdownBtn i');
                if (ic) ic.style.backgroundColor = e.target.value;
            });
        }

        // Color palettes
        document.getElementById('textColorPalette').addEventListener('click', (e) => {
            if (e.target.classList.contains('palette-color')) {
                const color = e.target.dataset.color;
                execWithRestore('foreColor', false, color);
                const p = document.getElementById('textColorPicker'); if (p) p.value = color;
                document.querySelector('#textColorDropdownBtn i').style.color = color;
                document.getElementById('textColorDropdownContent').style.display = 'none';
            }
        });

        // Text primary colors (single row)
        const textPrimary = document.getElementById('textColorPrimaryPalette');
        if (textPrimary) {
            textPrimary.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    execWithRestore('foreColor', false, color);
                    const p = document.getElementById('textColorPicker');
                    if (p) p.value = color;
                    const ic = document.querySelector('#textColorDropdownBtn i');
                    if (ic) ic.style.color = color;
                    const dd = document.getElementById('textColorDropdownContent');
                    if (dd) dd.style.display = 'none';
                }
            });
        }

        // Text standard colors (single row)
        const textStd = document.getElementById('textColorStandardPalette');
        if (textStd) {
            textStd.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    execWithRestore('foreColor', false, color);
                    const p = document.getElementById('textColorPicker');
                    if (p) p.value = color;
                    const ic = document.querySelector('#textColorDropdownBtn i');
                    if (ic) ic.style.color = color;
                    const dd = document.getElementById('textColorDropdownContent');
                    if (dd) dd.style.display = 'none';
                }
            });
        }

        document.getElementById('bgColorPalette').addEventListener('click', (e) => {
            if (e.target.classList.contains('palette-color')) {
                const color = e.target.dataset.color;
                const bgCmd = document.queryCommandSupported && document.queryCommandSupported('hiliteColor') ? 'hiliteColor' : 'backColor';
                execWithRestore(bgCmd, false, color);
                const p = document.getElementById('bgColorPicker'); if (p) p.value = color;
                document.querySelector('#bgColorDropdownBtn i').style.backgroundColor = color;
                document.getElementById('bgColorDropdownContent').style.display = 'none';
            }
        });

        // Background primary colors
        const bgPrimary = document.getElementById('bgColorPrimaryPalette');
        if (bgPrimary) {
            bgPrimary.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    const bgCmd = document.queryCommandSupported && document.queryCommandSupported('hiliteColor') ? 'hiliteColor' : 'backColor';
                    execWithRestore(bgCmd, false, color);
                    const p = document.getElementById('bgColorPicker');
                    if (p) p.value = color;
                    const ic = document.querySelector('#bgColorDropdownBtn i');
                    if (ic) ic.style.backgroundColor = color;
                    const dd = document.getElementById('bgColorDropdownContent');
                    if (dd) dd.style.display = 'none';
                }
            });
        }

        // Background standard colors
        const bgStd = document.getElementById('bgColorStandardPalette');
        if (bgStd) {
            bgStd.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    const bgCmd = document.queryCommandSupported && document.queryCommandSupported('hiliteColor') ? 'hiliteColor' : 'backColor';
                    execWithRestore(bgCmd, false, color);
                    const p = document.getElementById('bgColorPicker');
                    if (p) p.value = color;
                    const ic = document.querySelector('#bgColorDropdownBtn i');
                    if (ic) ic.style.backgroundColor = color;
                    const dd = document.getElementById('bgColorDropdownContent');
                    if (dd) dd.style.display = 'none';
                }
            });
        }

        // Color dropdowns
        const textDropdownBtn = document.getElementById('textColorDropdownBtn');
        textDropdownBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const content = document.getElementById('textColorDropdownContent');
            // Flip up if not enough space below
            const dd = textDropdownBtn.closest('.dropdown');
            if (dd && content) {
                content.style.display = 'block';
                const btnRect = textDropdownBtn.getBoundingClientRect();
                const spaceBelow = (window.innerHeight - btnRect.bottom);
                const needed = content.offsetHeight + 12;
                if (spaceBelow < needed) dd.classList.add('drop-up'); else dd.classList.remove('drop-up');
                // Toggle after measurement
                content.style.display = (content.style.display === 'none' ? 'block' : content.style.display);
                if (content.style.display === 'block' && dd.classList.contains('drop-up')) {
                    // keep open; clicking again will close
                }
            } else if (content) {
                content.style.display = content.style.display === 'none' ? 'block' : 'none';
            }
            document.getElementById('bgColorDropdownContent').style.display = 'none';
        });

        const bgDropdownBtn = document.getElementById('bgColorDropdownBtn');
        bgDropdownBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const content = document.getElementById('bgColorDropdownContent');
            const dd = bgDropdownBtn.closest('.dropdown');
            if (dd && content) {
                content.style.display = 'block';
                const btnRect = bgDropdownBtn.getBoundingClientRect();
                const spaceBelow = (window.innerHeight - btnRect.bottom);
                const needed = content.offsetHeight + 12;
                if (spaceBelow < needed) dd.classList.add('drop-up'); else dd.classList.remove('drop-up');
                content.style.display = (content.style.display === 'none' ? 'block' : content.style.display);
            } else if (content) {
                content.style.display = content.style.display === 'none' ? 'block' : 'none';
            }
            document.getElementById('textColorDropdownContent').style.display = 'none';
        });

        // Alignment buttons
        document.getElementById('alignLeftBtn').addEventListener('click', () => {
            execWithRestore('justifyLeft');
            setListMarkerPositionForSelection('left');
            try { this.saveState(); } catch (_) {}
        });

        document.getElementById('alignCenterBtn').addEventListener('click', () => {
            execWithRestore('justifyCenter');
            setListMarkerPositionForSelection('center');
            try { this.saveState(); } catch (_) {}
        });

        document.getElementById('alignRightBtn').addEventListener('click', () => {
            execWithRestore('justifyRight');
            setListMarkerPositionForSelection('right');
            try { this.saveState(); } catch (_) {}
        });

        document.getElementById('alignJustifyBtn').addEventListener('click', () => {
            execWithRestore('justifyFull');
            setListMarkerPositionForSelection('justify');
            try { this.saveState(); } catch (_) {}
        });

        // List buttons
        document.getElementById('bulletListBtn').addEventListener('click', () => {
            execWithRestore('insertUnorderedList');
        });

        document.getElementById('numberListBtn').addEventListener('click', () => {
            execWithRestore('insertOrderedList');
        });

        // Helper: pick a safe inner text block to nudge
        const getIndentTargetBlock = () => {
            const sel = window.getSelection();
            if (!sel || sel.rangeCount === 0) return null;
            const range = sel.getRangeAt(0);
            let node = range.startContainer;
            if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
            const editableHost = document.getElementById('editableContent');
            if (!editableHost) return null;
            // If inside list item, return that LI to allow style-based nudge
            if (node.closest && node.closest('li')) {
                const li = node.closest('li');
                return li && editableHost.contains(li) ? li : null;
            }
            // If inside a list but not in an LI (e.g., cursor around the list), target UL/OL as group
            if (node.closest) {
                const listEl = node.closest('ul,ol');
                if (listEl && editableHost.contains(listEl)) return listEl;
            }
            // Find closest text block (prioritize text elements)
            let block = node.closest && node.closest('p,h1,h2,h3,h4,h5,h6,li');
            // If not found, allow a safe DIV that is not a layout container
            if ((!block) && node.closest) {
                const candidateDiv = node.closest('div');
                const forbiddenDivs = ['newsletter-section','two-column-layout','gallery-section','syc-item','newsletter-container','column'];
                if (candidateDiv && editableHost.contains(candidateDiv) && candidateDiv !== editableHost) {
                    const isForbidden = forbiddenDivs.some(cls => candidateDiv.classList && candidateDiv.classList.contains(cls));
                    if (!isForbidden) block = candidateDiv;
                }
            }
            if (!block || !editableHost.contains(block)) return null;
            // Never target the editable host itself
            if (block === editableHost) return null;
            // Avoid large layout containers
            const forbidden = ['newsletter-section','two-column-layout','gallery-section','syc-item','newsletter-container'];
            if (forbidden.some(cls => block.classList && block.classList.contains(cls))) return null;
            return block;
        };

        document.getElementById('indentBtn').addEventListener('click', () => {
            // If inside a list, keep native behavior (nest list items)
            this.restoreSelection();
            const sel = window.getSelection();
            const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
            const root = document.getElementById('editableContent');
            if (range && root) {
                const lis = Array.from(root.querySelectorAll('li'));
                const hits = lis.filter(li => {
                    try { return range.intersectsNode(li); } catch (_) { return false; }
                });
                if (hits.length > 1) {
                    hits.forEach(li => {
                        const cur = parseFloat(window.getComputedStyle(li).marginLeft) || 0;
                        const next = Math.min(cur + 32, 320);
                        li.style.marginLeft = next + 'px';
                    });
                    this.saveSelection();
                    this.saveState();
                    return;
                }
            }
            const target = getIndentTargetBlock();
            if (target && (target.tagName === 'LI' || target.tagName === 'UL' || target.tagName === 'OL')) {
                // Nudge the list item or list as a whole to the right (bullet follows)
                const cur = parseFloat(window.getComputedStyle(target).marginLeft) || 0;
                const next = Math.min(cur + 32, 320);
                target.style.marginLeft = next + 'px';
                this.saveSelection();
                this.saveState();
                return;
            }
            if (!target) { execWithRestore('indent'); return; }
            // Apply indent to an inner wrapper so the container layout doesn't shift
            let wrap = target.querySelector('span[data-indent-wrap="1"]');
            if (!wrap) {
                wrap = document.createElement('span');
                wrap.setAttribute('data-indent-wrap', '1');
                wrap.style.display = 'inline-block';
                wrap.style.width = '100%';
                wrap.style.boxSizing = 'border-box';
                // Move existing children into the wrapper
                while (target.firstChild) {
                    wrap.appendChild(target.firstChild);
                }
                target.appendChild(wrap);
            }
            const curPad = parseFloat(window.getComputedStyle(wrap).paddingLeft) || 0;
            const next = Math.min(curPad + 32, 320);
            wrap.style.paddingLeft = next + 'px';
            this.saveSelection();
            this.saveState();
        });

        document.getElementById('outdentBtn').addEventListener('click', () => {
            // If inside a list, keep native behavior (un-nest list items)
            this.restoreSelection();
            const sel = window.getSelection();
            const range = sel && sel.rangeCount > 0 ? sel.getRangeAt(0) : null;
            const root = document.getElementById('editableContent');
            if (range && root) {
                const lis = Array.from(root.querySelectorAll('li'));
                const hits = lis.filter(li => {
                    try { return range.intersectsNode(li); } catch (_) { return false; }
                });
                if (hits.length > 1) {
                    hits.forEach(li => {
                        const cur = parseFloat(window.getComputedStyle(li).marginLeft) || 0;
                        const next = Math.max(cur - 32, 0);
                        li.style.marginLeft = next + 'px';
                    });
                    this.saveSelection();
                    this.saveState();
                    return;
                }
            }
            const target = getIndentTargetBlock();
            if (target && (target.tagName === 'LI' || target.tagName === 'UL' || target.tagName === 'OL')) {
                const cur = parseFloat(window.getComputedStyle(target).marginLeft) || 0;
                const next = Math.max(cur - 32, 0);
                target.style.marginLeft = next + 'px';
                this.saveSelection();
                this.saveState();
                return;
            }
            if (!target) { execWithRestore('outdent'); return; }
            let wrap = target.querySelector('span[data-indent-wrap="1"]');
            if (!wrap) {
                // Nothing to outdent; fallback to native for lists already handled above
                this.saveSelection();
                this.saveState();
                return;
            }
            const curPad = parseFloat(window.getComputedStyle(wrap).paddingLeft) || 0;
            const next = Math.max(curPad - 32, 0);
            wrap.style.paddingLeft = next + 'px';
            if (next === 0) {
                // unwrap to keep DOM clean
                while (wrap.firstChild) {
                    target.insertBefore(wrap.firstChild, wrap);
                }
                wrap.remove();
            }
            this.saveSelection();
            this.saveState();
        });

        // Link button
        document.getElementById('linkBtn').addEventListener('click', () => {
            const url = prompt('Entrez l\'URL du lien:');
            if (url) {
                execWithRestore('createLink', false, url);
            }
        });

        // Remove format button
        document.getElementById('removeFormatBtn').addEventListener('click', () => {
            execWithRestore('removeFormat');
        });

        // Show/hide toolbar on text selection within editable area, and follow selection
        const positionRichToolbar = () => {
            const selection = window.getSelection();
            const hasText = selection && selection.rangeCount > 0 && selection.toString().length > 0;
            const inEditable = hasText && isSelectionInEditable();

            const imageTb = document.getElementById('imageToolbar');
            const sectionTb = document.getElementById('sectionToolbar');
            const videoTb = document.getElementById('videoToolbar');
            const tableTb = document.getElementById('tableToolbar');
            const anotherToolbarOpen =
                (imageTb && imageTb.style.display === 'block') ||
                (sectionTb && sectionTb.style.display === 'block') ||
                (videoTb && videoTb.style.display === 'block') ||
                (tableTb && tableTb.style.display === 'block');
            if (anotherToolbarOpen) { toolbar.style.display = 'none'; return; }

            if (!inEditable) {
                if (isActiveInToolbar() || isInteractingWithToolbar) return;
                toolbar.style.display = 'none';
                return;
            }

            const range = selection.getRangeAt(0);
            const rect = range.getBoundingClientRect();
            const scrollY = window.scrollY || document.documentElement.scrollTop || 0;
            const scrollX = window.scrollX || document.documentElement.scrollLeft || 0;
            const margin = 10;
            toolbar.style.display = 'flex';

            // Prefer above; if not enough space, place below
            let top = rect.top + scrollY - toolbar.offsetHeight - margin;
            if (top < scrollY + 8) top = rect.bottom + scrollY + margin;

            // Clamp horizontally to viewport
            const vw = window.innerWidth || document.documentElement.clientWidth;
            const w = toolbar.offsetWidth || 600;
            let left = rect.left + scrollX;
            if (left < 8) left = 8;
            const maxLeft = scrollX + vw - w - 8;
            if (left > maxLeft) left = Math.max(8, maxLeft);

            toolbar.style.left = left + 'px';
            toolbar.style.top = top + 'px';
            this.saveSelection();
        };

        document.addEventListener('selectionchange', positionRichToolbar);
        window.addEventListener('scroll', positionRichToolbar, { passive: true });
        window.addEventListener('resize', positionRichToolbar);
    }

    // ===== Section Toolbar =====
    showSectionToolbar(sectionEl) {
        // Normalize to the actual section element using closest()
        const sectionSelector = '.newsletter-section, .gallery-section, .two-column-layout, .syc-item, .cta-section';
        const top = (sectionEl && sectionEl.closest && sectionEl.closest(sectionSelector)) || sectionEl;
        this.currentEditingSection = top;
        const toolbar = document.getElementById('sectionToolbar');
        const rect = this.currentEditingSection.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
        toolbar.style.top = `${rect.top + scrollTop - toolbar.offsetHeight - 8}px`;
        toolbar.style.left = `${rect.left + scrollLeft}px`;
        toolbar.style.display = 'block';

        // If this section has a persisted background color, reapply it and sync the icon
        try {
            const savedColor = sectionEl && sectionEl.dataset ? sectionEl.dataset.sectionBg : '';
            if (savedColor) {
                this.reapplySectionBackground(sectionEl);
                const icon = document.querySelector('#sectionBgColorDropdownBtn i');
                if (icon) icon.style.backgroundColor = savedColor;
            }
        } catch (_) { /* no-op */ }

        // Wire once
        if (!this._sectionToolbarWired) {
            const upBtn = document.getElementById('sectionMoveUpBtn');
            const downBtn = document.getElementById('sectionMoveDownBtn');
            const widthBtn = document.getElementById('sectionWidthBtn');
            const widthOptions = document.getElementById('sectionWidthOptions');
            const alignLeftBtn = document.getElementById('sectionAlignLeftBtn');
            const alignCenterBtn = document.getElementById('sectionAlignCenterBtn');
            const alignRightBtn = document.getElementById('sectionAlignRightBtn');
            const deleteBtn = document.getElementById('sectionDeleteBtn');

            upBtn && upBtn.addEventListener('click', () => this.moveSection(-1));
            downBtn && downBtn.addEventListener('click', () => this.moveSection(1));
            widthBtn && widthBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (!widthOptions) return;
                widthOptions.style.display = widthOptions.style.display === 'none' ? 'block' : 'none';
            });
            if (widthOptions) {
                widthOptions.querySelectorAll('.dropdown-item').forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        const size = e.currentTarget.getAttribute('data-size');
                        this.setSectionWidth(size);
                        widthOptions.style.display = 'none';
                    });
                });
            }
            alignLeftBtn && alignLeftBtn.addEventListener('click', () => this.alignSection('left'));
            alignCenterBtn && alignCenterBtn.addEventListener('click', () => this.alignSection('center'));
            alignRightBtn && alignRightBtn.addEventListener('click', () => this.alignSection('right'));
            deleteBtn && deleteBtn.addEventListener('click', () => this.deleteSection());
            // Section background color dropdown wiring
            const secBgBtn = document.getElementById('sectionBgColorDropdownBtn');
            const secBgDropdown = document.getElementById('sectionBgColorDropdownContent');
            const secBgPalette = document.getElementById('sectionBgColorPalette');
            const secBgPrimary = document.getElementById('sectionBgPrimaryPalette');
            const secBgIcon = document.querySelector('#sectionBgColorDropdownBtn i');

            const applySectionBg = (color) => {
                if (!this.currentEditingSection) return;
                this.currentEditingSection.style.background = '';
                this.currentEditingSection.style.backgroundColor = color;
                // Persist chosen color on the section to survive future edits
                this.currentEditingSection.dataset.sectionBg = color;
                // Apply via helper so rules are consistent
                this.reapplySectionBackground(this.currentEditingSection);
                if (secBgIcon) secBgIcon.style.backgroundColor = color;
                this.saveState();
                this.updateLastModified();
                this.autoSaveToLocalStorage();
                this.lastAction = 'Couleur de fond de section modifiée';
            };

            secBgBtn && secBgBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (!secBgDropdown) return;
                secBgDropdown.style.display = secBgDropdown.style.display === 'none' ? 'block' : 'none';
            });

            secBgPalette && secBgPalette.addEventListener('click', (e) => {
                if (e.target.classList && e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    applySectionBg(color);
                    if (secBgDropdown) secBgDropdown.style.display = 'none';
                }
            });
            secBgPrimary && secBgPrimary.addEventListener('click', (e) => {
                if (e.target.classList && e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    applySectionBg(color);
                    if (secBgDropdown) secBgDropdown.style.display = 'none';
                }
            });
            this._sectionToolbarWired = true;
        }
    }

    hideSectionToolbar() {
        const toolbar = document.getElementById('sectionToolbar');
        if (toolbar) toolbar.style.display = 'none';
        this.currentEditingSection = null;
    }

    // ===== Table Toolbar =====
    showTableToolbar(tableEl) {
        return;
    }

    hideTableToolbar() {
        const toolbar = document.getElementById('tableToolbar');
        if (toolbar) toolbar.style.display = 'none';
        this.currentEditingTable = null;
    }

    // ===== Table Context Menu (Right-Click) =====
    setupTableContextMenu() {
        try { } catch (_) { /* no-op */ }
    }

    showTableContextMenu(x, y) {
        const menu = this.tableContextMenu;
        if (!menu) return;
        menu.style.display = 'block';
        // Keep within viewport
        const vw = window.innerWidth, vh = window.innerHeight;
        const rect = menu.getBoundingClientRect();
        const left = Math.min(x, vw - rect.width - 8);
        const top = Math.min(y, vh - rect.height - 8);
        menu.style.left = left + 'px';
        menu.style.top = top + 'px';
        // Hide other toolbars to reduce interference
        this.hideTableToolbar();
    }

    hideTableContextMenu() {
        const menu = this.tableContextMenu;
        if (menu) menu.style.display = 'none';
    }

    // Split current cell into two stacked cells (adds a new row below)
    splitTableCellHorizontally() {
        return;
    }

    // Split current cell into two side-by-side cells (adds a new column after)
    splitTableCellVertically() {
        return;
    }

    // Selection helpers
    selectTableRow() {
        return;
    }

    selectTableColumn() {
        return;
    }

    selectTableElement() {
        return;
    }

    // Visual marking helpers (mirror existing hover/selection styles)
    markTableCell(cell, on) {
        return;
    }

    clearTableCellMarks(table) {
        return;
    }

    // Remove all plain text from the current table (keep elements like images/links)
    clearTablePlainText() {
        return;
    }

    // ===== Webinar Toolbar =====
    showWebinarToolbar(sectionEl) {
        this.currentWebinarSection = sectionEl;
        const toolbar = document.getElementById('webinarToolbar');
        const rect = sectionEl.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
        toolbar.style.top = `${rect.top + scrollTop - toolbar.offsetHeight - 8}px`;
        toolbar.style.left = `${rect.left + (rect.width / 2) - (toolbar.offsetWidth / 2)}px`;
        toolbar.style.display = 'block';

        if (!this._webinarToolbarWired) {
            // Color picker + palette + dropdown (mirrors text toolbar behavior)
            const bgPicker = document.getElementById('webinarBgColorPicker');
            const bgPalette = document.getElementById('webinarBgColorPalette');
            const bgBtnIcon = document.querySelector('#webinarBgColorDropdownBtn i');
            const bgDropdown = document.getElementById('webinarBgColorDropdownContent');
            const bgBtn = document.getElementById('webinarBgColorDropdownBtn');
            const bgStdPalette = document.getElementById('webinarBgStandardPalette');
            const bgPrimaryPalette = document.getElementById('webinarBgPrimaryPalette');

            const applyWebinarBg = (color) => {
                if (!this.currentWebinarSection) return;
                this.currentWebinarSection.style.background = '';
                this.currentWebinarSection.style.backgroundColor = color;
                if (bgPicker) bgPicker.value = color;
                if (bgBtnIcon) bgBtnIcon.style.backgroundColor = color;
                this.saveState();
            };

            bgPicker && bgPicker.addEventListener('input', (e) => applyWebinarBg(e.target.value));

            bgPalette && bgPalette.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    applyWebinarBg(color);
                    if (bgDropdown) bgDropdown.style.display = 'none';
                }
            });

            // Standard colors for webinar background
            bgStdPalette && bgStdPalette.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    applyWebinarBg(color);
                    if (bgDropdown) bgDropdown.style.display = 'none';
                }
            });

            // Primary colors for webinar background
            bgPrimaryPalette && bgPrimaryPalette.addEventListener('click', (e) => {
                if (e.target.classList.contains('palette-color')) {
                    const color = e.target.dataset.color;
                    applyWebinarBg(color);
                    if (bgDropdown) bgDropdown.style.display = 'none';
                }
            });

            bgBtn && bgBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (!bgDropdown) return;
                bgDropdown.style.display = bgDropdown.style.display === 'none' ? 'block' : 'none';
            });

            document.addEventListener('click', (e) => {
                if (bgDropdown && !e.target.closest('#webinarBgColorDropdownContent') && !e.target.closest('#webinarBgColorDropdownBtn')) {
                    bgDropdown.style.display = 'none';
                }
            });

            const imgBtn = document.getElementById('webinarInsertImageBtn');
            const imgInput = document.getElementById('webinarImageInput');
            imgBtn && imgBtn.addEventListener('click', (ev) => { ev.stopPropagation(); imgInput && imgInput.click(); });
            imgInput && imgInput.addEventListener('change', async (ev) => {
                const file = ev.target.files && ev.target.files[0];
                if (!file) return;

                // Capture the current mouse position for image insertion
                const sel = window.getSelection();
                if (sel.rangeCount > 0) {
                    this.lastMouseRange = sel.getRangeAt(0);
                }

                try {
                    const dataUrl = await this.compressImageFile(file, { maxWidth: 1600, maxHeight: 1200, quality: 0.9 });
                    this.insertImage(dataUrl, file.name);
                } catch (_) {
                    const reader = new FileReader();
                    reader.onload = () => this.insertImage(reader.result, file.name);
                    reader.readAsDataURL(file);
                }
                ev.target.value = '';
            });

            const videoUrlBtn = document.getElementById('webinarInsertVideoUrlBtn');
            videoUrlBtn && videoUrlBtn.addEventListener('click', () => {
                const url = prompt('Entrez l\'URL de la vidéo (YouTube, Vimeo, MP4)');
                if (!url) return;
                if (this.currentWebinarSection) {
                    const r = document.createRange();
                    r.selectNodeContents(this.currentWebinarSection);
                    r.collapse(false);
                    this.lastMouseRange = r;
                }
                this.insertVideo(url);
            });

            const videoLocalBtn = document.getElementById('webinarInsertVideoLocalBtn');
            const videoInput = document.getElementById('webinarVideoInput');
            videoLocalBtn && videoLocalBtn.addEventListener('click', (ev) => { ev.stopPropagation(); videoInput && videoInput.click(); });
            videoInput && videoInput.addEventListener('change', (ev) => {
                const file = ev.target.files && ev.target.files[0];
                if (!file) return;
                const url = URL.createObjectURL(file);
                if (this.currentWebinarSection) {
                    const r = document.createRange();
                    r.selectNodeContents(this.currentWebinarSection);
                    r.collapse(false);
                    this.lastMouseRange = r;
                }
                this.insertVideo(url);
                ev.target.value = '';
            });

            // Delete webinar section
            const deleteBtn = document.getElementById('webinarDeleteBtn');
            deleteBtn && deleteBtn.addEventListener('click', () => {
                if (this.currentWebinarSection && confirm('Supprimer cette section Webinar ?')) {
                    this.currentWebinarSection.remove();
                    this.hideWebinarToolbar();
                    this.saveState();
                    this.updateLastModified();
                    this.autoSaveToLocalStorage();
                }
            });

            this._webinarToolbarWired = true;
        }
    }

    hideWebinarToolbar() {
        const toolbar = document.getElementById('webinarToolbar');
        if (toolbar) toolbar.style.display = 'none';
        this.currentWebinarSection = null;
    }

    moveSection(direction) {
        if (!this.currentEditingSection) return;
        // Always work with the actual section element
        const sectionSelector = '.newsletter-section, .gallery-section, .two-column-layout, .syc-item, .cta-section';
        const section = (this.currentEditingSection.closest && this.currentEditingSection.closest(sectionSelector)) || this.currentEditingSection;
        this.currentEditingSection = section;

        const isSection = (el) => !!(el && el.classList && (
            el.classList.contains('newsletter-section') ||
            el.classList.contains('gallery-section') ||
            el.classList.contains('two-column-layout') ||
            el.classList.contains('syc-item') ||
            el.classList.contains('cta-section')
        ));

        // Build an ordered list of section siblings within the same parent
        const parent = section && section.parentNode;
        if (!parent) return;
        const siblings = Array.from(parent.children).filter(isSection);
        const idx = siblings.indexOf(section);
        if (idx === -1) return;

        if (direction < 0 && idx > 0) {
            const target = siblings[idx - 1];
            parent.insertBefore(section, target); // place just before the immediate previous eligible
        } else if (direction > 0 && idx < siblings.length - 1) {
            const target = siblings[idx + 1];
            // Move down by inserting this section after the next eligible sibling
            parent.insertBefore(section, target.nextSibling);
        }
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Section déplacée';
        // Keep it in view after moving
        try {
            section.scrollIntoView({ behavior: 'smooth', block: 'center' });
        } catch (_) {}
        this.showSectionToolbar(section);
    }

    alignSection(alignment) {
        if (!this.currentEditingSection) return;
        const section = this.currentEditingSection;
        section.style.marginLeft = '';
        section.style.marginRight = '';
        section.style.textAlign = '';

        switch (alignment) {
            case 'left':
                section.style.marginLeft = '0';
                section.style.marginRight = 'auto';
                section.style.textAlign = 'left';
                break;
            case 'center':
                section.style.marginLeft = 'auto';
                section.style.marginRight = 'auto';
                section.style.textAlign = 'center';
                break;
            case 'right':
                section.style.marginLeft = 'auto';
                section.style.marginRight = '0';
                section.style.textAlign = 'right';
                break;
        }
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Section alignée';
        this.showSectionToolbar(section);
    }

    async deleteSection() {
        if (!this.currentEditingSection) return;
        if (await this.confirmWithCancel('Supprimer cette section ?')) {
            this.currentEditingSection.remove();
            this.hideSectionToolbar();
            this.saveState();
            this.updateLastModified();
            this.autoSaveToLocalStorage();
            this.lastAction = 'Section supprimée';
        }
    }

    setSectionWidth(size) {
        if (!this.currentEditingSection) return;
        const section = this.currentEditingSection;
        // preserve alignment margins; only change width/max-width
        if (size === 'auto') {
            section.style.width = '';
            section.style.maxWidth = '';
        } else {
            const pct = parseInt(size, 10);
            section.style.width = pct + '%';
            section.style.maxWidth = '100%';
        }
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Largeur de section ajustée';
        this.showSectionToolbar(section);
    }

    // ===== Video Toolbar =====
    showVideoToolbar(videoEl) {
        this.currentEditingVideo = videoEl;
        const toolbar = document.getElementById('videoToolbar');
        const rect = videoEl.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
        toolbar.style.top = `${rect.top + scrollTop - toolbar.offsetHeight - 8}px`;
        toolbar.style.left = `${rect.left + scrollLeft}px`;
        toolbar.style.display = 'block';

        if (!this._videoToolbarWired) {
            document.getElementById('videoSizeBtn').addEventListener('click', (e) => {
                e.stopPropagation();
                const dd = document.getElementById('videoSizeOptions');
                dd.style.display = dd.style.display === 'none' ? 'block' : 'none';
            });
            document.querySelectorAll('#videoSizeOptions .dropdown-item').forEach(btn => {
                btn.addEventListener('click', (e) => {
                    const size = e.currentTarget.getAttribute('data-size');
                    this.setVideoSize(size);
                    document.getElementById('videoSizeOptions').style.display = 'none';
                });
            });
            document.getElementById('videoAlignLeftBtn').addEventListener('click', () => this.alignVideo('left'));
            document.getElementById('videoAlignCenterBtn').addEventListener('click', () => this.alignVideo('center'));
            document.getElementById('videoAlignRightBtn').addEventListener('click', () => this.alignVideo('right'));
            document.getElementById('videoDeleteBtn').addEventListener('click', async () => {
                if (await this.confirmWithCancel('Supprimer cette vidéo ?')) {
                    this.deleteVideo();
                }
            });
            this._videoToolbarWired = true;
        }
    }

    hideVideoToolbar() {
        const toolbar = document.getElementById('videoToolbar');
        if (toolbar) toolbar.style.display = 'none';
        this.currentEditingVideo = null;
    }

    setVideoSize(size) {
        if (!this.currentEditingVideo) return;
        const el = this.currentEditingVideo;
        if (size === 'auto') {
            el.style.width = '';
            el.style.maxWidth = '100%';
            el.style.height = '';
        } else {
            const pct = parseInt(size, 10);
            el.style.width = pct + '%';
            el.style.height = 'auto';
            el.style.maxWidth = '100%';
        }
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Taille de vidéo ajustée';
        this.showVideoToolbar(el);
    }

    alignVideo(alignment) {
        if (!this.currentEditingVideo) return;
        const el = this.currentEditingVideo;
        const wrapper = el.parentElement && el.parentElement.classList.contains('video-align-wrapper') ? el.parentElement : null;
        const container = wrapper || el;
        const baseMargin = '10px';
        container.style.display = 'block';
        container.style.marginTop = baseMargin;
        container.style.marginBottom = baseMargin;
        container.style.marginLeft = '';
        container.style.marginRight = '';
        container.style.textAlign = '';
        el.style.marginTop = '';
        el.style.marginBottom = '';
        el.style.marginLeft = '';
        el.style.marginRight = '';

        switch (alignment) {
            case 'left':
                container.style.marginLeft = '0';
                container.style.marginRight = 'auto';
                container.style.textAlign = 'left';
                el.style.display = 'block';
                el.style.marginLeft = '0';
                el.style.marginRight = 'auto';
                break;
            case 'center':
                container.style.marginLeft = 'auto';
                container.style.marginRight = 'auto';
                container.style.textAlign = 'center';
                el.style.display = 'block';
                el.style.marginLeft = 'auto';
                el.style.marginRight = 'auto';
                break;
            case 'right':
                container.style.marginLeft = 'auto';
                container.style.marginRight = '0';
                container.style.textAlign = 'right';
                el.style.display = 'block';
                el.style.marginLeft = 'auto';
                el.style.marginRight = '0';
                break;
        }
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Vidéo alignée';
        this.showVideoToolbar(el);
    }

    async deleteVideo() {
        if (!this.currentEditingVideo) return;
        const node = this.currentEditingVideo;
        // If iframe inside a wrapper, remove wrapper; otherwise remove node
        const parent = node.parentElement;
        if (parent && parent.classList.contains('video-align-wrapper')) {
            parent.remove();
        } else {
            // Fallback: remove the image directly
            node.remove();
        }
        
        this.hideVideoToolbar();
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Vidéo supprimée';
    }

    // Selection helpers
    saveSelection() {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return;
        const range = selection.getRangeAt(0);
        // Only persist selections that are inside the editable content area
        try {
            const startNode = range.startContainer;
            const el = (startNode && startNode.nodeType === Node.ELEMENT_NODE)
                ? startNode
                : (startNode && startNode.parentElement);
            if (el && el.closest && el.closest('#editableContent')) {
                this.savedSelection = range.cloneRange();
            }
        } catch (_) {
            // Fallback: do not update savedSelection if we cannot verify context
        }
    }

    // Insert raw HTML at the current selection inside #editableContent
    insertHTMLAtCursor(html) {
        const selection = window.getSelection();
        if (selection && selection.rangeCount > 0) {
            const range = selection.getRangeAt(0);
            // Ensure the insertion happens inside the editor
            const node = range.startContainer;
            const host = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
            if (host && host.closest && host.closest('#editableContent')) {
                range.deleteContents();
                const div = document.createElement('div');
                div.innerHTML = html;
                const fragment = document.createDocumentFragment();
                while (div.firstChild) {
                    fragment.appendChild(div.firstChild);
                }
                range.insertNode(fragment);
                return;
            }
        }
        // Fallback append to editor
        const editable = document.getElementById('editableContent');
        editable.innerHTML += html;
    }

    // Insert a DOM element at the current cursor or last mouse position range
    insertElementAtCursor(element) {
        const selection = window.getSelection();
        const editable = document.getElementById('editableContent');
        let range = null;

        // Prefer lastMouseRange if it points inside the editor
        if (this.lastMouseRange) {
            try {
                const node = this.lastMouseRange.startContainer;
                const host = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
                if (host && host.closest && host.closest('#editableContent')) {
                    range = this.lastMouseRange.cloneRange();
                }
            } catch (_) {}
        }

        // If no saved range, try to compute from the last mouse position inside the editor
        if (!range && this.lastMousePosition && typeof this.lastMousePosition.x === 'number') {
            try {
                const r = this.computeRangeFromPoint(this.lastMousePosition.x, this.lastMousePosition.y);
                if (r) range = r.cloneRange ? r.cloneRange() : r; // support native Range
            } catch (_) {}
        }

        // Otherwise, use current selection if inside editor
        if (!range && selection && selection.rangeCount > 0) {
            const r = selection.getRangeAt(0);
            const node = r.startContainer;
            const host = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
            if (host && host.closest && host.closest('#editableContent')) {
                range = r.cloneRange();
            }
        }

        if (range) {
            range.collapse(true);
            range.insertNode(element);
            // Place caret after inserted element to keep typing natural
            const spacer = document.createElement('p');
            spacer.innerHTML = '<br>';
            element.after(spacer);
            const newRange = document.createRange();
            newRange.selectNodeContents(spacer);
            newRange.collapse(true);
            selection.removeAllRanges();
            selection.addRange(newRange);
            this.lastMouseRange = null;
        } else {
            // Append at end as a safe fallback
            editable.appendChild(element);
            const spacer = document.createElement('p');
            spacer.innerHTML = '<br>';
            editable.appendChild(spacer);
            const newRange = document.createRange();
            newRange.selectNodeContents(spacer);
            newRange.collapse(true);
            selection.removeAllRanges();
            selection.addRange(newRange);
            this.lastMouseRange = null;
        }
    }

    // Compute an insertion Range from viewport coordinates within the editor
    computeRangeFromPoint(x, y) {
        const editableEl = document.getElementById('editableContent');
        if (!editableEl) return null;
        const containerRect = editableEl.getBoundingClientRect();
        if (x < containerRect.left || x > containerRect.right || y < containerRect.top || y > containerRect.bottom) {
            return null;
        }

        const snapRangeToBlockBoundary = (node) => {
            if (!node) return null;
            let child = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
            while (child && child.parentElement !== editableEl) {
                child = child.parentElement;
            }
            const range = document.createRange();
            if (!child || child === editableEl) {
                const last = editableEl.lastChild;
                if (last) range.setStartAfter(last); else range.setStart(editableEl, editableEl.childNodes.length);
                range.collapse(true);
                return range;
            }
            const childRect = child.getBoundingClientRect();
            const placeAfter = y > (childRect.top + childRect.height / 2);
            if (placeAfter) range.setStartAfter(child); else range.setStartBefore(child);
            range.collapse(true);
            return range;
        };

        if (document.caretRangeFromPoint) {
            const r = document.caretRangeFromPoint(x, y);
            if (r) return snapRangeToBlockBoundary(r.startContainer);
        }
        if (document.caretPositionFromPoint) {
            const pos = document.caretPositionFromPoint(x, y);
            if (pos && pos.offsetNode != null) {
                return snapRangeToBlockBoundary(pos.offsetNode);
            }
        }
        const target = document.elementFromPoint(x, y);
        if (!target) return null;
        const container = target.closest('#editableContent');
        if (!container) return null;
        return snapRangeToBlockBoundary(target);
    }
    restoreSelection() {
        if (!this.savedSelection) return;
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(this.savedSelection);
    }

    // Reapply a section's persisted background color to itself and common inner containers.
    // Also normalize inner editable content to transparent so white inline highlights do not show as bars.
    reapplySectionBackground(sectionEl) {
        if (!sectionEl) return;
        const color = sectionEl.dataset ? sectionEl.dataset.sectionBg : '';
        if (!color) return;
        try {
            sectionEl.style.background = '';
            sectionEl.style.backgroundColor = color;
            // Apply same background to columns and image placeholders
            const innerTargets = sectionEl.querySelectorAll('.column, .image-placeholder');
            innerTargets.forEach((el) => {
                el.style.background = '';
                el.style.backgroundColor = color;
            });
            // Keep the typing surface transparent so text tools don't create white bars
            const editables = sectionEl.querySelectorAll('[contenteditable="true"]');
            editables.forEach((ed) => {
                ed.style.background = 'transparent';
                ed.style.backgroundColor = 'transparent';
                // Also clear accidental white background on immediate children/spans/paragraphs
                const descendants = ed.querySelectorAll('*');
                descendants.forEach((node) => {
                    if (node && node.style && node.style.backgroundColor) {
                        const bg = node.style.backgroundColor.trim().toLowerCase();
                        if (bg === 'white' || bg === '#fff' || bg === '#ffffff' || bg === 'rgb(255, 255, 255)') {
                            node.style.backgroundColor = 'transparent';
                        }
                    }
                });
            });
        } catch (_) { /* no-op */ }
    }

    setupTableToolbar() {
        const toolbar = document.getElementById('tableToolbar');
        if (toolbar) toolbar.style.display = 'none';
    }

    // ===== Table Operations =====
    insertTableRow() { return; }

    insertTableColumn() { return; }

    deleteTableRow() { return; }

    deleteTableColumn() { return; }

    changeTableBackgroundColor(color) { return; }

    showTableProperties() { return; }
    
    setupImageToolbar() {
        const toolbar = document.getElementById('imageToolbar');
        // Ensure toolbar controls display on a single line
        const tools = toolbar && toolbar.querySelector('.image-tools');
        if (tools) {
            tools.style.display = 'flex';
            tools.style.flexDirection = 'row';
            tools.style.alignItems = 'center';
            tools.style.gap = '10px';
            tools.style.flexWrap = 'wrap';
        }
        // Ensure icon buttons are fully clickable; allow wrap instead of horizontal scrollbar
        const row = toolbar && toolbar.querySelector('.image-tools-row');
        if (row) {
            row.style.display = 'flex';
            row.style.flexDirection = 'row';
            row.style.alignItems = 'center';
            row.style.gap = '10px';
            row.style.flexWrap = 'wrap';
            row.style.overflowX = 'visible';
            row.style.whiteSpace = 'normal';
            // Prevent children from expanding unevenly
            const rowKids = row.querySelectorAll(':scope > *');
            rowKids.forEach(k => {
                k.style.flex = '0 0 auto';
            });
        }
        
        // Convert toolbar controls to icon-only with tooltips, preserving behavior
        const toIconOnly = (btn, iconHtml, title) => {
            if (!btn) return;
            btn.innerHTML = iconHtml;
            if (title) {
                btn.setAttribute('title', title);
                btn.setAttribute('aria-label', title);
            }
        };
        toIconOnly(document.getElementById('cropImageBtn'), '<i class="fas fa-crop"></i>', 'Recadrer');
        toIconOnly(document.getElementById('positionBtn'), '<i class="fas fa-arrows-alt"></i> <i class="fas fa-caret-down"></i>', 'Position');
        toIconOnly(document.getElementById('resetImageBtn'), '<i class="fas fa-undo"></i>', 'Réinitialiser');
        toIconOnly(document.getElementById('deleteImageBtn'), '<i class="fas fa-trash"></i>', 'Supprimer');
        toIconOnly(document.getElementById('rotationBtn'), '<i class="fas fa-sync"></i> <i class="fas fa-caret-down"></i>', 'Rotation');
        toIconOnly(document.getElementById('changeImageBtn'), '<i class="fas fa-image"></i>', "Changer l'image");
        toIconOnly(document.getElementById('linkImageBtn'), '<i class="fas fa-link"></i>', 'Lien');

        // Advanced gap (per-side) dropdown
        const gapAdvBtn = document.getElementById('gapAdvancedBtn');
        const gapAdvOptions = document.getElementById('gapAdvancedOptions');
        if (gapAdvBtn && gapAdvOptions) {
            gapAdvBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                gapAdvOptions.style.display = gapAdvOptions.style.display === 'none' ? 'block' : 'none';
            });

            const applyAdvancedGap = () => {
                if (!this.currentEditingImage) return;
                const w = this.currentEditingImage.closest('.image-wrapper');
                if (!w) return;
                const top = parseInt(document.getElementById('gapTopInput')?.value || '0', 10) || 0;
                const right = parseInt(document.getElementById('gapRightInput')?.value || '0', 10) || 0;
                const bottom = parseInt(document.getElementById('gapBottomInput')?.value || '0', 10) || 0;
                const left = parseInt(document.getElementById('gapLeftInput')?.value || '0', 10) || 0;

                if (w.classList.contains('position-inline')) {
                    // Keep inline centering: only apply top/bottom, preserve auto sides
                    w.style.marginTop = top + 'px';
                    w.style.marginBottom = bottom + 'px';
                    w.style.marginLeft = 'auto';
                    w.style.marginRight = 'auto';
                } else {
                    // Floats/absolute: apply full per-side margins
                    w.style.margin = `${top}px ${right}px ${bottom}px ${left}px`;
                }
                try { this.saveState(); this.lastAction = 'Espacement personnalisé appliqué'; } catch (_) {}
            };

            const applyBtn = document.getElementById('applyGapBtn');
            if (applyBtn) {
                applyBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    applyAdvancedGap();
                    gapAdvOptions.style.display = 'none';
                });
            }
        }

        // Rotation dropdown toggle
        document.getElementById('rotationBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            const rotationOptions = document.getElementById('rotationOptions');
            rotationOptions.style.display = rotationOptions.style.display === 'none' ? 'block' : 'none';
        });
        
        // Rotate right 90 degrees
        document.getElementById('rotateRight90Btn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.rotateImage(90);
                document.getElementById('rotationOptions').style.display = 'none';
            }
        });
        
        // Rotate left 90 degrees
        document.getElementById('rotateLeft90Btn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.rotateImage(-90);
                document.getElementById('rotationOptions').style.display = 'none';
            }
        });
        
        // Flip vertical
        document.getElementById('flipVerticalBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.flipImage('vertical');
                document.getElementById('rotationOptions').style.display = 'none';
            }
        });
        
        // Flip horizontal
        document.getElementById('flipHorizontalBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.flipImage('horizontal');
                document.getElementById('rotationOptions').style.display = 'none';
            }
        });
        
        // Reset image button
        document.getElementById('resetImageBtn').addEventListener('click', () => {
            if (this.currentEditingImage && this.originalImageSrc) {
                // Reset to original image
                this.currentEditingImage.src = this.originalImageSrc;
                this.currentEditingImage.style.width = 'auto';
                this.currentEditingImage.style.height = 'auto';
                this.currentEditingImage.style.transform = 'none';
                
                // Store the original dimensions for future reference
                this.originalImageWidth = this.currentEditingImage.naturalWidth;
                this.originalImageHeight = this.currentEditingImage.naturalHeight;
                
                this.saveState();
                this.lastAction = 'Image réinitialisée';
            }
        });

        // Delete image button
        document.getElementById('deleteImageBtn').addEventListener('click', (ev) => {
            if (this.currentEditingImage) {
                if (confirm('Êtes-vous sûr de vouloir supprimer cette image ?')) {
                    // Find the image wrapper and remove it
                    const wrapper = this.currentEditingImage.closest('.image-wrapper');
                    if (wrapper) {
                        ev && ev.preventDefault && ev.preventDefault();
                        ev && ev.stopPropagation && ev.stopPropagation();
                        this.deleteImageWrapperSafe(wrapper);
                    } else {
                        // Fallback: remove the image directly
                        this.deleteImageWrapperSafe(this.currentEditingImage);
                    }
                    
                    // Hide the image toolbar
                    document.getElementById('imageToolbar').style.display = 'none';
                    
                    // Clear current editing image reference
                    this.currentEditingImage = null;
                }
            }
        });

        // Change image button — replace current image source without altering other tools
        const changeImageBtn = document.getElementById('changeImageBtn');
        if (changeImageBtn) {
            changeImageBtn.addEventListener('click', () => {
                if (!this.currentEditingImage) return;
                const input = document.createElement('input');
                input.type = 'file';
                input.accept = 'image/*';
                input.onchange = async (e) => {
                    const file = e.target.files && e.target.files[0];
                    if (!file) return;
                    const applySrc = (dataUrl) => {
                        try {
                            this.currentEditingImage.src = dataUrl;
                            if (file && file.name) this.currentEditingImage.alt = file.name;
                            this.currentEditingImage.dataset.originalSrc = dataUrl;
                            this.originalImageSrc = dataUrl;
                            // Reset transforms/sizing for a clean state; user can resize again
                            this.currentEditingImage.style.transform = 'none';
                            this.currentEditingImage.style.width = 'auto';
                            this.currentEditingImage.style.height = 'auto';
                            // Update stored natural dimensions once loaded
                            const imgRef = this.currentEditingImage;
                            const onLoad = () => {
                                this.originalImageWidth = imgRef.naturalWidth;
                                this.originalImageHeight = imgRef.naturalHeight;
                                imgRef.removeEventListener('load', onLoad);
                                this.saveState();
                                this.updateLastModified();
                                this.autoSaveToLocalStorage();
                                this.lastAction = 'Image remplacée';
                            };
                            imgRef.addEventListener('load', onLoad);
                        } catch (_) { /* no-op */ }
                    };
                    try {
                        const dataUrl = await this.compressImageFile(file, { maxWidth: 1600, maxHeight: 1200, quality: 0.9 });
                        applySrc(dataUrl);
                    } catch (_) {
                        const reader = new FileReader();
                        reader.onload = (ev) => applySrc(ev.target.result);
                        reader.readAsDataURL(file);
                    }
                };
                input.click();
            });
        }

        // Link image button — add/update/remove hyperlink around the selected image
        const linkBtn = document.getElementById('linkImageBtn');
        if (linkBtn) {
            linkBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                if (!this.currentEditingImage) return;

                const img = this.currentEditingImage;
                // Find existing anchor that directly wraps the image
                let anchor = (img.parentElement && img.parentElement.tagName === 'A') ? img.parentElement : img.closest && img.closest('a');
                let currentHref = anchor ? (anchor.getAttribute('href') || '') : '';

                const url = prompt("Entrez l'URL du lien (laissez vide pour supprimer):", currentHref || '');
                if (url === null) return; // user canceled
                const trimmed = (url || '').trim();

                // Remove link if empty
                if (trimmed === '') {
                    if (anchor) {
                        const parent = anchor.parentNode;
                        try {
                            parent.insertBefore(img, anchor);
                            anchor.remove();
                        } catch (_) { /* no-op */ }
                        this.saveState();
                        this.updateLastModified();
                        this.autoSaveToLocalStorage();
                        this.lastAction = "Lien d'image supprimé";
                    }
                    return;
                }

                // Add or update link
                if (!anchor) {
                    anchor = document.createElement('a');
                    try {
                        const parent = img.parentNode;
                        parent.insertBefore(anchor, img);
                        anchor.appendChild(img);
                    } catch (_) { /* no-op */ }
                }
                try {
                    anchor.setAttribute('href', trimmed);
                    anchor.setAttribute('target', '_blank');
                    anchor.setAttribute('rel', 'noopener noreferrer');
                } catch (_) { /* no-op */ }

                // Persist state
                this.saveState();
                this.updateLastModified();
                this.autoSaveToLocalStorage();
                this.lastAction = "Lien d'image mis à jour";
            });
        }

        
        // Crop image button
        document.getElementById('cropImageBtn').addEventListener('click', () => {
            if (this.currentEditingImage) {
                this.startImageCropping();
            }
        });
        
        // Cancel crop button
        document.getElementById('cancelCropBtn').addEventListener('click', () => {
            this.cancelImageCropping();
        });
        
        // Apply crop button
        document.getElementById('applyCropBtn').addEventListener('click', () => {
            this.applyImageCropping();
        });
        
        // Position dropdown toggle
        document.getElementById('positionBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            const positionOptions = document.getElementById('positionOptions');
            positionOptions.style.display = positionOptions.style.display === 'none' ? 'block' : 'none';
        });
        
        // Free position removed: no listener wiring to avoid enabling absolute drag
        
        // Float left
        document.getElementById('floatLeftBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.setImagePosition('float-left');
                document.getElementById('positionOptions').style.display = 'none';
            }
        });
        
        // Float right
        document.getElementById('floatRightBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.setImagePosition('float-right');
                document.getElementById('positionOptions').style.display = 'none';
            }
        });
        
        // Inline
        document.getElementById('inlineBtn').addEventListener('click', (e) => {
            e.stopPropagation(); // Prevent event from bubbling up
            if (this.currentEditingImage) {
                this.setImagePosition('inline');
                document.getElementById('positionOptions').style.display = 'none';
            }
        });
        
        // Close dropdowns when clicking outside
        document.addEventListener('click', (e) => {
            // Handle rotation dropdown
            const rotationOptions = document.getElementById('rotationOptions');
            const rotationBtn = document.getElementById('rotationBtn');
            
            if (rotationOptions && 
                rotationOptions.style.display === 'block' && 
                !rotationBtn.contains(e.target) && 
                !rotationOptions.contains(e.target)) {
                rotationOptions.style.display = 'none';
            }

            // Close advanced gap dropdown
            const gapAdvBtnEl = document.getElementById('gapAdvancedBtn');
            const gapAdvOptsEl = document.getElementById('gapAdvancedOptions');
            if (gapAdvOptsEl && gapAdvOptsEl.style.display === 'block' && gapAdvBtnEl &&
                !gapAdvBtnEl.contains(e.target) && !gapAdvOptsEl.contains(e.target)) {
                gapAdvOptsEl.style.display = 'none';
            }
            
            // Handle position dropdown
            const positionOptions = document.getElementById('positionOptions');
            const positionBtn = document.getElementById('positionBtn');
            
            if (positionOptions && 
                positionOptions.style.display === 'block' && 
                !positionBtn.contains(e.target) && 
                !positionOptions.contains(e.target)) {
                positionOptions.style.display = 'none';
            }
        });
    }
    
    rotateImage(degrees) {
        if (!this.currentEditingImage) return;
        // Resolve the actual IMG element in case a link or wrapper was selected
        let imgEl = this.currentEditingImage;
        if (imgEl && imgEl.tagName && imgEl.tagName.toUpperCase() !== 'IMG') {
            try {
                const inner = imgEl.querySelector && imgEl.querySelector('img');
                if (inner) imgEl = inner;
            } catch (_) {}
        }
        if (!imgEl || (imgEl.tagName && imgEl.tagName.toUpperCase() !== 'IMG')) return;

        // Read current rotation from data attribute (robust even if transform is a matrix)
        let currentRotation = 0;
        if (imgEl.dataset && imgEl.dataset.rotationDegrees) {
            currentRotation = parseInt(imgEl.dataset.rotationDegrees) || 0;
        } else {
            const inline = imgEl.style && imgEl.style.transform || '';
            const m = inline.match(/rotate\((-?\d+)deg\)/);
            if (m) currentRotation = parseInt(m[1]) || 0;
        }

        // Compute and store new rotation
        const newRotation = ((currentRotation + degrees) % 360 + 360) % 360;
        if (imgEl.dataset) imgEl.dataset.rotationDegrees = String(newRotation);

        // Preserve other inline transforms by stripping any rotate() then appending
        const existing = imgEl.style && imgEl.style.transform ? imgEl.style.transform : '';
        const base = existing.replace(/rotate\((-?\d+)deg\)/, '').trim();
        const finalTransform = (base ? base + ' ' : '') + `rotate(${newRotation}deg)`;
        imgEl.style.transform = finalTransform;
        imgEl.style.transformOrigin = 'center center';

        // Persist state and refresh handles
        try {
            this.saveState();
            this.lastAction = 'Image tournée';
            this.removeResizeHandles();
            this.addResizeHandlesToImage(imgEl);
        } catch (_) {}
    }
    
    // Central entry-point used by click handlers to activate the floating image tools
    selectImage(image) {
        if (!image) return;
        // If a wrapper or container is passed, resolve to the inner <img>
        try {
            if (image.tagName !== 'IMG' && image.querySelector) {
                const inner = image.querySelector('img');
                if (inner) image = inner;
            }
        } catch (_) {}
        this.showImageEditingTools(image);
    }
    
    flipImage(direction) {
        if (!this.currentEditingImage) return;
        // Resolve the actual IMG element
        let imgEl = this.currentEditingImage;
        if (imgEl && imgEl.tagName && imgEl.tagName.toUpperCase() !== 'IMG') {
            try {
                const inner = imgEl.querySelector && imgEl.querySelector('img');
                if (inner) imgEl = inner;
            } catch (_) {}
        }
        if (!imgEl || (imgEl.tagName && imgEl.tagName.toUpperCase() !== 'IMG')) return;

        // Track flip state in data attributes so repeated clicks toggle correctly
        const flipX = direction === 'horizontal';
        const flipY = direction === 'vertical';
        const currentFlipX = imgEl.dataset && imgEl.dataset.flipX === '1';
        const currentFlipY = imgEl.dataset && imgEl.dataset.flipY === '1';

        let newFlipX = currentFlipX;
        let newFlipY = currentFlipY;
        if (flipX) newFlipX = !currentFlipX;
        if (flipY) newFlipY = !currentFlipY;
        if (imgEl.dataset) {
            imgEl.dataset.flipX = newFlipX ? '1' : '0';
            imgEl.dataset.flipY = newFlipY ? '1' : '0';
        }

        // Build transform: preserve existing non-flip transforms (e.g., rotate, scale) by stripping previous scaleX/scaleY flips
        const existing = imgEl.style && imgEl.style.transform ? imgEl.style.transform : '';
        const base = existing
            .replace(/scaleX\((-?\d*\.?\d+)\)/g, '')
            .replace(/scaleY\((-?\d*\.?\d+)\)/g, '')
            .replace(/\s{2,}/g, ' ')
            .trim();

        const flips = [];
        if (newFlipX) flips.push('scaleX(-1)');
        if (newFlipY) flips.push('scaleY(-1)');
        const finalTransform = (flips.join(' ') + ' ' + base).trim();
        imgEl.style.transform = finalTransform || 'none';
        imgEl.style.transformOrigin = 'center center';

        try {
            this.saveState();
            this.lastAction = 'Image retournée';
            this.removeResizeHandles();
            this.addResizeHandlesToImage(imgEl);
        } catch (_) {}
    }
    
    showImageEditingTools(image) {
        // Store reference to the current image being edited
        this.currentEditingImage = image;
        
        // Store original dimensions for resizing
        this.originalImageWidth = image.naturalWidth;
        this.originalImageHeight = image.naturalHeight;
        
        // Store original image source for reset functionality (refresh per selection)
        // Use the data attribute if available, otherwise use current src
        this.originalImageSrc = image.dataset.originalSrc || image.src;
        
        // Add a class to highlight the image being edited
        image.classList.add('image-being-edited');
        
        // Position the toolbar near the image
        const toolbar = document.getElementById('imageToolbar');
        const rect = image.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        
        // Ensure toolbar is not inside section columns (flex children)
        if (toolbar.parentElement !== document.body) {
            try { document.body.appendChild(toolbar); } catch (_) {}
        }
        // Force absolute positioning to avoid affecting layout
        toolbar.style.position = 'absolute';
        toolbar.style.zIndex = '9999';
        toolbar.style.display = 'block';
        toolbar.style.width = 'auto';
        toolbar.style.pointerEvents = 'auto';
        toolbar.style.margin = '0';
        
        // Compute centered and constrained position so it stays on-screen
        try {
            const viewportWidth = window.innerWidth || document.documentElement.clientWidth;
            const toolbarWidth = toolbar.offsetWidth;
            const toolbarHeight = toolbar.offsetHeight;
            let left = rect.left + (rect.width / 2) - (toolbarWidth / 2);
            const minLeft = 10;
            const maxLeft = Math.max(minLeft, viewportWidth - toolbarWidth - 10);
            left = Math.min(Math.max(left, minLeft), maxLeft);

            // Prefer above image; add clearance so it doesn't cover top controls/handles
            const clearance = 36; // leave space for rotation/resize handles and a margin
            let top = rect.top + scrollTop - toolbarHeight - clearance;
            const minTop = scrollTop + 10;
            if (top < minTop) {
                // Not enough space above, place below with same clearance
                top = rect.bottom + scrollTop + clearance;
            }

            toolbar.style.left = left + 'px';
            toolbar.style.top = top + 'px';
        } catch (_) {
            // Fallback: align to left edge
            const clearance = 36;
            toolbar.style.left = rect.left + 'px';
            toolbar.style.top = (rect.top + scrollTop - toolbar.offsetHeight - clearance) + 'px';
        }
        
        // Setup toolbar if not already done
        if (!this.imageToolbarInitialized) {
            this.setupImageToolbar();
            this.imageToolbarInitialized = true;
        }
        
        // Hide rotation dropdown initially
        document.getElementById('rotationOptions').style.display = 'none';
        
        // Hide crop-related buttons initially
        document.getElementById('cropImageBtn').style.display = 'block';
        document.getElementById('cancelCropBtn').style.display = 'none';
        document.getElementById('applyCropBtn').style.display = 'none';
        
        // Add resize handles to the image
        this.addResizeHandlesToImage(image);
        
        // Ensure proper positioning after a short delay to allow for DOM updates
        setTimeout(() => {
            this.updateResizeHandlePositions();
        }, 10);
    }
    
    addResizeHandlesToImage(image) {
        // Remove any existing resize handles first
        this.removeResizeHandles();
        
        // Get the parent wrapper or create one if it doesn't exist
        let wrapper = image.parentElement;
        const galleryContainer = image.closest && image.closest('.gallery-image-container');
        const isInGallery = !!galleryContainer;
        if (!wrapper.classList.contains('image-wrapper')) {
            wrapper = document.createElement('div');
            wrapper.className = 'image-wrapper';
            if (isInGallery) {
                // Fill the square container without changing layout
                wrapper.style.cssText = 'position: absolute; top:0; left:0; right:0; bottom:0; width:100%; height:100%; display:block; margin:0;';
            } else {
                // Respect previously chosen positioning. If centered inline, keep flex centering.
                if (image.closest && image.closest('.image-wrapper.position-inline')) {
                    wrapper.style.cssText = 'position: relative; display: flex; justify-content: center; align-items: center; width:100%; margin: 10px 0;';
                } else {
                    wrapper.style.cssText = 'position: relative; display: inline-block; margin: 10px 0;';
                }
            }
            image.parentNode.insertBefore(wrapper, image);
            wrapper.appendChild(image);
        } else {
            // If wrapper exists and image is inside gallery, ensure natural flow so container height follows image
            if (isInGallery) {
                wrapper.style.position = 'relative';
                wrapper.style.top = '';
                wrapper.style.left = '';
                wrapper.style.right = '';
                wrapper.style.bottom = '';
                wrapper.style.width = '100%';
                wrapper.style.height = 'auto';
                wrapper.style.display = 'block';
                wrapper.style.margin = '0';
            }
        }

        // If this image is in inline-centered mode, keep centering styles intact when (re)adding handles
        try {
            if (!isInGallery && wrapper.classList.contains('position-inline')) {
                // Shrink wrapper to image and center it
                wrapper.style.display = 'block';
                wrapper.style.width = 'fit-content';
                wrapper.style.marginLeft = 'auto';
                wrapper.style.marginRight = 'auto';
                wrapper.style.float = 'none';
                wrapper.style.cssFloat = 'none';
                image.style.float = 'none';
                image.style.display = 'block';
                image.style.width = image.style.width || 'auto';
            }
        } catch (_) {}

        // Ensure wrapper positioning
        if (!isInGallery) {
            // Do not override absolute positioning
            if (!wrapper.classList.contains('position-absolute')) {
                wrapper.style.position = 'relative';
            }
            // Keep existing layout for absolute/float/inline modes; only default to inline-block otherwise
            const keepDisplay = wrapper.classList.contains('position-inline') ||
                                wrapper.classList.contains('position-absolute') ||
                                wrapper.classList.contains('float-left') ||
                                wrapper.classList.contains('float-right');
            if (!keepDisplay) {
                wrapper.style.display = 'inline-block';
            }
        } else {
            wrapper.style.position = 'relative';
            wrapper.style.display = 'block';
        }
        
        // Create resize handles
        const handlePositions = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
        handlePositions.forEach(pos => {
            const handle = document.createElement('div');
            handle.className = `resize-handle ${pos}`;
            handle.dataset.position = pos;
            handle.style.position = 'absolute';
            handle.style.zIndex = '1000';
            wrapper.appendChild(handle);
            
            // Add event listeners for resizing
            this.setupResizeHandleEvents(handle, image);
        });
        
        // Create rotation handle
        const rotationHandle = document.createElement('div');
        rotationHandle.className = 'rotation-handle';
        rotationHandle.innerHTML = '<i class="fas fa-sync-alt"></i>';
        rotationHandle.style.position = 'absolute';
        rotationHandle.style.zIndex = '1001';
        wrapper.appendChild(rotationHandle);
        
        // Add event listener for rotation
        this.setupRotationHandleEvents(rotationHandle, image);
        
        // Store reference to the wrapper
        this.currentImageWrapper = wrapper;
        
        // Update handle positions after wrapper is set up
        this.updateResizeHandlePositions();
    }
    
    updateResizeHandlePositions() {
        if (!this.currentImageWrapper) return;
        
        const handles = this.currentImageWrapper.querySelectorAll('.resize-handle');
        handles.forEach(handle => {
            // Ensure handles are properly positioned
            handle.style.position = 'absolute';
            handle.style.zIndex = '1000';
        });
    }
    
    setupRotationHandleEvents(handle, image) {
        handle.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            
            // Disable dragging temporarily if the image is in absolute position
            const wrapper = image.closest('.image-wrapper');
            if (wrapper && wrapper.classList.contains('position-absolute')) {
                wrapper.setAttribute('data-temp-draggable', wrapper.getAttribute('draggable') || 'false');
                wrapper.setAttribute('draggable', 'false');
            }
            
            // Get the center point of the image
            const rect = image.getBoundingClientRect();
            const centerX = rect.left + rect.width / 2;
            const centerY = rect.top + rect.height / 2;
            
            // Calculate the initial angle
            const startAngle = Math.atan2(e.clientY - centerY, e.clientX - centerX);
            
            // Get current rotation from transform style or default to 0
            let currentRotation = 0;
            const transform = image.style.transform;
            if (transform) {
                const match = transform.match(/rotate\((-?\d+)deg\)/);
                if (match) {
                    currentRotation = parseInt(match[1]);
                }
            }
            
            // Mouse move handler for rotation
            const mouseMoveHandler = (e) => {
                // Calculate the new angle
                const newAngle = Math.atan2(e.clientY - centerY, e.clientX - centerX);
                
                // Calculate the angle difference in degrees
                let angleDiff = (newAngle - startAngle) * (180 / Math.PI);
                
                // Apply the new rotation
                const newRotation = currentRotation + angleDiff;
                image.style.transform = `rotate(${newRotation}deg)`;
                
                // Update the handle position
                handle.style.transform = `rotate(${newRotation}deg)`;
            };
            
            // Mouse up handler to stop rotation
            const mouseUpHandler = () => {
                document.removeEventListener('mousemove', mouseMoveHandler);
                document.removeEventListener('mouseup', mouseUpHandler);
                
                // Restore draggable functionality if it was temporarily disabled
                const wrapper = image.closest('.image-wrapper');
                if (wrapper && wrapper.hasAttribute('data-temp-draggable')) {
                    wrapper.setAttribute('draggable', wrapper.getAttribute('data-temp-draggable'));
                    wrapper.removeAttribute('data-temp-draggable');
                }
                
                // Keep CSS-based rotation (non-destructive); just persist state
                this.saveState();
                this.lastAction = 'Image tournée';
            };
            
            // Add event listeners
            document.addEventListener('mousemove', mouseMoveHandler);
            document.addEventListener('mouseup', mouseUpHandler);
            
            // Store references to remove later
            this.resizeEventHandlers = {
                mousemove: mouseMoveHandler,
                mouseup: mouseUpHandler
            };
        });
    }
    
    applyRotationToImageData() {
        if (!this.currentEditingImage) return;
        
        // Get the current rotation angle
        let currentRotation = 0;
        const transform = this.currentEditingImage.style.transform;
        if (transform) {
            const match = transform.match(/rotate\((-?\d+)deg\)/);
            if (match) {
                currentRotation = parseInt(match[1]);
            }
        }
        
        // Normalize the rotation angle to be between 0 and 360
        currentRotation = ((currentRotation % 360) + 360) % 360;
        
        // Only apply the rotation if it's significant (to avoid unnecessary processing)
        if (Math.abs(currentRotation) < 0.5) {
            this.currentEditingImage.style.transform = 'none';
            return;
        }
        
        // Create a canvas to perform the rotation
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        
        // Calculate the dimensions needed for the rotated image
        const angleInRadians = currentRotation * Math.PI / 180;
        const imgWidth = this.currentEditingImage.naturalWidth;
        const imgHeight = this.currentEditingImage.naturalHeight;
        
        // Calculate the dimensions of the rotated canvas
        const sin = Math.abs(Math.sin(angleInRadians));
        const cos = Math.abs(Math.cos(angleInRadians));
        const newWidth = Math.ceil(imgWidth * cos + imgHeight * sin);
        const newHeight = Math.ceil(imgWidth * sin + imgHeight * cos);
        
        canvas.width = newWidth;
        canvas.height = newHeight;
        
        // Move to the center of the canvas
        ctx.translate(newWidth / 2, newHeight / 2);
        
        // Rotate the canvas
        ctx.rotate(angleInRadians);
        
        // Draw the image centered on the canvas
        ctx.drawImage(this.currentEditingImage, -imgWidth / 2, -imgHeight / 2);
        
        // Apply the rotated image
        const rotatedImageDataUrl = canvas.toDataURL('image/png');
        this.currentEditingImage.src = rotatedImageDataUrl;
        
        // Reset the transform since the rotation is now in the image data
        this.currentEditingImage.style.transform = 'none';
        
        // Update the original dimensions
        this.originalImageWidth = newWidth;
        this.originalImageHeight = newHeight;
    }
    
    removeResizeHandles() {
        // Remove all resize handles
        const handles = document.querySelectorAll('.resize-handle');
        handles.forEach(handle => {
            handle.remove();
        });
        
        // Remove rotation handle
        const rotationHandle = document.querySelector('.rotation-handle');
        if (rotationHandle) {
            rotationHandle.remove();
        }
        
        // Remove event listeners
        if (this.resizeEventHandlers) {
            document.removeEventListener('mousemove', this.resizeEventHandlers.mousemove);
            document.removeEventListener('mouseup', this.resizeEventHandlers.mouseup);
            this.resizeEventHandlers = null;
        }
    }
    
    hideResizeAndRotationHandles() {
        // Hide all resize handles
        const handles = document.querySelectorAll('.resize-handle');
        handles.forEach(handle => {
            handle.style.display = 'none';
        });
        
        // Hide rotation handle
        const rotationHandle = document.querySelector('.rotation-handle');
        if (rotationHandle) {
            rotationHandle.style.display = 'none';
        }
    }
    
    setImagePosition(positionType) {
        if (!this.currentEditingImage) return;
        
        // Get the parent wrapper
        let wrapper = this.currentEditingImage.closest('.image-wrapper');
        if (!wrapper) {
            // If no wrapper exists, the image might be directly in the content
            wrapper = document.createElement('div');
            wrapper.className = 'image-wrapper';
            this.currentEditingImage.parentNode.insertBefore(wrapper, this.currentEditingImage);
            wrapper.appendChild(this.currentEditingImage);
        }
        
        // Remove any existing position classes
        wrapper.classList.remove('position-absolute', 'float-left', 'float-right', 'position-inline');
        
        // Reset common coordinates
        wrapper.style.position = '';
        wrapper.style.left = '';
        wrapper.style.top = '';
        wrapper.style.zIndex = '';

        // Helper: clear flex centering and width from previous inline mode
        const resetWrapperLayout = () => {
            wrapper.style.display = '';
            wrapper.style.justifyContent = '';
            wrapper.style.alignItems = '';
            wrapper.style.textAlign = '';
            wrapper.style.float = '';
            // Do not force 100% width outside inline mode
            wrapper.style.width = '';
        };
        
        // Apply the new position type
        switch (positionType) {
            case 'absolute':
                // Compute current on-screen position BEFORE changing layout
                let anchor = null;
                try {
                    anchor = wrapper.closest('.column') || wrapper.closest('.newsletter-section') || document.getElementById('editableContent');
                } catch (_) {
                    anchor = document.getElementById('editableContent');
                }
                const preRect = wrapper.getBoundingClientRect();
                const containerRect = anchor ? anchor.getBoundingClientRect() : { left: 0, top: 0 };

                // Set up for absolute positioning
                resetWrapperLayout();
                wrapper.classList.add('position-absolute');
                wrapper.style.position = 'absolute';
                // Make wrapper size to content for free movement
                wrapper.style.width = 'auto';

                // Ensure same anchoring in preview by setting inline position on the anchor
                if (anchor && getComputedStyle(anchor).position === 'static') {
                    anchor.style.position = 'relative';
                }
                // Apply left/top based on pre-change coordinates relative to anchor
                const relLeft = Math.max(0, preRect.left - containerRect.left);
                const relTop = Math.max(0, preRect.top - containerRect.top);
                wrapper.style.left = relLeft + 'px';
                wrapper.style.top = relTop + 'px';
                wrapper.style.zIndex = '1';
                
                // Preserve any user-resized dimensions; only clear problematic 100% width
                try {
                    const img = this.currentEditingImage;
                    if (img) {
                        img.style.display = 'block';
                        // If width was forced to 100% in previous modes, clear it; otherwise keep current pixel/auto width
                        if (img.style.width && img.style.width.trim() === '100%') {
                            img.style.width = '';
                        }
                        // If no explicit size is set, freeze current rendered size to avoid snapping back
                        const hasExplicitW = !!img.style.width && img.style.width.trim() !== '';
                        const hasExplicitH = !!img.style.height && img.style.height.trim() !== '';
                        if (!hasExplicitW || !hasExplicitH) {
                            const r = img.getBoundingClientRect();
                            if (!hasExplicitW) img.style.width = Math.max(1, Math.round(r.width)) + 'px';
                            if (!hasExplicitH) img.style.height = Math.max(1, Math.round(r.height)) + 'px';
                        }
                        // Do not touch height to preserve user resizing
                        img.style.objectFit = '';
                        img.style.float = 'none';
                    }
                } catch (_) {}
                // Make the wrapper draggable
                this.makeImageDraggable(wrapper);
                break;
                
            case 'float-left':
                resetWrapperLayout();
                wrapper.classList.add('float-left');
                // Apply float styles directly so it works without CSS class definitions
                wrapper.style.cssFloat = 'left';
                wrapper.style.float = 'left';
                wrapper.style.display = 'block';
                wrapper.style.width = 'auto';
                const imgL = this.currentEditingImage; if (imgL) { imgL.style.float = 'none'; }
                // Remove draggable functionality
                this.removeImageDraggable(wrapper);
                break;
                
            case 'float-right':
                resetWrapperLayout();
                wrapper.classList.add('float-right');
                wrapper.style.cssFloat = 'right';
                wrapper.style.float = 'right';
                wrapper.style.display = 'block';
                wrapper.style.width = 'auto';
                const imgR = this.currentEditingImage; if (imgR) { imgR.style.float = 'none'; }
                // Remove draggable functionality
                this.removeImageDraggable(wrapper);
                break;
                
            case 'inline':
                wrapper.classList.add('position-inline');
                // Remove draggable functionality
                this.removeImageDraggable(wrapper);
                // Center the WRAPPER so handles fit the image
                try {
                    const img = this.currentEditingImage;
                    if (img) {
                        // Wrapper shrinks to content and centers itself
                        wrapper.style.display = 'block';
                        wrapper.style.width = 'fit-content';
                        wrapper.style.marginLeft = 'auto';
                        wrapper.style.marginRight = 'auto';
                        wrapper.style.float = 'none';
                        wrapper.style.cssFloat = 'none';
                        // Image uses natural width so wrapper fits it
                        img.style.float = 'none';
                        img.style.display = 'block';
                        img.style.width = 'auto';
                    }
                } catch (_) {}
                // Center the image in view for better UX
                try { this.centerImageInView(this.currentEditingImage); } catch (_) {}
                break;
        }
        
        this.saveState();
        this.lastAction = 'Position de l\'image ajustée';
    }

    // Smoothly center the given element within the viewport/editor
    centerImageInView(el) {
        if (!el) return;
        try {
            // Prefer smooth center scroll when supported
            el.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' });
        } catch (_) {
            // Fallback: immediate center without smooth behavior
            try { el.scrollIntoView({ block: 'center', inline: 'center' }); } catch (__) {}
        }
    }
    
    makeImageDraggable(wrapper) {
        // Remove any existing drag handlers
        this.removeImageDraggable(wrapper);
        
        // Add draggable attribute
        wrapper.setAttribute('draggable', 'true');
        
        // Store initial position for drag operation
        let startX, startY, startLeft, startTop;
        // Anchor container used for bounds (column -> section -> editableContent)
        let anchorEl = null;
        
        // Add drag start event
        const dragStartHandler = (e) => {
            // Resolve anchor container for precise bounds
            try {
                anchorEl = wrapper.closest('.column') || wrapper.closest('.newsletter-section') || document.getElementById('editableContent');
            } catch (_) {
                anchorEl = document.getElementById('editableContent');
            }
            if (anchorEl && getComputedStyle(anchorEl).position === 'static') {
                anchorEl.style.position = 'relative';
            }
            if (anchorEl) {
                // Ensure the image can move beyond current text height
                anchorEl.style.overflow = 'visible';
                // Prime a minimal height to prevent immediate clipping
                const primed = Math.max(anchorEl.scrollHeight, wrapper.offsetTop + wrapper.offsetHeight + 16);
                anchorEl.style.minHeight = primed + 'px';
            }

            // Store the initial position
            startX = e.clientX;
            startY = e.clientY;
            startLeft = parseInt(wrapper.style.left);
            startTop = parseInt(wrapper.style.top);
            e.preventDefault();
            e.stopPropagation();
        };
        
        // Add drag end event
        const dragEndHandler = (e) => {
            // Remove dragging class
            wrapper.classList.remove('dragging');
        };
        
        // Add drag event
        const dragHandler = (e) => {
            e.preventDefault();
        };
        
        // Add dragover event to the editable content area
        const dragOverHandler = (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
        };
        
        // Add drop event to the editable content area
        const dropHandler = (e) => {
            e.preventDefault();
            
            // Calculate the new position
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            
            wrapper.style.left = (startLeft + dx) + 'px';
            wrapper.style.top = (startTop + dy) + 'px';
            
            this.saveState();
            this.lastAction = 'Image déplacée';
        };
        
        // Add mouse events for more precise dragging
        const mouseDownHandler = (e) => {
            // Only handle primary button (left-click)
            if (e.button !== 0) return;
            
            // Check if the click is on the image itself (not on a handle)
            if (e.target.tagName === 'IMG' || e.target === wrapper) {
                // Prevent default to avoid text selection
                e.preventDefault();
                
                // Store the initial position
                startX = e.clientX;
                startY = e.clientY;
                startLeft = parseInt(wrapper.style.left);
                startTop = parseInt(wrapper.style.top);
                
                // Add a class to indicate dragging
                wrapper.classList.add('dragging');
                
                // Add temporary event listeners for mouse move and up
                document.addEventListener('mousemove', mouseMoveHandler);
                document.addEventListener('mouseup', mouseUpHandler);
            }
        };
        
        const mouseMoveHandler = (e) => {
            // Calculate the new position
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;

            let nextLeft = startLeft + dx;
            let nextTop = startTop + dy;

            // Constrain to anchor bounds if in absolute mode
            if (wrapper.classList.contains('position-absolute')) {
                const container = anchorEl || document.getElementById('editableContent');
                if (container) {
                    const cRect = container.getBoundingClientRect();
                    const maxLeft = Math.max(0, cRect.width - wrapper.offsetWidth);
                    // Use current max of visual height and scrollHeight to allow dragging below current text
                    const effectiveHeight = Math.max(cRect.height, container.scrollHeight);
                    let maxTop = Math.max(0, effectiveHeight - wrapper.offsetHeight);
                    nextLeft = Math.min(Math.max(0, nextLeft), maxLeft);
                    nextTop = Math.max(0, nextTop);
                    // Expand container min-height when dragging lower
                    const needed = nextTop + wrapper.offsetHeight + 16;
                    if (needed > (parseInt(container.style.minHeight) || 0)) {
                        container.style.minHeight = needed + 'px';
                        // Recompute maxTop after expanding
                        maxTop = Math.max(0, Math.max(container.getBoundingClientRect().height, container.scrollHeight) - wrapper.offsetHeight);
                    }
                    nextTop = Math.min(nextTop, maxTop);
                }
            }

            wrapper.style.left = nextLeft + 'px';
            wrapper.style.top = nextTop + 'px';
        };
        
        const mouseUpHandler = (e) => {
            // Remove dragging class
            wrapper.classList.remove('dragging');
            
            // Remove temporary event listeners
            document.removeEventListener('mousemove', mouseMoveHandler);
            document.removeEventListener('mouseup', mouseUpHandler);
            
            this.saveState();
            this.lastAction = 'Image déplacée';
        };
        
        // Add event listeners
        wrapper.addEventListener('dragstart', dragStartHandler);
        wrapper.addEventListener('dragend', dragEndHandler);
        wrapper.addEventListener('drag', dragHandler);
        document.addEventListener('dragover', dragOverHandler);
        document.addEventListener('drop', dropHandler);
        
        // Add mouse event listeners for more precise dragging
        wrapper.addEventListener('mousedown', mouseDownHandler);
        
        // Store the event handlers for later removal
        wrapper.dragHandlers = {
            dragstart: dragStartHandler,
            dragend: dragEndHandler,
            drag: dragHandler,
            dragover: dragOverHandler,
            drop: dropHandler,
            mousedown: mouseDownHandler,
            mousemove: mouseMoveHandler,
            mouseup: mouseUpHandler
        };
    }
    
    removeImageDraggable(wrapper) {
        // Remove draggable attribute
        wrapper.removeAttribute('draggable');
        
        // Remove event listeners if they exist
        if (wrapper.dragHandlers) {
            wrapper.removeEventListener('dragstart', wrapper.dragHandlers.dragstart);
            wrapper.removeEventListener('dragend', wrapper.dragHandlers.dragend);
            wrapper.removeEventListener('drag', wrapper.dragHandlers.drag);
            document.removeEventListener('dragover', wrapper.dragHandlers.dragover);
            document.removeEventListener('drop', wrapper.dragHandlers.drop);
            wrapper.removeEventListener('mousedown', wrapper.dragHandlers.mousedown);
            document.removeEventListener('mousemove', wrapper.dragHandlers.mousemove);
            document.removeEventListener('mouseup', wrapper.dragHandlers.mouseup);
            
            // Clear the handlers
            wrapper.dragHandlers = null;
        }
    }
    
    showResizeAndRotationHandles() {
        // Show all resize handles
        const handles = document.querySelectorAll('.resize-handle');
        handles.forEach(handle => {
            handle.style.display = 'block';
        });
        
        // Show rotation handle
        const rotationHandle = document.querySelector('.rotation-handle');
        if (rotationHandle) {
            rotationHandle.style.display = 'block';
        }
    }
    
    setupResizeHandleEvents(handle, image) {
        handle.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            
            const position = handle.dataset.position;
            const startX = e.clientX;
            const startY = e.clientY;
            const startWidth = image.offsetWidth;
            const startHeight = image.offsetHeight;
            const aspectRatio = startWidth / startHeight;
            const keepAspectRatio = document.getElementById('keepAspectRatio').checked;
            
            // Mouse move handler for resizing
            const mouseMoveHandler = (e) => {
                let newWidth = startWidth;
                let newHeight = startHeight;
                
                // Calculate new dimensions based on handle position and mouse movement
                switch (position) {
                    case 'e':
                        newWidth = startWidth + (e.clientX - startX);
                        if (keepAspectRatio) {
                            newHeight = newWidth / aspectRatio;
                        }
                        break;
                    case 'w':
                        newWidth = startWidth - (e.clientX - startX);
                        if (keepAspectRatio) {
                            newHeight = newWidth / aspectRatio;
                        }
                        break;
                    case 's':
                        newHeight = startHeight + (e.clientY - startY);
                        if (keepAspectRatio) {
                            newWidth = newHeight * aspectRatio;
                        }
                        break;
                    case 'n':
                        newHeight = startHeight - (e.clientY - startY);
                        if (keepAspectRatio) {
                            newWidth = newHeight * aspectRatio;
                        }
                        break;
                    case 'se':
                        newWidth = startWidth + (e.clientX - startX);
                        newHeight = startHeight + (e.clientY - startY);
                        if (keepAspectRatio) {
                            // Use the larger change to determine the new size
                            const widthChange = (e.clientX - startX) / startWidth;
                            const heightChange = (e.clientY - startY) / startHeight;
                            if (Math.abs(widthChange) > Math.abs(heightChange)) {
                                newHeight = newWidth / aspectRatio;
                            } else {
                                newWidth = newHeight * aspectRatio;
                            }
                        }
                        break;
                    case 'sw':
                        newWidth = startWidth - (e.clientX - startX);
                        newHeight = startHeight + (e.clientY - startY);
                        if (keepAspectRatio) {
                            // Use the larger change to determine the new size
                            const widthChange = (e.clientX - startX) / startWidth;
                            const heightChange = (e.clientY - startY) / startHeight;
                            if (Math.abs(widthChange) > Math.abs(heightChange)) {
                                newHeight = newWidth / aspectRatio;
                            } else {
                                newWidth = newHeight * aspectRatio;
                            }
                        }
                        break;
                    case 'ne':
                        newWidth = startWidth + (e.clientX - startX);
                        newHeight = startHeight - (e.clientY - startY);
                        if (keepAspectRatio) {
                            // Use the larger change to determine the new size
                            const widthChange = (e.clientX - startX) / startWidth;
                            const heightChange = (e.clientY - startY) / startHeight;
                            if (Math.abs(widthChange) > Math.abs(heightChange)) {
                                newHeight = newWidth / aspectRatio;
                            } else {
                                newWidth = newHeight * aspectRatio;
                            }
                        }
                        break;
                    case 'nw':
                        newWidth = startWidth - (e.clientX - startX);
                        newHeight = startHeight - (e.clientY - startY);
                        if (keepAspectRatio) {
                            // Use the larger change to determine the new size
                            const widthChange = (e.clientX - startX) / startWidth;
                            const heightChange = (e.clientY - startY) / startHeight;
                            if (Math.abs(widthChange) > Math.abs(heightChange)) {
                                newHeight = newWidth / aspectRatio;
                            } else {
                                newWidth = newHeight * aspectRatio;
                            }
                        }
                        break;
                }
                
                // Ensure minimum size
                newWidth = Math.max(20, newWidth);
                newHeight = Math.max(20, newHeight);
                
                // Apply new dimensions
                image.style.width = newWidth + 'px';
                image.style.height = newHeight + 'px';
                
                // Update the slider to reflect the new size (if present)
                const newPercentage = Math.round((newWidth / this.originalImageWidth) * 100);
                const resizeSlider = document.getElementById('resizeSlider');
                const resizePercentage = document.getElementById('resizePercentage');
                if (resizeSlider) resizeSlider.value = newPercentage;
                if (resizePercentage) resizePercentage.textContent = newPercentage + '%';
            };
            
            // Mouse up handler to stop resizing
            const mouseUpHandler = () => {
                document.removeEventListener('mousemove', mouseMoveHandler);
                document.removeEventListener('mouseup', mouseUpHandler);
                this.saveState();
                this.lastAction = 'Taille de l\'image ajustée';
            };
            
            // Add event listeners
            document.addEventListener('mousemove', mouseMoveHandler);
            document.addEventListener('mouseup', mouseUpHandler);
            
            // Store references to remove later
            this.resizeEventHandlers = {
                mousemove: mouseMoveHandler,
                mouseup: mouseUpHandler
            };
        });
    }
    
    hideImageEditingTools() {
        const toolbar = document.getElementById('imageToolbar');
        toolbar.style.display = 'none';
        
        if (this.currentEditingImage) {
            this.currentEditingImage.classList.remove('image-being-edited');
            this.currentEditingImage = null;
        }
        
        // Reset original image source reference
        this.originalImageSrc = null;
        
        // Remove resize handles
        this.removeResizeHandles();
        
        this.cancelImageCropping();
    }
    
    startImageCropping() {
        if (!this.currentEditingImage) return;
        
        // Hide resize and rotation handles when cropping
        this.hideResizeAndRotationHandles();
        
        // Temporarily disable dragging if the image is in absolute position
        const wrapper = this.currentEditingImage.closest('.image-wrapper');
        if (wrapper && wrapper.classList.contains('position-absolute')) {
            wrapper.setAttribute('data-temp-draggable', wrapper.getAttribute('draggable') || 'false');
            wrapper.setAttribute('draggable', 'false');
            
            // Also remove mousedown handler temporarily
            if (wrapper.dragHandlers && wrapper.dragHandlers.mousedown) {
                wrapper.removeEventListener('mousedown', wrapper.dragHandlers.mousedown);
            }
        }
        
        // Create crop overlay
        const overlay = document.createElement('div');
        overlay.className = 'crop-overlay';
        
        // Position the overlay over the image
        const rect = this.currentEditingImage.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
        
        // Ensure we're using absolute positioning with correct coordinates
        overlay.style.position = 'absolute';
        overlay.style.left = Math.round(rect.left + scrollLeft) + 'px';
        overlay.style.top = Math.round(rect.top + scrollTop) + 'px';
        overlay.style.width = Math.round(rect.width) + 'px';
        overlay.style.height = Math.round(rect.height) + 'px';
        overlay.style.zIndex = '1002';
        overlay.style.pointerEvents = 'auto';
        
        // Create resize handles
        const handles = ['tl', 'tr', 'bl', 'br'];
        handles.forEach(pos => {
            const handle = document.createElement('div');
            handle.className = `crop-handle ${pos}`;
            handle.style.position = 'absolute';
            handle.style.width = '10px';
            handle.style.height = '10px';
            handle.style.backgroundColor = '#007bff';
            handle.style.border = '2px solid #ffffff';
            handle.style.borderRadius = '50%';
            handle.style.zIndex = '1003';
            overlay.appendChild(handle);
        });
        
        // Add the overlay to the document
        document.body.appendChild(overlay);
        this.cropOverlay = overlay;
        
        // Show crop control buttons
        document.getElementById('cropImageBtn').style.display = 'none';
        document.getElementById('cancelCropBtn').style.display = 'block';
        document.getElementById('applyCropBtn').style.display = 'block';
        
        // Setup drag and resize events
        this.setupCropEvents(overlay);
    }
    
    setupCropEvents(overlay) {
        let isDragging = false;
        let isResizing = false;
        let currentHandle = null;
        let startX, startY, startWidth, startHeight, startLeft, startTop;
        
        // Mouse down on overlay (for moving)
        overlay.addEventListener('mousedown', (e) => {
            if (e.target === overlay) {
                isDragging = true;
                startX = e.clientX;
                startY = e.clientY;
                startLeft = parseInt(overlay.style.left);
                startTop = parseInt(overlay.style.top);
                e.preventDefault();
                e.stopPropagation();
            }
        });
        
        // Mouse down on handle (for resizing)
        const handles = overlay.querySelectorAll('.crop-handle');
        handles.forEach(handle => {
            handle.addEventListener('mousedown', (e) => {
                isResizing = true;
                currentHandle = handle;
                startX = e.clientX;
                startY = e.clientY;
                startWidth = parseInt(overlay.style.width);
                startHeight = parseInt(overlay.style.height);
                startLeft = parseInt(overlay.style.left);
                startTop = parseInt(overlay.style.top);
                e.preventDefault();
                e.stopPropagation();
            });
        });
        
        // Mouse move (for both moving and resizing)
        const mouseMoveHandler = (e) => {
            if (isDragging) {
                // Move the overlay, constrained to image bounds
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                
                // Get image bounds to constrain movement
                const imgRect = this.currentEditingImage.getBoundingClientRect();
                const imgScrollTop = window.pageYOffset || document.documentElement.scrollTop;
                const imgScrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
                const imgLeft = imgRect.left + imgScrollLeft;
                const imgTop = imgRect.top + imgScrollTop;
                const imgRight = imgLeft + imgRect.width;
                const imgBottom = imgTop + imgRect.height;
                
                const overlayWidth = parseInt(overlay.style.width);
                const overlayHeight = parseInt(overlay.style.height);
                
                const newLeft = Math.max(imgLeft, Math.min(startLeft + dx, imgRight - overlayWidth));
                const newTop = Math.max(imgTop, Math.min(startTop + dy, imgBottom - overlayHeight));
                
                overlay.style.left = newLeft + 'px';
                overlay.style.top = newTop + 'px';
            } else if (isResizing && currentHandle) {
                // Resize the overlay based on which handle is being dragged
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                
                // Get image bounds to constrain crop area
                const imgRect = this.currentEditingImage.getBoundingClientRect();
                const imgScrollTop = window.pageYOffset || document.documentElement.scrollTop;
                const imgScrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
                const imgLeft = imgRect.left + imgScrollLeft;
                const imgTop = imgRect.top + imgScrollTop;
                const imgRight = imgLeft + imgRect.width;
                const imgBottom = imgTop + imgRect.height;
                
                if (currentHandle.classList.contains('br')) {
                    const newWidth = Math.max(20, Math.min(startWidth + dx, imgRight - startLeft));
                    const newHeight = Math.max(20, Math.min(startHeight + dy, imgBottom - startTop));
                    overlay.style.width = newWidth + 'px';
                    overlay.style.height = newHeight + 'px';
                } else if (currentHandle.classList.contains('bl')) {
                    const newLeft = Math.max(imgLeft, Math.min(startLeft + dx, startLeft + startWidth - 20));
                    const newWidth = startLeft + startWidth - newLeft;
                    const newHeight = Math.max(20, Math.min(startHeight + dy, imgBottom - startTop));
                    overlay.style.width = newWidth + 'px';
                    overlay.style.left = newLeft + 'px';
                    overlay.style.height = newHeight + 'px';
                } else if (currentHandle.classList.contains('tr')) {
                    const newWidth = Math.max(20, Math.min(startWidth + dx, imgRight - startLeft));
                    const newTop = Math.max(imgTop, Math.min(startTop + dy, startTop + startHeight - 20));
                    const newHeight = startTop + startHeight - newTop;
                    overlay.style.width = newWidth + 'px';
                    overlay.style.height = newHeight + 'px';
                    overlay.style.top = newTop + 'px';
                } else if (currentHandle.classList.contains('tl')) {
                    const newLeft = Math.max(imgLeft, Math.min(startLeft + dx, startLeft + startWidth - 20));
                    const newTop = Math.max(imgTop, Math.min(startTop + dy, startTop + startHeight - 20));
                    const newWidth = startLeft + startWidth - newLeft;
                    const newHeight = startTop + startHeight - newTop;
                    overlay.style.width = newWidth + 'px';
                    overlay.style.left = newLeft + 'px';
                    overlay.style.height = newHeight + 'px';
                    overlay.style.top = newTop + 'px';
                }
            }
        };
        
        // Mouse up (end drag/resize)
        const mouseUpHandler = () => {
            isDragging = false;
            isResizing = false;
            currentHandle = null;
        };
        
        // Add event listeners to document
        document.addEventListener('mousemove', mouseMoveHandler);
        document.addEventListener('mouseup', mouseUpHandler);
        
        // Store references to remove later
        this.cropEventHandlers = {
            mousemove: mouseMoveHandler,
            mouseup: mouseUpHandler
        };
    }
    
    cancelImageCropping() {
        if (this.cropOverlay) {
            document.body.removeChild(this.cropOverlay);
            this.cropOverlay = null;
        }
        
        // Remove event listeners
        if (this.cropEventHandlers) {
            document.removeEventListener('mousemove', this.cropEventHandlers.mousemove);
            document.removeEventListener('mouseup', this.cropEventHandlers.mouseup);
            this.cropEventHandlers = null;
        }
        
        // Reset UI
        document.getElementById('cropImageBtn').style.display = 'block';
        document.getElementById('cancelCropBtn').style.display = 'none';
        document.getElementById('applyCropBtn').style.display = 'none';
        
        // Restore draggable functionality if it was temporarily disabled
        if (this.currentEditingImage) {
            const wrapper = this.currentEditingImage.closest('.image-wrapper');
            if (wrapper && wrapper.hasAttribute('data-temp-draggable')) {
                wrapper.setAttribute('draggable', wrapper.getAttribute('data-temp-draggable'));
                wrapper.removeAttribute('data-temp-draggable');
                
                // Re-add mousedown handler if it exists
                if (wrapper.dragHandlers && wrapper.dragHandlers.mousedown) {
                    wrapper.addEventListener('mousedown', wrapper.dragHandlers.mousedown);
                }
            }
        }
        
        // Show resize and rotation handles again
        this.showResizeAndRotationHandles();
    }
    
    applyImageCropping() {
        if (!this.currentEditingImage || !this.cropOverlay) return;
        
        // Get the crop dimensions
        const imgRect = this.currentEditingImage.getBoundingClientRect();
        const cropRect = this.cropOverlay.getBoundingClientRect();
        const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
        const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
        
        // Validate that we have valid dimensions
        if (imgRect.width <= 0 || imgRect.height <= 0 || cropRect.width <= 0 || cropRect.height <= 0) {
            alert('Dimensions de recadrage invalides. Veuillez réessayer.');
            this.cancelImageCropping();
            return;
        }
        
        // Calculate crop coordinates relative to the image
        // We need to scale the coordinates based on the actual image dimensions vs displayed dimensions
        const scaleX = this.currentEditingImage.naturalWidth / imgRect.width;
        const scaleY = this.currentEditingImage.naturalHeight / imgRect.height;
        
        const cropLeft = Math.max(0, (cropRect.left - imgRect.left) * scaleX);
        const cropTop = Math.max(0, (cropRect.top - imgRect.top) * scaleY);
        const cropWidth = Math.min(cropRect.width * scaleX, this.currentEditingImage.naturalWidth - cropLeft);
        const cropHeight = Math.min(cropRect.height * scaleY, this.currentEditingImage.naturalHeight - cropTop);
        
        // Validate crop dimensions
        if (cropWidth <= 0 || cropHeight <= 0) {
            alert('Zone de recadrage trop petite. Veuillez sélectionner une zone plus grande.');
            this.cancelImageCropping();
            return;
        }
        
        // Create a canvas to perform the crop
        const canvas = document.createElement('canvas');
        // Ensure we have positive dimensions for the canvas
        canvas.width = Math.max(1, Math.round(cropWidth));
        canvas.height = Math.max(1, Math.round(cropHeight));
        const ctx = canvas.getContext('2d');
        
        // Draw the cropped portion of the image onto the canvas
        try {
            // Ensure all parameters are valid numbers
            const sourceX = Math.max(0, Math.round(cropLeft));
            const sourceY = Math.max(0, Math.round(cropTop));
            const sourceWidth = Math.min(Math.round(cropWidth), this.currentEditingImage.naturalWidth - sourceX);
            const sourceHeight = Math.min(Math.round(cropHeight), this.currentEditingImage.naturalHeight - sourceY);
            
            // Only proceed if we have valid dimensions
            if (sourceWidth > 0 && sourceHeight > 0 && sourceX < this.currentEditingImage.naturalWidth && sourceY < this.currentEditingImage.naturalHeight) {
                ctx.drawImage(
                    this.currentEditingImage,
                    sourceX, sourceY, sourceWidth, sourceHeight,
                    0, 0, canvas.width, canvas.height
                );
            } else {
                throw new Error('Invalid crop dimensions or position');
            }
            
            // Replace the original image with the cropped version
            const croppedImageDataUrl = canvas.toDataURL('image/png');
            
            // If the image had an explicit fixed height before cropping, reset it to preserve aspect ratio
            const hadFixedHeight = this.currentEditingImage && this.currentEditingImage.style && this.currentEditingImage.style.height && this.currentEditingImage.style.height !== '' && this.currentEditingImage.style.height !== 'auto';
            // Preserve current width (prefer explicit style width; fallback to current rendered width)
            const preservedStyleWidth = this.currentEditingImage && this.currentEditingImage.style ? this.currentEditingImage.style.width : '';
            const preservedRenderedWidth = this.currentEditingImage ? this.currentEditingImage.offsetWidth : null;

            // Store the new image as the original for future resizing
            const newImg = new Image();
            newImg.onload = () => {
                this.originalImageWidth = newImg.width;
                this.originalImageHeight = newImg.height;
                
                // Update the resize slider to 100% since we're starting fresh with the cropped image (if present)
                const resizeSlider = document.getElementById('resizeSlider');
                const resizePercentage = document.getElementById('resizePercentage');
                if (resizeSlider) resizeSlider.value = 100;
                if (resizePercentage) resizePercentage.textContent = '100%';

                // After the actual <img> element loads the cropped data, adjust styles to avoid deformation
                if (this.currentEditingImage) {
                    // Attach a one-time onload to ensure dimensions are applied after the new data URL is rendered
                    const imgEl = this.currentEditingImage;
                    const applyDimensionFix = () => {
                        try {
                            if (hadFixedHeight) {
                                // Keep width as-is (style if set, otherwise preserve current rendered px width), and reset height to auto
                                if (preservedStyleWidth && preservedStyleWidth.trim() !== '') {
                                    imgEl.style.width = preservedStyleWidth;
                                } else if (preservedRenderedWidth && !isNaN(preservedRenderedWidth)) {
                                    imgEl.style.width = preservedRenderedWidth + 'px';
                                }
                                imgEl.style.height = 'auto';
                            }
                            // Explicitly remove rounded corners and shadow after crop
                            imgEl.style.borderRadius = '0';
                            imgEl.style.boxShadow = 'none';
                        } catch (_) {}
                        // remove handler
                        try { imgEl.onload = null; } catch (_) {}
                    };
                    // If the image might load instantly, set the handler before changing src below
                    imgEl.onload = applyDimensionFix;
                }
            };
            newImg.src = croppedImageDataUrl;
            
            // Apply the cropped image to the current image
            this.currentEditingImage.src = croppedImageDataUrl;
            
            // Clean up
            this.cancelImageCropping();
            
            // Show resize and rotation handles again
            this.showResizeAndRotationHandles();
            
            this.saveState();
            this.lastAction = 'Image recadrée';
        } catch (error) {
            console.error('Error during image cropping:', error);
            const msg = (error && (error.name || '')).toLowerCase().includes('security') || (String(error || '').toLowerCase().includes('tainted'))
                ? "Impossible de recadrer cette image car elle provient d'un domaine externe sans autorisation (CORS). Téléchargez l'image localement ou utilisez une image hébergée avec CORS activé, puis réessayez."
                : "Une erreur est survenue lors du recadrage de l'image. Veuillez réessayer.";
            alert(msg);
            this.cancelImageCropping();
        }
    }

    openImageSelector(targetContainer) {
        console.log('openImageSelector called with:', targetContainer);
        this.currentTargetContainer = targetContainer;
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.onchange = async (e) => {
            console.log('File selected:', e.target.files[0]);
            const file = e.target.files[0];
            if (file) {
                try {
                    const dataUrl = await this.compressImageFile(file, { maxWidth: 1600, maxHeight: 1200, quality: 0.9 });
                    console.log('File loaded, calling insertImageIntoColumn');
                    this.insertImageIntoColumn(dataUrl, file.name, targetContainer);
                } catch (err) {
                    const reader = new FileReader();
                    reader.onload = (ev) => this.insertImageIntoColumn(ev.target.result, file.name, targetContainer);
                    reader.readAsDataURL(file);
                }
            }
        };
        console.log('Triggering file input click');
        input.click();
    }

    insertImageIntoColumn(src, altText, targetContainer) {
        console.log('insertImageIntoColumn called with:', { src: src.substring(0, 50) + '...', altText, targetContainer });
        const img = document.createElement('img');
        img.src = src;
        img.alt = altText;
        img.style.width = '100%';
        img.style.height = 'auto';
        img.style.maxHeight = '400px';
        img.style.objectFit = 'contain';
        img.style.borderRadius = '8px';
        img.style.boxShadow = '0 2px 8px rgba(0,0,0,0.1)';
        
        // Replace the placeholder content with the image
        console.log('Replacing placeholder content with image');
        targetContainer.innerHTML = '';
        targetContainer.appendChild(img);
        
        // Make the image clickable for editing
        img.addEventListener('click', (e) => {
            e.stopPropagation();
            this.selectImage(img);
        });
        
        console.log('Image inserted successfully');
        this.saveState();
        this.updateLastModified();
        this.autoSaveToLocalStorage();
        this.lastAction = 'Image insérée';
    }

    insertImage(src, altText) {
        const img = document.createElement('img');
        // If inserting a remote URL image, attempt CORS-safe loading to support cropping
        try {
            const isHttp = /^https?:\/\//i.test(src);
            if (isHttp) {
                // Only set crossOrigin for cross-origin URLs
                const a = document.createElement('a');
                a.href = src;
                if (a.host && a.host !== window.location.host) {
                    img.crossOrigin = 'anonymous';
                    img.referrerPolicy = 'no-referrer';
                }
            }
        } catch (_) {}
        img.src = src;
        img.alt = altText;
        try {
            // Hint for preview: remember the local filename if provided
            if (altText && /\.[a-z0-9]{2,4}$/i.test(String(altText))) {
                img.setAttribute('data-local-filename', String(altText));
            }
        } catch (_) {}
        
        // Apply default image styles
        img.style.cssText = 'max-width: 100%; height: auto; margin: 10px 0; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);';
        
        // Store original image source for reset functionality
        img.dataset.originalSrc = src;
        
        // Create a wrapper for the image to contain resize handles and delete button
        const wrapper = document.createElement('div');
        wrapper.className = 'image-wrapper position-inline';
        wrapper.style.cssText = 'position: relative; display: block; margin: 15px auto;';
        
        // Add mouse position tracking for the wrapper
        wrapper.addEventListener('mousedown', (e) => {
            // Store the mouse position when clicking on the image
            this.lastMousePosition = { x: e.clientX, y: e.clientY };
            e.stopPropagation();
        });
        wrapper.appendChild(img);
        
        // Create delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.innerHTML = '<i class="fas fa-times"></i>';
        deleteBtn.className = 'image-delete-btn';
        deleteBtn.title = 'Supprimer l\'image';
        deleteBtn.style.cssText = `
            position: absolute;
            top: -10px;
            right: -10px;
            width: 24px;
            height: 24px;
            border-radius: 50%;
            background-color: #dc3545;
            color: white;
            border: 2px solid white;
            cursor: pointer;
            display: none;
            z-index: 1000;
            font-size: 12px;
            line-height: 1;
            box-shadow: 0 2px 4px rgba(0,0,0,0.3);
        `;

        // Clear default gallery title on first focus/click (same behavior as two-column)
        try {
            const h3 = gallerySection.querySelector('.gallery-title[contenteditable]');
            if (h3) {
                const clear = () => {
                    try {
                        if (h3.dataset.cleared) return;
                        if ((h3.textContent || '').trim() !== "Galerie d'images") return;
                        h3.innerHTML = '<br>';
                        h3.dataset.cleared = '1';
                        const range = document.createRange();
                        range.selectNodeContents(h3);
                        range.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                        try { h3.focus(); } catch (_) {}
                    } catch (_) {}
                };
                h3.addEventListener('focus', clear);
                h3.addEventListener('click', clear);
            }
        } catch (_) { /* no-op */ }
        
        // Add delete button click event
        deleteBtn.addEventListener('click', async (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (await this.confirmWithCancel('Êtes-vous sûr de vouloir supprimer cette image ?')) {
                this.deleteImageWrapperSafe(wrapper);
            }
        });
        
        wrapper.appendChild(deleteBtn);
        
        // Show/hide delete button on hover
        wrapper.addEventListener('mouseenter', () => {
            deleteBtn.style.display = 'block';
        });
        
        wrapper.addEventListener('mouseleave', () => {
            deleteBtn.style.display = 'none';
        });
        
        // Add click event to make the image editable (disabled for gallery images)
        img.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (img.closest && img.closest('.gallery-section')) {
                return; // gallery images managed by gallery tools only
            }
            const rich = document.getElementById('richTextToolbar');
            if (rich) rich.style.display = 'none';
            const tableTb = document.getElementById('tableToolbar');
            if (tableTb) tableTb.style.display = 'none';
            this.showImageEditingTools(img);
        });
        
        // Add keyboard support for deletion
        wrapper.addEventListener('keydown', (e) => {
            if (e.key === 'Delete' || e.key === 'Backspace') {
                e.preventDefault();
                e.stopPropagation();
                (async () => { if (await this.confirmWithCancel('Êtes-vous sûr de vouloir supprimer cette image ?')) { this.deleteImageWrapperSafe(wrapper); } })();
            }
        });
        
        // Make wrapper focusable for keyboard events
        wrapper.setAttribute('tabindex', '0');
        
        // Enable flow drag-to-reposition (not absolute move)
        try { this.makeImageFlowReorderable(wrapper); } catch (_) {}

        this.insertElementAtCursor(wrapper);
        // Immediately select the inserted image so floating tools (e.g., rotation) work
        try { this.showImageEditingTools(img); } catch (_) {}
        this.saveState();
        this.lastAction = 'Image insérée';
    }

    wrapVideoElementForEditing(mediaEl) {
        if (!mediaEl) return null;
        const wrapper = document.createElement('div');
        wrapper.className = 'video-align-wrapper';
        wrapper.style.cssText = 'position: relative; display: block; margin: 0 auto; max-width: 100%;';
        wrapper.setAttribute('tabindex', '0');
        const handle = document.createElement('button');
        handle.type = 'button';
        handle.className = 'video-toolbar-handle';
        handle.innerHTML = '<i class="fas fa-ellipsis-v"></i>';
        handle.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.showVideoToolbar(mediaEl);
        });
        wrapper.addEventListener('click', (e) => {
            if (e.target === wrapper) {
                e.preventDefault();
                this.showVideoToolbar(mediaEl);
            }
        });
        wrapper.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                this.showVideoToolbar(mediaEl);
            }
        });
        try {
            mediaEl.addEventListener('click', (e) => {
                e.stopPropagation();
                this.showVideoToolbar(mediaEl);
            });
        } catch (_) {}
        wrapper.appendChild(handle);
        wrapper.appendChild(mediaEl);
        return wrapper;
    }

    insertVideo(url) {
        const trimmedUrl = (url || '').trim();
        if (!trimmedUrl) return;
        let mediaEl = null;
        if (trimmedUrl.includes('youtube.com') || trimmedUrl.includes('youtu.be')) {
            const videoId = this.extractYouTubeId(trimmedUrl);
            const iframe = document.createElement('iframe');
            iframe.src = videoId ? `https://www.youtube.com/embed/${videoId}` : trimmedUrl;
            iframe.setAttribute('frameborder', '0');
            iframe.setAttribute('allow', 'accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share');
            iframe.allowFullscreen = true;
            iframe.style.border = 'none';
            mediaEl = iframe;
        } else if (trimmedUrl.includes('vimeo.com')) {
            const videoId = this.extractVimeoId(trimmedUrl);
            const iframe = document.createElement('iframe');
            iframe.src = videoId ? `https://player.vimeo.com/video/${videoId}` : trimmedUrl;
            iframe.setAttribute('frameborder', '0');
            iframe.setAttribute('allow', 'autoplay; fullscreen; picture-in-picture');
            iframe.allowFullscreen = true;
            iframe.style.border = 'none';
            mediaEl = iframe;
        } else {
            const video = document.createElement('video');
            video.controls = true;
            video.preload = 'metadata';
            const source = document.createElement('source');
            source.src = trimmedUrl;
            const lower = trimmedUrl.toLowerCase();
            if (lower.endsWith('.webm')) source.type = 'video/webm';
            else if (lower.endsWith('.ogg') || lower.endsWith('.ogv')) source.type = 'video/ogg';
            else source.type = 'video/mp4';
            video.appendChild(source);
            mediaEl = video;
        }
        if (!mediaEl) return;
        const wrapper = this.wrapVideoElementForEditing(mediaEl);
        this.insertElementAtCursor(wrapper || mediaEl);
        this.normalizeVideoStyles(wrapper || mediaEl);
        this.saveState();
        this.lastAction = 'Vidéo insérée';
        if (!window.__disableVideoPoster) {
            try {
                const host = document.getElementById('editableContent');
                const vids = host ? host.querySelectorAll('video') : [];
                const lastVideo = vids && vids[vids.length - 1];
                if (lastVideo && !lastVideo.getAttribute('poster')) {
                    const ensurePoster = async (v) => {
                        try {
                            const makePoster = () => {
                                try {
                                    const vw = v.videoWidth || 0, vh = v.videoHeight || 0;
                                    if (!vw || !vh) return '';
                                    const w = 800, h = Math.max(1, Math.round(w / (vw / (vh || 1))));
                                    const canvas = document.createElement('canvas');
                                    canvas.width = w; canvas.height = h;
                                    const ctx = canvas.getContext('2d');
                                    ctx.drawImage(v, 0, 0, w, h);
                                    return canvas.toDataURL('image/jpeg', 0.86);
                                } catch { return ''; }
                            };
                            const drawNow = () => {
                                const dataUrl = makePoster();
                                if (dataUrl) v.setAttribute('poster', dataUrl);
                            };
                            if (lastVideo.readyState >= 2) {
                                try { v.currentTime = Math.min(0.1, (v.seekable && v.seekable.length ? v.seekable.end(0) : 0.1)); } catch {}
                                v.addEventListener('seeked', drawNow, { once: true });
                            } else {
                                v.addEventListener('loadeddata', () => {
                                    try { v.currentTime = Math.min(0.1, (v.seekable && v.seekable.length ? v.seekable.end(0) : 0.1)); } catch {}
                                    v.addEventListener('seeked', drawNow, { once: true });
                                }, { once: true });
                            }
                            setTimeout(drawNow, 1500);
                        } catch { /* ignore */ }
                    };
                    ensurePoster(lastVideo);
                }
            } catch (_) { /* ignore */ }
        }
    }

    insertLocalVideo(src, name) {
        const video = document.createElement('video');
        video.controls = true;
        video.preload = 'auto';
        if (name) {
            try { video.setAttribute('data-local-filename', String(name)); } catch (_) {}
        }
        const persistent = document.createElement('source');
        const safeName = String(name || '').replace(/^[\\\/]+/, '');
        const mediaPath = safeName ? ('media/' + encodeURIComponent(safeName)) : '';
        if (mediaPath) persistent.src = mediaPath;
        try {
            const lower = String(name || '').toLowerCase();
            if (lower.endsWith('.mp4')) persistent.type = 'video/mp4';
            else if (lower.endsWith('.webm')) persistent.type = 'video/webm';
            else if (lower.endsWith('.ogg') || lower.endsWith('.ogv')) persistent.type = 'video/ogg';
        } catch (_) {}
        if (mediaPath) video.appendChild(persistent);

        const blobSource = document.createElement('source');
        blobSource.src = src;
        if (name) {
            try { blobSource.setAttribute('data-local-filename', String(name)); } catch (_) {}
        }
        try {
            const lower = String(name || '').toLowerCase();
            if (lower.endsWith('.mp4')) blobSource.type = 'video/mp4';
            else if (lower.endsWith('.webm')) blobSource.type = 'video/webm';
            else if (lower.endsWith('.ogg') || lower.endsWith('.ogv')) blobSource.type = 'video/ogg';
        } catch (_) {}
        video.appendChild(blobSource);

        const wrapper = this.wrapVideoElementForEditing(video);
        this.insertElementAtCursor(wrapper || video);
        this.normalizeVideoStyles(wrapper || video);
        this.saveState();
        this.lastAction = 'Vidéo insérée';

        if (!window.__disableVideoPoster) {
            try {
                const ensurePoster = (v) => {
                    const draw = () => {
                        try {
                            const vw = v.videoWidth || 0, vh = v.videoHeight || 0;
                            if (!vw || !vh) return;
                            const w = 800, h = Math.max(1, Math.round(w / (vw / (vh || 1))));
                            const canvas = document.createElement('canvas');
                            canvas.width = w; canvas.height = h;
                            const ctx = canvas.getContext('2d');
                            ctx.drawImage(v, 0, 0, w, h);
                            const dataUrl = canvas.toDataURL('image/jpeg', 0.86);
                            if (dataUrl) v.setAttribute('poster', dataUrl);
                        } catch { /* ignore */ }
                    };
                    if (v.readyState >= 2) {
                        try { v.currentTime = Math.min(0.1, (v.seekable && v.seekable.length ? v.seekable.end(0) : 0.1)); } catch {}
                        v.addEventListener('seeked', draw, { once: true });
                    } else {
                        v.addEventListener('loadeddata', () => {
                            try { v.currentTime = Math.min(0.1, (v.seekable && v.seekable.length ? v.seekable.end(0) : 0.1)); } catch {}
                            v.addEventListener('seeked', draw, { once: true });
                        }, { once: true });
                    }
                    setTimeout(draw, 1500);
                };
                ensurePoster(video);
            } catch (_) { /* ignore */ }
        }
    }

    // Normalize all embedded videos/iframes to be 70% width, centered, with proper height
    normalizeVideoStyles(root) {
        const host = root || document.getElementById('editableContent');
        if (!host) return;
        const nodes = host.querySelectorAll('iframe, video');
        nodes.forEach((n) => {
            try {
                // Remove fixed attributes that might constrain responsive layout
                if (n.hasAttribute('width')) n.removeAttribute('width');
                if (n.hasAttribute('height')) n.removeAttribute('height');
                // Apply consistent sizing and centering
                n.style.width = '70%';
                n.style.margin = '10px auto';
                n.style.display = 'block';
                // Ensure proper height behavior
                n.style.height = 'auto';
                // Help browsers maintain aspect ratio (works for iframes and videos in modern browsers)
                try { n.style.aspectRatio = '16 / 9'; } catch (_) {}
            } catch (_) {}
        });
    }

    extractYouTubeId(url) {
        const regExp = /^.*(youtu.be\/|v\/|u\/\w\/|embed\/|watch\?v=|&v=)([^#&?]*).*/;
        const match = url.match(regExp);
        return (match && match[2].length === 11) ? match[2] : null;
    }

    extractVimeoId(url) {
        const regExp = /vimeo.com\/(\d+)/;
        const match = url.match(regExp);
        return match ? match[1] : null;
    }

    insertTable() { return; }

    insertArticleSection() {
        const articleSection = document.createElement('div');
        articleSection.className = 'newsletter-section article-section';
        articleSection.style.cssText = 'margin: 30px 0; padding: 25px; background-color: #ffffff; border-left: 4px solid #007bff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);';
        
        articleSection.innerHTML = `
            <div class="article-header" style="margin-bottom: 20px;">
                <h2 contenteditable="true" style="color: #333; margin: 0 0 10px 0; font-size: 24px;">Titre de l'article</h2>
                <div class="article-meta" style="color: #666; font-size: 14px; border-bottom: 1px solid #eee; padding-bottom: 15px;">
                    <span contenteditable="true">Par: Nom de l'auteur</span> | 
                    <span contenteditable="true">Catégorie: Actualités</span> | 
                    <span>${new Date().toLocaleDateString('fr-FR')}</span>
                </div>
            </div>
            <div class="article-content" style="line-height: 1.8;">
                <p contenteditable="true" style="margin-bottom: 15px;">Cliquez ici pour écrire le contenu de votre article. Vous pouvez ajouter plusieurs paragraphes, des images et du formatage.</p>
                <p contenteditable="true" style="margin-bottom: 15px;">Deuxième paragraphe de votre article. Continuez à développer votre contenu ici.</p>
            </div>
            <div class="article-footer" style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #eee;">
                <p contenteditable="true" style="color: #888; font-size: 14px;">Publié le: <span style="font-weight: bold;">${new Date().toLocaleDateString('fr-FR')}</span></p>
            </div>
        `;

        // Apply two-column style pretext clearing to Article default texts
        try {
            const clearOnFocus = (el, isDefault) => {
                const clear = () => {
                    try {
                        if (el.dataset.cleared) return;
                        if (!isDefault()) return;
                        // Clear content but keep block/format context
                        if (el.tagName === 'P' || el.tagName === 'H1' || el.tagName === 'H2' || el.tagName === 'H3') {
                            el.innerHTML = '<br>';
                        } else {
                            el.textContent = '';
                        }
                        el.dataset.cleared = '1';
                        // Place caret at start
                        const range = document.createRange();
                        range.selectNodeContents(el);
                        range.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                        try { el.focus(); } catch (_) {}
                    } catch (_) { /* no-op */ }
                };
                el.addEventListener('focus', clear, true);
                el.addEventListener('click', clear, true);
            };

            const titleEl = articleSection.querySelector('h2[contenteditable]');
            if (titleEl) clearOnFocus(titleEl, () => (titleEl.textContent || '').trim() === "Titre de l'article");

            const pEls = articleSection.querySelectorAll('.article-content p[contenteditable]');
            const defaults = [
                'Cliquez ici pour écrire le contenu de votre article. Vous pouvez ajouter plusieurs paragraphes, des images et du formatage.',
                "Deuxième paragraphe de votre article. Continuez à développer votre contenu ici."
            ];
            pEls.forEach((p, idx) => {
                const expected = defaults[idx] || '';
                clearOnFocus(p, () => (p.textContent || '').trim() === expected);
            });

            const authorEl = articleSection.querySelector('.article-meta span[contenteditable]:nth-child(1)');
            if (authorEl) clearOnFocus(authorEl, () => (authorEl.textContent || '').trim().startsWith('Par: '));
            const catEl = articleSection.querySelector('.article-meta span[contenteditable]:nth-child(3)');
            if (catEl) clearOnFocus(catEl, () => (catEl.textContent || '').trim().startsWith('Catégorie: '));
        } catch (_) { /* no-op */ }
        
        this.insertElementAtCursor(articleSection);
        this.saveState();
        try { this.showSectionToolbar(articleSection); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section article insérée';
    }

    insertGallerySection() {
        const gallerySection = document.createElement('div');
        gallerySection.className = 'gallery-section';
        
        gallerySection.innerHTML = `
            <h3 class="gallery-title" contenteditable="true">Galerie d'images</h3>
            <div class="gallery-grid">
                <div class="gallery-item add-image-placeholder">
                    <i class="fas fa-plus"></i>
                    <p contenteditable="false">Cliquez pour ajouter des images</p>
                </div>
            </div>
            <input type="file" class="gallery-upload" multiple accept="image/*" style="display: none;">
            <button class="add-more-btn" style="display: none; margin-top: 10px; padding: 8px 16px; background: #0a9bcd; color: white; border: none; border-radius: 4px; cursor: pointer;">
                <i class="fas fa-plus"></i> Ajouter plus d'images
            </button>
        `;
        
        const galleryGrid = gallerySection.querySelector('.gallery-grid');
        const addImagePlaceholder = gallerySection.querySelector('.add-image-placeholder');
        const fileInput = gallerySection.querySelector('.gallery-upload');
        
        const addMoreBtn = gallerySection.querySelector('.add-more-btn');
        
        // Function to handle adding images
        const handleAddImage = () => {
            fileInput.click();
        };
        
        // Click handler for adding images (initial add button)
        addImagePlaceholder.addEventListener('click', handleAddImage);
        
        // Click handler for the "add more" button
        addMoreBtn.addEventListener('click', handleAddImage);
        
        // Handle file selection
        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            if (files.length > 0) {
                // Hide the initial add button and show the "add more" button
                addImagePlaceholder.style.display = 'none';
                addMoreBtn.style.display = 'block';
                
                files.forEach(file => {
                    if (file.type.startsWith('image/')) {
                        this.addImageToGallery(file, galleryGrid);
                    }
                });
                // Reset file input
                fileInput.value = '';
            }
        });
        
        this.insertElementAtCursor(gallerySection);
        this.saveState();
        try { this.showSectionToolbar(gallerySection); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section galerie insérée';
    }

    addImageToGallery(file, galleryGrid) {
        const handleLoaded = (dataUrl) => {
            // Create image container with proper classes for editing
            const imageContainer = document.createElement('div');
            imageContainer.className = 'gallery-item';
            imageContainer.contentEditable = false;
            imageContainer.style.position = 'relative';
            
            // Create image element with same properties as regular images
            const img = document.createElement('img');
            img.src = dataUrl;
            img.alt = file.name || 'Gallery Image';
            img.draggable = false;
            img.style.width = '100%';
            img.style.height = '100%';
            img.style.objectFit = 'contain';
            img.style.borderRadius = '8px';
            // Enable standard image tools on click (floating, inline, absolute, etc.)
            img.addEventListener('click', (e) => {
                e.stopPropagation();
                try { this.showImageEditingTools(img); } catch (_) {}
            });
            
            // Add image to container only (no overlay)
            imageContainer.appendChild(img);
            
            // Remove old hover overlay behavior (no-op)
            
            // Remove custom delete/replace overlay controls (standard image toolbar will handle actions)
            
            // Allow clicks to propagate to image tools via img handler above
            
            // Add drag and drop for reordering
            imageContainer.draggable = true;
            imageContainer.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('text/plain', '');
                imageContainer.classList.add('dragging');
                setTimeout(() => {
                    imageContainer.style.opacity = '0.4';
                }, 0);
            });
            
            imageContainer.addEventListener('dragend', () => {
                imageContainer.classList.remove('dragging');
                imageContainer.style.opacity = '1';
            });
            
            // Create image wrapper
            const imageWrapper = document.createElement('div');
            imageWrapper.className = 'gallery-image-wrapper';
            imageWrapper.style.position = 'relative';
            imageWrapper.style.flex = '1';
            imageWrapper.style.display = 'flex';
            imageWrapper.style.flexDirection = 'column';
            
            // Add image to wrapper
            const imageContainerInner = document.createElement('div');
            imageContainerInner.className = 'gallery-image-container';
            imageContainerInner.style.position = 'relative';
            imageContainerInner.style.width = '100%';
            // Keep original image aspect ratio: no forced square
            imageContainerInner.style.paddingBottom = '';
            imageContainerInner.style.overflow = 'hidden';
            
            // Style the image to preserve aspect ratio
            img.style.position = 'static';
            img.style.top = '';
            img.style.left = '';
            img.style.width = '100%';
            img.style.height = 'auto';
            img.style.objectFit = 'contain';
            
            // Add image to container
            imageContainerInner.appendChild(img);
            // No overlay appended
            
            // Create description area
            const description = document.createElement('div');
            description.className = 'gallery-description';
            description.contentEditable = true;
            description.setAttribute('data-placeholder', 'Cliquez pour ajouter une description');
            description.style.userSelect = 'text';
            description.style.pointerEvents = 'auto';
            
            // Add focus/blur handlers for the description
            const handleFocus = (e) => {
                e.stopPropagation();
                if (description.textContent === 'Cliquez pour ajouter une description') {
                    description.textContent = '';
                }
                // Prevent the gallery item click handler from firing
                e.stopImmediatePropagation();
            };
            
            const handleBlur = (e) => {
                e.stopPropagation();
                if (!description.textContent.trim()) {
                    description.textContent = 'Cliquez pour ajouter une description';
                } else {
                    this.saveState();
                    this.lastAction = 'Description de galerie modifiée';
                }
            };
            
            // Prevent Enter key from creating new lines
            const handleKeyDown = (e) => {
                e.stopPropagation();
                if (e.key === 'Enter') {
                    e.preventDefault();
                    description.blur();
                }
            };
            
            // Add event listeners
            const stopProp = (e) => {
                e.stopPropagation();
                e.stopImmediatePropagation();
            };
            
            description.addEventListener('mousedown', stopProp, true);
            description.addEventListener('click', stopProp, true);
            description.addEventListener('dblclick', stopProp, true);
            description.addEventListener('focus', handleFocus, true);
            description.addEventListener('blur', handleBlur, true);
            description.addEventListener('keydown', handleKeyDown, true);
            
            // Set initial text
            description.textContent = 'Cliquez pour ajouter une description';
            
            // Add elements to container
            imageWrapper.appendChild(imageContainerInner);
            imageWrapper.appendChild(description);
            imageContainer.appendChild(imageWrapper);
            
            // Remove overlay-related pointer events and hover behavior
            
            // Insert before the add image placeholder
            const addButton = galleryGrid.querySelector('.add-image-placeholder');
            if (addButton) {
                galleryGrid.insertBefore(imageContainer, addButton);
            } else {
                galleryGrid.appendChild(imageContainer);
            }
            
            this.saveState();
            this.lastAction = 'Image de galerie ajoutée';
        };

        // Load original image data without compression for gallery inserts
        const reader = new FileReader();
        reader.onload = (e) => handleLoaded(e.target.result);
        reader.readAsDataURL(file);
    }

    insertMultiImagesSection() {
        const section = document.createElement('div');
        section.className = 'newsletter-section multi-images-section';

        section.innerHTML = `
            <div class="multi-grid">
                <div class="multi-item multi-add">
                    <i class="fas fa-plus"></i>
                    <p>Ajouter des images</p>
                </div>
            </div>
            <input type="file" class="multi-upload" accept="image/*" multiple style="display:none;" />
        `;

        const grid = section.querySelector('.multi-grid');
        const addTile = section.querySelector('.multi-add');
        const input = section.querySelector('.multi-upload');

        // Columns by count: 1/2/3 columns based on number of images
        const updateColumns = () => {
            const items = grid.querySelectorAll('.multi-item:not(.multi-add)');
            const count = items.length;
            const cols = Math.min(Math.max(count, 1), 3);
            grid.style.gridTemplateColumns = `repeat(${cols}, 1fr)`;
            items.forEach(it => { it.style.gridColumn = ''; it.style.gridRowEnd = ''; });
            try { addTile.style.display = (count > 0) ? 'none' : 'flex'; } catch (_) {}
            try { section.classList.toggle('has-items', count > 0); } catch (_) {}
        };

        const attachObservers = (item, mediaEl) => {
            try {
                if (mediaEl) mediaEl.addEventListener('load', () => updateColumns(), { once: false });
                if (window.ResizeObserver) {
                    const ro = new ResizeObserver(() => updateColumns());
                    ro.observe(item);
                }
            } catch (_) {}
        };

        const pick = () => input.click();
        addTile.addEventListener('click', pick);

        input.addEventListener('change', (e) => {
            const files = Array.from(e.target.files || []);
            if (!files.length) return;
            files.forEach(f => {
                if (f.type && f.type.startsWith('image/')) {
                    const reader = new FileReader();
                    reader.onload = (ev) => {
                        const card = document.createElement('div');
                        card.className = 'multi-item';
                        card.contentEditable = false;

                        const img = document.createElement('img');
                        img.src = ev.target.result;
                        img.alt = f.name || 'Image';
                        img.draggable = false;
                        img.addEventListener('click', (ev2) => { ev2.stopPropagation(); try { this.showImageEditingTools(img); } catch (_) {} });
                        // Auto-assign tall variant based on intrinsic dimensions
                        img.addEventListener('load', () => {
                            try {
                                const nh = img.naturalHeight || 0;
                                const nw = img.naturalWidth || 0;
                                if (nh >= 900 || (nw > 0 && nh / nw > 1.2)) {
                                    card.classList.add('tall');
                                }
                            } catch (_) {}
                            updateColumns();
                        });

                        const del = document.createElement('button');
                        del.className = 'multi-delete';
                        del.setAttribute('type', 'button');
                        del.innerHTML = '<i class="fas fa-times"></i>';
                        del.addEventListener('click', (evd) => {
                            evd.stopPropagation();
                            if (card && card.parentNode) {
                                card.parentNode.removeChild(card);
                                updateColumns();
                                // If no images remain, show the add placeholder again
                                try {
                                    const remaining = grid.querySelectorAll('.multi-item:not(.multi-add)').length;
                                    if (remaining === 0) addTile.style.display = 'flex';
                                } catch (_) {}
                                this.saveState();
                                this.lastAction = 'Image multi supprimée';
                            }
                        });

                        const addInline = document.createElement('button');
                        addInline.className = 'multi-add-inline';
                        addInline.setAttribute('type', 'button');
                        addInline.title = 'Ajouter des images';
                        addInline.setAttribute('aria-label', 'Ajouter des images');
                        addInline.innerHTML = '<i class="fas fa-image"></i>';
                        addInline.addEventListener('click', (eai) => { eai.stopPropagation(); pick(); });

                        card.appendChild(img);
                        card.appendChild(del);
                        card.appendChild(addInline);
                        grid.insertBefore(card, addTile);
                        attachObservers(card, img);
                        updateColumns();
                        this.saveState();
                        this.lastAction = 'Images multi ajoutées';
                    };
                    reader.readAsDataURL(f);
                }
            });
            input.value = '';
            updateColumns();
        });

        this.insertElementAtCursor(section);
        this.saveState();
        try { this.showSectionToolbar(section); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section multi images insérée';

        // Initial layout
        updateColumns();
        window.addEventListener('resize', () => updateColumns());
    }
    
    // Helper method to handle drag over for gallery items
    handleGalleryDragOver(e) {
        e.preventDefault();
        const afterElement = this.getDragAfterElement(e.clientY);
        const draggable = document.querySelector('.dragging');
        if (afterElement == null) {
            draggable.parentNode.appendChild(draggable);
        } else {
            draggable.parentNode.insertBefore(draggable, afterElement);
        }
    }
    
    // Helper method to get the element after which to place the dragged item
    getDragAfterElement(y) {
        const draggableElements = [...document.querySelectorAll('.gallery-item:not(.dragging)')];
        
        return draggableElements.reduce((closest, child) => {
            const box = child.getBoundingClientRect();
            const offset = y - box.top - box.height / 2;
            
            if (offset < 0 && offset > closest.offset) {
                return { offset: offset, element: child };
            } else {
                return closest;
            }
        }, { offset: Number.NEGATIVE_INFINITY }).element;
    }

    showImageModal(src, alt) {
        // Create modal
        const modal = document.createElement('div');
        modal.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.9);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10000;
            cursor: pointer;
        `;
        
        // Create image
        const img = document.createElement('img');
        img.src = src;
        img.alt = alt;
        img.style.cssText = `
            max-width: 90%;
            max-height: 90%;
            object-fit: contain;
            border-radius: 8px;
            cursor: default;
        `;
        
        // Create close button
        const closeBtn = document.createElement('button');
        closeBtn.innerHTML = '×';
        closeBtn.style.cssText = `
            position: absolute;
            top: 20px;
            right: 30px;
            background: none;
            border: none;
            color: white;
            font-size: 40px;
            cursor: pointer;
            z-index: 10001;
        `;
        
        // Close modal events
        const closeModal = () => {
            document.body.removeChild(modal);
        };
        
        modal.addEventListener('click', closeModal);
        closeBtn.addEventListener('click', closeModal);
        img.addEventListener('click', (e) => e.stopPropagation());
        
        modal.appendChild(img);
        modal.appendChild(closeBtn);
        document.body.appendChild(modal);
    }

    insertQuoteSection() {
        const quoteSection = document.createElement('div');
        quoteSection.className = 'newsletter-section quote-section';
        quoteSection.style.cssText = 'margin: 30px 0; padding: 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px; color: white; position: relative;';
        
        quoteSection.innerHTML = `
            <div style="position: absolute; top: 15px; left: 25px; font-size: 48px; opacity: 0.3;">
                <i class="fas fa-tags"></i>
            </div>
            <blockquote style="margin: 0; padding-left: 60px; font-size: 18px; line-height: 1.6; font-style: italic;">
                <p contenteditable="true" style="margin-bottom: 20px; color: white;">Mettez en avant votre promotion ici. Décrivez l'offre, les dates et les conditions principales.</p>
            </blockquote>
            <footer style="text-align: right; margin-top: 20px; padding-top: 15px; border-top: 1px solid rgba(255,255,255,0.3);">
                <p contenteditable="true" style="color: rgba(255,255,255,0.9); font-size: 16px;">— Conditions</p>
            </footer>
        `;

        // Clear default quote text and footer label on first focus/click
        try {
            const q = quoteSection.querySelector('blockquote p[contenteditable]');
            if (q) {
                const qDefault = "Mettez en avant votre promotion ici. Décrivez l'offre, les dates et les conditions principales.";
                const clearQ = () => {
                    try {
                        if (q.dataset.cleared) return;
                        if ((q.textContent || '').trim() !== qDefault) return;
                        q.innerHTML = '<br>';
                        q.dataset.cleared = '1';
                        const range = document.createRange();
                        range.selectNodeContents(q);
                        range.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                        try { q.focus(); } catch (_) {}
                    } catch (_) {}
                };
                q.addEventListener('focus', clearQ);
                q.addEventListener('click', clearQ);
            }

            const foot = quoteSection.querySelector('footer p[contenteditable]');
            if (foot) {
                const clearF = () => {
                    try {
                        if (foot.dataset.cleared) return;
                        if ((foot.textContent || '').trim() !== '— Conditions') return;
                        foot.innerHTML = '<br>';
                        foot.dataset.cleared = '1';
                        const range = document.createRange();
                        range.selectNodeContents(foot);
                        range.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                        try { foot.focus(); } catch (_) {}
                    } catch (_) {}
                };
                foot.addEventListener('focus', clearF);
                foot.addEventListener('click', clearF);
            }
        } catch (_) { /* no-op */ }
        
        this.insertElementAtCursor(quoteSection);
        this.saveState();
        try { this.showSectionToolbar(quoteSection); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section citation insérée';
    }

    insertCTASection() {
        const ctaSection = document.createElement('div');
        ctaSection.className = 'newsletter-section cta-section';
        ctaSection.style.cssText = 'margin: 30px 0; padding: 40px; background: linear-gradient(135deg, #ff6b6b, #ee5a24); border-radius: 12px; text-align: center; color: white;';
        
        ctaSection.innerHTML = `
            <h3 contenteditable="true" style="color: white; margin: 0 0 15px 0; font-size: 28px; font-weight: bold;">Annonce</h3>
            <p contenteditable="true" style="color: rgba(255,255,255,0.9); font-size: 18px; margin-bottom: 25px; line-height: 1.6;">Publiez ici une annonce importante. Modifiez ce texte selon votre besoin.</p>
            <span class="webinar-button" contenteditable="false" style="display: inline-block !important; background: white !important; color: #ee5a24 !important; padding: 15px 30px !important; border-radius: 50px !important; text-decoration: none !important; font-weight: bold !important; font-size: 16px !important; transition: transform 0.3s ease !important; box-shadow: 0 4px 15px rgba(0,0,0,0.2) !important; border: none !important; white-space: normal !important; word-break: break-word !important; max-width: 100% !important; text-align: center !important;">
                <span class="btn-text" contenteditable="true" style="display:inline;">Bouton</span>
            </span>
        `;

        // Apply two-column style pretext clearing to CTA default texts
        try {
            const clearOnFocus = (el, isDefault) => {
                const clear = () => {
                    try {
                        if (el.dataset.cleared) return;
                        if (!isDefault()) return;
                        el.innerHTML = '<br>';
                        el.dataset.cleared = '1';
                        const range = document.createRange();
                        range.selectNodeContents(el);
                        range.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                        try { el.focus(); } catch (_) {}
                    } catch (_) { /* no-op */ }
                };
                el.addEventListener('focus', clear, true);
                el.addEventListener('click', clear, true);
            };

            const h3 = ctaSection.querySelector('h3[contenteditable]');
            if (h3) clearOnFocus(h3, () => (h3.textContent || '').trim() === 'Annonce');
            const p = ctaSection.querySelector('p[contenteditable]');
            if (p) clearOnFocus(p, () => (p.textContent || '').trim().startsWith('Publiez ici'));
            const btn = ctaSection.querySelector('.webinar-button');
            const btnText = ctaSection.querySelector('.webinar-button .btn-text[contenteditable]');
            if (btnText) {
                clearOnFocus(btnText, () => (btnText.textContent || '').trim() === 'Bouton');
                // Prevent deleting the button element when clearing its text
                const placeCaretInside = () => {
                    try {
                        const r = document.createRange();
                        r.selectNodeContents(btnText);
                        r.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(r);
                    } catch (_) {}
                };
                const keepButtonOnKeyDown = (e) => {
                    const key = e.key;
                    if (key === 'Enter') {
                        e.preventDefault();
                        e.stopPropagation();
                        return;
                    }
                    if (key === 'Backspace' || key === 'Delete') {
                        const sel = window.getSelection();
                        const inside = sel && sel.anchorNode && btnText.contains(sel.anchorNode);
                        if (!inside) return;
                        const txt = (btnText.textContent || '').trim();
                        // If selection spans content or last char would be removed, keep button
                        if (!sel.isCollapsed || txt.length <= 1) {
                            e.preventDefault();
                            e.stopPropagation();
                            btnText.innerHTML = '<br>';
                            placeCaretInside();
                        }
                    }
                };
                const keepButtonOnBeforeInput = (e) => {
                    const t = e.inputType || '';
                    if (t === 'deleteContentBackward' || t === 'deleteContentForward' || t === 'deleteByCut') {
                        const sel = window.getSelection();
                        const inside = sel && sel.anchorNode && btnText.contains(sel.anchorNode);
                        if (!inside) return;
                        const txt = (btnText.textContent || '').trim();
                        if (!sel || !sel.isCollapsed || txt.length <= 1) {
                            e.preventDefault();
                            btnText.innerHTML = '<br>';
                            placeCaretInside();
                        }
                    }
                };
                const normalizeOnInput = () => {
                    const txt = (btnText.textContent || '').trim();
                    if (!txt) {
                        btnText.innerHTML = '<br>';
                        placeCaretInside();
                    }
                };
                const normalizeOnBlur = () => {
                    const txt = (btnText.textContent || '').trim();
                    if (!txt) {
                        // Restore placeholder but keep the button element for clarity
                        btnText.textContent = 'Bouton';
                    }
                };
                btnText.addEventListener('beforeinput', keepButtonOnBeforeInput, true);
                btnText.addEventListener('keydown', keepButtonOnKeyDown, true);
                btnText.addEventListener('input', normalizeOnInput, true);
                btnText.addEventListener('blur', normalizeOnBlur, true);
                btnText.addEventListener('click', (e) => { e.stopPropagation(); }, true);
                // Prevent outer span from becoming editable due to browser quirks
                btn.setAttribute('contenteditable', 'false');

                // Guard against structural deletion when caret is adjacent to the button
                const guardStructure = (e) => {
                    const key = e.key;
                    if (key !== 'Backspace' && key !== 'Delete') return;
                    const sel = window.getSelection();
                    if (!sel || sel.rangeCount === 0) return;
                    const range = sel.getRangeAt(0);
                    const container = range.startContainer.nodeType === Node.ELEMENT_NODE ? range.startContainer : range.startContainer.parentNode;
                    // If caret is just after the button and Backspace is pressed, prevent removing the button
                    if (key === 'Backspace') {
                        // Determine node at caret - check previousSibling relative to a normalized container
                        const parent = container;
                        if (parent && parent.contains(btn)) {
                            // Walk up to the CTA section direct children context
                            let node = range.startContainer;
                            let offset = range.startOffset;
                            if (node.nodeType === Node.TEXT_NODE) {
                                // If inside text and not at start, allow default
                                if (offset > 0) return;
                                node = node.parentNode;
                                offset = Array.prototype.indexOf.call(node.parentNode ? node.parentNode.childNodes : [], node);
                            }
                            const siblings = node.parentNode ? node.parentNode.childNodes : [];
                            const prev = siblings[offset - 1] || node.previousSibling;
                            if (prev && (prev === btn || (prev.nodeType === 1 && prev.closest && prev.closest('.webinar-button') === btn))) {
                                e.preventDefault();
                                e.stopPropagation();
                                return;
                            }
                        }
                    }
                    // If caret is just before the button and Delete is pressed, prevent removing the button
                    if (key === 'Delete') {
                        let node = range.startContainer;
                        let offset = range.startOffset;
                        if (node.nodeType === Node.TEXT_NODE) {
                            if (offset < node.nodeValue.length) return; // not at end
                            node = node.parentNode;
                            offset = Array.prototype.indexOf.call(node.parentNode ? node.parentNode.childNodes : [], node) + 1;
                        }
                        const siblings = node.parentNode ? node.parentNode.childNodes : [];
                        const next = siblings[offset] || node.nextSibling;
                        if (next && (next === btn || (next.nodeType === 1 && next.closest && next.closest('.webinar-button') === btn))) {
                            e.preventDefault();
                            e.stopPropagation();
                            return;
                        }
                    }
                };
                ctaSection.addEventListener('keydown', guardStructure, true);
            }
        } catch (_) { /* no-op */ }
        
        this.insertElementAtCursor(ctaSection);
        this.saveState();
        try { this.showSectionToolbar(ctaSection); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section CTA insérée';
    }

    insertContactSection() {
        const contactSection = document.createElement('div');
        contactSection.className = 'newsletter-section contact-section';
        contactSection.style.cssText = 'margin: 30px 0; padding: 30px; background-color: #f8f9fa; border-radius: 8px; border: 1px solid #e9ecef;';
        
        contactSection.innerHTML = `
            <div style="text-align: center; margin-bottom: 25px;">
                <i class="fas fa-address-card" style="font-size: 36px; color: #6c757d; margin-bottom: 15px;"></i>
                <h3 contenteditable="true" style="color: #333; margin: 0; font-size: 24px;">Contactez-nous</h3>
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px;">
                <div style="text-align: center; padding: 20px;">
                    <i class="fas fa-envelope" style="font-size: 24px; color: #007bff; margin-bottom: 10px;"></i>
                    <h4 style="margin: 0 0 5px 0; color: #333;">Email</h4>
                    <p contenteditable="true" style="margin: 0; color: #666;">contact@exemple.com</p>
                </div>
                <div style="text-align: center; padding: 20px;">
                    <i class="fas fa-phone" style="font-size: 24px; color: #28a745; margin-bottom: 10px;"></i>
                    <h4 style="margin: 0 0 5px 0; color: #333;">Téléphone</h4>
                    <p contenteditable="true" style="margin: 0; color: #666;">+33 169852600</p>
                </div>
                <div style="text-align: center; padding: 20px;">
                    <i class="fas fa-map-marker-alt" style="font-size: 24px; color: #dc3545; margin-bottom: 10px;"></i>
                    <h4 style="margin: 0 0 5px 0; color: #333;">Adresse</h4>
                    <p contenteditable="true" style="margin: 0; color: #666;">14 mail du Commandant Cousteau 91300 Massy, France.</p>
                </div>
            </div>
        `;

        // Clear default contact texts and heading on first focus/click
        try {
            const clearOnFocus = (el, isDefault) => {
                const clear = () => {
                    try {
                        if (el.dataset.cleared) return;
                        if (!isDefault()) return;
                        el.innerHTML = '<br>';
                        el.dataset.cleared = '1';
                        const range = document.createRange();
                        range.selectNodeContents(el);
                        range.collapse(true);
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(range);
                        try { el.focus(); } catch (_) {}
                    } catch (_) {}
                };
                el.addEventListener('focus', clear);
                el.addEventListener('click', clear);
            };

            const email = contactSection.querySelector('div:nth-of-type(1) p[contenteditable]');
            if (email) clearOnFocus(email, () => (email.textContent || '').trim() === 'contact@exemple.com');
            const phone = contactSection.querySelector('div:nth-of-type(2) p[contenteditable]');
            // Phone: keep as real content, do not auto-clear
            const addr = contactSection.querySelector('div:nth-of-type(3) p[contenteditable]');
            // Address: keep as real content, do not auto-clear
        } catch (_) { /* no-op */ }
        
        this.insertElementAtCursor(contactSection);
        this.saveState();
        try { this.showSectionToolbar(contactSection); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section contact insérée';
    }

    insertTwoColumnSection() {
        const twoColumnSection = document.createElement('div');
        twoColumnSection.className = 'newsletter-section two-column-layout syc-item';
        twoColumnSection.style.cssText = 'display: flex; gap: 20px; margin: 30px 0; padding: 20px; background-color: #f8f9fa; border-radius: 8px;';

        twoColumnSection.innerHTML = `
            <div class="column" style="flex: 1; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div contenteditable="true">
                    <h3>Titre 1</h3>
                    <p>Contenu de l'élément 1. Ajoutez du texte, des images et d'autres éléments ici.</p>
                </div>
            </div>
            <div class="column" style="flex: 1; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div class="image-placeholder" contenteditable="false" style="border: 2px dashed #ccc; padding: 40px; text-align: center; cursor: pointer; border-radius: 8px;">
                    <i class="fas fa-image" style="font-size: 24px; color: #6c757d; margin-bottom: 10px;"></i>
                    <p style="color: #6c757d; margin: 0;">Cliquez pour ajouter une image</p>
                </div>
            </div>
        `;

        // Disappearing notice text for left column: clear on first focus/click if unchanged
        try {
            const leftColumn = twoColumnSection.querySelector('.column');
            const leftEditable = leftColumn && leftColumn.querySelector('[contenteditable="true"]');
            if (leftEditable) {
                const defaultTextMatcher = () => (leftEditable.textContent || '').trim().startsWith('Titre 1')
                    && (leftEditable.textContent || '').includes("Contenu de l'élément 1");
                const clearIfDefault = () => {
                    if (!leftEditable.dataset.cleared && defaultTextMatcher()) {
                        // Replace with a blank paragraph to retain height and caret
                        leftEditable.innerHTML = '<p><br></p>';
                        leftEditable.dataset.cleared = '1';
                        try {
                            // Place caret inside the blank paragraph
                            const p = leftEditable.querySelector('p');
                            if (p) {
                                const range = document.createRange();
                                range.selectNodeContents(p);
                                range.collapse(true);
                                const sel = window.getSelection();
                                sel.removeAllRanges();
                                sel.addRange(range);
                                leftEditable.focus();
                            }
                        } catch (_) {}
                    }
                };
                // Prevent bubbling so clicks/keys in left text do not trigger image/section handlers
                ['mousedown','click','keydown','keyup'].forEach(evt => {
                    leftEditable.addEventListener(evt, (e) => {
                        e.stopPropagation();
                    });
                });
                leftEditable.addEventListener('focus', clearIfDefault, true);
                leftEditable.addEventListener('click', clearIfDefault, true);
            }
        } catch (_) {}

        const imagePlaceholder = twoColumnSection.querySelector('.image-placeholder');
        imagePlaceholder.addEventListener('click', () => {
            const fileInput = document.createElement('input');
            fileInput.type = 'file';
            fileInput.accept = 'image/*';
            fileInput.style.display = 'none';

            fileInput.onchange = (e) => {
                const file = e.target.files[0];
                if (file) {
                    const reader = new FileReader();
                    reader.onload = (event) => {
                        const img = document.createElement('img');
                        img.src = event.target.result;
                        img.alt = 'Image';
                        img.setAttribute('contenteditable', 'false');
                        img.style.cssText = 'width:100%;height:auto;border-radius:8px;display:block;';
                        // Make the image clickable for editing later
                        img.addEventListener('click', (ev) => {
                            ev.stopPropagation();
                            try { this.selectImage(img); } catch (_) {}
                        });
                        // Remove any element(s) under the placeholder in the right column (text blocks, wrappers, etc.)
                        // This avoids leaving misleading default text below the image
                        try {
                            let sib = imagePlaceholder.nextElementSibling;
                            while (sib) {
                                const next = sib.nextElementSibling;
                                sib.remove();
                                sib = next;
                            }
                        } catch (_) {}
                        imagePlaceholder.replaceWith(img);
                        // Safety pass: remove any elements below the image in the same right column
                        try {
                            const column = img.closest('.column');
                            if (column) {
                                const children = Array.from(column.children);
                                let imgSeen = false;
                                for (const child of children) {
                                    if (child === img || (child.querySelector && child.querySelector('img') === img)) {
                                        imgSeen = true;
                                        continue;
                                    }
                                    if (imgSeen) {
                                        try { child.remove(); } catch (_) {}
                                    }
                                }
                            }
                        } catch (_) {}
                        // Do not auto-open image tools. They will open on user click, preserving expected UX.
                        this.saveState();
                        this.lastAction = 'Image ajoutée';
                    };
                    reader.readAsDataURL(file);
                }
            };
            fileInput.click();
        });

        this.insertElementAtCursor(twoColumnSection);
        this.saveState();
        try { this.showSectionToolbar(twoColumnSection); } catch (_) {}
        document.getElementById('sectionOptions').style.display = 'none';
        this.lastAction = 'Section deux colonnes insérée';
    }

    saveState() {
        // Create a temporary div to manipulate the content
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = document.getElementById('editableContent').innerHTML;
        
        // Remove all add buttons from the gallery sections before saving
        const gallerySections = tempDiv.querySelectorAll('.gallery-section');
        gallerySections.forEach(section => {
            const addButton = section.querySelector('.add-image-placeholder');
            if (addButton) {
                addButton.remove();
            }
            const addMoreBtn = section.querySelector('.add-more-btn');
            if (addMoreBtn) {
                addMoreBtn.remove();
            }
            // Also remove the file input
            const fileInput = section.querySelector('.gallery-upload');
            if (fileInput) {
                fileInput.remove();
            }
        });
        
        const content = tempDiv.innerHTML;
        
        // Remove future history if we're not at the end
        if (this.currentHistoryIndex < this.history.length - 1) {
            this.history = this.history.slice(0, this.currentHistoryIndex + 1);
        }
        
        // Add new state
        this.history.push(content);
        this.currentHistoryIndex++;
        
        // Limit history size
        if (this.history.length > this.maxHistorySize) {
            this.history.shift();
            this.currentHistoryIndex--;
        }
    }

    // Cleans editor-only markup and converts blob: media to persistent data URLs
    async sanitizeForExport(container) {
        // Remove editor-only wrappers/buttons
        container.querySelectorAll('.image-wrapper').forEach(wrapper => {
            const img = wrapper.querySelector('img');
            if (img) {
                wrapper.replaceWith(img);
            } else {
                wrapper.remove();
            }
        });
        container.querySelectorAll('.image-delete-btn, .crop-overlay').forEach(el => el.remove());

        // Remove contenteditable attributes
        container.querySelectorAll('[contenteditable]')
            .forEach(el => el.removeAttribute('contenteditable'));

        // Convert blob: URLs on img/video/source to data URLs, so saved file doesn't depend on runtime blobs
        const mediaNodes = Array.from(container.querySelectorAll('img, video, source'));
        for (const node of mediaNodes) {
            const srcAttr = node.tagName === 'SOURCE' ? 'src' : 'src';
            const src = node.getAttribute(srcAttr);
            if (!src || !src.startsWith('blob:')) continue;
            try {
                const fetched = await fetch(src);
                if (!fetched.ok) throw new Error('HTTP ' + fetched.status);
                const blob = await fetched.blob();
                const dataUrl = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result);
                    reader.onerror = reject;
                    reader.readAsDataURL(blob);
                });
                node.setAttribute(srcAttr, dataUrl);
            } catch (err) {
                console.warn('Failed to inline blob media, removing src to avoid 404:', src, err);
                node.removeAttribute(srcAttr);
            }
        }
    }

    undo() {
        if (this.currentHistoryIndex > 0) {
            this.currentHistoryIndex--;
            document.getElementById('editableContent').innerHTML = this.history[this.currentHistoryIndex];
            this.updateLastModified();
            this.lastAction = 'Édition annulée';
        }
    }

    redo() {
        if (this.currentHistoryIndex < this.history.length - 1) {
            this.currentHistoryIndex++;
            document.getElementById('editableContent').innerHTML = this.history[this.currentHistoryIndex];
            this.updateLastModified();
            this.lastAction = 'Édition rétablie';
        }
    }

    refresh() {
        location.reload();
    }

    clear() {
        if (confirm('Êtes-vous sûr de vouloir effacer tout le contenu ?')) {
            document.getElementById('editableContent').innerHTML = '';
            this.saveState();
            this.updateLastModified();
            this.lastAction = 'Contenu effacé';
        }
    }

    async confirmWithCancel(message) {
        return new Promise((resolve) => {
            try {
                let overlay = document.getElementById('kyo-confirm-overlay');
                if (!overlay) {
                    overlay = document.createElement('div');
                    overlay.id = 'kyo-confirm-overlay';
                    overlay.style.position = 'fixed';
                    overlay.style.inset = '0';
                    overlay.style.background = 'rgba(0,0,0,0.35)';
                    overlay.style.display = 'flex';
                    overlay.style.alignItems = 'center';
                    overlay.style.justifyContent = 'center';
                    overlay.style.zIndex = '99999';
                    const modal = document.createElement('div');
                    modal.id = 'kyo-confirm-modal';
                    modal.style.background = '#fff';
                    modal.style.borderRadius = '8px';
                    modal.style.boxShadow = '0 10px 25px rgba(0,0,0,0.2)';
                    modal.style.width = 'min(420px, 92vw)';
                    modal.style.maxWidth = '92vw';
                    modal.style.padding = '16px';
                    modal.style.fontFamily = 'inherit';
                    const msg = document.createElement('div');
                    msg.id = 'kyo-confirm-message';
                    msg.style.margin = '8px 0 16px 0';
                    msg.style.color = '#333';
                    msg.style.fontSize = '14px';
                    const actions = document.createElement('div');
                    actions.style.display = 'flex';
                    actions.style.justifyContent = 'flex-end';
                    actions.style.gap = '8px';
                    const cancelBtn = document.createElement('button');
                    cancelBtn.textContent = 'Annuler';
                    cancelBtn.style.padding = '8px 12px';
                    cancelBtn.style.border = '1px solid #ddd';
                    cancelBtn.style.background = '#fff';
                    cancelBtn.style.borderRadius = '6px';
                    const okBtn = document.createElement('button');
                    okBtn.textContent = 'Supprimer';
                    okBtn.style.padding = '8px 12px';
                    okBtn.style.border = 'none';
                    okBtn.style.background = '#dc3545';
                    okBtn.style.color = '#fff';
                    okBtn.style.borderRadius = '6px';
                    actions.appendChild(cancelBtn);
                    actions.appendChild(okBtn);
                    modal.appendChild(msg);
                    modal.appendChild(actions);
                    overlay.appendChild(modal);
                    document.body.appendChild(overlay);
                }
                const msgEl = overlay.querySelector('#kyo-confirm-message');
                if (msgEl) msgEl.textContent = message || '';
                overlay.style.display = 'flex';
                const cleanup = () => { overlay.style.display = 'none'; };
                const onCancel = () => { cleanup(); resolve(false); };
                const onOk = () => { cleanup(); resolve(true); };
                const cancelBtn = overlay.querySelector('button:nth-of-type(1)');
                const okBtn = overlay.querySelector('button:nth-of-type(2)');
                cancelBtn.onclick = onCancel;
                okBtn.onclick = onOk;
            } catch (_) { resolve(false); }
        });
    }

    async save() {
        try {
            console.log('Save method called');
            const editableContent = document.getElementById('editableContent');
            if (!editableContent) {
                throw new Error('Éditeur de contenu introuvable');
            }
            
            // Get the content from the editor
            let content = editableContent.innerHTML;
            let pageTitle = '';
            try {
                // Prefer any element with inline style 52px (tolerate missing space)
                const px52El = editableContent.querySelector('*[style*="font-size: 52px"], *[style*="font-size:52px"]');
                if (px52El) {
                    const t = (px52El.innerText || px52El.textContent || '').trim();
                    if (t) pageTitle = t;
                }
            } catch (_) { /* no-op */ }
            if (!pageTitle) {
                pageTitle = document.title.trim() ||
                            (document.querySelector('h1, h2, h3')?.textContent.trim() || 'newsletter_sans_titre');
            }
            const fileName = pageTitle.replace(/[\\/:*?"<>|]/g, '_'); // Remove invalid filename characters
            
            // Create a temporary div to clean up the content
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = content;
            // Pre-save validation: check for 52px and 22px font sizes and image filenames starting with 'article'
            try {
                const allEls = tempDiv.querySelectorAll('*');
                let has52 = false;
                let has22 = false;
                for (const el of allEls) {
                    const styleAttr = (el.getAttribute && el.getAttribute('style')) || '';
                    if (!has52 && /font-size\s*:\s*52px/i.test(styleAttr)) has52 = true;
                    if (!has22 && /font-size\s*:\s*22px/i.test(styleAttr)) has22 = true;
                    if (el.style) {
                        const fs = (el.style.fontSize || '').toLowerCase();
                        if (!has52 && fs === '52px') has52 = true;
                        if (!has22 && fs === '22px') has22 = true;
                    }
                    if (has52 && has22) break;
                }
                let conformingImageCount = 0;
                tempDiv.querySelectorAll('img').forEach(img => {
                    const src = (img.getAttribute('src') || '').trim();
                    if (!src) return;
                    try {
                        // Prefer filename from URL when not data/blob
                        if (!/^(data:|blob:)/i.test(src)) {
                            const clean = src.split('?')[0].split('#')[0];
                            const parts = clean.split(/[\/\\]/);
                            const name = (parts[parts.length - 1] || '').toLowerCase();
                            if (/^article\d+/i.test(name)) { conformingImageCount++; return; }
                        }
                        // For data/blob or when URL doesn't encode the name, try attributes
                        const candidates = [
                            img.getAttribute('data-filename'),
                            img.getAttribute('data-original-name'),
                            img.getAttribute('data-name'),
                            img.getAttribute('data-original-src'),
                            img.getAttribute('alt'),
                            img.getAttribute('title')
                        ];
                        for (const c of candidates) {
                            const v = (c || '').trim();
                            if (!v) continue;
                            const parts2 = v.split(/[\/\\]/);
                            const base = (parts2[parts2.length - 1] || '').toLowerCase();
                            if (/^article\d+/i.test(base)) { conformingImageCount++; return; }
                        }
                    } catch (_) { /* ignore */ }
                });
                // Validate video sources as well
                const badVideos = [];
                let skippedFirstVideoElement = false; // exempt the entire first <video> element
                tempDiv.querySelectorAll('video').forEach(vid => {
                    if (!skippedFirstVideoElement) { skippedFirstVideoElement = true; return; }
                    const vsrc = (vid.getAttribute('src') || '').trim();
                    if (vsrc && !/^(data:|blob:)/i.test(vsrc)) {
                        try {
                            const clean = vsrc.split('?')[0].split('#')[0];
                            const parts = clean.split(/[\/\\]/);
                            const name = (parts[parts.length - 1] || '').toLowerCase();
                            const ok = /^article\d+/i.test(name);
                            if (!ok) badVideos.push(name || vsrc);
                        } catch (_) { /* ignore */ }
                    }
                    // Also check nested <source> tags
                    vid.querySelectorAll('source[src]').forEach(srcEl => {
                        const s = (srcEl.getAttribute('src') || '').trim();
                        if (!s || /^(data:|blob:)/i.test(s)) return;
                        try {
                            const clean = s.split('?')[0].split('#')[0];
                            const parts = clean.split(/[\/\\]/);
                            const name = (parts[parts.length - 1] || '').toLowerCase();
                            const ok = /^article\d+/i.test(name);
                            if (!ok) badVideos.push(name || s);
                        } catch (_) { /* ignore */ }
                    });
                });
                const problems = [];
                // New rule set:
                // 1) must have 52px title AND 22px subtitle
                // 2) must have EITHER at least one image named articleN OR at least one video (local <video> or known video iframe)
                const videoTagCount = tempDiv.querySelectorAll('video').length;
                const iframeVideoCount = Array.from(tempDiv.querySelectorAll('iframe[src]'))
                    .filter(ifr => /(?:youtube\.com|youtu\.be|vimeo\.com|dailymotion\.com|player\.wistia\.net|loom\.com)/i
                        .test((ifr.getAttribute('src') || ''))).length;
                const hasAnyVideo = (videoTagCount + iframeVideoCount) > 0;
                if (!has52) problems.push("titre 52px manquant");
                if (!has22) problems.push("sous-titre 22px manquant");
                if (!(conformingImageCount >= 1 || hasAnyVideo)) {
                    problems.push("au moins une image 'articleN' ou une vidéo requise");
                }
                if (problems.length) {
                    alert('Validation avant sauvegarde échouée:\n' + problems.join('\n'));
                    return; // Abort save
                }
            } catch (_) { /* if validation crashes, do not block save */ }
            console.log('Content extracted for saving');

            let imageSummaryText = '';
            
            // Auto-generate poster images for local videos (blob or file) so final HTML has a thumbnail
            // We capture from the live video elements in the editor, then set 'poster' on the temp copy
            try {
                const editableEl = editableContent;
                const liveVideos = Array.from(editableEl.querySelectorAll('video'));
                const tempVideos = Array.from(tempDiv.querySelectorAll('video'));

                async function captureFrameFromLiveVideo(videoEl, seekTime = 0.1, targetWidth = 800) {
                    return new Promise((resolve) => {
                        try {
                            const onError = () => resolve('');
                            const cleanup = () => {
                                videoEl.removeEventListener('error', onError);
                                videoEl.removeEventListener('loadeddata', onLoaded);
                                videoEl.removeEventListener('seeked', onSeeked);
                            };
                            const draw = () => {
                                try {
                                    const vw = videoEl.videoWidth || 0;
                                    const vh = videoEl.videoHeight || 0;
                                    if (!vw || !vh) { resolve(''); return; }
                                    const ratio = vw / vh;
                                    const w = targetWidth;
                                    const h = Math.max(1, Math.round(w / (ratio || 1)));
                                    const canvas = document.createElement('canvas');
                                    canvas.width = w; canvas.height = h;
                                    const ctx = canvas.getContext('2d');
                                    ctx.drawImage(videoEl, 0, 0, w, h);
                                    const dataUrl = canvas.toDataURL('image/jpeg', 0.86);
                                    resolve(dataUrl || '');
                                } catch { resolve(''); }
                            };
                            const onSeeked = () => { cleanup(); draw(); };
                            const onLoaded = () => {
                                try {
                                    // Seek slightly into the video to avoid black frame
                                    if (!isNaN(seekTime) && videoEl.seekable && videoEl.seekable.length > 0) {
                                        try { videoEl.currentTime = Math.min(seekTime, videoEl.seekable.end(0) || seekTime); } catch {}
                                        videoEl.addEventListener('seeked', onSeeked, { once: true });
                                    } else {
                                        cleanup(); draw();
                                    }
                                } catch { cleanup(); resolve(''); }
                            };
                            videoEl.addEventListener('error', onError, { once: true });
                            if (videoEl.readyState >= 2) { onLoaded(); }
                            else { videoEl.addEventListener('loadeddata', onLoaded, { once: true }); }
                            setTimeout(() => { cleanup(); resolve(''); }, 2000);
                        } catch { resolve(''); }
                    });
                }

                for (let i = 0; i < liveVideos.length; i++) {
                    const liveV = liveVideos[i];
                    const tempV = tempVideos[i];
                    if (!tempV) continue;
                    // Skip if poster already present
                    if (tempV.hasAttribute('poster') && (tempV.getAttribute('poster') || '').trim().length > 0) continue;
                    // Try to capture a frame from the live element
                    try {
                        const posterUrl = await captureFrameFromLiveVideo(liveV);
                        if (posterUrl) {
                            tempV.setAttribute('poster', posterUrl);
                        }
                    } catch (_) { /* ignore and continue */ }
                }
            } catch (e) {
                console.debug('Video poster auto-generation skipped:', e);
            }
            
            // Remove all editor-specific elements before saving
            const elementsToRemove = [
                '.add-image-placeholder',
                '.add-more-btn',
                '.image-toolbar',
                '.resize-handle',
                '.rotation-handle'
            ];
            
            elementsToRemove.forEach(selector => {
                tempDiv.querySelectorAll(selector).forEach(el => el.remove());
            });

            // Extra cleanup for gallery editor controls/placeholders that might remain in content
            try {
                // Remove common gallery control wrappers or inputs
                const extraSelectors = [
                    '.gallery-upload',
                    '.gallery-controls',
                    '.gallery-editor-only',
                    '.gallery-actions',
                    'input[type="file"]'
                ];
                extraSelectors.forEach(sel => tempDiv.querySelectorAll(sel).forEach(el => el.remove()));

                // Remove buttons by visible label to be robust even without specific classes
                const buttonTextsToRemove = [
                    'Ajouter une image',
                    'Modifier la galerie',
                    'Ajouter plus d\'images'
                ];
                tempDiv.querySelectorAll('button, a').forEach(el => {
                    const txt = (el.innerText || el.textContent || '').trim().toLowerCase();
                    if (buttonTextsToRemove.some(t => txt.includes(t.toLowerCase()))) {
                        el.remove();
                    }
                });

                // Remove obvious placeholder tiles inside galleries
                tempDiv.querySelectorAll('[class*="placeholder"]').forEach(el => el.remove());
                tempDiv.querySelectorAll('[data-editor-only], [data-placeholder], [data-editable-control]').forEach(el => el.remove());
            } catch (_) { /* best-effort cleanup */ }
            
            // Remove ALL contenteditable attributes to make content completely non-editable
            tempDiv.querySelectorAll('[contenteditable]').forEach(el => {
                el.removeAttribute('contenteditable');
            });
            
            // Remove any remaining editor-specific attributes and classes
            tempDiv.querySelectorAll('[data-placeholder]').forEach(el => {
                el.removeAttribute('data-placeholder');
            });
            
            try {
                const MAX_IMG_WIDTH = 800;
                const JPEG_QUALITY = 0.85;
                function __blobFromUrl(u) {
                    return fetch(u, { cache: 'no-store' }).then(r => r.ok ? r.blob() : Promise.reject(new Error('fetch failed')));
                }
                function __blobToDataURL(blob) {
                    return new Promise((resolve, reject) => {
                        try {
                            const fr = new FileReader();
                            fr.onload = () => resolve(fr.result || '');
                            fr.onerror = () => reject(new Error('read'));
                            fr.readAsDataURL(blob);
                        } catch (e) { resolve(''); }
                    });
                }
                function __createImageFromBlob(blob) {
                    return new Promise((resolve, reject) => {
                        try {
                            const url = URL.createObjectURL(blob);
                            const img = new Image();
                            img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
                            img.onerror = () => { URL.revokeObjectURL(url); resolve(null); };
                            img.src = url;
                        } catch (e) { resolve(null); }
                    });
                }
                async function __resizeImgBlob(blob, origType) {
                    try {
                        const img = await __createImageFromBlob(blob);
                        if (!img) return blob;
                        const w = img.naturalWidth || img.width || 0;
                        const h = img.naturalHeight || img.height || 0;
                        if (!w || !h) return blob;
                        const targetW = Math.min(MAX_IMG_WIDTH, w);
                        if (targetW === w) return blob;
                        const targetH = Math.round(h * (targetW / w));
                        const canvas = document.createElement('canvas');
                        canvas.width = targetW; canvas.height = targetH;
                        const ctx = canvas.getContext('2d');
                        ctx.drawImage(img, 0, 0, targetW, targetH);
                        const usePng = !!(origType && /png/i.test(origType));
                        const mime = usePng ? 'image/png' : 'image/jpeg';
                        const quality = usePng ? undefined : JPEG_QUALITY;
                        return new Promise((resolve, reject) => {
                            canvas.toBlob(b => resolve(b || blob), mime, quality);
                        });
                    } catch (_) { return blob; }
                }
                const imgs = Array.from(tempDiv.querySelectorAll('img'));
                let __resized = 0, __unchanged = 0, __skippedCors = 0;
                for (const im of imgs) {
                    try {
                        const src = (im.getAttribute('src') || '').trim();
                        if (!src) continue;
                        if (/^data:/i.test(src)) { __unchanged++; continue; }
                        let blob;
                        try { blob = await __blobFromUrl(src); } catch { __skippedCors++; continue; }
                        if (!blob) continue;
                        const resized = await __resizeImgBlob(blob, blob.type || '');
                        if (resized === blob) { __unchanged++; }
                        else { __resized++; }
                        const dataUrl = await __blobToDataURL(resized);
                        if (dataUrl) {
                            im.setAttribute('src', dataUrl);
                            im.removeAttribute('srcset');
                            im.removeAttribute('sizes');
                        }
                    } catch (_) { /* ignore */ }
                }
                imageSummaryText = `${__resized} image(s) redimensionnée(s), ${__unchanged} inchangée(s), ${__skippedCors} ignorée(s) (CORS)`;
            } catch (_) { /* ignore */ }
            
            // Get the cleaned content and strip editor-only UI (e.g. video toolbar handles)
            const cleanedContent = (typeof sanitizeFullHtmlForHistory === 'function')
                ? sanitizeFullHtmlForHistory(tempDiv.innerHTML)
                : tempDiv.innerHTML;
            
            // Indicate manual save as the last action
            this.lastAction = 'Sauvegarde manuelle';

            // Save to history first with the cleaned content
            const historyItem = this.saveToHistory(fileName, cleanedContent);
            if (historyItem) {
                console.log('Saved to history:', historyItem);
            } else {
                console.warn('Failed to save to history');
            }
            
            // Create a complete HTML document with the cleaned content
            const fullHTML = `<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="robots" content="noindex,nofollow,noarchive,noimageindex">
    <title>${fileName}</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <style>
        body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
            color: #333;
        }
        .newsletter-container {
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background: #fff;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
        }
        img {
            max-width: 100%;
            height: auto;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
        }
        table, th, td {
            border: 1px solid #ddd;
        }
        th, td {
            padding: 10px;
            text-align: left;
        }
        .newsletter-section {
            margin: 20px 0;
            padding: 15px;
            border: 1px solid #eee;
            border-radius: 5px;
        }
        .gallery-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
            gap: 10px;
            margin: 15px 0;
        }
        .gallery-item {
            position: relative;
            overflow: hidden;
            border-radius: 5px;
        }
        .gallery-item img {
            width: 100%;
            height: 150px;
            object-fit: cover;
            transition: transform 0.3s ease;
        }
        .gallery-item:hover img {
            transform: scale(1.05);
        }
        .two-column-layout {
            display: flex;
            gap: 20px;
        }
        .column {
            flex: 1;
        }
        @media (max-width: 768px) {
            .two-column-layout {
                flex-direction: column;
            }
        }
        /* Ensure images maintain their aspect ratio */
        img {
            max-width: 100%;
            height: auto;
        }
        /* Make sure tables are responsive */
        table {
            width: 100% !important;
            max-width: 100%;
            table-layout: fixed;
        }
    </style>
</head>
<body>
    <div class="newsletter-container">
        ${cleanedContent}
    </div>
</body>
</html>`;

            // Try to use the File System Access API (modern browsers)
            if (window.showSaveFilePicker) {
                try {
                    let chosenHandle = null;
                    while (true) {
                        const baseSuggested = fileName.replace(/[^\w\-.]/g, '_');
                        const validBase = /^article\d+/i.test(baseSuggested.replace(/\.(html?)$/i, ''));
                        const proposed = validBase ? `${baseSuggested.replace(/\.(html?)$/i, '')}.html` : 'article1_edito.html';
                        const handle = await window.showSaveFilePicker({
                            suggestedName: proposed,
                            types: [{
                                description: 'Fichier HTML',
                                accept: { 'text/html': ['.html'] }
                            }]
                        });
                        const chosenName = (handle && handle.name) ? handle.name : proposed;
                        const baseNoExt = chosenName.replace(/\.(html?)$/i, '');
                        if (!/^article\d+/i.test(baseNoExt)) {
                            alert("Nom de fichier invalide: il doit commencer par 'articleN' (ex: article1_edito.html)\n\nLa sauvegarde est annulée.");
                            // Abort save entirely (Option A)
                            return;
                        }
                        chosenHandle = handle;
                        break;
                    }
                    // Final safety check before writing
                    try {
                        const finalName = (chosenHandle && chosenHandle.name ? chosenHandle.name : '').replace(/\.(html?)$/i, '');
                        if (!/^article\d+/i.test(finalName)) {
                            // Do not write if name is still invalid for any reason
                            return;
                        }
                    } catch(_) { return; }
                    const writable = await chosenHandle.createWritable();
                    await writable.write(fullHTML);
                    await writable.close();
                    
                    alert('Newsletter sauvegardée avec succès !' + (imageSummaryText ? '\n' + imageSummaryText : ''));
                    return;
                } catch (err) {
                    if (err && err.name === 'AbortError') {
                        // User cancelled the save dialog
                        return;
                    }
                    console.error('Error saving file:', err);
                    // Fall through to the download method
                }
            }
            
            // Fallback for older browsers
            try {
                // Ensure filename chosen/suggested is valid; prompt user if not
                let baseName = fileName.replace(/[^\w\-.]/g, '_');
                baseName = baseName.replace(/\.(html?)$/i, '');
                if (!/^article\d+/i.test(baseName)) {
                    const entered = prompt("Entrez un nom de fichier commençant par 'articleN' (ex: article1_edito)", 'article1_edito');
                    if (!entered) {
                        alert('Sauvegarde annulée');
                        return;
                    }
                    baseName = (entered || '').trim().replace(/[^\w\-.]/g, '_').replace(/\.(html?)$/i, '');
                    if (!/^article\d+/i.test(baseName)) {
                        alert("Nom de fichier invalide: il doit commencer par 'articleN'");
                        return;
                    }
                }
                const finalDownloadName = `${baseName}.html`;

                const blob = new Blob([fullHTML], { type: 'text/html;charset=utf-8' });
                const url = URL.createObjectURL(blob);
                
                const a = document.createElement('a');
                a.href = url;
                a.download = finalDownloadName;
                
                // Add to document, trigger click, then remove
                document.body.appendChild(a);
                a.click();
                
                // Clean up
                setTimeout(() => {
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                }, 100);
                
                console.log('Newsletter saved successfully');
                try { alert('Newsletter sauvegardée avec succès !' + (imageSummaryText ? '\n' + imageSummaryText : '')); } catch {}
            } catch (error) {
                console.error('Error during save:', error);
                alert('Erreur lors de la sauvegarde: ' + (error.message || 'Erreur inconnue'));
            }
        } catch (error) {
            console.error('Error in save process:', error);
        }
    }

    // Persist a snapshot to localStorage 'newsletterHistory'
    saveToHistory(fileName, content) {
        try {
            // Generate unique ID
            const id = Date.now().toString();

            // Build a possibly-sanitized copy for persistence to avoid quota issues
            const originalContent = content || '';
            let contentForHistory = originalContent;
            try {
                const isHeavy = originalContent.length > 300 * 1024 || /\ssrc=\"(?:data:|blob:)/i.test(originalContent);
                if (isHeavy) {
                    let slim = originalContent
                        .replace(/\s+src=\"data:[^\"]+\"/gi, '')
                        .replace(/\s+src=\"blob:[^\"]+\"/gi, '')
                        .replace(/<source([^>]*)src=\"[^\"]+\"([^>]*)>/gi, '<source$1$2>')
                        .replace(/<video[\s\S]*?<\/video>/gi, '<div class="video-placeholder" data-omitted="true"></div>');
                    const MAX_BYTES = 300 * 1024; // 300KB per entry
                    if (slim.length > MAX_BYTES) slim = slim.slice(0, MAX_BYTES);
                    contentForHistory = slim;
                }
            } catch (_) { /* keep original if sanitize fails */ }

            // Persist the full, unsanitized content separately in IndexedDB (non-blocking)
            try {
                if (typeof saveFullContentToIDB === 'function') {
                    saveFullContentToIDB(id, originalContent);
                }
            } catch (_) { /* ignore */ }

            // Compute preview from the original HTML
            const previewText = (content || '').replace(/<[^>]*>?/gm, '').substring(0, 150) + '...';

            // Derive a meaningful title from content using inline font-size:52px, then H1/H2/H3, then text
            let computedName = '';
            try {
                const temp = document.createElement('div');
                temp.innerHTML = content || '';
                const all = temp.querySelectorAll('*');
                let found = null;
                for (const el of all) {
                    const styleAttr = (el.getAttribute && el.getAttribute('style')) || '';
                    if (styleAttr && /font-size\s*:\s*52px/i.test(styleAttr)) { found = el; break; }
                    if (el.style && (el.style.fontSize || '').toLowerCase() === '52px') { found = el; break; }
                }
                if (found && found.textContent) computedName = found.textContent.trim();
                if (!computedName) {
                    const heading = temp.querySelector('h1, h2, h3');
                    if (heading) computedName = (heading.textContent || '').trim();
                }
                if (!computedName) computedName = (temp.textContent || '').trim().substring(0, 80);
            } catch (_) {}
            if (!computedName) computedName = fileName || 'Sans nom';

            // Prepare item
            const historyItem = {
                id,
                name: computedName,
                content: contentForHistory,
                date: new Date().toLocaleString('fr-FR'),
                preview: previewText,
                timestamp: Date.now(),
                lastAction: this.lastAction || 'Action inconnue'
            };

            // Read existing, push to front, clamp to 200 and save
            let history = [];
            try {
                const raw = localStorage.getItem('newsletterHistory');
                history = raw ? JSON.parse(raw) : [];
                if (!Array.isArray(history)) history = [];
            } catch (_) { history = []; }

            history.unshift(historyItem);
            if (history.length > 200) history = history.slice(0, 200);
            // Try to persist; if quota exceeded, trim and/or slim down content
            try {
                localStorage.setItem('newsletterHistory', JSON.stringify(history));
            } catch (e) {
                if (e && (e.name === 'QuotaExceededError' || e.code === 22)) {
                    // Remove oldest items until it fits
                    let saved = false;
                    while (history.length > 0 && !saved) {
                        history.pop();
                        try {
                            localStorage.setItem('newsletterHistory', JSON.stringify(history));
                            saved = true;
                        } catch (_) { /* keep trimming */ }
                    }
                    if (!saved) {
                        // Last resort: create a sanitized "slim" copy only for storage purposes
                        let slim = { ...historyItem };
                        try {
                            // Start from already sanitized-or-original contentForHistory, guard size
                            let slimContent = contentForHistory;
                            slimContent = slimContent.slice(0, 200000); // ~200KB
                            slim.content = slimContent;
                        } catch(_) { slim.content = ''; }
                        try {
                            localStorage.setItem('newsletterHistory', JSON.stringify([slim]));
                        } catch(_) {
                            // Give up but avoid throwing; history won't be updated this round
                            console.warn('History not saved due to storage quota, even after slimming');
                            try {
                                sessionStorage.setItem('newsletterHistoryFallback', JSON.stringify([slim]));
                            } catch (_) {}
                            try { window.__historyBuffer = [slim]; } catch (_) {}
                        }
                    }
                } else {
                    throw e;
                }
            }

            return historyItem;
        } catch (error) {
            console.error('Error saving to history:', error);
            return null;
        }
    }

    showHistory() {
        console.log('showHistory method called');
        
        try {
            // Get history from localStorage
            let history = [];
            try {
                const historyData = localStorage.getItem('newsletterHistory');
                history = historyData ? JSON.parse(historyData) : [];
                console.log('Loaded history from localStorage:', history);
            } catch (e) {
                console.error('Error parsing history from localStorage:', e);
                history = [];
            }
            
            // Clean up invalid entries and sort by timestamp (newest first)
            history = history
                .filter(item => {
                    const isValid = item && 
                                  typeof item === 'object' && 
                                  item.name && 
                                  item.content && 
                                  typeof item.name === 'string' && 
                                  typeof item.content === 'string';
                    if (!isValid) {
                        console.warn('Removing invalid history item:', item);
                    }
                    return isValid;
                })
                .sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
            
            // Update localStorage with cleaned history
            localStorage.setItem('newsletterHistory', JSON.stringify(history));
            
            console.log('Processed history data:', history);
            
            const historyList = document.getElementById('historyList');
            console.log('History list element:', historyList);
            
            if (!historyList) {
                throw new Error('historyList element not found');
            }
            
            if (history.length === 0) {
                historyList.innerHTML = `
                    <div style="text-align: center; padding: 20px; color: #666;">
                        <i class="fas fa-inbox" style="font-size: 48px; opacity: 0.5; margin-bottom: 10px;"></i>
                        <p>Aucun historique disponible</p>
                    </div>
                `;
            } else {
                // Create a document fragment for better performance
                const fragment = document.createDocumentFragment();
                
                // Add a title
                const title = document.createElement('h3');
                title.textContent = 'Historique des sauvegardes';
                title.style.margin = '0 0 15px 0';
                title.style.color = '#333';
                title.style.borderBottom = '1px solid #eee';
                title.style.paddingBottom = '10px';
                fragment.appendChild(title);
                
                // Add each history item
                history.forEach(item => {
                    if (!item || !item.name || !item.content) {
                        console.warn('Skipping invalid history item:', item);
                        return;
                    }
                    
                    const itemElement = document.createElement('div');
                    itemElement.className = 'history-item';
                    itemElement.style.cssText = `
                        border: 1px solid #e0e0e0;
                        border-radius: 8px;
                        padding: 12px 15px;
                        margin-bottom: 10px;
                        cursor: pointer;
                        transition: all 0.2s ease;
                        background: white;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
                    `;
                    
                    // Add hover effect
                    itemElement.onmouseover = () => {
                        itemElement.style.borderColor = '#0a9bcd';
                        itemElement.style.boxShadow = '0 4px 8px rgba(0,0,0,0.1)';
                    };
                    itemElement.onmouseout = () => {
                        itemElement.style.borderColor = '#e0e0e0';
                        itemElement.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
                    };
                    
                    // Format the date
                    let formattedDate = 'Date inconnue';
                    try {
                        const date = item.date ? new Date(item.date) : new Date(item.timestamp);
                        if (!isNaN(date.getTime())) {
                            formattedDate = date.toLocaleString('fr-FR', {
                                day: '2-digit',
                                month: '2-digit',
                                year: 'numeric',
                                hour: '2-digit',
                                minute: '2-digit'
                            });
                        }
                    } catch (e) {
                        console.warn('Error formatting date:', e);
                    }
                    
                    // Create the item content
                    itemElement.innerHTML = `
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <h4 style="margin: 0 0 5px 0; color: #0a9bcd; font-size: 16px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 70%;">
                                ${item.name || 'Sans nom'}
                            </h4>
                            <span style="color: #888; font-size: 12px;">${formattedDate}</span>
                        </div>
                        <p style="margin: 5px 0 0 0; color: #666; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
                            ${item.preview || 'Aperçu non disponible'}
                        </p>
                    `;
                    
                    // Add click handler
                    itemElement.onclick = () => {
                        this.loadFromHistory(item.name, item.content);
                    };
                    
                    fragment.appendChild(itemElement);
                });
                
                // Clear and append the fragment
                historyList.innerHTML = '';
                historyList.appendChild(fragment);
            }
            
            // Show the modal
            const modal = document.getElementById('historyModal');
            if (!modal) {
                throw new Error('historyModal element not found');
            }
            
            // Add close button if not exists
            if (!document.getElementById('closeHistoryModal')) {
                const closeBtn = document.createElement('button');
                closeBtn.id = 'closeHistoryModal';
                closeBtn.innerHTML = '&times;';
                closeBtn.style.cssText = `
                    position: absolute;
                    top: 10px;
                    right: 15px;
                    background: none;
                    border: none;
                    font-size: 24px;
                    cursor: pointer;
                    color: #666;
                `;
                closeBtn.onclick = () => {
                    // Fade out the modal content first
                    const modalContent = modal.querySelector('.history-modal');
                    if (modalContent) {
                        modalContent.style.opacity = '0';
                        modalContent.style.transform = 'translateY(-20px)';
                    }
                    
                    // Then fade out the overlay
                    modal.style.opacity = '0';
                    
                    // Hide everything after animation completes
                    setTimeout(() => {
                        modal.style.display = 'none';
                        modal.style.pointerEvents = 'none';
                    }, 300);
                };
                
                // Add close button to modal header
                const modalHeader = modal.querySelector('.modal-header');
                if (modalHeader) {
                    modalHeader.appendChild(closeBtn);
                }
            }
            
            // Show the modal with animation
            modal.style.display = 'flex';
            modal.style.pointerEvents = 'auto';
            
            // Force reflow
            void modal.offsetWidth;
            
            // Fade in the overlay
            modal.style.opacity = '1';
            
            // Get the modal content
            const modalContent = modal.querySelector('.history-modal');
            if (modalContent) {
                // Reset and animate the modal content
                modalContent.style.opacity = '0';
                modalContent.style.transform = 'translateY(-20px)';
                
                // Animate in
                setTimeout(() => {
                    modalContent.style.opacity = '1';
                    modalContent.style.transform = 'translateY(0)';
                }, 10);
            }
            
            console.log('History modal displayed');
            
        } catch (error) {
            console.error('Error in showHistory:', error);
            alert('Erreur lors de l\'affichage de l\'historique: ' + (error.message || 'Erreur inconnue'));
        }
    }

    loadFromHistory(name, content) {
        try {
            if (confirm(`Charger "${name}" ? Le contenu actuel sera remplacé.`)) {
                // Unescape the content if it's a string with escape sequences
                let unescapedContent = content;
                if (typeof content === 'string') {
                    unescapedContent = content
                        .replace(/\\'/g, "'")
                        .replace(/\\"/g, '"')
                        .replace(/\\n/g, '\n')
                        .replace(/\\r/g, '\r');
                }
                
                // Update the editor content
                document.getElementById('editableContent').innerHTML = unescapedContent;
                // Apply standard video sizing to history-loaded content
                try { this.normalizeVideoStyles(); } catch (_) {}
                
                // Update the title if available in the history item
                if (name && name !== 'newsletter_sans_titre') {
                    const titleInput = document.getElementById('newsletterTitle');
                    if (titleInput) {
                        titleInput.value = name;
                    }
                }
                
                // Save the current state
                this.saveState();
                this.updateLastModified();
                this.autoSaveToLocalStorage();
                
                // Close the history modal
                const modal = document.getElementById('historyModal');
                if (modal) {
                    modal.style.display = 'none';
                }
                
                console.log('Content loaded and restored from history:', name);
                alert(`Contenu "${name}" chargé avec succès !`);
            }
        } catch (error) {
            console.error('Error loading from history:', error);
            alert('Erreur lors du chargement de l\'historique: ' + error.message);
        }
    }

    updateLastModified() {
        document.getElementById('lastUpdate').textContent = new Date().toLocaleString('fr-FR');
    }

    // Restore function to recover from localStorage if page is refreshed
    restoreFromLocalStorage() {
        try {
            // Only restore if versions match
            const lsVersion = localStorage.getItem('currentNewsletterVersion');
            const ssVersion = sessionStorage.getItem('currentNewsletterVersion');
            let savedContent = null;
            if (lsVersion === this.storageVersion) {
                savedContent = localStorage.getItem('currentNewsletterContent');
            }
            if (!savedContent && ssVersion === this.storageVersion) {
                savedContent = sessionStorage.getItem('currentNewsletterContentSession');
                if (savedContent) this._autosaveUsingSession = true;
            }
            if (!savedContent) {
                // Clear incompatible old autosave to prevent future overwrites
                localStorage.removeItem('currentNewsletterContent');
                sessionStorage.removeItem('currentNewsletterContentSession');
                localStorage.removeItem('currentNewsletterVersion');
                sessionStorage.removeItem('currentNewsletterVersion');
            }
            const savedFileName = localStorage.getItem('currentNewsletterFileName');
            
            if (savedContent) {
                document.getElementById('editableContent').innerHTML = savedContent;
                console.log('Content restored from local/session storage');
                // Apply standard video sizing to restored content
                try { this.normalizeVideoStyles(); } catch (_) {}
            }
            
            // File name functionality removed
        } catch (error) {
            console.error('Error restoring from storage:', error);
        }
    }

    // Auto-save current content to localStorage with sessionStorage fallback
    autoSaveToLocalStorage() {
        const content = document.getElementById('editableContent').innerHTML;
        if (this._autosaveUsingSession) {
            try {
                sessionStorage.setItem('currentNewsletterContentSession', content);
                sessionStorage.setItem('currentNewsletterVersion', this.storageVersion);
            } catch (_) {}
            return;
        }
        try {
            localStorage.setItem('currentNewsletterContent', content);
            localStorage.setItem('currentNewsletterVersion', this.storageVersion);
        } catch (error) {
            const message = (error && (error.name || '')).toString();
            if (message.includes('QuotaExceededError') || message.includes('QUOTA') || message.includes('NS_ERROR_DOM_QUOTA_REACHED')) {
                try {
                    sessionStorage.setItem('currentNewsletterContentSession', content);
                    sessionStorage.setItem('currentNewsletterVersion', this.storageVersion);
                    this._autosaveUsingSession = true;
                    console.info('Autosave switched to sessionStorage due to quota. Manual Save unaffected.');
                } catch (_) {
                    console.info('Autosave paused: both localStorage and sessionStorage quotas exceeded. Manual Save still works.');
                }
            } else {
                console.error('Error auto-saving to localStorage:', error);
            }
        }
    }

    // Utility: compress an image File to a DataURL
    compressImageFile(file, { maxWidth = 1600, maxHeight = 1200, quality = 0.9 } = {}) {
        return new Promise((resolve, reject) => {
            try {
                const img = new Image();
                const reader = new FileReader();
                reader.onload = (e) => {
                    img.onload = () => {
                        let width = img.width;
                        let height = img.height;
                        const ratio = Math.min(maxWidth / width, maxHeight / height, 1);
                        const canvas = document.createElement('canvas');
                        canvas.width = Math.round(width * ratio);
                        canvas.height = Math.round(height * ratio);
                        const ctx = canvas.getContext('2d');
                        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
                        const mime = file.type && file.type.startsWith('image/') ? (file.type.includes('png') ? 'image/png' : 'image/jpeg') : 'image/jpeg';
                        const dataUrl = canvas.toDataURL(mime, quality);
                        resolve(dataUrl);
                    };
                    img.onerror = () => reject(new Error('Image load failed'));
                    img.src = e.target.result;
                };
                reader.onerror = () => reject(new Error('File read failed'));
                reader.readAsDataURL(file);
            } catch (err) {
                reject(err);
            }
        });
    }
}

// Initialize the editor when the page loads
let editor;
document.addEventListener('DOMContentLoaded', () => {
    editor = new NewsletterEditor();
    // Expose globally for inline modal to access lastAction and saveToHistory
    try { window.editor = editor; } catch (_) {}
});

// --- IndexedDB helpers for full snapshot HTML (preserves media src) ---

// Sanitize HTML before storing full snapshots for preview/live use.
// This removes editor-only UI elements (e.g. video toolbar handles)
// from the stored HTML without touching the live editor DOM.
function sanitizeFullHtmlForHistory(html) {
    try {
        const wrapper = document.createElement('div');
        wrapper.innerHTML = String(html || '');
        wrapper.querySelectorAll('.video-toolbar-handle').forEach(el => {
            if (el && el.parentNode) el.parentNode.removeChild(el);
        });
        return wrapper.innerHTML;
    } catch (_) {
        return String(html || '');
    }
}

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

function saveFullContentToIDB(id, html) {
    openHistoryDB()
        .then(db => {
            const tx = db.transaction('historyFull', 'readwrite');
            const cleanHtml = sanitizeFullHtmlForHistory(html);
            tx.objectStore('historyFull').put({ id: String(id), content: String(cleanHtml || '') });
            tx.oncomplete = () => { try { db.close(); } catch (_) {} };
            tx.onerror = () => { try { db.close(); } catch (_) {} };
        })
        .catch(() => {});
}

function getFullContentFromIDB(id) {
    return openHistoryDB()
        .then(db => new Promise(resolve => {
            const tx = db.transaction('historyFull', 'readonly');
            const req = tx.objectStore('historyFull').get(String(id));
            req.onsuccess = () => {
                try { db.close(); } catch (_) {}
                resolve((req.result && req.result.content) || '');
            };
            req.onerror = () => {
                try { db.close(); } catch (_) {}
                resolve('');
            };
        }))
        .catch(() => '');
}