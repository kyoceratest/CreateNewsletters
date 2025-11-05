// History modal and sidebar initialization
// All identifiers and comments in English; on-screen text remains in French.
(function(){
  document.addEventListener('DOMContentLoaded', function() {
    // Sidebar toggle (mobile)
    const sidebar = document.querySelector('.editor-sidebar');
    const overlay = document.getElementById('sidebarOverlay');
    const burger = document.getElementById('sidebarHamburger');
    if (sidebar && overlay && burger) {
      const openSidebar = () => {
        sidebar.classList.add('active');
        overlay.classList.add('active');
        document.body.style.overflow = 'hidden';
        sidebar.setAttribute('aria-expanded', 'true');
        burger.setAttribute('aria-expanded', 'true');
      };
      const closeSidebar = () => {
        sidebar.classList.remove('active');
        overlay.classList.remove('active');
        document.body.style.overflow = '';
        sidebar.setAttribute('aria-expanded', 'false');
        burger.setAttribute('aria-expanded', 'false');
      };
      const toggleSidebar = (e) => {
        e.stopPropagation();
        const isOpen = sidebar.classList.contains('active');
        isOpen ? closeSidebar() : openSidebar();
      };
      burger.addEventListener('click', toggleSidebar);
      overlay.addEventListener('click', closeSidebar);
      document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSidebar(); });
    }

    // History Modal guard
    const historyModal = document.getElementById('historyModal');
    if (!historyModal) return;
    if (window.__historyModalWired) return;
    window.__historyModalWired = true;

    // Elements
    const closeHistoryModal = document.getElementById('closeHistoryModal');
    const historyBtn = document.getElementById('historyBtn');
    const cancelHistoryBtn = document.getElementById('cancelHistoryBtn');
    const previewHistoryBtn = document.getElementById('previewHistoryBtn');
    const deleteHistoryBtn = document.getElementById('deleteHistoryBtn');
    const historySearch = document.getElementById('historySearch');
    const historyList = document.getElementById('historyList');
    const selectedCountEl = document.getElementById('selectedCount');
    const saveSnapshotBtn = document.getElementById('saveSnapshotBtn');

    // Data state
    window.historyData = Array.isArray(window.historyData) ? window.historyData : [];
    let selectedItems = new Set();
    let allHistoryItems = [];

    // Storage -> in-memory mapping
    function refreshHistoryFromStorage() {
      try {
        const raw = localStorage.getItem('newsletterHistory');
        const persisted = raw ? JSON.parse(raw) : [];
        const fallbackRaw = sessionStorage.getItem('newsletterHistoryFallback');
        const fallbackList = fallbackRaw ? JSON.parse(fallbackRaw) : [];
        const bufferList = Array.isArray(window.__historyBuffer) ? window.__historyBuffer : [];
        const merged = []
          .concat(Array.isArray(persisted) ? persisted : [])
          .concat(Array.isArray(fallbackList) ? fallbackList : [])
          .concat(Array.isArray(bufferList) ? bufferList : []);
        if (Array.isArray(merged)) {
          historyData.length = 0;
          historyData.push(...merged.map(item => ({
            id: String(item.id),
            title: item.name || 'Sans nom',
            date: item.date || (item.timestamp ? new Date(item.timestamp).toLocaleString('fr-FR') : ''),
            content: item.content || '',
            lastAction: item.lastAction || 'Action inconnue'
          })));
        }
      } catch (_) {}
    }

    // Render list
    function renderHistoryItems(items = historyData) {
      if (!historyList) return;
      historyList.innerHTML = '';
      items.forEach(item => {
        const displayTitle = computeDisplayTitle(item);
        const isSelected = selectedItems.has(item.id);
        const itemElement = document.createElement('div');
        itemElement.className = `history-item ${isSelected ? 'selected' : ''}`;
        itemElement.setAttribute('role', 'option');
        itemElement.setAttribute('aria-selected', isSelected);
        itemElement.setAttribute('tabindex', '0');
        itemElement.setAttribute('data-id', item.id);
        itemElement.innerHTML = `
          <label class="history-checkbox">
            <input type="checkbox" class="history-select" ${isSelected ? 'checked' : ''} data-id="${item.id}">
            <span class="checkmark"></span>
          </label>
          <div class="history-item-content">
            <div class="history-title">${displayTitle}</div>
            <div class="history-lastaction" style="color:#666; font-size:13px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">${item.lastAction || 'Action inconnue'}</div>
            <div class="history-date">${item.date}</div>
          </div>
          <div class="history-actions">
            <button class="btn-icon preview-btn restore-btn" data-id="${item.id}" title="Restaurer"><i class="fas fa-rotate-left"></i></button>
            <button class="btn-icon delete-btn" data-id="${item.id}" title="Supprimer"><i class="fas fa-trash"></i></button>
          </div>`;
        historyList.appendChild(itemElement);
      });
      updateSelectionUI();
    }

    // Compute display title: prefer the element with the largest inline font-size (>=46px),
    // then any element with inline font-size:52px, then h1/h2/h3, then stored title
    function computeDisplayTitle(item) {
      try {
        const html = item && item.content ? String(item.content) : '';
        if (html) {
          const tmp = document.createElement('div');
          tmp.innerHTML = html;
          const all = tmp.querySelectorAll('*');

          // 1) Find element with largest inline font-size
          let bestEl = null;
          let bestSize = 0;
          for (const el of all) {
            const styleAttr = (el.getAttribute && el.getAttribute('style')) || '';
            let sizePx = 0;
            // Match px values
            const mPx = styleAttr.match(/font-size\s*:\s*(\d+(?:\.\d+)?)\s*px/i);
            if (mPx) {
              sizePx = parseFloat(mPx[1]);
            } else {
              // Match rem and convert to px assuming root 16px (best-effort)
              const mRem = styleAttr.match(/font-size\s*:\s*(\d+(?:\.\d+)?)\s*rem/i);
              if (mRem) sizePx = parseFloat(mRem[1]) * 16;
              else {
                // Also support CSS font shorthand e.g., font: 700 52px/1.2 "Montserrat"
                const mSh = styleAttr.match(/font\s*:\s*[^;]*?(\d+(?:\.\d+)?)\s*px/i);
                if (mSh) sizePx = parseFloat(mSh[1]);
              }
            }
            if (sizePx > bestSize) {
              bestSize = sizePx;
              bestEl = el;
            }
          }
          if (bestEl && bestSize >= 46 && bestEl.textContent && bestEl.textContent.trim()) {
            return bestEl.textContent.trim();
          }

          // 2) Explicit 52px fallback
          let hit = null;
          for (const el of all) {
            const styleAttr = (el.getAttribute && el.getAttribute('style')) || '';
            if (styleAttr && /font-size\s*:\s*52px/i.test(styleAttr)) { hit = el; break; }
            if (el.style && (el.style.fontSize || '').toLowerCase() === '52px') { hit = el; break; }
          }
          if (hit && hit.textContent && hit.textContent.trim()) return hit.textContent.trim();

          const h = tmp.querySelector('h1, h2, h3');
          if (h && (h.textContent || '').trim()) return h.textContent.trim();
        }
      } catch (_) { /* ignore */ }
      // Fallbacks: stored title then empty string
      return (item && item.title) ? item.title : '';
    }

    function updateSelectionUI() {
      const selectedCount = selectedItems.size;
      if (selectedCountEl) selectedCountEl.textContent = selectedCount;
      allHistoryItems = Array.from(historyList.querySelectorAll('.history-item'));
      allHistoryItems.forEach(item => {
        const itemId = item.getAttribute('data-id');
        const isSelected = selectedItems.has(itemId);
        item.setAttribute('aria-selected', isSelected);
        isSelected ? item.classList.add('selected') : item.classList.remove('selected');
      });
      if (previewHistoryBtn) previewHistoryBtn.disabled = selectedCount === 0;
      if (deleteHistoryBtn) deleteHistoryBtn.disabled = selectedCount === 0;
    }

    function toggleItemSelection(item) {
      const itemId = item.getAttribute('data-id');
      const checkbox = item.querySelector('.history-select');
      if (selectedItems.has(itemId)) {
        selectedItems.delete(itemId);
        item.classList.remove('selected');
        item.setAttribute('aria-selected', 'false');
        if (checkbox) checkbox.checked = false;
      } else {
        selectedItems.add(itemId);
        item.classList.add('selected');
        item.setAttribute('aria-selected', 'true');
        if (checkbox) checkbox.checked = true;
      }
      updateSelectionUI();
    }

    function filterHistoryItems(searchTerm) {
      const filtered = historyData.filter(item =>
        item.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
        item.date.includes(searchTerm)
      );
      renderHistoryItems(filtered);
    }

    if (historyList) {
      historyList.addEventListener('click', (e) => {
        const item = e.target.closest('.history-item');
        const delBtn = e.target.closest('.delete-btn');
        const restoreBtn = e.target.closest('.restore-btn');
        if (item) {
          if (restoreBtn) {
            const id = restoreBtn.getAttribute('data-id');
            restoreItem(id);
          } else if (delBtn) {
            const id = delBtn.getAttribute('data-id');
            deleteItems([id]);
          } else {
            toggleItemSelection(item);
          }
        }
      });
      historyList.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          const item = e.target.closest('.history-item');
          if (item) toggleItemSelection(item);
        }
      });
    }

    if (historySearch) {
      let searchTimeout;
      historySearch.addEventListener('input', (e) => {
        clearTimeout(searchTimeout);
        searchTimeout = setTimeout(() => filterHistoryItems(e.target.value), 300);
      });
    }

    function openPreview(itemIds) {
      if (itemIds.length === 0) return;
      const item = historyData.find(i => String(i.id) === String(itemIds[0]));
      if (!item) return;
      const previewWindow = window.open(`preview.html?historyId=${item.id}`, 'previewWindow');
      const checkLoaded = setInterval(() => {
        if (!previewWindow) { clearInterval(checkLoaded); return; }
        if (previewWindow.document.readyState === 'complete') {
          clearInterval(checkLoaded);
          previewWindow.postMessage({
            type: 'previewContent',
            content: item.content || `<div class="preview-content"><h1>${item.title}</h1><p>Date de création : ${item.date}</p></div>`,
            title: item.title,
            date: item.date
          }, '*');
        }
      }, 100);
    }

    async function restoreItem(itemId) {
      const item = historyData.find(i => String(i.id) === String(itemId));
      if (!item) return;
      const editable = document.getElementById('editableContent');
      if (editable) {
        let html = '';
        try {
          if (typeof getFullContentFromIDB === 'function') {
            html = await getFullContentFromIDB(item.id);
          }
        } catch (_) { html = ''; }
        if (!html) html = item.content || '';
        editable.innerHTML = '';
        const wrapper = document.createElement('div');
        wrapper.innerHTML = html;
        while (wrapper.firstChild) editable.appendChild(wrapper.firstChild);
      }
      closeModal();
    }

    function deleteItems(itemIds) {
      if (itemIds.length === 0) return;
      try {
        const raw = localStorage.getItem('newsletterHistory');
        const persisted = raw ? JSON.parse(raw) : [];
        const updated = Array.isArray(persisted) ? persisted.filter(h => !itemIds.includes(String(h.id))) : [];
        localStorage.setItem('newsletterHistory', JSON.stringify(updated));
        // Reflect in UI state
        itemIds.forEach(id => {
          const idx = historyData.findIndex(i => String(i.id) === String(id));
          if (idx !== -1) historyData.splice(idx, 1);
        });
        // Remove DOM nodes
        itemIds.forEach(id => {
          const node = Array.from(historyList.querySelectorAll('.history-item'))
            .find(el => String(el.getAttribute('data-id')) === String(id));
          if (node) node.remove();
        });
        selectedItems.clear();
        updateSelectionUI();
        alert(`${itemIds.length} élément(s) supprimé(s) avec succès.`);
      } catch (e) {
        console.warn('Failed to persist history deletion:', e);
      }
    }

    function openModal() {
      historyModal.style.display = 'flex';
      historyModal.setAttribute('aria-hidden', 'false');
      document.body.style.overflow = 'hidden';
      // Make modal interactive and visible
      historyModal.style.opacity = '1';
      historyModal.style.pointerEvents = 'auto';
      const modalContent = historyModal.querySelector('.history-modal');
      if (modalContent) {
        modalContent.style.opacity = '1';
        modalContent.style.transform = 'translateY(0)';
      }
      // Reset search and selection
      if (historySearch) historySearch.value = '';
      selectedItems.clear();
      // Reload latest history from storage, migrate titles, and render
      refreshHistoryFromStorage();
      migrateHistoryTitles();
      renderHistoryItems();
      // Focus on search input shortly after opening
      setTimeout(() => { historySearch && historySearch.focus(); }, 100);
    }

    function closeModal() {
      historyModal.style.opacity = '0';
      historyModal.style.pointerEvents = 'none';
      historyModal.style.display = 'none';
      historyModal.setAttribute('aria-hidden', 'true');
      document.body.style.overflow = '';
      const modalContent = historyModal.querySelector('.history-modal');
      if (modalContent) {
        modalContent.style.opacity = '0';
        modalContent.style.transform = 'translateY(-20px)';
      }
      const historyBtn = document.getElementById('historyBtn');
      historyBtn && historyBtn.focus();
    }

    if (closeHistoryModal) closeHistoryModal.addEventListener('click', closeModal);
    if (cancelHistoryBtn) cancelHistoryBtn.addEventListener('click', closeModal);
    historyModal.addEventListener('click', (e) => { if (e.target === historyModal) closeModal(); });
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && historyModal.style.display === 'flex') closeModal();
    });

    if (historyBtn) {
      historyBtn.addEventListener('click', (e) => {
        e.preventDefault();
        openModal();
      });
    }

    if (previewHistoryBtn) previewHistoryBtn.addEventListener('click', () => openPreview(Array.from(selectedItems)));

    if (deleteHistoryBtn) deleteHistoryBtn.addEventListener('click', () => deleteItems(Array.from(selectedItems)));

    if (saveSnapshotBtn) {
      saveSnapshotBtn.addEventListener('click', () => {
        try {
          const editable = document.getElementById('editableContent');
          if (!editable) { alert("Zone d'édition introuvable."); return; }
          const content = editable.innerHTML;
          let computedName = '';
          try {
            const tmp = document.createElement('div');
            tmp.innerHTML = content || '';
            const all = tmp.querySelectorAll('*');
            let hit = null;
            for (const el of all) {
              const styleAttr = (el.getAttribute && el.getAttribute('style')) || '';
              if (styleAttr && /font-size\s*:\s*52px/i.test(styleAttr)) { hit = el; break; }
              if (el.style && (el.style.fontSize || '').toLowerCase() === '52px') { hit = el; break; }
            }
            if (hit && hit.textContent) computedName = hit.textContent.trim();
            if (!computedName) {
              const h = tmp.querySelector('h1, h2, h3');
              if (h) computedName = (h.textContent || '').trim();
            }
            if (!computedName) computedName = (tmp.textContent || '').trim().substring(0, 80);
          } catch (_) {}
          if (!computedName) computedName = document.title && document.title.trim() ? document.title.trim() : 'instantané_' + new Date().toLocaleString('fr-FR');

          let savedOk = false;
          try {
            const rawExisting = localStorage.getItem('newsletterHistory');
            let existing = rawExisting ? JSON.parse(rawExisting) : [];
            if (!Array.isArray(existing)) existing = [];
            const nameToMatch = (computedName || '').trim();
            const dupIndex = existing.findIndex(e => (e && (e.name || '')).trim() === nameToMatch);
            if (dupIndex !== -1) {
              const confirmReplace = window.confirm(`Un instantané nommé "${computedName}" existe déjà. Voulez-vous le remplacer ?`);
              if (confirmReplace) {
                const preview = content.replace(/<[^>]*>?/gm, '').substring(0, 150) + '...';
                const lastAction = (window.editor && window.editor.lastAction) ? window.editor.lastAction : 'Action inconnue';
                const oldId = existing[dupIndex].id || Date.now().toString();
                existing[dupIndex] = {
                  ...existing[dupIndex],
                  id: String(oldId),
                  name: computedName,
                  content,
                  date: new Date().toLocaleString('fr-FR'),
                  preview,
                  timestamp: Date.now(),
                  lastAction
                };
                try { if (typeof saveFullContentToIDB === 'function') saveFullContentToIDB(String(oldId), content); } catch (_) {}
                localStorage.setItem('newsletterHistory', JSON.stringify(existing));
                savedOk = true;
              } else {
                savedOk = false;
              }
            }
          } catch (_) { /* ignore duplicate check errors and fallback to normal flow */ }

          if (!savedOk) {
            if (window.editor && typeof window.editor.saveToHistory === 'function') {
              const res = window.editor.saveToHistory(computedName, content);
              if (res) savedOk = true;
            }
          }
          if (!savedOk) {
            try {
              const raw = localStorage.getItem('newsletterHistory');
              let history = raw ? JSON.parse(raw) : [];
              if (!Array.isArray(history)) history = [];
              const id = Date.now().toString();
              const preview = content.replace(/<[^>]*>?/gm, '').substring(0, 150) + '...';
              const lastAction = (window.editor && window.editor.lastAction) ? window.editor.lastAction : 'Action inconnue';
              const item = { id, name: computedName, content, date: new Date().toLocaleString('fr-FR'), preview, timestamp: Date.now(), lastAction };
              history.unshift(item);
              if (history.length > 200) history = history.slice(0, 200);
              localStorage.setItem('newsletterHistory', JSON.stringify(history));
              try { if (typeof saveFullContentToIDB === 'function') saveFullContentToIDB(id, content); } catch (_) {}
              savedOk = true;
            } catch (e) { console.warn('Direct save to newsletterHistory failed:', e); }
          }
          if (savedOk) { refreshHistoryFromStorage(); renderHistoryItems(); alert('Instantané sauvegardé avec succès.'); }
          else { alert("Échec de la sauvegarde de l'instantané."); }
        } catch (e) {
          console.error('Save snapshot error:', e);
          alert("Erreur lors de la sauvegarde de l'instantané.");
        }
      });
    }

    // Footer select/deselect all
    const selectAllFooterBtn = document.getElementById('selectAllFooterBtn');
    const deselectAllFooterBtn = document.getElementById('deselectAllFooterBtn');
    if (selectAllFooterBtn && historyList) {
      selectAllFooterBtn.addEventListener('click', () => {
        const checkboxes = historyList.querySelectorAll('.history-select');
        checkboxes.forEach(cb => { const id = cb.getAttribute('data-id'); cb.checked = true; selectedItems.add(id); });
        updateSelectionUI();
      });
    }
    if (deselectAllFooterBtn && historyList) {
      deselectAllFooterBtn.addEventListener('click', () => {
        const checkboxes = historyList.querySelectorAll('.history-select');
        checkboxes.forEach(cb => { const id = cb.getAttribute('data-id'); cb.checked = false; selectedItems.delete(id); });
        updateSelectionUI();
      });
    }

    // Initial render
    refreshHistoryFromStorage();
    migrateHistoryTitles();
    renderHistoryItems();
  });
})();

// One-time, lightweight migration to ensure the first line uses the computed 52px (or largest) title
function migrateHistoryTitles() {
  try {
    const raw = localStorage.getItem('newsletterHistory');
    let list = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(list) || list.length === 0) return;
    let changed = false;
    const updated = list.map(it => {
      try {
        const displayTitle = (typeof computeDisplayTitle === 'function') ? computeDisplayTitle({ content: it.content, title: it.name }) : (it.name || '');
        if (displayTitle && displayTitle !== it.name) {
          changed = true;
          return { ...it, name: displayTitle };
        }
      } catch (_) {}
      return it;
    });
    if (changed) localStorage.setItem('newsletterHistory', JSON.stringify(updated));
  } catch (_) { /* ignore */ }
}
