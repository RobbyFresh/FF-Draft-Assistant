/* global fetch, document */
(function () {
  const statusEl = document.getElementById('status');
  const fileMetaEl = document.getElementById('file-meta');
  const titleEl = document.getElementById('page-title');
  const tabsEl = document.getElementById('tabs');
  const sheetContainerEl = document.getElementById('sheet-container');
  const myTeamListEl = null; // removed
  const takenListEl = document.getElementById('taken-list');
  const rosterGridEl = document.getElementById('roster-grid');
  const addSlotBtn = document.getElementById('add-slot');
  const removeSlotBtn = document.getElementById('remove-slot');
  const slotTypeSelect = document.getElementById('slot-type');
  const resetAllBtn = document.getElementById('reset-all');
  const paneSplitter = document.getElementById('pane-splitter');
  const myTeamEl = document.getElementById('my-team');
  let playerSearchInput = null;
  const refreshBtn = document.getElementById('refresh-page');

  // Global bottom scrollbar elements
  let globalScroll = null;
  let globalSpacer = null;
  let isSyncing = false;

  function ensureGlobalScroller() {
    if (globalScroll) return;
    globalScroll = document.createElement('div');
    globalScroll.className = 'global-hscrollbar';
    globalSpacer = document.createElement('div');
    globalSpacer.className = 'global-hscrollbar-spacer';
    globalScroll.appendChild(globalSpacer);
    document.body.appendChild(globalScroll);

    // When bottom bar scrolls, move the active wrapper
    globalScroll.addEventListener('scroll', () => {
      if (isSyncing) return; isSyncing = true;
      const active = getActiveWrapper();
      if (active) active.scrollLeft = globalScroll.scrollLeft;
      isSyncing = false;
    });

    window.addEventListener('resize', updateGlobalScrollerWidth);
  }

  function getActiveWrapper() {
    const wrappers = [...sheetContainerEl.querySelectorAll('.table-scroll')];
    return wrappers.find((el) => el.style.display !== 'none');
  }

  function updateGlobalScrollerWidth() {
    const active = getActiveWrapper();
    if (!active) return;
    const table = active.querySelector('table');
    if (!table) return;
    // Match spacer to table scroll width
    globalSpacer.style.width = table.scrollWidth + 'px';
    globalScroll.scrollLeft = active.scrollLeft;
  }

  // Safe attribute value escaper for CSS attribute selectors
  function escapeAttrValue(value) {
    return String(value).replace(/["\\]/g, '\\$&');
  }

  const getMyTeamKey = (player) => `myteam:${player}`;
  const getTakenKey = (player) => `taken:${player}`;
  const getLegacyDraftKey = (player) => `draftedAny:${player}`; // migrate to My Team
  const isInMyTeam = (player) => localStorage.getItem(getMyTeamKey(player)) === '1' || localStorage.getItem(getLegacyDraftKey(player)) === '1';
  const isTaken = (player) => localStorage.getItem(getTakenKey(player)) !== null;
  function setMyTeam(player, val) { if (val) localStorage.setItem(getMyTeamKey(player), '1'); else localStorage.removeItem(getMyTeamKey(player)); localStorage.removeItem(getLegacyDraftKey(player)); }
  function setTaken(player, val) { if (val) localStorage.setItem(getTakenKey(player), String(Date.now())); else localStorage.removeItem(getTakenKey(player)); }

  function updateAllCheckboxesForPlayer(playerName) {
    const esc = escapeAttrValue(playerName);
    const myBoxes = document.querySelectorAll(`input[type="checkbox"][data-role="myteam"][data-player="${esc}"]`);
    const takenBoxes = document.querySelectorAll(`input[type="checkbox"][data-role="taken"][data-player="${esc}"]`);
    const myChecked = isInMyTeam(playerName);
    const takenChecked = isTaken(playerName);
    myBoxes.forEach((inp) => { inp.checked = myChecked; });
    takenBoxes.forEach((inp) => { inp.checked = takenChecked; });
  }

  function updateRowVisibilityForPlayer(playerName) {
    const hide = isInMyTeam(playerName) || isTaken(playerName);
    const esc = escapeAttrValue(playerName);
    const rows = document.querySelectorAll(`tr[data-player="${esc}"]`);
    rows.forEach((tr) => { tr.style.display = hide ? 'none' : ''; });
  }

  function rebuildMyTeamUI() {
    if (!takenListEl) return;
    const mine = [];
    const taken = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const k = localStorage.key(i);
      if (!k) continue;
      if (k.startsWith('myteam:')) mine.push(k.slice('myteam:'.length));
      else if (k.startsWith('taken:')) taken.push({ name: k.slice('taken:'.length), ts: Number(localStorage.getItem(k)) || 0 });
      else if (k.startsWith('draftedAny:')) mine.push(k.slice('draftedAny:'.length)); // legacy
    }
    mine.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
    taken.sort((a, b) => b.ts - a.ts);
    takenListEl.innerHTML = '';
    const makeLi = (name, list) => {
      const li = document.createElement('li');
      const span = document.createElement('span');
      span.textContent = name;
      const btn = document.createElement('button');
      btn.className = 'my-team-remove';
      btn.textContent = 'Remove';
      btn.addEventListener('click', () => {
        // Removing from either list clears both flags
        setMyTeam(name, false);
        setTaken(name, false);
        updateAllCheckboxesForPlayer(name);
        updateRowVisibilityForPlayer(name);
        rebuildMyTeamUI();
      });
      li.appendChild(span);
      li.appendChild(btn);
      list.appendChild(li);
    };
    taken.forEach((item) => makeLi(item.name, takenListEl));
  }

  function setMyTeamState(playerName, on) {
    if (on) { setMyTeam(playerName, true); setTaken(playerName, false); }
    else { setMyTeam(playerName, false); unassignPlayer(playerName); }
    updateAllCheckboxesForPlayer(playerName);
    updateRowVisibilityForPlayer(playerName);
    rebuildMyTeamUI();
    renderRoster();
  }

  function setTakenState(playerName, on) {
    if (on) { setTaken(playerName, true); setMyTeam(playerName, false); unassignPlayer(playerName); }
    else { setTaken(playerName, false); }
    updateAllCheckboxesForPlayer(playerName);
    updateRowVisibilityForPlayer(playerName);
    rebuildMyTeamUI();
    renderRoster();
  }

  // Roster management
  const DEFAULT_ROSTER = [
    { type: 'QB', count: 1 },
    { type: 'RB', count: 2 },
    { type: 'WR', count: 2 },
    { type: 'TE', count: 2 },
    { type: 'FLEX', count: 1 },
    { type: 'SFLX', count: 0 },
    { type: 'K', count: 1 },
    { type: 'DST', count: 1 },
    { type: 'BENCH', count: 7 },
  ];
  const ROSTER_KEY = 'rosterSlots';
  function loadRoster() {
    try {
      const raw = localStorage.getItem(ROSTER_KEY);
      if (!raw) return DEFAULT_ROSTER.slice();
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) return DEFAULT_ROSTER.slice();
      return parsed.filter((r) => r && r.type && Number.isFinite(r.count));
    } catch { return DEFAULT_ROSTER.slice(); }
  }
  function saveRoster(arr) {
    localStorage.setItem(ROSTER_KEY, JSON.stringify(arr));
  }
  function addRosterSlot(type) {
    const roster = loadRoster();
    const idx = roster.findIndex((r) => r.type === type);
    if (idx >= 0) roster[idx].count += 1; else roster.push({ type, count: 1 });
    saveRoster(roster);
    renderRoster();
  }
  function removeRosterSlot(type) {
    const roster = loadRoster();
    const idx = roster.findIndex((r) => r.type === type);
    if (idx >= 0) {
      roster[idx].count -= 1;
      if (roster[idx].count <= 0) roster.splice(idx, 1);
      saveRoster(roster);
      renderRoster();
    }
  }
  function renderRoster() {
    if (!rosterGridEl) return;
    const roster = loadRoster();
    const assigns = loadAssignments();
    const removed = reconcileAssignmentsWithRoster(roster, assigns);
    saveAssignments(assigns);
    // Clear any players that were ejected due to capacity shrink
    removed.forEach((name) => setMyTeam(name, false));

    rosterGridEl.innerHTML = '';
    roster.forEach((row) => {
      const rowEl = document.createElement('div');
      rowEl.className = 'roster-row';
      const title = document.createElement('div');
      title.className = 'roster-row-title';
      title.textContent = `${row.type} (${row.count})`;
      const slots = document.createElement('div');
      slots.className = 'slot-list';
      const filled = assigns[row.type] || [];
      for (let i = 0; i < row.count; i += 1) {
        const occupant = filled[i] || '';
        const chip = document.createElement('div');
        chip.className = 'chip';
        if (occupant) {
          // show name with position color
          const posForOcc = (() => {
            // find POS heuristically from eligible slot
            if (row.type === 'QB') return 'QB';
            if (row.type === 'RB') return 'RB';
            if (row.type === 'WR') return 'WR';
            if (row.type === 'TE') return 'TE';
            if (row.type === 'K') return 'K';
            if (row.type === 'DST') return 'DST';
            // For FLEX/SFLX/BENCH we cannot easily infer; leave uncolored
            return '';
          })();
          if (posForOcc) chip.classList.add(posClass(posForOcc));
          chip.textContent = occupant;
        } else {
          chip.textContent = row.type;
        }
        const rm = document.createElement('button');
        rm.className = 'remove';
        rm.textContent = '×';
        rm.title = occupant ? `Remove ${occupant}` : 'Remove slot';
        rm.addEventListener('click', () => {
          if (occupant) {
            setMyTeamState(occupant, false);
          } else {
            removeRosterSlot(row.type);
          }
        });
        chip.appendChild(rm);
        slots.appendChild(chip);
      }
      rowEl.appendChild(title);
      rowEl.appendChild(slots);
      rosterGridEl.appendChild(rowEl);
    });
  }
  if (addSlotBtn && removeSlotBtn && slotTypeSelect) {
    addSlotBtn.addEventListener('click', () => addRosterSlot(slotTypeSelect.value));
    removeSlotBtn.addEventListener('click', () => removeRosterSlot(slotTypeSelect.value));
  }
  if (resetAllBtn) {
    resetAllBtn.addEventListener('click', () => {
      // Clear selections
      const names = [];
      for (let i = 0; i < localStorage.length; i += 1) {
        const k = localStorage.key(i);
        if (!k) continue;
        if (k.startsWith('myteam:')) names.push(k.slice('myteam:'.length));
        if (k.startsWith('taken:')) names.push(k.slice('taken:'.length));
      }
      names.forEach((n) => {
        setMyTeam(n, false);
        setTaken(n, false);
        updateAllCheckboxesForPlayer(n);
        updateRowVisibilityForPlayer(n);
      });
      // Reset roster to default
      saveRoster(DEFAULT_ROSTER.slice());
      // Clear assignments
      saveAssignments({});
      renderRoster();
      rebuildMyTeamUI();
    });
  }

  // Resizable panes in sidebar
  if (paneSplitter && myTeamEl) {
    let dragging = false;
    let startY = 0;
    let startTopFraction = 0.5;
    function readFractions() {
      const cs = getComputedStyle(myTeamEl);
      const topPart = cs.getPropertyValue('--pane-top').trim() || '1fr';
      const bottomPart = cs.getPropertyValue('--pane-bottom').trim() || '1fr';
      // naive fractions; default 1fr/1fr => 0.5
      if (topPart.endsWith('fr') && bottomPart.endsWith('fr')) {
        const t = parseFloat(topPart); const b = parseFloat(bottomPart);
        if (t + b > 0) return t / (t + b);
      }
      return 0.5;
    }
    function setFractions(f) {
      const clamped = Math.min(0.85, Math.max(0.15, f));
      myTeamEl.style.setProperty('--pane-top', clamped + 'fr');
      myTeamEl.style.setProperty('--pane-bottom', (1 - clamped) + 'fr');
      localStorage.setItem('paneTopFrac', String(clamped));
    }
    const saved = Number(localStorage.getItem('paneTopFrac'));
    if (!Number.isNaN(saved) && saved > 0 && saved < 1) setFractions(saved);

    function onMove(e) {
      if (!dragging) return;
      const rect = myTeamEl.getBoundingClientRect();
      const total = rect.height;
      const delta = (e.clientY - startY) / total;
      setFractions(startTopFraction + delta);
      e.preventDefault();
    }
    function onUp() {
      dragging = false;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    }
    paneSplitter.addEventListener('mousedown', (e) => {
      dragging = true;
      startY = e.clientY;
      startTopFraction = readFractions();
      window.addEventListener('mousemove', onMove);
      window.addEventListener('mouseup', onUp);
      e.preventDefault();
    });
    // Keyboard: up/down to adjust
    paneSplitter.addEventListener('keydown', (e) => {
      const step = 0.05;
      if (e.key === 'ArrowUp') { setFractions(readFractions() - step); e.preventDefault(); }
      if (e.key === 'ArrowDown') { setFractions(readFractions() + step); e.preventDefault(); }
    });
  }

  // Assignment storage: which players occupy which slots
  const ASSIGN_KEY = 'assignments';
  function loadAssignments() {
    try {
      const raw = localStorage.getItem(ASSIGN_KEY);
      const parsed = raw ? JSON.parse(raw) : {};
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch { return {}; }
  }
  function saveAssignments(obj) { localStorage.setItem(ASSIGN_KEY, JSON.stringify(obj)); }
  function clearAssignmentsForPlayer(player) {
    const a = loadAssignments();
    Object.keys(a).forEach((slot) => {
      a[slot] = (a[slot] || []).filter((p) => p !== player);
    });
    saveAssignments(a);
  }

  // Reconcile: trim assignments if roster counts decreased
  function reconcileAssignmentsWithRoster(roster, assigns) {
    const removed = [];
    const wantCounts = Object.fromEntries(roster.map((r) => [r.type, r.count]));
    Object.keys(assigns).forEach((slot) => {
      const cap = wantCounts[slot] ?? 0;
      if ((assigns[slot] || []).length > cap) {
        const extra = assigns[slot].splice(cap);
        removed.push(...extra);
      }
    });
    return removed;
  }

  function positionEligibility(pos) {
    const p = String(pos || '').toUpperCase();
    if (p === 'QB') return ['QB', 'SFLX', 'BENCH'];
    if (p === 'RB') return ['RB', 'FLEX', 'SFLX', 'BENCH'];
    if (p === 'WR') return ['WR', 'FLEX', 'SFLX', 'BENCH'];
    if (p === 'TE') return ['TE', 'FLEX', 'SFLX', 'BENCH'];
    if (p === 'K') return ['K', 'BENCH'];
    if (p === 'DST' || p === 'DEF') return ['DST', 'BENCH'];
    return ['BENCH'];
  }

  function posClass(posRaw) {
    const p = String(posRaw || '').toUpperCase();
    if (p === 'QB') return 'pos-QB';
    if (p === 'WR') return 'pos-WR';
    if (p === 'RB') return 'pos-RB';
    if (p === 'TE') return 'pos-TE';
    if (p === 'K') return 'pos-K';
    if (p === 'DST' || p === 'DEF') return 'pos-DST';
    return '';
  }

  function tryAssignPlayer(player, pos) {
    const roster = loadRoster();
    const assigns = loadAssignments();
    const eligible = positionEligibility(pos);
    // keep order priority
    for (const slot of eligible) {
      const cap = roster.find((r) => r.type === slot)?.count || 0;
      const arr = assigns[slot] || [];
      if (arr.length < cap) {
        arr.push(player);
        assigns[slot] = arr;
        saveAssignments(assigns);
        return true;
      }
    }
    return false;
  }

  function unassignPlayer(player) {
    const assigns = loadAssignments();
    let changed = false;
    Object.keys(assigns).forEach((slot) => {
      const before = assigns[slot].length;
      assigns[slot] = (assigns[slot] || []).filter((p) => p !== player);
      if (assigns[slot].length !== before) changed = true;
    });
    if (changed) saveAssignments(assigns);
  }

  /**
   * Render the workbook data into tabs and tables
   * @param {{fileName: string, sheets: {name: string, rows: any[][]}[]}} data
   */
  function renderWorkbook(data) {
    titleEl.textContent = 'Fantasy Football 2025 Draft Assistant';
    if (fileMetaEl) fileMetaEl.remove();
    statusEl.textContent = '';

    tabsEl.innerHTML = '';
    sheetContainerEl.innerHTML = '';

    const norm = (v) => String(v ?? '').toLowerCase().trim();
    const sheets = (data.sheets || []).filter((s) => norm(s.name) !== '8.5 (archived)');

    if (!sheets || sheets.length === 0) {
      statusEl.textContent = 'No sheets found in the workbook.';
      return;
    }

    sheets.forEach((sheet, index) => {
      const tab = document.createElement('button');
      tab.className = 'tab';
      tab.setAttribute('role', 'tab');
      tab.setAttribute('aria-selected', index === 0 ? 'true' : 'false');
      tab.textContent = sheet.name;
      tab.addEventListener('click', () => selectSheet(index));
      tabsEl.appendChild(tab);
    });

    // Append BOOM / BUST toggle button at the end of the tabs row (right side)
    (function attachBoomToggle() {
      const boomBtn = document.createElement('button');
      boomBtn.id = 'boom-toggle';
      boomBtn.className = 'btn boom-btn';
      boomBtn.setAttribute('aria-pressed', 'false');
      boomBtn.title = 'Highlight Boom/Bust rows';
      boomBtn.textContent = 'BOOM / BUST';
      tabsEl.appendChild(boomBtn);

      function setBoomState(on) {
        boomBtn.setAttribute('aria-pressed', on ? 'true' : 'false');
        const wrappers = [...sheetContainerEl.querySelectorAll('.table-scroll')];
        wrappers.forEach((wrap) => {
          const rows = wrap.querySelectorAll('tbody tr');
          rows.forEach((tr) => {
            const hadBoom = tr.dataset.hasBoom === '1';
            const hadBust = tr.dataset.hasBust === '1';
            tr.classList.remove('row-boom', 'row-bust');
            if (on) {
              if (hadBoom) tr.classList.add('row-boom');
              if (hadBust) tr.classList.add('row-bust');
            }
          });
        });
      }
      boomBtn.addEventListener('click', () => {
        const pressed = boomBtn.getAttribute('aria-pressed') === 'true';
        setBoomState(!pressed);
      });
      setBoomState(false);
    })();

    // Append Search box at the far right of the tabs row
    (function attachSearchBox() {
      const container = document.createElement('div');
      container.className = 'search-container';
      const input = document.createElement('input');
      input.id = 'player-search';
      input.className = 'search-input';
      input.type = 'search';
      input.placeholder = 'Search players...';
      input.setAttribute('aria-label', 'Search players');
      container.appendChild(input);
      tabsEl.appendChild(container);
      playerSearchInput = input;
      playerSearchInput.addEventListener('input', () => applySearchFilter());
    })();

    // Create one table element per sheet (we will toggle visibility)
    sheets.forEach((sheet, index) => {
      const wrapper = document.createElement('div');
      wrapper.className = 'table-scroll';
      wrapper.style.display = index === 0 ? 'block' : 'none';

      const table = document.createElement('table');
      table.dataset.sheetIndex = String(index);

      const thead = document.createElement('thead');
      const tbody = document.createElement('tbody');

      // Determine header row intelligently (look for a row containing 'player name')
      const rows = sheet.rows || [];
      const normalize = (v) => String(v ?? '').toLowerCase().trim();
      let headerRowIndex = 0;
      for (let i = 0; i < rows.length; i += 1) {
        const norm = (rows[i] || []).map(normalize);
        if (norm.some((c) => c.includes('player') && c.includes('name'))) {
          headerRowIndex = i;
          break;
        }
      }

      const headerRow = document.createElement('tr');
      const headers = rows[headerRowIndex] || [];
      const columnCount = headers.length || Math.max(...rows.map((r) => r.length), 0);

      // Identify key columns
      const playerColIndex = headers.findIndex((h) => normalize(h).includes('player') && normalize(h).includes('name'));
      const isPosHeader = (h) => {
        const t = normalize(h);
        return t === 'pos' || t === 'position';
      };
      const posColIndex = headers.findIndex((h) => isPosHeader(h));

      // Build visible columns, filtering out prop bets or auction
      const allIndices = Array.from({ length: columnCount }, (_, i) => i);
      const shouldHideColumn = (h) => {
        const t = normalize(h);
        return (t.includes('prop') && t.includes('bet')) || t.includes('auction');
      };
      let visibleIndices = allIndices.filter((idx) => !shouldHideColumn(headers[idx] ?? ''));

      // Reorder so Player Name is directly left of POS, if both exist
      const playerIdxInVisible = visibleIndices.indexOf(playerColIndex);
      const posIdxInVisible = visibleIndices.indexOf(posColIndex);
      if (playerIdxInVisible !== -1 && posIdxInVisible !== -1) {
        // Remove player from current position
        visibleIndices.splice(playerIdxInVisible, 1);
        const newPosIdx = visibleIndices.indexOf(posColIndex);
        const insertAt = Math.max(0, newPosIdx);
        visibleIndices.splice(insertAt, 0, playerColIndex);
      }

      // Construct column descriptors and inject synthetic My Team/Taken columns
      const isSuperflex = normalize(sheet.name) === 'superflex';
      const draftedColIndex = headers.findIndex((h) => normalize(h) === 'drafted' || normalize(h) === 'my team');
      const columns = visibleIndices.map((idx) => {
        const h = headers[idx];
        if (normalize(h) === 'drafted') return { kind: 'myteamExisting', idx };
        return { kind: 'sheet', idx };
      });
      if (playerColIndex !== -1) {
        const playerColInColumns = columns.findIndex((c) => c.idx === playerColIndex);
        const insertAt = Math.max(0, playerColInColumns);
        // Only inject if there isn't already a My Team column
        if (draftedColIndex === -1) {
          columns.splice(insertAt, 0, { kind: 'myteam' });
        } else {
          // Convert the existing drafted column descriptor if present in visibleIndices
          const existingIdx = columns.findIndex((c) => c.kind === 'myteamExisting');
          if (existingIdx !== -1) {
            // leave as is
          }
        }
      }
      // Insert Taken to the right of My Team (existing or injected)
      const mtIndex = columns.findIndex((c) => c.kind === 'myteam' || c.kind === 'myteamExisting');
      if (mtIndex !== -1) {
        columns.splice(mtIndex + 1, 0, { kind: 'taken' });
      }

      const lastVisibleSheetIdx = (() => {
        const onlySheets = (columns.length ? columns : visibleIndices.map((idx) => ({ kind: 'sheet', idx }))).filter((c) => c.kind === 'sheet');
        return onlySheets.length ? onlySheets[onlySheets.length - 1].idx : -1;
      })();
      (columns.length ? columns : visibleIndices.map((idx) => ({ kind: 'sheet', idx }))).forEach((col, i) => {
        const th = document.createElement('th');
        if (col.kind === 'myteam' || col.kind === 'myteamExisting') {
          th.textContent = 'My Team';
        } else if (col.kind === 'taken') {
          th.textContent = 'Taken';
        } else {
          const colIdx = col.idx;
          const label = String(headers[colIdx] ?? `Column ${colIdx + 1}`);
          th.textContent = label;
        }
        // Optional: disable sort for checkbox-only columns
        if (!(col.kind === 'myteam' || col.kind === 'myteamExisting' || col.kind === 'taken')) {
          th.addEventListener('click', () => sortTableByColumn(table, i));
        }
        // Sticky sequence: My Team, Taken, Player Name
        if (col.kind === 'myteam' || col.kind === 'myteamExisting') {
          th.classList.add('sticky-left-0', 'col-checkbox');
        } else if (col.kind === 'taken') {
          th.classList.add('sticky-left-1', 'col-checkbox');
        } else if (col.kind === 'sheet' && col.idx === playerColIndex) {
          th.classList.add('sticky-left-2');
        }
        if (col.kind === 'sheet' && col.idx === lastVisibleSheetIdx) th.classList.add('col-wide');
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);

      // Data rows start after headerRowIndex
      for (let r = headerRowIndex + 1; r < rows.length; r += 1) {
        const tr = document.createElement('tr');
        const playerNameForRow = String((rows[r] && rows[r][playerColIndex]) ?? '');
        tr.dataset.player = playerNameForRow;
        if (playerNameForRow && (isInMyTeam(playerNameForRow) || isTaken(playerNameForRow))) {
          tr.style.display = 'none';
        }
        // click-to-select toggle
        tr.addEventListener('click', (e) => {
          // ignore clicks on inputs (checkbox toggles)
          if (e.target && (e.target.tagName === 'INPUT' || e.target.closest('input'))) return;
          tr.classList.toggle('row-selected');
        });
        // annotate Boom/Bust classes
        const boomColIdx = headers.findIndex((h) => normalize(h) === 'boom factor');
        if (boomColIdx !== -1) {
          const boomText = String((rows[r] && rows[r][boomColIdx]) ?? '').toLowerCase();
          const hasBoom = boomText.includes('boom');
          const isDesperation = boomText.includes('desperation');
          const isWeak = boomText.includes('weak');
          // store flags for reliable toggling later
          if (hasBoom || isDesperation) tr.dataset.hasBoom = '1';
          if (isWeak) tr.dataset.hasBust = '1';
        }
        (columns.length ? columns : visibleIndices.map((idx) => ({ kind: 'sheet', idx }))).forEach((col) => {
          const td = document.createElement('td');
          if (col.kind === 'myteam' || col.kind === 'myteamExisting' || col.kind === 'taken') {
            // Checkbox stored in localStorage per player across sheets
            const playerName = playerNameForRow;
            const input = document.createElement('input');
            input.type = 'checkbox';
            input.setAttribute('data-player', playerName);
            if (col.kind === 'myteam' || col.kind === 'myteamExisting') {
              // Seed from existing cell for legacy 'drafted' columns
              if (col.kind === 'myteamExisting' && !isInMyTeam(playerName)) {
                const original = String((rows[r] && rows[r][col.idx]) ?? '').trim().toLowerCase();
                const truthy = original === 'true' || original === 'yes' || original === 'y' || original === '1' || original === 'x' || original === '✓' || original === 'drafted';
                if (truthy) setMyTeam(playerName, true);
              }
              input.setAttribute('data-role', 'myteam');
              input.checked = isInMyTeam(playerName);
              input.addEventListener('change', () => {
                if (input.checked) {
                  // Try assign based on POS
                  const posValue = String((rows[r] && rows[r][posColIndex]) ?? '');
                  const ok = tryAssignPlayer(playerName, posValue);
                  if (!ok) { input.checked = false; return; }
                } else {
                  unassignPlayer(playerName);
                }
                setMyTeamState(playerName, input.checked);
              });
            } else {
              input.setAttribute('data-role', 'taken');
              input.checked = isTaken(playerName);
              input.addEventListener('change', () => { setTakenState(playerName, input.checked); });
            }
            td.appendChild(input);
            // Sticky positions for checkbox columns
            if (col.kind === 'myteam' || col.kind === 'myteamExisting') {
              td.classList.add('sticky-left-0', 'col-checkbox');
            } else if (col.kind === 'taken') {
              td.classList.add('sticky-left-1', 'col-checkbox');
            }
          } else {
            const colIdx = col.idx;
            const value = (rows[r] && rows[r][colIdx]) ?? '';
            const text = String(value);
            // color POS and also color the player name cell using POS
            if (colIdx === posColIndex) {
              td.textContent = text;
              td.classList.add(posClass(text));
            } else if (colIdx === playerColIndex) {
              td.textContent = text;
              const posText = String((rows[r] && rows[r][posColIndex]) ?? '');
              const cls = posClass(posText);
              if (cls) td.classList.add(cls);
              td.classList.add('sticky-left-2');
            } else {
              td.textContent = text;
            }
            if (col.kind === 'sheet' && col.idx === lastVisibleSheetIdx) td.classList.add('col-wide');
          }
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      }

      table.appendChild(thead);
      table.appendChild(tbody);
      wrapper.appendChild(table);
      // Sync horizontal scroll via the table element instead of wrapper
      wrapper.addEventListener('wheel', (e) => {
        if (e.shiftKey && globalScroll) {
          globalScroll.scrollLeft += e.deltaY;
          e.preventDefault();
        }
      }, { passive: false });

      sheetContainerEl.appendChild(wrapper);
    });

    function selectSheet(index) {
      // Toggle tab selected state
      [...tabsEl.children].forEach((tab, i) => {
        tab.setAttribute('aria-selected', i === index ? 'true' : 'false');
      });
      // Toggle table visibility
      const wrappers = [...sheetContainerEl.querySelectorAll('.table-scroll')];
      wrappers.forEach((el, i) => {
        el.style.display = i === index ? 'block' : 'none';
      });
      updateGlobalScrollerWidth();
      applySearchFilter();
    }

    // Show a small tip
    statusEl.textContent = 'Tip: Click My Team or Taken button to remove player from board. Scroll horizontally for more data.';

    // Ensure and initialize bottom-fixed scrollbar
    ensureGlobalScroller();
    // Defer to next frame so layout is ready
    requestAnimationFrame(updateGlobalScrollerWidth);
    rebuildMyTeamUI();
    renderRoster();

    // boom toggle and search box handled above
  }

  function applySearchFilter() {
    const q = (playerSearchInput && playerSearchInput.value || '').trim().toLowerCase();
    const active = getActiveWrapper();
    if (!active) return;
    const rows = active.querySelectorAll('tbody tr');
    rows.forEach((tr) => {
      const nameCell = tr.querySelector('td'); // first cell may be My Team checkbox, so better:
    });
    // Use dataset set earlier for player name; fallback to first text match in row
    rows.forEach((tr) => {
      if (!q) { tr.style.display = (tr.dataset.player && (isInMyTeam(tr.dataset.player) || isTaken(tr.dataset.player))) ? 'none' : ''; return; }
      const playerName = (tr.dataset.player || '').toLowerCase();
      const rowText = tr.textContent.toLowerCase();
      const match = playerName.includes(q) || rowText.includes(q);
      // Maintain hide if drafted/taken
      const hiddenByDraft = tr.dataset.player && (isInMyTeam(tr.dataset.player) || isTaken(tr.dataset.player));
      tr.style.display = (match && !hiddenByDraft) ? '' : 'none';
    });
  }

  function sortTableByColumn(table, columnIndex) {
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const ths = [...thead.querySelectorAll('th')];

    // Determine current sort state
    const current = ths[columnIndex];
    const isAsc = current.classList.contains('th-sort-asc');
    const newAsc = !isAsc;

    // Clear all sort indicators
    ths.forEach((th) => th.classList.remove('th-sort-asc', 'th-sort-desc'));
    current.classList.add(newAsc ? 'th-sort-asc' : 'th-sort-desc');

    // Helpers for robust sorting
    const normalizeText = (s) => (s || '').trim();
    const isEmptyCell = (s) => {
      const t = normalizeText(s);
      return t === '' || t === '-';
    };
    const parseNumber = (s) => {
      let t = normalizeText(s);
      let isParenNeg = false;
      if (/^\(.*\)$/.test(t)) { isParenNeg = true; t = t.slice(1, -1); }
      t = t.replace(/[,$%\s]/g, '').replace(/^\$/g, '');
      const n = Number(t);
      if (Number.isNaN(n)) return null;
      return isParenNeg ? -n : n;
    };

    // Collect rows and sort
    const rows = Array.from(tbody.querySelectorAll('tr'));
    rows.sort((a, b) => {
      const aText = normalizeText(a.children[columnIndex]?.textContent || '');
      const bText = normalizeText(b.children[columnIndex]?.textContent || '');

      const aEmpty = isEmptyCell(aText);
      const bEmpty = isEmptyCell(bText);
      if (aEmpty !== bEmpty) {
        // Empties always last
        return aEmpty ? 1 : -1;
      }

      const aNum = parseNumber(aText);
      const bNum = parseNumber(bText);
      const aIsNum = aNum !== null;
      const bIsNum = bNum !== null;

      // Numbers always before text
      if (aIsNum !== bIsNum) {
        return aIsNum ? -1 : 1;
      }

      if (aIsNum && bIsNum) {
        return newAsc ? aNum - bNum : bNum - aNum;
      }

      // Case-insensitive text compare
      return newAsc
        ? aText.localeCompare(bText, undefined, { sensitivity: 'base' })
        : bText.localeCompare(aText, undefined, { sensitivity: 'base' });
    });

    // Rebuild tbody
    tbody.innerHTML = '';
    rows.forEach((tr) => tbody.appendChild(tr));
  }

  async function init() {
    try {
      const apiPath = (window && window.location && window.location.hostname.includes('vercel.app')) ? '/api/workbook' : '/api/workbook';
      const res = await fetch(`${apiPath}?cb=${Date.now()}`, { cache: 'no-store' });
      const data = await res.json();
      if (!res.ok) {
        throw new Error(data?.error || 'Failed to load workbook');
      }
      renderWorkbook(data);
    } catch (err) {
      statusEl.textContent = (err && err.message) || 'Failed to load workbook';
    }
  }

  init();
  if (refreshBtn) {
    refreshBtn.addEventListener('click', () => {
      // Clear My Team and Taken selections, unhide rows, keep roster as-is
      statusEl.textContent = 'Clearing selections…';
      const names = [];
      for (let i = 0; i < localStorage.length; i += 1) {
        const k = localStorage.key(i);
        if (!k) continue;
        if (k.startsWith('myteam:')) names.push(k.slice('myteam:'.length));
        if (k.startsWith('taken:')) names.push(k.slice('taken:'.length));
      }
      // Remove selection flags
      names.forEach((n) => {
        setMyTeam(n, false);
        setTaken(n, false);
        updateAllCheckboxesForPlayer(n);
        updateRowVisibilityForPlayer(n);
      });
      // Clear slot assignments but keep roster configuration
      saveAssignments({});
      renderRoster();
      rebuildMyTeamUI();
      applySearchFilter();
      statusEl.textContent = 'Selections cleared.';
    });
  }
})();


