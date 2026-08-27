/* global fetch, document, DraftRoom */
(function () {
  const CURRENT_SEASON = '2026';
  const STORAGE_SEASON_KEY = 'draftWizard:season';
  const COLUMN_WIDTHS_KEY = 'draftWizard:columnWidths:2026:v1';
  const SIDEBAR_WIDTH_KEY = 'draftWizard:sidebarWidth:2026:v1';
  const LEAGUE_VIEW_KEY = 'draftWizard:leaguePaneView:2026:v1';
  const EDGE_DATA_URL = '/data/adp-edge-2026.json';
  const VARIANT_BY_SHEET = { STRD: 'standard', '.5 PPR NEW': 'halfPpr', 'FULL PPR NEW': 'ppr', 'SUPERFLEX NEW': 'superflex' };
  // Brightened official primary/identity colors preserve team recognition and
  // legibility on the dark draft board. Aliases share their franchise color.
  const TEAM_COLORS = Object.freeze({
    ARI: '#97233f', AZ: '#97233f', ATL: '#c21f3a', BAL: '#7b61b5', BUF: '#2f7fe5',
    CAR: '#00a5e6', CHI: '#d7511f', CIN: '#fb4f14', CLE: '#ff5a1f', DAL: '#869397',
    DEN: '#ff7a21', DET: '#1995d3', GB: '#ffb612', HOU: '#e52b38', IND: '#4f8ac9',
    JAC: '#00a5b5', JAX: '#00a5b5', KC: '#e31837', LAC: '#38a6e8', LAR: '#ffd100', LV: '#a5acaf',
    MIA: '#00a0a8', MIN: '#ffc62f', NE: '#d7263d', NO: '#d3bc8d', NYG: '#2878bd',
    NYJ: '#3f9c35', PHI: '#2a9d8f', PIT: '#f7c33d', SEA: '#69be28', SF: '#c9243f',
    TB: '#e02b31', TEN: '#5aa5e6', WAS: '#ffcf40', FA: '#94a3b8',
  });
  const PRIOR_SEASON = String(Number(CURRENT_SEASON) - 1);
  let edgeData = null;
  let edgeByIdentity = new Map();
  let pointsByIdentity = new Map();
  let playerCatalog = [];
  let playerByKey = new Map();
  let boomEnabled = false;
  let edgeEnabled = false;
  let activeVariant = 'standard';
  let activeProfile = null;
  let draftSession = null;
  let selectedPlayerKey = null;
  let editingProfileId = null;
  let importPreview = [];
  let toastTimer = null;
  let selectSheetByVariant = null;
  let leaguePaneView = localStorage.getItem(LEAGUE_VIEW_KEY) === 'team' ? 'team' : 'league';
  let columnWidths = {};
  try { columnWidths = JSON.parse(localStorage.getItem(COLUMN_WIDTHS_KEY) || '{}') || {}; } catch { columnWidths = {}; }

  const legacySnapshot = DraftRoom.captureLegacySelections(localStorage, CURRENT_SEASON);
  const draftStore = new DraftRoom.DraftStore(localStorage, CURRENT_SEASON);

  function migrateSeasonStorage() {
    if (localStorage.getItem(STORAGE_SEASON_KEY) === CURRENT_SEASON) return;
    const remove = [];
    for (let i = 0; i < localStorage.length; i += 1) {
      const key = localStorage.key(i);
      if (key && (key.startsWith('myteam:') || key.startsWith('taken:') || key.startsWith('draftedAny:'))) remove.push(key);
    }
    remove.forEach((key) => localStorage.removeItem(key));
    localStorage.removeItem('assignments');
    localStorage.setItem(STORAGE_SEASON_KEY, CURRENT_SEASON);
  }
  migrateSeasonStorage();

  const statusEl = document.getElementById('status');
  const fileMetaEl = document.getElementById('file-meta');
  const titleEl = document.getElementById('page-title');
  const tabsEl = document.getElementById('tabs');
  const sheetContainerEl = document.getElementById('sheet-container');
  const paneSplitter = document.getElementById('pane-splitter');
  const myTeamEl = document.getElementById('draft-sidebar');
  const sidebarResizer = document.getElementById('sidebar-resizer');
  let playerSearchInput = null;
  const refreshBtn = document.getElementById('refresh-page');
  const edgeDialog = document.getElementById('edge-dialog');
  const edgeDialogContent = document.getElementById('edge-dialog-content');
  const leagueSetupBtn = document.getElementById('league-setup');
  const importDraftBtn = document.getElementById('import-draft');
  const undoPickBtn = document.getElementById('undo-pick');
  const clockTeamEl = document.getElementById('clock-team');
  const clockPickEl = document.getElementById('clock-pick');
  const clockOverrideEl = document.getElementById('clock-team-override');
  const nextMyPickEl = document.getElementById('next-my-pick');
  const lastDraftedPlayerEl = document.getElementById('last-drafted-player');
  const profileSelectEl = document.getElementById('league-profile-select');
  const decisionListEl = document.getElementById('decision-list');
  const riskPlayerSearchEl = document.getElementById('risk-player-search');
  const riskSearchResultsEl = document.getElementById('risk-search-results');
  const riskInfoButtonEl = document.getElementById('risk-info-button');
  const riskInfoDialogEl = document.getElementById('risk-info-dialog');
  const leagueGridEl = document.getElementById('league-grid');
  const leaguePaneTitleEl = document.getElementById('league-pane-title');
  const leagueViewToggleEl = document.getElementById('league-view-toggle');
  const draftProgressEl = document.getElementById('draft-progress');
  const leagueDialog = document.getElementById('league-dialog');
  const leagueForm = document.getElementById('league-form');
  const importDialog = document.getElementById('import-dialog');
  const importForm = document.getElementById('import-form');
  const teamDialog = document.getElementById('team-dialog');
  const pickDialog = document.getElementById('pick-dialog');
  const draftToast = document.getElementById('draft-toast');

  const normalizeName = (value) => String(value || '').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\b(jr|sr|ii|iii|iv)\b/g, '').replace(/[^a-z0-9]/g, '');
  const edgeIdentity = (name, position) => `${normalizeName(name)}|${String(position || '').toUpperCase()}`;
  function formatFantasyPoints(value) {
    if (value === null || value === undefined || value === '') return '—';
    const numeric = Number(value);
    return Number.isFinite(numeric) ? String(Math.round(numeric * 10) / 10) : '—';
  }
  function positiveRank(value) {
    if (value === null || value === undefined || String(value).trim() === '') return null;
    const numeric = Number(value);
    return Number.isFinite(numeric) && numeric > 0 ? numeric : null;
  }

  // Current tiers are rebuilt independently for every scoring format. Each
  // group is centered on a useful draft-board size, then nudged to the
  // strongest nearby ADP gap so a tier represents a real market drop-off
  // instead of a stale workbook row boundary.
  function currentTierProfile(rank) {
    if (rank <= 24) return { min: 4, target: 5, max: 7, gapFloor: 0.75 };
    if (rank <= 72) return { min: 5, target: 7, max: 10, gapFloor: 1 };
    if (rank <= 144) return { min: 7, target: 10, max: 14, gapFloor: 1.5 };
    return { min: 9, target: 14, max: 18, gapFloor: 2.5 };
  }

  function assignCurrentFormatTiers(table, rankColumnIndex, variant) {
    if (!table || rankColumnIndex < 0) return;
    const rankedRows = [...table.querySelectorAll('tbody tr')].map((row) => {
      row.dataset.tier = '';
      row.removeAttribute('data-tier-model');
      const cell = row.children[rankColumnIndex];
      return {
        row,
        rank: cell?.dataset.sortMissing === 'true'
          ? null
          : positiveRank(cell?.dataset.sortValue || cell?.textContent),
      };
    }).filter((item) => item.rank !== null).sort((a, b) => a.rank - b.rank);

    let start = 0;
    let tierNumber = 1;
    while (start < rankedRows.length) {
      const remaining = rankedRows.length - start;
      const profile = currentTierProfile(rankedRows[start].rank);
      let groupSize = Math.min(profile.target, remaining);

      if (remaining > profile.max) {
        let best = null;
        const smallest = Math.min(profile.min, remaining);
        const largest = Math.min(profile.max, remaining - 1);
        for (let size = smallest; size <= largest; size += 1) {
          const current = rankedRows[start + size - 1];
          const next = rankedRows[start + size];
          const gap = next.rank - current.rank;
          if (gap < profile.gapFloor) continue;
          const distancePenalty = 1 + (Math.abs(size - profile.target) * 0.18);
          const score = gap / distancePenalty;
          if (!best || score > best.score) best = { size, score };
        }
        groupSize = best?.size || profile.target;
      } else {
        groupSize = remaining;
      }

      // Never split players who share the same market rank.
      while (
        start + groupSize < rankedRows.length
        && rankedRows[start + groupSize - 1].rank === rankedRows[start + groupSize].rank
      ) groupSize += 1;

      const tier = `Tier ${tierNumber}`;
      rankedRows.slice(start, start + groupSize).forEach(({ row }) => {
        row.dataset.tier = tier;
        row.dataset.tierModel = `${variant}:current-rank-gap-v1`;
        const player = playerByKey.get(row.dataset.playerKey);
        if (player?.variants?.[variant]) player.variants[variant].workbookTier = tier;
      });
      start += groupSize;
      tierNumber += 1;
    }
  }

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

  function setSidebarWidth(value, persist = true) {
    const max = Math.max(320, Math.min(720, window.innerWidth * 0.58));
    const width = Math.round(Math.max(280, Math.min(max, Number(value) || 360)));
    document.documentElement.style.setProperty('--sidebar-width', `${width}px`);
    if (persist) localStorage.setItem(SIDEBAR_WIDTH_KEY, String(width));
    requestAnimationFrame(updateGlobalScrollerWidth);
    return width;
  }

  const savedSidebarWidth = Number(localStorage.getItem(SIDEBAR_WIDTH_KEY));
  if (Number.isFinite(savedSidebarWidth) && savedSidebarWidth > 0) setSidebarWidth(savedSidebarWidth, false);

  if (sidebarResizer && myTeamEl) {
    let startX = 0;
    let startWidth = 0;
    const stop = () => {
      document.body.classList.remove('is-sidebar-resizing');
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', stop);
      window.removeEventListener('pointercancel', stop);
    };
    const move = (event) => {
      if (!document.body.classList.contains('is-sidebar-resizing')) return;
      setSidebarWidth(startWidth - (event.clientX - startX));
      event.preventDefault();
    };
    sidebarResizer.addEventListener('pointerdown', (event) => {
      startX = event.clientX;
      startWidth = myTeamEl.getBoundingClientRect().width;
      document.body.classList.add('is-sidebar-resizing');
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', stop);
      window.addEventListener('pointercancel', stop);
      event.preventDefault();
    });
    sidebarResizer.addEventListener('keydown', (event) => {
      const current = myTeamEl.getBoundingClientRect().width;
      if (event.key === 'ArrowLeft') { setSidebarWidth(current + 16); event.preventDefault(); }
      if (event.key === 'ArrowRight') { setSidebarWidth(current - 16); event.preventDefault(); }
    });
  }

  function columnStorageKey(sheetName, column) {
    return `${sheetName}|${column.kind === 'sheet' ? `sheet:${column.label}` : column.kind}`;
  }

  function applyColumnWidth(table, columnIndex, width) {
    const safeWidth = Math.round(Math.max(48, Math.min(640, Number(width) || 0)));
    if (!safeWidth) return;
    table.querySelectorAll('tr').forEach((row) => {
      const cell = row.children[columnIndex];
      if (!cell) return;
      cell.style.width = `${safeWidth}px`;
      cell.style.minWidth = `${safeWidth}px`;
      cell.style.maxWidth = `${safeWidth}px`;
      cell.classList.add('user-sized-column');
    });
  }

  function attachColumnResizer(th, table, columnIndex, storageKey) {
    const handle = document.createElement('span');
    handle.className = 'column-resizer';
    handle.setAttribute('role', 'separator');
    handle.setAttribute('aria-orientation', 'vertical');
    handle.setAttribute('aria-label', `Resize ${th.textContent || 'column'}`);
    handle.tabIndex = 0;
    const saved = Number(columnWidths[storageKey]);
    if (Number.isFinite(saved) && saved > 0) applyColumnWidth(table, columnIndex, saved);

    let startX = 0;
    let startWidth = 0;
    let currentWidth = 0;
    const stop = () => {
      if (!document.body.classList.contains('is-column-resizing')) return;
      document.body.classList.remove('is-column-resizing');
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', stop);
      window.removeEventListener('pointercancel', stop);
      if (currentWidth) {
        columnWidths[storageKey] = currentWidth;
        localStorage.setItem(COLUMN_WIDTHS_KEY, JSON.stringify(columnWidths));
      }
    };
    const move = (event) => {
      currentWidth = Math.round(Math.max(48, Math.min(640, startWidth + event.clientX - startX)));
      applyColumnWidth(table, columnIndex, currentWidth);
      updateGlobalScrollerWidth();
      event.preventDefault();
    };
    handle.addEventListener('pointerdown', (event) => {
      startX = event.clientX;
      startWidth = th.getBoundingClientRect().width;
      currentWidth = startWidth;
      document.body.classList.add('is-column-resizing');
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', stop);
      window.addEventListener('pointercancel', stop);
      event.stopPropagation();
      event.preventDefault();
    });
    handle.addEventListener('click', (event) => event.stopPropagation());
    handle.addEventListener('dblclick', (event) => {
      event.stopPropagation();
      delete columnWidths[storageKey];
      localStorage.setItem(COLUMN_WIDTHS_KEY, JSON.stringify(columnWidths));
      table.querySelectorAll('tr').forEach((row) => {
        const cell = row.children[columnIndex];
        if (!cell) return;
        cell.style.removeProperty('width');
        cell.style.removeProperty('min-width');
        cell.style.removeProperty('max-width');
        cell.classList.remove('user-sized-column');
      });
      updateGlobalScrollerWidth();
    });
    handle.addEventListener('keydown', (event) => {
      if (!['ArrowLeft', 'ArrowRight'].includes(event.key)) return;
      const width = th.getBoundingClientRect().width + (event.key === 'ArrowRight' ? 12 : -12);
      applyColumnWidth(table, columnIndex, width);
      columnWidths[storageKey] = Math.max(48, Math.min(640, Math.round(width)));
      localStorage.setItem(COLUMN_WIDTHS_KEY, JSON.stringify(columnWidths));
      updateGlobalScrollerWidth();
      event.preventDefault();
      event.stopPropagation();
    });
    th.appendChild(handle);
  }

  const isPlayerKeyDrafted = (key) => Boolean(key && (draftSession?.picks || []).some((pick) => pick.playerKey === key));

  // Preserve the user's old roster template as the starting point for the
  // first league profile; all active draft allocation lives in DraftRoom.
  const DEFAULT_ROSTER = DraftRoom.DEFAULT_ROSTER;
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

  function ensureDraftState() {
    let created = false;
    activeProfile = draftStore.activeProfile();
    if (!activeProfile) {
      activeProfile = draftStore.saveProfile(DraftRoom.makeProfile({
        season: Number(CURRENT_SEASON),
        name: 'My League',
        scoring: 'ppr',
        teamCount: 12,
        userSlot: 1,
        rosterTemplate: loadRoster(),
      }));
      created = true;
    }
    draftSession = draftStore.ensureSession(activeProfile);
    if (!Array.isArray(draftSession.watchlist)) draftSession.watchlist = [];
    return created;
  }

  function saveActiveSession(session) {
    if (!Array.isArray(session.watchlist)) session.watchlist = [];
    draftSession = draftStore.saveSession(session);
  }

  function showDialog(dialog) {
    if (!dialog) return;
    if (typeof dialog.showModal === 'function') dialog.showModal();
    else dialog.setAttribute('open', '');
  }

  function closeDialog(dialog) {
    if (!dialog) return;
    if (typeof dialog.close === 'function') dialog.close();
    else dialog.removeAttribute('open');
  }

  function showToast(message, allowUndo = false) {
    if (!draftToast) return;
    clearTimeout(toastTimer);
    document.getElementById('draft-toast-text').textContent = message;
    document.getElementById('toast-undo').hidden = !allowUndo;
    draftToast.hidden = false;
    toastTimer = setTimeout(() => { draftToast.hidden = true; }, 7000);
  }

  function currentScheduledTeam() {
    return DraftRoom.currentTeam(activeProfile, draftSession);
  }

  function nextMyOverallPick() {
    if (!activeProfile || !draftSession) return null;
    const current = DraftRoom.teamForPick(draftSession.nextOverall, activeProfile.teamCount);
    const currentTeamId = draftSession.overrideTeamId || current.teamId;
    const includeCurrent = currentTeamId !== activeProfile.userTeamId;
    return DraftRoom.nextPickForTeam(
      draftSession.nextOverall,
      activeProfile.userTeamId,
      activeProfile.teamCount,
      activeProfile.teamCount * activeProfile.rounds,
      includeCurrent,
    );
  }

  function playerForKey(key) { return playerByKey.get(key) || null; }

  function playerView(player, variant = activeVariant) {
    const value = player?.variants?.[variant] || {};
    return {
      playerKey: player.playerKey,
      id: player.id,
      name: player.name,
      position: player.position,
      team: player.team || null,
      marketRank: value.marketRank,
      boardRank: value.boardRank,
      edgeScore: value.edgeScore,
      edgeTier: value.edgeTier,
      tier: value.workbookTier,
      projectedPoints: value.projectedPoints,
    };
  }

  function activeIntel() {
    if (!activeProfile || !draftSession) return [];
    return DraftRoom.buildReturnIntel(
      playerCatalog.map((player) => playerView(player)).filter((player) => player.playerKey),
      activeProfile,
      draftSession,
      edgeData?.draftAvailability,
    );
  }

  function selectPlayer(key, scroll = false) {
    selectedPlayerKey = key;
    document.querySelectorAll('tr.row-selected').forEach((row) => row.classList.remove('row-selected'));
    const active = getActiveWrapper();
    if (!active || !key) return;
    const row = [...active.querySelectorAll('tbody tr')].find((candidate) => candidate.dataset.playerKey === key && candidate.style.display !== 'none');
    if (!row) return;
    row.classList.add('row-selected');
    if (scroll) row.scrollIntoView({ block: 'nearest' });
  }

  function applyPlayerHighlightState() {
    const watched = new Set(draftSession?.watchlist || []);
    document.querySelectorAll('tbody tr[data-player-key]').forEach((row) => {
      const key = row.dataset.playerKey;
      const isWatched = watched.has(key);
      row.classList.toggle('row-watched', isWatched);
      row.classList.toggle('row-selected', Boolean(selectedPlayerKey && key === selectedPlayerKey));
      const nameCell = row.querySelector('.player-name-cell');
      if (!nameCell || !['QB', 'RB', 'WR', 'TE'].includes(row.dataset.position)) return;
      nameCell.title = `Double-click to ${isWatched ? 'remove from' : 'add to'} Will He Make It Back?`;
      const watchButton = row.querySelector('.watch-player-btn');
      if (watchButton) {
        watchButton.classList.toggle('is-watched', isWatched);
        watchButton.setAttribute('aria-pressed', String(isWatched));
        watchButton.setAttribute('aria-label', `${isWatched ? 'Remove' : 'Add'} ${row.dataset.player} ${isWatched ? 'from' : 'to'} Will He Make It Back?`);
        watchButton.title = `${isWatched ? 'Remove from' : 'Add to'} Will He Make It Back?`;
        watchButton.textContent = isWatched ? '✓ Watch' : 'Watch';
      }
    });
  }

  function focusNextAvailableRow(referenceRow) {
    const active = getActiveWrapper();
    if (!active) return;
    const rows = [...active.querySelectorAll('tbody tr')];
    const start = Math.max(0, referenceRow ? rows.indexOf(referenceRow) : 0);
    const next = rows.slice(start + 1).find((row) => row.style.display !== 'none') || rows.find((row) => row.style.display !== 'none');
    selectPlayer(next?.dataset.playerKey || null, true);
  }

  function draftPlayer(targetPlayerKey, source = 'manual', teamId = null) {
    const player = playerForKey(targetPlayerKey);
    if (!player || !activeProfile || !draftSession) return;
    const referenceRow = [...(getActiveWrapper()?.querySelectorAll('tbody tr') || [])].find((row) => row.dataset.playerKey === targetPlayerKey);
    const result = DraftRoom.recordPick(activeProfile, draftSession, player, { source, teamId });
    if (result.error) { showToast(result.error); return; }
    saveActiveSession(result.session);
    const team = activeProfile.teams.find((candidate) => candidate.id === result.pick.teamId);
    refreshDraftRoomUI();
    focusNextAvailableRow(referenceRow);
    showToast(`${result.pick.name} → ${team?.name || result.pick.teamId}`, true);
  }

  function undoPick() {
    const result = DraftRoom.undoLastPick(draftSession);
    if (!result.pick) { showToast('There is no recorded pick to undo.'); return; }
    saveActiveSession(result.session);
    refreshDraftRoomUI();
    selectPlayer(result.pick.playerKey, true);
    showToast(`Undid ${result.pick.name}.`);
  }

  function updateClock() {
    if (!activeProfile || !draftSession) return;
    const coordinates = DraftRoom.teamForPick(draftSession.nextOverall, activeProfile.teamCount);
    const team = currentScheduledTeam();
    clockTeamEl.textContent = team?.name || 'Draft complete';
    clockTeamEl.classList.toggle('is-mine', team?.id === activeProfile.userTeamId);
    clockPickEl.textContent = `Round ${coordinates.round} · Pick ${coordinates.roundPick} · Overall ${draftSession.nextOverall}`;
    const nextMine = nextMyOverallPick();
    if (nextMine) {
      const mine = DraftRoom.teamForPick(nextMine, activeProfile.teamCount);
      nextMyPickEl.textContent = `${mine.round}.${mine.roundPick} (#${nextMine})`;
    } else nextMyPickEl.textContent = 'Complete';
    const lastDrafted = DraftRoom.lastDraftedPick(draftSession);
    lastDraftedPlayerEl.textContent = lastDrafted?.name || '—';
    lastDraftedPlayerEl.title = lastDrafted
      ? `${lastDrafted.name} · Pick #${lastDrafted.overall}`
      : 'No players drafted yet';
    undoPickBtn.disabled = !draftSession.picks.length;

    const previous = clockOverrideEl.value;
    clockOverrideEl.innerHTML = '';
    const scheduledOption = document.createElement('option');
    scheduledOption.value = '';
    scheduledOption.textContent = `Scheduled · ${activeProfile.teams[coordinates.slot - 1]?.name || coordinates.teamId}`;
    clockOverrideEl.appendChild(scheduledOption);
    activeProfile.teams.forEach((candidate) => {
      const option = document.createElement('option');
      option.value = candidate.id;
      option.textContent = `${candidate.slot}. ${candidate.name}`;
      clockOverrideEl.appendChild(option);
    });
    clockOverrideEl.value = draftSession.overrideTeamId || (previous && activeProfile.teams.some((candidate) => candidate.id === previous) ? previous : '');
  }

  function renderProfileSelect() {
    if (!profileSelectEl) return;
    profileSelectEl.innerHTML = '';
    draftStore.profiles().forEach((profile) => {
      const option = document.createElement('option');
      option.value = profile.id;
      option.textContent = profile.name;
      profileSelectEl.appendChild(option);
    });
    const add = document.createElement('option');
    add.value = '__new__';
    add.textContent = '+ New league';
    profileSelectEl.appendChild(add);
    profileSelectEl.value = activeProfile?.id || '__new__';
  }

  function renderDecisionPanel(intel) {
    if (!decisionListEl) return;
    decisionListEl.innerHTML = '';
    const byKey = new Map(intel.map((item) => [item.playerKey, item]));
    const watched = DraftRoom.sortByDisplayRank(
      (draftSession.watchlist || []).map((key) => byKey.get(key)).filter(Boolean),
    );
    watched.forEach((item) => {
      const card = document.createElement('div');
      card.className = 'decision-card-wrap';
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `decision-card decision-${item.label}`;
      button.dataset.playerKey = item.playerKey;
      const top = document.createElement('span');
      top.className = 'decision-topline';
      const name = document.createElement('strong');
      name.textContent = item.name;
      const badge = document.createElement('span');
      badge.className = `decision-label ${item.label}`;
      badge.textContent = item.label === 'take-now' ? 'TAKE NOW' : item.label === 'can-wait' ? 'CAN WAIT' : 'PIVOT';
      top.append(name, badge);
      const meta = document.createElement('span');
      meta.className = 'decision-meta';
      meta.textContent = `${item.position} · Return risk ${item.risk}%`;
      const why = document.createElement('span');
      why.className = 'decision-why';
      why.textContent = item.reasons[0];
      button.append(top, meta, why);
      button.title = item.reasons.join(' ');
      button.addEventListener('click', () => selectPlayer(item.playerKey, true));
      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'decision-remove';
      remove.textContent = '×';
      remove.title = `Stop tracking ${item.name}`;
      remove.setAttribute('aria-label', `Stop tracking ${item.name}`);
      remove.addEventListener('click', () => removeWatchPlayer(item.playerKey));
      card.append(button, remove);
      decisionListEl.appendChild(card);
    });
  }

  function availableWatchMatches(query) {
    const normalized = normalizeName(query);
    if (!normalized) return [];
    const watched = new Set(draftSession?.watchlist || []);
    return playerCatalog
      .filter((player) => ['QB', 'RB', 'WR', 'TE'].includes(player.position))
      .filter((player) => player.variants?.[activeVariant] && !watched.has(player.playerKey) && !isPlayerKeyDrafted(player.playerKey))
      .filter((player) => normalizeName(player.name).includes(normalized))
      .sort((a, b) => {
        const aName = normalizeName(a.name);
        const bName = normalizeName(b.name);
        const prefix = Number(bName.startsWith(normalized)) - Number(aName.startsWith(normalized));
        if (prefix) return prefix;
        return (a.variants[activeVariant]?.marketRank || 9999) - (b.variants[activeVariant]?.marketRank || 9999);
      })
      .slice(0, 8);
  }

  function renderRiskSearchResults() {
    if (!riskPlayerSearchEl || !riskSearchResultsEl) return;
    const matches = availableWatchMatches(riskPlayerSearchEl.value);
    riskSearchResultsEl.innerHTML = '';
    matches.forEach((player) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'risk-search-result';
      button.setAttribute('role', 'option');
      button.dataset.playerKey = player.playerKey;
      button.textContent = `${player.name} · ${player.position}${player.team ? ` · ${player.team}` : ''}`;
      button.addEventListener('click', () => addWatchPlayer(player.playerKey));
      riskSearchResultsEl.appendChild(button);
    });
    riskSearchResultsEl.hidden = !matches.length;
  }

  function addWatchPlayer(playerKey) {
    if (!draftSession || !playerForKey(playerKey) || isPlayerKeyDrafted(playerKey)) return;
    const keys = Array.isArray(draftSession.watchlist) ? draftSession.watchlist : [];
    if (!keys.includes(playerKey)) {
      draftSession.watchlist = [...keys, playerKey];
      saveActiveSession(draftSession);
    }
    riskPlayerSearchEl.value = '';
    renderRiskSearchResults();
    refreshDraftRoomUI();
  }

  function removeWatchPlayer(playerKey) {
    if (!draftSession) return;
    if (selectedPlayerKey === playerKey) selectedPlayerKey = null;
    draftSession.watchlist = (draftSession.watchlist || []).filter((key) => key !== playerKey);
    saveActiveSession(draftSession);
    refreshDraftRoomUI();
  }

  function toggleWatchPlayer(playerKey) {
    const player = playerForKey(playerKey);
    if (!player || !['QB', 'RB', 'WR', 'TE'].includes(player.position) || isPlayerKeyDrafted(playerKey)) return;
    if ((draftSession?.watchlist || []).includes(playerKey)) removeWatchPlayer(playerKey);
    else addWatchPlayer(playerKey);
  }

  function renderLeagueGrid() {
    if (!leagueGridEl || !activeProfile || !draftSession) return;
    leagueGridEl.innerHTML = '';
    const current = currentScheduledTeam();
    const total = activeProfile.teamCount * activeProfile.rounds;
    activeProfile.teams.forEach((team) => {
      const summary = DraftRoom.teamSummary(activeProfile, draftSession, team.id);
      const nextPick = DraftRoom.nextPickForTeam(draftSession.nextOverall, team.id, activeProfile.teamCount, total, true);
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'league-team-row';
      if (team.id === activeProfile.userTeamId) button.classList.add('is-mine');
      if (team.id === current?.id) button.classList.add('is-on-clock');
      const heading = document.createElement('span');
      heading.className = 'league-team-heading';
      const teamName = document.createElement('strong');
      teamName.textContent = `${team.slot}. ${team.name}`;
      const away = document.createElement('span');
      away.textContent = nextPick ? (nextPick === draftSession.nextOverall ? 'NOW' : `in ${nextPick - draftSession.nextOverall}`) : 'done';
      heading.append(teamName, away);
      const positions = document.createElement('span');
      positions.className = 'team-position-counts';
      DraftRoom.rosterNeeds(summary, activeProfile.rosterTemplate).forEach((slot) => {
        const count = document.createElement('span');
        count.textContent = `${slot.label} ${slot.filled}/${slot.total}`;
        count.classList.add(slot.complete ? 'need-filled' : 'need-open');
        count.title = slot.complete
          ? `${slot.type} starter slot${slot.total === 1 ? '' : 's'} filled`
          : `${slot.remaining} ${slot.type} starter slot${slot.remaining === 1 ? '' : 's'} open`;
        positions.appendChild(count);
      });
      const tendency = document.createElement('span');
      tendency.className = 'team-tendency';
      tendency.textContent = summary.tendency ? `Recent: ${summary.tendency}` : `${summary.picks.length} drafted`;
      button.append(heading, positions, tendency);
      button.addEventListener('click', () => openTeamDialog(team.id));
      leagueGridEl.appendChild(button);
    });
    draftProgressEl.textContent = `${draftSession.picks.length}/${total} picks`;
    draftProgressEl.title = draftSession.picks.slice(-8).map((pick) => `#${pick.overall} ${pick.name} → ${activeProfile.teams.find((team) => team.id === pick.teamId)?.name || pick.teamId}`).join('\n');
  }

  function renderMyTeamSidebar() {
    if (!leagueGridEl || !activeProfile || !draftSession) return;
    leagueGridEl.innerHTML = '';
    leagueGridEl.classList.add('my-team-roster-view');
    const summary = DraftRoom.teamSummary(activeProfile, draftSession, activeProfile.userTeamId);
    activeProfile.rosterTemplate.forEach((slot) => {
      if (!slot.count) return;
      const row = document.createElement('div');
      row.className = 'team-roster-row sidebar-roster-row';
      const label = document.createElement('strong');
      label.textContent = `${slot.type} (${slot.count})`;
      const chips = document.createElement('div');
      chips.className = 'team-roster-chips';
      for (let index = 0; index < slot.count; index += 1) {
        const pick = summary.assignments[slot.type]?.[index];
        const chip = document.createElement(pick ? 'button' : 'span');
        if (pick) chip.type = 'button';
        chip.className = `team-roster-chip ${pick ? posClass(pick.position) : 'is-empty'}`;
        chip.textContent = pick ? pick.name : slot.type;
        if (pick) chip.addEventListener('click', () => openPickDialog(pick.playerKey));
        chips.appendChild(chip);
      }
      row.append(label, chips);
      leagueGridEl.appendChild(row);
    });
    if (summary.assignments.EXTRA.length) {
      const row = document.createElement('div');
      row.className = 'team-roster-row sidebar-roster-row';
      const label = document.createElement('strong');
      label.textContent = 'Extra';
      const chips = document.createElement('div');
      chips.className = 'team-roster-chips';
      summary.assignments.EXTRA.forEach((pick) => {
        const chip = document.createElement('button');
        chip.type = 'button';
        chip.className = `team-roster-chip ${posClass(pick.position)}`;
        chip.textContent = pick.name;
        chip.addEventListener('click', () => openPickDialog(pick.playerKey));
        chips.appendChild(chip);
      });
      row.append(label, chips);
      leagueGridEl.appendChild(row);
    }
    draftProgressEl.textContent = `${summary.picks.length} drafted`;
    draftProgressEl.title = `${activeProfile.teams.find((team) => team.id === activeProfile.userTeamId)?.name || 'My Team'} roster`;
  }

  function renderLeaguePane() {
    if (!leaguePaneTitleEl || !leagueViewToggleEl || !leagueGridEl) return;
    leagueGridEl.classList.remove('my-team-roster-view');
    if (leaguePaneView === 'team') {
      leaguePaneTitleEl.textContent = 'My Team';
      leagueViewToggleEl.textContent = 'League Needs';
      leagueViewToggleEl.setAttribute('aria-label', 'Show league needs');
      renderMyTeamSidebar();
    } else {
      leaguePaneTitleEl.textContent = 'League Needs';
      leagueViewToggleEl.textContent = 'My Team';
      leagueViewToggleEl.setAttribute('aria-label', 'Show my full roster');
      renderLeagueGrid();
    }
  }

  function openPickDialog(targetPlayerKey) {
    const pick = draftSession.picks.find((candidate) => candidate.playerKey === targetPlayerKey);
    if (!pick) return;
    pickDialog.dataset.playerKey = targetPlayerKey;
    document.getElementById('pick-dialog-player').textContent = `${pick.name} · ${pick.position} · pick ${pick.round}.${pick.roundPick}`;
    const select = document.getElementById('pick-team-select');
    select.innerHTML = '';
    activeProfile.teams.forEach((team) => {
      const option = document.createElement('option');
      option.value = team.id;
      option.textContent = `${team.slot}. ${team.name}`;
      select.appendChild(option);
    });
    select.value = pick.teamId;
    showDialog(pickDialog);
  }

  function openTeamDialog(teamId) {
    const team = activeProfile.teams.find((candidate) => candidate.id === teamId);
    if (!team) return;
    document.getElementById('team-dialog-title').textContent = `${team.name} Roster`;
    const content = document.getElementById('team-dialog-content');
    content.innerHTML = '';
    const summary = DraftRoom.teamSummary(activeProfile, draftSession, teamId);
    activeProfile.rosterTemplate.forEach((slot) => {
      if (!slot.count) return;
      const row = document.createElement('div');
      row.className = 'team-roster-row';
      const label = document.createElement('strong');
      label.textContent = `${slot.type} (${slot.count})`;
      const chips = document.createElement('div');
      chips.className = 'team-roster-chips';
      for (let index = 0; index < slot.count; index += 1) {
        const pick = summary.assignments[slot.type]?.[index];
        const chip = document.createElement(pick ? 'button' : 'span');
        if (pick) chip.type = 'button';
        chip.className = `team-roster-chip ${pick ? posClass(pick.position) : 'is-empty'}`;
        chip.textContent = pick ? pick.name : slot.type;
        if (pick) chip.addEventListener('click', () => openPickDialog(pick.playerKey));
        chips.appendChild(chip);
      }
      row.append(label, chips);
      content.appendChild(row);
    });
    if (summary.assignments.EXTRA.length) {
      const extra = document.createElement('p');
      extra.className = 'edge-limit';
      extra.textContent = `Extra selections: ${summary.assignments.EXTRA.map((pick) => pick.name).join(', ')}`;
      content.appendChild(extra);
    }
    showDialog(teamDialog);
  }

  function updateReturnRiskBadges(intel) {
    document.querySelectorAll('.return-risk-badge').forEach((badge) => badge.remove());
    const active = getActiveWrapper();
    if (!active) return;
    const watched = new Set(draftSession?.watchlist || []);
    const byKey = new Map(intel.filter((item) => watched.has(item.playerKey)).map((item) => [item.playerKey, item]));
    active.querySelectorAll('tr[data-player-key]').forEach((row) => {
      const item = byKey.get(row.dataset.playerKey);
      if (!item) return;
      const cell = row.querySelector('.player-name-cell');
      if (!cell) return;
      const badge = document.createElement('span');
      badge.className = `return-risk-badge risk-${item.label === 'take-now' ? 'high' : item.label === 'can-wait' ? 'low' : 'medium'}`;
      badge.textContent = item.label === 'take-now' ? 'NOW' : item.label === 'can-wait' ? 'WAIT' : `${item.risk}%`;
      badge.title = `Estimated return risk ${item.risk}%. ${item.reasons.join(' ')}`;
      cell.appendChild(badge);
    });
  }

  function refreshDraftRoomUI() {
    if (!activeProfile || !draftSession) return;
    updateClock();
    renderProfileSelect();
    const intel = activeIntel();
    renderDecisionPanel(intel);
    renderLeaguePane();
    document.querySelectorAll('tbody tr[data-player]').forEach((row) => {
      row.style.display = isPlayerKeyDrafted(row.dataset.playerKey) ? 'none' : '';
    });
    applySearchFilter();
    renderRiskSearchResults();
    updateReturnRiskBadges(intel);
    applyPlayerHighlightState();
    document.querySelectorAll('table').forEach((table) => {
      applyTierSeparators(table);
      applyNextPickMarker(table);
    });
  }

  function rosterFieldValues() {
    const fields = [...document.querySelectorAll('#league-roster-fields input[data-slot-type]')];
    return fields.map((input) => ({ type: input.dataset.slotType, count: Math.max(0, Number(input.value) || 0) }));
  }

  function renderRosterFields(roster) {
    const container = document.getElementById('league-roster-fields');
    container.innerHTML = '';
    const byType = Object.fromEntries((roster || DraftRoom.DEFAULT_ROSTER).map((slot) => [slot.type, slot.count]));
    for (const type of ['QB', 'RB', 'WR', 'TE', 'FLEX', 'SFLX', 'K', 'DST', 'BENCH']) {
      const label = document.createElement('label');
      label.textContent = type === 'SFLX' ? 'Superflex' : type;
      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.max = '30';
      input.value = String(byType[type] || 0);
      input.dataset.slotType = type;
      label.appendChild(input);
      container.appendChild(label);
    }
  }

  function currentTeamNameFields() {
    return [...document.querySelectorAll('#league-team-fields input')].map((input) => input.value);
  }

  function renderTeamNameFields(teamCount, names = []) {
    const container = document.getElementById('league-team-fields');
    container.innerHTML = '';
    const userSlot = Number(document.getElementById('league-user-slot').value) || 1;
    for (let slot = 1; slot <= teamCount; slot += 1) {
      const label = document.createElement('label');
      label.textContent = `Team ${slot}`;
      const input = document.createElement('input');
      input.maxLength = 30;
      input.placeholder = slot === userSlot ? 'My Team' : `Team ${slot}`;
      input.value = names[slot - 1] || '';
      label.appendChild(input);
      container.appendChild(label);
    }
  }

  function openLeagueSetup(asNew = false) {
    const profile = asNew ? DraftRoom.makeProfile({ season: Number(CURRENT_SEASON), rosterTemplate: activeProfile?.rosterTemplate || loadRoster() }) : activeProfile;
    editingProfileId = asNew ? null : profile.id;
    document.getElementById('league-name').value = asNew ? '' : profile.name;
    document.getElementById('league-scoring').value = profile.scoring;
    document.getElementById('league-team-count').value = profile.teamCount;
    document.getElementById('league-user-slot').value = Number(profile.userTeamId.replace('team-', '')) || 1;
    document.getElementById('league-rounds').value = profile.rounds;
    renderRosterFields(profile.rosterTemplate);
    renderTeamNameFields(profile.teamCount, asNew ? [] : profile.teams.map((team) => team.name));
    document.getElementById('delete-profile').hidden = asNew;
    showDialog(leagueDialog);
    document.getElementById('league-name').focus();
  }

  function addLegacyPreviewItems(items, startOverall) {
    if (!document.getElementById('include-legacy').checked) return items;
    const legacyNames = [
      ...(legacySnapshot?.mine || []).map((item) => ({ ...item, teamId: activeProfile.userTeamId })),
      ...(legacySnapshot?.taken || []).sort((a, b) => (a.timestamp || 0) - (b.timestamp || 0)).map((item) => ({ ...item, teamId: null })),
    ];
    const existingNames = new Set(items.map((item) => normalizeName(item.player?.name || item.raw)));
    for (const legacy of legacyNames) {
      if (existingNames.has(normalizeName(legacy.name))) continue;
      const matches = playerCatalog.filter((player) => normalizeName(player.name) === normalizeName(legacy.name));
      const player = matches.length === 1 ? playerView(matches[0]) : null;
      items.push({
        raw: `Legacy: ${legacy.name}`,
        overall: startOverall + items.length,
        teamId: legacy.teamId,
        player,
        playerKey: player?.playerKey || null,
        status: player && legacy.teamId ? 'ready' : player ? 'unassigned' : matches.length > 1 ? 'ambiguous' : 'unmatched',
        candidates: matches.slice(0, 5).map((candidate) => ({ playerKey: candidate.playerKey, name: candidate.name, position: candidate.position })),
        legacy: true,
      });
    }
    return items;
  }

  function updateImportApplyState() {
    const ready = importPreview.filter((item) => item.status === 'ready' && !item.skip).length;
    const blocked = importPreview.some((item) => item.status !== 'ready' && !item.skip);
    document.getElementById('apply-import').disabled = ready === 0 || blocked;
    const counts = importPreview.reduce((acc, item) => { acc[item.status] = (acc[item.status] || 0) + 1; return acc; }, {});
    document.getElementById('import-summary').textContent = importPreview.length
      ? `${ready} ready · ${(counts.ambiguous || 0) + (counts.unmatched || 0) + (counts.unassigned || 0)} need attention · ${counts.duplicate || 0} duplicate`
      : 'Paste picks, then choose Preview.';
  }

  function renderImportPreview() {
    const container = document.getElementById('import-preview');
    container.innerHTML = '';
    importPreview.forEach((item, index) => {
      const row = document.createElement('div');
      row.className = `import-row status-${item.status}`;
      const pick = document.createElement('span');
      pick.className = 'import-pick';
      pick.textContent = `#${item.overall}`;
      const playerCell = document.createElement('div');
      playerCell.className = 'import-player';
      if (item.status === 'ambiguous' && item.candidates.length) {
        const select = document.createElement('select');
        const prompt = document.createElement('option');
        prompt.value = '';
        prompt.textContent = 'Choose player…';
        select.appendChild(prompt);
        item.candidates.forEach((candidate) => {
          const option = document.createElement('option');
          option.value = candidate.playerKey;
          option.textContent = `${candidate.name} · ${candidate.position}`;
          select.appendChild(option);
        });
        select.addEventListener('change', () => {
          const player = playerForKey(select.value);
          if (player) {
            item.player = playerView(player);
            item.playerKey = player.playerKey;
            item.status = item.teamId ? 'ready' : 'unassigned';
            renderImportPreview();
          }
        });
        playerCell.appendChild(select);
      } else {
        const name = document.createElement('strong');
        name.textContent = item.player ? `${item.player.name} · ${item.player.position}` : item.raw;
        const status = document.createElement('span');
        status.textContent = item.status;
        playerCell.append(name, status);
      }
      const teamSelect = document.createElement('select');
      const unknown = document.createElement('option');
      unknown.value = '';
      unknown.textContent = 'Choose team…';
      teamSelect.appendChild(unknown);
      activeProfile.teams.forEach((team) => {
        const option = document.createElement('option');
        option.value = team.id;
        option.textContent = `${team.slot}. ${team.name}`;
        teamSelect.appendChild(option);
      });
      teamSelect.value = item.teamId || '';
      teamSelect.addEventListener('change', () => {
        item.teamId = teamSelect.value || null;
        if (item.player && !['duplicate', 'unmatched'].includes(item.status)) item.status = item.teamId ? 'ready' : 'unassigned';
        renderImportPreview();
      });
      const skipLabel = document.createElement('label');
      skipLabel.className = 'import-skip';
      const skip = document.createElement('input');
      skip.type = 'checkbox';
      skip.checked = Boolean(item.skip);
      skip.addEventListener('change', () => { item.skip = skip.checked; updateImportApplyState(); });
      skipLabel.append(skip, document.createTextNode('Skip'));
      row.append(pick, playerCell, teamSelect, skipLabel);
      container.appendChild(row);
      row.dataset.index = String(index);
    });
    updateImportApplyState();
  }

  function previewImport() {
    importPreview = DraftRoom.parseImportText(
      document.getElementById('import-text').value,
      playerCatalog.map((player) => playerView(player)),
      activeProfile,
      draftSession.nextOverall,
      draftSession.picks,
    );
    addLegacyPreviewItems(importPreview, draftSession.nextOverall);
    renderImportPreview();
  }

  function bindDraftRoomControls() {
    document.querySelectorAll('[data-close-dialog]').forEach((button) => {
      button.addEventListener('click', () => closeDialog(document.getElementById(button.dataset.closeDialog)));
    });
    leagueSetupBtn.addEventListener('click', () => openLeagueSetup(false));
    importDraftBtn.addEventListener('click', () => {
      importPreview = [];
      document.getElementById('import-text').value = '';
      document.getElementById('include-legacy').checked = false;
      document.getElementById('legacy-import-option').hidden = !(legacySnapshot?.mine?.length || legacySnapshot?.taken?.length);
      renderImportPreview();
      showDialog(importDialog);
    });
    undoPickBtn.addEventListener('click', undoPick);
    riskInfoButtonEl.addEventListener('click', () => showDialog(riskInfoDialogEl));
    leagueViewToggleEl.addEventListener('click', () => {
      leaguePaneView = leaguePaneView === 'team' ? 'league' : 'team';
      localStorage.setItem(LEAGUE_VIEW_KEY, leaguePaneView);
      renderLeaguePane();
    });
    riskPlayerSearchEl.addEventListener('input', renderRiskSearchResults);
    riskPlayerSearchEl.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter') return;
      const first = riskSearchResultsEl.querySelector('[data-player-key]');
      if (first) addWatchPlayer(first.dataset.playerKey);
      event.preventDefault();
    });
    riskPlayerSearchEl.addEventListener('focus', renderRiskSearchResults);
    document.addEventListener('click', (event) => {
      if (!event.target.closest('.decision-search')) riskSearchResultsEl.hidden = true;
    });
    document.getElementById('toast-undo').addEventListener('click', () => { draftToast.hidden = true; undoPick(); });
    clockOverrideEl.addEventListener('change', () => {
      draftSession.overrideTeamId = clockOverrideEl.value || null;
      saveActiveSession(draftSession);
      refreshDraftRoomUI();
    });
    profileSelectEl.addEventListener('change', () => {
      if (profileSelectEl.value === '__new__') { openLeagueSetup(true); renderProfileSelect(); return; }
      draftStore.setActiveProfile(profileSelectEl.value);
      activeProfile = draftStore.activeProfile();
      draftSession = draftStore.ensureSession(activeProfile);
      if (!Array.isArray(draftSession.watchlist)) draftSession.watchlist = [];
      if (selectSheetByVariant) selectSheetByVariant(activeProfile.scoring);
      refreshDraftRoomUI();
    });
    document.getElementById('league-team-count').addEventListener('input', () => {
      const count = Math.max(4, Math.min(20, Number(document.getElementById('league-team-count').value) || 12));
      document.getElementById('league-user-slot').max = String(count);
      renderTeamNameFields(count, currentTeamNameFields());
    });
    document.getElementById('league-user-slot').addEventListener('input', () => renderTeamNameFields(Number(document.getElementById('league-team-count').value) || 12, currentTeamNameFields()));
    leagueForm.addEventListener('submit', (event) => {
      event.preventDefault();
      const teamCount = Number(document.getElementById('league-team-count').value);
      const userSlot = Number(document.getElementById('league-user-slot').value);
      if (userSlot < 1 || userSlot > teamCount) { showToast('Your draft slot must be within the league size.'); return; }
      if (editingProfileId && draftSession.picks.length && (teamCount !== activeProfile.teamCount || userSlot !== Number(activeProfile.userTeamId.replace('team-', ''))) && !window.confirm('Changing team order will keep recorded picks but alter future snake order. Continue?')) return;
      const profile = draftStore.saveProfile({
        id: editingProfileId || undefined,
        createdAt: editingProfileId ? activeProfile.createdAt : undefined,
        name: document.getElementById('league-name').value,
        scoring: document.getElementById('league-scoring').value,
        teamCount,
        userSlot,
        rounds: Number(document.getElementById('league-rounds').value),
        rosterTemplate: rosterFieldValues(),
        teamNames: currentTeamNameFields(),
      });
      activeProfile = profile;
      draftSession = draftStore.ensureSession(profile);
      if (!Array.isArray(draftSession.watchlist)) draftSession.watchlist = [];
      closeDialog(leagueDialog);
      if (selectSheetByVariant) selectSheetByVariant(profile.scoring);
      refreshDraftRoomUI();
      showToast(`Saved ${profile.name}.`);
    });
    document.getElementById('delete-profile').addEventListener('click', () => {
      if (!editingProfileId || !window.confirm(`Delete ${activeProfile.name} and its draft session?`)) return;
      draftStore.deleteProfile(editingProfileId);
      ensureDraftState();
      closeDialog(leagueDialog);
      if (selectSheetByVariant) selectSheetByVariant(activeProfile.scoring);
      refreshDraftRoomUI();
    });
    document.getElementById('preview-import').addEventListener('click', previewImport);
    importForm.addEventListener('submit', (event) => {
      event.preventDefault();
      const blocked = importPreview.some((item) => item.status !== 'ready' && !item.skip);
      if (blocked) { showToast('Resolve or explicitly skip every highlighted import row.'); return; }
      let session = draftSession;
      let applied = 0;
      for (const item of importPreview.filter((candidate) => candidate.status === 'ready' && !candidate.skip)) {
        const result = DraftRoom.recordPick(activeProfile, session, item.player, { source: 'import', teamId: item.teamId });
        if (!result.error) { session = result.session; applied += 1; }
      }
      saveActiveSession(session);
      closeDialog(importDialog);
      refreshDraftRoomUI();
      showToast(`Imported ${applied} pick${applied === 1 ? '' : 's'}.`, applied > 0);
    });
    document.getElementById('include-legacy').addEventListener('change', () => { if (importPreview.length) previewImport(); });
    document.getElementById('pick-form').addEventListener('submit', (event) => {
      event.preventDefault();
      const result = DraftRoom.movePlayer(draftSession, pickDialog.dataset.playerKey, document.getElementById('pick-team-select').value);
      if (result.pick) saveActiveSession(result.session);
      closeDialog(pickDialog);
      closeDialog(teamDialog);
      refreshDraftRoomUI();
      if (result.pick) showToast(`Moved ${result.pick.name}.`);
    });
    document.getElementById('undraft-player').addEventListener('click', () => {
      const result = DraftRoom.undraftPlayer(draftSession, pickDialog.dataset.playerKey);
      if (result.pick) saveActiveSession(result.session);
      closeDialog(pickDialog);
      closeDialog(teamDialog);
      refreshDraftRoomUI();
      if (result.pick) { selectPlayer(result.pick.playerKey, true); showToast(`Undrafted ${result.pick.name}.`); }
    });
    document.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'z') { event.preventDefault(); undoPick(); return; }
      const target = event.target;
      const isSearch = target === playerSearchInput;
      if (target && ['TEXTAREA', 'SELECT'].includes(target.tagName)) return;
      if (target && target.tagName === 'INPUT' && !isSearch) return;
      if (event.key === 'Enter' && isSearch) {
        const first = [...(getActiveWrapper()?.querySelectorAll('tbody tr') || [])].find((row) => row.style.display !== 'none');
        if (first) { event.preventDefault(); draftPlayer(first.dataset.playerKey); }
        return;
      }
      if (event.key === 'Enter' && selectedPlayerKey && !document.querySelector('dialog[open]')) {
        event.preventDefault();
        draftPlayer(selectedPlayerKey);
      }
    });
  }

  const TIER_SEPARATOR_COLORS = ['#f472b6', '#a3e635', '#38bdf8', '#f59e0b', '#a78bfa', '#facc15', '#fb7185', '#2dd4bf'];
  function tierSeparatorColor(tier) {
    const numericTier = Number(String(tier || '').match(/\d+/)?.[0]);
    const colorIndex = Number.isFinite(numericTier) && numericTier > 1 ? numericTier - 2 : 0;
    return TIER_SEPARATOR_COLORS[colorIndex % TIER_SEPARATOR_COLORS.length];
  }
  function applyTierSeparators(table) {
    const rows = [...table.querySelectorAll('tbody tr')];
    const groups = [];

    rows.forEach((row) => {
      row.classList.remove('tier-break');
      row.style.removeProperty('--tier-separator-color');
    });

    // A tier is an ordered rank band. Suppress the lines for alternate sorts,
    // where the same tiers are necessarily interleaved and the rules would be
    // visually misleading.
    const sortedHeader = table.querySelector('th.th-sort-asc, th.th-sort-desc');
    if (!sortedHeader?.classList.contains('th-sort-asc') || sortedHeader.dataset.sortKind !== 'rank') return;

    rows.forEach((row) => {
      const tier = String(row.dataset.tier || '').trim();
      if (!tier) return;
      const currentGroup = groups[groups.length - 1];
      if (currentGroup && currentGroup.tier === tier) {
        currentGroup.rows.push(row);
      } else {
        groups.push({ tier, rows: [row] });
      }
    });

    groups.forEach((group, index) => {
      if (index === 0) return;
      const tierAboveStillAvailable = groups[index - 1].rows.some((row) => {
        return row.dataset.playerKey && !isPlayerKeyDrafted(row.dataset.playerKey);
      });
      if (!tierAboveStillAvailable) return;

      // Keep the separator visible when the first player in the lower tier is
      // taken by moving it to the first remaining row in that tier. If the
      // entire tier is taken, carry the line to the next visible lower tier.
      let separatorRow = null;
      for (let lowerIndex = index; lowerIndex < groups.length && !separatorRow; lowerIndex += 1) {
        separatorRow = groups[lowerIndex].rows.find((row) => row.style.display !== 'none') || null;
      }
      if (separatorRow) {
        separatorRow.classList.add('tier-break');
        separatorRow.style.setProperty('--tier-separator-color', tierSeparatorColor(group.tier));
      }
    });
  }

  function applyNextPickMarker(table) {
    if (!table) return;
    table.querySelectorAll('tbody tr.next-pick-line').forEach((row) => row.classList.remove('next-pick-line'));
    if (!activeProfile || !draftSession) return;
    const nextMine = nextMyOverallPick();
    const boardIndex = DraftRoom.nextPickBoardIndex(draftSession.nextOverall, nextMine);
    if (boardIndex === null || boardIndex < 0) return;
    const visible = [...table.querySelectorAll('tbody tr[data-player-key]')]
      .filter((row) => row.style.display !== 'none' && !isPlayerKeyDrafted(row.dataset.playerKey));
    const target = visible[boardIndex];
    if (target) target.classList.add('next-pick-line');
  }

  function openEdgeDetails(record, variant) {
    if (!record || !record.variants || !record.variants[variant] || !edgeDialogContent) return;
    const value = record.variants[variant];
    const sourceLinks = (value.sourceRefs || []).map((id) => edgeData.sources[id]).filter(Boolean)
      .map((source) => `<a href="${source.url}" target="_blank" rel="noopener noreferrer">${source.label}</a>`).join(' · ');
    const commonMetrics = ['currentTeam', 'leagueYears', 'line', 'lineMethod', 'olInjuryBurden', 'depthRole', 'projectedPprPoints', 'consensusProjectionRank', 'projectionVsMarketSlots', 'projectionProvider', 'injuryStatus', 'injuryDetail', 'injuryNewsUpdatedAt', 'projectionAvailabilityMultiplier'];
    const positionMetrics = {
      QB: ['games', 'passerRating', 'cpoe', 'passingTdRate', 'passingAdot', 'sackRate', 'rushYardsPerGame', 'rushTds', 'qbFloorSignal', 'incomingReceivingYards', 'projectedPassAttempts', 'projectedPassYards', 'projectedRushYards', 'projectedTotalTds'],
      RB: ['games', 'touches', 'targetShare', 'yardsPerCarry', 'rushYardsOverExpectedPerAttempt', 'twoYearReversionSignal', 'teammateCompetition', 'maxGamePprShare', 'projectedTouches', 'projectedRushYards', 'projectedTotalTds'],
      WR: ['games', 'targets', 'targetShare', 'airYardsShare', 'wopr', 'yardsPerTarget', 'receivingYardsPerOffenseSnap', 'twoYearReversionSignal', 'maxGameReceivingYardShare', 'targetVacancyPct', 'projectedQbChangeSignal', 'projectedStartingQb', 'projectedReceptions', 'projectedReceivingYards', 'projectedTotalTds'],
      TE: ['games', 'targets', 'targetShare', 'airYardsShare', 'wopr', 'yardsPerTarget', 'receivingYardsPerOffenseSnap', 'twoYearReversionSignal', 'maxGameReceivingYardShare', 'teammateCompetition', 'destinationTeTargetShare', 'projectedQbChangeSignal', 'projectedStartingQb', 'projectedReceptions', 'projectedReceivingYards', 'projectedTotalTds'],
    };
    const metricKeys = new Set([...(positionMetrics[record.position] || []), ...commonMetrics]);
    const metricLabels = {
      cpoe: 'CPOE',
      wopr: 'WOPR',
      projectedPprPoints: 'Projected PPR Points',
      projectedTotalTds: 'Projected Total TDs',
      rushYardsOverExpectedPerAttempt: 'Rush Yards Over Expected / Att',
      receivingYardsPerOffenseSnap: 'Receiving Yards / Offense Snap',
      yacOverExpected: 'YAC Over Expected',
      projectionVsMarketSlots: 'Projection vs Market Slots',
      passingTdRate: 'Pass TD Rate (%)',
      passingAdot: 'Passing aDOT',
      sackRate: 'Sack Rate (%)',
      rushYardsPerGame: 'Rush Yards / Game',
      rushTds: 'Rush TDs',
      qbFloorSignal: 'QB Floor Signal',
      maxGamePprShare: 'Largest Game Share of PPR (%)',
      maxGameReceivingYardShare: 'Largest Game Share of Rec Yards (%)',
      twoYearReversionSignal: 'Two-Year Reversion Signal',
      targetVacancyPct: 'Team Target Vacancy (%)',
      incomingReceivingYards: 'Incoming Receiver Prior Yards',
      teammateCompetition: 'Teammate Competition',
      projectedQbChangeSignal: 'Projected QB Change Signal',
      projectedStartingQb: 'Projected Starting QB',
      destinationTeTargetShare: 'Destination TE Target Share (%)',
      lineMethod: 'Current Line Method',
      olInjuryBurden: 'Current OL Injury Burden',
      injuryStatus: 'Current Injury Status',
      injuryDetail: 'Current Injury Detail',
      injuryNewsUpdatedAt: 'Injury Feed Updated',
      projectionAvailabilityMultiplier: 'Projection Availability Multiplier',
    };
    const metrics = Object.entries(value.displayedMetrics || {}).filter(([key, v]) => metricKeys.has(key) && v !== null && v !== '')
      .map(([key, val]) => `<div><dt>${metricLabels[key] || key.replace(/([A-Z])/g, ' $1')}</dt><dd>${val}</dd></div>`).join('');
    edgeDialogContent.innerHTML = `
      <h2 id="edge-dialog-title">${record.name} · ADP Edge ${value.score}</h2>
      <p class="edge-summary"><span class="edge-pill tier-${value.tier}">${value.tier === 'none' ? 'No tier' : value.tier}</span> Market pos. rank ${value.marketPositionRank ?? '—'} → model rank ${value.modelRank ?? '—'} (${value.expectedDelta >= 0 ? '+' : ''}${value.expectedDelta ?? 0})</p>
      <p>Confidence: ${record.confidence}% · ${variant} · generated ${new Date(edgeData.generatedAt).toLocaleString()}</p>
      <h3>Why</h3><ul>${(value.reasons || []).map((reason) => `<li>${reason}</li>`).join('')}</ul>
      <h3>Key inputs</h3><dl class="edge-metrics">${metrics}</dl>
      <p class="edge-sources"><strong>Sources:</strong> ${sourceLinks}</p>
      <p class="edge-limit">${edgeData.calibration?.limitation || ''}</p>`;
    if (typeof edgeDialog.showModal === 'function') edgeDialog.showModal();
    else edgeDialog.setAttribute('open', '');
  }

  /**
   * Render the workbook data into tabs and tables
   * @param {{fileName: string, sheets: {name: string, rows: any[][]}[]}} data
   */
  function renderWorkbook(data) {
    titleEl.textContent = `Fantasy Football ${data.season || CURRENT_SEASON} Draft Assistant`;
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

    const initialSheetIndex = Math.max(0, sheets.findIndex((sheet) => VARIANT_BY_SHEET[sheet.name] === activeProfile?.scoring));
    sheets.forEach((sheet, index) => {
      const tab = document.createElement('button');
      tab.className = 'tab';
      tab.setAttribute('role', 'tab');
      tab.setAttribute('aria-selected', index === initialSheetIndex ? 'true' : 'false');
      tab.textContent = sheet.name.replace(/\s+NEW$/i, '');
      tab.addEventListener('click', () => selectSheet(index));
      tabsEl.appendChild(tab);
    });

    // Independent contextual and model highlight controls.
    (function attachHighlightControls() {
      const controls = document.createElement('div');
      controls.className = 'highlight-controls';
      const boomBtn = document.createElement('button');
      boomBtn.id = 'boom-toggle';
      boomBtn.className = 'btn';
      boomBtn.setAttribute('aria-pressed', String(boomEnabled));
      boomBtn.title = 'Creator signals: Boom, Bust, and desperation opportunity';
      boomBtn.textContent = 'BOOM / BUST';
      const edgeBtn = document.createElement('button');
      edgeBtn.id = 'edge-toggle';
      edgeBtn.className = 'btn edge-toggle';
      edgeBtn.setAttribute('aria-pressed', String(edgeEnabled));
      edgeBtn.title = 'Highlight players the ADP Edge model expects to beat market position';
      edgeBtn.textContent = 'ADP EDGE';
      controls.append(boomBtn, edgeBtn);
      tabsEl.appendChild(controls);

      function applyHighlightState() {
        boomBtn.setAttribute('aria-pressed', String(boomEnabled));
        edgeBtn.setAttribute('aria-pressed', String(edgeEnabled));
        const wrappers = [...sheetContainerEl.querySelectorAll('.table-scroll')];
        wrappers.forEach((wrap) => {
          const rows = wrap.querySelectorAll('tbody tr');
          rows.forEach((tr) => {
            const hadBoom = tr.dataset.hasBoom === '1';
            const hadBust = tr.dataset.hasBust === '1';
            const desperation = tr.dataset.hasDesperation === '1';
            tr.classList.remove('row-boom', 'row-bust', 'row-desperation', 'edge-strong', 'edge-watch');
            if (boomEnabled) {
              if (hadBoom) tr.classList.add('row-boom');
              if (hadBust) tr.classList.add('row-bust');
              if (desperation) tr.classList.add('row-desperation');
            }
            if (edgeEnabled && tr.dataset.edgeTier === 'strong') tr.classList.add('edge-strong');
            if (edgeEnabled && tr.dataset.edgeTier === 'watch') tr.classList.add('edge-watch');
          });
        });
      }
      boomBtn.addEventListener('click', () => {
        boomEnabled = !boomEnabled;
        applyHighlightState();
      });
      edgeBtn.addEventListener('click', () => {
        edgeEnabled = !edgeEnabled;
        applyHighlightState();
      });
      applyHighlightState();
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
      wrapper.style.display = index === initialSheetIndex ? 'block' : 'none';

      const table = document.createElement('table');
      table.dataset.sheetIndex = String(index);

      const thead = document.createElement('thead');
      const tbody = document.createElement('tbody');

      // Determine the actual column header row, not the creator's explanatory text above it.
      const rows = sheet.rows || [];
      const normalize = (v) => String(v ?? '').toLowerCase().trim();
      let headerRowIndex = 0;
      for (let i = 0; i < rows.length; i += 1) {
        const norm = (rows[i] || []).map(normalize);
        if (norm.some((c) => c === 'player name')) {
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
      const teamColIndex = headers.findIndex((h) => normalize(h) === 'team');
      const cuffColIndex = headers.findIndex((h) => normalize(h) === 'cuff');
      const depthColIndex = headers.findIndex((h) => normalize(h).includes('team depth chart'));
      const ageColIndex = headers.findIndex((h) => normalize(h) === 'age');
      const primaryRankColIndex = headers.findIndex((h) => /(?:rk|rank)/.test(normalize(h)));
      const isSuperflex = normalize(sheet.name) === 'superflex new';
      const pointsVariant = VARIANT_BY_SHEET[sheet.name] || 'ppr';

      // Build visible columns, filtering out market/tier columns replaced by
      // the historical and projected fantasy-point columns, plus prop bets.
      const allIndices = Array.from({ length: columnCount }, (_, i) => i);
      const shouldHideColumn = (h) => {
        const t = normalize(h);
        const isSuperflexSecondaryRank = isSuperflex && (t === 'super flex strd' || t === 'super flex .5 ppr');
        return t === 'adp sleeper'
          || t === 'adp rt sports'
          || t === 'drafted'
          || t === 'my team'
          || t === '3rd yr wr'
          || t === 'rookie'
          || t.includes('tier')
          || isSuperflexSecondaryRank
          || (t.includes('prop') && t.includes('bet'))
          || t.includes('auction');
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

      // Construct column descriptors and inject synthetic points/Draft columns.
      const columns = visibleIndices.map((idx) => ({ kind: 'sheet', idx }));
      const rankColumnPosition = columns.findIndex((col) => col.kind === 'sheet' && col.idx === primaryRankColIndex);
      const pointsInsertAt = rankColumnPosition === -1 ? 0 : rankColumnPosition + 1;
      columns.splice(pointsInsertAt, 0, { kind: 'priorPoints' }, { kind: 'projectedPoints' });
      if (playerColIndex !== -1) {
        const playerColInColumns = columns.findIndex((c) => c.idx === playerColIndex);
        const insertAt = Math.max(0, playerColInColumns);
        columns.splice(insertAt, 0, { kind: 'draft' });
      }
      const playerColumnPosition = columns.findIndex((c) => c.kind === 'sheet' && c.idx === playerColIndex);
      if (playerColumnPosition !== -1) columns.splice(playerColumnPosition + 1, 0, { kind: 'edge' });
      const ageColumnPosition = columns.findIndex((c) => c.kind === 'sheet' && c.idx === ageColIndex);
      columns.splice(ageColumnPosition === -1 ? columns.length : ageColumnPosition, 0, { kind: 'leagueYears' });
      const cuffColumnPosition = columns.findIndex((c) => c.kind === 'sheet' && c.idx === cuffColIndex);
      columns.splice(cuffColumnPosition === -1 ? columns.length : cuffColumnPosition, 0, { kind: 'watch' });
      const defaultRankColumnPosition = columns.findIndex((col) => col.kind === 'sheet' && col.idx === primaryRankColIndex);

      const lastVisibleSheetIdx = (() => {
        const onlySheets = (columns.length ? columns : visibleIndices.map((idx) => ({ kind: 'sheet', idx }))).filter((c) => c.kind === 'sheet');
        return onlySheets.length ? onlySheets[onlySheets.length - 1].idx : -1;
      })();
      (columns.length ? columns : visibleIndices.map((idx) => ({ kind: 'sheet', idx }))).forEach((col, i) => {
        const th = document.createElement('th');
        if (col.kind === 'draft') {
          th.textContent = 'Draft';
          th.title = 'Assign this player to the fantasy team currently on the clock';
        } else if (col.kind === 'watch') {
          th.textContent = 'Watch';
          th.title = 'Add or remove this player from Will He Make It Back?';
          th.classList.add('col-watch');
        } else if (col.kind === 'edge') {
          th.textContent = 'Edge';
          th.title = 'ADP Edge 0–100 comparison score';
        } else if (col.kind === 'leagueYears') {
          th.textContent = 'NFL Year';
          th.title = 'Current NFL experience year (rookies = 1)';
          th.classList.add('league-years-column');
        } else if (col.kind === 'priorPoints') {
          th.textContent = `${PRIOR_SEASON} Total Fantasy Pts`;
          th.title = `${PRIOR_SEASON} total fantasy points in this view's scoring format`;
          th.classList.add('points-column');
        } else if (col.kind === 'projectedPoints') {
          th.textContent = `${CURRENT_SEASON} Projected Fantasy Pts`;
          th.title = `${CURRENT_SEASON} projected fantasy points in this view's scoring format`;
          th.classList.add('points-column');
        } else {
          const colIdx = col.idx;
          const label = colIdx === depthColIndex ? 'Team Depth Chart RB/WR/TE' : String(headers[colIdx] ?? `Column ${colIdx + 1}`);
          th.textContent = label;
          if (colIdx === depthColIndex) th.classList.add('depth-chart-column');
          if (colIdx === primaryRankColIndex) th.dataset.sortKind = 'rank';
        }
        // Action columns are not sortable.
        if (!['draft', 'watch'].includes(col.kind)) {
          th.addEventListener('click', () => sortTableByColumn(table, i));
        }
        // Sticky sequence: Draft, Player Name.
        if (col.kind === 'draft') {
          th.classList.add('sticky-left-0', 'col-draft');
        } else if (col.kind === 'sheet' && col.idx === playerColIndex) {
          th.classList.add('sticky-left-1');
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
        tr.dataset.tier = '';
        const rowPosition = String((rows[r] && rows[r][posColIndex]) ?? '').toUpperCase();
        const variant = VARIANT_BY_SHEET[sheet.name];
        const edgeRecord = edgeByIdentity.get(edgeIdentity(playerNameForRow, rowPosition));
        const pointsRecord = pointsByIdentity.get(edgeIdentity(playerNameForRow, rowPosition));
        const edgeValue = edgeRecord?.variants?.[variant];
        const canonicalKey = DraftRoom.playerKey({ id: edgeRecord?.id, name: playerNameForRow, position: rowPosition });
        tr.dataset.playerKey = canonicalKey;
        tr.dataset.position = rowPosition;
        if (isPlayerKeyDrafted(canonicalKey)) tr.style.display = 'none';
        let catalogPlayer = playerByKey.get(canonicalKey);
        if (!catalogPlayer) {
          catalogPlayer = { playerKey: canonicalKey, id: edgeRecord?.id || null, name: playerNameForRow, position: rowPosition, team: edgeRecord?.team || String((rows[r] && rows[r][teamColIndex]) || ''), variants: {} };
          playerByKey.set(canonicalKey, catalogPlayer);
          playerCatalog.push(catalogPlayer);
        }
        const currentRank = positiveRank(edgeValue?.marketRank);
        const workbookRank = positiveRank((rows[r] && rows[r][primaryRankColIndex]));
        catalogPlayer.variants[variant] = {
          marketRank: currentRank ?? workbookRank,
          boardRank: currentRank ?? workbookRank,
          edgeScore: Number.isFinite(Number(edgeValue?.score)) ? Number(edgeValue.score) : null,
          edgeTier: edgeValue?.tier || 'none',
          workbookTier: tr.dataset.tier || null,
          projectedPoints: Number(edgeRecord?.projectionPoints?.[variant]) || null,
        };
        // Click a row to make keyboard drafting unambiguous across all four tabs.
        tr.addEventListener('click', (event) => {
          if (event.target?.closest('button, input, select, a')) return;
          selectPlayer(canonicalKey);
        });
        if (edgeValue) tr.dataset.edgeTier = edgeValue.tier;

        // Reproduce the creator's workbook signals exactly.
        const boomColIdx = headers.findIndex((h) => normalize(h) === 'boom factor');
        const connectColIdx = headers.findIndex((h) => normalize(h).includes('boom connect'));
        const earlyColIdx = headers.findIndex((h) => normalize(h).includes('1st 5') && normalize(h).includes('sos'));
        const lineColIdx = headers.findIndex((h) => normalize(h) === 'oline ovrl');
        const boomText = String((rows[r] && rows[r][boomColIdx]) ?? '').trim().toUpperCase();
        const connectText = String((rows[r] && rows[r][connectColIdx]) ?? '').toLowerCase();
        const earlyText = String((rows[r] && rows[r][earlyColIdx]) ?? '').trim().toLowerCase();
        const lineText = String((rows[r] && rows[r][lineColIdx]) ?? '').trim().toUpperCase();
        if (boomText === 'BOOM') tr.dataset.hasBoom = '1';
        if (connectText.includes('desperation')) tr.dataset.hasDesperation = '1';
        if (earlyText === 'hard' && (lineText === 'OK' || lineText === 'BAD')) tr.dataset.hasBust = '1';
        (columns.length ? columns : visibleIndices.map((idx) => ({ kind: 'sheet', idx }))).forEach((col) => {
          const td = document.createElement('td');
          if (col.kind === 'draft') {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'draft-player-btn';
            button.textContent = 'Draft';
            button.dataset.role = 'draft';
            button.dataset.player = playerNameForRow;
            button.dataset.playerKey = canonicalKey;
            button.title = `Draft ${playerNameForRow} to the team on the clock`;
            button.setAttribute('aria-label', `Draft ${playerNameForRow} to the team on the clock`);
            button.addEventListener('click', (event) => { event.stopPropagation(); draftPlayer(canonicalKey); });
            td.appendChild(button);
            td.classList.add('sticky-left-0', 'col-draft');
          } else if (col.kind === 'watch') {
            td.className = 'watch-cell col-watch';
            if (['QB', 'RB', 'WR', 'TE'].includes(rowPosition)) {
              const button = document.createElement('button');
              button.type = 'button';
              button.className = 'watch-player-btn';
              button.dataset.playerKey = canonicalKey;
              button.textContent = 'Watch';
              button.title = 'Add to Will He Make It Back?';
              button.setAttribute('aria-label', `Add ${playerNameForRow} to Will He Make It Back?`);
              button.setAttribute('aria-pressed', 'false');
              button.addEventListener('click', (event) => {
                event.stopPropagation();
                toggleWatchPlayer(canonicalKey);
              });
              td.appendChild(button);
            } else td.textContent = '—';
          } else if (col.kind === 'edge') {
            td.className = 'edge-cell';
            if (edgeValue && Number.isFinite(edgeValue.score) && ['QB', 'RB', 'WR', 'TE'].includes(rowPosition)) {
              const pill = document.createElement('button');
              pill.className = `edge-pill tier-${edgeValue.tier}`;
              pill.textContent = String(edgeValue.score);
              pill.title = `Open ADP Edge details for ${playerNameForRow}`;
              pill.setAttribute('aria-label', `${playerNameForRow} ADP Edge score ${edgeValue.score}; open details`);
              pill.addEventListener('click', (event) => { event.stopPropagation(); openEdgeDetails(edgeRecord, variant); });
              td.dataset.sortValue = String(edgeValue.score);
              td.appendChild(pill);
            } else {
              td.textContent = '—';
            }
          } else if (col.kind === 'priorPoints' || col.kind === 'projectedPoints') {
            const pointSet = col.kind === 'priorPoints' ? pointsRecord?.priorSeasonPoints : pointsRecord?.projectionPoints;
            const value = pointSet?.[pointsVariant];
            td.className = 'points-cell';
            td.textContent = formatFantasyPoints(value);
            if (value !== null && value !== undefined && value !== '') td.dataset.sortValue = String(value);
          } else if (col.kind === 'leagueYears') {
            const value = pointsRecord?.leagueYears;
            td.className = 'league-years-cell';
            td.textContent = value === null || value === undefined ? '—' : String(value);
            if (Number.isFinite(Number(value))) td.dataset.sortValue = String(value);
          } else {
            const colIdx = col.idx;
            const value = (rows[r] && rows[r][colIdx]) ?? '';
            const text = String(value);
            // color POS and also color the player name cell using POS
            if (colIdx === primaryRankColIndex) {
              if (edgeRecord?.availability?.seasonEnding) {
                td.textContent = 'OUT';
                td.dataset.sortMissing = 'true';
                td.title = `${edgeRecord.availability.injuryDetail || 'Season-ending injury'}; removed from current ranks. Previous provider rank: ${edgeValue.reportedMarketRankBeforeInjury ?? '—'}`;
                td.classList.add('rank-updated');
              } else if (positiveRank(edgeValue?.marketRank) !== null) {
                const rank = positiveRank(edgeValue.marketRank);
                td.textContent = rank.toFixed(rank % 1 ? 1 : 0);
                td.dataset.sortValue = String(rank);
                td.title = `${edgeValue.marketRankSource || 'Current market rank'}${edgeValue.marketRankUpdatedAt ? ` · ${new Date(edgeValue.marketRankUpdatedAt).toLocaleString()}` : ''}; workbook: ${text || 'blank'}`;
                if (Number(text) !== rank) td.classList.add('rank-updated');
              } else if (positiveRank(value) !== null) {
                td.textContent = text;
                td.dataset.sortValue = String(positiveRank(value));
              } else {
                td.textContent = '—';
                td.dataset.sortMissing = 'true';
              }
            } else if (colIdx === posColIndex) {
              td.textContent = text;
              td.classList.add(posClass(text));
            } else if (colIdx === teamColIndex) {
              const displayedTeam = String(edgeRecord?.team || text || 'FA').toUpperCase();
              td.textContent = displayedTeam;
              td.classList.add('team-abbr');
              td.style.setProperty('--team-color', TEAM_COLORS[displayedTeam] || '#cbd5e1');
              if (edgeRecord?.team && edgeRecord.team !== text) {
                td.classList.add('team-updated');
                td.title = `Current directory: ${edgeRecord.team} (${edgeRecord.rosterStatus}); workbook: ${text || 'blank'}`;
              }
            } else if (colIdx === depthColIndex) {
              const currentDepthRole = String(edgeRecord?.metrics?.depthRole || '').trim();
              td.textContent = currentDepthRole || text || '—';
              td.classList.add('depth-chart-cell');
              if (currentDepthRole && currentDepthRole !== text) {
                td.classList.add('depth-updated');
                td.title = `Current directory: ${currentDepthRole}; workbook: ${text || 'blank'}`;
              }
            } else if (colIdx === playerColIndex) {
              td.textContent = text;
              const posText = String((rows[r] && rows[r][posColIndex]) ?? '');
              const cls = posClass(posText);
              if (cls) td.classList.add(cls);
              td.classList.add('sticky-left-1', 'player-name-cell');
              if (['QB', 'RB', 'WR', 'TE'].includes(rowPosition)) {
                td.title = 'Double-click to add to Will He Make It Back?';
                td.addEventListener('dblclick', (event) => {
                  event.preventDefault();
                  event.stopPropagation();
                  toggleWatchPlayer(canonicalKey);
                });
              }
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
      [...headerRow.children].forEach((th, columnIndex) => {
        const column = columns[columnIndex];
        if (!column || column.kind === 'draft') return;
        const descriptor = {
          ...column,
          label: column.kind === 'sheet' ? (normalize(headers[column.idx]) || String(column.idx)) : column.kind,
        };
        attachColumnResizer(th, table, columnIndex, columnStorageKey(sheet.name, descriptor));
      });
      assignCurrentFormatTiers(table, defaultRankColumnPosition, pointsVariant);
      if (defaultRankColumnPosition >= 0) sortTableByColumn(table, defaultRankColumnPosition, true);
      else {
        applyTierSeparators(table);
        applyNextPickMarker(table);
      }
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
      [...tabsEl.querySelectorAll('.tab')].forEach((tab, i) => {
        tab.setAttribute('aria-selected', i === index ? 'true' : 'false');
      });
      // Toggle table visibility
      const wrappers = [...sheetContainerEl.querySelectorAll('.table-scroll')];
      wrappers.forEach((el, i) => {
        el.style.display = i === index ? 'block' : 'none';
      });
      activeVariant = VARIANT_BY_SHEET[sheets[index]?.name] || 'ppr';
      updateGlobalScrollerWidth();
      refreshDraftRoomUI();
    }
    selectSheetByVariant = (variant) => {
      const index = sheets.findIndex((sheet) => VARIANT_BY_SHEET[sheet.name] === variant);
      selectSheet(index >= 0 ? index : 0);
    };

    // Show a small tip
    statusEl.textContent = `${data.fileName} · data ${data.dataVersion}. Edge pills open an auditable explanation.`;

    // Ensure and initialize bottom-fixed scrollbar
    ensureGlobalScroller();
    // Defer to next frame so layout is ready
    requestAnimationFrame(updateGlobalScrollerWidth);
    selectSheet(initialSheetIndex);

    // boom toggle and search box handled above
  }

  function applySearchFilter() {
    const q = (playerSearchInput && playerSearchInput.value || '').trim().toLowerCase();
    const active = getActiveWrapper();
    if (!active) return;
    const rows = active.querySelectorAll('tbody tr');
    // Use dataset set earlier for player name; fallback to first text match in row
    rows.forEach((tr) => {
      if (!q) { tr.style.display = isPlayerKeyDrafted(tr.dataset.playerKey) ? 'none' : ''; return; }
      const playerName = (tr.dataset.player || '').toLowerCase();
      const rowText = tr.textContent.toLowerCase();
      const match = playerName.includes(q) || rowText.includes(q);
      // Maintain hide if drafted/taken
      const hiddenByDraft = isPlayerKeyDrafted(tr.dataset.playerKey);
      tr.style.display = (match && !hiddenByDraft) ? '' : 'none';
    });
    const table = active.querySelector('table');
    applyTierSeparators(table);
    applyNextPickMarker(table);
  }

  function sortTableByColumn(table, columnIndex, forceAscending = null) {
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const ths = [...thead.querySelectorAll('th')];

    // Determine current sort state
    const current = ths[columnIndex];
    const isAsc = current.classList.contains('th-sort-asc');
    const newAsc = forceAscending === null ? !isAsc : Boolean(forceAscending);

    // Clear all sort indicators
    ths.forEach((th) => th.classList.remove('th-sort-asc', 'th-sort-desc'));
    current.classList.add(newAsc ? 'th-sort-asc' : 'th-sort-desc');

    // Helpers for robust sorting
    const normalizeText = (s) => (s || '').trim();
    const isEmptyCell = (s) => {
      const t = normalizeText(s);
      return t === '' || t === '-' || t === '—' || t.toLowerCase() === 'n/a';
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
      const aCell = a.children[columnIndex];
      const bCell = b.children[columnIndex];
      const aText = normalizeText(aCell?.dataset.sortValue || aCell?.textContent || '');
      const bText = normalizeText(bCell?.dataset.sortValue || bCell?.textContent || '');

      const rankSort = current.dataset.sortKind === 'rank';
      const aParsedRank = rankSort ? parseNumber(aText) : null;
      const bParsedRank = rankSort ? parseNumber(bText) : null;
      const aEmpty = aCell?.dataset.sortMissing === 'true' || isEmptyCell(aText) || (rankSort && (aParsedRank === null || aParsedRank <= 0));
      const bEmpty = bCell?.dataset.sortMissing === 'true' || isEmptyCell(bText) || (rankSort && (bParsedRank === null || bParsedRank <= 0));
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
    applyTierSeparators(table);
    applyNextPickMarker(table);
  }

  async function init() {
    try {
      const apiPath = (window && window.location && window.location.hostname.includes('vercel.app')) ? '/api/workbook' : '/api/workbook';
      const [workbookResponse, edgeResponse] = await Promise.all([
        fetch(`${apiPath}?cb=${Date.now()}`, { cache: 'no-store' }),
        fetch(`${EDGE_DATA_URL}?cb=${Date.now()}`, { cache: 'no-store' }),
      ]);
      const data = await workbookResponse.json();
      if (!workbookResponse.ok) throw new Error(data?.error || 'Failed to load workbook');
      if (!edgeResponse.ok) throw new Error('Failed to load the offline ADP Edge dataset');
      edgeData = await edgeResponse.json();
      if (Number(data.season) !== Number(CURRENT_SEASON) || Number(edgeData.season) !== Number(CURRENT_SEASON)) throw new Error('Season metadata mismatch');
      edgeByIdentity = new Map((edgeData.players || []).map((player) => [edgeIdentity(player.name, player.position), player]));
      pointsByIdentity = new Map((edgeData.points || []).map((player) => [edgeIdentity(player.name, player.position), player]));
      playerCatalog = (edgeData.players || []).map((player) => ({
        playerKey: DraftRoom.playerKey(player),
        id: player.id || null,
        name: player.name,
        position: player.position,
        team: player.team || null,
        variants: Object.fromEntries(Object.entries(player.variants || {}).map(([variant, value]) => [variant, {
          marketRank: value.marketRank,
          boardRank: null,
          edgeScore: value.score,
          edgeTier: value.tier,
          workbookTier: null,
          projectedPoints: Number(player.projectionPoints?.[variant]) || null,
        }])),
      }));
      playerByKey = new Map(playerCatalog.map((player) => [player.playerKey, player]));
      const createdProfile = ensureDraftState();
      bindDraftRoomControls();
      renderWorkbook(data);
      if (createdProfile) requestAnimationFrame(() => openLeagueSetup(false));
    } catch (err) {
      statusEl.textContent = (err && err.message) || 'Failed to load workbook';
    }
  }

  init();
  if (refreshBtn) {
    refreshBtn.addEventListener('click', () => {
      const pickCount = draftSession?.picks.length || 0;
      const watchCount = draftSession?.watchlist?.length || 0;
      if (!activeProfile || (!pickCount && !watchCount)) { showToast('This draft session is already empty.'); return; }
      if (!window.confirm(`Clear ${pickCount} recorded pick${pickCount === 1 ? '' : 's'} and ${watchCount} tracked player${watchCount === 1 ? '' : 's'} from ${activeProfile.name}? The league profile will be kept.`)) return;
      draftSession = draftStore.clearSession(activeProfile);
      selectedPlayerKey = null;
      refreshDraftRoomUI();
      showToast('Draft session cleared.');
    });
  }
})();
