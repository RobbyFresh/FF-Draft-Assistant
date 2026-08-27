const assert = require('assert');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const { loadWorkbookToJson } = require('../server');

const root = path.resolve(__dirname, '..');
const workbookPath = path.join(root, 'Fantasy Football Cheat Sheet with Boom Outlier 2026.xlsx');
const publicWorkbookPath = path.join(root, 'public', 'Fantasy Football Cheat Sheet with Boom Outlier 2026.xlsx');
const dataset = JSON.parse(fs.readFileSync(path.join(root, 'public', 'data', 'adp-edge-2026.json'), 'utf8'));
const backtest = JSON.parse(fs.readFileSync(path.join(root, 'data', 'reports', 'backtest-2026.json'), 'utf8'));
const reviewInput = JSON.parse(fs.readFileSync(path.join(root, 'data', 'inputs', 'edge-review-2025.json'), 'utf8'));
const sha = (file) => crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
const norm = (v) => String(v || '').trim().toLowerCase();

assert.equal(sha(workbookPath), sha(publicWorkbookPath), 'served workbook must be byte-identical to the supplied workbook');
const api = loadWorkbookToJson();
assert.equal(api.season, 2026);
assert.equal(api.dataVersion, '2026.5');
assert.equal(api.fileName, 'Fantasy Football Cheat Sheet with Boom Outlier 2026.xlsx');
assert.deepEqual(api.sheets.map((s) => s.name), ['STRD', '.5 PPR NEW', 'FULL PPR NEW', 'SUPERFLEX NEW']);
assert(loadWorkbookToJson(path.join(root, 'missing-workbook.xlsx')).error.includes('not found'));

const workbook = XLSX.readFile(workbookPath, { cellStyles: true });
let boomCount = 0;
let bustCount = 0;
let desperationCount = 0;
let neutralCount = 0;
for (const sheetName of workbook.SheetNames) {
  const ws = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
  const headerRow = rows.findIndex((row) => row.some((v) => norm(v) === 'player name'));
  assert(headerRow >= 0, `${sheetName} header row`);
  const headers = rows[headerRow].map(norm);
  const player = headers.findIndex((h) => h === 'player name');
  const boom = headers.findIndex((h) => h === 'boom factor');
  const connect = headers.findIndex((h) => h.includes('boom connect'));
  const early = headers.findIndex((h) => h.includes('1st 5') && h.includes('sos'));
  const line = headers.findIndex((h) => h === 'oline ovrl');
  for (let index = headerRow + 1; index < rows.length; index += 1) {
    const row = rows[index];
    const playerName = String(row[player] || '');
    if (!playerName) continue;
    const isBoom = String(row[boom] || '').trim().toUpperCase() === 'BOOM';
    const isBust = norm(row[early]) === 'hard' && ['OK', 'BAD'].includes(String(row[line] || '').trim().toUpperCase());
    const isDesperation = norm(row[connect]).includes('desperation');
    const cell = ws[XLSX.utils.encode_cell({ r: index, c: player })];
    const fill = cell?.s?.fgColor?.rgb || '';
    assert.equal(fill === 'DFFEB6', isBoom, `${sheetName} ${playerName}: green fill must equal Boom Factor=BOOM`);
    assert.equal(fill === 'EABFBE', isBust, `${sheetName} ${playerName}: red fill must equal hard early SoS plus OK/BAD line`);
    boomCount += Number(isBoom);
    bustCount += Number(isBust);
    desperationCount += Number(isDesperation);
    neutralCount += Number(!isBoom && !isBust && !isDesperation);
  }
}
assert(boomCount > 0 && bustCount > 0 && desperationCount > 0 && neutralCount > 0, 'boom, bust, desperation, and neutral fixtures must all exist');

assert.equal(dataset.season, 2026);
assert.equal(dataset.schemaVersion, '1.4.0');
assert.equal(dataset.modelVersion, 'adp-edge-2026.5');
assert.equal(dataset.draftAvailability?.modelVersion, 'return-risk-2026.6');
assert(['QB', 'RB', 'WR', 'TE'].every((position) => dataset.draftAvailability.uncertaintyByPosition[position]), 'return-risk uncertainty must cover every scored position');
assert.equal(backtest.accepted, true, 'rolling backtest must beat the market baseline');
assert(backtest.aggregate.lift > 0, 'rolling backtest lift must be positive');
assert.deepEqual(backtest.retainedHistoricalCorePositions, ['QB'], 'only validated upside positions may retain the historical core');
assert.deepEqual(backtest.retainedRegressionRiskPositions, ['TE'], 'TE is the only independently validated downside position');
assert(backtest.regressionRiskAggregate.lift > 0, 'retained regression-risk model must beat its matched baseline');
assert.equal(reviewInput.players.length, 20, '2025 review database must contain both supplied cohorts');
assert.deepEqual(backtest.reviewCohorts.summary.breakout, { requested: 12, matched: 12, legacyCaught: 2, updatedCaught: 2 });
assert.equal(backtest.reviewCohorts.summary.regression.requested, 8);
assert.equal(backtest.reviewCohorts.summary.regression.matched, 8);
assert.equal(backtest.reviewCohorts.summary.regression.updatedCaught, 1);
assert.equal(dataset.calibration.backtest.accepted, true);
assert(dataset.calibration.backtest.regressionRisk.lift > 0);
assert.deepEqual(dataset.variantMap, { STRD: 'standard', '.5 PPR NEW': 'halfPpr', 'FULL PPR NEW': 'ppr', 'SUPERFLEX NEW': 'superflex' });
assert(dataset.marketCoverage.currentMarketVariants > dataset.marketCoverage.workbookFallbackVariants * 20, 'current ADP must dominate workbook fallback coverage');
assert(dataset.sourcePrecedence[0].includes('injury'), 'injury/news state must have top source precedence');
assert(Object.values(dataset.sourceFreshness).every((source) => source.fetchedAt), 'every production feed must expose a fetch timestamp');
assert(Array.isArray(dataset.points) && dataset.points.length >= 800, 'points dataset must cover the draft board');
const pointIdentities = new Set();
for (const pointRecord of dataset.points) {
  const identity = `${pointRecord.normalizedName}|${pointRecord.position}`;
  assert(!pointIdentities.has(identity), `duplicate points identity ${identity}`);
  pointIdentities.add(identity);
  for (const variant of Object.values(dataset.variantMap)) {
    assert(Object.hasOwn(pointRecord.priorSeasonPoints, variant), `${identity} prior-season ${variant} points`);
    assert(Object.hasOwn(pointRecord.projectionPoints, variant), `${identity} projected ${variant} points`);
  }
}
const lamarPoints = dataset.points.find((player) => player.name === 'Lamar Jackson' && player.position === 'QB');
assert(lamarPoints?.projectionPoints.ppr > 300, 'Lamar Jackson QB must retain his 2026 projection despite the same-name defensive back');
assert(lamarPoints.projectionSources.includes('Sleeper / RotoWire') && lamarPoints.projectionSources.includes('FantasyPros consensus'));
assert.equal(dataset.points.find((player) => player.name === 'Jadarian Price' && player.position === 'RB')?.leagueYears, 1);
assert.equal(dataset.points.find((player) => player.name === 'Rashee Rice' && player.position === 'WR')?.leagueYears, 4);
assert(dataset.calibration.coverageTop240 >= 95, 'top-240 join coverage must be at least 95%');
const identities = new Set();
let scored = 0;
for (const player of dataset.players) {
  const identity = `${player.normalizedName}|${player.position}`;
  assert(!identities.has(identity), `duplicate identity ${identity}`);
  identities.add(identity);
  assert(['QB', 'RB', 'WR', 'TE'].includes(player.position));
  assert(player.confidence >= 0 && player.confidence <= 100);
  for (const [variant, value] of Object.entries(player.variants)) {
    if (!Number.isFinite(value.score)) continue;
    scored += 1;
    assert(value.score >= 0 && value.score <= 100, `${identity} ${variant} score range`);
    assert(['strong', 'watch', 'none'].includes(value.tier));
    assert(!Object.hasOwn(value, 'regressionRiskFlag') && !Object.hasOwn(value, 'regressionRiskSignal'), 'warning-only fields must not ship');
    assert(value.reasons.length >= (player.availability.seasonEnding ? 1 : 2) && value.reasons.length <= 4);
    assert(value.sourceRefs.every((id) => dataset.sources[id]), `${identity} ${variant} source references`);
    if (value.tier === 'strong') {
      assert(value.score >= 70 && value.expectedDelta >= 3 && player.confidence >= 65);
    }
    if (value.tier === 'watch') {
      assert(value.score >= 58 && value.expectedDelta >= 1 && player.confidence >= 50);
    }
  }
}
assert(scored > 500, 'expected useful scoring coverage across four variants');
assert(dataset.players.every((player) => !Object.hasOwn(player, 'researchFlags')), 'warning-only player fields must be removed');
assert(['QB', 'RB', 'WR', 'TE'].every((position) => ['reversion', 'context', 'availability'].every((component) => dataset.positionWeights[position][component] > 0)), 'updated Edge must score reversion, context, and availability at every position');

const higgins = dataset.players.find((player) => player.name === 'Jayden Higgins' && player.position === 'WR');
assert(higgins?.availability.seasonEnding, 'Jayden Higgins season-ending ACL must override the projection feed');
assert.equal(higgins.projectionPoints.ppr, 0);
assert.equal(higgins.variants.ppr.marketRank, null);
assert.equal(higgins.variants.ppr.score, 0);

const tua = dataset.players.find((player) => player.name === 'Tua Tagovailoa' && player.position === 'QB');
assert(Number.isFinite(tua?.metrics.passingAdot) && Number.isFinite(tua?.metrics.sackRate) && Number.isFinite(tua?.metrics.qbFloorSignal), 'QB floor diagnostics must be present');
const jeudy = dataset.players.find((player) => player.name === 'Jerry Jeudy' && player.position === 'WR');
assert(Number.isFinite(jeudy?.metrics.maxGameReceivingYardShare) && Number.isFinite(jeudy?.metrics.projectedQbChangeSignal), 'WR concentration and QB-change diagnostics must be present');

const jonah = dataset.players.find((player) => player.name === 'Jonah Coleman');
assert(jonah?.rookie, 'Jonah Coleman should be recognized as a rookie from current roster/draft data');
assert.equal(jonah.metrics.games, null, 'rookies must not receive fabricated prior NFL games');
assert(jonah.metrics.projectedPprPoints > 0 && jonah.metrics.projectedOpportunity > 0, 'rookies should receive season projections');
assert(jonah.components.opportunity > 0, 'rookies should receive a positive projected-role signal without fabricated NFL history');
assert(jonah.variants.ppr.reasons.some((reason) => reason.startsWith('Projection ')));
assert.equal(jonah.metrics.workbookEnvironmentDisplayOnly, true);

const diggs = dataset.players.find((player) => player.name === 'Stefon Diggs');
assert.equal(diggs.team, 'WAS', 'current directory should supersede stale workbook team');
assert.equal(diggs.rosterStatus, 'Active');
assert.equal(diggs.metrics.games, 17, 'blank NFL IDs must fall back to the correct normalized-name stat row');
assert.equal(diggs.metrics.targetShare, 21.2);
assert(diggs.variants.ppr.sourceRefs.includes('sleeperPlayers') && diggs.variants.ppr.sourceRefs.includes('sleeperProjections'));

for (const player of dataset.players) {
  for (const value of Object.values(player.variants)) {
    if (!Array.isArray(value.reasons)) continue;
    for (const reason of value.reasons) {
      assert(!reason.startsWith('Schedule helps: hard'), `${player.name}: hard schedule cannot help`);
      assert(!reason.startsWith('Schedule limits: easy'), `${player.name}: easy schedule cannot limit`);
    }
  }
}

for (const file of ['server.js', 'api/workbook.js', 'public/index.html', 'public/app.js', 'README.md']) {
  const text = fs.readFileSync(path.join(root, file), 'utf8');
  assert(!text.includes('Boom Outlier 2025.xlsx') && !text.includes('Fantasy Football 2025 Draft'), `${file} contains a stale 2025 season/file reference`);
}
const appSource = fs.readFileSync(path.join(root, 'public', 'app.js'), 'utf8');
const cssSource = fs.readFileSync(path.join(root, 'public', 'styles.css'), 'utf8');
const htmlSource = fs.readFileSync(path.join(root, 'public', 'index.html'), 'utf8');
const draftRoomSource = fs.readFileSync(path.join(root, 'public', 'draft-room.js'), 'utf8');
assert(appSource.includes('applyTierSeparators(table)'), 'tier separators must be applied to every rendered/sorted table');
assert(appSource.includes('assignCurrentFormatTiers(table, defaultRankColumnPosition, pointsVariant)'), 'every scoring table must rebuild tiers from its own current rank pool');
assert(appSource.includes('current-rank-gap-v1'), 'current tier model must be explicit and auditable');
assert(!appSource.includes('fullPprTierByRank'), 'Superflex must not inherit stale Full-PPR workbook tiers');
assert(appSource.includes("sortedHeader.dataset.sortKind !== 'rank'"), 'tier lines must be suppressed when the table is not rank-sorted');
assert(!appSource.includes('Research-only warnings') && !appSource.includes('Validated regression warning'), 'warning-only UI must be removed');
assert(appSource.includes('rank-updated') && cssSource.includes('.rank-updated'), 'current market ranks must replace stale workbook ranks in the board');
assert(appSource.includes("sheet.name.replace(/\\s+NEW$/i, '')"), 'display tabs must omit the workbook-only NEW suffix');
assert(appSource.includes('sortTableByColumn(table, defaultRankColumnPosition, true)'), 'every scoring table must default to ascending rank order');
assert(appSource.includes("td.dataset.sortMissing = 'true'") && appSource.includes('numeric > 0'), 'zero and missing ranks must be treated as unavailable and sorted last');
assert(cssSource.includes('tr.tier-break td'), 'tier separator styling must exist');
assert(appSource.includes("t === 'adp sleeper'"), 'ADP Sleeper must be replaced in the rendered columns');
assert(appSource.includes("t === 'adp rt sports'"), 'ADP RT Sports must be hidden from the rendered columns');
assert(appSource.includes("t === '3rd yr wr'") && appSource.includes("t === 'rookie'"), 'legacy breakout flags must be hidden');
assert(appSource.includes("kind: 'leagueYears'"), 'one NFL experience-year column must replace the legacy flags');
assert(appSource.includes("t.includes('tier')"), 'tier columns must be replaced in the rendered columns');
assert(appSource.includes("kind: 'priorPoints'"), 'prior-season fantasy points column must be rendered');
assert(appSource.includes("kind: 'projectedPoints'"), 'projected fantasy points column must be rendered');
assert(appSource.includes("kind: 'draft'"), 'the board must expose one draft action instead of separate My Team/Taken actions');
assert(!appSource.includes("columns.splice(mtIndex + 1, 0, { kind: 'taken' })"), 'the legacy Taken action column must be removed');
assert(cssSource.includes('.draft-player-btn'), 'the one-click Draft action must be styled');
assert(cssSource.includes('.pos-K { color: #fbbf24; }'), 'kicker names and position labels must use a distinct amber color');
assert(!appSource.includes("event.key.toLowerCase() === 'd'"), 'typing D must never invoke the draft action');
assert(htmlSource.includes('/draft-room.js'), 'the offline draft engine must load before the board');
assert(cssSource.includes('width: 96px'), 'fantasy-point columns must use the compact width');
assert(htmlSource.includes('id="sidebar-resizer"') && appSource.includes('SIDEBAR_WIDTH_KEY'), 'the right sidebar must be user-resizable and persistent');
assert(appSource.includes('attachColumnResizer') && cssSource.includes('.column-resizer'), 'table columns must expose invisible resize handles');
assert(htmlSource.includes('id="league-view-toggle"') && appSource.includes('renderMyTeamSidebar'), 'the league pane must switch to a full My Team view');
assert(htmlSource.includes('id="risk-player-search"') && appSource.includes('availableWatchMatches'), 'Return Risk must use a manual available-player search');
assert(htmlSource.includes('id="risk-info-button"') && htmlSource.includes('id="risk-info-dialog"'), 'Return Risk statuses must have an accessible in-app explanation');
assert(htmlSource.includes('id="last-drafted-player"') && appSource.includes('DraftRoom.lastDraftedPick(draftSession)'), 'the clock bar must display the latest recorded player');
assert(htmlSource.includes('70%+ Return Risk') && htmlSource.includes('60–69% Return Risk') && htmlSource.includes('Below 60% Return Risk'), 'the status guide must document the current thresholds');
assert(!appSource.includes('const shortlisted = []'), 'Return Risk must not auto-fill an ADP Edge shortlist');
assert(!draftRoomSource.includes('ADP Edge') && !draftRoomSource.includes('edgeScore *'), 'Return Risk recommendations must remain independent from ADP Edge');
assert(appSource.includes('draftSession.watchlist') && appSource.includes('removeWatchPlayer'), 'manual watchlist choices must be stored and removable');
assert(appSource.includes("addEventListener('dblclick'") && appSource.includes('toggleWatchPlayer(canonicalKey)'), 'double-clicking a player name must toggle Return Risk tracking');
assert(appSource.includes('applyPlayerHighlightState()') && cssSource.includes('.row-watched td'), 'tracked Return Risk players must keep the soft-white row highlight');
assert(appSource.includes("{ kind: 'watch' }") && appSource.includes("className = 'watch-player-btn'"), 'the board must add a Watch action immediately before Cuff');
assert(appSource.includes('Team Depth Chart RB/WR/TE') && dataset.players.some((player) => player.position === 'RB' && /^RB\d+$/.test(player.metrics?.depthRole)), 'the depth-chart column must include current RB roles');
assert(appSource.includes('const TEAM_COLORS') && appSource.includes("classList.add('team-abbr')") && cssSource.includes('.team-abbr'), 'TEAM abbreviations must use the NFL team color palette');
assert(appSource.includes('applyNextPickMarker') && cssSource.includes('tr.next-pick-line td'), 'the board must mark the available-player depth at the next user pick');
assert(appSource.includes('DraftRoom.nextPickBoardIndex'), 'the white line must use the inclusive next-pick board slot');
assert(cssSource.includes('border-top: 5px solid rgba(255, 255, 255'), 'the next-pick marker must be a thicker white line than tier separators');

console.log(`Validated 4 sheets, ${boomCount} Boom rows, ${bustCount} Bust rows, ${desperationCount} desperation rows, and ${scored} Edge variants.`);
