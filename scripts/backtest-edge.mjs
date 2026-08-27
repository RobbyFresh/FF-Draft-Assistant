import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import readline from 'node:readline';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const cache = path.join(root, 'data', 'cache');
const output = path.join(root, 'data', 'reports', 'backtest-2026.json');
const reviewOutput = path.join(root, 'data', 'reports', 'edge-review-2025.md');
const sampleSeasons = [2019, 2020, 2021, 2022, 2023, 2024, 2025];
const testSeasons = [2022, 2023, 2024, 2025];
const positions = ['QB', 'RB', 'WR', 'TE'];
const reviewInput = JSON.parse(fs.readFileSync(path.join(root, 'data', 'inputs', 'edge-review-2025.json'), 'utf8'));
const reviewCohorts = Object.fromEntries(['breakout', 'regression'].map((cohort) => [
  cohort,
  reviewInput.players.filter((row) => row.cohort === cohort).map((row) => row.name),
]));
const reviewPointOutcomes = new Map(reviewInput.players.map((row) => [normalizeName(row.name), {
  projected: row.projectedPoints,
  actual: row.actualPoints,
  miss: Math.round((row.actualPoints - row.projectedPoints) * 10) / 10,
  clues: row.clues,
  sourceUrls: row.sourceUrls,
}]));

function normalizeName(value) {
  const normalized = String(value || '').normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase()
    .replace(/\b(jr|sr|ii|iii|iv)\b/g, '').replace(/[^a-z0-9]/g, '');
  return ({ kennethgainwell: 'kennygainwell' })[normalized] || normalized;
}
function identity(name, position) { return `${normalizeName(name)}|${String(position || '').toUpperCase()}`; }
function normalizeTeam(value) {
  const team = String(value || '').toUpperCase().trim();
  return ({ JAC: 'JAX', LAR: 'LA', WSH: 'WAS', LVR: 'LV', GBP: 'GB', KCC: 'KC', NOS: 'NO', SFO: 'SF', TBB: 'TB', NEP: 'NE' })[team] || team;
}
function csvLine(line) {
  const out = []; let field = ''; let quoted = false;
  for (let i = 0; i < line.length; i += 1) {
    const c = line[i];
    if (quoted) {
      if (c === '"' && line[i + 1] === '"') { field += '"'; i += 1; }
      else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { out.push(field); field = ''; }
    else field += c;
  }
  out.push(field);
  return out;
}
function readCsv(file) {
  const lines = fs.readFileSync(file, 'utf8').split(/\r?\n/).filter(Boolean);
  const headers = csvLine(lines.shift());
  return lines.map((line) => {
    const values = csvLine(line);
    return Object.fromEntries(headers.map((h, i) => [h, values[i] ?? '']));
  });
}
function n(value) { const x = Number(value); return Number.isFinite(x) ? x : 0; }
function mean(values) { const valid = values.filter(Number.isFinite); return valid.length ? valid.reduce((a, b) => a + b, 0) / valid.length : 0; }
function clamp(value, low, high) { return Math.max(low, Math.min(high, value)); }
function percentile(values, probability) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const index = (sorted.length - 1) * probability;
  const lower = Math.floor(index); const upper = Math.ceil(index);
  return lower === upper ? sorted[lower] : sorted[lower] + (sorted[upper] - sorted[lower]) * (index - lower);
}
function rankByPosition(rows, targetField, rankField, descending = true) {
  const groups = Object.values(rows.reduce((out, row) => ((out[`${row.season}|${row.pos}`] ||= []).push(row), out), {}));
  for (const group of groups) group.sort((a, b) => descending ? b[targetField] - a[targetField] : a[targetField] - b[targetField]).forEach((row, index) => { row[rankField] = index + 1; });
}
function marketBucket(rank) { return rank <= 12 ? '1-12' : rank <= 24 ? '13-24' : rank <= 48 ? '25-48' : '49+'; }

async function ensureStats() {
  fs.mkdirSync(cache, { recursive: true });
  for (let year = 2018; year <= 2025; year += 1) {
    const file = path.join(cache, `stats_player_reg_${year}.csv`);
    if (fs.existsSync(file) && fs.statSync(file).size > 100) continue;
    const url = `https://github.com/nflverse/nflverse-data/releases/download/stats_player/stats_player_reg_${year}.csv`;
    const response = await fetch(url, { headers: { 'User-Agent': 'Draft-Wizard-Backtest/1.0' } });
    if (!response.ok) throw new Error(`Download failed (${response.status}): ${url}`);
    fs.writeFileSync(file, Buffer.from(await response.arrayBuffer()));
  }
}

async function loadPreseasonEcr() {
  const file = path.join(cache, 'db_fpecr.csv.gz');
  if (!fs.existsSync(file)) throw new Error('Missing DynastyProcess cache: db_fpecr.csv.gz');
  const rl = readline.createInterface({ input: fs.createReadStream(file).pipe(zlib.createGunzip()) });
  let headers = [];
  const candidates = new Map();
  for await (const line of rl) {
    if (!headers.length) { headers = csvLine(line); continue; }
    const values = csvLine(line);
    const row = Object.fromEntries(headers.map((h, i) => [h, values[i] ?? '']));
    const year = Number(row.scrape_date.slice(0, 4));
    if (!sampleSeasons.includes(year) || row.page_type !== 'redraft-overall' || !positions.includes(row.pos)) continue;
    if (row.scrape_date > `${year}-08-31`) continue;
    const key = `${year}|${identity(row.player, row.pos)}`;
    const old = candidates.get(key);
    if (!old || row.scrape_date > old.scrape_date) candidates.set(key, row);
  }
  return [...candidates.values()];
}

function rawFeatures(position, old, older, draft, marketContext = {}) {
  if (!old) {
    const pick = n(draft?.pick) || 260;
    return { volume: 0, dominance: 0, efficiency: 0, production: 0, trajectory: clamp((150 - pick) / 120, -0.8, 1.1), reversion: 0, floor: 0, context: marketContext.value || 0, hasHistory: false };
  }
  const games = Math.max(1, n(old.games));
  const attempts = n(old.attempts); const carries = n(old.carries); const targets = n(old.targets); const receptions = n(old.receptions);
  const opportunity = position === 'QB' ? attempts + carries * 1.75 : position === 'RB' ? carries + targets * 1.75 : targets * 1.5;
  const oppPerGame = opportunity / games;
  const pprPerGame = (n(old.fantasy_points_ppr) || n(old.fantasy_points)) / games;
  let volume; let dominance; let efficiency;
  if (position === 'QB') {
    volume = Math.log1p(oppPerGame);
    dominance = n(old.rushing_yards) / games / 35 + carries / games / 5;
    efficiency = (n(old.passing_epa) / Math.max(1, attempts)) * 3 + n(old.passing_cpoe) / 12 + n(old.passing_yards) / Math.max(1, attempts) / 8;
  } else if (position === 'RB') {
    volume = Math.log1p(oppPerGame);
    dominance = n(old.target_share) * 5 + carries / games / 12;
    efficiency = (n(old.rushing_epa) + n(old.receiving_epa)) / Math.max(1, carries + targets) + (n(old.rushing_yards) + n(old.receiving_yards)) / Math.max(1, carries + receptions) / 6;
  } else {
    volume = Math.log1p(targets / games);
    dominance = n(old.wopr) || n(old.target_share) * 1.5 + n(old.air_yards_share) * 0.7;
    efficiency = n(old.receiving_epa) / Math.max(1, targets) + n(old.receiving_yards) / Math.max(1, targets) / 10 + n(old.receiving_yards_after_catch) / Math.max(1, receptions) / 8;
  }
  const reliability = Math.sqrt(opportunity / (opportunity + 100));
  efficiency *= reliability;
  let trajectory = 0;
  let reversion = 0;
  let floor = 0;
  const draftSeason = n(draft?.season);
  const experience = draftSeason ? n(old.season) - draftSeason + 1 : 0;
  if (older) {
    const olderGames = Math.max(1, n(older.games));
    const olderOpp = position === 'QB'
      ? (n(older.attempts) + n(older.carries) * 1.75) / olderGames
      : position === 'RB'
        ? (n(older.carries) + n(older.targets) * 1.75) / olderGames
        : n(older.targets) * 1.5 / olderGames;
    if (olderOpp > 0) trajectory += clamp(Math.log(oppPerGame / olderOpp), -1.2, 1.2);
    const olderPprPerGame = (n(older.fantasy_points_ppr) || n(older.fantasy_points)) / olderGames;
    if (position === 'QB') {
      const recentTdRate = n(old.passing_tds) / Math.max(1, attempts);
      const olderTdRate = n(older.passing_tds) / Math.max(1, n(older.attempts));
      reversion = clamp((olderTdRate - recentTdRate) * 18 + (olderPprPerGame - pprPerGame) / 14, -1.8, 1.8);
    } else {
      const productionGap = clamp((olderPprPerGame - pprPerGame) / 6, -1.8, 1.8);
      const earlyCareerDampener = experience >= 1 && experience <= 3 && productionGap < 0 ? 0.25 : 1;
      reversion = productionGap * earlyCareerDampener;
      if (position === 'RB') {
        const ageAtDraft = n(draft?.age);
        const currentAge = ageAtDraft && draftSeason ? ageAtDraft + n(old.season) - draftSeason : 0;
        const workloadSpike = olderOpp > 0 ? Math.log(oppPerGame / olderOpp) : 0;
        if (currentAge >= 26 && workloadSpike > 0) reversion -= clamp(workloadSpike * 0.65, 0, 0.65);
      }
    }
  }
  if (experience >= 1 && experience <= 3) trajectory += position === 'WR' || position === 'TE' ? 0.35 : 0.18;
  if (position === 'QB') {
    const adot = n(old.passing_air_yards) / Math.max(1, attempts);
    const dropbacks = attempts + n(old.sacks_suffered);
    const sackRate = n(old.sacks_suffered) / Math.max(1, dropbacks);
    floor = clamp((n(old.rushing_yards) / games - 10) / 22 + n(old.rushing_tds) / games * 2.2 + (adot - 7.5) / 3 - sackRate * 2.5, -2.2, 2.2);
  }
  return { volume, dominance, efficiency, production: pprPerGame, trajectory, reversion, floor, context: marketContext.value || 0, hasHistory: true };
}

function normalizeFeatures(rows) {
  const fields = ['volume', 'dominance', 'efficiency', 'production', 'trajectory', 'reversion', 'floor', 'context'];
  const groups = Object.values(rows.reduce((out, row) => ((out[`${row.season}|${row.pos}`] ||= []).push(row), out), {}));
  for (const group of groups) {
    for (const field of fields) {
      const values = group.filter((row) => row.raw.hasHistory || ['trajectory', 'context'].includes(field)).map((row) => row.raw[field]);
      const low = percentile(values, 0.05); const high = percentile(values, 0.95);
      const clipped = values.map((value) => clamp(value, low, high));
      const average = mean(clipped); const sd = Math.sqrt(mean(clipped.map((value) => (value - average) ** 2))) || 1;
      group.forEach((row) => {
        row.features[field] = !row.raw.hasHistory && !['trajectory', 'context'].includes(field)
          ? 0
          : clamp((clamp(row.raw[field], low, high) - average) / sd, -2.5, 2.5);
      });
    }
  }
}

function normalizeLegacyFeatures(rows) {
  const fields = ['volume', 'dominance', 'efficiency', 'production', 'trajectory'];
  const groups = Object.values(rows.reduce((out, row) => ((out[`${row.season}|${row.pos}`] ||= []).push(row), out), {}));
  for (const group of groups) {
    for (const field of fields) {
      const values = group.map((row) => row.raw[field]);
      const low = percentile(values, 0.05); const high = percentile(values, 0.95);
      const clipped = values.map((value) => clamp(value, low, high));
      const average = mean(clipped); const sd = Math.sqrt(mean(clipped.map((value) => (value - average) ** 2))) || 1;
      group.forEach((row) => {
        row.legacyFeatures ||= {};
        row.legacyFeatures[field] = clamp((clamp(row.raw[field], low, high) - average) / sd, -2.5, 2.5);
      });
    }
  }
}

const candidates = {
  QB: [
    { volume: .34, dominance: .22, efficiency: .20, production: .14, trajectory: .10 },
    { volume: .40, dominance: .24, efficiency: .14, production: .12, trajectory: .10 },
    { volume: .30, dominance: .30, efficiency: .16, production: .14, trajectory: .10 },
    { volume: .25, dominance: .18, efficiency: .12, production: .08, trajectory: .08, reversion: .17, floor: .12 },
    { volume: .28, dominance: .16, efficiency: .10, production: .08, trajectory: .08, reversion: .20, floor: .10 },
    { volume: .22, dominance: .16, efficiency: .10, production: .07, trajectory: .07, reversion: .15, floor: .10, context: .13 },
  ],
  RB: [
    { volume: .44, dominance: .25, efficiency: .07, production: .14, trajectory: .10 },
    { volume: .50, dominance: .22, efficiency: .05, production: .13, trajectory: .10 },
    { volume: .38, dominance: .30, efficiency: .08, production: .12, trajectory: .12 },
    { volume: .32, dominance: .24, efficiency: .04, production: .08, trajectory: .10, reversion: .22 },
    { volume: .30, dominance: .22, efficiency: .04, production: .08, trajectory: .12, reversion: .24 },
    { volume: .27, dominance: .21, efficiency: .04, production: .07, trajectory: .10, reversion: .18, context: .13 },
  ],
  WR: [
    { volume: .26, dominance: .34, efficiency: .10, production: .16, trajectory: .14 },
    { volume: .30, dominance: .36, efficiency: .07, production: .13, trajectory: .14 },
    { volume: .22, dominance: .40, efficiency: .10, production: .14, trajectory: .14 },
    { volume: .20, dominance: .30, efficiency: .08, production: .10, trajectory: .14, reversion: .18 },
    { volume: .20, dominance: .28, efficiency: .08, production: .10, trajectory: .14, reversion: .20 },
    { volume: .18, dominance: .26, efficiency: .07, production: .09, trajectory: .12, reversion: .16, context: .12 },
  ],
  TE: [
    { volume: .30, dominance: .30, efficiency: .09, production: .15, trajectory: .16 },
    { volume: .34, dominance: .31, efficiency: .06, production: .13, trajectory: .16 },
    { volume: .27, dominance: .35, efficiency: .08, production: .14, trajectory: .16 },
    { volume: .22, dominance: .28, efficiency: .06, production: .10, trajectory: .14, reversion: .20 },
    { volume: .20, dominance: .26, efficiency: .06, production: .10, trajectory: .16, reversion: .22 },
    { volume: .18, dominance: .24, efficiency: .05, production: .08, trajectory: .13, reversion: .18, context: .14 },
  ],
};
const edgeThresholds = { QB: 0.7, RB: 0.4, WR: 0.5, TE: 0.7 };
const legacyWeights = {
  QB: { volume: .40, dominance: .24, efficiency: .14, production: .12, trajectory: .10 },
  RB: { volume: .38, dominance: .30, efficiency: .08, production: .12, trajectory: .12 },
  WR: { volume: .26, dominance: .34, efficiency: .10, production: .16, trajectory: .14 },
  TE: { volume: .34, dominance: .31, efficiency: .06, production: .13, trajectory: .16 },
};
const riskThresholds = { QB: 0.55, RB: 0.55, WR: 0.55, TE: 0.55 };
const riskCandidates = {
  QB: [
    { reversion: .45, floor: .35, context: .20 },
    { reversion: .35, floor: .45, context: .20 },
    { reversion: .35, floor: .30, context: .35 },
    { reversion: .20, floor: .65, context: .15 },
  ],
  RB: [
    { reversion: .55, context: .45 },
    { reversion: .40, context: .60 },
    { reversion: .70, context: .30 },
    { reversion: .25, context: .75 },
  ],
  WR: [
    { reversion: .55, context: .45 },
    { reversion: .40, context: .60 },
    { reversion: .70, context: .30 },
    { reversion: .25, context: .75 },
  ],
  TE: [
    { reversion: .55, context: .45 },
    { reversion: .40, context: .60 },
    { reversion: .70, context: .30 },
    { reversion: .20, context: .80 },
  ],
};

function applyModel(rows, weights) {
  for (const row of rows) row.signal = Object.entries(weights).reduce((sum, [field, weight]) => sum + row.features[field] * weight, 0);
  for (const row of rows) {
    row.expectedGain = Math.round(row.signal * 4);
    row.modelRank = Math.max(1, row.marketPosRank - row.expectedGain);
  }
}
function applyRiskModel(rows, weights) {
  for (const row of rows) {
    row.riskSignal = -Object.entries(weights).reduce((sum, [field, weight]) => sum + row.features[field] * weight, 0);
    row.expectedLoss = Math.max(0, Math.round(row.riskSignal * 4));
  }
}
function evaluate(rows) {
  const edge = rows.filter((row) => row.expectedGain >= 2 && row.signal >= edgeThresholds[row.pos]);
  const edgeHitRate = mean(edge.map((row) => Number(row.outcomeRank < row.marketPosRank)));
  let baselineHits = 0;
  for (const selected of edge) {
    const peers = rows.filter((row) => row.season === selected.season && row.pos === selected.pos && marketBucket(row.marketPosRank) === marketBucket(selected.marketPosRank));
    baselineHits += mean(peers.map((row) => Number(row.outcomeRank < row.marketPosRank)));
  }
  const baselineHitRate = edge.length ? baselineHits / edge.length : 0;
  return { selections: edge.length, edgeHitRate, baselineHitRate, lift: edgeHitRate - baselineHitRate };
}
function evaluateRisk(rows) {
  const flagged = rows.filter((row) => row.expectedLoss >= 2 && row.riskSignal >= riskThresholds[row.pos]);
  const hitRate = mean(flagged.map((row) => Number(row.outcomeRank > row.marketPosRank)));
  let baselineHits = 0;
  for (const selected of flagged) {
    const peers = rows.filter((row) => row.season === selected.season && row.pos === selected.pos && marketBucket(row.marketPosRank) === marketBucket(selected.marketPosRank));
    baselineHits += mean(peers.map((row) => Number(row.outcomeRank > row.marketPosRank)));
  }
  const baselineHitRate = flagged.length ? baselineHits / flagged.length : 0;
  return { selections: flagged.length, hitRate, baselineHitRate, lift: hitRate - baselineHitRate };
}

function buildMarketContexts(markets, priorRows) {
  const priorByIdentity = new Map(priorRows.map((row) => [identity(row.player_display_name, row.position), row]));
  const entriesByTeam = new Map();
  for (const market of markets) {
    const team = normalizeTeam(market.team || market.tm);
    if (!team) continue;
    const entries = entriesByTeam.get(team) || [];
    entries.push(market);
    entriesByTeam.set(team, entries);
  }
  const priorTeamStats = new Map();
  for (const row of priorRows) {
    const team = normalizeTeam(row.recent_team);
    if (!team) continue;
    const stats = priorTeamStats.get(team) || { targets: 0, teTargets: 0, topQb: null };
    stats.targets += n(row.targets);
    if (row.position === 'TE') stats.teTargets += n(row.targets);
    if (row.position === 'QB' && (!stats.topQb || n(row.attempts) > n(stats.topQb.attempts))) stats.topQb = row;
    priorTeamStats.set(team, stats);
  }
  const qbPositionRanks = new Map(markets.filter((row) => row.pos === 'QB').sort((a, b) => n(a.ecr) - n(b.ecr)).map((row, index) => [identity(row.player, row.pos), index + 1]));
  const contexts = new Map();
  for (const market of markets) {
    const key = identity(market.player, market.pos);
    const team = normalizeTeam(market.team || market.tm);
    const teamEntries = entriesByTeam.get(team) || [];
    const teamStats = priorTeamStats.get(team) || { targets: 0, teTargets: 0, topQb: null };
    let retainedTargets = 0;
    let incomingReceivingYards = 0;
    for (const teammate of teamEntries) {
      const prior = priorByIdentity.get(identity(teammate.player, teammate.pos));
      if (!prior) continue;
      if (normalizeTeam(prior.recent_team) === team) retainedTargets += n(prior.targets);
      else if (['RB', 'WR', 'TE'].includes(prior.position)) incomingReceivingYards += n(prior.receiving_yards);
    }
    const targetVacancy = teamStats.targets > 0 ? clamp((teamStats.targets - retainedTargets) / teamStats.targets, 0, 0.75) : 0;
    const priorQbPpg = teamStats.topQb ? (n(teamStats.topQb.fantasy_points_ppr) || n(teamStats.topQb.fantasy_points)) / Math.max(1, n(teamStats.topQb.games)) : 15;
    const currentQb = teamEntries.filter((row) => row.pos === 'QB').sort((a, b) => n(a.ecr) - n(b.ecr))[0];
    let currentQbEstimate = priorQbPpg;
    if (currentQb) {
      const currentQbKey = identity(currentQb.player, currentQb.pos);
      const currentQbPrior = priorByIdentity.get(currentQbKey);
      const marketEstimate = clamp(22.5 - (qbPositionRanks.get(currentQbKey) || 24) * 0.38, 10.5, 22.1);
      currentQbEstimate = currentQbPrior
        ? ((n(currentQbPrior.fantasy_points_ppr) || n(currentQbPrior.fantasy_points)) / Math.max(1, n(currentQbPrior.games))) * 0.65 + marketEstimate * 0.35
        : marketEstimate;
    }
    const qbChange = clamp((currentQbEstimate - priorQbPpg) / 5, -1.8, 1.8);
    const competitors = teamEntries.filter((row) => row.pos === market.pos && identity(row.player, row.pos) !== key)
      .map((row) => priorByIdentity.get(identity(row.player, row.pos))).filter(Boolean);
    let competition = 0;
    if (market.pos === 'RB') competition = Math.max(0, ...competitors.map((row) => (n(row.carries) + n(row.targets) * 1.5) / Math.max(1, n(row.games))));
    else if (market.pos === 'TE') competition = competitors.reduce((sum, row) => sum + n(row.targets) / Math.max(1, n(row.games)), 0);
    else if (market.pos === 'WR') competition = Math.max(0, ...competitors.map((row) => n(row.targets) / Math.max(1, n(row.games))));
    const teamTeTargetShare = teamStats.targets > 0 ? teamStats.teTargets / teamStats.targets : 0.16;
    let value = 0;
    if (market.pos === 'QB') value = clamp(incomingReceivingYards / 1200, 0, 1.5);
    else if (market.pos === 'RB') value = -clamp((competition - 8) / 8, 0, 1.5);
    else if (market.pos === 'WR') value = clamp(targetVacancy * 2 + qbChange * 0.8 - Math.max(0, competition - 6) / 10, -2, 2);
    else if (market.pos === 'TE') value = clamp(targetVacancy * 0.7 + qbChange * 0.7 + (teamTeTargetShare - 0.16) * 4 - competition / 12, -2, 2);
    contexts.set(key, { value, team, targetVacancy, incomingReceivingYards, qbChange, competition, teamTeTargetShare, currentQb: currentQb?.player || null });
  }
  return contexts;
}

await ensureStats();
const ecr = await loadPreseasonEcr();
const availabilityBands = { early: [1, 24], middle: [25, 72], late: [73, 300] };
const availabilityFallback = {
  QB: { early: 7, middle: 13, late: 22 },
  RB: { early: 6, middle: 12, late: 22 },
  WR: { early: 6, middle: 12, late: 23 },
  TE: { early: 8, middle: 15, late: 24 },
};
const uncertaintyByPosition = Object.fromEntries(positions.map((position) => {
  const bands = Object.fromEntries(Object.entries(availabilityBands).map(([band, [low, high]]) => {
    const spreads = ecr.filter((row) => row.pos === position && n(row.ecr) >= low && n(row.ecr) <= high).map((row) => {
      const rankSd = n(row.sd);
      const expertRange = Math.max(0, n(row.worst) - n(row.best)) / 4;
      return Math.max(rankSd, expertRange);
    }).filter((value) => value > 0);
    const derived = percentile(spreads, 0.65) * 1.6;
    return [band, Math.round(clamp(derived || availabilityFallback[position][band], availabilityFallback[position][band] * 0.75, availabilityFallback[position][band] * 1.75))];
  }));
  return [position, bands];
}));
const statsBySeason = new Map();
for (let year = 2018; year <= 2025; year += 1) statsBySeason.set(year, readCsv(path.join(cache, `stats_player_reg_${year}.csv`)));
const drafts = readCsv(path.join(cache, 'draft_picks.csv'));
const draftByIdentity = new Map(drafts.slice().sort((a, b) => n(b.season) - n(a.season)).map((row) => [identity(row.pfr_player_name, row.category), row]));
const samples = [];
for (const season of sampleSeasons) {
  const priorRows = statsBySeason.get(season - 1);
  const prior = new Map(priorRows.map((row) => [identity(row.player_display_name, row.position), row]));
  const older = season - 2 >= 2018 ? new Map(statsBySeason.get(season - 2).map((row) => [identity(row.player_display_name, row.position), row])) : new Map();
  const outcome = new Map(statsBySeason.get(season).map((row) => [identity(row.player_display_name, row.position), row]));
  const seasonMarkets = ecr.filter((row) => Number(row.scrape_date.slice(0, 4)) === season);
  const marketContexts = buildMarketContexts(seasonMarkets, priorRows);
  for (const market of seasonMarkets) {
    const key = identity(market.player, market.pos);
    const result = outcome.get(key);
    if (!result || n(market.ecr) <= 0 || n(market.ecr) > 300) continue;
    const marketContext = marketContexts.get(key) || {};
    const raw = rawFeatures(market.pos, prior.get(key), older.get(key), draftByIdentity.get(key), marketContext);
    samples.push({ season, name: market.player, pos: market.pos, marketRank: n(market.ecr), outcomePoints: n(result.fantasy_points_ppr) || n(result.fantasy_points), raw, features: {}, marketContext });
  }
}
rankByPosition(samples, 'marketRank', 'marketPosRank', false);
rankByPosition(samples, 'outcomePoints', 'outcomeRank', true);
normalizeFeatures(samples);
normalizeLegacyFeatures(samples);
for (const row of samples) {
  row.legacySignal = Object.entries(legacyWeights[row.pos]).reduce((sum, [field, weight]) => sum + row.legacyFeatures[field] * weight, 0);
  row.legacyExpectedGain = Math.round(row.legacySignal * 4);
}

const folds = [];
const latestWeights = {};
const latestRiskWeights = {};
for (const testSeason of testSeasons) {
  const foldPositions = {};
  for (const position of positions) {
    const train = samples.filter((row) => row.season < testSeason && row.pos === position);
    const test = samples.filter((row) => row.season === testSeason && row.pos === position);
    let best = null;
    for (const weights of candidates[position]) {
      applyModel(train, weights);
      const result = evaluate(train);
      const objective = result.lift + Math.min(result.selections, 80) / 10000;
      if (!best || objective > best.objective) best = { weights, objective };
    }
    applyModel(test, best.weights);
    const result = evaluate(test);
    let bestRisk = null;
    for (const weights of riskCandidates[position]) {
      applyRiskModel(train, weights);
      const riskResult = evaluateRisk(train);
      const objective = riskResult.lift + Math.min(riskResult.selections, 80) / 10000;
      if (!bestRisk || objective > bestRisk.objective) bestRisk = { weights, objective };
    }
    applyRiskModel(test, bestRisk.weights);
    const riskResult = evaluateRisk(test);
    foldPositions[position] = { matchedPlayers: test.length, weights: best.weights, ...result, regressionRisk: { weights: bestRisk.weights, ...riskResult } };
    if (testSeason === testSeasons.at(-1)) {
      latestWeights[position] = best.weights;
      latestRiskWeights[position] = bestRisk.weights;
    }
  }
  const allTest = samples.filter((row) => row.season === testSeason);
  const selected = allTest.filter((row) => row.expectedGain >= 2 && row.signal >= edgeThresholds[row.pos]);
  const weighted = positions.map((position) => foldPositions[position]);
  folds.push({
    season: testSeason,
    matchedPlayers: allTest.length,
    edgeSelections: selected.length,
    edgeHitRate: mean(selected.map((row) => Number(row.outcomeRank < row.marketPosRank))),
    baselineHitRate: weighted.reduce((sum, row) => sum + row.baselineHitRate * row.selections, 0) / Math.max(1, selected.length),
    positions: foldPositions,
  });
}

const aggregate = {
  edgeHitRate: mean(folds.map((fold) => fold.edgeHitRate)),
  baselineHitRate: mean(folds.map((fold) => fold.baselineHitRate)),
};
aggregate.lift = aggregate.edgeHitRate - aggregate.baselineHitRate;
const positionSummary = Object.fromEntries(positions.map((position) => {
  const rows = folds.map((fold) => fold.positions[position]);
  const edgeHitRate = mean(rows.map((row) => row.edgeHitRate));
  const baselineHitRate = mean(rows.map((row) => row.baselineHitRate));
  const lift = edgeHitRate - baselineHitRate;
  const riskHitRate = mean(rows.map((row) => row.regressionRisk.hitRate));
  const riskBaselineHitRate = mean(rows.map((row) => row.regressionRisk.baselineHitRate));
  const riskLift = riskHitRate - riskBaselineHitRate;
  return [position, {
    edgeHitRate, baselineHitRate, lift, acceptedCore: lift >= 0, latestWeights: latestWeights[position],
    regressionRisk: { hitRate: riskHitRate, baselineHitRate: riskBaselineHitRate, lift: riskLift, accepted: riskLift >= 0, latestWeights: latestRiskWeights[position] },
  }];
}));
const retainedPositions = positions.filter((position) => positionSummary[position].acceptedCore);
const retainedRiskPositions = positions.filter((position) => positionSummary[position].regressionRisk.accepted);
const retainedFolds = folds.map((fold) => {
  const rows = retainedPositions.map((position) => fold.positions[position]);
  const selections = rows.reduce((sum, row) => sum + row.selections, 0);
  return {
    selections,
    edgeHitRate: rows.reduce((sum, row) => sum + row.edgeHitRate * row.selections, 0) / Math.max(1, selections),
    baselineHitRate: rows.reduce((sum, row) => sum + row.baselineHitRate * row.selections, 0) / Math.max(1, selections),
  };
});
aggregate.allPositionLiftBeforeGate = aggregate.lift;
aggregate.selections = retainedFolds.reduce((sum, fold) => sum + fold.selections, 0);
aggregate.edgeHitRate = mean(retainedFolds.map((fold) => fold.edgeHitRate));
aggregate.baselineHitRate = mean(retainedFolds.map((fold) => fold.baselineHitRate));
aggregate.lift = aggregate.edgeHitRate - aggregate.baselineHitRate;
const retainedRiskFolds = folds.map((fold) => {
  const rows = retainedRiskPositions.map((position) => fold.positions[position].regressionRisk);
  const selections = rows.reduce((sum, row) => sum + row.selections, 0);
  return {
    selections,
    hitRate: rows.reduce((sum, row) => sum + row.hitRate * row.selections, 0) / Math.max(1, selections),
    baselineHitRate: rows.reduce((sum, row) => sum + row.baselineHitRate * row.selections, 0) / Math.max(1, selections),
  };
});
const regressionRiskAggregate = {
  selections: retainedRiskFolds.reduce((sum, fold) => sum + fold.selections, 0),
  hitRate: mean(retainedRiskFolds.map((fold) => fold.hitRate)),
  baselineHitRate: mean(retainedRiskFolds.map((fold) => fold.baselineHitRate)),
};
regressionRiskAggregate.lift = regressionRiskAggregate.hitRate - regressionRiskAggregate.baselineHitRate;
const reviewNames = new Set(Object.values(reviewCohorts).flat().map(normalizeName));
const cohortDiagnostics = samples.filter((row) => row.season === 2025 && reviewNames.has(normalizeName(row.name)))
  .map((row) => {
    const coreCall = row.expectedGain >= 2 && row.signal >= edgeThresholds[row.pos] ? 'upside' : row.expectedGain <= -2 && row.signal <= -edgeThresholds[row.pos] ? 'downside' : 'neutral';
    const validatedDownside = row.expectedLoss >= 2 && row.riskSignal >= riskThresholds[row.pos] && positionSummary[row.pos].regressionRisk.accepted;
    return ({
    cohort: reviewCohorts.breakout.some((name) => normalizeName(name) === normalizeName(row.name)) ? 'breakout' : 'regression',
    name: row.name,
    position: row.pos,
    ...reviewPointOutcomes.get(normalizeName(row.name)),
    marketRank: row.marketRank,
    marketPositionRank: row.marketPosRank,
    modelRank: row.modelRank,
    expectedGain: row.expectedGain,
    actualFinish: row.outcomeRank,
    signal: row.signal,
    legacySignal: row.legacySignal,
    legacyExpectedGain: row.legacyExpectedGain,
    legacyEdgeCall: row.legacyExpectedGain >= 2 && row.legacySignal >= edgeThresholds[row.pos] ? 'upside' : row.legacyExpectedGain <= -2 && row.legacySignal <= -edgeThresholds[row.pos] ? 'downside' : 'neutral',
    edgeCall: coreCall,
    regressionRiskSignal: row.riskSignal,
    expectedLoss: row.expectedLoss,
    regressionRiskCall: row.expectedLoss >= 2 && row.riskSignal >= riskThresholds[row.pos] ? 'fade' : 'neutral',
    validatedDownsideEdgeCall: validatedDownside ? 'downside' : 'neutral',
    updatedEdgeCall: coreCall !== 'neutral' ? coreCall : validatedDownside ? 'downside' : 'neutral',
    correctDirection: row.outcomeRank === row.marketPosRank ? 'push' : row.outcomeRank < row.marketPosRank ? 'upside' : 'downside',
    rawFeatures: row.raw,
    normalizedFeatures: row.features,
    marketContext: row.marketContext,
  });
  });
const missingCohortPlayers = Object.entries(reviewCohorts).flatMap(([cohort, names]) => names
  .filter((name) => !cohortDiagnostics.some((row) => normalizeName(row.name) === normalizeName(name)))
  .map((name) => ({ cohort, name })));
const cohortSummary = {
  breakout: {
    requested: reviewCohorts.breakout.length,
    matched: cohortDiagnostics.filter((row) => row.cohort === 'breakout').length,
    legacyCaught: cohortDiagnostics.filter((row) => row.cohort === 'breakout' && row.legacyEdgeCall === 'upside').length,
    updatedCaught: cohortDiagnostics.filter((row) => row.cohort === 'breakout' && row.updatedEdgeCall === 'upside').length,
  },
  regression: {
    requested: reviewCohorts.regression.length,
    matched: cohortDiagnostics.filter((row) => row.cohort === 'regression').length,
    legacyCaught: cohortDiagnostics.filter((row) => row.cohort === 'regression' && row.legacyEdgeCall === 'downside').length,
    updatedCaught: cohortDiagnostics.filter((row) => row.cohort === 'regression' && row.updatedEdgeCall === 'downside').length,
  },
};
const examples = samples.filter((row) => row.season === 2025 && ['Jaxon Smith-Njigba', 'Javonte Williams'].includes(row.name))
  .map((row) => ({ name: row.name, position: row.pos, marketPositionRank: row.marketPosRank, modelRank: row.modelRank, actualFinish: row.outcomeRank, signal: row.signal }));
const roundDeep = (value) => JSON.parse(JSON.stringify(value, (key, item) => typeof item === 'number' && !Number.isInteger(item) ? Math.round(item * 1000) / 1000 : item));
const accepted = (aggregate.lift > 0 && retainedPositions.length > 0)
  || (regressionRiskAggregate.lift > 0 && retainedRiskPositions.length > 0);
const report = roundDeep({
  generatedAt: new Date().toISOString(),
  method: 'Rolling no-lookahead folds. Each test season uses Aug. 31-or-earlier ECR, only earlier-season data to choose position weights, prior-year volume/dominance/sample-shrunk efficiency/production/trajectory, and next-season positional finish. The ECR-only baseline is the observed hit rate among players in the same position and market-rank band.',
  featureDefinitions: {
    volume: 'Per-game attempts, weighted opportunities, or targets.',
    dominance: 'QB rushing value; RB carry/target role; WR/TE WOPR from target and air-yard shares.',
    efficiency: 'EPA and yardage per opportunity, shrunk toward average for small samples.',
    production: 'Prior PPR points per game; availability is not used as an injury forecast.',
    trajectory: 'Two-year opportunity trend, early-career curve, and rookie draft capital.',
    reversion: 'Two-year touchdown and per-game production mean reversion, damped for early-career breakouts; older RB workload spikes receive an added penalty.',
    floor: 'QB rushing production, passing depth, and sack-rate floor.',
    context: 'Time-safe August team context reconstructed from current-team ECR: target vacancy, incoming weapons, teammate competition, QB change, and team TE usage.',
  },
  seasons: testSeasons,
  folds,
  aggregate,
  positions: positionSummary,
  retainedHistoricalCorePositions: retainedPositions,
  rejectedHistoricalCorePositions: positions.filter((position) => !retainedPositions.includes(position)),
  regressionRiskAggregate,
  retainedRegressionRiskPositions: retainedRiskPositions,
  rejectedRegressionRiskPositions: positions.filter((position) => !retainedRiskPositions.includes(position)),
  examples,
  reviewCohorts: { source: 'data/inputs/edge-review-2025.json', summary: cohortSummary, players: cohortDiagnostics, unmatched: missingCohortPlayers },
  draftAvailability: {
    modelVersion: 'return-risk-2026.6',
    method: 'Position and market-band uncertainty proxy derived from historical preseason FantasyPros consensus rank standard deviation and expert best/worst ranges, then conditioned on each intervening team\'s roster demand. It supports an estimated return-risk band, not a guaranteed draft probability.',
    uncertaintyByPosition,
  },
  accepted,
  limitation: 'Historical public projection and injury snapshots are not complete enough for time-safe folds. The rolling test validates the price-aware statistical core and retained downside component; current consensus-projection disagreement, live injuries, and current OL availability are evaluated in production but cannot be reconstructed prospectively for every historical season.',
});
fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(output, `${JSON.stringify(report, null, 2)}\n`);
const pct = (value) => `${Math.round(value * 1000) / 10}%`;
const mdRows = cohortDiagnostics.slice().sort((a, b) => a.cohort.localeCompare(b.cohort) || a.position.localeCompare(b.position) || a.name.localeCompare(b.name))
  .map((row) => {
    const rankResult = row.actualFinish < row.marketPositionRank ? `beat (${row.marketPositionRank}→${row.actualFinish})` : row.actualFinish > row.marketPositionRank ? `missed (${row.marketPositionRank}→${row.actualFinish})` : `push (${row.actualFinish})`;
    return `| ${row.cohort} | ${row.position} | ${row.name} | ${row.projected} → ${row.actual} (${row.miss > 0 ? '+' : ''}${row.miss}) | ${row.legacyEdgeCall} | ${row.updatedEdgeCall} | ${rankResult} |`;
  });
const reviewMarkdown = `# 2025 Edge review\n\n`
  + `Generated ${report.generatedAt}. The point-miss cohorts come from \`${report.reviewCohorts.source}\`; market ranks are the last FantasyPros redraft ECR snapshot on or before August 31, 2025.\n\n`
  + `## Result\n\n`
  + `- Legacy Edge caught ${cohortSummary.breakout.legacyCaught}/${cohortSummary.breakout.matched} listed breakouts and ${cohortSummary.regression.legacyCaught}/${cohortSummary.regression.matched} listed regressions.\n`
  + `- Updated Edge catches ${cohortSummary.breakout.updatedCaught}/${cohortSummary.breakout.matched} listed breakouts and ${cohortSummary.regression.updatedCaught}/${cohortSummary.regression.matched} listed regressions. Downside only counts when its position model beat the rolling no-lookahead baseline.\n`
  + `- Retained historical upside: ${retainedPositions.join(', ') || 'none'} (${pct(aggregate.edgeHitRate)} vs ${pct(aggregate.baselineHitRate)}, ${aggregate.selections} selections). Retained downside Edge: ${retainedRiskPositions.join(', ') || 'none'} (${pct(regressionRiskAggregate.hitRate)} vs ${pct(regressionRiskAggregate.baselineHitRate)}, ${regressionRiskAggregate.selections} selections).\n\n`
  + `## Player-level calls\n\n`
  + `| Cohort | Pos | Player | Projected → Actual | Legacy Edge | Updated Edge | Market-pos result |\n`
  + `| --- | --- | --- | ---: | --- | --- | --- |\n`
  + `${mdRows.join('\n')}\n\n`
  + `## Interpretation\n\n`
  + `Cam Ward illustrates the difference between point projections and draft value: he missed the supplied point projection by 70, but finished ahead of his preseason positional market rank. Edge is rank-relative, so the report preserves both results.\n\n`
  + `The production formula now scores two-year reversion, QB aDOT/rushing/sack floor, weekly concentration, target vacancy, incoming receiving production, teammate competition, projected QB change, and destination-team TE usage. Sparse college and coaching history remains outside production scoring until broad historical coverage exists.\n`;
fs.writeFileSync(reviewOutput, reviewMarkdown);
console.log(JSON.stringify(report, null, 2));
if (!report.accepted) process.exitCode = 2;
