import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import XLSX from 'xlsx';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const args = process.argv.slice(2);
const option = (name, fallback) => {
  const i = args.indexOf(name);
  return i >= 0 && args[i + 1] ? args[i + 1] : fallback;
};
const season = Number(option('--season', '2026'));
const workbookPath = path.resolve(ROOT, option('--workbook', `Fantasy Football Cheat Sheet with Boom Outlier ${season}.xlsx`));
const cacheDir = path.resolve(ROOT, option('--cache', 'data/cache'));
const outputPath = path.resolve(ROOT, option('--output', `public/data/adp-edge-${season}.json`));
const reportPath = path.resolve(ROOT, option('--report', `data/reports/unmatched-${season}.json`));
const backtestPath = path.resolve(ROOT, 'data/reports/backtest-2026.json');
const backtest = fs.existsSync(backtestPath) ? JSON.parse(fs.readFileSync(backtestPath, 'utf8')) : null;
const refresh = args.includes('--refresh');
const priorSeason = season - 1;

const SOURCES = {
  workbook: {
    label: `Fantasy Football Cheat Sheet with Boom Outlier ${season}`,
    url: 'https://www.reddit.com/r/fantasyfootballadvice/comments/1vndd5x/created_a_free_cheat_sheet_with_everything/',
    license: 'Creator-supplied workbook; used locally with attribution.',
  },
  nflverseStats: {
    label: `nflverse player statistics (${priorSeason})`,
    url: 'https://nflreadr.nflverse.com/reference/load_player_stats.html',
    license: 'CC BY 4.0',
  },
  nflverseWeekly: {
    label: `nflverse weekly player statistics (${priorSeason})`,
    url: 'https://nflreadr.nflverse.com/reference/load_player_stats.html',
    license: 'CC BY 4.0',
  },
  nflverseNgs: {
    label: 'NFL Next Gen Stats via nflverse',
    url: 'https://nflreadr.nflverse.com/articles/dictionary_nextgen_stats.html',
    license: 'Public NFL NGS data redistributed by nflverse.',
  },
  nflverseRoster: {
    label: `nflverse rosters (${season})`,
    url: 'https://github.com/nflverse/nflverse-data/releases/tag/rosters',
    license: 'CC BY 4.0',
  },
  nflverseSnaps: {
    label: `nflverse snap counts (${priorSeason})`,
    url: 'https://github.com/nflverse/nflverse-data/releases/tag/snap_counts',
    license: 'CC BY 4.0',
  },
  nflverseDraft: {
    label: 'nflverse draft picks',
    url: 'https://github.com/nflverse/nflverse-data/releases/tag/draft_picks',
    license: 'CC BY 4.0',
  },
  dynastyProcess: {
    label: 'DynastyProcess historical FantasyPros ECR archive',
    url: 'https://github.com/DynastyProcess/data',
    license: 'Open data; source providers retain their respective rights.',
  },
  sleeperProjections: {
    label: `Sleeper-hosted season projections (${season})`,
    url: `https://api.sleeper.com/projections/nfl/${season}?season_type=regular`,
    license: 'Publicly accessible projection feed; underlying projection provider is identified in each record.',
  },
  sleeperPlayers: {
    label: 'Sleeper current NFL player/injury directory',
    url: 'https://api.sleeper.app/v1/players/nfl',
    license: 'Public Sleeper API data.',
  },
  fantasyProsProjections: {
    label: `FantasyPros consensus projections (${season})`,
    url: 'https://www.fantasypros.com/nfl/projections/',
    license: 'Public consensus projection tables; source sites and update dates are identified by FantasyPros.',
  },
};

const downloads = {
  [`stats_player_reg_${priorSeason}.csv`]: `https://github.com/nflverse/nflverse-data/releases/download/stats_player/stats_player_reg_${priorSeason}.csv`,
  [`stats_player_reg_${priorSeason - 1}.csv`]: `https://github.com/nflverse/nflverse-data/releases/download/stats_player/stats_player_reg_${priorSeason - 1}.csv`,
  [`stats_player_week_${priorSeason}.csv`]: `https://github.com/nflverse/nflverse-data/releases/download/stats_player/stats_player_week_${priorSeason}.csv`,
  [`snap_counts_${priorSeason}.csv`]: `https://github.com/nflverse/nflverse-data/releases/download/snap_counts/snap_counts_${priorSeason}.csv`,
  [`roster_${season}.csv`]: `https://github.com/nflverse/nflverse-data/releases/download/rosters/roster_${season}.csv`,
  'draft_picks.csv': 'https://github.com/nflverse/nflverse-data/releases/download/draft_picks/draft_picks.csv',
  'ngs_passing.csv.gz': 'https://github.com/nflverse/nflverse-data/releases/download/nextgen_stats/ngs_passing.csv.gz',
  'ngs_receiving.csv.gz': 'https://github.com/nflverse/nflverse-data/releases/download/nextgen_stats/ngs_receiving.csv.gz',
  'ngs_rushing.csv.gz': 'https://github.com/nflverse/nflverse-data/releases/download/nextgen_stats/ngs_rushing.csv.gz',
  [`sleeper_projections_${season}.json`]: `https://api.sleeper.com/projections/nfl/${season}?season_type=regular`,
  [`sleeper_players_${season}.json`]: 'https://api.sleeper.app/v1/players/nfl',
  ...Object.fromEntries(['qb', 'rb', 'wr', 'te'].map((position) => [
    `fantasypros_${position}_${season}.html`,
    `https://www.fantasypros.com/nfl/projections/${position}.php?week=draft`,
  ])),
};

function normalizeName(value) {
  const normalized = String(value || '')
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\b(jr|sr|ii|iii|iv)\b/g, '')
    .replace(/[^a-z0-9]/g, '');
  return ({ kennethgainwell: 'kennygainwell' })[normalized] || normalized;
}
function normalizeTeam(value) {
  const team = String(value || '').toUpperCase().trim();
  return ({ JAC: 'JAX', LAR: 'LA', WSH: 'WAS', LVR: 'LV', GBP: 'GB', KCC: 'KC', NOS: 'NO', SFO: 'SF', TBB: 'TB', NEP: 'NE' })[team] || team;
}
function normalizePosition(value) {
  const position = String(value || '').toUpperCase().trim();
  if (position === 'HB' || position === 'FB') return 'RB';
  if (position === 'DEF') return 'DST';
  return position;
}
function identityKey(name, position) {
  return `${normalizeName(name)}|${normalizePosition(position)}`;
}
function sleeperFantasyPosition(player) {
  const positions = Array.isArray(player?.fantasy_positions) ? player.fantasy_positions.map(normalizePosition) : [];
  return positions.find((position) => ['QB', 'RB', 'WR', 'TE', 'K'].includes(position)) || normalizePosition(player?.position);
}
function num(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(String(value).replace(/[$,%]/g, '').trim());
  return Number.isFinite(n) ? n : null;
}
function flag(value) {
  const t = String(value || '').trim().toLowerCase();
  return Boolean(t && t !== '-' && t !== 'no' && t !== 'false' && t !== '0');
}
function clamp(value, low, high) { return Math.max(low, Math.min(high, value)); }
function round(value, digits = 1) {
  const m = 10 ** digits;
  return Math.round(value * m) / m;
}
function mean(values) {
  const valid = values.filter(Number.isFinite);
  return valid.length ? valid.reduce((a, b) => a + b, 0) / valid.length : null;
}
function percentile(values, probability) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const index = (sorted.length - 1) * probability;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  if (lower === upper) return sorted[lower];
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (index - lower);
}
function textTier(value, map, fallback = 0) {
  const t = String(value || '').toUpperCase().trim();
  for (const [needle, score] of map) if (t.includes(needle)) return score;
  return fallback;
}
function csvParse(text) {
  const rows = [];
  let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const c = text[i];
    if (quoted) {
      if (c === '"' && text[i + 1] === '"') { field += '"'; i += 1; }
      else if (c === '"') quoted = false;
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
    else field += c;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  const headers = rows.shift() || [];
  return rows.filter((r) => r.some(Boolean)).map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ''])));
}
function readCsv(file) {
  const buffer = fs.readFileSync(file);
  const text = file.endsWith('.gz') ? zlib.gunzipSync(buffer).toString('utf8') : buffer.toString('utf8');
  return csvParse(text);
}
async function ensureData() {
  fs.mkdirSync(cacheDir, { recursive: true });
  for (const [name, url] of Object.entries(downloads)) {
    const target = path.join(cacheDir, name);
    if (!refresh && fs.existsSync(target) && fs.statSync(target).size > 100) continue;
    const response = await fetch(url, { headers: { 'User-Agent': 'Draft-Wizard-Edge-Builder/1.0' } });
    if (!response.ok) throw new Error(`Download failed (${response.status}): ${url}`);
    const temp = `${target}.download`;
    fs.writeFileSync(temp, Buffer.from(await response.arrayBuffer()));
    if (fs.statSync(temp).size < 100) throw new Error(`Downloaded file is unexpectedly small: ${name}`);
    fs.renameSync(temp, target);
  }
}
function indexRows(rows, fields) {
  const out = new Map();
  for (const row of rows) {
    for (const field of fields) {
      const key = normalizeName(row[field]);
      if (key && !out.has(key)) out.set(key, row);
    }
  }
  return out;
}
function indexRowsByIdentity(rows, fields, getPosition) {
  const out = new Map();
  for (const row of rows) {
    const position = normalizePosition(getPosition(row));
    if (!position) continue;
    for (const field of fields) {
      const key = identityKey(row[field], position);
      if (normalizeName(row[field]) && !out.has(key)) out.set(key, row);
    }
  }
  return out;
}
function newestNgs(rows, expectedPosition) {
  const filtered = rows.filter((r) => Number(r.season) === priorSeason && r.season_type === 'REG' && Number(r.week) === 0);
  return indexRowsByIdentity(filtered, ['player_display_name', 'player_short_name'], (row) => row.player_position || expectedPosition);
}
function decodeHtml(value) {
  return String(value || '')
    .replaceAll('&amp;', '&').replaceAll('&#039;', "'").replaceAll('&#39;', "'")
    .replaceAll('&quot;', '"').replaceAll('&nbsp;', ' ').replace(/<[^>]+>/g, '').trim();
}
function parseFantasyProsProjections(html, position) {
  const rows = [];
  for (const match of html.matchAll(/<tr class="mpb-player-[\s\S]*?<\/tr>/gi)) {
    const rowHtml = match[0];
    const nameMatch = rowHtml.match(/fp-player-name="([^"]+)"/i);
    if (!nameMatch) continue;
    const teamMatch = rowHtml.match(/<\/a>\s*([A-Z]{2,3})\s*<\/td>/i);
    const values = [...rowHtml.matchAll(/<td class="center"[^>]*>([\s\S]*?)<\/td>/gi)]
      .map((item) => num(decodeHtml(item[1]).replaceAll(',', '')));
    if (position === 'QB' && values.length >= 10) {
      rows.push({
        name: decodeHtml(nameMatch[1]), position, team: teamMatch?.[1] || '',
        passAtt: values[0], passCmp: values[1], passYards: values[2], passTds: values[3], interceptions: values[4],
        rushAtt: values[5], rushYards: values[6], rushTds: values[7], fumblesLost: values[8], standardPoints: values[9], receptions: 0,
      });
    } else if (position === 'RB' && values.length >= 8) {
      rows.push({
        name: decodeHtml(nameMatch[1]), position, team: teamMatch?.[1] || '',
        rushAtt: values[0], rushYards: values[1], rushTds: values[2], receptions: values[3],
        receivingYards: values[4], receivingTds: values[5], fumblesLost: values[6], standardPoints: values[7],
      });
    } else if (position === 'WR' && values.length >= 8) {
      rows.push({
        name: decodeHtml(nameMatch[1]), position, team: teamMatch?.[1] || '',
        receptions: values[0], receivingYards: values[1], receivingTds: values[2],
        rushAtt: values[3], rushYards: values[4], rushTds: values[5], fumblesLost: values[6], standardPoints: values[7],
      });
    } else if (position === 'TE' && values.length >= 5) {
      rows.push({
        name: decodeHtml(nameMatch[1]), position, team: teamMatch?.[1] || '',
        receptions: values[0], receivingYards: values[1], receivingTds: values[2], fumblesLost: values[3], standardPoints: values[4],
      });
    }
  }
  return rows;
}
function sha256(file) { return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex'); }
function canonicalHeader(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}
function headerMap(row) {
  return Object.fromEntries((row || []).map((h, i) => [canonicalHeader(h), i]));
}
function getCell(row, headers, ...names) {
  for (const name of names) {
    const key = canonicalHeader(name);
    if (headers[key] !== undefined) return row[headers[key]];
  }
  return '';
}

await ensureData();
if (!fs.existsSync(workbookPath)) throw new Error(`Workbook not found: ${workbookPath}`);

const statsRows = readCsv(path.join(cacheDir, `stats_player_reg_${priorSeason}.csv`));
const priorStatsRows = readCsv(path.join(cacheDir, `stats_player_reg_${priorSeason - 1}.csv`));
const weeklyStatsRows = readCsv(path.join(cacheDir, `stats_player_week_${priorSeason}.csv`));
const snapRows = readCsv(path.join(cacheDir, `snap_counts_${priorSeason}.csv`));
const rosterRows = readCsv(path.join(cacheDir, `roster_${season}.csv`));
const draftRows = readCsv(path.join(cacheDir, 'draft_picks.csv'));
const ngsPass = newestNgs(readCsv(path.join(cacheDir, 'ngs_passing.csv.gz')), 'QB');
const ngsRecv = newestNgs(readCsv(path.join(cacheDir, 'ngs_receiving.csv.gz')), 'WR');
const ngsRush = newestNgs(readCsv(path.join(cacheDir, 'ngs_rushing.csv.gz')), 'RB');
const sleeperProjectionRows = JSON.parse(fs.readFileSync(path.join(cacheDir, `sleeper_projections_${season}.json`), 'utf8'));
const sleeperPlayerRows = Object.values(JSON.parse(fs.readFileSync(path.join(cacheDir, `sleeper_players_${season}.json`), 'utf8')));
const fantasyProsRows = ['QB', 'RB', 'WR', 'TE'].flatMap((position) => parseFantasyProsProjections(
  fs.readFileSync(path.join(cacheDir, `fantasypros_${position.toLowerCase()}_${season}.html`), 'utf8'), position,
));
const statsById = new Map(statsRows.filter((r) => r.player_id).map((r) => [r.player_id, r]));
const statsByIdentity = indexRowsByIdentity(statsRows, ['player_display_name', 'player_name'], (row) => row.position);
const priorStatsByIdentity = indexRowsByIdentity(priorStatsRows, ['player_display_name', 'player_name'], (row) => row.position);
const rosterByIdentity = indexRowsByIdentity(rosterRows, ['full_name', 'football_name'], (row) => row.position || row.depth_chart_position);
const draftByIdentity = indexRowsByIdentity(draftRows.slice().sort((a, b) => Number(b.season) - Number(a.season)), ['pfr_player_name'], (row) => row.category || row.position);
const sleeperPlayersByIdentity = indexRowsByIdentity(sleeperPlayerRows, ['full_name'], (row) => sleeperFantasyPosition(row));
const fantasyProsByIdentity = indexRowsByIdentity(fantasyProsRows, ['name'], (row) => row.position);
const sleeperProjectionsByIdentity = new Map();
for (const projection of sleeperProjectionRows) {
  if (projection.game_id !== 'season' || projection.category !== 'proj' || String(projection.season) !== String(season)) continue;
  const fullName = projection.player?.full_name || `${projection.player?.first_name || ''} ${projection.player?.last_name || ''}`.trim();
  const position = sleeperFantasyPosition(projection.player) || projection.position;
  const key = identityKey(fullName, position);
  const old = sleeperProjectionsByIdentity.get(key);
  if (normalizeName(fullName) && normalizePosition(position) && (!old || Number(projection.updated_at) > Number(old.updated_at))) sleeperProjectionsByIdentity.set(key, projection);
}

const VARIANT_ADP_FIELDS = {
  standard: 'adp_std',
  halfPpr: 'adp_half_ppr',
  ppr: 'adp_ppr',
  superflex: 'adp_2qb',
};
function liveSleeperAdp(projection, variant) {
  return num(projection?.stats?.[VARIANT_ADP_FIELDS[variant]]);
}
function availabilityAdjustment(player, projection) {
  const injuryStatus = String(player?.injury_status || '').trim();
  const injuryDetail = `${player?.injury_body_part || ''} ${player?.injury_notes || ''}`.trim();
  const newsUpdatedAt = num(player?.news_updated) || 0;
  const projectionUpdatedAt = num(projection?.updated_at) || 0;
  const injuryIsNewer = Boolean(injuryStatus && newsUpdatedAt > projectionUpdatedAt);
  const seriousSurgery = /\b(acl|achilles)\b/i.test(injuryDetail) && /surgery/i.test(injuryDetail);
  const hasSeasonProjection = ['pts_std', 'pts_half_ppr', 'pts_ppr'].some((field) => num(projection?.stats?.[field]) !== null);
  const seasonEnding = injuryStatus.toUpperCase() === 'IR' && seriousSurgery && (injuryIsNewer || !hasSeasonProjection);
  let multiplier = 1;
  if (seasonEnding) multiplier = 0;
  else if (injuryIsNewer && ['IR', 'PUP'].includes(injuryStatus.toUpperCase())) multiplier = 13 / 17;
  else if (injuryIsNewer && injuryStatus.toUpperCase() === 'OUT') multiplier = 16 / 17;
  else if (injuryIsNewer && injuryStatus.toUpperCase() === 'DOUBTFUL') multiplier = 0.9;
  else if (injuryIsNewer && injuryStatus.toUpperCase() === 'QUESTIONABLE') multiplier = 0.97;
  return {
    multiplier,
    seasonEnding,
    injuryIsNewer,
    injuryStatus: injuryStatus || null,
    injuryDetail: injuryDetail || null,
    newsUpdatedAt: newsUpdatedAt ? new Date(newsUpdatedAt).toISOString() : null,
    projectionUpdatedAt: projectionUpdatedAt ? new Date(projectionUpdatedAt).toISOString() : null,
  };
}
function currentDepthRole(position, player, fallback = '') {
  const normalizedPosition = normalizePosition(position);
  const order = num(player?.depth_chart_order);
  if (['RB', 'WR', 'TE'].includes(normalizedPosition) && order !== null && order >= 1) {
    return `${normalizedPosition}${Math.max(1, Math.round(order))}`;
  }
  return String(fallback || '').trim();
}
const snapsAgg = new Map();
for (const row of snapRows) {
  const key = identityKey(row.player, row.position);
  const current = snapsAgg.get(key) || { snaps: 0, games: 0 };
  current.snaps += num(row.offense_snaps) || 0;
  current.games += 1;
  snapsAgg.set(key, current);
}
const weeklyAgg = new Map();
for (const row of weeklyStatsRows.filter((item) => item.season_type === 'REG')) {
  const key = identityKey(row.player_display_name || row.player_name, row.position);
  const current = weeklyAgg.get(key) || { games: 0, pprTotal: 0, pprMax: 0, receivingYardsTotal: 0, receivingYardsMax: 0 };
  const ppr = num(row.fantasy_points_ppr) || num(row.fantasy_points) || 0;
  const receivingYards = num(row.receiving_yards) || 0;
  current.games += 1;
  current.pprTotal += ppr;
  current.pprMax = Math.max(current.pprMax, ppr);
  current.receivingYardsTotal += receivingYards;
  current.receivingYardsMax = Math.max(current.receivingYardsMax, receivingYards);
  weeklyAgg.set(key, current);
}

const wb = XLSX.readFile(workbookPath, { cellDates: true });
const sheetToVariant = {
  STRD: 'standard',
  '.5 PPR NEW': 'halfPpr',
  'FULL PPR NEW': 'ppr',
  'SUPERFLEX NEW': 'superflex',
};
const playerMap = new Map();
for (const sheetName of wb.SheetNames) {
  const variant = sheetToVariant[sheetName];
  if (!variant) continue;
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', raw: true, blankrows: false });
  const headerIndex = rows.findIndex((r) => r.some((v) => String(v).toLowerCase().trim() === 'player name'));
  const headers = headerMap(rows[headerIndex]);
  for (const row of rows.slice(headerIndex + 1)) {
    const name = String(getCell(row, headers, 'Player Name') || '').trim();
    const position = normalizePosition(getCell(row, headers, 'POS'));
    if (!name || !position) continue;
    const team = String(getCell(row, headers, 'TEAM') || '').toUpperCase().trim();
    const key = identityKey(name, position);
    const entry = playerMap.get(key) || { name, normalizedName: normalizeName(name), team, position, variants: {}, workbook: {} };
    entry.team = entry.team || team;
    entry.workbook = {
      age: num(getCell(row, headers, 'Age')),
      rookie: flag(getCell(row, headers, 'ROOKIE')),
      thirdYear: flag(getCell(row, headers, '3RD YEAR WR', '3rd Yr WR')),
      cuff: getCell(row, headers, 'Cuff'),
      passBlock: getCell(row, headers, 'Pass Block', 'Oline Pass Blk'),
      runBlock: getCell(row, headers, 'Run Block', 'Oline Run Blk'),
      oline: getCell(row, headers, 'OLine Ovrl'),
      depth: getCell(row, headers, 'Depth Chart', 'Team Depth Chart WR/TE'),
      defense: getCell(row, headers, 'Defense', 'Team Def'),
      earlySos: getCell(row, headers, 'First 5 weeks SoS (Top 5 Easiest/Hardest)', '1st 5 games SoS by position'),
      fullSos: getCell(row, headers, 'Full SoS', 'Full SoS by position 1-14'),
      boomFactor: getCell(row, headers, 'Boom Factor'),
      boomConnect: getCell(row, headers, 'Boom Connect', 'QB and/or WR/TE Boom Connect'),
    };
    const sleeper = num(getCell(row, headers, 'Sleeper ADP', 'ADP Sleeper'));
    const rts = num(getCell(row, headers, 'RT Sports ADP', 'ADP RT Sports'));
    const sheetRank = num(getCell(row, headers, 'RK', 'Super Flex Full PPR RK', 'STRD RK', '.5 PPR RK', 'FULL PPR RK'));
    const currentProjection = sleeperProjectionsByIdentity.get(key);
    const currentAdp = liveSleeperAdp(currentProjection, variant);
    const workbookMarketRank = variant === 'superflex' ? sheetRank : mean([sleeper, rts]) ?? sheetRank;
    const marketRank = currentAdp ?? workbookMarketRank;
    entry.variants[variant] = {
      sleeperAdp: sleeper,
      rtSportsAdp: rts,
      marketRank,
      sheetRank,
      workbookMarketRank,
      marketRankSource: currentAdp !== null ? 'Sleeper current ADP' : 'Workbook fallback',
      marketRankUpdatedAt: currentAdp !== null && currentProjection?.updated_at ? new Date(Number(currentProjection.updated_at)).toISOString() : null,
    };
    playerMap.set(key, entry);
  }
}

// The workbook is a view, not the player universe. Add any currently drafted
// player in the live Sleeper feed so a new depth-chart riser cannot be omitted.
for (const [key, projection] of sleeperProjectionsByIdentity) {
  const position = normalizePosition(sleeperFantasyPosition(projection.player) || projection.position);
  if (!['QB', 'RB', 'WR', 'TE'].includes(position)) continue;
  const hasDraftableAdp = Object.values(VARIANT_ADP_FIELDS).some((field) => {
    const value = num(projection.stats?.[field]);
    return value !== null && value <= 300;
  });
  if (!hasDraftableAdp) continue;
  const name = String(projection.player?.full_name || `${projection.player?.first_name || ''} ${projection.player?.last_name || ''}`).trim();
  if (!name) continue;
  const entry = playerMap.get(key) || {
    name,
    normalizedName: normalizeName(name),
    team: normalizeTeam(projection.team || projection.player?.team),
    position,
    variants: {},
    workbook: {},
  };
  for (const variant of Object.values(sheetToVariant)) {
    const currentAdp = liveSleeperAdp(projection, variant);
    if (currentAdp === null) continue;
    const previous = entry.variants[variant] || {};
    entry.variants[variant] = {
      ...previous,
      marketRank: currentAdp,
      marketRankSource: 'Sleeper current ADP',
      marketRankUpdatedAt: projection.updated_at ? new Date(Number(projection.updated_at)).toISOString() : null,
    };
  }
  playerMap.set(key, entry);
}

const teamContexts = new Map();
const meaningful = (value) => Boolean(String(value || '').trim() && String(value || '').trim() !== '-');
for (const entry of playerMap.values()) {
  if (!entry.team || entry.team === 'FA') continue;
  const key = `${entry.team}|${entry.position}`;
  const context = teamContexts.get(key) || {};
  for (const field of ['passBlock', 'runBlock', 'oline', 'defense', 'earlySos', 'fullSos']) {
    if (!meaningful(context[field]) && meaningful(entry.workbook[field])) context[field] = entry.workbook[field];
  }
  teamContexts.set(key, context);
}

const teamQbQuality = new Map();
for (const s of statsRows.filter((r) => r.position === 'QB')) {
  const attempts = num(s.attempts) || 0;
  if (attempts < 50) continue;
  const rate = (num(s.passing_epa) || 0) / attempts + (num(s.passing_cpoe) || 0) / 20;
  const team = normalizeTeam(s.recent_team);
  const old = teamQbQuality.get(team);
  if (!old || attempts > old.attempts) teamQbQuality.set(team, { value: rate, attempts });
}

// Current line context is rebuilt from completed-season team performance and
// today's player directory. Workbook line grades are intentionally display-only.
const teamLineRaw = new Map();
for (const row of statsRows) {
  const team = normalizeTeam(row.recent_team);
  if (!team) continue;
  const current = teamLineRaw.get(team) || { passAttempts: 0, sacks: 0, rbCarries: 0, rbRushingEpa: 0 };
  if (row.position === 'QB') {
    current.passAttempts += num(row.attempts) || 0;
    current.sacks += num(row.sacks_suffered) || 0;
  } else if (row.position === 'RB') {
    current.rbCarries += num(row.carries) || 0;
    current.rbRushingEpa += num(row.rushing_epa) || 0;
  }
  teamLineRaw.set(team, current);
}
function standardizeTeamMetric(rows, getValue) {
  const values = rows.map(([, value]) => getValue(value)).filter(Number.isFinite);
  const average = mean(values) || 0;
  const sd = Math.sqrt(mean(values.map((value) => (value - average) ** 2)) || 1);
  return new Map(rows.map(([team, value]) => [team, clamp((getValue(value) - average) / sd, -2.2, 2.2)]));
}
const lineRows = [...teamLineRaw.entries()].filter(([, value]) => value.passAttempts >= 100 && value.rbCarries >= 50);
const passLineBaselines = standardizeTeamMetric(lineRows, (value) => -(value.sacks / Math.max(1, value.passAttempts + value.sacks)));
const runLineBaselines = standardizeTeamMetric(lineRows, (value) => value.rbRushingEpa / Math.max(1, value.rbCarries));
const olInjuryBurden = new Map();
const OL_POSITIONS = new Set(['OL', 'OT', 'T', 'G', 'C']);
const OL_STATUS_WEIGHT = { IR: 1, PUP: 0.85, OUT: 0.65, DOUBTFUL: 0.45, QUESTIONABLE: 0.2 };
for (const player of sleeperPlayerRows) {
  if (!OL_POSITIONS.has(normalizePosition(player.position))) continue;
  const team = normalizeTeam(player.team);
  const statusWeight = OL_STATUS_WEIGHT[String(player.injury_status || '').toUpperCase()] || 0;
  if (!team || !statusWeight) continue;
  const depthOrder = num(player.depth_chart_order);
  const depthWeight = depthOrder === null ? 0.7 : depthOrder <= 1 ? 1 : depthOrder === 2 ? 0.55 : 0.25;
  olInjuryBurden.set(team, (olInjuryBurden.get(team) || 0) + statusWeight * depthWeight);
}
function currentLineSignal(team, type) {
  const baseline = (type === 'run' ? runLineBaselines : passLineBaselines).get(normalizeTeam(team)) || 0;
  const injuryPenalty = clamp((olInjuryBurden.get(normalizeTeam(team)) || 0) * 0.22, 0, 1.1);
  return clamp(baseline - injuryPenalty, -2.5, 2.5);
}

const DEPTH_MAP = [['STARTER', 1.1], ['RB1', 1.1], ['WR1', 1.1], ['TE1', 1], ['QB1', 1], ['BACKUP', -0.8], ['RB2', 0.15], ['WR2', 0.3], ['TE2', 0.1]];
const POSITION_WEIGHTS = {
  QB: { opportunity: 0.14, efficiency: 0.05, environment: 0.06, trajectory: 0.05, reversion: 0.05, floor: 0.08, context: 0.04, concentration: 0.00, availability: 0.03, projection: 0.50 },
  RB: { opportunity: 0.11, efficiency: 0.03, environment: 0.06, trajectory: 0.06, reversion: 0.08, floor: 0.00, context: 0.11, concentration: 0.00, availability: 0.05, projection: 0.50 },
  WR: { opportunity: 0.13, efficiency: 0.05, environment: 0.05, trajectory: 0.06, reversion: 0.06, floor: 0.00, context: 0.09, concentration: 0.03, availability: 0.03, projection: 0.50 },
  TE: { opportunity: 0.12, efficiency: 0.04, environment: 0.05, trajectory: 0.07, reversion: 0.08, floor: 0.00, context: 0.10, concentration: 0.01, availability: 0.03, projection: 0.50 },
};
function projectionMean(...values) {
  return mean(values.filter(Number.isFinite));
}
function projectedStat(sleeperStats, fantasyPros, sleeperField, fantasyProsField) {
  return projectionMean(num(sleeperStats?.[sleeperField]), num(fantasyPros?.[fantasyProsField]));
}
function getLeagueYears(sleeperPlayer, roster, draft) {
  const sleeperYears = num(sleeperPlayer?.years_exp);
  if (sleeperYears !== null) return Math.max(1, Math.round(sleeperYears) + 1);
  const rookieYear = num(sleeperPlayer?.rookie_year) ?? num(roster?.rookie_year) ?? num(draft?.season);
  return rookieYear === null ? null : Math.max(1, season - rookieYear + 1);
}
function entryCurrentTeam(entry) {
  const identity = identityKey(entry.name, entry.position);
  return String(
    sleeperPlayersByIdentity.get(identity)?.team
    || sleeperProjectionsByIdentity.get(identity)?.team
    || fantasyProsByIdentity.get(identity)?.team
    || rosterByIdentity.get(identity)?.team
    || entry.team
    || 'FA'
  ).toUpperCase();
}
function entryProjectedPpr(entry) {
  const identity = identityKey(entry.name, entry.position);
  const sleeperStats = sleeperProjectionsByIdentity.get(identity)?.stats || {};
  const fantasyPros = fantasyProsByIdentity.get(identity);
  const fantasyProsStandard = fantasyPros ? num(fantasyPros.standardPoints) : null;
  const fantasyProsPpr = fantasyProsStandard === null ? null : fantasyProsStandard + (num(fantasyPros.receptions) || 0);
  return projectionMean(num(sleeperStats.pts_ppr), fantasyProsPpr);
}

const priorTeamUsage = new Map();
for (const stat of statsRows) {
  const team = normalizeTeam(stat.recent_team);
  if (!team) continue;
  const usage = priorTeamUsage.get(team) || { targets: 0, teTargets: 0, topQb: null };
  usage.targets += num(stat.targets) || 0;
  if (stat.position === 'TE') usage.teTargets += num(stat.targets) || 0;
  if (stat.position === 'QB' && (!usage.topQb || (num(stat.attempts) || 0) > (num(usage.topQb.attempts) || 0))) usage.topQb = stat;
  priorTeamUsage.set(team, usage);
}
const currentEntriesByTeam = new Map();
for (const entry of playerMap.values()) {
  if (!['QB', 'RB', 'WR', 'TE'].includes(entry.position)) continue;
  const team = normalizeTeam(entryCurrentTeam(entry));
  if (!team || team === 'FA') continue;
  const entries = currentEntriesByTeam.get(team) || [];
  entries.push(entry);
  currentEntriesByTeam.set(team, entries);
}
const playerMarketContexts = new Map();
for (const entry of playerMap.values()) {
  if (!['QB', 'RB', 'WR', 'TE'].includes(entry.position)) continue;
  const identity = identityKey(entry.name, entry.position);
  const team = normalizeTeam(entryCurrentTeam(entry));
  const teamEntries = currentEntriesByTeam.get(team) || [];
  const teamUsage = priorTeamUsage.get(team) || { targets: 0, teTargets: 0, topQb: null };
  let retainedTargets = 0;
  let incomingReceivingYards = 0;
  for (const teammate of teamEntries) {
    const teammateStat = statsByIdentity.get(identityKey(teammate.name, teammate.position));
    if (!teammateStat) continue;
    if (normalizeTeam(teammateStat.recent_team) === team) retainedTargets += num(teammateStat.targets) || 0;
    else incomingReceivingYards += num(teammateStat.receiving_yards) || 0;
  }
  const targetVacancy = teamUsage.targets > 0 ? clamp((teamUsage.targets - retainedTargets) / teamUsage.targets, 0, 0.75) : 0;
  const priorQbPpg = teamUsage.topQb
    ? ((num(teamUsage.topQb.fantasy_points_ppr) || num(teamUsage.topQb.fantasy_points) || 0) / Math.max(1, num(teamUsage.topQb.games) || 1))
    : 15;
  const currentQb = teamEntries.filter((item) => item.position === 'QB').sort((a, b) => {
    const aProjection = entryProjectedPpr(a) || 0;
    const bProjection = entryProjectedPpr(b) || 0;
    if (aProjection !== bProjection) return bProjection - aProjection;
    const aRank = Math.min(...Object.values(a.variants).map((variant) => variant.marketRank).filter(Number.isFinite), 999);
    const bRank = Math.min(...Object.values(b.variants).map((variant) => variant.marketRank).filter(Number.isFinite), 999);
    return aRank - bRank;
  })[0];
  const currentQbPpg = currentQb ? (entryProjectedPpr(currentQb) || priorQbPpg * 17) / 17 : priorQbPpg;
  const qbChange = clamp((currentQbPpg - priorQbPpg) / 5, -1.8, 1.8);
  const competitors = teamEntries.filter((item) => item.position === entry.position && identityKey(item.name, item.position) !== identity)
    .map((item) => statsByIdentity.get(identityKey(item.name, item.position))).filter(Boolean);
  let competition = 0;
  if (entry.position === 'RB') competition = Math.max(0, ...competitors.map((stat) => ((num(stat.carries) || 0) + (num(stat.targets) || 0) * 1.5) / Math.max(1, num(stat.games) || 1)));
  else if (entry.position === 'TE') competition = competitors.reduce((sum, stat) => sum + (num(stat.targets) || 0) / Math.max(1, num(stat.games) || 1), 0);
  else if (entry.position === 'WR') competition = Math.max(0, ...competitors.map((stat) => (num(stat.targets) || 0) / Math.max(1, num(stat.games) || 1)));
  const teamTeTargetShare = teamUsage.targets > 0 ? teamUsage.teTargets / teamUsage.targets : 0.16;
  let value = 0;
  if (entry.position === 'QB') value = clamp(incomingReceivingYards / 1200, 0, 1.5);
  else if (entry.position === 'RB') value = -clamp((competition - 8) / 8, 0, 1.5);
  else if (entry.position === 'WR') value = clamp(targetVacancy * 2 + qbChange * 0.8 - Math.max(0, competition - 6) / 10, -2, 2);
  else if (entry.position === 'TE') value = clamp(targetVacancy * 0.7 + qbChange * 0.7 + (teamTeTargetShare - 0.16) * 4 - competition / 12, -2, 2);
  playerMarketContexts.set(identity, {
    value, targetVacancy, incomingReceivingYards, qbChange, competition, teamTeTargetShare,
    currentQb: currentQb?.name || null,
  });
}
const records = [];
const points = [];
const unmatched = [];

for (const entry of playerMap.values()) {
  const identity = identityKey(entry.name, entry.position);
  const marketContext = playerMarketContexts.get(identity) || {};
  const weekly = weeklyAgg.get(identity);
  const roster = rosterByIdentity.get(identity);
  const sleeperPlayer = sleeperPlayersByIdentity.get(identity);
  const projection = sleeperProjectionsByIdentity.get(identity);
  const fantasyPros = fantasyProsByIdentity.get(identity);
  const projectionStats = projection?.stats || {};
  const availability = availabilityAdjustment(sleeperPlayer, projection);
  const currentTeam = String(sleeperPlayer?.team || projection?.team || fantasyPros?.team || roster?.team || entry.team || 'FA').toUpperCase();
  const gsisId = roster?.gsis_id || '';
  const stat = (gsisId ? statsById.get(gsisId) : null) || statsByIdentity.get(identity);
  const priorStat = priorStatsByIdentity.get(identity);
  const draft = draftByIdentity.get(identity);
  const leagueYears = getLeagueYears(sleeperPlayer, roster, draft);
  const receptions = num(stat?.receptions) || 0;
  const priorStandardPoints = num(stat?.fantasy_points);
  const priorPprPoints = num(stat?.fantasy_points_ppr) ?? (priorStandardPoints === null ? null : priorStandardPoints + receptions);
  const fpStandard = num(fantasyPros?.standardPoints);
  const fpReceptions = num(fantasyPros?.receptions) || 0;
  const sleeperProjectedPoints = {
    standard: num(projectionStats.pts_std),
    halfPpr: num(projectionStats.pts_half_ppr),
    ppr: num(projectionStats.pts_ppr),
    superflex: num(projectionStats.pts_ppr),
  };
  const fantasyProsProjectedPoints = {
    standard: fpStandard,
    halfPpr: fpStandard === null ? null : fpStandard + fpReceptions * 0.5,
    ppr: fpStandard === null ? null : fpStandard + fpReceptions,
    superflex: fpStandard === null ? null : fpStandard + fpReceptions,
  };
  const projectedPoints = Object.fromEntries(Object.entries({
    standard: projectionMean(sleeperProjectedPoints.standard, fantasyProsProjectedPoints.standard),
    halfPpr: projectionMean(sleeperProjectedPoints.halfPpr, fantasyProsProjectedPoints.halfPpr),
    ppr: projectionMean(sleeperProjectedPoints.ppr, fantasyProsProjectedPoints.ppr),
    superflex: projectionMean(sleeperProjectedPoints.superflex, fantasyProsProjectedPoints.superflex),
  }).map(([variant, value]) => [variant, availability.seasonEnding ? 0 : value === null ? null : value * availability.multiplier]));
  points.push({
    name: entry.name,
    normalizedName: entry.normalizedName,
    position: entry.position,
    leagueYears,
    priorSeasonPoints: {
      standard: priorStandardPoints === null ? null : round(priorStandardPoints),
      halfPpr: priorStandardPoints === null ? null : round(priorStandardPoints + receptions * 0.5),
      ppr: priorPprPoints === null ? null : round(priorPprPoints),
      superflex: priorPprPoints === null ? null : round(priorPprPoints),
    },
    projectionPoints: Object.fromEntries(Object.entries(projectedPoints).map(([variant, value]) => [variant, value === null ? null : round(value)])),
    projectionSources: [Number.isFinite(sleeperProjectedPoints.ppr) ? 'Sleeper / RotoWire' : null, Number.isFinite(fantasyProsProjectedPoints.ppr) ? 'FantasyPros consensus' : null].filter(Boolean),
    availability,
  });
  if (!['QB', 'RB', 'WR', 'TE'].includes(entry.position)) continue;
  const snap = snapsAgg.get(identity);
  const ngs = entry.position === 'QB' ? ngsPass.get(identity) : entry.position === 'RB' ? ngsRush.get(identity) : ngsRecv.get(identity);
  const teamContext = teamContexts.get(`${currentTeam}|${entry.position}`) || {};
  const w = { ...teamContext, ...Object.fromEntries(Object.entries(entry.workbook).filter(([, value]) => meaningful(value))) };
  for (const field of ['passBlock', 'runBlock', 'oline', 'defense', 'earlySos', 'fullSos']) {
    if (!meaningful(entry.workbook[field]) && meaningful(teamContext[field])) w[field] = teamContext[field];
  }
  const games = num(stat?.games) || 0;
  const displayDepthRole = currentDepthRole(entry.position, sleeperPlayer, w.depth);
  const isRookie = leagueYears === 1 || Boolean(w.rookie || Number(roster?.rookie_year) === season || Number(draft?.season) === season);
  const targets = num(stat?.targets) || 0;
  const carries = num(stat?.carries) || 0;
  const attempts = num(stat?.attempts) || 0;
  const fantasy = num(stat?.fantasy_points_ppr) || num(stat?.fantasy_points) || 0;
  const opportunityDenom = entry.position === 'QB' ? Math.max(1, attempts + carries) : Math.max(1, targets + carries);
  const targetShare = num(stat?.target_share);
  const airShare = num(stat?.air_yards_share);
  const wopr = num(stat?.wopr);
  const snapPerGame = snap ? snap.snaps / Math.max(1, snap.games) : null;
  const depth = textTier(w.depth, DEPTH_MAP, 0);
  const projectedGamesRaw = num(projectionStats.gp) || (fantasyPros ? 18 : null);
  const projectedGames = projectedGamesRaw === null ? null : projectedGamesRaw * availability.multiplier;
  const adjustedProjectedStat = (sleeperField, fantasyProsField) => {
    const value = projectedStat(projectionStats, fantasyPros, sleeperField, fantasyProsField);
    return value === null ? null : value * availability.multiplier;
  };
  const projectedPassAttempts = adjustedProjectedStat('pass_att', 'passAtt');
  const projectedRushAttempts = adjustedProjectedStat('rush_att', 'rushAtt');
  const projectedReceptions = adjustedProjectedStat('rec', 'receptions');
  const projectedPassYards = adjustedProjectedStat('pass_yd', 'passYards');
  const projectedRushYards = adjustedProjectedStat('rush_yd', 'rushYards');
  const projectedReceivingYards = adjustedProjectedStat('rec_yd', 'receivingYards');
  const projectedPassTds = adjustedProjectedStat('pass_td', 'passTds');
  const projectedRushTds = adjustedProjectedStat('rush_td', 'rushTds');
  const projectedReceivingTds = adjustedProjectedStat('rec_td', 'receivingTds');
  const projectedOpportunity = entry.position === 'QB'
    ? (projectedPassAttempts || 0) + (projectedRushAttempts || 0) * 1.75
    : entry.position === 'RB'
      ? (projectedRushAttempts || 0) + (projectedReceptions || 0) * 1.75
      : (projectedReceptions || 0) * 1.5;
  const historicalOpportunity = entry.position === 'QB' ? attempts + carries * 1.75 : entry.position === 'RB' ? carries + targets * 1.75 : targets * 1.5;
  const historicalOpportunityPerGame = stat ? historicalOpportunity / Math.max(1, games) : null;
  const comparableHistoricalOpportunity = entry.position === 'QB' ? attempts + carries * 1.75 : entry.position === 'RB' ? carries + receptions * 1.75 : receptions * 1.5;
  const comparableHistoricalOpportunityPerGame = stat ? comparableHistoricalOpportunity / Math.max(1, games) : null;
  const projectedOpportunityPerGame = projectedGames && projectedOpportunity ? projectedOpportunity / projectedGames : null;
  const opportunityGrowth = comparableHistoricalOpportunityPerGame && projectedOpportunityPerGame
    ? clamp(Math.log(projectedOpportunityPerGame / comparableHistoricalOpportunityPerGame), -1.25, 1.25)
    : projectedOpportunityPerGame ? clamp((projectedOpportunityPerGame - (entry.position === 'QB' ? 28 : entry.position === 'RB' ? 11 : entry.position === 'TE' ? 3.5 : 5)) / 10, -1.25, 1.25) : 0;
  let volume = 0;
  if (stat) {
    if (entry.position === 'QB') volume = (attempts / Math.max(1, games) - 30) / 7 + (carries / Math.max(1, games) - 3) / 4;
    else if (entry.position === 'RB') volume = (historicalOpportunityPerGame - 15) / 7 + ((targetShare ?? 0.08) - 0.10) * 5;
    else if (entry.position === 'WR') volume = (targets / Math.max(1, games) - 6.5) / 3 + ((targetShare ?? 0.16) - 0.18) * 5 + ((airShare ?? 0.20) - 0.22) * 3;
    else volume = (targets / Math.max(1, games) - 4.2) / 2.4 + ((targetShare ?? 0.12) - 0.14) * 5 + ((airShare ?? 0.12) - 0.14) * 2;
  }
  const snapSignal = snapPerGame === null ? 0 : (snapPerGame - (entry.position === 'RB' ? 28 : 38)) / 24;
  const opportunity = clamp(depth * 0.35 + volume * 0.55 + snapSignal * 0.15 + opportunityGrowth * 0.75, -3, 3);
  let efficiency = 0;
  if (stat) {
    const reliability = Math.sqrt(opportunityDenom / (opportunityDenom + 100));
    if (entry.position === 'QB') efficiency = (((num(stat.passing_epa) || 0) / Math.max(1, attempts)) * 3 + (num(stat.passing_cpoe) || 0) / 10 + (num(stat.rushing_yards) || 0) / Math.max(1, games) / 40) * reliability;
    else efficiency = (((num(stat.receiving_epa) || 0) + (num(stat.rushing_epa) || 0)) / opportunityDenom + (fantasy / opportunityDenom - 1) * 0.45) * reliability;
  }
  if (ngs) {
    if (entry.position === 'QB') efficiency += ((num(ngs.passer_rating) || 85) - 90) / 30 + (num(ngs.completion_percentage_above_expectation) || 0) / 12;
    else if (entry.position === 'RB') efficiency += (num(ngs.rush_yards_over_expected_per_att) || 0) / 1.5;
    else efficiency += (num(ngs.avg_yac_above_expectation) || 0) / 2 + ((num(ngs.avg_separation) || 3) - 3) / 1.5;
  }
  efficiency = clamp(efficiency, -2.5, 2.5);
  const passLine = currentLineSignal(currentTeam, 'pass');
  const runLine = currentLineSignal(currentTeam, 'run');
  const qb = teamQbQuality.get(currentTeam)?.value || 0;
  let environment = entry.position === 'RB'
    ? runLine * 0.75
    : passLine * 0.6 + (entry.position === 'WR' || entry.position === 'TE' ? qb * 0.6 : 0);
  environment = clamp(environment, -2.5, 2.5);
  const age = w.age ?? num(roster?.birth_date ? season - Number(roster.birth_date.slice(0, 4)) : null);
  let trajectory = 0;
  if (isRookie) trajectory += 0.35;
  if (leagueYears && leagueYears >= 2 && leagueYears <= 3) trajectory += entry.position === 'WR' || entry.position === 'TE' ? 0.4 : 0.2;
  if (draft && Number(draft.season) >= season - 2) trajectory += clamp((130 - (num(draft.pick) || 180)) / 120, -0.4, 0.8);
  if (priorStat && stat) {
    const priorGames = Math.max(1, num(priorStat.games) || 1);
    const priorOpp = entry.position === 'QB'
      ? ((num(priorStat.attempts) || 0) + (num(priorStat.carries) || 0) * 1.75) / priorGames
      : entry.position === 'RB'
        ? ((num(priorStat.carries) || 0) + (num(priorStat.targets) || 0) * 1.75) / priorGames
        : (num(priorStat.targets) || 0) * 1.5 / priorGames;
    if (priorOpp > 0 && historicalOpportunityPerGame) trajectory += clamp(Math.log(historicalOpportunityPerGame / priorOpp), -0.7, 0.7) * 0.55;
  }
  if (age) {
    const peak = entry.position === 'QB' ? 29 : entry.position === 'TE' ? 27 : 25;
    trajectory += age <= peak ? 0.25 : -Math.max(0, age - peak - 2) * 0.16;
  }
  trajectory = clamp(trajectory, -2.2, 2.2);
  let reversion = 0;
  if (stat && priorStat) {
    const statGames = Math.max(1, games);
    const olderGames = Math.max(1, num(priorStat.games) || 1);
    const recentPpg = fantasy / statGames;
    const olderPpg = (num(priorStat.fantasy_points_ppr) || num(priorStat.fantasy_points) || 0) / olderGames;
    if (entry.position === 'QB') {
      const recentTdRate = (num(stat.passing_tds) || 0) / Math.max(1, attempts);
      const olderTdRate = (num(priorStat.passing_tds) || 0) / Math.max(1, num(priorStat.attempts) || 0);
      reversion = clamp((olderTdRate - recentTdRate) * 18 + (olderPpg - recentPpg) / 14, -1.8, 1.8);
    } else {
      const productionGap = clamp((olderPpg - recentPpg) / 6, -1.8, 1.8);
      const earlyCareerDampener = leagueYears && leagueYears <= 3 && productionGap < 0 ? 0.25 : 1;
      reversion = productionGap * earlyCareerDampener;
      if (entry.position === 'RB' && age && age >= 26) {
        const recentOpp = (carries + targets * 1.75) / statGames;
        const olderOpp = ((num(priorStat.carries) || 0) + (num(priorStat.targets) || 0) * 1.75) / olderGames;
        if (olderOpp > 0 && recentOpp > olderOpp) reversion -= clamp(Math.log(recentOpp / olderOpp) * 0.65, 0, 0.65);
      }
    }
  }
  let qbFloor = 0;
  const passAdot = entry.position === 'QB' && attempts ? (num(stat?.passing_air_yards) || 0) / attempts : null;
  const qbDropbacks = entry.position === 'QB' ? attempts + (num(stat?.sacks_suffered) || 0) : 0;
  const sackRate = entry.position === 'QB' && qbDropbacks ? (num(stat?.sacks_suffered) || 0) / qbDropbacks : null;
  if (entry.position === 'QB' && stat) {
    qbFloor = clamp(((num(stat.rushing_yards) || 0) / Math.max(1, games) - 10) / 22
      + (num(stat.rushing_tds) || 0) / Math.max(1, games) * 2.2
      + ((passAdot ?? 7.5) - 7.5) / 3 - (sackRate || 0) * 2.5, -2.2, 2.2);
  }
  const contextSignal = marketContext.value || 0;
  const concentrationSignal = (entry.position === 'WR' || entry.position === 'TE') && weekly?.receivingYardsTotal
    ? -clamp((weekly.receivingYardsMax / weekly.receivingYardsTotal - 0.12) / 0.12, 0, 1.8)
    : 0;
  const availabilitySignal = availability.multiplier === 1 ? 0 : clamp((availability.multiplier - 1) * 4, -2.5, 0);
  const available = [Boolean(stat), Boolean(roster || sleeperPlayer), Boolean(snap), Boolean(ngs), Boolean(draft), Boolean(projection || fantasyPros), Boolean(w.depth), Boolean(w.oline), Boolean(w.earlySos)];
  let completeness = available.filter(Boolean).length / available.length;
  if (isRookie) completeness = Math.max(completeness, (available.filter(Boolean).length + 1.5) / (available.length + 1.5));
  completeness = clamp(completeness, 0, 1);
  const raw = 0;
  const canonicalId = gsisId || (sleeperPlayer?.player_id ? `sleeper-${sleeperPlayer.player_id}` : `ff-${entry.normalizedName}-${entry.position.toLowerCase()}`);
  const record = {
    id: canonicalId,
    normalizedName: entry.normalizedName,
    name: entry.name,
    team: currentTeam,
    workbookTeam: entry.team,
    position: entry.position,
    leagueYears,
    rookie: isRookie,
    rosterStatus: String(sleeperPlayer?.status || roster?.status || (currentTeam === 'FA' ? 'Free Agent' : 'Unknown')),
    availability,
    currentDirectoryMatched: Boolean(sleeperPlayer),
    confidence: Math.round(completeness * 100),
    joinMethod: gsisId && stat?.player_id === gsisId ? 'nfl-id' : sleeperPlayer && stat ? 'current-directory+normalized-name' : stat ? 'normalized-name' : sleeperPlayer ? 'current-directory-only' : roster ? 'roster-only' : 'workbook-only',
    components: { opportunity: round(opportunity, 2), efficiency: round(efficiency, 2), environment: round(environment, 2), trajectory: round(trajectory, 2), reversion: round(reversion, 2), floor: round(qbFloor, 2), context: round(contextSignal, 2), concentration: round(concentrationSignal, 2), availability: round(availabilitySignal, 2) },
    componentRaw: { opportunity, efficiency, environment, trajectory, reversion, floor: qbFloor, context: contextSignal, concentration: concentrationSignal, availability: availabilitySignal },
    raw,
    projectionPoints: projectedPoints,
    projectionSourceAvailability: { sleeper: Number.isFinite(sleeperProjectedPoints.ppr), fantasyPros: Boolean(fantasyPros) },
    projectionRaw: projectedGames ? projectedPoints.ppr / projectedGames : null,
    contextContributions: {
      line: round(entry.position === 'RB' ? runLine * 0.75 : passLine * 0.6, 2),
      schedule: 0,
      gameScript: 0,
      qbContext: round((entry.position === 'WR' || entry.position === 'TE') ? qb * 0.6 : 0, 2),
    },
    metrics: {
      currentTeam,
      rosterStatus: String(sleeperPlayer?.status || roster?.status || (currentTeam === 'FA' ? 'Free Agent' : 'Unknown')),
      games: games || null,
      snapsPerGame: snapPerGame === null ? null : round(snapPerGame, 1),
      targetShare: targetShare === null ? null : round(targetShare * 100, 1),
      airYardsShare: airShare === null ? null : round(airShare * 100, 1),
      wopr: wopr === null ? null : round(wopr, 2),
      targets: stat ? targets : null,
      touches: stat ? carries + receptions : null,
      receptions: stat ? receptions : null,
      yardsPerTarget: stat && targets ? round((num(stat.receiving_yards) || 0) / targets, 2) : null,
      yardsPerCarry: stat && carries ? round((num(stat.rushing_yards) || 0) / carries, 2) : null,
      receivingYardsPerOffenseSnap: stat && snap?.snaps ? round((num(stat.receiving_yards) || 0) / snap.snaps, 2) : null,
      passerRating: entry.position === 'QB' && num(ngs?.passer_rating) !== null ? round(num(ngs.passer_rating), 1) : null,
      cpoe: entry.position === 'QB' && num(ngs?.completion_percentage_above_expectation) !== null ? round(num(ngs.completion_percentage_above_expectation), 1) : null,
      rushYardsOverExpectedPerAttempt: entry.position === 'RB' && num(ngs?.rush_yards_over_expected_per_att) !== null ? round(num(ngs.rush_yards_over_expected_per_att), 2) : null,
      averageSeparation: (entry.position === 'WR' || entry.position === 'TE') && num(ngs?.avg_separation) !== null ? round(num(ngs.avg_separation), 2) : null,
      yacOverExpected: (entry.position === 'WR' || entry.position === 'TE') && num(ngs?.avg_yac_above_expectation) !== null ? round(num(ngs.avg_yac_above_expectation), 2) : null,
      fantasyPointsPerOpportunity: stat ? round(fantasy / opportunityDenom, 2) : null,
      passingTdRate: entry.position === 'QB' && attempts ? round((num(stat?.passing_tds) || 0) / attempts * 100, 1) : null,
      passingAdot: passAdot === null ? null : round(passAdot, 1),
      sackRate: sackRate === null ? null : round(sackRate * 100, 1),
      rushYardsPerGame: entry.position === 'QB' && stat ? round((num(stat.rushing_yards) || 0) / Math.max(1, games), 1) : null,
      rushTds: entry.position === 'QB' && stat ? num(stat.rushing_tds) || 0 : null,
      qbFloorSignal: entry.position === 'QB' && stat ? round(qbFloor, 2) : null,
      maxGamePprShare: weekly?.pprTotal ? round(weekly.pprMax / weekly.pprTotal * 100, 1) : null,
      maxGameReceivingYardShare: weekly?.receivingYardsTotal ? round(weekly.receivingYardsMax / weekly.receivingYardsTotal * 100, 1) : null,
      twoYearReversionSignal: stat && priorStat ? round(reversion, 2) : null,
      targetVacancyPct: Number.isFinite(marketContext.targetVacancy) ? round(marketContext.targetVacancy * 100, 1) : null,
      incomingReceivingYards: Number.isFinite(marketContext.incomingReceivingYards) ? round(marketContext.incomingReceivingYards) : null,
      teammateCompetition: Number.isFinite(marketContext.competition) ? round(marketContext.competition, 2) : null,
      projectedQbChangeSignal: Number.isFinite(marketContext.qbChange) ? round(marketContext.qbChange, 2) : null,
      projectedStartingQb: marketContext.currentQb || null,
      destinationTeTargetShare: entry.position === 'TE' && Number.isFinite(marketContext.teamTeTargetShare) ? round(marketContext.teamTeTargetShare * 100, 1) : null,
      qbContext: (entry.position === 'WR' || entry.position === 'TE') && teamQbQuality.has(currentTeam) ? round(qb, 2) : null,
      line: round(entry.position === 'RB' ? runLine : passLine, 2),
      lineMethod: `${priorSeason} team ${entry.position === 'RB' ? 'RB rushing EPA/carry' : 'QB sack rate'} adjusted for current OL injuries`,
      olInjuryBurden: round(olInjuryBurden.get(normalizeTeam(currentTeam)) || 0, 2),
      earlySchedule: String(w.earlySos || ''),
      fullSchedule: String(w.fullSos || ''),
      workbookEnvironmentDisplayOnly: true,
      depthRole: displayDepthRole,
      draftPick: num(draft?.pick),
      leagueYears,
      projectedGames,
      projectedOpportunity: projectedOpportunity ? round(projectedOpportunity, 1) : null,
      projectedOpportunityGrowthPct: comparableHistoricalOpportunityPerGame && projectedOpportunityPerGame ? round((projectedOpportunityPerGame / comparableHistoricalOpportunityPerGame - 1) * 100, 1) : null,
      projectedPassAttempts: projectedPassAttempts === null ? null : round(projectedPassAttempts, 1),
      projectedRushAttempts: projectedRushAttempts === null ? null : round(projectedRushAttempts, 1),
      projectedReceptions: projectedReceptions === null ? null : round(projectedReceptions, 1),
      projectedTouches: entry.position === 'RB' ? round((projectedRushAttempts || 0) + (projectedReceptions || 0), 1) : null,
      projectedPassYards: projectedPassYards === null ? null : round(projectedPassYards, 1),
      projectedRushYards: projectedRushYards === null ? null : round(projectedRushYards, 1),
      projectedReceivingYards: projectedReceivingYards === null ? null : round(projectedReceivingYards, 1),
      projectedTotalTds: (projectedPassTds || projectedRushTds || projectedReceivingTds) ? round((projectedPassTds || 0) + (projectedRushTds || 0) + (projectedReceivingTds || 0), 1) : null,
      projectedPprPoints: projectedPoints.ppr === null ? null : round(projectedPoints.ppr),
      projectionProvider: [projection?.company ? `Sleeper/${projection.company}` : null, fantasyPros ? 'FantasyPros consensus' : null].filter(Boolean).join(' + ') || null,
      projectionUpdatedAt: projection?.updated_at ? new Date(Number(projection.updated_at)).toISOString() : null,
      injuryStatus: availability.injuryStatus,
      injuryDetail: availability.injuryDetail,
      injuryNewsUpdatedAt: availability.newsUpdatedAt,
      projectionAvailabilityMultiplier: round(availability.multiplier, 3),
      seasonEndingInjury: availability.seasonEnding,
    },
    variants: entry.variants,
  };
  records.push(record);
  if (!stat && !isRookie) unmatched.push({ name: entry.name, team: currentTeam, position: entry.position, rosterMatched: Boolean(roster || sleeperPlayer), reason: 'No prior-season player-stat match' });
}

const positionGroups = Object.groupBy ? Object.groupBy(records, (r) => r.position) : records.reduce((a, r) => ((a[r.position] ||= []).push(r), a), {});
for (const [position, group] of Object.entries(positionGroups)) {
  for (const component of ['opportunity', 'efficiency', 'environment', 'trajectory', 'reversion', 'floor', 'context', 'concentration', 'availability']) {
    const values = group.map((r) => r.componentRaw[component]).filter(Number.isFinite);
    const low = percentile(values, 0.05);
    const high = percentile(values, 0.95);
    const clipped = values.map((value) => clamp(value, low, high));
    const average = mean(clipped) || 0;
    const sd = Math.sqrt(mean(clipped.map((value) => (value - average) ** 2)) || 1);
    group.forEach((r) => { r.components[component] = round(clamp((clamp(r.componentRaw[component], low, high) - average) / sd, -2.5, 2.5), 2); });
  }
  const weights = POSITION_WEIGHTS[position];
  group.forEach((r) => {
    r.raw = Object.entries(weights)
      .filter(([component]) => component !== 'projection')
      .reduce((sum, [component, weight]) => sum + (r.components[component] || 0) * weight, 0);
  });
}
for (const variant of Object.values(sheetToVariant)) {
  for (const [position, group] of Object.entries(positionGroups)) {
    const eligible = group.filter((r) => !r.availability.seasonEnding && Number.isFinite(r.variants[variant]?.marketRank) && r.variants[variant].marketRank <= 240);
    eligible.sort((a, b) => a.variants[variant].marketRank - b.variants[variant].marketRank);
    eligible.forEach((r, i) => { r.variants[variant].marketPositionRank = i + 1; });
    eligible.filter((r) => Number.isFinite(r.projectionPoints[variant]))
      .sort((a, b) => b.projectionPoints[variant] - a.projectionPoints[variant])
      .forEach((r, i) => { r.variants[variant].projectionPositionRank = i + 1; });
    const projectionDeltas = eligible
      .filter((r) => Number.isFinite(r.variants[variant].projectionPositionRank))
      .map((r) => r.variants[variant].marketPositionRank - r.variants[variant].projectionPositionRank);
    const projectionLow = percentile(projectionDeltas, 0.05);
    const projectionHigh = percentile(projectionDeltas, 0.95);
    const projectionClipped = projectionDeltas.map((value) => clamp(value, projectionLow, projectionHigh));
    const projectionAverage = mean(projectionClipped) || 0;
    const projectionSd = Math.sqrt(mean(projectionClipped.map((value) => (value - projectionAverage) ** 2)) || 1);
    const weights = POSITION_WEIGHTS[position];
    eligible.forEach((r) => {
      const v = r.variants[variant];
      const projectionDelta = Number.isFinite(v.projectionPositionRank) ? v.marketPositionRank - v.projectionPositionRank : null;
      v.projectionScore = projectionDelta === null ? 0 : clamp((clamp(projectionDelta, projectionLow, projectionHigh) - projectionAverage) / projectionSd, -2.5, 2.5);
      v.projectionWeight = projectionDelta === null ? 0 : weights.projection;
      const redistributedCore = projectionDelta === null ? 1 / (1 - weights.projection) : 1;
      v.combinedRaw = r.raw * redistributedCore + v.projectionScore * v.projectionWeight;
    });
    eligible.sort((a, b) => {
      const aProjection = -a.variants[variant].marketPositionRank + a.variants[variant].combinedRaw * 3.6;
      const bProjection = -b.variants[variant].marketPositionRank + b.variants[variant].combinedRaw * 3.6;
      return bProjection - aProjection || a.variants[variant].marketRank - b.variants[variant].marketRank;
    });
    eligible.forEach((r, i) => { r.variants[variant].modelRank = i + 1; });
    for (const r of eligible) {
      const v = r.variants[variant];
      const gain = v.marketPositionRank - v.modelRank;
      const probability = Math.round(100 / (1 + Math.exp(-(gain / 3.5 + v.combinedRaw * 0.3))));
      const score = clamp(probability, 0, 100);
      const tier = score >= 70 && gain >= 3 && r.confidence >= 65 ? 'strong' : score >= 58 && gain >= 1 && r.confidence >= 50 ? 'watch' : 'none';
      const variantComponents = { ...r.components, projection: round(v.projectionScore, 2), price: round(gain / 4, 2) };
      const reasonComponents = {
        price: variantComponents.price,
        projection: Number.isFinite(r.projectionPoints[variant]) ? variantComponents.projection : 0,
        opportunity: r.components.opportunity,
        efficiency: r.components.efficiency,
        trajectory: r.components.trajectory,
        line: r.components.environment,
        reversion: r.components.reversion,
        floor: r.components.floor,
        context: r.components.context,
        concentration: r.components.concentration,
        availability: r.components.availability,
      };
      const components = Object.entries(reasonComponents).sort((a, b) => Math.abs(b[1]) - Math.abs(a[1]));
      const selectedComponents = [];
      const addComponent = (name) => {
        const found = components.find(([componentName, componentValue]) => componentName === name && Math.abs(componentValue) >= 0.18);
        if (found && !selectedComponents.some(([componentName]) => componentName === name)) selectedComponents.push(found);
      };
      if (Number.isFinite(v.projectionPositionRank)) addComponent('projection');
      const strongestPositive = components.find(([, componentValue]) => componentValue >= 0.18);
      const strongestNegative = components.find(([, componentValue]) => componentValue <= -0.18);
      if (strongestPositive) addComponent(strongestPositive[0]);
      if (strongestNegative) addComponent(strongestNegative[0]);
      for (const [componentName] of components) {
        if (selectedComponents.length >= 4) break;
        addComponent(componentName);
      }
      const reasons = [];
      for (const [name, value] of selectedComponents) {
        if (reasons.length >= 4) break;
        const direction = value > 0 ? 'helps' : 'limits';
        const roleDetail = r.metrics.depthRole && r.metrics.depthRole !== '-' ? r.metrics.depthRole : r.metrics.games ? `${r.metrics.snapsPerGame ?? 'limited'} snaps/game` : 'role not established; no prior NFL sample';
        const lineDetail = r.metrics.line && r.metrics.line !== '-' ? r.metrics.line : 'unknown line';
        const projectedPoints = r.projectionPoints[variant];
        const projectedDetail = `${round(projectedPoints, 1)} consensus ${variant === 'halfPpr' ? 'half-PPR' : variant === 'standard' ? 'standard' : 'PPR'} points; projection rank ${v.projectionPositionRank} vs market pos. rank ${v.marketPositionRank}`;
        const label = name[0].toUpperCase() + name.slice(1);
        const detail = name === 'opportunity' ? roleDetail
          : name === 'efficiency' ? `${r.metrics.fantasyPointsPerOpportunity ?? 'no prior NFL sample'} fantasy points/opportunity`
            : name === 'projection' ? projectedDetail
              : name === 'line' ? `${lineDetail} current line signal (${r.metrics.lineMethod})`
                : name === 'reversion' ? 'two-year per-game production and workload direction'
                  : name === 'floor' ? `${r.metrics.rushYardsPerGame ?? 0} QB rush yards/game, ${r.metrics.passingAdot ?? 'unknown'} aDOT, ${r.metrics.sackRate ?? 'unknown'}% sack rate`
                    : name === 'context' ? `${r.metrics.targetVacancyPct ?? 0}% target vacancy, ${r.metrics.teammateCompetition ?? 0} teammate-competition signal, QB-change ${r.metrics.projectedQbChangeSignal ?? 0}`
                      : name === 'concentration' ? `${r.metrics.maxGameReceivingYardShare ?? 0}% of receiving yards in the largest game`
                        : name === 'availability' ? `${r.metrics.injuryStatus || 'active'}; ${round(r.metrics.projectionAvailabilityMultiplier * 100)}% projection availability`
                          : name === 'price' ? `market pos. rank ${v.marketPositionRank}, model rank ${v.modelRank}`
                            : (r.metrics.draftPick ? `pick ${r.metrics.draftPick}` : 'age/experience curve');
        reasons.push(`${label} ${direction}: ${detail}.`);
      }
      if (reasons.length < 2) reasons.push('Market disagreement is modest; use as a tie-breaker, not a stand-alone pick.');
      v.score = score;
      v.tier = tier;
      v.expectedDelta = gain;
      v.reasons = reasons.slice(0, 4);
      v.components = variantComponents;
      v.displayedMetrics = {
        ...r.metrics,
        consensusProjectionRank: v.projectionPositionRank || null,
        projectionVsMarketSlots: Number.isFinite(v.projectionPositionRank) ? v.marketPositionRank - v.projectionPositionRank : null,
      };
      v.sourceRefs = ['workbook', 'nflverseStats', 'nflverseWeekly', 'nflverseNgs', 'nflverseRoster', 'nflverseSnaps', 'nflverseDraft'];
      if (r.projectionSourceAvailability.sleeper) v.sourceRefs.push('sleeperProjections');
      if (r.projectionSourceAvailability.fantasyPros) v.sourceRefs.push('fantasyProsProjections');
      if (r.currentDirectoryMatched) v.sourceRefs.push('sleeperPlayers');
      delete v.sleeperAdp;
      delete v.rtSportsAdp;
      delete v.sheetRank;
      delete v.workbookMarketRank;
      delete v.combinedRaw;
    }
  }
}

// Preserve the stale provider number for auditability, but remove confirmed
// season-ending injuries from every active rank and give them no Edge score.
for (const r of records.filter((record) => record.availability.seasonEnding)) {
  for (const [variant, value] of Object.entries(r.variants)) {
    value.reportedMarketRankBeforeInjury = value.marketRank ?? null;
    value.marketRank = null;
    value.marketRankSource = `${value.marketRankSource || 'Provider'} (suppressed by confirmed season-ending injury)`;
    value.marketPositionRank = null;
    value.projectionPositionRank = null;
    value.modelRank = null;
    value.projectionScore = null;
    value.projectionWeight = 0;
    value.score = 0;
    value.tier = 'none';
    value.expectedDelta = null;
    value.reasons = [`Season-ending ${r.availability.injuryDetail || 'injury'} removes this player from the active market and model ranks.`];
    value.components = { ...r.components, projection: -2.5, price: 0 };
    value.displayedMetrics = { ...r.metrics, consensusProjectionRank: null, projectionVsMarketSlots: null };
    value.sourceRefs = ['sleeperPlayers', 'sleeperProjections'];
    delete value.sleeperAdp;
    delete value.rtSportsAdp;
    delete value.sheetRank;
    delete value.workbookMarketRank;
  }
}

for (const r of records) {
  delete r.raw;
  delete r.componentRaw;
  delete r.projectionSourceAvailability;
}
const generatedAt = new Date().toISOString();
const top240 = records.filter((r) => Object.values(r.variants).some((v) => Number.isFinite(v.marketRank) && v.marketRank <= 240));
const matched = top240.filter((r) => r.joinMethod !== 'workbook-only').length;
const freshnessFiles = {
  currentMarketAndProjections: `sleeper_projections_${season}.json`,
  currentPlayersAndInjuries: `sleeper_players_${season}.json`,
  fantasyProsQbProjections: `fantasypros_qb_${season}.html`,
  fantasyProsRbProjections: `fantasypros_rb_${season}.html`,
  fantasyProsWrProjections: `fantasypros_wr_${season}.html`,
  fantasyProsTeProjections: `fantasypros_te_${season}.html`,
  completedSeasonStats: `stats_player_reg_${priorSeason}.csv`,
  completedSeasonWeeklyStats: `stats_player_week_${priorSeason}.csv`,
  currentRoster: `roster_${season}.csv`,
};
const sourceFreshness = Object.fromEntries(Object.entries(freshnessFiles).map(([label, fileName]) => {
  const file = path.join(cacheDir, fileName);
  return [label, { fileName, fetchedAt: fs.existsSync(file) ? fs.statSync(file).mtime.toISOString() : null }];
}));
const currentMarketVariants = records.flatMap((record) => Object.values(record.variants)).filter((value) => value.marketRankSource === 'Sleeper current ADP').length;
const workbookFallbackVariants = records.flatMap((record) => Object.values(record.variants)).filter((value) => value.marketRankSource === 'Workbook fallback').length;
const output = {
  schemaVersion: '1.4.0',
  modelVersion: 'adp-edge-2026.5',
  dataVersion: `2026-${generatedAt.slice(0, 10).replaceAll('-', '')}`,
  season,
  generatedAt,
  workbook: { fileName: path.basename(workbookPath), sha256: sha256(workbookPath), sheets: Object.keys(sheetToVariant) },
  calibration: {
    method: 'Position-specific, winsorized model anchored to current market price. Every displayed Edge input is scored: current consensus projection rank, opportunity, sample-shrunk efficiency, two-year trajectory/reversion, QB floor, teammate/QB context, receiving concentration, current line performance/injuries, and projection availability.',
    window: backtest ? `Rolling no-lookahead ${backtest.seasons[0]}–${backtest.seasons.at(-1)} preseason folds; current build uses the latest completed season for player signals.` : 'Current build uses the latest completed season for player signals.',
    baseline: 'Current Sleeper scoring-specific ADP, with workbook rank used only when the current feed has no player/variant value.',
    backtest: backtest ? {
      accepted: backtest.accepted,
      edgeHitRate: backtest.aggregate.edgeHitRate,
      baselineHitRate: backtest.aggregate.baselineHitRate,
      lift: backtest.aggregate.lift,
      regressionRisk: backtest.regressionRiskAggregate,
      report: 'data/reports/backtest-2026.json',
    } : null,
    coverageTop240: round(matched / Math.max(1, top240.length) * 100, 1),
    limitation: 'The 0–100 Edge score is a comparison aid, not a literal probability. Injury/news state overrides an older projection; confirmed season-ending injuries are removed from active ranks. Workbook O-line, defense, and SOS values are display-only. Current line context is a reproducible proxy from 2025 team performance plus current OL injuries; current consensus projections absorb defense and schedule changes.',
  },
  draftAvailability: backtest?.draftAvailability || {
    modelVersion: 'return-risk-2026.5',
    method: 'Fallback position and market-band uncertainty used with roster-conditioned intervening-pick demand when the historical FantasyPros archive is unavailable.',
    uncertaintyByPosition: {
      QB: { early: 7, middle: 13, late: 22 },
      RB: { early: 6, middle: 12, late: 22 },
      WR: { early: 6, middle: 12, late: 23 },
      TE: { early: 8, middle: 15, late: 24 },
    },
  },
  positionWeights: POSITION_WEIGHTS,
  thresholds: {
    eligibility: 'Market/ranking position within top 240',
    strong: { score: 70, expectedPositionalGain: 3, completeness: 65 },
    watch: { score: 58, expectedPositionalGain: 1, completeness: 50 },
  },
  variantMap: sheetToVariant,
  sources: SOURCES,
  sourceFreshness,
  sourcePrecedence: ['current injury/news state', 'current Sleeper ADP and projections', 'current FantasyPros consensus projections', `${priorSeason} nflverse/NGS performance`, 'workbook display/fallback'],
  marketCoverage: { currentMarketVariants, workbookFallbackVariants },
  seasonEndingOverrides: records.filter((record) => record.availability.seasonEnding).map((record) => ({ name: record.name, position: record.position, team: record.team, injury: record.availability.injuryDetail, updatedAt: record.availability.newsUpdatedAt })),
  points,
  players: records,
};

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.mkdirSync(path.dirname(reportPath), { recursive: true });
fs.writeFileSync(outputPath, `${JSON.stringify(output, null, 2)}\n`);
fs.writeFileSync(reportPath, `${JSON.stringify({ season, generatedAt, top240Count: top240.length, matchedTop240: matched, coverageTop240: output.calibration.coverageTop240, unmatched }, null, 2)}\n`);
console.log(`Generated ${path.relative(ROOT, outputPath)} with ${records.length} scored players.`);
console.log(`Top-240 coverage: ${output.calibration.coverageTop240}% (${matched}/${top240.length}).`);
console.log(`Join exceptions: ${unmatched.length}; see ${path.relative(ROOT, reportPath)}.`);
