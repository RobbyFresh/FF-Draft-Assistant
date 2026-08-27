/* global window */
(function (root, factory) {
  const api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.DraftRoom = api;
})(typeof window !== 'undefined' ? window : globalThis, function () {
  const POSITIONS = ['QB', 'RB', 'WR', 'TE', 'K', 'DST'];
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

  const clamp = (value, low, high) => Math.max(low, Math.min(high, value));
  const normalizeName = (value) => String(value || '')
    .normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\b(jr|sr|ii|iii|iv)\b/g, '')
    .replace(/[^a-z0-9]/g, '');
  const normalizePosition = (value) => {
    const position = String(value || '').trim().toUpperCase();
    if (position === 'DEF') return 'DST';
    if (position === 'HB' || position === 'FB') return 'RB';
    return position;
  };
  const playerKey = (player) => player?.id
    ? `nfl:${player.id}`
    : `name:${normalizeName(player?.name)}|${normalizePosition(player?.position)}`;

  function clone(value) { return JSON.parse(JSON.stringify(value)); }

  function rosterTotal(roster) {
    return roster.reduce((sum, slot) => sum + Math.max(0, Number(slot.count) || 0), 0);
  }

  function normalizeRoster(roster) {
    const byType = new Map();
    for (const item of Array.isArray(roster) ? roster : DEFAULT_ROSTER) {
      const type = String(item?.type || '').toUpperCase();
      const count = clamp(Math.floor(Number(item?.count) || 0), 0, 30);
      if (!['QB', 'RB', 'WR', 'TE', 'FLEX', 'SFLX', 'K', 'DST', 'BENCH'].includes(type)) continue;
      byType.set(type, (byType.get(type) || 0) + count);
    }
    return ['QB', 'RB', 'WR', 'TE', 'FLEX', 'SFLX', 'K', 'DST', 'BENCH']
      .map((type) => ({ type, count: byType.get(type) || 0 }))
      .filter((slot) => slot.count > 0 || slot.type === 'SFLX');
  }

  function buildTeams(teamCount, userSlot, names = []) {
    return Array.from({ length: teamCount }, (_, index) => ({
      id: `team-${index + 1}`,
      slot: index + 1,
      name: String(names[index] || (index + 1 === userSlot ? 'My Team' : `Team ${index + 1}`)).trim() || `Team ${index + 1}`,
    }));
  }

  function makeProfile(input = {}) {
    const teamCount = clamp(Math.floor(Number(input.teamCount) || 12), 4, 20);
    const userSlot = clamp(Math.floor(Number(input.userSlot) || 1), 1, teamCount);
    const rosterTemplate = normalizeRoster(input.rosterTemplate);
    const teams = buildTeams(teamCount, userSlot, input.teamNames || input.teams?.map((team) => team.name));
    const now = new Date().toISOString();
    return {
      id: input.id || `league-${Date.now().toString(36)}`,
      schemaVersion: 1,
      season: Number(input.season) || 2026,
      name: String(input.name || 'My League').trim() || 'My League',
      scoring: ['standard', 'halfPpr', 'ppr', 'superflex'].includes(input.scoring) ? input.scoring : 'ppr',
      teamCount,
      userTeamId: `team-${userSlot}`,
      rounds: clamp(Math.floor(Number(input.rounds) || rosterTotal(rosterTemplate)), 1, 40),
      rosterTemplate,
      teams,
      createdAt: input.createdAt || now,
      updatedAt: now,
    };
  }

  function normalizeProfile(input) {
    return makeProfile({
      ...input,
      userSlot: Number(String(input?.userTeamId || '').replace('team-', '')) || input?.userSlot,
      teamNames: input?.teams?.map((team) => team.name),
    });
  }

  function teamForPick(overallPick, teamCount) {
    const overall = Math.max(1, Math.floor(Number(overallPick) || 1));
    const count = Math.max(1, Math.floor(Number(teamCount) || 1));
    const round = Math.floor((overall - 1) / count) + 1;
    const roundPick = ((overall - 1) % count) + 1;
    const slot = round % 2 === 1 ? roundPick : count - roundPick + 1;
    return { round, roundPick, slot, teamId: `team-${slot}` };
  }

  function nextPickForTeam(fromOverall, teamId, teamCount, totalPicks, includeCurrent = true) {
    const start = Math.max(1, Math.floor(Number(fromOverall) || 1) + (includeCurrent ? 0 : 1));
    for (let overall = start; overall <= totalPicks; overall += 1) {
      if (teamForPick(overall, teamCount).teamId === teamId) return overall;
    }
    return null;
  }

  // Zero-based row index for the player available at the user's next pick.
  // Example: from overall 46 to 51, pick 51 is the sixth visible player, index 5.
  function nextPickBoardIndex(currentOverall, nextUserPick) {
    const current = Math.max(1, Math.floor(Number(currentOverall) || 1));
    const next = Math.floor(Number(nextUserPick) || 0);
    return next >= current ? next - current : null;
  }

  function newSession(profile) {
    const now = new Date().toISOString();
    return {
      schemaVersion: 1,
      id: `draft-${Date.now().toString(36)}`,
      profileId: profile.id,
      status: 'active',
      nextOverall: 1,
      overrideTeamId: null,
      picks: [],
      watchlist: [],
      startedAt: now,
      updatedAt: now,
    };
  }

  function eligibleSlots(position) {
    const pos = normalizePosition(position);
    if (pos === 'QB') return ['QB', 'SFLX', 'BENCH'];
    if (pos === 'RB') return ['RB', 'FLEX', 'SFLX', 'BENCH'];
    if (pos === 'WR') return ['WR', 'FLEX', 'SFLX', 'BENCH'];
    if (pos === 'TE') return ['TE', 'FLEX', 'SFLX', 'BENCH'];
    if (pos === 'K') return ['K', 'BENCH'];
    if (pos === 'DST') return ['DST', 'BENCH'];
    return ['BENCH'];
  }

  function allocateRoster(picks, rosterTemplate) {
    const roster = normalizeRoster(rosterTemplate);
    const capacity = Object.fromEntries(roster.map((slot) => [slot.type, slot.count]));
    const assignments = Object.fromEntries(roster.map((slot) => [slot.type, []]));
    assignments.EXTRA = [];
    const ordered = [...(picks || [])].sort((a, b) => a.overall - b.overall);
    for (const pick of ordered) {
      const slot = eligibleSlots(pick.position).find((candidate) => (assignments[candidate]?.length || 0) < (capacity[candidate] || 0)) || 'EXTRA';
      assignments[slot].push(pick);
    }
    const remaining = Object.fromEntries(roster.map((slot) => [slot.type, Math.max(0, slot.count - assignments[slot.type].length)]));
    const positionCounts = Object.fromEntries(POSITIONS.map((position) => [position, ordered.filter((pick) => normalizePosition(pick.position) === position).length]));
    return { assignments, remaining, positionCounts };
  }

  function rosterNeeds(allocation, rosterTemplate) {
    return normalizeRoster(rosterTemplate)
      .filter((slot) => slot.type !== 'BENCH' && slot.count > 0)
      .map((slot) => {
        const filled = Math.min(slot.count, allocation?.assignments?.[slot.type]?.length || 0);
        return {
          type: slot.type,
          label: slot.type === 'FLEX' ? 'FLX' : slot.type,
          filled,
          total: slot.count,
          remaining: Math.max(0, slot.count - filled),
          complete: filled >= slot.count,
        };
      });
  }

  function picksForTeam(session, teamId) {
    return (session?.picks || []).filter((pick) => pick.teamId === teamId);
  }

  function hasStarterPath(profile, position) {
    const pos = normalizePosition(position);
    const configured = new Set(profile.rosterTemplate.filter((slot) => slot.count > 0).map((slot) => slot.type));
    if (pos === 'QB') return configured.has('QB') || configured.has('SFLX');
    if (['RB', 'WR', 'TE'].includes(pos)) return configured.has(pos) || configured.has('FLEX') || configured.has('SFLX');
    return configured.has(pos);
  }

  function roundUrgency(position, round, profile) {
    const pos = normalizePosition(position);
    const hasSuperflex = profile.rosterTemplate.some((slot) => slot.type === 'SFLX' && slot.count > 0);
    const hasDedicatedTe = profile.rosterTemplate.some((slot) => slot.type === 'TE' && slot.count > 0);
    const hasTeFlexPath = profile.rosterTemplate.some((slot) => ['FLEX', 'SFLX'].includes(slot.type) && slot.count > 0);
    if (!hasStarterPath(profile, pos)) return 0.22;
    if (pos === 'QB') return hasSuperflex ? (round <= 5 ? 1.2 : 1) : (round <= 2 ? 0.55 : round <= 8 ? 1 : 0.8);
    if (pos === 'TE' && !hasDedicatedTe) return hasTeFlexPath ? 1 : 0.22;
    if (pos === 'TE') return round <= 3 ? 0.7 : 0.95;
    if (pos === 'K' || pos === 'DST') return round <= 10 ? 0.05 : 0.75;
    return 1;
  }

  function teamNeedScore(profile, session, teamId, position, round, existingAllocation = null) {
    const pos = normalizePosition(position);
    const allocation = existingAllocation || allocateRoster(picksForTeam(session, teamId), profile.rosterTemplate);
    const hasDedicatedTe = profile.rosterTemplate.some((slot) => slot.type === 'TE' && slot.count > 0);
    const openPriorityStarters = ['QB', 'RB', 'WR', 'TE', 'FLEX', 'SFLX']
      .reduce((sum, type) => sum + (allocation.remaining[type] || 0), 0);
    let need = 0.06;
    if ((allocation.remaining[pos] || 0) > 0) need = 1;
    else if (pos === 'QB' && (allocation.remaining.SFLX || 0) > 0) need = 0.94;
    else if (['RB', 'WR'].includes(pos) && (allocation.remaining.FLEX || 0) > 0) need = 0.78;
    else if (pos === 'TE' && (allocation.remaining.FLEX || 0) > 0) need = hasDedicatedTe ? 0.66 : 0.78;
    else if (['RB', 'WR'].includes(pos) && (allocation.remaining.SFLX || 0) > 0) need = 0.52;
    else if (pos === 'TE' && (allocation.remaining.SFLX || 0) > 0) need = hasDedicatedTe ? 0.44 : 0.52;
    else if ((allocation.remaining.BENCH || 0) > 0) need = openPriorityStarters > 0 ? 0.1 : 0.28;
    const positionCount = allocation.positionCounts[pos] || 0;
    if (positionCount >= 4 && !(allocation.remaining[pos] || 0)) need *= 0.55;
    return clamp(need * roundUrgency(pos, round, profile), 0, 1.2);
  }

  // In a one-QB lineup, an early backup QB is much less likely than a team's
  // first. Tight end behaves similarly, though the discount is softer because
  // a second TE can still fill FLEX or a multi-TE lineup. Superflex preserves
  // demand for a second QB and only discounts the third.
  function duplicatePositionFactor(profile, session, teamId, position, round, existingAllocation = null) {
    const pos = normalizePosition(position);
    if (pos !== 'QB' && pos !== 'TE') return 1;
    const allocation = existingAllocation || allocateRoster(picksForTeam(session, teamId), profile.rosterTemplate);
    const positionCount = allocation.positionCounts[pos] || 0;
    if (positionCount === 0) return 1;
    const hasSuperflex = profile.rosterTemplate.some((slot) => slot.type === 'SFLX' && slot.count > 0);
    const hasDedicatedTe = profile.rosterTemplate.some((slot) => slot.type === 'TE' && slot.count > 0);
    if ((allocation.remaining[pos] || 0) > 0) return 1;
    if (pos === 'QB' && hasSuperflex && (allocation.remaining.SFLX || 0) > 0) return 1;
    if (pos === 'TE' && ((allocation.remaining.FLEX || 0) > 0 || (allocation.remaining.SFLX || 0) > 0)) {
      return 0.72;
    }
    if (pos === 'QB') {
      if (round <= 8) return positionCount === 1 ? 0.12 : 0.06;
      if (round <= 11) return positionCount === 1 ? 0.3 : 0.14;
      return positionCount === 1 ? 0.6 : 0.32;
    }
    if (round <= 8) return positionCount === 1 ? 0.28 : 0.12;
    if (round <= 11) return positionCount === 1 ? 0.5 : 0.24;
    return positionCount === 1 ? 0.75 : 0.42;
  }

  function teamSummary(profile, session, teamId) {
    const picks = picksForTeam(session, teamId);
    const allocation = allocateRoster(picks, profile.rosterTemplate);
    const recent = picks.slice(-3).map((pick) => normalizePosition(pick.position));
    let tendency = '';
    if (recent.length >= 2 && recent.at(-1) === recent.at(-2)) tendency = `${recent.at(-1)} ×${recent.filter((position) => position === recent.at(-1)).length}`;
    return { picks, ...allocation, tendency };
  }

  function currentTeam(profile, session) {
    const scheduled = teamForPick(session.nextOverall, profile.teamCount);
    return profile.teams.find((team) => team.id === (session.overrideTeamId || scheduled.teamId)) || profile.teams[scheduled.slot - 1];
  }

  function draftedPlayerKeys(session) { return new Set((session?.picks || []).map((pick) => pick.playerKey)); }

  function lastDraftedPick(session) {
    return (session?.picks || []).reduce((latest, pick) => {
      if (!latest) return pick;
      return Number(pick.overall) >= Number(latest.overall) ? pick : latest;
    }, null);
  }

  function recordPick(profile, session, player, options = {}) {
    const key = player.playerKey || playerKey(player);
    if (!key || draftedPlayerKeys(session).has(key)) return { session, error: 'Player is already drafted.' };
    const scheduled = teamForPick(session.nextOverall, profile.teamCount);
    const teamId = options.teamId || session.overrideTeamId || scheduled.teamId;
    if (!profile.teams.some((team) => team.id === teamId)) return { session, error: 'Unknown fantasy team.' };
    const pick = {
      overall: session.nextOverall,
      round: scheduled.round,
      roundPick: scheduled.roundPick,
      teamId,
      playerKey: key,
      playerId: player.id || null,
      name: player.name,
      position: normalizePosition(player.position),
      source: options.source || 'manual',
      timestamp: new Date().toISOString(),
    };
    const next = clone(session);
    next.picks.push(pick);
    next.nextOverall += 1;
    next.overrideTeamId = null;
    next.updatedAt = new Date().toISOString();
    if (next.nextOverall > profile.teamCount * profile.rounds) next.status = 'complete';
    return { session: next, pick };
  }

  function undoLastPick(session) {
    if (!session?.picks?.length) return { session, pick: null };
    const next = clone(session);
    const lastIndex = next.picks.reduce((winner, pick, index, array) => pick.overall > array[winner].overall ? index : winner, 0);
    const [pick] = next.picks.splice(lastIndex, 1);
    next.nextOverall = Math.min(next.nextOverall, pick.overall);
    next.status = 'active';
    next.overrideTeamId = null;
    next.updatedAt = new Date().toISOString();
    return { session: next, pick };
  }

  function undraftPlayer(session, targetPlayerKey) {
    const next = clone(session);
    const removed = next.picks.find((pick) => pick.playerKey === targetPlayerKey);
    if (!removed) return { session, pick: null };
    next.picks = next.picks.filter((pick) => pick.playerKey !== targetPlayerKey);
    if (removed.overall === next.nextOverall - 1) next.nextOverall = removed.overall;
    next.status = 'active';
    next.updatedAt = new Date().toISOString();
    return { session: next, pick: removed };
  }

  function movePlayer(session, targetPlayerKey, teamId) {
    const next = clone(session);
    const pick = next.picks.find((item) => item.playerKey === targetPlayerKey);
    if (!pick) return { session, pick: null };
    pick.teamId = teamId;
    pick.source = 'corrected';
    next.updatedAt = new Date().toISOString();
    return { session: next, pick };
  }

  function marketBand(rank) {
    if (rank <= 24) return 'early';
    if (rank <= 72) return 'middle';
    return 'late';
  }

  const FALLBACK_SIGMA = {
    QB: { early: 7, middle: 13, late: 22 },
    RB: { early: 6, middle: 12, late: 22 },
    WR: { early: 6, middle: 12, late: 23 },
    TE: { early: 8, middle: 15, late: 24 },
  };

  function projectionRankMap(players) {
    const projected = players
      .filter((player) => ['RB', 'WR', 'TE'].includes(normalizePosition(player.position)))
      .filter((player) => Number.isFinite(Number(player.projectedPoints)) && Number(player.projectedPoints) > 0)
      .filter((player) => Number.isFinite(Number(player.marketRank)) && Number(player.marketRank) > 0)
      .sort((a, b) => Number(b.projectedPoints) - Number(a.projectedPoints) || Number(a.marketRank) - Number(b.marketRank));
    const marketSlots = projected.map((player) => Number(player.marketRank)).sort((a, b) => a - b);
    return new Map(projected.map((player, index) => [player.playerKey, marketSlots[index]]));
  }

  function sortByDisplayRank(players) {
    const rank = (player) => {
      const boardRank = Number(player?.boardRank);
      if (Number.isFinite(boardRank) && boardRank > 0) return boardRank;
      const marketRank = Number(player?.marketRank);
      return Number.isFinite(marketRank) && marketRank > 0 ? marketRank : Number.POSITIVE_INFINITY;
    };
    return [...(players || [])].sort((a, b) => rank(a) - rank(b)
      || String(a?.name || '').localeCompare(String(b?.name || ''), undefined, { sensitivity: 'base' }));
  }

  // If an opponent has multiple picks before the user returns, later picks
  // should see a roster advanced by the earlier selection. Use configured
  // needs to choose a representative alternative player for that earlier
  // pick; this is an expected roster path, not a claim about the exact player.
  function expectedAlternativePosition(profile, allocation, round) {
    const preference = { RB: 1, WR: 0.99, QB: 0.94, TE: 0.86 };
    if (!profile.rosterTemplate.some((slot) => slot.type === 'TE' && slot.count > 0)) preference.TE = 0.68;
    return ['QB', 'RB', 'WR', 'TE'].map((position) => {
      const need = teamNeedScore(profile, null, 'projected-team', position, round, allocation)
        * duplicatePositionFactor(profile, null, 'projected-team', position, round, allocation);
      return { position, score: need * preference[position], count: allocation.positionCounts[position] || 0 };
    }).sort((a, b) => b.score - a.score || a.count - b.count
      || ['QB', 'TE', 'RB', 'WR'].indexOf(a.position) - ['QB', 'TE', 'RB', 'WR'].indexOf(b.position))[0]?.position || 'WR';
  }

  function buildReturnIntel(players, profile, session, availabilityModel = {}) {
    const current = session.nextOverall;
    const scheduled = teamForPick(current, profile.teamCount);
    const isUsersTurn = (session.overrideTeamId || scheduled.teamId) === profile.userTeamId;
    const totalPicks = profile.teamCount * profile.rounds;
    const nextUserPick = nextPickForTeam(current, profile.userTeamId, profile.teamCount, totalPicks, !isUsersTurn);
    if (!nextUserPick || nextUserPick <= current) return [];
    const intervening = [];
    const firstOpponentPick = isUsersTurn ? current + 1 : current;
    for (let overall = firstOpponentPick; overall < nextUserPick; overall += 1) intervening.push(teamForPick(overall, profile.teamCount));
    const drafted = draftedPlayerKeys(session);
    const validRank = (value) => value !== null && value !== undefined && String(value).trim() !== '' && Number.isFinite(Number(value)) && Number(value) > 0;
    const available = players.filter((player) => !drafted.has(player.playerKey) && ['QB', 'RB', 'WR', 'TE'].includes(normalizePosition(player.position)) && validRank(player.marketRank));
    const projectedFlexRanks = projectionRankMap(available);
    const hasDedicatedTe = profile.rosterTemplate.some((slot) => slot.type === 'TE' && slot.count > 0);
    const recent = session.picks.slice(-6);
    const userCoords = teamForPick(nextUserPick, profile.teamCount);
    const rosterAllocations = new Map(profile.teams.map((team) => [
      team.id,
      allocateRoster(picksForTeam(session, team.id), profile.rosterTemplate),
    ]));
    const projectedPicks = new Map(profile.teams.map((team) => [team.id, [...picksForTeam(session, team.id)]]));
    const interveningAllocations = intervening.map((pick) => {
      const teamPicks = projectedPicks.get(pick.teamId) || [];
      const allocation = allocateRoster(teamPicks, profile.rosterTemplate);
      const expectedPosition = expectedAlternativePosition(profile, allocation, pick.round);
      teamPicks.push({
        overall: pick.overall,
        teamId: pick.teamId,
        playerKey: `projected:${pick.overall}`,
        name: 'Projected intervening selection',
        position: expectedPosition,
      });
      projectedPicks.set(pick.teamId, teamPicks);
      return allocation;
    });

    return available.map((player) => {
      const pos = normalizePosition(player.position);
      const marketRank = Number(player.marketRank);
      const boardRank = validRank(player.boardRank) ? Number(player.boardRank) : marketRank;
      const projectionFlexRank = projectedFlexRanks.get(player.playerKey) || null;
      const formatRank = pos === 'TE' && !hasDedicatedTe && projectionFlexRank
        ? marketRank * 0.35 + projectionFlexRank * 0.65
        : marketRank;
      const formatBoardRank = pos === 'TE' && !hasDedicatedTe && projectionFlexRank
        ? boardRank * 0.35 + projectionFlexRank * 0.65
        : boardRank;
      const band = marketBand(formatRank);
      const sigma = Number(availabilityModel?.uncertaintyByPosition?.[pos]?.[band]) || FALLBACK_SIGMA[pos][band];
      const picksUntil = nextUserPick - current;
      const boardSlot = picksUntil + 1;
      const boardSpread = Math.max(3, Math.min(8, picksUntil * 0.9));
      const marketSpread = Math.max(5, Math.min(24, sigma * 0.65));
      const boardPressure = 1 / (1 + Math.exp(-(nextUserPick - formatBoardRank) / boardSpread));
      const marketPressure = 1 / (1 + Math.exp(-(nextUserPick - formatRank) / marketSpread));
      const rankPressure = validRank(player.boardRank)
        ? boardPressure * 0.62 + marketPressure * 0.38
        : marketPressure;
      const needScores = intervening.map((pick, index) => teamNeedScore(profile, session, pick.teamId, pos, pick.round, interveningAllocations[index]));
      const demandScores = intervening.map((pick, index) => clamp(
        needScores[index] * duplicatePositionFactor(profile, session, pick.teamId, pos, pick.round, interveningAllocations[index]),
        0,
        1.2,
      ));
      const needPressure = demandScores.length ? demandScores.reduce((sum, value) => sum + value, 0) / demandScores.length : 0;
      const needTeams = new Set(intervening.filter((pick, index) => demandScores[index] >= 0.65).map((pick) => pick.teamId)).size;
      const needPicks = demandScores.filter((score) => score >= 0.65).length;
      const baselineRankHazard = intervening.length
        ? 1 - Math.pow(1 - rankPressure, 1 / intervening.length)
        : rankPressure;
      const rosterRankPressure = 1 - intervening.reduce((probability, pick, index) => {
        // Rank/ADP estimates this player's baseline availability. Roster need
        // may scale that player-specific hazard, but must never create an
        // independent floor or substitute generic positional demand for the
        // chance that this particular player is selected.
        // ADP/board rank already reflects the league-wide chance that this
        // specific player is selected. Team need should move that baseline,
        // not nearly erase it whenever a roster has another open starter.
        const demandFactor = clamp(0.55 + demandScores[index] * 1.05, 0.55, 1.6);
        return probability * (1 - clamp(baselineRankHazard * demandFactor, 0, 0.58));
      }, 1);
      const sameTier = available.filter((candidate) => normalizePosition(candidate.position) === pos && candidate.tier && candidate.tier === player.tier).length;
      const tierCliff = player.tier ? (sameTier <= 1 ? 1 : sameTier === 2 ? 0.68 : sameTier === 3 ? 0.42 : 0.18) : 0.2;
      const runPressure = recent.length ? recent.filter((pick) => normalizePosition(pick.position) === pos).length / recent.length : 0;
      const tendencyPressure = intervening.length ? intervening.reduce((sum, pick, index) => {
        const teamPicks = picksForTeam(session, pick.teamId).slice(-4);
        const tendency = teamPicks.length ? teamPicks.filter((teamPick) => normalizePosition(teamPick.position) === pos).length / teamPicks.length : 0;
        return sum + tendency * demandScores[index];
      }, 0) / intervening.length : 0;
      const contextMultiplier = 1 + tierCliff * 0.12 + runPressure * 0.08 + tendencyPressure * 0.05;
      const risk = Math.round(clamp(rosterRankPressure * contextMultiplier * 100, 2, 97));
      const rosterFit = teamNeedScore(profile, session, profile.userTeamId, pos, scheduled.round, rosterAllocations.get(profile.userTeamId));
      const priority = Math.round(clamp(risk * 0.72 + tierCliff * 18 + rosterFit * 10, 0, 100));
      const label = risk >= 70 ? 'take-now' : risk >= 60 ? 'pivot' : 'can-wait';
      const reasons = [
        `${intervening.length} opponent selection${intervening.length === 1 ? '' : 's'} occur before your ${userCoords.round}.${userCoords.roundPick} pick.`,
        `${needPicks} ${pos}-need opportunit${needPicks === 1 ? 'y' : 'ies'} across ${needTeams} intervening team${needTeams === 1 ? '' : 's'}; board rank ${Math.round(boardRank * 10) / 10}, market ADP ${Math.round(marketRank * 10) / 10}, next pick #${nextUserPick}.`,
      ];
      if (pos === 'TE' && !hasDedicatedTe) {
        reasons.push(projectionFlexRank
          ? `No dedicated TE slot: projected points imply flex-demand rank ${Math.round(projectionFlexRank * 10) / 10} (adjusted rank ${Math.round(formatRank * 10) / 10}).`
          : 'No dedicated TE slot: TE demand is limited to open flex or bench spots.');
      }
      if (tierCliff >= 0.65) reasons.push(sameTier <= 1 ? `Last available ${pos} in this tier.` : `Only ${sameTier} ${pos}s remain in this tier.`);
      if (runPressure >= 0.5) reasons.push(`${pos} run: ${recent.filter((pick) => normalizePosition(pick.position) === pos).length} of the last ${recent.length} picks.`);
      return { ...player, risk, priority, label, nextUserPick, picksUntil, boardSlot, opponentPicks: intervening.length, needTeams, needPicks, needPressure, rankPressure, rosterRankPressure, tierCliff, projectionFlexRank, formatRank, reasons: reasons.slice(0, 3) };
    }).sort((a, b) => b.priority - a.priority || b.risk - a.risk || a.marketRank - b.marketRank);
  }

  function parseImportText(text, players, profile, startOverall, existingPicks = []) {
    const byName = new Map();
    for (const player of players) {
      const key = normalizeName(player.name);
      if (!byName.has(key)) byName.set(key, []);
      byName.get(key).push(player);
    }
    const existing = new Set(existingPicks.map((pick) => pick.playerKey));
    const seen = new Set();
    return String(text || '').split(/\r?\n/).map((raw) => raw.trim()).filter(Boolean).map((raw, index) => {
      const overall = startOverall + index;
      let line = raw.replace(/^\s*(?:#?\d+|\d+\.\d+)\s*[.)-]?\s*/, '').trim();
      let explicitTeam = null;
      for (const team of profile.teams) {
        const escaped = team.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const match = line.match(new RegExp(`^${escaped}\\s*[-–—,:|\\t]+\\s*(.+)$`, 'i'));
        if (match) { explicitTeam = team; line = match[1].trim(); break; }
      }
      line = line.replace(/\s+\((?:QB|RB|WR|TE|K|DST|DEF)(?:\s*[-,][^)]+)?\)\s*$/i, '').trim();
      const exact = byName.get(normalizeName(line)) || [];
      const fuzzy = exact.length ? exact : players.filter((player) => {
        const normalized = normalizeName(player.name);
        const candidate = normalizeName(line);
        return candidate.length >= 5 && (normalized.includes(candidate) || candidate.includes(normalized));
      });
      const player = fuzzy.length === 1 ? fuzzy[0] : null;
      const scheduled = teamForPick(overall, profile.teamCount);
      const key = player?.playerKey || (player ? playerKey(player) : null);
      let status = player ? 'ready' : fuzzy.length > 1 ? 'ambiguous' : 'unmatched';
      if (key && (existing.has(key) || seen.has(key))) status = 'duplicate';
      if (key) seen.add(key);
      return {
        raw,
        overall,
        teamId: explicitTeam?.id || scheduled.teamId,
        player,
        playerKey: key,
        status,
        candidates: fuzzy.slice(0, 5).map((candidate) => ({ playerKey: candidate.playerKey, name: candidate.name, position: candidate.position })),
      };
    });
  }

  function captureLegacySelections(storage, season) {
    const key = `draftWizard:legacySelections:${season}:v1`;
    const old = storage.getItem(key);
    if (old) {
      try { return JSON.parse(old); } catch { /* rebuild below */ }
    }
    const mine = [];
    const taken = [];
    for (let index = 0; index < storage.length; index += 1) {
      const storageKey = storage.key(index);
      if (!storageKey) continue;
      if (storageKey.startsWith('myteam:')) mine.push({ name: storageKey.slice(7) });
      if (storageKey.startsWith('taken:')) taken.push({ name: storageKey.slice(6), timestamp: Number(storage.getItem(storageKey)) || null });
    }
    const snapshot = { season: Number(season), capturedAt: new Date().toISOString(), mine, taken };
    if (mine.length || taken.length) storage.setItem(key, JSON.stringify(snapshot));
    return snapshot;
  }

  class DraftStore {
    constructor(storage, season) {
      this.storage = storage;
      this.season = String(season);
      this.profileKey = `draftWizard:leagueProfiles:${this.season}:v1`;
      this.sessionKey = `draftWizard:draftSessions:${this.season}:v1`;
      this.activeKey = `draftWizard:activeLeague:${this.season}:v1`;
    }
    read(key, fallback) {
      try { return JSON.parse(this.storage.getItem(key)) ?? fallback; } catch { return fallback; }
    }
    write(key, value) { this.storage.setItem(key, JSON.stringify(value)); }
    profiles() { return this.read(this.profileKey, []).map(normalizeProfile); }
    activeProfileId() { return this.storage.getItem(this.activeKey); }
    activeProfile() {
      const profiles = this.profiles();
      return profiles.find((profile) => profile.id === this.activeProfileId()) || profiles[0] || null;
    }
    saveProfile(input) {
      const profile = makeProfile({ ...input, season: Number(this.season) });
      const profiles = this.profiles();
      const index = profiles.findIndex((candidate) => candidate.id === profile.id);
      if (index >= 0) profiles[index] = profile; else profiles.push(profile);
      this.write(this.profileKey, profiles);
      this.storage.setItem(this.activeKey, profile.id);
      return profile;
    }
    setActiveProfile(profileId) { this.storage.setItem(this.activeKey, profileId); }
    sessions() { return this.read(this.sessionKey, {}); }
    session(profileId) { return this.sessions()[profileId] || null; }
    saveSession(session) {
      const sessions = this.sessions();
      sessions[session.profileId] = session;
      this.write(this.sessionKey, sessions);
      return session;
    }
    ensureSession(profile) { return this.session(profile.id) || this.saveSession(newSession(profile)); }
    clearSession(profile) {
      const sessions = this.sessions();
      delete sessions[profile.id];
      this.write(this.sessionKey, sessions);
      return this.saveSession(newSession(profile));
    }
    deleteProfile(profileId) {
      const profiles = this.profiles().filter((profile) => profile.id !== profileId);
      this.write(this.profileKey, profiles);
      const sessions = this.sessions();
      delete sessions[profileId];
      this.write(this.sessionKey, sessions);
      if (this.activeProfileId() === profileId) {
        if (profiles[0]) this.storage.setItem(this.activeKey, profiles[0].id);
        else this.storage.removeItem(this.activeKey);
      }
    }
  }

  return {
    DEFAULT_ROSTER,
    DraftStore,
    allocateRoster,
    buildReturnIntel,
    captureLegacySelections,
    currentTeam,
    draftedPlayerKeys,
    eligibleSlots,
    lastDraftedPick,
    makeProfile,
    movePlayer,
    newSession,
    nextPickBoardIndex,
    nextPickForTeam,
    normalizeName,
    normalizePosition,
    parseImportText,
    playerKey,
    recordPick,
    rosterNeeds,
    sortByDisplayRank,
    teamForPick,
    teamNeedScore,
    teamSummary,
    undoLastPick,
    undraftPlayer,
  };
});
