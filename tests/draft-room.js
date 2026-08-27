const assert = require('assert');
const DraftRoom = require('../public/draft-room.js');

class MemoryStorage {
  constructor() { this.data = new Map(); }
  get length() { return this.data.size; }
  key(index) { return [...this.data.keys()][index] ?? null; }
  getItem(key) { return this.data.has(key) ? this.data.get(key) : null; }
  setItem(key, value) { this.data.set(key, String(value)); }
  removeItem(key) { this.data.delete(key); }
}

const profile = DraftRoom.makeProfile({
  id: 'league-test',
  season: 2026,
  name: 'Test League',
  teamCount: 12,
  userSlot: 4,
  rounds: 17,
  scoring: 'ppr',
  rosterTemplate: DraftRoom.DEFAULT_ROSTER,
});

assert.equal(DraftRoom.teamForPick(1, 12).teamId, 'team-1');
assert.equal(DraftRoom.teamForPick(12, 12).teamId, 'team-12');
assert.equal(DraftRoom.teamForPick(13, 12).teamId, 'team-12');
assert.equal(DraftRoom.teamForPick(24, 12).teamId, 'team-1');
assert.equal(DraftRoom.teamForPick(25, 12).teamId, 'team-1');
assert.equal(DraftRoom.nextPickForTeam(5, 'team-4', 12, 204, true), 21);
assert.equal(DraftRoom.nextPickForTeam(4, 'team-4', 12, 204, false), 21);
assert.equal(DraftRoom.nextPickBoardIndex(46, 51), 5, 'pick #51 must be the sixth available row when #46 is on the clock');

let session = DraftRoom.newSession(profile);
assert.deepEqual(session.watchlist, [], 'a new draft must start with an empty manual return-risk watchlist');
const gibbs = { id: 'gibbs', name: 'Jahmyr Gibbs', position: 'RB' };
let result = DraftRoom.recordPick(profile, session, gibbs);
assert.equal(result.pick.teamId, 'team-1');
assert.equal(result.session.nextOverall, 2);
session = result.session;
assert.equal(DraftRoom.lastDraftedPick(session).name, 'Jahmyr Gibbs');
session.overrideTeamId = 'team-9';
result = DraftRoom.recordPick(profile, session, { id: 'bijan', name: 'Bijan Robinson', position: 'RB' });
assert.equal(result.pick.teamId, 'team-9', 'one-pick override should supersede scheduled team');
assert.equal(result.session.overrideTeamId, null, 'override should clear after one pick');
session = result.session;
assert.equal(DraftRoom.lastDraftedPick(session).name, 'Bijan Robinson');
assert(DraftRoom.recordPick(profile, session, gibbs).error, 'duplicate canonical player must be rejected');
assert(!DraftRoom.draftedPlayerKeys(session).has('nfl:gibbs-defender'), 'same-name players with different NFL IDs must remain distinct');

let corrected = DraftRoom.movePlayer(session, 'nfl:gibbs', 'team-4');
assert.equal(corrected.pick.teamId, 'team-4');
let undone = DraftRoom.undoLastPick(corrected.session);
assert.equal(undone.pick.name, 'Bijan Robinson');
assert.equal(undone.session.nextOverall, 2);
assert.equal(DraftRoom.lastDraftedPick(undone.session).name, 'Jahmyr Gibbs', 'undo must reveal the previous last-drafted player');
let undrafted = DraftRoom.undraftPlayer(corrected.session, 'nfl:gibbs');
assert.equal(undrafted.pick.name, 'Jahmyr Gibbs');

const smallProfile = DraftRoom.makeProfile({
  id: 'small', teamCount: 4, userSlot: 1, rounds: 8, scoring: 'ppr',
  rosterTemplate: [{ type: 'QB', count: 1 }, { type: 'RB', count: 2 }, { type: 'WR', count: 2 }, { type: 'TE', count: 1 }, { type: 'FLEX', count: 1 }, { type: 'BENCH', count: 3 }],
});
const teamPicks = [
  { overall: 1, teamId: 'team-1', playerKey: 'mine', name: 'Mine', position: 'WR' },
  { overall: 2, teamId: 'team-2', playerKey: 'rb2', name: 'RB Two', position: 'RB' },
  { overall: 3, teamId: 'team-3', playerKey: 'rb3', name: 'RB Three', position: 'RB' },
  { overall: 4, teamId: 'team-4', playerKey: 'rb4', name: 'RB Four', position: 'RB' },
];
const allocation = DraftRoom.allocateRoster(teamPicks.filter((pick) => pick.teamId === 'team-2').concat([
  { overall: 5, teamId: 'team-2', playerKey: 'rb2b', name: 'RB Two B', position: 'RB' },
  { overall: 6, teamId: 'team-2', playerKey: 'wr2', name: 'WR Two', position: 'WR' },
]), smallProfile.rosterTemplate);
assert.equal(allocation.assignments.RB.length, 2);
assert.equal(allocation.assignments.WR.length, 1);
assert.equal(allocation.remaining.FLEX, 1);

const needsRoster = [
  { type: 'QB', count: 1 }, { type: 'RB', count: 2 }, { type: 'WR', count: 2 },
  { type: 'TE', count: 1 }, { type: 'FLEX', count: 2 }, { type: 'BENCH', count: 4 },
];
const needsAllocation = DraftRoom.allocateRoster([
  { overall: 1, position: 'QB' },
  { overall: 2, position: 'RB' }, { overall: 3, position: 'RB' },
  { overall: 4, position: 'WR' }, { overall: 5, position: 'WR' },
  { overall: 6, position: 'WR' }, { overall: 7, position: 'WR' },
], needsRoster);
assert.deepEqual(
  DraftRoom.rosterNeeds(needsAllocation, needsRoster).map((slot) => [slot.label, slot.filled, slot.total, slot.complete]),
  [['QB', 1, 1, true], ['RB', 2, 2, true], ['WR', 2, 2, true], ['TE', 0, 1, false], ['FLX', 2, 2, true]],
  'league needs must report filled configured slots rather than raw position totals',
);

const noTeRoster = [
  { type: 'QB', count: 1 }, { type: 'RB', count: 1 }, { type: 'WR', count: 1 },
  { type: 'FLEX', count: 1 }, { type: 'SFLX', count: 1 }, { type: 'BENCH', count: 3 },
];
const noTeAllocation = DraftRoom.allocateRoster([
  { overall: 1, position: 'QB' }, { overall: 2, position: 'QB' }, { overall: 3, position: 'TE' },
], noTeRoster);
const noTeNeeds = DraftRoom.rosterNeeds(noTeAllocation, noTeRoster);
assert(!noTeNeeds.some((slot) => slot.type === 'TE'), 'a disabled TE slot must not appear in League Needs');
assert.equal(noTeNeeds.find((slot) => slot.type === 'FLEX').filled, 1, 'TE may fill an enabled FLEX slot');
assert.equal(noTeNeeds.find((slot) => slot.type === 'SFLX').filled, 1, 'a second QB must fill an enabled SFLX slot');

const riskPlayers = [
  { playerKey: 'target', name: 'Target Back', position: 'RB', marketRank: 10, edgeScore: 70, tier: 'Tier 2' },
  { playerKey: 'other', name: 'Other Back', position: 'RB', marketRank: 14, edgeScore: 55, tier: 'Tier 3' },
  { playerKey: 'late', name: 'Late Back', position: 'RB', marketRank: 90, edgeScore: 70, tier: 'Tier 8' },
  { playerKey: 'missing-rank', name: 'Missing Rank', position: 'RB', marketRank: null, boardRank: null, tier: null },
  { playerKey: 'zero-rank', name: 'Zero Rank', position: 'WR', marketRank: 0, boardRank: 0, tier: null },
];
const needySession = { ...DraftRoom.newSession(smallProfile), nextOverall: 9, picks: teamPicks };
const filledSession = { ...needySession, picks: [...teamPicks] };
for (const teamId of ['team-2', 'team-3', 'team-4']) {
  for (let index = 0; index < 2; index += 1) filledSession.picks.push({ overall: 20 + filledSession.picks.length, teamId, playerKey: `${teamId}-filled-${index}`, name: 'Filled RB', position: 'RB' });
}
const needyIntel = DraftRoom.buildReturnIntel(riskPlayers, smallProfile, needySession);
const filledIntel = DraftRoom.buildReturnIntel(riskPlayers, smallProfile, filledSession);
assert(!needyIntel.some((player) => ['missing-rank', 'zero-rank'].includes(player.playerKey)), 'null and zero ranks must not enter return-risk ordering');
assert(needyIntel.find((player) => player.playerKey === 'target').risk > filledIntel.find((player) => player.playerKey === 'target').risk, 'opponent RB needs should increase return risk');
assert(needyIntel.find((player) => player.playerKey === 'target').risk > needyIntel.find((player) => player.playerKey === 'late').risk, 'earlier ADP should increase return risk');

// Regression for the turn shown in the live board: user is on #46 and next
// selects at #51, while teams 1 and 2 each pick twice and still need a QB.
const turnProfile = DraftRoom.makeProfile({
  id: 'turn-regression', teamCount: 12, userSlot: 3, rounds: 16, scoring: 'ppr', rosterTemplate: DraftRoom.DEFAULT_ROSTER,
});
const turnSession = {
  ...DraftRoom.newSession(turnProfile),
  nextOverall: 46,
  picks: [
    { overall: 1, teamId: 'team-1', playerKey: 't1-rb', name: 'T1 RB', position: 'RB' },
    { overall: 2, teamId: 'team-1', playerKey: 't1-wr1', name: 'T1 WR1', position: 'WR' },
    { overall: 3, teamId: 'team-1', playerKey: 't1-wr2', name: 'T1 WR2', position: 'WR' },
    { overall: 4, teamId: 'team-2', playerKey: 't2-rb', name: 'T2 RB', position: 'RB' },
    { overall: 5, teamId: 'team-2', playerKey: 't2-wr1', name: 'T2 WR1', position: 'WR' },
    { overall: 6, teamId: 'team-2', playerKey: 't2-wr2', name: 'T2 WR2', position: 'WR' },
    { overall: 43, teamId: 'team-3', playerKey: 'my-qb', name: 'My QB', position: 'QB' },
    { overall: 44, teamId: 'team-3', playerKey: 'my-rb1', name: 'My RB1', position: 'RB' },
    { overall: 45, teamId: 'team-3', playerKey: 'my-rb2', name: 'My RB2', position: 'RB' },
  ],
};
const turnPlayers = [
  { playerKey: 'jayden', name: 'Jayden Daniels', position: 'QB', marketRank: 64.5, boardRank: 51, tier: 'Tier 5' },
  { playerKey: 'luther', name: 'Luther Burden III', position: 'WR', marketRank: 42.5, boardRank: 47, tier: 'Tier 5' },
  ...Array.from({ length: 4 }, (_, index) => ({ playerKey: `turn-qb-${index}`, name: `Turn QB ${index}`, position: 'QB', marketRank: 70 + index, boardRank: 60 + index, tier: 'Tier 5' })),
  ...Array.from({ length: 4 }, (_, index) => ({ playerKey: `turn-wr-${index}`, name: `Turn WR ${index}`, position: 'WR', marketRank: 70 + index, boardRank: 60 + index, tier: 'Tier 5' })),
];
const turnIntel = DraftRoom.buildReturnIntel(turnPlayers, turnProfile, turnSession, {
  uncertaintyByPosition: { QB: { middle: 22 }, WR: { middle: 21 } },
});
const jaydenRisk = turnIntel.find((player) => player.playerKey === 'jayden');
const lutherRisk = turnIntel.find((player) => player.playerKey === 'luther');
assert.equal(jaydenRisk.picksUntil, 5, '#51 is five selections away from #46');
assert.equal(jaydenRisk.boardSlot, 6, '#51 must be displayed as the sixth available row');
assert.equal(jaydenRisk.opponentPicks, 4, 'four opponent picks occur after the current user pick');
assert.equal(jaydenRisk.needTeams, 2);
assert.equal(jaydenRisk.needPicks, 4, 'both QB-needy teams pick twice before #51');
assert(jaydenRisk.risk < 70 && jaydenRisk.risk >= 60);
assert.equal(jaydenRisk.label, 'pivot', '60–69% urgency must remain a pivot under the stricter threshold');
assert(lutherRisk.risk >= 70);
assert.equal(lutherRisk.label, 'take-now', 'a WR already above the next-pick rank should not be a generic pivot');

// Regression for the #51 screenshot: broad WR/FLEX demand must not be
// mistaken for the probability that a specific WR with ADP 155.1 is taken
// during the 18 opponent selections before pick #70.
const striblingProfile = DraftRoom.makeProfile({
  id: 'stribling-regression', teamCount: 12, userSlot: 3, rounds: 16, scoring: 'ppr',
  rosterTemplate: [
    { type: 'QB', count: 1 }, { type: 'RB', count: 2 }, { type: 'WR', count: 2 },
    { type: 'TE', count: 1 }, { type: 'FLEX', count: 2 }, { type: 'K', count: 1 },
    { type: 'DST', count: 1 }, { type: 'BENCH', count: 6 },
  ],
});
const striblingSession = { ...DraftRoom.newSession(striblingProfile), nextOverall: 51 };
const stribling = DraftRoom.buildReturnIntel([{
  playerKey: 'stribling', name: "De'Zhaun Stribling", position: 'WR',
  marketRank: 155.1, boardRank: 155.1, projectedPoints: 128.4, tier: 'Tier 10',
}], striblingProfile, striblingSession, {
  uncertaintyByPosition: { WR: { late: 40 } },
})[0];
assert.equal(stribling.nextUserPick, 70);
assert.equal(stribling.opponentPicks, 18, 'only picks #52–69 occur before the user selects at #70');
assert.equal(stribling.label, 'can-wait');
assert(stribling.risk < 10, `ADP 155.1 must remain low risk at pick #51, received ${stribling.risk}%`);
assert(stribling.reasons[0].startsWith('18 opponent selections'), 'the card must not count the current user pick as an opponent selection');

// The screenshot case: both intervening teams already roster a QB. Market and
// board rank still matter, but four early backup-QB opportunities should not
// overpower the actual rosters and produce a Take Now recommendation.
const filledTurnSession = {
  ...turnSession,
  picks: turnSession.picks.concat([
    { overall: 10, teamId: 'team-1', playerKey: 't1-qb', name: 'T1 QB', position: 'QB' },
    { overall: 11, teamId: 'team-2', playerKey: 't2-qb', name: 'T2 QB', position: 'QB' },
  ]),
};
const burrow = { playerKey: 'burrow', name: 'Joe Burrow', position: 'QB', marketRank: 58, boardRank: 45, tier: 'Tier 4' };
const burrowIntel = DraftRoom.buildReturnIntel([burrow], turnProfile, filledTurnSession, {
  uncertaintyByPosition: { QB: { middle: 22 } },
})[0];
assert.equal(burrowIntel.needTeams, 0, 'teams with a QB should not count as early backup-QB needs');
assert.equal(burrowIntel.needPicks, 0, 'repeated picks by QB-filled teams should stay discounted');
assert(burrowIntel.risk < 60, `filled QB rooms should make Burrow a wait, received ${burrowIntel.risk}%`);
assert.equal(burrowIntel.label, 'can-wait');

const tightEnds = [
  { playerKey: 'target-te', name: 'Target TE', position: 'TE', marketRank: 43, boardRank: 46, tier: 'Tier 4' },
  { playerKey: 'other-te', name: 'Other TE', position: 'TE', marketRank: 55, boardRank: 55, tier: 'Tier 4' },
];
const openTeIntel = DraftRoom.buildReturnIntel(tightEnds, turnProfile, turnSession)[0];
const filledTeSession = {
  ...turnSession,
  picks: turnSession.picks.concat([
    { overall: 12, teamId: 'team-1', playerKey: 't1-te', name: 'T1 TE', position: 'TE' },
    { overall: 13, teamId: 'team-2', playerKey: 't2-te', name: 'T2 TE', position: 'TE' },
    { overall: 14, teamId: 'team-1', playerKey: 't1-te2', name: 'T1 TE2', position: 'TE' },
    { overall: 15, teamId: 'team-2', playerKey: 't2-te2', name: 'T2 TE2', position: 'TE' },
  ]),
};
const filledTeIntel = DraftRoom.buildReturnIntel(tightEnds, turnProfile, filledTeSession)[0];
assert(openTeIntel.risk > filledTeIntel.risk, 'filled configured TE slots should reduce early/middle-round backup-TE risk');
assert.equal(filledTeIntel.needTeams, 0, 'a backup TE should not be treated like an open configured TE need');

const lineupProfile = DraftRoom.makeProfile({
  id: 'lineup-priority', teamCount: 4, userSlot: 1, rounds: 10, scoring: 'ppr',
  rosterTemplate: [
    { type: 'QB', count: 1 }, { type: 'RB', count: 1 }, { type: 'WR', count: 1 },
    { type: 'TE', count: 1 }, { type: 'FLEX', count: 1 }, { type: 'BENCH', count: 4 },
  ],
});
const lineupSession = {
  ...DraftRoom.newSession(lineupProfile),
  picks: [
    { overall: 1, teamId: 'team-2', playerKey: 'lp-qb', position: 'QB' },
    { overall: 2, teamId: 'team-2', playerKey: 'lp-rb1', position: 'RB' },
    { overall: 3, teamId: 'team-2', playerKey: 'lp-rb2', position: 'RB' },
    { overall: 4, teamId: 'team-2', playerKey: 'lp-wr', position: 'WR' },
  ],
};
assert(
  DraftRoom.teamNeedScore(lineupProfile, lineupSession, 'team-2', 'TE', 5)
    > DraftRoom.teamNeedScore(lineupProfile, lineupSession, 'team-2', 'QB', 5),
  'a missing TE starter must outweigh a backup QB once FLEX is filled',
);

const superflexProfile = DraftRoom.makeProfile({
  id: 'superflex-priority', teamCount: 4, userSlot: 1, rounds: 10, scoring: 'superflex',
  rosterTemplate: [{ type: 'QB', count: 1 }, { type: 'SFLX', count: 1 }, { type: 'BENCH', count: 4 }],
});
const oneQbSession = {
  ...DraftRoom.newSession(superflexProfile),
  picks: [{ overall: 1, teamId: 'team-2', playerKey: 'sf-qb1', position: 'QB' }],
};
assert(
  DraftRoom.teamNeedScore(superflexProfile, oneQbSession, 'team-2', 'QB', 4) > 0.85,
  'an open SFLX slot must preserve strong second-QB demand',
);

const noQbProfile = DraftRoom.makeProfile({
  id: 'no-qb-priority', teamCount: 4, userSlot: 1, rounds: 8, scoring: 'ppr',
  rosterTemplate: [{ type: 'RB', count: 1 }, { type: 'WR', count: 1 }, { type: 'FLEX', count: 1 }, { type: 'BENCH', count: 4 }],
});
assert(
  DraftRoom.teamNeedScore(noQbProfile, DraftRoom.newSession(noQbProfile), 'team-2', 'QB', 4)
    < DraftRoom.teamNeedScore(noQbProfile, DraftRoom.newSession(noQbProfile), 'team-2', 'RB', 4),
  'QB must be bench-only demand when neither QB nor SFLX is configured',
);

const noTeProfile = DraftRoom.makeProfile({
  id: 'no-te-risk', teamCount: 4, userSlot: 1, rounds: 10, scoring: 'ppr',
  rosterTemplate: [
    { type: 'QB', count: 1 }, { type: 'RB', count: 1 }, { type: 'WR', count: 1 },
    { type: 'FLEX', count: 2 }, { type: 'BENCH', count: 4 },
  ],
});
const dedicatedTeProfile = DraftRoom.makeProfile({
  ...noTeProfile,
  id: 'dedicated-te-risk',
  rosterTemplate: [...noTeProfile.rosterTemplate, { type: 'TE', count: 1 }],
});
const formatSession = { ...DraftRoom.newSession(noTeProfile), nextOverall: 9 };
const formatPlayers = [
  { playerKey: 'high-proj-te', name: 'High Projection TE', position: 'TE', marketRank: 30, boardRank: 30, projectedPoints: 250, tier: 'Tier 4' },
  { playerKey: 'low-proj-te', name: 'Low Projection TE', position: 'TE', marketRank: 10, boardRank: 10, projectedPoints: 80, tier: 'Tier 4' },
  { playerKey: 'projection-rb', name: 'Projection RB', position: 'RB', marketRank: 12, boardRank: 12, projectedPoints: 200, tier: 'Tier 4' },
  { playerKey: 'projection-wr', name: 'Projection WR', position: 'WR', marketRank: 18, boardRank: 18, projectedPoints: 150, tier: 'Tier 4' },
];
const noTeFormatIntel = DraftRoom.buildReturnIntel(formatPlayers, noTeProfile, formatSession);
const dedicatedTeIntel = DraftRoom.buildReturnIntel(formatPlayers, dedicatedTeProfile, { ...formatSession, profileId: dedicatedTeProfile.id });
const noTeHigh = noTeFormatIntel.find((player) => player.playerKey === 'high-proj-te');
const noTeLow = noTeFormatIntel.find((player) => player.playerKey === 'low-proj-te');
assert(noTeHigh.risk > noTeLow.risk, 'without a TE slot, projected flex points must outweigh TE ADP when they disagree sharply');
assert(noTeHigh.reasons.some((reason) => reason.includes('No dedicated TE slot')), 'no-TE adjustments must be auditable in the player explanation');
assert(
  dedicatedTeIntel.find((player) => player.playerKey === 'high-proj-te').needPressure > noTeHigh.needPressure,
  'a dedicated TE slot must create more positional pressure than flex-only TE eligibility',
);

const importPlayers = riskPlayers.concat([{ playerKey: 'nfl:gibbs', id: 'gibbs', name: 'Jahmyr Gibbs', position: 'RB' }]);
const imported = DraftRoom.parseImportText('Team 2 — Jahmyr Gibbs\nTarget Back\nUnknown Person', importPlayers, smallProfile, 1, []);
assert.equal(imported[0].teamId, 'team-2');
assert.equal(imported[0].status, 'ready');
assert.equal(imported[1].teamId, 'team-2', 'player-only lines should infer snake order');
assert.equal(imported[2].status, 'unmatched');

const storage = new MemoryStorage();
storage.setItem('taken:Old Player', '123');
const legacy = DraftRoom.captureLegacySelections(storage, 2026);
assert.equal(legacy.taken[0].name, 'Old Player');
const store = new DraftRoom.DraftStore(storage, 2026);
store.saveProfile(profile);
assert.equal(store.activeProfile().name, 'Test League');
store.saveSession(DraftRoom.newSession(profile));
assert(store.session(profile.id));

console.log('Validated snake order, roster allocation, pick corrections, import parsing, storage, and return-risk behavior.');
