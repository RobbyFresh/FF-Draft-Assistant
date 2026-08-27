# 2026 Fantasy Football Draft Wizard

This site serves the supplied `Fantasy Football Cheat Sheet with Boom Outlier 2026.xlsx` as a four-tab draft board and layers an offline, auditable ADP Edge model over the creator's data. The workbook remains unchanged; generated research lives in `public/data/adp-edge-2026.json`.

## Run locally

```powershell
npm install
npm run build:edge
npm run backtest:edge
npm test
npm start
```

Open `http://localhost:3000`. The Express and Vercel workbook APIs both return `season`, `dataVersion`, `fileName`, and the original `sheets` array.

## Refresh the Edge data

```powershell
node scripts/build-edge-data.mjs --season 2026 --workbook "Fantasy Football Cheat Sheet with Boom Outlier 2026.xlsx" --refresh
```

The refresh downloads public files into ignored `data/cache/`, validates non-empty downloads, joins by NFL GSIS ID when possible and normalized name/position otherwise, and applies this precedence: current injury/news state, current scoring-specific Sleeper ADP and projections, current FantasyPros consensus projections, completed-season nflverse/NGS data, then workbook display/fallback values. A newer confirmed season-ending injury removes the player from active ranks and zeros the projection.

- `public/data/adp-edge-2026.json` — the versioned browser dataset.
- `data/reports/unmatched-2026.json` — auditable join exceptions and coverage.

Omit `--refresh` to reuse the cache. The browser never calls a third-party service during a draft.

## Draft-board behavior

- `League Setup` saves reusable, browser-local league profiles with scoring, snake position, team names, rounds, and one shared roster template. Profiles and draft sessions are isolated, so clearing a draft does not erase league settings.
- Each board row has one `Draft` action. It assigns the player to the team on the clock, advances the snake order, hides the player across all four tabs, and enables immediate Undo. Select a row and press `Enter` for keyboard entry; `Ctrl+Z` undoes the last pick. Letter keys remain available for typing in player search.
- The compact clock bar keeps the assignment controls close to the on-the-clock team and displays the most recently drafted player's name beside the user's next pick. Drafting, importing, undoing, undrafting, and clearing all refresh that value from the recorded pick history.
- The right-side Draft Room switches between compact league needs and the user's complete starter/FLEX/Superflex/bench roster. League Needs is generated from the active roster template, displays filled/configured counts for every enabled starter slot (for example `WR 2/2` and `FLX 1/2`), omits disabled slots, and turns completed slots green. Select any league team for its full allocation, then select a drafted player to move or undraft them. Drag the sidebar's left edge to resize it; drag an invisible table-header boundary to resize that column. Both choices persist locally.
- “Will He Make It Back?” is an empty manual watchlist until the user searches for and adds available players or double-clicks a QB/RB/WR/TE name on the board. Double-clicking the name again removes the player; tracked rows retain the same soft-white highlight used by normal row selection. Tracked cards are always displayed from lowest to highest rank for the active STRD, .5 PPR, FULL PPR, or SUPERFLEX tab, independent of the order they were added or their risk label. Return Risk blends the visible board rank with market ADP, distributes that player-specific probability across the opponent picks before the user's next turn, and scales each opportunity by the relevant team's configured roster need. When the same opponent picks more than once, its projected roster advances after the earlier selection so an open QB/WR/etc. does not suppress every later opportunity. Generic positional demand never creates an independent risk floor, while roster need adjusts rather than nearly erases the player-specific rank/ADP baseline; tier cliffs, recent runs, and team tendencies only scale that baseline. Open configured starters are prioritized over bench-only depth, early/mid-round backup QBs and TEs are discounted, and Superflex preserves strong QB2 demand. In leagues without a dedicated TE slot, TE demand comes only from FLEX/SFLX openings and uses a projection-led flex rank against RB/WR/TE alternatives; once a TE qualifies for an open FLEX on projected points, it receives the same slot-need weight as an RB or WR rather than a second no-TE penalty. `Take Now` is 70%+, `Pivot` is 60–69%, and `Can Wait` is below 60%. Cards keep the visible explanation to a concise opponent-pick sentence; the full need/ADP detail remains in the card tooltip. The info button beside the heading explains these labels in the app.
- A `Watch` action immediately before `Cuff` mirrors the player-name double-click toggle. The current-directory depth column covers RB/WR/TE roles, and TEAM abbreviations use distinct, high-contrast franchise colors.
- A thick white board rule marks the player occupying the user's next-pick slot. For example, from #46 to #51 the turn is five selections away but #51 is the sixth available row, so the line appears above row six. It moves after every selection and follows the current tab, sort, and search order.
- `Import Picks` accepts chronological `Team Name — Player` lines or player-only lines that follow the configured snake order. A preview blocks ambiguous, unmatched, unassigned legacy, or duplicate rows until each is corrected or explicitly skipped.
- `BOOM / BUST` reproduces the creator's signals: green for `Boom Factor = BOOM`, amber when Boom Connect contains `Desperation`, and red when early schedule is `hard` while `OLine Ovrl` is `OK` or `BAD`.
- `ADP EDGE` independently adds a bright Strong or lighter Watch rail. Edge pills remain sortable and open the current ranks, scored components, confidence, generation time, and source links. There is no separate warning layer.
- Standard, half-PPR, PPR, and Superflex use the current scoring-specific Sleeper ADP feed. Workbook rank is a fallback only for missing feed values. The visible rank cell is replaced with the current value; blue dotted text identifies an update and shows the workbook value on hover.
- `2025 Total Fantasy Pts` comes from completed nflverse scoring and `2026 Projected Fantasy Pts` blends available Sleeper/RotoWire and FantasyPros consensus stat lines. Current injury/news data overrides an older projection. Both columns are compact and scoring-format aware.
- `NFL Year` replaces the workbook's separate rookie/third-year columns. It comes from the current Sleeper player directory, with rookies displayed as 1.
- Projection positional rank versus current market positional rank is 50% of Edge. The other 50% scores opportunity, sample-shrunk efficiency, two-year trajectory/reversion, QB aDOT/rushing/sack floor, weekly concentration, target vacancy, incoming receiving production, teammate competition, projected QB change, destination-team TE usage, availability, and a current line proxy. The proxy combines 2025 team sack rate or RB rushing EPA/carry with current OL injury status. Workbook O-line, defense, and SOS cells remain visible reference fields but have zero Edge weight; current consensus projections absorb schedule and defensive changes.
- Current Sleeper team/status data overrides stale workbook display values without modifying the supplied workbook. Updated team cells are blue with a dotted underline and explain the workbook/current values on hover.
- Existing name-only My Team/Taken selections are preserved in a versioned legacy snapshot before migration. The import preview can surface them, but it never guesses an opponent owner.

All live-draft state uses versioned local browser records: `draftWizard:leagueProfiles:2026:v1`, `draftWizard:draftSessions:2026:v1`, and `draftWizard:activeLeague:2026:v1`. Canonical NFL IDs are used when available, with normalized name plus position as the fallback.

See [docs/adp-edge-model.md](docs/adp-edge-model.md) for formulas, sources, attribution, limitations, and join inspection.

## Verification

`npm run backtest:edge` runs rolling, no-lookahead 2022–2025 preseason folds against archived ECR and the following season's positional finish. It rejects position/direction models that fail their same-position/ECR-band baseline. The retained upside core is QB (22.9% hit rate versus 20.8% baseline across 14 selections); the retained downside component is TE (53.6% versus 45.6% across 32 selections). On the supplied 2025 names, updated Edge caught 2/12 breakouts and 1/8 regressions. That modest result is reported rather than tuned away. The command writes `data/reports/edge-review-2025.md`, which records every call and keeps absolute point misses separate from positional-rank outcomes. The same historical archive supplies position/market-band uncertainty for Return Risk. `npm test` also verifies feed precedence, season-ending injury removal, current-rank display, snake turns, overrides, roster allocation, corrections, import parsing, storage isolation, and ADP/opponent-need direction.
