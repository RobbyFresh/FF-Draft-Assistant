# ADP Edge 2026 model card

## Purpose

ADP Edge is a same-tier comparison aid for QB, RB, WR, and TE. It looks for evidence that a player's expected 2026 role and scoring rank are better than the current scoring-specific market price. K and DST are not scored. The 0–100 number is a comparison score, not a guarantee or a literal probability.

## Inputs and attribution

- The creator's 2026 workbook supplies the four display tabs and fallback age/depth/rank fields. Its O-line, defense, and schedule fields remain visible but have zero production Edge weight. The creator's method is described in the [2026 Reddit post](https://www.reddit.com/r/fantasyfootballadvice/comments/1vndd5x/created_a_free_cheat_sheet_with_everything/).
- [nflverse player statistics](https://nflreadr.nflverse.com/reference/load_player_stats.html) provide 2025 fantasy points, attempts, carries, receptions, targets, target share, air-yard share, WOPR, EPA, yards, first downs, two-year opportunity/production trends, passing depth, sacks, and weekly production concentration.
- [NFL Next Gen Stats through nflverse](https://nflreadr.nflverse.com/articles/dictionary_nextgen_stats.html) provide passer rating/CPOE, receiver separation/YAC over expectation, and RB rushing yards over expectation where samples are available.
- nflverse [rosters](https://github.com/nflverse/nflverse-data/releases/tag/rosters), [snap counts](https://github.com/nflverse/nflverse-data/releases/tag/snap_counts), and [draft picks](https://github.com/nflverse/nflverse-data/releases/tag/draft_picks) provide IDs, current team/status, offensive participation, and rookie capital. nflverse's relevant timing and limitations are documented in its [data schedule](https://nflreadr.nflverse.com/articles/nflverse_data_schedule.html).
- Sleeper's public [player directory](https://api.sleeper.app/v1/players/nfl) supplies current team, injury/news state, and years of NFL experience. Its [2026 projection feed](https://api.sleeper.com/projections/nfl/2026?season_type=regular) supplies scoring-specific ADP and stat projections and identifies RotoWire as the underlying provider in current records.
- [FantasyPros 2026 consensus projections](https://www.fantasypros.com/nfl/projections/) supply a second projection source for the publicly visible leading players. Standard, half-PPR, and PPR totals are averaged across available sources; missing providers are not treated as zero.
- The [DynastyProcess FantasyPros archive](https://github.com/DynastyProcess/data) supplies time-safe historical preseason ECR for validation.

## Fantasy-point columns

The `2025 Total Fantasy Pts` column uses completed 2025 nflverse totals. Standard is the nflverse standard total; half-PPR adds 0.5 per reception; PPR adds one point per reception. Superflex uses the PPR total because the workbook's Superflex tab is Full-PPR.

The `2026 Projected Fantasy Pts` column is a consensus of the available Sleeper/RotoWire and FantasyPros stat-line totals in the active scoring format. Current injury/news state has higher precedence than an older projection. A confirmed season-ending IR injury with ACL/Achilles surgery removes the player from active market/model ranks and sets projected points to zero. Joins use player ID when available and normalized name plus position otherwise.

## Position-specific formula

All components are winsorized at the 5th/95th percentiles and standardized within position. Efficiency is sample-size shrunk so a small hot streak cannot dominate.

The most important input is not raw projected points. It is the player's consensus projected positional rank minus the player's market positional rank. That prevents the model from simply rewarding stars who already have expensive ADPs.

| Pos. | Opportunity | Efficiency | Current line | Trajectory | Reversion | QB floor | Team context | Concentration | Availability | Projection-vs-market |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| QB | 14% | 5% | 6% | 5% | 5% | 8% | 4% | 0% | 3% | 50% |
| RB | 11% | 3% | 6% | 6% | 8% | 0% | 11% | 0% | 5% | 50% |
| WR | 13% | 5% | 5% | 6% | 6% | 0% | 9% | 3% | 3% | 50% |
| TE | 12% | 4% | 5% | 7% | 8% | 0% | 10% | 1% | 3% | 50% |

The inputs inside those components differ by position:

- QB: pass attempts, rushing value, EPA per attempt, CPOE/passer rating, aDOT, sack rate, two-year direction, incoming receiving production, current line context, and projected pass/rush totals.
- RB: carries, targets, touches, target share, RYOE, two-year workload reversion, teammate competition, current run-line context, depth role, draft capital, and projected carries/receptions/touchdowns.
- WR: targets per game, target/air-yard share, WOPR, yards and EPA per target, separation/YAC over expectation, target vacancy, projected QB change, single-game concentration, current pass-line context, and projected receptions/yards/touchdowns.
- TE: the same receiving structure at TE-specific baselines, plus destination-team TE usage, QB change, TE-room competition, current pass-line context, and projected role.
- Rookies: draft capital, experience year, projected workload/scoring, depth and team context. Missing NFL history lowers confidence but is not scored as poor production.

Projected role growth compares like with like: projected touches/receptions against prior touches/receptions. Injury status is not forecast from prior games missed; it is applied only when current injury/news data is newer than the projection or the current projection feed has removed the player after a confirmed serious IR injury.

Players are re-ranked within position around their market rank. Expected delta is market positional rank minus model positional rank. The displayed score is a bounded logistic transformation of expected rank gain and component strength. Strong requires score ≥70, expected gain ≥3, completeness ≥65%, and top-240 eligibility. Watch requires score ≥58, gain ≥1, completeness ≥50%, and top-240 eligibility.

## Historical validation

`npm run backtest:edge` runs rolling no-lookahead 2022–2025 folds. Each test fold freezes DynastyProcess redraft ECR at August 31 or earlier, selects candidate position weights only from earlier seasons, and evaluates the following completed-season positional finish. Volume, dominance, efficiency, production, and trajectory are compared with an ECR-only baseline made from the same position and market-rank band.

This validation gates upside and downside separately. The retained historical upside core is QB: 22.9% of its 14 selections beat market position versus 20.8% for matched ECR-band peers (+2.1 percentage points). RB, WR, and TE upside cores were rejected. The retained downside component is TE: 53.6% of selected TEs underperformed market position versus 45.6% for matched peers (+7.9 points across 32 retained-fold selections). The production score still uses the full, capped 50% statistical/context half at every position, but the failed standalone gates are disclosed and are not presented as proven alpha.

The supplied 2025 review cohort is intentionally not used to choose production weights. Updated Edge caught 2 of 12 listed breakouts and 1 of 8 listed point regressions; Jonnu Smith is the regression caught by the independently retained TE downside model. Cam Ward illustrates why the report stores both targets: he missed the supplied point projection while still beating his preseason positional market rank. See `data/reports/edge-review-2025.md` and `data/inputs/edge-review-2025.json`.

The backtest cannot time-safely validate current injury snapshots, OL availability, or 2026 projection disagreement because complete historical preseason snapshots are not publicly available. Current line context is therefore a reproducible proxy: completed-2025 team QB sack rate for pass protection or RB rushing EPA/carry for run blocking, adjusted by current OL injury status. Workbook line/defense/SOS labels have zero score weight. See `data/reports/backtest-2026.json` for every fold and player-level call.

## Return Risk and live-draft intelligence

Return Risk is deliberately separate from ADP Edge. Edge asks whether a player is undervalued; Return Risk estimates whether the player is likely to be selected before the user's next snake pick.

The offline build derives position- and market-band uncertainty from the historical FantasyPros archive's consensus rank standard deviation and expert best/worst ranges. During a draft, the browser blends that market signal with the active tab's visible board rank, then compounds position pressure separately for every opponent selection before the user's next turn. This preserves repeated opportunities when the same needy team selects twice around a snake turn. Open native/FLEX/Superflex slots, round-specific positional behavior, remaining players in the same workbook tier, recent position runs, and lightly weighted team tendencies also contribute. The pane starts empty and calculates these labels only for available players the user manually adds through its search box; it does not auto-fill from ADP Edge. `Take Now` means 70%+ estimated Return Risk, `Pivot` means 60–69%, and `Can Wait` means below 60%. The visible card copy is limited to the pick-distance sentence; detailed need and rank context remains in its tooltip. The UI calls the result an estimate. It is not trained on a complete archive of individual home-league draft rooms and should not be read as a literal calibrated probability.

## Reproducibility and inspection

Run `npm run build:edge` to use cached inputs, or run the documented refresh command to download current public inputs. The browser makes no third-party calls during a live draft. Every scored record includes join method, confidence, component values, displayed metrics, explanations, generation time, and source references. `data/reports/unmatched-2026.json` records unmatched completed-season stat rows; top-240 current identity coverage is validated separately.

The supplied workbook is never rewritten.

## Limitations

Projection providers can be wrong together, public coverage varies, and preseason roles change. Receiving yards per offensive snap is included as a reproducible route-efficiency proxy; it is deliberately not mislabeled as true YPRR because the free route-run data is incomplete. College sack rates, complete prospect production, coaching-history target shares, and transaction timing do not yet have broad time-safe coverage. The injury override is deliberately narrow; ambiguous injuries and suspensions still require draft-day judgment.
