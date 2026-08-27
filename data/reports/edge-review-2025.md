# 2025 Edge review

Generated 2026-08-20T18:00:26.817Z. The point-miss cohorts come from `data/inputs/edge-review-2025.json`; market ranks are the last FantasyPros redraft ECR snapshot on or before August 31, 2025.

## Result

- Legacy Edge caught 2/12 listed breakouts and 1/8 listed regressions.
- Updated Edge catches 2/12 listed breakouts and 1/8 listed regressions. Downside only counts when its position model beat the rolling no-lookahead baseline.
- Retained historical upside: QB (22.9% vs 20.8%, 14 selections). Retained downside Edge: TE (53.6% vs 45.6%, 32 selections).

## Player-level calls

| Cohort | Pos | Player | Projected → Actual | Legacy Edge | Updated Edge | Market-pos result |
| --- | --- | --- | ---: | --- | --- | --- |
| breakout | QB | Drake Maye | 285 → 352 (+67) | neutral | neutral | beat (16→2) |
| breakout | QB | Matthew Stafford | 274 → 350.4 (+76.4) | neutral | neutral | beat (22→3) |
| breakout | QB | Trevor Lawrence | 279 → 338.2 (+59.2) | neutral | neutral | beat (18→4) |
| breakout | RB | Kenneth Gainwell | 66 → 221.3 (+155.3) | neutral | downside | beat (77→16) |
| breakout | RB | Rico Dowdle | 84 → 216.3 (+132.3) | upside | upside | beat (48→18) |
| breakout | RB | Travis Etienne Jr. | 132 → 253.9 (+121.9) | neutral | neutral | beat (31→10) |
| breakout | TE | Harold Fannin Jr. | 54 → 186.4 (+132.4) | downside | neutral | beat (34→6) |
| breakout | TE | Kyle Pitts Sr. | 142 → 210.8 (+68.8) | neutral | neutral | beat (18→2) |
| breakout | TE | Oronde Gadsden II | 57 → 131.4 (+74.4) | downside | neutral | beat (39→14) |
| breakout | WR | Alec Pierce | 71 → 183.3 (+112.3) | neutral | neutral | beat (73→28) |
| breakout | WR | Jaxon Smith-Njigba | 241 → 359.9 (+118.9) | upside | upside | beat (12→2) |
| breakout | WR | Parker Washington | 44 → 172.7 (+128.7) | neutral | neutral | beat (103→27) |
| regression | QB | Cam Ward | 257 → 187 (-70) | downside | neutral | beat (27→21) |
| regression | QB | Tua Tagovailoa | 276 → 161 (-115) | neutral | neutral | missed (21→24) |
| regression | RB | Chuba Hubbard | 258 → 125 (-133) | upside | upside | missed (18→38) |
| regression | RB | Saquon Barkley | 326 → 232 (-94) | upside | upside | missed (3→14) |
| regression | TE | Evan Engram | 169 → 103 (-66) | neutral | neutral | missed (7→26) |
| regression | TE | Jonnu Smith | 135 → 85 (-50) | upside | downside | missed (17→30) |
| regression | WR | Jerry Jeudy | 232 → 121 (-111) | upside | neutral | missed (33→50) |
| regression | WR | Justin Jefferson | 316 → 202 (-114) | upside | upside | missed (3→21) |

## Interpretation

Cam Ward illustrates the difference between point projections and draft value: he missed the supplied point projection by 70, but finished ahead of his preseason positional market rank. Edge is rank-relative, so the report preserves both results.

The production formula now scores two-year reversion, QB aDOT/rushing/sack floor, weekly concentration, target vacancy, incoming receiving production, teammate competition, projected QB change, and destination-team TE usage. Sparse college and coaching history remains outside production scoring until broad historical coverage exists.
