# AMD Financial Analysis Scripts

## Scripts

### Projection (5-Year)
- `build_amd_projection_engine.py` — Complete engine, self-testing
- Run: `python build_amd_projection_engine.py`
- Output: `AMD_5Year_Projection_v6.xlsx`
- Self-test verifies: file size, tabs, Bull SPL ~$3,073, Base SPL ~$1,337

### Income Statement
- `build_amd_simple.js` — Complete engine (on Drive, not GitHub)
- Drive ID: 15RmUwcgcsNc_sc3YoD4S0P7SHkF10LNf

## Key Q1 2026 Outputs
| Scenario | 2030 SPL | 2030 SPH | 5-Yr CAGR |
|---|---|---|---|
| Bull (35%) | $3,073 | $3,713 | +46.5% |
| Base (45%) | $1,337 | $1,504 | +24.1% |
| Bear (20%) | $181 | $238 | -16.8% |
| Expected | $1,713 | $2,024 | +30.4% |

## Version Control
- Every update = new version file (v7, v8, etc.)
- Never overwrite existing versions
- Self-test must pass before committing
