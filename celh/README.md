# CELH Financial Analysis Scripts

## Scripts

### Income Statement
- `build_celsius_simple_v4_final.py` — Complete engine, self-testing
- Run: `python build_celsius_simple_v4_final.py`
- Output: `CELH_Q1_2026_Income_Statement_v4.xlsx`
- Self-test: prints PASSED or FAILED with details

### Projection
- `build_celh_projection_final.py` — Complete engine, self-testing  
- Run: `python build_celh_projection_final.py`
- Output: `CELH_5Year_Projection_v3.xlsx`
- Self-test: prints PASSED or FAILED with details

## Version Control
- Every update creates a new version file (v5, v6, etc.)
- Never overwrite existing versions
- Self-test must pass before committing

## Key Q1 2026 Outputs (verify after each run)
### Income Statement
- Revenue: $782.6M
- Gross Margin: 48.3%
- Adj. EBITDA: $195.5M (25.0% margin)
- Adj. EPS: $0.41

### Projection 2030
- Bull: SPL=$217 SPH=$271
- Base: SPL=$110 SPH=$138
- Bear: SPL=$39 SPH=$51
