# Macro Daily Snapshot (口径A) - Cloud Runner (No API Keys)

This repo runs a daily collector on GitHub Actions and outputs:
- out/daily_wide.csv
- out/daily_wide.xlsx

**Mode A (daily snapshot):** non-daily macro releases are forward-filled into a daily panel (step function).

## Sources (no API keys)
- Stooq (market daily)
- HKMA Open API (JSON)
- Hong Kong C&SD (JSON)
- US BLS Public API v1 (POST, no key)
- China NBS PressRelease RSS + HTML parsing

## Run locally
```bash
python3 macro_daily_collector_v3_keyless.py --registry registry_full_A_v3_keyless.json --out out --start 2026-04-01
```

## Run in GitHub Actions
- Enable Actions
- The workflow runs daily and on manual dispatch.
- Download artifacts from the workflow run page or grab files from `out/` in the repo.

## Notes
Some items in your original long list are not reliably machine-readable from stable public sources (e.g., China fund daily sales totals, certain fund holdings/cash ratios). Keep those as manual columns or point them to a custom JSON endpoint.
