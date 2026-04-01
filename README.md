# Voyage-Manager

Vessel voyage expense tracker and reporting tool.

## Summary

See **[SUMMARY.md](SUMMARY.md)** for a full breakdown of all voyage expenses, including per-voyage details, category analysis, top vendors, and urgency distribution.

### Quick Stats

| Metric | Value |
|---|---|
| Total Voyages | 6 |
| Period | July 2026 – January 2027 |
| Total Sailing Days | 139 |
| Total OPEX (IDR) | Rp 3,136,244,937 |
| Total OPEX (USD) | $675,351.00 |

## Regenerating the Summary

```bash
pip install openpyxl
python generate_summary.py
```

This reads `Salinan dari Expense kapal.xlsx` and writes `SUMMARY.md`.