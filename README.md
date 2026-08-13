# Amazon Catalog Command Center

Tableau portfolio project: catalog P&L attainment + Prime Day war room, built from **synthesized** performance data.

**Question:** Did the SEO catalog hit active revenue/GP targets, and did Prime Day buy profitable volume?

## Disclaimer

Data is synthesized for portfolio demonstration. Brand names in the clean files are **fictional** (VitalCore, PawVista, Medora) — not affiliated with any real consumer brand.

## Headline KPIs (validated)

€5.56M income · €1.09M GP · ~90.5% of active revenue target · ~19.6% margin

## Clean data (`data/clean/` — join on `asin`)

| File | Role |
|---|---|
| `monthly_asin_pnl.csv` | Main P&L fact (18 months) |
| `asin_dim.csv` | Product, brand, owners |
| `prime_day_plan_vs_actual.csv` | Prime Day plan vs Day-1 |
| `hourly_prime.csv` | Intra-day pace |
| `campaign_budget.csv` | Campaign spend vs budget |
| `cpu_mix.csv` | PPC vs organic / CPU |
| `data_dictionary.csv` | Field definitions |
| `validation_summary.csv` | Totals check vs source headers |

## Workbook

- **Local:** Tableau Desktop workbook *Amazon Catalog Command Center* (Executive + Prime Day dashboards + story)
- **Public:** [Amazon Catalog Command Center — Story](https://public.tableau.com/app/profile/aneesh.kumar2834/viz/AmazonCatalogCommandCenter/CatalogPerformanceStory)

Design follows the same dark visual standard as Projects 2–3.

## Reproduce clean CSVs

```bash
python3 scripts/extract_tableau_data.py
```

(Requires local source PDF/Excel outside this repo; output is anonymized on write.)

## Plan of attack

Open `POA-interactive.html` for the build checklist.
