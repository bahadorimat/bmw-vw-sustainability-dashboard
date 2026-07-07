# German Automotive Sustainability Dashboard
## BMW Group vs. Volkswagen Group — CSRD & KPI Analysis (FY 2025)

A comparative sustainability analysis of BMW Group and Volkswagen Group using real audited data extracted directly from their 2025 CSRD/ESRS sustainability reports. The project combines GHG emissions analysis, sustainability KPI benchmarking, and ESRS topic coverage into a single Python dashboard.

---

## What is CSRD?

The **Corporate Sustainability Reporting Directive (CSRD)** is an EU law that requires large companies to disclose sustainability data following the **European Sustainability Reporting Standards (ESRS)**. Both BMW and Volkswagen are subject to CSRD from 2024 onward. Their reports are independently audited — making the data reliable for analysis.

---

## Data Sources

| Field | BMW Group | Volkswagen Group |
|-------|-----------|-----------------|
| Report | BMW Group Report 2025 | VW ESRS Sustainability Report 2025 |
| Page reference | p. 128–129 (GHG), KPI sections | p. 273 (Scope 3), KPI tables |
| GHG standard | GHG Protocol Scope 3 Standard | GHG Protocol Scope 3 Standard |
| CSRD standard | ESRS Set 1 (2023) | ESRS Set 1 (2023) |
| Auditor | PricewaterhouseCoopers GmbH | Ernst & Young GmbH |
| Audit level | Reasonable assurance | Reasonable assurance |

---

## Key Findings

| Metric | BMW 2025 | VW 2025 | Ratio |
|--------|---------|---------|-------|
| Scope 1+2 (Mt CO₂e) | 0.811 | 2.80 | VW 3.5x higher |
| Scope 3 Total (Mt CO₂e) | 127.54 | 883.74 | VW 6.9x higher |
| Total Energy (million MWh) | 6.18 | 19.2 | VW 3.1x higher |
| Renewable Energy Share | **49.4%** | 40.6% | BMW leads |
| Water Withdrawal (million m³) | 5.70 | 19.9 | VW 3.5x higher |
| Total Waste (thousand tonnes) | 851.8 | 2,547 | VW 3.0x higher |
| Employees | 154,540 | 602,659 | VW 3.9x larger |
| Scope 3 per employee (t CO₂e) | **825** | 1,466 | BMW 1.8x more efficient |

> **Key insight:** For both companies, 70–85% of all Scope 3 emissions come from the use phase of sold vehicles (Cat 11) — confirming that EV transition is the single most important climate lever for the automotive sector.

---

## What the Script Produces

Running `sustainability_dashboard.py` generates 4 figures and one Excel workbook saved to `output/`:

- **Figure 1** — GHG emissions comparison: Scope 1+2, Scope 3 and total for both companies (2024 vs 2025)
- **Figure 2** — Scope 3 category breakdown: donut charts showing which categories dominate for each company
- **Figure 3** — KPI scorecard: side-by-side comparison of energy, renewables, water, waste and per-employee emissions
- **Figure 4** — CSRD/ESRS topic coverage: which ESRS topics each company reports on and considers material
- **Excel workbook** — full data tables with 3 sheets: GHG emissions, KPI scorecard, data sources

---

## CSRD / ESRS Coverage

| ESRS Topic | BMW Reports | BMW Material | VW Reports | VW Material |
|------------|-------------|-------------|------------|-------------|
| E1 — Climate Change | ✅ | ✅ | ✅ | ✅ |
| E2 — Pollution | ✅ | ✅ | ✅ | ✅ |
| E3 — Water & Marine | ✅ | ❌ | ✅ | ✅ |
| E4 — Biodiversity | ✅ | ❌ | ✅ | ✅ |
| E5 — Resource Use | ✅ | ✅ | ✅ | ✅ |
| S1 — Own Workforce | ✅ | ✅ | ✅ | ✅ |
| S2 — Value Chain Workers | ✅ | ✅ | ✅ | ✅ |
| S3 — Communities | ❌ | ❌ | ✅ | ✅ |
| S4 — Consumers | ❌ | ❌ | ✅ | ❌ |
| G1 — Business Conduct | ✅ | ✅ | ✅ | ✅ |

---

## How to Run

```bash
pip install pandas matplotlib openpyxl numpy

# Open in Spyder and press F5
# or from terminal:
python sustainability_dashboard.py
```

---

## Project Structure

```
bmw-vw-sustainability-dashboard/
├── sustainability_dashboard.py
├── requirements.txt
├── README.md
└── .gitignore
```

---

## Methodology

| Item | Detail |
|------|--------|
| GHG standard | GHG Protocol Corporate Value Chain (Scope 3) Standard (2011) |
| CSRD standard | ESRS Set 1 — European Sustainability Reporting Standards (2023) |
| Scope 2 method | Market-based |
| Climate science | IPCC AR6 Global Warming Potential (GWP100) |
| Boundary | Both companies: global operations |

---

**Matin Bahadori**
M.Sc. Umwelttechnik — BTU Cottbus-Senftenberg
