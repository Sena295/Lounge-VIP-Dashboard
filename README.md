# VIP Lounges Performance

Interactive HTML dashboard for tracking airport VIP lounge partner performance, paired with the Python pipeline that builds the input workbook, collects exchange rates, and generates the report.

---

### Overview

This project automates the reporting workflow for VIP lounge partnership performance across multiple airports:

- **Input generation**: `criar_planilha_input.py` builds a structured input workbook from the raw partner access base, one sheet per lounge, preserving manually entered fields (paying passengers, currency overrides) across regenerations.
- **Exchange rate collection**: `fx_collector.py` fetches daily EUR/USD rates from investing.com (with the European Central Bank as fallback), audits the data for invalid or duplicate dates, and outputs both the legacy and corrected monthly averages.
- **Report generation**: `gerar_relatorio.py` reads the input workbook and renders the standalone HTML dashboard.

---

### Dashboard

Open `VIP_Lounges_Performance (In).html` directly in any browser, no server or build step required. All data and rendering logic are self-contained in the file.

**Highlights**

- Revenue and passenger volume by lounge and partner airline
- Paying passenger percentage
- Published vs. corrected EUR/USD rate comparison, with data quality alerts
- Year-over-year and month-over-month breakdowns
- Responsive layout, print-friendly

---

### Automation Scripts
| `criar_planilha_input.py` | Rebuilds the VIP lounge input workbook from the latest partner access base, keeping manual entries intact |
| `fx_collector.py` | Collects and audits daily EUR/USD rates, flagging invalid dates, duplicates, and text-formatted values that Excel's `AVERAGE` silently ignores |
| `gerar_relatorio.py` | Generates the final HTML dashboard from the input workbook |

Built with `openpyxl` for spreadsheet generation and parsing, and plain `urllib` for the exchange rate collection, no external HTTPS dependencies.

Paths for the source and output folders are read from environment variables (`VIP_LOUNGE_SOURCE_PATH`, `VIP_LOUNGE_OUTPUT_PATH`) rather than hardcoded, so the scripts can run against any equivalent folder structure.

---

### Tech Stack

`Python` · `openpyxl` · `urllib` · `HTML` · `CSS` · `JavaScript`

--

### Note on Data

The dataset shown in the dashboard is illustrative and does not represent real financial figures or actual partner agreements
