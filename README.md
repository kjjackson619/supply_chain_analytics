Hosted at https://docs.google.com/spreadsheets/d/1UWeiE45m07DOPI3U6edua7siG5AswZci-zoXDOoc8rc/edit?usp=sharing

# 📦 Supply Chain Analytics Workbook

An advanced Excel / Google Sheets portfolio project demonstrating analyst-level spreadsheet skills across lookup functions, aggregation, scenario modeling, data validation, and conditional formatting — built on a realistic ocean import operations dataset.

---

## Workbook Structure

### 📋 README (Tab 1)
Index of all sheets, skill inventory, and the industry-standard color-coding convention used throughout:
- **Blue text** — hardcoded inputs the user can change
- **Black text** — formula-calculated values
- **Green text** — cross-sheet links pulling from another tab
- **Amber background** — key assumption cells

---

### 📦 Raw Orders (Tab 2)
80 rows of transaction-level ocean import data. The source-of-truth layer for all downstream analysis.

**Skills demonstrated:**
- Structured Excel Table with named reference (`OrdersTable`)
- `IFERROR` wrapping on division formulas to prevent `#DIV/0!`
- Formula columns: Total Landed Cost, Delay Flag, On-Time %, Cost per TEU
- **Conditional formatting:** traffic-light coloring on Status column (Delivered/Delayed/Cancelled), color scale on Freight Cost, data bars on Total Landed Cost

**Fields include:** Order ID, Origin/Destination Port, Carrier, Category, Incoterms, TEU count, Unit Value, Freight Cost, Insurance, Duties, Total Landed Cost, Ship Date, ETA, Actual Arrival, Transit Days, Status, Delay Flag

---

### 📊 Dashboard (Tab 3)
Executive summary pulling live from Raw Orders via cross-sheet formulas. All values update automatically when source data changes.

**Skills demonstrated:**
- Cross-sheet formula references (`='📦 Raw Orders'!column`)
- `COUNTIF`, `COUNTIFS` — shipment counts by carrier and status
- `SUMIF`, `SUMIFS` — revenue aggregation by carrier and category
- `AVERAGEIF`, `AVERAGEIFS` — average freight rate and transit days by carrier
- Percentage-of-total calculations with absolute references (`$D$20:$D$25`)
- Color scale conditional formatting on On-Time % (red → yellow → green)
- Data bars on Category % share column

**Panels:**
| Section | Formula Types Used |
|---|---|
| KPI Cards (6 metrics) | `COUNTA`, `SUM`, `AVERAGEIF`, `COUNTIF` + cross-sheet refs |
| Carrier Performance (6 carriers) | `COUNTIF`, `AVERAGEIF`, `COUNTIFS`, `SUMIF`, `AVERAGEIFS` |
| Revenue by Category (6 categories) | `COUNTIF`, `SUMIF`, `AVERAGEIF`, `%` of total |

---

### 🔍 VLOOKUP & INDEX/MATCH (Tab 4)
Side-by-side demonstration of the three primary lookup methods solving identical problems, plus a nested IF / IFS section.

**Skills demonstrated:**
- `VLOOKUP` with exact match (`FALSE`) and `IFERROR` wrapping
- `INDEX` / `MATCH` combination for flexible two-way lookup
- `XLOOKUP` with default fallback value (Excel 365 / Google Sheets compatible)
- Structured Table references (`ProductTable[Column Name]`) in all formulas
- Nested `IF` chains for multi-tier classification
- `IF` + comparison logic for composite alert flags (🔴 HIGH / 🟡 MEDIUM / 🟢 LOW)

---

### 📈 Pivot Analysis (Tab 5)
Manual pivot-style cross-tabulation and monthly trend analysis — built with formulas, not Excel Pivot Tables, to show the underlying logic.

**Skills demonstrated:**
- `SUMIFS` matrix: Origin Port × Product Category cross-tab (5×6 grid)
- Row and column totals using `SUM` with mixed references
- `SUMPRODUCT` for date-based aggregation (monthly filtering without helper columns)
- `AVERAGEIFS`-equivalent via `SUMPRODUCT` for conditional averages
- Color scale heat map on the cross-tab matrix (white → blue gradient by value)
- Red → green color scale on % Delayed column

---

### 🌊 Freight Cost Model (Tab 6)
A dynamic ocean import landed cost calculator comparing four scenarios with all assumptions isolated in editable input cells.

**Skills demonstrated:**
- Assumption-driven model architecture — all inputs in one block, all formulas reference those cells
- Named range documentation (column labels showing intended named range names)
- Multi-scenario comparison: Base Case / Optimistic / Stressed / Express
- Freight-to-value ratio formula with automatic vs-target comparison
- Savings vs. Base Case calculation
- `IF` statements generating automatic threshold alerts ("⚠ Over Target" / "✓ Within Target")
- Color scale on freight-to-value ratios across all scenarios

**Scenarios modeled:**

| Scenario | Description |
|---|---|
| Base Case | Standard market rates at current assumptions |
| Optimistic | 12% freight reduction, favorable surcharges |
| Stressed | 20% freight increase, elevated surcharges |
| Express | Premium rate multiplier (1.65×) for time-sensitive cargo |

---

### ⚠️ Risk Flags (Tab 7)
Automated risk scoring with formula-driven alert generation and enforced data validation.

**Skills demonstrated:**
- **Data Validation:** Carrier dropdown list (restricts input to valid values with custom error message)
- **Data Validation:** Numeric range constraint on Delay Days (-10 to 60, whole numbers only)
- Composite risk score formula: `=ROUND(MIN(100, MAX(0, delay*2.5 + variance*100)), 0)`
- Multi-tier `IF` chain for risk classification (CRITICAL / ELEVATED / MODERATE / LOW)
- `AVERAGE`, `MAX`, `MIN`, `STDEV` on risk scores
- `PERCENTILE` (75th) for threshold analysis
- `AVERAGEIF` for conditional average (delayed shipments only)
- Color scale conditional formatting on risk scores (0–100 range)

---

## Domain Context

This workbook is built around **ocean import operations** — a data-rich domain where analysts regularly track:
- **Freight cost variance** — budget vs. actual per TEU, per lane, per carrier
- **Transit time performance** — on-time rate, delay days, carrier benchmarking
- **Landed cost modeling** — freight + duties + insurance + drayage for true cost visibility
- **Origin diversification** — exposure by country/port for supply chain risk

---

## Skills Evidenced

`VLOOKUP` `INDEX/MATCH` `XLOOKUP` `SUMIFS` `AVERAGEIFS` `COUNTIFS` `SUMPRODUCT` `Nested IF` `IFERROR` `PERCENTILE` `STDEV` `Data Validation` `Conditional Formatting` `Excel Tables` `Scenario Modeling` `Cross-Sheet References` `Named Ranges` `Supply Chain Analytics` `Freight Cost Analysis` `Ocean Import Operations`
