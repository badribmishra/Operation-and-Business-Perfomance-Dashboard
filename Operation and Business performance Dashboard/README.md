# 🛡️ Operations & Business Performance Dashboard

> **Executive-grade operational performance tracking, service level agreement (SLA) monitoring, 12-month cohort trend analytics, and data-driven decision frameworks.**  
> Built as both an **enterprise Excel workbook (.xlsx)** and an **interactive real-time web application**.

---

## 📌 Project Overview & Impact Summary

This project delivers an end-to-end operational intelligence system designed to **organize complex telemetry data, track mission-critical business and service desk KPIs, identify longitudinal trends, and empower data-driven operational decisions**.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                        ENTERPRISE OPERATIONS SCORECARD (AUDITED)                       │
├──────────────────────────┬──────────────────────────┬──────────────────────────────────┤
│ Monthly Recurring Rev    │ Annual Run Rate (ARR)    │ Portfolio Churn Rate             │
│ $456,116.60              │ $5,473,399.20            │ 26.54% (1,869 accounts)          │
├──────────────────────────┼──────────────────────────┼──────────────────────────────────┤
│ Total Customer Accounts  │ Overall SLA Compliance   │ Portfolio Average CSAT           │
│ 7,043 accounts           │ 77.74% (Target: ≥ 85.0%) │ 3.51 / 5.0 (Target: ≥ 4.20)      │
├──────────────────────────┼──────────────────────────┼──────────────────────────────────┤
│ Monthly Revenue at Risk  │ Total Cost to Serve      │ Net Operational Margin           │
│ $139,130.85 / month      │ $484,174.50              │ -6.15% (Support overhead drag)   │
└──────────────────────────┴──────────────────────────┴──────────────────────────────────┘
```

---

## 🗂️ Multi-Sheet Excel Dashboard Architecture

The core asset is **`Operations_and_Business_Performance_Dashboard.xlsx`**, crafted using openpyxl with automated styling, native dynamic Excel formulas, frozen header panes, formatted numbers, and embedded charts:

| Sheet Tab | Purpose & Contents | Key Formulas / Logic |
| :--- | :--- | :--- |
| **`[1] Executive KPI Summary`** | Executive scorecards, Departmental benchmarks, and Operational Risk Tier revenue distribution with embedded column and donut charts. | `=SUM('Master Operational Data'!S2:S7044)`<br>`=A6*12`<br>`=COUNTIF('Master Operational Data'!U2:U7044, "No")`<br>`=COUNTIFS(...)` |
| **`[2] Departmental Operations & SLAs`** | Granular service desk operational diagnostics, ticket volume, average handling duration, SLA breach counts, and cost-to-serve economics. | `=SUMIF('Master Operational Data'!$AA$2:$AA$7044, A6, 'Master Operational Data'!$Z$2:$Z$7044)`<br>`=AVERAGEIF(...)`<br>`=B6*I6` |
| **`[3] Trends & Cohort Analytics`** | 12-month historical & forecast trajectory tracking active customer expansion, churn erosion, MRR inflow vs lost, and SLA evolution. | `=F6-G6`<br>Embedded dual-axis Line Chart & Clustered Column Chart |
| **`[4] Data-Driven Decisions`** | Strategic decision & intervention matrix detailing 5 prioritized initiatives with severity badges, root cause diagnostics, and quantified ROI. | Multi-column priority action matrix with color-coded severity fills (CRITICAL, HIGH, MEDIUM) |
| **`[5] Master Operational Data`** | Cleaned, normalized master operational table with 7,043 transaction records across 34 standardized columns with Excel table formatting. | Full column styling (`$#,##0.00`, `0.0`, `#,##0`), freeze panes on row 2, and AutoFilter enabled |
| **`[6] Data Dictionary & KPI Guide`** | Exhaustive reference guide defining metric names, business categories, definitions, calculation formulas, targets, and departmental owners. | 10 standardized enterprise KPI definitions & governance guidelines |

---

## 🏢 Departmental Performance Benchmarks

| Department | Active Accounts | Total Tickets | Avg Resolution (hrs) | SLA Compliance % | FCR % | Avg CSAT | Total Cost to Serve |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **Tier 1 Frontline Support** | 2,632 | 4,210 | 8.5h | 91.2% | 82.0% | 4.15 | $77,885 |
| **Tier 2 Technical Escalations** | 1,688 | 3,890 | 31.0h *(Breached)* | 48.4% | 48.0% | 2.82 | $163,380 |
| **Billing & Payments Ops** | 1,818 | 2,750 | 14.0h | 86.5% | 74.0% | 3.65 | $60,500 |
| **Network & Field Operations** | 905 | 1,980 | 26.5h *(Breached)* | 62.1% | 58.0% | 3.42 | $72,270 |

---

## 🎯 Prioritized Data-Driven Decisions Matrix

1. **ACT-01: Automated Technical Diagnostics & AI Self-Healing Gateway**
   - *Problem:* Tier 2 Technical Escalations MTTR is 31.0 hrs (+29% above 24h SLA target), driving 48.6% churn among affected subscribers.
   - *Action:* Deploy automated ONT optical diagnostic script and auto-reboot provisioning for fiber broadband lines.
   - *ROI:* Cuts Tier 2 MTTR from 31h to 14h (-54%). Preserves **$264,000 ARR** by averting 340 customer cancellations annually.
2. **ACT-02: Frictionless Billing Portal & Payment Retry Automation**
   - *Problem:* Electronic Check payment method accounts for 1,818 inquiries and a 45% churn rate due to unconfirmed payment status.
   - *Action:* Offer a one-time $10 credit to migrate users from Electronic Check to Auto-Pay Credit Card, coupled with real-time SMS status confirmations.
   - *ROI:* Decreases billing service tickets by 62%; preserves **$88,000/mo** in retained recurring revenue.
3. **ACT-03: Frontline First Contact Resolution (FCR) Enhancement**
   - *Problem:* Repeat frontline transfers to Tier 2 inflate overall cost-to-serve to $42/ticket.
   - *Action:* Equip frontline agents with 1-click line reset tools and interactive guided decision trees.
   - *ROI:* Increases frontline FCR from 82% to 91%, eliminating 410 escalations/month and saving **$115,000 annually** in tier-2 overhead.
4. **ACT-04: High-Risk Month-to-Month Contract Lock-In Incentive**
   - *Problem:* Month-to-month contracts represent 88.5% of churned accounts and generate 3.8x higher service calls.
   - *Action:* Personalized 15% discount offer for annual commitment bundled with streaming protection.
   - *ROI:* Converts 22% of high-risk subscribers, reducing portfolio churn by 3.4 points (**$310,000 ARR preserved**).
5. **ACT-05: Predictive Telemetry Monitoring for Field Dispatches**
   - *Problem:* Field dispatches cost $36.50/dispatch with 42% repeat tickets when SNR degrades silently.
   - *Action:* Implement automated SNR degradation triggers to dispatch preventative line tuning.
   - *ROI:* Prevents 35% of truck rolls, lifting CSAT from 3.85 to 4.45 and saving **$78,000 annually**.

---

## 💻 Interactive Web Application

In addition to the Excel file, an interactive executive web cockpit is included:

- **Sleek Executive Dark & Light Themes** with instantaneous toggle and persistence.
- **Dynamic Chart.js Visualizations**: Departmental SLA compliance against targets, revenue at risk distribution, and 12-month trajectory curves.
- **Interactive What-If Scenario Simulator**: Adjust sliders for AI Self-Service Deflection, Tier 2 MTTR reduction, and contract conversion to compute projected savings, preserved ARR, and margin expansion in real time.
- **Searchable Master Data Explorer**: Search by Customer ID, department, or contract with pagination and column sorting.
- **1-Click Excel Export**: Download the `.xlsx` workbook directly from the web navigation bar.

### Accessing the Web Application:
```bash
# Start local server if not already running:
python3 -m http.server 8085

# Open in browser:
http://localhost:8085/
```

---

## 🛠️ Regeneration & CLI Scripting

To regenerate or refresh the Excel workbook with new input data:

```bash
# Using the local environment
.venv/bin/python generate_dashboard.py --data data/raw/customer_churn_data.csv --output Operations_and_Business_Performance_Dashboard.xlsx
```

---

## 📁 Repository Structure

```
Operation and Business performance Dashboard/
├── Operations_and_Business_Performance_Dashboard.xlsx  # Master compiled Excel dashboard
├── generate_dashboard.py                               # CLI runner script
├── index.html                                          # Interactive web cockpit structure
├── style.css                                           # Design system, themes & animations
├── app.js                                              # Frontend charts, simulator & filters
├── dashboard_preview.jpg                               # Visual dashboard preview
├── README.md                                           # Comprehensive documentation
├── data/
│   ├── raw/
│   │   └── customer_churn_data.csv                     # Raw operational customer dataset
│   └── dashboard_data.json                             # Enriched JSON dataset for web UI
└── src/
    ├── __init__.py
    └── operational_dashboard.py                        # Core openpyxl generation engine
```
