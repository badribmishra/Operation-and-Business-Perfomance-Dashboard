#!/usr/bin/env python3
"""
Operations & Business Performance Dashboard Generator CLI.
Compiles a corporate-grade Excel workbook (.xlsx) to organize data, track KPIs,
analyze 12-month operational trends, and deliver data-driven decision frameworks.

Usage:
    python generate_dashboard.py
    python generate_dashboard.py --data data/raw/customer_churn_data.csv --output Operations_Dashboard.xlsx
"""

import argparse
import os
import sys
import time

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from src.operational_dashboard import generate_full_operational_dashboard


def main():
    parser = argparse.ArgumentParser(
        description="Generate Executive Operations & Business Performance Excel Dashboard"
    )
    parser.add_argument(
        "--data",
        type=str,
        default=os.path.join(CURRENT_DIR, "data", "raw", "customer_churn_data.csv"),
        help="Path to input operational dataset (.csv)"
    )
    parser.add_argument(
        "--output",
        type=str,
        default=os.path.join(CURRENT_DIR, "Operations_and_Business_Performance_Dashboard.xlsx"),
        help="Path to destination Excel workbook (.xlsx)"
    )
    args = parser.parse_args()

    print("=" * 80)
    print(" 🛡️  OPERATIONS & BUSINESS PERFORMANCE EXECUTIVE DASHBOARD")
    print("=" * 80)
    print(f"📂 Source Data:   {args.data}")
    print(f"📊 Output Excel:  {args.output}")
    print("⏳ Processing records, synthesizing SLA telemetry, and compiling sheets...")

    t0 = time.time()
    df_ops, df_trends, decisions = generate_full_operational_dashboard(
        raw_data_path=args.data,
        output_excel_path=args.output
    )
    elapsed = time.time() - t0

    # Calculate summary highlights
    total_accounts = len(df_ops)
    total_mrr = df_ops["MonthlyCharges"].sum()
    arr = total_mrr * 12
    churn_rate = (df_ops["Churn"] == "Yes").mean() * 100
    sla_compliance = (df_ops["SLAMet"] == "Yes").mean() * 100
    avg_csat = df_ops["CSATScore"].mean()
    cost_to_serve = df_ops["CostToServe"].sum()
    net_margin_pct = ((total_mrr - cost_to_serve) / total_mrr) * 100
    at_risk_mrr = df_ops.loc[df_ops["Churn"] == "Yes", "MonthlyCharges"].sum()

    print(f"\n✨ Compilation completed in {elapsed:.2f} seconds!")
    print("-" * 80)
    print("📈 EXECUTIVE PERFORMANCE SCORECARD:")
    print(f" • Total Customer Accounts:       {total_accounts:,}")
    print(f" • Monthly Recurring Revenue:     ${total_mrr:,.2f}")
    print(f" • Annual Run Rate (ARR):         ${arr:,.2f}")
    print(f" • Overall Churn Rate:            {churn_rate:.2f}%")
    print(f" • Monthly Revenue at Risk:       ${at_risk_mrr:,.2f}")
    print(f" • Overall SLA Compliance:        {sla_compliance:.2f}% (Target: >=85.0%)")
    print(f" • Portfolio Average CSAT:        {avg_csat:.2f} / 5.0")
    print(f" • Total Cost to Serve:           ${cost_to_serve:,.2f}")
    print(f" • Net Operational Margin:        {net_margin_pct:.2f}%")
    print("-" * 80)
    print("📋 WORKBOOK SHEET DIRECTORY:")
    print(" [1] 🛡️ Executive KPI Summary        - Scorecards, Dept Summary & Risk Segmentation")
    print(" [2] ⚙️ Departmental Operations & SLAs - Granular handling time, SLA breach & unit costs")
    print(" [3] 📈 Trends & Cohort Analytics      - 12-Month historical MRR, churn & SLA trajectory")
    print(" [4] 🎯 Data-Driven Decisions         - Actionable interventions, ROI & root cause matrix")
    print(" [5] 📋 Master Operational Data        - Formatted master table (7,043 rows x 34 cols)")
    print(" [6] 📖 Data Dictionary & KPI Guide   - Benchmark targets, calculation formulas & ownership")
    print("=" * 80)


if __name__ == "__main__":
    main()
