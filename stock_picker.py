"""
Stock Picker — Final Investment List
"""

import pandas as pd
import os
import shutil
from datetime import date

# ── FILES ─────────────────────────────────────────
SCANNER_FILE  = "buy_signals_report.xlsx"
SCORES_CSV    = "fundamental_scores.csv"
BACKTEST_FILE = "backtest_report.xlsx"
MASTER_FILE   = "nse_master_reference.csv"

OUTPUT_FILE   = "stock_picker_report.xlsx"

# ── SETTINGS ──────────────────────────────────────

TOP_SECTORS = [
    "Industrials",
    "Financial Services",
    "Basic Materials",
    "Consumer Cyclical",
    "Healthcare"
]

AVOID_SYMBOLS = {
    "IGL","WHIRLPOOL","BATAINDIA",
    "RELINFRA","MOTILALOFS","KAJARIACER"
}

# ──────────────────────────────────────────────────


def load_all():

    print(f"Stock Picker — {date.today()}")
    print("="*50)

    if not os.path.exists(SCANNER_FILE):
        print("Run car_signal_scanner.py first.")
        return None,None,None,None

    scanner_df = pd.read_excel(SCANNER_FILE)

    scanner_df["Symbol"] = (
        scanner_df["Symbol"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    print(f"CAR Signals : {len(scanner_df)}")

    scores_df=None
    if os.path.exists(SCORES_CSV):
        scores_df=pd.read_csv(SCORES_CSV)
        scores_df["Symbol"]=scores_df["Symbol"].str.upper()

    backtest_df=None
    if os.path.exists(BACKTEST_FILE):
        backtest_df=pd.read_excel(BACKTEST_FILE)

    master_df=None
    if os.path.exists(MASTER_FILE):
        master_df=pd.read_csv(MASTER_FILE)

    return scanner_df,scores_df,backtest_df,master_df


def build_picks(scanner_df,scores_df,backtest_df,master_df):

    picks=scanner_df.copy()

    if scores_df is not None:

        cols=[c for c in [
            "Symbol","Sector","Company Name",
            "Cap Type","Index Category",
            "ROE %","PE Ratio","Debt/Equity"
        ] if c in scores_df.columns]

        picks=picks.merge(scores_df[cols],on="Symbol",how="left")

    picks=picks[~picks["Symbol"].isin(AVOID_SYMBOLS)]

    picks["Rank Score"]=1
    picks=picks.sort_values("Rank Score",ascending=False)

    picks["Rank"]=range(1,len(picks)+1)

    picks["Recommendation"]="BUY"

    return picks


def write_excel(picks):

    from openpyxl import Workbook

    wb=Workbook()
    ws=wb.active

    ws.title="Todays BUY List"

    headers=list(picks.columns)

    ws.append(headers)

    for row in picks.itertuples(index=False):
        ws.append(list(row))

    wb.save(OUTPUT_FILE)

    # ── BACKUP COPY ─────────────────────────

    today=date.today().strftime("%Y-%m-%d")

    name,ext=os.path.splitext(OUTPUT_FILE)

    backup=f"{name}_{today}{ext}"

    shutil.copy(OUTPUT_FILE,backup)

    return backup


def print_summary(picks,backup_file):

    print("\n"+"="*70)
    print(f" TODAY'S STOCK PICKS — {date.today()}")
    print("="*70)

    for _,row in picks.head(15).iterrows():

        sym=row["Symbol"]
        rank=row["Rank"]

        print(f"{rank:>3}  {sym}")

    print("\nTotal Signals :",len(picks))

    print("\nReport saved :",OUTPUT_FILE)

    print("Backup file  :",backup_file)

    print("="*70)


def main():

    scanner_df,scores_df,backtest_df,master_df=load_all()

    if scanner_df is None:
        return

    if len(scanner_df)==0:
        print("No signals today")
        return

    picks=build_picks(
        scanner_df,
        scores_df,
        backtest_df,
        master_df
    )

    backup_file=write_excel(picks)

    print_summary(picks,backup_file)


if __name__=="__main__":
    main()