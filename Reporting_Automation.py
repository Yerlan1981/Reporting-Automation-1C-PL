import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 250)
pd.options.display.float_format = "{:,.2f}".format

INPUT_DIR = Path("data_input")
files = list(INPUT_DIR.glob("*.xlsx"))


def clean_block(block, account):
    block = block[[0, 4, 5, "indent"]].copy()
    block.columns = ["article", "debit", "credit", "indent"]
    block["article"] = block["article"].astype(str).str.strip()
    block["debit"] = pd.to_numeric(block["debit"], errors="coerce")
    block["credit"] = pd.to_numeric(block["credit"], errors="coerce")

    rows = block[(block["indent"] > 0) & (block["article"] != "Головное подразделение")]
    level = rows["indent"].min()
    rows = rows[rows["indent"] == level]
    rows = rows[rows["debit"].notna() | rows["credit"].notna()]

    rows = rows.assign(account=account)
    return rows[["account", "article", "debit", "credit"]]


def load_osv(path):
    df = pd.read_excel(path, sheet_name="ОСВ", header=None)

    ws = load_workbook(path)["ОСВ"]
    df["indent"] = [ws.cell(row=i + 1, column=1).alignment.indent for i in range(len(df))]

    first_col = df[0].astype(str)
    headers = first_col[first_col.str.contains("по счету")]
    accounts = headers.str.extract(r"по счету (\d+)")[0]

    starts = headers.index.tolist()
    ends = starts[1:] + [len(df)]

    tables = []
    for account, start, end in zip(accounts, starts, ends):
        tables.append(clean_block(df.iloc[start:end], account))

    osv = pd.concat(tables, ignore_index=True)
    osv["amount"] = osv["credit"].where(osv["account"].str.startswith("6"), osv["debit"])
    osv["amount"] = osv["amount"].fillna(0)
    return osv


all_tables = []
for file in files:
    table = load_osv(file)
    table["period"] = file.stem[-7:].replace("_", "-")
    all_tables.append(table)

osv = pd.concat(all_tables, ignore_index=True)

print(osv.groupby(["period", "account"])["amount"].sum())


