import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 250)
pd.options.display.float_format = "{:,.2f}".format

INPUT_DIR = Path("data_input")

files = list(INPUT_DIR.glob("*.xlsx"))
print(files)

df = pd.read_excel(files[0], sheet_name="ОСВ", header=None)

ws = load_workbook(files[0])["ОСВ"]
df["indent"] = [ws.cell(row=i + 1, column=1).alignment.indent for i in range(len(df))]

first_col = df[0].astype(str)

headers = first_col[first_col.str.contains("по счету")]

accounts = headers.str.extract(r"по счету (\d+)")[0]

starts = headers.index.tolist()
ends = starts[1:] + [len(df)]

blocks = {}
for account, start, end in zip(accounts, starts, ends):
    blocks[account] = df.iloc[start:end]
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


tables = []
for account in blocks:
    tables.append(clean_block(blocks[account], account))

osv = pd.concat(tables, ignore_index=True)
print(osv)
print(osv.groupby("account")[["debit", "credit"]].sum())
