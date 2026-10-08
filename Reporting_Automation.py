import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
import re

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 250)
pd.options.display.float_format = "{:,.2f}".format

INPUT_DIR = Path("data_input")
files = [f for f in INPUT_DIR.glob("*.xlsx") if not f.name.startswith("~$")]


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

MONTHS = {
    "Январь": 1, "Февраль": 2, "Март": 3, "Апрель": 4,
    "Май": 5, "Июнь": 6, "Июль": 7, "Август": 8,
    "Сентябрь": 9, "Октябрь": 10, "Ноябрь": 11, "Декабрь": 12,
}

def get_period(df):
    first_col = df[0].astype(str)
    found = first_col[first_col.str.contains(r"за \S+ \d{4}")].iloc[0]
    month_name, year = re.search(r"за (\S+) (\d{4})", found).groups()
    return f"{year}-{MONTHS[month_name]:02d}"

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
    osv["period"] = get_period(df)
    return osv

all_tables = []
for file in files:
    table = load_osv(file)
    all_tables.append(table)

osv = pd.concat(all_tables, ignore_index=True)

print(osv.groupby(["period", "account"])["amount"].sum())

mapping = pd.read_excel("mapping.xlsx", dtype={"account": str})
mapping["article"] = mapping["article"].str.strip()
merged = osv.merge(mapping, on=["account", "article"], how="left")
fallback = pd.DataFrame(
    [
        ("6010", "Доход", "Доход - прочее", 1),
        ("6200", "Доход", "Доход - прочее", 1),
        ("7010", "Себестоимость", "Прочие", 1),
        ("7110", "Расходы по реализации", "Прочее", 1),
        ("7210", "Административные расходы", "Прочее - услуги", 1),
        ("6100", "Прочее", "Прочее", -1),
        ("7300", "Прочее", "Финансовые расходы", 1),
        ("7400", "Прочее", "Курсовые разницы", 1),
        ("7710", "Прочее", "КПН", 1),
    ],
    columns=["account", "fb_section", "fb_pl_line", "fb_sign"],
)

merged = merged.merge(fallback, on="account", how="left")
merged["recognized"] = merged["pl_line"].notna()

merged["section"] = merged["section"].fillna(merged["fb_section"])
merged["pl_line"] = merged["pl_line"].fillna(merged["fb_pl_line"])
merged["sign"] = merged["sign"].fillna(merged["fb_sign"])

unrecognized = merged[(~merged["recognized"]) & (merged["amount"] != 0)]
print("Нераспознанные статьи с суммой:")
print(unrecognized[["account", "article", "amount", "pl_line"]])

settings_table = pd.read_excel("settings.xlsx")
settings = dict(zip(settings_table["key"], settings_table["value"]))

share_by_line = {
    "ТМЗ - основное производство": settings["gas_share"],
    "ТМЗ - логистические расходы": settings["logistics_share"],
}

merged["share"] = merged["pl_line"].map(share_by_line).fillna(1)
merged["final_amount"] = merged["amount"] * merged["sign"] * merged["share"]

report = (
    merged.groupby(["period", "section", "pl_line"], as_index=False)["final_amount"]
    .sum()
)
report["final_amount"] = report["final_amount"] / 1000

totals = report.groupby("section")["final_amount"].sum()

income = totals["Доход"]
cost = totals["Себестоимость"]
selling = totals["Расходы по реализации"]
admin = totals["Административные расходы"]

gross_profit = income - cost
ebitda = gross_profit - selling - admin

def line(section, pl_line):
    mask = (report["section"] == section) & (report["pl_line"] == pl_line)
    return report.loc[mask, "final_amount"].sum()


depreciation = line("Прочее", "Амортизация - Амортизация") + line("Прочее", "Амортизация - Амортизация - офис")
operating_profit = ebitda - depreciation

other_result = line("Прочее", "Курсовые разницы") + line("Прочее", "Выбытие активов") + line("Прочее", "Прочее")
ebit = operating_profit - other_result

finance = line("Прочее", "Финансовые расходы") + line("Прочее", "Начисленные финансовые расходы")
ebt = ebit - finance

tax = line("Прочее", "КПН")
net_profit = ebt - tax

print("Валовая прибыль:", round(gross_profit, 2))
print("EBITDA:", round(ebitda, 2))
print("Операционная прибыль:", round(operating_profit, 2))
print("EBIT:", round(ebit, 2))
print("EBT:", round(ebt, 2))
print("Чистая прибыль:", round(net_profit, 2))

layout = pd.read_excel("layout.xlsx")

full = layout.merge(report, on=["section", "pl_line"], how="left")
full["final_amount"] = full["final_amount"].fillna(0)

print(full[["block", "pl_line", "final_amount"]].to_string())

