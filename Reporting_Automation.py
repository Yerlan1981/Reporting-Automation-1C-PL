import pandas as pd
from pathlib import Path

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 250)

INPUT_DIR = Path("data_input")

files = list(INPUT_DIR.glob("*.xlsx"))
print(files)

df = pd.read_excel(files[0], sheet_name="ОСВ", header=None)

first_col = df[0].astype(str)

headers = first_col[first_col.str.contains("по счету")]
print(headers)

accounts = headers.str.extract(r"по счету (\d+)")[0]
print(accounts)

starts = headers.index.tolist()
ends = starts[1:] + [len(df)]

blocks = {}
for account, start, end in zip(accounts, starts, ends):
    blocks[account] = df.iloc[start:end]
    print(account, start, end, len(blocks[account]))
    