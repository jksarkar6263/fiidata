import pandas as pd
import requests
from datetime import datetime, timedelta

# -------------------------------
# STEP 1 — FIND LATEST NSE FILE
# -------------------------------
session = requests.Session()

headers = {
    "User-Agent": "Mozilla/5.0",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/"
}

session.get("https://www.nseindia.com", headers=headers)

base = "https://nsearchives.nseindia.com/content/fo/fii_stats_{}.xls"

file_content = None
file_date = None

for i in range(7):
    d = datetime.now() - timedelta(days=i)
    date_str = d.strftime("%d-%b-%Y")
    url = base.format(date_str)

    r = session.get(url, headers=headers)

    if r.status_code == 200:
        file_content = r.content
        file_date = date_str
        print("Found NSE file:", date_str)
        break

if file_content is None:
    raise Exception("No NSE file found in last 7 days")

with open("temp.xls", "wb") as f:
    f.write(file_content)

# -------------------------------
# STEP 2 — READ NSE XLS
# -------------------------------
df = pd.read_excel("temp.xls", header=None)
df = df.fillna("")

# -------------------------------
# STEP 3 — ADD NET COLUMNS (Sheet2 formulas)
# -------------------------------
BUY_CONTRACT_COL = 2
BUY_AMOUNT_COL   = 3
SELL_CONTRACT_COL = 4
SELL_AMOUNT_COL   = 5

net_contracts = []
net_amounts = []

for i in range(len(df)):
    try:
        buy_c = float(df.iloc[i, BUY_CONTRACT_COL])
        sell_c = float(df.iloc[i, SELL_CONTRACT_COL])
        buy_a = float(df.iloc[i, BUY_AMOUNT_COL])
        sell_a = float(df.iloc[i, SELL_AMOUNT_COL])

        net_contracts.append(buy_c - sell_c)
        net_amounts.append(buy_a - sell_a)
    except:
        net_contracts.append("")
        net_amounts.append("")

df["NET Contracts"] = net_contracts
df["NET Amount"] = net_amounts

# -------------------------------
# STEP 4 — CONVERT TO HTML TABLE
# -------------------------------
table_html = df.to_html(index=False, header=False, border=0)

# Fix header wording globally
table_html = table_html.replace("Amt in Crores", "Amount (₹ Crores)")

# -------------------------------
# STEP 5 — BUILD WEBPAGE
# -------------------------------
html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>FII Derivative Data</title>

<style>
body {{
    font-family: Arial, Helvetica, sans-serif;
    background:white;
}}

.container {{
    max-width:770px;
    margin:auto;
}}

table {{
    border-collapse: collapse;
    font-size: 12px;
    width:100%;
}}

td {{
    padding: 6px 10px;
    border:1px solid #ccc;
    text-align:center;
}}
</style>

</head>
<body>

<div class="container">
<p><b>Last updated: {file_date}</b></p>
{table_html}
</div>

</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("index.html generated successfully")
