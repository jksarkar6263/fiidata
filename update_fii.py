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

df = df.replace(
    ["Amt in Crores", "Amount (in Crores)", "Amount (Crores)"],
    "Amount (in ₹ Crores)"
)

# -------------------------------
# STEP 3 — ADD NET COLUMNS
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

df.insert(5, "NET Contracts", net_contracts)
df.insert(6, "NET Amount", net_amounts)

# -------------------------------
# HEADER FIX LIKE SHEET2
# -------------------------------
for i in range(len(df)):
    row = " ".join([str(x) for x in df.iloc[i]])

    if "BUY" in row.upper() and "SELL" in row.upper():
        df.iloc[i,5] = "NET"
        df.iloc[i,6] = "NET"

    if "CONTRACT" in row.upper():
        df.iloc[i,5] = "No. of Contracts"
        df.iloc[i,6] = "Amount (in ₹ Crores)"

# -------------------------------
# STYLE HELPERS
# -------------------------------
def number_color(val):
    try:
        v = float(str(val).replace(",", ""))
        if v > 0: return "green"
        if v < 0: return "red"
    except:
        pass
    return "black"

def is_category_row(text):
    text = text.upper()
    keys = ["INDEX FUTURES","INDEX OPTIONS","STOCK FUTURES","STOCK OPTIONS"]
    return any(k in text for k in keys)

# -------------------------------
# BUILD TABLE MANUALLY ⭐
# -------------------------------
table_html = "<table>"

for r in range(len(df)):
    row_values = df.iloc[r].tolist()
    row_text = " ".join([str(x) for x in row_values])

    # highlight category rows
    if is_category_row(row_text):
        table_html += "<tr class='category'>"
    else:
        table_html += "<tr>"

    for c, val in enumerate(row_values):

        style = ""

        # first column bold & left align
        if c == 0:
            style += "font-weight:bold;text-align:left;"

        # NET columns = F & G (index 5 & 6)
        if c == 5 or c == 6:
            style += f"font-weight:bold;color:{number_color(val)};"

        text = str(val)

        # rotate credit text
        if "jayfromstockmarketsinindia" in text.lower():
            text = f"<div class='rotate'>{text}</div>"

        table_html += f"<td style='{style}'>{text}</td>"

    table_html += "</tr>"

table_html += "</table>"

# -------------------------------
# FINAL HTML PAGE
# -------------------------------
html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">

<style>
body {{ font-family: Arial; background:white; }}
.container {{ max-width:770px; margin:auto; }}

table {{ border-collapse:collapse; width:100%; font-size:11px; }}
td {{ border:1px solid #d0d7e5; padding:6px 8px; text-align:center; }}

.category {{ background:#e8eefc; }}

.rotate {{
    transform:rotate(-45deg);
    white-space:nowrap;
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

with open("index.html","w",encoding="utf-8") as f:
    f.write(html)

print("index.html generated successfully")
