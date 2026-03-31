import pandas as pd
import requests
from datetime import datetime, timedelta

# -------------------------------
# STEP 1 — DOWNLOAD LATEST NSE FILE
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
    raise Exception("No NSE file found")

with open("temp.xls", "wb") as f:
    f.write(file_content)

# -------------------------------
# STEP 2 — READ XLS
# -------------------------------
df = pd.read_excel("temp.xls", header=None)
df = df.fillna("")

df = df.replace(
    ["Amt in Crores", "Amount (in Crores)", "Amount (Crores)"],
    "Amount (₹ Crores)"
)

# -------------------------------
# STEP 3 — CALCULATE NET
# -------------------------------
net_contracts = []
net_amounts = []

for i in range(len(df)):
    try:
        buy_contracts  = float(df.iloc[i,1])
        buy_amount     = float(df.iloc[i,2])
        sell_contracts = float(df.iloc[i,3])
        sell_amount    = float(df.iloc[i,4])

        net_contracts.append(buy_contracts - sell_contracts)
        net_amounts.append(buy_amount - sell_amount)
    except:
        net_contracts.append("")
        net_amounts.append("")

df.insert(5, "NET Contracts", net_contracts)
df.insert(6, "NET Amount", net_amounts)

# -------------------------------
# STEP 4 — FORMAT NUMBERS
# -------------------------------
def format_contract(val):
    try:
        return f"{int(float(val)):,}"
    except:
        return val

def format_amount(val):
    try:
        return f"{float(val):,.2f}"
    except:
        return val

for r in range(len(df)):
    for c in range(len(df.columns)):
        if r < 2:
            continue
        if c in [1,3,5,7]:
            df.iat[r,c] = format_contract(df.iat[r,c])
        if c in [2,4,6,8]:
            df.iat[r,c] = format_amount(df.iat[r,c])

# -------------------------------
# STEP 5 — COLOR FUNCTION
# -------------------------------
def color_net(val):
    try:
        v = float(str(val).replace(",", ""))
        if v > 0: return "green"
        if v < 0: return "red"
    except:
        pass
    return "black"

# -------------------------------
# STEP 6 — BUILD HTML TABLE
# -------------------------------
table_html = "<table class='fii'>"

# TOP BAR
table_html += f"""
<tr class='topbar'>
<td colspan='5' class='left bold'>
DETAILED FII DERIVATIVES DATA FOR {file_date}
</td>
<td colspan='4' class='num bold'>
Last updated on {file_date}
</td>
</tr>
"""

# HEADERS
table_html += """
<tr class='subhead'>
  <th rowspan='2' class='credit firstcol'>
    <div class='rotate'>jayfromstockmarketsinindia</div>
  </th>
  <th colspan='2'>BUY</th>
  <th colspan='2'>SELL</th>
  <th colspan='2'>NET</th>
  <th colspan='2'>OPEN INTEREST</th>
</tr>

<tr class='subsub'>
  <th>No. of Contracts</th>
  <th>Amount<br>(₹ Crores)</th>
  <th>No. of Contracts</th>
  <th>Amount<br>(₹ Crores)</th>
  <th>No. of Contracts</th>
  <th>Amount<br>(₹ Crores)</th>
  <th>No. of Contracts</th>
  <th>Amount<br>(₹ Crores)</th>
</tr>
"""

major_rows = ["INDEX FUTURES","INDEX OPTIONS","STOCK FUTURES","STOCK OPTIONS"]

for r in range(2, len(df)):
    row = df.iloc[r].tolist()
    name = str(row[0]).strip().upper()

    if name == "":
        continue

    if name in major_rows:
        table_html += "<tr class='separator'><td colspan='9'></td></tr>"
        table_html += "<tr class='category'>"
    else:
        table_html += "<tr>"

    table_html += f"<td class='left bold'>{row[0]}</td>"

    for i in range(1,9):
        val = row[i]
        style="text-align:right;"
        if i in [5,6]:
            style += f"background:#dde5ff;font-weight:bold;color:{color_net(val)};"
        table_html += f"<td style='{style}'>{val}</td>"

    table_html += "</tr>"

# -------------------------------
# STEP 7 — NOTES (MERGED + TIGHT ROWS)
# -------------------------------
notes = [
    "Notes:",
    "Both buy and sell positions have been considered",
    "Options Value (Buy/Sell) = Strike price * Qty",
    "Futures Value (Buy/Sell) = Traded Price * Qty",
    "Value & Open Interest at the end of day:",
    "Options Value (End of day) = Underlying Close Price * Qty",
    "Futures Value (End of day) = Closing Futures Price * Qty (daily settlement price)"
]

table_html += "<tr class='separator'><td colspan='9'></td></tr>"

for n in notes:
    if n == "Notes:":
        table_html += f"<tr class='noteshead'><td colspan='9'>{n}</td></tr>"
    else:
        table_html += f"<tr class='notes'><td colspan='9'>{n}</td></tr>"

table_html += "</table>"

# -------------------------------
# STEP 8 — FINAL PAGE
# -------------------------------
html = f"""
<html>
<head>
<style>
body {{font-family:Arial;background:#eef2ff}}
.container {{max-width:770px;margin:auto}}
table {{width:100%;border-collapse:collapse;font-size:11px}}
td,th {{border:1px solid #cfd6e6;padding:6px}}
.topbar {{background:#dbe4ff;font-weight:bold;font-size:14px}}
.subhead {{background:#244c9a;color:white}}
.subsub {{background:#4f74c9;color:white}}
.category {{background:#dde5ff;font-weight:bold}}
.separator td{{height:4px;background:#4f74c9;border:none;padding:0}}

.noteshead td{{font-size:10px;font-weight:bold;padding:3px}}
.notes td{{font-size:9px;padding:2px}}
</style>
</head>
<body>
<div class="container">
{table_html}
</div>
</body>
</html>
"""

with open("index.html","w",encoding="utf-8") as f:
    f.write(html)

print("SUCCESS — FILE GENERATED")
