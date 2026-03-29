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
# STEP 2 — READ ORIGINAL XLS
# -------------------------------
df = pd.read_excel("temp.xls", header=None)
df = df.fillna("")

# Fix header wording
df = df.replace(
    ["Amt in Crores", "Amount (in Crores)", "Amount (Crores)"],
    "Amount (in ₹ Crores)"
)

# -------------------------------
# STEP 3 — CALCULATE NET (CORRECT LOGIC)
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

# Insert NET columns AFTER Sell columns → position 5 & 6
df.insert(5, "NET Contracts", net_contracts)
df.insert(6, "NET Amount", net_amounts)

# -------------------------------
# STEP 4 — COLOR FUNCTION
# -------------------------------
def number_color(val):
    try:
        v = float(str(val).replace(",", ""))
        if v > 0:
            return "green"
        elif v < 0:
            return "red"
    except:
        pass
    return "black"

# -------------------------------
# STEP 5 — BUILD HTML TABLE MANUALLY
# -------------------------------
table_html = "<table>"

for r in range(len(df)):
    table_html += "<tr>"
    for c, val in enumerate(df.iloc[r]):
        style = "text-align:center; padding:6px; border:1px solid #ccc;"

        # First column bold left
        if c == 0:
            style += "font-weight:bold; text-align:left;"

        # NET columns (F & G → index 5 & 6)
        if c in [5,6]:
            style += "font-weight:bold; font-size:13px;"
            style += f"color:{number_color(val)};"

        table_html += f"<td style='{style}'>{val}</td>"
    table_html += "</tr>"

table_html += "</table>"

# -------------------------------
# STEP 6 — FINAL WEBPAGE
# -------------------------------
html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{font-family:Arial}}
.container {{max-width:770px;margin:auto}}
table {{border-collapse:collapse;width:100%;font-size:12px}}
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
