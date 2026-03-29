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
    "Amount (₹ Crores)"
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
# STEP 4 — NUMBER FORMATTING ENGINE
# -------------------------------

def format_contract(val):
    """Format contract numbers → integer with commas"""
    try:
        num = float(val)
        return f"{int(num):,}"
    except:
        return val

def format_amount(val):
    """Format amount → 2 decimals with commas"""
    try:
        num = float(val)
        return f"{num:,.2f}"
    except:
        return val

# Apply formatting to dataframe cells
for r in range(len(df)):
    for c in range(len(df.columns)):

        value = df.iat[r, c]

        # Skip header rows (first 2 rows in Sheet2)
        if r < 2:
            continue

        # Column positions (Excel structure)
        # 1 = Buy Contracts
        # 2 = Buy Amount
        # 3 = Sell Contracts
        # 4 = Sell Amount
        # 5 = NET Contracts
        # 6 = NET Amount
        # 7 = OI Contracts
        # 8 = OI Amount

        if c in [1,3,5,7]:      # Contracts columns
            df.iat[r, c] = format_contract(value)

        if c in [2,4,6,8]:      # Amount columns
            df.iat[r, c] = format_amount(value)

# -------------------------------
# STEP 5 — COLOR FUNCTION
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
# STEP 6 — BUILD TRUE EXCEL-LAYOUT TABLE
# -------------------------------
def color_net(val):
    try:
        v = float(str(val).replace(",", ""))
        if v > 0:
            return "green"
        elif v < 0:
            return "red"
    except:
        return "black"
    return "black"

table_html = """
<table class='fii'>
"""

# ================= HEADER ROW 1 =================


# ================= HEADER ROW 2 =================
table_html += """
<tr class='tophead'>
  <th rowspan='3' class='credit'>
    <div class='rotate'>jayfromstockmarketsinindia</div>
  </th>
  <th colspan='2'>BUY</th>
  <th colspan='2'>SELL</th>
  <th colspan='2'>NET</th>
  <th colspan='2'>OPEN INTEREST</th>
</tr>
"""

# ================= HEADER ROW 3 =================
table_html += """
<tr class='subhead'>
  <th>No. of Contracts</th><th>Amount (₹ Crores)</th>
  <th>No. of Contracts</th><th>Amount (₹ Crores)</th>
  <th>No. of Contracts</th><th>Amount (₹ Crores)</th>
  <th>No. of Contracts</th><th>Amount (₹ Crores)</th>
</tr>
"""

# ================= DATA ROWS =================
major_rows = ["INDEX FUTURES","INDEX OPTIONS","STOCK FUTURES","STOCK OPTIONS"]

for r in range(2, len(df)):  # skip header rows from XLS
    row = df.iloc[r].tolist()
    name = str(row[0]).upper()

    if name.strip() == "":
        continue

    # highlight category rows
    if any(k in name for k in major_rows):
        table_html += "<tr class='category'>"
    else:
        table_html += "<tr>"

    # first column (segment name)
    table_html += f"<td class='left bold'>{row[0]}</td>"

    # remaining numeric columns
    for i in range(1,9):
        val = row[i]
        style = ""

        # NET columns colored
        if i in [5,6]:
            style += f"color:{color_net(val)};font-weight:bold;font-size:12px;"

        table_html += f"<td style='{style}'>{val}</td>"

    table_html += "</tr>"

table_html += "</table>"

# -------------------------------
# STEP 7 — FINAL WEBPAGE
# -------------------------------

html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
 font-family: Arial, Helvetica, sans-serif;
 background:white;
}}

.container{{
 max-width:770px;
 margin:auto;
}}

table.fii{{
 width:100%;
 border-collapse:collapse;
 font-size:11px;
}}

td,th{{
 border:1px solid #cfd6e6;
 padding:6px 6px;
 text-align:right;
}}

.tophead th{{
 background:#002a6e;
 color:white;
 font-size:14px;
}}

.midhead th{{
 background:#244c9a;
 color:white;
 font-size:12px;
}}

.subhead th{{
 background:#4f74c9;
 color:white;
 font-size:11px;
}}

.left{{ text-align:left; }}
.bold{{ font-weight:bold; }}

.category{{
 background:#e8eefc;
 font-weight:bold;
}}

.rotate{{
 transform:rotate(-30deg);
 white-space:nowrap;
 font-weight:bold;
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
