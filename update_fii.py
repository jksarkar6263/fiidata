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

# visit NSE homepage to get cookies
session.get("https://www.nseindia.com", headers=headers)

base = "https://nsearchives.nseindia.com/content/fo/fii_stats_{}.xls"

file_content = None
file_date = None

# NSE uploads only on trading days → check last 7 days
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

# save downloaded file
with open("temp.xls", "wb") as f:
    f.write(file_content)

# -------------------------------
# STEP 2 — READ SHEET 2 EXACTLY
# -------------------------------
df = pd.read_excel("temp.xls", sheet_name="Sheet2", header=None)

# Replace NaN with blanks
df = df.fillna("")

# -------------------------------
# STEP 3 — CONVERT SHEET TO HTML
# -------------------------------
table_html = df.to_html(index=False, header=False, border=0)

# -------------------------------
# STEP 4 — BUILD WEBPAGE
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

table {{
    border-collapse: collapse;
    font-size: 12px;
}}

td {{
    padding: 6px 10px;
}}
</style>

</head>
<body>

{table_html}

</body>
</html>
"""

# GitHub Pages needs index.html
with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("index.html generated successfully")
