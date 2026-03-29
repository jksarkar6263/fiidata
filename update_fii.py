import pandas as pd
import requests
from datetime import datetime, timedelta
import base64

# -------------------------------
# STEP 1 — FIND LATEST NSE FILE
# -------------------------------
session = requests.Session()

headers = {
    "User-Agent": "Mozilla/5.0",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/"
}

# visit homepage to get cookies
session.get("https://www.nseindia.com", headers=headers)

base = "https://nsearchives.nseindia.com/content/fo/fii_stats_{}.xls"

file_content = None
file_date = None

for i in range(7):  # look back 7 days
    d = datetime.now() - timedelta(days=i)
    date_str = d.strftime("%d-%b-%Y")
    url = base.format(date_str)

    r = session.get(url, headers=headers)

    if r.status_code == 200:
        file_content = r.content
        file_date = date_str
        print("Found file:", date_str)
        break

if file_content is None:
    raise Exception("No NSE file found")

# -------------------------------
# STEP 2 — READ XLS
# -------------------------------
with open("temp.xls", "wb") as f:
    f.write(file_content)

df = pd.read_excel("temp.xls")

# -------------------------------
# STEP 3 — BASIC CLEANING
# -------------------------------
df.fillna("", inplace=True)

# -------------------------------
# STEP 4 — GENERATE HTML TABLE
# -------------------------------
def color_value(val):
    try:
        v = float(val)
        if v > 0:
            return "green"
        elif v < 0:
            return "red"
    except:
        pass
    return "black"

html = f"""
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
  font-family: Arial;
}}
.container {{
  max-width:770px;
  margin:auto;
}}
table {{
  width:100%;
  border-collapse:collapse;
  font-size:11px;
}}
th {{
  background:#002550;
  color:white;
  padding:6px;
}}
td {{
  padding:4px;
  text-align:center;
  border:1px solid #ccc;
}}
</style>
</head>
<body>
<div class="container">
<p>Last Updated on: {file_date}</p>
<table>
"""

# headers
html += "<tr>"
for col in df.columns:
    html += f"<th>{col}</th>"
html += "</tr>"

# rows
for _, row in df.iterrows():
    html += "<tr>"
    for val in row:
        color = color_value(val)
        html += f"<td style='color:{color};font-weight:bold'>{val}</td>"
    html += "</tr>"

html += "</table></div></body></html>"

with open("FII_Data.html", "w", encoding="utf-8") as f:
    f.write(html)

print("HTML generated!")
