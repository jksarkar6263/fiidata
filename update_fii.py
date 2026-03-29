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
    raise Exception("No NSE file found in last 7 days")

with open("temp.xls", "wb") as f:
    f.write(file_content)

# -------------------------------
# STEP 2 — READ SHEET2 EXACTLY
# -------------------------------
df = pd.read_excel("temp.xls", header=None)
df = df.fillna("")

# -------------------------------
# STEP 3 — HELPERS
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

def is_category_row(text):
    text = text.upper()
    keywords = ["INDEX FUTURES", "INDEX OPTIONS", "STOCK FUTURES", "STOCK OPTIONS"]
    return any(k in text for k in keywords)

# -------------------------------
# STEP 4 — BUILD TABLE MANUALLY
# -------------------------------
table_html = "<table>"

for r in range(len(df)):
    row_values = df.iloc[r].tolist()
    row_text = " ".join([str(x) for x in row_values])

    # highlight major rows
    if is_category_row(row_text):
        table_html += "<tr class='category'>"
    else:
        table_html += "<tr>"

    for c, val in enumerate(row_values):
        style = ""

        # first column bold
        if c == 0:
            style += "font-weight:bold; text-align:left;"

        # NET columns (last two columns)
        if c >= len(row_values) - 2:
            style += "font-weight:bold; font-size:13px;"
            style += f"color:{number_color(val)};"

        text = str(val)

        # rotate credit text
        if "jayfromstockmarketsinindia" in text.lower():
            text = f"<div class='rotate'>{text}</div>"

        table_html += f"<td style='{style}'>{text}</td>"

    table_html += "</tr>"

table_html += "</table>"

# -------------------------------
# STEP 5 — FINAL WEBPAGE
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

.container {{
    max-width:770px;
    margin:auto;
}}

table {{
    width:100%;
    border-collapse:collapse;
    font-size:11px;
}}

td {{
    border:1px solid #d0d7e5;
    padding:6px 8px;
    text-align:center;
}}

.category {{
    background:#e8eefc;
}}

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

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)

print("index.html generated successfully")
