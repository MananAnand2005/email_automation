from flask import Flask, render_template, request, send_file
import pandas as pd
import win32com.client as win32
import pythoncom
import os, zipfile, io, time, re
from datetime import datetime
from dateutil import parser

app = Flask(__name__)

# ✅ Static Owners File
OWNERS_FILE = "data/owners.csv"

# ✅ Ensure owners file exists
if not os.path.exists(OWNERS_FILE):
    os.makedirs("data", exist_ok=True)
    pd.DataFrame(columns=["Name", "Email"]).to_csv(OWNERS_FILE, index=False)

# -------------------------
# Helper Functions
# -------------------------

def format_acv(value):
    try:
        if pd.isna(value) or value == "":
            return ""
        num = float(value)
        k_value = round(num / 1000, 1)
        return f"${int(k_value)}k" if k_value.is_integer() else f"${k_value}k"
    except:
        return str(value)

def generate_styled_table(df, sort_col, sort_ord, rules):
    header_style = "background-color:#002451;color:white;padding:8px;border:1px solid #000;font-family:Calibri;font-size:11pt;"
    cell_style = "padding:8px;border:1px solid #000;font-family:Calibri;font-size:11pt;"
    
    # Sorting
    if sort_col and sort_col in df.columns:
        df = df.sort_values(by=sort_col, ascending=(sort_ord == 'asc'))

    html = '<table style="border-collapse:collapse;width:100%;border:1px solid #000;"><tr>'
    for c in df.columns:
        html += f'<th style="{header_style}">{c}</th>'
    html += '</tr>'

    if df.empty:
        return html + "</table>"

    for _, row in df.iterrows():
        html += '<tr>'

        for col in df.columns:
            val = row[col]

            # ✅ Value formatting
            if pd.isna(val):
                display_val = ""
            elif "ACV" in str(col).upper():
                display_val = format_acv(val)
            elif "DATE" in str(col).upper():
                dt = pd.to_datetime(val, errors='coerce')
                display_val = dt.strftime('%m/%d/%Y') if pd.notnull(dt) else str(val)
            else:
                display_val = str(val)

            style = cell_style

            # ✅ FULL RULE ENGINE
            for r_col, r_op, r_val in rules:
                if col == r_col:
                    trigger = False
                    try:
                        if r_op == 'eq':
                            targets = [x.strip() for x in str(r_val).split(',')]
                            if str(val) in targets:
                                trigger = True

                        elif r_op == 'neq':
                            if str(val) != str(r_val):
                                trigger = True

                        elif r_op == 'gt':
                            if pd.to_numeric(val, errors='coerce') > float(r_val):
                                trigger = True

                        elif r_op == 'lt':
                            if pd.to_numeric(val, errors='coerce') < float(r_val):
                                trigger = True

                        elif r_op == 'gte':
                            if pd.to_numeric(val, errors='coerce') >= float(r_val):
                                trigger = True

                        elif r_op == 'lte':
                            if pd.to_numeric(val, errors='coerce') <= float(r_val):
                                trigger = True

                        elif r_op in ['before', 'after', 'on']:
                            dt = pd.to_datetime(val, errors='coerce')
                            if not pd.isna(dt):
                                target = datetime.now() if str(r_val).lower() == 'today' else parser.parse(str(r_val))

                                if r_op == 'before' and dt < target:
                                    trigger = True
                                elif r_op == 'after' and dt > target:
                                    trigger = True
                                elif r_op == 'on' and dt.date() == target.date():
                                    trigger = True

                    except:
                        pass

                    if trigger:
                        style += " color:red; font-weight:bold;"

            html += f'<td style="{style}">{display_val}</td>'

        html += '</tr>'

    return html + '</table>'


# -------------------------
# Main Route
# -------------------------

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        pythoncom.CoInitialize()

        try:
            temp_dir = f"drafts_{int(time.time())}"
            os.makedirs(temp_dir, exist_ok=True)
            temp_dir = os.path.abspath(temp_dir)

            # ✅ Load static owners
            owners_df = pd.read_csv(OWNERS_FILE)
            owners_df['FirstName'] = owners_df['Name'].astype(str).apply(lambda x: x.split()[0])
            owners_df['FullName'] = owners_df['Name']

            # ✅ Get rules
            rules = list(zip(
                request.form.getlist('rule_col[]'),
                request.form.getlist('rule_op[]'),
                request.form.getlist('rule_val[]')
            ))

            # ✅ Load Excel
            data_file = request.files['data']
            xl = pd.ExcelFile(data_file)

            outlook = win32.Dispatch("Outlook.Application")
            zip_buf = io.BytesIO()

            match_column = request.form['match_col']

            with zipfile.ZipFile(zip_buf, 'w') as zf:
                for _, owner_row in owners_df.iterrows():
                    owner_name = str(owner_row['Name']).strip()
                    owner_email = owner_row['Email']

                    mail = outlook.CreateItem(0)
                    mail.To = owner_email
                    mail.CC = request.form.get('cc', '')
                    mail.Subject = request.form['subject'].format_map(owner_row.to_dict())

                    def replace_table(match):
                        sheet = match.group(1)
                        df = xl.parse(sheet)

                        # ✅ STRONG MATCHING (fix for spaces etc.)
                        filtered = df[
                            df[match_column].astype(str).str.strip() == owner_name
                        ]

                        return generate_styled_table(
                            filtered,
                            request.form.get('sort_col'),
                            request.form.get('sort_order'),
                            rules
                        )

                    body = re.sub(r'\{TABLE:(.*?)\}', replace_table, request.form['message'])
                    mail.HTMLBody = body.format(**owner_row.to_dict())

                    mail.Save()

                    # ✅ SAFE FILE NAME
                    safe_name = re.sub(r'[\\/*?:"<>|]', "", owner_name)[:50]
                    filename = f"{safe_name}.msg"

                    path = os.path.abspath(os.path.join(temp_dir, filename))
                    time.sleep(0.05)

                    mail.SaveAs(path)

                    zf.write(path, filename)
                    mail.Close(0)

            zip_buf.seek(0)
            return send_file(zip_buf, download_name='Drafts.zip', as_attachment=True)

        finally:
            pythoncom.CoUninitialize()

    return render_template('index.html')


# -------------------------
# Run App
# -------------------------

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)