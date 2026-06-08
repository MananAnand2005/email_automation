from flask import Flask, render_template, request, send_file
import pandas as pd
import win32com.client as win32
import pythoncom
import os, zipfile, io, time, re
from datetime import datetime
from dateutil import parser

app = Flask(__name__)

# --- Helper for ACV Formatting ---
def format_acv(value):
    try:
        if pd.isna(value) or value == "":
            return ""
        num = float(value)
        k_value = round(num / 1000, 1)
        return f"${int(k_value)}k" if k_value.is_integer() else f"${k_value}k"
    except:
        return str(value)

# --- Table Generator ---
def generate_styled_table(df, sort_col, sort_ord, rules):
    header_style = "background-color:#002451; color:white; padding:8px; border:1px solid #000; font-family:Calibri; font-size:11pt;"
    cell_style = "padding:8px; border:1px solid #000; font-family:Calibri; font-size:11pt;"
    
    # Sorting
    if sort_col and sort_col in df.columns:
        df = df.sort_values(by=sort_col, ascending=(sort_ord == 'asc'))
    
    html = '<table style="border-collapse:collapse; width:100%; border:1px solid #000;"><tr>'
    for c in df.columns:
        html += f'<th style="{header_style}">{c}</th>'
    html += '</tr>'
    
    if df.empty:
        return html + "</table>"
    
    for _, row in df.iterrows():
        html += '<tr>'
        for col in df.columns:
            val = row[col]

            # ✅ NaN → blank
            if pd.isna(val):
                display_val = ""

            # ✅ ACV formatting
            elif "ACV" in str(col).upper():
                display_val = format_acv(val)

            # ✅ DATE formatting (based on column name)
            elif "DATE" in str(col).upper():
                dt = pd.to_datetime(val, errors='coerce')
                if pd.notnull(dt):
                    display_val = dt.strftime('%m/%d/%Y')
                else:
                    display_val = str(val)

            # ✅ Default
            else:
                display_val = str(val)

            style = cell_style

            # ✅ Rules engine
            for r_col, r_op, r_val in rules:
                if col == r_col:
                    trigger = False
                    try:
                        if r_op == 'eq':
                            targets = [x.strip() for x in str(r_val).split(',')]
                            if str(val) in targets:
                                trigger = True
                        elif r_op == 'neq' and str(val) != str(r_val):
                            trigger = True
                        elif r_op == 'gt':
                            if pd.to_numeric(val, errors='coerce') > float(r_val):
                                trigger = True
                        elif r_op == 'lt':
                            if pd.to_numeric(val, errors='coerce') < float(r_val):
                                trigger = True
                        elif r_op in ['before', 'after', 'on']:
                            dt = pd.to_datetime(val, errors='coerce')
                            if not pd.isna(dt):
                                target = datetime.now() if r_val.lower() == 'today' else parser.parse(r_val)
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

# --- Route Logic ---
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        pythoncom.CoInitialize()
        try:
            unique_id = str(int(time.time()))
            temp_dir = os.path.join(os.getcwd(), f"drafts_{unique_id}")
            os.makedirs(temp_dir)
            
            owners_df = pd.read_csv(request.files['owners'])
            owners_df['FirstName'] = owners_df[request.form['owner_col']].apply(lambda x: str(x).split(' ')[0])
            owners_df['FullName'] = owners_df[request.form['owner_col']]
            
            rules = list(zip(
                request.form.getlist('rule_col[]'), 
                request.form.getlist('rule_op[]'), 
                request.form.getlist('rule_val[]')
            ))
            
            outlook = win32.Dispatch("Outlook.Application")
            zip_buf = io.BytesIO()
            
            with zipfile.ZipFile(zip_buf, 'w') as zf:
                for _, row in owners_df.iterrows():
                    mail = outlook.CreateItem(0)
                    mail.To = row[request.form['email_col']]
                    mail.CC = request.form.get('cc', '')
                    mail.Subject = request.form['subject'].format(**row.to_dict())
                    
                    def replace_table(match):
                        sheet = match.group(1)
                        try:
                            df = pd.read_excel(request.files['data'], sheet_name=sheet)
                            filtered = df[df[request.form['data_col']] == row[request.form['owner_col']]]
                            return generate_styled_table(
                                filtered,
                                request.form.get('sort_col'),
                                request.form.get('sort_order'),
                                rules
                            )
                        except Exception as e:
                            return f"<p>Error in {sheet}: {str(e)}</p>"
                    
                    body = re.sub(r'\{TABLE:(.*?)\}', replace_table, request.form['message'])
                    mail.HTMLBody = body.format(**row.to_dict())
                    
                    mail.Save()
                    
                    filename = f"{str(row[request.form['owner_col']]).replace('/', '-')}.msg"
                    path = os.path.join(temp_dir, filename)
                    mail.SaveAs(path)
                    
                    zf.write(path, os.path.basename(path))
                    mail.Close(0)
            
            zip_buf.seek(0)
            return send_file(zip_buf, download_name='Outlook_Drafts.zip', as_attachment=True)
            
        finally:
            pythoncom.CoUninitialize()
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)