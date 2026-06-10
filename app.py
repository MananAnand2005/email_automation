from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import win32com.client as win32
import pythoncom
import os, zipfile, io, time, re, json, unicodedata
from datetime import datetime
from dateutil import parser

app = Flask(__name__)

# ---------------------------------------------------------
# 🛠️ CONFIGURATION & DIRECTORY SETUP
# ---------------------------------------------------------
OWNERS_FILE = "data/owners.csv"
TEMPLATES_FOLDER = "templates_db"

os.makedirs(TEMPLATES_FOLDER, exist_ok=True)
os.makedirs("data", exist_ok=True)

if not os.path.exists(OWNERS_FILE):
    pd.DataFrame(columns=["Name", "Email"]).to_csv(OWNERS_FILE, index=False)

# ---------------------------------------------------------
# 🔬 DEEP NORMALIZATION ENGINE (The Fix)
# ---------------------------------------------------------

def normalize_text(s):
    """
    Ultra-robust text normalization to fix Broadcaster .isin() failures.
    - Handles NaN/Null.
    - Unicode NFKC Normalization (fixes diverse character encodings).
    - Removes Non-Breaking Spaces (\xa0) and Zero-Width Spaces.
    - Standardizes to lowercase and stripped whitespace.
    """
    if pd.isna(s) or s == "" or s is None:
        return ""
    # Normalize unicode characters to a standard form
    s = unicodedata.normalize("NFKC", str(s))
    # Remove common invisible/problematic Excel characters
    s = s.replace("\xa0", " ").replace("\u200b", "").replace("\r", "").replace("\n", " ")
    return s.strip().lower()

# ---------------------------------------------------------
# 📊 FORMATTING ENGINE
# ---------------------------------------------------------

def format_acv(value):
    """ION Standard ACV Formatting: 25000 -> $25k | 20500 -> $20.5k"""
    try:
        if pd.isna(value) or value == "":
            return ""
        num = float(value)
        k_value = round(num / 1000, 1)
        return f"${int(k_value)}k" if k_value.is_integer() else f"${k_value}k"
    except (ValueError, TypeError):
        return str(value)

# ---------------------------------------------------------
# 🛡️ THE LOSSLESS TABLE ENGINE
# ---------------------------------------------------------

def generate_styled_table(df, sort_col, sort_ord, rules):
    """
    Lossless HTML Table Generator:
    - Calibri 11pt styling.
    - Rule Engine for cell-level Red/Bold highlighting.
    - Preserves all columns including Match Column and Blank Comments.
    """
    if sort_col and sort_col in df.columns:
        df = df.sort_values(by=sort_col, ascending=(sort_ord == 'asc'))

    h_style = "background-color:#002451;color:white;padding:10px;border:1px solid #cccccc;font-family:Calibri;font-size:11pt;font-weight:bold;text-align:left;"
    c_style = "padding:10px;border:1px solid #cccccc;font-family:Calibri;font-size:11pt;"
    
    html = '<table style="border-collapse:collapse;width:100%;border:1px solid #cccccc;"><thead><tr>'
    for col_name in df.columns:
        html += f'<th style="{h_style}">{col_name}</th>'
    html += f'<th style="{h_style}">Comments/Update</th></tr></thead><tbody>'

    if df.empty:
        html += f"<tr><td colspan='{len(df.columns) + 1}' style='{c_style};text-align:center;'>No matching data found in Excel.</td></tr>"
    else:
        today = datetime.now()
        for _, row in df.iterrows():
            html += '<tr>'
            for col in df.columns:
                val = row[col]
                
                # Format value
                if pd.isna(val): disp = ""
                elif "ACV" in str(col).upper(): disp = format_acv(val)
                elif "DATE" in str(col).upper():
                    dt = pd.to_datetime(val, errors='coerce')
                    disp = dt.strftime('%m/%d/%Y') if pd.notnull(dt) else str(val)
                else: disp = str(val)

                # Execute Highlighting Rules
                style, red = c_style, False
                for r_col, r_op, r_val in rules:
                    if col == r_col:
                        try:
                            # Use Deep Normalization for string comparisons in rules
                            cv = normalize_text(val)
                            rv = normalize_text(r_val)
                            if r_op == 'eq':
                                if cv in [normalize_text(x) for x in str(r_val).split(',')]: red = True
                            elif r_op == 'neq' and cv != rv: red = True
                            elif r_op == 'gt' and pd.to_numeric(val) > float(r_val): red = True
                            elif r_op == 'lt' and pd.to_numeric(val) < float(r_val): red = True
                            elif r_op == 'gte' and pd.to_numeric(val) >= float(r_val): red = True
                            elif r_op == 'lte' and pd.to_numeric(val) <= float(r_val): red = True
                            elif r_op in ['before', 'after', 'on']:
                                cdt = pd.to_datetime(val, errors='coerce')
                                if not pd.isna(cdt):
                                    tdt = today if rv == 'today' else parser.parse(rv)
                                    if r_op == 'before' and cdt < tdt: red = True
                                    elif r_op == 'after' and cdt > tdt: red = True
                                    elif r_op == 'on' and cdt.date() == tdt.date(): red = True
                        except: pass
                
                if red: style += " color:red; font-weight:bold;"
                html += f'<td style="{style}">{disp}</td>'
            
            html += f'<td style="{c_style}"></td></tr>'

    return html + "</tbody></table>"

# -------------------------------------------------------------------------------------------
# 🛠️ PLACEHOLDER ENGINE
# -------------------------------------------------------------------------------------------

def apply_placeholders(text, name, email):
    ctx = {'Name': name, 'FullName': name, 'FirstName': name.split()[0], 'Email': email}
    for k, v in ctx.items():
        text = text.replace("{" + k + "}", str(v))
    return text

# -------------------------------------------------------------------------------------------
# 🌐 FLASK ROUTES
# -------------------------------------------------------------------------------------------

@app.route('/')
def index():
    owners = pd.read_csv(OWNERS_FILE).to_dict('records')
    templates = [f.replace('.json', '') for f in os.listdir(TEMPLATES_FOLDER)]
    return render_template('index.html', owners=owners, templates=templates)

@app.route('/scan_excel', methods=['POST'])
def scan_excel():
    try:
        f = request.files['data']
        stream = io.BytesIO(f.read())
        xl = pd.ExcelFile(stream, engine='openpyxl')
        m_col = request.form['match_col']
        found = set()
        for s in xl.sheet_names:
            df = xl.parse(s)
            if m_col in df.columns:
                found.update(df[m_col].dropna().astype(str).str.strip().unique())
        return jsonify(list(found))
    except Exception as e:
        return jsonify([])

@app.route('/save_template', methods=['POST'])
def save_tpl():
    data = request.json
    with open(f"{TEMPLATES_FOLDER}/{data['name']}.json", 'w') as f:
        json.dump(data, f)
    return jsonify({"status": "Success"})

@app.route('/load_template/<name>')
def load_tpl(name):
    with open(f"{TEMPLATES_FOLDER}/{name}.json", 'r') as f:
        return jsonify(json.load(f))

# -------------------------------------------------------------------------------------------
# 🚀 MAIN PROCESSING ROUTE (Deep-Normalization Logic Implementation)
# -------------------------------------------------------------------------------------------

@app.route('/process', methods=['POST'])
def process():
    pythoncom.CoInitialize()
    try:
        # 1. UI Configuration Capture
        tool_mode = request.form.get('tool_mode')
        broadcast_type = request.form.get('broadcast_type')
        match_col = request.form['match_col']
        sort_col = request.form.get('sort_col')
        sort_ord = request.form.get('sort_order', 'asc')
        subject_raw = request.form['subject']
        message_raw = request.form['message']
        cc_list = request.form.get('cc', '')
        
        # Recipient Deep Normalization
        recipients_ui = request.form.getlist('recipients[]')
        normalized_pool = set(normalize_text(n) for n in recipients_ui)

        # Highlight Rules
        rules = list(zip(request.form.getlist('rule_col[]'), 
                         request.form.getlist('rule_op[]'), 
                         request.form.getlist('rule_val[]')))

        # 2. Excel Handling (In-memory stream with Openpyxl engine)
        excel_file = request.files['data']
        excel_bytes = io.BytesIO(excel_file.read())
        xl_file = pd.ExcelFile(excel_bytes, engine='openpyxl')

        # Owners Dir
        owners_map = {row['Name']: row['Email'] for _, row in pd.read_csv(OWNERS_FILE).iterrows()}
        outlook = win32.Dispatch("Outlook.Application")
        zip_mem = io.BytesIO()

        with zipfile.ZipFile(zip_mem, 'w') as zf:
            
            # --- BRANCH A: MASS BROADCASTER (ONE CONSOLIDATED EMAIL) ---
            if tool_mode == 'broadcaster' and broadcast_type == 'single':
                mail = outlook.CreateItem(0)
                mail.Subject = subject_raw
                mail.CC = cc_list
                
                def mass_pool_replacer(match):
                    sheet_name = match.group(1)
                    df = xl_file.parse(sheet_name)
                    if match_col not in df.columns: return f"Error: '{match_col}' not found"
                    
                    # DEEP FILTERING: Use normalization on BOTH sides
                    col_series = df[match_col].apply(normalize_text)
                    filtered = df[col_series.isin(normalized_pool)]
                    return generate_styled_table(filtered, sort_col, sort_ord, rules)

                mail.HTMLBody = re.sub(r'\{TABLE:(.*?)\}', mass_pool_replacer, message_raw)
                mail.Save()
                abs_path = os.path.abspath("Broadcast_Group_Draft.msg")
                mail.SaveAs(abs_path); zf.write(abs_path, "Broadcast_Group_Draft.msg")

            # --- BRANCH B: INDIVIDUALIZED DRAFTS (SENDER MODE OR BROADCASTER MULTI) ---
            else:
                for person_name in recipients_ui:
                    email_addr = owners_map.get(person_name, "")
                    if tool_mode == 'sender' and not email_addr: continue
                    
                    mail = outlook.CreateItem(0)
                    mail.To = email_addr
                    mail.CC = cc_list
                    mail.Subject = apply_placeholders(subject_raw, person_name, email_addr)

                    def logic_replacer(match):
                        sheet_name = match.group(1)
                        df = xl_file.parse(sheet_name)
                        if match_col not in df.columns: return f"Error: '{match_col}' not found"
                        
                        # Apply same Deep Normalization
                        col_series = df[match_col].apply(normalize_text)
                        
                        if tool_mode == 'sender':
                            # INDIVIDUAL: Matches only THIS person
                            filtered = df[col_series == normalize_text(person_name)]
                        else:
                            # BROADCASTER MULTI: Matches the whole UI POOL
                            filtered = df[col_series.isin(normalized_pool)]
                        
                        return generate_styled_table(filtered, sort_col, sort_ord, rules)

                    msg_w_tables = re.sub(r'\{TABLE:(.*?)\}', logic_replacer, message_raw)
                    mail.HTMLBody = apply_placeholders(msg_w_tables, person_name, email_addr)
                    mail.Save()
                    
                    fn = re.sub(r'[\\/*?:\u0022<>|]', '', person_name)[:45] + ".msg"
                    fpath = os.path.abspath(fn)
                    time.sleep(0.05); mail.SaveAs(fpath); zf.write(fpath, fn)

        zip_mem.seek(0)
        return send_file(zip_mem, download_name='ION_Email_Engine_Drafts.zip', as_attachment=True)

    except Exception as e:
        print(f"Logic Fault: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        pythoncom.CoUninitialize()

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True, threaded=True)