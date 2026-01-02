import streamlit as st
import pandas as pd
import io
import os
import json
import time
import base64
import tempfile
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit.components.v1 as components
from reportlab.lib.pagesizes import A4, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter, Transformation, PageObject

# --- CONFIGURATION ---
st.set_page_config(page_title="Hike Warehouse Manager", layout="wide")

# --- DATABASE CONNECTION ---
def get_db_connection():
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("Secrets not configured!"); st.stop()
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        sheet_url = st.secrets["database"]["sheet_url"]
        return client.open_by_url(sheet_url)
    except Exception as e:
        st.error(f"Database Error: {e}"); st.stop()

# --- AUTHENTICATION ---
def load_users():
    try:
        sh = get_db_connection(); ws = sh.worksheet("Users")
        data = ws.get_all_records()
        if not data: return {"admin": "admin123"} 
        return {str(row['username']): str(row['password']) for row in data}
    except: return {"admin": "admin123"}

def save_users(users_dict):
    try:
        sh = get_db_connection(); ws = sh.worksheet("Users"); ws.clear()
        rows = [["username", "password"]] + [[u, p] for u, p in users_dict.items()]
        ws.update(rows)
    except Exception as e: st.error(f"DB Error: {e}")

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    st.title("🔒 Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Login", type="primary"):
        users = load_users()
        if u in users and users[u] == p:
            st.session_state['logged_in'] = True; st.session_state['username'] = u; st.rerun()
        else: st.error("Invalid Credentials")
    st.stop()

# --- SIDEBAR ---
with st.sidebar:
    st.write(f"👤 **{st.session_state['username']}**")
    if st.button("Logout"): st.session_state['logged_in'] = False; st.rerun()
    
    st.divider()
    st.header("🖨️ Printing Mode")
    print_mode = st.radio("Select Method:", ["Web (Kiosk/Popup)", "Local (Windows App)"], index=0)
    
    if print_mode == "Web (Kiosk/Popup)":
        st.info("Uses browser settings. For silent print, use Kiosk Shortcut.")
    else:
        st.warning("Requires 'app.py' running locally on Windows.")

# --- FILE PATHS ---
FILES_DIR = "consignment_files"
CACHE_FILE = "master_data.csv" 
if not os.path.exists(FILES_DIR): os.makedirs(FILES_DIR)

# --- PERMANENT DATA (GOOGLE SHEETS) ---
def load_history():
    try:
        sh = get_db_connection(); ws = sh.worksheet("History")
        all_rows = ws.get_all_values()
        if not all_rows or len(all_rows) < 2: return []
        history = []
        for row in all_rows[1:]:
            json_str = "".join([cell for cell in row[3:] if cell])
            if not json_str: continue
            try:
                con_obj = json.loads(json_str)
                if 'data' in con_obj: con_obj['data'] = pd.DataFrame(con_obj['data'])
                if 'original_data' in con_obj: con_obj['original_data'] = pd.DataFrame(con_obj['original_data'])
                history.append(con_obj)
            except: pass
        return history
    except: return []

def save_history(history_list):
    try:
        sh = get_db_connection(); ws = sh.worksheet("History"); ws.clear()
        rows = [["id", "date", "channel", "data_chunks"]]
        for h in history_list:
            h_copy = h.copy()
            if 'data' in h_copy and isinstance(h_copy['data'], pd.DataFrame): h_copy['data'] = h_copy['data'].to_dict('records')
            if 'original_data' in h_copy and isinstance(h_copy['original_data'], pd.DataFrame): h_copy['original_data'] = h_copy['original_data'].to_dict('records')
            full_json = json.dumps(h_copy)
            chunk_size = 40000
            chunks = [full_json[i:i+chunk_size] for i in range(0, len(full_json), chunk_size)]
            rows.append([h['id'], h['date'], h['channel']] + chunks)
        ws.update(rows)
    except Exception as e: st.error(f"Save Error: {e}")

def load_address_data(sheet_name, default_cols):
    try:
        sh = get_db_connection(); ws = sh.worksheet(sheet_name)
        data = ws.get_all_records()
        if not data: return pd.DataFrame(columns=default_cols)
        return pd.DataFrame(data).astype(str)
    except: return pd.DataFrame(columns=default_cols)

def save_address_data(sheet_name, df):
    try:
        sh = get_db_connection(); ws = sh.worksheet(sheet_name); ws.clear()
        ws.update([df.columns.values.tolist()] + df.values.tolist())
    except Exception as e: st.error(f"Save Error: {e}")

def sync_data():
    try:
        sheet_url = st.secrets["database"]["sheet_url"]
        csv_url = sheet_url.replace('/edit?gid=', '/export?format=csv&gid=').split('#')[0]
        df = pd.read_csv(csv_url, dtype={'EAN': str})
        if 'PPCN' not in df.columns: return False, "Column 'PPCN' missing."
        df.to_csv(CACHE_FILE, index=False)
        return True, "✅ Master Data Synced!"
    except Exception as e: return False, f"❌ Sync Failed: {e}"

def load_master_data():
    return pd.read_csv(CACHE_FILE, dtype={'EAN': str}) if os.path.exists(CACHE_FILE) else pd.DataFrame()

# --- PRINTING HELPERS ---
def get_merged_labels_path(c_id): return os.path.join(FILES_DIR, c_id, "merged_labels.pdf")
def save_uploaded_file(uploaded_file, c_id, file_type):
    c_dir = os.path.join(FILES_DIR, c_id); 
    if not os.path.exists(c_dir): os.makedirs(c_dir)
    file_path = os.path.join(c_dir, f"{file_type}.pdf")
    with open(file_path, "wb") as f: f.write(uploaded_file.getbuffer())
    return file_path
def get_stored_file(c_id, file_type):
    file_path = os.path.join(FILES_DIR, c_id, f"{file_type}.pdf")
    return file_path if os.path.exists(file_path) else None
def extract_box_pdf_page(merged_pdf_path, box_index):
    try:
        reader = PdfReader(merged_pdf_path); writer = PdfWriter()
        if box_index >= len(reader.pages): return None, None
        writer.add_page(reader.pages[box_index])
        output_bytes = io.BytesIO(); writer.write(output_bytes)
        return output_bytes.getvalue(), writer
    except: return None, None

# --- NEW AGGRESSIVE PRINT TRIGGER ---
def trigger_browser_print(pdf_bytes):
    """
    Forces the browser to load the PDF in a hidden iframe and print it.
    Works for both Normal (Popup) and Kiosk (Silent) modes.
    """
    base64_pdf = base64.b64encode(pdf_bytes).decode('utf-8')
    
    # 1. We create an embedded Iframe
    # 2. We use 'onload' to wait for PDF rendering
    # 3. We use a backup timeout just in case onload fails
    html_code = f"""
        <iframe id="pdf_print_frame" 
                src="data:application/pdf;base64,{base64_pdf}"
                style="position: fixed; width: 1px; height: 1px; bottom: 0; right: 0; border: none;">
        </iframe>
        <script>
            var iframe = document.getElementById('pdf_print_frame');
            
            function doPrint() {{
                try {{
                    iframe.contentWindow.focus();
                    iframe.contentWindow.print();
                }} catch(e) {{
                    console.error("Print failed: " + e);
                }}
            }}

            // Attempt 1: On Load
            iframe.onload = function() {{
                setTimeout(doPrint, 500); 
            }};
            
            // Attempt 2: Backup Timer (in case onload misses)
            setTimeout(doPrint, 1500);
        </script>
    """
    components.html(html_code, height=0, width=0)

# Local Fallback
def print_local_windows(pdf_path, printer_name):
    try:
        import win32api
        win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
        return True, "Sent to Local Printer"
    except Exception as e: return False, str(e)

def get_windows_printers():
    try:
        import win32print
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        return [p[2] for p in printers]
    except: return ["Default"]

# --- GENERATORS ---
def generate_confirm_consignment_csv(df):
    output = io.BytesIO(); sorted_df = df.sort_values(by='SKU Id'); rows = []; box_counter = 1
    for _, row in sorted_df.iterrows():
        try: num_boxes = int(row['Editable Boxes']); ppcn = int(float(row['PPCN'])) if float(row['PPCN']) > 0 else 1; fsn = row.get('FSN', '')
        except: num_boxes=0; ppcn=1; fsn=''
        nominal_val = 350 * ppcn
        for _ in range(num_boxes):
            rows.append({'BOX NUMBER': box_counter, 'BOX NAME': box_counter, 'LENGTH (cm)': 75, 'BREADTH (cm)': 55, 'HEIGHT (cm)': 40, 'WEIGHT (kg)': 10, 'NOMINAL VALUE (INR)': nominal_val, 'FSN': fsn, 'QUANTITY': ppcn}); box_counter += 1
    export_df = pd.DataFrame(rows); export_df.to_csv(output, index=False); return output.getvalue()

def generate_merged_box_labels(df, c_details, sender, receiver, flipkart_pdf_path, progress_bar=None, save_path=None):
    if not flipkart_pdf_path: return None
    with open(flipkart_pdf_path, "rb") as f: pdf_bytes = f.read()
    box_data = []; total_boxes = int(df['Editable Boxes'].sum()); current_box = 1; sorted_df = df.sort_values(by='SKU Id')
    for _, row in sorted_df.iterrows():
        try: boxes = int(row['Editable Boxes'])
        except: boxes = 0
        for _ in range(boxes):
            box_data.append({'num': current_box, 'total': total_boxes, 'sku': str(row['SKU Id']), 'qty': row['PPCN'], 'fsn': str(row.get('FSN', '')), 'id': c_details['id'], 'ch': c_details['channel']}); current_box += 1
    writer = PdfWriter(); w_a4, h_a4 = A4; half_h = h_a4/2; SHIFT_UP = 25*mm
    for i, box in enumerate(box_data):
        if progress_bar: progress_bar.progress(int((i+1)/len(box_data)*100))
        packet = io.BytesIO(); c = canvas.Canvas(packet, pagesize=A4)
        def draw_grid_table(y_top):
            row_h=10*mm; y_h=y_top; y_d=y_top-row_h; x=10*mm; x1=30*mm; x2=85*mm; x3=175*mm; xe=w_a4-10*mm
            c.setLineWidth(1); c.line(x,y_h+row_h,xe,y_h+row_h); c.line(x,y_h,xe,y_h); c.line(x,y_d,xe,y_d)
            c.line(x,y_d,x,y_h+row_h); c.line(x1,y_d,x1,y_h+row_h); c.line(x2,y_d,x2,y_h+row_h); c.line(x3,y_d,x3,y_h+row_h); c.line(xe,y_d,xe,y_h+row_h)
            c.setFont("Helvetica-Bold", 12); c.drawString(x+2*mm,y_h+3*mm,"SR NO."); c.drawString(x1+2*mm,y_h+3*mm,"FSN"); c.drawString(x2+2*mm,y_h+3*mm,"SKU ID"); c.drawString(x3+2*mm,y_h+3*mm,"QTY")
            c.setFont("Helvetica", 12); c.drawString(x+2*mm,y_d+3*mm,"1."); c.drawString(x1+2*mm,y_d+3*mm,box['fsn']); c.drawString(x2+2*mm,y_d+3*mm,box['sku'][:35])
            c.setFont("Helvetica-Bold", 14); c.drawString(x3+2*mm,y_d+3*mm,str(int(float(box['qty']))) ); return y_d
        def draw_slip(y_base):
            c.setFont("Helvetica-Bold", 30); c.drawCentredString(w_a4/2, y_base+45*mm, "PACKING SLIP"); db_y = draw_grid_table(y_base+32*mm)
            c.setFont("Helvetica-Bold", 30); c.drawCentredString(w_a4/2, db_y-5*mm, f"BOX NO.- {box['num']}         BOX NAME- {box['num']}")
        draw_slip(240*mm); c.setLineWidth(2); c.line(0, 210*mm, w_a4, 210*mm); draw_slip(155*mm); c.setLineWidth(1); c.line(0, half_h, w_a4, half_h)
        c.save(); packet.seek(0); custom_page = PdfReader(packet).pages[0]
        fk_idx = i // 2; is_top = (i%2==0); temp_reader = PdfReader(io.BytesIO(pdf_bytes))
        result_page = PageObject.create_blank_page(width=w_a4, height=h_a4); result_page.merge_page(custom_page)
        if fk_idx < len(temp_reader.pages):
            fk_page = temp_reader.pages[fk_idx]; fk_h = fk_page.mediabox.height; fk_w = fk_page.mediabox.width
            shift = (-(0.65*float(fk_h))+float(SHIFT_UP)) if is_top else (-(0.2*float(fk_h))+float(SHIFT_UP))
            op = Transformation().translate(tx=0, ty=shift); fk_page.add_transformation(op)
            if not is_top: fk_page.mediabox.lower_left=(0,0); fk_page.mediabox.upper_right=(fk_w, (0.4*float(fk_h))+float(SHIFT_UP))
            result_page.merge_page(fk_page)
        writer.add_page(result_page)
    if save_path:
        with open(save_path, "wb") as f: writer.write(f)
    out=io.BytesIO(); writer.write(out); return out.getvalue()

def generate_consignment_data_pdf(df, c_details):
    buffer = io.BytesIO(); doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm); elements = []
    elements.append(Paragraph(f"<b>Consignment ID:</b> {c_details['id']}", getSampleStyleSheet()['Heading2']))
    elements.append(Paragraph(f"<b>Pickup Date:</b> {c_details['date']}", getSampleStyleSheet()['Normal'])); elements.append(Spacer(1, 10))
    sorted_df = df.sort_values(by='SKU Id'); data = [['SKU', 'QTY', 'No. of Box']]; t_qty, t_box = 0, 0
    for _, row in sorted_df.iterrows():
        qty = int(row['Editable Qty']); box = int(row['Editable Boxes']); t_qty += qty; t_box += box
        data.append([str(row['SKU Id']), str(qty), str(box)])
    data.append(['TOTAL', str(t_qty), str(t_box)])
    table = Table(data, colWidths=[110*mm, 30*mm, 30*mm]); table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (1,0), (-1,-1), 'CENTER')]))
    elements.append(table); doc.build(elements); return buffer.getvalue()

def generate_bartender_full(df):
    try:
        output = io.BytesIO(); master_df = load_master_data()
        if master_df.empty: return None
        if 'SKU Id' not in df.columns: return None
        export_df = pd.merge(df[['SKU Id', 'Editable Qty']], master_df, left_on='SKU Id', right_on='SKU', how='left')
        export_df['QTY'] = export_df['Editable Qty']; export_df = export_df.drop(columns=['SKU Id', 'Editable Qty'], errors='ignore')
        if 'EAN' in export_df.columns: export_df['EAN'] = export_df['EAN'].astype(str).str.replace(r'\.0$', '', regex=True)
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, index=False); workbook = writer.book; worksheet = writer.sheets['Sheet1']
            if 'EAN' in export_df.columns: worksheet.set_column(export_df.columns.get_loc('EAN'), export_df.columns.get_loc('EAN'), 20, workbook.add_format({'num_format': '@'}))
        return output.getvalue()
    except: return None

def generate_excel_simple(df, cols, filename):
    output = io.BytesIO(); temp_df = df.copy(); temp_df['Qty']=temp_df.get('Editable Qty',0); temp_df['Boxes']=temp_df.get('Editable Boxes',0)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: temp_df[[c for c in cols if c in temp_df.columns]].to_excel(writer, index=False)
    return output.getvalue()

# --- APP NAVIGATION ---
def nav(page): st.session_state['page'] = page; st.rerun()
def home_button(): 
    if st.sidebar.button("🏠 Home", use_container_width=True): nav('home')

# --- INITIALIZATION ---
if 'page' not in st.session_state: st.session_state['page'] = 'home'
if 'consignments' not in st.session_state: st.session_state['consignments'] = load_history()
addr_cols = ['Code', 'Address1', 'Address2', 'City', 'State', 'Pincode', 'GST', 'Channel']

# 1. HOME
if st.session_state['page'] == 'home':
    st.title("Hike Warehouse Manager 🚀")
    if st.session_state['consignments']:
        df_hist = pd.DataFrame([{ 'Date': pd.to_datetime(c['date']), 'Boxes': int(c['data']['Editable Boxes'].sum()), 'Qty': int(c['data']['Editable Qty'].sum()), 'Channel': c['channel'] } for c in st.session_state['consignments']])
        m1, m2, m3 = st.columns(3); m1.metric("📦 Boxes", df_hist['Boxes'].sum()); m2.metric("👟 Pairs", df_hist['Qty'].sum()); m3.metric("📅 Last", df_hist['Date'].max().strftime('%d-%b-%Y') if not df_hist.empty else "N/A")
        st.divider(); c1, c2 = st.columns(2)
        with c1: st.subheader("Volume"); st.bar_chart(df_hist.groupby('Channel')['Boxes'].sum(), color="#FF4B4B")
        with c2: st.subheader("Recent"); st.dataframe(df_hist.sort_values(by='Date', ascending=False)[['Date', 'Channel', 'Boxes']].head(5), hide_index=True)
    else: st.info("No data yet.")
    st.divider(); st.subheader("Manage Channels")
    c1,c2,c3 = st.columns(3)
    if c1.button("🛒 Flipkart", use_container_width=True): st.session_state['current_channel']='Flipkart'; nav('channel')
    if c2.button("📦 Amazon", use_container_width=True): st.session_state['current_channel']='Amazon'; nav('channel')
    if c3.button("🛍️ Myntra", use_container_width=True): st.session_state['current_channel']='Myntra'; nav('channel')
    with st.sidebar:
        st.header("Settings")
        if st.button("🔄 Sync Data"):
            s, m = sync_data()
            if s: st.success(m)
            else: st.error(m)
    st.divider()
    with st.expander("📂 View History"):
        if st.session_state['consignments']:
            for c in reversed(st.session_state['consignments']):
                ca, cb = st.columns([4,1]); ca.write(f"**{c['channel']}** | {c['id']} | {c['date']}")
                if cb.button("Open", key=f"ho_{c['id']}"): st.session_state['curr_con']=c; nav('view_saved')

# 2. CHANNEL
elif st.session_state['page'] == 'channel':
    home_button(); st.title(f"{st.session_state['current_channel']}")
    cons = [c for c in st.session_state['consignments'] if c['channel'] == st.session_state['current_channel']]
    if cons:
        for c in reversed(cons[-5:]):
             if st.button(f"📄 Open {c['id']} ({c['date']})", key=f"ch_{c['id']}"): st.session_state['curr_con'] = c; nav('view_saved')
    else: st.info("No saved consignments.")
    st.divider(); 
    if st.button("➕ Create New", type="primary"): nav('add')

# 3. ADD
elif st.session_state['page'] == 'add':
    home_button(); st.title("New Consignment"); c_id = st.text_input("Consignment ID"); p_date = st.date_input("Pickup Date")
    df_s = load_address_data("Senders", addr_cols); df_r = load_address_data("Receivers", addr_cols)
    c1, c2 = st.columns(2)
    with c1:
        s_sel = st.selectbox("Sender", df_s['Code'].tolist() + ["+ Add New"])
        if s_sel == "+ Add New":
            with st.form("ns"):
                ns = {k: st.text_input(k) for k in addr_cols if k!='Channel'}; ns['Channel']='All'
                if st.form_submit_button("Save"): save_address_data("Senders", pd.concat([df_s, pd.DataFrame([ns])], ignore_index=True)); st.rerun()
    with c2:
        r_list = df_r[df_r['Channel']==st.session_state['current_channel']]['Code'].tolist() if not df_r.empty else []
        r_sel = st.selectbox("Receiver", r_list + ["+ Add New"])
        if r_sel == "+ Add New":
            with st.form("nr"):
                nr = {k: st.text_input(k) for k in addr_cols if k!='Channel'}; nr['Channel']=st.session_state['current_channel']
                if st.form_submit_button("Save"): save_address_data("Receivers", pd.concat([df_r, pd.DataFrame([nr])], ignore_index=True)); st.rerun()
    uploaded = st.file_uploader("Upload CSV", type='csv')
    if uploaded and c_id and s_sel != "+ Add New":
        if st.button("Process"):
            if c_id in [c['id'] for c in st.session_state['consignments']]: st.error("ID exists!"); st.stop()
            df_m = load_master_data(); 
            if df_m.empty: st.error("Sync Data!"); st.stop()
            df_raw = pd.read_csv(uploaded); uploaded.seek(0); df_c = pd.read_csv(uploaded)
            
            # Auto-strip whitespace
            df_c.columns = df_c.columns.str.strip()
            if 'SKU Id' in df_c.columns:
                df_c['SKU Id'] = df_c['SKU Id'].astype(str).str.strip()
            
            # --- VALIDATION ---
            csv_skus = set(df_c['SKU Id'].astype(str))
            master_skus = set(df_m['SKU'].astype(str))
            missing = [s for s in csv_skus if s not in master_skus]
            
            if missing:
                st.error("🚨 STOP! Found SKUs in file that are NOT in Master Data:")
                st.dataframe(pd.DataFrame(missing, columns=["Missing SKU IDs"]), use_container_width=True)
                st.error("Please add these to Master Data and Sync before proceeding.")
                st.stop()

            merged = pd.merge(df_c, df_m, left_on='SKU Id', right_on='SKU', how='left')
            merged['Editable Qty'] = merged['Quantity Sent'].fillna(0); merged['PPCN'] = pd.to_numeric(merged['PPCN'], errors='coerce').fillna(1)
            merged['Editable Boxes'] = (merged['Editable Qty'] / merged['PPCN']).apply(lambda x: float(x)).round(2)
            st.session_state['curr_con'] = {'id': c_id, 'date': str(p_date), 'channel': st.session_state['current_channel'], 'data': merged, 'original_data': df_raw, 'sender': df_s[df_s['Code']==s_sel].iloc[0].to_dict(), 'receiver': df_r[df_r['Code']==r_sel].iloc[0].to_dict(), 'saved': False, 'printed_boxes': []}
            nav('preview')

# 4. PREVIEW
elif st.session_state['page'] == 'preview':
    home_button(); pkg = st.session_state['curr_con']; st.title(f"Review: {pkg['id']}")
    disp = pkg['data'][['SKU Id', 'Product Name', 'Editable Qty', 'Editable Boxes']].copy(); disp['Editable Boxes'] = disp['Editable Boxes'].astype(int)
    st.dataframe(disp, hide_index=True, use_container_width=True)
    if st.button("💾 SAVE", type="primary"): pkg['saved'] = True; st.session_state['consignments'].append(pkg); save_history(st.session_state['consignments']); nav('view_saved')

# 5. SCAN & PRINT PAGE
elif st.session_state['page'] == 'scan_print':
    pkg = st.session_state['curr_con']; c_id = pkg['id']; merged_pdf_path = get_merged_labels_path(c_id)
    if 'scan_box_data' not in st.session_state or st.session_state.get('scan_c_id') != c_id:
        box_data = []; current_box = 1; sorted_df = pkg['data'].sort_values(by='SKU Id')
        for _, row in sorted_df.iterrows():
            try: boxes = int(row['Editable Boxes'])
            except: boxes = 0
            for _ in range(boxes):
                box_data.append({'Box No': current_box, 'SKU': str(row['SKU Id']), 'FSN': str(row.get('FSN', '')), 'EAN': str(row.get('EAN', '')).replace('.0',''), 'Qty': int(row['PPCN'])})
                current_box += 1
        st.session_state['scan_box_data'] = pd.DataFrame(box_data); st.session_state['scan_c_id'] = c_id; st.session_state['last_printed_box'] = None

    df_boxes = st.session_state['scan_box_data']

    def handle_print_action(box_num):
        bytes_data, writer = extract_box_pdf_page(merged_pdf_path, int(box_num)-1)
        if not bytes_data: st.error("PDF Error"); return False
        
        if print_mode == "Local (Direct USB)":
            if not HAS_WIN32: st.error("Local Mode requires Windows + pywin32"); return False
            try:
                import win32api
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    writer.write(tmp); tmp_path = tmp.name
                printer = st.session_state.get('selected_printer_local')
                if printer:
                    win32api.ShellExecute(0, "printto", tmp_path, f'"{printer}"', ".", 0)
                    time.sleep(1); 
                    try: os.remove(tmp_path) 
                    except: pass
                    return True
                else: st.warning("Select local printer!"); return False
            except ImportError: st.error("Pywin32 not installed"); return False
        else:
            st.session_state['web_print_trigger'] = bytes_data
            return True

    def process_scan():
        scan_val = st.session_state.scan_input.strip(); 
        if not scan_val: return
        matches = df_boxes[(df_boxes['SKU'] == scan_val) | (df_boxes['FSN'] == scan_val) | (df_boxes['EAN'] == scan_val)]
        if matches.empty: st.toast(f"❌ Not found: {scan_val}", icon="⚠️")
        else:
            printed_set = set(pkg.get('printed_boxes', []))
            valid_boxes = matches[~matches['Box No'].isin(printed_set)]
            if valid_boxes.empty: st.toast(f"✅ All boxes printed!", icon="ℹ️")
            else:
                target_box = valid_boxes.iloc[0]['Box No']
                if handle_print_action(target_box):
                    st.session_state['last_printed_box'] = int(target_box)
                    if 'printed_boxes' not in pkg: pkg['printed_boxes'] = []
                    pkg['printed_boxes'].append(int(target_box))
                    save_history(st.session_state['consignments'])
                    st.toast(f"🖨️ Printing Box {target_box}...", icon="✅")
        st.session_state.scan_input = ""

    def trigger_reprint_manual(box_num):
        if handle_print_action(box_num): st.session_state['last_printed_box'] = int(box_num); st.toast(f"🖨️ Re-printing Box {box_num}...", icon="✅")

    c_back, c_spacer, c_pr = st.columns([1, 4, 2])
    with c_back: 
        if st.button("🔙 Back", use_container_width=True): nav('view_saved')
    
    with c_pr:
        if print_mode == "Local (Direct USB)":
            try:
                import win32print
                printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            except: printers = ["Default"]
            if 'selected_printer_local' not in st.session_state: st.session_state['selected_printer_local'] = printers[0] if printers else None
            st.selectbox("Select Printer", printers, key='selected_printer_local', label_visibility="collapsed")
        elif print_mode == "Web (Kiosk/Popup)":
            st.caption("🌐 Using Browser Print Trigger")

    st.divider(); st.text_input("SCAN BARCODE", key='scan_input', on_change=process_scan)

    if 'web_print_trigger' in st.session_state:
        trigger_browser_print(st.session_state['web_print_trigger']); del st.session_state['web_print_trigger']

    last_p = st.session_state.get('last_printed_box')
    if last_p: st.info(f"🖨️ Sent to Printer: **BOX {last_p}**", icon="✨")

    printed_set = set(pkg.get('printed_boxes', [])); display_df = df_boxes.copy()
    display_df['Status'] = display_df['Box No'].apply(lambda x: '✅ PRINTED' if x in printed_set else 'WAITING')
    def highlight_rows(row):
        if row['Box No'] == st.session_state.get('last_printed_box'): return ['background-color: #fff3cd'] * len(row)
        elif row['Status'] == '✅ PRINTED': return ['background-color: #d4edda'] * len(row)
        return [''] * len(row)
    
    st.subheader("Box List"); event = st.dataframe(display_df.style.apply(highlight_rows, axis=1), use_container_width=True, hide_index=True, height=500, on_select="rerun", selection_mode="single-row")
    if event.selection.rows:
        sel_box = display_df.iloc[event.selection.rows[0]]['Box No']
        if st.button(f"🖨️ Reprint Box {sel_box}", type="primary", use_container_width=True): trigger_reprint_manual(sel_box)

# 6. VIEW SAVED
elif st.session_state['page'] == 'view_saved':
    home_button(); pkg = st.session_state['curr_con']; c_id = pkg['id']; st.title(f"Consignment: {c_id}")
    
    # SAFE DOWNLOAD BUTTONS (Wrapped in Try-Except to prevent crash)
    r1, r2, r3 = st.columns(3)
    with r1: csv_b=io.BytesIO(); pkg['original_data'].to_csv(csv_b, index=False); st.download_button("⬇ CSV", csv_b.getvalue(), f"{c_id}.csv")
    with r2: st.download_button("⬇ Data PDF", generate_consignment_data_pdf(pkg['data'], pkg), f"Data_{c_id}.pdf")
    with r3: st.download_button("⬇ Confirm CSV", generate_confirm_consignment_csv(pkg['data']), f"Confirm_{c_id}.csv")
    
    r4, r5 = st.columns(2)
    with r4: 
        try:
            bt_data = generate_bartender_full(pkg['data'])
            if bt_data: st.download_button("⬇ Bartender", bt_data, f"Bartender_{c_id}.xlsx")
            else: st.warning("Bartender: Data error or missing SKU Id")
        except Exception as e: st.error(f"Error: {e}")
    with r5: st.download_button("⬇ Ewaybill", generate_excel_simple(pkg['data'], ['SKU Id', 'Editable Qty', 'Cost Price'], f"Eway_{c_id}.xlsx"), f"Eway_{c_id}.xlsx")
    
    st.divider(); st.subheader("Labels & Printing")
    uc1, uc2 = st.columns([1, 1])
    with uc1:
        f_lbl = st.file_uploader("Upload Labels PDF", type=['pdf'], key='u_lbl')
        if f_lbl and st.button("Merge Labels"):
            save_uploaded_file(f_lbl, c_id, 'box_labels'); p_bar = st.progress(0)
            try: generate_merged_box_labels(pkg['data'], pkg, pkg['sender'], pkg['receiver'], get_stored_file(c_id, 'box_labels'), p_bar, get_merged_labels_path(c_id)); st.success("Merged!"); st.rerun()
            except Exception as e: st.error(str(e))
    with uc2:
        if os.path.exists(get_merged_labels_path(c_id)):
            with open(get_merged_labels_path(c_id), "rb") as f: st.download_button("⬇ Download Merged PDF", f, f"Merged_{c_id}.pdf")
            st.divider(); 
            if st.button("🖨️ SCAN & PRINT", type="primary", use_container_width=True): nav('scan_print')
    st.divider(); 
    with st.expander("Danger Zone"):
        if st.button(f"Delete {c_id}"): 
            st.session_state['consignments']=[c for c in st.session_state['consignments'] if c['id']!=c_id]
            save_history(st.session_state['consignments']); nav('home')
