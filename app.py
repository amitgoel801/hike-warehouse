import streamlit as st
import pandas as pd
import io
import os
import json
import time
import tempfile
from reportlab.lib.pagesizes import A4, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from pypdf import PdfReader, PdfWriter, Transformation, PageObject

# --- WINDOWS PRINTING IMPORTS ---
try:
    import win32print
    import win32api
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# --- CONFIGURATION ---
st.set_page_config(page_title="Hike Warehouse Manager", layout="wide")

# --- FILE PATHS ---
CACHE_FILE = "master_data.csv"
HISTORY_FILE = "consignment_history.json"
SENDERS_FILE = "senders.xlsx"
RECEIVERS_FILE = "receivers.xlsx"
FILES_DIR = "consignment_files"
SHEET_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRdLEddTZgmuUSswPp3A_HM7DGH8UCUWEmqd-cIbbJ7nb_Eq4YvZxO0vjWESlxX-9Y6VWRcVLPFlIVp/pub?gid=0&single=true&output=csv"

if not os.path.exists(FILES_DIR): os.makedirs(FILES_DIR)

# --- PRINTER HELPERS ---
def get_printers():
    if not HAS_WIN32: return ["Error: pywin32 not installed"]
    try:
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        return [p[2] for p in printers]
    except: return ["Default Printer"]

def send_pdf_to_printer(pdf_path, printer_name):
    if not HAS_WIN32:
        st.error("Cannot print: pywin32 library missing.")
        return False
    try:
        # "printto" sends the file to the specified printer
        win32api.ShellExecute(0, "printto", pdf_path, f'"{printer_name}"', ".", 0)
        return True
    except Exception as e:
        return False, str(e)

def extract_and_print_box(merged_pdf_path, box_index, printer_name):
    """
    Extracts a specific page (box label) from the merged PDF and sends it to the printer.
    box_index is 0-based (Box 1 = index 0).
    """
    try:
        reader = PdfReader(merged_pdf_path)
        writer = PdfWriter()
        
        # Security check
        if box_index >= len(reader.pages):
            return False, "Box Index out of range"

        writer.add_page(reader.pages[box_index])
        
        # Create a temp file for the single page
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            writer.write(tmp)
            tmp_path = tmp.name
        
        # Print
        success = send_pdf_to_printer(tmp_path, printer_name)
        
        # Cleanup (Wait a bit for spooler to grab it)
        time.sleep(2) 
        try: os.remove(tmp_path)
        except: pass
        
        if success is False: return False, "Unknown Error"
        return True, "Sent to Printer"
    except Exception as e:
        return False, str(e)

# --- DATA MANAGERS ---
def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, 'r') as f:
            try:
                history = json.load(f)
                for h in history:
                    if 'data' in h: h['data'] = pd.DataFrame(h['data'])
                    if 'original_data' in h: h['original_data'] = pd.DataFrame(h['original_data'])
                    if 'printed_boxes' not in h: h['printed_boxes'] = [] 
                return history
            except: return []
    return []

def save_history(history_list):
    serializable_list = []
    for h in history_list:
        h_copy = h.copy()
        if 'data' in h_copy and isinstance(h_copy['data'], pd.DataFrame):
            h_copy['data'] = h_copy['data'].to_dict('records')
        if 'original_data' in h_copy and isinstance(h_copy['original_data'], pd.DataFrame):
            h_copy['original_data'] = h_copy['original_data'].to_dict('records')
        serializable_list.append(h_copy)
    with open(HISTORY_FILE, 'w') as f: json.dump(serializable_list, f)

def load_address_data(file_path, default_cols):
    if os.path.exists(file_path): return pd.read_excel(file_path, dtype=str)
    return pd.DataFrame(columns=default_cols)

def save_address_data(file_path, df): df.to_excel(file_path, index=False)

def sync_data():
    try:
        df = pd.read_csv(SHEET_URL, dtype={'EAN': str})
        if 'PPCN' not in df.columns: return False, "Column 'PPCN' missing."
        df.to_csv(CACHE_FILE, index=False)
        return True, "✅ Master Data Synced!"
    except Exception as e: return False, f"❌ Sync Failed: {e}"

def load_master_data():
    return pd.read_csv(CACHE_FILE, dtype={'EAN': str}) if os.path.exists(CACHE_FILE) else pd.DataFrame()

# --- FILE HELPERS ---
def save_uploaded_file(uploaded_file, c_id, file_type):
    c_dir = os.path.join(FILES_DIR, c_id)
    if not os.path.exists(c_dir): os.makedirs(c_dir)
    file_path = os.path.join(c_dir, f"{file_type}.pdf")
    with open(file_path, "wb") as f: f.write(uploaded_file.getbuffer())
    return file_path

def get_stored_file(c_id, file_type):
    file_path = os.path.join(FILES_DIR, c_id, f"{file_type}.pdf")
    return file_path if os.path.exists(file_path) else None

def get_merged_labels_path(c_id):
    return os.path.join(FILES_DIR, c_id, "merged_labels.pdf")

# --- GENERATORS (PDFs etc) ---
def generate_confirm_consignment_csv(df):
    output = io.BytesIO()
    sorted_df = df.sort_values(by='SKU Id')
    rows = []
    box_counter = 1
    for _, row in sorted_df.iterrows():
        try: 
            num_boxes = int(row['Editable Boxes'])
            ppcn = int(float(row['PPCN'])) if float(row['PPCN']) > 0 else 1
            fsn = row.get('FSN', '')
        except: num_boxes=0; ppcn=1; fsn=''
        nominal_val = 350 * ppcn
        for _ in range(num_boxes):
            rows.append({
                'BOX NUMBER': box_counter, 'BOX NAME': box_counter,
                'LENGTH (cm)': 75, 'BREADTH (cm)': 55, 'HEIGHT (cm)': 40, 'WEIGHT (kg)': 10,
                'NOMINAL VALUE (INR)': nominal_val, 'FSN': fsn, 'QUANTITY': ppcn
            })
            box_counter += 1
    export_df = pd.DataFrame(rows)
    export_df.to_csv(output, index=False)
    return output.getvalue()

def generate_merged_box_labels(df, c_details, sender, receiver, flipkart_pdf_path, progress_bar=None, save_path=None):
    if not flipkart_pdf_path: return None
    with open(flipkart_pdf_path, "rb") as f: pdf_bytes = f.read()
    
    box_data = []
    total_boxes = int(df['Editable Boxes'].sum())
    current_box = 1
    sorted_df = df.sort_values(by='SKU Id')
    for _, row in sorted_df.iterrows():
        try: boxes = int(row['Editable Boxes'])
        except: boxes = 0
        for _ in range(boxes):
            box_data.append({
                'num': current_box, 'total': total_boxes,
                'sku': str(row['SKU Id']), 'qty': row['PPCN'],
                'fsn': str(row.get('FSN', '')),
                'id': c_details['id'], 'ch': c_details['channel']
            })
            current_box += 1

    writer = PdfWriter()
    w_a4, h_a4 = A4 
    half_h = h_a4 / 2
    SHIFT_UP = 25 * mm
    total_items = len(box_data)
    
    for i, box in enumerate(box_data):
        if progress_bar:
            progress = int((i + 1) / total_items * 100)
            progress_bar.progress(progress, text=f"Processing Box {i+1} of {total_items}...")
        
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=A4)
        
        def draw_grid_table(y_top):
            row_h = 10*mm; y_header = y_top; y_data = y_top - row_h
            x_start = 10*mm; x_c1 = 30*mm; x_c2 = 85*mm; x_c3 = 175*mm; x_end = w_a4 - 10*mm 
            c.setLineWidth(1)
            c.line(x_start, y_header + row_h, x_end, y_header + row_h) 
            c.line(x_start, y_header, x_end, y_header) 
            c.line(x_start, y_data, x_end, y_data) 
            c.line(x_start, y_data, x_start, y_header + row_h)
            c.line(x_c1, y_data, x_c1, y_header + row_h)
            c.line(x_c2, y_data, x_c2, y_header + row_h)
            c.line(x_c3, y_data, x_c3, y_header + row_h)
            c.line(x_end, y_data, x_end, y_header + row_h)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x_start + 2*mm, y_header + 3*mm, "SR NO."); c.drawString(x_c1 + 2*mm, y_header + 3*mm, "FSN"); c.drawString(x_c2 + 2*mm, y_header + 3*mm, "SKU ID"); c.drawString(x_c3 + 2*mm, y_header + 3*mm, "QTY")
            c.setFont("Helvetica", 12)
            c.drawString(x_start + 2*mm, y_data + 3*mm, "1."); c.drawString(x_c1 + 2*mm, y_data + 3*mm, box['fsn']) 
            c.setFont("Helvetica", 12); c.drawString(x_c2 + 2*mm, y_data + 3*mm, box['sku'][:35]) 
            c.setFont("Helvetica-Bold", 14); c.drawString(x_c3 + 2*mm, y_data + 3*mm, str(int(float(box['qty'])))) 
            return y_data 

        def draw_slip(y_base):
            c.setFont("Helvetica-Bold", 30); c.drawCentredString(w_a4/2, y_base + 45*mm, "PACKING SLIP")
            data_bottom_y = draw_grid_table(y_base + 32*mm)
            c.setFont("Helvetica-Bold", 30); box_txt = f"BOX NO.- {box['num']}         BOX NAME- {box['num']}"
            c.drawCentredString(w_a4/2, data_bottom_y - 5*mm, box_txt)

        draw_slip(240*mm)
        c.setLineWidth(2); c.line(0, 210*mm, w_a4, 210*mm)
        draw_slip(155*mm)
        c.setLineWidth(1); c.line(0, half_h, w_a4, half_h)
        c.save(); packet.seek(0)
        custom_page = PdfReader(packet).pages[0]
        
        fk_page_idx = i // 2
        is_top_label = (i % 2 == 0)
        temp_reader = PdfReader(io.BytesIO(pdf_bytes))
        num_fk_pages = len(temp_reader.pages)
        result_page = PageObject.create_blank_page(width=w_a4, height=h_a4)
        result_page.merge_page(custom_page)
        
        if fk_page_idx < num_fk_pages:
            fk_page = temp_reader.pages[fk_page_idx]
            fk_h = fk_page.mediabox.height; fk_w = fk_page.mediabox.width
            if is_top_label:
                shift_amount = -(0.65 * float(fk_h)) + float(SHIFT_UP)
                op = Transformation().translate(tx=0, ty=shift_amount)
                fk_page.add_transformation(op)
            else:
                shift_amount = -(0.2 * float(fk_h)) + float(SHIFT_UP)
                op = Transformation().translate(tx=0, ty=shift_amount)
                fk_page.add_transformation(op)
                fk_page.mediabox.lower_left = (0, 0)
                fk_page.mediabox.upper_right = (fk_w, (0.4 * float(fk_h)) + float(SHIFT_UP))
            result_page.merge_page(fk_page)
        writer.add_page(result_page)

    if save_path:
        with open(save_path, "wb") as f: writer.write(f)
    output = io.BytesIO(); writer.write(output)
    return output.getvalue()

def generate_consignment_data_pdf(df, c_details):
    buffer = io.BytesIO(); doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=10*mm, leftMargin=10*mm, topMargin=10*mm, bottomMargin=10*mm); elements = []; styles = getSampleStyleSheet()
    elements.append(Paragraph(f"<b>Consignment ID:</b> {c_details['id']}", styles['Heading2']))
    elements.append(Paragraph(f"<b>Pickup Date:</b> {c_details['date']}", styles['Normal']))
    elements.append(Spacer(1, 10))
    sorted_df = df.sort_values(by='SKU Id')
    data = [['SKU', 'QTY', 'No. of Box']]; t_qty, t_box = 0, 0
    for _, row in sorted_df.iterrows():
        qty = int(row['Editable Qty']); box = int(row['Editable Boxes']); t_qty += qty; t_box += box
        data.append([str(row['SKU Id']), str(qty), str(box)])
    data.append(['TOTAL', str(t_qty), str(t_box)])
    table = Table(data, colWidths=[110*mm, 30*mm, 30*mm])
    table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (1,0), (-1,-1), 'CENTER'), ('ALIGN', (0,0), (0,-1), 'LEFT'), ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'), ('BACKGROUND', (0,-1), (-1,-1), colors.whitesmoke)]))
    elements.append(table)
    doc.build(elements); return buffer.getvalue()

def generate_challan(df, c_details, sender, receiver):
    buffer = io.BytesIO(); c = canvas.Canvas(buffer, pagesize=A4); w, h = A4
    c.setFont("Helvetica-Bold", 14); c.drawString(10*mm, h-15*mm, f"Consignment ID: {c_details['id']}")
    c.setFont("Helvetica-Bold", 20); c.drawCentredString(w/2, h-25*mm, "DELIVERY CHALLAN")
    c.setLineWidth(1); c.rect(10*mm, h-85*mm, w-20*mm, 50*mm)
    def draw_addr(x, y, data, lbl):
        c.setFont("Helvetica-Bold", 10); c.drawString(x, y, lbl); c.drawString(x, y-5*mm, str(data.get('Code','')))
        c.setFont("Helvetica", 10); c.drawString(x, y-10*mm, str(data.get('Address1',''))); c.drawString(x, y-15*mm, f"{data.get('City','')}, {data.get('State','')}"); c.drawString(x, y-20*mm, f"GST: {data.get('GST','')}")
    draw_addr(15*mm, h-40*mm, sender, "FROM:"); draw_addr(110*mm, h-40*mm, receiver, "TO:")
    c.drawString(15*mm, h-95*mm, f"Date: {c_details['date']}")
    data = [['S.No', 'SKU', 'Product', 'Qty', 'Boxes']]; 
    for i, row in df.iterrows(): data.append([str(i+1), str(row['SKU Id']), str(row.get('Product Name',''))[:25], str(int(row['Editable Qty'])), str(int(row['Editable Boxes']))])
    table = Table(data, colWidths=[15*mm, 60*mm, 70*mm, 20*mm, 20*mm])
    table.setStyle(TableStyle([('GRID', (0,0), (-1,-1), 1, colors.black), ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('ALIGN', (0,0), (-1,-1), 'CENTER')]))
    table.wrapOn(c, w, h); table.drawOn(c, 10*mm, h-110*mm - (len(data)*7*mm))
    c.save()
    return buffer.getvalue()

def generate_appointment_letter(c_details, sender, receiver):
    buffer = io.BytesIO(); c = canvas.Canvas(buffer, pagesize=A4); w, h = A4
    c.setFont("Helvetica-Bold", 20); c.drawCentredString(w/2, h-30*mm, "APPOINTMENT LETTER")
    c.setFont("Helvetica", 12)
    c.drawString(20*mm, h-60*mm, f"Date: {c_details['date']}")
    c.drawString(20*mm, h-70*mm, f"To: {receiver.get('Code')} ({receiver.get('City')})")
    c.drawString(20*mm, h-80*mm, f"From: {sender.get('Code')} ({sender.get('City')})")
    c.drawString(20*mm, h-100*mm, f"Subject: Delivery Appointment for Consignment {c_details['id']}")
    c.drawString(20*mm, h-120*mm, "Dear Team,"); c.drawString(20*mm, h-130*mm, "Please accept the delivery of the mentioned consignment.")
    c.drawString(20*mm, h-150*mm, "Vehicle No: _________________"); c.drawString(20*mm, h-160*mm, "Driver Name: ________________")
    c.save()
    return buffer.getvalue()

def generate_excel_simple(df, cols, filename):
    output = io.BytesIO(); valid_cols = [c for c in cols if c in df.columns]
    temp_df = df.copy()
    if 'Qty' in cols and 'Qty' not in temp_df.columns: temp_df['Qty'] = temp_df['Editable Qty']
    if 'Boxes' in cols and 'Boxes' not in temp_df.columns: temp_df['Boxes'] = temp_df['Editable Boxes']
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer: temp_df[valid_cols].to_excel(writer, index=False)
    return output.getvalue()

def generate_bartender_full(df):
    output = io.BytesIO(); master_df = load_master_data()
    export_df = pd.merge(df[['SKU Id', 'Editable Qty']], master_df, left_on='SKU Id', right_on='SKU', how='left')
    export_df['QTY'] = export_df['Editable Qty']
    if 'SKU Id' in export_df.columns: export_df = export_df.drop(columns=['SKU Id'])
    if 'Editable Qty' in export_df.columns: export_df = export_df.drop(columns=['Editable Qty'])
    if 'EAN' in export_df.columns: export_df['EAN'] = export_df['EAN'].astype(str).str.replace(r'\.0$', '', regex=True)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        export_df.to_excel(writer, index=False)
        workbook = writer.book; worksheet = writer.sheets['Sheet1']
        text_format = workbook.add_format({'num_format': '@'})
        if 'EAN' in export_df.columns:
            col_idx = export_df.columns.get_loc('EAN')
            worksheet.set_column(col_idx, col_idx, 20, text_format)
    return output.getvalue()

# --- APP NAVIGATION ---
def nav(page): st.session_state['page'] = page; st.rerun()
def home_button(): 
    if st.sidebar.button("🏠 Home", use_container_width=True): nav('home')

# --- INITIALIZATION ---
if 'page' not in st.session_state: st.session_state['page'] = 'home'
if 'consignments' not in st.session_state: st.session_state['consignments'] = load_history()
addr_cols = ['Code', 'Address1', 'Address2', 'City', 'State', 'Pincode', 'GST', 'Channel']
if not os.path.exists(SENDERS_FILE): save_address_data(SENDERS_FILE, pd.DataFrame([{'Code': 'MAIN', 'Address1': 'Addr', 'City': 'City', 'Channel': 'All'}]))
if not os.path.exists(RECEIVERS_FILE): save_address_data(RECEIVERS_FILE, pd.DataFrame(columns=addr_cols))

# 1. HOME
if st.session_state['page'] == 'home':
    st.title("Hike Warehouse Manager 🚀")
    
    # --- DASHBOARD LOGIC ---
    if st.session_state['consignments']:
        df_hist = pd.DataFrame([
            {
                'Date': pd.to_datetime(c['date']),
                'Boxes': int(c['data']['Editable Boxes'].sum()),
                'Qty': int(c['data']['Editable Qty'].sum()),
                'Channel': c['channel']
            }
            for c in st.session_state['consignments']
        ])
        
        # Metrics
        total_boxes = df_hist['Boxes'].sum()
        total_qty = df_hist['Qty'].sum()
        total_cons = len(df_hist)
        
        # Layout
        m1, m2, m3 = st.columns(3)
        m1.metric("📦 Total Boxes Sent", total_boxes)
        m1.caption(f"Across {total_cons} consignments")
        m2.metric("👟 Total Pairs/Qty", total_qty)
        m3.metric("📅 Last Shipment", df_hist['Date'].max().strftime('%d-%b-%Y') if not df_hist.empty else "N/A")
        
        st.divider()
        
        # Charts
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Volume by Channel")
            if not df_hist.empty:
                st.bar_chart(df_hist.groupby('Channel')['Boxes'].sum(), color="#FF4B4B")
        
        with c2:
            st.subheader("Recent Activity")
            st.dataframe(
                df_hist.sort_values(by='Date', ascending=False)[['Date', 'Channel', 'Boxes', 'Qty']].head(5),
                hide_index=True,
                use_container_width=True
            )
            
    else:
        st.info("No consignment data found. Create your first consignment to see analytics!")

    st.divider()

    # --- ACTION BUTTONS ---
    st.subheader("Manage Channels")
    c1,c2,c3 = st.columns(3)
    if c1.button("🛒 Flipkart", use_container_width=True): st.session_state['current_channel']='Flipkart'; nav('channel')
    if c2.button("📦 Amazon", use_container_width=True): st.session_state['current_channel']='Amazon'; nav('channel')
    if c3.button("🛍️ Myntra", use_container_width=True): st.session_state['current_channel']='Myntra'; nav('channel')
    
    with st.sidebar:
        st.header("Settings")
        if st.button("🔄 Sync Master Data"):
            s, m = sync_data()
            if s: st.success(m)
            else: st.error(m)
        
        st.divider()
        st.caption(f"📦 History Size: {len(st.session_state['consignments'])} Records")
        if not HAS_WIN32:
            st.error("⚠️ pywin32 not detected. Scanning will work, but Printing will fail.")

    # Recent Consignments List (Collapsible)
    st.divider()
    with st.expander("📂 View Full Consignment History"):
        if st.session_state['consignments']:
            for c in reversed(st.session_state['consignments']):
                col_a, col_b = st.columns([4, 1])
                with col_a:
                    st.write(f"**{c['channel']}** | {c['id']} | {c['date']}")
                with col_b:
                    if st.button("Open", key=f"home_open_{c['id']}"): 
                        st.session_state['curr_con'] = c; nav('view_saved')

# 2. CHANNEL
elif st.session_state['page'] == 'channel':
    home_button(); st.title(f"{st.session_state['current_channel']}")
    st.subheader(f"Saved {st.session_state['current_channel']} Consignments")
    channel_cons = [c for c in st.session_state['consignments'] if c['channel'] == st.session_state['current_channel']]
    if channel_cons:
        for c in reversed(channel_cons[-5:]):
             if st.button(f"📄 Open {c['id']} ({c['date']})", key=f"ch_{c['id']}"):
                st.session_state['curr_con'] = c; nav('view_saved')
    else: st.info("No saved consignments for this channel yet.")
    st.divider()
    if st.button("➕ Create New Consignment", type="primary"): nav('add')

# 3. ADD
elif st.session_state['page'] == 'add':
    home_button(); st.title("New Consignment")
    c_id = st.text_input("Consignment ID")
    p_date = st.date_input("Pickup Date")
    df_s = load_address_data(SENDERS_FILE, addr_cols); df_r = load_address_data(RECEIVERS_FILE, addr_cols)
    c1, c2 = st.columns(2)
    with c1:
        s_sel = st.selectbox("Sender", df_s['Code'].tolist() + ["+ Add New"])
        if s_sel == "+ Add New":
            with st.form("ns"):
                ns = {k: st.text_input(k) for k in addr_cols if k!='Channel'}; ns['Channel']='All'
                if st.form_submit_button("Save"): save_address_data(SENDERS_FILE, pd.concat([df_s, pd.DataFrame([ns])], ignore_index=True)); st.rerun()
    with c2:
        r_list = df_r[df_r['Channel']==st.session_state['current_channel']]['Code'].tolist()
        r_sel = st.selectbox("Receiver", r_list + ["+ Add New"])
        if r_sel == "+ Add New":
            with st.form("nr"):
                nr = {k: st.text_input(k) for k in addr_cols if k!='Channel'}; nr['Channel']=st.session_state['current_channel']
                if st.form_submit_button("Save"): save_address_data(RECEIVERS_FILE, pd.concat([df_r, pd.DataFrame([nr])], ignore_index=True)); st.rerun()

    uploaded = st.file_uploader("Upload CSV", type='csv')
    if uploaded and c_id and s_sel != "+ Add New":
        if st.button("Process"):
            existing_ids = [c['id'] for c in st.session_state['consignments']]
            if c_id in existing_ids: st.error(f"⚠️ Consignment ID '{c_id}' already created!"); st.stop()
            df_m = load_master_data()
            if df_m.empty: st.error("Sync Data!"); st.stop()
            df_raw = pd.read_csv(uploaded); uploaded.seek(0); df_c = pd.read_csv(uploaded)
            merged = pd.merge(df_c, df_m, left_on='SKU Id', right_on='SKU', how='left')
            merged['Editable Qty'] = merged['Quantity Sent'].fillna(0)
            merged['PPCN'] = pd.to_numeric(merged['PPCN'], errors='coerce').fillna(1)
            merged['Editable Boxes'] = (merged['Editable Qty'] / merged['PPCN']).apply(lambda x: float(x)).round(2)
            st.session_state['curr_con'] = {'id': c_id, 'date': str(p_date), 'channel': st.session_state['current_channel'], 'data': merged, 'original_data': df_raw, 'sender': df_s[df_s['Code']==s_sel].iloc[0].to_dict(), 'receiver': df_r[df_r['Code']==r_sel].iloc[0].to_dict(), 'saved': False, 'printed_boxes': []}
            nav('preview')

# 4. PREVIEW
elif st.session_state['page'] == 'preview':
    home_button(); pkg = st.session_state['curr_con']; st.title(f"Review: {pkg['id']}")
    disp = pkg['data'][['SKU Id', 'Product Name', 'Editable Qty', 'Editable Boxes']].copy()
    disp['Editable Boxes'] = disp['Editable Boxes'].astype(int)
    st.dataframe(disp, hide_index=True, use_container_width=True)
    if st.button("💾 SAVE CONSIGNMENT", type="primary"):
        pkg['saved'] = True; st.session_state['consignments'].append(pkg); save_history(st.session_state['consignments']); nav('view_saved')

# 5. SCAN & PRINT PAGE
elif st.session_state['page'] == 'scan_print':
    # --- SETUP & LOGIC ---
    pkg = st.session_state['curr_con']
    c_id = pkg['id']
    merged_pdf_path = get_merged_labels_path(c_id)

    # 1. Expand Box Data for the Table
    if 'scan_box_data' not in st.session_state or st.session_state.get('scan_c_id') != c_id:
        box_data = []
        current_box = 1
        sorted_df = pkg['data'].sort_values(by='SKU Id')
        for _, row in sorted_df.iterrows():
            try: boxes = int(row['Editable Boxes'])
            except: boxes = 0
            for _ in range(boxes):
                box_data.append({
                    'Box No': current_box,
                    'SKU': str(row['SKU Id']),
                    'FSN': str(row.get('FSN', '')),
                    'EAN': str(row.get('EAN', '')).replace('.0',''),
                    'Qty': int(row['PPCN'])
                })
                current_box += 1
        st.session_state['scan_box_data'] = pd.DataFrame(box_data)
        st.session_state['scan_c_id'] = c_id
        st.session_state['last_printed_box'] = None 

    df_boxes = st.session_state['scan_box_data']
    
    # 2. Logic for Printing/Scanning
    def process_scan():
        scan_val = st.session_state.scan_input.strip()
        if not scan_val: return
        
        matches = df_boxes[
            (df_boxes['SKU'] == scan_val) | 
            (df_boxes['FSN'] == scan_val) | 
            (df_boxes['EAN'] == scan_val)
        ]

        if matches.empty:
            st.toast(f"❌ Product not found: {scan_val}", icon="⚠️")
        else:
            # FIFO Logic
            printed_set = set(pkg.get('printed_boxes', []))
            valid_boxes = matches[~matches['Box No'].isin(printed_set)]
            
            if valid_boxes.empty:
                st.toast(f"✅ All boxes for {scan_val} already printed!", icon="ℹ️")
            else:
                target_box = valid_boxes.iloc[0]['Box No']
                
                # PRINTING
                printer = st.session_state.get('selected_printer')
                if printer:
                    success, msg = extract_and_print_box(merged_pdf_path, int(target_box)-1, printer)
                    if success:
                        st.session_state['last_printed_box'] = int(target_box)
                        if 'printed_boxes' not in pkg: pkg['printed_boxes'] = []
                        pkg['printed_boxes'].append(int(target_box))
                        save_history(st.session_state['consignments'])
                        st.toast(f"🖨️ Printed Box {target_box}", icon="✅")
                    else:
                        st.toast(f"❌ Print Failed: {msg}", icon="🔥")
                else:
                    st.toast("⚠️ Select a printer first (Top Right)!", icon="⚠️")
        
        st.session_state.scan_input = ""

    # 3. Reprint Logic (Direct)
    def trigger_reprint(box_num):
        printer = st.session_state.get('selected_printer')
        if printer:
            success, msg = extract_and_print_box(merged_pdf_path, int(box_num)-1, printer)
            if success: 
                st.session_state['last_printed_box'] = int(box_num)
                st.toast(f"🖨️ Re-printed Box {box_num}", icon="✅")
            else: st.toast(f"❌ Reprint Failed: {msg}", icon="🔥")
        else: st.toast("Select Printer first (Top Right)")

    # --- UI LAYOUT ---
    
    # Header & Top Right Printer
    c_back, c_spacer, c_print = st.columns([1, 4, 2])
    with c_back: 
        if st.button("🔙 Back", use_container_width=True): nav('view_saved')
    
    with c_print:
        printers = get_printers()
        if 'selected_printer' not in st.session_state: st.session_state['selected_printer'] = printers[0] if printers else None
        st.selectbox("Select Printer", printers, key='selected_printer', label_visibility="collapsed")
        if not HAS_WIN32: st.caption("❌ Driver Missing")

    st.divider()

    # Input Area
    st.text_input("SCAN BARCODE (EAN / SKU / FSN)", key='scan_input', on_change=process_scan, placeholder="Click here and scan...", help="Press Enter after scanning")

    # Notification Area (Yellow Box)
    last_p = st.session_state.get('last_printed_box')
    if last_p:
        st.info(f"🖨️ Last Printed: **BOX {last_p}**", icon="✨")

    # Data Table Preparation
    printed_set = set(pkg.get('printed_boxes', []))
    display_df = df_boxes.copy()
    display_df['Status'] = display_df['Box No'].apply(lambda x: '✅ PRINTED' if x in printed_set else 'WAITING')
    
    def highlight_rows(row):
        box_num = row['Box No']
        if box_num == st.session_state.get('last_printed_box'):
            return ['background-color: #fff3cd'] * len(row) # Yellow
        elif row['Status'] == '✅ PRINTED':
            return ['background-color: #d4edda'] * len(row) # Green
        return [''] * len(row)

    st.subheader("Box List")
    st.caption("💡 Click a row to see Reprint options")

    event = st.dataframe(
        display_df.style.apply(highlight_rows, axis=1),
        use_container_width=True,
        hide_index=True,
        height=500,
        on_select="rerun",
        selection_mode="single-row"
    )

    if event.selection.rows:
        selected_idx = event.selection.rows[0]
        selected_box = display_df.iloc[selected_idx]['Box No']
        col_act1, col_act2 = st.columns([3, 1])
        with col_act1:
            st.warning(f"Selected: **Box {selected_box}**")
        with col_act2:
            if st.button(f"🖨️ Reprint Box {selected_box}", type="primary", use_container_width=True):
                trigger_reprint(selected_box)

# 6. VIEW SAVED
elif st.session_state['page'] == 'view_saved':
    home_button(); pkg = st.session_state['curr_con']
    c_id = pkg['id']
    
    st.title(f"Consignment: {c_id}")
    
    st.subheader("1. Download Files (Generated)")
    
    r1c1, r1c2, r1c3 = st.columns(3)
    with r1c1: orig_csv = io.BytesIO(); pkg['original_data'].to_csv(orig_csv, index=False); st.download_button("⬇ Consignment CSV", orig_csv.getvalue(), f"{c_id}.csv", "text/csv")
    with r1c2: st.download_button("⬇ Consignment Data PDF", generate_consignment_data_pdf(pkg['data'], pkg), f"Data_{c_id}.pdf")
    with r1c3: st.download_button("⬇ Confirm Consignment Upload (CSV)", generate_confirm_consignment_csv(pkg['data']), f"Confirm_{c_id}.csv", "text/csv")

    r2c1, r2c2, r2c3 = st.columns(3)
    with r2c1: st.download_button("⬇ Product Labels (Bartender)", generate_bartender_full(pkg['data']), f"Bartender_All_{c_id}.xlsx")
    with r2c2: st.download_button("⬇ Ewaybill Data (Excel)", generate_excel_simple(pkg['data'], ['SKU Id', 'Editable Qty', 'Cost Price'], f"Eway_{c_id}.xlsx"), f"Eway_{c_id}.xlsx")
    
    st.divider()
    
    st.subheader("2. File Repository & Merged Labels")
    
    # BOX LABELS (THE BIG MERGE)
    uc1, uc2 = st.columns([1, 1])
    with uc1:
        f_lbl = st.file_uploader("Upload Flipkart Box Labels PDF", type=['pdf'], key='u_lbl')
        if f_lbl: 
            if st.button("Process & Merge Labels"):
                save_uploaded_file(f_lbl, c_id, 'box_labels')
                progress_bar = st.progress(0, text="Starting Merge...")
                path_lbl = get_stored_file(c_id, 'box_labels')
                path_merged = get_merged_labels_path(c_id)
                try:
                    generate_merged_box_labels(pkg['data'], pkg, pkg['sender'], pkg['receiver'], path_lbl, progress_bar, save_path=path_merged)
                    progress_bar.progress(100, text="Completed!")
                    time.sleep(1)
                    st.success("Uploaded & Merged!")
                    st.rerun()
                except Exception as e: st.error(f"Error merging: {e}")

    with uc2:
        path_lbl = get_stored_file(c_id, 'box_labels')
        path_merged = get_merged_labels_path(c_id)
        
        # Priority: Show Merged File if exists
        if os.path.exists(path_merged):
            with open(path_merged, "rb") as f:
                st.download_button("⬇ Download MERGED Box Labels", f, f"Merged_Labels_{c_id}.pdf", "application/pdf")
            
            # --- SCAN & PRINT BUTTON ---
            st.divider()
            if st.button("🖨️ SCAN & PRINT BOX LABELS", type="primary", use_container_width=True):
                nav('scan_print')
            # --------------------------

        elif path_lbl:
            st.warning("Labels uploaded but not merged yet. Click 'Process & Merge'.")
            st.button("🖨️ SCAN & PRINT BOX LABELS", disabled=True, key='dis_scan')
        else:
            st.button("⬇ Download Box Labels", disabled=True, key='d_lbl_dis', help="Upload PDF first")
            st.button("🖨️ SCAN & PRINT BOX LABELS", disabled=True, key='dis_scan_2')

    # APPOINTMENT LETTER & CHALLAN (Combined row for space)
    st.divider()
    c_apt, c_chal = st.columns(2)
    
    with c_apt:
        st.write("Appointment Letter")
        f_apt = st.file_uploader("Upload Appt PDF", type=['pdf'], key='u_apt')
        if f_apt: 
            if st.button("Save Appt"): save_uploaded_file(f_apt, c_id, 'appointment'); st.rerun()
        path_apt = get_stored_file(c_id, 'appointment')
        if path_apt:
             with open(path_apt, "rb") as f: st.download_button("⬇ Download Appt", f, f"Appt_{c_id}.pdf")

    with c_chal:
        st.write("Challan")
        f_ch = st.file_uploader("Upload Challan PDF", type=['pdf'], key='u_ch')
        if f_ch: 
            if st.button("Save Challan"): save_uploaded_file(f_ch, c_id, 'challan'); st.rerun()
        path_ch = get_stored_file(c_id, 'challan')
        if path_ch:
             with open(path_ch, "rb") as f: st.download_button("⬇ Download Challan", f, f"Challan_{c_id}.pdf")

    st.divider(); st.subheader("3. Edit Qty & Print Labels")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a: 
        if st.button("Select All"): pkg['data']['Select'] = True
    with col_b:
        if st.button("Deselect All"): pkg['data']['Select'] = False
        
    if 'Select' not in pkg['data'].columns: pkg['data'].insert(0, 'Select', False)
    sorted_df = pkg['data'].sort_values(by='SKU Id').reset_index(drop=True)

    edited_df = st.data_editor(sorted_df, column_config={"Select": st.column_config.CheckboxColumn("Print?", default=False), "Editable Qty": st.column_config.NumberColumn("Qty", min_value=0), "Editable Boxes": st.column_config.NumberColumn("Boxes", min_value=0)}, disabled=["SKU Id", "Product Name", "PPCN"], column_order=["Select", "SKU Id", "Product Name", "Editable Qty", "Editable Boxes"], hide_index=True, use_container_width=True, key="editor")

    if st.button("🔄 Update Qty based on Box Count", type="secondary"):
        for i, row in edited_df.iterrows():
            ppcn = float(row['PPCN']) if row['PPCN'] > 0 else 1
            curr_box = float(row['Editable Boxes'])
            edited_df.at[i, 'Editable Qty'] = int(curr_box * ppcn)
        st.success("Qty updated!")
        pkg['data'] = edited_df; st.rerun()

    c1, c2 = st.columns(2); sel_rows = edited_df[edited_df['Select']==True]
    with c1:
        if not sel_rows.empty: st.download_button("⬇ Download Bartender (Selected)", generate_bartender_full(sel_rows), f"Bartender_Sel_{c_id}.xlsx")
        else: st.warning("Select rows to download.")
    with c2:
        if st.button("💾 Save Changes to Consignment"):
            pkg['data'] = edited_df
            for idx, h in enumerate(st.session_state['consignments']):
                if h['id'] == pkg['id']: st.session_state['consignments'][idx] = pkg
            save_history(st.session_state['consignments']); st.success("Saved!"); st.rerun()

    # --- DANGER ZONE (Bottom of Page) ---
    st.divider()
    with st.expander("🚫 Danger Zone"):
        st.caption("Deleting a consignment is permanent. Please be careful.")
        if st.button(f"🗑️ Delete Consignment {c_id}", type="primary"):
            st.session_state['consignments'] = [c for c in st.session_state['consignments'] if c['id'] != c_id]
            save_history(st.session_state['consignments'])
            nav('home')
