import streamlit as st
import streamlit.components.v1 as components
import uuid
import re
import io
import json
import datetime
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# ==========================================
# PAGE CONFIG 
# ==========================================
st.set_page_config(page_title="Lumina: Comparative Manuscript Suite", layout="wide")

# ==========================================
# TIME-BASED MACOS UI/UX INJECTION
# ==========================================
# Determine if it is Dark Mode based on time (7 PM to 7 AM is Dark Mode)
current_hour = datetime.datetime.now().hour
is_dark_mode = current_hour < 7 or current_hour >= 19

if is_dark_mode:
    # macOS Dark Mode (Mojave/Monterey)
    ui_css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; color: #f5f5f7 !important; }
        .stApp { background-color: #1c1c1e; }
        header, #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"] { background: rgba(40, 40, 40, 0.6) !important; backdrop-filter: blur(24px); border-right: 1px solid rgba(255, 255, 255, 0.1); }
        .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>div { background-color: rgba(60, 60, 60, 0.8) !important; color: white !important; border: 1px solid rgba(255, 255, 255, 0.1) !important; border-radius: 12px !important; }
        .stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus { border: 1px solid #0a84ff !important; box-shadow: 0 0 0 4px rgba(10, 132, 255, 0.15) !important; }
        .stButton>button { background: #2c2c2e; border: 1px solid rgba(255,255,255,0.1); border-radius: 10px; color: #f5f5f7; }
        .stButton>button:hover { background: #3a3a3c; transform: translateY(-1px); }
        .stButton>button[kind="primary"] { background: linear-gradient(180deg, #0a84ff 0%, #0060c0 100%); color: white !important; border: none; }
        .streamlit-expanderHeader { background-color: rgba(60,60,60,0.5); border-radius: 12px; font-weight: 600; color: #f5f5f7; }
        [data-testid="stExpander"] { background-color: rgba(40, 40, 40, 0.6); backdrop-filter: blur(10px); border-radius: 16px; border: 1px solid rgba(255,255,255,0.05); }
        [data-testid="stMetricValue"], [data-testid="stMetricLabel"] { color: #f5f5f7; }
    </style>
    """
else:
    # macOS Light Mode (Tahoe)
    ui_css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; color: #1d1d1f; }
        .stApp { background-color: #f5f5f7; background-image: radial-gradient(at 0% 0%, hsla(253,16%,7%,0.03) 0, transparent 50%), radial-gradient(at 50% 0%, hsla(225,39%,30%,0.03) 0, transparent 50%), radial-gradient(at 100% 0%, hsla(339,49%,30%,0.03) 0, transparent 50%); }
        header, #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"] { background: rgba(255, 255, 255, 0.6) !important; backdrop-filter: blur(24px); border-right: 1px solid rgba(0, 0, 0, 0.05); }
        .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>div { background-color: rgba(255, 255, 255, 0.8) !important; border: 1px solid rgba(0, 0, 0, 0.08) !important; border-radius: 12px !important; }
        .stTextInput>div>div>input:focus, .stTextArea>div>div>textarea:focus { border: 1px solid #0066cc !important; box-shadow: 0 0 0 4px rgba(0, 102, 204, 0.15) !important; }
        .stButton>button { background: #ffffff; border: 1px solid rgba(0,0,0,0.1); border-radius: 10px; color: #1d1d1f; font-weight: 500; }
        .stButton>button:hover { background: #fdfdfd; transform: translateY(-1px); }
        .stButton>button[kind="primary"] { background: linear-gradient(180deg, #007aff 0%, #0056b3 100%); color: white !important; border: none; }
        .streamlit-expanderHeader { background-color: rgba(255,255,255,0.5); border-radius: 12px; font-weight: 600; }
        [data-testid="stExpander"] { background-color: rgba(255, 255, 255, 0.6); backdrop-filter: blur(10px); border-radius: 16px; border: 1px solid rgba(0,0,0,0.05); }
        [data-testid="stMetricValue"] { color: #1d1d1f; }
        [data-testid="stMetricLabel"] { color: #86868b; }
    </style>
    """
st.markdown(ui_css, unsafe_allow_html=True)

# ==========================================
# PERSISTENCE & AUTOSAVE HELPERS
# ==========================================
RULES_FILE = "taxonomy_rules.json"
AUTOSAVE_FILE = "autosave_lumina_project.json"

def hex_to_rgb(hex_code):
    try:
        hex_code = hex_code.lstrip('#')
        return RGBColor(int(hex_code[0:2], 16), int(hex_code[2:4], 16), int(hex_code[4:6], 16))
    except:
        return RGBColor(26, 35, 126)

def save_rules_to_json(rules):
    serializable = [{'name': r['name'], 'hex_code': r.get('hex_code', '#000000'), 'keywords': r['keywords']} for r in rules]
    with open(RULES_FILE, "w", encoding="utf-8") as f:
        json.dump(serializable, f, indent=4)

def load_rules_from_json():
    if not os.path.exists(RULES_FILE): return []
    try:
        with open(RULES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return [{'name': r['name'], 'color': hex_to_rgb(r['hex_code']), 'hex_code': r['hex_code'], 'keywords': r['keywords']} for r in data]
    except: return []

def load_autosave():
    if not os.path.exists(AUTOSAVE_FILE): return None
    try:
        with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except: return None

def trigger_autosave():
    project_data = {
        "metadata": {
            "title": st.session_state.get('b_title', 'Untitled'), 
            "author": st.session_state.get('b_author', ''), 
            "year": st.session_state.get('b_year', ''),
            "is_rtl": st.session_state.get('is_rtl', False),
            "custom_labels": st.session_state.get('labels_raw', 'Variant A, Variant B, Commentary')
        },
        "rules": [{'name': r['name'], 'hex_code': r['hex_code'], 'keywords': r['keywords']} for r in st.session_state.rules],
        "chapters": st.session_state.chapters
    }
    with open(AUTOSAVE_FILE, "w", encoding="utf-8") as f:
        json.dump(project_data, f, indent=4)

# ==========================================
# XML INJECTORS
# ==========================================
def set_rtl_formatting(paragraph):
    pPr = paragraph._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT

def illuminate_run(run, hex_color):
    run.font.color.rgb = hex_to_rgb(hex_color)
    run.font.bold = True

# ==========================================
# HTML LIVE PREVIEW GENERATOR
# ==========================================
def generate_html_preview(text, rules, is_rtl=False, dark_mode=False):
    if not text: return ""
    all_kws = []
    for r in rules:
        for kw in r['keywords']:
            clean_kw = kw.strip()
            if clean_kw: all_kws.append((clean_kw, r.get('hex_code', '#000000')))
    all_kws.sort(key=lambda x: len(x[0]), reverse=True)

    escaped_text = text.replace('\n', '<br>')
    for kw, hex_code in all_kws:
        h_style = f'color: {hex_code}; font-weight: bold;'
        escaped_text = re.sub(f'({re.escape(kw)})', rf'<span style="{h_style}">\1</span>', escaped_text, flags=re.IGNORECASE)
            
    align = "right" if is_rtl else "left"
    font = "'Traditional Arabic', Arial, sans-serif" if is_rtl else "Georgia, serif"
    size = "22px" if is_rtl else "18px"
    direction = "rtl" if is_rtl else "ltr"
    
    # Adapt Preview colors based on Light/Dark Mode
    bg_color = "#2c2c2e" if dark_mode else "#f9f9f9"
    text_color = "#f5f5f7" if dark_mode else "black"
    border_color = "#555555" if dark_mode else "#dddddd"
    
    return f'<div style="text-align: {align}; font-family: {font}; font-size: {size}; direction: {direction}; line-height: 1.6; padding: 10px; background: {bg_color}; border-left: 4px solid {border_color}; margin-bottom: 10px; color: {text_color}; border-radius: 8px;">{escaped_text}</div>'

# ==========================================
# KNOWLEDGE WEB ENGINE (VIS.JS)
# ==========================================
def render_interactive_concordance(chapters, rules):
    nodes = []
    edges = []
    for i, r in enumerate(rules):
        nodes.append({"id": f"rule_{i}", "label": r['name'], "color": r.get('hex_code', '#333'), "font": {"color": "white"}, "shape": "box"})
        
    all_kws = []
    for i, r in enumerate(rules):
        for kw in r['keywords']:
            if kw.strip(): all_kws.append((kw.strip(), f"rule_{i}"))
    all_kws.sort(key=lambda x: len(x[0]), reverse=True)

    keyword_lookup = {kw.lower(): rule_id for kw, rule_id in all_kws}
    pattern_parts = [re.escape(kw) for kw, _ in all_kws]
    master_pattern = re.compile(r'(' + '|'.join(pattern_parts) + r')', re.IGNORECASE) if pattern_parts else None

    if master_pattern:
        for c in chapters:
            if 'sections' not in c: continue
            for s in c['sections']:
                sec_id = f"sec_{s['id']}"
                nodes.append({"id": sec_id, "label": f"{c['title']}\n{s['title']}", "color": "#e0e0e0", "shape": "ellipse"})
                src_txt = s.get('source_text', s.get('arabic_text',''))
                all_text = " ".join([str(src_txt), str(s.get('commentary', ''))] + [str(t.get('text', '')) for t in s.get('translations', [])])
                found_rules = set()
                for match in master_pattern.finditer(all_text):
                    r_id = keyword_lookup.get(match.group().lower())
                    if r_id: found_rules.add(r_id)
                for r_id in found_rules:
                    edges.append({"from": sec_id, "to": r_id})

    # Adjust Web Background for Dark Mode
    bg_web = "#1c1c1e" if is_dark_mode else "#ffffff"

    html_code = f"""
    <html>
    <head>
        <script type="text/javascript" src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
        <style type="text/css">#mynetwork {{width: 100%; height: 600px; border: 1px solid rgba(150,150,150,0.3); background-color: {bg_web}; border-radius: 12px;}}</style>
    </head>
    <body style="margin: 0; padding: 0;">
        <div id="mynetwork"></div>
        <script type="text/javascript">
            const nodes = new vis.DataSet({json.dumps(nodes)});
            const edges = new vis.DataSet({json.dumps(edges)});
            const container = document.getElementById('mynetwork');
            const data = {{nodes: nodes, edges: edges}};
            const options = {{
                nodes: {{borderWidth: 2, size: 30}},
                edges: {{color: 'lightgray', smooth: true}},
                physics: {{barnesHut: {{gravitationalConstant: -2000, centralGravity: 0.3, springLength: 150}}}}
            }};
            const network = new vis.Network(container, data, options);
        </script>
    </body>
    </html>
    """
    components.html(html_code, height=620)

# ==========================================
# TYPESETTING ENGINE (UNIVERSAL DOCX)
# ==========================================
def initialize_universal_styles(doc, is_rtl=False):
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(6), Inches(9)
    section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.75)
    section.left_margin, section.right_margin = Inches(0.8), Inches(0.5)
    styles = doc.styles

    try: src = styles.add_style('SourceText', WD_STYLE_TYPE.PARAGRAPH)
    except: src = styles['SourceText']
    src.font.name = 'Arial' if is_rtl else 'Georgia'
    src.font.size = Pt(16) if is_rtl else Pt(12)
    src.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_rtl else WD_ALIGN_PARAGRAPH.LEFT
    src.paragraph_format.space_after = Pt(12)

    try: tag = styles.add_style('VarTag', WD_STYLE_TYPE.PARAGRAPH)
    except: tag = styles['VarTag']
    tag.font.name = 'Garamond'
    tag.font.size = Pt(9)
    tag.font.bold = True
    tag.font.small_caps = True 
    tag.paragraph_format.space_after = Pt(2)
    tag.paragraph_format.space_before = Pt(8)

    try: body = styles.add_style('ParallelContent', WD_STYLE_TYPE.PARAGRAPH)
    except: body = styles['ParallelContent']
    body.font.name = 'Garamond'
    body.font.size = Pt(11)
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    body.paragraph_format.line_spacing = 1.0 
    body.paragraph_format.left_indent = Inches(0.25)
    body.paragraph_format.first_line_indent = Inches(-0.25)
    body.paragraph_format.space_after = Pt(8)

    try: comm = styles.add_style('AcadCommentary', WD_STYLE_TYPE.PARAGRAPH)
    except: comm = styles['AcadCommentary']
    comm.font.name = 'Garamond'
    comm.font.size = Pt(9.5)
    comm.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    comm.paragraph_format.space_after = Pt(12)

    try: conc = styles.add_style('ConcIndex', WD_STYLE_TYPE.PARAGRAPH)
    except: conc = styles['ConcIndex']
    conc.font.name = 'Garamond'
    conc.font.size = Pt(11)
    conc.paragraph_format.tab_stops.add_tab_stop(Inches(4.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

def build_secure_manuscript(metadata, chapters_data, categories_data):
    doc = Document()
    is_rtl = metadata.get('is_rtl', False)
    initialize_universal_styles(doc, is_rtl)

    all_kws = []
    for cat in categories_data:
        for kw in cat['keywords']:
            if kw.strip(): all_kws.append((kw.strip(), cat['name'], cat.get('hex_code', '#000000')))
    all_kws.sort(key=lambda x: len(x[0]), reverse=True)

    keyword_lookup = {kw.lower(): {'category': cat, 'original': kw, 'hex_code': hex} for kw, cat, hex in all_kws}
    pattern_parts = [re.escape(kw) for kw, _, _ in all_kws]
    master_pattern = re.compile(r'(' + '|'.join(pattern_parts) + r')', re.IGNORECASE) if pattern_parts else None
    concordance = {cat['name']: {} for cat in categories_data}

    def process_and_style(text, style_name, index_ref, p=None, force_rtl=False):
        if not text: return
        if p is None: p = doc.add_paragraph(style=style_name)
        if force_rtl: set_rtl_formatting(p)
        if not master_pattern:
            p.add_run(text)
            return
        last_idx = 0
        for match in master_pattern.finditer(text):
            start, end = match.span()
            matched = match.group()
            if start > last_idx: p.add_run(text[last_idx:start])
            run = p.add_run(matched)
            k_info = keyword_lookup.get(matched.lower())
            if k_info:
                illuminate_run(run, k_info['hex_code'])
                cat, orig = k_info['category'], k_info['original']
                if orig not in concordance[cat]: concordance[cat][orig] = set()
                concordance[cat][orig].add(index_ref)
            last_idx = end
        if last_idx < len(text): p.add_run(text[last_idx:])

    # 1. Front Matter
    doc.add_paragraph().add_run(f"\n\n\n\n{metadata.get('title','Untitled')}").bold = True
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].runs[0].font.size = Pt(28)
    doc.add_paragraph(f"\nBy {metadata.get('author','Anonymous Author')}").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    doc.add_paragraph(f"\n\nCopyright © {metadata.get('year','')} by {metadata.get('author','')}\n\nAll rights reserved.")
    doc.add_page_break()

    # 2. Key
    doc.add_heading("Research Taxonomy: Color Key", level=1)
    for cat in categories_data:
        p_label = doc.add_paragraph()
        run = p_label.add_run(f"  {cat['name'].upper()}  ")
        run.font.color.rgb = hex_to_rgb(cat.get('hex_code', '#000000'))
        run.font.bold = True; run.font.size = Pt(13)
    doc.add_page_break()

    # 3. Content
    for chap in chapters_data:
        doc.add_paragraph().add_run(f"\n\n\n{chap['title']}").bold = True
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[-1].runs[0].font.size = Pt(22)
        doc.add_page_break()
        for sec in chap.get('sections', []):
            i_ref = f"{chap['title']} : {sec['title']}"
            hdr = doc.add_heading(sec['title'], level=2)
            hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            tbl = doc.add_table(rows=1, cols=2)
            tbl.autofit = False; tbl.columns[0].width = Inches(3.2); tbl.columns[1].width = Inches(2.3)
            s_text = sec.get('source_text', sec.get('arabic_text',''))
            if s_text:
                p_src = tbl.rows[0].cells[1].paragraphs[0]
                p_src.style = 'SourceText'; process_and_style(s_text, 'SourceText', i_ref, p_src, force_rtl=is_rtl)
            f_trans = True
            for tr in sec.get('translations', []):
                if tr.get('text'):
                    p_tag = tbl.rows[0].cells[0].paragraphs[0] if f_trans else tbl.rows[0].cells[0].add_paragraph()
                    p_tag.style = 'VarTag'; f_trans = False; p_tag.add_run(f"[{tr['name']}]")
                    p_b = tbl.rows[0].cells[0].add_paragraph(style='ParallelContent')
                    process_and_style(tr['text'], 'ParallelContent', i_ref, p_b)
            if sec.get('commentary') or any(s['text'] for s in sec.get('sources',[])):
                doc.add_paragraph("________________________________________")
                if sec.get('commentary'):
                    doc.add_paragraph("AUTHOR'S COMMENTARY:", style='VarTag')
                    process_and_style(sec['commentary'], 'AcadCommentary', i_ref)
                if sec.get('sources'):
                    for s in sec['sources']:
                        if s['text']: process_and_style(s['text'], 'AcadCommentary', i_ref)
            doc.add_page_break()

    # 4. Conc
    doc.add_heading("Appendix: Concordance Index", level=1)
    for c_n, kw_d in concordance.items():
        if not kw_d: continue
        doc.add_heading(c_n, level=2)
        for kw, secs in sorted(kw_d.items()):
            doc.add_paragraph(f"{kw}\t{', '.join(sorted(list(secs)))}", style='ConcIndex')

    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream

# ==========================================
# STREAMLIT UI
# ==========================================
PRELOADED_COLORS = {
    "Deep Blue": "#1A237E", "Forest Green": "#1B5E20", "Goldenrod": "#B8860B",
    "Royal Purple": "#4A148C", "Dark Orange": "#E65100", "Saddle Brown": "#5D4037", "Crimson Red": "#B71C1C"
}

if 'rules' not in st.session_state: st.session_state.rules = load_rules_from_json()
if 'chapters' not in st.session_state:
    sv = load_autosave()
    if sv and 'chapters' in sv:
        st.session_state.chapters = sv['chapters']
        st.session_state.b_title = sv.get('metadata',{}).get('title', 'Comparative Study')
        st.session_state.b_author = sv.get('metadata',{}).get('author', '')
        st.session_state.b_year = sv.get('metadata',{}).get('year', str(datetime.datetime.now().year))
        st.session_state.is_rtl = sv.get('metadata',{}).get('is_rtl', False)
        st.session_state.labels_raw = sv.get('metadata',{}).get('custom_labels', 'Yusuf Ali, Sahih, Variant A')
    else:
        st.session_state.chapters = [{"id":str(uuid.uuid4()),"title":"Chapter 1","sections":[{"id":str(uuid.uuid4()),"title":"Section 1","source_text":"","translations":[{"id":str(uuid.uuid4()),"name":"Yusuf Ali","text":""}],"commentary":"","sources":[{"id":str(uuid.uuid4()),"text":""}]}]}]
        st.session_state.b_title, st.session_state.b_author, st.session_state.b_year = "Comparative Study", "", str(datetime.datetime.now().year)
        st.session_state.is_rtl, st.session_state.labels_raw = False, "Yusuf Ali, Sahih, Variant A"

LABELS = [l.strip() for l in st.session_state.get('labels_raw','').split(',') if l.strip()] + ["Custom..."]

# Control Center (Toggles)
c1, c2, c3 = st.columns(3)
zen = c1.toggle("🧘 Zen Mode", value=st.session_state.get('zen_mode', False))
show_dash = c2.toggle("📊 Show Dashboard", value=not zen)
show_side = c3.toggle("🛠️ Sidebar Controls", value=not zen)

st.session_state.zen_mode = zen

if show_side:
    with st.sidebar:
        if is_dark_mode:
            st.caption("🌙 Dark Mode Active")
        else:
            st.caption("☀️ Light Mode Active")
            
        st.header("📚 Project Details")
        st.session_state.b_title = st.text_input("Project Title", st.session_state.get('b_title', "Comparative Study"))
        st.session_state.b_author = st.text_input("Author Name", st.session_state.get('b_author', ""))
        st.session_state.b_year = st.text_input("Year", st.session_state.get('b_year', str(datetime.datetime.now().year)))
        st.session_state.is_rtl = st.toggle("Primary Text is RTL (Arabic/Hebrew)", value=st.session_state.get('is_rtl', False))
        st.session_state.labels_raw = st.text_input("Variant Labels (CSV)", st.session_state.get('labels_raw',''))
        
        st.divider(); st.header("🎨 Taxonomy")
        c_n = st.text_input("Category Name")
        col_c = st.selectbox("Color", list(PRELOADED_COLORS.keys()))
        kw_s = st.text_area("Keywords (CSV)")
        if st.button("➕ Add Rule"):
            if c_n and kw_s:
                st.session_state.rules.append({'name':c_n, 'hex_code':PRELOADED_COLORS[col_c], 'keywords':[k.strip() for k in kw_s.split(',') if k.strip()]})
                save_rules_to_json(st.session_state.rules); st.rerun()
        
        if st.session_state.rules:
            for i, r in enumerate(st.session_state.rules):
                with st.expander(f"{r['name']}"):
                    if st.button(f"Delete", key=f"dr_{i}"): 
                        st.session_state.rules.pop(i)
                        save_rules_to_json(st.session_state.rules); st.rerun()
            
            # --- FIXED: Added global Reset Rules button ---
            if st.button("🗑️ Reset All Rules", type="secondary", use_container_width=True):
                st.session_state.rules = []
                save_rules_to_json(st.session_state.rules); st.rerun()

        st.divider(); st.header("📂 Backup & Restore")
        up_rules = st.file_uploader("Import Rules (JSON)", type="json")
        if up_rules and st.button("⬆️ Apply Rules"):
            try:
                data = json.load(up_rules)
                new_l_rules = []
                for r in data:
                    h = r.get('hex_code', '#1A237E')
                    new_l_rules.append({'name': r.get('name', 'Untitled'), 'color': hex_to_rgb(h), 'hex_code': h, 'keywords': r.get('keywords', [])})
                st.session_state.rules = new_l_rules
                save_rules_to_json(st.session_state.rules); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
            
        up_proj = st.file_uploader("Load Project (JSON)", type="json")
        if up_proj and st.button("⬆️ Load Project Bundle"):
            try:
                d = json.load(up_proj)
                if "rules" in d: 
                    new_p_rules = []
                    for r in d["rules"]:
                        h = r.get('hex_code', '#1A237E')
                        new_p_rules.append({'name': r.get('name', 'Untitled'), 'color': hex_to_rgb(h), 'hex_code': h, 'keywords': r.get('keywords', [])})
                    st.session_state.rules = new_p_rules
                if "chapters" in d: 
                    chaps = d["chapters"]
                    for c in chaps:
                        for s in c.get('sections', []):
                            if 'arabic_text' in s and 'source_text' not in s: s['source_text'] = s.pop('arabic_text')
                    st.session_state.chapters = chaps
                m = d.get("metadata", {})
                st.session_state.b_title = m.get('title', st.session_state.b_title)
                st.session_state.b_author = m.get('author', st.session_state.b_author)
                st.session_state.is_rtl = m.get('is_rtl', False)
                st.session_state.labels_raw = m.get('custom_labels', st.session_state.labels_raw)
                trigger_autosave(); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
        
        if st.session_state.chapters:
            pd = {"metadata":{"title":st.session_state.b_title,"author":st.session_state.b_author,"year":st.session_state.b_year,"is_rtl":st.session_state.is_rtl,"custom_labels":st.session_state.labels_raw},"rules":[{'name':r['name'],'hex_code':r['hex_code'],'keywords':r['keywords']} for r in st.session_state.rules],"chapters":st.session_state.chapters}
            st.download_button("⬇️ Export Full Project (.json)", json.dumps(pd, indent=4), file_name="Project.json")
        st.success("🟢 Auto-Save Active")

if show_dash:
    st.markdown("### 📊 Scholar's Dashboard")
    cs = st.columns(4)
    tot_s = sum(len(c.get('sections',[])) for c in st.session_state.chapters)
    ak = [k for r in st.session_state.rules for k in r['keywords']]
    ic = 0
    if ak:
        try:
            pt = re.compile(r'(' + '|'.join([re.escape(k) for k in ak]) + r')', re.IGNORECASE)
            for c in st.session_state.chapters:
                for s in c.get('sections', []):
                    tx = str(s.get('source_text','')) + " " + str(s.get('commentary','')) + " " + "".join([str(tr.get('text','')) for tr in s.get('translations', [])])
                    ic += len(pt.findall(tx))
        except: pass
    cs[0].metric("Chapters", len(st.session_state.chapters)); cs[1].metric("Sections", tot_s); cs[2].metric("Taxonomies", len(st.session_state.rules)); cs[3].metric("Illuminations", ic)
    st.divider()

# --- FIXED TAB LOGIC: Cleaned tuple unpacking ---
if not zen:
    t1, t2 = st.tabs(["📝 Manuscript Builder", "🕸️ Knowledge Web"])
else:
    t1, t2 = st.tabs(["📝 Zen Mode Active", "🕸️ Disabled"])

with t1:
    for ci, ch in enumerate(st.session_state.chapters):
        if not zen:
            c1, c2 = st.columns([5,1])
            ch['title'] = c1.text_input("Chapter Title", ch['title'], key=f"ct_{ch['id']}")
            if c2.button("🗑️", key=f"dc_{ch['id']}"): st.session_state.chapters.pop(ci); st.rerun()
        
        for si, sc in enumerate(ch['sections']):
            with st.expander(f"🔹 {ch['title']} : {sc['title']}", expanded=(si == len(ch['sections'])-1)):
                e, p = st.tabs(["📝 Edit Data", "👁️ Live Preview"])
                with e:
                    sc['title'] = st.text_input("Section Title", sc['title'], key=f"st_{sc['id']}")
                    lvl = "Source Text (Anchor)" if not st.session_state.is_rtl else "Source Text (RTL/Arabic)"
                    sc['source_text'] = st.text_area(lvl, sc.get('source_text', sc.get('arabic_text','')), key=f"sr_{sc['id']}", height=80)
                    if not zen:
                        st.divider(); st.markdown("##### Parallel Texts / Variants")
                        for tr in sc.get('translations', []):
                            c1, c2 = st.columns([1,4])
                            tr['name'] = c1.selectbox("Label", LABELS, index=(LABELS.index(tr['name']) if tr['name'] in LABELS else 0), key=f"ln_{tr['id']}")
                            tr['text'] = c2.text_area("Variant Text", tr['text'], key=f"lt_{tr['id']}", height=68)
                        if st.button("➕ Add Variant", key=f"av_{sc['id']}"):
                            sc['translations'].append({"id":str(uuid.uuid4()), "name":LABELS[0], "text":""}); st.rerun()
                    st.divider(); st.markdown("##### Academic Commentary")
                    sc['commentary'] = st.text_area("Observations", sc['commentary'], key=f"ci_{sc['id']}", height=120)
                    if not zen:
                        for s in sc.get('sources', []): s['text'] = st.text_input("Source/Citation", s['text'], key=f"si_{s['id']}")
                        if st.button("➕ Add Source", key=f"as_s_{sc['id']}"): sc['sources'].append({"id":str(uuid.uuid4()), "text":""}); st.rerun()
                        if st.button("🗑️ Delete Section", key=f"ds_{sc['id']}"): ch['sections'].pop(si); st.rerun()
                with p:
                    st.markdown("##### Source Text Highlighting")
                    st.markdown(generate_html_preview(sc.get('source_text',''), st.session_state.rules, st.session_state.is_rtl, is_dark_mode), unsafe_allow_html=True)
                    st.markdown("##### Variants & Commentary Preview")
                    vo = "".join([f"<b>[{t['name']}]</b><br>{t['text']}<br><br>" for t in sc.get('translations',[]) if t['text']])
                    if sc['commentary']:
                        vo += f"<b>[Commentary]</b><br>{sc['commentary']}"
                    st.markdown(generate_html_preview(vo, st.session_state.rules, False, is_dark_mode), unsafe_allow_html=True)
        if not zen:
            if st.button(f"📜 + Add Section to {ch['title']}", key=f"as_{ch['id']}"):
                ch['sections'].append({"id":str(uuid.uuid4()), "title":f"Section {len(ch['sections'])+1}", "source_text":"", "translations":[{"id":str(uuid.uuid4()), "name":LABELS[0], "text":""}], "commentary":"","sources":[{"id":str(uuid.uuid4()),"text":""}]}); st.rerun()
        st.divider()
        
    if not zen:
        c1, c2 = st.columns(2)
        if c1.button("📘 + CREATE NEW CHAPTER", use_container_width=True):
            st.session_state.chapters.append({"id":str(uuid.uuid4()), "title":f"Chapter {len(st.session_state.chapters)+1}", "sections":[{"id":str(uuid.uuid4()), "title":"Section 1", "source_text":"", "translations":[{"id":str(uuid.uuid4()), "name":LABELS[0], "text":""}], "commentary":"","sources":[{"id":str(uuid.uuid4()),"text":""}]}]}); st.rerun()
        if c2.button("🚀 COMPILE MASTER MANUSCRIPT", type="primary", use_container_width=True):
            mt = {'title':st.session_state.get('b_title',''), 'author':st.session_state.get('b_author',''), 'year':st.session_state.get('b_year', str(datetime.datetime.now().year)), 'is_rtl': st.session_state.is_rtl}
            s = build_secure_manuscript(mt, st.session_state.chapters, st.session_state.rules)
            st.download_button("📥 Download (.docx)", s, f"{mt['title'].replace(' ','_')}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

with t2:
    if not zen:
        render_interactive_concordance(st.session_state.chapters, st.session_state.rules)
    else:
        st.info("🕸️ The Knowledge Web is disabled while Zen Mode is active.")

trigger_autosave()
