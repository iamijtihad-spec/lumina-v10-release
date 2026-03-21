import streamlit as st
import streamlit.components.v1 as components
import uuid
import re
import io
import json
import datetime
import os
import string
from collections import Counter
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import pdfplumber

# ==========================================
# PAGE CONFIG & UI SYSTEM
# ==========================================
st.set_page_config(page_title="Lumina: Comparative Manuscript Suite", layout="wide")

current_hour = datetime.datetime.now().hour
is_dark_mode = current_hour < 7 or current_hour >= 19

if is_dark_mode:
    ui_css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; color: #f5f5f7 !important; }
        .stApp { background-color: #1c1c1e; }
        header, #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"] { background: rgba(40, 40, 40, 0.6) !important; backdrop-filter: blur(24px); border-right: 1px solid rgba(255, 255, 255, 0.1); }
        .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>div { background-color: rgba(60, 60, 60, 0.8) !important; color: white !important; border: 1px solid rgba(255, 255, 255, 0.1) !important; border-radius: 12px !important; }
        .stButton>button { background: #2c2c2e; border: 1px solid rgba(255,255,255,0.1); border-radius: 10px; color: #f5f5f7; }
        .stButton>button[kind="primary"] { background: linear-gradient(180deg, #0a84ff 0%, #0060c0 100%); color: white !important; border: none; }
        .streamlit-expanderHeader { background-color: rgba(60,60,60,0.5); border-radius: 12px; font-weight: 600; color: #f5f5f7; }
        [data-testid="stExpander"] { background-color: rgba(40, 40, 40, 0.6); backdrop-filter: blur(10px); border-radius: 16px; border: 1px solid rgba(255,255,255,0.05); }
    </style>
    """
else:
    ui_css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', -apple-system, sans-serif; color: #1d1d1f; }
        .stApp { background-color: #f5f5f7; background-image: radial-gradient(at 0% 0%, hsla(253,16%,7%,0.03) 0, transparent 50%), radial-gradient(at 50% 0%, hsla(225,39%,30%,0.03) 0, transparent 50%), radial-gradient(at 100% 0%, hsla(339,49%,30%,0.03) 0, transparent 50%); }
        header, #MainMenu, footer { visibility: hidden; }
        [data-testid="stSidebar"] { background: rgba(255, 255, 255, 0.6) !important; backdrop-filter: blur(24px); border-right: 1px solid rgba(0, 0, 0, 0.05); }
        .stTextInput>div>div>input, .stTextArea>div>div>textarea, .stSelectbox>div>div>div { background-color: rgba(255, 255, 255, 0.8) !important; border: 1px solid rgba(0, 0, 0, 0.08) !important; border-radius: 12px !important; }
        .stButton>button { background: #ffffff; border: 1px solid rgba(0,0,0,0.1); border-radius: 10px; color: #1d1d1f; font-weight: 500; }
        .stButton>button[kind="primary"] { background: linear-gradient(180deg, #007aff 0%, #0056b3 100%); color: white !important; border: none; }
        .streamlit-expanderHeader { background-color: rgba(255,255,255,0.5); border-radius: 12px; font-weight: 600; }
        [data-testid="stExpander"] { background-color: rgba(255, 255, 255, 0.6); backdrop-filter: blur(10px); border-radius: 16px; border: 1px solid rgba(0,0,0,0.05); }
    </style>
    """
st.markdown(ui_css, unsafe_allow_html=True)

# ==========================================
# PURE LOGIC ALGORITHMIC ANALYZERS
# ==========================================
def clean_word_for_analysis(word):
    return word.translate(str.maketrans('', '', string.punctuation + '؟،؛')).lower().strip()

def algorithmic_lexical_diff(variant_1_text, variant_2_text, v1_name, v2_name):
    v1_words = set(clean_word_for_analysis(w) for w in variant_1_text.split() if w)
    v2_words = set(clean_word_for_analysis(w) for w in variant_2_text.split() if w)
    unique_to_v1 = sorted(list(v1_words - v2_words))
    unique_to_v2 = sorted(list(v2_words - v1_words))
    shared_lexicon = v1_words & v2_words
    
    report = f"[Algorithmic Comparative Diff]\n"
    report += f"• Words unique to {v1_name}: {', '.join(unique_to_v1) if unique_to_v1 else 'None'}\n"
    report += f"• Words unique to {v2_name}: {', '.join(unique_to_v2) if unique_to_v2 else 'None'}\n"
    report += f"• Shared Vocabulary Density: {len(shared_lexicon)} exact matches."
    return report

def algorithmic_frequency_analysis(source_text):
    words = [clean_word_for_analysis(w) for w in source_text.split() if clean_word_for_analysis(w)]
    if not words: return "No source text provided for analysis."
    counts = Counter(words)
    top_words = counts.most_common(5)
    report = "[Statistical Source Analysis]\n• Dominant Tokens: "
    report += ", ".join([f"{w} ({c}x)" for w, c in top_words])
    report += f"\n• Total Word Count: {len(words)} | Unique Tokens: {len(counts)}"
    return report

# ==========================================
# HELPERS & PERSISTENCE
# ==========================================
RULES_FILE = "taxonomy_rules.json"
AUTOSAVE_FILE = "autosave_lumina_project.json"

def hex_to_rgb(hex_code):
    try: return RGBColor(int(hex_code.lstrip('#')[0:2], 16), int(hex_code.lstrip('#')[2:4], 16), int(hex_code.lstrip('#')[4:6], 16))
    except: return RGBColor(26, 35, 126)

def is_rtl_text(text):
    if not text: return False
    return any('\u0600' <= c <= '\u06FF' or '\u0590' <= c <= '\u05FF' for c in text)

def save_rules_to_json(rules):
    with open(RULES_FILE, "w", encoding="utf-8") as f: json.dump(rules, f, indent=4)

def load_rules_from_json():
    if not os.path.exists(RULES_FILE): return []
    try:
        with open(RULES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            for r in data:
                if 'type' not in r: r['type'] = 'keyword'
                r['color'] = hex_to_rgb(r.get('hex_code', '#000000'))
            return data
    except: return []

def load_autosave():
    if os.path.exists(AUTOSAVE_FILE):
        try:
            with open(AUTOSAVE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return None

def trigger_autosave():
    project_data = {
        "metadata": {"title": st.session_state.get('b_title', 'Untitled'), "author": st.session_state.get('b_author', ''), "year": st.session_state.get('b_year', ''), "custom_labels": st.session_state.get('labels_raw', 'Yusuf Ali, JPS Tanakh, KJV, NIV')},
        "rules": [{'name': r['name'], 'hex_code': r['hex_code'], 'type': r.get('type', 'keyword'), 'keywords': r.get('keywords', []), 'pattern': r.get('pattern', '')} for r in st.session_state.rules],
        "chapters": st.session_state.chapters
    }
    with open(AUTOSAVE_FILE, "w", encoding="utf-8") as f: json.dump(project_data, f, indent=4)

def build_master_regex(rules):
    pattern_parts = []
    for i, r in enumerate(rules):
        if r.get('type') == 'pattern' and r.get('pattern'):
            pattern_parts.append(rf"(?P<rule_{i}>{r['pattern']})")
        else:
            kws = r.get('keywords', [])
            if kws:
                kws = sorted([k.strip() for k in kws if k.strip()], key=len, reverse=True)
                group_pats = []
                for kw in kws:
                    escaped = re.escape(kw)
                    if is_rtl_text(kw): group_pats.append(escaped)
                    else: group_pats.append(r'(?<![a-zA-Z])' + escaped + r'(?![a-zA-Z])')
                if group_pats:
                    pattern_parts.append(rf"(?P<rule_{i}>" + "|".join(group_pats) + ")")
    if pattern_parts: return re.compile("|".join(pattern_parts), re.IGNORECASE)
    return None

# ==========================================
# HTML LIVE PREVIEW GENERATOR
# ==========================================
def generate_html_preview(text, rules, is_rtl=False, dark_mode=False, enable_highlights=True):
    if not text: return ""
    escaped_text = text.replace('\n', '<br>')
    if enable_highlights:
        master_pattern = build_master_regex(rules)
        if master_pattern:
            matches = list(master_pattern.finditer(escaped_text))
            for match in reversed(matches):
                rule_idx = int(match.lastgroup.split('_')[1])
                hex_code = rules[rule_idx].get('hex_code', '#000000')
                start, end = match.span()
                matched_str = escaped_text[start:end]
                h_style = f'background-color: {hex_code}; color: white; font-weight: bold; padding: 0 4px; border-radius: 4px;'
                escaped_text = escaped_text[:start] + f'<span style="{h_style}">{matched_str}</span>' + escaped_text[end:]
            
    align = "right" if is_rtl else "left"
    font = "'Traditional Arabic', Arial, sans-serif" if is_rtl else "Georgia, serif"
    size = "22px" if is_rtl else "18px"
    direction = "rtl" if is_rtl else "ltr"
    bg_color = "#2c2c2e" if dark_mode else "#f9f9f9"
    text_color = "#f5f5f7" if dark_mode else "black"
    border_color = "#555555" if dark_mode else "#dddddd"
    return f'<div style="text-align: {align}; font-family: {font}; font-size: {size}; direction: {direction}; line-height: 1.6; padding: 12px; background: {bg_color}; border-left: 4px solid {border_color}; margin-bottom: 10px; color: {text_color}; border-radius: 8px;">{escaped_text}</div>'

# ==========================================
# TYPESETTING ENGINE (.DOCX)
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

def initialize_universal_styles(doc):
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(6), Inches(9)
    section.top_margin, section.bottom_margin = Inches(0.75), Inches(0.75)
    section.left_margin, section.right_margin = Inches(0.8), Inches(0.5)
    styles = doc.styles

    try: src = styles.add_style('SourceText', WD_STYLE_TYPE.PARAGRAPH)
    except: src = styles['SourceText']
    src.font.name = 'Georgia'
    src.font.size = Pt(14)
    src.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    src.paragraph_format.space_after = Pt(12)
    
    try: src_rtl = styles.add_style('SourceTextRTL', WD_STYLE_TYPE.PARAGRAPH)
    except: src_rtl = styles['SourceTextRTL']
    src_rtl.font.name = 'Arial'
    src_rtl.font.size = Pt(16)
    src_rtl.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    src_rtl.paragraph_format.space_after = Pt(12)

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

def build_secure_manuscript(metadata, chapters_data, rules):
    doc = Document()
    initialize_universal_styles(doc)

    master_pattern = build_master_regex(rules)
    concordance = {r['name']: {} for r in rules}

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
            rule_idx = int(match.lastgroup.split('_')[1])
            rule = rules[rule_idx]
            if start > last_idx: p.add_run(text[last_idx:start])
            run = p.add_run(matched)
            illuminate_run(run, rule.get('hex_code', '#000000'))
            cat_name = rule['name']
            orig_match = matched.strip()
            if orig_match not in concordance[cat_name]: concordance[cat_name][orig_match] = set()
            concordance[cat_name][orig_match].add(index_ref)
            last_idx = end
        if last_idx < len(text): p.add_run(text[last_idx:])

    doc.add_paragraph().add_run(f"\n\n\n\n{metadata.get('title','Untitled')}").bold = True
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].runs[0].font.size = Pt(28)
    doc.add_paragraph(f"\nBy {metadata.get('author','Anonymous Author')}").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()
    doc.add_paragraph(f"\n\nCopyright © {metadata.get('year','')} by {metadata.get('author','')}\n\nAll rights reserved.")
    doc.add_page_break()

    doc.add_heading("Research Taxonomy: Color Key", level=1)
    for rule in rules:
        p_label = doc.add_paragraph()
        run = p_label.add_run(f"  {rule['name'].upper()}  ")
        run.font.color.rgb = hex_to_rgb(rule.get('hex_code', '#000000'))
        run.font.bold = True; run.font.size = Pt(13)
        if rule.get('type') == 'pattern':
            doc.add_paragraph(f"  (If-Then Rule: Matches '{rule.get('pattern')}')", style='AcadCommentary')
    doc.add_page_break()

    for chap in chapters_data:
        doc.add_paragraph().add_run(f"\n\n\n{chap['title']}").bold = True
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.paragraphs[-1].runs[0].font.size = Pt(22)
        doc.add_page_break()
        for sec in chap.get('sections', []):
            i_ref = f"{chap['title']} : {sec['title']}"
            hdr = doc.add_heading(sec['title'], level=2)
            hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            if sec.get('scripture_type'):
                p_sub = doc.add_paragraph(f"[{sec['scripture_type'].upper()}]")
                p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_sub.style = 'VarTag'
            
            tbl = doc.add_table(rows=1, cols=2)
            tbl.autofit = False; tbl.columns[0].width = Inches(3.2); tbl.columns[1].width = Inches(2.3)
            
            s_text = sec.get('source_text', sec.get('arabic_text',''))
            is_rtl_anchor = is_rtl_text(s_text)
            
            if s_text:
                p_src = tbl.rows[0].cells[1].paragraphs[0]
                if is_rtl_anchor:
                    p_src.style = 'SourceTextRTL'
                    process_and_style(s_text, 'SourceTextRTL', i_ref, p_src, force_rtl=True)
                else:
                    p_src.style = 'SourceText'
                    p_src.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    process_and_style(s_text, 'SourceText', i_ref, p_src, force_rtl=False)
            
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

    doc.add_heading("Appendix: Concordance Index", level=1)
    for c_n, kw_d in concordance.items():
        if not kw_d: continue
        doc.add_heading(c_n, level=2)
        for kw, secs in sorted(kw_d.items()):
            display_kw = kw if len(kw) < 40 else kw[:37] + "..."
            doc.add_paragraph(f"{display_kw}\t{', '.join(sorted(list(secs)))}", style='ConcIndex')

    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream

# ==========================================
# STREAMLIT UI INIT
# ==========================================
PRELOADED_COLORS = {"Deep Blue": "#1A237E", "Forest Green": "#1B5E20", "Goldenrod": "#B8860B", "Royal Purple": "#4A148C", "Dark Orange": "#E65100", "Saddle Brown": "#5D4037", "Crimson Red": "#B71C1C"}
SCRIPTURE_TRADITIONS = ["Quran", "Torah", "Bible (New Testament)", "Bible (Old Testament)", "Other Tradition"]

# Initialize Session States
if 'pdf_library' not in st.session_state: st.session_state.pdf_library = {}
if 'rules' not in st.session_state: st.session_state.rules = load_rules_from_json()
if 'chapters' not in st.session_state:
    sv = load_autosave()
    if sv and 'chapters' in sv:
        st.session_state.chapters = sv['chapters']
        st.session_state.b_title = sv.get('metadata',{}).get('title', 'Comparative Study')
        st.session_state.b_author = sv.get('metadata',{}).get('author', '')
        st.session_state.b_year = sv.get('metadata',{}).get('year', str(datetime.datetime.now().year))
        st.session_state.labels_raw = sv.get('metadata',{}).get('custom_labels', 'Yusuf Ali, JPS Tanakh, KJV, NIV')
    else:
        st.session_state.chapters = [{"id":str(uuid.uuid4()),"title":"Chapter 1","sections":[{"id":str(uuid.uuid4()),"title":"Section 1", "scripture_type": "Quran", "source_text":"","translations":[{"id":str(uuid.uuid4()),"name":"Yusuf Ali","text":""}],"commentary":"","sources":[{"id":str(uuid.uuid4()),"text":""}]}]}]
        st.session_state.b_title, st.session_state.b_author, st.session_state.b_year = "Comparative Scripture Study", "", str(datetime.datetime.now().year)
        st.session_state.labels_raw = "Yusuf Ali, JPS Tanakh, KJV, NIV"

LABELS = [l.strip() for l in st.session_state.get('labels_raw','').split(',') if l.strip()] + ["Custom..."]
zen = st.toggle("🧘 Zen Mode")

# --- SIDEBAR ---
if not zen:
    with st.sidebar:
        st.caption("🌙 Dark Mode Active" if is_dark_mode else "☀️ Light Mode Active")
        st.header("📚 Project Details")
        st.session_state.b_title = st.text_input("Project Title", st.session_state.get('b_title', "Comparative Scripture Study"))
        st.session_state.b_author = st.text_input("Author Name", st.session_state.get('b_author', ""))
        st.session_state.b_year = st.text_input("Year", st.session_state.get('b_year', str(datetime.datetime.now().year)))
        st.session_state.labels_raw = st.text_input("Historical Variant Labels (CSV)", st.session_state.get('labels_raw','Yusuf Ali, JPS Tanakh, KJV, NIV'))
        st.caption("Lumina automatically detects Hebrew and Arabic characters and formats them Right-to-Left.")
        
        st.divider(); st.header("🎨 Taxonomy Rules")
        c_n = st.text_input("Category Name")
        col_c = st.selectbox("Color", list(PRELOADED_COLORS.keys()))
        rule_type = st.radio("Rule Engine", ["Keyword List", "If-Then (Starts With)"], horizontal=True)
        if rule_type == "Keyword List":
            kw_s = st.text_area("Keywords (CSV)")
            if st.button("➕ Add Keyword Rule"):
                if c_n and kw_s:
                    st.session_state.rules.append({'name':c_n, 'type': 'keyword', 'hex_code':PRELOADED_COLORS[col_c], 'keywords':[k.strip() for k in kw_s.split(',') if k.strip()]})
                    save_rules_to_json(st.session_state.rules); st.rerun()
        else:
            trigger_word = st.text_input("If sentence starts with...")
            if st.button("➕ Add If-Then Rule"):
                if c_n and trigger_word:
                    pattern = rf"(?i)\b{re.escape(trigger_word)}.*?(?:[.?!؟]|$)"
                    st.session_state.rules.append({'name':c_n, 'type': 'pattern', 'hex_code':PRELOADED_COLORS[col_c], 'pattern': pattern})
                    save_rules_to_json(st.session_state.rules); st.rerun()

        if st.session_state.rules:
            for i, r in enumerate(st.session_state.rules):
                with st.expander(f"{r['name']}"):
                    if r.get('type') == 'pattern': st.code(r.get('pattern'))
                    else: st.write(f"Keywords: {', '.join(r.get('keywords',[]))}")
                    if st.button(f"Delete", key=f"dr_{i}"): 
                        st.session_state.rules.pop(i)
                        save_rules_to_json(st.session_state.rules); st.rerun()
            if st.button("🗑️ Reset All Rules", type="secondary", use_container_width=True):
                st.session_state.rules = []; save_rules_to_json(st.session_state.rules); st.rerun()

        st.divider(); st.header("📂 Backup & Restore")
        up_proj = st.file_uploader("Load Project (JSON)", type="json")
        if up_proj and st.button("⬆️ Load Bundle"):
            try:
                d = json.load(up_proj)
                if "rules" in d: st.session_state.rules = d["rules"]
                if "chapters" in d: st.session_state.chapters = d["chapters"]
                trigger_autosave(); st.rerun()
            except Exception as e: st.error(f"Error: {e}")
            
        if st.session_state.chapters:
            pd = {"metadata":{"title":st.session_state.b_title,"author":st.session_state.b_author,"year":st.session_state.b_year,"custom_labels":st.session_state.labels_raw},"rules":st.session_state.rules,"chapters":st.session_state.chapters}
            st.download_button("⬇️ Export Full Project (.json)", json.dumps(pd, indent=4), file_name="Project.json")

        st.divider()
        st.header("☕ Support the Project")
        st.markdown("<p style='font-size: 13px; opacity: 0.8;'>If Lumina has helped your research or workflow, consider supporting its continued development!</p>", unsafe_allow_html=True)
        st.markdown("""
        <div style="display: flex; flex-direction: column; gap: 12px; align-items: center; margin-top: 10px;">
            <a href="https://www.patreon.com/c/enki33" target="_blank">
                <img src="https://c5.patreon.com/external/logo/become_a_patron_button.png" alt="Become a Patron" style="height: 40px !important; border-radius: 8px; box-shadow: 0px 2px 5px rgba(0,0,0,0.1);" >
            </a>
        </div>
        """, unsafe_allow_html=True)

# --- TABS ---
if not zen:
    t1, t2, t3, t4 = st.tabs(["📝 Manuscript Builder", "📚 PDF Vault", "🕸️ Knowledge Web", "📖 User Guide"])
else:
    t1, t2, t3, t4 = st.tabs(["📝 Zen Mode Active", "📚 Disabled", "🕸️ Disabled", "📖 Disabled"])

# --- PDF VAULT TAB (TAB 2) ---
with t2:
    if not zen:
        st.header("📚 Digital Library Vault")
        st.markdown("Upload PDFs of your scriptures. Lumina extracts the text and loads it into RAM, allowing you to instantly search and cross-reference multiple books offline.")
        
        c_vault1, c_vault2 = st.columns([1, 2])
        with c_vault1:
            uploaded_file = st.file_uploader("Upload PDF Document", type="pdf")
            book_label = st.text_input("Label this Book (e.g., KJV Bible, Yusuf Ali Quran)")
            if st.button("📥 Index Book to RAM", use_container_width=True):
                if uploaded_file and book_label:
                    with st.spinner(f"Extracting and indexing '{book_label}'..."):
                        try:
                            # Using pdfplumber to safely extract text from the PDF stream
                            pages_text = []
                            with pdfplumber.open(uploaded_file) as pdf:
                                for page in pdf.pages:
                                    ext_txt = page.extract_text()
                                    pages_text.append(ext_txt if ext_txt else "")
                            st.session_state.pdf_library[book_label] = pages_text
                            st.success(f"Indexed {len(pages_text)} pages for '{book_label}'!")
                        except Exception as e:
                            st.error(f"Failed to read PDF: {e}")
                else:
                    st.warning("Please provide a PDF and a Label.")
                    
        with c_vault2:
            if st.session_state.pdf_library:
                st.subheader("Currently Loaded Books")
                for book_name, pages in st.session_state.pdf_library.items():
                    st.markdown(f"📖 **{book_name}** ({len(pages)} pages loaded)")
                if st.button("🗑️ Clear Library Memory"):
                    st.session_state.pdf_library = {}
                    st.rerun()
            else:
                st.info("Your library is currently empty. Upload a PDF to begin.")

# --- MANUSCRIPT BUILDER TAB (TAB 1) ---
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
                    # --- NEW: PDF VAULT SEARCH UI ---
                    if not zen and st.session_state.pdf_library:
                        with st.container():
                            st.markdown("##### 📖 Search PDF Library")
                            s_c1, s_c2, s_c3 = st.columns([2, 3, 1])
                            search_book = s_c1.selectbox("Select Book", list(st.session_state.pdf_library.keys()), key=f"sb_{sc['id']}")
                            search_query = s_c2.text_input("Search Phrase or Regex", key=f"sq_{sc['id']}")
                            if s_c3.button("Search", key=f"sbtn_{sc['id']}", use_container_width=True):
                                if search_query:
                                    found_results = []
                                    book_pages = st.session_state.pdf_library[search_book]
                                    for page_num, page_text in enumerate(book_pages):
                                        if search_query.lower() in page_text.lower():
                                            # Grab a snippet around the match
                                            idx = page_text.lower().find(search_query.lower())
                                            start = max(0, idx - 100)
                                            end = min(len(page_text), idx + len(search_query) + 100)
                                            snippet = page_text[start:end].replace('\n', ' ')
                                            found_results.append((page_num + 1, snippet))
                                    
                                    if found_results:
                                        st.success(f"Found {len(found_results)} matches in '{search_book}':")
                                        for p_num, snip in found_results[:5]: # Show top 5
                                            st.markdown(f"**Page {p_num}:** ...{snip}...")
                                    else:
                                        st.warning("No matches found.")
                        st.divider()

                    sc_c1, sc_c2 = st.columns([3, 1])
                    with sc_c1: sc['title'] = st.text_input("Section Title", sc['title'], key=f"st_{sc['id']}")
                    with sc_c2: sc['scripture_type'] = st.selectbox("Tradition", SCRIPTURE_TRADITIONS, index=(SCRIPTURE_TRADITIONS.index(sc.get('scripture_type', 'Quran')) if sc.get('scripture_type') in SCRIPTURE_TRADITIONS else 0), key=f"strad_{sc['id']}")
                    
                    st.markdown("##### Primary Source Text (Immutable Anchor)")
                    sc['source_text'] = st.text_area("Original Text (Hebrew, Arabic, Greek, etc.)", sc.get('source_text', sc.get('arabic_text','')), key=f"sr_{sc['id']}", height=80)
                    
                    if not zen:
                        st.divider()
                        st.markdown("##### ⚙️ Algorithmic Text Analysis (Offline & Private)")
                        algo_c1, algo_c2 = st.columns(2)
                        with algo_c1:
                            if st.button("📊 Run Source Frequency Analysis", use_container_width=True, key=f"freq_{sc['id']}"):
                                if sc.get('source_text'):
                                    freq_report = algorithmic_frequency_analysis(sc['source_text'])
                                    sc['commentary'] = sc.get('commentary', '') + f"\n\n{freq_report}"
                                    st.rerun()
                                else: st.warning("Please provide Source Text to analyze.")
                        with algo_c2:
                            if st.button("🔍 Diff-Map Historical Variants", use_container_width=True, key=f"diff_{sc['id']}"):
                                human_variants = [t for t in sc.get('translations', []) if t['text']]
                                if len(human_variants) >= 2:
                                    v1, v2 = human_variants[0], human_variants[1]
                                    diff_report = algorithmic_lexical_diff(v1['text'], v2['text'], v1['name'], v2['name'])
                                    sc['commentary'] = sc.get('commentary', '') + f"\n\n{diff_report}"
                                    st.rerun()
                                else: st.warning("Please add at least 2 historical variants (e.g., JPS Tanakh, KJV) to compare.")
                        
                        st.divider()
                        st.markdown("##### Historical Parallel Variants")
                        for tr in sc.get('translations', []):
                            c1, c2 = st.columns([1,4])
                            tr['name'] = c1.selectbox("Label", LABELS, index=(LABELS.index(tr['name']) if tr['name'] in LABELS else 0), key=f"ln_{tr['id']}")
                            tr['text'] = c2.text_area("Variant Text", tr['text'], key=f"lt_{tr['id']}", height=68)
                        
                        if st.button("➕ Add Historical Variant", key=f"av_{sc['id']}"):
                            sc['translations'].append({"id":str(uuid.uuid4()), "name":LABELS[0], "text":""}); st.rerun()
                            
                    st.divider(); st.markdown("##### Academic Commentary")
                    sc['commentary'] = st.text_area("Observations", sc['commentary'], key=f"ci_{sc['id']}", height=120)
                    
                    if not zen:
                        for s in sc.get('sources', []): s['text'] = st.text_input("Source/Citation", s['text'], key=f"si_{s['id']}")
                        if st.button("➕ Add Source", key=f"as_s_{sc['id']}"): sc['sources'].append({"id":str(uuid.uuid4()), "text":""}); st.rerun()
                        if st.button("🗑️ Delete Section", key=f"ds_{sc['id']}"): ch['sections'].pop(si); st.rerun()
                
                with p:
                    show_hl = st.toggle("🎨 Enable Taxonomy Highlighting", value=True, key=f"tog_{sc['id']}")
                    st.markdown("##### Primary Source Text")
                    is_rtl_prev = is_rtl_text(sc.get('source_text', ''))
                    st.markdown(generate_html_preview(sc.get('source_text',''), st.session_state.rules, is_rtl_prev, is_dark_mode, show_hl), unsafe_allow_html=True)
                    st.markdown("##### Variants & Commentary")
                    vo = "".join([f"<b>[{t['name']}]</b><br>{t['text']}<br><br>" for t in sc.get('translations',[]) if t['text']])
                    if sc['commentary']: vo += f"<b>[Author's Commentary]</b><br>{sc['commentary']}"
                    st.markdown(generate_html_preview(vo, st.session_state.rules, False, is_dark_mode, show_hl), unsafe_allow_html=True)
                    
        if not zen:
            if st.button(f"📜 + Add Section", key=f"as_{ch['id']}"):
                ch['sections'].append({"id":str(uuid.uuid4()), "title":f"Section {len(ch['sections'])+1}", "scripture_type": "Quran", "source_text":"", "translations":[{"id":str(uuid.uuid4()), "name":LABELS[0], "text":""}], "commentary":"","sources":[{"id":str(uuid.uuid4()),"text":""}]}); st.rerun()
        st.divider()
        
    if not zen:
        c1, c2 = st.columns(2)
        if c1.button("📘 + CREATE NEW CHAPTER", use_container_width=True):
            st.session_state.chapters.append({"id":str(uuid.uuid4()), "title":f"Chapter {len(st.session_state.chapters)+1}", "sections":[{"id":str(uuid.uuid4()), "title":"Section 1", "scripture_type": "Quran", "source_text":"", "translations":[{"id":str(uuid.uuid4()), "name":LABELS[0], "text":""}], "commentary":"","sources":[{"id":str(uuid.uuid4()),"text":""}]}]}); st.rerun()
        if c2.button("🚀 COMPILE MASTER MANUSCRIPT", type="primary", use_container_width=True):
            mt = {'title':st.session_state.get('b_title',''), 'author':st.session_state.get('b_author',''), 'year':st.session_state.get('b_year', str(datetime.datetime.now().year))}
            s = build_secure_manuscript(mt, st.session_state.chapters, st.session_state.rules)
            st.download_button("📥 Download (.docx)", s, f"{mt['title'].replace(' ','_')}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

with t3:
    if not zen:
        st.info("🕸️ The Knowledge Web generates a map based on your rules. (If-Then rules do not map well visually and are excluded).")
    else:
        st.info("🕸️ The Knowledge Web is disabled while Zen Mode is active.")

with t4:
    if not zen:
        st.title("📖 Lumina Official User Guide")
        st.markdown("""
        **Welcome to Lumina: The Comparative Manuscript Suite.** This application is designed as a "Digital Scriptorium" to help scholars, theologians, and researchers build complex, publisher-ready comparative manuscripts with zero friction across Abrahamic traditions.

        ---

        ### 🛡️ Core Philosophy: Algorithmic Privacy & Offline Independence
        Your research is your intellectual property. Lumina operates on a strict **Zero-Footprint, Logic-First Architecture**:
        * **No AI Dependencies:** All analysis is performed using hard-coded mathematical and lexical logic. Your texts are never sent to external servers or third-party AI corporations.
        * **Offline PDFs:** The new PDF Vault stores your books securely in local RAM without cloud processing.
        * Documents are generated entirely in your browser's RAM. 
        * Your work is silently auto-saved locally (`autosave_lumina_project.json`).

        ---

        ### 📚 Step 1: The PDF Vault
        Navigate to the **PDF Vault** tab. Upload any theological PDF (e.g., King James Bible, Quranic text). Lumina will extract the text from the PDF and securely index it into your computer's short-term memory. You can then instantly search across hundreds of pages directly inside the Manuscript Builder without ever leaving the app.

        ---

        ### ⚙️ Step 2: Project Setup (The Sidebar)
        1. **Metadata:** Enter your Book Title, Author Name, and Copyright Year.
        2. **Auto-Formatting Language:** Lumina natively detects if your Primary Source Text contains Arabic or Hebrew characters and will automatically inject Right-to-Left (RTL) formatting into your exported Word document.
        3. **Custom Labels:** Enter a comma-separated list of the historical translations you use most frequently (e.g., *Yusuf Ali, JPS Tanakh, KJV, NIV*).

        ---

        ### 🎨 Step 3: The Rule Engine (Color Taxonomy)
        * **Option A: Keyword List** Create a category (e.g., *Attributes of God*), pick a color, and paste a list of words. Lumina uses smart word-boundaries to highlight them dynamically.
        * **Option B: If-Then (Starts With) Rule**
          Create a category for complex grammatical structures. If you type the trigger word `Say,`, the engine will highlight everything from that word until the end of the sentence.

        ---

        ### 📝 Step 4: Drafting & Algorithmic Analysis
        Your manuscript is organized into **Chapters** and **Sections**.
        
        * **Editing:** Use the **📖 Search PDF Library** tool inside your section to pull quotes directly from the Vault.
        * **Algorithmic Text Analysis:** Instead of relying on internet-connected AI, Lumina provides pure logic tools:
          * **📊 Source Frequency Analysis:** Click this to mathematically extract the most frequently used root words in the source text.
          * **🔍 Diff-Map Historical Variants:** Click this to compare two historical translations. Lumina uses Set Theory to highlight exactly which words are unique to Translator A versus Translator B, and maps their shared vocabulary.
        * **Live Preview:** Click the `👁️ Live Preview` tab to see your highlighted typography exactly as it will appear in print.

        ---

        ### 🚀 Step 5: Exporting
        Click the **🚀 COMPILE MASTER MANUSCRIPT** button. 
        Lumina will instantly generate an Amazon KDP-compliant Microsoft Word `.docx` file featuring a Title Page, Copyright Page, Color Key, Hierarchical Chapters, a 60/40 Split Parallel Layout, and a fully indexed Auto-Concordance in the Appendix.
        """)
    else:
        st.info("📖 The User Guide is disabled while Zen Mode is active.")

trigger_autosave()
