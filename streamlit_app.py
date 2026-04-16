# -*- coding: utf-8 -*-
import subprocess, sys
subprocess.run([sys.executable, "-m", "pip", "install", "python-docx", "openpyxl"],
               capture_output=True)

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# ── RTL helper ──────────────────────────────────────────────────────────────
def set_rtl_paragraph(para):
    pPr = para._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'right')
    pPr.append(jc)

def set_rtl_run(run):
    rPr = run._r.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)
    run.font.name = 'David'

def set_doc_rtl(doc):
    settings = doc.settings.element
    bidi = OxmlElement('w:bidi')
    settings.append(bidi)

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Excel → Word", page_icon="📄", layout="centered")

st.markdown("""
<style>
    body, .stApp { direction: rtl; }
    .stTextInput input, .stFileUploader, .stTable, .stDataFrame { direction: rtl; text-align: right; }
    h1, h2, h3, p, label { text-align: right; }
</style>
""", unsafe_allow_html=True)

st.title("📄 המרת שאלות ותשובות מ-Excel ל-Word")
st.markdown("העלה קובץ Excel עם שאלות בעמודה **A** ותשובות בעמודה **D**")

uploaded_file = st.file_uploader("בחר קובץ Excel", type=["xlsx", "xls"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, header=None, dtype=str)
    except Exception as e:
        st.error(f"שגיאה בקריאת הקובץ: {e}")
        st.stop()

    total_rows = len(df_raw)
    st.success(f"✅ הקובץ נטען — {total_rows} שורות")

    row_12_idx = 11
    if total_rows > row_12_idx:
        row_12 = df_raw.iloc[[row_12_idx]]
        rest = pd.concat([
            df_raw.iloc[1:row_12_idx],
            df_raw.iloc[row_12_idx + 1:]
        ])
        df_ordered = pd.concat([row_12, rest], ignore_index=True)
    else:
        df_ordered = df_raw.copy()

    questions = df_ordered.iloc[:, 0]
    answers   = df_ordered.iloc[:, 3] if df_ordered.shape[1] > 3 else pd.Series([""] * len(df_ordered))

    qa_pairs = []
    for q, a in zip(questions, answers):
        q_str = str(q).strip() if pd.notna(q) else ""
        a_str = str(a).strip() if pd.notna(a) else ""
        if q_str and q_str.lower() != "nan":
            qa_pairs.append((q_str, a_str if a_str.lower() != "nan" else ""))

    if not qa_pairs:
        st.warning("לא נמצאו שאלות תקינות בעמודה A")
        st.stop()

    st.markdown("### תצוגה מקדימה (5 ראשונות)")
    preview_data = [{"שאלה": q, "תשובה": a} for q, a in qa_pairs[:5]]
    st.table(preview_data)
    st.caption(f'סה"כ {len(qa_pairs)} שאלות ותשובות')

    doc_title = st.text_input("כותרת המסמך", value="שאלות ותשובות")

    if st.button("🔄 צור קובץ Word", type="primary"):
        doc = Document()
        set_doc_rtl(doc)

        for section in doc.sections:
            section.top_margin    = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin   = Inches(1.2)
            section.right_margin  = Inches(1.2)

        # כותרת
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_rtl_paragraph(title_para)
        title_run = title_para.add_run(doc_title)
        title_run.bold = True
        title_run.font.size = Pt(22)
        title_run.font.color.rgb = RGBColor(0x2E, 0x40, 0x57)
        set_rtl_run(title_run)
        doc.add_paragraph()

        for i, (q, a) in enumerate(qa_pairs, start=1):
            # שאלה
            q_para = doc.add_paragraph()
            set_rtl_paragraph(q_para)
            q_para.paragraph_format.space_before = Pt(10)
            q_para.paragraph_format.space_after  = Pt(2)
            q_num = q_para.add_run(f"{i}. ")
            q_num.bold = True
            q_num.font.size = Pt(13)
            q_num.font.color.rgb = RGBColor(0x2E, 0x40, 0x57)
            set_rtl_run(q_num)
            q_text = q_para.add_run(q)
            q_text.bold = True
            q_text.font.size = Pt(13)
            q_text.font.color.rgb = RGBColor(0x2E, 0x40, 0x57)
            set_rtl_run(q_text)

            # תשובה
            a_para = doc.add_paragraph()
            set_rtl_paragraph(a_para)
            a_para.paragraph_format.space_after = Pt(6)
            a_mark = a_para.add_run("תשובה: ")
            a_mark.bold = True
            a_mark.font.size = Pt(12)
            a_mark.font.color.rgb = RGBColor(0x27, 0xAE, 0x60)
            set_rtl_run(a_mark)
            a_text = a_para.add_run(a)
            a_text.font.size = Pt(12)
            a_text.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            set_rtl_run(a_text)

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        st.success(f"✅ המסמך נוצר עם {len(qa_pairs)} שאלות ותשובות!")
        st.download_button(
            label="⬇️ הורד קובץ Word",
            data=buf,
            file_name=f"{doc_title}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
