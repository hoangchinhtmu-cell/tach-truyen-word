import io
import re
from copy import deepcopy
from pathlib import Path

import streamlit as st
from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt

st.set_page_config(page_title="Tách truyện Word", page_icon="📖", layout="centered")

WORD_RE = re.compile(r"\S+")

def word_count(text):
    return len(WORD_RE.findall(text or ""))

def split_sentences(text):
    if not text or not text.strip():
        return []
    parts = re.split(r"(?<=[.!?。！？…])(?=\s+|$)", text.strip())
    return [p for p in parts if p.strip()]

def clone_paragraph(src_p, dst_doc):
    """
    Tạo một paragraph MỚI thay vì chép nguyên XML của paragraph cũ.
    Cách này tạo ra cấu trúc DOCX sạch:
        <w:p>...</w:p>
        <w:p>...</w:p>
    để các website đọc Word vẫn nhận đúng từng đoạn.
    Đồng thời cố gắng giữ style, định dạng paragraph và định dạng từng run.
    """
    new_p = dst_doc.add_paragraph()

    # Giữ style của paragraph nếu có.
    try:
        if src_p.style and src_p.style.name:
            new_p.style = src_p.style.name
    except Exception:
        pass

    # Sao chép paragraph properties (căn lề, thụt đầu dòng, giãn dòng...)
    try:
        if src_p._p.pPr is not None:
            new_p._p.insert(0, deepcopy(src_p._p.pPr))
    except Exception:
        pass

    # Sao chép từng run, không chép nguyên <w:p>.
    # Nhờ vậy paragraph mới luôn là một <w:p> độc lập.
    for run in src_p.runs:
        new_p._p.append(deepcopy(run._r))

    return new_p

def clear_default_paragraph(doc):
    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)

def paragraph_units(doc, max_words):
    units = []

    for p in doc.paragraphs:
        text = p.text or ""

        if not text.strip():
            units.append(("paragraph", p, 0))
            continue

        wc = word_count(text)

        if wc <= max_words:
            units.append(("paragraph", p, wc))
        else:
            sentences = split_sentences(text)
            if len(sentences) <= 1:
                units.append(("paragraph", p, wc))
            else:
                for s in sentences:
                    units.append(("text", s, word_count(s)))

    return units

def build_chunks(doc, target, tolerance):
    min_words = max(1, target - tolerance)
    max_words = target + tolerance

    units = paragraph_units(doc, max_words)

    chunks = []
    current = []
    current_words = 0

    for kind, obj, wc in units:
        if wc == 0:
            if current:
                current.append((kind, obj))
            continue

        # Khi chunk đã đạt mức tối thiểu, nếu thêm đơn vị này
        # vượt quá mức tối đa thì kết thúc chunk trước đó.
        if current and current_words >= min_words and current_words + wc > max_words:
            chunks.append(current)
            current = []
            current_words = 0

        current.append((kind, obj))
        current_words += wc

    if current:
        chunks.append(current)

    return chunks

def chunk_count(chunk):
    total = 0
    for kind, obj in chunk:
        total += word_count(obj.text if kind == "paragraph" else obj)
    return total

def add_chapter_title(doc, number):
    """
    Thêm tiêu đề Chương X ở đầu mỗi chương.
    Mỗi chương bắt đầu ở trang mới.
    """
    if len(doc.paragraphs) > 0 and doc.paragraphs[0].text == "":
        p = doc.paragraphs[0]
    else:
        p = doc.add_paragraph()

    p.style = doc.styles["Heading 1"]
    run = p.add_run(f"Chương {number}")
    run.bold = True
    run.font.size = Pt(16)

def add_page_break(doc):
    p = doc.add_paragraph()
    p.add_run().add_break(WD_BREAK.PAGE)

def build_one_word_document(chunks):
    """
    Tạo DUY NHẤT một file Word.
    Cấu trúc:
      Chương 1
      [nội dung]
      --- sang trang ---
      Chương 2
      [nội dung]
      ...
    """
    out = Document()
    clear_default_paragraph(out)

    for idx, chunk in enumerate(chunks, 1):
        if idx > 1:
            add_page_break(out)

        add_chapter_title(out, idx)

        for kind, obj in chunk:
            if kind == "paragraph":
                clone_paragraph(obj, out)
            else:
                # Câu được tách từ paragraph dài: tạo paragraph Word thật.
                out.add_paragraph(obj)

    return out

st.title("📖 Tách truyện Word")
st.caption(
    "Chia một file Word thành nhiều chương nhưng gộp tất cả chương vào **một file Word duy nhất**."
)

uploaded = st.file_uploader("Chọn file truyện .docx", type=["docx"])

if uploaded is None:
    st.info("👆 Hãy tải file Word lên để bắt đầu.")
    st.stop()

try:
    raw = uploaded.getvalue()
    doc = Document(io.BytesIO(raw))
except Exception as e:
    st.error(f"Không thể đọc file Word: {e}")
    st.stop()

total_words = sum(word_count(p.text) for p in doc.paragraphs)

st.success(
    f"Đã đọc **{uploaded.name}** · khoảng **{total_words:,} từ** · "
    f"**{len(doc.paragraphs):,} đoạn văn**."
)

st.subheader("⚙️ Cài đặt chia chương")

c1, c2 = st.columns(2)

with c1:
    target = st.number_input(
        "Số từ mục tiêu / chương",
        min_value=100,
        max_value=50000,
        value=5000,
        step=100,
    )

with c2:
    tolerance = st.number_input(
        "Dung sai ± từ",
        min_value=0,
        max_value=10000,
        value=500,
        step=100,
    )

min_words = max(1, target - tolerance)
max_words = target + tolerance

st.info(
    f"🎯 Mục tiêu: **{target:,} từ/chương** · "
    f"Khoảng cho phép: **{min_words:,}–{max_words:,} từ**"
)

if st.button("🔍 Tạo bản xem trước", use_container_width=True):
    with st.spinner("Đang phân tích và chia chương..."):
        chunks = build_chunks(doc, target, tolerance)

    st.session_state["chunks"] = chunks
    st.session_state["source_name"] = uploaded.name
    st.session_state.pop("output_bytes", None)

if "chunks" in st.session_state:
    chunks = st.session_state["chunks"]
    counts = [chunk_count(c) for c in chunks]

    st.success(
        f"Dự kiến **{len(chunks):,} chương** · "
        f"Trung bình **{sum(counts)/len(counts):,.0f} từ/chương**."
    )

    rows = []
    for i, wc in enumerate(counts, 1):
        rows.append({
            "Chương": f"Chương {i}",
            "Số từ": f"{wc:,}",
            "Trạng thái": "✓" if min_words <= wc <= max_words else "⚠",
        })

    st.dataframe(rows, use_container_width=True, hide_index=True)

    if st.button(
        "✂️ Tách và tạo 1 file Word",
        type="primary",
        use_container_width=True,
    ):
        with st.spinner("Đang tạo file Word..."):
            out_doc = build_one_word_document(chunks)
            output = io.BytesIO()
            out_doc.save(output)
            output_bytes = output.getvalue()

        st.session_state["output_bytes"] = output_bytes

        original = Path(
            st.session_state.get("source_name", "truyen.docx")
        ).stem
        st.session_state["output_name"] = f"{original}_da_tach.docx"

        st.success(
            f"✅ Đã tạo xong **1 file Word gồm {len(chunks):,} chương**."
        )

    if "output_bytes" in st.session_state:
        st.download_button(
            "⬇️ TẢI FILE WORD ĐÃ TÁCH",
            data=st.session_state["output_bytes"],
            file_name=st.session_state["output_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

st.divider()
st.caption(
    "V5: tất cả chương được gộp vào một file Word; mỗi chương có tiêu đề "
    "'Chương X' và bắt đầu ở trang mới."
)
