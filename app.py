import io
import re
import zipfile
from copy import deepcopy
from pathlib import Path

import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Tách truyện Word",
    page_icon="📖",
    layout="centered",
)

# ----------------------------
# Helpers
# ----------------------------

WORD_RE = re.compile(r"\S+")

def word_count(text: str) -> int:
    return len(WORD_RE.findall(text or ""))

def split_sentences(text: str):
    """
    Tách câu tương đối an toàn, giữ nguyên nội dung.
    Ưu tiên các dấu kết thúc câu tiếng Việt/Trung.
    """
    text = text or ""
    if not text.strip():
        return []

    # Giữ dấu câu cùng câu. Khoảng trắng sau dấu câu là điểm ngắt.
    parts = re.split(r"(?<=[.!?。！？…])(?=\s+|$)", text)
    return [p for p in parts if p.strip()]

def clone_paragraph(src_p, dst_doc):
    """
    Copy toàn bộ paragraph XML để giữ định dạng, runs, style,
    alignment, spacing... tốt nhất có thể.
    """
    new_p = dst_doc.add_paragraph()
    new_p._p.getparent().remove(new_p._p)
    new_p._p.addnext(deepcopy(src_p._p))
    return dst_doc.paragraphs[-1]

def clear_default_paragraph(doc):
    if doc.paragraphs:
        p = doc.paragraphs[0]
        p._element.getparent().remove(p._element)

def make_doc_from_paragraphs(paragraphs):
    out = Document()
    clear_default_paragraph(out)
    for p in paragraphs:
        clone_paragraph(p, out)
    return out

def add_text_paragraph(out, text):
    out.add_paragraph(text)

def make_doc_from_units(units):
    """
    units:
      ("paragraph", source_paragraph)
      ("text", text)
    """
    out = Document()
    clear_default_paragraph(out)

    for kind, obj in units:
        if kind == "paragraph":
            clone_paragraph(obj, out)
        else:
            add_text_paragraph(out, obj)

    return out

def paragraph_units(doc):
    """
    Mỗi paragraph là một đơn vị. Nếu paragraph quá dài so với
    max_words thì chia thành các câu.
    """
    units = []

    for p in doc.paragraphs:
        text = p.text or ""

        # Đoạn trống: giữ như một paragraph để không phá bố cục.
        if not text.strip():
            units.append(("paragraph", p, 0))
            continue

        wc = word_count(text)

        if wc <= 0:
            units.append(("paragraph", p, 0))
        elif wc <= st.session_state.get("max_words", 5500):
            units.append(("paragraph", p, wc))
        else:
            # Paragraph quá dài: tách theo câu.
            sentences = split_sentences(text)
            if len(sentences) <= 1:
                units.append(("paragraph", p, wc))
            else:
                for s in sentences:
                    units.append(("text", s, word_count(s)))

    return units

def build_chunks(doc, target, tolerance):
    """
    Chia gần target, trong khoảng target +/- tolerance.
    Mỗi chunk ưu tiên kết thúc tại cuối paragraph/câu.
    Thuật toán không chốt quá sớm chỉ vì một paragraph nhỏ.
    """
    min_words = max(1, target - tolerance)
    max_words = target + tolerance

    # Dùng session_state chỉ để paragraph_units biết max.
    st.session_state["max_words"] = max_words
    units = paragraph_units(doc)

    chunks = []
    current = []
    current_words = 0

    for item in units:
        kind, obj, wc = item

        # Paragraph trống: chỉ giữ trong chunk hiện tại nếu chunk có nội dung.
        if wc == 0:
            if current:
                current.append((kind, obj))
            continue

        # Nếu đã đạt tối thiểu và thêm đơn vị này vượt max,
        # chốt chunk hiện tại trước.
        if current and current_words >= min_words and current_words + wc > max_words:
            chunks.append(current)
            current = []
            current_words = 0

        current.append((kind, obj))
        current_words += wc

        # Nếu đã đạt target, chỉ chốt ngay nếu:
        # - đã vượt/đạt max; hoặc
        # - đơn vị tiếp theo sẽ có nguy cơ vượt max.
        # Việc này được xử lý ở đầu vòng lặp, nên không chốt sớm ở đây.

    if current:
        chunks.append(current)

    return chunks

def chunk_count(chunk):
    total = 0
    for kind, obj in chunk:
        total += word_count(obj.text if kind == "paragraph" else obj)
    return total

def export_chunk(chunk):
    return make_doc_from_units([(kind, obj) for kind, obj, *_ in chunk])

def prepare_zip(chunks):
    buf = io.BytesIO()

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for idx, chunk in enumerate(chunks, 1):
            doc = export_chunk(chunk)
            dbuf = io.BytesIO()
            doc.save(dbuf)
            z.writestr(f"Chuong_{idx:03d}.docx", dbuf.getvalue())

    return buf.getvalue()

# ----------------------------
# UI
# ----------------------------

st.title("📖 Tách truyện Word")
st.caption("Chia một file Word thành các chương theo số từ bạn chọn, ưu tiên kết thúc ở cuối đoạn/câu.")

uploaded = st.file_uploader(
    "Chọn file truyện .docx",
    type=["docx"],
    help="File Word được xử lý trong phiên làm việc hiện tại.",
)

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

st.subheader("📋 Xem trước")

if st.button("🔍 Tạo bản xem trước", use_container_width=True):
    chunks = build_chunks(doc, target, tolerance)
    st.session_state["chunks"] = chunks
    st.session_state["target"] = target
    st.session_state["tolerance"] = tolerance
    st.session_state["source_name"] = uploaded.name
    st.session_state.pop("zip_bytes", None)

if "chunks" in st.session_state:
    chunks = st.session_state["chunks"]
    counts = [chunk_count(c) for c in chunks]

    st.success(
        f"Dự kiến **{len(chunks):,} chương** · "
        f"Trung bình **{sum(counts)/len(counts):,.0f} từ/chương**."
    )

    # Bảng gọn, không tạo None.
    rows = []
    for i, wc in enumerate(counts, 1):
        status = "✓" if min_words <= wc <= max_words else "⚠"
        rows.append({"Chương": f"{i:03d}", "Số từ": f"{wc:,}", "Trạng thái": status})

    st.dataframe(rows, use_container_width=True, hide_index=True)

    if st.button("✂️ Tách và tạo ZIP", type="primary", use_container_width=True):
        with st.spinner("Đang tạo các file Word..."):
            zip_bytes = prepare_zip(chunks)

        st.session_state["zip_bytes"] = zip_bytes

        original = Path(st.session_state.get("source_name", "truyen.docx")).stem
        st.session_state["zip_name"] = f"{original}_da_tach.zip"

        st.success(
            f"✅ Đã tạo xong **{len(chunks):,} file Word** "
            f"({len(zip_bytes) / (1024 * 1024):.1f} MB)."
        )

    if "zip_bytes" in st.session_state:
        st.download_button(
            label="⬇️ TẢI FILE ZIP",
            data=st.session_state["zip_bytes"],
            file_name=st.session_state["zip_name"],
            mime="application/zip",
            use_container_width=True,
        )

st.divider()
st.caption(
    "V3: ưu tiên điểm kết thúc ở cuối paragraph/câu, không chốt quá sớm, "
    "và giữ định dạng Word tốt nhất có thể."
)
