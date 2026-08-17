import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from copy import deepcopy
from pathlib import Path
import io, re, zipfile

st.set_page_config(page_title="Tách truyện Word", page_icon="📖", layout="centered")

def wc(text):
    return len(re.findall(r"\S+", text))

def copy_run_fragment(src_run, text, dst_p):
    r = dst_p.add_run(text)
    r._r.get_or_add_rPr().extend(deepcopy(src_run._r.rPr)) if src_run._r.rPr is not None else None
    # copy run-level attributes such as rStyle/rFonts/etc via the full rPr
    if src_run._r.rPr is not None:
        old = r._r.rPr
        for child in list(old):
            old.remove(child)
        for child in list(src_run._r.rPr):
            old.append(deepcopy(child))
    return r

def copy_para_properties(src_p, dst_p):
    if src_p._p.pPr is not None:
        dst_p._p.insert(0, deepcopy(src_p._p.pPr))

def add_text_range_preserve_runs(src_p, start, end, out_doc, keep_empty=False):
    dst_p = out_doc.add_paragraph()
    # remove default pPr/runs if any
    for child in list(dst_p._p):
        if child.tag.endswith("}pPr"):
            dst_p._p.remove(child)
    copy_para_properties(src_p, dst_p)

    pos=0
    for run in src_p.runs:
        t=run.text or ""
        r0, r1 = pos, pos+len(t)
        a=max(start,r0); b=min(end,r1)
        if a < b:
            copy_run_fragment(run, t[a-r0:b-r0], dst_p)
        pos=r1
    if keep_empty and not dst_p.text:
        dst_p.add_run("")
    return dst_p

def sentence_units_for_paragraph(p):
    text=p.text
    if not text.strip():
        return []
    # Prefer sentence endings; newlines are also safe boundaries.
    matches=list(re.finditer(r'[.!?。！？…]+(?=\s|$)|\n+', text))
    units=[]
    start=0
    for m in matches:
        end=m.end()
        frag=text[start:end]
        if frag.strip():
            units.append((p,start,end,wc(frag)))
        start=end
    if start < len(text) and text[start:].strip():
        units.append((p,start,len(text),wc(text[start:])))
    return units

def build_units(doc):
    units=[]
    for p in doc.paragraphs:
        if not p.text.strip():
            continue
        units.extend(sentence_units_for_paragraph(p))
    return units

def make_chunk_doc(units):
    out=Document()
    # remove initial empty paragraph
    if out.paragraphs:
        p=out.paragraphs[0]
        p._element.getparent().remove(p._element)
    for src,start,end,_ in units:
        add_text_range_preserve_runs(src,start,end,out)
    return out

def split_units(units,target,min_w,max_w):
    chunks=[]; cur=[]; n=0
    for u in units:
        w=u[3]
        if cur and n+w > max_w:
            chunks.append(cur); cur=[]; n=0
        cur.append(u); n+=w
        if n >= target and n >= min_w:
            # Don't force exact target: the next sentence stays in the next chapter.
            chunks.append(cur); cur=[]; n=0
    if cur: chunks.append(cur)
    return chunks

def chunk_words(c):
    return sum(u[3] for u in c)

st.title("📖 Tách truyện Word")
st.write("Chia file Word thành các chương theo số từ, ưu tiên kết thúc ở cuối câu và giữ định dạng chữ.")

f=st.file_uploader("Chọn file .docx", type=["docx"])

if f:
    doc=Document(io.BytesIO(f.getvalue()))
    total=sum(wc(p.text) for p in doc.paragraphs)
    st.success(f"Đã đọc **{f.name}** — khoảng **{total:,} từ**.")

    c1,c2=st.columns(2)
    with c1:
        target=st.number_input("Số từ mục tiêu / chương",100,100000,5000,100)
    with c2:
        tol=st.number_input("Dung sai ± từ",0,20000,500,100)

    min_w=max(1,target-tol); max_w=target+tol
    st.info(f"Khoảng mục tiêu: **{min_w:,}–{max_w:,} từ/chương**. App ưu tiên cắt sau dấu câu.")

    if st.button("🔍 Phân tích & xem trước",use_container_width=True):
        units=build_units(doc)
        chunks=split_units(units,target,min_w,max_w)
        st.session_state["chunks"]=chunks
        st.session_state["source"]=f.name

    if "chunks" in st.session_state:
        chunks=st.session_state["chunks"]
        st.subheader(f"📚 Dự kiến {len(chunks)} chương")
        rows=[]
        for i,c in enumerate(chunks,1):
            n=chunk_words(c)
            status="✓" if min_w<=n<=max_w else "⚠"
            rows.append({"Chương":f"{i:03d}","Số từ":f"{n:,}","":status})
        st.dataframe(rows,use_container_width=True,hide_index=True)

        if st.button("✂️ Tách và tạo ZIP",type="primary",use_container_width=True):
            buf=io.BytesIO()
            with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as z:
                for i,c in enumerate(chunks,1):
                    out=make_chunk_doc(c)
                    b=io.BytesIO(); out.save(b)
                    z.writestr(f"Chuong_{i:03d}.docx",b.getvalue())
            buf.seek(0)
            stem=Path(st.session_state["source"]).stem
            st.download_button("⬇️ Tải ZIP",buf.getvalue(),f"{stem}_da_tach.zip","application/zip",use_container_width=True)
            st.success("Đã tạo xong.")

st.divider()
st.caption("V2: phù hợp với file truyện có nhiều dòng nằm trong cùng một paragraph; chia theo câu và sao chép định dạng run khi cắt giữa paragraph.")
