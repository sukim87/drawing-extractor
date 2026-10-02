import io
import re
import pandas as pd
import openpyxl
import streamlit as st
import fitz  # PyMuPDF

st.set_page_config(page_title="Pipe Support 도면 데이터 자율 추출기", layout="wide")
st.title("📐 Pipe Support 도면 데이터 자율 추출기")

uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

def extract_numbers_from_str(s):
    """문자열에서 숫자(소수점 포함) 추출"""
    m = re.search(r'(\d+(?:\.\d+)?)', str(s))
    return float(m.group(1)) if m else None

def parse_robust_drawing(page, identifier):
    """
    라벨 위치 기반 유연 탐색 로직 (텍스트 씹힘 방지)
    """
    data = {
        "구분": identifier,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (N/mm)": None,
        "4. Hot Load (N)": None,
        "5. Cold Load (N)": None,
        "6. Movement (mm)": None,
    }

    # 전체 페이지 텍스트 & 단어 바운딩 박스 추출
    raw_text = page.get_text("text")
    words = page.get_text("words")  # (x0, y0, x1, y1, word, block_no, line_no, word_no)

    # 1. SUPPORT TAG NO
    # 'SUPPORT TAG' 근처 단어에서 식별자 검색
    tag_rects = page.search_for("SUPPORT TAG") or page.search_for("TAG NO")
    if tag_rects:
        r = tag_rects[0]
        # 해당 라벨 우측 및 아래쪽 범위를 넉넉하게 잡아서 탐색
        nearby = [w[4] for w in words if (r.x0 - 20 <= w[0] <= r.x1 + 300) and (r.y0 - 10 <= w[1] <= r.y1 + 80)]
        for w in nearby:
            m = re.search(r'([0-9A-Z]{5,}-[A-Z0-9]+)', w)
            if m:
                data["1. SUPPORT TAG NO"] = m.group(1)
                break
    
    # 예비책: 전체 단어 중 11LAB... 또는 SH... 패턴 직접 탐색
    if not data["1. SUPPORT TAG NO"]:
        m_direct = re.search(r'\b(11LAB[A-Z0-9-]+)\b', raw_text) or re.search(r'\b(S[HM]\d*-[A-Z0-9-]+)\b', raw_text)
        if m_direct:
            data["1. SUPPORT TAG NO"] = m_direct.group(1)

    # 2. Type & Size (BOM 영역의 VT-030-17LCG 패턴 탐색)
    vt_match = re.search(r'\b(VT-\d{3}-[\w]+)\b', raw_text, re.IGNORECASE)
    if vt_match:
        data["2. Type & Size"] = vt_match.group(1).upper()

    # 3. SPRING RATE
    sp_rects = page.search_for("SPRING RATE")
    if sp_rects:
        r = sp_rects[0]
        nearby = [w[4] for w in words if (r.x0 - 20 <= w[0] <= r.x1 + 250) and (r.y0 - 5 <= w[1] <= r.y1 + 40)]
        for w in nearby:
            val = extract_numbers_from_str(w)
            if val is not None:
                data["3. Spring Rate (N/mm)"] = val
                break

    # 4. HOT LOAD
    hl_rects = page.search_for("HOT LOAD")
    if hl_rects:
        r = hl_rects[0]
        nearby = [w[4] for w in words if (r.x0 - 20 <= w[0] <= r.x1 + 250) and (r.y0 - 5 <= w[1] <= r.y1 + 40)]
        for w in nearby:
            val = extract_numbers_from_str(w)
            if val is not None:
                data["4. Hot Load (N)"] = val
                break

    # 5. COLD LOAD
    cl_rects = page.search_for("COLD LOAD")
    if cl_rects:
        r = cl_rects[0]
        nearby = [w[4] for w in words if (r.x0 - 20 <= w[0] <= r.x1 + 250) and (r.y0 - 5 <= w[1] <= r.y1 + 40)]
        for w in nearby:
            val = extract_numbers_from_str(w)
            if val is not None:
                data["5. Cold Load (N)"] = val
                break

    # 6. MOVEMENT (+Y UP 근처 탐색)
    y_rects = page.search_for("+Y") or page.search_for("THERMAL")
    if y_rects:
        r = y_rects[0]
        nearby = [w[4] for w in words if (r.x1 <= w[0] <= r.x1 + 200) and (r.y0 - 15 <= w[1] <= r.y1 + 15)]
        for w in nearby:
            val = extract_numbers_from_str(w)
            if val is not None:
                data["6. Movement (mm)"] = val
                break

    return data

def process_pdf(pdf_file):
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)
    rows = []

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            identifier = f"{pdf_file.name} (p.{page_idx+1})"
            parsed = parse_robust_drawing(page, identifier)
            rows.append(parsed)
        doc.close()
    except Exception as e:
        st.error(f"⚠ '{pdf_file.name}' 분석 중 오류: {e}")

    return rows

if uploaded_pdfs:
    if st.button("🚀 업로드 도면 분석 및 엑셀 리스트업"):
        all_results = []
        progress_bar = st.progress(0)

        for idx, pdf in enumerate(uploaded_pdfs):
            results = process_pdf(pdf)
            all_results.extend(results)
            progress_bar.progress((idx + 1) / len(uploaded_pdfs))

        df = pd.DataFrame(all_results)
        st.subheader("📋 분석 결과 리스트")
        st.dataframe(df, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Support_List')
        excel_data = output.getvalue()

        st.download_button(
            label="📥 엑셀 파일 다운로드 (.xlsx)",
            data=excel_data,
            file_name="Pipe_Support_Extracted_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
