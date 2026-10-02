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

def parse_rabigh_drawing_text(text, identifier):
    """
    Rabigh 2 / BHI 도면 특화 키워드 추적 파싱 로직
    """
    # 텍스트 내의 과도한 공백 및 줄바꿈 정리
    clean_text = " ".join(text.split())
    
    data = {
        "구분": identifier,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (N/mm)": None,
        "4. Hot Load (N)": None,
        "5. Cold Load (N)": None,
        "6. Movement (mm)": None,
    }

    # 1. SUPPORT TAG NO (SUPPORT TAG NO. 뒤 또는 11LAB... 식별자)
    tag_m = re.search(r'SUPPORT\s*TAG\s*\.?\s*NO\.?\s*([A-Z0-9-]+)', clean_text, re.IGNORECASE)
    if tag_m:
        data["1. SUPPORT TAG NO"] = tag_m.group(1).strip()
    else:
        tag_direct = re.search(r'\b(\d{2}[A-Z]{3}\d{2}[A-Z]{2}\d{3}-[A-Z0-9]+)\b', clean_text)
        if tag_direct:
            data["1. SUPPORT TAG NO"] = tag_direct.group(1).strip()

    # 2. Type & Size (VIS SUPPORT 또는 VT- 규격)
    type_m = re.search(r'\b(VT-\d{3}-[A-Z0-9]+)\b', clean_text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(1).strip()

    # 3. SPRING RATE
    sp_m = re.search(r'SPRING\s*RATE\s*(\d+(?:\.\d+)?)', clean_text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (N/mm)"] = float(sp_m.group(1))

    # 4. HOT LOAD
    hl_m = re.search(r'HOT\s*LOAD\s*(\d+(?:\.\d+)?)', clean_text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (N)"] = float(hl_m.group(1))

    # 5. COLD LOAD
    cl_m = re.search(r'COLD\s*LOAD\s*(\d+(?:\.\d+)?)', clean_text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (N)"] = float(cl_m.group(1))

    # 6. MOVEMENT (+Y UP 또는 THERMAL 수치)
    # +Y UP 뒤의 수치를 우선 매칭
    mov_m = re.search(r'\+Y\s*(?:UP)?\s*(\d+(?:\.\d+)?)', clean_text, re.IGNORECASE)
    if mov_m:
        data["6. Movement (mm)"] = float(mov_m.group(1))

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
            
            # PDF 텍스트 직접 추출
            text = page.get_text("text")
            parsed = parse_rabigh_drawing_text(text, identifier)
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
