import re
import io
import pandas as pd
import pdfplumber
import openpyxl
import streamlit as st
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image

st.set_page_config(page_title="도면 데이터 자동 추출기", layout="wide")
st.title("📐 Pipe Support 도면 일괄 분석 및 엑셀 리스트업")

uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요 (스캔본/이미지 PDF 지원)", 
    type=["pdf"], 
    accept_multiple_files=True
)

def extract_text_from_pdf(pdf_file):
    """일반 PDF 텍스트 추출 시도 후, 실패 시 OCR 엔진 가동"""
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)
    
    full_text = ""
    # 1차: pdfplumber로 텍스트 추출 시도
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                full_text += t + "\n"
                
    # 2차: 텍스트가 없는 경우(스캔/이미지 PDF) OCR 처리
    if not full_text.strip():
        images = convert_from_bytes(pdf_bytes)
        for img in images:
            ocr_text = pytesseract.image_to_string(img, lang='eng+kor')
            full_text += ocr_text + "\n"
            
    return full_text

def extract_7_fields(pdf_file):
    text = extract_text_from_pdf(pdf_file)
    
    data = {
        "파일명": pdf_file.name,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (N/mm)": None,
        "4. Hot Load (N)": None,
        "5. Cold Load (N)": None,
        "6. Hot Point (mm)": None,
        "7. Cold Point (mm)": None
    }
    
    # 1. SUPPORT TAG NO.
    tag_match = re.search(r'SUPPORT\s*TAG\.?\s*NO\.?\s*[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
    if tag_match:
        data["1. SUPPORT TAG NO"] = tag_match.group(1).strip()
        
    # 2. Type & Size
    type_match = re.search(r'(?:VT|VC|VA|CS|CH|CHT)-[\w-]+', text)
    if type_match:
        data["2. Type & Size"] = type_match.group(0).strip()
        
    # 3. SPRING RATE
    spring_match = re.search(r'SPRING\s*RATE\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if spring_match:
        data["3. Spring Rate (N/mm)"] = float(spring_match.group(1))
        
    # 4. HOT LOAD
    hot_load_match = re.search(r'HOT\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if hot_load_match:
        data["4. Hot Load (N)"] = float(hot_load_match.group(1))
        
    # 5. COLD LOAD
    cold_load_match = re.search(r'COLD\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if cold_load_match:
        data["5. Cold Load (N)"] = float(cold_load_match.group(1))
        
    # 6. HOT POINT
    hot_point_match = re.search(r'HOT\s*POINT\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if hot_point_match:
        data["6. Hot Point (mm)"] = float(hot_point_match.group(1))
        
    # 7. COLD POINT
    cold_point_match = re.search(r'COLD\s*POINT\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if cold_point_match:
        data["7. Cold Point (mm)"] = float(cold_point_match.group(1))

    return data

if uploaded_pdfs:
    if st.button("🚀 업로드한 도면 전체 분석 및 엑셀 생성"):
        results = []
        progress_bar = st.progress(0)
        
        for idx, pdf in enumerate(uploaded_pdfs):
            res = extract_7_fields(pdf)
            results.append(res)
            progress_bar.progress((idx + 1) / len(uploaded_pdfs))
            
        df = pd.DataFrame(results)
        
        st.subheader("📋 추출 결과 요약")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Support_List')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 엑셀 리스트 다운로드 (.xlsx)",
            data=excel_data,
            file_name="Pipe_Support_Drawing_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
