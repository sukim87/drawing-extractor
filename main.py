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
st.title("📐 도면 양식 자율 인식 Pipe Support 데이터 리스트업")

uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

def parse_text_to_fields(text, identifier):
    """텍스트 맥락 추적 기반 핵심 7가지 수치 추출"""
    data = {
        "구분": identifier,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (N/mm)": None,
        "4. Hot Load (N)": None,
        "5. Cold Load (N)": None,
        "6. Hot Point (mm)": None,
        "7. Cold Point (mm)": None,
        "비고": "정상"
    }
    
    # 1. SUPPORT TAG NO
    tag_m = re.search(r'(?:SUPPORT\s*TAG|TAG\s*NO|TAG)\.?\s*[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
    if tag_m:
        data["1. SUPPORT TAG NO"] = tag_m.group(1).strip()
        
    # 2. Type & Size
    type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC)-[A-Z0-9-]+', text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(0).strip()
        
    # 3. SPRING RATE
    sp_m = re.search(r'SPRING\s*RATE[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (N/mm)"] = float(sp_m.group(1))
        
    # 4. HOT LOAD
    hl_m = re.search(r'HOT\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (N)"] = float(hl_m.group(1))
        
    # 5. COLD LOAD
    cl_m = re.search(r'COLD\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (N)"] = float(cl_m.group(1))
        
    # 6. HOT POINT
    hp_m = re.search(r'HOT\s*POINT[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if hp_m:
        data["6. Hot Point (mm)"] = float(hp_m.group(1))
        
    # 7. COLD POINT
    cp_m = re.search(r'COLD\s*POINT[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if cp_m:
        data["7. Cold Point (mm)"] = float(cp_m.group(1))

    return data

def process_single_pdf_safe(pdf_file):
    """손상되거나 깨진 PDF도 예외 처리 및 OCR 기반 안전 분석"""
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)
    
    extracted_rows = []
    
    # 1차 시도: pdfplumber 표준 텍스트 추출
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                if text.strip():
                    identifier = f"{pdf_file.name} (p.{page_idx+1})"
                    extracted_rows.append(parse_text_to_fields(text, identifier))
    except Exception as e:
        st.warning(f"⚠️ '{pdf_file.name}' 파일의 PDF 구조에 오류가 감지되어 이미지 OCR 판독 모드로 전환합니다.")

    # 2차 시도 (1차 실패 또는 스캔본 PDF인 경우): PDF -> Image -> OCR 판독
    if not extracted_rows:
        try:
            images = convert_from_bytes(pdf_bytes)
            for idx, img in enumerate(images):
                ocr_text = pytesseract.image_to_string(img, lang='eng+kor')
                identifier = f"{pdf_file.name} (도면 #{idx+1})"
                extracted_rows.append(parse_text_to_fields(ocr_text, identifier))
        except Exception as ocr_err:
            st.error(f"❌ '{pdf_file.name}' 파일을 판독할 수 없습니다: PDF 자체가 완전히 손상되었습니다.")
            extracted_rows.append({
                "구분": pdf_file.name,
                "비고": "파일 손상으로 읽기 실패"
            })
            
    return extracted_rows

if uploaded_pdfs:
    if st.button("🚀 업로드 도면 분석 및 엑셀 리스트업"):
        results = []
        progress_bar = st.progress(0)
        
        for idx, pdf in enumerate(uploaded_pdfs):
            rows = process_single_pdf_safe(pdf)
            results.extend(rows)
            progress_bar.progress((idx + 1) / len(uploaded_pdfs))
            
        df = pd.DataFrame(results)
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
