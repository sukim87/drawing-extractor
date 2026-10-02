import re
import io
import pandas as pd
import openpyxl
import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image

st.set_page_config(page_title="도면 데이터 자동 추출기", layout="wide")
st.title("📐 Pipe Support 도면 데이터 자율 추출기")

uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

def parse_text_to_fields(text, identifier):
    """텍스트에서 표 안의 핵심 수치 및 TAG 정보 추출"""
    data = {
        "구분": identifier,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (kgf/mm)": None,
        "4. Hot Load (kgf)": None,
        "5. Cold Load (kgf)": None,
        "6. Movement (mm)": None,
        "HANGER MK. NO.": None
    }
    
    # HANGER MK. NO. / TAG NO
    mk_m = re.search(r'HANGER\s*MK\.?\s*NO\.?\s*[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
    if mk_m:
        data["HANGER MK. NO."] = mk_m.group(1).strip()
        data["1. SUPPORT TAG NO"] = mk_m.group(1).strip()
        
    # TYPE OR SIZE (예: VT-120-09LCG)
    type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC)-[\w-]+', text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(0).strip()
        
    # SPRING RATE (kgf/mm)
    sp_m = re.search(r'SPRING\s*RATE\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (kgf/mm)"] = float(sp_m.group(1))
        
    # HOT LOAD (kgf)
    hl_m = re.search(r'HOT\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (kgf)"] = float(hl_m.group(1))
        
    # COLD LOAD (kgf)
    cl_m = re.search(r'COLD\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (kgf)"] = float(cl_m.group(1))
        
    # MOVEMENT (mm)
    mov_m = re.search(r'PIPE\s*MOVEMENT[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if mov_m:
        data["6. Movement (mm)"] = float(mov_m.group(1))

    return data

def process_pdf(pdf_file):
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)
    
    rows = []
    
    try:
        # PyMuPDF로 PDF 열기 (손상 및 그래픽 렌더링에 강력함)
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            text = page.get_text()
            
            # PDF 내부 텍스트가 있을 경우
            if text and len(text.strip()) > 20:
                identifier = f"{pdf_file.name} (p.{page_idx+1})"
                rows.append(parse_text_to_fields(text, identifier))
            else:
                # 텍스트가 없는 스캔본/벡터 이미지 도면은 고해상도 렌더링 후 OCR
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_text = pytesseract.image_to_string(img, lang='eng+kor')
                identifier = f"{pdf_file.name} (p.{page_idx+1}-OCR)"
                rows.append(parse_text_to_fields(ocr_text, identifier))
                
        doc.close()
    except Exception as e:
        st.error(f"⚠️️ '{pdf_file.name}' 파싱 중 오류 발생: {e}")
        
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
