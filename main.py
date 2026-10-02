import re
import io
import pandas as pd
import openpyxl
import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageEnhance

st.set_page_config(page_title="Pipe Support 도면 데이터 자율 추출기", layout="wide")
st.title("📐 Pipe Support 도면 데이터 자율 추출기")

uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

def preprocess_image_for_ocr(img):
    """OCR 인식률 극대화를 위한 이미지 전처리 (대비 강조)"""
    gray = img.convert('L')
    enhancer = ImageEnhance.Contrast(gray)
    enhanced = enhancer.enhance(2.0)
    threshold = 180
    binarized = enhanced.point(lambda p: 255 if p > threshold else 0)
    return binarized

def safe_float(value_str):
    """날짜('26.08.18')나 문자가 포함된 경우 안전하게 float으로 변환하는 함수"""
    if not value_str:
        return None
    try:
        # 날짜처럼 점(.)이 2개 이상 있는 문자열 제거
        clean_str = str(value_str).strip()
        if clean_str.count('.') > 1:
            return None
        # 순수 숫자와 점(.)만 추출
        clean_str = re.sub(r'[^0-9\.]', '', clean_str)
        if clean_str:
            return float(clean_str)
    except ValueError:
        return None
    return None

def parse_text_to_fields(text, identifier):
    """도면 표 구조에 특화된 정규표현식 파싱"""
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
    
    # 1. HANGER MK. NO. / SUPPORT TAG NO (예: SH7-3404-03, SM7-5356-01)
    mk_m = re.search(r'(?:HANGER\s*MK\.?\s*NO\.?|TAG\s*NO\.?|MARK\s*NO\.?)[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
    if mk_m:
        val = mk_m.group(1).strip()
        data["HANGER MK. NO."] = val
        data["1. SUPPORT TAG NO"] = val
    else:
        tag_direct = re.search(r'\b[A-Z0-9]{2,5}-\d{3,5}-\d{2,3}\b', text)
        if tag_direct:
            data["1. SUPPORT TAG NO"] = tag_direct.group(0).strip()
            data["HANGER MK. NO."] = tag_direct.group(0).strip()

    # 2. Type & Size (예: VT-120-C91CG, VT-080-081, VT-030-091CG 등)
    type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC|JLSTW|VB09)-[\w-]+', text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(0).strip()

    # 3. SPRING RATE (kgf/mm)
    sp_m = re.search(r'SPRING\s*RATE\s*[:\s]*(\d+\.?\d*)', text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (kgf/mm)"] = safe_float(sp_m.group(1))

    # 4. HOT LOAD (kgf)
    hl_m = re.search(r'HOT\s*LOAD\s*[:\s]*(\d+\.?\d*)', text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (kgf)"] = safe_float(hl_m.group(1))

    # 5. COLD LOAD (kgf)
    cl_m = re.search(r'COLD\s*LOAD\s*[:\s]*(\d+\.?\d*)', text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (kgf)"] = safe_float(cl_m.group(1))

    # 6. MOVEMENT / TRAVEL (mm)
    mov_m = re.search(r'(?:THERMAL\s*MAXIMUM|PIPE\s*MOVEMENT|MAX\.\s*TRAVEL)\s*[:\s]*(\d+\.?\d*)', text, re.IGNORECASE)
    if mov_m:
        data["6. Movement (mm)"] = safe_float(mov_m.group(1))

    return data

def process_pdf(pdf_file):
    pdf_bytes = pdf_file.read()
    pdf_file.seek(0)
    
    rows = []
    
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        for page_idx in range(len(doc)):
            page = doc[page_idx]
            text = page.get_text()
            
            identifier = f"{pdf_file.name} (p.{page_idx+1}-OCR)"
            
            # PDF 텍스트 파싱
            if text and len(text.strip()) > 50:
                parsed = parse_text_to_fields(text, identifier)
            else:
                # OCR 파싱
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                processed_img = preprocess_image_for_ocr(img)
                ocr_text = pytesseract.image_to_string(processed_img, lang='eng')
                parsed = parse_text_to_fields(ocr_text, identifier)
                
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
