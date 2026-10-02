import re
import io
import pandas as pd
import openpyxl
import streamlit as st
import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageEnhance, ImageOps

st.set_page_config(page_title="Pipe Support 도면 데이터 자율 추출기", layout="wide")
st.title("📐 Pipe Support 도면 데이터 자율 추출기")

uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

def preprocess_image_for_ocr(img):
    """OCR 인식률 극대화를 위한 이미지 전처리 (흑백 전환 + 대비 강조)"""
    gray = img.convert('L')
    # 대비 강조
    enhancer = ImageEnhance.Contrast(gray)
    enhanced = enhancer.enhance(2.0)
    # 이진화 (Thresholding)
    threshold = 180
    binarized = enhanced.point(lambda p: 255 if p > threshold else 0)
    return binarized

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
    
    # 1. HANGER MK. NO. / SUPPORT TAG NO (예: SH7-3404-03, H260616HD-3)
    mk_m = re.search(r'(?:HANGER\s*MK\.?\s*NO\.?|TAG\s*NO\.?|MARK\s*NO\.?)[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
    if mk_m:
        val = mk_m.group(1).strip()
        data["HANGER MK. NO."] = val
        data["1. SUPPORT TAG NO"] = val
    else:
        # 서브패턴 (SH7-3404-03 과 같은 지원 태그 형식 직접 추출)
        tag_direct = re.search(r'\b[A-Z0-9]{2,5}-\d{3,5}-\d{2,3}\b', text)
        if tag_direct:
            data["1. SUPPORT TAG NO"] = tag_direct.group(0).strip()
            data["HANGER MK. NO."] = tag_direct.group(0).strip()

    # 2. Type & Size (예: VT-120-09LCG, VB09-J, JSLD-130-F 등)
    type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC|JLSTW|VB09)-[\w-]+', text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(0).strip()

    # 3. SPRING RATE (kgf/mm) - 예: 1.34
    sp_m = re.search(r'SPRING\s*RATE[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (kgf/mm)"] = float(sp_m.group(1))

    # 4. HOT LOAD (kgf) - 예: 343.3
    hl_m = re.search(r'HOT\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (kgf)"] = float(hl_m.group(1))

    # 5. COLD LOAD (kgf) - 예: 357.6
    cl_m = re.search(r'COLD\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (kgf)"] = float(cl_m.group(1))

    # 6. MOVEMENT / TRAVEL (mm) - 예: 10.7, 89.3
    mov_m = re.search(r'(?:THERMAL\s*MAXIMUM|MOVEMENT|MAX\.\s*TRAVEL)[^\d]*([\d\.]+)', text, re.IGNORECASE)
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
            text = page.get_text()
            
            # PDF 자체 텍스트가 풍부한 경우 우선 추출
            if text and len(text.strip()) > 50:
                identifier = f"{pdf_file.name} (p.{page_idx+1})"
                parsed = parse_text_to_fields(text, identifier)
                
                # 텍스트 추출로 항목이 안 채워지면 이미지 OCR 2차 수행
                if not parsed["1. SUPPORT TAG NO"] and not parsed["3. Spring Rate (kgf/mm)"]:
                    pix = page.get_pixmap(dpi=300)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    processed_img = preprocess_image_for_ocr(img)
                    ocr_text = pytesseract.image_to_string(processed_img, lang='eng')
                    parsed = parse_text_to_fields(ocr_text, identifier + "-OCR")
                
                rows.append(parsed)
            else:
                # 스캔본 도면인 경우 고해상도 이미지 전처리 후 OCR 수행
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                processed_img = preprocess_image_for_ocr(img)
                ocr_text = pytesseract.image_to_string(processed_img, lang='eng')
                identifier = f"{pdf_file.name} (p.{page_idx+1}-OCR)"
                rows.append(parse_text_to_fields(ocr_text, identifier))
                
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
