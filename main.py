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
    """도면 OCR 인식률을 극대화하는 이미지 전처리"""
    # 1. 흑백 전환
    gray = img.convert('L')
    
    # 2. 선명도 및 대비 2.5배 강화
    enhancer = ImageEnhance.Contrast(gray)
    enhanced = enhancer.enhance(2.5)
    
    # 3. 임계값(Thresholding)을 이용한 이진화 (글자 명확화)
    threshold = 170
    binarized = enhanced.point(lambda p: 255 if p > threshold else 0)
    return binarized

def crop_table_area(img):
    """도면의 표(Hanger Setting Data / BOM)가 위치한 우측 하단 50% 영역 크롭"""
    width, height = img.size
    # (left, upper, right, lower)
    crop_box = (int(width * 0.45), int(height * 0.40), width, height)
    return img.crop(crop_box)

def safe_float(value_str):
    """날짜 형식('26.08.18')이나 노이즈 텍스트 예외 처리"""
    if not value_str:
        return None
    try:
        clean_str = str(value_str).strip()
        # 점(.)이 2개 이상 들어간 날짜 표기 제거
        if clean_str.count('.') > 1:
            return None
        clean_str = re.sub(r'[^0-9\.]', '', clean_str)
        if clean_str:
            return float(clean_str)
    except ValueError:
        return None
    return None

def parse_text_to_fields(text, identifier):
    """도면 표 구조에 맞춘 정규표현식 파싱"""
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
    
    # 1. SUPPORT TAG NO / HANGER MK. NO. (예: SH7-3404-03, SM7-5356-01, H260616HD-1)
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
    type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC|JLSTW|VB09|VS)-[\w-]+', text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(0).strip()

    # 3. SPRING RATE (kgf/mm)
    sp_m = re.search(r'SPRING\s*RATE[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (kgf/mm)"] = safe_float(sp_m.group(1))

    # 4. HOT LOAD (kgf)
    hl_m = re.search(r'HOT\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (kgf)"] = safe_float(hl_m.group(1))

    # 5. COLD LOAD (kgf)
    cl_m = re.search(r'COLD\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (kgf)"] = safe_float(cl_m.group(1))

    # 6. MOVEMENT / TRAVEL (mm)
    mov_m = re.search(r'(?:THERMAL\s*MOVEMENT|MAX\.\s*TRAVEL|MOVEMENT)[^\d]*([\d\.]+)', text, re.IGNORECASE)
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
            identifier = f"{pdf_file.name} (p.{page_idx+1})"
            
            # 1단계: PDF 내부 텍스트 직접 추출 시도
            parsed = parse_text_to_fields(text, identifier)
            
            # 주요 값(TAG NO 또는 Spring Rate 등)이 비어있는 경우 2단계 고해상도 OCR 수행
            if not parsed["1. SUPPORT TAG NO"] or not parsed["3. Spring Rate (kgf/mm)"]:
                # 300 DPI 고해상도 렌더링
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                
                # 표 영역만 크롭 후 전처리
                cropped_img = crop_table_area(img)
                processed_img = preprocess_image_for_ocr(cropped_img)
                
                # OCR 실행 (PSM 6: 단일 텍스트 블록 모드 적용)
                custom_config = r'--oem 3 --psm 6'
                ocr_text = pytesseract.image_to_string(processed_img, lang='eng', config=custom_config)
                
                # 전체 이미지 추출 텍스트와 크롭 OCR 텍스트 병합 후 다시 파싱
                combined_text = text + "\n" + ocr_text
                parsed = parse_text_to_fields(combined_text, identifier + "-OCR")
                
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
