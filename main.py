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
    """도면 OCR 인식률 극대화를 위한 전처리"""
    gray = img.convert('L')
    enhancer = ImageEnhance.Contrast(gray)
    enhanced = enhancer.enhance(2.5)
    threshold = 170
    binarized = enhanced.point(lambda p: 255 if p > threshold else 0)
    return binarized

def parse_drawing_text(text_content, identifier):
    """도면 표 전체 텍스트 대상 고정밀 정규식 파싱"""
    
    # 공백 정제 (줄바꿈 문자를 공백 1개로 통일하여 줄바꿈으로 인한 파싱 문제 해결)
    clean_text = " ".join(text_content.split())
    
    data = {
        "구분": identifier,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (kgf/mm)": None,
        "4. Hot Load (kgf)": None,
        "5. Cold Load (kgf)": None,
        "6. Movement (mm)": None,
    }

    # 1. HANGER MK. NO. / SUPPORT TAG NO (예: SH7-3404-03, SM7-5356-02 등)
    # HANGER MK. NO. 뒤에 나오는 태그 탐색 후 없으면 일반 패턴 매칭
    tag_m = re.search(r'HANGER\s*MK\.?\s*NO\.?[\s:]*([A-Z0-9-]+)', clean_text, re.IGNORECASE)
    if tag_m:
        data["1. SUPPORT TAG NO"] = tag_m.group(1).strip()
    else:
        tag_direct = re.search(r'\b(S[HM]\d*-\d{3,5}-\d{2,3})\b', clean_text)
        if tag_direct:
            data["1. SUPPORT TAG NO"] = tag_direct.group(1).strip()

    # 2. Type & Size (BOM 내 VT-120-09LCG, VT-060-04LCG 등)
    # O, o, C 등으로 오인식되는 문자 대응
    type_m = re.search(r'\b(VT-\d{3}-[A-Z0-9]+)\b', clean_text, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(1).strip().upper()

    # 3. SPRING RATE (kgf/mm)
    sp_m = re.search(r'SPRING\s*RATE[^\d]*(\d+\.?\d*)', clean_text, re.IGNORECASE)
    if sp_m:
        data["3. Spring Rate (kgf/mm)"] = float(sp_m.group(1))

    # 4. HOT LOAD (kgf)
    hl_m = re.search(r'HOT\s*LOAD[^\d]*(\d+\.?\d*)', clean_text, re.IGNORECASE)
    if hl_m:
        data["4. Hot Load (kgf)"] = float(hl_m.group(1))

    # 5. COLD LOAD (kgf)
    cl_m = re.search(r'COLD\s*LOAD[^\d]*(\d+\.?\d*)', clean_text, re.IGNORECASE)
    if cl_m:
        data["5. Cold Load (kgf)"] = float(cl_m.group(1))

    # 6. MOVEMENT / THERMAL MAXIMUM (+Y UP 수치 추적)
    # THERMAL MAXIMUM 이후 나타나는 +Y UP 방향 값 또는 숫자를 파싱
    mov_m = re.search(r'(?:\+Y\s*UP|UP\s*\+Y)[^\d]*(\d+\.?\d*)', clean_text, re.IGNORECASE)
    if mov_m:
        data["6. Movement (mm)"] = float(mov_m.group(1))
    else:
        # 서브패턴: THERMAL MAXIMUM 부근 수치
        mov_sub = re.search(r'THERMAL\s*MAXIMUM[^\d]*(\d+\.?\d*)', clean_text, re.IGNORECASE)
        if mov_sub:
            data["6. Movement (mm)"] = float(mov_sub.group(1))

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
            
            # 1. PyMuPDF 텍스트 extraction
            text = page.get_text()

            # 2. 고해상도 OCR 수행 (도면 특성상 OCR 텍스트 병합이 필수적임)
            pix = page.get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            processed_img = preprocess_image_for_ocr(img)
            
            # PSM 11 (Sparse Text) - 도면처럼 표와 무작위 위치 텍스트 파싱에 적합
            custom_config = r'--oem 3 --psm 11'
            ocr_text = pytesseract.image_to_string(processed_img, lang='eng', config=custom_config)

            # PDF 원본 텍스트 + OCR 텍스트 결합 후 통합 분석
            full_text = text + "\n" + ocr_text
            parsed = parse_drawing_text(full_text, identifier)

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
