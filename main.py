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

def extract_first_number(text):
    """문자열에서 첫 번째 소수/정수 숫자만 추출"""
    if not text:
        return None
    m = re.search(r'(\d+\.?\d*)', str(text))
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return None
    return None

def parse_drawing_text(text_content, identifier):
    """도면 표 위치 구조 기반 정밀 파싱 알고리즘"""
    data = {
        "구분": identifier,
        "1. SUPPORT TAG NO": None,
        "2. Type & Size": None,
        "3. Spring Rate (kgf/mm)": None,
        "4. Hot Load (kgf)": None,
        "5. Cold Load (kgf)": None,
        "6. Movement (mm)": None,
    }

    # 1. SUPPORT TAG NO (HANGER MK. NO.)
    tag_m = re.search(r'(?:HANGER\s*MK\.?\s*NO\.?|PURCHASER\'S\s*DOC\.?\s*NO\.?)[\s:]*([A-Z0-9-]+)', text_content, re.IGNORECASE)
    if tag_m:
        data["1. SUPPORT TAG NO"] = tag_m.group(1).strip()
    else:
        # 패턴 매칭 (예: SH7-3404-03, SM7-5356-02)
        tag_direct = re.search(r'\b[A-Z0-9]{2,4}-\d{3,5}-\d{2,3}\b', text_content)
        if tag_direct:
            data["1. SUPPORT TAG NO"] = tag_direct.group(0).strip()

    # 2. Type & Size (예: VT-120-09LCG, VT-060-04LCG 등)
    type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC|VB09|VS)-[\w-]+', text_content, re.IGNORECASE)
    if type_m:
        data["2. Type & Size"] = type_m.group(0).strip()

    # 줄 단위 분석 (아래 칸 수치 추적)
    lines = [line.strip() for line in text_content.split('\n') if line.strip()]
    
    for i, line in enumerate(lines):
        # 3. SPRING RATE
        if "SPRING RATE" in line.upper():
            # 같은 줄이나 다음 1~2줄 내에서 숫자 탐색
            search_area = " ".join(lines[i:i+3])
            val = extract_first_number(re.sub(r'SPRING\s*RATE', '', search_area, flags=re.IGNORECASE))
            if val is not None and not data["3. Spring Rate (kgf/mm)"]:
                data["3. Spring Rate (kgf/mm)"] = val

        # 4. HOT LOAD
        if "HOT LOAD" in line.upper():
            search_area = " ".join(lines[i:i+3])
            val = extract_first_number(re.sub(r'HOT\s*LOAD', '', search_area, flags=re.IGNORECASE))
            if val is not None and not data["4. Hot Load (kgf)"]:
                data["4. Hot Load (kgf)"] = val

        # 5. COLD LOAD
        if "COLD LOAD" in line.upper():
            search_area = " ".join(lines[i:i+3])
            val = extract_first_number(re.sub(r'COLD\s*LOAD', '', search_area, flags=re.IGNORECASE))
            if val is not None and not data["5. Cold Load (kgf)"]:
                data["5. Cold Load (kgf)"] = val

        # 6. MOVEMENT (+Y UP 행 추적)
        if "+Y" in line.upper() or "UP" in line.upper():
            val = extract_first_number(line)
            if val is not None and not data["6. Movement (mm)"]:
                data["6. Movement (mm)"] = val

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

            # 텍스트 추출 시도
            parsed = parse_drawing_text(text, identifier)

            # 값이 비어있을 경우 OCR 수행
            if not parsed["1. SUPPORT TAG NO"] or not parsed["3. Spring Rate (kgf/mm)"]:
                pix = page.get_pixmap(dpi=300)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                processed_img = preprocess_image_for_ocr(img)
                
                custom_config = r'--oem 3 --psm 6'
                ocr_text = pytesseract.image_to_string(processed_img, lang='eng', config=custom_config)
                
                parsed = parse_drawing_text(text + "\n" + ocr_text, identifier)

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
