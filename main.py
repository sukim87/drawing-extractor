import re
import io
import pandas as pd
import pdfplumber
import openpyxl
import streamlit as st

st.set_page_config(page_title="도면 데이터 자동 추출기", layout="wide")
st.title("📐 Pipe Support 도면 일괄 분석 및 엑셀 리스트업")

# 1. 파일 업로드 (여러 도면 동시 선택 가능)
uploaded_pdfs = st.file_uploader(
    "도면 PDF 파일들을 업로드하세요 (여러 개 선택 가능)", 
    type=["pdf"], 
    accept_multiple_files=True
)

def extract_7_fields(pdf_file):
    """PDF 도면에서 요구되는 7가지 핵심 항목을 정밀 추출하는 함수"""
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
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            
            # 1. SUPPORT TAG NO.
            if not data["1. SUPPORT TAG NO"]:
                tag_match = re.search(r'SUPPORT\s*TAG\.?\s*NO\.?\s*[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
                if tag_match:
                    data["1. SUPPORT TAG NO"] = tag_match.group(1).strip()
            
            # 2. Type & Size (서포트 모델 타입)
            if not data["2. Type & Size"]:
                type_match = re.search(r'(?:VT|VC|VA|CS|CH|CHT)-[\w-]+', text)
                if type_match:
                    data["2. Type & Size"] = type_match.group(0).strip()
            
            # 3. SPRING RATE
            if data["3. Spring Rate (N/mm)"] is None:
                spring_match = re.search(r'SPRING\s*RATE\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
                if spring_match:
                    data["3. Spring Rate (N/mm)"] = float(spring_match.group(1))
            
            # 4. HOT LOAD
            if data["4. Hot Load (N)"] is None:
                hot_load_match = re.search(r'HOT\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
                if hot_load_match:
                    data["4. Hot Load (N)"] = float(hot_load_match.group(1))
            
            # 5. COLD LOAD
            if data["5. Cold Load (N)"] is None:
                cold_load_match = re.search(r'COLD\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
                if cold_load_match:
                    data["5. Cold Load (N)"] = float(cold_load_match.group(1))
            
            # 6. HOT POINT
            if data["6. Hot Point (mm)"] is None:
                hot_point_match = re.search(r'HOT\s*POINT\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
                if hot_point_match:
                    data["6. Hot Point (mm)"] = float(hot_point_match.group(1))
            
            # 7. COLD POINT
            if data["7. Cold Point (mm)"] is None:
                cold_point_match = re.search(r'COLD\s*POINT\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
                if cold_point_match:
                    data["7. Cold Point (mm)"] = float(cold_point_match.group(1))

    return data

# 2. 실행 및 결과 다운로드
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
        
        # 메모리 상에서 엑셀 파일 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Support_List')
        excel_data = output.getvalue()
        
        # 3. 엑셀 다운로드 버튼
        st.download_button(
            label="📥 엑셀 리스트 다운로드 (.xlsx)",
            data=excel_data,
            file_name="Pipe_Support_Drawing_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
