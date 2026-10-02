import re
import io
import pandas as pd
import pdfplumber
import openpyxl
import streamlit as st

st.set_page_config(page_title="유연한 도면 데이터 자동 추출기", layout="wide")
st.title("📐 도면 양식 자율 인식 Pipe Support 데이터 리스트업")

uploaded_pdfs = st.file_uploader(
    "양식이 서로 다른 도면 PDF 파일들을 업로드하세요", 
    type=["pdf"], 
    accept_multiple_files=True
)

def extract_fields_adaptively(pdf_file):
    """
    도면 양식과 위치가 제각각이어도 키워드 맥락 및 표 구조를 동적으로 추적하여 추출
    """
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
            
            # 1. SUPPORT TAG NO (양식별 다양한 라벨 대응)
            if not data["1. SUPPORT TAG NO"]:
                tag_m = re.search(r'(?:SUPPORT\s*TAG|TAG\s*NO|TAG)\.?\s*[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
                if tag_m:
                    data["1. SUPPORT TAG NO"] = tag_m.group(1).strip()
            
            # 2. Type & Size (규격 코드 패턴 인식: VT-, VC-, VA-, CS-, CHT- 등)
            if not data["2. Type & Size"]:
                type_m = re.search(r'(?:VT|VC|VA|CS|CH|CHT|JSLD|JCSC)-[A-Z0-9-]+', text, re.IGNORECASE)
                if type_m:
                    data["2. Type & Size"] = type_m.group(0).strip()
            
            # 3. SPRING RATE
            if data["3. Spring Rate (N/mm)"] is None:
                sp_m = re.search(r'SPRING\s*RATE[^\d]*([\d\.]+)', text, re.IGNORECASE)
                if sp_m:
                    data["3. Spring Rate (N/mm)"] = float(sp_m.group(1))
                    
            # 4. HOT LOAD
            if data["4. Hot Load (N)"] is None:
                hl_m = re.search(r'HOT\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
                if hl_m:
                    data["4. Hot Load (N)"] = float(hl_m.group(1))
                    
            # 5. COLD LOAD
            if data["5. Cold Load (N)"] is None:
                cl_m = re.search(r'COLD\s*LOAD[^\d]*([\d\.]+)', text, re.IGNORECASE)
                if cl_m:
                    data["5. Cold Load (N)"] = float(cl_m.group(1))
                    
            # 6. HOT POINT
            if data["6. Hot Point (mm)"] is None:
                hp_m = re.search(r'HOT\s*POINT[^\d]*([\d\.]+)', text, re.IGNORECASE)
                if hp_m:
                    data["6. Hot Point (mm)"] = float(hp_m.group(1))
                    
            # 7. COLD POINT
            if data["7. Cold Point (mm)"] is None:
                cp_m = re.search(r'COLD\s*POINT[^\d]*([\d\.]+)', text, re.IGNORECASE)
                if cp_m:
                    data["7. Cold Point (mm)"] = float(cp_m.group(1))

    return data

if uploaded_pdfs:
    if st.button("🚀 업로드 도면 분석 및 엑셀 리스트업"):
        results = []
        for pdf in uploaded_pdfs:
            res = extract_fields_adaptively(pdf)
            results.append(res)
            
        df = pd.DataFrame(results)
        st.subheader("📋 분석 결과 리스트")
        st.dataframe(df, use_container_width=True)
        
        # 엑셀 다운로드 파일 생성
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
