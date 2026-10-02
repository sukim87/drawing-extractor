import re
import openpyxl
import pdfplumber

def extract_drawing_data(pdf_path):
    """PDF 도면에서 HANGER & SNUBBER DATA 수치를 정밀 추출하는 함수"""
    extracted_data = {
        "SUPPORT_TAG": None,
        "SPRING_RATE": None,
        "HOT_LOAD": None,
        "COLD_LOAD": None,
        "HOT_POINT": None,
        "COLD_POINT": None
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        text = page.extract_text()
        
        # 1. SUPPORT TAG NO. 추출
        tag_match = re.search(r'SUPPORT\s*TAG\.?\s*NO\.?\s*[:\s]*([A-Z0-9-]+)', text, re.IGNORECASE)
        if tag_match:
            extracted_data["SUPPORT_TAG"] = tag_match.group(1).strip()
            
        # 2. SPRING RATE (N/mm) 추출
        spring_match = re.search(r'SPRING\s*RATE\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if spring_match:
            extracted_data["SPRING_RATE"] = float(spring_match.group(1))
            
        # 3. HOT LOAD (N) 추출
        hot_load_match = re.search(r'HOT\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if hot_load_match:
            extracted_data["HOT_LOAD"] = float(hot_load_match.group(1))
            
        # 4. COLD LOAD (N) 추출
        cold_load_match = re.search(r'COLD\s*LOAD\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if cold_load_match:
            extracted_data["COLD_LOAD"] = float(cold_load_match.group(1))
            
        # 5. HOT POINT (mm) 추출
        hot_point_match = re.search(r'HOT\s*POINT\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if hot_point_match:
            extracted_data["HOT_POINT"] = float(hot_point_match.group(1))
            
        # 6. COLD POINT (mm) 추출
        cold_point_match = re.search(r'COLD\s*POINT\s*[:\s]*([\d\.]+)', text, re.IGNORECASE)
        if cold_point_match:
            extracted_data["COLD_POINT"] = float(cold_point_match.group(1))
            
    return extracted_data


def update_excel_sheet(excel_path, data):
    """추출된 데이터를 엑셀 시트에 매칭하여 기입하는 함수"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    
    tag_to_find = data["SUPPORT_TAG"]
    if not tag_to_find:
        print("❌ 도면에서 SUPPORT TAG NO를 찾지 못했습니다.")
        return

    print(f"🔍 찾은 SUPPORT TAG NO: {tag_to_find}")
    print(f"📊 추출 데이터: {data}")
    
    matched_row = None
    for row in range(5, ws.max_row + 1):
        cell_value = str(ws.cell(row=row, column=4).value or '').strip()
        if tag_to_find in cell_value or cell_value in tag_to_find:
            matched_row = row
            break
            
    if matched_row:
        print(f"✅ 엑셀 {matched_row}번 행 매칭 성공")
        
        # 엑셀 열 기입
        ws.cell(row=matched_row, column=7, value=data["SPRING_RATE"])
        ws.cell(row=matched_row, column=9, value=data["HOT_LOAD"])
        ws.cell(row=matched_row, column=11, value=data["COLD_LOAD"])
        ws.cell(row=matched_row, column=13, value=data["HOT_POINT"])
        ws.cell(row=matched_row, column=15, value=data["COLD_POINT"])
        
        output_path = "updated_" + excel_path
        wb.save(output_path)
        print(f"🎉 기입 완료! 저장 파일: {output_path}")
    else:
        print(f"⚠️ 엑셀에서 '{tag_to_find}'와 일치하는 항목을 찾지 못했습니다.")


if __name__ == "__main__":
    pdf_file = "3PAGE.pdf"
    excel_file = "5. #23,22. Spring Hanger 성능시험지(LB) .xlsx"
    
    parsed_data = extract_drawing_data(pdf_file)
    update_excel_sheet(excel_file, parsed_data)