import streamlit as st
from docxtpl import DocxTemplate
from supabase import create_client
import zipfile
import io
import os
import base64
import subprocess

# 1. Supabase 접속
url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]
supabase = create_client(url, key)

# HWPX 텍스트 치환 함수
def process_hwpx(template_path, data):
    with zipfile.ZipFile(template_path, 'r') as zin:
        with io.BytesIO() as out_zip:
            with zipfile.ZipFile(out_zip, 'w') as zout:
                for item in zin.infolist():
                    buffer = zin.read(item.filename)
                    if item.filename.startswith('Contents/section') and item.filename.endswith('.xml'):
                        content = buffer.decode('utf-8')
                        for key, value in data.items():
                            content = content.replace(f"{{{{{key}}}}}", str(value))
                        buffer = content.encode('utf-8')
                    zout.writestr(item, buffer)
            return out_zip.getvalue()

# PDF 변환 함수 (LibreOffice 사용)
def convert_to_pdf(input_data, file_extension):
    temp_filename = f"temp_file{file_extension}"
    with open(temp_filename, "wb") as f:
        f.write(input_data)
    
    # LibreOffice 명령어로 PDF 변환
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', temp_filename])
    
    pdf_filename = "temp_file.pdf"
    if os.path.exists(pdf_filename):
        with open(pdf_filename, "rb") as f:
            pdf_bytes = f.read()
        return pdf_bytes
    return None

st.title("⚡ EV-CON: 실시간 서류 미리보기 시스템")

# 2. 입력 폼 (변수명 동일)
with st.form("계약_입력_폼"):
    c1, c2, c3 = st.columns(3)
    # ... (기존 입력 항목들 동일)
    사업구분 = st.text_input("사업구분", value="한국환경공단 이사장")
    아파트명 = st.text_input("아파트명")
    주소 = st.text_input("주소")
    사업자번호 = st.text_input("사업자번호")
    관리소전화 = st.text_input("관리소전화")
    설치수량 = st.number_input("설치수량", min_value=0)
    설치금액 = st.number_input("설치금액", min_value=0)
    
    col1, col2 = st.columns(2)
    미리보기_실행 = col1.form_submit_button("🔍 실시간 서류 미리보기")
    저장생성 = col2.form_submit_button("💾 DB저장 및 최종 다운로드")

데이터 = {
    "사업구분": 사업구분, "아파트명": 아파트명, "주소": 주소, "사업자번호": 사업자번호,
    "관리소전화": 관리소전화, "설치수량": 설치수량, "설치금액": 설치금액
}

# 3. 미리보기 로직 (화면에 PDF 출력)
if 미리보기_실행:
    if not 아파트명:
        st.warning("아파트명을 먼저 입력해주세요.")
    else:
        with st.spinner('서류 양식을 불러오는 중...'):
            # HWPX 생성 후 PDF로 변환
            hwpx_bin = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            pdf_bin = convert_to_pdf(hwpx_bin, ".hwpx")
            
            if pdf_bin:
                st.subheader(f"📄 {아파트명} 신청서 미리보기")
                base64_pdf = base64.b64encode(pdf_bin).decode('utf-8')
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)

# 4. 저장 및 다운로드 (기존 로직 동일)
if 저장생성:
    supabase.table("contracts").insert(데이터).execute()
    st.success("DB 저장 완료!")
    # ... (다운로드 버튼 생성)
