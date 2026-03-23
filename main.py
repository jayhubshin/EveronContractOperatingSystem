import streamlit as st
from docxtpl import DocxTemplate
from supabase import create_client
import pandas as pd
import zipfile
import io
import os
import base64
import subprocess

# 1. Supabase 접속 설정
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)
except:
    st.error("Streamlit Secrets에 SUPABASE_URL과 SUPABASE_KEY를 설정해주세요.")

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
                            # 숫자형 데이터는 천단위 콤마 추가하여 치환
                            display_value = f"{value:,}" if isinstance(value, int) and value > 1000 else str(value)
                            content = content.replace(f"{{{{{key}}}}}", display_value)
                        buffer = content.encode('utf-8')
                    zout.writestr(item, buffer)
            return out_zip.getvalue()

# PDF 변환 함수 (미리보기용)
def convert_to_pdf(input_data, file_extension):
    temp_filename = f"temp_file{file_extension}"
    with open(temp_filename, "wb") as f:
        f.write(input_data)
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', temp_filename])
    pdf_filename = "temp_file.pdf"
    if os.path.exists(pdf_filename):
        with open(pdf_filename, "rb") as f:
            pdf_bytes = f.read()
        return pdf_bytes
    return None

st.set_page_config(page_title="EV-CON", layout="wide")
st.title("⚡ EV-CON: 에버온 계약 지원 시스템")

# 2. 12가지 항목 입력 폼
with st.form("계약_입력_폼"):
    st.subheader("📝 상세 계약 정보 입력")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        사업구분 = st.selectbox("사업구분", ["한국환경공단 이사장", "주식회사 에버온인프라", "기타"])
        아파트명 = st.text_input("아파트명")
        주소 = st.text_input("주소")
        사업자번호 = st.text_input("사업자번호")
        
    with c2:
        관리소전화 = st.text_input("관리소전화")
        설치수량 = st.number_input("설치수량", min_value=0, step=1)
        주차면수 = st.number_input("주차면수", min_value=0, step=1)
        설치단가 = st.number_input("설치단가", min_value=0, step=1000, value=3500000)
        
    with c3:
        # 설치금액 자동 계산 (수량 * 단가)
        계약년수 = st.number_input("계약년수", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간(월)", min_value=0)
        프로모션요금 = st.number_input("프로모션요금(원)", min_value=0)
        설치금액 = st.number_input("설치금액(원)", min_value=0, value=설치수량 * 설치단가)

    col1, col2 = st.columns(2)
    미리보기_실행 = col1.form_submit_button("🔍 실시간 서류 미리보기")
    저장생성 = col2.form_submit_button("💾 DB저장 및 최종 다운로드")

# DB 및 템플릿용 데이터 묶음
데이터 = {
    "사업구분": 사업구분, "아파트명": 아파트명, "주소": 주소, "사업자번호": 사업자번호,
    "관리소전화": 관리소전화, "설치수량": 설치수량, "주차면수": 주차면수, "설치단가": 설치단가,
    "설치금액": 설치금액, "계약년수": 계약년수, "프로모션기간": 프로모션기간, "프로모션요금": 프로모션요금
}

# 3. 미리보기 로직
if 미리보기_실행:
    if not 아파트명:
        st.warning("아파트명을 먼저 입력해주세요.")
    else:
        with st.spinner('미리보기 생성 중...'):
            hwpx_bin = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            pdf_bin = convert_to_pdf(hwpx_bin, ".hwpx")
            if pdf_bin:
                base64_pdf = base64.b64encode(pdf_bin).decode('utf-8')
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)

# 4. 저장 및 다운로드 (upsert 로직으로 변경)
if 저장생성:
    try:
        # '아파트명'을 기준으로 중복이면 업데이트, 없으면 삽입 (on_conflict='아파트명')
        # 주의: Supabase 테이블 설정에서 '아파트명'이 PK(Primary Key)여야 작동합니다.
        supabase.table("contracts").upsert(데이터, on_conflict="아파트명").execute()
        
        st.success(f"✅ '{아파트명}' 데이터가 업데이트(또는 저장) 되었습니다!")
        
        # 파일 생성 및 다운로드 버튼 로직
        hwpx_output = process_hwpx("templates/신청서_양식.hwpx", 데이터)
        doc = DocxTemplate("templates/계약서_양식.docx")
        doc.render(데이터)
        docx_io = io.BytesIO()
        doc.save(docx_io)
        
        d1, d2 = st.columns(2)
        d1.download_button("📂 신청서(HWP) 받기", hwpx_output, f"{아파트명}_신청서.hwpx")
        d2.download_button("📂 계약서(워드) 받기", docx_io.getvalue(), f"{아파트명}_계약서.docx")
        
    except Exception as e:
        st.error(f"오류 발생: {e}")
