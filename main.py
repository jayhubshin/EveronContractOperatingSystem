import streamlit as st
from docxtpl import DocxTemplate
from supabase import create_client
import pandas as pd
import zipfile
import io

# 1. Supabase 접속 (Secrets 설정값 불러오기)
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)
except:
    st.error("Supabase 접속 정보를 Secrets에 먼저 설정해주세요.")

# [핵심] HWPX 파일 내부의 텍스트를 교체하는 함수
def process_hwpx(template_path, data):
    with zipfile.ZipFile(template_path, 'r') as zin:
        with io.BytesIO() as out_zip:
            with zipfile.ZipFile(out_zip, 'w') as zout:
                for item in zin.infolist():
                    buffer = zin.read(item.filename)
                    # 실제 본문 데이터가 들어있는 XML 파일들만 골라 텍스트 치환
                    if item.filename.startswith('Contents/section') and item.filename.endswith('.xml'):
                        content = buffer.decode('utf-8')
                        for key, value in data.items():
                            # {{변수명}}을 실제 입력값으로 교체
                            content = content.replace(f"{{{{{key}}}}}", str(value))
                        buffer = content.encode('utf-8')
                    zout.writestr(item, buffer)
            return out_zip.getvalue()

st.set_page_config(page_title="EV-CON", layout="wide")
st.title("⚡ EV-CON: 에버온 계약 지원 시스템")

# 2. 입력 화면 구성
with st.form("계약_데이터_폼"):
    st.subheader("📝 계약 정보 입력")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        사업구분 = st.text_input("사업구분", value="한국환경공단 이사장")
        아파트명 = st.text_input("아파트명")
        주소 = st.text_input("주소")
        사업자번호 = st.text_input("사업자번호")
        
    with c2:
        관리소전화 = st.text_input("관리소전화")
        설치수량 = st.number_input("설치수량", min_value=0, step=1)
        주차면수 = st.number_input("주차면수", min_value=0, step=1)
        설치단가 = st.number_input("설치단가", min_value=0, step=1000)
        
    with c3:
        설치금액 = st.number_input("설치금액", min_value=0)
        계약년수 = st.number_input("계약년수", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간", min_value=0)
        프로모션요금 = st.number_input("프로모션요금", min_value=0)
    
    col1, col2 = st.columns(2)
    미리보기 = col1.form_submit_button("🔍 입력 내용 확인")
    저장생성 = col2.form_submit_button("💾 DB저장 및 서류 생성")

# 변수 매핑 (한글 변수명 그대로 사용)
데이터 = {
    "사업구분": 사업구분, "아파트명": 아파트명, "주소": 주소, "사업자번호": 사업자번호,
    "관리소전화": 관리소전화, "설치수량": 설치수량, "주차면수": 주차면수, "설치단가": 설치단가,
    "설치금액": 설치금액, "계약년수": 계약년수, "프로모션기간": 프로모션기간, "프로모션요금": 프로모션요금
}

# 3. 미리보기 로직
if 미리보기:
    st.info("아래 내용으로 서류가 작성됩니다.")
    st.table(pd.Series(데이터, name="입력값"))

# 4. 저장 및 파일 생성 로직
if 저장생성:
    if not 아파트명:
        st.error("아파트명을 입력해야 저장할 수 있습니다.")
    else:
        try:
            # A. Supabase 저장
            supabase.table("contracts").insert(데이터).execute()
            
            # B. 한글(HWPX) 신청서 생성
            hwpx_data = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            
            # C. 워드(DOCX) 계약서 생성
            doc = DocxTemplate("templates/계약서_양식.docx")
            doc.render(데이터)
            docx_io = io.BytesIO()
            doc.save(docx_io)
            
            st.success(f"✅ {아파트명} 계약 정보 저장 및 서류 생성 완료!")
            
            # 다운로드 버튼
            d1, d2 = st.columns(2)
            d1.download_button("📂 신청서(HWP) 다운로드", data=hwpx_data, file_name=f"{아파트명}_신청서.hwpx")
            d2.download_button("📂 계약서(워드) 다운로드", data=docx_io.getvalue(), file_name=f"{아파트명}_계약서.docx")
            
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
