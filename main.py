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
    st.error("Streamlit Secrets에 접속 정보(URL, KEY)를 설정해주세요.")

# [함수] HWPX 텍스트 치환 (한글 양식용)
def process_hwpx(template_path, data):
    try:
        with zipfile.ZipFile(template_path, 'r') as zin:
            with io.BytesIO() as out_zip:
                with zipfile.ZipFile(out_zip, 'w') as zout:
                    for item in zin.infolist():
                        buffer = zin.read(item.filename)
                        if item.filename.startswith('Contents/section') and item.filename.endswith('.xml'):
                            content = buffer.decode('utf-8')
                            for key, value in data.items():
                                # 숫자 데이터는 1,000 단위 콤마 추가
                                disp = f"{value:,}" if isinstance(value, int) and value > 999 else str(value)
                                content = content.replace(f"{{{{{key}}}}}", disp)
                            buffer = content.encode('utf-8')
                        zout.writestr(item, buffer)
                return out_zip.getvalue()
    except: return None

# [함수] PDF 변환 (미리보기용)
def convert_to_pdf(input_data, file_extension):
    import uuid
    uid = str(uuid.uuid4())[:8]
    tmp_in, tmp_out = f"temp_{uid}{file_extension}", f"temp_{uid}.pdf"
    try:
        with open(tmp_in, "wb") as f: f.write(input_data)
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', tmp_in, '--outdir', '.'], check=True)
        if os.path.exists(tmp_out):
            with open(tmp_out, "rb") as f: pdf_bytes = f.read()
            os.remove(tmp_in); os.remove(tmp_out)
            return pdf_bytes
    except: return None

# --- UI 설정 ---
st.set_page_config(page_title="EV-CON", layout="wide")
st.title("⚡ EV-CON: 에버온 계약 지원 시스템")

# 2. 사이드바 설정 (디폴트: 저장 없이 서류만 생성)
st.sidebar.header("⚙️ 시스템 설정")
저장옵션 = st.sidebar.radio(
    "데이터 저장 방식",
    ["DB 저장 및 서류 생성", "저장 없이 서류만 생성"],
    index=1, # 1번 인덱스(저장 없이)가 기본 선택
    help="확정된 계약 건만 'DB 저장'을 선택하세요."
)

# 3. 입력 폼
with st.form("계약_입력_폼"):
    st.subheader("📝 상세 계약 정보 입력")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        사업구분 = st.selectbox("사업구분", ["한국환경공단 이사장", "주식회사 에버온인프라", "기타"])
        아파트명 = st.text_input("아파트명 (고유키)", placeholder="중복 시 기존 데이터 덮어씀")
        주소 = st.text_input("주소")
        사업자번호 = st.text_input("사업자번호")
        
    with c2:
        관리소전화 = st.text_input("관리소전화")
        설치수량 = st.number_input("설치수량 (기)", min_value=0, step=1, value=0)
        주차면수 = st.number_input("주차면수 (면)", min_value=0, step=1, value=0)
        
        # [수정] 단가 선택 상태 유지 로직
        단가_옵션 = ["3,500,000", "2,500,000", "직접입력"]
        
        # selectbox에 key를 부여하여 상태를 강제 고정합니다.
        단가선택 = st.selectbox("설치단가 선택", 단가_옵션, index=0, key="price_select")
        
        if 단가선택 == "직접입력":
            # 직접 입력 시에도 값이 초기화되지 않도록 key를 부여합니다.
            설치단가 = st.number_input("단가 직접 입력 (원)", min_value=0, step=10000, key="custom_price")
        else:
            설치단가 = int(단가선택.replace(",", ""))
        
    with c3:
        계약년수 = st.number_input("계약년수 (년)", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간 (월)", min_value=0)
        프로모션요금 = st.number_input("프로모션요금 (원)", min_value=0)
        
        # [자동계산] 설치금액 = 설치수량 * 설치단가
        계산금액 = 설치수량 * 설치단가
        설치금액 = st.number_input("최종 설치금액 (원)", min_value=0, value=계산금액)

    col_btn1, col_btn2 = st.columns(2)
    미리보기_실행 = col_btn1.form_submit_button("🔍 서류 미리보기 (PDF)")
    생성_실행 = col_btn2.form_submit_button("🚀 서류 생성 및 저장 실행")

# 데이터 매핑
데이터 = {
    "사업구분": 사업구분, "아파트명": 아파트명, "주소": 주소, "사업자번호": 사업자번호,
    "관리소전화": 관리소전화, "설치수량": 설치수량, "주차면수": 주차면수, "설치단가": 설치단가,
    "설치금액": 설치금액, "계약년수": 계약년수, "프로모션기간": 프로모션기간, "프로모션요금": 프로모션요금
}

# 4. 미리보기 (웨일 브라우저 차단 대비 요약표 병행)
if 미리보기_실행:
    if not 아파트명:
        st.warning("아파트명을 입력해주세요.")
    else:
        st.subheader("📋 입력 데이터 확인")
        st.table(pd.Series(데이터, name="내용"))
        
        with st.spinner('미리보기 생성 중...'):
            hwpx_bin = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            pdf_bin = convert_to_pdf(hwpx_bin, ".hwpx") if hwpx_bin else None
            
            if pdf_bin:
                base64_pdf = base64.b64encode(pdf_bin).decode('utf-8')
                st.markdown(f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>', unsafe_allow_html=True)
            else:
                st.info("💡 브라우저 환경에 따라 미리보기가 제한될 수 있습니다. 아래 버튼으로 서류를 생성해 확인하세요.")

# 5. 생성 및 저장 로직
if 생성_실행:
    if not 아파트명:
        st.error("아파트명은 필수 입력 항목입니다.")
    else:
        try:
            # DB 저장 분기 (Upsert: 중복 시 덮어쓰기)
            if 저장옵션 == "DB 저장 및 서류 생성":
                supabase.table("contracts").upsert(데이터, on_conflict="아파트명").execute()
                st.success(f"✅ '{아파트명}' 데이터가 DB에 업데이트 되었습니다.")
            else:
                st.info("ℹ️ DB 저장 없이 서류 파일만 생성합니다.")

            # 파일 생성 (HWPX)
            hwpx_out = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            
            # 파일 생성 (DOCX)
            doc = DocxTemplate("templates/계약서_양식.docx")
            doc.render(데이터)
            docx_io = io.BytesIO()
            doc.save(docx_io)
            
            st.write("---")
            st.subheader("📥 서류 다운로드")
            d1, d2 = st.columns(2)
            if hwpx_out:
                d1.download_button("📂 신청서(HWP) 받기", hwpx_out, f"{아파트명}_신청서.hwpx")
            d2.download_button("📂 계약서(워드) 받기", docx_io.getvalue(), f"{아파트명}_계약서.docx")
            
        except Exception as e:
            st.error(f"실행 중 오류 발생: {e}")
