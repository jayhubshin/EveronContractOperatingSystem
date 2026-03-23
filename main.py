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

# [함수] HWPX 내부 텍스트 치환 (한글 양식용)
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
                                # 숫자형 데이터는 보기 좋게 콤마 추가
                                disp = f"{value:,}" if isinstance(value, int) and value > 999 else str(value)
                                content = content.replace(f"{{{{{key}}}}}", disp)
                            buffer = content.encode('utf-8')
                        zout.writestr(item, buffer)
                return out_zip.getvalue()
    except FileNotFoundError:
        st.error(f"템플릿 파일을 찾을 수 없습니다: {template_path}")
        return None

# [함수] PDF 변환 (실시간 미리보기용)
def convert_to_pdf(input_data, file_extension):
    import uuid
    uid = str(uuid.uuid4())[:8]
    tmp_in = f"temp_{uid}{file_extension}"
    tmp_out = f"temp_{uid}.pdf"
    
    try:
        with open(tmp_in, "wb") as f:
            f.write(input_data)
        # LibreOffice 명령어로 PDF 변환 (outdir 지정 필수)
        subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', tmp_in, '--outdir', '.'], check=True)
        
        if os.path.exists(tmp_out):
            with open(tmp_out, "rb") as f:
                pdf_bytes = f.read()
            os.remove(tmp_in)
            os.remove(tmp_out)
            return pdf_bytes
    except:
        return None
    return None

# --- UI 시작 ---
st.set_page_config(page_title="EV-CON", layout="wide")
st.title("⚡ EV-CON: 에버온 계약 지원 시스템")

# 2. 상단 옵션 설정 (라디오 버튼)
st.sidebar.header("⚙️ 시스템 설정")
저장옵션 = st.sidebar.radio(
    "데이터 저장 방식",
    ["DB 저장 및 서류 생성", "저장 없이 서류만 생성"],
    index=0,
    help="기존 데이터 덮어쓰기를 원하시면 'DB 저장'을 선택하세요."
)

# 3. 12가지 항목 입력 폼
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
        설치수량 = st.number_input("설치수량 (기)", min_value=0, step=1)
        주차면수 = st.number_input("주차면수 (면)", min_value=0, step=1)
        설치단가 = st.number_input("설치단가 (원)", min_value=0, step=1000, value=3500000)
        
    with c3:
        계약년수 = st.number_input("계약년수 (년)", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간 (월)", min_value=0)
        프로모션요금 = st.number_input("프로모션요금 (원)", min_value=0)
        설치금액 = st.number_input("최종 설치금액 (원)", min_value=0, value=설치수량 * 설치단가)

    col_btn1, col_btn2 = st.columns(2)
    미리보기_실행 = col_btn1.form_submit_button("🔍 입력 내용 및 서류 미리보기")
    생성_실행 = col_btn2.form_submit_button("🚀 서류 생성 및 저장 실행")

# 데이터 매핑 (모든 변수 정제)
데이터 = {
    "사업구분": 사업구분, "아파트명": 아파트명, "주소": 주소, "사업자번호": 사업자번호,
    "관리소전화": 관리소전화, "설치수량": 설치수량, "주차면수": 주차면수, "설치단가": 설치단가,
    "설치금액": 설치금액, "계약년수": 계약년수, "프로모션기간": 프로모션기간, "프로모션요금": 프로모션요금
}

# 4. 미리보기 로직 (신청서 + 계약서 통합)
if 미리보기_실행:
    if not 아파트명:
        st.warning("아파트명을 입력해야 미리보기가 가능합니다.")
    else:
        st.subheader("📋 입력 데이터 요약")
        st.table(pd.Series(데이터, name="내용"))
        
        with st.spinner('실시간 서류 양식(신청서 & 계약서) 생성 중...'):
            # --- [1] 신청서 (HWPX -> PDF) 변환 ---
            hwpx_bin = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            pdf_hwpx = convert_to_pdf(hwpx_bin, ".hwpx") if hwpx_bin else None
            
            # --- [2] 계약서 (DOCX -> PDF) 변환 ---
            doc = DocxTemplate("templates/계약서_양식.docx")
            doc.render(데이터)
            docx_io = io.BytesIO()
            doc.save(docx_io)
            pdf_docx = convert_to_pdf(docx_io.getvalue(), ".docx")
            
            # --- 화면 출력 ---
            t1, t2 = st.tabs(["📄 1. 보조금 신청서 미리보기", "📄 2. 설치운영 계약서 미리보기"])
            
            with t1:
                if pdf_hwpx:
                    base64_pdf1 = base64.b64encode(pdf_hwpx).decode('utf-8')
                    pdf_display1 = f'<embed src="data:application/pdf;base64,{base64_pdf1}" width="100%" height="800" type="application/pdf">'
                    st.markdown(pdf_display1, unsafe_allow_html=True)
                else:
                    st.info("💡 신청서 미리보기를 불러올 수 없습니다. 다운로드 후 확인해주세요.")
            
            with t2:
                if pdf_docx:
                    base64_pdf2 = base64.b64encode(pdf_docx).decode('utf-8')
                    pdf_display2 = f'<embed src="data:application/pdf;base64,{base64_pdf2}" width="100%" height="800" type="application/pdf">'
                    st.markdown(pdf_display2, unsafe_allow_html=True)
                else:
                    st.info("💡 계약서 미리보기를 불러올 수 없습니다. 다운로드 후 확인해주세요.")

# 5. 생성 및 저장 로직
if 생성_실행:
    if not 아파트명:
        st.error("아파트명은 필수 입력 항목입니다.")
    else:
        try:
            # DB 저장 분기
            if 저장옵션 == "DB 저장 및 서류 생성":
                supabase.table("contracts").upsert(데이터, on_conflict="아파트명").execute()
                st.success(f"✅ '{아파트명}' 데이터가 DB에 업데이트 되었습니다.")
            else:
                st.info("ℹ️ DB 저장 없이 파일만 생성합니다.")

            # 파일 생성 (HWPX)
            hwpx_out = process_hwpx("templates/신청서_양식.hwpx", 데이터)
            
            # 파일 생성 (DOCX) - 워드 내부 변수 정제 주입
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
            st.error(f"실행 중 오류가 발생했습니다: {e}")
