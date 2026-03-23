import streamlit as st
from docxtpl import DocxTemplate
from supabase import create_client
import pandas as pd
import os
import io

# 1. Supabase 접속 설정 (Streamlit Secrets 활용)
try:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    supabase = create_client(url, key)
except Exception as e:
    st.error("Supabase 접속 정보를 확인해주세요. (Secrets 설정 필요)")

# 페이지 설정
st.set_page_config(page_title="EV-CON: 계약 지원 시스템", layout="wide")

st.title("⚡ EV-CON: 전기차 충전기 계약 자동화")
st.info("입력하신 정보는 DB에 저장되며, 신청서(HWP형식)와 계약서(DOCX)가 자동으로 생성됩니다.")

# 2. 입력 폼 (팀장님 요청 한글 변수명 적용)
with st.form("계약서_입력_폼"):
    st.subheader("📝 기본 정보 및 계약 조건 입력")
    c1, c2, c3 = st.columns(3)
    
    with c1:
        사업구분 = st.selectbox("사업구분", ["한국환경공단 이사장", "주식회사 에버온인프라", "기타"])
        아파트명 = st.text_input("아파트명", placeholder="예: 래미안아름숲아파트")
        주소 = st.text_input("주소", placeholder="도로명 주소 입력")
        사업자번호 = st.text_input("사업자번호", placeholder="000-00-00000")
        
    with c2:
        관리소전화 = st.text_input("관리소전화", placeholder="02-000-0000")
        설치수량 = st.number_input("설치수량 (기)", min_value=0, value=0)
        주차면수 = st.number_input("주차면수 (면)", min_value=0, value=0)
        설치단가 = st.number_input("설치단가 (원)", min_value=0, value=3500000, step=100000)
        
    with c3:
        # 설치금액 자동 계산 표시 (참고용)
        자동계산_금액 = 설치수량 * 설치단가
        설치금액 = st.number_input("최종 설치금액 (원)", min_value=0, value=자동계산_금액)
        계약년수 = st.number_input("계약년수 (년)", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간 (월)", min_value=0, value=6)
        프로모션요금 = st.number_input("프로모션요금 (원)", min_value=0, value=0)

    # 폼 하단 버튼
    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        미리보기 = st.form_submit_button("🔍 입력 내용 미리보기")
    with col_btn2:
        저장_및_생성 = st.form_submit_button("💾 DB 저장 및 서류 생성")

# 데이터 꾸러미 생성
계약데이터 = {
    "사업구분": 사업구분,
    "아파트명": 아파트명,
    "주소": 주소,
    "사업자번호": 사업자번호,
    "관리소전화": 관리소전화,
    "설치수량": 설치수량,
    "주차면수": 주차면수,
    "설치단가": 설치단가,
    "설치금액": 설치금액,
    "계약년수": 계약년수,
    "프로모션기간": 프로모션기간,
    "프로모션요금": 프로모션요금
}

# 3. 미리보기 기능
if 미리보기:
    if not 아파트명:
        st.warning("아파트명을 입력해주세요.")
    else:
        st.write("---")
        st.subheader("👀 작성 내용 확인")
        df_preview = pd.DataFrame([계약데이터]).T
        df_preview.columns = ["입력값"]
        st.table(df_preview)

# 4. 저장 및 파일 생성 로직
if 저장_및_생성:
    if not 아파트명:
        st.error("아파트명은 필수 입력 항목입니다.")
    else:
        try:
            # A. Supabase DB 저장
            supabase.table("contracts").insert(계약데이터).execute()
            st.success(f"✅ {아파트명} 데이터가 성공적으로 DB에 저장되었습니다.")

            # B. 서류 생성 (워드 템플릿 방식)
            # 템플릿 파일이 templates 폴더 안에 있어야 합니다.
            # 1. 신청서 생성
            doc1 = DocxTemplate("templates/신청서_양식.docx")
            doc1.render(계약데이터)
            doc1_io = io.BytesIO()
            doc1.save(doc1_io)
            doc1_io.seek(0)

            # 2. 계약서 생성
            doc2 = DocxTemplate("templates/계약서_양식.docx")
            doc2.render(계약데이터)
            doc2_io = io.BytesIO()
            doc2.save(doc2_io)
            doc2_io.seek(0)

            st.write("---")
            st.subheader("📥 서류 다운로드")
            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    label="📄 1. 보조금 신청서 다운로드",
                    data=doc1_io,
                    file_name=f"{아파트명}_신청서.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            with d2:
                st.download_button(
                    label="📄 2. 설치운영 계약서 다운로드",
                    data=doc2_io,
                    file_name=f"{아파트명}_계약서.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            st.balloons()
            
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
