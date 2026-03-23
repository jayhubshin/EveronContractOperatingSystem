import streamlit as st
from docxtpl import DocxTemplate
# 필요한 경우 hwp 관련 라이브러리 추가

def 실행():
    st.title("에버온 계약 지원 시스템")

    with st.form("입력창"):
        # 1. 화면 입력 (팀장님이 정하신 항목명 그대로)
        사업구분 = st.text_input("사업구분")
        아파트명 = st.text_input("아파트명")
        주소 = st.text_input("주소")
        사업자번호 = st.text_input("사업자번호")
        관리소전화 = st.text_input("관리소전화")
        설치수량 = st.number_input("설치수량", min_value=0)
        주차면수 = st.number_input("주차면수", min_value=0)
        설치단가 = st.number_input("설치단가", min_value=0)
        설치금액 = st.number_input("설치금액", min_value=0)
        계약년수 = st.number_input("계약년수", min_value=0)
        프로모션기간 = st.number_input("프로모션기간", min_value=0)
        프로모션요금 = st.number_input("프로모션요금", min_value=0)

        제출 = st.form_submit_button("파일 생성")

        if 제출:
            # 2. 데이터 묶기
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

            # 3. 워드 템플릿에 주입 (예시)
            doc = DocxTemplate("templates/계약서_양식.docx")
            doc.render(계약데이터)
            doc.save("결과_계약서.docx")
            
            st.success(f"{아파트명} 서류 생성 완료!")

if __name__ == "__main__":
    실행()
