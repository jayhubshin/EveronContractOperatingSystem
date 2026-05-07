import streamlit as st
from docxtpl import DocxTemplate
import pandas as pd
import zipfile
import io
import os
import base64
import httpx

# ── 페이지 설정 ─────────────────────────────────────────
st.set_page_config(page_title="EV-CON", page_icon="⚡", layout="wide")
st.title("⚡ EV-CON: 에버온 계약 지원 시스템")

# ── Supabase 연결 ───────────────────────────────────────
supabase = None
try:
    from supabase import create_client
    supabase = create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_KEY"]
    )
except Exception:
    pass

# ── GitHub 템플릿 로드 ──────────────────────────────────
@st.cache_data(ttl=300)
def fetch_template(filename: str) -> bytes | None:
    # 로컬 templates/ 폴더 우선
    local = os.path.join("templates", filename)
    if os.path.exists(local):
        with open(local, "rb") as f:
            return f.read()
    # GitHub API에서 로드
    try:
        repo   = st.secrets["GITHUB_REPO"]
        token  = st.secrets["GITHUB_TOKEN"]
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        url = (
            f"https://api.github.com/repos/{repo}"
            f"/contents/templates/{filename}?ref={branch}"
        )
        headers = {
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
        }
        resp = httpx.get(url, headers=headers, timeout=20)
        if resp.status_code == 200:
            b64 = resp.json().get("content", "").replace("\n", "")
            return base64.b64decode(b64)
        else:
            st.warning(f"템플릿 로드 실패 ({filename}): {resp.status_code}")
            return None
    except Exception as e:
        st.warning(f"템플릿 로드 오류: {e}")
        return None

# ── HWPX 메일머지 ───────────────────────────────────────
def process_hwpx(template_bytes: bytes, data: dict) -> bytes | None:
    try:
        in_zip  = io.BytesIO(template_bytes)
        out_zip = io.BytesIO()
        with zipfile.ZipFile(in_zip, "r") as zin, \
             zipfile.ZipFile(out_zip, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                buf = zin.read(item.filename)
                if (item.filename.startswith("Contents/section")
                        and item.filename.endswith(".xml")):
                    content = buf.decode("utf-8")
                    for k, v in data.items():
                        display = (
                            f"{v:,}" if isinstance(v, int) and v > 999
                            else str(v)
                        )
                        content = content.replace(f"{{{{{k}}}}}", display)
                    buf = content.encode("utf-8")
                zout.writestr(item, buf)
        return out_zip.getvalue()
    except Exception as e:
        st.error(f"HWPX 오류: {e}")
        return None

# ── DOCX 메일머지 ───────────────────────────────────────
def process_docx(template_bytes: bytes, data: dict) -> bytes | None:
    try:
        import uuid
        tmp = f"/tmp/tpl_{uuid.uuid4().hex[:8]}.docx"
        with open(tmp, "wb") as f:
            f.write(template_bytes)
        doc = DocxTemplate(tmp)
        doc.render(data)
        out = io.BytesIO()
        doc.save(out)
        os.remove(tmp)
        return out.getvalue()
    except Exception as e:
        st.error(f"DOCX 오류: {e}")
        return None

# ── 사이드바 ────────────────────────────────────────────
st.sidebar.header("⚙️ 시스템 설정")
저장옵션 = st.sidebar.radio(
    "데이터 저장 방식",
    ["DB 저장 및 서류 생성", "저장 없이 서류만 생성"],
    index=1
)

st.sidebar.markdown("---")
st.sidebar.markdown("**📁 템플릿 상태**")
for fn in ["신청서_양식.hwpx", "계약서_양식.docx"]:
    tpl = fetch_template(fn)
    if tpl:
        st.sidebar.success(f"✅ {fn}")
    else:
        st.sidebar.error(f"❌ {fn}")

if st.sidebar.button("🔄 템플릿 새로고침"):
    st.cache_data.clear()
    st.rerun()

# ── 입력 폼 ─────────────────────────────────────────────
with st.form("계약입력"):
    st.subheader("📝 계약 정보 입력")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("**🏢 사업장 정보**")
        사업구분   = st.selectbox("사업구분", [
            "한국환경공단 이사장",
            "주식회사 에버온인프라",
            "기타"
        ])
        아파트명   = st.text_input("아파트명 (고유키) *")
        주소       = st.text_input("주소")
        사업자번호 = st.text_input("사업자번호")
        관리소전화 = st.text_input("관리소전화")

    with c2:
        st.markdown("**🔌 설치 정보**")
        설치수량 = st.number_input("설치수량 (기)", min_value=0, step=1, value=0)
        주차면수 = st.number_input("주차면수 (면)", min_value=0, step=1, value=0)

        단가선택 = st.selectbox("설치단가", [
            "3,500,000",
            "2,500,000",
            "직접입력"
        ])
        if 단가선택 == "직접입력":
            설치단가 = st.number_input("단가 직접입력 (원)", min_value=0, step=10000, value=0)
        else:
            설치단가 = int(단가선택.replace(",", ""))

        calc = 설치수량 * 설치단가
        st.metric("💰 자동계산 금액", f"{calc:,} 원")
        최종설치금액 = st.number_input(
            "최종 설치금액 (원)",
            min_value=0,
            value=calc
        )

    with c3:
        st.markdown("**📋 계약 조건**")
        계약년수     = st.number_input("계약년수 (년)", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간 (월)", min_value=0, value=0)
        프로모션요금 = st.number_input("프로모션요금 (원)", min_value=0, value=0)
        st.markdown("---")
        st.metric("설치 수량", f"{설치수량} 기")
        st.metric("계약 기간", f"{계약년수} 년")

    st.markdown("---")
    col1, col2 = st.columns(2)
    미리보기 = col1.form_submit_button("🔍 미리보기", use_container_width=True)
    생성실행 = col2.form_submit_button(
        "🚀 서류 생성 및 다운로드",
        use_container_width=True,
        type="primary"
    )

# ── 데이터 구성 ─────────────────────────────────────────
데이터 = {
    "사업구분":       사업구분,
    "아파트명":       아파트명,
    "주소":           주소,
    "사업자번호":     사업자번호,
    "관리소전화":     관리소전화,
    "설치수량":       설치수량,
    "주차면수":       주차면수,
    "설치단가":       설치단가,
    "설치금액":       최종설치금액,
    "계약년수":       계약년수,
    "프로모션기간":   프로모션기간,
    "프로모션요금":   프로모션요금,
    "설치단가_fmt":   f"{설치단가:,}",
    "설치금액_fmt":   f"{최종설치금액:,}",
    "프로모션요금_fmt": f"{프로모션요금:,}",
}

# ── 미리보기 ────────────────────────────────────────────
if 미리보기:
    if not 아파트명:
        st.warning("⚠️ 아파트명을 입력해주세요.")
    else:
        st.subheader("📋 입력 데이터 확인")
        rows = [
            (k, f"{v:,} 원" if isinstance(v, int) and v > 0 else str(v))
            for k, v in 데이터.items()
            if not k.endswith("_fmt")
        ]
        st.table(pd.DataFrame(rows, columns=["항목", "내용"]))

# ── 서류 생성 ───────────────────────────────────────────
if 생성실행:
    if not 아파트명:
        st.error("❌ 아파트명은 필수입니다.")
    else:
        with st.spinner("📄 서류 생성 중..."):

            # DB 저장
            if 저장옵션 == "DB 저장 및 서류 생성":
                if supabase:
                    try:
                        save = {
                            k: v for k, v in 데이터.items()
                            if not k.endswith("_fmt")
                        }
                        supabase.table("contracts").upsert(
                            save, on_conflict="아파트명"
                        ).execute()
                        st.success("✅ DB 저장 완료!")
                    except Exception as e:
                        st.error(f"DB 저장 오류: {e}")
                else:
                    st.warning("⚠️ Supabase 미연결 — 서류만 생성합니다.")

            st.subheader("📥 서류 다운로드")
            col_a, col_b = st.columns(2)

            # HWPX
            with col_a:
                tpl = fetch_template("신청서_양식.hwpx")
                if tpl:
                    result = process_hwpx(tpl, 데이터)
                    if result:
                        st.download_button(
                            label="📂 신청서 (HWPX) 다운로드",
                            data=result,
                            file_name=f"{아파트명}_신청서.hwpx",
                            mime="application/octet-stream",
                            use_container_width=True,
                        )
                    else:
                        st.error("HWPX 생성 실패")
                else:
                    st.error("신청서 템플릿 로드 실패")

            # DOCX
            with col_b:
                tpl = fetch_template("계약서_양식.docx")
                if tpl:
                    result = process_docx(tpl, 데이터)
                    if result:
                        st.download_button(
                            label="📂 계약서 (DOCX) 다운로드",
                            data=result,
                            file_name=f"{아파트명}_계약서.docx",
                            mime="application/vnd.openxmlformats-officedocument"
                                 ".wordprocessingml.document",
                            use_container_width=True,
                        )
                    else:
                        st.error("DOCX 생성 실패")
                else:
                    st.error("계약서 템플릿 로드 실패")
