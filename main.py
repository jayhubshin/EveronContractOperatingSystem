import streamlit as st
from docxtpl import DocxTemplate
import zipfile
import io
import os
import base64
import httpx

# ── 페이지 설정 ─────────────────────────────────────────
st.set_page_config(page_title="EV-CON", page_icon="⚡", layout="centered")
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
    local = os.path.join("templates", filename)
    if os.path.exists(local):
        with open(local, "rb") as f:
            return f.read()
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

# ── 국세청 사업자 조회 ──────────────────────────────────
def lookup_business(biz_no: str) -> dict | None:
    try:
        api_key = st.secrets.get("NTS_API_KEY", "")
        if not api_key:
            st.warning("⚠️ NTS_API_KEY가 Secrets에 없습니다.")
            return None

        biz_no_clean = biz_no.replace("-", "").replace(" ", "").strip()
        if len(biz_no_clean) != 10:
            st.warning("⚠️ 사업자번호 10자리를 확인해주세요.")
            return None

        url  = "https://api.odcloud.kr/api/nts-businessman/v1/status"
        body = {"b_no": [biz_no_clean]}

        resp = httpx.post(
            url,
            params={"serviceKey": api_key},
            json=body,
            headers={
                "Content-Type": "application/json",
                "Accept": "application/json",
            },
            timeout=10
        )

        if resp.status_code == 200:
            items = resp.json().get("data", [])
            if not items:
                return None
            item     = items[0]
            상태코드  = item.get("b_stt_cd", "")
            상태_map  = {
                "01": "✅ 계속사업자",
                "02": "⚠️ 휴업자",
                "03": "❌ 폐업자"
            }
            return {
                "상태":   상태_map.get(상태코드, "알 수 없음"),
                "상태코드": 상태코드,
            }
        else:
            st.error(f"API 오류: {resp.status_code} / {resp.text}")
            return None

    except Exception as e:
        st.error(f"사업자 조회 오류: {e}")
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
                        content = content.replace(f"{{{{{k}}}}}", str(v))
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

# ── 세션 초기화 ─────────────────────────────────────────
if "biz_상호" not in st.session_state:
    st.session_state["biz_상호"] = ""
if "biz_주소" not in st.session_state:
    st.session_state["biz_주소"] = ""

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

# ════════════════════════════════════════════════════════
#  입력 섹션 (폼 없이 실시간)
# ════════════════════════════════════════════════════════

st.subheader("📝 계약 정보 입력")
st.divider()

# ── 1. 사업장 정보 ──────────────────────────────────────
st.markdown("**🏢 사업장 정보**")

사업구분 = st.selectbox("사업구분", [
    "한국환경공단 이사장",
    "주식회사 에버온인프라",
    "기타"
])

# 사업자번호 + 조회버튼
col_biz, col_btn = st.columns([3, 1])
with col_biz:
    사업자번호 = st.text_input(
        "사업자번호",
        placeholder="예) 123-45-67890"
    )
with col_btn:
    st.markdown("<br/>", unsafe_allow_html=True)
    조회버튼 = st.button("🔍 조회", use_container_width=True)

# 조회 실행
if 조회버튼:
    if not 사업자번호:
        st.warning("사업자번호를 입력해주세요.")
    else:
        with st.spinner("국세청 조회 중..."):
            result = lookup_business(사업자번호)
            if result:
                상태코드 = result.get("상태코드", "")
                if 상태코드 == "01":
                    st.success(f"{result.get('상태')} — 상호/주소를 직접 입력해주세요.")
                elif 상태코드 == "02":
                    st.warning(f"{result.get('상태')}")
                elif 상태코드 == "03":
                    st.error(f"{result.get('상태')} — 계약 불가 사업자입니다.")
            else:
                st.error("❌ 조회 결과가 없습니다.")

아파트명 = st.text_input(
    "아파트명 *",
    value=st.session_state["biz_상호"],
    placeholder="예) 래미안 강남 1단지"
)

주소 = st.text_input(
    "주소",
    value=st.session_state["biz_주소"],
    placeholder="예) 서울특별시 강남구 테헤란로 123"
)

관리소전화 = st.text_input(
    "관리소전화",
    placeholder="예) 02-1234-5678"
)

st.divider()

# ── 2. 설치 정보 ────────────────────────────────────────
st.markdown("**🔌 설치 정보**")

설치수량 = st.number_input(
    "설치수량 (기)",
    min_value=0, step=1, value=0
)

주차면수 = st.number_input(
    "주차면수 (면)",
    min_value=0, step=1, value=0
)

단가선택 = st.selectbox("설치단가", [
    "3,500,000",
    "2,500,000",
    "직접입력"
])

if 단가선택 == "직접입력":
    설치단가 = st.number_input(
        "단가 직접입력 (원)",
        min_value=0, step=10000, value=0
    )
else:
    설치단가 = int(단가선택.replace(",", ""))

# 자동계산 표시
calc = 설치수량 * 설치단가
st.info(f"💡 자동계산: {설치수량}기 × {설치단가:,}원 = **{calc:,}원**")

최종설치금액 = st.number_input(
    "최종 설치금액 (원)",
    min_value=0,
    value=calc,
    step=1,
    format="%d"        # ← 정수 표시
)

# 콤마 포함 표시
st.markdown(f"**💰 {최종설치금액:,} 원**")


st.divider()

# ── 3. 계약 조건 ────────────────────────────────────────
st.markdown("**📋 계약 조건**")

계약년수     = st.number_input("계약년수 (년)", min_value=0, value=7)
프로모션기간 = st.number_input("프로모션기간 (월)", min_value=0, value=0)
프로모션요금 = st.number_input("프로모션요금 (원)", min_value=0, value=0)

st.divider()

# ── 데이터 구성 ─────────────────────────────────────────
데이터 = {
    "사업구분":       사업구분,
    "아파트명":       아파트명,
    "주소":           주소,
    "사업자번호":     사업자번호,
    "관리소전화":     관리소전화,
    "설치수량":       설치수량,
    "주차면수":       주차면수,
    "설치단가":       f"{설치단가:,}",
    "설치금액":       f"{최종설치금액:,}",
    "계약년수":       계약년수,
    "프로모션기간":   프로모션기간,
    "프로모션기간월": 프로모션기간,
    "프로모션요금":   f"{프로모션요금:,}",
    "프로모션요금원": f"{프로모션요금:,}",
}

# ── 서류 생성 버튼 ──────────────────────────────────────
생성실행 = st.button(
    "🚀 서류 생성 및 다운로드",
    use_container_width=True,
    type="primary"
)

if 생성실행:
    if not 아파트명:
        st.error("❌ 아파트명은 필수입니다.")
    else:
        with st.spinner("📄 서류 생성 중..."):

            # DB 저장
            if 저장옵션 == "DB 저장 및 서류 생성":
                if supabase:
                    try:
                        supabase.table("contracts").upsert(
                            데이터, on_conflict="아파트명"
                        ).execute()
                        st.success("✅ DB 저장 완료!")
                    except Exception as e:
                        st.error(f"DB 저장 오류: {e}")
                else:
                    st.warning("⚠️ Supabase 미연결")

            st.subheader("📥 서류 다운로드")
            col_a, col_b = st.columns(2)

            # HWPX
            with col_a:
                tpl = fetch_template("신청서_양식.hwpx")
                if tpl:
                    result = process_hwpx(tpl, 데이터)
                    if result:
                        st.download_button(
                            label="📂 신청서 (HWPX)",
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
                            label="📂 계약서 (DOCX)",
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
