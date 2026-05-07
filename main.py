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
    st.divider()

    c1, c2, c3 = st.columns(3)

    with c1:
        사업구분   = st.selectbox("사업구분", [
            "한국환경공단 이사장",
            "주식회사 에버온인프라",
            "기타"
        ])
        아파트명   = st.text_input("아파트명 *")
        주소       = st.text_input("주소")
        사업자번호 = st.text_input("사업자번호")
        관리소전화 = st.text_input("관리소전화")

    with c2:
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
        최종설치금액 = st.number_input("최종 설치금액 (원)", min_value=0, value=calc)
        st.caption(f"💡 {설치수량}기 × {설치단가:,}원 = {calc:,}원")

    with c3:
        계약년수     = st.number_input("계약년수 (년)", min_value=0, value=7)
        프로모션기간 = st.number_input("프로모션기간 (월)", min_value=0, value=0)
        프로모션요금 = st.number_input("프로모션요금 (원)", min_value=0, value=0)

    st.divider()
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
# ── 미리보기 ────────────────────────────────────────────
if 미리보기:
    if not 아파트명:
        st.warning("⚠️ 아파트명을 입력해주세요.")
    else:
        tab1, tab2 = st.tabs(["📄 신청서 미리보기", "📝 계약서 미리보기"])

        with tab1:
            st.markdown(f"""
            <div style="
                background:white; color:#111; padding:48px 56px;
                border-radius:12px; font-family:'맑은 고딕','Malgun Gothic',sans-serif;
                max-width:800px; margin:0 auto; box-shadow:0 4px 24px rgba(0,0,0,0.15);
                line-height:1.8;
            ">
                <div style="text-align:center; margin-bottom:36px;">
                    <h2 style="font-size:22px; font-weight:900; letter-spacing:4px; margin:0;">
                        전기자동차 충전기 설치 신청서
                    </h2>
                    <div style="margin-top:8px; font-size:13px; color:#555;">
                        {사업구분}
                    </div>
                </div>

                <table style="width:100%; border-collapse:collapse; font-size:14px; margin-bottom:28px;">
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700; width:30%;">아파트명</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{아파트명}</td>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700; width:30%;">사업자번호</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{사업자번호}</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">주소</td>
                        <td colspan="3" style="padding:10px 14px; border:1px solid #ccc;">{주소}</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">관리소 전화</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{관리소전화}</td>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">주차면수</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{주차면수} 면</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">설치 수량</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{설치수량} 기</td>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">설치 단가</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{설치단가:,} 원</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">설치 금액</td>
                        <td colspan="3" style="padding:10px 14px; border:1px solid #ccc; font-weight:700; color:#c00; font-size:16px;">
                            ₩ {최종설치금액:,} 원
                        </td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">계약 기간</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{계약년수} 년</td>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">프로모션 기간</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{프로모션기간} 월</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">프로모션 요금</td>
                        <td colspan="3" style="padding:10px 14px; border:1px solid #ccc;">{프로모션요금:,} 원</td>
                    </tr>
                </table>

                <div style="margin-top:48px; text-align:center; font-size:13px; color:#777;">
                    위와 같이 전기자동차 충전기 설치를 신청합니다.
                </div>
                <div style="margin-top:48px; display:flex; justify-content:space-between; font-size:14px;">
                    <div style="text-align:center; width:45%;">
                        신청인 : {아파트명} <br/><br/>
                        <div style="border-top:1px solid #333; padding-top:8px; margin-top:32px;">서명</div>
                    </div>
                    <div style="text-align:center; width:45%;">
                        수신 : {사업구분} <br/><br/>
                        <div style="border-top:1px solid #333; padding-top:8px; margin-top:32px;">직인</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        with tab2:
            st.markdown(f"""
            <div style="
                background:white; color:#111; padding:48px 56px;
                border-radius:12px; font-family:'맑은 고딕','Malgun Gothic',sans-serif;
                max-width:800px; margin:0 auto; box-shadow:0 4px 24px rgba(0,0,0,0.15);
                line-height:1.8;
            ">
                <div style="text-align:center; margin-bottom:36px;">
                    <h2 style="font-size:22px; font-weight:900; letter-spacing:4px; margin:0;">
                        전기자동차 충전기 설치 계약서
                    </h2>
                </div>

                <p style="font-size:14px; margin-bottom:24px;">
                    <b>"{아파트명}"</b> (이하 "갑")과 <b>{사업구분}</b> (이하 "을")은
                    아래와 같이 전기자동차 충전기 설치에 관한 계약을 체결한다.
                </p>

                <h3 style="font-size:15px; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:16px;">
                    제1조 (계약 목적)
                </h3>
                <p style="font-size:14px; margin-bottom:24px;">
                    본 계약은 전기자동차 충전기 설치 및 운영에 관한 제반 사항을 규정함을 목적으로 한다.
                </p>

                <h3 style="font-size:15px; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:16px;">
                    제2조 (계약 내용)
                </h3>
                <table style="width:100%; border-collapse:collapse; font-size:14px; margin-bottom:24px;">
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700; width:35%;">설치 장소</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{주소}</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">설치 수량</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{설치수량} 기</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">설치 단가</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{설치단가:,} 원</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">총 설치 금액</td>
                        <td style="padding:10px 14px; border:1px solid #ccc; font-weight:700; color:#c00;">
                            ₩ {최종설치금액:,} 원
                        </td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">계약 기간</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{계약년수} 년</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">프로모션 기간</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{프로모션기간} 월</td>
                    </tr>
                    <tr>
                        <td style="background:#f5f5f5; padding:10px 14px; border:1px solid #ccc; font-weight:700;">프로모션 요금</td>
                        <td style="padding:10px 14px; border:1px solid #ccc;">{프로모션요금:,} 원</td>
                    </tr>
                </table>

                <h3 style="font-size:15px; border-bottom:2px solid #333; padding-bottom:6px; margin-bottom:16px;">
                    제3조 (계약 기간)
                </h3>
                <p style="font-size:14px; margin-bottom:24px;">
                    계약 기간은 설치 완료일로부터 <b>{계약년수}년</b>으로 하며,
                    프로모션 기간 <b>{프로모션기간}개월</b> 동안은
                    월 <b>{프로모션요금:,}원</b>을 적용한다.
                </p>

                <div style="margin-top:60px;">
                    <table style="width:100%; font-size:14px; border:none;">
                        <tr>
                            <td style="width:50%; text-align:center; padding:12px; border:none;">
                                <div><b>갑 (신청인)</b></div>
                                <div style="margin-top:8px;">{아파트명}</div>
                                <div style="margin-top:4px; font-size:12px; color:#555;">{주소}</div>
                                <div style="margin-top:32px; border-top:1px solid #333; padding-top:8px;">서명 / 인</div>
                            </td>
                            <td style="width:50%; text-align:center; padding:12px; border:none;">
                                <div><b>을 (사업자)</b></div>
                                <div style="margin-top:8px;">{사업구분}</div>
                                <div style="margin-top:32px; border-top:1px solid #333; padding-top:8px;">서명 / 인</div>
                            </td>
                        </tr>
                    </table>
                </div>
            </div>
            """, unsafe_allow_html=True)


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
