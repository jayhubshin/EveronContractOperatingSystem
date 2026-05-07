import os
import io
import uuid
import zipfile
import base64
import subprocess
import httpx

from fastapi import FastAPI, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from docxtpl import DocxTemplate
from dotenv import load_dotenv

load_dotenv()

app = FastAPI(title="EV-CON API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── 환경변수 ────────────────────────────────────────────────
SUPABASE_URL   = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY   = os.getenv("SUPABASE_KEY", "")
GITHUB_REPO    = os.getenv("GITHUB_REPO", "")        # "username/repo"
GITHUB_TOKEN   = os.getenv("GITHUB_TOKEN", "")
GITHUB_BRANCH  = os.getenv("GITHUB_BRANCH", "main")

# Supabase 클라이언트 (선택적)
supabase = None
if SUPABASE_URL and SUPABASE_KEY:
    from supabase import create_client
    supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# 정적 파일 서빙
app.mount("/static", StaticFiles(directory="static"), name="static")

TEMPLATES_DIR = "templates"
os.makedirs(TEMPLATES_DIR, exist_ok=True)
os.makedirs("outputs", exist_ok=True)


# ── 데이터 모델 ─────────────────────────────────────────────
class ContractData(BaseModel):
    사업구분:     str = ""
    아파트명:     str = ""
    주소:         str = ""
    사업자번호:   str = ""
    관리소전화:   str = ""
    설치수량:     int = 0
    주차면수:     int = 0
    설치단가:     int = 0
    설치금액:     int = 0
    계약년수:     int = 0
    프로모션기간: int = 0
    프로모션요금: int = 0
    saveMode:     str = "local"


# ── GitHub에서 템플릿 다운로드 ──────────────────────────────
async def fetch_template_from_github(filename: str) -> bytes | None:
    """
    GitHub raw content API로 templates/ 폴더의 파일을 바이너리로 가져옵니다.
    """
    url = (
        f"https://api.github.com/repos/{GITHUB_REPO}/contents/templates/{filename}"
        f"?ref={GITHUB_BRANCH}"
    )
    headers = {
        "Authorization": f"token {GITHUB_TOKEN}",
        "Accept": "application/vnd.github.v3+json",
    }
    async with httpx.AsyncClient() as client:
        resp = await client.get(url, headers=headers)
        if resp.status_code != 200:
            print(f"GitHub 파일 조회 실패 ({filename}): {resp.status_code} {resp.text}")
            return None
        data = resp.json()
        # GitHub API는 base64로 인코딩된 content를 반환
        content_b64 = data.get("content", "").replace("\n", "")
        return base64.b64decode(content_b64)


async def get_template(filename: str) -> bytes | None:
    """
    로컬 캐시 우선, 없으면 GitHub에서 다운로드.
    """
    local_path = os.path.join(TEMPLATES_DIR, filename)
    if os.path.exists(local_path):
        with open(local_path, "rb") as f:
            return f.read()
    # GitHub에서 가져오기
    data = await fetch_template_from_github(filename)
    if data:
        # 로컬 캐시 저장
        with open(local_path, "wb") as f:
            f.write(data)
    return data


# ── GitHub 템플릿 강제 새로고침 엔드포인트 ─────────────────
@app.post("/api/refresh-templates")
async def refresh_templates():
    """로컬 캐시를 삭제하고 GitHub에서 최신 템플릿을 다시 받습니다."""
    results = {}
    for filename in ["계약서_양식.docx", "신청서_양식.hwpx"]:
        local_path = os.path.join(TEMPLATES_DIR, filename)
        if os.path.exists(local_path):
            os.remove(local_path)
        data = await fetch_template_from_github(filename)
        if data:
            with open(local_path, "wb") as f:
                f.write(data)
            results[filename] = "✅ 갱신 완료"
        else:
            results[filename] = "❌ 갱신 실패"
    return JSONResponse(results)


# ── HWPX 메일머지 (텍스트 치환) ────────────────────────────
def process_hwpx(template_bytes: bytes, data: dict) -> bytes | None:
    """
    HWPX(ZIP 구조) 내부 XML에서 {{키}} 패턴을 데이터로 치환합니다.
    """
    try:
        in_zip  = io.BytesIO(template_bytes)
        out_zip = io.BytesIO()
        with zipfile.ZipFile(in_zip, 'r') as zin, \
             zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                buffer = zin.read(item.filename)
                # Contents/section*.xml 파일만 치환
                if (item.filename.startswith("Contents/section")
                        and item.filename.endswith(".xml")):
                    content = buffer.decode("utf-8")
                    for key, value in data.items():
                        if isinstance(value, int) and value > 999:
                            display = f"{value:,}"
                        else:
                            display = str(value)
                        content = content.replace(f"{{{{{key}}}}}", display)
                    buffer = content.encode("utf-8")
                zout.writestr(item, buffer)
        return out_zip.getvalue()
    except Exception as e:
        print(f"HWPX 처리 오류: {e}")
        return None


# ── DOCX 메일머지 (docxtpl 사용) ───────────────────────────
def process_docx(template_bytes: bytes, data: dict) -> bytes | None:
    """
    docxtpl의 Jinja2 방식 {{ 키 }} 치환을 사용합니다.
    """
    try:
        tmp_path = f"outputs/tmp_{uuid.uuid4().hex[:8]}.docx"
        with open(tmp_path, "wb") as f:
            f.write(template_bytes)
        doc = DocxTemplate(tmp_path)
        doc.render(data)
        out = io.BytesIO()
        doc.save(out)
        os.remove(tmp_path)
        return out.getvalue()
    except Exception as e:
        print(f"DOCX 처리 오류: {e}")
        return None


# ── 데이터 딕셔너리 준비 (saveMode 제외) ───────────────────
def build_context(contract: ContractData) -> dict:
    ctx = contract.model_dump(exclude={"saveMode"})
    # 금액 포맷 추가 (템플릿에서 {{설치금액_fmt}} 형태로도 사용 가능)
    ctx["설치단가_fmt"]   = f"{ctx['설치단가']:,}"
    ctx["설치금액_fmt"]   = f"{ctx['설치금액']:,}"
    ctx["프로모션요금_fmt"] = f"{ctx['프로모션요금']:,}"
    return ctx


# ── 엔드포인트: HWPX 다운로드 ──────────────────────────────
@app.post("/api/generate/hwpx")
async def generate_hwpx(contract: ContractData):
    tmpl_bytes = await get_template("신청서_양식.hwpx")
    if not tmpl_bytes:
        return JSONResponse({"error": "신청서_양식.hwpx 템플릿을 찾을 수 없습니다."}, status_code=404)

    ctx = build_context(contract)
    result = process_hwpx(tmpl_bytes, ctx)
    if not result:
        return JSONResponse({"error": "HWPX 메일머지 실패"}, status_code=500)

    # DB 저장 (옵션)
    if contract.saveMode == "db" and supabase:
        try:
            supabase.table("contracts").upsert(
                contract.model_dump(exclude={"saveMode"}),
                on_conflict="아파트명"
            ).execute()
        except Exception as e:
            print(f"DB 저장 오류: {e}")

    filename = f"{contract.아파트명}_신청서.hwpx"
    return StreamingResponse(
        io.BytesIO(result),
        media_type="application/octet-stream",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{filename}",
            "Content-Length": str(len(result)),
        }
    )


# ── 엔드포인트: DOCX 다운로드 ──────────────────────────────
@app.post("/api/generate/docx")
async def generate_docx(contract: ContractData):
    tmpl_bytes = await get_template("계약서_양식.docx")
    if not tmpl_bytes:
        return JSONResponse({"error": "계약서_양식.docx 템플릿을 찾을 수 없습니다."}, status_code=404)

    ctx = build_context(contract)
    result = process_docx(tmpl_bytes, ctx)
    if not result:
        return JSONResponse({"error": "DOCX 메일머지 실패"}, status_code=500)

    filename = f"{contract.아파트명}_계약서.docx"
    return StreamingResponse(
        io.BytesIO(result),
        media_type="application/octet-stream",
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{filename}",
            "Content-Length": str(len(result)),
        }
    )


# ── 엔드포인트: 두 파일 동시 생성 ─────────────────────────
@app.post("/api/generate/all")
async def generate_all(contract: ContractData):
    ctx = build_context(contract)
    apt = contract.아파트명
    errors = []
    files  = {}

    # HWPX
    hwpx_tmpl = await get_template("신청서_양식.hwpx")
    if hwpx_tmpl:
        hwpx_bytes = process_hwpx(hwpx_tmpl, ctx)
        if hwpx_bytes:
            files["hwpx"] = base64.b64encode(hwpx_bytes).decode()
        else:
            errors.append("HWPX 메일머지 실패")
    else:
        errors.append("신청서_양식.hwpx 템플릿 없음")

    # DOCX
    docx_tmpl = await get_template("계약서_양식.docx")
    if docx_tmpl:
        docx_bytes = process_docx(docx_tmpl, ctx)
        if docx_bytes:
            files["docx"] = base64.b64encode(docx_bytes).decode()
        else:
            errors.append("DOCX 메일머지 실패")
    else:
        errors.append("계약서_양식.docx 템플릿 없음")

    # DB 저장
    if contract.saveMode == "db" and supabase:
        try:
            supabase.table("contracts").upsert(
                contract.model_dump(exclude={"saveMode"}),
                on_conflict="아파트명"
            ).execute()
        except Exception as e:
            errors.append(f"DB 저장 오류: {e}")

    return JSONResponse({
        "apt": apt,
        "files": files,   # base64 인코딩된 파일 데이터
        "errors": errors,
    })


# ── 루트 ────────────────────────────────────────────────────
@app.get("/")
async def root():
    return FileResponse("static/index.html")


# ── 템플릿 상태 확인 ────────────────────────────────────────
@app.get("/api/template-status")
async def template_status():
    status = {}
    for fn in ["계약서_양식.docx", "신청서_양식.hwpx"]:
        local = os.path.join(TEMPLATES_DIR, fn)
        status[fn] = "로컬 캐시 있음" if os.path.exists(local) else "GitHub에서 로드 필요"
    return JSONResponse(status)
