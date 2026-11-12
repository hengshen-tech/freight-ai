"""货代 AI 工作台 —— FastAPI 服务。启动：uvicorn backend.app:app --port 8000"""
import os
import re
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from . import agents, db, memory, settings

FRONT = db.ROOT / "frontend" / "index.html"
ALLOWED_EXT = {".xlsx", ".xlsm", ".xls", ".csv", ".pdf", ".png", ".jpg", ".jpeg", ".txt", ".docx"}
MAX_UPLOAD = 25 * 1024 * 1024      # 单个文件上限 25MB


@asynccontextmanager
async def lifespan(_app):
    for k in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_BASE_URL", "CLAUDE_MODEL", "REVIEW_MODEL"):
        if os.environ.get(k) == "":        # compose 里留空的变量按未设置处理
            del os.environ[k]
    db.init()
    memory.init()
    agents.JOBS_DIR.mkdir(parents=True, exist_ok=True)
    agents.CHAT_DIR.mkdir(parents=True, exist_ok=True)
    agents.load_jobs()
    yield


app = FastAPI(title="货代 AI 工作台", lifespan=lifespan)


@app.get("/")
def index():
    return FileResponse(FRONT, headers={"Cache-Control": "no-cache"})


@app.get("/api/health")
def health():
    return {
        "api_key": settings.configured(),
        "model": settings.model(),
        "skills": sorted(p.parent.name for p in agents.SKILLS_DIR.glob("*/SKILL.md")),
        "sample_template": agents.SAMPLE_TEMPLATE.exists(),
    }


# ---------- 设置 ----------
@app.get("/api/settings")
def get_settings():
    return {**settings.public(), "rules": settings.rules()}


@app.put("/api/settings")
def put_settings(data: dict = Body(...)):
    try:
        out = settings.save(data)
    except ValueError as e:
        raise HTTPException(400, str(e))
    if "rules" in data:
        settings.save_rules(str(data["rules"] or ""))
    return {**out, "rules": settings.rules()}


@app.post("/api/settings/test")
async def test_settings():
    if not settings.configured():
        return {"ok": False, "message": "还没有填写 API Key。"}
    return await agents.test_connection()


# ---------- 工作台 / 运单 ----------
@app.get("/api/dashboard")
def dashboard():
    return db.dashboard()


@app.post("/api/shipments")
def create_shipment(data: dict = Body(...)):
    try:
        return db.create_shipment(data)
    except Exception as e:
        raise HTTPException(400, f"创建失败：{e}")


@app.patch("/api/shipments/{sid}")
def update_shipment(sid: int, data: dict = Body(...)):
    try:
        s = db.update_shipment(sid, data)
    except Exception as e:
        raise HTTPException(400, f"更新失败：{e}")
    if not s:
        raise HTTPException(404, "运单不存在")
    return s


@app.post("/api/shipments/{sid}/advance")
def advance(sid: int):
    s = db.get_shipment(sid)
    if not s:
        raise HTTPException(404, "运单不存在")
    nxt = s["stages"][min(s["stage_index"] + 1, len(s["stages"]) - 1)]
    return db.update_shipment(sid, {"status": nxt})


@app.delete("/api/shipments/{sid}")
def delete_shipment(sid: int):
    db.delete_shipment(sid)
    return {"ok": True}


@app.post("/api/todos/{tid}/toggle")
def toggle_todo(tid: int):
    db.toggle_todo(tid)
    return {"ok": True}


# ---------- 上传 ----------
def _safe_name(name: str) -> str:
    name = Path((name or "file").replace("\\", "/")).name
    name = re.sub(r"[^\w.\-()（）一-鿿 ]", "_", name).strip(" .") or "file"
    if Path(name).suffix.lower() not in ALLOWED_EXT:
        raise HTTPException(400, f"不支持的文件类型：{name}")
    return name


def _check(files) -> list[str]:
    """建任务之前先校验全部文件名，避免留下空任务。"""
    return [_safe_name(f.filename) for f in files if f and f.filename]


async def _store(files, folder: Path) -> list[str]:
    saved = []
    for f in files:
        if not (f and f.filename):
            continue
        data = await f.read()
        if len(data) > MAX_UPLOAD:
            raise HTTPException(400, f"文件过大（上限 25MB）：{f.filename}")
        name = _safe_name(f.filename)
        stem, suf, n = Path(name).stem, Path(name).suffix, 1
        while name in saved:               # 同一批里重名的文件不互相覆盖
            n += 1
            name = f"{stem}_{n}{suf}"
        (folder / name).write_bytes(data)
        saved.append(name)
    return saved


def _need_key():
    if not settings.configured():
        raise HTTPException(400, "还没有配置模型的 API Key，请先到「设置」填写。")


# ---------- AI 任务 ----------
@app.post("/api/jobs")
async def create_job(kind: str = Form(...), instructions: str = Form(""), shipment_id: int | None = Form(None),
                     use_sample_template: bool = Form(False), customer: str = Form(""),
                     files: list[UploadFile] = File(default=[]), template: UploadFile | None = File(None)):
    if kind not in ("fill", "audit", "hs"):
        raise HTTPException(400, "未知任务类型")
    _need_key()
    names = _check(files)
    has_tpl = bool(template and template.filename)
    if has_tpl:
        _safe_name(template.filename)
    ship = db.get_shipment(shipment_id) if shipment_id else None
    if kind in ("fill", "audit") and not names:
        raise HTTPException(400, "请至少上传一份原始单证")
    if kind == "fill" and not has_tpl and not use_sample_template:
        raise HTTPException(400, "请上传报关模板，或勾选使用内置示例模板")
    if kind == "hs" and not instructions.strip():
        raise HTTPException(400, "请填写商品信息")

    title = {"fill": "报关单填制", "audit": "单证审核", "hs": "HS 归类"}[kind]
    title += f" · {ship['ref']}" if ship else (f" · {instructions.strip()[:16]}" if kind == "hs" else f" · {names[0]}")
    customer = customer.strip() or (ship or {}).get("customer") or ""
    job = agents.new_job(kind, title, ship["id"] if ship else None, customer=customer)
    wd = agents.workdir(job)
    try:
        await _store(files, wd / "input")
        if kind == "fill":
            if has_tpl:
                await _store([template], wd / "template")
            else:
                agents._prepare(wd, with_template=True)
    except HTTPException as e:
        job["status"], job["error"] = "error", e.detail
        agents._save(job)
        raise

    extra = ""
    if kind == "hs":
        extra = instructions.strip()
    else:
        if ship:
            extra += "\n本票在系统里的运单信息（仅作参考，与单证冲突时以单证为准并在问题清单指出）：\n" + \
                     str({k: v for k, v in ship.items() if k not in ("stages", "stage_index", "id") and v})
        extra += agents.identify_text(customer, with_template=(kind == "fill"))
        if instructions.strip():
            extra += "\n操作员的补充要求：\n" + instructions.strip()
        if kind == "fill":      # 同一模板 / 同一客户之前确认正确的做法，直接带进任务目录
            tpl = next((p for p in sorted((wd / "template").iterdir()) if p.is_file()), None)
            hits, text = memory.attach(wd, customer, tpl)
            job["experience"] = [{"id": m["id"], "match": m["match"], "customer": m.get("customer", ""),
                                  "template_name": m.get("template_name", "")} for m in hits]
            extra += text
    tools = {"fill": agents.FILE_TOOLS, "audit": agents.FILE_TOOLS, "hs": ["Skill", "Write", "WebSearch"]}[kind]
    skills = ["customs-declaration"] if kind == "hs" else ["xlsx", "customs-declaration"]
    agents.start(job, agents.PROMPTS[kind].replace("{extra}", extra), tools=tools, skills=skills)
    return agents.public(job)


@app.get("/api/jobs")
def list_jobs():
    jobs = [j for j in agents.JOBS.values() if j["kind"] != "chat"]
    return [agents.public(j, brief=True) for j in sorted(jobs, key=lambda j: j["created"], reverse=True)]


@app.get("/api/jobs/{jid}")
def get_job(jid: str):
    if jid not in agents.JOBS:
        raise HTTPException(404, "任务不存在")
    return agents.public(agents.JOBS[jid])


@app.post("/api/jobs/{jid}/followup")
async def followup(jid: str, data: dict = Body(...)):
    job = agents.JOBS.get(jid)
    msg = (data.get("message") or "").strip()
    if not job or job["kind"] == "chat":
        raise HTTPException(404, "任务不存在")
    if not msg:
        raise HTTPException(400, "请填写要修改的内容")
    if job["status"] == "running" or job.get("_learning"):
        raise HTTPException(409, "任务还在处理中，请稍等片刻再试")
    _need_key()
    agents.follow_up(job, msg)
    return agents.public(job)


@app.post("/api/jobs/{jid}/customer")
def set_job_customer(jid: str, data: dict = Body(...)):
    """操作员更正 AI 识别的客户。"""
    job = agents.JOBS.get(jid)
    name = str(data.get("customer") or "").strip()
    if not job or job["kind"] == "chat":
        raise HTTPException(404, "任务不存在")
    if not name:
        raise HTTPException(400, "请填写客户名")
    agents.set_customer(job, name)
    return agents.public(job)


@app.get("/api/customers")
def customers():
    return agents.known_customers()


@app.post("/api/jobs/{jid}/confirm")
async def confirm(jid: str):
    """操作员确认结果无误。确认后后台会把这次的做法总结成模板经验，供以后同类模板复用。"""
    job = agents.JOBS.get(jid)
    if not job or job["kind"] != "fill":
        raise HTTPException(404, "只有报关单填制任务可以确认")
    if job["status"] == "running":
        raise HTTPException(409, "任务还在进行中，请等它结束")
    if job["status"] != "done" or not (job.get("result") or {}).get("output_files"):
        raise HTTPException(400, "任务没有成功产出文件，不能确认")
    job["confirmed"] = True
    agents._save(job)
    if job.get("session_id") and settings.configured():
        agents.learn(job)
    return agents.public(job)


# ---------- 记忆与经验 ----------
@app.get("/api/memory")
def get_memory():
    return {"general": memory.general(), "customers": memory.list_customers(),
            "playbooks": memory.list_playbooks(with_text=True)}


@app.put("/api/memory")
def put_memory(data: dict = Body(...)):
    if "general" in data:
        memory.save_general(str(data["general"] or ""))
    if data.get("customer"):
        memory.save_customer_notes(str(data["customer"]), str(data.get("notes") or ""))
    return get_memory()


@app.delete("/api/memory/playbooks/{pid}")
def delete_playbook(pid: str):
    if not memory.delete_playbook(pid):
        raise HTTPException(404, "经验不存在")
    return {"ok": True}


@app.get("/api/jobs/{jid}/files/{name}")
def download(jid: str, name: str):
    job = agents.JOBS.get(jid)
    path = agents.workdir(job) / "output" / Path(name).name if job else None
    if not path or not path.is_file():
        raise HTTPException(404, "文件不存在")
    return FileResponse(path, filename=path.name)


# ---------- 对话 ----------
@app.get("/api/chats")
def list_chats():
    return agents.list_convs()


@app.delete("/api/chats/{conv}")
def delete_chat(conv: str):
    if not agents.valid_conv(conv):
        raise HTTPException(400, "对话 ID 不合法")
    if agents.running_job(conv=conv):
        raise HTTPException(409, "这个对话还在处理中，稍后再删")
    agents.delete_conv(conv)
    return {"ok": True}


@app.get("/api/chat/{conv}")
def get_chat(conv: str):
    if not agents.valid_conv(conv):
        raise HTTPException(400, "对话 ID 不合法")
    run = agents.running_job(conv=conv)
    return {"messages": agents.load_conv(conv)["messages"], "running_job": run["id"] if run else None}


@app.post("/api/chat")
async def chat(message: str = Form(""), conv: str = Form(...), files: list[UploadFile] = File(default=[])):
    if not agents.valid_conv(conv):
        raise HTTPException(400, "对话 ID 不合法")
    _check(files)
    msg = message.strip()
    if not msg and not any(f and f.filename for f in files):
        raise HTTPException(400, "消息不能为空")
    if agents.running_job(conv=conv):
        raise HTTPException(409, "上一条还在处理中，请稍候")
    _need_key()
    job = agents.new_job("chat", msg[:30], conv=conv)
    try:
        saved = await _store(files, agents.workdir(job) / "input")
    except HTTPException:
        agents.JOBS.pop(job["id"], None)        # 上传失败不能留下一个“进行中”的任务，否则这个对话会一直被占用
        raise
    data = agents.load_conv(conv)
    data["messages"].append({"role": "user", "text": msg, "files": saved})
    agents.save_conv(conv, data)
    prompt = msg or "请先看一下我上传的文件，告诉我它们是什么单证、关键信息有哪些，以及你建议下一步做什么。"
    if saved:
        prompt += "\n\n（本次上传的文件已保存到 input/：" + "、".join(saved) + "）"
    agents.start(job, prompt, tools=agents.FILE_TOOLS, skills=["xlsx", "customs-declaration"],
                 use_tms=True, resume=data.get("session_id"), max_turns=60)
    return agents.public(job)


@app.get("/api/chat/{conv}/files/{name}")
def chat_download(conv: str, name: str):
    if not agents.valid_conv(conv):
        raise HTTPException(400, "对话 ID 不合法")
    path = agents.CHAT_DIR / conv / "output" / Path(name).name
    if not path.is_file():
        raise HTTPException(404, "文件不存在")
    return FileResponse(path, filename=path.name)
