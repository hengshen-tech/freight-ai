import asyncio
import json
import logging
import shutil
import time
from pathlib import Path

import pytest
from claude_agent_sdk import AssistantMessage, ResultMessage, TextBlock, ToolResultBlock, ToolUseBlock, UserMessage
from fastapi.testclient import TestClient

from backend import agents, memory
from backend.app import app

SAMPLES = Path(__file__).resolve().parent.parent / "samples"
DOCS = ["样例_商业发票.xlsx", "样例_装箱单.xlsx", "样例_销售合同.xlsx"]
PLAYBOOK = "# 报关单模板经验\n\n| 栏位 | 单元格 | 来源 |\n|---|---|---|\n| 合同协议号 | 报关单!H8 | 合同 B3 |\n" + "明细行从第 20 行开始，预留 8 行，H 列是公式不要写。" * 3


def _result(text="完成", session="sess-1"):
    return ResultMessage(subtype="success", duration_ms=1200, duration_api_ms=900, is_error=False, num_turns=4,
                         session_id=session, total_cost_usd=0.0123, usage={"input_tokens": 100, "output_tokens": 20}, result=text)


async def _call(job, name, args):
    t = next(t for t in agents._mem_tools(job) if t.name == name)
    return await t.handler(args)


STATE = {"reviews": 0, "fixes": 0, "fill_wrong": False, "tamper": False, "ai_customer": "宁波恒益电器有限公司"}


@pytest.fixture()
def client(monkeypatch):
    calls = []
    STATE.update(reviews=0, fixes=0, fill_wrong=False, tamper=False, ai_customer="宁波恒益电器有限公司")

    async def fake_query(*, prompt, options):
        """扮演 Agent：按提示词类型做一遍真实 Agent 会做的文件操作和工具调用。"""
        wd = Path(options.cwd)
        calls.append({"prompt": prompt, "options": options})
        job = next((j for j in agents.JOBS.values() if agents.workdir(j) == wd and (j["status"] == "running" or j.get("_learning"))), None)
        yield AssistantMessage(content=[TextBlock(text="开始处理"), ToolUseBlock(id="t1", name="Skill", input={"skill": "xlsx"})], model="fake")
        yield UserMessage(content=[ToolResultBlock(tool_use_id="t1", content="skill loaded", is_error=False)])
        if prompt.startswith("你是报关单复核员"):
            STATE["reviews"] += 1
            assert options.resume is None                          # 复核员是全新会话，看不到填表过程
            assert "mem" not in (options.mcp_servers or {})        # 复核员不能写记忆
            from openpyxl import load_workbook
            out = next((wd / "output").glob("*.xlsx"))
            val = load_workbook(out).active["H8"].value
            good = val == "HY-SC-2026-0917"
            if STATE["tamper"]:
                out.write_bytes(b"broken"); (wd / "output" / "extra.txt").write_text("x")
            (wd / "review.json").write_text(json.dumps({
                "verdict": "pass" if good else "fail", "summary": "逐栏核对无误" if good else "合同协议号与合同不符",
                "checks": [{"item": "合同协议号", "cell": "报关单!H8", "status": "ok" if good else "error",
                            "detail": f"成品 {val}，合同 B3 为 HY-SC-2026-0917"},
                           {"item": "成交方式", "status": "ok", "detail": "FOB，与发票一致"}]}, ensure_ascii=False), "utf-8")
            yield AssistantMessage(content=[ToolUseBlock(id="r1", name="Bash", input={"command": "python check.py"})], model="fake")
            yield UserMessage(content=[ToolResultBlock(tool_use_id="r1", content="checked", is_error=False)])
        elif prompt.startswith("独立复核发现下面这些问题"):
            STATE["fixes"] += 1
            from openpyxl import load_workbook
            out = next((wd / "output").glob("*.xlsx"))
            wb = load_workbook(out); wb.active["H8"] = "HY-SC-2026-0917"; wb.save(out)
        elif "沉淀成以后可以复用的经验" in prompt:
            tpl = next((wd / "template").iterdir())
            out = next((wd / "output").iterdir())
            r = await _call(job, "save_experience", {"template_file": f"template/{tpl.name}", "output_file": f"output/{out.name}",
                                                     "playbook": PLAYBOOK, "script_file": "work/fill.py"})
            yield AssistantMessage(content=[ToolUseBlock(id="t2", name="mcp__mem__save_experience", input={"template_file": tpl.name})], model="fake")
            yield UserMessage(content=[ToolResultBlock(tool_use_id="t2", content=r["content"], is_error=r.get("is_error", False))])
        elif "记住" in prompt:
            r = await _call(job, "remember", {"note": "监管方式默认填一般贸易(0110)"})
            await _call(job, "remember", {"note": "境内货源地固定填宁波", "customer": "宁波恒益电器"})
            yield AssistantMessage(content=[ToolUseBlock(id="t2", name="mcp__mem__remember", input={"note": "监管方式默认填一般贸易(0110)"})], model="fake")
            yield UserMessage(content=[ToolResultBlock(tool_use_id="t2", content=r["content"], is_error=False)])
        elif (wd / "template").exists() and any((wd / "template").iterdir()) and "result.json" in prompt:
            tpl = next((wd / "template").iterdir())
            (wd / "work" / "fill.py").write_text("print('fill')", "utf-8")
            from openpyxl import load_workbook
            wb = load_workbook(tpl); wb.active["H8"] = "HY-SC-2026-0917"
            if STATE["fill_wrong"]:
                wb.active["H8"] = "WRONG-NO"; STATE["fill_wrong"] = False
            wb.save(wd / "output" / "报关单_已填.xlsx")
            (wd / "result.json").write_text(json.dumps({
                "summary": "已填 2 个栏位", "output_files": ["output/报关单_已填.xlsx"],
                "customer": {"name": STATE["ai_customer"], "basis": "销售合同卖方"},
                "fields": [{"field": "合同协议号", "cell": "报关单!H8", "value": "HY-SC-2026-0917", "source": "合同 B3", "confidence": "high"}],
                "issues": [{"level": "warning", "message": "监管方式为默认值"}]}, ensure_ascii=False), "utf-8")
            yield AssistantMessage(content=[ToolUseBlock(id="t2", name="Bash", input={"command": "python work/fill.py"})], model="fake")
            yield UserMessage(content=[ToolResultBlock(tool_use_id="t2", content=[{"type": "text", "text": "fill\n"}], is_error=False)])
            yield AssistantMessage(content=[ToolUseBlock(id="t3", name="Bash", input={"command": "bad"})], model="fake")
            yield UserMessage(content=[ToolResultBlock(tool_use_id="t3", content="command not found", is_error=True)])
        yield _result()

    monkeypatch.setattr(agents, "query", fake_query)
    with TestClient(app) as c:
        c.calls = calls
        yield c


def wait_learned(client, jid, timeout=10):
    """确认后的总结在后台进行，等它结束。"""
    end = time.time() + timeout
    while time.time() < end:
        if not agents.JOBS[jid].get("_learning"):
            return client.get(f"/api/jobs/{jid}").json()
        time.sleep(0.05)
    raise AssertionError("后台总结超时")


def wait(client, jid, timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        j = client.get(f"/api/jobs/{jid}").json()
        if j["status"] != "running":
            return j
        time.sleep(0.05)
    raise AssertionError("任务超时")


def post_fill(client, customer="宁波恒益电器", template=None):
    files = [("files", (n, (SAMPLES / n).read_bytes())) for n in DOCS]
    data = {"kind": "fill", "customer": customer}
    if template:
        files.append(("template", template))
    else:
        data["use_sample_template"] = "true"
    r = client.post("/api/jobs", data=data, files=files)
    assert r.status_code == 200, r.text
    return r.json()["id"]


def test_fingerprint_ignores_filename_but_not_labels(tmp_path):
    from openpyxl import load_workbook
    a, b, c = tmp_path / "a.xlsx", tmp_path / "另一个名字.xlsx", tmp_path / "c.xlsx"
    shutil.copy(SAMPLES / "报关单模板.xlsx", a)
    shutil.copy(a, b)
    wb = load_workbook(a); wb.active["A3"] = "境内发货人（改）"; wb.save(c)
    assert memory.fingerprint(a) == memory.fingerprint(b) != memory.fingerprint(c)
    bad = tmp_path / "x.xls"; bad.write_bytes(b"not excel")
    assert memory.fingerprint(bad).startswith("f")


def test_settings_never_return_key(client):
    r = client.put("/api/settings", json={"model": "m1", "api_key": "sk-secret-123456", "base_url": "https://gw.example.com/", "rules": "规则A"}).json()
    assert r["key_set"] and r["key_hint"] == "…3456" and "api_key" not in r and r["base_url"] == "https://gw.example.com"
    assert "sk-secret" not in client.get("/api/settings").text
    from backend import settings
    env = settings.agent_env()
    assert env["ANTHROPIC_AUTH_TOKEN"] == "sk-secret-123456" and env["ANTHROPIC_API_KEY"] == "" and env["ANTHROPIC_DEFAULT_HAIKU_MODEL"] == "m1"
    assert client.put("/api/settings", json={"base_url": "ftp://x"}).status_code == 400
    client.put("/api/settings", json={"clear_key": True, "base_url": "", "model": ""})


def test_upload_validation_creates_no_job(client):
    before = len(client.get("/api/jobs").json())
    r = client.post("/api/jobs", data={"kind": "audit"}, files=[("files", ("evil.py", b"x"))])
    assert r.status_code == 400 and len(client.get("/api/jobs").json()) == before
    assert client.post("/api/jobs", data={"kind": "fill"}, files=[("files", ("a.xlsx", b"x"))]).status_code == 400


def test_fill_followup_confirm_and_reuse(client, caplog):
    records = []
    h = logging.Handler(); h.emit = lambda rec: records.append(rec.getMessage())
    agents.log.addHandler(h)
    try:
        jid = post_fill(client)
        j = wait(client, jid)
    finally:
        agents.log.removeHandler(h)
    assert j["status"] == "done" and j["result"]["output_files"] == ["报关单_已填.xlsx"] and j["experience"] == []
    rv = j["review"]
    assert rv["verdict"] == "pass" and rv["rounds"] == 0 and STATE["fixes"] == 0
    assert any(c["by"] == "复核员" and c["item"] == "合同协议号" for c in rv["checks"])
    assert any(c["by"] == "程序校验" and c["item"] == "模板原有内容" and c["status"] == "ok" for c in rv["checks"])
    assert any(e["kind"] == "review" for e in j["events"])
    assert j["result"]["fields"][0]["cell"] == "报关单!H8" and j["customer"] == "宁波恒益电器"
    assert client.get(f"/api/jobs/{jid}/files/报关单_已填.xlsx").status_code == 200
    assert client.get(f"/api/jobs/{jid}/files/..%2Fjob.json").status_code == 404
    # 日志：开始、提示词、工具输入、工具输出、报错、结束
    text = "\n".join(records)
    for tag in ("START", "PROMPT", "TOOL", "RESULT", "ERROR", "DONE", "REVIEW", "R.TOOL"):
        assert f" {tag}" in text, tag
    assert '"command": "python work/fill.py"' in text and "command not found" in text and "turns=4" in text
    assert any(e["kind"] == "result" for e in j["events"])          # 工具报错也显示在页面的执行过程里

    # 没确认之前，Agent 不能自己保存经验
    r = asyncio.run(_call(agents.JOBS[jid], "save_experience", {"template_file": "template/报关单模板.xlsx", "playbook": PLAYBOOK}))
    assert r.get("is_error")

    # 追加修改：续接同一会话
    assert client.post(f"/api/jobs/{jid}/followup", json={"message": "运费按 1200 美元填"}).status_code == 200
    j = wait(client, jid)
    fu = next(c for c in client.calls if "运费按 1200 美元填" in c["prompt"])
    assert fu["options"].resume == "sess-1"
    assert j["review"]["verdict"] == "pass" and STATE["reviews"] == 2          # 首次和追加修改后各复核一次

    assert j["confirmed"] is False          # 改过之后要重新确认

    # 确认无误 → 立即返回；经验在后台总结，不改任务状态，也不往页面的执行过程里写东西
    n_events = len(j["events"])
    r = client.post(f"/api/jobs/{jid}/confirm").json()
    assert r["status"] == "done" and r["confirmed"]
    j = wait_learned(client, jid)
    assert j["status"] == "done" and j["confirmed"] and j["playbook"] and len(j["events"]) == n_events
    pbs = client.get("/api/memory").json()["playbooks"]
    assert len(pbs) == 1 and pbs[0]["customer"] == "宁波恒益电器" and pbs[0]["has_script"] and "H8" in pbs[0]["playbook"]
    d = memory.PLAYBOOKS / pbs[0]["id"]
    assert (d / "template.xlsx").exists() and (d / "example_output.xlsx").exists()

    # 同一客户、同一模板再来一票：经验自动带进任务目录和提示词
    jid2 = post_fill(client)
    j2 = wait(client, jid2)
    assert j2["experience"][0]["match"] == "exact"
    wd2 = agents.workdir(agents.JOBS[jid2])
    assert (wd2 / "experience" / "1" / "playbook.md").exists() and (wd2 / "experience" / "1" / "fill.py").exists()
    first = next(c for c in client.calls if "experience/1/" in c["prompt"])
    assert "模板指纹与本次模板完全一致" in first["prompt"]
    assert client.get("/api/memory").json()["playbooks"][0]["uses"] == 1

    # 同一客户换了模板：作为参考
    from openpyxl import load_workbook
    import io
    wb = load_workbook(SAMPLES / "报关单模板.xlsx"); wb.active["A3"] = "发货人"; buf = io.BytesIO(); wb.save(buf)
    j3 = wait(client, post_fill(client, template=("新模板.xlsx", buf.getvalue())))
    assert j3["experience"][0]["match"] == "customer"
    # 别的客户、别的模板：不带经验
    j4 = wait(client, post_fill(client, customer="另一家", template=("新模板.xlsx", buf.getvalue())))
    assert j4["experience"] == []

    assert client.delete(f"/api/memory/playbooks/{pbs[0]['id']}").status_code == 200
    assert client.get("/api/memory").json()["playbooks"] == []


def test_chat_remembers_and_memory_reaches_next_prompt(client):
    conv = "ctest-000001"
    r = client.post("/api/chat", data={"conv": conv, "message": "记住：监管方式默认一般贸易"}, files=[("files", (DOCS[0], (SAMPLES / DOCS[0]).read_bytes()))])
    assert r.status_code == 200
    wait(client, r.json()["id"])
    d = client.get(f"/api/chat/{conv}").json()
    assert [m["role"] for m in d["messages"]] == ["user", "assistant"] and d["messages"][0]["files"] == [DOCS[0]]
    assert any("记住" in s for s in d["messages"][1]["steps"])
    m = client.get("/api/memory").json()
    assert "监管方式默认填一般贸易(0110)" in m["general"] and m["customers"][0]["name"] == "宁波恒益电器"
    # 重复的不再记
    assert asyncio.run(_call({"id": "x", "kind": "chat", "conv": conv}, "remember", {"note": "监管方式默认填一般贸易(0110)"}))["content"][0]["text"] == "已经记过了"
    # 下一个任务的系统提示词里带着通用记忆和该客户的备注
    wait(client, post_fill(client))
    sp = next(c for c in reversed(client.calls) if "把原始单证的数据填进报关模板" in c["prompt"])["options"].system_prompt
    assert "监管方式默认填一般贸易(0110)" in sp and "境内货源地固定填宁波" in sp
    # 第二条消息续接会话
    r = client.post("/api/chat", data={"conv": conv, "message": "你好"})
    wait(client, r.json()["id"])
    assert client.calls[-1]["options"].resume == "sess-1"
    # 页面上可以改、可以清空
    client.put("/api/memory", json={"general": "", "customer": "宁波恒益电器", "notes": ""})
    m = client.get("/api/memory").json()
    assert m["general"] == "" and m["customers"] == []


def test_mem_tool_paths_cannot_escape_workdir(client):
    jid = post_fill(client)
    wait(client, jid)
    job = agents.JOBS[jid]; job["_can_save"] = True
    r = asyncio.run(_call(job, "save_experience", {"template_file": "../../settings.json", "playbook": PLAYBOOK}))
    assert r.get("is_error") and "工作目录" in r["content"][0]["text"]


def test_failed_learning_is_silent_and_can_be_retried(client, monkeypatch):
    jid = post_fill(client, customer="重试客户")
    wait(client, jid)
    real = agents.query

    async def no_save(*, prompt, options):          # Agent 这次没有调用 save_experience
        yield _result("我总结好了")
    monkeypatch.setattr(agents, "query", no_save)
    assert client.post(f"/api/jobs/{jid}/confirm").status_code == 200
    j = wait_learned(client, jid)
    assert j["status"] == "done" and j["confirmed"] and not j["playbook"] and j["error"] == "" and j["result"]["output_files"]
    monkeypatch.setattr(agents, "query", real)
    assert client.post(f"/api/jobs/{jid}/confirm").status_code == 200      # 再确认一次会重新总结
    assert wait_learned(client, jid)["playbook"]


def test_secrets_are_redacted_from_logs_and_events(client, monkeypatch):
    records = []
    h = logging.Handler(); h.emit = lambda rec: records.append(rec.getMessage())
    agents.log.addHandler(h)
    client.put("/api/settings", json={"api_key": "sk-super-secret-000111"})

    async def leaky(*, prompt, options):
        yield AssistantMessage(content=[TextBlock(text="key 是 sk-super-secret-000111"), ToolUseBlock(id="t", name="Bash", input={"command": "env"})], model="f")
        yield UserMessage(content=[ToolResultBlock(tool_use_id="t", content="ANTHROPIC_AUTH_TOKEN=sk-super-secret-000111", is_error=True)])
        yield _result("sk-super-secret-000111")
    monkeypatch.setattr(agents, "query", leaky)
    try:
        r = client.post("/api/chat", data={"conv": "csecret-0001", "message": "hi"})
        j = wait(client, r.json()["id"])
    finally:
        agents.log.removeHandler(h)
        client.put("/api/settings", json={"clear_key": True})
    blob = "\n".join(records) + json.dumps(j, ensure_ascii=False) + client.get("/api/chat/csecret-0001").text
    assert "sk-super-secret" not in blob and "***" in blob


def test_conversation_list_and_delete(client):
    r = client.post("/api/chat", data={"conv": "clist-000001", "message": "第一句话"})
    wait(client, r.json()["id"])
    lst = client.get("/api/chats").json()
    assert lst[0]["id"] == "clist-000001" and lst[0]["title"] == "第一句话" and lst[0]["count"] == 2
    assert client.delete("/api/chats/..%2F..").status_code in (400, 404)
    assert client.delete("/api/chats/clist-000001").status_code == 200
    assert all(c["id"] != "clist-000001" for c in client.get("/api/chats").json())
    assert client.get("/api/chat/clist-000001").json()["messages"] == []


def test_review_catches_error_and_fill_agent_fixes_it(client):
    STATE["fill_wrong"] = True                       # 填表 Agent 把合同号填错
    j = wait(client, post_fill(client, customer="复核客户"))
    rv = j["review"]
    assert j["status"] == "done" and rv["verdict"] == "pass" and rv["rounds"] == 1
    assert STATE["reviews"] == 2 and STATE["fixes"] == 1
    fix = next(c for c in client.calls if c["prompt"].startswith("独立复核发现下面这些问题"))
    assert "合同协议号" in fix["prompt"] and fix["options"].resume == "sess-1"      # 交回同一个填表会话改
    from openpyxl import load_workbook
    out = agents.workdir(agents.JOBS[j["id"]]) / "output" / "报关单_已填.xlsx"
    assert load_workbook(out).active["H8"].value == "HY-SC-2026-0917"


def test_reviewer_cannot_change_the_output(client):
    STATE["tamper"] = True                           # 复核员违规改了成品、还多写了个文件
    j = wait(client, post_fill(client, customer="只读客户"))
    out_dir = agents.workdir(agents.JOBS[j["id"]]) / "output"
    from openpyxl import load_workbook
    assert load_workbook(out_dir / "报关单_已填.xlsx").active["H8"].value == "HY-SC-2026-0917"   # 已恢复
    assert [p.name for p in out_dir.iterdir()] == ["报关单_已填.xlsx"]
    assert j["result"]["output_files"] == ["报关单_已填.xlsx"]


def test_precheck_flags_template_damage_and_false_claims(tmp_path):
    from openpyxl import load_workbook
    from backend import review
    for d in ("template", "output"):
        (tmp_path / d).mkdir()
    shutil.copy(SAMPLES / "报关单模板.xlsx", tmp_path / "template" / "t.xlsx")
    wb = load_workbook(SAMPLES / "报关单模板.xlsx"); ws = wb.active
    ws["A3"] = "被改掉的标签"; ws["C3"] = "宁波恒益电器有限公司"; ws["H20"] = 123      # 改了标签、覆盖了公式
    wb.save(tmp_path / "output" / "o.xlsx")
    checks = review.precheck(tmp_path, {"fields": [
        {"field": "境内发货人", "cell": "报关单!C3", "value": "宁波恒益电器有限公司"},
        {"field": "合同协议号", "cell": "报关单!H8", "value": "HY-SC-2026-0917"},      # 声称填了，其实是空的
        {"field": "商品明细", "cell": "A20:J22", "value": "3 行"}]})                    # 区域不校验
    bad = {(c["item"], c["cell"]) for c in checks if c["status"] == "error"}
    assert ("模板原有内容被改动", "报关单!A3") in bad and ("模板原有内容被改动", "报关单!H20") in bad and ("合同协议号", "报关单!H8") in bad
    assert review.verdict(checks) == "fail"
    assert review._norm("22,950.00") == review._norm(22950) and review._norm(" A  B ") == "A B"


def test_frontend_script_parses(tmp_path):
    """前端是单文件、没有构建步骤，这里至少保证脚本语法正确。"""
    import subprocess
    node = shutil.which("node")
    if not node:
        pytest.skip("没有安装 node")
    html = (SAMPLES.parent / "frontend" / "index.html").read_text("utf-8")
    js = tmp_path / "app.js"
    js.write_text(html[html.index("<script>") + 8: html.rindex("</script>")], "utf-8")
    r = subprocess.run([node, "--check", str(js)], capture_output=True, text=True)
    assert r.returncode == 0, r.stderr[:500]


def test_customer_is_identified_from_documents_and_reused(client):
    # 操作员没填客户：采用 Agent 识别的结果，提示词里要求它识别并对齐已有客户名
    jid = post_fill(client, customer="")
    j = wait(client, jid)
    assert j["customer"] == "宁波恒益电器有限公司" and j["customer_source"] == "ai" and j["customer_basis"] == "销售合同卖方"
    p1 = next(c["prompt"] for c in client.calls if "把原始单证的数据填进报关模板" in c["prompt"])
    assert "客户识别：没有指定客户" in p1 and "find_experience" in p1 and "演示客户·宁波恒益电器" in p1    # 已有客户名（来自运单）
    assert "宁波恒益电器有限公司" in client.get("/api/customers").json()
    assert client.get("/api/jobs").json()[0]["customer"] == "宁波恒益电器有限公司"
    # 确认后沉淀的经验记在识别出的客户名下
    client.post(f"/api/jobs/{jid}/confirm")
    wait_learned(client, jid)
    assert client.get("/api/memory").json()["playbooks"][0]["customer"] == "宁波恒益电器有限公司"

    # 下一票：已识别过的客户名出现在“已有客户”里，供 Agent 对齐
    j2 = wait(client, post_fill(client, customer=""))
    p2 = [c["prompt"] for c in client.calls if "把原始单证的数据填进报关模板" in c["prompt"]][-1]
    assert "宁波恒益电器有限公司" in p2 and j2["experience"][0]["match"] == "exact"

    # 操作员指定的客户优先，不被 Agent 的识别结果覆盖
    j3 = wait(client, post_fill(client, customer="人工指定的客户"))
    assert j3["customer"] == "人工指定的客户" and j3["customer_source"] == "operator"
    p3 = [c["prompt"] for c in client.calls if "把原始单证的数据填进报关模板" in c["prompt"]][-1]
    assert "客户：人工指定的客户（已指定）" in p3 and "客户识别：" not in p3

    # 识别错了可以更正；之后的追加修改不会再被 Agent 改回去
    r = client.post(f"/api/jobs/{jid}/customer", json={"customer": "更正后的客户"}).json()
    assert r["customer"] == "更正后的客户" and r["customer_source"] == "operator"
    client.post(f"/api/jobs/{jid}/followup", json={"message": "再检查一遍"})
    assert wait(client, jid)["customer"] == "更正后的客户"
    assert client.post(f"/api/jobs/{jid}/customer", json={"customer": " "}).status_code == 400

    # 识别不出来：保持未识别
    STATE["ai_customer"] = None
    j4 = wait(client, post_fill(client, customer=""))
    assert j4["customer"] == "" and j4["customer_source"] == ""


def test_failed_chat_upload_does_not_block_the_conversation(client, monkeypatch):
    from backend import app as appmod
    monkeypatch.setattr(appmod, "MAX_UPLOAD", 10)
    conv = "cupload-0001"
    r = client.post("/api/chat", data={"conv": conv, "message": "看看"}, files=[("files", ("big.xlsx", b"x" * 50))])
    assert r.status_code == 400 and "过大" in r.json()["detail"]
    r = client.post("/api/chat", data={"conv": conv, "message": "你好"})       # 之前这里会一直返回 409
    assert r.status_code == 200
    wait(client, r.json()["id"])


def test_review_crash_still_delivers_the_fill_result(client, monkeypatch):
    from backend import review
    monkeypatch.setattr(review, "snapshot", lambda wd: (_ for _ in ()).throw(OSError("磁盘满了")))
    j = wait(client, post_fill(client, customer="出错客户"))
    assert j["status"] == "done" and j["result"]["output_files"] and j["review"]["verdict"] == "warn"
    assert "复核没有完成" in j["review"]["summary"]
