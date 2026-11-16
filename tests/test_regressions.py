import asyncio
import io
import json
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from claude_agent_sdk import ResultMessage
from fastapi import UploadFile
from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from backend import agents, memory, review
from backend.app import app, _store

CORE_CHECKS = [{"item": item, "status": "ok", "detail": "与原始单证一致，已核对全部适用数据"}
               for item in ("关键字段", "商品明细", "金额", "件数重量", "文件可用性")]


@pytest.fixture
def client():
    with TestClient(app) as c:
        yield c


@pytest.mark.parametrize("pid", ["..", "aaaaaaaaaaaa__escape"])
def test_playbook_deletion_cannot_escape_memory(client, tmp_path, monkeypatch, pid):
    books = tmp_path / "memory" / "playbooks"
    books.mkdir(parents=True)
    keep = tmp_path / "memory" / "MEMORY.md"
    keep.write_text("keep")
    outside = tmp_path / "outside"
    outside.mkdir()
    (outside / "keep.txt").write_text("keep")
    if pid != "..":
        (books / pid).symlink_to(outside, target_is_directory=True)
    monkeypatch.setattr(memory, "PLAYBOOKS", books)
    response = client.delete("/api/memory/playbooks/" + pid.replace(".", "%2e"))
    assert response.status_code == 404
    assert keep.read_text() == "keep"
    assert (outside / "keep.txt").read_text() == "keep"


def test_upload_preserves_files_from_previous_messages(tmp_path):
    (tmp_path / "invoice.txt").write_bytes(b"original")
    (tmp_path / "invoice_2.txt").write_bytes(b"earlier")
    uploads = [UploadFile(filename="invoice.txt", file=io.BytesIO(value))
               for value in (b"second", b"third")]
    saved = asyncio.run(_store(uploads, tmp_path))
    assert saved == ["invoice_3.txt", "invoice_4.txt"]
    assert (tmp_path / "invoice.txt").read_bytes() == b"original"
    assert (tmp_path / "invoice_2.txt").read_bytes() == b"earlier"
    assert [(tmp_path / n).read_bytes() for n in saved] == [b"second", b"third"]


@pytest.mark.parametrize("insert_row", [False, True])
def test_precheck_detects_cached_formula_errors(tmp_path, insert_row):
    for name in ("template", "output"):
        (tmp_path / name).mkdir()
    wb = Workbook()
    wb.active["A1"] = "=1/0"
    wb.save(tmp_path / "template" / "t.xlsx")
    if insert_row:
        wb.active["A2"] = "new row"
    out = tmp_path / "output" / "o.xlsx"
    wb.save(out)
    # 模拟 LibreOffice 重算后写入的公式缓存；openpyxl 本身不计算公式。
    with zipfile.ZipFile(out) as z:
        entries = {name: z.read(name) for name in z.namelist()}
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheet = ET.fromstring(entries["xl/worksheets/sheet1.xml"])
    cell = sheet.find('.//s:c[@r="A1"]', ns)
    cell.set("t", "e")
    cell.find("s:v", ns).text = "#DIV/0!"
    entries["xl/worksheets/sheet1.xml"] = ET.tostring(sheet)
    with zipfile.ZipFile(out, "w") as z:
        for name, content in entries.items():
            z.writestr(name, content)
    assert load_workbook(out, data_only=True).active["A1"].value == "#DIV/0!"
    checks = review.precheck(tmp_path, {})
    assert any(c["status"] == "error" and c["cell"] == "Sheet!A1" for c in checks)


@pytest.mark.parametrize("case, expected", [
    ("valid", "ok"), ("label", "error"), ("empty_field", "error"),
    ("missing_sheet", "warn"), ("formula_error", "error"), ("corrupt", "error"),
])
def test_xls_core_precheck(tmp_path, case, expected):
    fixtures = Path(__file__).parent / "fixtures"
    for folder in ("template", "output"):
        (tmp_path / folder).mkdir()
        shutil.copy(fixtures / "core.xls", tmp_path / folder / "core.xls")
    out = tmp_path / "output" / "core.xls"
    cell = "Missing!B2" if case == "missing_sheet" else "Main!C2" if case == "empty_field" else "Main!B2"
    result = {"fields": [{"field": "字段", "cell": cell, "value": "VALUE"}]}
    if case == "label":
        data = out.read_bytes()
        assert b"LABEL" in data
        out.write_bytes(data.replace(b"LABEL", b"BROKE"))
    elif case == "formula_error":
        shutil.copy(fixtures / "formula_error.xls", out)
    elif case == "corrupt":
        out.write_bytes(b"broken")
    # 附带的 xlsx 版本不能取代与模板同格式的主成品。
    Workbook().save(tmp_path / "output" / "a_companion.xlsx")
    checks = review.precheck(tmp_path, result)
    assert review.verdict(checks) == ("pass" if expected == "ok" else "fail" if expected == "error" else "warn")
    assert any(c["status"] == expected for c in checks)
    assert not any("跳过" in c["detail"] for c in checks)


@pytest.mark.parametrize("verdict, checks, expected", [
    ("pass", [], "warn"),
    ("pass", [None, {}, {"item": "无明细"}], "warn"),
    ("pass", [{"item": "合同", "status": "ok", "detail": "一致"}], "warn"),
    ("pass", [{"item": "合同", "status": "error"}], "fail"),
    ("fail", [], "fail"),
    ("fail", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "fail"),
    ("warn", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "warn"),
    ("unknown", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "warn"),
    ("pass", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "warn"),
    ("pass", CORE_CHECKS, "pass"),
    ("pass", CORE_CHECKS[:-1], "warn"),
    ("pass", [CORE_CHECKS[0]] * 10, "warn"),
])
def test_review_requires_complete_checks_and_honors_verdict(client, monkeypatch, verdict, checks, expected):
    job = agents.new_job("fill", "review regression")
    wd = agents.workdir(job)
    wb = Workbook()
    wb.active["A1"] = "label"
    wb.save(wd / "template" / "t.xlsx")
    wb.save(wd / "output" / "o.xlsx")

    async def report(*, prompt, options):
        (Path(options.cwd) / "review.json").write_text(json.dumps({
            "verdict": verdict, "summary": "报告结论", "checks": checks,
        }))
        yield ResultMessage(subtype="success", duration_ms=1, duration_api_ms=1,
                            is_error=False, num_turns=1, session_id="review-test", result="done")

    monkeypatch.setattr(agents, "query", report)
    rv = asyncio.run(agents._review_once(job))
    assert rv["verdict"] == expected
    job["status"] = "done"
    agents._save(job)


@pytest.mark.parametrize("interruption", ["sdk_error", "turn_limit"])
def test_interrupted_review_preserves_findings_and_restores_output(client, monkeypatch, interruption):
    job = agents.new_job("fill", "partial core review")
    wd = agents.workdir(job)
    out = wd / "output" / "keep.txt"
    out.write_text("original")
    closed = []

    async def interrupted(*, prompt, options):
        assert options.max_turns == 200
        out.write_text("tampered")
        (wd / "output" / "extra.txt").write_text("extra")
        (wd / "review.json").write_text(json.dumps({
            "verdict": "fail", "summary": "已确认金额不一致",
            "checks": [{"item": "金额", "status": "error", "detail": "成品100，原始发票200"}],
        }))
        try:
            if interruption == "sdk_error":
                raise RuntimeError("connection lost")
            yield ResultMessage(subtype="error_max_turns", duration_ms=1, duration_api_ms=1,
                                is_error=True, num_turns=200, session_id="partial", result="error_max_turns")
        finally:
            closed.append(True)

    monkeypatch.setattr(agents, "query", interrupted)
    rv = asyncio.run(agents._review_once(job))
    assert rv["verdict"] == "fail"
    assert any(c["item"] == "金额" and c["status"] == "error" for c in rv["checks"])
    assert any(c["item"] == "独立复核" and c["status"] == "warn" for c in rv["checks"])
    assert any(c["item"] == "复核报告完整性" and "关键字段" in c["detail"] for c in rv["checks"])
    assert closed and out.read_text() == "original"
    assert not (wd / "output" / "extra.txt").exists()
    job["status"] = "done"
    agents._save(job)


def test_jobs_with_same_timestamp_show_newest_first(client, monkeypatch):
    monkeypatch.setattr(agents, "JOBS", {})
    first = agents.new_job("audit", "first")
    second = agents.new_job("audit", "second")
    for job in (first, second):
        job["created"] = "2026-11-12T12:00:00"
        job["status"] = "done"
    assert [j["id"] for j in client.get("/api/jobs").json()] == [second["id"], first["id"]]
