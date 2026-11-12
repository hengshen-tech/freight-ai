import asyncio
import io
import json
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


@pytest.mark.parametrize("verdict, checks, expected", [
    ("pass", [], "warn"),
    ("pass", [None, {}, {"item": "无明细"}], "warn"),
    ("pass", [{"item": "合同", "status": "ok", "detail": "一致"}], "warn"),
    ("pass", [{"item": "合同", "status": "error"}], "fail"),
    ("fail", [], "fail"),
    ("fail", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "fail"),
    ("warn", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "warn"),
    ("unknown", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "warn"),
    ("pass", [{"item": str(i), "status": "ok", "detail": "一致"} for i in range(10)], "pass"),
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


def test_jobs_with_same_timestamp_show_newest_first(client, monkeypatch):
    monkeypatch.setattr(agents, "JOBS", {})
    first = agents.new_job("audit", "first")
    second = agents.new_job("audit", "second")
    for job in (first, second):
        job["created"] = "2026-11-12T12:00:00"
        job["status"] = "done"
    assert [j["id"] for j in client.get("/api/jobs").json()] == [second["id"], first["id"]]
