"""长期记忆与制单经验库 —— 都是 data/memory/ 下的普通文件，人可以直接看、改、删。

思路沿用 Anthropic 的做法（记忆工具的 /memories 目录、Claude Code 的 MEMORY.md）：
记忆就是一个目录里的 Markdown 文件，每次任务开始时带上索引，Agent 需要时再读细节、学到东西再写回去。

data/memory/
  MEMORY.md                      通用记忆：操作员说过的、以后还用得上的话（每次任务都带上）
  customers/<客户>.md            某个客户的备注（任务关联到该客户时带上）
  playbooks/<指纹>__<客户>/      一份「模板经验」：某客户某个模板，人工确认正确后沉淀下来
      meta.json                  客户、模板指纹、模板文件名、使用次数、来源任务
      playbook.md                栏位→单元格映射、取数规律、踩过的坑
      fill.py                    （可选）当时写的可复用填表脚本
      template.*  example_output.*   当时的模板和确认正确的成品，供对照
"""
import hashlib
import json
import re
import shutil
from datetime import datetime
from pathlib import Path

from . import db

MEM = db.DATA / "memory"
GENERAL = MEM / "MEMORY.md"
CUSTOMERS = MEM / "customers"
PLAYBOOKS = MEM / "playbooks"
MAX_INJECT = 6000          # 注入系统提示词的记忆上限（字符）


def init():
    for d in (MEM, CUSTOMERS, PLAYBOOKS):
        d.mkdir(parents=True, exist_ok=True)


def slug(name: str) -> str:
    s = re.sub(r"[^\w一-鿿-]+", "_", (name or "").strip()).strip("_")
    return s[:60] or "未指定客户"


def _now():
    return datetime.now().isoformat(timespec="seconds")


# ---------------- 记忆（操作员说过的话） ----------------
def general() -> str:
    return GENERAL.read_text("utf-8") if GENERAL.exists() else ""


def save_general(text: str):
    init()
    GENERAL.write_text(text or "", "utf-8")


def customer_notes(name: str) -> str:
    f = CUSTOMERS / f"{slug(name)}.md"
    return f.read_text("utf-8") if name and f.exists() else ""


def save_customer_notes(name: str, text: str):
    init()
    f = CUSTOMERS / f"{slug(name)}.md"
    if (text or "").strip():
        f.write_text(text, "utf-8")
    elif f.exists():
        f.unlink()


def list_customers() -> list[dict]:
    init()
    return [{"name": f.stem, "notes": f.read_text("utf-8")} for f in sorted(CUSTOMERS.glob("*.md"))]


def remember(note: str, customer: str = "") -> str:
    """追加一条记忆。带 customer 记到该客户名下，否则记到通用记忆。重复的不再记。"""
    init()
    note = " ".join((note or "").split())
    if not note:
        raise ValueError("记忆内容不能为空")
    if len(note) > 500:
        raise ValueError("单条记忆请控制在 500 字以内，只记结论")
    f = CUSTOMERS / f"{slug(customer)}.md" if customer.strip() else GENERAL
    old = f.read_text("utf-8") if f.exists() else ""
    if note in old:
        return "已经记过了"
    with f.open("a", encoding="utf-8") as fh:
        fh.write(("" if not old or old.endswith("\n") else "\n") + f"- [{datetime.now():%Y-%m-%d}] {note}\n")
    return f"已记住（{'客户 ' + customer if customer.strip() else '通用'}）"


def prompt_block(customer: str = "") -> str:
    """拼进系统提示词的记忆内容。"""
    parts = []
    g = general().strip()
    if g:
        parts.append("## 通用记忆\n" + g[-MAX_INJECT:])
    c = customer_notes(customer).strip()
    if c:
        parts.append(f"## 客户「{customer}」的备注\n" + c[-MAX_INJECT:])
    names = [x["name"] for x in list_customers()]
    if names and not c:
        parts.append("## 有备注的客户（需要时用 recall 工具查看）\n" + "、".join(names[:80]))
    return "\n\n".join(parts)


# ---------------- 模板指纹 ----------------
def fingerprint(path: Path) -> str:
    """模板指纹：Excel 取「各 sheet 名 + 所有非空单元格的位置和文字」做哈希，
    这样同一份空白模板无论文件名、修改时间怎么变，指纹都相同。读不了的格式退回文件哈希。"""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=False)
        h = hashlib.sha1()
        for ws in wb.worksheets:
            h.update(f"#{ws.title}\n".encode())
            for row in ws.iter_rows():
                for c in row:
                    if c.value not in (None, ""):
                        h.update(f"{c.coordinate}={str(c.value).strip()}\n".encode())
        return h.hexdigest()[:12]
    except Exception:
        return "f" + hashlib.sha1(Path(path).read_bytes()).hexdigest()[:11]


# ---------------- 经验库（playbook） ----------------
def _meta(d: Path) -> dict | None:
    try:
        m = json.loads((d / "meta.json").read_text("utf-8"))
        m["id"] = d.name
        return m
    except Exception:
        return None


def list_playbooks(with_text=False) -> list[dict]:
    init()
    out = []
    for d in sorted(PLAYBOOKS.iterdir()):
        m = _meta(d) if d.is_dir() else None
        if m:
            if with_text:
                pb = d / "playbook.md"
                m["playbook"] = pb.read_text("utf-8") if pb.exists() else ""
                m["has_script"] = (d / "fill.py").exists()
            out.append(m)
    return sorted(out, key=lambda m: m.get("updated", ""), reverse=True)


def find(customer: str, fp: str) -> list[dict]:
    """找可用经验：模板指纹完全一致的排最前（同客户优先），其次是同客户的其他模板。"""
    cs = slug(customer) if (customer or "").strip() else None
    hits = []
    for m in list_playbooks():
        same_fp, same_c = m.get("fingerprint") == fp, cs is not None and slug(m.get("customer", "")) == cs
        if same_fp or same_c:
            m["match"] = "exact" if same_fp else "customer"
            hits.append((0 if same_fp and same_c else 1 if same_fp else 2, m))
    return [m for _, m in sorted(hits, key=lambda x: x[0])][:3]


def attach(wd: Path, customer: str, template: Path | None) -> tuple[list[dict], str]:
    """把匹配到的经验复制进任务工作目录的 experience/，返回（命中的经验, 给 Agent 的说明文字）。"""
    fp = fingerprint(template) if template and template.exists() else ""
    hits = find(customer, fp)
    if not hits:
        return [], ""
    lines = ["\n本次任务有可复用的历史经验（均为人工确认正确后沉淀），已放在 experience/ 下。先读它们再动手："]
    for i, m in enumerate(hits, 1):
        src, dst = PLAYBOOKS / m["id"], wd / "experience" / str(i)
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(src, dst)
        if m["match"] == "exact":
            how = "模板指纹与本次模板完全一致：按 playbook.md 的映射直接填，有 fill.py 就改参数复用；仍要回读核对和做单单一致校验"
        else:
            how = "同一客户、但模板与本次不同：只作参考（单证版式、取数习惯大概率相同），单元格位置必须按本次模板重新确认"
        lines.append(f"- experience/{i}/ ：客户「{m.get('customer') or '未指定'}」，模板 {m.get('template_name', '')}，已用 {m.get('uses', 0)} 次。{how}。")
        m["uses"] = m.get("uses", 0) + 1
        m["last_used"] = _now()
        (src / "meta.json").write_text(json.dumps({k: v for k, v in m.items() if k not in ("id", "match")}, ensure_ascii=False, indent=1), "utf-8")
    lines.append("如果发现经验里的说法与本次实际不符，以本次实际为准，并在问题清单里指出哪条经验过时了。")
    return hits, "\n".join(lines)


def save_playbook(*, customer: str, template: Path, output: Path | None, playbook: str,
                  script: Path | None = None, source: str = "") -> dict:
    """沉淀一份模板经验。同一客户 + 同一模板指纹会更新原有经验而不是新建。"""
    init()
    if len((playbook or "").strip()) < 80:
        raise ValueError("playbook 内容太短，至少要写清栏位→单元格映射和取数规律")
    if not template.is_file():
        raise ValueError(f"找不到模板文件：{template.name}")
    fp = fingerprint(template)
    d = PLAYBOOKS / f"{fp}__{slug(customer)}"
    d.mkdir(parents=True, exist_ok=True)
    old = _meta(d) or {}
    (d / "playbook.md").write_text(playbook.strip() + "\n", "utf-8")
    shutil.copy(template, d / ("template" + template.suffix.lower()))
    if output and output.is_file():
        shutil.copy(output, d / ("example_output" + output.suffix.lower()))
    if script and script.is_file():
        shutil.copy(script, d / "fill.py")
    meta = {"customer": (customer or "").strip(), "fingerprint": fp, "template_name": template.name,
            "created": old.get("created", _now()), "updated": _now(), "uses": old.get("uses", 0),
            "sources": (old.get("sources", []) + ([source] if source else []))[-20:]}
    (d / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1), "utf-8")
    return {**meta, "id": d.name}


def delete_playbook(pid: str) -> bool:
    if not re.fullmatch(r"[0-9a-f]{12}__[\w-]{1,60}", pid):
        return False
    d = PLAYBOOKS / pid
    if not d.is_symlink() and d.is_dir() and d.resolve().parent == PLAYBOOKS.resolve():
        shutil.rmtree(d)
        return True
    return False
