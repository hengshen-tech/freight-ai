"""成品校验（确定性部分）：不靠模型，直接打开成品 Excel 核对。

复核分两层：
1. 这里的程序校验 —— 模板原有内容有没有被改动、填表 Agent 声称填了的格子是不是真的填了。
2. agents.py 里的复核 Agent —— 一个独立的、没看过填表过程的 Agent，从原始单证重新取数，逐栏核对成品。
"""
import hashlib
import re
import shutil
from pathlib import Path

CELL = re.compile(r"^(?:'?(?P<sheet>[^'!]+)'?!)?\$?(?P<col>[A-Za-z]{1,3})\$?(?P<row>\d+)$")
EXCEL = {".xlsx", ".xlsm"}


def _check(item, status, detail, cell=""):
    return {"item": item, "status": status, "detail": detail, "cell": cell, "by": "程序校验"}


def _norm(v) -> str:
    """把值规整成便于比较的文本：数字去掉千分位和多余的 0，文本去掉空白差异。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return f"{float(v):.6f}".rstrip("0").rstrip(".")
    s = " ".join(str(v).split())
    t = s.replace(",", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", t):
        return f"{float(t):.6f}".rstrip("0").rstrip(".")
    return s


def first_file(folder: Path, exts=None):
    files = sorted(p for p in folder.glob("*") if p.is_file() and (not exts or p.suffix.lower() in exts))
    return files[0] if files else None


def precheck(wd: Path, result: dict) -> list[dict]:
    """返回一组核对结果。成品或模板不是 xlsx 时只给一条说明，交给复核 Agent。"""
    tpl, out = first_file(wd / "template", EXCEL), first_file(wd / "output", EXCEL)
    if not tpl or not out:
        return [_check("程序校验", "ok", "模板或成品不是 xlsx，跳过逐格校验，由复核员人工式核对。")]
    try:
        from openpyxl import load_workbook
        wt, wo = load_workbook(tpl), load_workbook(out)
        values = load_workbook(out, data_only=True)
    except Exception as e:
        return [_check("成品能否打开", "error", f"成品或模板无法打开：{e}")]

    checks = []
    missing = [n for n in wt.sheetnames if n not in wo.sheetnames]
    if missing:
        checks.append(_check("工作表完整", "error", "成品里缺少模板的工作表：" + "、".join(missing)))

    # 1) 模板原有的文字 / 公式不应被改动
    changed, filled, shifted = [], 0, []
    for ws_t in wt.worksheets:
        if ws_t.title not in wo.sheetnames:
            continue
        ws_o = wo[ws_t.title]
        if ws_o.max_row != ws_t.max_row:           # 明细行不够时会插行，位置整体下移，逐格比对没有意义
            shifted.append(ws_t.title)
            continue
        for row in ws_t.iter_rows():
            for c in row:
                o = ws_o[c.coordinate].value
                if c.value not in (None, ""):
                    if _norm(o) != _norm(c.value):
                        changed.append((f"{ws_t.title}!{c.coordinate}", c.value, o))
                elif o not in (None, ""):
                    filled += 1
    # 公式文本用于检查模板；缓存值用于检查计算错误，两者必须分别读取。
    # 独立扫描所有成品工作表，插行或新增工作表也不能跳过公式错误。
    for ws_o in values.worksheets:
        for row in ws_o.iter_rows():
            for c in row:
                if c.data_type == "e":
                    checks.append(_check("公式错误", "error", f"单元格里是错误值 {c.value}", f"{ws_o.title}!{c.coordinate}"))
    for where, old, new in changed[:8]:
        checks.append(_check("模板原有内容被改动", "error", f"模板里是「{old}」，成品里变成「{new if new not in (None, '') else '空'}」", where))
    if len(changed) > 8:
        checks.append(_check("模板原有内容被改动", "error", f"另有 {len(changed) - 8} 处模板内容被改动"))
    for name in shifted:
        checks.append(_check("行数变化", "warn", f"工作表「{name}」的行数与模板不同（可能因商品行较多插了行），请留意表尾合计和签章区是否完好"))
    if not changed and not shifted:
        checks.append(_check("模板原有内容", "ok", f"标签、公式均未被改动；共填入 {filled} 个原本空白的单元格"))

    # 2) 填表 Agent 声称填了的格子，是否真的是那个值
    claimed = same = 0
    for f in (result or {}).get("fields") or []:
        m = CELL.match(str(f.get("cell") or "").strip())
        if not m:
            continue                                # 区域、文字描述等无法定位的不校验
        sheet = m["sheet"] if m["sheet"] in wo.sheetnames else wo.sheetnames[0]
        coord = f"{m['col'].upper()}{m['row']}"
        actual, want = wo[sheet][coord].value, f.get("value")
        claimed += 1
        where = f"{sheet}!{coord}"
        if actual in (None, "") and want not in (None, ""):
            checks.append(_check(f.get("field") or "栏位", "error", f"说明里称已填「{want}」，但成品里这个单元格是空的", where))
        elif isinstance(actual, str) and actual.startswith("="):
            same += 1                               # 公式格：值要重算后才有，交给复核员
        elif _norm(actual) != _norm(want) and _norm(want) not in _norm(actual) and _norm(actual) not in _norm(want):
            checks.append(_check(f.get("field") or "栏位", "warn", f"说明里写的是「{want}」，成品里实际是「{actual}」", where))
        else:
            same += 1
    if claimed:
        checks.append(_check("说明与成品一致", "ok" if same == claimed else "warn", f"说明里列出的 {claimed} 个栏位中，{same} 个与成品实际内容一致"))
    wt.close()
    wo.close()
    values.close()
    return checks


def verdict(checks: list[dict]) -> str:
    st = {c.get("status") for c in checks}
    return "fail" if "error" in st else "warn" if "warn" in st else "pass"


# —— 复核期间保护成品：复核员只读，万一动了就恢复 ——
def snapshot(wd: Path) -> dict:
    snap = wd / ".review_snapshot"
    if snap.exists():
        shutil.rmtree(snap)
    shutil.copytree(wd / "output", snap)
    return {p.name: hashlib.sha1(p.read_bytes()).hexdigest() for p in snap.glob("*") if p.is_file()}


def restore_if_changed(wd: Path, hashes: dict) -> list[str]:
    out, snap, touched = wd / "output", wd / ".review_snapshot", []
    for name, h in hashes.items():
        f = out / name
        if not f.exists() or hashlib.sha1(f.read_bytes()).hexdigest() != h:
            shutil.copy(snap / name, f)
            touched.append(name)
    for f in out.glob("*"):
        if f.is_file() and f.name not in hashes:    # 复核员新建的文件不算成品
            f.unlink()
            touched.append(f.name)
    shutil.rmtree(snap, ignore_errors=True)
    return touched
