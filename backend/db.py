"""SQLite 存储：运单（shipment）与待办（todo）。"""
import os
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA = Path(os.getenv("DATA_DIR") or ROOT / "data")      # 测试时可用 DATA_DIR 指到临时目录
DATA.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA / "freight.db"

# 出口/进口各自的作业节点，工作台按这个顺序推进
STAGES = {
    "export": ["接单", "订舱", "单证收集", "报关申报", "海关放行", "已开船", "已到港", "完成"],
    "import": ["接单", "换单", "单证收集", "报关申报", "海关放行", "提货派送", "完成"],
}
FIELDS = ["ref", "direction", "customer", "counterparty", "pol", "pod", "mode", "incoterm",
          "goods", "status", "exception", "etd", "cutoff", "bl_no", "note"]


def conn():
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS shipments(
          id INTEGER PRIMARY KEY AUTOINCREMENT, ref TEXT UNIQUE, direction TEXT DEFAULT 'export',
          customer TEXT, counterparty TEXT, pol TEXT, pod TEXT, mode TEXT DEFAULT '海运整箱',
          incoterm TEXT, goods TEXT, status TEXT DEFAULT '接单', exception TEXT DEFAULT '',
          etd TEXT, cutoff TEXT, bl_no TEXT, note TEXT DEFAULT '', updated TEXT);
        CREATE TABLE IF NOT EXISTS todos(
          id INTEGER PRIMARY KEY AUTOINCREMENT, shipment_id INTEGER, text TEXT,
          done INTEGER DEFAULT 0, created TEXT);
        """)
        if c.execute("SELECT COUNT(*) FROM shipments").fetchone()[0] == 0:
            _seed(c)


def _seed(c):
    """首次启动写入几票演示数据，方便看工作台效果。可在界面里删除或直接删 data/freight.db。"""
    d = lambda n: (date.today() + timedelta(days=n)).isoformat()
    demo = [
        ("EX-DEMO-001", "export", "演示客户·宁波恒益电器", "NORDLICHT HANDELS GMBH", "宁波", "汉堡", "海运整箱", "FOB", "LED 灯具", "单证收集", "", d(6), d(2), ""),
        ("EX-DEMO-002", "export", "演示客户·苏州瑞禾纺织", "PACIFIC LINEN CO.", "上海", "洛杉矶", "海运拼箱", "CIF", "棉制床单", "报关申报", "", d(4), d(0), "DEMO2611001"),
        ("EX-DEMO-003", "export", "演示客户·东莞嘉拓五金", "BRISA FERRAGENS LTDA", "深圳盐田", "桑托斯", "海运整箱", "CFR", "不锈钢紧固件", "报关申报", "海关查验", d(3), d(-1), "DEMO2611002"),
        ("IM-DEMO-001", "import", "演示客户·青岛海悦食品", "AUSTRAL DAIRY PTY", "墨尔本", "青岛", "海运整箱", "CIF", "全脂奶粉", "换单", "", d(-12), "", "DEMO2610087"),
        ("EX-DEMO-004", "export", "演示客户·宁波恒益电器", "NORDLICHT HANDELS GMBH", "宁波", "鹿特丹", "空运", "EXW", "电源适配器", "已开船", "", d(-2), d(-5), "DEMO2611000"),
    ]
    for r in demo:
        c.execute("""INSERT INTO shipments(ref,direction,customer,counterparty,pol,pod,mode,incoterm,
                     goods,status,exception,etd,cutoff,bl_no,updated) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (*r, datetime.now().isoformat(timespec="seconds")))
    for sid, text in [(1, "催客户补齐箱单和申报要素"), (3, "配合海关查验：联系场站安排掏箱"), (4, "向船代换取提货单")]:
        c.execute("INSERT INTO todos(shipment_id,text,created) VALUES(?,?,?)",
                  (sid, text, datetime.now().isoformat(timespec="seconds")))


def list_shipments(keyword: str = ""):
    with conn() as c:
        rows = [dict(r) for r in c.execute("SELECT * FROM shipments ORDER BY id DESC")]
    if keyword:
        rows = [r for r in rows if keyword.lower() in " ".join(str(v) for v in r.values()).lower()]
    for r in rows:
        st = STAGES.get(r["direction"], STAGES["export"])
        r["stages"] = st
        r["stage_index"] = st.index(r["status"]) if r["status"] in st else 0
    return rows


def get_shipment(key):
    """key 可以是 id 或业务编号 ref。"""
    for r in list_shipments():
        if str(r["id"]) == str(key) or r["ref"] == str(key):
            return r
    return None


def create_shipment(data: dict):
    data = {k: v for k, v in data.items() if k in FIELDS and v not in (None, "")}
    data.setdefault("direction", "export")
    if not data.get("ref"):
        prefix = "IM" if data["direction"] == "import" else "EX"
        data["ref"] = f"{prefix}-{datetime.now():%y%m%d-%H%M%S}"
    data["updated"] = datetime.now().isoformat(timespec="seconds")
    cols = ",".join(data)
    with conn() as c:
        cur = c.execute(f"INSERT INTO shipments({cols}) VALUES({','.join('?' * len(data))})", list(data.values()))
    return get_shipment(cur.lastrowid)   # 提交后再读


def update_shipment(key, data: dict):
    s = get_shipment(key)
    if not s:
        return None
    data = {k: v for k, v in data.items() if k in FIELDS and v is not None}
    if "status" in data and data["status"] not in s["stages"]:
        raise ValueError(f"状态必须是其中之一：{s['stages']}")
    if data:
        data["updated"] = datetime.now().isoformat(timespec="seconds")
        sets = ",".join(f"{k}=?" for k in data)
        with conn() as c:
            c.execute(f"UPDATE shipments SET {sets} WHERE id=?", [*data.values(), s["id"]])
    return get_shipment(s["id"])


def delete_shipment(sid: int):
    with conn() as c:
        c.execute("DELETE FROM shipments WHERE id=?", (sid,))
        c.execute("DELETE FROM todos WHERE shipment_id=?", (sid,))


def list_todos(include_done=False):
    q = """SELECT t.*, s.ref FROM todos t LEFT JOIN shipments s ON s.id=t.shipment_id
           {} ORDER BY t.done, t.id DESC""".format("" if include_done else "WHERE t.done=0")
    with conn() as c:
        return [dict(r) for r in c.execute(q)]


def add_todo(shipment_id, text: str):
    with conn() as c:
        c.execute("INSERT INTO todos(shipment_id,text,created) VALUES(?,?,?)",
                  (shipment_id, text, datetime.now().isoformat(timespec="seconds")))


def toggle_todo(tid: int):
    with conn() as c:
        c.execute("UPDATE todos SET done=1-done WHERE id=?", (tid,))


def dashboard():
    ships = list_shipments()
    today = date.today().isoformat()
    active = [s for s in ships if s["status"] != "完成"]
    urgent = [s for s in active if s["cutoff"] and s["cutoff"] <= today
              and s["stage_index"] < s["stages"].index("海关放行")]
    return {
        "kpi": {
            "active": len(active),
            "to_declare": len([s for s in active if s["status"] in ("单证收集", "报关申报")]),
            "exception": len([s for s in active if s["exception"]]),
            "cutoff_due": len(urgent),
        },
        "shipments": ships,
        "todos": list_todos(),
    }
