"""模型连接设置与公司规则，保存在 data/ 下（Docker 里是持久卷）。"""
import json
import os

from . import db

FILE = db.DATA / "settings.json"
RULES = db.DATA / "company_rules.md"
DEFAULTS = {"model": "", "api_key": "", "base_url": "", "auth_mode": "auto"}


def load() -> dict:
    s = dict(DEFAULTS)
    if FILE.exists():
        try:
            s.update({k: v for k, v in json.loads(FILE.read_text("utf-8")).items() if k in DEFAULTS})
        except Exception:
            pass
    return s


def save(data: dict) -> dict:
    s = load()
    for k in ("model", "base_url", "auth_mode"):
        if k in data:
            s[k] = str(data[k] or "").strip()
    s["base_url"] = s["base_url"].rstrip("/")
    if s["base_url"] and not s["base_url"].startswith(("http://", "https://")):
        raise ValueError("接口地址必须以 http:// 或 https:// 开头")
    if s["auth_mode"] not in ("auto", "api_key", "bearer"):
        s["auth_mode"] = "auto"
    if data.get("clear_key"):
        s["api_key"] = ""
    elif (data.get("api_key") or "").strip():      # 留空表示不修改已保存的 key
        s["api_key"] = data["api_key"].strip()
    FILE.write_text(json.dumps(s, ensure_ascii=False, indent=1), "utf-8")
    try:
        os.chmod(FILE, 0o600)
    except OSError:
        pass
    return public()


def public() -> dict:
    """返回给前端的设置：永远不回传完整 key。"""
    s = load()
    key = s.pop("api_key")
    s["key_set"] = bool(key)
    s["key_hint"] = ("…" + key[-4:]) if len(key) >= 8 else ""
    s["env_key"] = bool(os.getenv("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_AUTH_TOKEN"))
    s["env_model"] = os.getenv("CLAUDE_MODEL", "")
    s["env_base_url"] = os.getenv("ANTHROPIC_BASE_URL", "")
    return s


def configured() -> bool:
    p = public()
    return p["key_set"] or p["env_key"]


def model() -> str:
    return load()["model"] or os.getenv("CLAUDE_MODEL", "")


def agent_env() -> dict:
    """传给 Claude Agent SDK 子进程的环境变量。界面里填的值优先于容器环境变量。"""
    s, env = load(), {}
    if s["base_url"]:
        env["ANTHROPIC_BASE_URL"] = s["base_url"]
    if s["api_key"]:
        mode = s["auth_mode"]
        if mode == "auto":     # 官方接口用 x-api-key；自定义网关一般用 Bearer
            mode = "bearer" if s["base_url"] else "api_key"
        if mode == "bearer":
            env["ANTHROPIC_AUTH_TOKEN"], env["ANTHROPIC_API_KEY"] = s["api_key"], ""
        else:
            env["ANTHROPIC_API_KEY"], env["ANTHROPIC_AUTH_TOKEN"] = s["api_key"], ""
    m = model()
    if m:
        env["ANTHROPIC_MODEL"] = m
        if s["base_url"] or os.getenv("ANTHROPIC_BASE_URL"):
            # 走自定义网关时，后台小模型等各档位也指向同一个模型名，避免网关不认识默认模型
            for tier in ("HAIKU", "SONNET", "OPUS"):
                env[f"ANTHROPIC_DEFAULT_{tier}_MODEL"] = m
    return env


def rules() -> str:
    return RULES.read_text("utf-8") if RULES.exists() else ""


def save_rules(text: str):
    RULES.write_text(text or "", "utf-8")
