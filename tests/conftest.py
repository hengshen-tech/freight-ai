"""测试不需要模型：用一个假的 query() 扮演 Agent，验证除模型以外的整条链路。
运行：pip install pytest httpx && pytest -q"""
import os
import sys
import tempfile
from pathlib import Path

os.environ["DATA_DIR"] = tempfile.mkdtemp(prefix="freight-test-")     # 必须在导入 backend 之前设置
os.environ["ANTHROPIC_API_KEY"] = "test-key"
os.environ["AGENT_LOG"] = "full"
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
