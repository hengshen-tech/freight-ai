#!/usr/bin/env bash
# 货代 AI 工作台 —— 启动脚本
#
#   ./start.sh              构建并启动（Docker），等服务就绪后打印访问地址
#   ./start.sh --cn         同上，构建时使用国内 apt / pip 镜像源
#   ./start.sh stop         停止（数据保留）
#   ./start.sh restart      重启（不重新构建）
#   ./start.sh logs [行数]   跟随日志，默认从最后 200 行开始
#   ./start.sh status       查看运行状态
#   ./start.sh local        不用 Docker，直接在本机用 Python 启动
#
# 可用环境变量（也可以写进同目录的 .env 文件）：
#   PORT=8000               对外端口
#   BASE_IMAGE=...          基础镜像，拉不到 Docker Hub 时可换成镜像站地址
#   ANTHROPIC_API_KEY / ANTHROPIC_BASE_URL / CLAUDE_MODEL   不填则启动后在网页「设置」里配置
set -euo pipefail
cd "$(dirname "$0")"

NAME=freight-ai
say()  { printf '\033[1m%s\033[0m\n' "$*"; }
fail() { printf '\033[31m%s\033[0m\n' "$*" >&2; exit 1; }

# 端口：环境变量优先，其次 .env，默认 8000
if [ -z "${PORT:-}" ] && [ -f .env ]; then
  PORT=$(grep -E '^PORT=' .env | tail -1 | cut -d= -f2- | tr -d '"'"'"' \r' || true)
fi
PORT=${PORT:-8000}
export PORT

need_docker() {
  command -v docker >/dev/null 2>&1 || fail "没有找到 docker，请先安装 Docker（或用 ./start.sh local 在本机直接运行）。"
  docker info >/dev/null 2>&1 || fail "Docker 没有在运行，请先启动 Docker。"
  if docker compose version >/dev/null 2>&1; then DC=(docker compose)
  elif command -v docker-compose >/dev/null 2>&1; then DC=(docker-compose)
  else fail "没有找到 docker compose，请安装 Docker Compose 插件。"; fi
}

wait_ready() {
  say "等待服务就绪…"
  for _ in $(seq 1 90); do
    state=$(docker inspect --format '{{if .State.Health}}{{.State.Health.Status}}{{else}}{{.State.Status}}{{end}}' "$NAME" 2>/dev/null || echo missing)
    case "$state" in
      healthy) return 0 ;;
      exited|dead|missing) "${DC[@]}" logs --tail 40 || true; fail "容器没有正常启动（状态：$state），上面是最后的日志。" ;;
    esac
    sleep 2
  done
  "${DC[@]}" logs --tail 40 || true
  fail "等了 3 分钟服务仍未就绪，上面是最后的日志。"
}

cmd=${1:-start}
case "$cmd" in
  start|--cn|up)
    need_docker
    args=()
    if [ "$cmd" = "--cn" ] || [ "${2:-}" = "--cn" ]; then
      args+=(--build-arg APT_MIRROR=mirrors.aliyun.com --build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple/)
      say "使用国内镜像源构建"
    fi
    [ -n "${BASE_IMAGE:-}" ] && args+=(--build-arg "BASE_IMAGE=$BASE_IMAGE")
    say "构建镜像（首次需要几分钟）…"
    "${DC[@]}" build ${args[@]+"${args[@]}"}   # 这种写法兼容 macOS 自带的旧版 bash
    "${DC[@]}" up -d
    wait_ready
    say "已启动： http://localhost:$PORT"
    echo "首次使用：打开页面 → 「设置」→ 填模型名称 / API Key / 接口地址 → 「保存并测试连接」"
    echo "查看日志：./start.sh logs      停止：./start.sh stop"
    ;;
  stop|down)
    need_docker; "${DC[@]}" down; say "已停止。数据保留在卷 freight-data 里。" ;;
  restart)
    need_docker; "${DC[@]}" restart; wait_ready; say "已重启： http://localhost:$PORT" ;;
  logs)
    need_docker; "${DC[@]}" logs -f --tail "${2:-200}" ;;
  status|ps)
    need_docker; "${DC[@]}" ps ;;
  local)
    PY=$(command -v python3 || command -v python || true)
    [ -n "$PY" ] || fail "没有找到 Python（需要 3.10 及以上）。"
    "$PY" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)' || fail "Python 版本过低，需要 3.10 及以上。"
    if [ ! -d .venv ]; then say "创建虚拟环境 .venv 并安装依赖…"; "$PY" -m venv .venv; fi
    .venv/bin/pip install -q -r requirements.txt -r requirements-extra.txt
    command -v soffice >/dev/null 2>&1 || echo "提示：本机没有 LibreOffice（soffice），Excel 公式无法重算校验，其余功能不受影响。"
    say "启动： http://localhost:$PORT   （Ctrl+C 停止）"
    exec .venv/bin/python -m uvicorn backend.app:app --host 0.0.0.0 --port "$PORT"
    ;;
  -h|--help|help)
    sed -n '2,16p' "$0" | sed 's/^# \{0,1\}//' ;;
  *)
    fail "不认识的命令：$cmd（用 ./start.sh help 查看用法）" ;;
esac
