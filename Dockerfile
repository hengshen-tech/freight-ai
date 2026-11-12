# 货代 AI 工作台
# 国内网络构建慢时可指定镜像源，例如：
#   docker compose build --build-arg APT_MIRROR=mirrors.aliyun.com --build-arg PIP_INDEX_URL=https://mirrors.aliyun.com/pypi/simple/
# 拉不到 Docker Hub 时，可以把基础镜像换成镜像站的地址：--build-arg BASE_IMAGE=<镜像站>/library/python:3.12-slim
ARG BASE_IMAGE=python:3.12-slim
FROM ${BASE_IMAGE}

ARG APT_MIRROR=""
ARG PIP_INDEX_URL=""
# 额外的系统软件包（空格分隔），例如 OCR： --build-arg EXTRA_APT="tesseract-ocr tesseract-ocr-chi-sim"
ARG EXTRA_APT=""

# LibreOffice：xlsx Skill 用它重算并校验 Excel 公式
RUN if [ -n "$APT_MIRROR" ]; then sed -i "s|deb.debian.org|$APT_MIRROR|g" /etc/apt/sources.list.d/debian.sources; fi \
 && apt-get update \
 && apt-get install -y --no-install-recommends libreoffice-calc tzdata $EXTRA_APT \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
# Python 依赖和 Agent 处理单证用的库都在构建时装好；要加库就写进 requirements-extra.txt
COPY requirements.txt requirements-extra.txt ./
RUN pip install --no-cache-dir ${PIP_INDEX_URL:+-i "$PIP_INDEX_URL"} -r requirements.txt -r requirements-extra.txt

COPY . .

# 以非 root 用户运行；data/ 是唯一可写目录（数据库、任务文件、设置、Agent 会话记录）
RUN useradd -m -u 1000 app && mkdir -p /app/data && chown -R app:app /app/data
USER app
ENV HOME=/home/app \
    CLAUDE_CONFIG_DIR=/app/data/.claude \
    TZ=Asia/Shanghai \
    PYTHONUNBUFFERED=1

VOLUME /app/data
EXPOSE 8000
HEALTHCHECK --interval=10s --timeout=5s --start-period=20s --retries=6 \
  CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8000/api/health')" || exit 1
CMD ["uvicorn", "backend.app:app", "--host", "0.0.0.0", "--port", "8000"]
