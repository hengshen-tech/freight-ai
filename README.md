# 货代 AI 工作台（MVP）

面向货运代理 / 报关业务的 AI 原生工作台。AI 能力由 **Claude Agent SDK** 驱动：
Excel 的读写由 Agent 通过 **Claude Code 的 `xlsx` Skill** 完成，报关业务规则放在 `customs-declaration` Skill 里。

> **当前状态**：页面、接口、任务流程、复核、记忆等都有自动化测试覆盖（用假 Agent 扮演模型）。
> 但**还没有用真实模型跑通过**，Docker 镜像也**还没有实际构建过**——开发环境里没有可用的 API Key，也拉不到基础镜像。
> 第一次部署时请先用 `samples/` 里的样例走一遍，再接入真实单证。

## 功能

| 模块 | 说明 |
|---|---|
| 工作台 | 在办运单、作业节点进度、截关预警、查验 / 异常、待办；顶部一句话指令直接交给 AI |
| AI 助手 | 多轮对话，带对话列表，可拖入单证。能查票、建单、改状态、加待办，也能在对话里建单、审单、填报关单并给出下载 |
| 报关单智能填制 | 上传多份原始单证 + 一份报关模板 → 填表 → 独立复核 →（有错则自动改正再复核）→ 下载。每个栏位带来源和把握程度，附问题清单和复核清单；完成后可继续让 AI 修改，核对后点「确认无误」 |
| 单证一致性审核 | 发票 / 箱单 / 合同 / 提单交叉核对（金额、数量、件重、抬头、条款） |
| HS 归类助手 | 候选编码、归类依据、申报要素、监管条件 |
| 任务记录 | 每次 AI 任务的过程与结果留档；可按客户、类型筛选和搜索 |
| 设置 | 模型名称、API Key、接口地址、鉴权方式、连接测试；公司制单规则 |

后台自动进行、操作员无感知的部分：从单证识别客户、记住操作员交代过的要点、把确认无误的做法沉淀成模板经验并在下次复用。

## 启动

```bash
./start.sh            # 构建并启动，就绪后打印访问地址
./start.sh --cn       # 同上，构建时使用国内 apt / pip 镜像源
./start.sh logs       # 跟随日志
./start.sh status     # 运行状态
./start.sh restart    # 重启（不重新构建）
./start.sh stop       # 停止，数据保留
./start.sh local      # 不用 Docker，在本机用 Python 直接运行
```

启动后打开 http://localhost:8000 → 「设置」→ 填模型名称 / API Key / 接口地址 → 「保存并测试连接」。
然后进入「报关单智能填制」，上传 `samples/` 里的三份样例单证（虚构数据），勾选「使用内置示例模板」试跑。
首次启动会写入几票演示运单，在工作台里可以逐条删除。

`start.sh` 只是对 `docker compose` 的封装，直接用 `docker compose up -d --build` 也可以。

- **改了代码或依赖后**：再执行一次 `./start.sh` 即可重新构建。
- **拉不到 Docker Hub**：`BASE_IMAGE=<镜像站>/library/python:3.12-slim ./start.sh --cn`。
- **镜像较大**：里面装了 LibreOffice（xlsx Skill 用它重算公式），首次构建要几分钟。
- **本机运行（`local`）**：需要 Python 3.10+；没有 LibreOffice 时公式无法重算校验，其余功能不受影响。

### 配置项

都可以写进项目目录下的 `.env` 文件，然后重新执行 `./start.sh`。模型相关的三项也可以不填，启动后在网页「设置」里配置（网页里填的优先）。

| 变量 | 默认 | 说明 |
|---|---|---|
| `PORT` | 8000 | 对外端口 |
| `ANTHROPIC_API_KEY` | 空 | 模型的 Key |
| `ANTHROPIC_BASE_URL` | 空 | 接口地址，留空为 Anthropic 官方接口 |
| `CLAUDE_MODEL` | 空 | 模型名称，留空用默认模型 |
| `REVIEW` | on | 填制后是否由独立的复核 Agent 再核一遍（on / off） |
| `REVIEW_AUTOFIX` | 1 | 复核发现错误时自动改正并重新复核的轮数，0 为只报告不改 |
| `REVIEW_MODEL` | 空 | 复核用的模型，留空则与填表相同 |
| `AGENT_LOG` | full | Agent 日志详细程度：full / brief / off |
| `AGENT_LOG_MAX` | 2000 | 单条日志里输入 / 输出最多打印的字符数 |
| `BASE_IMAGE` | python:3.12-slim | 构建用的基础镜像 |

### 接口地址

Claude Agent SDK 使用 Anthropic Messages API。「接口地址」留空即官方接口；填自建网关或中转时，对方必须兼容这个协议，地址不要带 `/v1` 结尾。
鉴权方式「自动」在官方接口用 `x-api-key`，自定义地址用 `Bearer`，连不上时可手动切换。

### 数据与备份

所有数据都在 Docker 命名卷 `freight-data`（容器内 `/app/data`）里，删容器不丢，`docker compose down -v` 才会清空。

```
data/
  freight.db          运单、待办
  settings.json       模型设置（含 API Key，明文）
  company_rules.md    公司规则
  jobs/<任务ID>/      每个任务的工作目录：input/ template/ output/ work/ result.json review.json job.json
  chat/<对话ID>/      每个对话的工作目录和记录
  memory/             长期记忆与模板经验
  .claude/            Agent 的会话记录（续接对话、追加修改时要用）
```

备份：`docker run --rm -v freight-data:/data -v "$PWD":/backup alpine tar czf /backup/freight-data.tgz -C /data .`

### 日志

```
11:06:55 [agent] 1112110655-7f28 START  kind=fill model=... resume=否 customer=... cwd=...
11:06:55 [agent] 1112110655-7f28 TOOL   Bash  {"command": "python work/fill.py", ...}
11:06:55 [agent] 1112110655-7f28 RESULT Bash  已写入 28 个单元格
11:06:58 [agent] 1112110655-7f28 DONE   subtype=success turns=14 85.2s cost=$0.4120 in=... out=...
```

每行带任务 ID。各标签的含义：

| 标签 | 含义 |
|---|---|
| `START` / `PROMPT` | 一轮 Agent 开始，以及发给它的任务说明 |
| `TEXT` | 模型说的话 |
| `TOOL` / `RESULT` / `ERROR` | 工具调用及输入、工具输出、工具报错 |
| `DONE` / `FAILED` | 一轮结束（轮数、耗时、花费）/ 失败原因 |
| `CUSTOM` | 从单证识别出的客户 |
| `REVIEW` / `R.TOOL` / `R.RES` / `R.TEXT` | 复核的开始与结论、复核 Agent 的工具调用、输出、说的话 |
| `LEARN` | 后台总结模板经验 |

日志里会出现单证内容，注意存放和访问权限；API Key 会被替换成 `***`；compose 里已设置滚动（20MB × 5）。

### 预装的库

镜像里已装好 Agent 处理单证常用的库：openpyxl、pandas、xlrd、markitdown、pymupdf、pdfplumber、pypdf、python-docx、Pillow，以及 LibreOffice。
要加别的库：写进 `requirements-extra.txt`（一行一个）后重新 `./start.sh`。
要加系统软件（比如 OCR）：`docker compose build --build-arg EXTRA_APT="tesseract-ocr tesseract-ocr-chi-sim" && docker compose up -d`。

## 代码结构

```
start.sh             启动脚本
backend/app.py       FastAPI 接口（运单、任务、对话、设置、记忆、下载）
backend/agents.py    Claude Agent SDK 封装：提示词、Skill 加载、工具、填表 → 复核 → 改正流程、对话、日志
backend/review.py    成品的程序校验（模板是否被改动、声称填的格子是否真的填了）
backend/memory.py    长期记忆与模板经验库
backend/settings.py  模型连接设置与公司规则
backend/db.py        SQLite：运单、待办
frontend/index.html  单文件前端，无需构建
.claude/skills/xlsx                  Anthropic 官方 xlsx Skill（Apache-2.0，来自 github.com/anthropics/skills）
.claude/skills/customs-declaration   报关填制 / 审单的通用规则（按通用报关常识编写，建议请报关员审一遍）
samples/             样例单证与报关单模板；scripts/make_samples.py 可重新生成
tests/               不需要模型的测试
Dockerfile, docker-compose.yml, requirements.txt, requirements-extra.txt
```

运行测试：`pip install -r requirements.txt pytest httpx && pytest -q`。测试用一个假的 Agent 扮演模型，验证除模型本身以外的整条链路。

## 工作原理

### Agent 怎么运行

每个任务、每个对话都有独立的工作目录，Agent 以它为 `cwd` 运行：

```python
ClaudeAgentOptions(
    cwd=work_dir,
    setting_sources=["project"],                 # 从 <cwd>/.claude/skills 发现 Skill
    skills=["xlsx", "customs-declaration"],      # 只开放这两个
    tools=["Skill", "Read", "Write", "Edit", "Bash", "Glob", "Grep"],
    allowed_tools=[...],                         # 服务端无人值守，预先批准
    mcp_servers={"tms": ..., "mem": ...},        # 进程内工具：运单系统（对话用）、记忆与经验
    env={...}, model=...,                        # 来自「设置」页
    resume=session_id,                           # 对话 / 追加修改时续接同一会话
)
```

Agent 先加载两个 Skill，再用 Bash 跑 Python 读取单证、复制模板并填写、回读校验。
表单式任务最后写出 `result.json`（结论、客户、已填栏位、问题清单）供页面展示；对话则直接回答并附上产出文件。

进提示词的内容有三层：`customs-declaration` Skill 里的通用规则、「设置 → 公司规则」里人写的规定、AI 自己积累的记忆。
适配你们自己的模板不用改代码：直接上传模板即可；固定习惯写进公司规则。

### 客户识别

报关单填制和单证审核时可以不填客户：Agent 从单证里识别（出口一般是卖方 / 发货人，进口一般是买方 / 收货人），
并对齐到系统里已有的客户名（运单、历史任务、客户备注、模板经验里出现过的），避免同一家公司出现多个写法。
识别出客户后它会自己去查这个客户的备注和模板经验。操作员填了客户则以操作员为准；识别错了可以在任务页点「更正」。
文件按任务存放，没有另外按客户复制一份；按客户查看用任务记录里的筛选。

### 独立复核

报关单填制（含之后的追加修改）完成后，后端自动再走一遍复核，全部结束任务才算完成：

1. **程序校验**（不靠模型）：打开成品，检查模板原有的标签和公式有没有被改动、有没有公式错误值、填表 Agent 声称填的每个格子是不是真的是那个值。
2. **复核 Agent**：一个全新的会话，看不到填表过程，只读。它从原始单证重新取数，逐栏核对成品，写出核对清单。它不能写记忆；复核期间成品如被改动会自动恢复。
3. **自动改正**：复核出现“错误”级问题时，把问题清单交回原来的填表会话改正，然后重新复核一次。仍有错误就原样展示，由人处理。

结果显示在任务页的「独立复核」卡片里。开启复核后，一次填制的耗时和花费大约翻倍；对话里让 AI 填的报关单不走这个流程。
商品行较多、模板被插行时，程序校验无法逐格比对，只给一条提醒。

### 记忆与模板经验

做法沿用 Anthropic 自己的思路——记忆就是一个目录里的 Markdown 文件：每次任务开始时带上，Agent 学到东西再写回去，人可以看和改。

```
data/memory/
  MEMORY.md                    通用记忆：每次任务都带进系统提示词
  customers/<客户>.md          客户备注：任务关联到该客户时带上
  playbooks/<模板指纹>__<客户>/
      playbook.md              栏位→单元格映射、取数规律、踩过的坑
      fill.py                  当时的填表脚本（如有）
      template.*  example_output.*  meta.json
```

- **记忆**：操作员在对话或任务里说了以后还用得上的话（固定做法、偏好、客户特殊要求、对 AI 的纠正），Agent 在后台记一条。
- **模板经验**：任务页点「确认无误」后，后台让 Agent 在同一会话里把这次的做法总结成 playbook。成败只写进日志；不点确认就不会保存；确认后又让 AI 修改的，需要重新确认。
- **复用**：新任务进来时，后端给模板算指纹（sheet 名 + 所有标签文字的位置和内容，与文件名无关）。
  指纹完全一致 → 把 playbook 和脚本放进任务目录，Agent 直接按映射填；同一客户但模板变了 → 作为参考；都没有 → 按常规流程。
- **纠正**：这些内容不在侧边栏里。需要查看、修改、删除时，从「设置」页底部的“高级”链接进入。

## 已知限制

**上线前必须处理**

- **没有登录与权限**：任何能访问页面的人都能看到全部数据、改模型设置、删记忆。不要直接暴露到公网；至少放在内网或加一层反向代理鉴权。
- **Agent 没有硬隔离**：它有执行命令的权限并会处理客户上传的文件。容器已用非 root 用户运行、代码目录只读，但 Agent 仍能读到 `data/` 下其他任务的文件和保存的 API Key。建议给这个 Key 设置用量上限。
- **记忆可能被单证内容污染**：提示词要求把单证内容当数据而非指令，但这不是硬性保证。客户文件里的“请记住……”有可能被记下来并影响之后的任务，建议定期查看记忆。
- **AI 产出是草单**：申报前须由持证报关员复核；HS 归类仅供参考。

**目前没有的**

- 运行中的任务不能取消，没有超时、并发上限和单任务花费上限。
- 不能在页面上预览和修改成品 Excel，需要下载后查看。
- 运单不能编辑，也没有运单详情页（只能新建、推进、删除）。
- PDF 和扫描件可以上传，但没有专门的识别流程。
- 更正任务的客户不会迁移已沉淀的模板经验。
- 没有对接邮件、单一窗口、船公司或物流轨迹；没有费用对账和客户门户。
