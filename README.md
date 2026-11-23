# 学术 PPT 模板与格式适配 Skill

把**固定论文/材料**套进**组织方提供的真实 pptx 模板**：识别真实的母版、版式、主题与占位符，写入正确的版式，保留图号与图表指向，等比缩放图表，守住可读字号下限，生成讲稿备注，并独立检查内容保真、格式继承和真实渲染。交付可编辑 `slides.pptx` + 页级证据映射 + 变更记录 + 真实 PDF/PNG 预览 + `result.json`。

- 触发条件与执行硬规则见 [SKILL.md](SKILL.md)。
- 从侦察证据编写计划的逐步方法见 [计划编写指南](references/plan_guide.md)。
- 交付合同、`result.json` 字段、状态与退出码见 [结果合同](references/result_contract.md)。
- 支持范围、明确不支持项与降级处置见 [限制与降级](references/limitations.md)。
- 合成接口示例见 [examples](examples/README.md)；全部标 `example=true` / EXAMPLE，**不是正式案例、预制答案、实测结果或评分**。

## 1. 提交件结构与安装

提交件根目录结构（与 `CONTRACT.md` §2 一致）：

```
SKILL.md                 # 触发边界与流程（YAML frontmatter: name + description）
README.md                # 本文件
GITHUB_URL.txt           # 由提交方维护：第 1 行公开链接，第 2 行可选相对路径
scripts/                 # common.py / discover.py / adapt_pptx.py / run.py
                         # verify_pptx.py / render_report.py / finalize.py
references/              # plan_guide.md / result_contract.md / limitations.md
examples/                # case.json / slides_plan.json / README.md（均 example=true）
```

安装：把本目录作为一个 Skill 放到运行环境的 `/skills/weekly-skill`（平台由组织方安装）。脚本只用 Python 3 标准库、离线、不写 `__pycache__`（以 `python -B` 运行）。下文以 `$SKILL` 表示该安装目录。

外部工具（`soffice`/`office-renderer`、`pdftoppm`、`pdffonts`、PDF 文本提取工具等）按运行环境探测；探测不到就如实记录 `skipped`/`failed`，**不伪造预览或字体状态**。

## 2. 环境、输入与输出

| 项目 | 约定 |
|---|---|
| 运行方式 | Python 3 标准库，离线；目标运行环境为 Linux 容器，脚本本身 Windows/POSIX 兼容 |
| 输入（只读） | `/input/case.json`、`/input/task.md` 与实际材料（模板、论文/材料、图、可选起始 deck） |
| 工作区 | `/app/work`：侦察结果、可编辑计划、审计报告 |
| 输出区 | `/app/output`：`result.json`、可编辑 deck、证据副本、真实预览 |
| 预算 | 64 model requests / 900 秒 / 单响应 16384 token |
| Skill 资源上限 | ≤50 个 UTF-8 文本文件，单文件 ≤500KB，合计 ≤2MB，**无符号链接**；二进制 pptx/PDF/PNG 不得打进 Skill 本体 |

输出清单（详见 [结果合同](references/result_contract.md)）：

| 路径 | 内容 |
|---|---|
| `/app/output/result.json` | **必需**；`case_id` + 原 `output_contract` 原样保留 + 状态、证据、逐页映射 |
| `/app/output/slides.pptx` | 默认可编辑 deck；必须含模板 master/layout/theme 部件 |
| `/app/output/slides_plan.json` | 页级证据映射对外副本（每页 `claims` + `evidence` + `locator`） |
| `/app/output/evidence/*.json`、`evidence/change_records.jsonl` | `fidelity_report` / `format_audit` / `render_report` / `sources` / `validation_report` / 变更记录 |
| `/app/output/preview/slides.pdf`、`preview/page-001.png…` | **真实**渲染；无渲染器则缺席并在 `render_report.status=skipped` 说明 |
| `/app/output/README.md` | 本次运行的产物清单、状态与缺口 |

## 3. 完整调用步骤

```bash
export SKILL=/skills/weekly-skill
python "$SKILL/scripts/run.py" prepare --input /input --work /app/work --output /app/output
# Agent 阅读 /app/work/REVIEW_PACKET.md、case_manifest.json、template_spec.json、
# source_inventory.json，按实际证据修改 /app/work/slides_plan.json。
python "$SKILL/scripts/run.py" finish  --work /app/work --output /app/output
```

预算紧张或希望零干预一次跑完时：

```bash
python "$SKILL/scripts/run.py" all --input /input --work /app/work --output /app/output
```

`all` 等价于 `prepare → 自动推导保守计划 → finish`；自动计划与手写计划走**完全相同**的保真/格式/渲染门禁。

### Windows 调用示例

PowerShell（先设变量，避免写死本机路径）：

```powershell
$env:SKILL = '<skill-install-dir>'     # 换成实际安装目录；不要在资源中写死本机绝对路径
python "$env:SKILL\scripts\run.py" prepare --input /input --work /app/work --output /app/output
python "$env:SKILL\scripts\run.py" finish  --work /app/work --output /app/output
python "$env:SKILL\scripts\run.py" all     --input /input --work /app/work --output /app/output
```

cmd.exe：

```bat
set "SKILL=<skill-install-dir>"
python "%SKILL%\scripts\run.py" prepare --input /input --work /app/work --output /app/output
python "%SKILL%\scripts\run.py" finish  --work /app/work --output /app/output
```

本地 Windows 上 `/input`、`/app/work`、`/app/output` 不存在时，请把它们替换成本机的真实输入/工作/输出目录；容器内运行时保持题面路径。

### 退出码

| 码 | 含义 |
|---|---|
| `0` | `result.json` 已写出且 `status ∈ {ok, partial}`（**仍需读报告**，退出码不等于质量通过） |
| `2` | 参数校验失败 / `/input/case.json` 不可读；同时写出诚实的 `needs_human` 结果 |
| `3` | `status=needs_human`：只有三类——`fidelity_report.ok=false`、关键格式项 fail、审计模块缺失（渲染器不可用**不**属此类，落 `partial` + exit `0`） |

### 模块函数签名（冻结，供接口核对）

只有 `run.py` 的 `prepare`/`finish`/`all` 三个子命令是冻结 CLI；下面是审计/汇总模块的 Python 签名，供你把它们当作库独立调用，**不要**给它们添加未冻结的命令行选项：

| 模块 | 签名 | 写出 |
|---|---|---|
| `verify_pptx` | `verify(work_dir, output_dir) -> dict` | `<work>/fidelity_report.json`、`<work>/format_audit.json` |
| `render_report` | `render(work_dir, output_dir) -> dict` | `<work>/render_report.json`（真实件入 `<output>/preview/`） |
| `finalize` | `build(work_dir, output_dir) -> dict` | `<output>/result.json`（+ `slides_plan.json`、`README.md`） |
| `finalize` | `validate(work_dir, output_dir) -> dict` | `<output>/evidence/validation_report.json`（含 `ok: bool`） |

正常流程应优先用 `prepare`/`finish` 入口，避免漏掉中间步骤。审计模块只读 `<work>`/`<output>` 中 engine 产出的 pptx，异常必须返回 `{'ok': False, ...}`，不得静默跳过。

## 4. 逐页复核步骤（别人怎么核对）

1. 先核对合同：`/app/output/result.json` 的 `identification.case_id` 是否等于输入 `case_id`；原 `output_contract` 键与嵌套形状是否**原样保留**；`delivery.audit.unfilled_fields` 里每一项是否都有说明与降级。
2. 打开 `/app/output/slides.pptx`，逐页对照 `slides_plan.json`：该页 `role/title/bullets/figures` 是否与计划一致；每条 `bullets[].evidence` 能否在 `source_inventory.units[].id` 找到；`figures[].cite` 是否指向真实源单元。
3. 逐页核对继承链：用 `format_audit.inheritance[].{slide,layout,master,theme,placeholder_idx}` 与 `result.template_inheritance.slide_layout_links` 检查每页是否真的挂在模板版式上；占位符 `idx` 是否在 `template_spec.layouts[]` 中存在。
4. 核对图与图号：`format_audit.aspect[]` 的 `aspect_delta` 应接近 0（等比）；正文图号与 `numbering`（默认 `preserve_source`）一致；`fidelity_report.checks` 中 `figure_numbering`、`caption_text_exact` 不为 `fail`。
5. 核对溢出与字号：`format_audit.overflow[]` 每项 `action` 应是 `font_ladder` 或 `split`；不应出现低于 `floor` 仍 `unresolved` 的项；`font_size_floor` 检查为 `pass`。字号下限按**形状角色**判定：优先 `<p:ph>` 的 `idx`/`type`，无 `<p:ph>` 时按固定 name（`Title` 2400 / `Content Placeholder` 1400 / `Substituted Body` 1400 / `Figure Caption` 1200，1/100 pt，见 [结果合同 §4.1](references/result_contract.md)）；**不得**用「文本以图号开头」这类文本模式判角色。
6. 核对真实预览：`render_report.status` 为 `rendered` 时，`preview/slides.pdf` 与 `preview/page-001.png…` **必须真实存在且非空**，页数与 pptx 页数一致（`consistency.ok=true`）；否则应为 `skipped`/`failed` 并写明原因，不得用其他文件冒充。
7. 核对保真与变更：`fidelity_report.ok=true` 且 `violations[]` 无未解释项；`change_records.jsonl` 的每条改动都能对上 `adaptation_actions.json` 的 `before/after/evidence`。
8. 最后跑一次合同自检：读 `/app/output/evidence/validation_report.json`，确认 `ok` 与未填字段，再看 `result.json.delivery.status` 是否与第 6 节的缺口描述一致。

任何一项 `fail`/`skipped`/`unchecked` 都**不得**当成通过；发现不一致时回到 `/app/work` 的原始报告定位，而不是改文字结论。

## 5. 资源上限自检

提交前在 Skill 目录上跑一次（只读，不修改任何文件）：

Linux / macOS：

```bash
python -c "import os,sys;r=sys.argv[1];fs=[os.path.join(d,f) for d,_,g in os.walk(r) for f in g];print('files=%d (limit 50)'%len(fs));print('max_bytes=%d (limit 500000)'%max(os.path.getsize(f) for f in fs));print('total_bytes=%d (limit 2000000)'%sum(os.path.getsize(f) for f in fs));print('symlinks=%d (limit 0)'%sum(1 for d,_,g in os.walk(r) for n in g if os.path.islink(os.path.join(d,n))))" /skills/weekly-skill
```

Windows PowerShell：

```powershell
python -c "import os,sys;r=sys.argv[1];fs=[os.path.join(d,f) for d,_,g in os.walk(r) for f in g];print('files=%d (limit 50)'%len(fs));print('max_bytes=%d (limit 500000)'%max(os.path.getsize(f) for f in fs));print('total_bytes=%d (limit 2000000)'%sum(os.path.getsize(f) for f in fs));print('symlinks=%d (limit 0)'%sum(1 for d,_,g in os.walk(r) for n in g if os.path.islink(os.path.join(d,n))))" "$env:SKILL"
```

要求：`files≤50`、`max_bytes≤500000`、`total_bytes≤2000000`、`symlinks=0`，且所有文件都是 UTF-8 文本。示例本身只做接口演示，不能替代真实案例的侦察产物。

## 6. 状态、失败与接管

| 情况 | 处理 |
|---|---|
| 全部关键产物 + 保真 + 格式 + 真实预览 + 合同就绪 | 状态才可为 `ok` |
| 有真实可用 deck，但渲染/字体/部分规格未验证 | `partial`，清楚列出缺口 |
| 保真失败（`fidelity_report.ok=false`）、关键格式项 fail、审计模块缺失 | `needs_human`，exit `3`；仍保留已有中间产物与原件 |
| 渲染器不可用（`render_report.status=skipped`）或渲染失败（`failed`） | `partial`，exit `0`；deck 可用，预览缺席并如实记录，**不**借用其他文件冒充预览 |
| 请求正文改写、跨模板合并、在线下载模板、宏 deck 编辑 | 只读保留并登记 `unsupported`/不确定性；见 [限制与降级](references/limitations.md) |

## 7. 版本与维护

- Skill 名称：`academic-ppt-template-adaptation`；版本：`1.0.0`（与 `scripts/common.py` 的 `SKILL_NAME`/`SKILL_VERSION` 一致，`sources.json` 会记录脚本 SHA）。
- 接口以 `CONTRACT.md` 为唯一冻结源。README、`references/`、`examples/` 只解释冻结接口，不引入隐式功能；接口变化时必须同步更新命令、字段、路径与失败处置。
- 维护原则：缺证据时优先诚实降级，而不是把说明写得"看起来成功"。平台上的正式案例、对照白板和评分均由组织方提供或生成；本目录的合成示例只用于理解接口，不能预填正式答案，也不能用来声称保真、视觉通过、增益、分数或第三方认证。
