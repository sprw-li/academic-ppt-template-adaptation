# 学术 PPT 模板与格式适配 Skill

把**固定论文/材料**套进**组织方提供的真实 pptx 模板**：识别真实的母版、版式、主题与占位符，写入正确的版式，保留图号与图表指向，等比缩放图表，守住可读字号下限，生成讲稿备注，并独立检查内容保真、格式继承和真实渲染。题面要求"actual starting deck where edit/adapt requires one"与"fixed audience/time/slide count"时，本 Skill 会走 `mode=adapt_start_deck` 保留起始 deck 的**全部文本**（可重排/降号/拆页），逐项遵从 `case.json.constraints`，把计划里的表格渲染成真正可编辑的 `<a:tbl>`，并写出可朗读、有来源的讲稿。交付可编辑 `slides.pptx` + 页级证据映射 + 变更记录 + 真实 PDF/PNG 预览 + `result.json`。

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

### 有起始 deck 时：走 `adapt_start_deck`

题面要求"在已有 deck 上编辑/适配"时，**不要**照 `from_template` 从零重建——那样会丢掉起始 deck 的内容。流程一样，只是 `prepare` 会把模式写进计划：

```bash
python "$SKILL/scripts/run.py" prepare --input /input --work /app/work --output /app/output
# 读 /app/work/start_deck_spec.json（v1.3）确认起始 deck 真实可解析：
#   ok=true / slide_count>0 / slide_details[].text_units[] 与 pictures[]
# 再读 /app/work/slides_plan.json，确认 mode=adapt_start_deck、start_deck_path 指向该 deck，
# 并把每页的 source_slide / preserve_start_text=true 与 constraints_applied 补全。
python "$SKILL/scripts/run.py" finish --work /app/work --output /app/output
```

- **触发条件**（三者同时成立）：`case_manifest.json` 的 `detected.start_deck` 唯一命中、`start_deck_spec.json.ok=true`、`slide_count>0`，且任务说明要求适配已有 deck。任一不成立就退回 `mode=from_template`（`start_deck_path=null`），并在 plan `notes` 说明为什么没用起始 deck。
- **要覆盖自动判定**：只改 `/app/work/slides_plan.json` 里的 `mode` 与 `start_deck_path`，不要加命令行选项；证据冲突时把两条来源都写进 `notes` 与 `open_items.uncertainty`，审计侧 `start_deck_preserved` 记 `skipped`。
- **内容规则**：起始 deck 每页文本必须保留（可重排/降号/拆页），起始 deck 里有而计划未列的页默认追加保留；确需删页必须记 `start_deck_slide_dropped` + 理由 + 该页文本 hash。
- **约束规则**：`case.json.constraints` 的 `audience`/`time_minutes`/`target_slide_count`/`font_size_floor_pt`/`caption_numbering` 逐项落到计划与 deck，并在 `constraints_applied` 记 `{value, source}`；`target_slide_count` 是硬上限。
- **表格规则**：计划 `tables[]` 渲染成真正的 `<a:tbl>`，不许截图或纯文本替代。
- **讲稿规则**：每页 `notes` 写成可朗读句子，`notes_evidence` 非空且内容 token 可回溯到源。

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
9. **（v1.3）核验起始 deck 内容没丢**：先看 `format_audit.checks[]` 里 `start_deck_preserved` 的状态与 `evidence[]`，再把 `start_deck_spec.json.slide_details[].text_units[].text` 逐条与 `/app/output/slides.pptx` 的正文对照；允许页序变化、字号变小、拆页（标题带 "(续)"），**不允许**整句/整段消失。确需删页的，`adaptation_actions.json` 必须有 `start_deck_slide_dropped` + 理由 + 该页文本 hash；计划外的起 deck 页应为 `start_deck_slide_preserved`。可用下面的命令把起始 deck 每页文本打印出来逐条核对：

```bash
python -c "import json,sys,io;d=json.load(io.open(sys.argv[1],encoding='utf-8'));[print(s['index'],s.get('title'),[u['text'] for u in s.get('text_units',[])]) for s in d['slide_details']]" /app/work/start_deck_spec.json
```

10. **（v1.3）核验表格是真 `<a:tbl>` 而不是图片/纯文本**：`format_audit.checks[]` 的 `table_rendered` 应为 `pass`（或计划确实无表格时 `skipped` 且进了 `unchecked`）；同时直接用下面命令数一遍输出 deck 里的 `<a:tbl>` 与 `<a:tblGrid>`，数量应等于 `slides_plan.json` 的 `tables[]` 条数（拆表时按拆出的块数计，并能在 `adaptation_actions.json` 找到 `table_split`）：

```bash
python -c "import re,sys,zipfile;z=zipfile.ZipFile(sys.argv[1]);parts=[n for n in z.namelist() if re.match(r'ppt/slides/slide[0-9]+\.xml$',n)];data={n:z.read(n).decode('utf-8') for n in parts};print('slide_parts=%d a_tbl=%d a_tblGrid=%d'%(len(parts),sum(v.count('<a:tbl>') for v in data.values()),sum(v.count('<a:tblGrid>') for v in data.values())))" /app/output/slides.pptx
```

11. **（v1.3）核验约束与讲稿**：`result.json.plan.constraints_applied` 五项与 `case.json.constraints` 对上（`audience`/`time_minutes` 不得是占位串，`slide_count <= target_slide_count`）；`format_audit` 的 `constraints_honoured` 为 `pass`；`fidelity_report.notes_evidence[]` 每页 `unsourced_tokens` 为空，且 `notes` 打开就在 deck 的备注页里（不是只写在计划里）。

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
| v1.3 四项新检查任一项 `fail`（`start_deck_preserved` / `table_rendered` / `constraints_honoured` / `notes_evidence_present`） | 按**关键格式项 fail** 处理：`needs_human`，exit `3`；缺口与冲突写进 `open_items`，不要靠改文字修绿 |
| v1.3 四项新检查只出现 `skipped`（无起始 deck / 无表格 / 案例无约束 / 案例不要求讲稿） | 降为 `partial`，exit `0`；`skipped` 项必须进 `format_audit.unchecked[]` 或该项 `detail`，**不是** `pass` |
| 起始 deck 内容与源材料冲突 | 以源材料为事实基准；冲突原文与出处写进 `open_items.uncertainty`，起始 deck 文本仍按 `start_deck_preserved` 规则保留 |
| 渲染器不可用（`render_report.status=skipped`）或渲染失败（`failed`） | `partial`，exit `0`；deck 可用，预览缺席并如实记录，**不**借用其他文件冒充预览 |
| 请求正文改写、跨模板合并、在线下载模板、宏 deck 编辑 | 只读保留并登记 `unsupported`/不确定性；见 [限制与降级](references/limitations.md) |

## 7. 版本与维护

- Skill 名称：`academic-ppt-template-adaptation`；版本：`1.0.0`（与 `scripts/common.py` 的 `SKILL_NAME`/`SKILL_VERSION` 一致，`sources.json` 会记录脚本 SHA）。
- 接口以 `CONTRACT.md` 为唯一冻结源。README、`references/`、`examples/` 只解释冻结接口，不引入隐式功能；接口变化时必须同步更新命令、字段、路径与失败处置。
- 维护原则：缺证据时优先诚实降级，而不是把说明写得"看起来成功"。平台上的正式案例、对照白板和评分均由组织方提供或生成；本目录的合成示例只用于理解接口，不能预填正式答案，也不能用来声称保真、视觉通过、增益、分数或第三方认证。
