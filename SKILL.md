---
name: academic-ppt-template-adaptation
description: Adapt a fixed paper or research-material pack into a real organizer-provided academic PPTX template while preserving every scientific content unit. Use when the task is to read /input/case.json and its real materials (template.pptx, paper.pdf/docx/md/txt, figure files, optional start deck) and produce an editable slides.pptx that inherits the template master, layout, theme and placeholders, plus speaker notes, a page-level evidence map, change records and real rendered PDF/PNG previews under /app/output with result.json. It reads the true OOXML master/layout/theme/placeholder structure, fills placeholders in the correct layout, keeps source figure numbering by default, scales images proportionally, holds a readable font floor (splits a slide instead of squeezing), and reports honest ok/partial/needs_human status. Do not use it for free-form slide design with no source material or no template, for rewriting or summarizing the science beyond what the sources state, for editing legacy .ppt or macro .pptm decks, for merging several templates, for downloading templates at runtime, or for producing prepared answers, whiteboards or scores for the fixed cases.
---

# 学术 PPT 模板与格式适配（真实母版/版式/主题继承 + 页级证据映射）

把**实际输入**中的论文或研究材料，套进**组织方提供的真实 pptx 模板**，交付仍可编辑的演示文件、讲稿备注、逐页证据映射和真实渲染预览。不是写一份排版建议，也不是把论文整段粘进默认母版。

**一句话原则：先保真，再谈外观。** 科学文字、数字、单位、引用、图号与图表指向都不因排版而变；模板的母版/版式/主题继承必须真实发生且可被他人查证。

## 1. 目标与硬边界

- 输入（只读）：`/input/case.json`、`/input/task.md` 与实际材料文件。角色由 case、任务说明和实际内容共同确定，**不能只按文件名猜**。
- 中间产物：`/app/work`（侦察、计划、审计报告）。
- 交付区：`/app/output`，必需 `result.json`，默认可编辑 `slides.pptx`，真实预览在 `preview/`。
- 安装位置：Skill 装在 `/skills/weekly-skill`；**不得修改** `/input` 与 `/skills/weekly-skill`，首次运行**不得联网**。
- 预算：64 model requests / 900 秒 / 单响应 16384 token。**先把诚实 partial 落盘，再逐步改善**。
- 资源：Skill 本体是可安装文本资源；**不得把二进制 pptx/PDF/PNG 打包进 Skill 本体**，不得预制答案。

## 2. 必须遵守的铁律

1. **先保真再外观。** 不改写正文事实，不删/并/重建图、表、公式、引用与图号。`fidelity_report.json` 的 `ok=false` 是**硬失败**：不得声明 `ok` 或"完成"。
2. **模板继承必须真实且可查。** 交付 pptx 必须含模板的 master / slideLayout / theme 部件，`slide → layout → master → theme` 链接解析成功、占位符 `idx` 真实存在；证据写进 `format_audit.json` 的 `inheritance[]` 与 `result.json` 的 `template_inheritance`。只把文字粘进默认母版＝不合格。
3. **绝不把字号压到不可读来硬塞。** 只允许沿 `font_ladder` 降到下限 `floor`（`common.FONT_FLOOR`：title 2400 / body 1400 / caption 1200 / table 1200 / notes 1000，单位 1/100 pt）。到下限仍溢出就**拆页**（`text_split`），并把 `format_audit.overflow[].action` 记为 `split`；不得出现低于下限仍 `unresolved` 的硬塞。
4. **图表等比不拉伸。** `figures[].aspect_lock=true` 为默认。按容器与原图真实宽高比缩放（contain），必要时裁切；`format_audit.aspect[]` 记录 `scale` 与 `aspect_delta`。非等比拉伸＝格式失败。
5. **图号默认保留源编号。** `numbering.mode="preserve_source"`，`map` 只做 `F1=图1` 一类**映射**，不改写源编号。只有任务显式要求时才用 `relabel`，并在 notes 说明依据。
6. **无渲染器就诚实 `skipped`。** 绝不伪造 PDF/PNG；只有**真实存在且非空**的文件才能登记 `sha256`/`bytes`。任一空文件或借用其他文件的预览都算伪造。
7. **不得预制答案。** Skill 本体只含流程、规则与接口示例；`examples/**` 全部标 `example=true` 与「示例 / EXAMPLE」，**不是**正式案例、不是答案、不代表评分或任何实测结果。
8. **预算内先产出诚实 partial。** 时间/请求紧张时用 `run.py all` 一次性走完全流程（auto 计划走同一套门禁）；宁可交付带明确缺口的 `partial`，也不要为"全绿"编造证据。
9. **报告字段与命令必须与冻结接口一致。** 不新增未冻结选项、不虚构文件名或状态、不编造 hash/字体/渲染结果。
10. **跳过不等于通过。** `skipped`/`unchecked`/`unsupported` 一律不得计作 `pass`；未检查项必须出现在报告里。

## 3. 标准运行（两命令 + 一次性）

```bash
SKILL=/skills/weekly-skill
# 1) 侦察：只读 /input，写 /app/work，不排版
python "$SKILL/scripts/run.py" prepare --input /input --work /app/work --output /app/output
# 2) Agent 读 REVIEW_PACKET.md 等，按需改 /app/work/slides_plan.json
# 3) 套版 + 审计 + 渲染 + 汇总，写 /app/output
python "$SKILL/scripts/run.py" finish  --work /app/work --output /app/output
# 0 干预一次性路径（预算紧张时）：等价 prepare → auto-derive plan → finish
python "$SKILL/scripts/run.py" all     --input /input --work /app/work --output /app/output
```

- 三个子命令**当且仅当** `/app/output/result.json` 已写出且 `status ∈ {ok, partial}` 时 exit `0`；`needs_human` 时 exit `3`（仍必须写出诚实的 `result.json` 与已有产物）；参数校验失败或 `/input/case.json` 不可读时 exit `2` 并写 `needs_human` 结果。
- `prepare` 只侦察、复制安全副本、写诚实中间产物，**不做排版**。即使后续未完成，也不要删除这份中间产物。
- `finish` 固定顺序：套版 → `verify_pptx.verify` → `render_report.render` → `finalize.build` → `finalize.validate`；`run.py` 负责把 `<work>/*.json` 复制到 `<output>/evidence/`。
- 缺任一审计模块时明确报 `auditor module missing` → `status=needs_human` + exit `3`；**不得自实现、不得静默跳过**。
- **状态语义（v1.2）**：只有三类算 `needs_human`（exit `3`）——① `fidelity_report.ok=false`；② 关键格式项 fail；③ 审计模块缺失。**渲染器不可用只降级为 `partial` + exit `0`**（交付物可用、只是缺预览），但 `render_report.json` 的 `status` 仍必须如实写 `skipped`/`failed`。`prepare` 单独调用时只要侦察产物齐备即 exit `0`。

### 模块函数签名（冻结，供接口核对，不是要求 Agent 改代码）

```python
verify_pptx.verify(work_dir, output_dir) -> dict      # 写 <work>/fidelity_report.json + format_audit.json
render_report.render(work_dir, output_dir) -> dict     # 写 <work>/render_report.json；真实件入 <output>/preview/
finalize.build(work_dir, output_dir) -> dict           # 写 <output>/result.json（+ slides_plan.json、README.md）
finalize.validate(work_dir, output_dir) -> dict         # 写 <output>/evidence/validation_report.json，含 ok: bool
```

审计模块只读 `<work>`/`<output>` 中 engine 产出的 pptx；异常或缺文件必须返回 `{'ok': False, ...}` 并写明原因。

## 4. `prepare` 之后实际要读的文件

按顺序读，别只看一个摘要就假定全文：

| `/app/work` 文件 | 你要确认什么 |
|---|---|
| `REVIEW_PACKET.md` | 文件角色、真实要求摘要、`template_spec` 要点（版式名/占位符）、冲突与不确定入口 |
| `case_manifest.json` | `case_id`、原 `output_contract`（原样保留）、`required_fields`、每个输入的 `rel/kind/size/sha256`、`budget` |
| `template_spec.json` | 真实 `slide_size`（`src` 指向实际 XPath）、`masters[]`、`layouts[].name/placeholders[]`、`theme[].fonts/colors`、`evidence[]` |
| `source_inventory.json` | 源 `kind/sha256`、`sections[]`、`units[]`（含 `locator`）、`figures[]`（`label/caption/asset/width/height`）、`tables[]`、`extraction.status` |
| `slides_plan.json` | 待你按证据编辑的唯一计划；自动生成的只是保守壳，不是答案 |
| `plan_resolution.json` | 本次计划来自 `agent` / `auto_derived` / `unusable` 及理由 |

长文用检索定位（`source_inventory.units[].locator`），不必把全文塞进上下文；**绝不能只凭文件名或错误领域标签补造内容**。

## 5. 计划关键字段（完整字段见 `references/plan_guide.md`）

- `confirmed_roles`：只有核查材料后才能置 `true`；角色冲突或候选不唯一时保持 `false` 并降级。
- 每页 `layout_name` + `layout_part`：必须来自真实 `template_spec.layouts[]`；只按真实版式名匹配，不猜。
- 形状命名（v1.2 §4.2）：无 `<p:ph>` 的替代文本框**必须**用固定 `p:cNvPr@name`——`Title`（title，下限 2400）/ `Content Placeholder`（body 有占位符，1400）/ `Substituted Body`（body，占位符缺失时的显式文本框，1400）/ `Figure Caption`（caption，1200）；角色判定**优先** `<p:ph>` 的 `idx/type`，没有 `<p:ph>` 才看 name，**禁止**用"文本以图号开头"这类文本模式判角色。
- `bullets[].evidence`：每条要点指向 `source_inventory` 的 `units[].id`（`u0012`）或 `figures[].id`。
- `figures[].aspect_lock`：默认 `true`；`source` 指向 `F1`，`asset` 是真实相对路径，`cite` 指向源单元。
- `notes` / `notes_evidence`：讲稿备注与其来源，保证备注不是凭空生成。
- `font_ladder` / `numbering` / `contract_values`：字号阶梯（含 `floor`）、图号策略（默认 `preserve_source`）、原合同 dotted path → 真实字段的映射。
- `continuation_of`：拆页后第 2 页指向被拆的原页号，便于复核"拆页而非删字"。

## 6. 验收顺序与状态表（详见 `references/result_contract.md`）

`finish` 不会只依赖上游自报成功。完成后**必须**逐份阅读：

| 文件 | 你要确认什么 |
|---|---|
| `adaptation_actions.json` | 每个改动有 `before/after/evidence`；没有"文本改写"伪装成排版 |
| `fidelity_report.json` | `ok=true`；`metrics.unsourced_deck_strings` 与 `violations[]` 无未解释项；无改写、丢字、图号错位 |
| `format_audit.json` | `status` 与每项 `status` 有实测 `expected/observed/evidence`；未检查项没被算成 PASS |
| `render_report.json` | `status` 真实；登记进 `artifacts`/`pages` 的文件确实存在且非空；字体证据来自真实工具 |
| `result.json` | 保留原 `output_contract`、artifacts/hash 指向最终文件、`status` 与缺口没被覆盖 |
| `evidence/validation_report.json` | `finalize.validate` 的合同自检结果（`ok` 及未填字段） |

| 状态 | 判据（需同时满足） | 退出码 | 优先读 |
|---|---|---|---|
| **ok** | `result.json` 写出；审计模块均返回 ok；`fidelity_report.ok=true`；`format_audit.status=pass`；`render_report.status=rendered` 且有真实非空 PDF/PNG；pptx 含模板 master/layout/theme；`delivery.audit.unfilled_fields` 为空 | 0 | `result.json` + `fidelity/format/render` + `sources.json` |
| **partial** | 有真实可用 deck 与真实证据；`fidelity_report.ok=true` 且无关键格式项 fail；但存在**非阻断**缺口：`format_audit.status=partial`（有 `skipped`/`unchecked`）、`render_report.status=skipped`（无渲染器）或 `failed`（渲染器存在但失败）、`unfilled_fields` 非空、或目标页数未达但已说明 | 0 | 同上 + `pipeline_report.json` |
| **needs_human** | 三类（与 `CONTRACT.md` §4 v1.2 逐字对应）：① `fidelity_report.ok=false`；② 关键格式项 fail；③ 审计模块缺失（`auditor module missing`） | 3 | `result.json` + 已产出的 work json；缺口写进 `open_items.uncertainty/limitations` |

> `needs_human` 的编码细节：`/input/case.json` 不可读或参数校验失败时 exit `2`，但**仍必须写出 `needs_human` 结果**；其余三类 exit `3`。角色/模板无法确认、无法产出可用 deck 的极端情形不属于上表三类，也不得据此伪造完整的 `ok`/`partial` 交付——仍要写出诚实的 `result.json` 与已有中间产物，并把缺口记入 `open_items`。

> **诚实降级（v1.2）**：`fidelity_report.ok=false`、关键格式项 fail、审计模块缺失 → 降级并 exit `3`（落 `needs_human`）；**渲染 `status=skipped`/`failed` 只降为 `partial` + exit `0`**，但状态字段必须如实写 `skipped`/`failed`。`skipped` 不是 `pass`。若 engine 的实现与本节某一格不同，**以 `CONTRACT.md` 为准**并报告，不要靠文字把状态写成通过。

失败后不要用越来越宽的规则"修绿"。回到安全副本与 `adaptation_actions.json`，保留失败证据；可编辑 pptx 仍要留下。

## 7. 异常与降级（逐条）

| 情况 | 诚实处置 |
|---|---|
| **模板缺版式** | 只用 `template_spec.layouts[]` 里真实存在的版式；需要的版式不存在时退到语义最近的真实版式（如 `title`/`blank`）并记 `placeholder_substitute` + `notes`；无法承载时如实记入 `open_items` 并降级，**不得**伪造完整的 `ok`/`partial` 交付。 |
| **占位符缺失** | 不在 `template_spec` 中的 `idx`/`ph_type` 不得凭空写。优先用同版式其他真实占位符；否则用该版式的文本框等价替代，并**必须**按 v1.2 §4.2 用固定 `p:cNvPr@name`（`Title` / `Content Placeholder` / `Substituted Body` / `Figure Caption`），审计据此判角色与字号下限；结构不可安全承载时降级并记录。 |
| **文本溢出** | 先按 `font_ladder` 逐级降到 `floor`；仍溢出就 `text_split` 拆页（`continuation_of` 指向原页）；两种情况都写 `format_audit.overflow[]`。**不允许低于下限硬塞**；无法拆解时如实 `unresolved` 并降级。 |
| **字体缺失** | 主题字体只在 `theme[].fonts` 有真实证据时保留/设置；渲染后字体被替换时，把 `substitutions[]` 写进 `render_report.fonts`（`pdffonts`/`soffice-log` 证据）；不假装字体已嵌入，也不因缺失改版面结构。 |
| **图表源失效** | `figures[].asset` 不存在或不可解析时，保留 `source` 与 `caption`/`cite`，`asset=null`，在 `figures[]`/`open_items` 说明缺图；**不得**用占位图冒充原图，不得编造数据。 |
| **渲染失败 / 无渲染器** | 探测不到渲染器：`render_report.status="skipped"`，`preview/` 缺席，`attempts[]` 写明探测过程，**不得伪造预览**；按 v1.2 该情形降为 **`partial` + exit `0`**（交付物可用、只缺预览）。渲染器存在但失败：`status="failed"` 并记渲染器与错误，同样为 `partial`。 |
| **源非可解析格式** | `source_inventory.extraction.status=failed` 时保留原始材料与定位信息，`text_available=false`，只做有证据支持的最小交付并记 `limitations`；不得依据文件名臆造正文。 |
| **源是 legacy `.ppt` / 宏 `.pptm`** | 保留原件，不尝试编辑；按 `limitations.md` 的未支持路径处理并降级。 |

## 8. 可复用材料

- `README.md`：安装、Linux/Windows 独立调用、输入输出清单、逐页复核步骤、资源上限自检。
- `references/plan_guide.md`：`slides_plan.json` 全字段、证据优先级、冲突处置、占位符缺失时的固定形状命名（v1.2 §4.2）、auto-derived 兜底规则。
- `references/result_contract.md`：`result.json` 全字段、原 `output_contract` 与 `contract_path_map`、`unfilled_fields`、状态语义、形状命名与角色判定、真实性要求。
- `references/limitations.md`：支持范围、明确不支持项与每类诚实处置。
- `examples/`：标 `example=true` / EXAMPLE 的合成最小接口示例，**不是**正式案例、不是答案、不代表评分。

平台案例的事实优先级**高于**示例；严禁把示例中的路径、版式名或数值套到真实材料上。
