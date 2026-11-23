# 支持范围、明确不支持项与降级处置

本文件说明本 Skill **能做什么**、**明确不做什么**，以及每一类失败该如何**诚实处置**。所有降级都必须落在报告字段里（`render_report.status`、`format_audit.status`、`fidelity_report.ok`、`result.json.delivery.audit.unfilled_fields`、`open_items`），而不是只写在说明文字里。

## 1. 支持范围

### 1.1 输入

| 类别 | 支持 | 说明 |
|---|---|---|
| 模板 | `.pptx`、`.potx` | 必须能作为 ZIP/OOXML 解析出 `presentation.xml` 与 master/layout/theme 部件；`template_spec.json` 只登记**真实字节**中读到的值。 |
| 源（正文主线） | `.pdf`、`.docx`、`.md`、`.txt` | 提取结果写进 `source_inventory.json`；PDF 收到的是**真实 PDF**，不假定存在预抽取答案层。 |
| 源（其他） | `.tex`、`.csv`/`.tsv`、`.xlsx`、`.bib`、`.json`、代码/笔记本 | 能解析出结构时作为补充证据；不能解析就如实记录 `extraction.status=failed`。 |
| 图 | PNG/JPEG/GIF/BMP/SVG（读真实头部尺寸） | 等比放入容器；EMF/WMF/TIFF/WebP 视渲染环境而定，读不到尺寸时按不可验证处理。 |
| 起始 deck | `.pptx` | 任务要求在已有 deck 上编辑时，作为 `start_deck` 处理；仍须验证模板继承。 |
| 输出 | 可编辑 `.pptx` + 讲稿备注 + 真实 PDF/PNG 预览 | 预览**仅在有可用渲染器时**存在。 |

### 1.2 能力

- 识别真实 `slide_size`、`masters[]`、`layouts[].name/type/placeholders[]`、`theme[].fonts/colors`、`notes_master`，并保留继承链。
- 按真实版式名与占位符 `idx` 填内容；占位符缺失时做**有记录的等价替代**。
- 沿 `font_ladder` 下调字号到下限；仍溢出则**拆页**（不压到不可读）。
- 图表按容器**等比**缩放/裁切，记录 `scale` 与 `aspect_delta`。
- 默认保留源图号（`numbering.mode="preserve_source"`），只做映射不改写。
- 生成讲稿备注与逐页 `claims`/`evidence`/`locator`。
- 独立审计：保真、格式继承、溢出、图表、批量一致性；真实渲染与字体证据。

## 2. 明确不支持（不要尝试，直接按 §3 降级）

| 不支持项 | 原因 | 处置 |
|---|---|---|
| **legacy `.ppt`** | 二进制旧格式，无法用标准库安全改写 | 保留原件，只读登记；`needs_human`。 |
| **宏 `.pptm` / 模板宏** | 改写可能破坏 VBA 部件与签名 | 保留原件，不编辑；登记 `unsupported` + `needs_human`。 |
| **跨模板合并** | 多 master/版式冲突无法用单一继承链表达 | 只按其中一个真实模板交付，其余登记 `limitations`；必要时 `needs_human`。 |
| **无源文件的凭空生成** | 会编造科学内容 | 拒绝生成正文；只产出诚实中间产物与 `needs_human`。 |
| **正文改写 / 摘要重写** | 改变事实 | 不改写；`fidelity_report.ok=false` 视为硬失败。 |
| **在线下载模板** | 首次运行不得联网 | 不下载；缺模板时如实降级。 |
| **编造数据 / 补造图表数值** | 科学诚信 | 一律禁止；缺数据就记缺口。 |
| **PDF/LaTeX 反向排版** | 超出套版范围 | 只读保留并登记 `unsupported`。 |
| **跨向视觉"认证"** | 本 Skill 不做人工审美判定 | 只提供可复核证据；不声称视觉通过或评分。 |

## 3. 逐类异常与降级处置

### 3.1 模板缺版式

- 版式**只能**取 `template_spec.layouts[]` 中真实存在的 `name`。需要的版式不存在时退到语义最近的真实版式（`title` / `blank` / 通用内容页），并在 `adaptation_actions.json` 记 `placeholder_substitute`、在 plan `notes` 说明。
- 该版式无法承载必需内容（例如需要图表版式但模板只有纯文字版式）时，如实记 `unresolved`，`status` 降级。

### 3.2 占位符缺失

- 不在 `template_spec.layouts[].placeholders[]` 中的 `idx`/`ph_type` **不得凭空写入**。
- 优先用同版式其他真实占位符；否则用该版式的文本框做等价替代，**必须**按 v1.2 §4.2 给固定 `p:cNvPr@name`（`Title` / `Content Placeholder` / `Substituted Body` / `Figure Caption`），并记录 `before/after/evidence` 与限制。角色判定优先 `<p:ph>` 的 `idx/type`，没有 `<p:ph>` 才看 name；**禁止**用「文本以图号开头」这类文本模式判角色。
- 结构无法安全承载时（如标题被母版固定为不可覆盖），如实记入 `open_items` 并把缺口写明，**不得**伪造完整的 `ok`/`partial` 交付。

### 3.3 文本溢出

- 顺序固定：**沿 `font_ladder` 降到 `floor` → 仍溢出则拆页**。
- `floor` 不得低于 `common.FONT_FLOOR`（title 2400 / body 1400 / caption 1200 / table 1200 / notes 1000，1/100 pt）。
- 每种处置都写进 `format_audit.overflow[]`，`action ∈ {font_ladder, split, unresolved}`；**不允许**低于下限仍然硬塞，也不允许把 `unresolved` 写成 `font_ladder`。
- 拆页产生的续页用 `continuation_of` 指向原页，并在 plan `notes` 说明"拆页而非删字"。

### 3.4 字体缺失

- 主题字体只在 `theme[].fonts` 有真实证据时保留/设置；不假设本机装了哪套字体。
- 渲染后发生字体替换时，把 `requested` / `used` / `evidence` 写进 `render_report.fonts.substitutions[]`。
- **不**因字体缺失改动版面结构、不假装字体已嵌入；必要时在 `open_items.limitations` 说明印刷风险。

### 3.5 图表源失效

- `figures[].asset` 不存在或不可解析时：保留 `source`、`caption`、`cite`，`asset=null`，在 `figures[]`/`open_items` 说明缺图。
- **不得**用占位图、截图或编造图形冒充原图；不得因缺图改写图号或题注。
- 图仍可等比放置但尺寸无法核实（读不到真实宽高）时，记 `aspect` 为不可验证，`fidelity_report` 相应检查记 `skipped` 而非 `pass`。

### 3.6 渲染失败 / 无渲染器

- 探测不到渲染器：`render_report.status="skipped"`，`preview/` 缺席，`attempts[]` 写明探测过程；**不得伪造预览**；按 v1.2 该情形降为 **`partial` + exit `0`**（交付物可用、只缺预览）。
- 渲染器存在但失败：`status="failed"`，`attempts[]` 记渲染器与错误；可编辑 deck 照常交付，`status` 降为 `partial`。
- 渲染成功但页数不一致：`consistency.ok=false`，如实记录，不删页凑数。

### 3.7 源非可解析格式 / 提取失败

- `source_inventory.extraction.status="failed"`、`text_available=false` 时：保留原始材料与可核实的定位信息，只做有证据支持的最小交付，把缺口写进 `open_items.limitations`。
- **不得**依据文件名、领域标签或"看起来像什么"补造正文、数字或图号。
- PDF 走真实解析路径；解析工具不可用时如实记 `failed`，不编造 token 覆盖指标。

### 3.8 源与模板角色不明

- 候选不唯一或角色冲突时保持 `confirmed_roles=false`，`plan_resolution.plan_source` 可能为 `unusable`；如实记入 `open_items` 并给出诚实的中间产物，**不得**伪造完整的 `ok`/`partial` 交付。
- 都不得用文件名赌博；把缺哪个事实写进 `open_items.uncertainty`。

### 3.9 预算耗尽

- 64 请求 / 900 秒内**先产出诚实 partial**：至少保留 `/app/work` 侦察产物与已有 `/app/output` 中间件。
- 若只完成侦察，写 `pipeline_report.json` 与 `needs_human`/`partial` 结果，说明下一步；不得为赶进度跳过审计或伪造 `ok`。
- 预算紧张时用 `run.py all` 一次性路径，auto 计划走同一套门禁。

## 4. 诚实性红线（违反即视为造假）

1. 登记不存在的文件/哈希/页数/字体/渲染结果。
2. 把 `skipped`/`unchecked`/`unsupported` 写成 `pass`。
3. 用占位图冒充原图，或编造图表数据。
4. 改写正文事实却声明保真通过。
5. 把示例、白板或历史运行结果说成本次正式案例结果或分数。
6. 把退出码 0 当作质量通过。

## 5. 相关文件

- 计划字段与证据优先级：[plan_guide.md](plan_guide.md)
- 交付字段、状态与证据真实性：[result_contract.md](result_contract.md)
- 示例（`example=true`，不是答案）：[../examples/README.md](../examples/README.md)
