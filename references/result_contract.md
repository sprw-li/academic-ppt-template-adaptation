# 交付合同：`result.json`、状态与证据真实性

本文件解释 `/app/output` 的交付合同：`result.json` 全字段、原 `output_contract` 的保留与映射、`unfilled_fields` 降级规则、`ok`/`partial`/`needs_human` 状态语义，以及"哪些证据必须真实"。字段以 `CONTRACT.md` §6 / §6.1 为准。

## 1. `/app/output` 产物清单

| 文件 | 要求 |
|---|---|
| `result.json` | **必需**。顶层 `case_id`；原 `output_contract` 原样保留；结构见 §2 |
| `<deck_name>`（默认 `slides.pptx`） | 可编辑 pptx；**必须**含模板的 master / slideLayout / theme 部件 |
| `slides_plan.json` | 页级证据映射的对外副本（每页 `claims` + `evidence` + `locator`） |
| `evidence/fidelity_report.json` | 内容保真审计副本 |
| `evidence/format_audit.json` | 格式/继承/溢出/图表审计副本 |
| `evidence/render_report.json` | 渲染与字体证据副本 |
| `evidence/change_records.jsonl` | 逐条变更记录副本 |
| `evidence/sources.json` | 来源、哈希、工具链、复跑命令副本 |
| `evidence/validation_report.json` | `finalize.validate` 的合同自检结果（`ok` 与未填字段） |
| `preview/<deck>.pdf`、`preview/page-001.png…` | **真实**渲染；无渲染器则缺席并在 `render_report.status=skipped` 说明 |
| `README.md` | 本次运行的产物清单、状态、缺口 |

`<output>/evidence/*` 的副本由 `run.py` 统一从 `<work>` 复制；审计模块不负责 evidence 目录布局。

## 2. `result.json` 全字段

```json
{"identification":{"case_id":"...","task_type":"academic_ppt_template_adaptation",
                   "skill":"academic-ppt-template-adaptation","skill_version":"1.0.0"},
 "delivery":{"status":"ok|partial|needs_human","primary_output":"slides.pptx","artifacts":[
    {"rel":"slides.pptx","sha256":"...","bytes":123,"kind":"pptx","editable":true},
    {"rel":"preview/slides.pdf","sha256":"...","bytes":123,"kind":"pdf","rendered":true}],
   "audit":{"fidelity_ok":true,"format_status":"pass","render_status":"rendered","unfilled_fields":[]}},
 "evidence":{"sources":{...},"change_records":{...},"fidelity":{...},"format_audit":{...},"render":{...}},
 "slides":[{"n":1,"role":"title","title":"...","claims":["..."],
            "evidence":[{"kind":"section|figure|table|unit","ref":"S1|F1|u0044","locator":"p.3 §2.1","cite":"..."}]}],
 "plan":{"template":"rel","source":"rel","slide_count":16,"layout_usage":{"标题和内容":9,"空白":2}},
 "template_inheritance":{"masters":["..."],"layouts":["..."],"theme":"ppt/theme/theme1.xml",
                         "slide_layout_links":[{"slide":1,"layout":"ppt/slideLayouts/slideLayout1.xml"}]},
 "open_items":{"uncertainty":["..."],"limitations":["..."]},
 "runtime":{"prepare_s":0.0,"finish_s":0.0,"python":"3.x","renderer":"...",
            "budget":{"model_requests":64,"seconds":900}},
 "output_contract":{... 原样保留 ...}}
```

| 顶层键 | 语义与要求 |
|---|---|
| `identification` | `case_id` **必须**等于输入 `case_id`（平台硬要求）；`task_type` 固定 `academic_ppt_template_adaptation`；`skill`/`skill_version` 与 `scripts/common.py` 的 `SKILL_NAME`/`SKILL_VERSION` 一致。 |
| `delivery.status` | `ok` / `partial` / `needs_human`，判据见 §4。 |
| `delivery.primary_output` | 主产物相对路径，默认可编辑 deck。 |
| `delivery.artifacts[]` | 每个产物的 `rel`、真实 `sha256`、真实 `bytes`、`kind`；可编辑件带 `editable:true`，渲染件带 `rendered:true`。**只有真实存在的非空文件才能登记。** |
| `delivery.audit` | `fidelity_ok` 来自 `fidelity_report.ok`；`format_status` 来自 `format_audit.status`；`render_status` 来自 `render_report.status`；`unfilled_fields` 见 §3。 |
| `evidence` | 指向各审计报告的摘要/引用（`sources` / `change_records` / `fidelity` / `format_audit` / `render`）。 |
| `slides[]` | 逐页映射：`n` / `role` / `title` / `claims[]` / `evidence[]{kind,ref,locator,cite}`。`locator` 是源材料中的真实定位（如 `p.3 §2.1`、`line 42`、`图1`）。 |
| `plan` | `template`/`source` 相对路径、实际 `slide_count`、`layout_usage`（每个版式用了多少页）；**v1.3 增补** `mode`、`start_deck_path` 与 `constraints_applied`，见 §2.1。 |
| `template_inheritance` | 真实继承链：`masters[]`、`layouts[]`、`theme`、逐页 `slide_layout_links[]`。**不得填写不存在的部件。** |
| `open_items` | `uncertainty[]`（未能确认的事实）与 `limitations[]`（明确的能力/证据缺口）。 |
| `runtime` | `prepare_s`/`finish_s` 实测秒数、`python` 版本、`renderer`、`budget`。 |
| `output_contract` | 原输入的 `output_contract` **原样保留**（键、嵌套形状、值都不动）。 |
| `contract_path_map` | 见 §3。 |

### 2.1 v1.3 增量：`plan.constraints_applied` 与相关落点

题面要求"fixed audience/time/slide count"，因此 `case.json.constraints` 必须逐项遵从并留痕。冻结落点是 **`result.json.plan.constraints_applied`**（与 `/app/output/slides_plan.json` 的同名字段同值）：

```json
"plan": {
  "template": "template.pptx", "source": "paper.md",
  "slide_count": 16, "layout_usage": {"标题和内容": 9, "空白": 2},
  "mode": "adapt_start_deck",
  "start_deck_path": "start_deck.pptx",
  "constraints_applied": {
    "audience": {"value": "…", "source": "case.json:constraints.audience"},
    "time_minutes": {"value": 15, "source": "case.json:constraints.time_minutes"},
    "target_slide_count": {"value": 16, "source": "case.json:constraints.target_slide_count"},
    "font_size_floor_pt": {"value": {"body": 14, "caption": 12}, "source": "case.json:constraints.font_size_floor_pt"},
    "caption_numbering": {"value": "preserve_source", "source": "case.json:constraints.caption_numbering"}
  }
}
```

- 这是**只增不改**的 v1.3 扩展：`template`/`source`/`slide_count`/`layout_usage` 语义不变，不新增未冻结的顶层键。
- 每项必须是 `{value, source}`；`source` 用 `case.json:constraints.<key>` 形式指向真实约束，案例未提供的项写 `null` 并在 `open_items`/`notes` 说明**不要编造**。
- `target_slide_count` 是硬上限：审计 `constraints_honoured` 会比较 `result.json.plan.slide_count <= target_slide_count`。
- 计划的 `mode` 与 `start_deck_path` 也随 `plan` 暴露出来（`from_template` 时 `start_deck_path=null`），便于审计确认起始 deck 是否被真正使用。

## 3. 原合同保留、`contract_path_map` 与 `unfilled_fields`

**保留**：`result.json.output_contract` 必须是输入 `case.json.output_contract` 的**原样快照**——不删除、不改名、不扁平化、不"翻译"键名。缺失或改动都算合同违规。

**映射**：对原合同里可被本交付满足的**每个叶子**，在 `contract_path_map` 写出一条

```json
{"delivery.artifacts": "delivery.artifacts", "identification.case_id": "identification.case_id"}
```

- 左边 = 原合同中真实存在的 dotted path 叶子；右边 = 本 `result.json` 中承载它的真实字段。
- 语义是"可机器核对的对应关系"，**不是**允许编造文字、编造文件状态或编造测量值。
- 指不到的路径不要放进映射。

**降级（`unfilled_fields`）**：不能填的叶子必须列进 `delivery.audit.unfilled_fields`（每条说明缺什么、为什么、影响哪一项产物），并据此**降级 `delivery.status`**。**不得静默补编**，也不得为了"全绿"把未知值编成已知。

> 例：原合同要求某个本案例没有的字段（源材料里不存在该数据）时，如实列入 `unfilled_fields`，`status` 至少降为 `partial`；若该叶子是主交付的关键条件，则降为 `needs_human`。

## 4. 状态语义（`ok` / `partial` / `needs_human`）

| 状态 | 判据（需同时满足） | 退出码 | 该读哪些 json |
|---|---|---|---|
| **ok** | `result.json` 已写出；审计模块均返回 ok；`fidelity_report.ok=true`；`format_audit.status=pass`；`render_report.status=rendered` 且有真实非空 PDF/PNG；deck 含模板 master/layout/theme；`unfilled_fields` 为空 | `0` | `result.json`、`fidelity_report.json`、`format_audit.json`、`render_report.json`、`sources.json`、`change_records.jsonl` |
| **partial** | 有真实可编辑 deck 与真实证据；`fidelity_report.ok=true` 且无关键格式项 fail；存在**非阻断**缺口：`format_audit.status=partial`（有 `skipped`/`unchecked`）、`render_report.status=skipped`（无渲染器）或 `failed`（渲染器存在但失败）、`unfilled_fields` 非空、目标页数未达但已说明 | `0` | 同上 + `pipeline_report.json`、`validation_report.json` |
| **needs_human** | 三类（与 `CONTRACT.md` §4 v1.2 逐字对应）：① `fidelity_report.ok=false`；② 关键格式项 fail；③ 审计模块缺失（`auditor module missing`） | `3`（参数/读取错误 `2`） | `result.json` + 已产出的 work json；缺口写进 `open_items` |

关键约束与优先级：

1. **`fidelity_report.ok=false` 是硬失败**：不得声明 `ok` 或"完成"。内容改写的证据要保留，别用宽泛规则"修绿"。
2. **`skipped` / `unchecked` / `unsupported` 都不等于 `pass`**：未检查项不得计 PASS，必须在报告里出现。
3. **诚实降级（v1.2）**：只有三类落 `needs_human` + exit `3`——`fidelity_report.ok=false`、关键格式项 fail、审计模块缺失。**渲染器不可用只降级为 `partial` + exit `0`**（交付物可用、只是缺预览），但 `render_report.json` 的 `status` 仍必须如实写 `skipped`/`failed`；`render_report.status=failed`（渲染器存在但失败、deck 完好）同样为 `partial`。
4. **`/input/case.json` 不可读或参数校验失败** → exit `2`，但仍必须写出诚实的 `needs_human` 结果。角色/模板无法确认、无法产出可用 deck 的极端情形不属于上表三类，也**不得**据此伪造完整的 `ok`/`partial` 交付：仍要写出 `result.json` 与已有中间产物，并把缺口记入 `open_items`。
5. 退出码只是流程结果：`exit 0` **不**代表保真、格式或视觉通过，必须逐份读报告。
6. 若 engine 的实现与本节某一格不同，**以 `CONTRACT.md` 为准**并报告差异；不要靠文字把状态写成通过。

### 4.1 形状命名与角色判定（v1.2 §4.2）

`format_audit` 判定形状角色以决定字号下限时，顺序固定：

1. **优先**看 `<p:ph>` 的 `idx`/`type`。
2. 无 `<p:ph>` 的文本框（版式没有对应占位符时的替代路径）看固定 `p:cNvPr@name`：

| `p:cNvPr@name` | 角色 | 字号下限（1/100 pt） |
|---|---|---|
| `Title` | title | 2400 |
| `Content Placeholder` | body（有占位符时） | 1400 |
| `Substituted Body` | body（占位符缺失时的显式文本框） | 1400 |
| `Figure Caption` | caption | 1200 |

3. **禁止**用「文本以图号开头」这类文本模式判角色（实测误报来源）。

### 4.2 `format_audit` 的 v1.3 四项新检查

四项新检查与 v1.2 的检查**同一结构**（`{name, status, expected, observed, evidence[], detail}`）并入 `format_audit.checks[]`，并同样遵守 `skipped ≠ pass`：

| `name` | 判据 | 何时 `skipped` | 失败含义 |
|---|---|---|---|
| `start_deck_preserved` | `mode=adapt_start_deck` 时，起始 deck 每页的每个文本单元都出现在输出 deck（含拆页），或该页在 `adaptation_actions.json` 有 `start_deck_slide_dropped` + 理由（及文本 hash） | `mode=from_template`（无起始 deck / 坏 zip / 任务要求新建） | 有起始 deck 的文本被静默丢弃＝丢事实，按关键格式项 fail |
| `table_rendered` | 计划 `tables[]` 的每个表格都在输出 deck 里以真正的 `<a:tbl>` 出现，行列数与 `header`/`rows` 一致，单元格文本逐字一致 | 计划里没有任何 `tables[]`（无表可渲染） | 表格被降级成图片/纯文本＝交付能力缺失，关键格式项 fail |
| `constraints_honoured` | `plan.slide_count <= target_slide_count`；`audience`/`time_minutes` 与 `case.json.constraints` 一致且非空（case 提供时）；字号下限来自 `font_size_floor_pt` | case 未提供任何可核对的 `constraints` 字段 | 约束被忽略（超页数、编造听众/时长、压到下限以下）＝关键格式项 fail |
| `notes_evidence_present` | 每页 `notes_evidence` 非空，且 `notes` 的内容 token 能在其 evidence 指向的 `units[]/sections[]/figures[]/tables[]` 源文本中找到（允许虚词/连接词） | case 不要求讲稿且计划中确实无备注（须在 `detail` 写明） | 讲稿无来源或为空＝编造风险，关键格式项 fail |

- 四项都要求 `evidence[]` 指向真实部件/字段（如 `ppt/slides/slide3.xml`、`result.json.plan.constraints_applied`、`slides_plan.json:slides[2].tables[0]`），**不接受**"看起来对"的文字结论。
- 每项都要有**负例对照**：注入 `start_deck_text_dropped` / `table_flattened_to_text` / `constraint_ignored` / `notes_unsourced` 时必须检得 `fail`；未注入时必须 `pass`（或按上表 `skipped`）。
- 四项的 `status` 会参与 `format_audit.status` 汇总：任一 `fail` → 关键格式项 fail（`needs_human`）；只有 `skipped` → `format_audit.status=partial`，且 `skipped` 的项要同时出现在 `format_audit.unchecked[]` 或该项 `detail`。

### 4.3 `fidelity_report.notes_evidence` 的语义

`fidelity_report` 不再只校验"讲稿不是空的"，而是给出**逐页讲稿 → 来源**的对照表，供 `notes_evidence_present` 与保真检查共用：

```json
"notes_evidence": [
  {"slide": 3, "evidence": ["u0012", "S2"], "notes_sha256": "…",
   "tokens_checked": 18, "unsourced_tokens": [], "status": "pass|fail|skipped"}
]
```

- `slide`：页码，与 `result.json.slides[].n` 一致；`evidence`：该页 `notes_evidence` 的原值（顺序保留）。
- `tokens_checked` / `unsourced_tokens`：实际比对过的内容 token 数与**找不到来源**的 token 列表；`unsourced_tokens` 非空即 `status=fail`，并把该页写进 `violations[]`。
- `notes_sha256`：该页讲稿文本的实算哈希，用于证明审计对象是交付 deck 里的备注而不是计划自报。
- 语义边界：允许虚词、连接词与"本页先讲…再看下一页"这类**结构性过渡语**；不允许出现源材料没有的数字、单位、结论、图号。`status=skipped` 只用于"该案例不要求讲稿且 deck 确实无备注"，且必须写 `detail`。

## 5. 哪些证据必须真实

| 证据 | 真实性要求 |
|---|---|
| **哈希** | `artifacts[].sha256` / `sources.json` 中的 `template_sha256`、`source_sha256`、`deck_sha256`、`scripts_sha256` 都必须由**最终文件实算**；不得复用旧值或抄上游自报。 |
| **渲染** | `render_report.status=rendered` 时，`preview/slides.pdf` 与 `preview/page-001.png…` 必须真实存在、非空、可打开，并登记实算 `sha256`/`bytes`；`consistency` 里 pptx 页数与 PDF 页数一致。**没有渲染器就 `skipped`**，不得借用其他文件冒充预览。 |
| **字体** | `render_report.fonts` 必须来自真实工具输出（`pdffonts` / `soffice-log` / `none`）。被替换的字体写进 `substitutions[]`（`requested` / `used` / `evidence`）；不假装字体已嵌入。 |
| **模板继承** | `template_spec.evidence[].src` 必须指向真实部件与属性；`format_audit.inheritance[]` 与 `result.template_inheritance` 里的部件路径必须存在于交付 pptx 中；`placeholder_idx` 必须真实存在。 |
| **页级定位** | `result.json.slides[].evidence[].locator` 必须是源材料中的真实定位（页码/行号/图号/部件名）；索引 `ref` 必须能在 `source_inventory` 找到。 |
| **变更记录** | `change_records.jsonl` 每条含 `ts`/`slide`/`kind`/`target`/`before`/`after`/`evidence`，且与 `adaptation_actions.json` 的逐项动作对得上。 |
| **运行时长/工具链** | `runtime.prepare_s` / `finish_s` 与 `render_report.renderer` 是实测值；探测不到就写 `skipped`/`none`，不编版本号。 |
| **起始 deck 保真（v1.3）** | `mode=adapt_start_deck` 时 `start_deck_spec.json.ok=true` 且 `slide_details[]` 来自真实字节；`start_deck_slide_dropped` 必须带理由与**该页文本的实算 hash**；不得用"计划里没列"当作删除理由。 |
| **表格（v1.3）** | `<a:tbl>` 必须由 `ppt/slides/*.xml` 的真实 XML 解析出；行列数与单元格文本必须与计划 `tables[]` 逐字一致，且 `tables[].source` 能在 `source_inventory.tables[]` 找到；不得用图片或文本框冒充。 |
| **约束（v1.3）** | `constraints_applied[].value` 必须与 `case.json.constraints` 实读值一致，`source` 指向真实路径；`audience`/`time_minutes` 不得编造，`target_slide_count` 不得超页。 |
| **讲稿（v1.3）** | `notes` 文本取自交付 deck 的 `ppt/notesSlides/*.xml`（与 `notes_sha256` 一致）；`notes_evidence[]` 必须能在 `source_inventory` 找到对应 id；讲稿不得含源材料没有的数字或结论。 |

## 6. 输出前的最小自检

1. `result.json` 可被 JSON 解析；`identification.case_id` 与输入一致；`output_contract` 与输入快照逐键一致。
2. `contract_path_map` 的每个右侧字段在 `result.json` 中真实存在；`unfilled_fields` 无遗漏。
3. `delivery.artifacts[]` 的文件都真实存在，`sha256`/`bytes` 实算一致。
4. `delivery.status` 与 §4 的判据一致（尤其：保真失败或关键格式项 fail 时不得为 `ok`；无渲染器时 `render_report.status=skipped`，状态应为 `partial` 而非 `needs_human`）。
5. `open_items.uncertainty`/`limitations` 覆盖 `partial`/`needs_human` 的全部缺口。
6. `/app/output/slides_plan.json`、`evidence/*` 与 `<work>` 中的报告内容一致（副本未过期）。
7. （v1.3）`result.json.plan.mode` 与 `start_deck_path` 自洽：`adapt_start_deck` 时输出 deck 含起始 deck 的每一页文本（或在 `adaptation_actions.json` 有 `start_deck_slide_dropped` + 理由与 hash）；`from_template` 时 `start_deck_path=null`。
8. （v1.3）`format_audit.checks[]` 含 `start_deck_preserved`/`table_rendered`/`constraints_honoured`/`notes_evidence_present` 四项，`skipped` 的都进了 `unchecked` 或 `detail`；`plan.tables[]` 非空时输出 deck 里确有对应 `<a:tbl>`。
9. （v1.3）`result.json.plan.constraints_applied` 五项与 `case.json.constraints` 一致，`plan.slide_count <= target_slide_count`；`fidelity_report.notes_evidence[]` 每页 `unsourced_tokens` 为空（或已写进 `violations[]` 并降级）。

## 7. 相关文件

- 计划字段与证据优先级：[plan_guide.md](plan_guide.md)
- 支持范围与降级：[limitations.md](limitations.md)
- 合成交付示例：[../examples/](../examples/README.md)（`example=true`，不含真实交付或分数）
