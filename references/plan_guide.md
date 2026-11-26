# `slides_plan.json` 计划编写指南

本文件说明 `/app/work/slides_plan.json` 的**全部字段**、证据优先级、常见冲突处置，以及与 engine 兜底推导（`plan_source=auto_derived`）一致的规则。

计划是 Agent 与 engine 之间唯一的排版意图接口。它**不是**答案：`prepare` 写出的计划只是保守壳，必须按真实侦察证据编辑；不编辑时 `finish` 可能退回 `auto_derived` 推导（见 §7）。

计划形态以 `CONTRACT.md` §5.4 为基础，**v1.3 增量**（`mode` / `start_deck_path` / `constraints_applied` / `slides[].source_slide` / `slides[].preserve_start_text` / `slides[].tables[]` / `notes_evidence[]` 的必填语义）见 `ADDENDUM-v1.3.md` §A2，本文件是二者的可读说明。写入时用 UTF-8、`ensure_ascii=False`，不要添加未冻结的键。

## 1. 顶层字段

| 字段 | 类型 | 说明 |
|---|---|---|
| `plan_version` | string | 固定 `"1"`，与冻结版本对应。 |
| `example` | bool | 交付计划固定 `false`；只有 `examples/` 下的示例允许 `true`。 |
| `source_path` | string | 源材料在 `/input` 下的**相对路径**，必须与 `case_manifest.inputs[].rel` 逐字一致。 |
| `template_path` | string | 目标模板的相对路径，同样必须来自 manifest；不得是猜测的文件名。 |
| `mode` | `from_template` \| `adapt_start_deck` | v1.3 运行模式。有唯一可解析起始 deck 且任务要求适配时为 `adapt_start_deck`，否则 `from_template`；选定规则见 §2.1，只按真实证据判定，不按文件名猜。 |
| `start_deck_path` | string \| null | v1.3：起始 deck 的真实相对路径（与 `case_manifest.detected.start_deck` 一致）；`mode=from_template` 时为 `null`。v1.2 的同义旧键 `start_deck` 不再使用，计划里以 `start_deck_path` 为准。 |
| `deck_name` | string | 输出 deck 文件名，默认 `slides.pptx`。 |
| `language` | `zh` \| `en` \| `mixed` | 讲稿与标题的主要语言，影响换行估算与字号选择，不改变源文字。 |
| `audience` | string | 听众与场合，**必须**取自 `case.json.constraints.audience`；案例没给就置 `null` 并在 `notes` 说明，不得凭空编造。 |
| `time_minutes` | int \| null | 汇报时长，取自 `case.json.constraints.time_minutes`；与 `target_slide_count` 一起决定内容密度。 |
| `target_slide_count` | int \| null | 目标页数，取自 `case.json.constraints.target_slide_count`，是**硬上限**：实际页数不得超过；因拆页减少可用内容时在 `notes` 说明。 |
| `constraints_applied` | object | v1.3："约束取自哪、值是多少"的逐项留痕，见 §2.2。 |
| `confirmed_roles` | bool | 只有核查材料后（见 §2）才能置 `true`。角色冲突或候选不唯一时必须保持 `false`。 |
| `slides` | array | 页级计划，见 §3。 |
| `font_ladder` | object | 字号阶梯，见 §4。 |
| `numbering` | object | 图号策略，见 §5。 |
| `contract_values` | object | 原合同 dotted path → 本交付真实字段的映射，见 §6。 |
| `notes` | string[] | 决策依据、偏离说明、未决项。**不得**用 notes 绕过字段校验。 |

## 2. 确认源与模板角色（`confirmed_roles`）

1. 读 `case_manifest.json` 的 `detected.template` / `detected.source` / `detected.start_deck` / `detected.figures` 与每个输入的 `rel/kind/size/sha256`。
2. 读 `REVIEW_PACKET.md` 与 `task.md`，确认"任务实际要求套用哪个模板、用哪份材料"。
3. `source_path`、`template_path` 必须精确对应 manifest 的相对路径。多个貌似候选、角色不明、版本冲突时，**不要用文件名赌博**：保持 `confirmed_roles=false`，把缺哪个事实写进 `notes` 与 `open_items.uncertainty`，结果降级。
4. **不要**把 `confirmed_roles` 当成"要点一下的确认框"机械填 `true`。

若任务实际只需套用一个明确模板，不要因其他无关附件停滞；用任务与材料的真实证据说明选择，并保留其余材料的只读副本。

### 2.1 运行模式与起始 deck（`mode` / `start_deck_path`，v1.3）

题面明确要求"actual starting deck where edit/adapt requires one"，所以**有起始 deck 却不用它**是不合格的。是否走适配只按证据判定：

| 事实（全部来自真实侦察产物） | `mode` | `start_deck_path` |
|---|---|---|
| `case_manifest.detected.start_deck` 唯一命中，`start_deck_spec.json.ok=true` 且 `slide_count>0`，任务要求"编辑/适配已有 deck" | `adapt_start_deck` | 该起始 deck 的相对路径 |
| 没有起始 deck，或 `start_deck_spec.json.ok=false`（坏 zip / 解析失败），或任务明确要求"从模板新建" | `from_template` | `null` |
| 多个候选、角色冲突、任务与材料互相矛盾 | 按可核证的那条执行（通常是"有唯一真实起始 deck"→适配），**另一条写进 `notes` 与 `result.json.open_items.uncertainty`**，并在审计里让 `start_deck_preserved` 记 `skipped` 而不是 `pass` | 按证据填或 `null` |

- `start_deck_path` 必须与 manifest 里的 `rel` 逐字一致；**不要**用 `start_deck_spec.json.path` 之外的猜测文件名。
- 只解析到 `start_deck_spec.ok=true` 才允许走 `adapt_start_deck`；坏 zip 一律 `from_template` 并如实记降级。
- `mode=adapt_start_deck` 时，`slides[]` 用 `source_slide` 把输出页绑到起始 deck 的页；`mode=from_template` 时 `source_slide` 必须为 `null`。

### 2.2 约束遵从与 `constraints_applied`（v1.3）

`case.json.constraints` 一旦提供值就必须**逐项**落到计划与 deck，并在 `constraints_applied` 里记来源与值：

```json
"constraints_applied": {
  "audience": {"value": "…", "source": "case.json:constraints.audience"},
  "time_minutes": {"value": 15, "source": "case.json:constraints.time_minutes"},
  "target_slide_count": {"value": 16, "source": "case.json:constraints.target_slide_count"},
  "font_size_floor_pt": {"value": {"body": 14, "caption": 12}, "source": "case.json:constraints.font_size_floor_pt"},
  "caption_numbering": {"value": "preserve_source", "source": "case.json:constraints.caption_numbering"}
}
```

| 约束 | 落地要求 |
|---|---|
| `audience` | 写进 `result.json.identification`/计划的 `audience`，用于详略与术语密度；**不得编造**听众。 |
| `time_minutes` | 与 `target_slide_count` 共同决定内容密度；讲稿总长度要能在时长内讲完（按语言语速估算），偏离写 `notes`。 |
| `target_slide_count` | **硬上限**，实际页数（含拆页）不得超过；无法在页数内讲完时缩减页数而不是超页，并把取舍写进 `notes`。 |
| `font_size_floor_pt` | 是 `font_ladder.floor` 的**来源**；`floor` 换算成 1/100 pt（14 pt → 1400）后不得低于该值，也不得低于 `common.FONT_FLOOR`。 |
| `caption_numbering` | 默认 `preserve_source`：图/表号沿用源编号，`numbering.map` 只做映射不改写。 |

- 案例未提供的约束项：`constraints_applied` 里写 `null`（或省略该键）并在计划 `notes` 说明"案例未提供，未套用"，不得凭空补一个值。
- `constraints_applied` 会随计划进入 `result.json.plan.constraints_applied`，是审计 `constraints_honoured` 的核对入口。

## 3. `slides[]` 页级字段

| 字段 | 类型 | 说明 |
|---|---|---|
| `n` | int | 页码，从 1 起，唯一且递增。 |
| `role` | enum | `title` \| `outline` \| `method` \| `result` \| `figure` \| `table` \| `discussion` \| `conclusion` \| `references`。角色决定该页承担的论文主线位置。 |
| `layout_name` | string | **必须**是 `template_spec.layouts[].name` 的真实值（逐字，含中英文与空格）。 |
| `layout_part` | string | 对应的真实部件路径，如 `ppt/slideLayouts/slideLayout2.xml`。与 `layout_name` 同源，用于生成后可查证。 |
| `title` | string | 标题文字；来自源材料，不因版式而改写。过长时按 §4 处理，不截断语义。 |
| `subtitle` | string | 副标题；无证据时留空，不编造。 |
| `bullets` | array | `{text, level, evidence[]}`，见下。 |
| `figures` | array | `{source, asset, caption, cite, aspect_lock}`，见下。 |
| `tables` | array | v1.3：`{source, caption, header, rows, evidence[], caption_number}`，见下；非空时该页必须含真正的 `<a:tbl>`。 |
| `source_slide` | int \| null | v1.3：仅 `adapt_start_deck`，起始 deck 的 **1-based** 页号；该页文本必须保留（可重排/降号，不可删改事实）。 |
| `preserve_start_text` | bool | v1.3：`adapt_start_deck` 下为 `true`，表示起始 deck 该页文本单元必须出现在输出（含拆页）；`from_template` 下为 `false`。 |
| `notes` | string | 讲稿正文：**可朗读的句子**（本页主张 → 依据 → 与下一页衔接），不是 `来源：locator` 串。 |
| `notes_evidence` | string[] | 讲稿来源单元 id（`uXXXX` / `S2` / `F1` / `T1`），**必填非空**；讲稿内容 token 必须能在其源文本中找到。 |
| `claims` | string[] | 该页**页级主张**，一句话一条，供他人逐页对照源材料。 |
| `evidence` | string[] | 该页所有证据 id 的并集（单元、图、表）。 |
| `continuation_of` | int \| null | 由拆页产生的续页指向原页号；非续页为 `null`。 |

### `bullets[]`

- `text`：要点正文，**来自源材料**，不新增未在源中出现的断言。
- `level`：缩进层级，从 `0` 起；层级只影响视觉与换行估算，不改变内容。
- `evidence[]`：每条要点指向 `source_inventory.units[].id`（如 `u0012`）或 `figures[].id`（如 `F1`）。**没有证据的要点不要写。**

### `figures[]`

- `source`：`source_inventory.figures[].id`（如 `F1`）。
- `asset`：图的真实相对路径（如 `fig/fig1.png`）；源中等价物存在时才写，失效时写 `null` 并在 `notes`/`open_items` 说明。
- `caption`：图注文字，**必须**与源题注一致；不重写、不翻译、不改图号。
- `cite`：正文中引用该图的源单元 id。
- `aspect_lock`：默认 `true`。为 `true` 时按原图真实宽高比等比缩放（contain），必要时裁切，**不得**非等比拉伸。

### `tables[]`（v1.3，可编辑表格）

```json
{"source": "T1", "caption": "表1 …", "header": ["列1", "列2"], "rows": [["a", "b"]],
 "evidence": ["u0044"], "caption_number": "表1"}
```

- `source`：`source_inventory.tables[].id`（如 `T1`）。拿不到真实 `header/rows` 时**不要编造**：留空数组并在 `notes`/`open_items` 说明"源未提供单元格数据"。
- `header` / `rows`：真实单元格文本，逐字来自源（`source_inventory.tables[].header/rows`）；`rows` 是二维数组，行数/列数必须自洽。
- `caption` / `caption_number`：题注与编号沿用源（默认 `preserve_source`），不重写、不翻译。
- `evidence[]`：该表数据来源单元 id，便于逐单元格回溯。
- 交付要求：每张表在 slide 里渲染成**真正的** `<p:graphicFrame><a:graphic><a:graphicData uri="…/table"><a:tbl>`，行列数与 `header`/`rows` 一致、单元格文字逐字一致；**禁止**降级成图片截图或纯文本行。列宽/字号策略与"表过宽"的处置见 `limitations.md` §3.10。
- 审计入口：`format_audit.checks[]` 的 `table_rendered`。

### 起始 deck 来源页（`source_slide` / `preserve_start_text`）

- `source_slide` 指向 `start_deck_spec.json.slide_details[].index`（1-based）；该页的 `text_units[]` 与 `pictures[]` 是**必须被承载**的内容。
- 允许：重排页序、按 `font_ladder` 降字号、因溢出拆页（续页 `continuation_of` 指向本页 `n`）。
- **不允许**：改写/摘要/合并句子、删数字或单位、丢文本单元。确需丢弃一页时，`adapt_pptx.py` 必须在 `adaptation_actions.json` 记 `start_deck_slide_dropped` + 理由 + 该页文本 hash；计划里应在 `notes` 写明同一理由，否则审计判 `start_deck_preserved=fail`。
- 起始 deck 里存在、而计划未列出的页：**必须保留**为一页（默认追加在末尾），动作记 `start_deck_slide_preserved`。

### 占位符缺失时的替代路径（v1.2 §4.2 形状命名）

不在 `template_spec.layouts[].placeholders[]` 中的 `idx`/`ph_type` **不得凭空写入**。优先用同版式其他真实占位符；确实没有对应占位符时，用该版式的文本框等价替代，并**必须**给固定 `p:cNvPr@name`，审计据此判角色与字号下限：

| `p:cNvPr@name` | 角色 | 字号下限（1/100 pt） |
|---|---|---|
| `Title` | title | 2400 |
| `Content Placeholder` | body（有占位符时） | 1400 |
| `Substituted Body` | body（占位符缺失时的显式文本框） | 1400 |
| `Figure Caption` | caption | 1200 |

角色判定顺序：**优先**看 `<p:ph>` 的 `idx`/`type`；没有 `<p:ph>` 才看上面的 name。**禁止**用「文本以图号开头」这类文本模式判角色——这是实测踩过的误报来源（图注会被误判成 body，从而用错字号下限）。替代动作仍要在 `adaptation_actions.json` 记 `placeholder_substitute` 与 `before/after/evidence`，并在 plan `notes` 说明限制。

## 4. `font_ladder` 与溢出处理

```json
"font_ladder": {"title": [3200, 2800, 2400], "body": [2000, 1800, 1600, 1400],
                "caption": [1400, 1200], "floor": {"body": 1400, "caption": 1200}}
```

- 数值单位是 **1/100 pt（OOXML 原值）**，与 `template_spec` 的 `sz` 同单位。
- 阶梯只能**向下**取；`floor` 不得超过 `common.FONT_FLOOR`：title 2400 / body 1400 / caption 1200 / table 1200 / notes 1000（1/100 pt）。
- 处理顺序（与 engine 一致）：**先沿阶梯降字号 → 到下限仍溢出就拆页**。溢出估算用确定性的 `common.estimate_wrapped_lines` / `common.capacity_lines`（Helvetica 宽度表 + CJK 全角 1.0 em），不依赖本机字体文件。
- 拆页时新页 `continuation_of` 指向原页 `n`，并在 plan `notes` 记录"因溢出拆页，非删字"。
- **禁止**低于 `floor` 仍然硬塞；无法在版心内安全承载时如实记 `unresolved` 并降级，而不是继续压字号。

## 5. `numbering`

```json
"numbering": {"mode": "preserve_source", "map": ["F1=图1"]}
```

- `mode="preserve_source"`（默认）：保留源编号，`map` 只建立"源 id → 实际显示编号"的对应，**不改写**源编号文字。
- `mode="relabel"`：仅在任务**显式要求**重编号时使用，并在 plan `notes` 写明依据；重编号后 `fidelity_report` 的 `figure_numbering` 检查仍需与正文一致。
- `map` 的每一对必须能在 `source_inventory.figures[].label` 找到依据。

## 6. `contract_values`

把原合同的 dotted path 指向本交付的**真实标准字段**，例如：

```json
"contract_values": {"delivery.artifacts": "delivery.artifacts"}
```

- 左边是 `/input/case.json` `output_contract` 中真实存在的叶子路径；右边是 `result.json` 中能承载它的字段。
- 不能填的叶子**不要**用 `contract_values` 编文字或编文件状态；让 `finalize` 把它们列入 `delivery.audit.unfilled_fields` 并降级。
- 指向关系会被记录为 `result.json` 的 `contract_path_map`（见 `references/result_contract.md`）。

## 7. 证据优先级与冲突处置

**优先级（高 → 低）：**

1. **显式指令**：`/input/task.md` 与 `case.json.instructions` 中直接写明的页数、语言、输出名、必含内容。
2. **`case.json.constraints`**：正式约束（如字号不可低于下限、图号策略、页面尺寸）。
3. **模板实际证据**：`template_spec.json` 中来自真实 pptx 字节的 `slide_size` / `layouts[].name` / `placeholders[]` / `theme` / `evidence[].xpath`。
4. **源材料结构**：`source_inventory.json` 的 `sections` / `units` / `figures` / `tables`，决定主线与页序。

冲突处置规则：

- **不静默选一条**。把两条来源都写进 plan `notes`，选可核证的那条（通常更高优先级），并在 `open_items.uncertainty` 记录另一条。
- 显式指令与模板证据冲突（例如要求某版式但模板没有）：以指令为准去**表述目标**，以模板为准去**实现方式**——退到语义最近的真实版式，记录 `placeholder_substitute` 与限制。
- 源材料自相矛盾（同一图两个编号、两个题注）：保留两种可核证的写法之一并在 `notes`/`uncertainty` 说明，**不得**取平均或自行发明第三个。
- **起始 deck 内容与源材料冲突（v1.3）**：起始 deck 往往早于源材料定稿。规则是——**始终以源材料（`paper`/源正文）为事实基准**：数字、单位、图号、结论一律取源材料的值。起始 deck 的文本单元仍**必须保留**（`preserve_start_text=true`），但若其措辞与源材料矛盾，不得把冲突文本当事实输出：优先保留源材料一致的表述并说明改了什么，或在无法两全时保留起始 deck 原文但把冲突记进 `open_items.uncertainty`（写清两边原文与出处），审计侧如实标 `start_deck_preserved` 与保真检查的结果。**禁止**静默二选一，禁止用一个折中值把两边都掩盖。
- **起始 deck 页数 vs `target_slide_count`**：起始 deck 页数已超上限时，先合并/拆页重排承载内容，仍超就按"删页需 `start_deck_slide_dropped` + 理由"处理，并把取舍写进 `notes`；不得只删内容不留痕。
- 页数与时长冲突：以指令为准，偏离在 `notes` 说明（拆页导致的增加必须说明）。

## 8. 预算内没有可靠计划时的兜底（`auto_derived`）

若 `finish` 发现 `slides_plan.json` 缺失、角色未确认或关键路径为空，它**不会**直接失败，而是从同一批真实侦察证据推导一份**保守计划**，并把决策写进 `plan_resolution.json`（`plan_source: agent | auto_derived | unusable`）；原计划（如有）另存为 `agent_plan_backup` 指向的路径。推导规则全部基于证据，与 engine 一致：

- **只在恰好各有一个显式证据候选时**确认 `source_path` / `template_path`；候选不唯一或角色冲突时保持 `unusable`，如实记入 `open_items` 并给出诚实的中间产物（该情形**不属于** `CONTRACT.md` §4 v1.2 的三类 `needs_human`，但**不得**伪造完整的 `ok`/`partial` 交付）。
- 版式与样式**只按真实名称匹配**：`layout_name` 取 `template_spec.layouts[].name` 的精确值，不做模糊猜测；匹配不到就不生成该页或退到真实存在的通用版式并记录。
- 页角色只在证据唯一时确认（例如源中唯一的一级标题 + 首页 → `title`）；不确定时用中性角色并记 `notes`。
- 溢出处理顺序固定为**先阶梯后拆页**，下限使用 `common.FONT_FLOOR`。
- **只做证据支持的改动**：不新增无证据要点、不改写正文、不臆造图号与题注；`contract_values` 只按叶子名映射到真实标准字段。
- 图号默认 `preserve_source`。
- **模式与约束也按证据推导（v1.3）**：`detected.start_deck` 唯一、`start_deck_spec.ok=true` 且 `slide_count>0`、任务要求适配 → `mode=adapt_start_deck` 并填 `start_deck_path`，否则 `from_template`；`audience` / `time_minutes` / `target_slide_count` / `font_size_floor_pt` / `caption_numbering` 只从 `case.json.constraints` 取，取不到的记 `notes`，不编造。`target_slide_count` 当硬上限用。
- **表格与讲稿**：`source_inventory.tables[]` 有真实 `header/rows` 时才生成 `tables[]`（否则留空并说明）；`notes_evidence` 非空，`notes` 写可朗读句子，不写定位串。

兜底计划与手写计划走**完全相同的保真/格式/预览门禁**。它降低"计划填错就零交付"的阻断风险，**不降低验收标准**；要精确控制映射仍应显式编写计划。

## 9. 提交前的最小自检

1. `source_path` / `template_path` 与 `case_manifest.inputs[].rel` 逐字一致，且 `confirmed_roles` 有证据。
2. 每个 `layout_name` 都能在 `template_spec.layouts[].name` 找到；每个 `layout_part` 都真实存在。
3. 每条 `bullets[].evidence` / `notes_evidence` / `figures[].cite` 都能在 `source_inventory` 找到对应 id。
4. `font_ladder` 未低于 `floor`，`floor` 未低于 `common.FONT_FLOOR`。
5. `numbering.mode` 与任务要求一致；默认 `preserve_source`。
6. `contract_values` 左侧叶子在 `output_contract` 中真实存在，右侧字段在 `result.json` 中真实存在。
7. 版式缺占位符而用替代文本框时，`p:cNvPr@name` 只用 §3 的四个固定名字，且角色判定没有依赖文本模式。
8. plan `notes` 写清所有偏离、冲突与未决项。
9. （v1.3）`mode` 与 `start_deck_path` 一致：`adapt_start_deck` 时 `start_deck_path` 非空且与 manifest 逐字一致，每页 `source_slide` 都是起始 deck 的真实页号且 `preserve_start_text=true`；`from_template` 时两者分别为 `from_template`/`null` 且 `source_slide=null`。
10. （v1.3）`constraints_applied` 覆盖 `audience`/`time_minutes`/`target_slide_count`/`font_size_floor_pt`/`caption_numbering` 五项，每项都有 `value` 与 `source`；未提供的项为 `null` 并在 `notes` 说明。
11. （v1.3）每个 `tables[]` 的 `source` 在 `source_inventory.tables[]` 存在，`header`/`rows` 与源数据逐字一致（或明确留空并说明）；不要放"示意数字"。
12. （v1.3）每页 `notes_evidence` 非空，`notes` 是可朗读句子；抽查 2–3 个内容词能在 evidence 源文本里找到。

## 10. 相关文件

- 交付字段与状态语义：[result_contract.md](result_contract.md)
- 支持范围与降级：[limitations.md](limitations.md)
- 合成最小计划示例：[../examples/slides_plan.json](../examples/slides_plan.json)（`example=true`，不是答案）
