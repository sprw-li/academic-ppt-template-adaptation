# `slides_plan.json` 计划编写指南

本文件说明 `/app/work/slides_plan.json` 的**全部字段**、证据优先级、常见冲突处置，以及与 engine 兜底推导（`plan_source=auto_derived`）一致的规则。

计划是 Agent 与 engine 之间唯一的排版意图接口。它**不是**答案：`prepare` 写出的计划只是保守壳，必须按真实侦察证据编辑；不编辑时 `finish` 可能退回 `auto_derived` 推导（见 §7）。

计划形态以 `CONTRACT.md` §5.4 为准。写入时用 UTF-8、`ensure_ascii=False`，不要添加未冻结的键。

## 1. 顶层字段

| 字段 | 类型 | 说明 |
|---|---|---|
| `plan_version` | string | 固定 `"1"`，与冻结版本对应。 |
| `example` | bool | 交付计划固定 `false`；只有 `examples/` 下的示例允许 `true`。 |
| `source_path` | string | 源材料在 `/input` 下的**相对路径**，必须与 `case_manifest.inputs[].rel` 逐字一致。 |
| `template_path` | string | 目标模板的相对路径，同样必须来自 manifest；不得是猜测的文件名。 |
| `start_deck` | string \| null | 需要"在已有 deck 上编辑"时指向起始 deck 的真实相对路径；否则 `null`。 |
| `deck_name` | string | 输出 deck 文件名，默认 `slides.pptx`。 |
| `language` | `zh` \| `en` \| `mixed` | 讲稿与标题的主要语言，影响换行估算与字号选择，不改变源文字。 |
| `audience` | string | 听众与场合（来自任务说明或 case），用于判断详略与页数。 |
| `time_minutes` | int | 汇报时长；与 `target_slide_count` 一起决定内容密度。 |
| `target_slide_count` | int | 目标页数；实际页数可因拆页增加，但必须在 notes 说明偏离。 |
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
| `notes` | string | 讲稿备注文本；应由 `notes_evidence` 支撑。 |
| `notes_evidence` | string[] | 备注来源单元 id（`uXXXX` / `F1`）。 |
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
- 页数与时长冲突：以指令为准，偏离在 `notes` 说明（拆页导致的增加必须说明）。

## 8. 预算内没有可靠计划时的兜底（`auto_derived`）

若 `finish` 发现 `slides_plan.json` 缺失、角色未确认或关键路径为空，它**不会**直接失败，而是从同一批真实侦察证据推导一份**保守计划**，并把决策写进 `plan_resolution.json`（`plan_source: agent | auto_derived | unusable`）；原计划（如有）另存为 `agent_plan_backup` 指向的路径。推导规则全部基于证据，与 engine 一致：

- **只在恰好各有一个显式证据候选时**确认 `source_path` / `template_path`；候选不唯一或角色冲突时保持 `unusable`，如实记入 `open_items` 并给出诚实的中间产物（该情形**不属于** `CONTRACT.md` §4 v1.2 的三类 `needs_human`，但**不得**伪造完整的 `ok`/`partial` 交付）。
- 版式与样式**只按真实名称匹配**：`layout_name` 取 `template_spec.layouts[].name` 的精确值，不做模糊猜测；匹配不到就不生成该页或退到真实存在的通用版式并记录。
- 页角色只在证据唯一时确认（例如源中唯一的一级标题 + 首页 → `title`）；不确定时用中性角色并记 `notes`。
- 溢出处理顺序固定为**先阶梯后拆页**，下限使用 `common.FONT_FLOOR`。
- **只做证据支持的改动**：不新增无证据要点、不改写正文、不臆造图号与题注；`contract_values` 只按叶子名映射到真实标准字段。
- 图号默认 `preserve_source`。

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

## 10. 相关文件

- 交付字段与状态语义：[result_contract.md](result_contract.md)
- 支持范围与降级：[limitations.md](limitations.md)
- 合成最小计划示例：[../examples/slides_plan.json](../examples/slides_plan.json)（`example=true`，不是答案）
