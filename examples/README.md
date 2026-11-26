# 合成接口示例（EXAMPLE；非正式案例、非答案）

**`example=true`。本目录只展示输入合同与计划的**形状**，不含真实 pptx/PDF/PNG、起始 deck 字节、正式案例、答案、运行产物或任何分数。**

本目录中每个文件的顶层或 `notes` 都标了 `example=true` / `EXAMPLE` / 示例；所有标识、版式名、单元 id、图/表 id、`source_slide` 页号、约束值、字号与合同路径都是**为演示字段而虚构**的。没有据此宣称过任何真实操作通过，也不代表评分。

## 文件

| 文件 | 演示目的 |
|---|---|
| [case.json](case.json) | 顶层 `example=true` 与 `EXAMPLE-ppt-710fbc72-001` 标识；展示 `materials[]`（`rel`/`kind`/`role_hint`，含 `start_deck` 与 `table`）、`instructions`、多层嵌套 `output_contract`，以及 **v1.3 的结构化 `constraints`**（`audience`/`time_minutes`/`target_slide_count`/`font_size_floor_pt`/`caption_numbering`） |
| [slides_plan.json](slides_plan.json) | `example=true` 与 `EXAMPLE` 标识；展示 **v1.3 计划字段**：`mode=adapt_start_deck`、`start_deck_path`、`constraints_applied`、`slides[].source_slide`/`preserve_start_text`、`slides[].tables[]`、`notes`/`notes_evidence`，以及 v1.2 的 `layout_name`/`layout_part`、`bullets[].evidence`、`figures[].aspect_lock`、`font_ladder`、`numbering`、`contract_values` |

## 关键声明

- **不是正式案例**：本示例不对应平台任何一个固定案例，不能当成某题的输入。
- **不是答案**：示例中的标题、要点、讲稿、图注、表数据、单元 id 全为合成文本；不能预填进真实 `/app/work/slides_plan.json`。
- **不是实测结果**：本目录没有 deck、没有报告、没有渲染件、没有 hash，因此不能用来声称保真、格式、视觉、约束或合同通过。
- **不代表评分**：本示例与平台评分、白板、对照结果无关。
- **不可直接运行**：`case.json` 中每个材料的 `available_in_this_example=false`，特意说明这些文件在本目录里并**不**存在（包括 `start_deck.pptx` 与 `data/table1.csv`）。

## v1.3 示例覆盖了什么

| 能力 | 示例里的形态 | 真实执行时必须替换为 |
|---|---|---|
| 起始 deck 适配 | `mode=adapt_start_deck`、`start_deck_path="start_deck.pptx"`、每页 `source_slide` 1–4、`preserve_start_text=true` | 先读 `/app/work/start_deck_spec.json`：`detected.start_deck` 唯一、`ok=true`、`slide_count>0` 才走该模式；`source_slide` 取 `slide_details[].index`，`text_units[]` 必须全部承载（见 [plan_guide §2.1](../references/plan_guide.md)） |
| 约束遵从 | `constraints_applied` 五项，每项 `{value, source}`，`source` 形如 `case.json:constraints.audience` | 值必须来自本案例 `case.json.constraints` 的**实读值**；案例未提供的项写 `null` 并说明，不得编造听众/时长 |
| 可编辑表格 | `slides[2].tables[0]`：`source="T1"`、`header`/`rows` 真实单元格文本、`caption_number="表1"` | 数据必须来自本案例 `source_inventory.tables[].header/rows`；输出 deck 必须含真正的 `<a:tbl>`，禁止截图或纯文本替代 |
| 真实讲稿 | 每页 `notes` 是可朗读句子、`notes_evidence` 非空（`uXXXX`/`S1`/`T1`） | 讲稿内容 token 必须能在 evidence 指向的源文本里找到；不得引入源材料没有的数字或结论 |

示例里刻意保持的保守选择：`confirmed_roles=false`（没有真实模板/源证据可确认角色）、`font_ladder.floor` 不低于 `common.FONT_FLOOR`、`numbering.mode="preserve_source"`（只做 `F1=图1`、`T1=表1` 的映射，不改写源编号）、`continuation_of=null`（本示例没有因溢出拆页）。

## 字段形状的用法

真实操作必须先做侦察，再把**真实证据**填进计划：

```bash
export SKILL=/skills/weekly-skill
python "$SKILL/scripts/run.py" prepare --input /input --work /app/work --output /app/output
```

然后读 `/app/work/REVIEW_PACKET.md`、`case_manifest.json`、`template_spec.json`、`source_inventory.json`、`start_deck_spec.json`，按需修改 `/app/work/slides_plan.json`，再：

```bash
python "$SKILL/scripts/run.py" finish --work /app/work --output /app/output
```

合成计划里采用的**假设证据**与真实执行时必须替换为：

| 合成观察（EXAMPLE） | 仅用于演示的字段 | 真实执行时必须替换为 |
|---|---|---|
| 模板有 `slideLayout1.xml` / `slideLayout2.xml` | `slides[].layout_part` | 本案例 `template_spec.layouts[].part` 的真实值；且每页 `layout_name` 与 `layouts[].name` 逐字一致 |
| 起始 deck 有 4 页 | `slides[].source_slide`、`preserve_start_text` | 本案例 `start_deck_spec.json.slide_details[].index`；起始 deck 里计划未列的页要保留或记 `start_deck_slide_dropped` |
| 源材料有单元 `u0012`/`u0013`/`u0044`/`u0050` 与章节 `S1` | `bullets[].evidence`、`notes_evidence`、`cite`、`evidence` | 本案例 `source_inventory.units[].id`/`sections[].id` 与真实 `locator` |
| 源材料有图 `F1`、表 `T1` | `figures[].source`/`caption`/`cite`、`tables[].source`/`header`/`rows`、`numbering.map` | 本案例 `source_inventory.figures[]`/`tables[]` 的真实 `id`/`label`/`caption`/单元格数据 |
| 存在 `fig/fig1.png`、`data/table1.csv` | `figures[].asset`、表格数据来源 | 本案例材料的真实相对路径；失效时写 `null`/留空并说明 |
| 示例合同含 `delivery.artifacts` 叶子 | `contract_values` | 本案例 `output_contract` 中**确实存在**的叶子路径；不能复制本示例的键名 |
| 示例约束值（15 分钟、4 页、14/12 pt） | `constraints`、`constraints_applied` | 本案例 `case.json.constraints` 的实读值；`target_slide_count` 是硬上限 |

**注意**：`contract_values` 的右侧不是用户可见文字，而是 `result.json` 中承载该叶子的真实标准字段。任何案例都不能照抄本示例的 dotted path——只有本示例的 `output_contract` 里恰好有 `delivery.artifacts`，别的案例不一定有。

## 检查示例本身

```bash
python -m json.tool "$SKILL/examples/case.json"
python -m json.tool "$SKILL/examples/slides_plan.json"
```

这两条只检查 JSON 语法。示例不含真实材料与侦察输出，因此**不能**做角色确认、版式存在性、证据命中、起始 deck 保真、表格渲染、约束遵从、讲稿来源、保真、视觉质量或正式运行成功验证。

## 相关文件

- 计划字段与证据优先级（含 v1.3 模式/约束/表格/讲稿）：[../references/plan_guide.md](../references/plan_guide.md)
- 交付字段、`format_audit` 四项新检查与证据真实性：[../references/result_contract.md](../references/result_contract.md)
- 支持范围与降级（PDF 无文本层、起始 deck 版式不兼容、表格过宽）：[../references/limitations.md](../references/limitations.md)
