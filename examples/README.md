# 合成接口示例（EXAMPLE；非正式案例、非答案）

**`example=true`。本目录只展示输入合同与计划的**形状**，不含真实 pptx/PDF/PNG、正式案例、答案、运行产物或任何分数。**

本目录中每个文件的顶层或 `notes` 都标了 `example=true` / `EXAMPLE` / 示例；所有标识、版式名、单元 id、图 id、字号与合同路径都是**为演示字段而虚构**的。没有据此宣称过任何真实操作通过，也不代表评分。

## 文件

| 文件 | 演示目的 |
|---|---|
| [case.json](case.json) | 顶层 `example=true` 与 `EXAMPLE-ppt-710fbc72-001` 标识；展示 `materials[]`（`rel`/`kind`/`role_hint`）、`instructions`、多层嵌套 `output_contract` 与 `constraints` |
| [slides_plan.json](slides_plan.json) | `example=true` 与 `EXAMPLE` 标识；展示 `layout_name`/`layout_part`、`bullets[].evidence`、`figures[].aspect_lock`、`notes`/`notes_evidence`、`font_ladder`、`numbering`、`contract_values` |

## 关键声明

- **不是正式案例**：本示例不对应平台任何一个固定案例，不能当成某题的输入。
- **不是答案**：示例中的标题、要点、图注、单元 id 全为合成文本；不能预填进真实 `/app/work/slides_plan.json`。
- **不是实测结果**：本目录没有 deck、没有报告、没有渲染件、没有 hash，因此不能用来声称保真、格式、视觉或合同通过。
- **不代表评分**：本示例与平台评分、白板、对照结果无关。
- **不可直接运行**：`case.json` 中每个材料的 `available_in_this_example=false`，特意说明这些文件在本目录里并**不**存在。

## 字段形状的用法

真实操作必须先做侦察，再把**真实证据**填进计划：

```bash
export SKILL=/skills/weekly-skill
python "$SKILL/scripts/run.py" prepare --input /input --work /app/work --output /app/output
```

然后读 `/app/work/REVIEW_PACKET.md`、`case_manifest.json`、`template_spec.json`、`source_inventory.json`，按需修改 `/app/work/slides_plan.json`，再：

```bash
python "$SKILL/scripts/run.py" finish --work /app/work --output /app/output
```

合成计划里采用的**假设证据**与真实执行时必须替换为：

| 合成观察（EXAMPLE） | 仅用于演示的字段 | 真实执行时必须替换为 |
|---|---|---|
| 模板有 `slideLayout1.xml` / `slideLayout2.xml` | `slides[].layout_part` | 本案例 `template_spec.layouts[].part` 的真实值；且每页 `layout_name` 与 `layouts[].name` 逐字一致 |
| 源材料有单元 `u0012`/`u0013`/`u0044` | `bullets[].evidence`、`notes_evidence`、`cite`、`evidence` | 本案例 `source_inventory.units[].id` 与对应的真实 `locator` |
| 源材料有图 `F1`，标签为 `图1` | `figures[].source`/`caption`/`cite`、`numbering.map` | 本案例 `source_inventory.figures[]` 的真实 `id`/`label`/`caption`/`asset` |
| 存在 `fig/fig1.png` | `figures[].asset` | 本案例材料的真实相对路径；失效时写 `null` 并说明 |
| 示例合同含 `delivery.artifacts` 叶子 | `contract_values` | 本案例 `output_contract` 中**确实存在**的叶子路径；不能复制本示例的键名 |
| 计划声明 `confirmed_roles=false` | `confirmed_roles` | 真实案例核对材料后决定；候选不唯一时必须保持 `false` |

**注意**：`contract_values` 的右侧不是用户可见文字，而是 `result.json` 中承载该叶子的真实标准字段。任何案例都不能照抄本示例的 dotted path——只有本示例的 `output_contract` 里恰好有 `delivery.artifacts`，别的案例不一定有。

## 逐项说明

- `source_path` / `template_path` 是输入根相对路径；本示例的 `confirmed_roles=false` 表示"没有真实证据可确认角色"，不代表真实案例也应为 `false`。
- `slides[].role` 只用冻结枚举：`title|outline|method|result|figure|table|discussion|conclusion|references`。
- `figures[].aspect_lock=true` 表示等比缩放；示例没有真实图片，因此无法验证 `aspect_delta`。
- `font_ladder` 的 `floor` 不低于 `common.FONT_FLOOR`（title 2400 / body 1400 / caption 1200 / table 1200 / notes 1000，1/100 pt）；示例不代表真实模板的可用字号。
- `numbering.mode="preserve_source"` 只建立 `F1=图1` 的映射，不改写源图号。
- `continuation_of=null` 表示本示例没有因溢出拆页；真实拆页的续页必须指向原页 `n`。

## 检查示例本身

```bash
python -m json.tool "$SKILL/examples/case.json"
python -m json.tool "$SKILL/examples/slides_plan.json"
```

这两条只检查 JSON 语法。示例不含真实材料与侦察输出，因此**不能**做角色确认、版式存在性、证据命中、保真、视觉质量或正式运行成功验证。

## 相关文件

- 计划字段与证据优先级：[../references/plan_guide.md](../references/plan_guide.md)
- 交付字段、状态与证据真实性：[../references/result_contract.md](../references/result_contract.md)
- 支持范围与降级：[../references/limitations.md](../references/limitations.md)
