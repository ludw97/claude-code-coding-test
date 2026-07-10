# 工作流改动日志 — 2026-07-04

## 一、物业管理费及维修费分析 skill

### 1. Bug 修复：`scripts/property_analysis.py` — 2025 旧格式 campus_map 缺少奉贤院区映射

**文件**：`.claude/skills/物业管理费及维修费分析/scripts/property_analysis.py`

**问题**：`detect_format()` 对 2025 旧格式的 `campus_map` 仅配置了 `{"杨浦院区": "杨浦"}`，缺少 `"奉贤院区": "奉贤"`，导致上年同期奉贤院区数据在取数时全部丢失。

**修改**（两处）：
- 第 124 行：`campus_map` 由 `{"杨浦院区": "杨浦"}` 改为 `{"杨浦院区": "杨浦", "奉贤院区": "奉贤"}`
- 第 145 行（fallback 分支）：同上

### 2. 配置同步：`references/category_mapping.json` — 2025 campus_map 修正

**文件**：`.claude/skills/物业管理费及维修费分析/references/category_mapping.json`

**修改**：
- 第 30 行：`format_detection.2025.campus_map` 由 `{"杨浦院区": "杨浦"}` 改为 `{"杨浦院区": "杨浦", "奉贤院区": "奉贤"}`

### 3. 文档更新：`SKILL.md`

**文件**：`.claude/skills/物业管理费及维修费分析/SKILL.md`

**改动**：
- 新增「格式检测」章节，统一描述 2026/2025 两种格式的列布局、前缀规则、campus_map
- campus_map 项补充 2025 格式的奉贤院区映射
- 删除验证项"上年数据仅来自杨浦院区"（已不成立）
- 新增「扩展 > 格式检测」说明

---

## 二、成本分析 skill

### 1. 文档更新：`SKILL.md`

**文件**：`.claude/skills/成本分析/SKILL.md`

**改动**：
- Slide 编号修正：`50→43, 51→44, 52→45, 53→46, 54→47, 55→48, 56→49, 57→50, 65→59`（与 `mapping.json` 实际 slide ID 对齐）
- 移除 Slide 75 引用（不在当前 mapping 配置中）
- 卫生材料明细：更新年份检测逻辑描述（sheet 名称优先而非列数判断），说明两期文件均可含院区列
- 新增「与现有技能的关系」章节（先运行物业维修费分析 → 再运行成本分析）
- 示例文件名更新为 6 月场景

---

## 三、执行记录

| 步骤 | 执行的 skill | 输入文件 | 输出文件 |
|------|-------------|----------|----------|
| 1 | 物业管理费及维修费分析 | `202606物业维修费.xlsx` / `202506物业维修费.xls` / `6月分析底稿_updated_0704.xlsx` | `6月分析底稿_updated_0704.xlsx` |
| 2 | 成本分析 | `6月分析底稿_updated_0704.xlsx` / `新华医院2026年6月运营分析.pptx` | `新华医院2026年6月运营分析_updated_0704.pptx` |

---

## 四、新 skill 创建（2026-07-04）

### 1. 创建：物业管理费及维修费分析累积月份 skill

**目录**：`.claude/skills/物业管理费及维修费分析累积月份/`

**内容**：基于 **物业管理费及维修费分析** skill 完整复制，包含：
- `SKILL.md`：name 改为 `物业管理费及维修费分析累积月份`，标题改为「物业管理费及维修费分析（累积月份）」，触发场景改为累积分析相关关键词
- `scripts/property_analysis.py`：与原 skill 完全相同
- `references/category_mapping.json`：与原 skill 完全相同

### 2. 创建：成本分析累积月份 skill

**目录**：`.claude/skills/成本分析累积月份/`

**内容**：基于 **成本分析** skill 完整复制，包含：
- `SKILL.md`：name 改为 `成本分析累积月份`，标题改为「成本分析技能（累积月份）」，触发场景改为累积分析相关关键词
- `scripts/cost_analysis.py`：与原 skill 完全相同
- `references/mapping.json`：与原 skill 完全相同

---

## 五、累积月份 skill 修改（2026-07-04）

### 1. 修改：物业管理费及维修费分析累积月份 skill — 从单月改为累计数据

**目标**：将 skill 从处理单月明细数据改为处理1月至当前月份的累计明细数据，与上年同期累计进行同比对比分析。

**文件变更**：

#### `SKILL.md`
- **Frontmatter description**：强化累计数据描述，新增"与单月版本不同"说明
- **正文开头**：新增醒目提示框，解释与单月版本的关键区别（累计数据 vs 单月数据）
- **输入章节**：
  - "本年当月" → "本年累计（1月至当前月份）"
  - 示例文件名从 `202606物业及维修.xlsx` → `2026年1-6月物业维修费.xlsx`
  - 补充说明每个文件包含1月至当前月的全部交易数据
- **执行章节**：
  - 脚本路径从 `物业管理费及维修费分析/` → `物业管理费及维修费分析累积月份/`
  - 示例改为累计文件名
- **更新范围章节**：
  - 物业管理 C/D/E 列：当月合计/奉贤院区/上年同期 → 本年累计/奉贤累计/上年同期累计
  - 维修护费 C/D/E 列：当月合计/奉贤院区/上年同期 → 本年累计/奉贤累计/上年同期累计
  - 后勤维修 D/E/F 列：当月合计/奉贤院区/上年同期 → 本年累计/奉贤累计/上年同期累计
  - 商品服务 C6/C7/D6/D7/E6/E7：标注为累计数据
- **验证章节**：所有"当月"改为"本年累计"，新增"本年累计数据应 ≥ 任一单月数据"检查项
- **与现有技能的关系**：重写为对比单月版本与累积版本的区别，引用「成本分析累积月份」skill

#### `scripts/property_analysis.py`
- **文件 docstring**：更新为累积月份版本描述，说明与单月版本的区别
- **参数帮助文本**：
  - `current_file`：本年当月 → 本年累计（1月至当前月份）
  - `prev_file`：上年同期 → 上年同期累计（1月至当前月份）
  - `description`：增加"累积月份"标识
- **核心逻辑**：`load_categories()`、`load_maintenance_data()` 等数据处理函数保持不变，因为无论单月还是累计文件，都是汇总文件中的所有交易行。差异仅在于输入文件的数据范围。
- **输出日志**：分类汇总打印保持原样，数据含义取决于输入文件的范围。

#### `references/category_mapping.json`
- 无需修改（纯配置数据，列映射和分类规则对单月和累计通用）

---

### 2. 修改：成本分析累积月份 skill — 从单月改为累计数据

**目标**：将 skill 从取附表本月数（col 1）改为取累计数（col 2），使 PPT 输出年初至今累计值而非单月值。

**核心设计决策**：不为累积版本创建独立脚本，而是在原有 `cost_analysis.py` 中通过 `cumulative` 参数区分。该参数从 `main()` → `CostAnalysisUpdater` → `ExcelDataLoader` + `ComputedValueResolver` 逐层传递：

- `ExcelDataLoader.__init__(self, excel_path, sources_config, cumulative=False)` — 控制附表取数列号
- `ComputedValueResolver.__init__(self, loader, data_month, detail_loader, cumulative=False)` — 控制 detail_sum 聚合模式
- `CostAnalysisUpdater.__init__(self, ..., cumulative=False)` — 透传标志
- `main()` 中：`CostAnalysisUpdater(..., cumulative=True)` 固化累计模式

#### `scripts/cost_analysis.py`

**类初始化变更**：

| 类 | 改动 |
|----|------|
| `ExcelDataLoader` | 新增 `cumulative=False` 参数，存入 `self.cumulative` |
| `ComputedValueResolver` | 新增 `cumulative=False` 参数，存入 `self.cumulative` |
| `CostAnalysisUpdater` | 新增 `cumulative=False` 参数，传递至 loader 和 resolver |
| `main()` | 硬编码 `cumulative=True` |

**修复函数变更**（`ExcelDataLoader` 方法，列号根据 `self.cumulative` 切换）：

| 方法 | 改动 |
|------|------|
| `_repair_medical_cost` | 所有 `_get_from_cache` 调用：`val_col = 2 if self.cumulative else 1`，影响 cur/fx/prev 取数及所有比率计算（rows 57-78 涉及 col 4-6 的比率） |
| `_repair_personnel` | `val_col = 2 if self.cumulative else 1`；上年同期附表3 的 `prev_col3 = 4 if self.cumulative else 3`、`prev_col5 = 6 if self.cumulative else 5`（用于基本工资/津贴补贴/伙食补助费/绩效工资/其他工资福利支出的取数） |
| `_repair_goods_services` | `_s3_val_col = 2 if self.cumulative else 1`，传入 `_read_s3()` 的默认 col 参数 |
| `_repair_property` | 无需改动（前端分析 sheet 列布局不变） |
| `_repair_maintenance` | 无需改动（前端分析 sheet 列布局不变） |
| `_repair_energy` | 无需改动（已是累计数列，col 2 当期、col 4+6 同期） |

**ComputedValueResolver 变更**：

| 方法 | 改动 |
|------|------|
| `resolve()` → `detail_sum` 分支 | 累计模式下：循环 `for m in range(1, month + 1)` 逐月聚合求和，而非仅取 `f"{year}-{month:02d}"` 单月。`compute_spec.get("cumulative", False)` 也支持 mapping.json 级粒度覆盖 |

#### `references/mapping.json`

**改动的 Slide**：4（收支结余）、43（医疗成本结构）、44（人员经费）、45（卫生材料费）、46（检验试剂）、47（商品服务）、49（物业管理费）、50（维修护费）。Slide 48（能源成本）和 59（预算完成）已是累计数据无需改动。

| 维度 | 具体变更 |
|------|---------|
| **附表 xl/xl_col** | 所有 `"xl": 1` → `"xl": 2`（从本月数列改为累计数列），涉及 附表1/附表2-成本/附表3/奉贤附表/上年同期附表 |
| **表头 headers** | `"2026年{month}月"` → `"2026年1-{month}月"`（Slide 4、43-50） |
| **文本 template** | 同理，模板中的月份显示从当月改为累计区间 |
| **detail_sum** | Slide 45（卫生材料费）和 Slide 46（检验试剂）的所有 `detail_sum` 计算列新增 `"cumulative": true` |
| **texts vars** | 所有 text shape 的 var 引用，附表取数列号同步更新 |

#### `SKILL.md`

| 章节 | 改动 |
|------|------|
| Frontmatter description | "数据取数规则同单月版本" → "附表 sheet 取累计数列（col 2），数据表 sheet 取收入（已是累计数据），÷10000 转万元" |
| 数据规则表 | 附表行：`是` → `是（取 col 2 累计数，非本月数 col 1）` |
| 关键业务逻辑 > Slide 4 | 医疗成本/收支结余："当月值" → "累计值"，标注 col 2 |
| 明细数据聚合 > 期间过滤 | 新增累计模式说明：汇总 1 月至当前月份逐月聚合求和 |
| Slide 45 描述 | "仅取对应月份数据" → "累计模式汇总 1 月至当前月份，单月模式仅取当月数据" |
| 与现有技能的关系 | 开头新增一行：标注本 skill 是「成本分析」的累计数据版本 |

#### Bug 修复：`_repair_assay_income` 取数列号未适配累计模式

**问题**：`_repair_assay_income` 函数用于修复附表2中"化验收入"的取数——附表2收入侧有两处"化验收入"（门急诊收入下 + 住院收入下），加载时后者覆盖前者，因此需重新读取原始 Excel 将两处值求和。但函数硬编码了 `row[1].value`（col 1 = 本月数），未根据 `self.cumulative` 标志切换为 col 2（累计数）。

**影响**：Slide 46 化验收入仅取了本月数列的部分数据（可能为住院收入的化验收入），遗漏门急诊收入的化验收入，导致数值被严重低估，成本收入比虚高（如修复前 70.9% → 修复后 36.0%）。

**修改**（`cost_analysis.py:885-924`）：
- 新增 `val_col = 2 if self.cumulative else 1`，读取、写入、日志均使用 `val_col`
- 日志增加列标签（"累计数" / "本月数"）以便区分

	#### Bug 修复：`_repair_medical_cost` 和 `_repair_personnel` 守卫条件导致取数错误

	**问题**：两个修复方法均通过 `_is_row_empty()` 检查目标行是否已有数据，若已填充则直接跳过修复。但底稿中「医疗成本」和「人员经费」sheet 被上游工具预填了**本月数**（如人员经费=19,934 万元），而非预期的**累计数**（146,042 万元）。守卫条件命中（非空），修复被跳过，脚本将本月数写入 PPT，导致 Slide 43 成本结构占比和 Slide 44 人员经费绝对值全部取错。

	**影响**：
	- Slide 43（医疗成本结构表）：人员经费占比从正确值 36.4% 被压低至 31.8%，其他科目占比同步偏移
	- Slide 44（人员经费明细）：所有行（人员经费/工资总额/基本工资等/绩效工资/社保公积金）均取值本月数，而非累计数。如人员经费医院合计 19,934 → 应为 146,042，偏差约 7 倍
	- Slide 44 文本段落：叙事中的人员经费总额、同比增幅均基于错误本月数

	**修改**（`cost_analysis.py`）：

	##### `_repair_medical_cost()`（约第 187-189 行）

	移除守卫条件，确保每次运行都从附表2累计数列重新填充：
	```python
	# 删除以下 2 行：
	# if not self._is_row_empty(cache, "人员经费", [2, 4, 6]):
	#     return
	```

	##### `_repair_personnel()`（约第 275-277 行）

	移除守卫条件，确保每次运行都从附表3累计数列重新填充：
	```python
	# 删除以下 2 行：
	# if not self._is_row_empty(cache, "人员经费", [2, 3, 4]):
	#     return
	```

	**设计理由**：`_is_row_empty` 守卫是出于性能优化考虑（避免重复计算），但牺牲了数据正确性。当底稿数据已存在但与预期模式不符（本月数 vs 累计数）时，守卫无法识别语义错误。移除后，每次运行都强制从附表源（附表2/附表3 累计数列）重新填充前端分析 sheet，确保数据正确性。性能影响可忽略（仅涉及约 15 行数据的重填）。

	**验证**：修复后重新运行，Slide 44 人员经费医院合计从 19,934 → 146,042（与 附表3 C7=1,460,419,924.94 元 / 10000 = 146,042 万元 一致）。Slide 43 人员经费占比从 31.8% → 36.4%。

---

## 六、成本分析累积月份 skill 修改（2026-07-05）

### 1. 新功能：跨行引用计算 — Slide 4「其他盈余」取数逻辑改为表内计算

**背景**：Slide 4「其他盈余」行原先从 Excel 附表1 独立取数（`一、医疗业务盈余` − `（一）医疗服务盈余`），与表中「收支结余」「医疗服务盈余」两行来源不同，存在取数路径不一致的风险。改为从 PPT 表内直接引用同行列的值做减法，确保数据一致性。

**文件变更**：

#### `scripts/cost_analysis.py`

**1.1 `ComputedValueResolver.resolve()` — 新增 `row_values` 参数 + `subtract_table_rows` 计算类型**

- 方法签名：`resolve(self, compute_spec, resolved_cols=None)` → `resolve(self, compute_spec, resolved_cols=None, row_values=None)`
- 新增 `subtract_table_rows` 分支（约第 1195 行 `subtract_refs` 之后）：
  - 参数：`a_label`（被减数行标签）、`b_label`（减数行标签）、`col`（取值列号）
  - 逻辑：`row_values[a_label][col] − row_values[b_label][col]`
  - 若 `row_values` 为 None 或任一值缺失 → 返回 None
- 设计理由：不同于已有的 `subtract_refs`（独立访问 Excel 两个数据源），`subtract_table_rows` 直接引用 PPT 表中其他行已解析的值，确保上下游数据一致性。

**1.2 `_process_slide()` — 新增 `row_values` 字典跨行共享**

- 在表处理循环前创建 `row_values = {}` 字典
- 传递给每个 `_process_table_row()` 调用
- 该字典以行标签为 key，值为 `{ppt_col: resolved_value}` 字典

**1.3 `_process_table_row()` — 接收并回存 `row_values`**

- 方法签名：`_process_table_row(self, table, row_cfg, slide_cfg)` → `_process_table_row(self, table, row_cfg, slide_cfg, row_values=None)`
- Pass 2 解析计算列时，将 `row_values` 传入 `self.resolver.resolve()`
- 在行解析完成后：`row_values[label] = resolved` 存入结果供后续行引用
- 行按 mapping.json 的 `rows` 数组顺序处理，先处理的行对后处理的行可见

#### `references/mapping.json`

**1.4 Slide 4「其他盈余」cols 改为 `subtract_table_rows`**

修改前（独立取数）：
```json
{ "ppt": 1, "compute": { "type": "subtract_refs",
  "a": { "source": "当月附表1", "xl_label": "一、医疗业务盈余", "xl_col": 2 },
  "b": { "source": "当月附表1", "xl_label": "（一）医疗服务盈余", "xl_col": 2 }
} }
```
（col 2、3 同理分别指向 当月奉贤附表1 / 上年同期附表1）

修改后（表内引用）：
```json
{ "ppt": 1, "compute": { "type": "subtract_table_rows",
  "a_label": "收支结余", "b_label": "医疗服务盈余", "col": 1 } }
{ "ppt": 2, "compute": { "type": "subtract_table_rows",
  "a_label": "收支结余", "b_label": "医疗服务盈余", "col": 2 } }
{ "ppt": 3, "compute": { "type": "subtract_table_rows",
  "a_label": "收支结余", "b_label": "医疗服务盈余", "col": 3 } }
{ "ppt": 4, "compute": { "type": "subtract", "a_col": 1, "b_col": 3 } }
```

- col 4（增减）保持不变：`subtract` a_col=1, b_col=3（同行内计算）
- `desc` 字段更新为 `"收支结余 - 医疗服务盈余，从表内同行取值计算"`

**1.5 行顺序调换**：`收支结余` 行移至 `其他盈余` 行之前

原因：`subtract_table_rows` 依赖被引用行先完成解析。原顺序为「医疗服务盈余 → 其他盈余 → 收支结余」，`其他盈余` 处理时 `收支结余` 尚未解析，`row_values["收支结余"]` 不存在，所有列返回 None。调换后处理顺序为「医疗收入 → 减：医疗成本 → 医疗服务盈余 → 收支结余 → 其他盈余」，依赖关系满足。

影响：`find_row_by_label()` 按标签搜索 PPT 表格中的实际行位置，不受 mapping.json 中 rows 数组顺序影响，因此调换不影响写入位置。

### 执行记录

| 步骤 | 执行的 skill | 输入文件 | 输出文件 |
|------|-------------|----------|----------|
| 1 | 成本分析累积月份 | `1-6月分析底稿_updated_0704.xlsx` / `新华医院2026年1-6月运营分析_updated.pptx` | `新华医院2026年1-6月运营分析_updated_updated_0705.pptx` |

**验证结果**：Slide 4「其他盈余」4 列全部更新成功：
- Col 1（医院合计）：2,934 → 16,361
- Col 2（奉贤院区）：-176 → 343
- Col 3（上年同期）：9,417 → 24,532
- Col 4（增减）：-6,483 → -8,171（= 16,361 − 24,532 ✓）

---

## 七、成本分析 skill（单月版）同步修改（2026-07-05）

### 1. 新功能：跨行引用计算 — Slide 4「其他盈余」取数逻辑改为表内计算

**背景**：与第六节「成本分析累积月份」的改动完全对应。将单月版 skill 的 Slide 4「其他盈余」也从 Excel 附表1 独立取数改为 PPT 表内引用，确保与累积版行为一致。

**文件变更**：

#### `scripts/cost_analysis.py`（3 处，与累积版相同）

- **`ComputedValueResolver.resolve()`**：新增 `row_values=None` 参数 + `subtract_table_rows` 计算类型
- **`_process_slide()`**：新增 `row_values = {}` 字典跨行共享
- **`_process_table_row()`**：接收 `row_values` 参数、传入 resolver、解析完成后存入结果

#### `references/mapping.json`（2 处，与累积版相同）

- **Slide 4「其他盈余」cols**：独立 `subtract_refs`（xl_col=1 本月数）→ 表内 `subtract_table_rows`
- **行顺序调换**：`收支结余` 行移至 `其他盈余` 之前

### 执行记录

| 步骤 | 执行的 skill | 输入文件 | 输出文件 |
|------|-------------|----------|----------|
| 1 | 成本分析 | `6月分析底稿_updated_0704.xlsx` / `新华医院2026年6月运营分析底版.pptx` | `新华医院2026年6月运营分析底版_updated_0705.pptx` |

**验证结果**：Slide 4「其他盈余」4 列全部更新成功：
- Col 1（医院合计）：69 → 250
- Col 2（奉贤院区）：-31 → -50
- Col 3（上年同期）：15 → 3,687
- Col 4（增减）：54 → -3,437（= 250 − 3,687 ✓）
