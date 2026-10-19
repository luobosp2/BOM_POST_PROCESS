# BOM Post Process

对 Altium Designer 23 导出的 `.xlsx` BOM 进行封装校验、参数解析，在人工完成选型说明后分别生成焊接表、采购 BOM Excel 和标准化 Word BOM。

## 工作流程

```text
AD 原始 BOM
    → validate 检查封装、位号和数量，提醒同料号不同 Footprint 的条目合并
    → gen_value 复制原工作簿，追加 name、note、not_bought、not_welded
    → 人工修改名称、选型备注及采购/焊接标记并保存
        ├─ gen_soldering → 焊接表 Excel（只看 not_welded）
        └─ gen_bom → 采购 BOM Excel（只看 not_bought）
                        → gen_bom_standard → 标准化 Word、PDF
```

**人工调整是独立步骤，因此不提供 `all` 命令。** 两种 Excel 导出直接读取同一个人工修改后的参数表，保留名称和备注；Word 再读取采购 BOM。导出阶段不重新解析参数，也不重新判断 NC。再次执行 `gen_value` 时，已有 `name` 和两个标记列保留人工内容（包括空白），已有非空 `note` 原样保留，只给空白备注补充可解析的参数。

## 安装

需要 Python 3.10 或以上版本。Excel 处理只依赖 `openpyxl`，不需要安装 Microsoft Excel。**Word 导出需要 Windows 和已安装、激活的桌面版 Microsoft Word**，通过实际排版核验页数和表格位置。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e .
```

不能激活虚拟环境时，可直接使用 `.\.venv\Scripts\python.exe -m bom_post_process ...`。`bom-post-process` 和 `python -m bom_post_process` 是等价入口。

## 使用样例

```powershell
# 1. 校验，保存完整报告
bom-post-process validate input/DAQ_1010A_V1R0.xlsx -o output/校验.json

# 2. 解析参数，输出 output/DAQ_1010A_V1R0_已解析.xlsx
bom-post-process gen_value input/DAQ_1010A_V1R0.xlsx

# 3. 用 Excel 打开已解析文件，修改 name、note、not_bought、not_welded，保存并关闭

# 4. 导出焊接表；人工备注会原样保留
bom-post-process gen_soldering output/DAQ_1010A_V1R0_已解析.xlsx

# 5. 从同一份人工修改后的文件导出采购 BOM
bom-post-process gen_bom output/DAQ_1010A_V1R0_已解析.xlsx

# 6. 从采购 BOM 导出 Word，同时生成 PDF 和分页核验 JSON
# 可先在 BOM Excel 中补充 PCB 占位行的数量等信息
bom-post-process gen_bom_standard output/DAQ_1010A_V1R0_BOM.xlsx
```

所有子命令支持 `-o/--output` 指定输出文件。`validate` 默认只显示终端报告；其余命令默认写入当前目录的 `output/`。默认文件名依次为 `_已解析.xlsx`、`_焊接表.xlsx`、`_BOM.xlsx`、`_标准化.docx`。

所有子命令支持 `--sheet 工作表名`：`validate`、`gen_value` 选择 AD 表；`gen_soldering`、`gen_bom` 选择含四个人工编辑列的 AD 表；`gen_bom_standard` 选择导出的六列表，默认优先 `BOM` 工作表，也支持自定义采购表名。标准化命令拒绝“焊接表”和“未焊接元件”，避免用焊接范围代替采购范围。

输入文件不会修改，已有输出不会覆盖。重复执行时用 `-o` 换输出文件名。Word 的 DOCX、PDF、核验 JSON 仅在内容和分页检查通过后发布。

退出码：`0` 表示命令完成且无校验错误（可能有未覆盖提示）；`1` 表示校验错误或处理失败；`2` 表示命令行参数错误。

## 输入与转换规则

原始表第一行为以下列名，列可重新排列，识别时忽略表头大小写和首尾空格：

```text
Comment | Description | Designator | Footprint | LibRef | Quantity
```

自动选择唯一具有这些表头的工作表；存在多张时用 `--sheet` 指定。`input/DAQ_1010A_V1R0.xlsx` 就是单张原始表输入。

`gen_value` 复制整个工作簿，在选中的原表末尾依次增加小写 `name`、`note`、`not_bought`、`not_welded` 四列。原 AD 表头、列顺序、单元格内容、字体、边框、行高、已有列宽、筛选、打印设置和其他工作表均保留；不新增中文参数表，不合并或删除任何条目。新列沿用原表文本单元格样式。

以样例为例，输出表头为：

```text
Comment | Description | Designator | Footprint | LibRef | Quantity | name | note | not_bought | not_welded
```

| 列 | 自动填写与人工用途 |
| --- | --- |
| `name` | 预填标准电阻、电容等可识别元件的名称；其他元件使用 LibRef 作为初值，供人工调整。导出 Excel 和 Word 的“名称”直接取此列。 |
| `note` | 填入标准电阻、电容参数；其他器件或无法解析的厂商编码留空，供人工填写选型说明。 |
| `not_bought` | 初始留空；不需要采购时填写任意文字。只影响采购 BOM。 |
| `not_welded` | Comment 末尾为 `/NC` 或 `(NC)`（忽略大小写）的条目预填 `NC`，其余留空。只影响焊接表，可手动清空或更改。 |

`Description` 仍在原列，不自动复制进备注。已有完整四列时直接复用；对于上一版本末尾只有 `note` 的文件，会在保留备注的基础上补齐四列。涉及复杂引用、合并等无法安全迁移的旧文件会明确报错，要求先整理列或从原始表重新解析。

`gen_soldering` 和 `gen_bom` 在导出阶段才转换为 `序号、规格型号/代号、名称、数量、项目代号、备注` 六列：Comment 对应规格型号，`name` 对应名称，Quantity 对应数量，Designator 对应项目代号，`note` 原样对应备注。数据保持源表顺序，导出序号连续生成；名称即使留空也不会被 LibRef 覆盖。

- `gen_soldering` 只排除 `not_welded` 非空的行，不检查 `not_bought`；输出仅含“焊接表”，不再输出未焊接条目的副表。
- `gen_bom` 只排除 `not_bought` 非空的行，不检查 `not_welded`；输出仅含 `BOM`。不焊接但要采购的元件仍可进入采购表。
- 两列中的任意非空值都表示排除，包括 `0`、`False`、`否`；只有清空单元格或仅空白字符才表示不排除。
- mark 点通常是 PCB 图形，不需要单独采购或焊接；测试点可能是焊盘，也可能是独立器件，应按实际设计人工设置两列，不根据名称自动剔除。
- 采购 BOM 首条补充 `PCB制板` 占位，数量留空、备注为 `PCB制板的占位符`。若参数表已存在 Comment 或 name 为 `PCB制板` 的行，就使用人工数据，不重复补行；该行被 `not_bought` 排除时也不会补回。
- Word 从采购 BOM 逐行读取名称、备注等内容；PCB 制板对应“外协件”分类，普通元件对应“外购件”分类。PCB 占位与分类标题同时保留。

参与解析的 AD 六列中的公式和 Excel 错误值会报错；其他列及其他工作表中的公式保持原样。Excel 导出读取的名称、备注、物料字段和本命令的标记列不接受公式或 Excel 错误，另一个标记列不参与判断；待筛选表不能含合并单元格。Word 阶段允许公式已有有效缓存值；无缓存值或错误值会停止导出。Word 读取单元格原始值，不应用 Excel 数字显示格式；需要保留前导零的编号或料号请存成文本。

## 校验与参数范围

`validate` 对符合本项目命名规则的电阻、电容比较 Comment 和 Footprint 中的封装尺寸及 C/R 器件类型，支持 `C0402`、`C 0402L`、`CAP_0402`、`0402C`、`R0402`、`RES-0402`、`0402R`、`0402` 等别名；还检查空位号、重复位号、数量与位号个数是否一致。

Comment 命名与封装一致性规则仅用于已识别的电阻、电容（包含极性电容）。芯片、连接器、LED、测试点、定位点等其他器件跳过该规则，不再因料号不符合电阻电容命名而产生 warning；所有条目仍检查位号、数量及必填的 Comment、Footprint。对已识别的电阻电容，只在完整匹配料号规则时比较封装，无法解析的厂商型号或未知封装别名仍提示未验证，不把 `M1206` 等芯片料号中的数字当成封装。

报告中 `checked` 是已比较行数，`skipped` 是未比较行数，包含正常跳过的其他器件及无法比较的电阻电容；并不是每个 skipped 条目都有 warning。`ok=true` 仅表示未发现错误，不表示所有元件均完成验证。`validate` 检查 AD 列，不验证人工填写的 note 中的选型说明。

`validate` 还会跨行检查电阻、电容：同一个 Comment 对应多个不同的 Footprint 时，每组给出一条合并提示（`same_comment_different_footprints`），列出相关行号、位号和原始 Footprint，供人工核对封装并合并 BOM 条目。元件身份通过现有料号规则或 LibRef 识别，包含极性电容及能由 LibRef 识别的厂商型号，不要求参数解析成功。

比较 Comment 和 Footprint 时忽略大小写和首尾空格，保留内部空格及 NC 后缀；例如 `R0402` 与 `0402R` 仍算不同 Footprint，带 `/NC` 的 Comment 不与普通 Comment 合为一组。空 Comment 或空 Footprint 由原有缺失检查处理，不参与分组。终端完整显示所有合并提示，其他普通提示最多显示 8 条；`validate -o` 的 JSON 保存完整报告。

合并提示属于 warning，不改变 `checked/skipped` 的含义；仅有提示时退出码仍为 `0`。校验不修改 Excel，也不会自动合并或阻止后续解析、焊接表、采购 BOM 和 Word 导出。样例 `input/DAQ_1010A_V1R0.xlsx` 会报告 5 条电容型号未验证提示和 6 组合并提示，总计 11 条提示、0 个错误；封装比较仍为 91 行已比较、51 行未比较，其中 46 行其他器件正常跳过。

| 类型 | 示例 Comment | 备注输出 |
| --- | --- | --- |
| 电容 | `0402-X5R104K250` | `100nF/25V/0402/X5R/10%` |
| 电容 | `0402-C0G4R7J500` | `4.7pF/50V/0402/C0G/5%` |
| 电容 | `0603-X5R476M6R3` | `47uF/6.3V/0603/X5R/20%` |
| 电阻 | `0402K1782` | `17.8k/0402/1%` |
| 电阻 | `0402K0R/NC` | `0/0402/1%` |

电容支持 X5R、X7R、C0G（COG、NP0 统一为 C0G），容量三位编码按 pF 解释，`R` 表示小数点；耐压三位编码除以 10，`6R3` 表示 6.3 V；J/K/M 精度分别为 5%/10%/20%。电阻四位编码按前三位有效数字和末位倍率解释，支持 `R` 小数与尾部 `R`。**电阻 1% 沿用旧项目约定，K/L/U/M 不作为精度码解释**；料号未给出的电阻耐压、功率不会填造。

## Word 模板与配置

默认使用包内 `resources/bom_template.docx` 和 `resources/bom_config.json`，安装后不依赖旧目录。模板保留首页签署栏、续页表尾、分类行以及“共 X 张 / 第 Y 张”。Word 实测行高，整条物料不跨页；分类标题与第一条物料保持同页。最终回读 Word 逐字段核对数据，并验证每页表格的物理页码。

可以复制配置文件填写，再通过 `gen_bom_standard --config 路径` 传入：

```json
{
  "PCB名称": "",
  "PCB订单号": "",
  "公司名": ""
}
```

空字符串保留模板占位文字。自定义模板可用 `--template` 指定，需保持原来的两张表、九个物理单元格、分类行和 `PCB订单号` 表尾标记。

## 项目结构与测试

```text
src/bom_post_process/
  cli.py          # 独立子命令
  components.py   # 电阻、电容命名规则与参数解析
  workbook.py     # Excel 读取、表头识别、共用 I/O
  validation.py   # 封装、位号和数量检查
  values.py       # 原表追加四列，填写名称、标准参数和 NC 标记
  xlsx_notes.py   # 局部更新工作表 XML，保留原工作簿格式
  reviewed_export.py # 读取人工修改结果，按指定标记筛选并映射六列
  soldering.py    # 按 not_welded 导出焊接表
  bom.py          # 按 not_bought 导出采购 BOM，补 PCB 占位
  standard.py     # Word 模板、分页与数据校验
  resources/      # 内置 Word 模板、配置、PowerShell worker
tests/
input/
output/           # 生成结果，Git 忽略
```

```powershell
python -m unittest discover -s tests -v
```

自动测试覆盖编码边界、封装冲突、四列保真与旧 note 迁移、NC 预填、两种标记独立筛选、人工名称和备注保留、PCB 占位、文件覆盖保护、Word 模板与分页异常、CLI 完整流程。单元测试不要求启动 Word；实际分页需在安装 Word 的 Windows 上执行导出命令。

旧代码目录已移除。运行所需的模板、配置和 Word 排版脚本均包含在包内资源中。
