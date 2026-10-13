# BOM Post Process

对 Altium Designer 23 导出的 `.xlsx` BOM 进行封装校验、参数解析，在人工完成选型说明后生成焊接表和标准化 Word BOM。

## 工作流程

```text
AD 原始 BOM
    → validate 检查封装、位号和数量
    → gen_value 复制原工作簿，在原表末尾追加 note 列
    → 人工编辑 note，填写元件选型和备注并保存
    → gen_soldering 分离焊接与未焊接元件
    → gen_bom_standard 从焊接表导出 Word
```

**人工调整是独立步骤，因此不提供 `all` 命令。** `gen_soldering` 和 `gen_bom_standard` 直接使用已编辑的备注，不会重新解析覆盖。再次执行 `gen_value` 时，已有非空 `note` 原样保留，只给空白项补充可解析的参数；需要重新解析某条备注时，先清空该格。

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

# 3. 用 Excel 打开已解析文件，在原工作表最后的 note 列填写选型备注，保存并关闭

# 4. 导出焊接表；人工备注会原样保留
bom-post-process gen_soldering output/DAQ_1010A_V1R0_已解析.xlsx

# 5. 从焊接表导出 Word，同时生成 PDF 和分页核验 JSON
bom-post-process gen_bom_standard output/DAQ_1010A_V1R0_焊接表.xlsx
```

所有子命令支持 `-o/--output` 指定输出文件。`validate` 默认只显示终端报告；其余命令默认写入当前目录的 `output/`。默认文件名依次为 `_已解析.xlsx`、`_焊接表.xlsx`、`_标准化.docx`。

所有子命令支持 `--sheet 工作表名`：`validate`、`gen_value` 选择 AD 表，`gen_soldering` 选择已编辑 `note` 的 AD 表（也兼容旧中文六列表），`gen_bom_standard` 选择导出的六列表（默认优先“焊接表”）。`gen_soldering` 默认用所选 AD 表的 Comment 判断 NC；需要从另一张保留的原始表判断时，用 `--ad-sheet` 指定。

输入文件不会修改，已有输出不会覆盖。重复执行时用 `-o` 换输出文件名。Word 的 DOCX、PDF、核验 JSON 仅在内容和分页检查通过后发布。

退出码：`0` 表示命令完成且无校验错误（可能有未覆盖提示）；`1` 表示校验错误或处理失败；`2` 表示命令行参数错误。

## 输入与转换规则

原始表第一行为以下列名，列可重新排列，识别时忽略表头大小写和首尾空格：

```text
Comment | Description | Designator | Footprint | LibRef | Quantity
```

自动选择唯一具有这些表头的工作表；存在多张时用 `--sheet` 指定。`input/DAQ_1010A_V1R0.xlsx` 就是单张原始表输入。

`gen_value` 复制整个工作簿，只在选中的原表末尾增加小写 `note` 列。原表头、列顺序、单元格内容、字体、边框、行高、已有列宽、筛选、打印设置和其他工作表均保留；不新增中文参数表，不合并或删除任何条目。新增的 `note` 列沿用原表文本单元格样式，列宽为 42；原有列宽不变。

以样例为例，输出表头为：

```text
Comment | Description | Designator | Footprint | LibRef | Quantity | note
```

标准电阻、电容的参数填入 `note`；其他器件或无法解析的厂商编码留空，供人工填写。`Description` 仍在原列，不自动复制进备注。已有 `note` 列时直接复用，不重复追加，也不覆盖人工内容。

`gen_soldering` 在导出阶段才转换为 `序号、规格型号/代号、名称、数量、项目代号、备注` 六列：Comment 对应规格型号，Quantity 对应数量，Designator 对应项目代号，`note` 原样对应备注；名称按支持规则和 LibRef 映射。数据保持源表顺序，序号按源记录位置生成，不重新计算参数。

- Comment 末尾 `/NC` 或 `(NC)` 表示不贴装，忽略大小写。通过位号匹配，把相关条目放入“未焊接元件”，其他条目放入“焊接表”。默认以当前已解析表为准，编辑 note 时保留 Comment 中的 NC 标记。
- 指定独立 `--ad-sheet` 时，NC 来源为该原始表；一行混有贴装和不贴装位号时要求先拆分。
- “规格型号/代号”或“名称”为 `PCB制板` 的占位项不输出。
- 兼容旧版已整理的六中文列表：直接复制人工名称、原序号、数量、位号、备注及格式，NC 从保留的 AD 原始表判断。
- 默认 Word 只读取“焊接表”，不把“未焊接元件”一并导出。

参与解析的 AD 六列中的公式和 Excel 错误值会报错；其他列及其他工作表中的公式保持原样。焊接导出的 note 不接受公式或 Excel 错误，待筛选表不能含合并单元格。Word 阶段允许公式已有有效缓存值；无缓存值或错误值会停止导出。Word 读取单元格原始值，不应用 Excel 数字显示格式；需要保留前导零的编号或料号请存成文本。

## 校验与参数范围

`validate` 对符合本项目命名规则的电阻、电容比较 Comment 和 Footprint 中的封装尺寸及 C/R 器件类型，支持 `C0402`、`C 0402L`、`CAP_0402`、`0402C`、`R0402`、`RES-0402`、`0402R`、`0402` 等别名；还检查空位号、重复位号、数量与位号个数是否一致。

只对完整匹配的料号进行封装比较，不把 `M1206` 等芯片料号中的数字当成封装。未知料号或封装别名会标记为未验证。报告中 `checked` 是已比较行数，`skipped` 是未比较行数；`ok=true` 仅表示未发现错误，不表示所有元件均完成验证。`validate` 检查 AD 列，不验证人工填写的 note 中的选型说明。

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
  values.py       # 原表末尾追加 note，填写标准参数
  xlsx_notes.py   # 局部更新工作表 XML，保留原工作簿格式
  soldering.py    # 保留人工内容，分离焊接/未焊接表
  standard.py     # Word 模板、分页与数据校验
  resources/      # 内置 Word 模板、配置、PowerShell worker
tests/
input/
output/           # 生成结果，Git 忽略
```

```powershell
python -m unittest discover -s tests -v
```

自动测试覆盖编码边界、封装冲突、NC 分离、人工备注保留、文件覆盖保护、Word 模板与分页异常、CLI 接口。单元测试不要求启动 Word；实际分页需在安装 Word 的 Windows 上执行导出命令。

旧代码目录已移除。运行所需的模板、配置和 Word 排版脚本均包含在包内资源中。
