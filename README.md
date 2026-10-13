# BOM Post Process

对 Altium Designer 23 导出的 BOM 进行检查、参数解析和标准化 Word 导出。

计划提供统一的 Python CLI：

- `validate`：检查 Comment 与 Footprint 是否冲突。
- `gen_value`：解析电阻、电容参数，填写备注并生成焊接表。
- `gen_bom_standard`：从焊接表生成标准化 Word BOM。
- `all`：按顺序执行全部处理。

`input/DAQ_1010A_V1R0.xlsx` 是输入样例；生成文件放在 `output/`。

旧脚本目录保留在本地用于迁移，不纳入 Git 版本管理。
