# 工具域与 Profile 参考

**文档版本**：`V2.4`\
**更新日期**：`2026-03-30`

***

## 1. 工具域总览

| 域 | 工具数 | 说明 |
|---|---:|------|
| server | 2 | 健康检查、状态查询 |
| workbook | 12 | 生命周期管理 |
| names | 4 | 命名范围 |
| sheet | 12 | 工作表操作 |
| range | 13 | 单元格操作 |
| formula | 4 | 公式操作 |
| format | 7 | 格式与样式 |
| vba | 8 | VBA 自动化 |
| recovery | 8 | 恢复与快照 |
| table | 8 | 表格对象管理 |
| analysis | 6 | 分析与扫描 |
| pq | 5 | Power Query |
| workbook_ops | 6 | 工作簿治理与导出 |
| chart/pivot/model | 0 | 待实现 |

**总计：约 75+ 工具**

***

## 2. Bundle 定义

### 2.1 Bundle 列表

| Bundle | 工具数 | 依赖 | 成熟度 | 说明 |
|--------|---:|------|------|------|
| foundation | 8 | - | stable | server + workbook 基础 |
| data | 8 | foundation | stable | table 表格管理 |
| analysis | 6 | foundation | stable | 分析工具集 |
| workbook_ops | 6 | foundation | experimental | 工作簿治理与导出 |
| edit_basic | 7 | foundation | stable | range/sheet 基础写操作 |
| edit_structure | 6 | foundation | stable | sheet.copy/move/hide/unhide + range.find_replace/autofit |
| calc_format | 11 | foundation | stable | formula + format |
| automation | 8 | foundation | stable | vba |
| recovery | 8 | foundation | stable | snapshot/backup/undo/audit |
| report | 6 | foundation | stable | chart/pivot/model（占位） |
| vba_first | ~28 | - | - | foundation + automation + recovery（调试用） |

### 2.2 Bundle 工具详情

**foundation（8工具）**
- server.health, server.status
- workbook.open_file, workbook.create_file, workbook.save_file, workbook.close_file, workbook.list_open, workbook.inspect

**data（8工具）**
- table.list_tables, table.create, table.inspect, table.resize, table.rename, table.set_style, table.toggle_total_row, table.delete

**analysis（6工具）**
- analysis.scan_structure, analysis.scan_formulas, analysis.scan_links, analysis.scan_hidden, analysis.export_report
- audit.list_operations

**workbook_ops（6工具）**
- workbook.save_as, workbook.refresh_all, workbook.calculate, workbook.list_links, workbook.export_pdf, sheet.export_csv

**edit_structure（6工具）**
- sheet.copy, sheet.move, sheet.hide, sheet.unhide
- range.find_replace, range.autofit

**calc_format（11工具）**
- formula.fill, formula.set_single, formula.get_dependencies, formula.repair
- format.set_style, format.auto_fit

***

## 3. Profile 定义

### 3.1 Profile 总览

| Profile | 工具数 | tool_budget | 说明 |
|---------|---:|---:|------|
| basic_edit | 35 | 40 | 日常编辑（推荐新手） |
| calc_format | 46 | 50 | 公式与格式处理 |
| automation | 40 | 45 | VBA 自动化与恢复 |
| data_workflow | 33 | 39 | Power Query / 数据流（**Trae AI 推荐**） |
| reporting | 32 | 37 | 报表与分析 |
| vba_first | ~28 | 30 | 调试专用 |
| all | 75+ | - | 全量开发调试 |

### 3.2 Profile 域覆盖矩阵

| 域 | basic_edit | calc_format | automation | data_workflow | reporting |
|---|:---:|:---:|:---:|:---:|:---:|
| server (2) | ✅ | ✅ | ✅ | ✅ | ✅ |
| workbook (12) | ✅ | ✅ | ✅ | ✅ | ✅ |
| names (4) | ✅ | ✅ | ✅ | ✅ | ✅ |
| sheet (12) | ✅ | ✅ | ❌ | ❌ | ❌ |
| range (13) | ✅ | ✅ | ❌ | ❌ | ❌ |
| formula (4) | ❌ | ✅ | ❌ | ❌ | ❌ |
| format (7) | ❌ | ✅ | ❌ | ❌ | ❌ |
| vba (8) | ❌ | ❌ | ✅ | ❌ | ❌ |
| recovery (8) | ❌ | ❌ | ✅ | ❌ | ❌ |
| analysis (6) | ❌ | ❌ | ❌ | ✅ | ✅ |
| pq (5) | ❌ | ❌ | ❌ | ✅ | ❌ |
| table (8) | ❌ | ❌ | ❌ | ✅ | ❌ |
| workbook_ops (6) | ❌ | ❌ | ❌ | ✅ | ❌ |

***

## 4. 场景与 Profile 推荐

### 4.1 场景推荐矩阵

| 场景 | 推荐 Profile | 工具数 | 说明 |
|------|------------|---:|------|
| 日常表格编辑 | basic_edit | 35 | 覆盖基础读写，不需要 PQ/VBA |
| 公式与样式处理 | calc_format | 46 | 包含 formula/format 域 |
| VBA 自动化开发 | automation | 40 | 包含 vba/recovery |
| AI 数据分析（Trae） | data_workflow | 33 | 低于 39 限制，推荐 AI 使用 |
| 报表生成 | reporting | 32 | analysis + report |
| 调试测试 | vba_first | ~28 | VBA 优先，工具少便于调试 |
| 全量功能测试 | all | 75+ | 仅用于开发调试 |

### 4.2 Trae AI 客户端配置

Trae AI 存在 39 工具截断限制。推荐配置：

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": [
        "run",
        "python",
        "-m",
        "excelforge.gateway.host",
        "--profile",
        "data_workflow"
      ],
      "cwd": "D:/Tools/AI/ExcelForge"
    }
  }
}
```

如需额外工具，可通过 `--enable-bundle` 添加：

```json
{
  "mcpServers": {
    "excel": {
      "command": "uv",
      "args": [
        "run",
        "python",
        "-m",
        "excelforge.gateway.host",
        "--profile",
        "data_workflow",
        "--enable-bundle",
        "edit_structure"
      ],
      "cwd": "D:/Tools/AI/ExcelForge"
    }
  }
}
```

**注意**：总工具数需保持在 39 以下。

***

## 5. 新增工具参考（V2.4）

### 5.1 table 域（8工具）

| 工具 | 说明 |
|------|------|
| table.list_tables | 列出工作表中的表格 |
| table.create | 将区域转换为表格 |
| table.inspect | 检查表格结构 |
| table.resize | 调整表格范围 |
| table.rename | 重命名表格 |
| table.set_style | 设置表格样式 |
| table.toggle_total_row | 开关总计行 |
| table.delete | 删除表格（保留数据） |

### 5.2 analysis 域（6工具）

| 工具 | 说明 |
|------|------|
| analysis.scan_structure | 扫描工作簿结构 |
| analysis.scan_formulas | 扫描公式分布 |
| analysis.scan_links | 扫描外部链接 |
| analysis.scan_hidden | 扫描隐藏元素 |
| analysis.export_report | 生成分析报告 |
| audit.list_operations | 列出操作审计记录 |

### 5.3 workbook_ops 域（6工具）

| 工具 | 说明 |
|------|------|
| workbook.save_as | 另存工作簿 |
| workbook.refresh_all | 刷新所有数据 |
| workbook.calculate | 重新计算 |
| workbook.list_links | 列出外部链接 |
| workbook.export_pdf | 导出 PDF |
| sheet.export_csv | 导出 CSV |

### 5.4 sheet 扩充（4工具）

| 工具 | 说明 |
|------|------|
| sheet.copy | 复制工作表 |
| sheet.move | 移动工作表 |
| sheet.hide | 隐藏工作表 |
| sheet.unhide | 取消隐藏 |

### 5.5 range 扩充（2工具）

| 工具 | 说明 |
|------|------|
| range.find_replace | 查找替换 |
| range.autofit | 自动调整列/行宽 |
