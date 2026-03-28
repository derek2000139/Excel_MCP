# ExcelForge

**ExcelForge** 是一个基于 MCP (Model Context Protocol) 的 Excel 操作工具集，让 AI 助手能够安全、高效地操作 Excel 文件。

## v2.2 新特性

- **统一入口**：单一 `excel-mcp` 入口替代多个独立 Gateway
- **Profile 机制**：按需加载工具集，减少资源占用
- **Bundle 装配**：灵活组合功能模块
- **Runtime 单例**：解决多进程隔离问题，所有操作共享同一 Excel 实例

## 快速开始

### 1. 安装

```bash
git clone https://github.com/your-repo/ExcelForge.git
cd ExcelForge
uv sync
```

### 2. 配置 MCP 客户端

在 Trae IDE 或其他 MCP 客户端中配置：

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "excelforge.gateway.host", "--config", "excel-mcp.yaml", "--profile", "automation"],
      "cwd": "你的ExcelForge项目路径"
    }
  }
}
```

### 3. 可选 Profile

| Profile | 工具数 | 说明 |
|---------|-------:|------|
| `basic_edit` | 25 | 基础编辑（默认） |
| `automation` | 18 | VBA + 恢复 |
| `calc_format` | 40 | 公式与格式 |
| `data_workflow` | 18 | PQ/Table/分析 |
| `reporting` | 18 | 图表/透视/模型 |
| `all` | 全部 | 开发调试用 |

## 配置文件说明

### 用户必需的配置文件

| 文件 | 说明 | 必需 |
|------|------|:----:|
| `excel-mcp.yaml` | 统一入口配置 | ✅ |
| `runtime-config.yaml` | Runtime 配置 | ✅ |
| `excelforge/gateway/profiles.yaml` | Profile 定义 | ✅ |
| `excelforge/gateway/bundles.yaml` | Bundle 定义 | ✅ |

### 兼容旧入口的配置文件（可选）

| 文件 | 说明 | 必需 |
|------|------|:----:|
| `excel-core-mcp.yaml` | Core 入口配置 | ❌ |
| `excel-vba-mcp.yaml` | VBA 入口配置 | ❌ |
| `excel-recovery-mcp.yaml` | Recovery 入口配置 | ❌ |
| `excel-pq-mcp.yaml` | PQ 入口配置 | ❌ |

> **推荐**：使用统一入口 `excel-mcp.yaml`，无需配置多个入口。

## 架构

```
┌─────────────────────────────────────────────────────────────┐
│                     MCP 客户端 (Trae IDE)                    │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                  excel-mcp (统一入口)                        │
│                  --profile automation                       │
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                  Runtime (单例进程)                          │
│                  Excel COM 对象                              │
└─────────────────────────────────────────────────────────────┘
```

## 功能特性

- **工作簿管理** - 打开、保存、关闭、创建 Excel 文件
- **工作表操作** - 创建、删除、重命名、查看结构、自动筛选
- **范围操作** - 读写数据、复制、清除、排序、合并单元格
- **公式支持** - 验证表达式、填充范围、获取依赖关系
- **格式设置** - 设置单元格样式、自动调整列宽
- **VBA 读写访问** - 查看工程、扫描代码、同步模块、执行宏
- **备份恢复** - 文件级备份、快照回滚
- **命名范围** - 列出和读取命名区域
- **数据验证和条件格式** - 读取工作表规则
- **审计日志** - 操作记录追踪

## MCP 工具列表

### 基础工具 (foundation bundle)

| 类别 | 工具 | 说明 |
|------|------|------|
| 服务器 | `server.get_status` | 获取服务状态 |
| 服务器 | `server.health` | 健康检查 |
| 工作簿 | `workbook.open_file` | 打开 Excel 文件 |
| 工作簿 | `workbook.create_file` | 创建新工作簿 |
| 工作簿 | `workbook.save_file` | 保存工作簿 |
| 工作簿 | `workbook.close_file` | 关闭工作簿 |
| 工作簿 | `workbook.inspect` | 查看工作簿信息 |
| 命名范围 | `names.inspect` | 查看命名范围 |
| 命名范围 | `names.manage` | 管理命名范围 |

### 编辑工具 (edit bundle)

| 类别 | 工具 | 说明 |
|------|------|------|
| 工作表 | `sheet.create_sheet` | 创建工作表 |
| 工作表 | `sheet.rename_sheet` | 重命名工作表 |
| 工作表 | `sheet.delete_sheet` | 删除工作表 |
| 工作表 | `sheet.set_auto_filter` | 设置自动筛选 |
| 工作表 | `sheet.get_conditional_formats` | 获取条件格式 |
| 工作表 | `sheet.get_data_validations` | 获取数据验证 |
| 范围 | `range.read_values` | 读取单元格值 |
| 范围 | `range.write_values` | 写入单元格值 |
| 范围 | `range.clear_contents` | 清除单元格内容 |
| 范围 | `range.copy` | 复制范围 |
| 范围 | `range.insert_rows` | 插入行 |
| 范围 | `range.delete_rows` | 删除行 |
| 范围 | `range.insert_columns` | 插入列 |
| 范围 | `range.delete_columns` | 删除列 |
| 范围 | `range.sort_data` | 排序数据 |
| 范围 | `range.merge` | 合并单元格 |

### 计算与格式工具 (calc_format bundle)

| 类别 | 工具 | 说明 |
|------|------|------|
| 公式 | `formula.evaluate` | 计算公式 |
| 格式 | `format.set_number_format` | 设置数字格式 |
| 格式 | `format.set_font` | 设置字体 |
| 格式 | `format.set_fill` | 设置填充 |
| 格式 | `format.set_border` | 设置边框 |
| 格式 | `format.set_alignment` | 设置对齐 |
| 格式 | `format.set_column_width` | 设置列宽 |
| 格式 | `format.set_row_height` | 设置行高 |

### VBA 工具 (automation bundle)

| 工具 | 说明 |
|------|------|
| `vba.inspect_project` | 查看 VBA 工程 |
| `vba.scan_code` | 扫描 VBA 代码 |
| `vba.sync_module` | 同步 VBA 模块 |
| `vba.remove_module` | 删除 VBA 模块 |
| `vba.execute` | 执行 VBA 宏 |
| `vba.compile` | 编译 VBA 工程 |

### 恢复工具 (recovery bundle)

| 工具 | 说明 |
|------|------|
| `rollback.manage` | 回滚管理 |
| `backups.manage` | 备份管理 |
| `snapshot.manage` | 快照管理 |

### 数据工具 (data bundle)

| 工具 | 说明 |
|------|------|
| `pq.list_connections` | 列出 PQ 连接 |
| `pq.list_queries` | 列出 PQ 查询 |
| `pq.get_code` | 获取 PQ 代码 |
| `pq.update_query` | 更新 PQ 查询 |
| `pq.refresh` | 刷新 PQ |

## 命令行使用

```bash
# 列出所有 Profile
python -m excelforge.gateway.host --list-profiles

# 列出所有 Bundle
python -m excelforge.gateway.host --list-bundles

# 使用指定 Profile 启动
python -m excelforge.gateway.host --config excel-mcp.yaml --profile automation

# 动态启用/禁用 Bundle
python -m excelforge.gateway.host --profile basic_edit --enable-bundle automation

# 打印 Runtime endpoint（用于诊断）
python -m excelforge.gateway.host --profile automation --print-runtime-endpoint
```

## 安全策略

### VBA 安全

- 写入 VBA 代码必须通过安全扫描
- 默认阻止 CRITICAL 和 HIGH 风险代码（如 `Shell`, `CreateObject("WScript.Shell")` 等）
- `MsgBox` 自动替换为 `Debug.Print` 以适应自动化环境

### 路径安全

- `runtime-config.yaml` 中的 `allowed_roots` 控制可访问的文件路径
- 支持通配符 `*` 表示允许所有路径

## 常见问题

### Q1: Runtime 启动失败

检查 `runtime-config.yaml` 中的配置是否正确，确保 Excel 已安装且可正常启动。

### Q2: Excel Worker 状态异常

执行 `server.health` 检查 Excel 状态。如果不是 ready 状态，尝试重启 Gateway。

### Q3: 工具调用超时

检查 `excel.ready` 字段。如果 Excel 正在计算，提高 `call_timeout_seconds` 配置。

### Q4: VBA 执行失败

确保 Excel 信任中心已启用"信任对 VBA 项目对象模型的访问"选项。

### Q5: 多个 Excel 进程残留

v2.2 已解决多进程隔离问题。如果仍有残留，手动关闭 Excel 进程后重试。

## 版本历史

| 版本 | 日期 | 主要更新 |
|------|------|----------|
| v0.1-v0.5 | 2026-03-22~24 | 基础功能开发 |
| v1.0.0 | 2026-03-24 | 工具组配置、简化工具链、Trae兼容 |
| v1.0.1 | 2026-03-24 | 工具合并、Profile机制、Worker健康检查、通配符路径 |
| v1.0.2 | 2026-03-26 | workbook.inspect返回增加index字段、修复三元表达式 |
| v2.0.0 | 2026-03-27 | Runtime重架构、ExcelWorker生命周期管理、named pipe通信 |
| v2.1.0 | 2026-03-28 | Runtime启动预热机制、Windows探活API修复、server.health增强 |
| **v2.2.0** | **2026-03-28** | **统一入口、Profile/Bundle装配、Runtime单例、解决多进程隔离问题** |

## 详细文档

- [V2开发记录文档.md](V2开发记录文档.md) - 开发过程问题追踪
- [设计文档V2.2 Profile 与 Bundle 工具装配优化（修订版）.md](设计文档V2.2%20Profile%20与%20Bundle%20工具装配优化（修订版）.md) - v2.2 设计文档
- [runtime-config.yaml](runtime-config.yaml) - Runtime 配置
- [excel-mcp.yaml](excel-mcp.yaml) - 统一入口配置
