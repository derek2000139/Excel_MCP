# ExcelForge

**ExcelForge** 是一个基于 MCP (Model Context Protocol) 的 Excel 操作工具集，让 AI 助手能够安全、高效地操作 Excel 文件。

## v2.0 架构

v2.0 采用多 MCP 网关架构，核心变化：

```
┌─────────────────────────────────────────────────────────┐
│                    MCP Clients (Trae/Cursor)            │
└─────────────────────────────────────────────────────────┘
                            │
        ┌───────────────────┼───────────────────┐
        ▼                   ▼                   ▼
┌───────────────┐  ┌───────────────┐  ┌─────────────────┐
│excel-core-mcp │  │ excel-vba-mcp │  │excel-recovery-mcp│
│   (Core)      │  │    (VBA)      │  │   (Recovery)     │
└───────┬───────┘  └───────┬───────┘  └────────┬────────┘
        │                  │                    │
        └──────────────────┼────────────────────┘
                           │ JSON-RPC
                           ▼
              ┌────────────────────────┐
              │  Runtime Service       │
              │  (Excel COM Objects)   │
              └────────────────────────┘
```

- **Runtime**：统一的 Excel COM 对象生命周期管理
- **Core Gateway**：工作簿、工作表、范围、公式、格式工具
- **VBA Gateway**：VBA 工程检查、代码同步、宏执行
- **Recovery Gateway**：快照、备份、回滚、审计、命名范围

## 功能特性

- 工作簿管理 - 打开、保存、关闭、创建 Excel 文件
- 工作表操作 - 创建、删除、重命名、查看结构、自动筛选
- 范围操作 - 读写数据、复制、清除、排序、合并单元格
- 公式支持 - 验证表达式、填充范围、获取依赖关系
- 格式设置 - 设置单元格样式、自动调整列宽
- VBA 读写访问 - 查看工程、扫描代码、同步模块、执行宏
- 备份恢复 - 文件级备份、快照回滚
- 命名范围 - 列出和读取命名区域
- 数据验证和条件格式 - 读取工作表规则
- 审计日志 - 操作记录追踪

## 版本

当前版本: **v2.1.0**

## 系统要求

- Windows 10/11
- Microsoft Excel Desktop（2016/2019/365）
- Python >= 3.11
- `uv` 包管理器

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/derek2000139/Excel_VBA_MCP.git
cd ExcelForge
```

### 2. 安装依赖

```bash
uv sync --extra dev
```

### 3. 启动服务

V2 支持两种启动方式：

#### 方式 A：一键启动（推荐）

Core Gateway 会自动启动 Runtime：

```bash
uv run python -m excelforge gateway-core --config ./excel-core-mcp.yaml
```

#### 方式 B：手动启动各组件

**终端 1 - 启动 Runtime：**

```bash
uv run python -m excelforge runtime --config ./runtime-config.yaml
```

**终端 2 - 启动 Core Gateway：**

```bash
uv run python -m excelforge gateway-core --config ./excel-core-mcp.yaml
```

**终端 3 - 启动 VBA Gateway（可选）：**

```bash
uv run python -m excelforge gateway-vba --config ./excel-vba-mcp.yaml
```

**终端 4 - 启动 Recovery Gateway（可选）：**

```bash
uv run python -m excelforge gateway-recovery --config ./excel-recovery-mcp.yaml
```

### 4. 在 IDE 中配置 MCP

参考 `mcp.example.json` 模板配置你的 IDE（MCP 客户端）：

```json
{
  "mcpServers": {
    "ExcelForge-Core": {
      "command": "uv",
      "args": [
        "run", "python", "-m", "excelforge",
        "gateway-core", "--config",
        "你的ExcelForge项目路径/excel-core-mcp.yaml"
      ],
      "cwd": "你的ExcelForge项目路径"
    },
    "ExcelForge-VBA": {
      "command": "uv",
      "args": [
        "run", "python", "-m", "excelforge",
        "gateway-vba", "--config",
        "你的ExcelForge项目路径/excel-vba-mcp.yaml"
      ],
      "cwd": "你的ExcelForge项目路径"
    },
    "ExcelForge-Recovery": {
      "command": "uv",
      "args": [
        "run", "python", "-m", "excelforge",
        "gateway-recovery", "--config",
        "你的ExcelForge项目路径/excel-recovery-mcp.yaml"
      ],
      "cwd": "你的ExcelForge项目路径"
    }
  }
}
```

## MCP 工具列表

### Core Gateway 工具（excel-core-mcp）

| 类别 | 工具 | 说明 |
|------|------|------|
| 工作簿 | `workbook.open_file` | 打开 Excel 文件 |
| 工作簿 | `workbook.create_file` | 创建新工作簿 |
| 工作簿 | `workbook.save_file` | 保存工作簿 |
| 工作簿 | `workbook.close_file` | 关闭工作簿 |
| 工作簿 | `workbook.inspect` | 查看工作簿信息 |
| 工作表 | `sheet.create_sheet` | 创建工作表 |
| 工作表 | `sheet.rename_sheet` | 重命名工作表 |
| 工作表 | `sheet.delete_sheet` | 删除工作表 |
| 工作表 | `sheet.inspect_structure` | 查看工作表结构 |
| 工作表 | `sheet.set_auto_filter` | 设置自动筛选 |
| 工作表 | `sheet.get_rules` | 获取条件格式/数据验证规则 |
| 范围 | `range.read_values` | 读取单元格值 |
| 范围 | `range.write_values` | 写入单元格值 |
| 范围 | `range.clear_contents` | 清除单元格内容 |
| 范围 | `range.copy_range` | 复制范围 |
| 范围 | `range.manage_rows` | 插入/删除行 |
| 范围 | `range.manage_columns` | 插入/删除列 |
| 范围 | `range.sort_data` | 排序数据 |
| 范围 | `range.merge_cells` | 合并单元格 |
| 范围 | `range.unmerge_cells` | 取消合并 |
| 范围 | `range.manage_merge` | 管理合并 |
| 公式 | `formula.fill_range` | 填充公式范围 |
| 公式 | `formula.set_single` | 设置单个单元格公式 |
| 公式 | `formula.get_dependencies` | 获取公式依赖 |
| 公式 | `formula.repair_references` | 修复公式引用 |
| 格式 | `format.manage` | 格式管理 |
| 服务器 | `server.status` | 获取服务状态 |
| 服务器 | `server.health` | 健康检查 |

### VBA Gateway 工具（excel-vba-mcp）

| 工具 | 说明 |
|------|------|
| `vba.inspect_project` | 查看 VBA 工程 |
| `vba.get_module_code` | 获取模块代码 |
| `vba.scan_code` | 扫描 VBA 代码 |
| `vba.sync_module` | 同步 VBA 模块 |
| `vba.remove_module` | 删除 VBA 模块 |
| `vba.execute` | 执行 VBA 宏 |
| `vba.compile` | 编译 VBA 工程 |

### Recovery Gateway 工具（excel-recovery-mcp）

| 工具 | 说明 |
|------|------|
| `rollback.undo_last` | 撤销最后操作 |
| `rollback.manage` | 快照管理 |
| `backups.manage` | 备份管理 |
| `audit.list_operations` | 审计日志 |
| `names.inspect` | 查看命名范围 |
| `names.create` | 创建命名范围 |
| `names.delete` | 删除命名范围 |

## 配置说明

### Runtime 配置（runtime-config.yaml）

```yaml
runtime:
  version: "2.1.0"
  pipe_name: "\\\\.\\pipe\\excelforge-runtime"

excel:
  visible: true              # Excel 窗口是否可见
  disable_events: true       # 禁用事件
  disable_alerts: true       # 禁用警告弹窗
  force_disable_macros: false # 强制禁用宏
  enable_warmup: true         # 启用预热机制
  startup_timeout_seconds: 120 # 启动预热超时

paths:
  allowed_roots:
    - "*"                    # 允许访问所有路径
  allowed_extensions:
    - ".xlsx"
    - ".xlsm"
    - ".xlsb"
```

### Gateway 配置（excel-*-mcp.yaml）

```yaml
gateway:
  id: "excel-core-mcp"
  display_name: "ExcelForge Core"
  runtime_data_dir: "./.runtime_data_v2"
  auto_start_runtime: true   # 自动启动 Runtime
  runtime_config_path: "./runtime-config.yaml"
  connect_timeout_seconds: 10
  call_timeout_seconds: 30
```

## 安全策略

### VBA 安全

- 写入 VBA 代码必须通过安全扫描
- 默认阻止 CRITICAL 和 HIGH 风险代码（如 `Shell`, `CreateObject("WScript.Shell")` 等）
- `MsgBox` 自动替换为 `Debug.Print`
- `InputBox` 被禁用以避免弹窗阻塞

### 路径访问控制

`runtime-config.yaml` 中的 `paths.allowed_roots` 控制可以访问哪些路径：

```yaml
paths:
  allowed_roots:
    - "*"           # 允许所有路径（测试用）
    - "C:/Users"    # 或指定具体路径
```

**安全建议**：生产环境建议限制为具体工作目录。

## 常见问题

### Q1: 提示"Runtime lock file not found"

Runtime 未启动。请先启动 Runtime 服务：
```bash
uv run python -m excelforge runtime --config ./runtime-config.yaml
```

### Q2: 提示"文件路径不允许"

修改 `runtime-config.yaml`，将 `allowed_roots` 改为 `"*"` 或添加目标路径。

### Q3: Excel Worker 状态异常

执行 `server.status` 或 `server.health` 检查 Excel 状态。如果不是 ready 状态，尝试重启 Gateway。

### Q4: 工具调用超时

检查 `excel_worker.state` 和 `excel.ready` 字段。如果 Excel 正在计算，提高 `call_timeout_seconds` 配置。

### Q5: VBA 执行失败

确保 Excel 信任中心已启用"信任对 VBA 项目对象模型的访问"选项。

## 版本历史

| 版本 | 日期 | 主要更新 |
|------|------|----------|
| v0.1-v0.5 | 2026-03-22~24 | 基础功能开发 |
| v1.0.0 | 2026-03-24 | 工具组配置、简化工具链、Trae兼容 |
| v1.0.1 | 2026-03-24 | 工具合并、Profile机制、Worker健康检查、通配符路径 |
| v1.0.2 | 2026-03-26 | workbook.inspect返回增加index字段、修复三元表达式为if-else、CLAUDE.md增加MCP设计决策说明 |
| **v2.0.0** | **2026-03-27** | **Runtime重架构、ExcelWorker生命周期管理、named pipe通信、snapshot/rollback/backup三大恢复机制** |
| **v2.1.0** | **2026-03-28** | **Runtime启动预热机制、Windows探活API修复、Dispatcher就绪拦截、server.health健康检查、首次请求0.43秒响应** |

## 详细文档

- [MULTI_MCP_QUICKSTART.md](MULTI_MCP_QUICKSTART.md) - 多网关快速启动指南
- [mcp.example.json](mcp.example.json) - MCP 客户端配置示例
- [runtime-config.yaml](runtime-config.yaml) - Runtime 配置
- [excel-*-mcp.yaml](excel-core-mcp.yaml) - 各网关配置
