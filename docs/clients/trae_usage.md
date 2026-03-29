# Trae AI 配置与使用指南

**文档版本**：`V2.4`\
**更新日期**：`2026-03-30`

***

## 1. Trae AI MCP 限制说明

### 1.1 工具数量限制

Trae AI 对 MCP 工具数量有 **39 个**的截断限制。当 profile 包含的工具数量超过 39 个时，后半部分工具会被截断，导致无法调用。

### 1.2 问题现象

- 某些工具（如 VBA、Recovery 等）无法识别或调用
- 明明定义了工具，但调用时报错 "Tool is not available"
- `server.health` 显示正常，但实际调用失败

### 1.3 验证方法

1. 打开 Trae AI 的工具列表
2. 搜索 `vba.` 或 `backup.` 等关键字
3. 如果显示"未找到工具"，说明该工具被截断了

***

## 2. 推荐配置

### 2.1 基础配置（推荐）

`data_workflow` profile 包含 33 个工具，低于 39 限制，且覆盖常用数据处理功能：

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
      "cwd": "YOUR_PROJECT_PATH/"
    }
  }
}
```

### 2.2 扩展配置

如需在 `data_workflow` 基础上增加特定工具，使用 `--enable-bundle`：

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
      "cwd": "YOUR_PROJECT_PATH/"
    }
  }
}
```

> **注意**：扩展后总工具数需保持在 39 以下。`data_workflow + edit_structure` = 33 + 6 = 39，刚好。

### 2.3 调试配置

如需使用 VBA 工具进行调试：

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
        "automation"
      ],
      "cwd": "YOUR_PROJECT_PATH/"
    }
  }
}
```

> `automation` profile 包含 40 个工具，会触发截断。建议调试时临时使用。

***

## 3. 工具验证

### 3.1 验证步骤

修改配置后，重启 Trae AI MCP 服务，然后：

1. 检查工具列表是否包含目标工具
2. 调用 `server.health` 确认 Runtime 正常
3. 调用目标工具测试功能

### 3.2 快速验证命令

```
server.health                    # 确认 Runtime 运行
workbook.list_open              # 验证 workbook 域
table.list_tables               # 验证 data bundle
analysis.scan_structure          # 验证 analysis bundle
workbook.save_as                # 验证 workbook_ops bundle
```

### 3.3 诊断参数

如遇问题，可使用以下诊断参数：

```bash
# 导出所有工具列表
python -m excelforge.gateway.host --dump-tools --profile data_workflow

# 导出带索引的工具列表
python -m excelforge.gateway.host --dump-tools-with-index --profile data_workflow

# 导出 Profile 解析过程
python -m excelforge.gateway.host --dump-profile-resolution --profile data_workflow
```

***

## 4. Profile 工具数量参考

| Profile | 工具数 | tool_budget | 建议 |
|---------|---:|---:|------|
| basic_edit | 35 | 40 | 低于 39，可直接使用 |
| calc_format | 46 | 50 | 超出限制，不推荐 |
| automation | 40 | 45 | 超出限制，不推荐 |
| **data_workflow** | **33** | **39** | **推荐** |
| reporting | 32 | 37 | 低于 39，可直接使用 |
| vba_first | ~28 | 30 | 低于 39，可直接使用 |

***

## 5. 常见问题

### 5.1 工具显示"未找到"

**原因**：profile 工具数超过 39，工具被截断。

**解决**：切换到 `data_workflow` profile，或减少 `--enable-bundle`。

### 5.2 Runtime 无法启动

**原因**：Python 环境问题或端口占用。

**解决**：
1. 确认 uv 环境可用
2. 使用 `--restart-runtime always` 强制重启
3. 检查日志文件：`C:\Users\<用户名>\.excelforge\logs\`

### 5.3 Excel 进程残留

**原因**：Runtime 异常退出。

**解决**：手动 kill Excel 进程，或重启电脑清理。
