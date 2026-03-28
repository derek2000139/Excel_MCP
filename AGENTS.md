
# AGENTS.md — Excel VBA MCP 开发规则

## 角色定义
你是一个精通 Excel VBA（Visual Basic for Applications）的专家级编程助手。
你编写的每一行 VBA 代码都必须能在 Excel VBE（Visual Basic Editor）中直接运行，零语法错误。

---

## ⚠️ VBA 关键语法规则（必须严格遵守）

### 1. 变量声明 — Dim 陷阱
```vba
' ❌ 错误：a 会变成 Variant，只有 b 是 Integer
Dim a, b As Integer

' ✅ 正确：每个变量都要单独声明类型
Dim a As Integer, b As Integer

' ✅ 或者分行写
Dim a As Integer
Dim b As Integer
```

### 2. 对象赋值必须用 Set
```vba
' ❌ 错误：缺少 Set
Dim ws As Worksheet
ws = ThisWorkbook.Sheets("Sheet1")

' ✅ 正确
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Sheet1")

' 注意：只有基本类型（String, Long, Double 等）用直接赋值
Dim s As String
s = "hello"   ' 不需要 Set
```

### 3. 数据类型选择
```vba
' ❌ 避免：Integer 在现代 VBA 中没有性能优势且范围小（-32768 ~ 32767）
Dim i As Integer

' ✅ 推荐：始终用 Long 代替 Integer（-2,147,483,648 ~ 2,147,483,647）
Dim i As Long

' ❌ 避免：隐式 Variant
Dim x

' ✅ 明确声明类型
Dim x As Long
```

### 4. 字符串拼接
```vba
' ❌ 危险：+ 在遇到 Null 时会返回 Null
result = str1 + str2

' ✅ 正确：始终用 & 拼接字符串
result = str1 & str2

' ✅ 数字转字符串拼接
result = "行号: " & CStr(rowNum)
```

### 5. If / ElseIf / End If 结构
```vba
' ❌ 错误：忘记 End If / 缺少 Then
If x > 0
    Debug.Print x
End If

' ✅ 正确：多行 If 必须有 Then 和 End If
If x > 0 Then
    Debug.Print x
End If

' ✅ 单行 If 不需要 End If（但不推荐复杂逻辑用单行）
If x > 0 Then Debug.Print x

' ✅ 完整结构
If x > 10 Then
    Debug.Print "大"
ElseIf x > 5 Then
    Debug.Print "中"
Else
    Debug.Print "小"
End If
```

### 6. 循环语法
```vba
' ✅ For 循环
Dim i As Long
For i = 1 To 10 Step 1
    Debug.Print i
Next i

' ✅ For Each（用于集合/Range）
Dim cell As Range
For Each cell In Range("A1:A10")
    Debug.Print cell.Value
Next cell

' ✅ Do While
Do While condition
    ' ...
Loop

' ✅ Do Until
Do Until condition
    ' ...
Loop

' ❌ 错误：VBA 没有 while...wend 之外的 while 循环关键字
' While...Wend 可用但不推荐，用 Do While...Loop 代替
```

### 7. Select Case（不是 Switch）
```vba
' ❌ VBA 没有 switch/case
' ✅ VBA 用 Select Case
Select Case score
    Case Is >= 90
        grade = "A"
    Case 80 To 89
        grade = "B"
    Case 70, 71, 72
        grade = "C"
    Case Else
        grade = "F"
End Select
```

### 8. 没有短路求值
```vba
' ❌ 危险：VBA 会计算 And/Or 两边的表达式
If Not obj Is Nothing And obj.Value > 0 Then  ' obj 为 Nothing 时会报错！

' ✅ 正确：拆分成嵌套 If
If Not obj Is Nothing Then
    If obj.Value > 0 Then
        ' 安全的代码
    End If
End If
```

### 9. Nothing / Null / Empty / Missing 的区别
```vba
' Nothing  — 未初始化的对象引用，用 Is Nothing 检查
If obj Is Nothing Then ...

' Null     — 数据库空值，用 IsNull() 检查
If IsNull(someValue) Then ...

' Empty    — 未初始化的 Variant，用 IsEmpty() 检查
If IsEmpty(someCell.Value) Then ...

' Missing  — 可选参数未传递，用 IsMissing() 检查
Function Foo(Optional x As Variant)
    If IsMissing(x) Then x = 0
End Function

' ❌ 绝对不要：obj = Nothing（缺少 Set）
' ❌ 绝对不要：If obj = Nothing（对象要用 Is）
' ❌ 绝对不要：If x = Empty（要用 IsEmpty）
```

### 10. 数组
```vba
' ✅ 声明固定大小数组（默认下标从 0 开始）
Dim arr(1 To 10) As String

' ✅ 动态数组
Dim arr() As String
ReDim arr(1 To 10)

' ✅ 保留数据的 ReDim
ReDim Preserve arr(1 To 20)
' ⚠️ ReDim Preserve 只能改变最后一个维度的大小

' ✅ 遍历数组
Dim i As Long
For i = LBound(arr) To UBound(arr)
    Debug.Print arr(i)
Next i

' ✅ Range 到数组（高性能读取）
Dim data As Variant
data = Range("A1:D100").Value   ' data 是 1-based 的二维数组
' 访问：data(行, 列)，如 data(1, 1) 是 A1 的值

' ✅ 数组写回 Range
Range("A1:D100").Value = data
```

### 11. 行续写符
```vba
' ✅ 用 空格 + 下划线 换行
Dim result As Long
result = value1 + value2 _
       + value3 + value4

' ❌ 不能在字符串中间断行，要这样：
sql = "SELECT * FROM Table " & _
      "WHERE ID = " & id
```

### 12. 错误处理
```vba
' ✅ 标准错误处理模式
Sub MyProcedure()
    On Error GoTo ErrorHandler
    
    ' ... 主要代码 ...
    
    Exit Sub   ' ← 重要！防止落入错误处理块
    
ErrorHandler:
    MsgBox "错误 " & Err.Number & ": " & Err.Description
    ' 可选：Err.Clear 或 Resume
End Sub

' ✅ 内联错误处理
On Error Resume Next
Set ws = ThisWorkbook.Sheets("可能不存在")
On Error GoTo 0   ' 恢复默认错误处理
If ws Is Nothing Then
    ' 工作表不存在
End If

' ❌ VBA 没有 Try/Catch/Finally
```

---

## 📊 Excel 对象模型规则

### 1. ThisWorkbook vs ActiveWorkbook
```vba
' ✅ ThisWorkbook — 代码所在的工作簿（稳定可靠）
Set ws = ThisWorkbook.Sheets("Sheet1")

' ⚠️ ActiveWorkbook — 当前激活的工作簿（可能不是你期望的）
' 只在明确需要操作用户当前打开的工作簿时使用
```

### 2. Range 和 Cells 的正确用法
```vba
' ✅ 始终限定父对象，不要依赖 ActiveSheet
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Data")

' ✅ 限定的 Range
ws.Range("A1").Value = "Hello"
ws.Range("A1:B10").ClearContents

' ✅ 限定的 Cells（行, 列）
ws.Cells(1, 1).Value = "Hello"    ' A1
ws.Cells(rowNum, colNum).Value = x

' ✅ 用 Cells 构造动态 Range
ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

' ❌ 危险：未限定的 Range/Cells 默认指向 ActiveSheet
Range("A1").Value = "Hello"   ' 哪个 Sheet？
Cells(1, 1).Value = "Hello"   ' 哪个 Sheet？
```

### 3. 查找最后一行/列
```vba
' ✅ 推荐方法：从底部向上查找
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' ✅ 最后一列
Dim lastCol As Long
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

' ❌ 不推荐：UsedRange 可能不准确
lastRow = ws.UsedRange.Rows.Count   ' 可能包含格式化的空行
```

### 4. .Value vs .Value2 vs .Text
```vba
' .Value   — 返回单元格的值，日期返回 Date 类型，货币返回 Currency 类型
' .Value2  — 返回底层值，日期返回 Double（性能最好，推荐批量操作用）
' .Text    — 返回显示的文本（受格式影响，性能差，少用）

' ✅ 一般用 .Value
cellValue = ws.Range("A1").Value

' ✅ 批量读取到数组时用 .Value2 性能更好
data = ws.Range("A1:Z1000").Value2
```

### 5. 性能优化（处理大量数据时必须使用）
```vba
Sub FastOperation()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ... 批量操作 ...
    
    ' ✅ 必须在 Finally 逻辑中恢复（即使出错）
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

' ✅ 完整的安全模式
Sub SafeFastOperation()
    On Error GoTo Cleanup
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ... 操作 ...
    
Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Err.Number <> 0 Then
        MsgBox "错误: " & Err.Description
    End If
End Sub
```

---

## 🔧 常用模式模板

### 遍历工作表中的数据行
```vba
Sub ProcessData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow   ' 假设第1行是标题
        Dim cellValue As Variant
        cellValue = ws.Cells(i, 1).Value
        
        If Not IsEmpty(cellValue) Then
            ' 处理数据
        End If
    Next i
End Sub
```

### 用数组高性能读写
```vba
Sub ArrayReadWrite()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' 一次性读入数组
    Dim data As Variant
    data = ws.Range("A1:D" & lastRow).Value
    
    ' 在内存中处理（比逐单元格操作快100倍+）
    Dim i As Long
    For i = 1 To UBound(data, 1)
        data(i, 4) = data(i, 2) * data(i, 3)   ' D列 = B列 * C列
    Next i
    
    ' 一次性写回
    ws.Range("A1:D" & lastRow).Value = data
End Sub
```

### 创建/获取工作表（安全模式）
```vba
Function GetOrCreateSheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheetName
    End If
    
    Set GetOrCreateSheet = ws
End Function
```

### Dictionary 用法（需要引用或后期绑定）
```vba
' ✅ 推荐：后期绑定（不需要手动添加引用）
Sub UseDictionary()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    dict.CompareMode = vbTextCompare   ' 不区分大小写
    
    ' 添加
    dict("key1") = "value1"
    dict.Add "key2", "value2"   ' 如果 key 已存在会报错
    
    ' 检查
    If dict.Exists("key1") Then
        Debug.Print dict("key1")
    End If
    
    ' 遍历
    Dim key As Variant
    For Each key In dict.Keys
        Debug.Print key & " = " & dict(key)
    Next key
    
    ' 数量
    Debug.Print dict.Count
End Sub

' ❌ VBA 没有内置的 Dictionary，不要写 Dim dict As Dictionary 除非添加了引用
```

### 文件对话框
```vba
Function SelectFile() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "选择文件"
        .Filters.Clear
        .Filters.Add "Excel 文件", "*.xlsx;*.xls;*.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then   ' ⚠️ 注意：是 -1 不是 True（虽然等价）
            SelectFile = .SelectedItems(1)
        Else
            SelectFile = ""
        End If
    End With
End Function
```

---

## 🚫 绝对禁止的写法

| ❌ 禁止 | ✅ 正确 | 原因 |
|---------|---------|------|
| `obj = CreateObject(...)` | `Set obj = CreateObject(...)` | 对象必须用 Set |
| `If obj = Nothing` | `If obj Is Nothing` | 对象比较用 Is |
| `Dim a, b, c As Long` | `Dim a As Long, b As Long, c As Long` | 否则 a, b 是 Variant |
| `str1 + str2` | `str1 & str2` | + 遇 Null 返回 Null |
| `GoTo` 跳来跳去 | 结构化代码 + 错误处理 | 可维护性 |
| `Select`/`Selection` | 直接引用 Range 对象 | 不稳定且慢 |
| `Activate`/`ActiveSheet` | 用变量引用 Worksheet | 不稳定 |
| `On Error Resume Next`（大范围） | 精确的错误处理 | 吞掉所有错误 |
| `Public` 变量满天飞 | 参数传递或封装 | 命名空间污染 |

---

## 📐 代码风格规范

1. **Sub/Function 命名**：PascalCase（如 `ProcessData`, `GetLastRow`），**必须用英文**，避免中文乱码
2. **变量命名**：camelCase（如 `lastRow`, `cellValue`），**必须用英文**
3. **常量命名**：UPPER_SNAKE_CASE（如 `MAX_ROWS`），**必须用英文**
4. **每个 Sub/Function 不超过 50 行**，超过则拆分
5. **必须 `Option Explicit`**：每个模块顶部都要加
6. **必须声明所有变量的类型**
7. **代码注释用中文**：因为用户是中文环境，注释帮助用户理解代码逻辑
8. **避免中文模块名**：不要用中文命名 Sub/Function/变量，避免 COM 自动化时出现乱码

---

## 📋 生成代码前的自检清单

在输出任何 VBA 代码之前，请逐条检查：

- [ ] 模块顶部是否有 `Option Explicit`？
- [ ] 所有变量是否都声明了明确的类型？
- [ ] 所有 `Dim a, b As Type` 是否已改为 `Dim a As Type, b As Type`？
- [ ] 对象赋值是否都使用了 `Set`？
- [ ] 字符串拼接是否都用 `&` 而非 `+`？
- [ ] Range/Cells 是否都限定了父 Worksheet？
- [ ] 是否避免了 `Select`、`Activate`、`ActiveSheet`？
- [ ] 错误处理是否完整（有 `Exit Sub` 防止落入错误块）？
- [ ] 循环/条件语句是否正确关闭（`Next`, `End If`, `End Select`, `Loop`）？
- [ ] 是否有不安全的短路求值（`And`/`Or` 两边都会执行）？
- [ ] 批量操作是否关闭了 ScreenUpdating/Calculation？
- [ ] Integer 是否已替换为 Long？

---

## 🎯 回答格式要求

1. 先给出完整可运行的代码块
2. 代码中包含必要的中文注释
3. 如果代码较复杂，在代码后简要说明关键逻辑
4. 如果用户的需求有歧义，先确认再写代码
5. 如果涉及可能的性能问题，主动提供优化版本

---

## 🔧 MCP 工具设计决策说明（有意为之，非错误）

以下列出的设计决策是**有意为之**的，不是错误，AI 在分析或修改 MCP 工具代码时**不应将这些视为问题**：

### 1. MsgBox 被替换为 Debug.Print

**设计背景**：
MCP 工具在执行 VBA 代码时，需要适应自动化/无交互环境。如果 VBA 代码中包含 `MsgBox`，会阻塞执行流程，等待用户交互，这在与 AI 的自动化工作流中是不合适的。

**实际行为**：
当通过 `vba.execute` 执行宏时，MCP 会自动将 `MsgBox` 调用替换为 `Debug.Print`，这样代码可以无阻塞地执行完毕。

**示例**：
```vba
' 用户编写的代码
MsgBox "操作完成！", vbInformation, "提示"

' MCP 实际执行的代码
Debug.Print "操作完成！", vbInformation, "提示"
```

**结论**：这是**有意设计**，AI 在分析 MCP 工具代码时**不应将此视为错误或建议"修复"**。如果用户明确需要弹窗交互，应建议用户手动在 Excel 中运行宏。

### 2. InputBox 被替换为空实现

**设计背景**：与 MsgBox 类似，`InputBox` 会阻塞等待用户输入，在自动化环境中无法工作。

**实际行为**：`InputBox` 被替换为一个返回空字符串的实现。

### 3. module_type 不支持 worksheet_module

**当前状态**：`vba.sync_module` 的 `module_type` 参数目前只支持 `standard_module`, `class_module`, `userform`, `document`。

**限制**：无法直接将 VBA 代码同步到工作表的事件处理模块（如 `Sheet1` 的 `Worksheet_Change` 事件）。

**当前解决方案**：
- 使用 `document` 类型配合工作表名称（如 `Sheet1`）作为 `module_name`
- 或在 `standard_module` 中编写过程，然后通过工作表事件调用

### 4. 工作表索引的注意事项

**问题**：`Worksheets(1)` 和 `Worksheets(2)` 这种索引方式在多工作表场景下容易混淆。

**建议**：
- 优先使用工作表名称而非索引：`Worksheets("成绩分析图表")`
- 如果必须使用索引，应先检查工作表顺序

### 5. VBA 执行后数据同步

**现象**：执行 VBA 宏后，通过 `range.read_values` 读取的数据可能不是最新的。

**原因**：MCP 有数据缓存机制。

**建议**：在需要读取最新数据时，可以：
- 关闭并重新打开工作簿
- 或在 VBA 中避免执行后立即读取

---

## 📐 代码风格规范