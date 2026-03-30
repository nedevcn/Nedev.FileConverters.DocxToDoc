# Nedev.FileConverters.DocxToDoc

一个将 OpenXML `.docx` 转换为二进制 `.doc` 的 .NET 库，不依赖 Office COM 自动化。

- 核心库目标框架：`net8.0`、`netstandard2.1`
- CLI 目标框架：`net8.0`
- 依赖：`OpenMcdf`、`System.Text.Encoding.CodePages`、`Nedev.FileConverters.Core`

---

## 代码现状与功能完成度（2026-03-30）

本项目已具备“可用于生产中的常见文档转换”能力，但并非 Word 完整渲染/排版引擎。  
当前以 **443 个测试全部通过（443/443）** 和全解 `Release` 构建通过为基准。

### 1) 核心转换链路（完成度：高）

- 同步/异步转换：`Convert`、`ConvertAsync`（路径与流双入口）
- 错误分层：`Validation / Reading / Parsing / Writing / Finalizing / Unknown`
- 性能与监控：基础日志、字节统计、图片/段落/表格计数
- CLI：单文件、批处理、递归、详细日志、错误码

### 2) 文本与段落（完成度：高）

- 段落、run、制表符、换行、符号字符、软/不间断连字符
- 常见字符样式：粗体、斜体、删除线、字号、下划线、字体、颜色
- 段落属性：对齐、缩进、段前后距、行距
- 字段：`fldSimple` 与复杂字段 begin/separate/end，含超链接字段

### 3) 表格（完成度：中高，启发式）

- `tblGrid/gridSpan/tcW/tblW` 解析与写出
- 单元格边距、边框、内边框、边框冲突规则
- 行高规则（auto/atLeast/exact）、垂直对齐、单元格间距
- 嵌套表格与表内浮动对象布局

> 说明：表格布局采用近似策略，不保证与 Word 像素级一致。

### 4) 图片与图形（完成度：中高，启发式）

- 解析并写出内联/浮动图片
- 支持超链接包裹图片
- OfficeArt/Escher 输出支持：PNG/JPEG/EMF/WMF
- 部分 VML fallback 路径可读

### 5) 页眉页脚、分节、脚注尾注、批注（完成度：中）

- 分节：页大小、页边距、首页不同、奇偶页设置、页码起始
- 页眉页脚：默认/首页/偶数页故事，支持文本、字段、超链接、图片、表格块
- 脚注尾注：引用锚点恢复、特殊分隔故事、多段纯文本写出
- 批注：锚点恢复、回复元数据读取，写出时进行兼容性折叠

### 6) 其他能力（完成度：中）

- 样式表、字体表、编号系统（abstract/instance）
- 文档属性（`docProps/core.xml` + `docProps/app.xml`）
- `altChunk`：文本/常见标记/RTF 可见文本提取，嵌入 DOCX 的可见段落与表格块保留
- VBA 宏：读取并嵌入 `Macros` 存储

---

## 缺失与不完整功能清单

以下为“当前版本已知边界”，属于设计上有意降级或尚未实现：

1. 版式保真不是 Word 引擎级别
- 段落换行/高度、浮动对象定位、复杂表格场景均为启发式
- 与 Word 在极端样例下可能出现分页/位置差异

2. `altChunk` 保真有限
- HTML/CSS 未做完整渲染
- RTF 仅可见文本优先，样式细节不完整
- 嵌入 DOCX 仅做“可见块”级保留，不是全语义回放

3. 批注与线程语义不完整
- 回复链在写出阶段会折叠到兼容文本故事
- 非法或不可锚定批注会被过滤/降级，而非完整保真

4. 脚注尾注标记保真有限
- 自定义标记依赖可连续可见文本恢复
- 更复杂的引用样式/版式细节未完整映射

5. 高级 Office 特性缺失
- 公式（OMML）仅有简化处理（非等价对象输出）
- SmartArt、修订痕迹的完全保真、复杂兼容行为未实现

6. 测试工程仍有工程化改进空间
- 存在模板遗留空测试 `UnitTest1.Test1`
- CLI 测试在未构建 CLI 时以 `return` 方式跳过，非显式 Skip 标记

---

## 改进计划（分阶段）

### P0（优先，先做稳定性与可验证性）

- 清理空测试，改为有效断言或删除
- CLI 测试改为显式 Skip/条件化 fixture，避免“静默跳过”
- 增加端到端样例集（真实 docx/doc 对照）与回归基线
- 在 CI 固化：`restore + build + test + pack`

### P1（优先，先补用户最敏感保真点）

- 表格布局：补齐更多 `tblLayout`、复杂合并、跨页场景样例驱动修正
- 浮动对象：增强相对定位与约束裁剪策略
- 页眉页脚：提升复杂混排（字段+图+表）一致性

### P2（中期，增强语义保真）

- `altChunk`：扩展 HTML/RTF 语义映射深度
- 批注：改进线程结构保留与写出映射
- 脚注尾注：提升自定义标记与复杂引用链还原

### P3（可选，高成本能力）

- 公式对象更完整映射
- 修订痕迹更高保真输出
- 高级对象（SmartArt 等）策略化降级与可配置处理

---

## 安装

```powershell
dotnet add package Nedev.FileConverters.DocxToDoc --version 0.1.0
```

或：

```xml
<PackageReference Include="Nedev.FileConverters.DocxToDoc" Version="0.1.0" />
```

---

## 使用方式

### 库调用

```csharp
using Nedev.FileConverters.DocxToDoc;

var converter = new DocxToDocConverter();
converter.Convert("input.docx", "output.doc");

using var input = File.OpenRead("input.docx");
using var output = File.Create("output.doc");
converter.Convert(input, output);
```

### CLI

```powershell
cd src\Nedev.FileConverters.DocxToDoc.Cli
dotnet build -c Release
dotnet bin\Release\net8.0\Nedev.FileConverters.DocxToDoc.Cli.dll <input.docx> <output.doc>
```

常用参数：

- `-b, --batch`：目录批量转换
- `-r, --recursive`：递归子目录
- `-o, --output`：输出文件或目录
- `-v, --verbose`：详细日志
- `-h, --help`：帮助

---

## 开发与验证

```powershell
dotnet restore
dotnet build .\Nedev.FileConverters.DocxToDoc.sln -c Release
dotnet test .\Nedev.FileConverters.DocxToDoc.sln
```

---

## 许可证

MIT，详见 [LICENSE](LICENSE)。
