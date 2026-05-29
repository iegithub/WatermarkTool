# 批量水印工具 - WatermarkTool

一个用于批量给 Excel 和 Word 文档添加水印的 Windows 桌面应用程序。

## 功能特性

- ✅ **批量处理**：支持同时处理多个 Excel (.xlsx) 和 Word (.docx) 文件
- ✅ **两种水印样式**：艺术字效果（阴影+渐变+描边）或 半透明文字
- ✅ **可视化预览**：拖拽调整水印位置，实时预览效果
- ✅ **关键字匹配**：根据文件名自动选择对应的水印文字
- ✅ **不影响原文**：Word水印通过页眉添加，Excel水印通过VML绘图层添加，不修改正文内容
- ✅ **不触发公式重算**：Excel水印方式不会导致公式重新计算

## 下载

### 方式1：GitHub Actions 自动构建
每次推送代码后，GitHub Actions 会自动编译并生成 EXE 文件。

1. 进入本仓库的 **Actions** 标签页
2. 点击最新的工作流运行记录
3. 在 **Artifacts** 部分下载 `WatermarkTool-EXE`

### 方式2：Release 页面
推送标签后会自动创建 Release 并上传 EXE：
```bash
git tag v1.0.0
git push origin v1.0.0
```

## 自行编译

### 环境要求
- Windows 10/11
- [.NET 6.0 SDK](https://dotnet.microsoft.com/download/dotnet/6.0)

### 编译步骤
```bash
# 克隆仓库
git clone https://github.com/你的用户名/WatermarkTool.git
cd WatermarkTool

# 编译并发布为单文件 EXE
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true

# 生成的 EXE 位置
# bin\Release\net6.0-windows\win-x64\publish\WatermarkTool.exe
```

## 使用说明

1. **添加文件**：点击"添加文件"或"添加文件夹"选择要处理的文档
2. **设置水印**：
   - 输入默认水印文字
   - 可选：添加关键字规则（文件名包含关键字时使用对应水印）
   - 调整字体、颜色、大小、透明度、旋转角度
   - 点击"预览位置"拖拽调整水印位置
3. **开始处理**：点击"开始批量添加水印"

## 技术架构

- **框架**：.NET 6.0 + Windows Forms
- **文档处理**：DocumentFormat.OpenXml (Open XML SDK)
- **Word水印**：通过页眉嵌入水印图片，不影响正文
- **Excel水印**：通过 VML 绘图覆盖层，不修改单元格数据

## 项目结构

```
WatermarkTool/
├── Models/              # 数据模型
│   ├── WatermarkSettings.cs
│   ├── KeywordRule.cs
│   └── ProcessResult.cs
├── Services/            # 业务逻辑
│   ├── WordWatermarkService.cs
│   ├── ExcelWatermarkService.cs
│   └── KeywordMatcher.cs
├── Forms/               # UI 窗体
│   ├── MainForm.cs
│   └── WatermarkPreviewForm.cs
└── Program.cs
```

## License

MIT License
