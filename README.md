# 欠料表处理工具

基于 **C# (.NET 8) + WPF + ClosedXML** 开发的 Windows 桌面软件，用于把 ERP 导出的欠料表，按「外协 / 外购」自动拆分生成回复单，并把供应商回复回填到原表。

> GitHub 仓库：`soldier25/qiliao-tool`（仓库 slug 仅支持英文，故使用 `qiliao-tool`）。

## 功能特性

- **导出回复单**：选择欠料表后自动加载判定列，无需手动找列；支持按「外购 / 外协」类别拆分、按采购员筛选；导出完成后自动打开输出文件夹。
- **回填回复**：默认目标列「采购交期回复」（按列名自动定位）；自动识别 Excel 日期序列号 / `DateTime` 并转换为真实日期，修复早期版本「回填后变成数字」的问题。
- **白底粉紫渐变主题**：以白色为基底，主按钮粉色、次按钮紫色，头部粉→紫渐变，简约清爽。
- **单文件、免安装**：发布为单个 `欠料表处理工具.exe`，内置 .NET 运行时，拷贝到任意 Windows 64 位电脑双击即可运行。

## 技术栈

- .NET 8 / WPF
- ClosedXML（Excel 读写）
- MVVM 结构

## 目录结构

```
欠料表处理工具.sln
欠料表回复工具.csproj
App.xaml / App.xaml.cs
MainWindow.xaml / MainWindow.xaml.cs
NativeCore.cs          # 原生库封装（qiliao_core）
QiliaoService.cs       # 导出 / 回填业务核心
Assets/app.ico         # 应用图标（粉紫渐变）
```

## 构建与运行

> 本机需安装 [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)（Windows x64）。

```powershell
# 在项目目录（含 .csproj 的目录）执行：
dotnet restore
dotnet build -c Release
dotnet run -c Release
```

## 发布为单文件 exe（推荐给同事使用）

```powershell
# 自包含单文件（内置 .NET 运行时，约 71 MB，目标机无需安装环境）：
dotnet publish -c Release -r win-x64 --self-contained true `
  -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
  -o ./publish
```

`publish/` 目录下生成的 `欠料表处理工具.exe` 即为独立可执行文件，可直接拷贝分发。

> 如需更小体积（约 10 MB），可去掉 `--self-contained true` 改为框架依赖版，但目标机需先安装 .NET 8 桌面运行时。

## 使用步骤

1. 打开软件，选择「导出」页签，点「选择欠料表」载入 Excel。
2. 选择类别（外购 / 外协 / 全部），按需填写采购员筛选。
3. 点「开始导出」，得到按供应商 / 采购员拆分的回复单，并自动打开输出文件夹。
4. 切换「回填」页签，选择原表与回复文件，点「开始回填」，结果自动打开。

## 版本说明

- **v2.15**：白底粉紫渐变主题；移除暗色主题与设置按钮、独立的导出/回填结果反馈区（信息统一进入底部处理日志）；新增粉紫应用图标；回填默认「采购交期回复」列并修复日期变数字问题；封装为真正单文件 exe（原生组件内嵌）。
- 早期版本（Python + openpyxl 脚本 v5.0）已重写为本桌面应用，详见 `更新说明.md`。
