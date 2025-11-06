# PDFQFZ

原始 WinForms 版本是一个辅助在 PDF 上批量加盖骑缝章的桌面工具：选择需要处理的 PDF 文件或文件夹，设定输出目录与印章图片，然后点击“盖章”即可生成加盖后的 PDF。

仓库现在额外提供一个跨平台的命令行实现，便于在无图形界面或自动化流水线中完成相同的盖章流程。

![PDFQFZ GUI](./pdfqfz.jpg)

## CLI 版本（PDFQFZ.CLI）
自 v1.0.2 起，官方发布包提供单文件版 `PDFQFZ.CLI.exe`（内含全部依赖）。

### 构建

```bash
dotnet build PDFQFZ.CLI/PDFQFZ.CLI.csproj
```

### 查看帮助

```bash
dotnet run --project PDFQFZ.CLI/PDFQFZ.CLI.csproj -- --help
```

### 常用参数
默认情况下 CLI 会直接覆盖输入的 PDF 文件。若需保留原件，请使用 `--output` 指定输出目录。
- `-i, --input <path>` 指定要处理的 PDF 文件或目录，搭配 `--recursive` 可递归扫描子目录。
- `-s, --stamp-image <path>` 指定印章图片，推荐使用 PNG。
- `-o, --output <dir>` （可选）将盖章后的 PDF 输出到其他目录，避免直接覆盖原文件。
- `--overwrite` 在输出目录中已存在同名文件时允许覆盖写入。
- `--suffix <text>` 配合 `--output` 使用时，为输出文件名追加后缀（默认为 `_stamped`）。
- `--size-mm <value>`：目标印章宽度（毫米），保持原始纵横比。
- `--rotation <deg>` / `--rotate-crop`：预旋转印章图片，并决定是否保留旋转后的完整边界。
- `--opacity <0-100>`：统一调整印章透明度；`--no-white-transparent` 禁用白色像素转透明。
- `--seam-*`：控制骑缝章的范围、位置、切片批次等。
- `--page-*`：控制普通印章的作用页、坐标、随机扰动及单页覆盖。
- `--encrypt-password <pwd>`：为输出文档设置打开密码。
- `--sign-mode pfx --sign-pfx <pfx> --sign-pass <pwd>`：使用现有 PFX 证书进行数字签名（目前支持 CMS 标准，暂不支持自动生成自签名证书）。

### 示例：每页右下角盖章并签名

以下示例命令会在每页右下角（X=85%，Y=20%）盖一个章，并使用无口令的 PFX 证书进行签名：

```powershell
PDFQFZ.CLI.exe `
  --input "C:\path\to\input.pdf" `
  --output "C:\path\to\output-folder" `
  --stamp-image "C:\path\to\stamp.png" `
  --page-scope all `
  --position 0.85,0.20 `
  --sign-mode pfx `
  --sign-pfx "C:\path\to\certificate.pfx" `
  --sign-pass ""
```

> 如果不需要数字签名，可省略 `--sign-*` 相关参数。

### 目录结构

- `PDFQFZ/`：原始 WinForms GUI 工程，仅供参考。
- `PDFQFZ.CLI/`：命令行项目源码。
- `pdfqfz.jpg`：GUI 截图。

### 注意事项

- `--flatten` 选项当前仅输出提示，实际不会执行页面栅格化。
- 数字签名暂不支持与输出加密（`--encrypt-password`）同时使用。
- CLI 使用 iTextSharp 和 `System.Drawing` 处理 PDF 与图像，运行环境需具备相应的原生依赖。
