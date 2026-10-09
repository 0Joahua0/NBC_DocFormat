## 本次更新 · {{VERSION}}

- 为 **GB/T 公文标准、学术论文、NBC格式** 三个预设分别增加“编辑”按钮。
- 支持修改字体、字号、行距、页边距、表格和页码等设置，保存后重启仍有效。
- 三个预设独立保存，可恢复各自默认值；取消编辑不会写入配置。
- 仅覆盖实际修改的字段，保留 NBC 日期样式、脚注及页眉页脚距离等专用设置。
- 更新中英文使用说明，统一发布页的下载表格与运行说明。

## 下载

根据操作系统和处理器选择一个文件。下方 Assets 中的 `Source code` 为源码压缩包。

| 系统 / 处理器 | 下载文件 |
| --- | --- |
| Windows 10/11 · 64 位 | [NBC_DocFormat_windows.exe](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_windows.exe) |
| Windows 7 SP1 / 8 · 64 位 | [NBC_DocFormat_windows_win7.exe](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_windows_win7.exe) |
| Linux · x86_64（Intel / AMD 等） | [NBC_DocFormat_linux_amd64.AppImage](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_linux_amd64.AppImage) |
| Linux · ARM64（飞腾 / 鲲鹏等） | [NBC_DocFormat_linux_aarch64.AppImage](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_linux_aarch64.AppImage) |
| macOS · Intel | [NBC_DocFormat_macos_intel.dmg](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_macos_intel.dmg) |
| macOS · Apple Silicon（M 系列） | [NBC_DocFormat_macos_apple_silicon.dmg](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_macos_apple_silicon.dmg) |

## 快速开始

1. 下载并打开对应系统的程序。
2. 选择一个或多个 `.docx` 文件，选择处理模式和格式预设。
3. 如需调整预设，点击卡片上的“编辑”，修改后保存。
4. 点击“开始处理”，结果另存为新文件，原文件不会被覆盖。

## 各系统运行说明

**Windows**：双击 `.exe` 运行。Windows 7/8 使用兼容版；不支持 32 位系统。

**Linux**：先执行 `uname -m` 确认架构，再添加执行权限并运行：

```bash
chmod +x NBC_DocFormat_linux_amd64.AppImage
./NBC_DocFormat_linux_amd64.AppImage
```

ARM64 用户将文件名替换为 `NBC_DocFormat_linux_aarch64.AppImage`。遇到 FUSE 错误可在运行命令末尾加上 `--appimage-extract-and-run`。当前未提供 LoongArch 预编译包。

**macOS**：打开 `.dmg`，将应用拖入“应用程序”文件夹。首次运行若被系统拦截，可在“系统设置 → 隐私与安全性”中选择“仍要打开”。

## 文档与格式支持

- 推荐使用 `.docx`。Windows 可借助本机 Microsoft Office 或 WPS Office 转换 `.doc` / `.wps`；macOS 和 Linux 请先另存为 `.docx`。
- 排版显示依赖本机字体，请安装文档所需的中文字体。
- 所有文档均在本地处理；正式使用前请检查输出结果。

更多说明见 [中文 README](https://github.com/{{REPOSITORY}}/blob/{{VERSION}}/README.md) · [English README](https://github.com/{{REPOSITORY}}/blob/{{VERSION}}/README_EN.md)。
