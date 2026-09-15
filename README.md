# NBC_DocFormat

面向 Word 公文材料的本地格式处理工具，支持一键处理、格式诊断、标点修复和可配置格式预设。

<p align="center">
  <img src="assets/imageforgithub.png" alt="NBC公文格式处理器" width="900">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%207%2F8%2F10%2F11%20%7C%20Linux-blue" alt="Platform">
  <img src="https://img.shields.io/badge/License-PolyForm%20Noncommercial%201.0.0-orange" alt="License">
  <img src="https://img.shields.io/badge/Language-Python-yellow" alt="Language">
</p>

## 项目来源与致谢

本项目基于原项目 [KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui) 修改开发。

原作者：[KaguraNanaga](https://github.com/KaguraNanaga)

感谢原作者提供的公文格式处理基础能力、桌面端实现和构建方案。本改版在原项目基础上进行了 NBC 格式适配、界面品牌调整和多平台客户端构建。原项目的版权与许可证声明请参见 [LICENSE](LICENSE)、[LICENSE-HISTORY.md](LICENSE-HISTORY.md) 和 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。

## 项目简介

NBC_DocFormat 用于处理 `.docx` 公文材料中的常见格式问题，包括标题层级、字体字号、段落缩进、行距、标点和表格格式。所有处理均在本地完成，不上传、不收集文档内容。

当前改版重点：

- 项目名称、软件标题和窗口标题统一为 `NBC_DocFormat`
- 内置格式预设中加入 NBC 格式
- 各级标题支持单独选择字号和加粗
- 行距支持按倍数或固定磅值设置
- 标点修复增加对字符 `·` 的检测和字体处理
- 功能按钮插图和界面主题调整为蓝色风格
- 提供 Windows 10/11、Windows 7/8 兼容、Linux x86_64、Linux ARM64 四类客户端

## 界面预览

<p align="center">
  <img src="assets/screenshot.png" alt="NBC_DocFormat软件界面截图" width="900">
</p>

## 核心能力

1. 智能一键处理：同时进行标点修复、格式统一和样式清理。
2. 格式诊断：检查文档中的格式问题，不修改原文件。
3. 标点修复：修正常见中英文标点混用问题。
4. NBC 格式预设：按 NBC 文稿材料格式要求进行排版处理。
5. 自定义格式：可配置页边距、字体、字号、加粗、行距、段前段后等参数。
6. 标题控制：一级至多级标题可分别设置字体、字号和加粗。
7. 行距控制：支持单倍/多倍行距，也支持固定磅值行距。
8. 表格处理：支持表格宽度、行高、字体和对齐方式优化。
9. 批量处理：可一次选择多个 `.docx` 文件并输出到指定目录。

## 客户端版本

构建产物位于 `NBC_DocFormat_release/` 目录：

| 系统 | 文件 | 说明 |
|---|---|---|
| Windows 10/11 | `NBC_DocFormat_windows.exe` | 64 位 Windows 10/11 推荐使用 |
| Windows 7/8 | `NBC_DocFormat_windows_win7.exe` | 64 位兼容版，建议 Windows 7 SP1 或更高版本 |
| Linux x86_64 | `NBC_DocFormat_linux_amd64.AppImage` | 适用于 Intel/AMD/兆芯/海光等 x86_64 架构 |
| Linux ARM64 | `NBC_DocFormat_linux_aarch64.AppImage` | 适用于飞腾、鲲鹏等 aarch64/ARM64 架构 |

### Windows 使用说明

下载对应 `.exe` 后双击运行即可。Windows 7/8 用户请使用 `NBC_DocFormat_windows_win7.exe`。

注意事项：

- Windows 7/8 兼容版为 64 位程序，不支持 32 位 Windows。
- Windows 7 建议使用 SP1。
- 处理 `.doc` / `.wps` 时需要本机安装 Microsoft Office 或 WPS Office；推荐优先使用 `.docx`。
- 如果 Win7/8 双击无反应或闪退，可尝试安装 Visual C++ Redistributable 2015-2022 x64。

### Linux 使用说明

先确认系统架构：

```bash
uname -m
```

选择对应 AppImage：

- 输出 `x86_64`：使用 `NBC_DocFormat_linux_amd64.AppImage`
- 输出 `aarch64` 或 `arm64`：使用 `NBC_DocFormat_linux_aarch64.AppImage`
- 输出 `loongarch64`：当前未提供预编译包，需要源码运行或单独构建

运行方式：

```bash
chmod +x NBC_DocFormat_linux_amd64.AppImage
./NBC_DocFormat_linux_amd64.AppImage
```

ARM64 用户将命令中的文件名替换为 `NBC_DocFormat_linux_aarch64.AppImage`。

银河麒麟、统信 UOS 等系统通常可直接运行对应架构的 AppImage。如遇 FUSE 相关错误，可安装系统的 `fuse` / `libfuse2` 包，或尝试：

```bash
./NBC_DocFormat_linux_amd64.AppImage --appimage-extract-and-run
```

## 使用方法

1. 点击输入区域选择一个或多个 `.docx` 文件。
2. 选择处理模式：智能一键处理、格式诊断或标点修复。
3. 选择格式预设，默认可使用 NBC 格式。
4. 需要细调时进入自定义格式设置。
5. 点击开始处理，程序会在原文件旁生成新文件，或输出到指定目录。

原文件不会被覆盖。

## 常见问题

**Win7/8 兼容版为什么还是打不开？**

请确认系统是 64 位 Windows，且 Windows 7 已安装 SP1。该兼容版使用 Python 3.8 和旧版 PyInstaller 构建，兼容性强于 Windows 10/11 版，但仍不支持 32 位系统。

**银河麒麟应该下载哪个 Linux 版本？**

在终端运行 `uname -m`。如果是 `x86_64`，下载 x86_64 版；如果是 `aarch64`，下载 ARM64 版；如果是 `loongarch64`，当前两个 AppImage 不能直接运行。

**为什么 Linux 上提示 Exec format error？**

通常是架构不匹配。请重新执行 `uname -m` 并下载对应架构的 AppImage。

**为什么处理后的文档字体不一致？**

请确认系统安装了所需中文字体，例如仿宋_GB2312、方正仿宋、黑体、楷体_GB2312 等。缺少字体时，Word 或 WPS 可能会自动替换显示字体。

**可以处理 `.doc` 或 `.wps` 吗？**

Windows 版本可借助 Microsoft Office 或 WPS Office 转换后处理。Linux 版本建议先手动另存为 `.docx`。

## 数据安全

本工具所有文档处理均在本地完成，不上传、不收集任何数据。处理结果仅供参考，建议在正式报送前人工复核。

详见 [DISCLAIMER.md](DISCLAIMER.md)。

## 使用许可

本项目基于 [KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui) 修改开发，并保留原项目许可证边界。

从 `v1.0.0` 起，本项目采用 [PolyForm Noncommercial License 1.0.0](LICENSE)：仅限个人和非商业用途免费使用、修改和按许可证要求分发；其他商业用途需要另行取得授权。原项目历史版本的许可证边界详见 [LICENSE-HISTORY.md](LICENSE-HISTORY.md)。
