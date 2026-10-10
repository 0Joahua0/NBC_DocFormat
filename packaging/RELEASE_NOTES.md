## 本次更新 · {{VERSION}}

- 调整粘贴文本、自定义及预设编辑弹窗的初始化顺序，等待控件创建完成、窗口可见后再抓取输入。
- 增加弹窗初始化阶段、控件状态和错误日志；构造失败时清理残留窗口并显示错误。麒麟 V11 白框问题仍需实机复测确认。
- 补齐窗口图标及任务栏匹配信息，新增带应用菜单入口的 x86_64 / ARM64 `.deb` 安装包。
- Linux 改用目录式运行库，`.deb` 安装后不再每次启动解压到随机临时路径。
- 增加 Linux X11 弹窗回归测试、Debian 包结构验证及安装后的启动检查。
- 中文 README 新增详细的格式识别与排版逻辑说明，覆盖规则优先级、具体阈值、NBC 默认参数、完整示例及常见误判。
- 更新安装、升级、卸载和界面问题排查说明。

**麒麟 V11 用户请注意**：本版包含窗口时序调整和诊断能力，但尚未确认解决实机反馈的白框问题。
若仍出现空白窗口，即使终端没有报错，也请提供 `~/.config/NBC_DocFormat/ui_diagnostics.log`；
设置了 `XDG_CONFIG_HOME` 时，日志位于该目录下的 `NBC_DocFormat` 文件夹。
日志记录初始化步骤和控件状态，便于进一步定位。固定安装不等于麒麟官方签名，是否仍需来源认证或应用信任取决于本机策略。

## 下载

根据操作系统和处理器选择一个文件。下方 Assets 中的 `Source code` 为源码压缩包。

| 系统 / 处理器 | 下载文件 |
| --- | --- |
| Windows 10/11 · 64 位 | [NBC_DocFormat_windows.exe](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_windows.exe) |
| Windows 7 SP1 / 8 · 64 位 | [NBC_DocFormat_windows_win7.exe](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_windows_win7.exe) |
| 麒麟 / Debian 系 Linux · x86_64 安装版 | [NBC_DocFormat_linux_amd64.deb](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_linux_amd64.deb) |
| Debian 系 Linux · ARM64 安装版 | [NBC_DocFormat_linux_arm64.deb](https://github.com/{{REPOSITORY}}/releases/download/{{VERSION}}/NBC_DocFormat_linux_arm64.deb) |
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

**麒麟 / Debian 系 Linux 安装版**：x86_64 下载 `.deb` 后双击安装，或执行
`sudo apt install ./NBC_DocFormat_linux_amd64.deb`；安装后从应用菜单打开，正常运行不需要管理员权限。
ARM64 选择 `NBC_DocFormat_linux_arm64.deb`。升级时安装新版 `.deb`，卸载使用 `sudo apt remove nbc-docformat`。
系统若要求来源签名或应用信任，仍需按本机安全策略处理；安装包不保证绕过麒麟的认证提示。

**Linux 便携版**：先执行 `uname -m` 确认架构，再添加执行权限并运行：

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
