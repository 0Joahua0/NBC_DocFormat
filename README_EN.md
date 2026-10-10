# NBC_DocFormat

A local Word document formatting tool for Chinese official-document workflows. It supports one-click processing, format diagnosis, punctuation repair, and configurable formatting presets.

[Download the latest release](https://github.com/0Joahua0/NBC_DocFormat/releases/latest) · [中文](README.md) · [Developer notes](README_DEV.md)

<p align="center">
  <img src="assets/imageforgithub.png" alt="NBC document format processor" width="900">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-blue" alt="Platform">
  <img src="https://img.shields.io/badge/License-PolyForm%20Noncommercial%201.0.0-orange" alt="License">
  <img src="https://img.shields.io/badge/Language-Python-yellow" alt="Language">
</p>

## Origin And Credits

This project is modified from the upstream project [KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui).

Original author: [KaguraNanaga](https://github.com/KaguraNanaga)

Thanks to the original author for the document-formatting foundation, desktop implementation, and packaging workflow. This fork adds NBC format adaptation, project rebranding, UI adjustments, and multi-platform client builds. See [LICENSE](LICENSE), [LICENSE-HISTORY.md](LICENSE-HISTORY.md), and [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md) for license and third-party notices.

## Overview

NBC_DocFormat processes common `.docx` formatting issues, including heading hierarchy, fonts, font sizes, paragraph indentation, line spacing, punctuation, and table layout. Document processing is local only; files are not uploaded or collected.

### What's new in v1.0.5

- Replaced decorative emoji in Linux UI labels with plain text to avoid a reproducible Ubuntu 22.04 / Tk 8.6.12 crash. Document content and user input are preserved. Whether this also resolves the Kylin blank-window report remains unverified.
- v1.0.4 failed the Linux pre-release tests and was not published as a Release; its changes are included below.
- Added x86_64 / ARM64 `.deb` installers with application-menu integration, window icons, and taskbar matching.
- Linux builds use a directory bundle; installed `.deb` applications no longer extract libraries into a new temporary directory on each launch.
- Adjusted the paste-text and preset-editor dialog lifecycle and added initialization, widget-state, and exception diagnostics.
- The reported blank dialogs on Kylin V11 still require verification on that system; this release does not claim a confirmed fix.
- Added a detailed [Chinese guide to format recognition](README.md#格式识别与排版逻辑), including rule priorities, thresholds, examples, and diagnostic limitations.

### What's new in v1.0.3

- Each of the three built-in presets now has an **Edit** button.
- Changes are saved independently, persist across restarts, and can be reset to each preset's defaults.
- Unchanged settings retain their original values, including NBC date styles, footnotes, and header/footer distances.
- Release download instructions are reorganized, with macOS builds documented alongside Windows and Linux.

Other features include:

- Project name, software title, and window title are unified as `NBC_DocFormat`
- NBC format preset is included
- Heading levels can configure font size and bold separately
- Line spacing can be configured by multiple or fixed point value
- Punctuation repair checks the `·` character and applies the required font handling
- Feature-card images and the UI theme are updated to a blue style
- Six client builds are available for Windows, Linux, and macOS; see the table below

## Interface Preview

<p align="center">
  <img src="assets/screenshot.png" alt="NBC_DocFormat interface screenshot" width="900">
</p>

## Download And Installation

Download the file for your system and processor from [GitHub Releases](https://github.com/0Joahua0/NBC_DocFormat/releases/latest). The `Source code` archives contain source files; use one of the client builds below to run the application directly.

| System | File | Notes |
|---|---|---|
| Windows 10/11 | `NBC_DocFormat_windows.exe` | Recommended for 64-bit Windows 10/11 |
| Windows 7/8 | `NBC_DocFormat_windows_win7.exe` | 64-bit compatible build; Windows 7 SP1 or later recommended |
| Kylin / Debian-based Linux x86_64 | `NBC_DocFormat_linux_amd64.deb` | Installer with application-menu integration |
| Debian-based Linux ARM64 | `NBC_DocFormat_linux_arm64.deb` | ARM64 installer |
| Linux x86_64 | `NBC_DocFormat_linux_amd64.AppImage` | For x86_64 systems such as Intel, AMD, Zhaoxin, and Hygon |
| Linux ARM64 | `NBC_DocFormat_linux_aarch64.AppImage` | For aarch64/ARM64 systems such as Phytium and Kunpeng |
| macOS Intel | `NBC_DocFormat_macos_intel.dmg` | For Intel-based Macs |
| macOS Apple Silicon | `NBC_DocFormat_macos_apple_silicon.dmg` | For Macs with M-series chips |

On Windows, double-click the `.exe`. Use the compatibility build for Windows 7 SP1 / 8; 32-bit Windows is not supported.

On macOS, open the `.dmg` and drag the app into Applications. If the first launch is blocked, use **System Settings → Privacy & Security → Open Anyway**.

## Usage

1. Select one or more `.docx` files.
2. Choose a mode: smart one-click processing, format diagnosis, or punctuation repair.
3. Choose a preset, such as the NBC format preset.
4. Click **编辑 (Edit)** on a preset card to adjust its settings, then click **保存设置 (Save Settings)**.
5. Start processing. The original files are not overwritten.

### Editing And Resetting Presets

- In smart processing mode, edit **GB/T official documents**, **academic papers**, or **NBC format** to adjust fonts, sizes, line spacing, margins, tables, and page numbers.
- Saving updates only that preset and persists across restarts. Other built-in and custom presets are unaffected.
- Click **恢复此预设默认值 (Restore Preset Defaults)** and save to reset a built-in preset.
- To discard changes, close the editor and choose **No** when asked to save.
- The **自定义 (Custom)** card manages separate templates, including creating, renaming, importing, and exporting presets.

Use `.docx` files on all platforms. Windows can convert `.doc` / `.wps` through a locally installed Microsoft Office or WPS Office; on macOS and Linux, save those files as `.docx` first. Install the Chinese fonts required by your documents for consistent display.

## Linux Notes

Version v1.0.5 adds `.deb` installers. Older release AppImages do not receive these changes automatically.
For Kylin V11 x86_64 or Debian-based desktops, install `NBC_DocFormat_linux_amd64.deb`
with the system package installer, or run:

```bash
sudo apt install ./NBC_DocFormat_linux_amd64.deb
```

Launch **NBC_DocFormat** from the application menu after installation. The application
and libraries stay under `/opt/nbc-docformat`, without extracting into a new temporary
directory on every launch. Normal use does not require root. Settings live in
`${XDG_CONFIG_HOME:-~/.config}/NBC_DocFormat` and are retained when upgrading or removing
the package. Install a newer `.deb` to upgrade; use `sudo apt remove nbc-docformat` to
uninstall. ARM64 users should choose `NBC_DocFormat_linux_arm64.deb`.

This package is not signed by Kylin. Source verification or execution approval may
still be required by the machine's security policy; managed machines may require
administrator approval. The app does not change security policy. Kylin-specific
compatibility and permission prompts still require validation on the target system.

For the portable AppImage, check your architecture first:

```bash
uname -m
```

Use `NBC_DocFormat_linux_amd64.AppImage` for `x86_64`, and `NBC_DocFormat_linux_aarch64.AppImage` for `aarch64` or `arm64`.

Run:

```bash
chmod +x NBC_DocFormat_linux_amd64.AppImage
./NBC_DocFormat_linux_amd64.AppImage
```

If AppImage reports a FUSE error, install `fuse` / `libfuse2`, or try:

```bash
./NBC_DocFormat_linux_amd64.AppImage --appimage-extract-and-run
```

LoongArch (`loongarch64`) is not covered by the current prebuilt AppImages.

## Data Safety

All document processing is performed locally. No document content is uploaded, collected, or tracked. Results are for reference and should be reviewed manually before formal submission.

## License

This project is modified from [KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui) and preserves the upstream license boundary.

Starting with `v1.0.0`, this project is licensed under the [PolyForm Noncommercial License 1.0.0](LICENSE). It is free for personal and noncommercial purposes. Commercial use requires separate permission. See [LICENSE-HISTORY.md](LICENSE-HISTORY.md) for the upstream license boundary.
