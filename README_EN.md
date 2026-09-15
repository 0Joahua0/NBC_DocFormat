# NBC_DocFormat

A local Word document formatting tool for Chinese official-document workflows. It supports one-click processing, format diagnosis, punctuation repair, and configurable formatting presets.

<p align="center">
  <img src="assets/imageforgithub.png" alt="NBC document format processor" width="900">
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%207%2F8%2F10%2F11%20%7C%20Linux-blue" alt="Platform">
  <img src="https://img.shields.io/badge/License-PolyForm%20Noncommercial%201.0.0-orange" alt="License">
  <img src="https://img.shields.io/badge/Language-Python-yellow" alt="Language">
</p>

## Origin And Credits

This project is modified from the upstream project [KaguraNanaga/docformat-gui](https://github.com/KaguraNanaga/docformat-gui).

Original author: [KaguraNanaga](https://github.com/KaguraNanaga)

Thanks to the original author for the document-formatting foundation, desktop implementation, and packaging workflow. This fork adds NBC format adaptation, project rebranding, UI adjustments, and multi-platform client builds. See [LICENSE](LICENSE), [LICENSE-HISTORY.md](LICENSE-HISTORY.md), and [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md) for license and third-party notices.

## Overview

NBC_DocFormat processes common `.docx` formatting issues, including heading hierarchy, fonts, font sizes, paragraph indentation, line spacing, punctuation, and table layout. Document processing is local only; files are not uploaded or collected.

Current changes include:

- Project name, software title, and window title are unified as `NBC_DocFormat`
- NBC format preset is included
- Heading levels can configure font size and bold separately
- Line spacing can be configured by multiple or fixed point value
- Punctuation repair checks the `·` character and applies the required font handling
- Feature-card images and the UI theme are updated to a blue style
- Four client builds are provided: Windows 10/11, Windows 7/8 compatible, Linux x86_64, and Linux ARM64

## Interface Preview

<p align="center">
  <img src="assets/screenshot.png" alt="NBC_DocFormat interface screenshot" width="900">
</p>

## Client Builds

Build outputs are placed in `NBC_DocFormat_release/`:

| System | File | Notes |
|---|---|---|
| Windows 10/11 | `NBC_DocFormat_windows.exe` | Recommended for 64-bit Windows 10/11 |
| Windows 7/8 | `NBC_DocFormat_windows_win7.exe` | 64-bit compatible build; Windows 7 SP1 or later recommended |
| Linux x86_64 | `NBC_DocFormat_linux_amd64.AppImage` | For x86_64 systems such as Intel, AMD, Zhaoxin, and Hygon |
| Linux ARM64 | `NBC_DocFormat_linux_aarch64.AppImage` | For aarch64/ARM64 systems such as Phytium and Kunpeng |

## Usage

1. Select one or more `.docx` files.
2. Choose a mode: smart one-click processing, format diagnosis, or punctuation repair.
3. Choose a preset, such as the NBC format preset.
4. Adjust custom settings when needed.
5. Start processing. The original files are not overwritten.

## Linux Notes

Check your architecture first:

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
