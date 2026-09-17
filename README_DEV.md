## 核心代码阅读路线

建议按下面的调用顺序阅读，而不是从 5000 多行的界面文件逐行向下看：

```text
NBC_DocFormat.py
  main()
    → DocFormatApp.run_operation()       校验界面输入、启动工作线程
    → DocFormatApp._do_operation()       批量调度和全局进度映射
    → DocFormatApp._process_single_file()
        ├─ analyze      → scripts/analyzer.py
        ├─ punctuation  → scripts/punctuation.py
        └─ smart        → 先修复标点，再调用 scripts/formatter.py

scripts/formatter.py
  format_document()
    → detect_para_type()                 推断段落语义类型
    → format_paragraph()                 应用段落与字符格式
        → set_font()                     写入中西文字体槽
    → 表格 / 页码 / 脚注 / 中文换行规则
```

如果只研究本次 NBC 字体修复，推荐依次查看：

1. `scripts/formatter.py::set_font`
2. `scripts/punctuation.py::_set_run_font_slots`
3. `scripts/punctuation.py::_split_run_around_characters`
4. `scripts/punctuation.py::normalize_chinese_quote_fonts`
5. `scripts/punctuation.py::normalize_middle_dot_fonts`
6. `scripts/formatter.py::format_paragraph`
7. `scripts/formatter.py::format_document`

## 理解 Word 字体的最小知识

DOCX 实际是一个 ZIP 包，正文位于其中的 WordprocessingML XML。常见层级是：

```text
w:document
  └─ w:p                 段落 paragraph
      └─ w:r             字符片段 run
          ├─ w:rPr       字符格式
          └─ w:t         文本
```

一个 `w:r` 不一定只有文字，也可能含图片、域、换行、制表符或 OLE 对象；
`w:r` 还可能包在超链接、内容控件或修订节点里。因此代码不能随意执行
`paragraph.text = ...` 或删除全部 run 后重建，否则可能丢失非文本 XML。

Word 也不是只用一个字体名。`w:rFonts` 至少有四个常见字体槽：

- `eastAsia`：中文、日文、韩文；
- `ascii`：ASCII 英文与数字；
- `hAnsi`：高位 ANSI 字符，中文弯引号经常被 Word 归到这里；
- `cs`：复杂文字脚本。

这就是为什么普通中文设为方正仿宋、英文数字设为 Times New Roman 后，中文
引号仍可能错误显示成 Times New Roman：引号被 Word 选用了 `hAnsi` 槽。
当前实现先把引号拆成独立 run，再为该 run 的四个字体槽写入所在位置的中文
字体，并设置 `w:hint="eastAsia"`。由于英文和数字留在其他 run，它们仍保持
Times New Roman。字符 `·` 采用相同的隔离方式，但字体固定为方正仿宋_GBK。

## 智能处理的数据流

`smart` 模式不是一次函数调用，而是两个串联阶段：

```text
原文
  → punctuation.process_paragraph()      文本与特殊标点字体
  → 临时 .docx
  → formatter.format_document()          段落分类与整篇版式
  → 输出文件
```

对于 `.doc/.wps`，外层还会先转成临时 `.docx`，处理完成后再尝试转回旧格式。
转换和文档处理都放在后台线程；日志、进度和弹窗等界面写操作通过线程安全
helper 或 `root.after()` 返回主线程。当前仍有少量 Tk 变量的 `.get()` 发生在
工作线程，维护时不要仿照或扩大这一做法。

## 维护时需要守住的边界

- 修改 run 时要保留图片、域、超链接、内容控件和修订节点。
- 中文引号跟随标题/正文/各级标题的中文字体；`·` 固定方正仿宋_GBK。
- 英文字母和阿拉伯数字使用各预设的 `font_en`，NBC 中为 Times New Roman。
- `Document.paragraphs` 不包含表格单元格；`Document.tables` 也只公开顶层表格。
- 脚注没有完整的 python-docx 高级 API，因此直接修改 `/word/footnotes.xml`。
- OOXML 的字号、缩进、行距、边框使用不同单位，改动前先看相关 helper 注释。
- 所有核心行为改动都应运行完整测试，尤其是 `test_nbc_preset.py`、
  `test_punctuation_middle_dot.py` 和 `test_formatter_media_attachment.py`。

## 关于 custom_settings.json

仓库里这个文件是**源码运行模式（`python NBC_DocFormat.py` / `bash install.sh`）的默认配置**，
内容与代码里的 DEFAULT_CUSTOM_SETTINGS 保持一致。

### 给开发者的提示

该文件会随源码一起提交，作为源码运行模式的默认配置。**如果你在调试时通过 GUI 保存了自定义设置，
会改动本地这个文件；提交前请确认它仍然是干净的默认值**，避免把个人调试配置误提交。

如果你确实要修改这份"默认配置"（比如调整 DEFAULT_CUSTOM_SETTINGS 后想同步到这里），
请一并更新代码中的默认值和相关测试。

### 用户配置文件的实际位置

- Windows/Linux 打包发布版：exe 同目录
- macOS 打包发布版：~/Library/Application Support/NBC_DocFormat/
- 开发模式（python 直接运行）：项目根目录（即这个文件本身）
