"""NBC preset values and output formatting."""

from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

from scripts.formatter import PRESETS, format_document


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FOOTNOTES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)


def _east_asian_font(paragraph):
    rfonts = paragraph.runs[0]._r.rPr.find(qn("w:rFonts"))
    assert rfonts is not None
    return rfonts.get(qn("w:eastAsia"))


def _add_test_footnote(document):
    footnotes_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{W_NS}">
  <w:footnote w:id="1">
    <w:p>
      <w:pPr><w:spacing w:line="560" w:lineRule="exact"/></w:pPr>
      <w:r><w:footnoteRef/></w:r>
      <w:r><w:t>脚注 ABC</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>""".encode("utf-8")
    part = Part(
        PackURI("/word/footnotes.xml"),
        FOOTNOTES_CONTENT_TYPE,
        footnotes_xml,
        document.part.package,
    )
    document.part.relate_to(part, RT.FOOTNOTES)

    paragraph = document.add_paragraph("含脚注的正文")
    reference = OxmlElement("w:footnoteReference")
    reference.set(qn("w:id"), "1")
    paragraph.add_run()._r.append(reference)


def test_nbc_replaces_legal_builtin_preset():
    assert "nbc" in PRESETS
    assert "legal" not in PRESETS

    preset = PRESETS["nbc"]
    assert preset["name"] == "NBC格式"
    assert preset["page"] == {
        "top": 3.4,
        "bottom": 2.8,
        "left": 2.8,
        "right": 2.8,
        "header_distance": 1.5,
        "footer_distance": 1.75,
    }
    assert preset["body"]["font_cn"] == "方正仿宋_GBK"
    assert preset["body"]["size"] == 18
    assert preset["body"]["line_spacing_type"] == "multiple"
    assert preset["body"]["line_spacing"] == 1.23
    assert preset["heading2"]["font_cn"] == "方正楷体_GBK"
    assert preset["heading2"]["bold"] is True
    assert preset["heading3"]["bold"] is True
    assert preset["date"]["font_cn"] == "方正楷体_GBK"
    assert preset["date"]["bold"] is True
    assert preset["date"]["align"] == "center"
    assert preset["date"]["line_spacing"] == 1.0
    assert preset["footnote"] == {
        "font_cn": "方正仿宋_GBK",
        "font_en": "Times New Roman",
        "size": 12,
        "bold": False,
        "line_spacing_type": "multiple",
        "line_spacing": 1.0,
        "space_before": 0,
        "space_after": 0,
    }


def test_nbc_preset_formats_docx_with_reference_layout(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"

    document = Document()
    document.add_paragraph("关于NBC格式测试的通知")
    document.add_paragraph("（2026年8月13日）")
    document.add_paragraph("一、一级标题")
    document.add_paragraph("（一）二级标题")
    document.add_paragraph("1.三级标题")
    document.add_paragraph("正文段落用于验证NBC格式。")
    document.save(source)

    format_document(
        str(source), str(output),
        preset_name="nbc",
        custom_settings={"page_number": False},
    )

    result = Document(output)
    section = result.sections[0]
    assert abs(section.top_margin.cm - 3.4) < 0.02
    assert abs(section.bottom_margin.cm - 2.8) < 0.02
    assert abs(section.left_margin.cm - 2.8) < 0.02
    assert abs(section.right_margin.cm - 2.8) < 0.02
    assert abs(section.header_distance.cm - 1.5) < 0.02
    assert abs(section.footer_distance.cm - 1.75) < 0.02

    body = next(p for p in result.paragraphs if "正文段落" in p.text)
    assert _east_asian_font(body) == "方正仿宋_GBK"
    assert body.runs[0].font.size.pt == 18
    assert body.paragraph_format.line_spacing_rule == WD_LINE_SPACING.MULTIPLE
    assert abs(float(body.paragraph_format.line_spacing) - 1.23) < 0.02

    heading2 = next(p for p in result.paragraphs if p.text.startswith("（一）"))
    assert _east_asian_font(heading2) == "方正楷体_GBK"
    assert heading2.runs[0].font.bold is True

    date = next(p for p in result.paragraphs if p.text == "（2026年8月13日）")
    assert _east_asian_font(date) == "方正楷体_GBK"
    assert date.runs[0].font.size.pt == 18
    assert date.runs[0].font.bold is True
    assert date.alignment == 1
    assert date.paragraph_format.line_spacing_rule == WD_LINE_SPACING.SINGLE
    assert abs(float(date.paragraph_format.line_spacing) - 1.0) < 0.02


def test_nbc_placeholder_date_is_detected_from_reference_template(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"

    document = Document()
    document.add_paragraph("重要文稿材料格式规范")
    document.add_paragraph("（202X年X月X日）")
    document.add_paragraph("材料抬头：")
    document.save(source)

    format_document(
        str(source), str(output),
        preset_name="nbc",
        custom_settings={"page_number": False},
    )

    date = next(p for p in Document(output).paragraphs if p.text == "（202X年X月X日）")
    assert _east_asian_font(date) == "方正楷体_GBK"
    assert date.runs[0].font.bold is True
    assert date.alignment == 1


def test_nbc_formats_real_word_footnotes(tmp_path):
    source = tmp_path / "source.docx"
    output = tmp_path / "output.docx"

    document = Document()
    document.add_paragraph("NBC脚注格式测试")
    _add_test_footnote(document)
    document.save(source)

    format_document(
        str(source), str(output),
        preset_name="nbc",
        custom_settings={"page_number": False},
    )

    with ZipFile(output) as package:
        root = etree.fromstring(package.read("word/footnotes.xml"))

    namespace = {"w": W_NS}
    footnote = root.xpath('./w:footnote[@w:id="1"]', namespaces=namespace)[0]
    spacing = footnote.find(".//w:pPr/w:spacing", namespace)
    assert spacing.get(qn("w:line")) == "240"
    assert spacing.get(qn("w:lineRule")) == "auto"

    text_run = footnote.xpath('.//w:r[w:t]', namespaces=namespace)[0]
    rfonts = text_run.find("w:rPr/w:rFonts", namespace)
    assert rfonts.get(qn("w:eastAsia")) == "方正仿宋_GBK"
    assert rfonts.get(qn("w:ascii")) == "Times New Roman"
    assert rfonts.get(qn("w:hAnsi")) == "Times New Roman"
    assert text_run.find("w:rPr/w:sz", namespace).get(qn("w:val")) == "24"
