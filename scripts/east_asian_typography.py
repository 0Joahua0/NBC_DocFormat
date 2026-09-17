"""为 Word/WPS 段落写入中文避头尾与标点换行规则。

本模块不替换任何标点字符，只控制换行行为：

- ``w:kinsoku=1``：启用中文行首/行尾禁则；
- ``w:overflowPunct=0``：关闭标点悬挂到版心外侧。

这些属性没有稳定的 python-docx 高级 API，因此直接操作 ``w:pPr``。OOXML 对
段落属性子元素有固定顺序，不能无条件追加到末尾；``_PPR_ORDER`` 用来把新节点
插入兼容 Word、WPS 及严格 OOXML 消费者的位置。
"""

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


_PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap",
    "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
)
_PPR_ORDER_INDEX = {name: index for index, name in enumerate(_PPR_ORDER)}


def _insert_paragraph_property(p_pr, element, local_name):
    """按照 WordprocessingML 的 ``w:pPr`` 子元素顺序插入属性。"""
    property_index = _PPR_ORDER_INDEX[local_name]
    for index, child in enumerate(p_pr):
        child_name = child.tag.rsplit("}", 1)[-1]
        if _PPR_ORDER_INDEX.get(child_name, -1) > property_index:
            p_pr.insert(index, element)
            return
    p_pr.append(element)


def _set_paragraph_boolean_property(paragraph, local_name, value):
    """设置段落布尔属性，并返回该属性是否发生变化。

    OOXML 中节点存在通常可代表真值，但这里连假值也显式写成 ``w:val="0"``，
    防止文档模板或样式继承把关闭项重新解释为开启。
    """
    p_pr = paragraph._p.get_or_add_pPr()
    tag = qn(f"w:{local_name}")
    element = p_pr.find(tag)
    if element is None:
        element = OxmlElement(f"w:{local_name}")
        _insert_paragraph_property(p_pr, element, local_name)
    desired_value = "1" if value else "0"
    if element.get(qn("w:val")) == desired_value:
        return False
    element.set(qn("w:val"), desired_value)
    return True


def apply_chinese_line_break_rules_to_paragraph(paragraph):
    """为一个非空段启用中文禁则并关闭悬挂标点。

    返回值按“段落”计数：任一属性发生变化即为 ``True``，不是修改属性的数量。
    实现不检测段落是否真的含中文，以保证整篇版式行为一致。
    """
    if not paragraph.text.strip():
        return False
    changed = _set_paragraph_boolean_property(paragraph, "kinsoku", True)
    return _set_paragraph_boolean_property(paragraph, "overflowPunct", False) or changed


def _iter_container_paragraphs(container, seen):
    """递归产出正文/表格中的段落，并按底层 ``w:p`` 身份去重。

    合并单元格可能让同一个 ``w:tc`` 从多个网格位置出现，嵌套表格也会造成
    重复入口；``seen`` 保证每个段落只处理一次。
    """
    for paragraph in container.paragraphs:
        paragraph_id = id(paragraph._p)
        if paragraph_id not in seen:
            seen.add(paragraph_id)
            yield paragraph
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_container_paragraphs(cell, seen)


def apply_chinese_line_break_rules(document):
    """把中文换行规则应用到正文、嵌套表格及独立页眉页脚。

    链接到前一节的页眉/页脚共享相同底层内容，因此跳过，避免重复处理。属性
    已是目标值时不会计数，所以函数可安全重复调用。返回发生至少一项变化的
    段落数，而非写入的 XML 节点数。
    """
    seen = set()
    changed_count = sum(
        apply_chinese_line_break_rules_to_paragraph(paragraph)
        for paragraph in _iter_container_paragraphs(document, seen)
    )
    story_names = (
        "header", "first_page_header", "even_page_header",
        "footer", "first_page_footer", "even_page_footer",
    )
    for section in document.sections:
        for story_name in story_names:
            story = getattr(section, story_name)
            if story.is_linked_to_previous:
                continue
            changed_count += sum(
                apply_chinese_line_break_rules_to_paragraph(paragraph)
                for paragraph in _iter_container_paragraphs(story, seen)
            )
    return changed_count
