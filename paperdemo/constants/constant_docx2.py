import argparse

from docx.oxml.ns import qn
from docx.shared import Inches, Cm, Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.section import WD_ORIENTATION, WD_SECTION_START


__all__ = [
    'ATTRIBUTE_DICT', 'PROPERTIES',
    'TITLE', 'AUTHOR',
]

ATTRIBUTE_DICT = {
    '页面方向-纵向': WD_ORIENTATION.PORTRAIT,
    '页面方向-横向': WD_ORIENTATION.LANDSCAPE,
    '段落对齐-左对齐': WD_ALIGN_PARAGRAPH.LEFT,
    '段落对齐-右对齐': WD_ALIGN_PARAGRAPH.RIGHT,
    '段落对齐-居中对齐': WD_ALIGN_PARAGRAPH.CENTER,
    '段落对齐-两端对齐': WD_ALIGN_PARAGRAPH.JUSTIFY,
    '段落对齐-分散对齐': WD_ALIGN_PARAGRAPH.DISTRIBUTE,
    '行间距-单倍行距': WD_LINE_SPACING.SINGLE,
    '行间距-1.5倍行距': WD_LINE_SPACING.ONE_POINT_FIVE,
    '行间距-两倍行距': WD_LINE_SPACING.DOUBLE,
    '行间距-多倍行距': WD_LINE_SPACING.MULTIPLE,
    '行间距-固定值': WD_LINE_SPACING.EXACTLY,
    '行间距-最小值': WD_LINE_SPACING.AT_LEAST,
}

PROPERTIES = {
    # doc >>>
    'settings_odd_and_even_pages_header_footer': None,
    # 奇偶页不同，默认是False，即默认奇偶页相同

    # section >>>
    'page_width': None,
    'page_height': None,
    'left_margin': None,
    'right_margin': None,
    'top_margin': None,
    'bottom_margin': None,
    'gutter': None,
    # 装订线
    'header_distance': None,
    # 页眉边距
    'footer_distance': None,
    # 页脚边距
    'orientation': None, 
    # 页面方向
    'start_type': None,
    'different_first_page_header_footer': None,
    'header_is_linked_to_previous': None,
    'footer_is_linked_to_previous': None,

    # paragraph >>>
    'alignment': None,
    
    # paragraph.paragraph_format >>>
    'paragraph_format_first_line_indent': None,
    'paragraph_format_left_indent': None,
    'paragraph_format_right_indent': None,
    'paragraph_format_space_before': None,
    'paragraph_format_space_after': None,
    'paragraph_format_line_spacing_rule': None,
    'paragraph_format_line_spacing': None,
    'paragraph_format_widow_control': None,
    # 孤行控制：防止页面顶端单独打印段落末行或页面底端单独打印段落首行
    'paragraph_format_keep_together': None,
    # 段中不分页
    'paragraph_format_keep_with_next': None,
    # 与下段同页
    'paragraph_format_page_break_before': None,
    # 段前分页

    # run.font >>>
    'font_name': None,
    'font_name_cn': None,
    'font_size': None,
    'font_bold': None,
    'font_italic': None,
    'font_shadow': None,
    # 阴影
    'font_outline': None,
    # 镂空
    'font_emboss': None,
    # 阳文
    'font_rtl': None,
    # 从右到左
    'font_underline': None,
    'font_math': None,
    # 公式格式
    'font_strike': None,
    # 删除线
    'font_double_strike': None,
    # 双删除线
    'font_highlight_color': None,
    # 突出显示颜色
    'font_superscript': None,
    # 上标
    'font_subscript': None,
    # 下标
    'font_imprint': None,
    # 印刷效果
    'font_hidden': None,
    # 隐藏
    'font_no_proof': None,
    # 忽略拼音
    'font_all_caps': None,
    # 全部大写
    'font_small_caps': None,
    # 全部小写
    'font_snap_to_grid': None,
    # 字符网格对齐
    'font_spec_vanish': None,
    # 隐藏段落标记
    'font_web_hidden': None,
    # 隐藏网格视图
}
PROPERTIES = argparse.Namespace(**PROPERTIES)

TITLE = {
    'title': '',
    'project_tag': '',
}
TITLE = argparse.Namespace(**TITLE)

AUTHOR = {
    'name': '',
    'unit_id': '',
}
AUTHOR = argparse.Namespace(**AUTHOR)

