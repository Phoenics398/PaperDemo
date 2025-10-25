"""
著作
"""


from copy import copy, deepcopy
from win32com.client import constants as cs

from paperdemo.constants.constant_docx1 import *

PUBLISHER_NAME = ''

# ======================================================================
# -------------------- 式样 --------------------
# 下面涉及到的所有设置都需要在这里设置默认值
cm_to_points = 28.35 # 1厘米为28.35磅
# 国家公文格式标准要求是上边距版心3.7cm，但是如果简单的把上边距设置为3.7cm，则因为文本的第一行本身有行距，会导致实际版心离上边缘较远，上下边距设置为3.3cm，是经过实验的，可以看看公文标准的图示。
# 版心指的是文字与边缘距离
styleAttr = {
    'PageSetup.TopMargin': 3.3 * cm_to_points, # 上边距3.3厘米
    'PageSetup.BottomMargin': 3.3 * cm_to_points, # 下边距3.3厘米
    'PageSetup.LeftMargin': 2.8 * cm_to_points, # 左边距2.8厘米
    'PageSetup.RightMargin': 2.6 * cm_to_points, # 右边距2.6厘米
    'PageSetup.LayoutMode': 1, # 指定行和字符网格
    # 'PageSetup.CharsLine': 28, # 每行28个字
    # 'PageSetup.LinesPage': 22, # 每页22行，会自动设置行间距
    'PageSetup.FooterDistance': 2.8 * cm_to_points, # 页码距下边缘2.8厘米
    'PageSetup.OddAndEvenPagesHeaderFooter': 0, # 首页页码相同
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.SpaceAfter': 0,
    'ParagraphFormat.SpaceBefore': 0,
    'ParagraphFormat.LeftIndent': 0,
    'ParagraphFormat.RightIndent': 0,
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.FirstLineIndent': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-固定值'],
    'ParagraphFormat.LineSpacing': 18,
    'Font.Name': 'Times New Roman', # 先设置英文字体，才能设置中文字体
    'Font.NameFarEast': '宋体',
    'Font.Size': 12,
    'Font.Bold': False,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

# -------------------- 页面 --------------------
pageAttr = {
    'Footers.SelfType': CONSTANTS['页码类型-除第一页外所有'],
    'Footers.LinkToPrevious': False, 
    'PageNumber.Alignment': CONSTANTS['页码对齐-居中'],
    'PageNumbers.ShowFirstPageNumber': True,
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
}

# -------------------- 参考文献 --------------------
referenceAttr = {
    'Reference.SelfSorting': 'Author',
    'Reference.SelfNumberStyle': '[]', 
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
}

# -------------------- 公式 --------------------
formulaAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
}

formulaTableAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Table.PreferredWidthType': CONSTANTS['首选度量单位-百分比'],
    'Table.PreferredWidth': 100,
    'Rows.Height': 18,
    'Cells.VerticalAlignment': CONSTANTS['单元格-文字中心对齐'],
}

# -------------------- 表格 --------------------
tableAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-固定值'],
    'ParagraphFormat.LineSpacing': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
    'Table.PreferredWidthType': CONSTANTS['首选度量单位-百分比'],
    'Table.PreferredWidth': 100,
    'Table.LeftPadding': 0,
    'Table.RightPadding': 0,
    'Rows.Height': 18,
    'Cells.VerticalAlignment': CONSTANTS['单元格-文字中心对齐'],
    # 'Cell.Borders.InsideLineStyle': CONSTANTS['线条-1'],
    # 'Cell.Borders.InsideLineWidth': CONSTANTS['线宽-4'],
    # 'Cell.Borders.OutsideLineStyle': CONSTANTS['线条-1'],
    # 'Cell.Borders.OutsideLineWidth': CONSTANTS['线宽-4'],
    'Borders.InsideLineStyle': CONSTANTS['线条-1'],
    'Borders.InsideLineWidth': CONSTANTS['线宽-4'],
    'Borders.OutsideLineStyle': CONSTANTS['线条-1'],
    'Borders.OutsideLineWidth': CONSTANTS['线宽-4'],
    'Borders.NoLeftBorder': True, # 不显示左边框
    'Borders.NoRightBorder': True, # 不显示右边框
    'Borders.NoVertical': False, # 不显示纵向框线
    'Borders.NoHorizontal': False, # 不显示横向框线
    'Border.LineStyle': CONSTANTS['线条-1'],
    'Border.LineWidth': CONSTANTS['线宽-4'],
}

# -------------------- 图片 --------------------
graphAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'InlineShape.LockAspectRatio': -1, # 锁定纵横比
    'InlineShape.Width': 28.35 * 12,
}

graphTableAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Table.PreferredWidthType': CONSTANTS['首选度量单位-百分比'],
    'Table.PreferredWidth': 80,
    'Cells.VerticalAlignment': CONSTANTS['单元格-文字中心对齐'],
    'InlineShape.LockAspectRatio': -1, # 锁定纵横比
    'InlineShape.Width': 28.35 * 6,
}

# -------------------- 注释 --------------------
noteAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
}

# # -------------------- 章节 --------------------
# section = paper.doc.sections[0]
# sectionAttr = copy(ATTRIBUTES)

# ======================================================================
# -------------------- 标题 --------------------
mainTitleAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 2,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '黑体',
    'Font.Size': 16,
    'Font.Bold': True,
}

mainTitleSuperscriptAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 2,
    'Font.Superscript': True,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '黑体',
    'Font.Size': 16,
    'Font.Bold': True,
}

mainTitleFootnoteAttr = {
    'Footnotes.NumberingRule': CONSTANTS['注释编号规则-每页重新编号'],
    'FootnoteOptions.NumberingRule': CONSTANTS['注释编号规则-每页重新编号'],
    'FootnoteOptions.SelfSymbol': '{Symbol 61472}{Symbol 61472}',
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
}

subTitleAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '仿宋',
    'Font.Size': 14,
}

# -------------------- 作者 --------------------
authorAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 1,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
}

corAuthorTagAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 1,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,    
}

# -------------------- 单位 --------------------
unitAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 12,
}

# -------------------- 摘要 --------------------
abstractAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
}

abstractEnAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Bold': True,
}

abstractContentAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
}

# -------------------- 关键词 --------------------
keywordsAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
}

keywordsEnAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Bold': True,
}

keywordsContentAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
}

# -------------------- 中图分类号等 --------------------
infoAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
}

# # -------------------- 基金项目 --------------------
# fundAttr = {
#     'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
#     'ParagraphFormat.LineUnitBefore': 0,
#     'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
#     'Font.Name': 'Times New Roman',
#     'Font.NameFarEast': '黑体',
#     'Font.Size': 10.5,
# }

# fundContentAttr = {
#     'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
#     'ParagraphFormat.LineUnitBefore': 0,
#     'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
#     'Font.Name': 'Times New Roman',
#     'Font.NameFarEast': '楷体',
#     'Font.Size': 10.5,
# }

# -------------------- 作者简介 --------------------
authorInfoAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 2,
    'ParagraphFormat.CharacterUnitRightIndent': 2,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
}

authorInfoContentAttr1 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 2,
    'ParagraphFormat.CharacterUnitRightIndent': 2,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
}

# authorInfoContentAttr2 = {
#     'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
#     'ParagraphFormat.LineUnitBefore': 0,
#     'ParagraphFormat.CharacterUnitLeftIndent': 6,
#     'Font.Name': 'Times New Roman',
#     'Font.NameFarEast': '楷体',
#     'Font.Size': 9,
# }

# -------------------- DOI --------------------
doiAttr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 2,
    'ParagraphFormat.CharacterUnitRightIndent': 2,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
}

# ======================================================================
# -------------------- 标题 --------------------
heading1Attr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0.5,
    'ParagraphFormat.LineUnitAfter': 0.5,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '宋体',
    'Font.Size': 14,
    'Heading1.SelfStyle': CONSTANTS['编号类型-1'],
    'Heading1.SelfStart': 1,
}

heading2Attr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
    'Heading1.SelfStyle': CONSTANTS['编号类型-1'],
    'Heading1.SelfStart': 1,
    'Heading2.SelfStyle': CONSTANTS['编号类型-1'],
    'Heading2.SelfStart': 1,
}

heading3Attr = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
    'Heading1.SelfStyle': CONSTANTS['编号类型-1'],
    'Heading1.SelfStart': 1,
    'Heading2.SelfStyle': CONSTANTS['编号类型-1'],
    'Heading2.SelfStart': 1,
    'Heading3.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading3.SelfStart': 1,
}

# -------------------- 正文 --------------------
# 正文
contentAttr1 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 上标
contentAttr2 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': True,
    'Font.Subscript': False,
}

# 下标
contentAttr3 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': True,
}

# 表格标题
contentAttr4 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0.5,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

# 表格注释
contentAttr5 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0.5,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

# 图片标题
contentAttr6 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0.5,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

# 注释、参考文献
contentAttr7 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 表格内容
contentAttr8 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 正文
contentAttr9 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 首页脚注
contentAttr10 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': -2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 首页脚注
contentAttr11 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': -2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 首页脚注
contentAttr12 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': -2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
} 

# 作者
contentAttr13 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

# 作者
contentAttr14 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
    'Font.Superscript': True,
    'Font.Subscript': False,
}

# 作者
contentAttr15 = {
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

