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
    'ParagraphFormat.LineSpacing': 24,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman', # 先设置英文字体，才能设置中文字体
    'Font.NameFarEast': '宋体',
    'Font.Size': 12,
    'Font.Bold': False,
    'Font.Superscript': False,
    'Font.Subscript': False,
}

# -------------------- 页面 --------------------
pageAttr = copy(styleAttr)
pageAttr.update({
    'PageSetup.TopMargin': 3.3 * cm_to_points, # 上边距3.3厘米
    'PageSetup.BottomMargin': 3.3 * cm_to_points, # 下边距3.3厘米
    'PageSetup.LeftMargin': 2.8 * cm_to_points, # 左边距2.8厘米
    'PageSetup.RightMargin': 2.6 * cm_to_points, # 右边距2.6厘米
    'PageSetup.LayoutMode': 1, # 指定行和字符网格
    # 'PageSetup.CharsLine': 28, # 每行28个字
    # 'PageSetup.LinesPage': 22, # 每页22行，会自动设置行间距
    'PageSetup.FooterDistance': 2.8 * cm_to_points, # 页码距下边缘2.8厘米
    'PageSetup.OddAndEvenPagesHeaderFooter': 0, # 首页页码相同
    'Footers.SelfType': CONSTANTS['页码类型-除第一页外所有'],
    'Footers.LinkToPrevious': False, 
    'PageNumber.Alignment': CONSTANTS['页码对齐-居中'],
    'PageNumbers.ShowFirstPageNumber': True,
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
})

# -------------------- 参考文献 --------------------
referenceAttr = copy(styleAttr)
referenceAttr.update({
    'Reference.SelfSorting': 'Author',
    'Reference.SelfNumberStyle': '[]', 
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
})

# -------------------- 公式 --------------------
formulaAttr = copy(styleAttr)
formulaAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
})

# -------------------- 公式表格 --------------------
formulaTableAttr = copy(styleAttr)
formulaTableAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Table.PreferredWidthType': CONSTANTS['首选度量单位-百分比'],
    'Table.PreferredWidth': 100,
    'Rows.Height': 18,
    'Cells.VerticalAlignment': CONSTANTS['单元格-文字中心对齐'],
})

# -------------------- 表格 --------------------
tableAttr = copy(styleAttr)
tableAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-固定值'],
    'ParagraphFormat.LineSpacing': 18,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 12,
    'Table.PreferredWidthType': CONSTANTS['首选度量单位-百分比'],
    'Table.PreferredWidth': 100,
    'Table.LeftPadding': 0,
    'Table.RightPadding': 0,
    # 'Rows.Height': 18,
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
})

# -------------------- 图片 --------------------
graphAttr = copy(styleAttr)
graphAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'InlineShape.LockAspectRatio': -1, # 锁定纵横比
    'InlineShape.Width': 28.35 * 15.5,
    # 'InlineShape.ScaleWidth': 100, 
})

graphAttr2 = copy(styleAttr)
graphAttr2.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'InlineShape.LockAspectRatio': -1, # 锁定纵横比
    'InlineShape.Width': 28.35 * 6.3,
})

graphAttr3 = copy(styleAttr)
graphAttr3.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'InlineShape.LockAspectRatio': -1, # 锁定纵横比
    'InlineShape.Width': 28.35 * 4.7,
})

graphTableAttr = copy(styleAttr)
graphTableAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.LineSpacingRule': CONSTANTS['段落行距-最小值'],
    'ParagraphFormat.LineSpacing': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Table.PreferredWidthType': CONSTANTS['首选度量单位-百分比'],
    'Table.PreferredWidth': 80,
    'Cells.VerticalAlignment': CONSTANTS['单元格-文字中心对齐'],
    'InlineShape.LockAspectRatio': -1, # 锁定纵横比
    'InlineShape.Width': 28.35 * 6,
})

# -------------------- 注释 --------------------
noteAttr = copy(styleAttr)
noteAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
})

# # -------------------- 章节 --------------------
# section = paper.doc.sections[0]
# sectionAttr = copy(ATTRIBUTES)

# ======================================================================
# -------------------- 标题 --------------------
mainTitleAttr = copy(styleAttr)
mainTitleAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 2,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '黑体',
    'Font.Size': 16,
    'Font.Bold': True,
})

mainTitleSuperscriptAttr = copy(styleAttr)
mainTitleSuperscriptAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 2,
    'Font.Superscript': True,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '黑体',
    'Font.Size': 16,
    'Font.Bold': True,
})

mainTitleFootnoteAttr = copy(styleAttr)
mainTitleFootnoteAttr.update({
    'Footnotes.NumberingRule': CONSTANTS['注释编号规则-每页重新编号'],
    'FootnoteOptions.NumberingRule': CONSTANTS['注释编号规则-每页重新编号'],
    'FootnoteOptions.SelfSymbol': '{Symbol 61472}{Symbol 61472}',
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
})

subTitleAttr = copy(styleAttr)
subTitleAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '仿宋',
    'Font.Size': 14,
})

# -------------------- 作者 --------------------
authorAttr = copy(styleAttr)
authorAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 1,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
})

corAuthorTagAttr = copy(styleAttr)
corAuthorTagAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 1,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,    
})

# -------------------- 单位 --------------------
unitAttr = copy(styleAttr)
unitAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitBefore': 0,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 12,
})

# -------------------- 摘要 --------------------
abstractAttr = copy(styleAttr)
abstractAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
})

abstractEnAttr = copy(styleAttr)
abstractEnAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Bold': True,
})

abstractContentAttr = copy(styleAttr)
abstractContentAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
})

# -------------------- 关键词 --------------------
keywordsAttr = copy(styleAttr)
keywordsAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
})

keywordsEnAttr = copy(styleAttr)
keywordsEnAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Bold': True,
})

keywordsContentAttr = copy(styleAttr)
keywordsContentAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
})

# -------------------- 中图分类号等 --------------------
infoAttr = copy(styleAttr)
infoAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 1,
    'ParagraphFormat.CharacterUnitRightIndent': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
})

# -------------------- 基金项目 --------------------
fundAttr = copy(styleAttr)
fundAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
})

fundContentAttr = copy(styleAttr)
fundContentAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
})

# -------------------- 作者简介 --------------------
authorInfoAttr = copy(styleAttr)
authorInfoAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 2,
    'ParagraphFormat.CharacterUnitRightIndent': 2,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
})

authorInfoContentAttr1 = copy(styleAttr)
authorInfoContentAttr1.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 2,
    'ParagraphFormat.CharacterUnitRightIndent': 2,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10.5,
})

# authorInfoContentAttr2 = copy(styleAttr)
# authorInfoContentAttr2.update({
#     'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
#     'ParagraphFormat.LineUnitBefore': 0,
#     'ParagraphFormat.CharacterUnitLeftIndent': 6,
#     'Font.Name': 'Times New Roman',
#     'Font.NameFarEast': '楷体',
#     'Font.Size': 9,
# })

# -------------------- DOI --------------------
doiAttr = copy(styleAttr)
doiAttr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 2,
    'ParagraphFormat.CharacterUnitRightIndent': 2,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
})

# ======================================================================
# -------------------- 标题 --------------------
heading1Attr = copy(styleAttr)
heading1Attr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 1,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevel1,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '宋体',
    'Font.Size': 16,
    'Heading1.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading1.SelfStart': 1,
})

heading2Attr = copy(styleAttr)
heading2Attr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevel2,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '宋体',
    'Font.Size': 14,
    'Heading1.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading1.SelfStart': 1,
    'Heading2.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading2.SelfStart': 1,
})

heading3Attr = copy(styleAttr)
heading3Attr.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevel3,
    'Font.Name': 'Times New Roman', 
    'Font.NameFarEast': '宋体',
    'Font.Size': 12,
    'Heading1.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading1.SelfStart': 1,
    'Heading2.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading2.SelfStart': 1,
    'Heading3.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading3.SelfStart': 1,
})

# -------------------- 目录 --------------------
tocAttr = copy(styleAttr)
tocAttr.update({
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 12,
}) 

# -------------------- 正文 --------------------
# 正文
contentAttr1 = copy(styleAttr)
contentAttr1.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 12,
    'Font.Superscript': False,
    'Font.Subscript': False,
}) 

# 上标
contentAttr2 = copy(styleAttr)
contentAttr2.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': True,
    'Font.Subscript': False,
})

# 下标
contentAttr3 = copy(styleAttr)
contentAttr3.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': True,
})

# 表格标题
contentAttr4 = copy(styleAttr)
contentAttr4.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0.5,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
    'Heading1.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading1.SelfStart': 1,
    'Heading2.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading2.SelfStart': 1,
    'Heading3.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading3.SelfStart': 1,
})

# 表格注释
contentAttr5 = copy(styleAttr)
contentAttr5.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0.5,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 9,
    'Font.Superscript': False,
    'Font.Subscript': False,
})

# 图片标题
contentAttr6 = copy(styleAttr)
contentAttr6.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0.5,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
    'Heading1.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading1.SelfStart': 1,
    'Heading2.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading2.SelfStart': 1,
    'Heading3.SelfStyle': CONSTANTS['编号类型-8'],
    'Heading3.SelfStart': 1,
})



# 注释、参考文献
contentAttr7 = copy(styleAttr)
contentAttr7.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
}) 

# 表格内容
contentAttr8 = copy(styleAttr)
contentAttr8.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
}) 

# 正文
contentAttr9 = copy(styleAttr)
contentAttr9.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
})

# 首页脚注
contentAttr10 = copy(styleAttr)
contentAttr10.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': -2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
}) 

# 首页脚注
contentAttr11 = copy(styleAttr)
contentAttr11.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': -2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '楷体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
}) 

# 首页脚注
contentAttr12 = copy(styleAttr)
contentAttr12.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-完全两端对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitLeftIndent': 0,
    'ParagraphFormat.CharacterUnitRightIndent': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': -2,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10,
    'Font.Superscript': False,
    'Font.Subscript': False,
}) 

# 作者
contentAttr13 = copy(styleAttr)
contentAttr13.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
    'Font.Superscript': False,
    'Font.Subscript': False,
})

# 作者
contentAttr14 = copy(styleAttr)
contentAttr14.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '黑体',
    'Font.Size': 12,
    'Font.Superscript': True,
    'Font.Subscript': False,
})

# 作者
contentAttr15 = copy(styleAttr)
contentAttr15.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-居中'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 1,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 10.5,
    'Font.Superscript': False,
    'Font.Subscript': False,
})

# 图表目录
contentAttr16 = copy(styleAttr)
contentAttr16.update({
    'ParagraphFormat.Alignment': CONSTANTS['段落对齐-分散对齐'],
    'ParagraphFormat.LineUnitAfter': 0,
    'ParagraphFormat.LineUnitBefore': 0,
    'ParagraphFormat.CharacterUnitFirstLineIndent': 0,
    'ParagraphFormat.OutlineLevel': cs.wdOutlineLevelBodyText,
    'Font.Name': 'Times New Roman',
    'Font.NameFarEast': '宋体',
    'Font.Size': 12,
    'Font.Superscript': False,
    'Font.Subscript': False,
    'Toc.Width': 100, # A4 图表目录宽度
}) 

