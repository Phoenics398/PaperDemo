from pdb import set_trace as sc

import os
import shutil
import argparse
import win32com
import win32com.client as client

from win32com.client import constants as cs


__all__ = [
    'CM_TO_POINTS', 'CONSTANTS', 'ATTRIBUTES', 
]

# fpath = "C:\\Users\\L\\AppData\\Local\\Temp\\gen_py\\3.10"
# if not os.path.isdir(fpath):
#     fpath = "C:\\Users\\Phoen\\AppData\\Local\\Temp\\gen_py\\3.10"

# for fx in os.listdir(fpath):
#     if os.path.isdir(os.path.join(fpath, fx)):
#         shutil.rmtree(os.path.join(fpath, fx))

# app = client.gencache.EnsureDispatch('KWPS.Application')
app = client.gencache.EnsureDispatch('Word.Application')
# app = client.DispatchEx('Word.Application')
app.Visible = 0

CM_TO_POINTS = 28.35 # 1厘米为28.35磅
# 1行 = 12磅 LinesToPoints()
# 1cm = 28.35磅 CentimetersToPoints()
# 1毫米 = 2.85磅 MillimetersToPoints()
# 1英寸 = 72磅 InchesToPoints()
# 像素 --> 磅 PixelsToPoints()

CONSTANTS = {
    '样式-正文字符样式': cs.wdStyleTypeCharacter,
    '样式-列表样式': cs.wdStyleTypeList,
    '样式-段落样式': cs.wdStyleTypeParagraph,
    '样式-表格样式': cs.wdStyleTypeTable,

    '段落对齐-居中': cs.wdAlignParagraphCenter,
    '段落对齐-分散对齐': cs.wdAlignParagraphDistribute,
    '段落对齐-完全两端对齐': cs.wdAlignParagraphJustify,
    '段落对齐-两端对齐高度压缩': cs.wdAlignParagraphJustifyHi,
    '段落对齐-两端对齐轻度压缩': cs.wdAlignParagraphJustifyLow,
    '段落对齐-两端对齐中度压缩': cs.wdAlignParagraphJustifyMed,
    '段落对齐-左对齐': cs.wdAlignParagraphLeft,
    '段落对齐-右对齐': cs.wdAlignParagraphRight,
    '段落对齐-两端对齐泰语布局': cs.wdAlignParagraphThaiJustify,

    '段落行距-1.5倍行距': cs.wdLineSpace1pt5,
    '段落行距-最小值': cs.wdLineSpaceAtLeast,
    '段落行距-双倍行距': cs.wdLineSpaceDouble,
    '段落行距-固定值': cs.wdLineSpaceExactly,
    '段落行距-多倍行距': cs.wdLineSpaceMultiple,
    '段落行距-单倍行距': cs.wdLineSpaceSingle,

    # '页面方向-纵向': WD_ORIENTATION.PORTRAIT,
    # '页面方向-横向': WD_ORIENTATION.LANDSCAPE,

    '页码类型-除第一页外所有': cs.wdHeaderFooterPrimary,
    # 返回文档或节中除第一页外所有页上的页眉或页脚。
    '页码类型-第一页': cs.wdHeaderFooterFirstPage,
    # 返回文档或节中的第一个页眉或页脚。
    '页码类型-偶数页': cs.wdHeaderFooterEvenPages,
    # 返回偶数页上的所有页眉或页脚。
    
    '页码对齐-居中': cs.wdAlignPageNumberCenter,
    '页码对齐-页脚内部左对齐': cs.wdAlignPageNumberInside,
    '页码对齐-左对齐': cs.wdAlignPageNumberLeft,
    '页码对齐-页脚外部右对齐': cs.wdAlignPageNumberOutside,
    '页码对齐-右对齐': cs.wdAlignPageNumberRight,

    '页码样式-阿拉伯语': cs.wdPageNumberStyleArabic,
    '页码样式-阿拉伯语全角': cs.wdPageNumberStyleArabicFullWidth,
    '页码样式-阿拉伯语字母1': cs.wdPageNumberStyleArabicLetter1,
    '页码样式-阿拉伯语字母2': cs.wdPageNumberStyleArabicLetter2,
    '页码样式-朝鲜文汉字读取': cs.wdPageNumberStyleHanjaRead,
    '页码样式-朝鲜文汉字读取数字': cs.wdPageNumberStyleHanjaReadDigit,
    '页码样式-希伯来语字母1': cs.wdPageNumberStyleHebrewLetter1,
    '页码样式-希伯来语字母2': cs.wdPageNumberStyleHebrewLetter2,
    '页码样式-印地语阿拉伯语': cs.wdPageNumberStyleHindiArabic,
    '页码样式-印地语基数文本': cs.wdPageNumberStyleHindiCardinalText,
    '页码样式-印地语字母1': cs.wdPageNumberStyleHindiLetter1,
    '页码样式-印地语字母2': cs.wdPageNumberStyleHindiLetter2,
    '页码样式-日语汉字': cs.wdPageNumberStyleKanji,
    '页码样式-日语汉字数字': cs.wdPageNumberStyleKanjiDigit,
    '页码样式-日语汉字传统': cs.wdPageNumberStyleKanjiTraditional,
    '页码样式-小写字母': cs.wdPageNumberStyleLowercaseLetter,
    '页码样式-小写罗马': cs.wdPageNumberStyleLowercaseRoman,
    '页码样式-带圈数字': cs.wdPageNumberStyleNumberInCircle,
    '页码样式-带划线数字': cs.wdPageNumberStyleNumberInDash,
    '页码样式-简体中文数字1': cs.wdPageNumberStyleSimpChinNum1,
    '页码样式-简体中文数字2': cs.wdPageNumberStyleSimpChinNum2,
    '页码样式-泰语阿拉伯语': cs.wdPageNumberStyleThaiArabic,
    '页码样式-泰语基数文本': cs.wdPageNumberStyleThaiCardinalText,
    '页码样式-泰语字母': cs.wdPageNumberStyleThaiLetter,
    '页码样式-繁体中文数字1': cs.wdPageNumberStyleTradChinNum1,
    '页码样式-繁体中文数字2': cs.wdPageNumberStyleTradChinNum2,
    '页码样式-大写字母': cs.wdPageNumberStyleUppercaseLetter,
    '页码样式-大写罗马': cs.wdPageNumberStyleUppercaseRoman,
    '页码样式-越南语基数文本': cs.wdPageNumberStyleVietCardinalText,

    '首选度量单位-自动选择': cs.wdPreferredWidthAuto,
    '首选度量单位-百分比': cs.wdPreferredWidthPercent,
    '首选度量单位-磅数': cs.wdPreferredWidthPoints,

    '单元格-文字上框线对齐': cs.wdCellAlignVerticalTop,
    '单元格-文字中心对齐': cs.wdCellAlignVerticalCenter,
    '单元格-文字下框线对齐': cs.wdCellAlignVerticalBottom,

    '线条-0': cs.wdLineStyleNone,
    # 无边框。
    '线条-1': cs.wdLineStyleSingle,
    # 单实线
    '线条-2': cs.wdLineStyleDot,
    # 点。
    '线条-3': cs.wdLineStyleDashSmallGap,
    # 划线后跟小间隙。
    '线条-4': cs.wdLineStyleDashLargeGap,
    # 划线后跟大间隙。
    '线条-5': cs.wdLineStyleDashDot,
    # 划线后跟点。
    '线条-6': cs.wdLineStyleDashDotDot,
    # 划线后跟两个点。
    '线条-7': cs.wdLineStyleDouble,
    # 双实线。
    '线条-8': cs.wdLineStyleTriple,
    # 三条细实线。
    '线条-9': cs.wdLineStyleThinThickSmallGap,
    # 里面是一条细实线，外面是一条粗实线，两条线的间隙较小。
    '线条-10': cs.wdLineStyleThickThinSmallGap,
    # 里面是一条粗实线，外面是一条细实线，两条线的间隙较小。
    '线条-11': cs.wdLineStyleThinThickThinSmallGap,
    # 最里面一条细实线，其次一条粗实线，最外面是一条细实线，所有线之间的间隙较小。
    '线条-12': cs.wdLineStyleThinThickMedGap,
    # 里面是一条细实线，外面是一条粗实线，两条线的间隙中等。
    '线条-13': cs.wdLineStyleThickThinMedGap,
    # 里面是一条粗实线，外面是一条细实线，两条线的间隙中等。
    '线条-14': cs.wdLineStyleThinThickThinMedGap,
    # 最里面是一条细实线，其次是一条粗实线，最外面是一条细实线，所有线之间的间隙中等。
    '线条-15': cs.wdLineStyleThinThickLargeGap,
    # 里面是一条细实线，外面是一条粗实线，两条线的间隙较大。
    '线条-16': cs.wdLineStyleThickThinLargeGap,
    # 里面是一条粗实线，外面是一条细实线，两条线的间隙较大。
    '线条-17': cs.wdLineStyleThinThickThinLargeGap,
    # 最里面是一条细实线，其次是一条粗实线，最外面是一条细实线，所有线之间的间隙较大。
    '线条-18': cs.wdLineStyleSingleWavy,
    # 波浪型单实线。
    '线条-19': cs.wdLineStyleDoubleWavy,
    # 波浪型双实线。
    '线条-20': cs.wdLineStyleDashDotStroked,
    # 划线后跟粗点，使边框的外观类似于理发店招牌。
    '线条-21': cs.wdLineStyleEmboss3D,
    # 边框呈现三维阳文效果。
    '线条-22': cs.wdLineStyleEngrave3D,
    # 边框呈现三维阴文效果。
    '线条-23': cs.wdLineStyleOutset,
    # 边框呈现凸起效果。
    '线条-24': cs.wdLineStyleInset,
    # 边框呈现凹进效果。

    '线宽-1': cs.wdLineWidth025pt, 
    # 0.25 磅。
    '线宽-2': cs.wdLineWidth050pt, 
    # 0.50 磅。
    '线宽-3': cs.wdLineWidth075pt, 
    # 0.75 磅。
    '线宽-4': cs.wdLineWidth100pt, 
    # 1.00 磅。 默认值。
    '线宽-5': cs.wdLineWidth150pt, 
    # 1.50 磅。
    '线宽-6': cs.wdLineWidth225pt, 
    # 2.25 磅。
    '线宽-7': cs.wdLineWidth300pt, 
    # 3.00 磅。
    '线宽-8': cs.wdLineWidth450pt, 
    # 4.50 磅。
    '线宽-9': cs.wdLineWidth600pt, 
    # 6.00 磅。

    '边框线-1': cs.wdBorderDiagonalUp, 
    # 从左下角开始的对角线边框。
    '边框线-2': cs.wdBorderDiagonalDown, 
    # 从左上角开始的对角线边框。
    '边框线-3': cs.wdBorderVertical, 
    # 纵向框线。
    '边框线-4': cs.wdBorderHorizontal, 
    # 横向框线。
    '边框线-5': cs.wdBorderRight, 
    # 右侧框线。
    '边框线-6': cs.wdBorderBottom, 
    # 底边框线。
    '边框线-7': cs.wdBorderLeft, 
    # 左侧框线。
    '边框线-8': cs.wdBorderTop, 
    # 上框线。

    '注释编号规则-连续分配编号': cs.wdRestartContinuous,
    '注释编号规则-每节重新编号': cs.wdRestartSection,
    '注释编号规则-每页重新编号': cs.wdRestartPage,

    '编号类型-1': "CHINESENUM1",
    # 一二三四
    '编号类型-2': "CHINESENUM2",
    # 壹贰叁肆
    '编号类型-3': "CHINESENUM3",
    # 一二三四
    '编号类型-4': "ALPHABETIC",
    # ABCD
    '编号类型-5': "alphabetic",
    # abcd
    '编号类型-6': "ROMAN",
    # IIIIIIIV
    '编号类型-7': "roman",
    # iiiiiiiv
    '编号类型-8': "ARABIC",
    # 1234
    '编号类型-9': "Arabic",
    # 1234
    '编号类型-10': "ZODIAC1",
    # 甲乙丙丁
    '编号类型-11': "ZODIAC2",
    # 子丑寅卯
    '编号类型-12': "ZODIAC3",
    # 甲子乙丑丙寅丁卯
    '编号类型-13': "GB1",
    # ⒈⒉⒊⒋
    '编号类型-14': "GB2",
    # ⑴⑵⑶⑷
    '编号类型-15': "GB3",
    # ①②③④
    '编号类型-16': "GB4",
    # ㈠㈡㈢㈣
}

ATTRIBUTES = {
    # Document >>>
    'Document._CodeName': None,
    # 仅供内部使用。
    'Document.ActiveTheme': None,
    # 返回活动主题名称以及主题格式选项为指定的文档。
    'Document.ActiveThemeDisplayName': None,
    # 返回指定文档的活动主题显示名称。
    'Document.ActiveWindow': None,
    # 返回表示 Window 活动窗口的 对象。
    'Document.ActiveWritingStyle[Object]': None,
    # 返回或设置指定文档中指定语言的写作风格。
    'Document.Application': None,
    # 返回一个Application对象，该对象表示 Microsoft Word 应用程序。
    'Document.AttachedTemplate': None,
    # 返回一个 Template 对象，该对象表示附加到指定文档的模板。
    'Document.AutoFormatOverride': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表是否自动设置格式替代格式设置限制的文档中的格式设置限制已生效。
    'Document.AutoHyphenation': None,
    # 确定是否为指定文档启用自动断字。
    'Document.Background': None,
    # 返回一个 Shape 对象，该对象代表指定文档的背景图像。
    'Document.Bibliography': None,
    # 返回文档中包含的书目引用。 此为只读属性。
    'Document.Bookmarks': None,
    # 返回一个 Bookmarks 集合，该集合代表文档中的所有书签。
    'Document.Broadcast': None,
    # 返回一个 Broadcast 对象，该对象代表广播会话，其中演示者可以通过 Web 向远程参与者呈现Word文档，而无需参与者安装丰富的客户端。
    'Document.BuiltInDocumentProperties': None,
    # 返回一个 DocumentProperties 集合，该集合表示指定文档的所有内置文档属性。
    'Document.Characters': None,
    # 返回一个 Characters 集合，该集合代表文档中的字符。
    'Document.ChartDataPointTrack': None,
    # 返回或设置 C#) 中的 布尔 (布尔 值，指定活动文档中的图表是否使用单元格引用数据点跟踪。 读写。
    'Document.ChildNodeSuggestions': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.ClickAndTypeParagraphStyle': None,
    # 返回或设置“即点即输”功能在指定文档中应用于文字的默认段落样式。
    'Document.CoAuthoring': None,
    # 返回一个 CoAuthoring 对象，该对象提供文档中共同创作相关对象模型的入口点。
    'Document.CodeName': None,
    # 返回指定文档的代码名称。
    'Document.CommandBars': None,
    # 返回一个CommandBars集合，该集合代表菜单栏和 Microsoft Word中的所有工具栏。
    'Document.Comments': None,
    # 返回一个 Comments 集合，该集合表示指定文档中的所有注释。
    'Document.Compatibility[WdCompatibility]': None,
    # 确定是否启用指定的兼容性选项。
    'Document.CompatibilityMode': None,
    # 返回一个 long 类型的值，该值指定 Word 2010 在打开文档时使用的兼容模式。
    'Document.ConsecutiveHyphensLimit': None,
    # 返回或设置能够以连字符结尾的连续行的最大数目。
    'Document.Container': None,
    # 返回包含指定 OLE 对象的容器应用程序的对象。
    'Document.Content': None,
    # 返回一个 Range 对象，该对象代表main文档文章。
    'Document.ContentControls': None,
    # 返回文档中的所有内容控件。 此为只读属性。
    'Document.ContentTypeProperties': None,
    # 返回存储在文档中的元数据，例如作者姓名、主题和公司。 此为只读属性。
    'Document.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Document.CurrentRsid': None,
    # 返回一个随机数，Word分配给文档中的更改。 此为只读属性。
    'Document.CustomDocumentProperties': None,
    # 返回一个 DocumentProperties 集合，该集合代表指定文档的所有自定义文档属性。
    'Document.CustomXMLParts': None,
    # 返回一个 CustomXMLParts#SameCHM 集合，该集合表示 XML 数据存储中的自定义 XML。 此为只读属性。
    'Document.DefaultTableStyle': None,
    # 返回一个 Object 类型的 值，该值代表应用于文档中所有新创建的表的表格样式。
    'Document.DefaultTabStop': None,
    # 返回或设置指定文档中默认制表位之间的间隔 (以磅为单位)。
    'Document.DefaultTargetFrame': None,
    # 返回或设置一个 字符串 ，表示用于显示通过超链接可到达网页的浏览器框架。
    'Document.DisableFeatures': None,
    # 确定是否禁用属性中指定的 DisableFeaturesIntroducedAfter 版本之后引入的所有功能。
    'Document.DisableFeaturesIntroducedAfter': None,
    # 禁用的 Microsoft Word 仅在文档中指定的版本之后引入的所有功能。
    'Document.DocID': None,
    # 仅供内部使用。
    'Document.DocumentInspectors': None,
    # 返回一个 DocumentInspectors 集合，使你能够查找隐藏的个人信息，例如作者姓名、公司名称和修订日期。 此为只读属性。
    'Document.DocumentLibraryVersions': None,
    # 返回一个 DocumentLibraryVersions 集合，该集合表示已启用版本控制且存储在服务器上的文档库中的共享文档版本的集合。
    'Document.DocumentTheme': None,
    # 返回一个 OfficeTheme 对象，该对象表示应用于文档的 Microsoft Office 主题。 此为只读属性。
    'Document.DoNotEmbedSystemFonts': None,
    # 确定 Microsoft Word是否嵌入常见系统字体。
    'Document.Email': None,
    # 返回一个 Email 对象，该对象包含当前文档的所有电子邮件相关属性。
    'Document.EmbedLinguisticData': None,
    # 确定 Microsoft Word是否嵌入语音和手写、存储东亚 IME 击键以及控制从设备接收的文本服务数据。
    'Document.EmbedSmartTags': None,
    # 确定 Microsoft Word是否将智能标记信息保存在文档中。
    'Document.EmbedTrueTypeFonts': None,
    # 如果 Microsoft Word在保存文档时将 TrueType 字体嵌入文档中，则返回True。
    'Document.EncryptionProvider': None,
    # 返回一个 String 类型的值，指定 Microsoft Office Word在加密文档时使用的算法加密提供程序的名称。 读/写。
    'Document.Endnotes': None,
    # 返回一个 Endnotes 集合，该集合代表区域、选定内容或文档中的所有尾注。
    'Document.EnforceStyle': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表是否在受保护的文档中实施格式设置限制。
    'Document.Envelope': None,
    # 返回一个 Envelope 对象，该对象表示指定文档中的信封功能和信封。
    'Document.FarEastLineBreakLanguage': None,
    # 返回或设置在指定的文档或模板中换行文本时使用的东亚语言。
    'Document.FarEastLineBreakLevel': None,
    # 返回或设置指定文档的行中断控制级别。
    'Document.Fields': None,
    # 返回一个只读 Fields 集合，该集合代表文档、区域或选定内容中的所有字段。
    'Document.Final': None,
    # 返回或设置 Boolean 值，该值指示文档是否是最终的。 读/写。
    'Document.Footnotes': None,
    # 返回一个 Footnotes 集合，该集合代表区域、选定内容或文档中的所有脚注。
    'Document.FormattingShowClear': None,
    # 确定 Microsoft Word 是否在“样式和格式”任务窗格中显示清除格式。
    'Document.FormattingShowFilter': None,
    # 返回或设置一个 WdShowFilter 常量，该常量代表“样式和格式设置”任务窗格中显示的样式和格式。
    'Document.FormattingShowFont': None,
    # 确定 Microsoft Word是否在“样式和格式”任务窗格中显示字体格式。
    'Document.FormattingShowNextLevel': None,
    # 返回或设置一个布尔值，该值代表使用上一个标题级别时，Microsoft Office Word是否显示下一个标题级别。 读/写。
    'Document.FormattingShowNumbering': None,
    # 确定 Microsoft Word在“样式和格式设置”任务窗格中是否显示数字格式。
    'Document.FormattingShowParagraph': None,
    # 确定 Microsoft Word是否在“样式和格式设置”任务窗格中显示段落格式。
    'Document.FormattingShowUserStyleName': None,
    # 返回或设置一个 boolean 类型的值 ，该值表示是否显示用户定义的样式。 读/写。
    'Document.FormFields': None,
    # 返回一个 FormFields 集合，该集合代表文档、区域或选定内容中的所有窗体字段。
    'Document.FormsDesign': None,
    # 如果指定的文档处于窗体设计模式，则返回 True 。
    'Document.Frames': None,
    # 返回一个 Frames 集合，该集合代表文档、区域或选定内容中的所有框架。
    'Document.Frameset': None,
    # 返回一个 Frameset 对象，该对象表示整个框架页或框架页上的单个框架。
    'Document.FullName': None,
    # 指定文档、模板或级联样式表的名称，包括驱动器或 Web 路径。
    'Document.GrammarChecked': None,
    # 确定是否已在指定的区域或文档上运行语法检查。
    'Document.GrammaticalErrors': None,
    # 返回一个ProofreadingErrors集合，该集合表示对指定文档或区域检查语法检查失败的句子。
    'Document.GridDistanceHorizontal': None,
    # 返回或设置 Microsoft Word在指定文档中绘制、移动自选图形或东亚字符以及调整其大小时所使用的不可见网格线之间的水平间距量。
    'Document.GridDistanceVertical': None,
    # 返回或设置 Microsoft Word在指定文档中绘制、移动自选图形或东亚字符和调整其大小时使用的不可见网格线之间的垂直间距量。
    'Document.GridOriginFromMargin': None,
    # 确定 Microsoft Word是否从页面左上角启动字符网格。
    'Document.GridOriginHorizontal': None,
    # 返回或设置相对于页面左边缘的点，你希望在指定文档中开始绘制、移动和调整自选图形或东亚字符的不可见网格。
    'Document.GridOriginVertical': None,
    # 返回或设置相对于页面顶部的点，你希望在指定文档中开始绘制、移动和调整自选图形或东亚字符的不可见网格。
    'Document.GridSpaceBetweenHorizontalLines': None,
    # 返回或设置 Microsoft Word 在页面视图中显示水平字符网格线的间隔。
    'Document.GridSpaceBetweenVerticalLines': None,
    # 返回或设置 Microsoft Word 在页面视图中显示垂直字符网格线的间隔。
    'Document.HasMailer': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.HasPassword': None,
    # 如果打开指定文档需要密码，则返回 True 。
    'Document.HasRoutingSlip': None,
    # 确定指定的文档是否附加了传送单。
    'Document.HasVBProject': None,
    # 返回一个 boolean 类型的值 ，该值代表文档是否具有附加的 Microsoft Visual Basic for Applications 项目。 此为只读属性。
    'Document.HTMLDivisions': None,
    # 返回一个 HTMLDivisions 对象，该对象代表 Web 文档中的 HTML 除法。
    'Document.HTMLProject': None,
    # 返回HTMLProject指定文档中表示顶级项目分支的对象，如Microsoft 脚本编辑器的项目资源管理器中所示。
    'Document.Hyperlinks': None,
    # 返回一个 Hyperlinks 集合，该集合代表指定文档、区域或选定内容中的所有超链接。
    'Document.HyphenateCaps': None,
    # 确定所有大写字母中的单词是否可以连字符。
    'Document.HyphenationZone': None,
    # 返回或设置断字区域的宽度，以磅为单位。
    'Document.Indexes': None,
    # 返回一个 Indexes 集合，该集合表示指定文档中的所有索引。
    'Document.InlineShapes': None,
    # 返回一个 InlineShapes 集合，该集合代表文档、区域或选定内容中的所有 InlineShape 对象。
    'Document.IsInAutosave': None,
    # 如此 如果最近触发的 Application.DocumentBeforeSave 事件 (Word) 事件是自动保存的结果，而不是用户手动保存的结果。 此为只读属性。
    'Document.IsMasterDocument': None,
    # 确定指定的文档是否为主控文档。
    'Document.IsSubdocument': None,
    # 确定是否在单独的文档窗口中打开指定的文档作为主控文档的子文档。
    'Document.JustificationMode': None,
    # 返回或设置指定文档字符间距的调整量。
    'Document.KerningByAlgorithm': None,
    # 确定 Microsoft 是否Word指定文档中的半角拉丁字符和标点符号。
    'Document.Kind': None,
    # 返回或设置 Microsoft Word 在自动设置指定文档的格式时使用的格式类型。
    'Document.LanguageDetected': None,
    # 返回或设置一个值，指定 Word 是否已检测到指定的文本的语言。
    'Document.ListParagraphs': None,
    # 返回一个 ListParagraphs 集合，该集合代表文档中的所有编号段落。
    'Document.Lists': None,
    # 返回一个 Lists 集合，该集合包含指定文档中的所有带格式列表。
    'Document.ListTemplates': None,
    # 返回一个 ListTemplates 集合，该集合代表指定文档的所有列表格式。
    'Document.LockQuickStyleSet': None,
    # 返回或设置一个 Boolean，它代表用户是否可以更改使用的快速样式集。 读/写。
    'Document.LockTheme': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表用户是否可以更改文档主题。 读/写。
    'Document.MailEnvelope': None,
    # 返回一个 MsoEnvelope 对象，该对象代表文档的电子邮件标题。
    'Document.Mailer': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.MailMerge': None,
    # 返回一个 MailMerge 对象，该对象表示指定文档的邮件合并功能。
    'Document.Name': None,
    # 返回指定对象的名称。
    'Document.NoLineBreakAfter': None,
    # 返回或设置行的首尾字符，该字符后，Microsoft Word 将不会中断。
    'Document.NoLineBreakBefore': None,
    # 返回或设置行的首尾字符，该字符之前的 Microsoft Word 不会中断。
    'Document.OMathBreakBin': None,
    # 返回或设置一个WdOMathBreakBin枚举值，该值表示当公式跨越两行或更多行时，Microsoft Office Word放置二进制运算符的位置。 读/写。
    'Document.OMathBreakSub': None,
    # 返回或设置一个WdOMathBreakSub枚举值，该值表示 Microsoft Office Word如何处理位于换行符之前的减法运算符。 读/写。
    'Document.OMathFontName': None,
    # 返回文档中用于显示公式的字体的名称。 读/写。
    'Document.OMathIntSubSupLim': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表积分的限制的默认位置。 读/写。
    'Document.OMathJc': None,
    # 返回或设置一组公式的默认理由（左、右、居中或居中为一组）。 读/写。
    'Document.OMathLeftMargin': None,
    # 返回或设置公式的左边距。 读/写。
    'Document.OMathNarySupSubLim': None,
    # 返回或设置一个 Boolean 类型的 值，该值代表除整型以外的对象限制 n的默认位置。 读/写。
    'Document.OMathRightMargin': None,
    # 返回或设置公式的右边距。 读/写。
    'Document.OMaths': None,
    # 返回 OMath 指定范围内的 对象。 此为只读属性。
    'Document.OMathSmallFrac': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表是否在文档所包含的公式中使用小型分数。 读/写。
    'Document.OMathWrap': None,
    # 返回换行到新行的公式的第二行的位置。 读/写。
    'Document.OpenEncoding': None,
    # 返回用于打开指定的文档的编码。
    'Document.OptimizeForWord97': None,
    # 确定 Microsoft Word是否优化当前文档以供在 Word 97 中查看。
    'Document.OriginalDocumentTitle': None,
    # 运行合法黑线文档比较函数后，返回原始文档的文档标题。 此为只读属性。
    'Document.PageSetup': None,
    # 返回与 PageSetup 指定文档关联的 对象。
    'Document.Paragraphs': None,
    # 返回一个 Paragraphs 集合，该集合表示指定文档中的所有段落。
    'Document.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Document.Password': None,
    # 设置打开指定的文档时必须提供一个密码。
    'Document.PasswordEncryptionAlgorithm': None,
    # 返回一个 字符串 ，表示 Microsoft Word 在用密码对文档加密时使用的算法。
    'Document.PasswordEncryptionFileProperties': None,
    # 如果 Microsoft Word加密受密码保护的文档的文件属性，则返回 True。
    'Document.PasswordEncryptionKeyLength': None,
    # 返回一个 Integer 类型的值，指示 Microsoft Word使用密码加密文档时使用的算法的密钥长度。
    'Document.PasswordEncryptionProvider': None,
    # 返回 Microsoft Word使用密码加密文档时使用的算法加密提供程序的名称。
    'Document.Path': None,
    # 返回指定对象的磁盘或 Web 路径。
    'Document.Permission': None,
    # 返回一个 Permission 对象，该对象表示指定文档中的权限设置。
    'Document.PrintFormsData': None,
    # 确定 Microsoft Word是否仅打印在相应联机表单中输入的数据打印到预打印的表单上。
    'Document.PrintFractionalWidths': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.PrintPostScriptOverText': None,
    # 确定使用 PostScript 打印机时，是否在文本和图形顶部打印 PRINT 字段指令 (（如文档中的 PostScript 命令) ）。
    'Document.PrintRevisions': None,
    # 确定是否随文档一起打印修订标记。
    'Document.ProtectionType': None,
    # 返回指定文档的保护类型。
    'Document.ReadabilityStatistics': None,
    # 返回一个 ReadabilityStatistics 集合，该集合表示指定文档的可读性统计信息。
    'Document.ReadingLayoutSizeX': None,
    # 返回或设置一个 Integer 类型的值，该值代表文档在阅读布局视图中显示并冻结以输入手写标记时该文档页的宽度。
    'Document.ReadingLayoutSizeY': None,
    # 返回或设置一个 Integer 类型的值，该值代表文档在阅读版式视图中显示并冻结以输入手写标记时该文档页的高度。
    'Document.ReadingModeLayoutFrozen': None,
    # 设置或返回一个 boolean 类型的值 ，该值代表是否在阅读版式视图中显示的页面冻结为指定大小以向文档插入手写的标记。
    'Document.ReadOnly': None,
    # 确定是否无法将文档的更改保存到原始文档。
    'Document.ReadOnlyRecommended': None,
    # 确定用户打开文档时，Word是否显示一个消息框，并建议将其以只读形式打开。
    'Document.RemoveDateAndTime': None,
    # 返回或设置一个 布尔值 ，指示文档是否存储所跟踪更改的日期和时间元数据。
    'Document.RemovePersonalInformation': None,
    # 确定 Microsoft Word在保存文档时是否从批注、修订和“属性”对话框中删除所有用户信息。
    'Document.Research': None,
    # 返回文档的研究服务。 此为只读属性。
    'Document.RevisedDocumentTitle': None,
    # 运行法律黑线文档比较函数后，返回修订后文档的文档标题。 此为只读属性。
    'Document.Revisions': None,
    # 返回一个 Revisions 集合，该集合代表文档或区域中的修订。
    'Document.Routed': None,
    # 确定指定的文档是否已路由到下一个收件人。
    'Document.RoutingSlip': None,
    # 返回一个 RoutingSlip 对象，该对象表示指定文档的传送单信息。
    'Document.Saved': None,
    # 确定指定的文档或模板自上次保存以来是否未更改。
    'Document.SaveEncoding': None,
    # 返回或设置保存文档时要使用的编码。
    'Document.SaveFormat': None,
    # 返回一个 Integer 类型的值，该值代表指定文档或文件转换器的文件格式。
    'Document.SaveFormsData': None,
    # 确定 Microsoft Word是否将表单中输入的数据保存为制表符分隔的记录，以便在数据库中使用。
    'Document.SaveSubsetFonts': None,
    # 确定 Microsoft Word是否将嵌入的 TrueType 字体的子集与文档一起保存。
    'Document.Scripts': None,
    # 返回一个 Scripts 集合，该集合表示指定对象中的 HTML 脚本集合。
    'Document.Sections': None,
    # 返回一个 Sections 集合，该集合代表指定文档中的节。
    'Document.Sentences': None,
    # 返回一个 Sentences 集合，该集合代表文档中的所有句子。
    'Document.ServerPolicy': None,
    # 返回为运行 Microsoft Office SharePoint Server 2007 的服务器上存储的文档指定的策略。 此为只读属性。
    'Document.Shapes': None,
    # 返回一个 Shapes 集合，该集合表示指定文档中的所有 Shape 对象。
    'Document.SharedWorkspace': None,
    # 返回一个 SharedWorkspace 对象，该对象表示指定文档所在的文档工作区。
    'Document.ShowGrammaticalErrors': None,
    # 确定语法错误是否由指定文档中的波浪绿线标记。
    'Document.ShowRevisions': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.ShowSpellingErrors': None,
    # 确定 Microsoft Word是否为文档中的拼写错误下划线。
    'Document.ShowSummary': None,
    # 确定是否显示指定文档的自动摘要。
    'Document.Signatures': None,
    # 返回一个 SignatureSet 对象，该对象代表文档的数字签名。
    'Document.SmartDocument': None,
    # 返回一个 SmartDocument 对象，该对象表示智能文档解决方案的设置。
    'Document.SmartTags': None,
    # 返回一个 SmartTags 对象，该对象代表文档中的智能标记。
    'Document.SmartTagsAsXMLProps': None,
    # 确定当包含智能标记的文档保存为 HTML 时，Microsoft Word是否创建包含智能标记信息的 XML 标头。
    'Document.SnapToGrid': None,
    # 确定在指定文档中绘制、移动或调整其大小时，自选图形或东亚字符是否自动与不可见网格对齐。
    'Document.SnapToShapes': None,
    # 确定 Microsoft Word是否自动将自选图形或东亚字符与不可见网格线对齐，这些网格线穿过指定文档中其他自选图形或东亚字符的垂直和水平边缘。
    'Document.SpellingChecked': None,
    # 确定是否已在整个指定区域或文档中检查拼写。
    'Document.SpellingErrors': None,
    # 返回一个 ProofreadingErrors 集合，该集合表示在指定文档或区域中标识为拼写错误的单词。
    'Document.StoryRanges': None,
    # 返回一个 StoryRanges 集合，该集合表示指定文档中的所有文章。
    'Document.Styles': None,
    # 返回 Styles 指定文档的集合。
    'Document.StyleSheets': None,
    # 返回一个 StyleSheets 对象，该对象表示附加到文档的 Web 样式表。
    'Document.StyleSortMethod': None,
    # 返回或设置在“样式”任务窗格中对样式进行排序时使用的排序方法。 读/写。
    'Document.Subdocuments': None,
    # 返回一个 Subdocuments 集合，该集合代表指定区域或文档中的所有子文档。
    'Document.SummaryLength': None,
    # 返回或设置摘要长度与文档长度的百分比。
    'Document.SummaryViewMode': None,
    # 返回或设置摘要的显示方式。
    'Document.Tables': None,
    # 返回一个 Tables 集合，该集合代表指定文档中的所有表。
    'Document.TablesOfAuthorities': None,
    # 返回一个 TablesOfAuthorities 集合，该集合代表指定文档中的引文表。
    'Document.TablesOfAuthoritiesCategories': None,
    # 返回一个 TablesOfAuthoritiesCategories 集合，该集合代表指定文档的可用引文类别表。
    'Document.TablesOfContents': None,
    # 返回一个 TablesOfContents 集合，该集合表示指定文档中的目录。
    'Document.TablesOfFigures': None,
    # 返回一个 TablesOfFigures 集合，该集合代表指定文档中的图表表。
    'Document.TextEncoding': None,
    # 返回或设置的代码页或字符集，则 Word 使用另存为编码的文本文件的文档。
    'Document.TextLineEnding': None,
    # 返回或设置一个WdLineEndingType常量，该常量指示 Microsoft Word如何在保存为文本文件的文档中标记换行符和段落分隔符。
    'Document.TrackFormatting': None,
    # 返回或设置一个 Boolean 类型的 值，该值代表在打开更改跟踪时是否跟踪格式更改。 读/写。
    'Document.TrackMoves': None,
    # 返回或设置一个Boolean 类型的 值，该值代表在打开“修订”时是否标记移动的文本。 读/写。
    'Document.TrackRevisions': None,
    # 确定是否在指定文档中跟踪更改。
    'Document.Type': None,
    # 返回文档的类型（模板或文档）。
    'Document.UpdateStylesOnOpen': None,
    # 确定是否更新指定文档中的样式，以匹配每次打开文档时附加模板中的样式。
    'Document.UseMathDefaults': None,
    # 返回或设置一个 boolean 类型的值 ，该值表示是否在创建新公式时使用默认数学设置。 读/写。
    'Document.UserControl': None,
    # 确定文档或应用程序是否由用户创建或打开。
    'Document.Variables': None,
    # 返回一个 Variables 集合，该集合表示存储在指定文档中的变量。
    'Document.VBASigned': None,
    # 确定是否对指定文档的Visual Basic for Applications (VBA) 项目进行了数字签名。
    'Document.VBProject': None,
    # 返回的 VBProject 对象所指定的模板或文档。
    'Document.Versions': None,
    # 返回一个 Versions 集合，该集合表示指定文档的所有版本。
    'Document.WebOptions': None,
    # 返回 WebOptions 对象，其中包含 Microsoft Word在将文档另存为网页或打开网页时使用的文档级属性。
    'Document.Windows': None,
    # 返回一个 Windows 集合，该集合代表指定文档的所有窗口 (例如，Sales.doc：1 和 Sales.doc：2) 。
    'Document.WordOpenXML': None,
    # 返回文档Word Open XML 内容的平面 XML 格式。 此为只读属性。
    'Document.Words': None,
    # 返回一个 Words 集合，该集合代表文档中的所有单词。
    'Document.WritePassword': None,
    # 该属性设置一个保存对指定文档所做的修改时所需的密码。
    'Document.WriteReserved': None,
    # 确定是否使用写入密码保护指定的文档。
    'Document.XMLHideNamespaces': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.XMLNodes': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.XMLSaveDataOnly': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.XMLSaveThroughXSLT': None,
    # 返回或设置一个 String 类型的 值，该值指定在用户保存文档时要应用的可扩展样式表语言转换 (XSLT) 的路径和文件名。
    'Document.XMLSchemaReferences': None,
    # 返回一个 XMLSchemaReferences 集合，该集合表示附加到文档的架构。
    'Document.XMLSchemaViolations': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Document.XMLShowAdvancedErrors': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表是否从内置的 Word 错误消息或 Office 附带的 Microsoft XML Core Services (MSXML) (MSXML) 5.0 组件生成错误消息文本。
    'Document.XMLUseXSLTWhenSaving': None,
    # 返回一个 boolean 类型的值 ，该值表示是否要保存文档通过可扩展样式表语言转换 (XSLT)。

    # PageSetup >>>
    'PageSetup.Application': None,
    # 返回一个Application对象，该对象表示 Microsoft Word 应用程序。
    'PageSetup.BookFoldPrinting': None,
    # 真正 的 Microsoft Word 对文档进行打印在一系列的小册子以便打印好的页面可以折叠和作为书籍阅读。
    'PageSetup.BookFoldPrintingSheets': None,
    # 返回或设置一个 Integer 类型的值，该值代表每个小册子的页数。
    'PageSetup.BookFoldRevPrinting': None,
    # 真 为 Microsoft Word 以逆序打印的书籍折页的双向或亚洲语言文档的打印。
    'PageSetup.BottomMargin': None,
    # 返回或设置页面的下边缘与正文文本边界之间的距离 (以磅为单位)。
    'PageSetup.CharsLine': None,
    # 返回或设置文档网格中每行的字符数。
    'PageSetup.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'PageSetup.DifferentFirstPageHeaderFooter': None,
    # 如此 如果不同的页眉或页脚使用第一页上。 可 真 、 假 或 wdUndefined 。
    'PageSetup.FirstPageTray': None,
    # 返回或设置用于文档或节的第一页的纸盒。
    'PageSetup.FooterDistance': None,
    # 返回或设置页脚和页面底部之间的距离 (以磅为单位)。
    'PageSetup.Gutter': None,
    # 返回或设置文档或节的每一页的额外页边距的大小（以磅为单位），以备装订需要。
    'PageSetup.GutterOnTop': None,
    # 仅供内部使用。
    'PageSetup.GutterPos': None,
    # 返回或设置文档中的装订线位于哪一侧。
    'PageSetup.GutterStyle': None,
    # 返回或设置是否 Microsoft Word 使用从右到左或从左到右语言基于当前文档的装订线。
    'PageSetup.HeaderDistance': None,
    # 返回或设置页眉与页面顶端之间的距离（以磅为单位）。
    'PageSetup.LayoutMode': None,
    # 返回或设置当前文档的布局模式。
    'PageSetup.LeftMargin': None,
    # 返回或设置页面左边的缘与正文文本的左边的界之间的距离 (以磅为单位)。
    'PageSetup.LineNumbering': None,
    # 返回或设置 对象， LineNumbering 该对象代表指定 PageSetup 对象的行号。
    'PageSetup.LinesPage': None,
    # 返回或设置文档网格中每页的行数。
    'PageSetup.MirrorMargins': None,
    # 如此 如果对开页的外侧页边距等宽。 可 真 、 假 或 wdUndefined 。
    'PageSetup.OddAndEvenPagesHeaderFooter': None,
    # 如此 如果 指定的 PageSetup 对象具有不同的页眉和页脚的奇数和偶数页。 可 真 、 假 或 wdUndefined 。
    'PageSetup.Orientation': None,
    # 返回或设置页面的方向。
    'PageSetup.OtherPagesTray': None,
    # 返回或设置用于文档或节中除第一页以外其他所有页的纸盒。
    'PageSetup.PageHeight': None,
    # 返回或设置页面高度（以磅为单位）。
    'PageSetup.PageWidth': None,
    # 返回或设置页面的宽度，以磅为单位。
    'PageSetup.PaperSize': None,
    # 返回或设置纸张大小。
    'PageSetup.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'PageSetup.RightMargin': None,
    # 返回或设置页面的右边缘与正文文本的右边界之间的距离 (以磅为单位)。
    'PageSetup.SectionDirection': None,
    # 返回或设置阅读顺序和对齐方式指定节。
    'PageSetup.SectionStart': None,
    # 返回或设置指定对象的分节符的类型。
    'PageSetup.ShowGrid': None,
    # 此成员仅适用于 Macintosh 并且不应使用。
    'PageSetup.SuppressEndnotes': None,
    # 如此 如果下一个没有隐藏尾注的节的末尾打印尾注。 在该节的尾注之前打印隐藏的尾注。
    'PageSetup.TextColumns': None,
    # 返回一个 TextColumns 集合，该集合代表指定 PageSetup 对象的一组文本列。
    'PageSetup.TopMargin': None,
    # 返回或设置页面的上边缘与正文文本上部边界之间的距离 (以磅为单位)。
    'PageSetup.TwoPagesOnOne': None,
    # 如此 如果 Microsoft Word 打印指定的文档的两页。
    'PageSetup.VerticalAlignment': None,
    # 返回或设置文档或节中每页文本的垂直对齐方式。

    # Paragraphs >>>
    'Paragraphs.AddSpaceBetweenFarEastAndAlpha': None,
    # 确定是否将 Microsoft Word设置为在指定段落的日语和拉丁文本之间自动添加空格。
    'Paragraphs.AddSpaceBetweenFarEastAndDigit': None,
    # 确定是否将 Microsoft Word 设置为在指定段落的日语文本和数字之间自动添加空格。
    'Paragraphs.Alignment': None,
    # 返回或设置一个 WdParagraphAlignment 常量，该常量表示指定段落的对齐方式。
    'Paragraphs.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Paragraphs.AutoAdjustRightIndent': None,
    # 确定是否将 Microsoft Word设置为自动调整指定段落的右缩进（如果每行指定了一组字符数）。
    'Paragraphs.BaseLineAlignment': None,
    # 返回或设置一个 WdBaselineAlignment 常量，该常量表示字体在行上的垂直位置。
    'Paragraphs.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Paragraphs.CharacterUnitFirstLineIndent': None,
    # 返回或设置首行或悬挂缩进的值 (以字符为单位)。
    'Paragraphs.CharacterUnitLeftIndent': None,
    # 返回或设置指定段落的左缩进值 (以字符为单位)。
    'Paragraphs.CharacterUnitRightIndent': None,
    # 该属性返回或设置指定段落的右缩进量（以字符为单位）。
    'Paragraphs.Count': None,
    # 返回指定集合中的项数。
    'Paragraphs.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Paragraphs.DisableLineHeightGrid': None,
    # 确定当每页指定一组行数时，Microsoft Word是否将指定段落中的字符与线条网格对齐。
    'Paragraphs.FarEastLineBreakControl': None,
    # 确定 Microsoft Word是否对指定的段落应用东亚换行规则。
    'Paragraphs.First': None,
    # 返回一个 Paragraph 对象，该对象代表集合中的 Paragraphs 第一项。
    'Paragraphs.FirstLineIndent': None,
    # 返回或设置首行的行或悬挂缩进的值 (以磅为单位)。
    'Paragraphs.Format': None,
    # 返回或设置一个 ParagraphFormat 对象，该对象表示指定段落或段落的格式设置。
    'Paragraphs.HalfWidthPunctuationOnTopOfLine': None,
    # 确定 Microsoft Word是否将行开头的标点符号更改为指定段落的半角字符。
    'Paragraphs.HangingPunctuation': None,
    # 确定是否为指定的段落启用了悬挂标点符号。
    'Paragraphs.Hyphenation': None,
    # 确定指定的段落是否包含在自动断字中。
    'Paragraphs.Item[Int32]': None,
    # 返回集合中的单个对象。
    'Paragraphs.KeepTogether': None,
    # 确定当 Microsoft Word重新分类文档时，指定段落中的所有行是否都保留在同一页上。
    'Paragraphs.KeepWithNext': None,
    # 确定当 Microsoft Word重新分类文档时，指定段落是否与后面的段落保持在同一页上。
    'Paragraphs.Last': None,
    # 以 对象的形式返回集合中的 Paragraphs 最后一项 Paragraph 。
    'Paragraphs.LeftIndent': None,
    # 返回或设置 一个 值，表示指定段落的左缩进值 (以磅为单位)。
    'Paragraphs.LineSpacing': None,
    # 返回或设置指定段落的行距 (以磅为单位)。
    'Paragraphs.LineSpacingRule': None,
    # 返回或设置指定段落的行距。
    'Paragraphs.LineUnitAfter': None,
    # 返回或设置指定段落的段后间距 (以网格线)。
    'Paragraphs.LineUnitBefore': None,
    # 返回或设置指定段落的段前间距 (以网格线) 的数量。
    'Paragraphs.NoLineNumber': None,
    # 确定是否为指定段落抑制行号。
    'Paragraphs.OutlineLevel': None,
    # 返回或设置指定段落的大纲级别。
    'Paragraphs.PageBreakBefore': None,
    # 确定是否在指定段落之前强制分页符。
    'Paragraphs.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Paragraphs.ReadingOrder': None,
    # 返回或设置指定段落的读取次序而不改变其对齐方式。
    'Paragraphs.RightIndent': None,
    # 返回或设置指定段落的右缩进量（以磅为单位）。
    'Paragraphs.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Paragraphs.SpaceAfter': None,
    # 返回或设置指定段落或文本栏后面的间距 (以磅为单位) 的数量。
    'Paragraphs.SpaceAfterAuto': None,
    # 确定 Microsoft Word是否自动设置指定段落后的间距量。
    'Paragraphs.SpaceBefore': None,
    # 返回或设置指定段落的段前间距 (以磅为单位)。
    'Paragraphs.SpaceBeforeAuto': None,
    # 确定 Microsoft Word是否自动设置指定段落之前的间距量。
    'Paragraphs.Style': None,
    # 返回或设置指定对象的样式。
    'Paragraphs.TabStops': None,
    # 返回或设置一个 TabStops 集合，该集合代表指定段落的所有自定义制表位。
    'Paragraphs.WidowControl': None,
    # 确定当 Microsoft Word重新分类文档时，指定段落中的第一行和最后一行是否与段落的其余部分保持在同一页上。
    'Paragraphs.WordWrap': None,
    # 确定 Microsoft Word是否在指定段落或文本框架中的单词中间换行拉丁文文本。

    # Paragraph >>>
    'Paragraph.AddSpaceBetweenFarEastAndAlpha': None,
    # 确定是否将 Microsoft Word设置为在指定段落的日语和拉丁文本之间自动添加空格。
    'Paragraph.AddSpaceBetweenFarEastAndDigit': None,
    # 确定是否将 Microsoft Word 设置为在指定段落的日语文本和数字之间自动添加空格。
    'Paragraph.Alignment': None,
    # 返回或设置一个 WdParagraphAlignment 常量，该常量表示指定段落的对齐方式。
    'Paragraph.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Paragraph.AutoAdjustRightIndent': None,
    # 确定是否将 Microsoft Word设置为自动调整指定段落的右缩进（如果每行指定了一组字符数）。
    'Paragraph.BaseLineAlignment': None,
    # 返回或设置一个 WdBaselineAlignment 常量，该常量表示字体在行上的垂直位置。
    'Paragraph.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Paragraph.CharacterUnitFirstLineIndent': None,
    # 返回或设置首行或悬挂缩进的值 (以字符为单位)。
    'Paragraph.CharacterUnitLeftIndent': None,
    # 返回或设置指定段落的左缩进值 (以字符为单位)。
    'Paragraph.CharacterUnitRightIndent': None,
    # 该属性返回或设置指定段落的右缩进量（以字符为单位）。
    'Paragraph.CollapsedState': None,
    # 返回或设置指定的段落当前是否处于折叠状态。 在 C#) 中读/写 布尔 (布尔 值。
    'Paragraph.CollapseHeadingByDefault': None,
    # 返回或设置在文档加载时是否默认折叠指定的段落。 在 C#) 中读/写 布尔 (布尔 值。
    'Paragraph.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Paragraph.DisableLineHeightGrid': None,
    # 确定当每页指定一组行数时，Microsoft Word是否将指定段落中的字符与线条网格对齐。
    'Paragraph.DropCap': None,
    # 返回一个 DropCap 对象，该对象代表指定段落的下沉大写字母。
    'Paragraph.FarEastLineBreakControl': None,
    # 确定 Microsoft Word是否对指定的段落应用东亚换行规则。
    'Paragraph.FirstLineIndent': None,
    # 返回或设置首行的行或悬挂缩进的值 (以磅为单位)。
    'Paragraph.Format': None,
    # 返回或设置一个 ParagraphFormat 对象，该对象表示指定段落或段落的格式设置。
    'Paragraph.HalfWidthPunctuationOnTopOfLine': None,
    # 确定 Microsoft Word是否将行开头的标点符号更改为指定段落的半角字符。
    'Paragraph.HangingPunctuation': None,
    # 确定是否为指定的段落启用了悬挂标点符号。
    'Paragraph.Hyphenation': None,
    # 确定指定的段落是否包含在自动断字中。
    'Paragraph.ID': None,
    # 返回或设置指定对象的标识标签当前文档保存为 Web 页时。
    'Paragraph.IsStyleSeparator': None,
    # 确定段落是否包含允许 Microsoft Word出现以联接不同段落样式的段落的特殊隐藏段落标记。
    'Paragraph.KeepTogether': None,
    # 确定当 Microsoft Word重新分类文档时，指定段落中的所有行是否都保留在同一页上。
    'Paragraph.KeepWithNext': None,
    # 确定当 Microsoft Word重新分类文档时，指定段落是否与后面的段落保持在同一页上。
    'Paragraph.LeftIndent': None,
    # 返回或设置一个 Single 类型的值，该值代表指定段落、表格行或 HTML 除法) (左缩进值。
    'Paragraph.LineSpacing': None,
    # 返回或设置指定段落的行距 (以磅为单位)。
    'Paragraph.LineSpacingRule': None,
    # 返回或设置指定段落的行距。
    'Paragraph.LineUnitAfter': None,
    # 返回或设置指定段落的段后间距 (以网格线)。
    'Paragraph.LineUnitBefore': None,
    # 返回或设置指定段落的段前间距 (以网格线) 的数量。
    'Paragraph.ListNumberOriginal[Int16]': None,
    # 返回一个 Integer 类型的值，该值代表段落的原始列表级别。 此为只读属性。
    'Paragraph.MirrorIndents': None,
    # 返回或设置一个 Integer 类型的值，该值代表左缩进和右缩进的宽度是否相同。 可以为 True、 False 或 wdUndefined。 读/写。
    'Paragraph.NoLineNumber': None,
    # 确定是否为指定段落抑制行号。
    'Paragraph.OutlineLevel': None,
    # 返回或设置指定段落的大纲级别。
    'Paragraph.PageBreakBefore': None,
    # 确定是否在指定段落之前强制分页符。
    'Paragraph.ParaID': None,
    # 仅供内部使用。
    'Paragraph.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Paragraph.Range': None,
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'Paragraph.ReadingOrder': None,
    # 返回或设置指定段落的读取次序而不改变其对齐方式。
    'Paragraph.RightIndent': None,
    # 返回或设置指定段落的右缩进量（以磅为单位）。
    'Paragraph.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Paragraph.SpaceAfter': None,
    # 返回或设置指定段落或文本栏后面的间距 (以磅为单位) 的数量。
    'Paragraph.SpaceAfterAuto': None,
    # 确定 Microsoft Word是否自动设置指定段落后的间距量。
    'Paragraph.SpaceBefore': None,
    # 返回或设置指定段落的段前间距 (以磅为单位)。
    'Paragraph.SpaceBeforeAuto': None,
    # 确定 Microsoft Word是否自动设置指定段落之前的间距量。
    'Paragraph.Style': None,
    # 返回或设置指定对象的样式。
    'Paragraph.TabStops': None,
    # 返回或设置一个 TabStops 集合，该集合代表指定段落的所有自定义制表位。
    'Paragraph.TextboxTightWrap': None,
    # 返回或设置一个 WdTextboxTightWrap 常量，该常量表示文本环绕形状或文本框的紧密程度。 读/写。
    'Paragraph.TextID': None,
    # 仅供内部使用。
    'Paragraph.WidowControl': None,
    # 确定当 Microsoft Word重新分类文档时，指定段落中的第一行和最后一行是否与段落的其余部分保持在同一页上。
    'Paragraph.WordWrap': None,
    # 确定 Microsoft Word是否在指定段落或文本框架中的单词中间换行拉丁文文本。

    # Range >>>
    'Range.Application': None,
    # 返回一个 Application 对象，该对象表示Microsoft Word应用程序。
    'Range.Bold': None,
    # 确定字体或区域的格式是否为粗体。
    'Range.BoldBi': None,
    # 确定字体或区域的格式是否为粗体。
    'Range.BookmarkID': None,
    # 返回包含指定选定内容或区域开头的书签编号;如果没有相应的书签，则返回 0 (零) 。
    'Range.Bookmarks': None,
    # 返回一个 Bookmarks 集合，该集合代表区域中的所有书签。
    'Range.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Range.CanEdit': None,
    # 仅供内部使用。
    'Range.CanPaste': None,
    # 仅供内部使用。
    'Range.Case': None,
    # 返回或设置一个 WdCharacterCase 常量，该常量代表指定区域中的文本大小写。
    'Range.Cells': None,
    # 返回一个 Cells 集合，该集合代表区域中的表格单元格。
    'Range.Characters': None,
    # 返回一个 Characters 集合，该集合代表区域中的字符。
    'Range.CharacterStyle': None,
    # 返回一个 Object 类型的 值，该值代表用于设置一个或多个字符格式的样式。 只读。
    'Range.CharacterWidth': None,
    # 返回或设置指定区域的字符宽度。
    'Range.Columns': None,
    # 返回一个 Columns 集合，该集合代表区域中的所有表列。
    'Range.CombineCharacters': None,
    # 确定指定的区域是否包含组合字符。
    'Range.Comments': None,
    # 返回一个 Comments 集合，该集合代表指定区域中的所有注释。
    'Range.Conflicts': None,
    # 获取一个 Conflicts 集合对象，该对象包含范围中的所有冲突对象。
    'Range.ContentControls': None,
    # 返回一个 ContentControls 集合，该集合表示范围中包含的内容控件。 只读。
    'Range.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Range.DisableCharacterSpaceGrid': None,
    # 确定Microsoft Word是否忽略范围每行的字符数。
    'Range.Document': None,
    # 返回与 Document 指定区域关联的 对象。
    'Range.Duplicate': None,
    # 返回一个 Range 对象，该对象表示指定区域的所有属性。
    'Range.Editors': None,
    # 返回一个 Editors 对象，该对象表示有权修改文档中所选内容或区域的所有用户。
    'Range.EmphasisMark': None,
    # 返回或设置字符或指定的字符字符串的着重号。
    'Range.End': None,
    # 返回或设置某区域中结束字符的位置。
    'Range.EndnoteOptions': None,
    # 返回一个 EndnoteOptions 对象，该对象代表区域或选定内容中的尾注。
    'Range.Endnotes': None,
    # 返回一个 Endnotes 集合，该集合代表区域中的所有尾注。
    'Range.EnhMetaFileBits': None,
    # 返回所选内容或文本范围的显示方式的图片表示形式。
    'Range.Fields': None,
    # 返回一个只读 Fields 集合，该集合代表区域中的所有字段。
    'Range.Find': None,
    # 返回一个 Find 对象，该对象包含查找操作的条件。
    'Range.FitTextWidth': None,
    # 返回或设置当前度量单位) 宽度 (，其中Microsoft Word适合当前范围内的文本。
    'Range.Font': None,
    # 返回或设置一个 Font 对象，该对象表示指定对象的字符格式设置。
    'Range.FootnoteOptions': None,
    # 返回一个 FootnoteOptions 对象，该对象代表区域中的脚注选项。
    'Range.Footnotes': None,
    # 返回一个 Footnotes 集合，该集合代表区域中的所有脚注。
    'Range.FormattedText': None,
    # 返回或设置一个 Range 对象，该对象包含指定区域或所选内容中的带格式文本。
    'Range.FormFields': None,
    # 返回一个 FormFields 集合，该集合代表区域中的所有窗体字段。
    'Range.Frames': None,
    # 返回一个 Frames 集合，该集合代表范围中的所有帧。
    'Range.GrammarChecked': None,
    # 确定是否已在指定范围内运行语法检查。
    'Range.GrammaticalErrors': None,
    # 返回一个ProofreadingErrors集合，该集合表示在指定范围上语法检查失败的句子。
    'Range.HighlightColorIndex': None,
    # 返回或设置指定区域的突出显示颜色。
    'Range.HorizontalInVertical': None,
    # 返回或设置水平垂直文本中的文本的格式。
    'Range.HTMLDivisions': None,
    # 返回一个 HTMLDivisions 对象，该对象代表 Web 文档中的 HTML 除法。
    'Range.Hyperlinks': None,
    # 返回一个 Hyperlinks 集合，该集合代表指定范围中的所有超链接。
    'Range.ID': None,
    # 返回或设置指定对象的标识标签当前文档保存为 Web 页时。
    'Range.Information[WdInformation]': None,
    # 返回有关指定选择或范围的信息。
    'Range.InlineShapes': None,
    # 返回一个 InlineShapes 集合，该集合代表文档、区域或选定内容中的所有 InlineShape 对象。
    'Range.IsEndOfRowMark': None,
    # 确定指定的区域是否折叠，并且是否位于表中的行尾标记处。
    'Range.Italic': None,
    # 确定区域的格式是否为斜体。
    'Range.ItalicBi': None,
    # 确定区域的格式是否为斜体。
    'Range.Kana': None,
    # 返回或设置日文文本的指定区域是平假名还是片假名。
    'Range.LanguageDetected': None,
    # 返回或设置一个值，指定 Word 是否已检测到指定的文本的语言。
    'Range.LanguageID': None,
    # 返回或设置指定对象的语言。
    'Range.LanguageIDFarEast': None,
    # 返回或设置指定对象的东亚语言。
    'Range.LanguageIDOther': None,
    # 返回或设置指定对象的语言。
    'Range.ListFormat': None,
    # 返回一个 ListFormat 对象，该对象表示区域的所有列表格式特征。
    'Range.ListParagraphs': None,
    # 返回一个 ListParagraphs 集合，该集合代表区域中的所有编号段落。
    'Range.ListStyle': None,
    # 返回一个 Object 类型的 值，该值代表用于设置项目符号列表或编号列表的格式。 只读。
    'Range.Locks': None,
    # 获取一个 CoAuthLocks 集合对象，该对象代表范围中的所有锁。
    'Range.NextStoryRange': None,
    # 返回一个 Range 对象，该对象引用下一篇文章，如下表所示。
    'Range.NoProofing': None,
    # 确定拼写和语法检查器是否忽略指定的文本。
    'Range.OMaths': None,
    # 返回一个 OMaths 集合，该集合代表 OMath 指定范围内的 对象。 只读。
    'Range.Orientation': None,
    # 返回或设置范围中文字的方向，当启用了文字方向功能。
    'Range.PageSetup': None,
    # 返回与 PageSetup 指定区域关联的 对象。
    'Range.ParagraphFormat': None,
    # 返回或设置一个 ParagraphFormat 对象，该对象代表指定区域的段落设置。
    'Range.Paragraphs': None,
    # 返回一个 Paragraphs 集合，该集合代表指定区域中的所有段落。
    'Range.ParagraphStyle': None,
    # 返回一个 Object 类型的值，该值代表用于设置段落格式的样式。 只读。
    'Range.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Range.ParentContentControl': None,
    # 返回一个 ContentControl 对象，该对象代表指定区域的父内容控件。 只读。
    'Range.PreviousBookmarkID': None,
    # 返回与指定区域位于同一位置或之前的最后一个书签的编号。
    'Range.ReadabilityStatistics': None,
    # 返回一个 ReadabilityStatistics 集合，该集合表示指定区域的可读性统计信息。
    'Range.Revisions': None,
    # 返回一个 Revisions 集合，该集合代表区域中的跟踪更改。
    'Range.Rows': None,
    # 返回一个 Rows 集合，该集合代表区域中的所有表行。
    'Range.Scripts': None,
    # 返回一个 Scripts 集合，该集合表示指定对象中的 HTML 脚本集合。
    'Range.Sections': None,
    # 返回一个 Sections 集合，该集合代表指定区域中的节。
    'Range.Sentences': None,
    # 返回一个 Sentences 集合，该集合代表范围中的所有句子。
    'Range.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Range.ShapeRange': None,
    # 返回一个 ShapeRange 集合，该集合代表指定范围中的所有 Shape 对象。
    'Range.ShowAll': None,
    # 确定是否显示所有非打印字符 (，例如隐藏文本、制表符、空格标记和段落标记) 。
    'Range.SmartTags': None,
    # 返回一个 SmartTags 对象，该对象代表区域中的智能标记。
    'Range.SpellingChecked': None,
    # 确定是否已在整个指定范围内检查拼写。
    'Range.SpellingErrors': None,
    # 返回一个 ProofreadingErrors 集合，该集合表示指定区域中标识为拼写错误的单词。
    'Range.Start': None,
    # 返回或设置范围的起始字符位置。
    'Range.StoryLength': None,
    # 返回包含指定的区域的文章中的字符数。
    'Range.StoryType': None,
    # 返回指定范围的故事类型。
    'Range.Style': None,
    # 返回或设置指定对象的样式。
    'Range.Subdocuments': None,
    # 返回一个 Subdocuments 集合，该集合代表指定范围中的所有子文档。
    'Range.SynonymInfo': None,
    # 返回一个 SynonymInfo 对象，该对象包含同义词库中有关同义词、反义词或指定字词或短语的相关字词和表达式的信息。
    'Range.Tables': None,
    # 返回一个 Tables 集合，该集合代表指定区域中的所有表。
    'Range.TableStyle': None,
    # 返回一个 Object 类型的值，该值代表用于设置表格格式的样式。 只读。
    'Range.Text': None,
    # 返回或设置指定区域中的文本。
    'Range.TextRetrievalMode': None,
    # 返回一个 TextRetrievalMode 对象，该对象控制如何从指定区域检索文本。
    'Range.TextVisibleOnScreen': None,
    # 返回 C#) 中的 整数 (int ，指示指定区域中的文本是否在屏幕上可见。 只读。
    'Range.TopLevelTables': None,
    # 返回一个 Tables 集合，该集合表示当前区域中最外层嵌套级别的表。
    'Range.TwoLinesInOne': None,
    # 返回或设置 Microsoft Word 设置在一个两行文本并指定将文本括起来的字符，如果有的话。
    'Range.Underline': None,
    # 返回或设置应用于区域的下划线类型。
    'Range.Updates': None,
    # 获取一个 CoAuthUpdates 集合对象，该对象表示范围中的所有可用更新。
    'Range.WordOpenXML': None,
    # 返回一个 String 类型的值，该值代表 Microsoft Office Word Open XML 格式中包含的 XML。 只读。
    'Range.Words': None,
    # 返回一个 Words 集合，该集合代表区域中的所有单词。
    'Range.XML[Boolean]': None,
    # 返回一个 String 类型的值，该值代表指定对象中的 XML 文本。
    'Range.XMLNodes': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    'Range.XMLParentNode': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。

    # ParagraphFormat >>>
    'ParagraphFormat.AddSpaceBetweenFarEastAndAlpha': None,
    # 如此 如果 Microsoft Word 设置为自动添加指定段落的日语和西文文字之间的空格。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.AddSpaceBetweenFarEastAndDigit': None,
    # 如此 如果 Microsoft Word 设置为自动添加日语文本和数字指定段落之间的间距。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Alignment': None,
    # 返回或设置一个 WdParagraphAlignment 常量，该常量表示指定段落的对齐方式。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Application': None,
    # 返回一个Application对象，该对象表示 Microsoft Word 应用程序。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.AutoAdjustRightIndent': None,
    # 如此 如果 Microsoft Word 设置为自动调整指定段落的右缩进，如果已指定每行指定一组字符数。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.BaseLineAlignment': None,
    # 返回或设置一个 WdBaselineAlignment 常量，该常量表示字体在行上的垂直位置。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.CharacterUnitFirstLineIndent': None,
    # 返回或设置首行或悬挂缩进的值 (以字符为单位)。 用正值设置首行缩进，并使用一个负值设置悬挂缩进。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.CharacterUnitLeftIndent': None,
    # 返回或设置指定段落的左缩进值 (以字符为单位)。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.CharacterUnitRightIndent': None,
    # 该属性返回或设置指定段落的右缩进量（以字符为单位）。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.CollapsedByDefault': None,
    # 返回或设置指定的段落格式是否默认折叠。 在 C#) 中读写 整数 (int 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.DisableLineHeightGrid': None,
    # 指定 true 如果 Microsoft Word 将与行网格时一组每页的行数指定段落中的字符。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Duplicate': None,
    # 返回一个只读 _ParagraphFormat 对象，该对象代表指定段落的段落格式。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.FarEastLineBreakControl': None,
    # 如此 如果 Microsoft Word 将东亚语言文字的换行规则应用于指定的段落。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.FirstLineIndent': None,
    # 返回或设置首行的行或悬挂缩进的值 (以磅为单位)。 用正数设置首行缩进的尺寸，用负数设置悬挂缩进的尺寸。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.HalfWidthPunctuationOnTopOfLine': None,
    # 如此 如果 Microsoft Word 更改为半角字符指定段落的标点符号在一行的开头。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.HangingPunctuation': None,
    # 为指定段落启用 true 如果标点。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Hyphenation': None,
    # 如此 如果指定段落的段包括在自动断字功能。 假 如果指定的段落不进行自动断字。 可以为 真 ，或者 wdUndefined 则 为 False 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.KeepTogether': None,
    # 如此 如果指定段落中的所有行都保持在同一页上时，Microsoft Word 对文档重新分页。 可 真 、 假 或 wdUndefined 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.KeepWithNext': None,
    # 如此 如果指定的段落保留在其后当 Microsoft Word 对文档重新分页的段落位于同一页上。 可 真 、 假 或 wdUndefined 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.LeftIndent': None,
    # 返回或设置 一个 值，表示指定段落的左缩进值 (以磅为单位)。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.LineSpacing': None,
    # 返回或设置指定段落的行距 (以磅为单位)。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.LineSpacingRule': None,
    # 返回或设置指定段落的行距。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.LineUnitAfter': None,
    # 返回或设置指定段落的段后间距 (以网格线)。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.LineUnitBefore': None,
    # 返回或设置指定段落的段前间距 (以网格线) 的数量。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.MirrorIndents': None,
    # 返回或设置一个 Integer 类型的值，该值代表左缩进和右缩进的宽度是否相同。 可以为 True、 False 或 wdUndefined。 读/写。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.NoLineNumber': None,
    # 如此 如果取消指定段落的行号。 可 真 、 假 或 wdUndefined 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.OutlineLevel': None,
    # 返回或设置指定段落的大纲级别。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.PageBreakBefore': None,
    # 如此 如果分页符强制在指定段落前。 可 真 、 假 或 wdUndefined 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.ReadingOrder': None,
    # 返回或设置指定段落的读取次序而不改变其对齐方式。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.RightIndent': None,
    # 返回或设置指定段落的右缩进量（以磅为单位）。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.SpaceAfter': None,
    # 返回或设置指定段落后) 以磅为单位的间距 (量。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.SpaceAfterAuto': None,
    # 如此 如果 Microsoft Word 自动设置指定段落的段后间距量。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.SpaceBefore': None,
    # 返回或设置指定段落的段前间距 (以磅为单位)。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.SpaceBeforeAuto': None,
    # 如此 如果 Microsoft Word 自动设置指定段落的段前间距。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.Style': None,
    # 返回或设置指定对象的样式。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.TabStops': None,
    # 返回或设置一个 TabStops 集合，该集合代表指定段落的所有自定义制表位。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.TextboxTightWrap': None,
    # 返回或设置一个 WdTextboxTightWrap 枚举值，该值表示文本环绕形状或文本框的紧密程度。 读/写。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.WidowControl': None,
    # 如此 如果指定段落中的第一个和最后一个行 Word 对文档重新分页时保留在其余的段落位于同一页上。 可以为 真 ，或者 wdUndefined 则 为 False 。
    # (继承自 _ParagraphFormat)
    'ParagraphFormat.WordWrap': None,
    # 如果 Microsoft Word 在指定段落或文本框架的西文单词中间断字换行，则该属性值为 True。
    # (继承自 _ParagraphFormat)

    # Font >>>
    'Font.AllCaps': None,
    # 如果字体格式为全部字母大写，则该属性值为 True。 返回 True、False 或 wdUndefined（当返回值既可为 True，也可为 False 时取该值）。 该属性可设置为 True、False 或 wdToggle（与当前设置相反）。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Animation': None,
    # 此对象、成员或枚举已被弃用并且不适合在您的代码中使用。
    # (继承自 _Font)
    'Font.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    # (继承自 _Font)
    'Font.Bold': None,
    # 如此 如果将字体或区域的格式设置为加粗格式。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.BoldBi': None,
    # 如此 如果将字体或区域的格式设置为加粗格式。 对于粗体和非粗体文本的混合) ，返回 True、 False 或 wdUndefined (。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    # (继承自 _Font)
    'Font.Color': None,
    # 返回或设置指定 Font 对象的 24 位颜色。
    # (继承自 _Font)
    'Font.ColorIndex': None,
    # 返回或设置指定 Border 或 Font 对象的颜色。
    # (继承自 _Font)
    'Font.ColorIndexBi': None,
    # 返回或设置从右到左语言文档中指定 Font 对象的颜色。
    # (继承自 _Font)
    'Font.ContextualAlternates': None,
    # 指定对指定的字体启用上下文替代字。
    # (继承自 _Font)
    'Font.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。 只读 Integer。
    # (继承自 _Font)
    'Font.DiacriticColor': None,
    # 返回或设置要用于指定 Font 对象的音调符号的 24 位颜色。 可以是 Visual Basic 的 RGB 函数返回的任何有效WdColor常量或值。 读/写。
    # (继承自 _Font)
    'Font.DisableCharacterSpaceGrid': None,
    # 如果 Microsoft Word忽略相应Font对象的每行字符数，则其值为 True。 读/写 Boolean。
    # (继承自 _Font)
    'Font.DoubleStrikeThrough': None,
    # 如果指定字体的格式设置为双删除线文本，则该属性值为 True。 返回 True、False 或 wdUndefined（当返回值既可为 True，也可为 False 时取该值）。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Duplicate': None,
    # 返回一个只读 Font 对象，该对象代表指定字体的字符格式。
    # (继承自 _Font)
    'Font.Emboss': None,
    # 如此 如果将指定的字体的格式设置为阳文。 返回 真 、 假 或 wdUndefined 。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.EmphasisMark': None,
    # 返回或设置字符或指定的字符字符串的着重号。
    # (继承自 _Font)
    'Font.Engrave': None,
    # 如果该字体的格式设置为阴文， 则返回 true 。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Fill': None,
    # 获取一个 FillFormat 对象，该对象包含指定文本范围使用的字体的填充格式属性。
    # (继承自 _Font)
    'Font.Glow': None,
    # 获取一个 GlowFormat 对象，该对象代表指定文本范围使用的字体的发光格式。
    # (继承自 _Font)
    'Font.Hidden': None,
    # 如此 如果字体的格式设置为隐藏文字。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Italic': None,
    # 如此 如果字体或区域的格式设置为倾斜格式。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.ItalicBi': None,
    # 如此 如果字体或区域的格式设置为倾斜格式。 为斜体和非斜体文本的混合) 返回 True、 False 或 wdUndefined (。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Kerning': None,
    # 返回或设置 Microsoft Word 将调整自动字距调整的最小字号。 读/写 单个 。
    # (继承自 _Font)
    'Font.Ligatures': None,
    # 获取或设置指定 Font 对象的连字设置。
    # (继承自 _Font)
    'Font.Line': None,
    # 获取一个 LineFormat 对象，该对象指定行的格式。
    # (继承自 _Font)
    'Font.Name': None,
    # 返回或设置指定对象的名称。 读/写 String。
    # (继承自 _Font)
    'Font.NameAscii': None,
    # 返回或设置用于西文文本 (字符代码为 0 (零) 到 127 个字符) 的字体。 读/写 String。
    # (继承自 _Font)
    'Font.NameBi': None,
    # 返回或设置从右到左语言的文档中的字体的名称。 读/写 String。
    # (继承自 _Font)
    'Font.NameFarEast': None,
    # 返回或设置一种东亚字体名称。 读/写 String。
    # (继承自 _Font)
    'Font.NameOther': None,
    # 返回或设置用于从 128 到 255 的字符代码的字符的字体。 读/写 String。
    # (继承自 _Font)
    'Font.NumberForm': None,
    # 返回或设置 OpenType 字体的数字形式设置。
    # (继承自 _Font)
    'Font.NumberSpacing': None,
    # 获取或设置字体的数字间距设置。
    # (继承自 _Font)
    'Font.Outline': None,
    # 如果字体格式为镂空，则该属性值为 True。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    # (继承自 _Font)
    'Font.Position': None,
    # 返回或设置 (以磅为单位) 的文本相对于基准线的位置。 正值将文本提升，负值将文本降低。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Reflection': None,
    # 获取一个 ReflectionFormat 对象，该对象表示形状的反射格式。
    # (继承自 _Font)
    'Font.Scaling': None,
    # 返回或设置应用于字体的缩放百分比。 本属性以当前字体大小的百分比水平拉长或压缩文字（缩放范围从 1 到 600）。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    # (继承自 _Font)
    'Font.Shadow': None,
    # 如果将指定字体设置为阴影格式，则该属性值为 True。 可 真 、 假 或 wdUndefined 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Size': None,
    # 返回或设置字体大小，以磅为单位。 读/写 单个 。
    # (继承自 _Font)
    'Font.SizeBi': None,
    # 返回或设置字体大小，以磅为单位。 读取/写入单。
    # (继承自 _Font)
    'Font.SmallCaps': None,
    # 如果字体格式为小型大写字母，则该属性值为 True。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Spacing': None,
    # 返回或设置字符之间的间距 (以磅为单位)。 读/写 单个 。
    # (继承自 _Font)
    'Font.StrikeThrough': None,
    # 如此 如果字体的格式设置为带删除线的文本。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.StylisticSet': None,
    # 获取或设置指定字体的风格集。
    # (继承自 _Font)
    'Font.Subscript': None,
    # 如此 如果字体的格式设置为下标。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 Integer 型，可读/写。
    # (继承自 _Font)
    'Font.Superscript': None,
    # 如此 如果字体格式为上标。 返回 True、 False 或 wdUndefined (true 和 False) 的混合。 可以设置为 真 、 假 或 wdToggle 。 读/写 Long。
    # (继承自 _Font)
    'Font.TextColor': None,
    # 获取一个 ColorFormat 对象，该对象表示指定字体的颜色。
    # (继承自 _Font)
    'Font.TextShadow': None,
    # 获取一个 ShadowFormat 对象，该对象指定指定字体的阴影格式。
    # (继承自 _Font)
    'Font.ThreeD': None,
    # 获取一个 ThreeDFormat 对象，该对象包含指定字体的三维效果格式属性。
    # (继承自 _Font)
    'Font.Underline': None,
    # 返回或设置应用于字体的下划线类型。 读/写 WdUnderline。
    # (继承自 _Font)
    'Font.UnderlineColor': None,
    # 返回或设置指定 Font 对象的下划线的 24 位颜色。
    # (继承自 _Font)

    # Tables >>>
    'Tables.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Tables.Count': None,
    # 返回指定集合中的项数。
    'Tables.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Tables.Item[Int32]': None,
    # 返回集合中的单个对象。
    'Tables.NestingLevel': None,
    # 返回指定表格的嵌套层。
    'Tables.Parent': None,
    # 返回一个对象，代表指定对象的父对象。    

    # Table >>>
    'Table.AllowAutoFit': None,
    # 使 Microsoft Word 可以自动调整表格中的单元格的大小以适应内容。
    'Table.AllowPageBreaks': None,
    # 允许跨页断行的指定的表格中的 Microsoft Word。
    'Table.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Table.ApplyStyleColumnBands': None,
    # 返回或设置 Boolean 值，该值表示如果所应用的预置表样式为列提供了样式带，则是否对表中的列应用样式带。 读/写。
    'Table.ApplyStyleFirstColumn': None,
    # 真正 的 Microsoft Word，以便第一列将格式应用于指定表格的第一列。
    'Table.ApplyStyleHeadingRows': None,
    # 真正 的 Microsoft Word 标题行格式应用于所选表的第一行。
    'Table.ApplyStyleLastColumn': None,
    # 如果为 True，则 Microsoft Word 对指定表格的最后一列应用最后一列的格式。
    'Table.ApplyStyleLastRow': None,
    # 真正 的 Microsoft Word，以便将应用最后一行的最后一个格式指定表中的行。
    'Table.ApplyStyleRowBands': None,
    # 返回或设置一个 boolean 类型的值 ，该值代表是否如果应用的预置的表样式为提供了样式带行到表中的行应用样式带。 读/写。
    'Table.AutoFormatType': None,
    # 返回已应用于指定表格的自动套用格式类型。
    'Table.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Table.BottomPadding': None,
    # 返回或设置的单个单元格或表格中的所有单元格内容的下方添加的间距 (以磅为单位)。
    'Table.Columns': None,
    # 返回一个 Columns 集合，该集合代表表中的所有表列。
    'Table.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Table.Descr': None,
    # 获取或设置包含指定表的说明的字符串。
    'Table.ID': None,
    # 返回或设置指定对象的标识标签当前文档保存为 Web 页时。
    'Table.LeftPadding': None,
    # 返回或设置要添加的单个单元格或表格中的所有单元格的内容左侧的间距 (以磅为单位)。
    'Table.NestingLevel': None,
    # 返回指定表格的嵌套层。
    'Table.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Table.PreferredWidth': None,
    # 返回或设置首选宽度 (（以磅为单位），或者作为指定单元格、单元格、列或表格) 窗口宽度的百分比。
    'Table.PreferredWidthType': None,
    # 返回或设置用于指定表格宽度的指定度量单位。
    'Table.Range': None,
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'Table.RightPadding': None,
    # 返回或设置要添加的单个单元格或表格中的所有单元格的内容右侧的间距 (以磅为单位)。
    'Table.Rows': None,
    # 返回一个 Rows 集合，该集合代表表中的所有表行。
    'Table.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Table.Spacing': None,
    # 返回或设置表格中单元格之间的间距 (以磅为单位)。
    'Table.Style': None,
    # 返回或设置指定对象的样式。
    'Table.TableDirection': None,
    # 返回或设置 Microsoft Word 对指定表格中的单元格进行排序的方向。
    'Table.Tables': None,
    # 返回一个 Tables 集合，该集合表示指定表中的所有表。
    'Table.Title': None,
    # 获取或设置一个字符串，其中包含指定表的标题。
    'Table.TopPadding': None,
    # 返回或设置单个单元格或表格中的所有单元格的内容上方要增加的间距 (以磅为单位)。
    'Table.Uniform': None,
    # 如此 如果表中的所有行都具有相同的列数。

    # Rows >>>
    'Rows.Alignment': None,
    # 返回或设置一个 WdRowAlignment 常量，该常量表示指定行的对齐方式。
    'Rows.AllowBreakAcrossPages': None,
    # 确定是否允许表格行或行中的文本拆分为分页符。
    'Rows.AllowOverlap': None,
    # 返回或设置一个值，该值指定是否允许指定行与其他行重叠。
    'Rows.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Rows.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Rows.Count': None,
    # 返回指定集合中的项数。
    'Rows.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Rows.DistanceBottom': None,
    # 返回或设置文档文字与指定表格的下边缘之间的距离 (以磅为单位)。
    'Rows.DistanceLeft': None,
    # 返回或设置文档文字与指定表格的左边的缘之间的距离 (以磅为单位)。
    'Rows.DistanceRight': None,
    # 返回或设置文档文字与指定表格的右边缘之间的距离 (以磅为单位)。
    'Rows.DistanceTop': None,
    # 返回或设置文档文字与指定表格的上边缘之间的距离 (以磅为单位)。
    'Rows.First': None,
    # 返回一个 Row 对象，该对象代表集合中的 Rows 第一项。
    'Rows.HeadingFormat': None,
    # 确定指定的行的格式是否为表标题。
    'Rows.Height': None,
    # 返回或设置表中指定行的高度。
    'Rows.HeightRule': None,
    # 返回或设置用于确定指定行高度的规则。
    # wdRowHeightAuto	    0	调整行高以适应该行中的最大高度值。
    # wdRowHeightAtLeast	1	行高至少是最小的指定值。
    # wdRowHeightExactly	2	行高是固定值。
    'Rows.HorizontalPosition': None,
    # 返回或设置行边缘与 属性指定的 RelativeHorizontalPosition 项之间的水平距离。
    'Rows.Item[Int32]': None,
    # 返回集合中的单个对象。
    'Rows.Last': None,
    # 以 对象的形式返回集合中的 Rows 最后一项 Row 。
    'Rows.LeftIndent': None,
    # 返回或设置用于指定的表格行的 单个 值，该值代表左缩进值 (以磅为单位)。
    'Rows.NestingLevel': None,
    # 返回指定行的嵌套级别。
    'Rows.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Rows.RelativeHorizontalPosition': None,
    # 指定一组行的水平位置的相对位置。
    'Rows.RelativeVerticalPosition': None,
    # 指定一组行的垂直位置的相对位置。
    'Rows.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Rows.SpaceBetweenColumns': None,
    # 返回或设置指定的一行或多行的相邻列中的文本之间的距离 (以磅为单位)。
    'Rows.TableDirection': None,
    # 返回或设置 Microsoft Word 对指定的表格或行中的单元格进行排序的方向。
    'Rows.VerticalPosition': None,
    # 返回或设置行边缘与 属性指定的 RelativeVerticalPosition 项之间的垂直距离。
    'Rows.WrapAroundText': None,
    # 确定文本是否应环绕指定的行。

    # Row >>>
    'Row.Alignment': None,
    # 返回或设置一个 WdRowAlignment 常量，该常量表示指定行的对齐方式。
    'Row.AllowBreakAcrossPages': None,
    # 确定是否允许表格行或行中的文本拆分为分页符。
    'Row.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Row.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Row.Cells': None,
    # 返回一个 Cells 集合，该集合代表列、行、选定内容或区域中的表格单元格。
    'Row.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Row.HeadingFormat': None,
    # 确定指定的行的格式是否为表标题。
    'Row.Height': None,
    # 返回或设置表格中指定行的高度。
    'Row.HeightRule': None,
    # 返回或设置用于确定指定行高度的规则。
    'Row.ID': None,
    # 返回或设置指定对象的标识标签当前文档保存为 Web 页时。
    'Row.Index': None,
    # 返回一个 Integer 类型的值，该值代表项在集合中的位置。
    'Row.IsFirst': None,
    # 确定指定的列或行是表中的第一个列或行。
    'Row.IsLast': None,
    # 确定指定的列或行是表中的最后一个列或行。
    'Row.LeftIndent': None,
    # 返回或设置用于指定的表格行的 单个 值，该值代表左缩进值 (以磅为单位)。
    'Row.NestingLevel': None,
    # 返回指定行的嵌套级别。
    'Row.Next': None,
    # 返回集合中的下一个对象。
    'Row.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Row.Previous': None,
    # 返回集合中的上一个对象。
    'Row.Range': None,
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'Row.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Row.SpaceBetweenColumns': None,
    # 返回或设置指定行相邻列中文本之间的距离 () 磅。

    # Columns >>>
    'Columns.Application': None,
    # 返回一个 Application 对象，该对象代表指定对象的创建者。
    'Columns.Borders': None,
    # 返回或设置一个 Borders 集合，该集合代表指定对象的所有边框。
    'Columns.Count': None,
    # 返回指定集合中的项数。
    'Columns.Creator': None,
    # 返回一个值，该值指示在其中创建了指定对象的应用程序。
    'Columns.First': None,
    # 返回一个 Column 对象，该对象代表集合中的 Columns 第一项。
    'Columns.Item[Int32]': None,
    # 返回集合中的单个对象。
    'Columns.Last': None,
    # 返回一个 Column 对象，该对象代表集合中的最后一 Columns 项。
    'Columns.NestingLevel': None,
    # 返回指定列的嵌套层。
    'Columns.Parent': None,
    # 返回指定对象的父对象。
    'Columns.PreferredWidth': None,
    # 返回或设置指定列的首选的宽度 (以磅为单位或表示为窗口宽度的百分比)。
    'Columns.PreferredWidthType': None,
    # 返回或设置用于指定列宽度的首选度量单位。
    'Columns.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Columns.Width': None,
    # 返回或设置指定对象的宽度（以磅为单位）。

    # Column >>>
    'Column.Application': None,
    # 返回一个 Application 对象，该对象代表指定对象的创建者。
    'Column.Borders': None,
    # 返回或设置一个 Borders 集合，该集合代表指定对象的所有边框。
    'Column.Cells': None,
    # 返回一个 Cells 集合，该集合代表列中的表格单元格。
    'Column.Creator': None,
    # 返回一个值，该值指示在其中创建了指定对象的应用程序。
    'Column.Index': None,
    # 返回一个值，该值表示项在集合中的位置。
    'Column.IsFirst': None,
    # 返回一个值，该值指示指定的列是否为表中的第一列。
    'Column.IsLast': None,
    # 返回一个值，该值指示指定的列是否是表中的最后一列。
    'Column.NestingLevel': None,
    # 返回指定列的嵌套层。
    'Column.Next': None,
    # 返回集合中的下一个对象。
    'Column.Parent': None,
    # 返回指定对象的父对象。
    'Column.PreferredWidth': None,
    # 返回或设置指定列的首选的宽度 (以磅为单位或表示为窗口宽度的百分比)。
    'Column.PreferredWidthType': None,
    # 返回或设置用于指定列宽度的首选度量单位。
    'Column.Previous': None,
    # 返回集合中的上一个对象。
    'Column.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Column.Width': None,
    # 返回或设置指定对象的宽度（以磅为单位）。

    # Cells >>>
    'Cells.Application': None,
    # 返回一个 Application 对象，该对象代表指定对象的创建者。
    'Cells.Borders': None,
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'Cells.Count': None,
    # 返回指定集合中的项数。
    'Cells.Creator': None,
    # 返回一个值，该值指示在其中创建了指定对象的应用程序。
    'Cells.Height': None,
    # 返回或设置表中指定单元格的高度。
    'Cells.HeightRule': None,
    # 返回或设置用于确定指定单元格高度的规则。
    'Cells.Item[Int32]': None,
    # 返回集合中的单个对象。
    'Cells.NestingLevel': None,
    # 返回指定单元格的嵌套层。
    'Cells.Parent': None,
    # 返回指定对象的父对象。
    'Cells.PreferredWidth': None,
    # 返回或设置指定单元格的首选的宽度 (以磅为单位或表示为窗口宽度的百分比)。
    'Cells.PreferredWidthType': None,
    # 返回或设置用于指定单元格的宽度度量单位。
    'Cells.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Cells.VerticalAlignment': None,
    # 返回或设置表格单元格中文本的垂直对齐方式。
    'Cells.Width': None,
    # 返回或设置指定对象的宽度（以磅为单位）。    
    
    # Cell >>>
    'Cell.Application': None,
    # 返回一个 Application 对象，该对象代表指定对象的创建者。
    'Cell.Borders': None,
    # 返回或设置一个 Borders 集合，该集合代表指定对象的所有边框。
    'Cell.BottomPadding': None,
    # 返回或设置 (以磅为单位的空间量，) 要添加到单元格内容下方。
    'Cell.Column': None,
    # 返回一个 Column 对象，该对象表示包含指定单元格的表列。
    'Cell.ColumnIndex': None,
    # 返回一个值，该值指示包含指定单元格的表列的编号。
    'Cell.Creator': None,
    # 返回一个值，该值指示在其中创建了指定对象的应用程序。
    'Cell.FitText': None,
    # 返回或设置一个值，该值指示 Microsoft Word是否在视觉上减小了在单元格中键入的文本的大小，使其适合列宽。
    'Cell.Height': None,
    # 除非) 另有说明，否则返回或设置指定对象的高度 (以磅为单位。
    'Cell.HeightRule': None,
    # 返回或设置用于确定指定单元格高度的规则。
    'Cell.ID': None,
    # 返回或设置用于确定指定单元格高度的规则。
    'Cell.LeftPadding': None,
    # 返回或设置要添加到单个单元格内容左侧) 以磅为单位的空间 (量。
    'Cell.NestingLevel': None,
    # 返回指定单元格的嵌套层。
    'Cell.Next': None,
    # 返回集合中的下一个对象。
    'Cell.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Cell.PreferredWidth': None,
    # 返回或设置指定单元格的首选的宽度 (以磅为单位或表示为窗口宽度的百分比)。
    'Cell.PreferredWidthType': None,
    # 返回或设置用于指定单元格的宽度度量单位。
    'Cell.Previous': None,
    # 返回集合中的上一个对象。
    'Cell.Range': None,
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'Cell.RightPadding': None,
    # 返回或设置要添加到单元格内容右侧) (磅的空间量。
    'Cell.Row': None,
    # 返回一个 Row 对象，该对象表示包含指定单元格的行。
    'Cell.RowIndex': None,
    # 返回包含指定单元格的行的数目。
    'Cell.Shading': None,
    # 返回一个 Shading 对象，该对象引用指定对象的底纹格式。
    'Cell.Tables': None,
    # 返回一个 Tables 集合，该集合代表指定单元格中的所有表。
    'Cell.TopPadding': None,
    # 返回或设置单个单元格或表格中的所有单元格的内容上方要增加的间距 (以磅为单位)。
    'Cell.VerticalAlignment': None,
    # 返回或设置表格单元格中文本的垂直对齐方式。
    'Cell.Width': None,
    # 返回或设置指定对象的宽度（以磅为单位）。
    'Cell.WordWrap': None,
    # 返回或设置一个值，该值指示 Microsoft Word是否将文本换行到多行并加长单元格，以便单元格宽度保持不变。

    # Borders >>>
    'Borders.AlwaysInFront': None,
    # 返回或设置一个值，该值指示页面边框是否显示在文档文本的前面。
    'Borders.Application': None,
    # 返回一个Application对象，该对象表示 Microsoft Word 应用程序。
    'Borders.Count': None,
    # 返回指定集合中的项数。
    'Borders.Creator': None,
    # 返回一个值，该值指示在其中创建了指定对象的应用程序。
    'Borders.DistanceFrom': None,
    # 返回或设置一个值，该值指示指定的页面边框从页面边缘测量还是从环绕的文本。
    'Borders.DistanceFromBottom': None,
    # 返回或设置文本与下边框之间的间距 (以磅为单位)。
    'Borders.DistanceFromLeft': None,
    # 返回或设置文本与左边的框之间的距离 (以磅为单位)。
    'Borders.DistanceFromRight': None,
    # 返回或设置文本的右边缘与右边框之间的间距 (以磅为单位)。
    'Borders.DistanceFromTop': None,
    # 返回或设置文本与上边框之间的间距 (以磅为单位)。
    'Borders.Enable': None,
    # 返回或设置指定对象的边框格式。
    'Borders.EnableFirstPageInSection': None,
    # 返回或设置一个值，该值指示是否为节中的第一页启用了页面边框。
    'Borders.EnableOtherPagesInSection': None,
    # 返回或设置一个值，该值指示是否为节中的所有页面启用页面边框，但第一页除外。
    'Borders.HasHorizontal': None,
    # 返回一个值，该值指示是否可以将水平边框应用于 对象。
    'Borders.HasVertical': None,
    # 返回一个值，该值指示垂直边框是否可以应用于指定的对象。
    'Borders.InsideColor': None,
    # 返回或设置一个值，该值指示内部边框的 24 位颜色。
    'Borders.InsideColorIndex': None,
    # 返回或设置内部边框。
    'Borders.InsideLineStyle': None,
    # 返回或设置指定对象的内边框。
    'Borders.InsideLineWidth': None,
    # 返回或设置对象的内边框的线条宽度。
    'Borders.Item[WdBorderType]': None,
    # 返回一个值，该值指示是否删除段落和表格边缘的垂直边框，以便水平边框可以连接到页面边框。
    'Borders.JoinBorders': None,
    # 返回或设置一个值，该值指示是否删除段落和表格边缘的垂直边框，以便水平边框可以连接到页面边框。
    'Borders.OutsideColor': None,
    # 返回或设置外部边框的 24 位颜色。
    'Borders.OutsideColorIndex': None,
    # 返回或设置外部边框的颜色。
    'Borders.OutsideLineStyle': None,
    # 返回或设置指定对象的外边框。
    'Borders.OutsideLineWidth': None,
    # 返回或设置对象外部边框的线条宽度。
    'Borders.Parent': None,
    # 返回指定对象的父对象。
    'Borders.Shadow': None,
    # 返回或设置一个值，该值指示指定的边框的格式是否为阴影。
    'Borders.SurroundFooter': None,
    # 返回或设置一个值，该值指示页面边框是否包含文档页脚。
    'Borders.SurroundHeader': None,
    # 返回或设置一个值，该值指示页面边框是否包含文档页眉。

    # Border >>>
    'Border.Application': None,
    # 返回一个Application对象，该对象表示 Microsoft Word 应用程序。
    'Border.ArtStyle': None,
    # 返回或设置文档的图形页面边框设计。
    'Border.ArtWidth': None,
    # 返回或设置指定艺术型边框的宽度 (以磅为单位)。
    'Border.Color': None,
    # 返回或设置指定 Border 对象的 24 位颜色。
    'Border.ColorIndex': None,
    # 返回或设置指定边框对象的颜色。
    'Border.Creator': None,
    # 返回一个值，该值指示在其中创建了指定对象的应用程序。
    'Border.Inside': None,
    # 返回可应用于指定对象的内部边框。
    'Border.LineStyle': None,
    # 返回或设置指定对象的边框线型。
    'Border.LineWidth': None,
    # 返回或设置对象边框的线条宽度。
    'Border.Parent': None,
    # 返回指定对象的父对象。
    'Border.Visible': None,
    # 返回或设置一个值，该值指示指定的 Border 对象是否可见。 

    # InlineShape >>>
    'InlineShape.AlternativeText': None, 
    # 返回或设置与 Web 页中的形状相关联的可选文字。
    'InlineShape.AnchorID': None, 
    # 代表文档的文字层中的对象。
    'InlineShape.Application': None, 
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'InlineShape.Borders': None, 
    # 返回一个 Borders 集合，该集合代表指定对象的所有边框。
    'InlineShape.Chart': None, 
    # 返回文档中内嵌形状集合中的图表。 此为只读属性。
    'InlineShape.Creator': None, 
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'InlineShape.EditID': None, 
    # 代表文档的文字层中的对象。
    'InlineShape.Field': None, 
    # 返回一个 Field 对象，该对象表示与指定形状关联的字段。
    'InlineShape.Fill': None, 
    # 返回一个 FillFormat 对象，该对象包含指定形状的填充格式属性。
    'InlineShape.Glow': None, 
    # 返回发光效果的格式属性。 此为只读属性。
    'InlineShape.GroupItems': None, 
    # 返回内联形状组合在一起的形状。 此为只读属性。
    'InlineShape.HasChart': None, 
    # 如果指定的形状是图表，则为 True。 此为只读属性。
    'InlineShape.HasSmartArt': None, 
    # 如果形状上存在 SmartArt 图表，则获取 True 。
    'InlineShape.Height': None, 
    # 返回或设置指定内联形状的高度。
    'InlineShape.HorizontalLineFormat': None, 
    # 返回一个 HorizontalLineFormat 对象，该对象包含指定 InlineShape 对象的水平线格式。
    'InlineShape.Hyperlink': None, 
    # 返回一个 Hyperlink 对象，该对象表示与指定 InlineShape 对象关联的超链接。
    'InlineShape.IsPictureBullet': None, 
    # 确定对象是否 InlineShape 为图片项目符号。
    'InlineShape.Line': None, 
    # 返回一个 LineFormat 对象，该对象包含指定形状的线条格式属性。
    'InlineShape.LinkFormat': None, 
    # 返回一个 LinkFormat 对象，该对象表示已链接到文件的指定 InlineShape 的链接选项。
    'InlineShape.LockAspectRatio': None, 
    # 确定在调整指定形状的大小时该形状是否保持其原始比例。
    'InlineShape.OLEFormat': None, 
    # 返回一个 OLEFormat 对象，该对象表示 OLE 特征 (，而不是链接指定的 InlineShape) 。
    'InlineShape.OWSAnchor': None, 
    # 仅供内部使用。
    'InlineShape.Parent': None, 
    # 返回一个对象，代表指定对象的父对象。
    'InlineShape.PictureFormat': None, 
    # 返回一个 PictureFormat 对象，该对象包含指定对象的图片格式属性。
    'InlineShape.Range': None, 
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'InlineShape.Reflection': None, 
    # 返回形状的反射格式。 此为只读属性。
    'InlineShape.ScaleHeight': None, 
    # 缩放指定的嵌入式图形相对于原始大小的高度。
    'InlineShape.ScaleWidth': None, 
    # 缩放指定的嵌入式图形相对于原始大小的宽度。
    'InlineShape.Script': None, 
    # 返回一个 Script 对象，该对象表示指定网页上的脚本或代码块。
    'InlineShape.Shadow': None, 
    # 返回指定形状的阴影格式。 此为只读属性。
    'InlineShape.SmartArt': None, 
    # 获取一个 SmartArt 对象，该对象提供一种处理与指定内联形状关联的 SmartArt 的方法。
    'InlineShape.SoftEdge': None, 
    # 返回形状的柔化边缘格式。 此为只读属性。
    'InlineShape.TextEffect': None, 
    # 返回一个 TextEffectFormat 对象，该对象包含指定形状的文本效果格式属性。
    'InlineShape.Title': None, 
    # 获取或设置一个值，该值包含指定内联形状的标题。
    'InlineShape.Type': None, 
    # 属性返回嵌入式图形的类型。
    'InlineShape.Width': None, 
    # 返回或设置指定对象的宽度（以磅为单位）。

    # Headers >>>
    'Headers.SelfType': None,
    # 自定义，返回文档或节中的指定页眉或页脚类型。
    'Headers.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Headers.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Headers.Exists': None,
    # 如此 如果 指定的 HeaderFooter 对象存在。
    'Headers.Index': None,
    # 返回一个 WdHeaderFooterIndex 常量，该常量代表文档或节中的指定页眉或页脚。
    'Headers.IsHeader': None,
    # 如此 如果 指定的 HeaderFooter 对象是标头。
    'Headers.LinkToPrevious': None,
    # 如果指定页眉或页脚链接至前一节中相应的页眉或页脚，则该属性值为 True。 链接页眉或页脚时，其内容与前一个页眉或页脚中的内容相同。
    'Headers.PageNumbers': None,
    # 返回一个 PageNumbers 集合，该集合代表指定页眉或页脚中包含的所有页码字段。
    'Headers.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Headers.Range': None,
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'Headers.Shapes': None,
    # 返回一个 Shapes 集合，该集合代表指定页眉或页脚中的所有 Shape 对象。 该集合可以包含绘图、形状、图片、OLE 对象、ActiveX 控件、文本对象和标注。
    
    # Footers >>>
    'Footers.SelfType': None,
    # 自定义，返回文档或节中的指定页眉或页脚类型。
    'Footers.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Footers.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Footers.Exists': None,
    # 如此 如果 指定的 HeaderFooter 对象存在。
    'Footers.Index': None,
    # 返回一个 WdHeaderFooterIndex 常量，该常量代表文档或节中的指定页眉或页脚。
    'Footers.IsHeader': None,
    # 如此 如果 指定的 HeaderFooter 对象是标头。
    'Footers.LinkToPrevious': None,
    # 如果指定页眉或页脚链接至前一节中相应的页眉或页脚，则该属性值为 True。 链接页眉或页脚时，其内容与前一个页眉或页脚中的内容相同。
    'Footers.PageNumbers': None,
    # 返回一个 PageNumbers 集合，该集合代表指定页眉或页脚中包含的所有页码字段。
    'Footers.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Footers.Range': None,
    # 返回一个 Range 对象，该对象表示包含在指定 对象中的文档部分。
    'Footers.Shapes': None,
    # 返回一个 Shapes 集合，该集合代表指定页眉或页脚中的所有 Shape 对象。 该集合可以包含绘图、形状、图片、OLE 对象、ActiveX 控件、文本对象和标注。

    # Footnotes >>>
    'Footnotes.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'Footnotes.ContinuationNotice': None,
    # 返回一个 Range 对象，该对象代表脚注继续通知。
    'Footnotes.ContinuationSeparator': None,
    # 返回一个 Range 对象，该对象表示脚注延续分隔符。
    'Footnotes.Count': None,
    # 返回指定集合中的项数。
    'Footnotes.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'Footnotes.Item[Int32]': None,
    # 返回集合中的单个对象。
    'Footnotes.Location': None,
    # 返回或设置所有脚注的位置。
    'Footnotes.NumberingRule': None,
    # 返回或设置脚注在分页符或分节符之后的编号方式。
    'Footnotes.NumberStyle': None,
    # 返回或设置所选内容、范围或文档中脚注的数字样式。 读/写。
    'Footnotes.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'Footnotes.Separator': None,
    # 返回一个 Range 对象，该对象表示脚注分隔符。
    'Footnotes.StartingNumber': None,
    # 返回或设置开始注释编号。

    # FootnoteOptions >>>
    'FootnoteOptions.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'FootnoteOptions.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'FootnoteOptions.LayoutColumns': None,
    # 当包含引用标记的节具有多个列时，返回或设置脚注在列中的布局方式。 Read-Write C#) 中的 整数 (int 。
    'FootnoteOptions.Location': None,
    # 返回或设置所有脚注的位置。
    'FootnoteOptions.NumberingRule': None,
    # 返回或设置在分页符或分节符之后的脚注或尾注的编号方式。
    'FootnoteOptions.NumberStyle': None,
    # 返回或设置文档中某个范围或所选脚注的数字样式。
    'FootnoteOptions.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'FootnoteOptions.StartingNumber': None,
    # 返回或设置开始注释编号。
    'FootnoteOptions.SelfSymbol': None,
    # 自定义脚注符号。

    # PageNumbers >>>
    'PageNumbers.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'PageNumbers.ChapterPageSeparator': None,
    # 返回或设置章节号和页码之间的分隔字符。 可以是常量之 WdSeparatorType 一。
    'PageNumbers.Count': None,
    # 返回指定集合中的项数。
    'PageNumbers.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'PageNumbers.DoubleQuote': None,
    # 如此 如果 Microsoft WordPageNumbers用双引号 (“) 指定对象。
    'PageNumbers.HeadingLevelForChapter': None,
    # 返回或设置应用于文档中章节标题的标题级别样式。 可以是介于 0 (零) 到 8 的数字，对应于标题级别 1 到 9。
    'PageNumbers.IncludeChapterNumber': None,
    # 为 页码或题注标签包含章节号时。
    'PageNumbers.Item[Int32]': None,
    # 返回集合中的单个对象。
    'PageNumbers.NumberStyle': None,
    # 返回或设置 对象的数字样式 PageNumbers 。
    'PageNumbers.Parent': None,
    # 返回一个对象，代表指定对象的父对象。
    'PageNumbers.RestartNumberingAtSection': None,
    # 如此 如果页码从 1 重新指定部分的开头开始。
    'PageNumbers.ShowFirstPageNumber': None,
    # 如此 如果在部分中的第一页显示页码。
    'PageNumbers.StartingNumber': None,
    # 返回或设置注释的起始编号，行号或页码。

    # PageNumber >>>
    'PageNumber.Alignment': None,
    # 返回或设置一个 WdPageNumberAlignment 常量，该常量表示页码的对齐方式。
    'PageNumber.Application': None,
    # 返回一个 Application 对象，该对象代表 Microsoft Word 应用程序。
    'PageNumber.Creator': None,
    # 返回一个 32 位整数，它指示在其中创建指定的对象的应用程序。
    'PageNumber.Index': None,
    # 返回一个 Integer 类型的值，该值代表项在集合中的位置。
    'PageNumber.Parent': None,
    # 返回一个对象，代表指定对象的父对象。

    # Reference >>>
    'Reference.SelfSorting': None,
    # 自定义，设置参考文献的排序方式，可选值有：'Author', 'Index'
    'Reference.SelfLabelStyle': None, 
    # 自定义，文中引用的编号样式，例如：'[]', '（）'
    'Reference.SelfNumberStyle': None, 
    # 自定义，引用的编号样式，例如：'.', '【】'
    'Reference.SelfInlineNumberStyle': None,
    # 自定义，调用内置编号类型作为引用的编号样式，例如：'arabic'

    # HeadingStyle >>>
    'Heading1.SelfStyle': None,
    'Heading2.SelfStyle': None,
    'Heading3.SelfStyle': None,
    'Heading4.SelfStyle': None,
    'Heading5.SelfStyle': None,
    'Heading6.SelfStyle': None,
    'Heading7.SelfStyle': None,
    'Heading8.SelfStyle': None,
    'Heading9.SelfStyle': None,
    
    'Heading1.SelfStart': None,
    'Heading2.SelfStart': None,
    'Heading3.SelfStart': None,
    'Heading4.SelfStart': None,
    'Heading5.SelfStart': None,
    'Heading6.SelfStart': None,
    'Heading7.SelfStart': None,
    'Heading8.SelfStart': None,
    'Heading9.SelfStart': None,

    'Borders.NoLeftBorder': False, 
    # 自定义，不显示左边框
    'Borders.NoRightBorder': False, 
    # 自定义，不显示右边框
    'Borders.NoVertical': False,
    # 自定义，不显示纵向框线
    'Borders.NoHorizontal': False, 
    # 自定义，不显示横向框线
}

