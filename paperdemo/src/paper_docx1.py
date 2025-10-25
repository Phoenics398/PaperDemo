# MS Word 自动排版
# 可使用【视图】-【宏】-【录制宏】先执行操作，然后保存宏，再使用【视图】-【宏】-【查看宏】-【编辑】查看代码。

from pdb import set_trace as sc

import os
import shutil
import pandas as pd
import argparse

appPath = os.path.dirname(os.environ.get('APPDATA', ''))
lsVer = ['3.8', '3.9', '3.10', '3.11', '3.12', '3.13']
for ver in lsVer:
    fpath = f"{appPath}\\Local\\Temp\\gen_py\\{ver}"
    if os.path.exists(fpath):
        for fx in os.listdir(fpath):
            if os.path.isdir(os.path.join(fpath, fx)):
                shutil.rmtree(os.path.join(fpath, fx))

import win32com
import win32com.client as client
import traceback as trb

from PIL import Image
from copy import copy, deepcopy
from tqdm import tqdm
from win32com.client import constants as cs

from paperdemo.constants.constant_docx1 import *


__all__ = [
    'STYLE', 'PAGE', 'REFERENCE', 'FORMULA', 'TABLE', 'GRAPH', 'NOTE',
    'AUTHOR', 
    'ADD_FOOTNOTE', 
    'ADD_PARAGRAPHS', 'ADD_PARAGRAPH', 'APPEND_PARAGRAPH', 
    'ADD_FORMULA', 'APPEND_FORMULA', 'ADD_FORMULA_WITH_TABLE',
    'ADD_TABLE', 'ADD_TABLE_WITH_NO_LR_BODERS', 'ADD_TABLE_MANUAL',
    'COPY_TABLE_FROM_DOCX', 'COPY_FROM_DOCX', 
    'ADD_GRAPH', 'APPEND_GRAPH', 'ADD_GRAPH_WITH_TABLE',
    'ADD_NOTE', 'ADD_NOTES', 'APPEND_NOTE',
    'ADD_REFERENCE_LABEL', 'APPEND_REFERENCE_LABEL', 'ADD_REFERENCES',

    'cv', 'addFunctions', 'PaperDocx', 
]

REF, FML, TBL, GRP, NTN = {}, {}, {}, {}, {}

STYLE = {
    '_tag': 'setStyle',
    'attr': None, 
    'content': '',
}
STYLE = argparse.Namespace(**STYLE)

PAGE = {
    '_tag':'setPage',
    'attr': None,
    'content': {},
}
PAGE = argparse.Namespace(**PAGE)

REFERENCE = {
    '_tag': 'setReference',
    'attr': None, 
    'content': {},
}
REFERENCE = argparse.Namespace(**REFERENCE)

FORMULA = {
    '_tag': 'setFormula',
    'attr': None, 
    'content': {},
}
FORMULA = argparse.Namespace(**FORMULA)

TABLE = {
    '_tag': 'setTable',
    'attr': None,
    'content': {},
}
TABLE = argparse.Namespace(**TABLE)

GRAPH = {
    '_tag':'setGraph',
    'attr': None,
    'content': {},
}
GRAPH = argparse.Namespace(**GRAPH)

NOTE = {
    '_tag':'setNote',
    'attr': None,
    'content': {},
}
NOTE = argparse.Namespace(**NOTE)

AUTHOR = {
    '_tag': 'addAuthor',
    'attr': None, 
    'name': [],
    'uid': [],
    'sep': '',
}
AUTHOR = argparse.Namespace(**AUTHOR)

ADD_FOOTNOTE = {
    '_tag': 'addFootnote',
    'attr': None,
    'words': "",
    'parts': [],
}
ADD_FOOTNOTE = argparse.Namespace(**ADD_FOOTNOTE)

ADD_PARAGRAPHS = {
    '_tag': 'addParagraphs',
    'attr': None, 
    'words': [],
}
ADD_PARAGRAPHS = argparse.Namespace(**ADD_PARAGRAPHS)

ADD_PARAGRAPH = {
    '_tag': 'addParagraph',
    'attr': None, 
    'words': '',
}
ADD_PARAGRAPH = argparse.Namespace(**ADD_PARAGRAPH)

APPEND_PARAGRAPH = {
    '_tag': 'appendParagraph',
    'attr': None, 
    'words': '',
}
APPEND_PARAGRAPH = argparse.Namespace(**APPEND_PARAGRAPH)

ADD_FORMULA = {
    '_tag': 'addFormula',
    'attr': None, 
    'title': '',
    'formula': '',
    'fid': '',
}
ADD_FORMULA = argparse.Namespace(**ADD_FORMULA)

APPEND_FORMULA = {
    '_tag': 'appendFormula',
    'attr': None,
    'title': '',
    'formula': '',
    'fid': '',
}
APPEND_FORMULA = argparse.Namespace(**APPEND_FORMULA)

ADD_FORMULA_WITH_TABLE = {
    '_tag': 'addFormulaWithTable',
    'attr': None, 
    'title': '',
    'formula': '',
    'fid': '',
}
ADD_FORMULA_WITH_TABLE = argparse.Namespace(**ADD_FORMULA_WITH_TABLE)

ADD_TABLE = {
    '_tag': 'addTable',
    'attr': None,
    'title': '',
    'table': None,
    'tid': '',
    'merge': [],
    'width': [],
}
ADD_TABLE = argparse.Namespace(**ADD_TABLE)

ADD_TABLE_WITH_NO_LR_BODERS = {
    '_tag': 'addTableWithNoLRBorders',
    'attr': None,
    'title': '',
    'table': None,
    'tid': '',
    'merge': [],
    'width': [],
}
ADD_TABLE_WITH_NO_LR_BODERS = argparse.Namespace(**ADD_TABLE_WITH_NO_LR_BODERS)

ADD_TABLE_MANUAL = {
    '_tag': 'addTableManual',
    'attr': None,
    'title': '',
    'table': None,
    'tid': '',
    'merge': [],
    'width': [],
    'columns': 0,
    'rows': 0,
    'cells': [],
    'lines': [],
}
ADD_TABLE_MANUAL = argparse.Namespace(**ADD_TABLE_MANUAL)

ADD_GRAPH = {
    '_tag': 'addGraph',
    'attr': None,
    'title': '',
    'graph': None,
    'gid': '',
}
ADD_GRAPH = argparse.Namespace(**ADD_GRAPH)

APPEND_GRAPH = {
    '_tag': 'appendGraph',
    'attr': None,
    'title': '',
    'graph': '',
    'gid': '',
}
APPEND_GRAPH = argparse.Namespace(**APPEND_GRAPH)

ADD_GRAPH_WITH_TABLE = {
    '_tag': 'addGraphWithTable',
    'attr': None,
    'title': '',
    'graph': '',
    'gid': '',
}
ADD_GRAPH_WITH_TABLE = argparse.Namespace(**ADD_GRAPH_WITH_TABLE)

ADD_NOTE = {
    '_tag': 'addNote',
    'attr': None,
    'note': '',
    'nid': '',
}
ADD_NOTE = argparse.Namespace(**ADD_NOTE)

ADD_NOTES = {
    '_tag': 'addNotes',
    'attr': None,
    'notes': [],
    'nid': '',
}
ADD_NOTES = argparse.Namespace(**ADD_NOTES)

APPEND_NOTE = {
    '_tag': 'appendNote',
    'attr': None,
    'note': '',
    'nid': '',
}
APPEND_NOTE = argparse.Namespace(**APPEND_NOTE)

ADD_REFERENCE_LABEL = {
    '_tag': 'addReferenceLabel',
    'attr': None,
    'words': '',
}
ADD_REFERENCE_LABEL = argparse.Namespace(**ADD_REFERENCE_LABEL)

APPEND_REFERENCE_LABEL = {
    '_tag': 'appendReferenceLabel',
    'attr': None,
    'words': '',
}
APPEND_REFERENCE_LABEL = argparse.Namespace(**APPEND_REFERENCE_LABEL)

ADD_REFERENCES = {
    '_tag': 'addReferences',
    'attr': None,
    'references': [],
    'rid': '',
    'sort': True,
}
ADD_REFERENCES = argparse.Namespace(**ADD_REFERENCES)

COPY_TABLE_FROM_DOCX = {
    '_tag': 'copyTableFromDocx',
    'attr': None,
    'xpath': '',
    'content': {},
    'split': [],
}
COPY_TABLE_FROM_DOCX = argparse.Namespace(**COPY_TABLE_FROM_DOCX)

COPY_FROM_DOCX = {
    '_tag': 'copyFromDocx',
    'attr': None,
    'xpath': '',
    'content': {},
    'split': [],
    'isSetAttributes': True,
}
COPY_FROM_DOCX = argparse.Namespace(**COPY_FROM_DOCX)

def cv(obj, maps):
    """
    解析字符串
    格式：
    wp|......<!>ap|......
    """
    flow = []
    tmp = obj.split('<!>')
    for ix in tmp:
        for kx, vx in maps.items():
            if ix.startswith(kx):
                bx = vars(copy(vx['dict']))
                bx['attr'] = vx['attr']
                ctx = ix.replace(kx, '')
                if 'content' in bx.keys():
                    bx['content'] = ctx
                elif 'formula' in bx.keys():
                    bx['formula'] = ctx
                else:
                    bx['words'] = ctx
                flow.append(argparse.Namespace(**bx))
    return flow

def addFunctions(paper, maps):
    """
    动态创建函数
    """
    globals().update({'paper': paper})
    res = {}
    for kx, vx in maps.items():
        code = f"""
def {kx[ : -1]}(ctx, **kwargs):
    import argparse
    from copy import copy, deepcopy
    bx = {vars(copy(vx['dict']))}
    bx['attr'] = {vx['attr']}
    bx.update(kwargs)
    if 'content' in bx.keys():
        bx['content'] = ctx
    elif 'formula' in bx.keys():
        bx['formula'] = ctx
    else:
        bx['words'] = ctx
    bx = argparse.Namespace(**bx)
    globals()['paper'].execute(bx)
        """
        # 使用exec执行代码字符串，并定义函数
        exec(code, globals())
        res[kx[ : -1]] = globals()[kx[ : -1]]

    return res

class PaperDocx(object):
    """"""
    def __init__(self, path=''):
        """
        初始化PaperDocx类的实例。

        :param path: 可选参数，指定要打开的Word文档的路径。如果未提供路径，则创建一个新的文档。
        """
        # 创建一个KWPS.Application对象，确保使用缓存的版本
        self.app = client.gencache.EnsureDispatch('Word.Application')
        # self.app = client.gencache.EnsureDispatch('KWPS.Application')
        # self.app = client.Dispatch('Word.Application')
        # 设置应用程序可见性为True，以便用户可以看到操作过程
        self.app.Visible = 1
        # 如果未提供路径，则创建一个新的文档
        if path == '':
            self.doc = self.app.Documents.Add()
        # 如果提供了路径，则打开指定的文档
        else:
            self.doc = self.app.Documents.Open(path)
        self.selection = self.app.Selection
        # 初始化属性字典为空
        self.attr = {}
        # 初始化文档样式为空
        self.docStyle = None

        self.toc = {
            'title': {},
            'graph': {},
            'table': {},
        }

        self.lsReference = []
        self.lsFormula = []
        self.lsTable = []
        self.lsGraph = []
        self.lsNote = []

        self.headings = pd.Series([-1, ] * 10)

        self.exdata = {}

    def _gotoEnd(self):
        """
        将光标移动到文档末尾。
        """
        # self.selection.GoTo(What=cs.wdGoToLine, Count=len(self.doc.Paragraphs))
        self.selection.GoTo(What=cs.wdGoToLine, Which=cs.wdGoToLast)

    def defAttributes(self, obj, attr=None):
        """
        定义格式属性
        """
        for i in obj:
            i.attr = copy(attr)
        return obj

    def setAttributes(self, obj, attr=None):
        """
        页面、段落、文字、页眉、页脚属性设置
        """
        if attr != None:
            for i in range(2):
                # # 如果传入的对象是段落对象，则将其范围添加到属性列表中
                # if obj.__class__ == self.doc.Paragraphs(1).__class__:
                #     lsAttr = [obj.Range, ]
                # # 否则，直接将对象添加到属性列表中
                # else:
                #     lsAttr = [obj, ]
                # # 遍历属性列表中的每个对象
                # for ax in lsAttr:
                # 遍历传入的属性字典中的每个键值对
                # for kx, vx in self.attr.items():
                #     # 将键按点号分割成列表
                #     px = kx.split('.')
                #     # 如果值不为None且键的第一个部分在对象的属性列表中
                #     if vx != None: 
                #         if px[0] in dir(obj):
                #             # 使用exec动态设置对象的属性
                #             try: exec(f'obj.{kx} = vx')
                #             except Exception: pass
                #         elif px[1] in dir(obj):
                #             px[1] = '.'.join(px[1 : ])
                #             try: exec(f'obj.{px[1]} = vx')
                #             except Exception: pass
                # 遍历传入的属性字典中的每个键值对
                for kx, vx in attr.items():
                    # 将键按点号分割成列表
                    px = kx.split('.')
                    # 如果值不为None且键的第一个部分在对象的属性列表中
                    if vx != None: 
                        if px[0] in dir(obj):
                            # 使用exec动态设置对象的属性
                            try: exec(f'obj.{kx} = vx')
                            except Exception: pass
                        elif px[1] in dir(obj):
                            px[1] = '.'.join(px[1 : ])
                            try: exec(f'obj.{px[1]} = vx')
                            except Exception: pass

    def setStyle(self, obj, attr=None):
        """
        添加式样
        """
        # 将传入的属性字典赋值给实例变量self.attr
        self.attr = attr
        # 在文档中添加一个新的样式，名称为'MyStyle'，类型为段落样式
        self.docStyle = self.doc.Styles.Add(Name='MyStyle', Type=CONSTANTS['样式-段落样式'])
        # 使用传入的属性字典设置新样式的属性
        self.setAttributes(self.doc, attr)
        self.setAttributes(self.docStyle, attr)
        # 返回新创建的样式对象
        return self.docStyle
    
    def setPage(self, page, attr=None):
        """
        添加页面
        """
        footer = self.doc.Sections(len(self.doc.Sections)).Footers(attr['Footers.SelfType'])
        footer.PageNumbers.Add(PageNumberAlignment=attr['PageNumber.Alignment'], FirstPage=attr['PageNumbers.ShowFirstPageNumber'])
        self.setAttributes(footer, attr)
        self.setAttributes(footer.Range, attr)

    def setReference(self, reference, attr=None):
        """
        添加引用
        """
        REF.update(reference['content'])

    def setFormula(self, formula, attr=None):
        """
        添加公式
        """
        FML.update(formula['content'])

    def setTable(self, table, attr=None):
        """
        添加表格
        """
        TBL.update(table['content'])

    def setGraph(self, group, attr=None):
        """
        添加分组
        """
        GRP.update(group['content'])

    def setNote(self, note, attr=None):
        """
        添加注释
        """
        NTN.update(note['content'])

    def setTableBorders(self, obj, attr=None):
        """
        表格框线设置
        """
        if not attr['Borders.NoHorizontal']:
            obj.Borders(cs.wdBorderHorizontal).LineStyle = attr['Border.LineStyle']
            obj.Borders(cs.wdBorderHorizontal).LineWidth = attr['Border.LineWidth']
        if not attr['Borders.NoVertical']:
            obj.Borders(cs.wdBorderVertical).LineStyle = attr['Border.LineStyle']
            obj.Borders(cs.wdBorderVertical).LineWidth = attr['Border.LineWidth']
        if not attr['Borders.NoLeftBorder']:
            obj.Borders.Item(cs.wdBorderLeft).LineStyle = attr['Border.LineStyle']
        else:
            obj.Borders.Item(cs.wdBorderLeft).LineStyle = cs.wdLineStyleNone
        if not attr['Borders.NoRightBorder']:
            obj.Borders.Item(cs.wdBorderRight).LineStyle = attr['Border.LineStyle']
        else:
            obj.Borders.Item(cs.wdBorderRight).LineStyle = cs.wdLineStyleNone

    def addAuthor(self, author, attr=None):
        """
        添加作者信息到文档中。

        :param author: 包含作者信息的字典，格式为 {'name': ['作者1', '作者2'], 'uid': ['单位1', '单位2'], 'sep': '分隔符'}
        :param attr: 可选参数，用于设置段落属性的字典
        :return: 返回包含所有作者信息的范围对象
        """
        # 获取作者姓名列表
        nameList = author['name']
        # 获取作者单位列表
        unitList = author['uid']
        # 获取作者之间的分隔符
        sep = author['sep']
        # 获取文档当前范围的结束位置
        locx = self.doc.Range().End
        # 遍历作者姓名列表
        for i, name in enumerate(nameList):
            # 在文档中添加一个新的段落
            tmp = self.doc.Paragraphs.Add()
            # 在段落中插入作者姓名
            tmp.Range.InsertBefore(f"{name}")
            # 如果提供了属性字典，则设置段落的属性
            self.setAttributes(tmp.Range, attr)
            # 在文档中添加一个新的段落
            tmp = self.doc.Paragraphs.Add()
            # 在段落中插入作者单位，并设置为上标
            tmp.Range.InsertBefore(f'{unitList[i]}')
            self.setAttributes(tmp.Range, attr)
            tmp.Range.Font.Superscript = True
            # 如果不是最后一个作者，则在文档中添加一个新的段落，并插入分隔符
            if i != len(nameList) - 1:
                tmp = self.doc.Paragraphs.Add()
                tmp.Range.InsertBefore(f"{sep}")
                self.setAttributes(tmp.Range, attr)
        # 获取文档当前范围的结束位置
        locy = self.doc.Range().End
        # 获取包含所有作者信息的范围对象
        tmp = self.doc.Range(locx - 1, locy - 2)
        # 查找并替换段落标记为空字符串，以去除多余的段落标记
        tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindStop, Replace=cs.wdReplaceAll, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)
        # 返回包含所有作者信息的范围对象
        return tmp
    
    def addFootnote(self, content, attr=None):
        """
        添加标题、页脚
        """
        if attr['FootnoteOptions.SelfSymbol'] != None:
            footnote = self.doc.Footnotes.Add(Range=self.doc.Paragraphs(len(self.doc.Paragraphs) - 1).Range, Reference=f"{attr['FootnoteOptions.SelfSymbol']}")
        else:
            footnote = self.doc.Footnotes.Add(Range=self.doc.Paragraphs(len(self.doc.Paragraphs) - 1).Range)

        if content['words'] != '':
            ctx = content['words']
            footnote.Range.InsertBefore(f"{ctx}")
            self.setAttributes(footnote, attr)
            self.setAttributes(footnote.Range, attr)

        if content['parts'] != []:
            parts = content['parts']
            for idx, part in enumerate(parts):
                part = vars(part)
                ctx = part['words']
                footnote.Range.InsertParagraphAfter()
                rtmp = footnote.Range.Paragraphs(len(footnote.Range.Paragraphs)).Range
                locx = rtmp.End
                if idx == 0:
                    rtmp.Text = "\r"
                rtmp.InsertBefore(f"{ctx}")
                self.setAttributes(rtmp, part['attr'])
                if part['_tag'] == 'appendParagraph':
                    locy = rtmp.End
                    tmp = footnote.Range.Paragraphs(len(footnote.Range.Paragraphs) - 1).Range 
                    tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindContinue, Replace=cs.wdReplaceOne, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)

    def addParagraphs(self, content, attr=None):
        """
        添加段落
        """
        for ctx in content['words']:
            # 在文档中添加一个新的段落
            tmp = self.doc.Paragraphs.Add()
            # 在段落中插入内容文本
            ctx = self.replaceTags(ctx, attr, tmp)
            tmp.Range.InsertBefore(f"{ctx}")
            # 如果提供了属性字典，则设置段落的属性
            self.setAttributes(tmp.Range, attr)
            # 返回添加段落内容后的段落对象
    
    def addParagraph(self, content, attr=None):
        """
        添加段落
        """
        # 在文档中添加一个新的段落
        tmp = self.doc.Paragraphs.Add()
        # 在段落中插入内容文本
        ctx = content['words']
        ctx = self.replaceTags(ctx, attr, tmp)
        tmp.Range.InsertBefore(f"{ctx}")
        # 如果提供了属性字典，则设置段落的属性
        self.setAttributes(tmp.Range, attr)
        # 返回添加段落内容后的段落对象
        return tmp
    
    def appendParagraph(self, content, attr=None):
        """
        添加段落
        """
        # 获取文档当前范围的结束位置
        locx = self.doc.Range().End
        # 在文档中添加一个新的段落
        tmp = self.doc.Paragraphs.Add()
        # 在段落中插入内容文本
        ctx = content['words']
        ctx = self.replaceTags(ctx, attr, tmp)
        tmp.Range.InsertBefore(f"{ctx}")
        # 如果提供了属性字典，则设置段落的属性
        self.setAttributes(tmp.Range, attr)
        # 获取文档当前范围的结束位置
        locy = self.doc.Range().End
        # 获取包含新添加段落的范围对象
        tmp = self.doc.Range(locx - 2, locy)
        # 查找并替换段落标记为空字符串，以去除多余的段落标记
        tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindContinue, Replace=cs.wdReplaceOne, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)
        self.setAttributes(tmp, attr)
        # 返回包含新添加段落的范围对象
        return tmp

    def addFormula(self, formula, attr=None):
        """
        添加公式
        """
        tmp = self.doc.Paragraphs.Add()
        xmp = self.doc.Paragraphs.Add()
        fmx = f"{formula['formula']}"
        # # WPS
        # frm = tmp.Range.OMaths.Add(tmp.Range)
        # fmo = tmp.Range.OMaths(1)
        # fmo.Range.Text = fmx
        # MS Word
        tmp.Range.Text = fmx
        frm = tmp.Range.OMaths.Add(tmp.Range)
        fmo = tmp.Range.OMaths(1)
        fmo.ConvertToMathText()
        fmo.BuildUp()
        self.setAttributes(fmo.Range, attr)
        self.setAttributes(tmp.Range, attr)

    def appendFormula(self, formula, attr=None):
        """
        添加公式
        """
        locx = self.doc.Range().End
        
        tmp = self.doc.Paragraphs.Add()
        xmp = self.doc.Paragraphs.Add()
        fmx = f"{formula['formula']}"
        tmp.Range.Text = fmx
        frm = tmp.Range.OMaths.Add(tmp.Range)
        fmo = tmp.Range.OMaths(1)
        fmo.ConvertToMathText()
        fmo.BuildUp()

        self.setAttributes(tmp.Range, attr)
    
        locy = self.doc.Range().End
        tmp = self.doc.Range(locx - 2, locy)
        tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindContinue, Replace=cs.wdReplaceOne, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)

        self.setAttributes(fmo.Range, attr)
        self.setAttributes(tmp, attr)
    
    def addFormulaWithTable(self, formula, attr=None):
        """
        添加公式
        """
        xmp = self.doc.Paragraphs.Add()        
        zmp = self.doc.Paragraphs.Add()
        ymp = self.doc.Paragraphs.Add()
        table = self.doc.Tables.Add(Range=zmp.Range, NumRows=1, NumColumns=2)

        tmp = table.Cell(1, 1)
        fmx = f"{formula['formula']}"
        # # WPS
        # frm = tmp.Range.OMaths.Add(tmp.Range)
        # fmo = tmp.Range.OMaths(1)
        # fmo.Range.Text = fmx
        # MS Word
        tmp.Range.Text = fmx
        frm = tmp.Range.OMaths.Add(tmp.Range)
        fmo = tmp.Range.OMaths(1)
        fmo.ConvertToMathText()
        fmo.BuildUp()

        self.setAttributes(tmp.Range, attr)

        tmp = table.Cell(1, 2)
        # rfml = dict(zip(FML.values(), FML.keys()))
        # ctx = f"@@FML.{rfml[formula['formula']]}"
        ctx = f"{formula['fid']}"
        ctx = self.replaceTags(ctx, attr, tmp)
        tmp.Range.Text = ctx
        self.setAttributes(tmp.Range, attr)

        xmp.Range.Delete()
        ymp.Range.Delete()

        self.setAttributes(table, attr)
        self.setAttributes(fmo.Range, attr)
        # self.setAttributes(zmp.Range, attr)
        table.Columns(1).PreferredWidth = 90
        table.Columns(2).PreferredWidth = 10

        return table

    def addTable(self, table, attr=None):
        """
        添加表格
        """
        df = pd.DataFrame(table['table'])
        xmp = self.doc.Paragraphs.Add()        
        tmp = self.doc.Paragraphs.Add()
        ymp = self.doc.Paragraphs.Add()
        tblx = self.doc.Tables.Add(Range=tmp.Range, NumRows=len(df.index), NumColumns=len(df.columns))
        dtCellAttr = {}
        for kx, vx in attr.items():
            if 'Cell.' in kx and vx != None:
                dtCellAttr.update({kx: vx})
        for i in range(len(df.index)):
            for j in range(len(df.columns)):
                zmp = tblx.Cell(i + 1, j + 1)
                zmp.Range.Text = f"{df.iloc[i, j]}"

                for kx, vx in dtCellAttr.items():
                    px = kx.split('.')
                    exec(f'zmp.{px[1]} = {vx}')

        if len(table['merge']) != 0:
            for jx in table['merge']:
                tblx.Cell(jx[0][0], jx[0][1]).Merge(tblx.Cell(jx[1][0], jx[1][1]))

        xmp.Range.Delete()
        ymp.Range.Delete()

        # self.setAttributes(tmp.Range, attr)
        self.setAttributes(tblx.Range, attr)
        self.setAttributes(tblx, attr)

        if len(table['width'])!= 0:
            for wx in table['width']:
                # tblx.Columns(wx[0]).PreferredWidth = wx[1]
                tblx.Cell(wx[0][0], wx[0][1]).Select()
                self.selection.SelectColumn()
                self.selection.Columns.PreferredWidthType = attr['Table.PreferredWidthType']
                self.selection.Columns.PreferredWidth = wx[1]

        return tblx
    
    def addTableWithNoLRBorders(self, table, attr=None):
        """
        添加表格
        """
        tblx = self.addTable(table, attr)
        # tblx.Columns(1).Borders.Item(cs.wdBorderLeft).LineStyle = cs.wdLineStyleNone
        # tblx.Columns(len(tblx.Columns)).Borders.Item(cs.wdBorderRight).LineStyle = cs.wdLineStyleNone
        tblx.Select()
        self.app.Selection.Borders.Item(cs.wdBorderLeft).LineStyle = cs.wdLineStyleNone
        self.app.Selection.Borders.Item(cs.wdBorderRight).LineStyle = cs.wdLineStyleNone
        return tblx
    
    def addTableManual(self, table, attr=None):
        """
        添加自定义表格
        """
        tmp = self.doc.Paragraphs.Add()
        tblx = self.doc.Tables.Add(Range=tmp.Range, NumRows=table['rows'], NumColumns=table['columns'])
        dtAttr = {}
        dtCellAttr = {}
        for kx, vx in attr.items():
            if 'Cell.' in kx and vx != None:
                dtCellAttr.update({kx: vx})
            else:
                dtAttr.update({kx: vx})

        for idx, ix in enumerate(table['cells']):
            self._gotoEnd()
            self.selection.GoTo(What=cs.wdGoToLine, Count=ix[1])
            self.selection.MoveEnd(Unit=cs.wdLine, Count=1) 
            self.selection.Cut()
            zmp = tblx.Cell(ix[0][0], ix[0][1])
            zmp.Select()
            self.selection.Paste()
            if idx!= len(table['cells']) - 1:
                self.selection.TypeBackspace()

        self._gotoEnd()
        self.selection.GoTo(What=cs.wdGoToLine, Count=-(table['rows'] + 1))
        self.selection.Delete()

        if len(table['merge']) != 0:
            for jx in table['merge']:
                tblx.Cell(jx[0][0], jx[0][1]).Merge(tblx.Cell(jx[1][0], jx[1][1]))

        for ix in table['cells']:
            zmp = tblx.Cell(ix[0][0], ix[0][1])
            self.setAttributes(zmp.Range, dtAttr)

        for ix in table['lines']:
            zmp = tblx.Cell(ix[0], ix[1])
            self.setAttributes(zmp, dtCellAttr)

        if len(table['width'])!= 0:
            for wx in table['width']:
                tblx.Cell(wx[0][0], wx[0][1]).Select()
                self.selection.SelectColumn()
                self.selection.Columns.PreferredWidthType = dtAttr['Table.PreferredWidthType']
                self.selection.Columns.PreferredWidth = wx[1]

        self._gotoEnd()
        return tblx

    def copyTableFromDocx(self, contents, attr=None):
        """
        复制自其他Word文档的内容到当前文档中。
        """
        # 打开源文档
        try:
            source_doc = self.app.Documents.Open(contents['xpath'])
            # 全选源文档内容
            for tbl in source_doc.Tables:
                # 将光标移动到源文档的开头
                source_doc.Activate()
                # 复制选中内容
                tbl.Select()
                self.app.Selection.Copy()
                # 切换至当前文档
                self.doc.Activate()
                # 将光标移动到当前文档末尾
                self._gotoEnd()
                locx = self.doc.Range().End
                # 粘贴复制的内容
                self.app.Selection.Paste()
                locy = self.doc.Range().End
                tmp = self.doc.Range(locx, locy)
                tmp.Select()
                self.replaceRangeTags(self.app.Selection.Range, attr)
                self.setAttributes(self.app.Selection.Range, attr)
                # if not attr['Borders.NoHorizontal']:
                #     self.app.Selection.Borders(cs.wdBorderHorizontal).LineStyle = attr['Border.LineStyle']
                #     self.app.Selection.Borders(cs.wdBorderHorizontal).LineWidth = attr['Border.LineWidth']
                # if not attr['Borders.NoVertical']:
                #     self.app.Selection.Borders(cs.wdBorderVertical).LineStyle = attr['Border.LineStyle']
                #     self.app.Selection.Borders(cs.wdBorderVertical).LineWidth = attr['Border.LineWidth']
                # if attr['Borders.NoLeftBorder']:
                #     self.app.Selection.Borders.Item(cs.wdBorderLeft).LineStyle = cs.wdLineStyleNone
                # if attr['Borders.NoRightBorder']:
                #     self.app.Selection.Borders.Item(cs.wdBorderRight).LineStyle = cs.wdLineStyleNone
                self.setTableBorders(self.app.Selection, attr)
                # for tblx in self.app.Selection.Tables:
                #     self.setAttributes(tblx.Range, attr)
                #     self.setAttributes(tblx, attr)
                #     self.setTableBorders(tblx, attr)
                self._gotoEnd()
            source_doc.Close(SaveChanges=0)
        
        except Exception:
            tx = copy(ADD_TABLE)
            tx.attr = attr
            lst = []
            with open(contents['xpath']) as f:
                cts = f.readlines()
                for ctx in cts:
                    if not ctx.startswith('----------'):
                        ctx = ctx.replace('\n', '')
                        lsx = []
                        for xx in contents['split']:
                            if xx[1] <= len(ctx):
                                tmp = ctx[xx[0] : xx[1]]
                                tmp = self.replaceTags(tmp, attr)
                                lsx.append(tmp.lstrip().rstrip())
                        lst.append(lsx)
            tx.table = lst
            table = self.addTable(vars(tx), tx.attr)
            self.setTableBorders(table, tx.attr)

    def copyFromDocx(self, contents, attr=None):
        """
        复制自其他Word文档的内容到当前文档中。
        """
        # 打开源文档
        source_doc = self.app.Documents.Open(contents['xpath'])
        # 全选源文档内容
        # 将光标移动到源文档的开头
        source_doc.Activate()
        # 复制选中内容
        source_doc.Select()
        self.app.Selection.Copy()
        # 切换至当前文档
        self.doc.Activate()
        # 将光标移动到当前文档末尾
        self._gotoEnd()
        locx = self.doc.Range().End
        # 粘贴复制的内容
        self.app.Selection.Paste()
        locy = self.doc.Range().End
        tmp = self.doc.Range(locx, locy)
        tmp.Select()
        self.replaceRangeTags(self.app.Selection.Range, attr)
        if contents['isSetAttributes']:
            self.setAttributes(self.app.Selection.Range, attr)
        # if not attr['Borders.NoHorizontal']:
        #     self.app.Selection.Borders(cs.wdBorderHorizontal).LineStyle = attr['Border.LineStyle']
        #     self.app.Selection.Borders(cs.wdBorderHorizontal).LineWidth = attr['Border.LineWidth']
        # if not attr['Borders.NoVertical']:
        #     self.app.Selection.Borders(cs.wdBorderVertical).LineStyle = attr['Border.LineStyle']
        #     self.app.Selection.Borders(cs.wdBorderVertical).LineWidth = attr['Border.LineWidth']
        # if attr['Borders.NoLeftBorder']:
        #     self.app.Selection.Borders.Item(cs.wdBorderLeft).LineStyle = cs.wdLineStyleNone
        # if attr['Borders.NoRightBorder']:
        #     self.app.Selection.Borders.Item(cs.wdBorderRight).LineStyle = cs.wdLineStyleNone
        # self.setTableBorders(self.app.Selection, attr)
        self._gotoEnd()
        source_doc.Close(SaveChanges=0)
        
    def addGraph(self, graph, attr=None):
        """
        添加图形
        """
        # # 获取页面宽度（单位：磅，1英寸=72磅）
        # page_setup = self.doc.PageSetup
        # page_width = page_setup.PageWidth  # 页面总宽度
        # left_margin = page_setup.LeftMargin
        # right_margin = page_setup.RightMargin
        # available_width = (page_width - left_margin - right_margin)  # 实际可用宽度

        # # 获取图片原始尺寸以保持清晰度
        # with Image.open(graph['graph']) as img:
        #     original_width, original_height = img.size

        tmp = self.doc.Paragraphs.Add()
        #在当前的段落中插入图片
        grp = tmp.Range.InlineShapes.AddPicture(graph['graph'], LinkToFile=False, SaveWithDocument=True)
        # grp.Width = available_width
        # grp.Height = original_height

        # attr['InlineShape.Width'] = grp.Width
        # attr['InlineShape.ScaleWidth'] = int(original_width // available_width)

        self.setAttributes(grp, attr)
        self.setAttributes(tmp.Range, attr)

        return grp

    def appendGraph(self, graph, attr=None):
        """
        添加图形
        """
        locx = self.doc.Range().End

        # grp = self.addGraph(graph, attr)
        tmp = self.doc.Paragraphs.Add()
        #在当前的段落中插入图片
        grp = tmp.Range.InlineShapes.AddPicture(graph['graph'], LinkToFile=False, SaveWithDocument=True)

        self.setAttributes(tmp.Range, attr)
    
        locy = self.doc.Range().End
        tmp = self.doc.Range(locx - 2, locy)
        tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindContinue, Replace=cs.wdReplaceOne, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)

        self.setAttributes(grp, attr)
        self.setAttributes(tmp, attr)

        return grp
    
    def addGraphWithTable(self, graph, attr=None):
        """
        添加图形
        """
        xmp = self.doc.Paragraphs.Add()
        zmp = self.doc.Paragraphs.Add()
        ymp = self.doc.Paragraphs.Add()
        table = self.doc.Tables.Add(Range=zmp.Range, NumRows=2, NumColumns=1)

        tmp = table.Cell(1, 1)
        grp = tmp.Range.InlineShapes.AddPicture(graph['graph'], LinkToFile=False, SaveWithDocument=True)
        self.setAttributes(tmp.Range, attr)

        tmp = table.Cell(2, 1)
        ctx = f"{graph['title']}"
        ctx = self.replaceTags(ctx, attr, tmp)
        tmp.Range.Text = ctx
        self.setAttributes(tmp.Range, attr)

        xmp.Range.Delete()
        ymp.Range.Delete()

        self.setAttributes(table, attr)
        self.setAttributes(grp, attr)
        # self.setAttributes(zmp.Range, attr)

        return grp
    
    def addNote(self, note, attr=None):
        """
        添加注释
        """
        # 9312 - 9321 对应 1 - 10
        tmp = self.doc.Paragraphs.Add()
        # 在当前的段落中插入注释
        ctx = note['nid']
        ctx = self.replaceTags(ctx, attr, tmp)
        tmp.Range.InsertSymbol(CharacterNumber=9311 + eval(ctx), Font=attr['Font.NameFarEast'], Unicode=True)
        tmp.Range.InsertAfter(f"{note['note']}")
        self.setAttributes(tmp.Range, attr)
        xmp = self.doc.Paragraphs.Add()

        return tmp
    
    def addNotes(self, notes, attr=None):
        """
        添加注释
        """
        for nid in self.lsNote:
            note = copy(ADD_NOTE)
            note.note = NTN[nid.split('.')[1]]
            note.nid = nid
            note = vars(note)
            self.addNote(note, attr)

    def appendNote(self, note, attr=None):
        """
        添加注释
        # 9312 - 9321 对应 1 - 10
        """
        # 获取文档当前范围的结束位置
        locx = self.doc.Range().End
        tmp = self.doc.Paragraphs.Add()
        # 在当前的段落中插入注释
        ctx = note['nid']
        ctx = self.replaceTags(ctx, attr, tmp)
        tmp.Range.InsertSymbol(CharacterNumber=9311 + eval(ctx), Font=attr['Font.NameFarEast'], Unicode=True)
        self.setAttributes(tmp.Range, attr)
        xmp = self.doc.Paragraphs.Add()
        locy = self.doc.Range().End
        tmp = self.doc.Range(locx - 2, locy)
        tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindContinue, Replace=cs.wdReplaceOne, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)
        self.setAttributes(tmp, attr)

        return tmp
    
    def addReferenceLabel(self, content, attr=None):
        """
        添加文中引用
        """
        # 在文档中添加一个新的段落
        tmp = self.doc.Paragraphs.Add()
        # 在段落中插入内容文本
        ctx = content['words']
        ctx = self.replaceTags(ctx, attr, tmp)
        if 'Reference.SelfLabelStyle' in attr:
            if len(attr['Reference.SelfLabelStyle']) == 1:
                ctx = f"{attr['Reference.SelfLabelStyle']}{ctx}"
            elif len(attr['Reference.SelfLabelStyle']) == 2:
                ctx = f"{attr['Reference.SelfLabelStyle'][0]}{ctx}{attr['Reference.SelfLabelStyle'][1]}"
            else:
                ctx = f"{attr['Reference.SelfLabelStyle'][0]}{ctx}{attr['Reference.SelfLabelStyle'][1 : ]}"
        tmp.Range.InsertBefore(f"{ctx}")
        # 如果提供了属性字典，则设置段落的属性
        self.setAttributes(tmp.Range, attr)
        # 返回添加段落内容后的段落对象
        return tmp

    def appendReferenceLabel(self, content, attr=None):
        """
        添加文中引用
        """
        locx = self.doc.Range().End
        tmp = self.doc.Paragraphs.Add()
        # 在段落中插入内容文本
        ctx = content['words']
        ctx = self.replaceTags(ctx, attr, tmp)
        if 'Reference.SelfLabelStyle' in attr:
            if len(attr['Reference.SelfLabelStyle']) == 1:
                ctx = f"{attr['Reference.SelfLabelStyle']}{ctx}"
            elif len(attr['Reference.SelfLabelStyle']) == 2:
                ctx = f"{attr['Reference.SelfLabelStyle'][0]}{ctx}{attr['Reference.SelfLabelStyle'][1]}"
            else:
                ctx = f"{attr['Reference.SelfLabelStyle'][0]}{ctx}{attr['Reference.SelfLabelStyle'][1 : ]}"
        tmp.Range.InsertBefore(f"{ctx}")
        # 如果提供了属性字典，则设置段落的属性
        self.setAttributes(tmp.Range, attr)
        # 返回添加段落内容后的段落对象
        locy = self.doc.Range().End
        tmp = self.doc.Range(locx - 2, locy)
        tmp.Find.Execute(FindText='^p', ReplaceWith='', Wrap=cs.wdFindContinue, Replace=cs.wdReplaceOne, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)
        return tmp

    def addReferences(self, references, attr=None):
        """
        添加引用
        """
        if references['sort']:
            drf = [i.split('.')[1] for i in self.lsReference]
            lrf = []
            for kx, vx in REF.items():
                if kx in drf:
                    lrf.append(vx['ctx'])
            lrf.sort()
            for idx, rid in enumerate(lrf):
                tmp = self.doc.Paragraphs.Add()
                if 'Reference.SelfInlineNumberStyle' in attr:
                    tmp.Range.Fields.Add(Range=tmp.Range, Text=f"= {idx + 1} \* {attr['Reference.SelfInlineNumberStyle']}")
                    self.doc.Paragraphs.Add()
                    px = copy(APPEND_PARAGRAPH)
                    px.attr = attr
                    px.words = rid
                    self.appendParagraph(vars(px), attr)
                else:
                    if 'Reference.SelfNumberStyle' in attr:
                        if len(attr['Reference.SelfNumberStyle']) == 1:
                            ctx = f"{idx + 1}{attr['Reference.SelfNumberStyle']}\t{rid}"
                        else:
                            ctx = f"{attr['Reference.SelfNumberStyle'][0]}{idx + 1}{attr['Reference.SelfNumberStyle'][1 : ]}\t{rid}"
                    else:
                        ctx = f"[{idx + 1}]\t{rid}"
                    tmp.Range.InsertBefore(f"{ctx}")
                    # 如果提供了属性字典，则设置段落的属性
                    self.setAttributes(tmp.Range, attr)
                    # 返回添加段落内容后的段落对象

        else:
            for idx, rid in enumerate(self.lsReference):
                tmp = self.doc.Paragraphs.Add()
                if 'Reference.SelfInlineNumberStyle' in attr:
                    tmp.Range.Fields.Add(Range=tmp.Range, Text=f"= {idx + 1} \* {attr['Reference.SelfInlineNumberStyle']}")
                    self.doc.Paragraphs.Add()
                    px = copy(APPEND_PARAGRAPH)
                    px.attr = attr
                    px.words = f"{REF[rid.split('.')[1]]['ctx']}"
                    self.appendParagraph(vars(px), attr)
                else:
                    if 'Reference.SelfNumberStyle' in attr:
                        if len(attr['Reference.SelfNumberStyle']) == 1:
                            ctx = f"{idx + 1}{attr['Reference.SelfNumberStyle']}\t{REF[rid.split('.')[1]]['ctx']}"
                        else:
                            ctx = f"{attr['Reference.SelfNumberStyle'][0]}{idx + 1}{attr['Reference.SelfNumberStyle'][1 : ]}\t{REF[rid.split('.')[1]]['ctx']}"
                    else:
                        ctx = f"[{idx + 1}]\t{REF[rid.split('.')[1]]['ctx']}"
                    tmp.Range.InsertBefore(f"{ctx}")
                    # 如果提供了属性字典，则设置段落的属性
                    self.setAttributes(tmp.Range, attr)
                    # 返回添加段落内容后的段落对象

    def replaceRangeTags(self, rangex, attr, obj=None):
        """
        替换文本中的标记
        """
        # 换行符替换
        if "@@_\_n" in rangex.Text:
            rangex.Find.Execute(FindText='@@_\_n', ReplaceWith='\n', Wrap=cs.wdFindStop, Replace=cs.wdReplaceAll, Format=False, Forward=True, MatchWildcards=False, MatchSoundsLike=False, MatchAllWordForms=False, MatchCase=False, MatchWholeWord=False)

    def replaceTags(self, ctx, attr, obj=None):
        """
        替换文本中的标记
        """
        # 换行符替换
        if "@@_\_n" in ctx:
            ctx = ctx.replace("@@_\_n", "\n")

        # 参考文献标记替换
        ref = argparse.Namespace(**REF)
        for key, value in ref.__dict__.items():
            if f"@@REF.{key}" in ctx:
                if f"@@REF.{key}" in self.lsReference:
                    idx = self.lsReference.index(f"@@REF.{key}") + 1
                else:
                    idx = len(self.lsReference) + 1
                    self.lsReference.append(f"@@REF.{key}")
                ctx = ctx.replace(f"@@REF.{key}_id", f"{idx}")
                if type(value['tag']) == str:
                    ctx = ctx.replace(f"@@REF.{key}", f"{value['tag']}")
                else:
                    for tdx, tag in enumerate(value['tag'], start=1):
                        ctx = ctx.replace(f"@@REF.{key}_tag{tdx}", f"{tag}")

        # 公式标记替换
        fml = argparse.Namespace(**FML)
        for key, value in fml.__dict__.items():
            if f"@@FML.{key}" in ctx:
                if f"@@FML.{key}" in self.lsFormula:
                    idx = self.lsFormula.index(f"@@FML.{key}") + 1
                else:
                    idx = len(self.lsFormula) + 1
                    self.lsFormula.append(f"@@FML.{key}")
                ctx = ctx.replace(f"@@FML.{key}", f"{idx}")

        # 表格标记替换
        tbl = argparse.Namespace(**TBL)
        for key, value in tbl.__dict__.items():
            if f"@@TBL.{key}" in ctx:
                if f"@@TBL.{key}" in self.lsTable:
                    idx = self.lsTable.index(f"@@TBL.{key}") + 1
                else:
                    idx = len(self.lsTable) + 1
                    self.lsTable.append(f"@@TBL.{key}")
                ctx = ctx.replace(f"@@TBL.{key}", f"{idx}")

        # 图片标记替换
        grp = argparse.Namespace(**GRP)
        for key, value in grp.__dict__.items():
            if f"@@GRP.{key}" in ctx:
                if f"@@GRP.{key}" in self.lsGraph:
                    idx = self.lsGraph.index(f"@@GRP.{key}") + 1
                else:
                    idx = len(self.lsGraph) + 1
                    self.lsGraph.append(f"@@GRP.{key}")
                ctx = ctx.replace(f"@@GRP.{key}", f"{idx}")
                
        # 注释标记替换
        ntn = argparse.Namespace(**NTN)
        for key, value in ntn.__dict__.items():
            if f"@@NTN.{key}" in ctx:
                if f"@@NTN.{key}" in self.lsNote:
                    idx = self.lsNote.index(f"@@NTN.{key}") + 1
                else:
                    idx = len(self.lsNote) + 1
                    self.lsNote.append(f"@@NTN.{key}")
                ctx = ctx.replace(f"@@NTN.{key}", f"{idx}")

        # 大纲级别
        for i in range(1, 11):
            if f"@@CH{i}" in ctx:
                ctx = ctx.replace(f"@@CH{i}", "")
                tag = attr[f'Heading{i}.SelfStyle']
                obj.Range.Fields.Add(Range=obj.Range, Text=f"= {self.headings[i]} \* {tag}")
                self.setAttributes(obj.Range, attr)
                self.doc.Paragraphs.Add()
            elif f"@@NH{i}" in ctx:
                ctx = ctx.replace(f"@@NH{i}", "")
                if self.headings[i] == -1:
                    self.headings[i] = attr[f'Heading{i}.SelfStart']
                else:
                    self.headings[i] += 1
                self.headings[i + 1 : ] = -1
                tag = attr[f'Heading{i}.SelfStyle']
                obj.Range.Fields.Add(Range=obj.Range, Text=f"= {self.headings[i]} \* {tag}")
                self.setAttributes(obj.Range, attr)
                self.doc.Paragraphs.Add()

        return ctx
    
    def insertPageBreak(self):
        """
        插入新页面
        """
        self.doc.Paragraphs.Add().Range.InsertBreak(cs.wdPageBreak)

    def insertSectionBreak(self):
        """
        插入新段落
        """
        self.doc.Paragraphs.Add().Range.InsertBreak(cs.wdSectionBreakNextPage)

    def insertSectionBreakWithNewPageNumber(self, attr=None):
        """
        插入新段落并新起一页
        """
        self.doc.Paragraphs.Add().Range.InsertBreak(cs.wdSectionBreakNextPage)
        # 获取新创建的节（最后一个节）
        new_section = self.doc.Sections(self.doc.Sections.Count)

        footer = new_section.Footers(attr['Footers.SelfType'])
        footer.PageNumbers.Add(PageNumberAlignment=attr['PageNumber.Alignment'], FirstPage=attr['PageNumbers.ShowFirstPageNumber'])
        self.setAttributes(footer, attr)
        self.setAttributes(footer.Range, attr)

        # 设置页码格式
        # 首先，断开与前一节的页眉页脚链接
        new_section.Headers(1).LinkToPrevious = False  # 1 = wdHeaderFooterPrimary
        new_section.Footers(1).LinkToPrevious = False
        
        # 设置页码从指定数字开始
        self._gotoEnd()
        # 获取活动窗口
        active_window = self.doc.Application.ActiveWindow
        
        # 如果窗口有拆分视图，关闭第二个窗格
        # wdPaneNone = 0
        if active_window.View.SplitSpecial != cs.wdPaneNone:  # 0 = wdPaneNone
            active_window.Panes(2).Close()
        
        # 如果当前视图是普通视图或大纲视图，切换到页面视图
        # wdNormalView = 1, wdOutlineView = 2, wdPrintView = 3
        active_pane = active_window.ActivePane
        if active_pane.View.Type == cs.wdNormalView or active_pane.View.Type == cs.wdOutlineView:  # wdNormalView or wdOutlineView
            active_pane.View.Type = cs.wdPrintView
        
        # 切换到当前页页脚视图
        # wdSeekCurrentPageFooter = 9
        active_pane.View.SeekView = cs.wdSeekCurrentPageFooter
        
        # 获取Selection对象
        selection = self.doc.Application.Selection
        
        # 设置页码属性
        page_numbers = selection.HeaderFooter.PageNumbers
        page_numbers.NumberStyle = cs.wdPageNumberStyleArabic  # wdPageNumberStyleArabic = 0 (阿拉伯数字)
        page_numbers.HeadingLevelForChapter = 0
        page_numbers.IncludeChapterNumber = False
        page_numbers.ChapterPageSeparator = cs.wdSeparatorHyphen  # wdSeparatorHyphen = 0
        page_numbers.RestartNumberingAtSection = True
        page_numbers.StartingNumber = 1
        
        # 返回到主文档视图
        # wdSeekMainDocument = 0
        active_pane.View.SeekView = cs.wdSeekMainDocument

    def insertToc(
        self, attr=None, 
        upper_level=1, 
        lower_level=3,
        use_hyperlinks=True,
        use_heading_styles=True,
    ):
        """
        插入目录
        """
        # self.insertPageBreak()
        self._gotoEnd()

        # 获取当前位置
        range_obj = self.selection.Range

        # 插入目录
        toc = self.doc.TablesOfContents.Add(
            Range=range_obj,
            UseHeadingStyles=use_heading_styles,
            UpperHeadingLevel=upper_level,
            LowerHeadingLevel=lower_level,
            UseFields=True,
            UseHyperlinks=use_hyperlinks,
            HidePageNumbersInWeb=False,
            IncludePageNumbers=True,
            RightAlignPageNumbers=True,
            UseOutlineLevels=True
        )

        # xattr = self.attr.copy()
        attr.pop('ParagraphFormat.LeftIndent', None)
        attr.pop('ParagraphFormat.RightIndent', None)
        attr.pop('ParagraphFormat.FirstLineIndent', None)
        attr.pop('ParagraphFormat.CharacterUnitLeftIndent', None)
        attr.pop('ParagraphFormat.CharacterUnitRightIndent', None)
        attr.pop('ParagraphFormat.CharacterUnitFirstLineIndent', None)
        self.setAttributes(toc.Range, attr)
        # self.attr.update(xattr)

        toc.Range.Select()
        self.app.Selection.Cut()

        # 获取 Word 应用的 Selection 对象
        selection = self.doc.Application.Selection

        # 确保从文档开始查找
        selection.HomeKey(Unit=6)  # 6 = wdStory（整个文档）
        
        # 获取查找对象
        find_obj = selection.Find
        find_obj.ClearFormatting()
        
        # 设置查找参数
        find_obj.Text = "@@TOC"
        find_obj.Forward = True
        find_obj.Wrap = 1  # 1 = wdFindContinue，查找到文档末尾后从头继续
        find_obj.MatchCase = False  # 不区分大小写
        find_obj.MatchWholeWord = False
        find_obj.MatchWildcards = False
        find_obj.MatchSoundsLike = False
        find_obj.MatchAllWordForms = False
        
        # 执行查找
        found = find_obj.Execute()

        self.app.Selection.Paste()

    def insertTocGraph(self, attr):
        """
        插入图目录
        """
        # self.insertPageBreak()
        self._gotoEnd()

        locx = self.doc.Range().End

        for kx, vx in self.toc['graph'].items():
            tmp = "%s:.<%s%s%s:.>%s%s" % ('{kx', attr['Toc.Width'], '}', '{vx', 4, '}')
            px = copy(ADD_PARAGRAPH)
            px.words = tmp.format(kx=kx, vx=vx)
            px = vars(px)
            self.addParagraph(px, attr)

        locy = self.doc.Range().End
        tmp = self.doc.Range(locx - 1, locy - 1)
        tmp.Select()
        self.app.Selection.Cut()

        # 获取 Word 应用的 Selection 对象
        selection = self.doc.Application.Selection

        # 确保从文档开始查找
        selection.HomeKey(Unit=6)  # 6 = wdStory（整个文档）
        
        # 获取查找对象
        find_obj = selection.Find
        find_obj.ClearFormatting()
        
        # 设置查找参数
        find_obj.Text = "@@FIG_TOC"
        find_obj.Forward = True
        find_obj.Wrap = 1  # 1 = wdFindContinue，查找到文档末尾后从头继续
        find_obj.MatchCase = False  # 不区分大小写
        find_obj.MatchWholeWord = False
        find_obj.MatchWildcards = False
        find_obj.MatchSoundsLike = False
        find_obj.MatchAllWordForms = False
        
        # 执行查找
        found = find_obj.Execute()

        self.app.Selection.Paste()

    def insertTocTable(self, attr):
        """
        插入表目录
        """
        # self.insertPageBreak()
        self._gotoEnd()

        locx = self.doc.Range().End

        for kx, vx in self.toc['table'].items():
            tmp = "%s:.<%s%s%s:.>%s%s" % ('{kx', attr['Toc.Width'], '}', '{vx', 4, '}')
            px = copy(ADD_PARAGRAPH)
            px.words = tmp.format(kx=kx, vx=vx)
            px = vars(px)
            self.addParagraph(px, attr)

        locy = self.doc.Range().End
        tmp = self.doc.Range(locx - 1, locy - 1)
        tmp.Select()
        self.app.Selection.Cut()

        # 获取 Word 应用的 Selection 对象
        selection = self.doc.Application.Selection

        # 确保从文档开始查找
        selection.HomeKey(Unit=6)  # 6 = wdStory（整个文档）
        
        # 获取查找对象
        find_obj = selection.Find
        find_obj.ClearFormatting()
        
        # 设置查找参数
        find_obj.Text = "@@TBL_TOC"
        find_obj.Forward = True
        find_obj.Wrap = 1  # 1 = wdFindContinue，查找到文档末尾后从头继续
        find_obj.MatchCase = False  # 不区分大小写
        find_obj.MatchWholeWord = False
        find_obj.MatchWildcards = False
        find_obj.MatchSoundsLike = False
        find_obj.MatchAllWordForms = False
        
        # 执行查找
        found = find_obj.Execute()

        self.app.Selection.Paste()

    def updateToc(self, tocType='title'):
        """
        更新目录
        tocType: str, 目录类型，可选值为 'title', 'graph', 'table'
        """
        range = self.doc.Paragraphs(self.doc.Paragraphs.Count - 1).Range
        range.Select()
        selection = self.doc.Application.Selection
        page_number = selection.Information(3)
        self.toc[tocType][range.Text[ : -1]] = page_number

    def execute(self, flow):
        """
        执行流程
        """
        # flow = [i for i in map(vars, flow)]
        flowx = vars(flow)
        # flowx['attr'] = vars(flowx['attr'])
        # exec(f"self.{flowx['_tag']}({flowx}, {flowx['attr']})")
        func = eval(f"self.{flowx['_tag']}")
        func(flowx, flowx['attr'])
    
    def create(self, lsFlow):
        """
        批量执行流程
        """
        for flow in tqdm(lsFlow, desc='Creating'):
            try:
                self.execute(flow)
            except Exception as e:
                print(f"ERROR: {flow}")
                print(f'{trb.format_exc()}')

    def save(self, path='', name='paper.docx'):
        """"""
        if path == '':
            self.doc.Save()
        else:
            if not os.path.exists(path):
                os.makedirs(path)
            if os.path.exists(os.path.join(path, name)):
                tmp = self.app.Documents.Open(os.path.join(path, name))
                tmp.Close(SaveChanges=0)
            self.doc.SaveAs(os.path.join(path, name))

    def close(self):
        """"""
        self.doc.Close()
        self.app.Quit()

    def saveAndClose(self, path='', name='paper.docx'):
        """"""
        self.save(path, name)
        self.close()

    # ======================================================================
    # def addTitle(self, title, attr=None):
    #     """
    #     添加标题
    #     """
    #     # 在文档中添加一个新的段落
    #     tmp = self.doc.Paragraphs.Add()
    #     # 在段落中插入标题文本
    #     tmp.Range.InsertBefore(f"{title['title']}")
    #     # 如果提供了属性字典，则设置段落的属性
    #     self.setAttributes(tmp.Range, attr)
    #     # 返回添加标题后的段落对象
    #     return tmp
    
    # def addTitleSuperscript(self, title, attr=None):
    #     """
    #     添加标题上标
    #     """
    #     # 获取文档中段落的数量
    #     # nx = len(self.doc.Paragraphs)
    #     # 在文档末尾添加一个新的段落
    #     # tmp = self.doc.Paragraphs(nx)
    #     # 获取文档的范围对象
    #     tmp = self.doc.Range()
    #     # 获取当前范围的结束位置
    #     blockStart = tmp.End
    #     # 将范围移动到文档末尾
    #     tmp.Move(blockStart)
    #     # 在文档末尾插入标题文本
    #     tmp.InsertAfter(f"{title['words']}")
    #     # 再次获取文档的范围对象
    #     tmp = self.doc.Range()
    #     # 获取当前范围的结束位置
    #     blockEnd = tmp.End
    #     # 获取包含标题文本的范围对象
    #     tmp = self.doc.Range(blockStart - 1, blockEnd)
    #     # 如果提供了属性字典，则设置段落的属性
    #     self.setAttributes(tmp, attr)
    #     # 返回包含标题文本的范围对象
    #     return tmp
    
    # def addAbstract(self, abstract, attr=None):
    #     """
    #     添加摘要
    #     """
    #     tmp = self.doc.Paragraphs.Add()
    #     tmp.Range.InsertBefore(f"{abstract['words']}")
    #     self.setAttributes(tmp.Range, attr)
    #     return tmp
    
    # def addAbstractContent(self, content, attr=None):
    #     """
    #     添加摘要内容
    #     """
    #     locx = self.doc.Range().End
    #     tmp = self.doc.Paragraphs.Add()
    #     tmp.Range.InsertBefore(f"{content['words']}")
    #     self.setAttributes(tmp.Range, attr)
    #     locy = self.doc.Range().End
    #     tmp = self.doc.Range(locx - 2, locy)
    #     tmp.Find.Execute('^p', False, False, False, False, False, False, True, 1, '', 2) 
    #     return tmp

