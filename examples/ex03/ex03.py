"""
Author: 何练（HE Lian, Leon HO）
Unit: 深圳职业技术大学（Shenzhen Polytechnic University, SZPU）
Email: Helian@szpu.edu.cn

本示例展示了 PaperDemo 操作 docx 的第三种方式。
"""

import os
import warnings
from copy import copy, deepcopy

from paperdemo.src.paper_docx1 import *
from paperdemo.constants.constant_docx1 import *

# 加载模板
import ex03_demo as dmo

warnings.filterwarnings("ignore")


# ======================================================================
CURRENT_FOLDER = os.path.abspath(os.path.dirname(__file__))

# -------------------- 实例化 --------------------
paper = PaperDocx()

# -------------------- 预定义 --------------------
# 对于重复性使用的内容，可以预先定义
maps = {
    'wt1|': {'dict': ADD_PARAGRAPH, 'attr': dmo.mainTitleAttr},
    'wh1|': {'dict': ADD_PARAGRAPH, 'attr': dmo.heading1Attr},
    'wp1|': {'dict': ADD_PARAGRAPH, 'attr': dmo.contentAttr1},
}    
tmp = addFunctions(paper, maps)
vars().update(tmp)

# ======================================================================
# -------------------- 式样 --------------------
style = copy(STYLE)
style.attr = dmo.styleAttr

# -------------------- 页码 --------------------
page = copy(PAGE)
page.attr = dmo.pageAttr

# -------------------- 正文 --------------------
wt1("身边的微党课系列——")

wt1("刺激消费信心 坚定制度自信")

wh1("一、理论引领：习近平总书记关于消费的重要论述指引方向")

wp1("关于刺激消费的问题，以习近平总书记为核心的党中央，早就指明了方向。")

f"{wp1("总书记指出：“要坚决贯彻落实扩大内需战略规划纲要，尽快形成完整内需体系，着力扩大有收入支撑的消费需求。”这一重要论述从理论高度阐明了消费在经济运行中的核心价值。")}{wp1("总书记还强调：“中国是全球第二大消费市场，拥有全球最大规模中等收入群体，蕴含着巨大投资和消费潜力。”这不仅为我们认识消费的重要地位提供了根本遵循，更让我们对破解刚才提到的消费疲软问题充满信心。")}"

# ======================================================================
# 保存文档
paperName = f"身边的微党课系列——刺激消费信心 坚定制度自信"
fileName = f'ex03_{paperName}.docx'
paper.saveAndClose(path=CURRENT_FOLDER, name=fileName)
