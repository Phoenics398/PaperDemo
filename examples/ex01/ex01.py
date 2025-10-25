"""
Author: 何练（HE Lian, Leon HO）
Unit: 深圳职业技术大学（Shenzhen Polytechnic University, SZPU）
Email: Helian@szpu.edu.cn

本示例展示了 PaperDemo 操作 docx 的第一种方式。
"""

import os
import warnings
from copy import copy, deepcopy

from paperdemo.src.paper_docx1 import *
from paperdemo.constants.constant_docx1 import *

# 加载模板
import ex01_demo as dmo

warnings.filterwarnings("ignore")


# ======================================================================
CURRENT_FOLDER = os.path.abspath(os.path.dirname(__file__))

# -------------------- 工作流 --------------------
# 工作流：撰写顺序
FLOW = []

# -------------------- 实例化 --------------------
paper = PaperDocx()


# ======================================================================
# -------------------- 式样 --------------------
style = copy(STYLE)
style.attr = dmo.styleAttr

# -------------------- 页码 --------------------
page = copy(PAGE)
page.attr = dmo.pageAttr

# -------------------- 正文 --------------------
t101 = copy(ADD_PARAGRAPH)
t101.attr = dmo.mainTitleAttr
t101.words = "身边的微党课系列——"

t102 = copy(ADD_PARAGRAPH)
t102.attr = dmo.mainTitleAttr
t102.words = "刺激消费信心 坚定制度自信"

h101 = copy(ADD_PARAGRAPH)
h101.attr = dmo.heading1Attr
h101.words = "一、理论引领：习近平总书记关于消费的重要论述指引方向"

p001 = copy(ADD_PARAGRAPH)
p001.attr = dmo.contentAttr1
p001.words = "关于刺激消费的问题，以习近平总书记为核心的党中央，早就指明了方向。"

x001 = "总书记指出：“要坚决贯彻落实扩大内需战略规划纲要，尽快形成完整内需体系，着力扩大有收入支撑的消费需求。”这一重要论述从理论高度阐明了消费在经济运行中的核心价值。"
x002 = "总书记还强调：“中国是全球第二大消费市场，拥有全球最大规模中等收入群体，蕴含着巨大投资和消费潜力。”这不仅为我们认识消费的重要地位提供了根本遵循，更让我们对破解刚才提到的消费疲软问题充满信心。"
p002 = copy(ADD_PARAGRAPHS)
p002.attr = dmo.contentAttr1
p002.words = [x001, x002]

# ======================================================================
FLOW += [
    style, page, 
]    
FLOW += [
    t101, t102,
    h101,
    p001, p002,
]
paper.create(lsFlow=FLOW)

# ======================================================================
# 保存文档
paperName = f"{t101.words}{t102.words}"
fileName = f'ex01_{paperName}.docx'
paper.saveAndClose(path=CURRENT_FOLDER, name=fileName)
