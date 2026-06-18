# -*- coding: utf-8 -*-
"""生成「日本-精英全日制（含学期）」留学服务报价单 Excel。

价格、条款等为可编辑的模板占位值，请按机构实际情况调整。
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

OUT = "日本-全日制-含学期-精英全日制-报价单.xlsx"

# ---------- 样式 ----------
BRAND = "1F3864"        # 主色（深蓝）
BRAND_LIGHT = "D9E1F2"  # 浅蓝
ACCENT = "C9A227"       # 金（精英档强调）
GREY = "F2F2F2"

thin = Side(style="thin", color="BFBFBF")
border_all = Border(left=thin, right=thin, top=thin, bottom=thin)

def font(size=11, bold=False, color="000000", name="微软雅黑"):
    return Font(name=name, size=size, bold=bold, color=color)

def fill(color):
    return PatternFill("solid", fgColor=color)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center", wrap_text=True)

wb = Workbook()
ws = wb.active
ws.title = "报价单"
ws.sheet_view.showGridLines = False

# 列宽：A 序号 / B-C 项目 / D 说明 / E 数量 / F 金额
widths = [6, 18, 16, 40, 10, 16]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

r = 1

def merge(cell_range):
    ws.merge_cells(cell_range)

def setc(coord, value, f=None, a=None, fl=None, b=True):
    c = ws[coord]
    c.value = value
    if f: c.font = f
    if a: c.alignment = a
    if fl: c.fill = fl
    if b: c.border = border_all
    return c

# ---------- 标题区 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "留学服务报价单", font(20, True, "FFFFFF"), center, fill(BRAND), b=False)
ws.row_dimensions[r].height = 40
r += 1

merge(f"A{r}:F{r}")
setc(f"A{r}", "日本方向 · 精英全日制（含学期）服务套餐", font(13, True, "FFFFFF"), center, fill(ACCENT), b=False)
ws.row_dimensions[r].height = 26
r += 1

# ---------- 机构与报价信息 ----------
today = datetime.date.today()
valid = today + datetime.timedelta(days=30)
quote_no = "JP-" + today.strftime("%Y%m%d") + "-001"

info_rows = [
    ("机构名称：", "＿＿＿＿＿＿＿＿＿＿留学", "报价单号：", quote_no),
    ("联系顾问：", "＿＿＿＿＿＿", "报价日期：", today.strftime("%Y-%m-%d")),
    ("联系电话：", "＿＿＿＿＿＿＿＿", "有效期至：", valid.strftime("%Y-%m-%d")),
    ("客户姓名：", "＿＿＿＿＿＿", "意向专业：", "＿＿＿＿＿＿＿＿"),
]
for k1, v1, k2, v2 in info_rows:
    setc(f"A{r}", k1, font(10, True, BRAND), right, fill(GREY))
    merge(f"B{r}:C{r}")
    setc(f"B{r}", v1, font(10), left)
    setc(f"D{r}", k2, font(10, True, BRAND), right, fill(GREY))
    merge(f"E{r}:F{r}")
    setc(f"E{r}", v2, font(10), left)
    ws.row_dimensions[r].height = 22
    r += 1

r += 1  # 空行

# ---------- 套餐定位 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "一、套餐定位", font(12, True, "FFFFFF"), left, fill(BRAND))
ws.row_dimensions[r].height = 24
r += 1
merge(f"A{r}:F{r}")
desc = ("适用于申请日本【大学院·修士 / 研究生（旁听）】的学生，全程一对一全日制托管服务，"
        "含语言学校在读学期的衔接规划。从背景提升、研究计划书、教授套磁、出愿到面试、签证全流程负责，"
        "为冲刺名校（旧帝大 / 早庆上理 / 难关国公立）的学生设计。")
setc(f"A{r}", desc, font(10), left)
ws.row_dimensions[r].height = 50
r += 1
r += 1

# ---------- 服务内容明细 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "二、服务内容明细", font(12, True, "FFFFFF"), left, fill(BRAND))
ws.row_dimensions[r].height = 24
r += 1

headers = ["序号", "服务模块", "子项", "服务内容说明", "次数/周期", "备注"]
for i, h in enumerate(headers, start=1):
    setc(f"{get_column_letter(i)}{r}", h, font(10, True, "FFFFFF"), center, fill(BRAND))
ws.row_dimensions[r].height = 22
r += 1

services = [
    ("1", "前期规划", "背景评估与定位", "学术背景、语言能力、研究方向综合评估，制定整体申请时间轴", "1 次", "签约后启动"),
    ("1", "前期规划", "选校选专业", "结合成绩与目标，制定冲刺/匹配/保底校名单", "不限", "动态调整"),
    ("2", "语言与考试", "学期衔接规划", "语言学校在读期间的学习与升学双线规划、出勤与成绩跟踪", "全程", "含学期"),
    ("2", "语言与考试", "日语/英语备考指导", "N1/N2、托福/托业备考规划与节点提醒", "全程", "不含授课"),
    ("3", "背景提升", "研究方向梳理", "结合目标教授研究领域，明确研究课题方向", "不限", ""),
    ("3", "背景提升", "科研/实习建议", "针对性背景提升方案与资源推荐", "1 套", "资源另计"),
    ("4", "核心文书", "研究计划书", "一对一指导撰写并反复修改至定稿（中日/英）", "1 篇", "精修不限轮次"),
    ("4", "核心文书", "出愿文书", "志望理由书、简历、个人陈述等全套文书", "全套", ""),
    ("5", "教授套磁", "套磁信撰写", "教授匹配、套磁邮件撰写与往来沟通指导", "不限教授", "精英专属"),
    ("5", "教授套磁", "面谈/旁听协调", "协助预约教授面谈、研究室访问指导", "按需", ""),
    ("6", "出愿申请", "网申与材料", "出愿系统填写、材料清单核对、邮寄指导", "全部院校", "含多校"),
    ("6", "出愿申请", "成绩认证", "学历学位认证、成绩单开具指导（JASSO/各校）", "全程", "官费另计"),
    ("7", "面试辅导", "模拟面试", "专业面试题库、模拟面试与复盘", "不限次", "精英专属"),
    ("8", "录取与签证", "在留资格申请", "在留资格认定证明书申请材料指导", "1 次", ""),
    ("8", "录取与签证", "签证办理", "赴日签证材料准备与递交指导", "1 次", "领馆费另计"),
    ("9", "赴日后", "行前指导", "住宿、入学、银行卡、保险等行前说明", "1 次", "增值服务"),
]
for row in services:
    for i, val in enumerate(row, start=1):
        col = get_column_letter(i)
        a = center if i in (1, 5) else left
        setc(f"{col}{r}", val, font(9), a)
    ws.row_dimensions[r].height = 30
    r += 1
r += 1

# ---------- 价格 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "三、套餐报价", font(12, True, "FFFFFF"), left, fill(BRAND))
ws.row_dimensions[r].height = 24
r += 1

price_headers = ["套餐档位", "服务内容", "服务周期", "原价(¥)", "优惠价(¥)", "备注"]
for i, h in enumerate(price_headers, start=1):
    setc(f"{get_column_letter(i)}{r}", h, font(10, True, "FFFFFF"), center, fill(BRAND))
ws.row_dimensions[r].height = 22
r += 1

price_rows = [
    ("精英全日制\n（含学期）", "上述全部服务\n全程一对一托管", "签约至入学\n(约12-18个月)", "￥＿＿＿＿", "￥＿＿＿＿", "本报价主推套餐"),
    ("标准全日制", "核心申请服务\n（不含背景提升资源）", "签约至入学", "￥＿＿＿＿", "￥＿＿＿＿", "可选"),
    ("单项服务", "研究计划书 / 套磁\n / 面试 单项", "按项目", "￥＿＿＿＿", "—", "可选"),
]
for idx, row in enumerate(price_rows):
    for i, val in enumerate(row, start=1):
        col = get_column_letter(i)
        a = center if i in (1, 3, 4, 5) else left
        f = font(10, True, ACCENT) if (idx == 0 and i == 1) else font(9)
        fl = fill(BRAND_LIGHT) if idx == 0 else None
        setc(f"{col}{r}", val, f, a, fl)
    ws.row_dimensions[r].height = 46
    r += 1
r += 1

# ---------- 付款方式 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "四、付款方式", font(12, True, "FFFFFF"), left, fill(BRAND))
ws.row_dimensions[r].height = 24
r += 1
pay = [
    "1. 签约首付：合同金额的 50%，签订服务协议时支付。",
    "2. 第二期款：教授内诺 / 出愿完成后支付 30%。",
    "3. 尾款：收到录取（合格通知）后支付 20%。",
    "4. 支持对公转账 / 微信 / 支付宝，款项以服务协议约定为准。",
]
for line in pay:
    merge(f"A{r}:F{r}")
    setc(f"A{r}", line, font(10), left)
    ws.row_dimensions[r].height = 20
    r += 1
r += 1

# ---------- 退费与保障 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "五、退费与服务保障", font(12, True, "FFFFFF"), left, fill(BRAND))
ws.row_dimensions[r].height = 24
r += 1
refund = [
    "1. 服务保障：全程一对一专属顾问 + 专业文书老师 + 日语外教（如套餐含）。",
    "2. 录取保障：若未获得任何院校录取，按服务协议约定办理退费（具体比例以合同为准）。",
    "3. 阶段退费：不同服务阶段对应不同退费比例，详见正式服务协议。",
    "4. 不含费用：各院校报名费、考试费、认证费、签证领馆费、第三方资源费等官方/第三方费用。",
]
for line in refund:
    merge(f"A{r}:F{r}")
    setc(f"A{r}", line, font(10), left)
    ws.row_dimensions[r].height = 22
    r += 1
r += 1

# ---------- 说明 ----------
merge(f"A{r}:F{r}")
setc(f"A{r}", "六、其他说明", font(12, True, "FFFFFF"), left, fill(BRAND))
ws.row_dimensions[r].height = 24
r += 1
notes = [
    "• 本报价单为意向沟通参考，最终以双方签署的《留学服务协议》为准。",
    "• 价格有效期 30 天，逾期请重新确认；优惠政策不与其他活动叠加。",
    "• 报价含税与否、发票事宜请与顾问确认。",
]
for line in notes:
    merge(f"A{r}:F{r}")
    setc(f"A{r}", line, font(9, color="595959"), left)
    ws.row_dimensions[r].height = 20
    r += 1
r += 1

# ---------- 签字区 ----------
merge(f"A{r}:C{r}")
setc(f"A{r}", "客户签字：________________", font(10), left, fill(GREY))
merge(f"D{r}:F{r}")
setc(f"D{r}", "机构（盖章）：________________", font(10), left, fill(GREY))
ws.row_dimensions[r].height = 36
r += 1
merge(f"A{r}:C{r}")
setc(f"A{r}", "日期：________________", font(10), left)
merge(f"D{r}:F{r}")
setc(f"D{r}", "日期：________________", font(10), left)
ws.row_dimensions[r].height = 30
r += 1

# 打印设置
ws.print_options.horizontalCentered = True
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4

wb.save(OUT)
print("saved:", OUT)
