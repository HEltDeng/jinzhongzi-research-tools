from pathlib import Path
import json
import re

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
OUT = ROOT / 'outputs/teacher-alumni-20261026'
FONT = 'Noto Sans CJK SC'
doc = Document()
section = doc.sections[0]
section.page_width = Cm(21.59)
section.page_height = Cm(27.94)
section.top_margin = Cm(1.8)
section.bottom_margin = Cm(1.8)
section.left_margin = Cm(2)
section.right_margin = Cm(2)
section.header_distance = Cm(.65)
section.footer_distance = Cm(.65)

for name, size, bold in [('Normal', 11, False), ('Title', 22, True),
                         ('Subtitle', 11, False), ('Heading 1', 16, True),
                         ('Heading 2', 13, True)]:
    style = doc.styles[name]
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.color.rgb = RGBColor(0, 0, 0)
    fonts = style._element.get_or_add_rPr().get_or_add_rFonts()
    for key in list(fonts.attrib):
        if 'theme' in key.lower():
            del fonts.attrib[key]
    for key in ['ascii', 'hAnsi', 'eastAsia', 'cs']:
        fonts.set(qn('w:' + key), FONT)
    for border in style._element.xpath('./w:pPr/w:pBdr'):
        border.getparent().remove(border)

normal = doc.styles['Normal'].paragraph_format
normal.line_spacing = Pt(17)
normal.space_after = Pt(7)
normal.widow_control = True
doc.styles['Title'].paragraph_format.space_after = Pt(5)
doc.styles['Subtitle'].paragraph_format.space_after = Pt(15)
for name in ['Heading 1', 'Heading 2']:
    pf = doc.styles[name].paragraph_format
    pf.keep_with_next = True
    pf.space_before = Pt(12)
    pf.space_after = Pt(8)

footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
footer.paragraph_format.space_after = Pt(0)
for instruction in ['PAGE', 'NUMPAGES']:
    if instruction == 'NUMPAGES':
        footer.add_run(' / ')
    field = OxmlElement('w:fldSimple')
    field.set(qn('w:instr'), instruction)
    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.text = '1'
    run.append(text)
    field.append(run)
    footer._p.append(field)

source = []

def para(text, style=None, keep=False):
    source.append(text)
    p = doc.add_paragraph(style=style)
    p.paragraph_format.keep_with_next = keep
    for fragment in re.split(r'(\*\*.*?\*\*)', text):
        if fragment:
            bold = fragment.startswith('**')
            r = p.add_run(fragment[2:-2] if bold else fragment)
            if bold:
                r.bold = True
    return p

def heading(text, level=1, new_page=False):
    source.append('\n' + '#' * (level + 1) + ' ' + text)
    p = doc.add_heading(text, level)
    if new_page:
        p.paragraph_format.page_break_before = True
        p.paragraph_format.space_before = Pt(0)
    return p

def comparison(rows):
    widths = [3.55, 7.15, 6.89]
    table = doc.add_table(rows=0, cols=3)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for col, width in zip(table.columns, widths):
        col.width = Cm(width)
    for index, values in enumerate(rows):
        source.append(' | '.join(values))
        row = table.add_row()
        properties = row._tr.get_or_add_trPr()
        properties.append(OxmlElement('w:cantSplit'))
        if index == 0:
            properties.append(OxmlElement('w:tblHeader'))
        for cell, value, width in zip(row.cells, values, widths):
            cell.width = Cm(width)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            pr = cell._tc.get_or_add_tcPr()
            borders = OxmlElement('w:tcBorders')
            for edge in ['top', 'left', 'bottom', 'right']:
                e = OxmlElement('w:' + edge)
                for key, val in [('val', 'single'), ('sz', '4'), ('color', 'D9D9D9')]:
                    e.set(qn('w:' + key), val)
                borders.append(e)
            pr.append(borders)
            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), 'E6EDF3' if index == 0 else 'FFFFFF')
            pr.append(shading)
            margins = OxmlElement('w:tcMar')
            for edge in ['top', 'left', 'bottom', 'right']:
                e = OxmlElement('w:' + edge)
                e.set(qn('w:w'), '100')
                e.set(qn('w:type'), 'dxa')
                margins.append(e)
            pr.append(margins)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(15.5)
            r = p.add_run(value)
            r.font.size = Pt(10.5)
            r.bold = index == 0
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(3)

para('第三届金种子杯选题汇报', 'Title')
para('精简版  2026 年 10 月 26 日', 'Subtitle')

para('**当前建议：优先在道路养护核验和充电站调度中确定一个主项目。**我们想把已有的软件开发能力用于交通、电气领域的具体任务，先做出可运行、能验证效果的小产品。隧道和材料方向是否上调，取决于老师能否帮助落实成套数据、试验条件和专业指导。')

heading('一 本次重点讨论的五个方案')
comparison([
    ['方向', '我们想做什么', '当前位置与关键条件'],
    ['道路养护核验', '核对同一维修位置的前后照片与复巡记录，辅助确认维修情况。', '优先候选\n需要连续影像、工单和工程人员复核。'],
    ['充电站调度', '在站点总功率受限时，改进功率分配，缩短等待并保障充电需求。', '优先候选\n需要站点数据、现行策略和运营方配合。'],
    ['隧道资料复核', '把同一施工段的掌子面、随钻等资料对齐，提示矛盾和缺项。', '有条件上调\n需要成套资料和工程师判断。'],
    ['再生材料试配', '针对旧料批次变化，推荐下一轮试验配方，减少无效试配。', '有条件上调\n需要实验室、历史记录和新批次试验。'],
    ['软件交付验收', '需求变更后，自动检查多个角色之间的关键业务流程是否仍然正确。', '自主启动备选\n需要外部团队验证需求与使用价值。'],
])

heading('二 选题调研形成的主要判断')
para('这些方向已经有相近产品或研究基础，机会在于把一个尚未解决好的具体任务做深。道路方向可聚焦维修证据核对，充电方向可聚焦真实容量限制下的调度收益；最终都要与用户现行做法比较。')
para('团队能够承担软件开发、数据处理、模型接入与产品演示；专业判断和真实验证依赖对应老师及业务方。目前已完成公开资料与竞品调研、方案设计，客户访谈、原始数据验证和试点合作仍待开展。')

heading('三 希望本次交流确定的事项')
para('**选方向：**请老师判断道路与充电哪个更有实际价值；若已有合适的隧道或材料课题，也请明确可切入的子任务。')
para('**落实验证：**确定最多两个候选方向的联系人、第一批数据或试验任务，以及判断方案值得继续的标准。随后用一至两周核实，收敛成一个主项目。')

heading('四 优先候选的具体想法', new_page=True)
heading('方案 A 道路养护核验', 2)
para('**我们的想法：**面向养护单位、检测机构和项目管理人员，把同一位置的维修前、维修后及复巡照片关联起来，辅助判断是否属于同一处病害、是否出现复发，以及哪些记录需要人工复核。核心交付是一份可追溯的核验记录。')
para('**选题调研：**现有产品已能识别道路病害、管理巡检与工单。我们拟聚焦跨时间、跨拍摄角度的维修证据核对，把位置、工单、影像变化和工程规则结合起来。需要先确认一线人员是否经常为这件事耗时、发生错配，现有系统是否已解决。')
para('**第一步怎么做：**争取一段道路的连续影像与约 30—50 条养护记录，做出资料导入、前后对照、疑点复核和结果导出四项功能。由工程人员核对关联结果，比较人工核验时间、错配和漏配情况；这批资料先用于核实需求和数据可用性。')
para('**需要考虑和请老师核对：**照片能支持哪些维修状态判断，哪些仍须现场检测？能否获取同一位置的连续资料，并请工程人员复核？谁使用、谁采购？若记录无法关联，应先调整数据采集方式；仅凭表面照片不能证明全部施工质量。')

heading('方案 B 充电站调度', 2)
para('**我们的想法：**面向充电站运营方，在站点总功率有限的条件下，根据车辆到达、离开和充电需求提出功率分配建议。先选择一种车辆和一个站点，做历史数据回放与调度建议，争取减少等待、降低成本并保障充电需求。')
para('**选题调研：**充电负荷管理已有成熟方案，单纯预测功率或展示大屏难以形成差异。我们的切口是把实际到站需求、容量约束和用户等待一起纳入调度。学校已有相关补能课题线索，值得请老师判断能否对接具体站点与任务。')
para('**第一步怎么做：**获取充电记录、枪位状态、站点容量、电价和现行策略；尽可能补充到站、排队及放弃充电的记录。在同一批数据上比较现行策略与新方案，观察等待时间、需求满足、电费和是否突破功率限制。')
para('**需要考虑和请老师核对：**运营方最在意排队、成本还是容量利用？数据能否支撑真实回放？只有已成交订单时，无法还原全部排队与流失需求。需要老师帮助核对电气约束、评价方法，并确认站点是否愿意提供数据及验证机会。')

heading('五 条件备选与自主启动路线', new_page=True)
heading('方案 C 隧道资料复核', 2)
para('**想法与调研判断：**把同一里程、同一施工循环的掌子面影像、随钻记录和地质资料关联，提示内容矛盾、缺项或异常。现有系统已有随钻数据分析和报告功能，我们需要验证的增量是资料对齐和异常解释能否帮助工程师更快复核。')
para('**需要老师核对：**是否有能对应到施工循环的成套数据，并能由工程师解释异常原因？首版可先做资料关联和疑点提示，再用人工复核结果评价。不能把掌子面照片直接当作前方隐伏地质或灾害预测依据；若只能拿到零散图片，暂不优先推进。')

heading('方案 D 再生材料试配', 2)
para('**想法与调研判断：**针对再生沥青旧料批次不同、性能波动的情况，结合材料与工艺限制，推荐下一轮最值得尝试的配方。已有实验记录软件和优化方法，差异应落实到批次变化下能否用更少试验找到合格配方。')
para('**需要老师核对：**是否有包含失败结果的历史试验记录、明确性能指标，以及可安排的新批次试验？首版可做试验推荐软件，与老师或工程师原有试配方法比较试验次数、成本和达标情况。只有落实实验室和实际试验，才考虑将其升为主方向。')

heading('方案 E 软件交付验收', 2)
para('**想法与调研判断：**面向小型开发团队，在需求变更后检查跨角色业务是否一致。例如用户取消预约后，名额是否释放、管理端记录是否同步、重复取消是否造成错误。现有自动化测试工具已很成熟，重点应放在业务规则与角色间状态核对。')
para('**启动方式与待核实点：**先用团队可控项目做两个关键流程，再请外部开发团队提供真实修改和历史缺陷，与人工验收及成熟工具比较。开发可自主启动，但需要确认是否节省验收时间、是否愿意持续使用或付费；请老师帮助核对评价方法和联系试用团队。')

heading('六 汇报后的推进安排')
para('**第一周：**按老师意见选定两个候选，分别访谈实际使用者、确认现行做法，并拿到一小批可用资料或明确试验安排。各整理一页需求、样本和验证记录。')
para('**第二周：**做最小可运行原型或第一轮试验，比较效果、实施成本和后续资源可得性，在 11 月上旬确定一个主项目。优先推进有明确问题、真实使用方和可验证改进的方向。')

doc.core_properties.title = '第三届金种子杯选题汇报 精简版'
doc.core_properties.subject = '老师汇报用候选方案与待核对事项'
doc.core_properties.author = ''
doc.core_properties.last_modified_by = ''
path = OUT / '03_第三届金种子杯选题汇报_精简版.docx'
doc.save(path)
(HERE / 'teacher-brief.txt').write_text('\n\n'.join(source), encoding='utf-8')
print(json.dumps({'path': str(path), 'characters': sum(map(len, source)),
                  'paragraphs': len(doc.paragraphs), 'tables': len(doc.tables)}, ensure_ascii=False))
