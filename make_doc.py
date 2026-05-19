from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# 기본 여백 설정
section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(3.0)
section.right_margin = Cm(3.0)

# 기본 폰트 설정
style = doc.styles['Normal']
style.font.name = '맑은 고딕'
style.font.size = Pt(10)
style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

def set_font(run, bold=False, size=10, color=None):
    run.font.name = '맑은 고딕'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_paragraph(text='', align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, size=10, space_before=0, space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        run = p.add_run(text)
        set_font(run, bold=bold, size=size)
    return p

def set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), val.get('val', 'single'))
            border.set(qn('w:sz'), str(val.get('sz', 4)))
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), val.get('color', '000000'))
            tcBorders.append(border)
    tcPr.append(tcBorders)

def add_table_row(table, label, value, header=False):
    row = table.add_row()
    row.height = Cm(0.8)

    label_cell = row.cells[0]
    value_cell = row.cells[1]

    label_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    value_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if header:
        set_cell_bg(label_cell, 'D9E1F2')
        set_cell_bg(value_cell, 'D9E1F2')
    else:
        set_cell_bg(label_cell, 'F2F2F2')
        set_cell_bg(value_cell, 'FFFFFF')

    lp = label_cell.paragraphs[0]
    lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lp.paragraph_format.space_before = Pt(0)
    lp.paragraph_format.space_after = Pt(0)
    lr = lp.add_run(label)
    set_font(lr, bold=True, size=9)

    vp = value_cell.paragraphs[0]
    vp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    vp.paragraph_format.space_before = Pt(0)
    vp.paragraph_format.space_after = Pt(0)
    vp.paragraph_format.left_indent = Cm(0.2)
    vr = vp.add_run(value)
    set_font(vr, bold=False, size=9)

# ── 제목 ──
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(6)
title_p.paragraph_format.space_after = Pt(16)
title_run = title_p.add_run('저축공제 만기 해지 및 공제금 수령의 건')
set_font(title_run, bold=True, size=16)

# ── 두문 ──
intro_p = doc.add_paragraph()
intro_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
intro_p.paragraph_format.space_after = Pt(4)
intro_run = intro_p.add_run(
    '당사가 가입한 수협보험 저축공제 계약의 만기가 도래함에 따라, '
    '아래와 같이 해지를 추진하고자 하오니 검토 후 재가하여 주시기 바랍니다.'
)
set_font(intro_run, size=10)

# ── 아 래 ──
arae_p = doc.add_paragraph()
arae_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
arae_p.paragraph_format.space_before = Pt(4)
arae_p.paragraph_format.space_after = Pt(10)
arae_run = arae_p.add_run('- 아     래 -')
set_font(arae_run, bold=True, size=10)

# ── 1. 목적 ──
add_paragraph('1. 목적', bold=True, size=10, space_before=4, space_after=4)
p_purpose = doc.add_paragraph()
p_purpose.paragraph_format.left_indent = Cm(0.5)
p_purpose.paragraph_format.space_after = Pt(10)
r = p_purpose.add_run('저축공제 계약 만기 도래에 따라, 해당 계약의 해지 처리 및 만기공제금을 수령하고자 합니다.')
set_font(r, size=10)

# ── 2. 계약 개요 ──
add_paragraph('2. 계약 개요', bold=True, size=10, space_before=4, space_after=6)

tbl1 = doc.add_table(rows=0, cols=2)
tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl1.style = 'Table Grid'
tbl1.columns[0].width = Cm(4.5)
tbl1.columns[1].width = Cm(10.0)

contract_rows = [
    ('상품명', '무배당 Sh VIP저축공제2301 (거치형)'),
    ('증권번호', '9831-0425-8029'),
    ('계약자 / 수익자', '주식회사 엔알비'),
    ('피공제자', '김동우'),
    ('계약일자', '2023년 3월 16일'),
    ('만기일자', '2026년 3월 16일'),
    ('공제기간 / 납입방식', '3년 / 일시납'),
    ('납입공제료', '200,000,000원'),
    ('관리점', '수협은행 테헤란로 지점'),
]
for label, value in contract_rows:
    add_table_row(tbl1, label, value)

doc.add_paragraph().paragraph_format.space_after = Pt(6)

# ── 3. 해지 사유 ──
add_paragraph('3. 해지 사유', bold=True, size=10, space_before=4, space_after=4)
p_reason = doc.add_paragraph()
p_reason.paragraph_format.left_indent = Cm(0.5)
p_reason.paragraph_format.space_after = Pt(10)
r = p_reason.add_run(
    '계약 만기 도래에 따른 만기공제금을 수령하고, 수령 자금을 '
    '건설공제조합 출자금 납입 등 운영 자금으로 활용하고자 합니다.'
)
set_font(r, size=10)

# ── 4. 관련 차입금 및 질권 현황 ──
add_paragraph('4. 관련 차입금 및 질권 현황', bold=True, size=10, space_before=4, space_after=4)
p_loan_intro = doc.add_paragraph()
p_loan_intro.paragraph_format.left_indent = Cm(0.5)
p_loan_intro.paragraph_format.space_after = Pt(6)
r = p_loan_intro.add_run(
    '해당 저축공제 계약은 아래 차입금의 담보(질권설정)로 설정되어 있으나, '
    '해당 차입금의 만기가 완료되었습니다.'
)
set_font(r, size=10)

tbl2 = doc.add_table(rows=0, cols=2)
tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl2.style = 'Table Grid'
tbl2.columns[0].width = Cm(4.5)
tbl2.columns[1].width = Cm(10.0)

loan_rows = [
    ('구분', '장기차입금'),
    ('은행명', '수협은행'),
    ('대출 계좌', '3210-1065-9544'),
    ('최초 차입일', '2023년 5월 17일'),
    ('갱신 만기일', '2026년 5월 18일'),
    ('상환 방식', '원금균등분할상환'),
    ('이자율', '4.9500%'),
    ('최초 차입금', '1,000,000,000원'),
    ('질권설정금액', '197,854,009원'),
]
for label, value in loan_rows:
    add_table_row(tbl2, label, value)

doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ── 주의 문구 ──
p_notice = doc.add_paragraph()
p_notice.paragraph_format.left_indent = Cm(0.5)
p_notice.paragraph_format.space_before = Pt(4)
p_notice.paragraph_format.space_after = Pt(10)
r_star = p_notice.add_run('※ ')
set_font(r_star, bold=True, size=9, color=(192, 0, 0))
r_notice = p_notice.add_run(
    '질권설정금액(197,854,009원)은 상기 수협은행 차입금과 연계된 담보이므로, '
    '해지 절차 진행 전 반드시 질권 해제 절차를 완료하여야 합니다.'
)
set_font(r_notice, bold=False, size=9, color=(192, 0, 0))

# ── 5. 첨부문서 ──
add_paragraph('5. 첨부문서', bold=True, size=10, space_before=4, space_after=4)
attachments = [
    '① SH 저축공제 미래해지환급금 예상조회서 (기준일자: 2026. 04. 03.)',
    '② 가입 당시 원안 문서',
]
for att in attachments:
    p_att = doc.add_paragraph()
    p_att.paragraph_format.left_indent = Cm(0.5)
    p_att.paragraph_format.space_before = Pt(0)
    p_att.paragraph_format.space_after = Pt(3)
    r = p_att.add_run(att)
    set_font(r, size=10)

# ── 끝 ──
p_end = doc.add_paragraph()
p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_end.paragraph_format.space_before = Pt(20)
r = p_end.add_run('- 끝 -')
set_font(r, bold=True, size=10)

doc.save('/home/user/NRB/저축공제_만기해지_기안서.docx')
print('완료')
