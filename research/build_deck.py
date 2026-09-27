"""Сборка презентации «Окупаемость модулей nanoCAD» на шаблоне Нормасофт.

Вход: skel.pptx — шаблон с уже продублированными слайдами (27 шт.):
  1 титул; 2,6,8,13,22 — разделители (из слайда «02 Учебный центр»);
  23,24,25,27 — слайды шаблона «О нас», «Учебный центр», «Компетенции», «Спасибо»;
  остальные — контентные (из слайда «Совместный план»).
Базовые цифры ROI взяты без изменений из трёх docx с формулами.
"""
import copy
import sys

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SRC, OUT = sys.argv[1], sys.argv[2]

BLUE = RGBColor(0x03, 0x86, 0xFC)
INK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0xF1, 0xF3, 0xF6)
MUTED = RGBColor(0x6B, 0x72, 0x80)
PALE = RGBColor(0xE3, 0xEF, 0xFE)
LINE = RGBColor(0xC9, 0xD3, 0xE0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
F_HEAD = "VK Sans Display DemiBold"
F_BODY = "Manrope Medium"
F_NAV = "Roboto Light"

prs = Presentation(SRC)
S = list(prs.slides)

# ---------------------------------------------------------------- helpers


def rm(shape):
    shape._element.getparent().remove(shape._element)


def by_name(slide, prefix):
    return [s for s in slide.shapes if s.name.startswith(prefix)]


def style_run(run, size, color=INK, font=F_BODY, bold=False):
    f = run.font
    f.size = Pt(size)
    f.name = font
    f.bold = bold
    f.color.rgb = color


def text(slide, x, y, w, h, paras, size=13, color=INK, font=F_BODY, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, bullet=False, space=4, line=None):
    """paras: str | list[str | list[(text, opts)]]"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", 0)
    if isinstance(paras, str):
        paras = [paras]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        para.space_after = Pt(space)
        if line:
            para.line_spacing = line
        runs = p if isinstance(p, list) else [(p, {})]
        for t, o in runs:
            r = para.add_run()
            r.text = t
            style_run(r, o.get("size", size), o.get("color", color), o.get("font", font), o.get("bold", False))
        if bullet:
            pPr = para._p.get_or_add_pPr()
            pPr.set("marL", str(Emu(Inches(0.27))))
            pPr.set("indent", str(-Emu(Inches(0.27))))
            buClr = etree.SubElement(pPr, qn("a:buClr"))
            etree.SubElement(buClr, qn("a:srgbClr")).set("val", "0386FC")
            etree.SubElement(pPr, qn("a:buFont")).set("typeface", F_BODY)
            etree.SubElement(pPr, qn("a:buChar")).set("char", "→")
    return tb


def box(slide, x, y, w, h, fill=GREY, line=None, shape=MSO_SHAPE.SNIP_1_RECTANGLE):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(1)
    s.shadow.inherit = False
    s.text_frame.text = ""
    return s


def pill(slide, x, y, w, label, active=True, h=0.42, size=13):
    s = box(slide, x, y, w, h, fill=BLUE if active else None, line=None if active else BLUE)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    style_run(r, size, WHITE if active else BLUE)
    return s


def num_badge(slide, x, y, label, size=0.5):
    s = box(slide, x, y, size, size, fill=BLUE)
    tf = s.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    style_run(r, 16, WHITE, F_HEAD)


def set_title(slide, title):
    t = by_name(slide, "Заголовок 1")[0]
    t.left, t.top, t.width, t.height = Inches(0.65), Inches(1.42), Inches(11.4), Inches(0.62)
    run = t.text_frame.paragraphs[0].runs[0]
    run.text = title
    run.font.size = Pt(24)
    for extra in t.text_frame.paragraphs[0].runs[1:]:
        extra._r.getparent().remove(extra._r)
    for body in by_name(slide, "Объект 2"):
        rm(body)


def notes(slide, txt):
    slide.notes_slide.notes_text_frame.text = txt


def source(slide, txt, y=6.62):
    text(slide, 0.65, y, 10.9, 0.3, txt, size=9, color=MUTED)


def style_chart(chart, size=10):
    chart.font.size = Pt(size)
    chart.font.name = F_BODY
    chart.font.color.rgb = INK


def set_paras(shape, lines):
    """Заменить текст, сохранив оформление первого абзаца."""
    tf = shape.text_frame
    first = tf.paragraphs[0]._p
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for r in first.findall(qn("a:r"))[1:]:
        first.remove(r)
    first.findall(qn("a:r"))[0].find(qn("a:t")).text = lines[0]
    prev = first
    for line in lines[1:]:
        np_ = copy.deepcopy(first)
        np_.findall(qn("a:r"))[0].find(qn("a:t")).text = line
        prev.addnext(np_)
        prev = np_


# ---------------------------------------------------------------- nav bar

NAV = [  # (label, target slide no., row)
    ("ВЫЗОВЫ", 2, 0), ("ЭКОНОМИЯ", 5, 0), ("АР и КР", 6, 0), ("ЭЛЕКТРО", 8, 0),
    ("СКС и ОПС", 10, 0), ("ДРУГИЕ РАЗДЕЛЫ", 12, 0),
    ("МЕТОДИКА", 13, 1), ("КЕЙСЫ", 18, 1), ("ОГРАНИЧЕНИЯ", 21, 1), ("НОРМАСОФТ", 22, 1),
    ("КОНТАКТЫ", 27, 1),
]
ACTIVE = {2: "ВЫЗОВЫ", 3: "ВЫЗОВЫ", 4: "ВЫЗОВЫ", 5: "ЭКОНОМИЯ", 6: "АР и КР", 7: "АР и КР",
          8: "ЭЛЕКТРО", 9: "ЭЛЕКТРО", 10: "СКС и ОПС", 11: "ЭКОНОМИЯ", 12: "ДРУГИЕ РАЗДЕЛЫ",
          13: "МЕТОДИКА", 14: "МЕТОДИКА", 15: "МЕТОДИКА", 16: "МЕТОДИКА", 17: "МЕТОДИКА",
          18: "КЕЙСЫ", 19: "КЕЙСЫ", 20: "КЕЙСЫ", 21: "ОГРАНИЧЕНИЯ", 22: "НОРМАСОФТ",
          23: "НОРМАСОФТ", 24: "НОРМАСОФТ", 25: "НОРМАСОФТ", 26: "НОРМАСОФТ", 27: "КОНТАКТЫ"}


def rebuild_nav(n, slide):
    for s in list(slide.shapes):
        if s.name.startswith(("Google Shape;229", "Google Shape;230", "Google Shape;231",
                              "Google Shape;232")):
            rm(s)
    for s in by_name(slide, "Google Shape;269"):   # иконка «домой» -> титул
        cNvPr = s._element.find(qn("p:nvSpPr")).find(qn("p:cNvPr"))
        for h in cNvPr.findall(qn("a:hlinkClick")):
            cNvPr.remove(h)
        s.click_action.target_slide = S[0]
    xs = [1.02, 1.02]
    for label, target, row in NAV:
        w = 0.34 + 0.105 * len(label)
        y = 0.24 if row == 0 else 0.70
        s = slide.shapes.add_shape(MSO_SHAPE.SNIP_1_RECTANGLE, Inches(xs[row]), Inches(y),
                                   Inches(w), Inches(0.37))
        s.name = "nav"
        s.shadow.inherit = False
        s.line.fill.background()
        active = ACTIVE.get(n) == label
        if active:
            s.fill.solid()
            s.fill.fore_color.rgb = WHITE
        else:
            s.fill.background()
        tf = s.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = label
        style_run(r, 11, BLUE if active else WHITE, F_NAV)
        s.click_action.target_slide = S[target - 1]
        xs[row] += w + 0.12


def drop_dead_slide_rels(slide):
    xml = etree.tostring(slide._element).decode()
    for rId, rel in list(slide.part.rels.items()):
        if rel.reltype.endswith("/slide") and f'"{rId}"' not in xml:
            slide.part.drop_rel(rId)


# ---------------------------------------------------------------- 1. титул
t = S[0]
pills = {s.shape_id: s for s in by_name(t, "Google Shape;321")}
for sid, label in {67: "Окупаемость", 68: "Модули", 65: "Обучение", 66: "Без 3D"}.items():
    set_paras(pills[sid], [label])
set_paras([s for s in t.shapes if s.shape_id == 70][0], ["Окупаемость модулей nanoCAD"])
sub = [s for s in t.shapes if s.shape_id == 69][0]
sub.top, sub.height = Inches(4.2), Inches(1.0)
set_paras(sub,
          ["Экономия времени и ФОТ проектного отдела — без перехода на 3D и полного переобучения штата"])
notes(t, "Презентация для руководителя проектной организации: сколько экономят специализированные "
         "модули nanoCAD и за сколько окупаются.")

# ---------------------------------------------------------------- разделители
DIVIDERS = {
    2: ("01", ["ТЯЖЁЛЫЕ ВРЕМЕНА —", "СОВРЕМЕННЫЕ РЕШЕНИЯ"],
        "Как сократить трудозатраты проектного отдела на 25–30% без перехода на 3D и полного переобучения штата"),
    6: ("02", ["ОКУПАЕМОСТЬ:", "АР и КР"],
        "Компонент СПДС Платформы nanoCAD: оформление, спецификации и ведомости по основным строительным разделам"),
    8: ("03", ["ОКУПАЕМОСТЬ:", "ЭЛЕКТРО, СКС, ОПС"],
        "Специализированные модули для инженерных разделов: расчёты, трассировка, журналы и спецификации"),
    13: ("04", ["МЕТОДИКА", "И ОБОСНОВАНИЕ"],
         "Базовые работы, экономия по этапам, подтверждение кейсами и ограничения расчёта"),
    22: ("05", ["НОРМАСОФТ:", "ВНЕДРЕНИЕ ПОД КЛЮЧ"],
         "Учебный центр, консультанты по всем разделам проектирования и собственные разработки"),
}
for n, (num, title, sub) in DIVIDERS.items():
    sl = S[n - 1]
    big = [s for s in by_name(sl, "Google Shape;529")]
    numshape = min(big, key=lambda s: s.width)
    titleshape = max(big, key=lambda s: s.width)
    set_paras(numshape, [num])
    titleshape.left, titleshape.width = Inches(4.2), Inches(8.22)
    set_paras(titleshape, title)
    for p in titleshape.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(34)
    subshape = by_name(sl, "TextBox 9")[0]
    subshape.left, subshape.width = Inches(5.1), Inches(7.32)
    set_paras(subshape, [sub])


# ---------------------------------------------------------------- 3. вызовы
sl = S[2]
set_title(sl, "Тяжёлые времена требуют современных решений")
pill(sl, 0.65, 2.2, 1.6, "Вызовы")
stats = [
    ("146 684 ₽", "средняя зарплата в проектировании (ОКВЭД 71, ЕМИСС, январь–апрель 2026): час инженера дорожает"),
    ("5,5%", "доля отрицательных заключений госэкспертизы в 2025 году — +1,8 п.п. к 2024 году"),
    ("> 70%", "всех переделок в строительстве вызваны ошибками проектирования (Love & Edwards)"),
]
y = 2.85
for big, desc in stats:
    text(sl, 0.65, y, 2.3, 0.6, big, size=26, color=BLUE, font=F_HEAD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, 3.0, y, 3.2, 0.9, desc, size=12, anchor=MSO_ANCHOR.MIDDLE)
    y += 1.18
pill(sl, 6.85, 2.2, 1.6, "Решение")
box(sl, 6.85, 2.85, 5.7, 3.55, fill=GREY)
text(sl, 7.15, 3.05, 5.1, 0.8, "Специализированные модули nanoCAD на привычной 2D-платформе",
     size=17, font=F_HEAD)
text(sl, 7.15, 3.95, 5.1, 2.4, [
    "Автоматизация рутины: спецификации, ведомости, кабельные журналы, расчёты",
    "Сокращение трудозатрат раздела на 25–30% — консервативная оценка",
    "Окупаемость лицензии и обучения — за 1–3 проекта",
    "Российское ПО: лицензионная чистота и защита данных",
], size=13, bullet=True, space=8)
source(sl, "Источники: ЕМИСС; итоги Главгосэкспертизы за 2025 год; Love P., Edwards D. Calculating total rework costs in Australian construction projects.")
notes(sl, "Ставка инженера растёт, требования экспертизы ужесточаются, ошибки проекта дорого обходятся на стройке. "
          "Ответ — автоматизировать рутину в привычной среде, а не перестраивать весь процесс.")

# ---------------------------------------------------------------- 4. без 3D
sl = S[3]
set_title(sl, "Экономия без 3D и полного переобучения штата")
cards = [
    ("01", "Привычный 2D-процесс",
     "Работа в DWG и в привычной логике чертежа. Не нужно перестраивать процессы отдела и ждать внедрения BIM-стандарта."),
    ("02", "Короткие курсы",
     "Обучение в УЦ Нормасофт: отдельный курс по нужному модулю, без полного переобучения. 15 000–25 100 ₽ на сотрудника."),
    ("03", "Эффект с первого проекта",
     "Спецификации, журналы и расчёты формируются автоматически сразу. Лицензия и обучение окупаются за 1–3 проекта."),
]
x = 0.65
for num, head, body in cards:
    box(sl, x, 2.2, 3.75, 2.75, fill=GREY)
    num_badge(sl, x + 0.3, 2.45, num)
    text(sl, x + 0.3, 3.15, 3.2, 0.45, head, size=16, font=F_HEAD)
    text(sl, x + 0.3, 3.65, 3.2, 1.3, body, size=12)
    x += 4.0
text(sl, 0.65, 5.2, 11, 0.35, "Затраты на внедрение одного рабочего места (лицензия + обучение)", size=12, color=MUTED)
x = 0.65
for val, lab in [("99 900 ₽", "АР и КР (СПДС)"), ("122 000 ₽", "Электро"), ("184 600 ₽", "СКС + ОПС")]:
    box(sl, x, 5.6, 3.75, 0.9, fill=None, line=BLUE)
    text(sl, x + 0.3, 5.6, 1.9, 0.9, val, size=20, color=BLUE, font=F_HEAD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, x + 2.15, 5.6, 1.5, 0.9, lab, size=12, anchor=MSO_ANCHOR.MIDDLE)
    x += 4.0
notes(sl, "Ключевой аргумент для директора: экономия достигается без перехода на 3D. "
          "Суммы внедрения — из расчётов по модулям: СПДС 69 900 + 30 000; Электро 96 900 + 25 100; СКС+ОПС 151 200 + 33 400.")

# ---------------------------------------------------------------- 5. где экономия
sl = S[4]
set_title(sl, "Где возникает экономия: объект «Детский сад на 200 мест»")
rows = [("КР", 1935.1, 35, 677.3), ("АР", 1607.7, 35, 562.7), ("ТХ", 1488.6, 30, 446.6),
        ("ОВ", 694.7, 22, 152.8), ("ВК", 565.7, 22, 124.4), ("ЭО", 466.4, 25, 116.6),
        ("ПБ", 426.7, 22, 93.9), ("СС", 347.3, 20, 69.5), ("ПОС", 307.6, 30, 92.3),
        ("ПЗУ", 297.7, 30, 89.3), ("АВТ", 267.9, 20, 53.6)]
cd = CategoryChartData()
cd.categories = [f"{c}  {p}%" for c, _, p, _ in rows][::-1]
cd.add_series("Затраты после внедрения, тыс. ₽", [round(v - s, 1) for _, v, _, s in rows][::-1])
cd.add_series("Экономия, тыс. ₽", [s for *_, s in rows][::-1])
gf = sl.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.55), Inches(2.1), Inches(7.6), Inches(4.45), cd)
ch = gf.chart
style_chart(ch, 10)
ch.has_title = False
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
plot = ch.plots[0]
plot.gap_width = 45
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = LINE
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = BLUE
s1 = plot.series[1]
s1.data_labels.show_value = True
s1.data_labels.number_format = '#,##0'
s1.data_labels.number_format_is_linked = False
s1.data_labels.position = XL_LABEL_POSITION.INSIDE_END
s1.data_labels.font.size = Pt(9)
s1.data_labels.font.color.rgb = WHITE
va = ch.value_axis
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = RGBColor(0xE5, 0xE7, 0xEB)
va.format.line.fill.background()
va.tick_labels.font.size = Pt(9)
va.tick_labels.font.color.rgb = MUTED
ch.category_axis.format.line.color.rgb = LINE
ch.category_axis.tick_labels.font.size = Pt(10)
tiles = [("9 923,8", "тыс. ₽ — базовая цена проектирования без НДС"),
         ("2 479,0", "тыс. ₽ — общая экономия по проекту"),
         ("25%", "от стоимости проекта — консервативная оценка")]
y = 2.2
for i, (big, lab) in enumerate(tiles):
    box(sl, 8.55, y, 4.0, 1.12, fill=PALE if i else GREY)
    text(sl, 8.8, y + 0.1, 3.6, 0.55, big, size=26, color=BLUE if i else INK, font=F_HEAD)
    text(sl, 8.8, y + 0.64, 3.6, 0.45, lab, size=11)
    y += 1.27
text(sl, 8.55, 6.0, 4.0, 0.55, "Разделы СМ, КОН, ООС, ОДИ, ЭЭ, ПЗ, ТБЭ — экономия не учитывается (0%).",
     size=10, color=MUTED)
source(sl, "Расчёт по калькулятору стоимости проектных работ: 848/пр, 707/пр, Ипр 1,68 (II кв. 2026). Длина полосы — затраты раздела, синяя часть — экономия.")
notes(sl, "Основной эффект — в АР и КР (СПДС: 35%), затем ТХ, ПОС и ПЗУ (30%). Инженерные разделы — 20–25%. "
          "Итого 25% от цены проекта, или 2,48 млн ₽ на одном объекте.")


# ---------------------------------------------------------------- ROI-слайды
def roi_slide(sl, title, bullets, objects, foot, note):
    set_title(sl, title)
    text(sl, 0.65, 2.2, 3.35, 0.4, "Что автоматизируется", size=15, font=F_HEAD)
    text(sl, 0.65, 2.75, 3.35, 3.6, bullets, size=12, bullet=True, space=9)
    x = 4.25
    for name, tdo, tpo, fot, loss, cost, roi, eff in objects:
        box(sl, x, 2.2, 3.98, 4.28, fill=GREY)
        text(sl, x + 0.28, 2.35, 3.45, 0.45, name, size=14, font=F_HEAD, anchor=MSO_ANCHOR.MIDDLE)
        text(sl, x + 0.28, 2.85, 3.45, 0.3, "Трудозатраты по разделу, ч", size=10, color=MUTED)
        text(sl, x + 0.28, 3.12, 3.45, 0.45,
             [[(f"{tdo}", {"color": MUTED}), ("  →  ", {"color": MUTED}), (f"{tpo} ч", {"color": BLUE})]],
             size=18, font=F_HEAD)
        yy = 3.7
        for lab, val in [("Экономия ФОТ в год", fot), ("Экономия на правках", loss),
                         ("Затраты на внедрение", cost)]:
            text(sl, x + 0.28, yy, 2.1, 0.3, lab, size=11)
            text(sl, x + 2.2, yy, 1.53, 0.3, val, size=11, align=PP_ALIGN.RIGHT)
            yy += 0.33
        box(sl, x + 0.28, 4.85, 3.42, 1.4, fill=WHITE)
        text(sl, x + 0.45, 4.93, 1.6, 0.3, "ROI за год", size=10, color=MUTED)
        text(sl, x + 0.45, 5.18, 1.7, 0.6, roi, size=28, color=BLUE, font=F_HEAD)
        text(sl, x + 2.0, 4.93, 1.55, 0.3, "Эффект за год", size=10, color=MUTED, align=PP_ALIGN.RIGHT)
        text(sl, x + 1.9, 5.25, 1.65, 0.5, eff, size=14, font=F_HEAD, align=PP_ALIGN.RIGHT)
        text(sl, x + 0.45, 5.8, 3.1, 0.4, "10 проектов в год, ставка 1 171 ₽/ч", size=9, color=MUTED)
        x += 4.2
    source(sl, foot)
    notes(sl, note)


roi_slide(S[6], "АР и КР: окупаемость компонента СПДС",
          ["Оформление по ГОСТ и СПДС параметрическими объектами",
           "Спецификации и ведомости — из данных чертежа",
           "Изменения автоматически попадают в таблицы и марки",
           "Готовые перечни и объёмы для сметного отдела"],
          [("Детский сад на 200 мест", 368, 288, "937 000 ₽", "140 500 ₽", "99 900 ₽", "978%", "977 600 ₽"),
           ("Котельная 25 Гкал/ч", 352, 248, "1 217 800 ₽", "234 000 ₽", "99 900 ₽", "1 353%", "1 351 900 ₽")],
          "Платформа nanoCAD с компонентом СПДС — 69 900 ₽; обучение — 30 000 ₽ (платформа + компонент). Окупаемость — 1–2 проекта.",
          "Экономия на оформлении документации, спецификациях и ведомостях АР и КР. Детсад: 46 → 36 дней, котельная: 44 → 31 день. "
          "Исправления и согласования: затраты снижаются до 50%.")

roi_slide(S[8], "Электро: окупаемость nanoCAD BIM Электро",
          ["Электротехнические и светотехнические расчёты",
           "Трассировка сетей и кабельные журналы",
           "Щиты, выбор аппаратов, однолинейные схемы",
           "Спецификации оборудования по ГОСТ"],
          [("Детский сад на 200 мест", 160, 112, "562 000 ₽", "117 000 ₽", "122 000 ₽", "457%", "557 000 ₽"),
           ("Котельная 25 Гкал/ч", 320, 240, "936 800 ₽", "234 000 ₽", "122 000 ₽", "860%", "1 048 800 ₽")],
          "nanoCAD BIM Электро — 71 500 ₽ + Платформа nanoCAD — 25 400 ₽; обучение — 25 100 ₽. Окупаемость — 2–3 проекта.",
          "Раздел ЭО. Детсад: 20 → 14 дней, котельная: 40 → 30 дней.")

roi_slide(S[9], "СКС и ОПС: окупаемость nanoCAD ОПС и СКС",
          ["Расстановка оборудования и трассировка кабелей",
           "Расчёты: АКБ, падение напряжения, заполнение лотков, оповещение",
           "Кабельные журналы и спецификации — автоматически",
           "Структурные схемы и планы по ГОСТ"],
          [("Детский сад на 200 мест", 344, 272, "843 100 ₽", "280 100 ₽", "184 600 ₽", "508%", "938 600 ₽"),
           ("Котельная 25 Гкал/ч", 400, 304, "1 124 400 ₽", "281 000 ₽", "184 600 ₽", "661%", "1 220 800 ₽")],
          "nanoCAD ОПС и СКС — по 62 900 ₽ + Платформа nanoCAD — 25 400 ₽; обучение — 2 курса по 16 700 ₽. Окупаемость — 2–3 проекта.",
          "Разделы ПБ, СС, АВТ. Детсад: 43 → 33–35 дней, котельная: 50 → 38 дней.")

# ---------------------------------------------------------------- 11. сводка
sl = S[10]
set_title(sl, "Сводка: эффект за год от одного рабочего места")
cd = CategoryChartData()
cd.categories = ["АР и КР (СПДС)", "Электро", "СКС + ОПС"]
cd.add_series("Детский сад на 200 мест", (977.6, 557.0, 938.6))
cd.add_series("Котельная 25 Гкал/ч", (1351.9, 1048.8, 1220.8))
gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.55), Inches(2.1), Inches(7.1), Inches(4.4), cd)
ch = gf.chart
style_chart(ch, 11)
ch.has_title = True
ch.chart_title.text_frame.text = "Эффект за год, тыс. ₽"
ch.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
ch.chart_title.text_frame.paragraphs[0].runs[0].font.bold = False
ch.has_legend = True
ch.legend.position = XL_LEGEND_POSITION.BOTTOM
ch.legend.include_in_layout = False
plot = ch.plots[0]
plot.gap_width = 70
plot.overlap = -10
for s, c in zip(plot.series, (LINE, BLUE)):
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = c
    s.data_labels.show_value = True
    s.data_labels.number_format = '#,##0'
    s.data_labels.number_format_is_linked = False
    s.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    s.data_labels.font.size = Pt(10)
va = ch.value_axis
va.visible = False
va.has_major_gridlines = False
ch.category_axis.format.line.color.rgb = LINE
tbl = sl.shapes.add_table(4, 4, Inches(7.95), Inches(2.2), Inches(4.6), Inches(2.3)).table
hdr = ["Модуль", "ROI детсад", "ROI котельная", "Окупаемость"]
data = [["АР и КР", "978%", "1 353%", "1–2 проекта"], ["Электро", "457%", "860%", "2–3 проекта"],
        ["СКС + ОПС", "508%", "661%", "2–3 проекта"]]
widths = [1.15, 1.05, 1.2, 1.2]
for i, w in enumerate(widths):
    tbl.columns[i].width = Inches(w)
for r, row in enumerate([hdr] + data):
    for c, v in enumerate(row):
        cell = tbl.cell(r, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE if r == 0 else (GREY if r % 2 else WHITE)
        cell.margin_left = cell.margin_right = Inches(0.06)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
        run = p.add_run()
        run.text = v
        style_run(run, 10 if r == 0 else 11, WHITE if r == 0 else (BLUE if c in (1, 2) else INK))
box(sl, 7.95, 4.85, 4.6, 1.55, fill=PALE)
text(sl, 8.2, 4.98, 4.15, 1.35,
     [[("Все модули окупаются за 1–3 проекта. ", {"font": F_HEAD, "size": 14})],
      "Расчёт — по консервативному сценарию: 20–30% экономии времени по разделу."],
     size=12, space=6)
source(sl, "Эффект = экономия ФОТ + экономия на исправлениях − затраты на внедрение; 10 однотипных проектов в год.")
notes(sl, "Сводка по шести расчётам. Цифры — из базовых расчётов по модулям, без изменений.")

# ---------------------------------------------------------------- 12. другие разделы
sl = S[11]
set_title(sl, "Справочно: другие разделы проектирования")
text(sl, 0.65, 2.1, 11.9, 0.5,
     [[("Сокращение времени и трудозатрат — до 25–30% ", {"font": F_HEAD, "size": 15, "color": BLUE}),
       ("(консервативная оценка, аналогичная расчётам выше)", {"size": 13, "color": MUTED})]])
cards = [
    ("nanoCAD BIM ОВ", "ОВ", "Теплотехнические и аэродинамические расчёты, подбор оборудования, схемы и спецификации"),
    ("nanoCAD BIM ВК", "ВК", "Гидравлические расчёты, аксонометрические схемы, спецификации систем водоснабжения и канализации"),
    ("GeoniCS", "ПЗУ, генплан, сети", "Цифровая модель рельефа, генплан и вертикальная планировка, профили наружных сетей и трасс"),
    ("nanoCAD Стройплощадка", "ПОС, ППР", "Стройгенпланы, зоны действия кранов и опасные зоны, ведомости для ПОС и ППР"),
]
x = 0.65
for name, sec, body in cards:
    box(sl, x, 2.8, 2.83, 3.55, fill=GREY)
    text(sl, x + 0.25, 3.0, 2.4, 0.75, name, size=15, font=F_HEAD)
    pill(sl, x + 0.25, 3.8, 2.33, sec, active=False, h=0.36, size=11)
    text(sl, x + 0.25, 4.35, 2.35, 1.3, body, size=11)
    text(sl, x + 0.25, 5.68, 2.35, 0.5, "до 25–30%", size=20, color=BLUE, font=F_HEAD)
    x += 3.03
source(sl, "В расчёте по детскому саду: ОВ и ВК — 22%, ПОС и ПЗУ — 30% экономии (слайд «Где возникает экономия»).")
notes(sl, "Справочная информация — отдельные расчёты окупаемости по этим модулям не приводим.")

# ---------------------------------------------------------------- 14. формула и допущения
sl = S[13]
set_title(sl, "Методика расчёта окупаемости")
box(sl, 0.65, 2.2, 5.9, 4.25, fill=GREY)
text(sl, 0.95, 2.4, 5.3, 0.4, "Формулы", size=15, font=F_HEAD)
formulas = [
    ("ROI = (Экономия ФОТ + L − Cвнедр) / Cвнедр × 100%", True),
    ("Экономия ФОТ = (Tдо − Tпосле) × N × Rчас", False),
    ("L — экономия на исправлениях = Nсред × Mошиб × Oкол-во", False),
    ("Cвнедр = Cлиц + Cобуч", False),
]
y = 2.95
for f, main in formulas:
    box(sl, 0.95, y, 5.3, 0.62, fill=BLUE if main else WHITE)
    text(sl, 1.12, y, 5.0, 0.62, f, size=13 if main else 12, color=WHITE if main else INK,
         anchor=MSO_ANCHOR.MIDDLE, font=F_HEAD if main else F_BODY)
    y += 0.78
text(sl, 0.95, 6.02, 5.3, 0.4, "T — часы на раздел; N — проектов в год; Rчас — стоимость часа инженера",
     size=10, color=MUTED)
text(sl, 6.95, 2.2, 5.6, 0.4, "Допущения", size=15, font=F_HEAD)
assumptions = [
    ("Ставка", "146 000 ₽/мес → 9 370 ₽/день → 1 171 ₽/ч (ЕМИСС, ОКВЭД 71; соцотчисления по 707/пр)"),
    ("Цена проекта", "Методики 848/пр и 707/пр, Ипр 1,68 (II кв. 2026), доля раздела — по 707/пр"),
    ("Себестоимость", "Рентабельность 10%; ФОТ + отчисления — 53,98% себестоимости"),
    ("Загрузка", "10 однотипных проектов в год"),
    ("Резерв", "Правки и согласования — около 20% трудового бюджета раздела; затраты на исправления снижаются до 50%"),
    ("Сценарий", "Консервативный: 20–30% экономии времени по разделу"),
]
y = 2.72
for k, v in assumptions:
    text(sl, 6.95, y, 1.45, 0.55, k, size=11, color=BLUE)
    text(sl, 8.4, y, 4.15, 0.6, v, size=11)
    y += 0.62
notes(sl, "Все расчёты по модулям используют одну методику и одни допущения. Источник стоимости — калькулятор ПИР.")


# ---------------------------------------------------------------- 15–16. базовые работы
def works_table(sl, x, y, w, title, rows, total=None):
    text(sl, x, y, w, 0.4, title, size=14, font=F_HEAD)
    n = len(rows) + 1 + (1 if total else 0)
    tbl = sl.shapes.add_table(n, 4, Inches(x), Inches(y + 0.45), Inches(w), Inches(0.36 * n)).table
    cw = [w - 3.2, 1.05, 1.05, 1.1]
    for i, c in enumerate(cw):
        tbl.columns[i].width = Inches(c)
    allrows = [("Этап работ", "Без модуля", "С модулем", "Экономия")] + rows + ([total] if total else [])
    for r, row in enumerate(allrows):
        is_total = total and r == n - 1
        for c, v in enumerate(row):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BLUE if r == 0 else (PALE if is_total else (GREY if r % 2 else WHITE))
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.03)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER
            run = p.add_run()
            run.text = v
            color = WHITE if r == 0 else (BLUE if c == 3 and v not in ("—", "0%") else INK)
            style_run(run, 10 if r == 0 else 10.5, color, F_HEAD if (is_total or (c == 3 and r)) else F_BODY)


sl = S[14]
set_title(sl, "Базовые работы и экономия по этапам: АР и КР, сметы")
works_table(sl, 0.65, 2.15, 6.3, "АР и КР — компонент СПДС, дни", [
    ("Анализ проектной документации", "2–3", "2–3", "0%"),
    ("Составление ведомостей материалов", "3–5", "1–3", "−50%"),
    ("Спецификация оборудования*", "5–10", "3–7", "−33%"),
    ("Сведение, оформление по ГОСТ, согласование", "3–5", "2–3", "−38%"),
], ("Итого", "13–23", "8–16", "−33%"))
works_table(sl, 7.25, 2.15, 5.3, "Сметный отдел, дни", [
    ("Перечни материалов, объёмы", "2–4", "0", "−100%"),
    ("Конъюнктурный анализ цен", "5–10", "5–10", "0%"),
    ("Корректировки смет", "1–3", "1", "−50%"),
])
text(sl, 7.25, 4.55, 5.3, 1.5,
     "Спецификации и ведомости из чертежа передаются сметчикам готовыми: ручной подсчёт объёмов исключается.",
     size=12)
source(sl, "* При использовании готовых баз оборудования. Экономия — по середине диапазона. Детский сад, типовой объём работ.", y=6.45)
notes(sl, "Таблицы работ — базовые, из материалов по модулям. Процент экономии по этапу — отношение середин диапазонов.")

sl = S[15]
set_title(sl, "Базовые работы и экономия по этапам: Электро, СКС и ОПС")
works_table(sl, 0.65, 2.15, 5.85, "Электро — nanoCAD BIM Электро, дни", [
    ("Сбор исходных данных, ТУ", "3–5", "3–5", "0%"),
    ("Трассировка сетей ЭО, ЭМ", "5–7", "3–5", "−33%"),
    ("ВРУ и щиты, селективность", "4–5", "2–3", "−44%"),
    ("Оформление РД и ПД", "5–7", "3–5", "−33%"),
    ("Электротехнический расчёт", "4–6", "2–3", "−50%"),
], ("Итого (без расчётов)", "17–24", "11–18", "−29%"))
works_table(sl, 6.7, 2.15, 5.85, "СКС и ОПС — nanoCAD ОПС и СКС, дни", [
    ("Графическая часть, кабельные журналы", "5–7", "3–5", "−33%"),
    ("Специализированные расчёты", "5–7", "2–3", "−58%"),
    ("Формирование спецификации", "4–5", "2–3", "−44%"),
    ("Согласование и сдача проекта", "5–7", "3–5", "−33%"),
], ("Итого", "19–26", "10–16", "−42%"))
box(sl, 0.65, 5.35, 11.9, 0.95, fill=PALE)
text(sl, 0.95, 5.35, 11.4, 0.95,
     [[("По этапам экономия достигает 29–42%. ", {"font": F_HEAD, "size": 14}),
       ("В расчётах окупаемости принято 21–30% по разделу — консервативный сценарий.", {"size": 13})]],
     anchor=MSO_ANCHOR.MIDDLE)
notes(sl, "Принятые в ROI значения: СКС/ОПС 43→34 и 50→38 дней (21–24%), Электро 20→14 и 40→30 (25–30%), "
          "СПДС 46→36 и 44→31 (22–30%).")

# ---------------------------------------------------------------- 17. виды работ с макс. экономией
sl = S[16]
set_title(sl, "Виды работ с наибольшей экономией")
tasks = [("Перечни и объёмы для смет", 100), ("Спецрасчёты СКС и ОПС", 58), ("Ведомости материалов (АР, КР)", 50),
         ("Электротехнический расчёт", 50), ("Корректировки смет", 50), ("Исправления по замечаниям", 50),
         ("ВРУ и щиты", 44), ("Спецификации СКС и ОПС", 44), ("Сведение и согласование (АР, КР)", 38),
         ("Спецификация оборудования (АР, КР)", 33), ("Трассировка и журналы", 33)]
cd = CategoryChartData()
cd.categories = [t for t, _ in tasks][::-1]
cd.add_series("Экономия времени, %", [v / 100 for _, v in tasks][::-1])
gf = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.55), Inches(2.05), Inches(7.4), Inches(4.5), cd)
ch = gf.chart
style_chart(ch, 10)
ch.has_title = False
ch.has_legend = False
plot = ch.plots[0]
plot.gap_width = 40
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = BLUE
dl = plot.series[0].data_labels
dl.show_value = True
dl.number_format = '0%'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
dl.font.size = Pt(10)
ch.value_axis.visible = False
ch.value_axis.has_major_gridlines = False
ch.value_axis.maximum_scale = 1.15
ch.value_axis.minimum_scale = 0
ch.category_axis.format.line.color.rgb = LINE
ch.category_axis.tick_labels.font.size = Pt(10)
box(sl, 8.3, 2.2, 4.25, 2.1, fill=GREY)
text(sl, 8.55, 2.35, 3.8, 1.9,
     [[("Максимальный эффект", {"font": F_HEAD, "size": 15})],
      "там, где результат формируется из данных чертежа автоматически: спецификации, ведомости, журналы, расчёты, внесение правок."],
     size=12, space=6)
box(sl, 8.3, 4.5, 4.25, 1.9, fill=PALE)
text(sl, 8.55, 4.62, 3.8, 1.7,
     [[("77–91%", {"font": F_HEAD, "size": 26, "color": BLUE})],
      "экономия на тех же типах задач в исследовании Autodesk для специализированных наборов AutoCAD"],
     size=12, space=4)
source(sl, "Середина диапазонов из таблиц базовых работ; «Исправления по замечаниям» — допущение методики (до 50%).")
notes(sl, "Самые быстрые выигрыши: всё, что раньше считалось вручную по чертежу.")

# ---------------------------------------------------------------- 18. кейсы nanoCAD
sl = S[17]
set_title(sl, "Подтверждение кейсами: пользователи nanoCAD")
cases = [
    ("−83%", "«Уфанет», телеком, СКС",
     ["2D-проект одного объекта: 4 ч → 40 мин (501 объект)", "−30% срока проектирования за счёт проверки коллизий",
      "−70% возвратов рабочей документации", "> 1,3 млн ₽ экономии материалов на пилоте"],
     "CNews, ComNews, 23.07.2026; CADmaster № 106"),
    ("+40%", "«АБТ», Санкт-Петербург, Электро",
     ["Эффективность проектирования выросла на 40%", "Подбор оборудования и кабелей расчётом по модели сети",
      "Автоматическая спецификация", "Меньше замечаний экспертизы и времени на доработку"],
     "nanocad.ru; «САПР и графика»; CADmaster № 103"),
    ("до −30%", "Параметрические 2D-элементы, КР",
     ["Сокращение сроков разработки документации до 30%", "Библиотека типовых параметрических элементов",
      "Настройка под стандарт предприятия", "Оперативный контроль качества проекта"],
     "Данные вендора, nanoCAD Конструкции PS"),
]
x = 0.65
for big, head, items, src in cases:
    box(sl, x, 2.15, 3.8, 4.3, fill=GREY)
    text(sl, x + 0.28, 2.3, 3.3, 0.7, big, size=32, color=BLUE, font=F_HEAD)
    text(sl, x + 0.28, 3.05, 3.3, 0.55, head, size=14, font=F_HEAD)
    text(sl, x + 0.28, 3.65, 3.3, 2.2, items, size=11, bullet=True, space=6)
    text(sl, x + 0.28, 5.95, 3.3, 0.4, src, size=9, color=MUTED)
    x += 4.05
notes(sl, "Кейсы подтверждают, что 25–30% экономии в наших расчётах — нижняя граница результатов реальных пользователей.")

# ---------------------------------------------------------------- 19. кейсы Autodesk
sl = S[18]
set_title(sl, "Кейсы: специализированные наборы AutoCAD")
cd = CategoryChartData()
cd.categories = ["Mechanical", "Architecture", "MEP", "Electrical"]
cd.add_series("Рост производительности", (0.55, 0.61, 0.85, 0.95))
gf = sl.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.55), Inches(2.45), Inches(6.2), Inches(3.3), cd)
ch = gf.chart
style_chart(ch, 11)
ch.has_title = False
ch.has_legend = False
plot = ch.plots[0]
plot.gap_width = 45
plot.series[0].format.fill.solid()
plot.series[0].format.fill.fore_color.rgb = BLUE
dl = plot.series[0].data_labels
dl.show_value = True
dl.number_format = '0%'
dl.number_format_is_linked = False
dl.position = XL_LABEL_POSITION.OUTSIDE_END
ch.value_axis.visible = False
ch.value_axis.has_major_gridlines = False
ch.value_axis.maximum_scale = 1.1
ch.value_axis.minimum_scale = 0
ch.category_axis.format.line.color.rgb = LINE
text(sl, 0.65, 2.1, 6, 0.35, "Общий рост производительности по сравнению с базовым AutoCAD", size=12, color=MUTED)
text(sl, 7.15, 2.1, 5.4, 0.4, "Экономия времени на отдельных задачах", size=15, font=F_HEAD)
items = [("−91%", "планы трубопроводов (MEP)"), ("−86%", "силовые и электрические планы (MEP)"),
         ("−84%", "новые электрические схемы (Electrical)"), ("−82%", "нумерация проводов (Electrical)"),
         ("−81%", "планы ОВ и воздуховодов (MEP)"), ("−77%", "правки существующих схем (Electrical)")]
y = 2.65
for big, lab in items:
    text(sl, 7.15, y, 1.1, 0.42, big, size=17, color=BLUE, font=F_HEAD, anchor=MSO_ANCHOR.MIDDLE)
    text(sl, 8.3, y, 4.25, 0.42, lab, size=12, anchor=MSO_ANCHOR.MIDDLE)
    y += 0.5
box(sl, 0.65, 5.9, 11.9, 0.55, fill=PALE)
text(sl, 0.9, 5.9, 11.4, 0.55,
     "Методика: независимый консультант по заказу Autodesk; один опытный пользователь, одинаковые задачи в базовом AutoCAD и в специализированном наборе.",
     size=11, anchor=MSO_ANCHOR.MIDDLE)
source(sl, "Autodesk: AutoCAD Toolsets Productivity Studies (Electrical, MEP, Architecture, Mechanical).", y=6.55)
notes(sl, "Мировой ориентир для специализированных инструментов поверх 2D-CAD. Это экономия на задачах, а не на всём разделе, "
          "поэтому она выше наших 25–30%.")

# ---------------------------------------------------------------- 20. консервативный сценарий
sl = S[19]
set_title(sl, "Наши расчёты выполнены по консервативному сценарию")
ranges = [("Принято в расчётах ROI", 21, 30), ("Сумма этапов работ (таблицы)", 29, 42),
          ("Кейсы пользователей nanoCAD", 30, 83), ("Бенчмарки Autodesk", 55, 95)]
cd = CategoryChartData()
cd.categories = [r[0] for r in ranges][::-1]
cd.add_series("start", [r[1] / 100 for r in ranges][::-1])
cd.add_series("Диапазон экономии", [(r[2] - r[1]) / 100 for r in ranges][::-1])
gf = sl.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(0.55), Inches(2.1), Inches(8.0), Inches(3.3), cd)
ch = gf.chart
style_chart(ch, 11)
ch.has_title = False
ch.has_legend = False
plot = ch.plots[0]
plot.gap_width = 55
plot.series[0].format.fill.background()
plot.series[1].format.fill.solid()
plot.series[1].format.fill.fore_color.rgb = BLUE
pt = plot.series[1].points[3]
pt.format.fill.solid()
pt.format.fill.fore_color.rgb = INK
va = ch.value_axis
va.minimum_scale = 0
va.maximum_scale = 1
va.major_unit = 0.25
va.tick_labels.number_format = '0%'
va.tick_labels.number_format_is_linked = False
va.tick_labels.font.size = Pt(10)
va.tick_labels.font.color.rgb = MUTED
va.has_major_gridlines = True
va.major_gridlines.format.line.color.rgb = RGBColor(0xE5, 0xE7, 0xEB)
va.format.line.fill.background()
ch.category_axis.format.line.color.rgb = LINE
labels = ["21–30%", "29–42%", "30–83%", "55–95%"]
y = 2.45
for lab in labels:
    text(sl, 8.7, y, 1.2, 0.5, lab, size=15, font=F_HEAD, color=BLUE, anchor=MSO_ANCHOR.MIDDLE)
    y += 0.7
box(sl, 10.05, 2.2, 2.5, 3.1, fill=GREY)
text(sl, 10.25, 2.35, 2.15, 2.9,
     [[("Исправления", {"font": F_HEAD, "size": 13})],
      "В расчётах — снижение затрат до 50%.", "У «Уфанет» — −70% возвратов документации."],
     size=11, space=6)
box(sl, 0.65, 5.6, 11.9, 0.8, fill=PALE)
text(sl, 0.9, 5.6, 11.4, 0.8,
     [[("Вывод: ", {"font": F_HEAD, "size": 14}),
       ("25–30% экономии по разделу — нижняя граница наблюдаемых результатов. Фактический эффект, как правило, выше.",
        {"size": 13})]], anchor=MSO_ANCHOR.MIDDLE)
notes(sl, "Сравниваем, что заложено в ROI, с таблицами этапов, кейсами nanoCAD и бенчмарками Autodesk. "
          "Бенчмарки задачные, кейсы — по отдельным процессам, поэтому выше.")

# ---------------------------------------------------------------- 21. ограничения
sl = S[20]
set_title(sl, "Ограничения методики")
lims = [
    ("01", "Параллельная работа над несколькими проектами",
     "Методика считает экономию на один проект. Если специалист одновременно ведёт несколько объектов, "
     "распределение его времени и фактический эффект по каждому проекту учесть и оценить невозможно."),
    ("02", "Просадка эффективности при переходе на новое ПО",
     "Временное снижение производительности в период освоения не учитывается: оно слишком зависит от методики "
     "внедрения. С Нормасофт переход быстрее и проще — короткие курсы по модулю, готовые шаблоны и базы, "
     "сопровождение консультантов."),
]
y = 2.2
for num, head, body in lims:
    box(sl, 0.65, y, 11.9, 1.85, fill=GREY)
    num_badge(sl, 0.95, y + 0.3, num, size=0.6)
    text(sl, 1.85, y + 0.25, 10.4, 0.45, head, size=16, font=F_HEAD)
    text(sl, 1.85, y + 0.78, 10.4, 1.0, body, size=12)
    y += 2.1
source(sl, "Ставки — средние по ОКВЭД 71 без региональной поправки; цены лицензий — годовые, на дату расчёта.", y=6.45)
notes(sl, "Честно показываем, что расчёт не учитывает. Второй пункт — мост к блоку о Нормасофт.")

# ---------------------------------------------------------------- 25. компетенции — новый заголовок
sl = S[24]
for s in sl.shapes:
    if s.has_text_frame and s.text_frame.text.startswith("О важности компетенций"):
        set_paras(s, ["Консультанты по всем разделам проектирования"])

# ---------------------------------------------------------------- 26. разработки и внедрение
sl = S[25]
set_title(sl, "Собственные разработки и сопровождение внедрения")
steps = [("01", "Аудит", "Разбираем процессы отдела и выбираем разделы с наибольшим эффектом"),
         ("02", "Пилот", "Запуск модуля на реальном проекте заказчика, замер экономии"),
         ("03", "Обучение", "Короткие курсы по модулю в УЦ Нормасофт — очно, онлайн или у вас"),
         ("04", "Тиражирование", "Распространение на отдел, настройка шаблонов, поддержка")]
x = 0.65
for i, (num, head, body) in enumerate(steps):
    box(sl, x, 2.2, 2.75, 2.35, fill=GREY)
    num_badge(sl, x + 0.25, 2.4, num)
    text(sl, x + 0.25, 3.02, 2.3, 0.4, head, size=15, font=F_HEAD)
    text(sl, x + 0.25, 3.45, 2.3, 1.0, body, size=11)
    if i < 3:
        text(sl, x + 2.77, 2.95, 0.3, 0.5, "→", size=20, color=BLUE, align=PP_ALIGN.CENTER)
    x += 3.05
devs = [("Плагины и утилиты", "Типовые и индивидуальные плагины под задачи отдела и стандарты предприятия"),
        ("Базы и шаблоны", "Наполнение баз оборудования, настройка шаблонов и спецификаций под ГОСТ и СТО"),
        ("Развитие", "Цифровые двойники производства, BIM VR — следующий шаг, когда он нужен")]
x = 0.65
for head, body in devs:
    box(sl, x, 4.85, 3.8, 1.6, fill=None, line=BLUE)
    text(sl, x + 0.25, 5.0, 3.3, 0.4, head, size=14, font=F_HEAD, color=BLUE)
    text(sl, x + 0.25, 5.42, 3.3, 0.95, body, size=11)
    x += 4.05
notes(sl, "Мы не просто продаём лицензии: помогаем получить экономию из расчёта на практике.")

def fix_img_layers(slide):
    """В шаблоне эффект яркости картинки (a14:imgLayer) ссылается на связь со слайдом, а не на hdphoto."""
    for layer in list(slide._element.iter("{http://schemas.microsoft.com/office/drawing/2010/main}imgLayer")):
        rId = layer.get(qn("r:embed"))
        rel = slide.part.rels.get(rId) if rId else None
        if rel is None or not rel.reltype.endswith("/hdphoto"):
            ext = layer.getparent().getparent()          # a:ext
            ext.getparent().remove(ext)


# ---------------------------------------------------------------- навигация и чистка
for n, sl in enumerate(S, 1):
    if n > 1:
        rebuild_nav(n, sl)
for sl in S:
    fix_img_layers(sl)
    drop_dead_slide_rels(sl)

prs.save(OUT)
print("saved", OUT)
