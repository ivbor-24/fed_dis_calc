"""Короткая версия для директора: 12 слайдов из полной презентации.

python3 make_short.py ROI_moduli_nanoCAD.pptx ROI_moduli_nanoCAD_short.pptx
"""
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

SRC, OUT = sys.argv[1], sys.argv[2]
BLUE = RGBColor(0x03, 0x86, 0xFC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# номера слайдов полной версии, в порядке показа
KEEP = [1, 3, 4, 5, 7, 9, 10, 11, 18, 20, 26, 27]
NAV = [("ВЫЗОВЫ", 2), ("ЭКОНОМИЯ", 4), ("ОКУПАЕМОСТЬ", 5), ("КЕЙСЫ", 9),
       ("НОРМАСОФТ", 11), ("КОНТАКТЫ", 12)]
ACTIVE = {2: "ВЫЗОВЫ", 3: "ВЫЗОВЫ", 4: "ЭКОНОМИЯ", 5: "ОКУПАЕМОСТЬ", 6: "ОКУПАЕМОСТЬ",
          7: "ОКУПАЕМОСТЬ", 8: "ОКУПАЕМОСТЬ", 9: "КЕЙСЫ", 10: "КЕЙСЫ", 11: "НОРМАСОФТ",
          12: "КОНТАКТЫ"}

def fix_img_layers(slide):
    """В шаблоне эффект яркости картинки (a14:imgLayer) ссылается на связь со слайдом, а не на hdphoto."""
    for layer in list(slide._element.iter("{http://schemas.microsoft.com/office/drawing/2010/main}imgLayer")):
        rId = layer.get(qn("r:embed"))
        rel = slide.part.rels.get(rId) if rId else None
        if rel is None or not rel.reltype.endswith("/hdphoto"):
            ext = layer.getparent().getparent()          # a:ext
            ext.getparent().remove(ext)


prs = Presentation(SRC)
lst = prs.slides._sldIdLst
ids = list(lst)
for el in ids:
    lst.remove(el)
for n in KEEP:
    lst.append(ids[n - 1])
for el in ids:
    if el not in lst:
        prs.part.drop_rel(el.get(qn("r:id")))
S = list(prs.slides)

for n, sl in enumerate(S, 1):
    if n == 1:
        continue
    for s in list(sl.shapes):
        if s.name == "nav":
            s._element.getparent().remove(s._element)
    x = 1.02
    for label, target in NAV:
        w = 0.4 + 0.11 * len(label)
        s = sl.shapes.add_shape(MSO_SHAPE.SNIP_1_RECTANGLE, Inches(x), Inches(0.47),
                                Inches(w), Inches(0.37))
        s.name = "nav"
        s.shadow.inherit = False
        s.line.fill.background()
        active = ACTIVE[n] == label
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
        r.font.size = Pt(12)
        r.font.name = "Roboto Light"
        r.font.color.rgb = BLUE if active else WHITE
        s.click_action.target_slide = S[target - 1]
        x += w + 0.15
    fix_img_layers(sl)
    xml = etree.tostring(sl._element).decode()
    for rId, rel in list(sl.part.rels.items()):
        if rel.reltype.endswith("/slide") and f'"{rId}"' not in xml:
            sl.part.drop_rel(rId)

# ограничения методики — в заметки к слайду о консервативном сценарии
nt = S[9].notes_slide.notes_text_frame
nt.text += ("\n\nОграничения методики: не учитывается параллельная работа специалиста над несколькими "
            "проектами и временная просадка эффективности при переходе — она зависит от методики "
            "внедрения; с Нормасофт переход быстрее за счёт коротких курсов и сопровождения.")
prs.save(OUT)
print("saved", OUT, len(S))
