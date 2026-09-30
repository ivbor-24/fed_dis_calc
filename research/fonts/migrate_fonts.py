"""Перевод pptx на фирменное семейство Manrope.

    python3 migrate_fonts.py in.pptx out.pptx      # миграция + отчёт «было → стало»
    python3 migrate_fonts.py in.pptx                # только отчёт по шрифтам

Тема: заголовки — Manrope (Bold задаётся атрибутом b="1"), основной текст — Manrope Medium.
Текст ссылается на шрифты темы (+mj-lt / +mn-lt), поэтому смена шрифта в будущем — правка одной темы.
Значковые шрифты (Wingdings, Symbol) не трогаем.
"""
import collections
import re
import sys
import zipfile

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
HEAD, BODY = "Manrope", "Manrope Medium"          # семейства в теме

# typeface -> (для a:latin, для a:ea/a:cs, сделать полужирным)
MAP = {
    "VK Sans Display DemiBold": ("+mj-lt", HEAD, True),
    "Aptos Display": ("+mj-lt", HEAD, True),
    "Roboto Light": ("Manrope", "Manrope", False),
    "Roboto": ("Manrope", "Manrope", False),
    "Manrope Medium": ("+mn-lt", BODY, False),
    "Aptos": ("+mn-lt", BODY, False),
    "Arial": ("+mn-lt", BODY, False),
    "Calibri": ("+mn-lt", BODY, False),
    "Times New Roman": ("+mn-lt", BODY, False),
}
BULLET_MAP = {"Arial": BODY, "Manrope Medium": BODY, "Roboto Light": "Manrope"}
KEEP = {"Wingdings", "Symbol"}
PARTS = re.compile(r"ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster|notesSlides/"
                   r"notesSlide|notesMasters/notesMaster|charts/chart|theme/theme)\d*\.xml$")


def fonts_in(xml):
    return re.findall(r'<a:(?:latin|ea|cs|buFont) typeface="([^"]+)"', xml)


def migrate_part(data, stats):
    root = etree.fromstring(data)
    if root.tag == f"{{{A}}}theme":
        for tag, face in (("majorFont", HEAD), ("minorFont", BODY)):
            latin = root.find(f".//{{{A}}}{tag}/{{{A}}}latin")
            latin.attrib.clear()
            latin.set("typeface", face)
    for el in root.iter(f"{{{A}}}latin", f"{{{A}}}ea", f"{{{A}}}cs", f"{{{A}}}buFont"):
        face = el.get("typeface", "")
        if not face or face in KEEP or face.startswith("+") or el.getparent().tag.endswith(("majorFont", "minorFont")):
            continue
        local = etree.QName(el).localname
        if local == "buFont":
            new, bold = BULLET_MAP.get(face), False
        elif face in MAP:
            latin_face, other_face, bold = MAP[face]
            new = latin_face if local == "latin" else other_face
        else:
            continue
        if new is None or new == face:
            continue
        for attr in ("panose", "pitchFamily", "charset"):
            el.attrib.pop(attr, None)
        el.set("typeface", new)
        stats[f"{face} → {new}{' + Bold' if bold and local == 'latin' else ''}"] += 1
        if bold and local == "latin":
            el.getparent().set("b", "1")     # a:rPr / a:defRPr / a:endParaRPr
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def report(path):
    c = collections.Counter()
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if PARTS.search(n):
                c.update(fonts_in(z.read(n).decode("utf8", "ignore")))
    return c


def main():
    src = sys.argv[1]
    before = report(src)
    print(f"== {src}\n  было:  " + ", ".join(f"{k} ×{v}" for k, v in before.most_common()))
    if len(sys.argv) < 3:
        return
    out, stats = sys.argv[2], collections.Counter()
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if PARTS.search(item.filename):
                data = migrate_part(data, stats)
            zout.writestr(item, data)
    after = report(out)
    print("  замены: " + ", ".join(f"{k} ×{v}" for k, v in stats.most_common()))
    print("  стало: " + ", ".join(f"{k} ×{v}" for k, v in after.most_common()))


if __name__ == "__main__":
    main()
