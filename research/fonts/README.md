# Фирменный шрифт Manrope

- `Manrope-*.ttf` — статические веса 200–800, собраны из переменного шрифта Google Fonts (SIL OFL 1.1, `OFL.txt`).
  Ставить на рабочие машины именно их: переменный файл PowerPoint поддерживает ненадёжно.
- Роли: заголовки — Manrope Bold (шрифт заголовков темы + полужирный), текст — Manrope Medium (основной шрифт темы),
  меню и подписи — Manrope Regular, крупные цифры — Manrope ExtraBold.
- `migrate_fonts.py` — перевод любой pptx на Manrope + отчёт «было → стало»:
  `python3 migrate_fonts.py in.pptx out.pptx` (без `out.pptx` — только отчёт).
- `Normasoft_template_Manrope.pptx` — шаблон Нормасофт после миграции.
- `EmbedFonts.bas` — макрос PowerPoint для пакетного встраивания шрифтов перед отправкой pptx наружу.

Ограничения Manrope: нет курсива; нет казахских Ә Ғ Қ Ң Ұ и символов ✓ ↗.
