"""
kp_pdf.py — генератор PDF коммерческого предложения.
Единый стиль с каталогом. Компактная вёрстка без принудительных разрывов.
"""

import io
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (Paragraph, Spacer, Table, TableStyle,
                                 HRFlowable, PageBreak, KeepTogether,
                                 BaseDocTemplate, PageTemplate, Frame,
                                 NextPageTemplate, Image as RLImage)
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image as PILImage

W, H = A4

# ── Палитра ───────────────────────────────────────────────────────────────────
C_DARK    = colors.HexColor('#1A2B4A')
C_BLUE    = colors.HexColor('#2563EB')
C_LBLUE   = colors.HexColor('#EFF6FF')
C_GOLD    = colors.HexColor('#F59E0B')
C_AMBER   = colors.HexColor('#FFFBEB')
C_GRAY    = colors.HexColor('#64748B')
C_LGRAY   = colors.HexColor('#F8FAFC')
C_BORDER  = colors.HexColor('#E2E8F0')
C_TEXT    = colors.HexColor('#1E293B')
C_WHITE   = colors.white
C_RED     = colors.HexColor('#C0392B')
C_GREEN   = colors.HexColor('#16A34A')

# ── Шрифты ────────────────────────────────────────────────────────────────────
_FONTS_REGISTERED = False

def _register_fonts():
    global _FONTS_REGISTERED
    if _FONTS_REGISTERED:
        return

    import sys

    # Наборы (regular, bold, italic) — проверяются по порядку
    FONT_SETS = [
        # Windows — Arial всегда установлен
        ('C:/Windows/Fonts/arial.ttf',
         'C:/Windows/Fonts/arialbd.ttf',
         'C:/Windows/Fonts/ariali.ttf'),
        # macOS
        ('/Library/Fonts/Arial.ttf',
         '/Library/Fonts/Arial Bold.ttf',
         '/Library/Fonts/Arial Italic.ttf'),
        ('/System/Library/Fonts/Supplemental/Arial.ttf',
         '/System/Library/Fonts/Supplemental/Arial Bold.ttf',
         '/System/Library/Fonts/Supplemental/Arial Italic.ttf'),
        # Linux — DejaVu
        ('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
         '/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf'),
        ('/usr/share/fonts/dejavu/DejaVuSans.ttf',
         '/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf',
         '/usr/share/fonts/dejavu/DejaVuSans-Oblique.ttf'),
    ]

    # PyInstaller .exe — ресурсы в sys._MEIPASS
    if hasattr(sys, '_MEIPASS'):
        mp = Path(sys._MEIPASS)
        FONT_SETS.insert(0, (str(mp / 'arial.ttf'),
                             str(mp / 'arialbd.ttf'),
                             str(mp / 'ariali.ttf')))

    # Локальная папка fonts/ рядом со скриптом
    local = Path(__file__).resolve().parent / 'fonts'
    FONT_SETS.insert(0, (str(local / 'arial.ttf'),
                         str(local / 'arialbd.ttf'),
                         str(local / 'ariali.ttf')))

    for r, b, i in FONT_SETS:
        if Path(r).exists() and Path(b).exists() and Path(i).exists():
            pdfmetrics.registerFont(TTFont('DV',  r))
            pdfmetrics.registerFont(TTFont('DVB', b))
            pdfmetrics.registerFont(TTFont('DVI', i))
            _FONTS_REGISTERED = True
            return

    # Fallback: apt-get на Linux
    try:
        import subprocess
        subprocess.run(['apt-get', 'install', '-y', 'fonts-dejavu'],
                       capture_output=True)
        p = '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
        if Path(p).exists():
            pdfmetrics.registerFont(TTFont('DV',  p))
            pdfmetrics.registerFont(TTFont('DVB', p.replace('Sans.', 'Sans-Bold.')))
            pdfmetrics.registerFont(TTFont('DVI', p.replace('Sans.', 'Sans-Oblique.')))
            _FONTS_REGISTERED = True
            return
    except Exception:
        pass

    raise RuntimeError(
        'Шрифты не найдены!\n'
        'Windows: проверьте C:\\Windows\\Fonts\\arial.ttf\n'
        'Linux:   установите пакет fonts-dejavu'
    )


def _s(name, **kw):
    d = dict(fontName='DV', fontSize=9, textColor=C_TEXT, leading=13)
    d.update(kw)
    return ParagraphStyle(name, **d)


# ── Изображение из байт ───────────────────────────────────────────────────────
def _img(img_bytes, max_w, max_h):
    if not img_bytes:
        return None
    try:
        pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGB')
        pw, ph = pil.size
        scale = min(max_w / pw, max_h / ph)
        buf = io.BytesIO()
        pil.save(buf, 'JPEG', quality=88)
        buf.seek(0)
        return RLImage(buf, width=pw * scale, height=ph * scale)
    except Exception:
        return None


def _fmt(n):
    try:
        v = float(str(n).replace(',', '.').replace(' ', ''))
        return f'{v:,.0f}'.replace(',', '\u00a0')  # nbsp как разделитель тысяч
    except Exception:
        return str(n)


# ── Обложка ───────────────────────────────────────────────────────────────────
def _cover_fn(data):
    equipment = data['equipment']
    works     = data['works']
    logo_b    = data['logo_bytes']
    kp_num    = data['kp_number']
    manager   = data['manager']
    company   = data['company']

    total_eq  = sum(p['price'] * p['qty'] for p in equipment)
    total_wk  = sum(w['price'] * w['qty'] for w in works)
    total_all = total_eq + total_wk

    def draw(canvas, doc):
        canvas.saveState()
        # Фон
        canvas.setFillColor(C_DARK)
        canvas.rect(0, 0, W, H, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor('#1E3A6E'))
        canvas.circle(W - 50*mm, H - 20*mm, 130*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor('#152540'))
        canvas.circle(-20*mm, 70*mm, 100*mm, fill=1, stroke=0)

        # Полосы
        canvas.setFillColor(C_GOLD)
        canvas.rect(0, H - 8*mm, W, 8*mm, fill=1, stroke=0)
        canvas.setFillColor(C_BLUE)
        canvas.rect(0, 0, W, 6*mm, fill=1, stroke=0)

        # Логотип
        if logo_b:
            try:
                pil = PILImage.open(io.BytesIO(logo_b))
                pw, ph = pil.size
                scale = min(44*mm / pw, 18*mm / ph)
                canvas.drawImage(io.BytesIO(logo_b), 20*mm, H - 35*mm,
                                 width=pw*scale, height=ph*scale,
                                 preserveAspectRatio=True, mask='auto')
            except Exception:
                pass

        # Контакты компании (под логотипом)
        canvas.setFillColor(colors.HexColor('#94A3B8'))
        canvas.setFont('DV', 7.5)
        lines = [l.strip() for l in company.split('\n') if l.strip()]
        for i, line in enumerate(lines[:3]):
            canvas.drawString(20*mm, H - 39*mm - i*9, line)

        canvas.setStrokeColor(colors.HexColor('#2D4A7A'))
        canvas.line(20*mm, H - 44*mm - len(lines[:3])*9,
                    W - 20*mm, H - 44*mm - len(lines[:3])*9)

        # Заголовок
        canvas.setFillColor(C_WHITE)
        canvas.setFont('DVB', 36)
        canvas.drawString(20*mm, H/2 + 44*mm, 'КОММЕРЧЕСКОЕ')
        canvas.drawString(20*mm, H/2 + 26*mm, 'ПРЕДЛОЖЕНИЕ')

        canvas.setFillColor(C_GOLD)
        canvas.setFont('DVB', 13)
        canvas.drawString(20*mm, H/2 + 12*mm, kp_num)
        canvas.setFillColor(C_BLUE)
        canvas.rect(20*mm, H/2 + 6*mm, 55*mm, 2*mm, fill=1, stroke=0)

        # Менеджер
        canvas.setFillColor(colors.HexColor('#1E3A6E'))
        canvas.roundRect(20*mm, H/2 - 18*mm, W - 40*mm, 22*mm,
                         5, fill=1, stroke=0)
        canvas.setFillColor(C_WHITE)
        canvas.setFont('DV', 8.5)
        mgr_lines = [l.strip() for l in manager.replace('<br>', '\n').split('\n')
                     if l.strip()]
        for i, line in enumerate(mgr_lines[:3]):
            canvas.drawString(26*mm, H/2 - 3*mm - i*10, line)

        # Статистика
        if works:
            stats = [(_fmt(len(equipment)), 'позиций\nоборудования'),
                     (_fmt(len(works)),     'позиций\nработ'),
                     (_fmt(total_all),      'руб.  итого')]
        else:
            stats = [(_fmt(len(equipment)), 'позиций\nоборудования'),
                     (_fmt(total_all),      'руб.  итого')]

        box_w = 52*mm
        gap   = (W - 40*mm - box_w * len(stats)) / max(len(stats) - 1, 1)
        for i, (val, lbl) in enumerate(stats):
            x = 20*mm + i * (box_w + gap)
            canvas.setFillColor(colors.HexColor('#0F1F3D'))
            canvas.roundRect(x, H/2 - 48*mm, box_w, 26*mm, 5, fill=1, stroke=0)
            canvas.setFillColor(C_GOLD)
            canvas.setFont('DVB', 13)
            canvas.drawCentredString(x + box_w/2, H/2 - 30*mm, val)
            canvas.setFillColor(colors.HexColor('#94A3B8'))
            canvas.setFont('DV', 7)
            for j, l in enumerate(lbl.split('\n')):
                canvas.drawCentredString(x + box_w/2, H/2 - 40*mm - j*8, l)

        canvas.setFillColor(colors.HexColor('#64748B'))
        canvas.setFont('DV', 7.5)
        canvas.drawString(20*mm, 14*mm,
                          'Цены действительны в течение 3 банковских дней')
        canvas.restoreState()
    return draw


def _page_fn(kp_num):
    def draw(canvas, doc):
        canvas.saveState()
        canvas.setFillColor(C_WHITE)
        canvas.rect(0, 0, W, H, fill=1, stroke=0)
        # Шапка
        canvas.setFillColor(C_DARK)
        canvas.rect(0, H - 14*mm, W, 14*mm, fill=1, stroke=0)
        canvas.setFillColor(C_GOLD)
        canvas.rect(0, H - 15.5*mm, W, 1.5*mm, fill=1, stroke=0)
        canvas.setFillColor(C_WHITE)
        canvas.setFont('DVB', 9)
        canvas.drawString(20*mm, H - 9*mm, f'КП • {kp_num}')
        canvas.setFont('DV', 8)
        canvas.drawRightString(W - 20*mm, H - 9*mm, f'Стр. {doc.page}')
        # Подвал
        canvas.setFillColor(C_DARK)
        canvas.rect(0, 0, W, 9*mm, fill=1, stroke=0)
        canvas.setFillColor(colors.HexColor('#64748B'))
        canvas.setFont('DV', 6.5)
        canvas.drawCentredString(W/2, 3*mm,
                                 'CTV • Системы безопасности и домофонии')
        canvas.restoreState()
    return draw


# ── Карточка товара ───────────────────────────────────────────────────────────
def _equip_card(item, idx):
    """Компактная карточка товара — помещается несколько на страницу."""
    img_rl = _img(item.get('img_bytes'), 34*mm, 44*mm)

    # Разбираем описание
    raw_desc  = item.get('desc', '')
    desc_lines = [l.strip() for l in raw_desc.split('\n') if l.strip()]
    spec_lines = [l for l in desc_lines if '\t' in l]
    narr_lines = [l for l in desc_lines if '\t' not in l]

    # Заголовок карточки — название крупно, артикул мелко
    name_p = Paragraph(item.get('name', '—'),
                        _s('n', fontName='DVB', fontSize=11,
                            textColor=C_DARK, leading=15))
    art_val = item.get('article', '')
    art_p   = Paragraph(f'Арт. {art_val}' if art_val else '',
                        _s('a', fontSize=7.5, textColor=C_GRAY))

    narr_txt = ' '.join(narr_lines[:2])
    narr_p   = (Paragraph(narr_txt, _s('nr', fontName='DVI', fontSize=8,
                                         textColor=C_GRAY, leading=11))
                if narr_txt else Spacer(1, 1))

    right = Table([[name_p], [art_p], [Spacer(1, 2)], [narr_p]],
                  colWidths=[114*mm])
    right.setStyle(TableStyle([
        ('VALIGN',        (0,0),(-1,-1),'TOP'),
        ('LEFTPADDING',   (0,0),(-1,-1), 0),
        ('RIGHTPADDING',  (0,0),(-1,-1), 0),
        ('TOPPADDING',    (0,0),(-1,-1), 1),
        ('BOTTOMPADDING', (0,0),(-1,-1), 1),
    ]))

    img_cell = img_rl if img_rl else Spacer(34*mm, 8)
    main = Table([[img_cell, right]], colWidths=[38*mm, 120*mm])
    main.setStyle(TableStyle([
        ('VALIGN',         (0,0),(-1,-1),'MIDDLE'),
        ('BACKGROUND',     (0,0),(-1,-1), C_LBLUE),
        ('TOPPADDING',     (0,0),(-1,-1), 6),
        ('BOTTOMPADDING',  (0,0),(-1,-1), 6),
        ('LEFTPADDING',    (0,0),(0,-1),  7),
        ('LEFTPADDING',    (1,0),(1,-1),  6),
        ('RIGHTPADDING',   (0,0),(-1,-1), 7),
        ('ROUNDEDCORNERS', [5,5,5,5]),
    ]))

    # Спецификации — до 5 строк
    block = [main]
    if spec_lines:
        rows = []
        for line in spec_lines[:5]:
            parts = line.split('\t', 1)
            if len(parts) == 2:
                rows.append([
                    Paragraph(parts[0].strip(),
                               _s('sk', fontSize=7, textColor=C_GRAY, leading=10)),
                    Paragraph(parts[1].strip(),
                               _s('sv', fontSize=7, fontName='DVB',
                                   textColor=C_TEXT, leading=10)),
                ])
        if rows:
            st = Table(rows, colWidths=[68*mm, 90*mm])
            st.setStyle(TableStyle([
                ('VALIGN',        (0,0),(-1,-1),'TOP'),
                ('TOPPADDING',    (0,0),(-1,-1), 2),
                ('BOTTOMPADDING', (0,0),(-1,-1), 2),
                ('LEFTPADDING',   (0,0),(-1,-1), 4),
                ('ROWBACKGROUNDS',(0,0),(-1,-1),[C_LGRAY, C_WHITE]),
            ]))
            block += [Spacer(1, 1), st]

    # Ценовая строка
    price = item.get('price', 0)
    qty   = item.get('qty',   1)
    total = item.get('total') or price * qty
    if qty != 1:
        price_txt = (f'Цена: <b>{_fmt(price)} руб.</b>  ×  {qty:g} шт.'
                     f'  =  <b>{_fmt(total)} руб.</b>')
    else:
        price_txt = f'Цена: <b>{_fmt(price)} руб.</b>'

    bar = Table([[Paragraph(price_txt,
                             _s('pr', fontSize=9.5, textColor=C_WHITE,
                                 alignment=TA_CENTER, leading=13))]],
                colWidths=[160*mm])
    bar.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), C_BLUE),
        ('TOPPADDING',    (0,0),(-1,-1), 5),
        ('BOTTOMPADDING', (0,0),(-1,-1), 5),
        ('ROUNDEDCORNERS',[4,4,4,4]),
    ]))
    block += [Spacer(1, 2), bar,
              Spacer(1, 4),
              HRFlowable(width='100%', thickness=0.5, color=C_BORDER),
              Spacer(1, 4)]

    return KeepTogether(block)


# ── Таблица работ ─────────────────────────────────────────────────────────────
def _works_table(works):
    if not works:
        return []

    th = _s('wth', fontName='DVB', fontSize=8.5, textColor=C_WHITE, leading=12)
    td = _s('wtd', fontSize=8.5, leading=12)
    tr = _s('wtr', fontSize=8.5, textColor=C_BLUE, alignment=TA_RIGHT,
             fontName='DVB', leading=12)

    header = [Paragraph('№',                   th),
              Paragraph('Наименование работ',  th),
              Paragraph('Кол-во',              th),
              Paragraph('Цена, руб.',          th),
              Paragraph('Сумма, руб.',         th)]

    rows = [header]
    rc   = []
    for i, w in enumerate(works, 1):
        price = w.get('price', 0)
        qty   = w.get('qty',   1)
        total = w.get('total') or price * qty
        rows.append([
            Paragraph(str(i),        td),
            Paragraph(w.get('name',''), td),
            Paragraph(f'{qty:g}',    td),
            Paragraph(_fmt(price),   td),
            Paragraph(_fmt(total),   tr),
        ])
        rc.append(('BACKGROUND', (0, i), (-1, i),
                   C_AMBER if i % 2 == 1 else C_WHITE))

    # Используем None для колонки с наименованием, чтобы она занимала всю оставшуюся ширину
    t = Table(rows, colWidths=[7*mm, None, 16*mm, 24*mm, 25*mm],
              splitByRow=True)
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1, 0), colors.HexColor('#92400E')),
        ('TOPPADDING',    (0,0),(-1,-1), 4),
        ('BOTTOMPADDING', (0,0),(-1,-1), 4),
        ('LEFTPADDING',   (0,0),(-1,-1), 4),
        ('RIGHTPADDING',  (0,0),(-1,-1), 4),
        ('GRID',          (0,0),(-1,-1), 0.4, C_BORDER),
        ('VALIGN',        (0,0),(-1,-1),'MIDDLE'),
    ] + rc))
    return [t]


# ── Итоговый блок ─────────────────────────────────────────────────────────────
def _totals(equipment, works):
    total_eq  = sum((p.get('total') or p['price'] * p['qty']) for p in equipment)
    total_wk  = sum((w.get('total') or w['price'] * w['qty']) for w in works)
    total_all = total_eq + total_wk

    els = []

    def _row(label, value, bg, fg=C_WHITE, size=11):
        t = Table([[
            Paragraph(label, _s('tl', fontName='DVB', fontSize=size,
                                  textColor=fg, alignment=TA_RIGHT, leading=16)),
            Paragraph(f'{_fmt(value)}\u00a0руб.',
                       _s('tv', fontName='DVB', fontSize=size+1,
                           textColor=fg, alignment=TA_RIGHT, leading=16)),
        ]], colWidths=[118*mm, 42*mm])
        t.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), bg),
            ('TOPPADDING',    (0,0),(-1,-1), 6),
            ('BOTTOMPADDING', (0,0),(-1,-1), 6),
            ('LEFTPADDING',   (0,0),(-1,-1), 10),
            ('RIGHTPADDING',  (0,0),(-1,-1), 10),
            ('ROUNDEDCORNERS',[4,4,4,4]),
        ]))
        return [t, Spacer(1, 3)]

    if works:
        els += _row('Итого оборудование:', total_eq,
                    colors.HexColor('#1E3A6E'))
        els += _row('Итого работы:',       total_wk,
                    colors.HexColor('#7C2D12'))
        els += _row('ОБЩАЯ СУММА:',        total_all,
                    C_DARK, size=13)
    else:
        els += _row('ИТОГО:',              total_eq,
                    C_DARK, size=13)
    return els


# ── Сводная таблица оборудования ──────────────────────────────────────────────
def _equip_summary(equipment):
    th = _s('eth', fontName='DVB', fontSize=8, textColor=C_WHITE, leading=11)
    td = _s('etd', fontSize=8, leading=11)
    tr = _s('etr', fontSize=8, fontName='DVB', textColor=C_BLUE,
             alignment=TA_RIGHT, leading=11)

    header = [Paragraph('№',           th),
              Paragraph('Наименование',th),
              Paragraph('Арт.',        th),
              Paragraph('Цена, руб.',  th),
              Paragraph('Кол-во',      th),
              Paragraph('Сумма, руб.', th)]
    rows = [header]
    rc   = []
    for i, p in enumerate(equipment, 1):
        price = p.get('price', 0)
        qty   = p.get('qty',   1)
        total = p.get('total') or price * qty
        rows.append([
            Paragraph(str(i),              td),
            Paragraph(p.get('name',''),    td),
            Paragraph(p.get('article',''), td),
            Paragraph(_fmt(price),         td),
            Paragraph(f'{qty:g}',          td),
            Paragraph(_fmt(total),         tr),
        ])
        rc.append(('BACKGROUND', (0, i), (-1, i),
                   C_LBLUE if i % 2 == 1 else C_WHITE))

    # Используем None для колонки с наименованием, чтобы она занимала всю оставшуюся ширину
    t = Table(rows, colWidths=[7*mm, None, 16*mm, 24*mm, 14*mm, 25*mm],
              splitByRow=True)
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1, 0), C_DARK),
        ('TOPPADDING',    (0,0),(-1,-1), 3),
        ('BOTTOMPADDING', (0,0),(-1,-1), 3),
        ('LEFTPADDING',   (0,0),(-1,-1), 4),
        ('RIGHTPADDING',  (0,0),(-1,-1), 4),
        ('GRID',          (0,0),(-1,-1), 0.4, C_BORDER),
        ('VALIGN',        (0,0),(-1,-1),'MIDDLE'),
    ] + rc))
    return [t]


def _sec(title, subtitle='', color=None):
    color = color or C_BLUE
    h = _s('sh', fontName='DVB', fontSize=14, textColor=C_DARK,
            leading=18, spaceAfter=2*mm)
    s = _s('ss', fontName='DVI', fontSize=8.5, textColor=C_GRAY,
            leading=12, spaceAfter=3*mm)
    return [Paragraph(title, h),
            *(([Paragraph(subtitle, s)] if subtitle else [])),
            HRFlowable(width='100%', thickness=2, color=color, spaceAfter=4*mm)]


# ── Основная функция ──────────────────────────────────────────────────────────
class _Doc(BaseDocTemplate):
    def __init__(self, path, cover, page_bg, **kw):
        BaseDocTemplate.__init__(self, path, **kw)
        self.addPageTemplates([
            PageTemplate('Cover',
                         [Frame(0,0,W,H, leftPadding=0, rightPadding=0,
                                topPadding=0, bottomPadding=0)],
                         onPage=cover),
            PageTemplate('Content',
                         [Frame(18*mm, 14*mm, W-36*mm, H-31*mm,
                                leftPadding=0, rightPadding=0,
                                topPadding=0, bottomPadding=0)],
                         onPage=page_bg),
        ])


def generate(data: dict, output_path: str,
             progress_cb=None) -> str:
    """
    data — словарь из parser.parse() (возможно отредактированный пользователем).
    output_path — путь для сохранения PDF.
    """
    def _p(step, msg):
        if progress_cb:
            progress_cb(step, 5, msg)

    _p(0, 'Шрифты…')
    _register_fonts()

    equipment = data.get('equipment', [])
    works     = data.get('works',     [])
    kp_num    = data.get('kp_number', 'Коммерческое предложение')

    _p(1, f'Оборудование: {len(equipment)}, работы: {len(works)}')

    story = [Spacer(1, H), NextPageTemplate('Content'), PageBreak()]

    # ── Раздел: Оборудование ─────────────────────────────────────────────────
    story += _sec('Оборудование',
                  'Полный перечень оборудования с фотографиями',
                  C_BLUE)
    for idx, item in enumerate(equipment):
        story.append(_equip_card(item, idx))

    # ── Раздел: Работы ───────────────────────────────────────────────────────
    if works:
        story += _sec('Монтажные работы',
                      'Установка и настройка оборудования',
                      C_GOLD)
        story += _works_table(works)
        story.append(Spacer(1, 6*mm))

    # ── Итоговая страница ────────────────────────────────────────────────────
    story.append(PageBreak())
    story += _sec('Итоговая таблица', '', C_BLUE)
    story += _equip_summary(equipment)
    story.append(Spacer(1, 5*mm))
    if works:
        story += _sec('Работы', '', C_GOLD)
        story += _works_table(works)
        story.append(Spacer(1, 5*mm))

    story += _totals(equipment, works)

    # Подвал
    story.append(Spacer(1, 5*mm))
    mgr = data.get('manager', '').replace('\n', '  |  ')
    footer = Table([[Paragraph(
        f'Цены действительны 3 банковских дня.  {mgr}',
        _s('ft', fontSize=8, textColor=C_WHITE,
            alignment=TA_CENTER, leading=12))]],
        colWidths=[160*mm])
    footer.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), colors.HexColor('#1E3A6E')),
        ('TOPPADDING',    (0,0),(-1,-1), 8),
        ('BOTTOMPADDING', (0,0),(-1,-1), 8),
        ('ROUNDEDCORNERS',[6,6,6,6]),
    ]))
    story.append(footer)

    _p(3, 'Рендер PDF…')
    doc = _Doc(
        output_path,
        cover    = _cover_fn(data),
        page_bg  = _page_fn(kp_num),
        pagesize = A4,
    )
    doc.build(story)
    _p(5, f'Готово: {output_path}')
    return output_path
