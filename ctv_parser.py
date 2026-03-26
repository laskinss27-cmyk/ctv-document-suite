"""
parser.py — универсальный парсер xlsx-файлов КП CTV.
Возвращает структурированные данные: шапка, товары с фото, работы.
"""

import zipfile, re, os, io
from pathlib import Path
import pandas as pd
import openpyxl
from PIL import Image as PILImage


def _clean(text):
    """Убирает HTML-теги и лишние пробелы."""
    if not text:
        return ''
    s = str(text).strip()
    s = re.sub(r'<br\s*/?>', '\n', s, flags=re.IGNORECASE)
    s = re.sub(r'<[^>]+>', '', s)
    s = re.sub(r'&amp;', '&', s)
    s = re.sub(r'&lt;', '<', s)
    s = re.sub(r'&gt;', '>', s)
    s = re.sub(r' +', ' ', s)
    return s.strip()


def _to_float(s):
    try:
        return float(str(s).replace(' ', '').replace(',', '.'))
    except Exception:
        return 0.0


def _extract_images(xlsx_path):
    """Извлекает логотип и фото товаров прямо из zip-структуры xlsx."""
    logo_bytes   = None
    row_to_bytes = {}

    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            names = z.namelist()
            media = {os.path.basename(f): z.read(f)
                     for f in names if 'xl/media' in f}

            # Ищем все drawing-файлы (может быть несколько листов)
            draw_files = [f for f in names if re.match(r'xl/drawings/drawing\d+\.xml$', f)]
            rel_files  = [f for f in names if re.match(r'xl/drawings/_rels/drawing\d+\.xml\.rels$', f)]

            for draw_path, rels_path in zip(sorted(draw_files), sorted(rel_files)):
                draw = z.read(draw_path).decode('utf-8', errors='replace')
                rels  = z.read(rels_path).decode('utf-8', errors='replace')

                rid_to_fname = dict(re.findall(
                    r'Id="(rId\d+)"[^>]+Target="\.\./media/([^"]+)"', rels))

                anchors = re.findall(
                    r'<xdr:from>\s*<xdr:col>(\d+)</xdr:col>\s*<xdr:colOff>[^<]*</xdr:colOff>\s*'
                    r'<xdr:row>(\d+)</xdr:row>.*?r:embed="(rId\d+)"',
                    draw, re.DOTALL)

                for col, row, rid in anchors:
                    fname = rid_to_fname.get(rid, '')
                    b = media.get(fname)
                    if not b:
                        continue
                    col_i = int(col)
                    row_i = int(row)
                    if col_i == 0 and row_i == 0:
                        logo_bytes = b          # логотип в A1
                    elif col_i == 1:
                        row_to_bytes[row_i] = b  # фото товара в колонке B

    except Exception as e:
        pass

    return logo_bytes, row_to_bytes


def parse(xlsx_path: str) -> dict:
    """
    Парсит КП из xlsx.
    Возвращает словарь:
    {
      'kp_number': str,
      'company':   str,    # контакты компании (строка 0)
      'manager':   str,    # менеджер (строка 1, колонка C)
      'logo_bytes': bytes | None,
      'equipment': [
          {
            'article': str,
            'name': str,
            'desc': str,           # полное описание из файла
            'price': float,
            'qty': float,
            'total': float,
            'img_bytes': bytes | None,
          }, ...
      ],
      'works': [
          {
            'name': str,
            'desc': str,
            'price': float,
            'qty': float,
            'total': float,
          }, ...
      ],
    }
    """
    xlsx_path = str(xlsx_path)
    logo_bytes, row_to_bytes = _extract_images(xlsx_path)

    df = pd.read_excel(xlsx_path, header=None, dtype=str).fillna('')

    # ── Шапка ────────────────────────────────────────────────────────────────
    company    = _clean(df.iloc[0, 0]) if len(df) > 0 else ''
    manager    = ''
    kp_number  = ''

    for i in range(min(5, len(df))):
        row = df.iloc[i]
        for j, cell in enumerate(row):
            val = _clean(str(cell))
            if 'Менеджер:' in val or 'Телефон:' in val:
                manager = val
            if 'Коммерческое предложение №' in val:
                kp_number = val

        # Менеджер обычно в col 1 строки 1
        if i == 1 and len(row) > 1:
            val2 = _clean(str(row[1]))
            if val2:
                manager = val2

    # ── Находим секции ───────────────────────────────────────────────────────
    equip_header = equip_data_start = None
    work_header  = work_data_start  = None
    equip_end    = None

    for idx in range(len(df)):
        cell = str(df.iloc[idx, 0]).strip()
        if re.match(r'^ОБОРУДОВАНИЕ', cell):
            equip_header = idx
        elif equip_header is not None and equip_data_start is None:
            # Следующая строка после заголовка колонок — данные
            cols_row = [str(c).strip() for c in df.iloc[idx] if str(c).strip()]
            if 'Артикул' in cols_row or 'Наименование' in cols_row:
                equip_data_start = idx + 1
            elif str(df.iloc[idx, 0]).strip() not in ('', 'nan'):
                # нет строки заголовков — данные сразу
                equip_data_start = idx

        if re.match(r'^РАБОТЫ', cell) and equip_header is not None:
            equip_end    = idx
            work_header  = idx

        if work_header is not None and work_data_start is None and idx > work_header:
            cols_row = [str(c).strip() for c in df.iloc[idx] if str(c).strip()]
            if any(k in cols_row for k in ('Наименование', 'Цена', 'Артикул')):
                work_data_start = idx + 1
            elif str(df.iloc[idx, 0]).strip() not in ('', 'nan', 'РАБОТЫ'):
                work_data_start = idx

    if equip_end is None:
        equip_end = len(df)

    # ── Оборудование ─────────────────────────────────────────────────────────
    equipment = []
    if equip_data_start is not None:
        for i in range(equip_data_start, equip_end):
            row = df.iloc[i]
            cell0 = str(row[0]).strip()

            # Пропускаем итоговые строки и пустые
            if not cell0 or cell0.lower().startswith('итого') or cell0 == 'nan':
                continue
            # Пропускаем строку заголовков если попалась
            if cell0 in ('Артикул', 'Наименование', 'Фото'):
                continue

            # Определяем колонки — в некоторых файлах нет колонки "Фото"
            # Формат A: Арт | Фото | Название | Описание | Цена | Кол-во | Стоим
            # Формат B: Арт | Название | Описание | Цена | Кол-во | Стоим
            r = [str(x).strip() for x in row]

            # Пробуем определить формат по наличию данных
            # Если r[1] это текст (название), значит нет колонки Фото
            def looks_like_name(s):
                return bool(s) and not s.replace('.','').replace(',','').isdigit() and len(s) > 2

            if len(r) >= 7 and looks_like_name(r[2]):
                # Формат с Фото: col0=арт, col1=фото(пусто), col2=название, col3=описание, col4=цена, col5=кол-во, col6=стоим
                article = _clean(r[0])
                name    = _clean(r[2])
                desc    = _clean(r[3])
                price   = _to_float(r[4])
                qty     = _to_float(r[5])
                total   = _to_float(r[6])
            elif len(r) >= 6 and looks_like_name(r[1]):
                # Формат без Фото: col0=арт, col1=название, col2=описание, col3=цена, col4=кол-во, col5=стоим
                article = _clean(r[0])
                name    = _clean(r[1])
                desc    = _clean(r[2])
                price   = _to_float(r[3])
                qty     = _to_float(r[4])
                total   = _to_float(r[5])
            elif len(r) >= 5 and looks_like_name(r[1]):
                article = _clean(r[0])
                name    = _clean(r[1])
                desc    = ''
                price   = _to_float(r[2])
                qty     = _to_float(r[3])
                total   = _to_float(r[4])
            else:
                continue

            if not name and not article:
                continue

            # Фото — ищем по 0-indexed row
            draw_row = i  # df index = 0-indexed row в drawing
            img_bytes = row_to_bytes.get(draw_row)

            equipment.append({
                'article':   article,
                'name':      name,
                'desc':      desc,
                'price':     price,
                'qty':       qty,
                'total':     total if total else price * qty,
                'img_bytes': img_bytes,
            })

    # ── Работы ───────────────────────────────────────────────────────────────
    works = []
    if work_data_start is not None:
        for i in range(work_data_start, len(df)):
            row = df.iloc[i]
            cell0 = str(row[0]).strip()

            if not cell0 or cell0 == 'nan':
                continue
            if cell0.lower().startswith('итого') or cell0.lower().startswith('транспорт'):
                break
            if re.match(r'^(ЦЕНЫ|ВНИМАНИЕ|Итого)', cell0):
                break
            if cell0 in ('Наименование', 'Артикул', 'Описание'):
                continue

            r = [str(x).strip() for x in row]

            # Форматы работ:
            # A: Название | Артикул | Описание | Цена | Кол-во | Стоим
            # B: Название | Описание | Цена | Кол-во | Стоим  (без Артикула)
            # C: Название | Цена | Кол-во | Стоим

            def is_num(s):
                return s.replace('.','').replace(',','').replace(' ','').isdigit() and len(s) > 0

            # Ищем цену — первое числовое значение начиная с col 1
            price = qty = total = 0.0
            desc = ''
            name = _clean(r[0])

            # Пробуем разные смещения
            for offset in range(1, min(len(r), 6)):
                if is_num(r[offset]):
                    # Нашли первое число — это цена
                    # Смотрим дальше
                    nums = []
                    for j in range(offset, len(r)):
                        if is_num(r[j]):
                            nums.append(_to_float(r[j]))
                    # Текст до первого числа — описание
                    texts = [_clean(r[k]) for k in range(1, offset) if _clean(r[k])]
                    desc = ' '.join(texts)

                    if len(nums) >= 3:
                        price, qty, total = nums[0], nums[1], nums[2]
                    elif len(nums) == 2:
                        price, qty = nums[0], nums[1]
                        total = price * qty
                    elif len(nums) == 1:
                        price = nums[0]
                        qty   = 1
                        total = price
                    break

            if not name or (price == 0 and qty == 0):
                continue

            works.append({
                'name':  name,
                'desc':  desc,
                'price': price,
                'qty':   qty,
                'total': total if total else price * qty,
            })

    return {
        'kp_number':  kp_number,
        'company':    company,
        'manager':    manager,
        'logo_bytes': logo_bytes,
        'equipment':  equipment,
        'works':      works,
    }


def img_bytes_to_thumbnail(img_bytes, size=(120, 120)):
    """Конвертирует байты изображения в миниатюру PNG для GUI."""
    if not img_bytes:
        return None
    try:
        pil = PILImage.open(io.BytesIO(img_bytes)).convert('RGB')
        pil.thumbnail(size, PILImage.LANCZOS)
        buf = io.BytesIO()
        pil.save(buf, 'PNG')
        return buf.getvalue()
    except Exception:
        return None
