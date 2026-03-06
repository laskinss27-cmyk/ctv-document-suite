"""
app.py — CTV Document Suite
Список: tk.Text + Frame-строки (крупно, адаптивно, кнопки всегда видны).
Скролл: колесо мыши через tk.Text — работает.
Разделитель: нативный tk.PanedWindow — без тормозов.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading, os, sys, io, copy
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import parser as kp_parser
import kp_pdf
import updater

try:
    import catalog_pdf as catalog_gen
    HAS_CATALOG = True
except ImportError:
    HAS_CATALOG = False

try:
    from PIL import Image as PILImage, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

BG       = '#0F172A'
CARD     = '#1E293B'
CARD2    = '#162032'
BORDER   = '#334155'
BLUE     = '#2563EB'
GOLD     = '#F59E0B'
TEXT     = '#F1F5F9'
SUB      = '#94A3B8'
SUCCESS  = '#22C55E'
ERR      = '#EF4444'
WHITE    = '#FFFFFF'
BTN_KP   = '#B45309'
ITEM_BG  = '#1E293B'


def _open_file(path):
    import subprocess, platform
    try:
        s = platform.system()
        if s == 'Windows':   os.startfile(path)
        elif s == 'Darwin':  subprocess.Popen(['open', path])
        else:                subprocess.Popen(['xdg-open', path])
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
#  ScrollList — tk.Text с embedded Frame-строками
#  Строки крупные, адаптивные. Скролл колесом — нативный через Text.
# ─────────────────────────────────────────────────────────────────────────────
class ScrollList(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=BG, **kw)
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=0)
        self.rowconfigure(0, weight=1)

        sb = ttk.Scrollbar(self, orient='vertical')
        sb.grid(row=0, column=1, sticky='ns')

        self._t = tk.Text(
            self,
            yscrollcommand=sb.set,
            bg=BG, relief='flat', bd=0,
            highlightthickness=0,
            state='disabled', wrap='none',
            cursor='arrow',
            selectbackground=BG,
            inactiveselectbackground=BG,
        )
        self._t.grid(row=0, column=0, sticky='nsew')
        sb.configure(command=self._t.yview)

        self._t.bind('<MouseWheel>', self._wheel)
        self._t.bind('<Button-4>',   self._wheel)
        self._t.bind('<Button-5>',   self._wheel)
        self._first = True

    def _wheel(self, e):
        if e.delta:
            self._t.yview_scroll(int(-1 * (e.delta / 120)), 'units')
        elif e.num == 4:
            self._t.yview_scroll(-1, 'units')
        else:
            self._t.yview_scroll(1, 'units')
        return 'break'

    def _fwd(self, w):
        """Прокидываем колесо от дочерних виджетов к Text (один уровень)."""
        for widget in [w] + w.winfo_children():
            widget.bind('<MouseWheel>', self._wheel, add='+')
            widget.bind('<Button-4>',   self._wheel, add='+')
            widget.bind('<Button-5>',   self._wheel, add='+')

    def add(self, widget):
        self._t.configure(state='normal')
        if not self._first:
            self._t.insert('end', '\n')
        self._first = False
        self._t.window_create('end', window=widget, stretch=True)
        self._t.configure(state='disabled')
        self._fwd(widget)

    def clear(self):
        self._t.configure(state='normal')
        self._t.delete('1.0', 'end')
        self._t.configure(state='disabled')
        self._first = True


# ─────────────────────────────────────────────────────────────────────────────
#  ItemEditorDialog
# ─────────────────────────────────────────────────────────────────────────────
class ItemEditorDialog(tk.Toplevel):

    def __init__(self, parent, item: dict, mode='equipment'):
        super().__init__(parent)
        self.result     = None
        self._item      = copy.deepcopy(item)
        self._mode      = mode
        self._img_bytes = item.get('img_bytes')

        self.title('Товар' if mode == 'equipment' else 'Работа')
        self.geometry('680x580')
        self.configure(bg=BG)
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self._build()
        self._fill()
        self.wait_window(self)

    def _lbl(self, p, t):
        tk.Label(p, text=t, font=('Helvetica', 9, 'bold'),
                 fg=TEXT, bg=BG).pack(anchor='w', pady=(10, 2))

    def _ent(self, p, var):
        e = tk.Entry(p, textvariable=var, font=('Helvetica', 10),
                     fg=TEXT, bg=CARD, insertbackground=TEXT, relief='flat',
                     highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=BLUE)
        e.pack(fill='x', ipady=7, ipadx=6)
        return e

    def _build(self):
        w = tk.Frame(self, bg=BG, padx=20, pady=14)
        w.pack(fill='both', expand=True)

        if self._mode == 'equipment':
            self._lbl(w, 'Артикул')
            self._v_art = tk.StringVar()
            self._ent(w, self._v_art)
            self._lbl(w, 'Наименование')
            self._v_name = tk.StringVar()
            self._ent(w, self._v_name)
            self._lbl(w, 'Описание / Характеристики')
            self._desc = scrolledtext.ScrolledText(
                w, height=8, font=('Helvetica', 9), fg=TEXT, bg=CARD,
                insertbackground=TEXT, relief='flat', wrap='word')
            self._desc.pack(fill='x')
            self._lbl(w, 'Фото')
            pr = tk.Frame(w, bg=BG)
            pr.pack(fill='x')
            self._img_lbl = tk.Label(pr, bg=CARD, width=80, height=80,
                                      text='Нет фото', fg=SUB,
                                      font=('Helvetica', 8))
            self._img_lbl.pack(side='left', padx=(0, 12))
            bc = tk.Frame(pr, bg=BG)
            bc.pack(side='left')
            tk.Button(bc, text='Загрузить фото…', font=('Helvetica', 9),
                      fg=WHITE, bg=CARD, activebackground=BORDER,
                      relief='flat', bd=0, padx=12, pady=6, cursor='hand2',
                      command=self._pick_photo).pack(anchor='w')
            tk.Button(bc, text='Удалить фото', font=('Helvetica', 9),
                      fg=SUB, bg=BG, relief='flat', bd=0,
                      padx=12, pady=4, cursor='hand2',
                      command=self._clear_photo).pack(anchor='w', pady=(4, 0))
        else:
            self._lbl(w, 'Наименование работы')
            self._v_name = tk.StringVar()
            self._ent(w, self._v_name)
            self._lbl(w, 'Описание (необязательно)')
            self._desc = scrolledtext.ScrolledText(
                w, height=4, font=('Helvetica', 9), fg=TEXT, bg=CARD,
                insertbackground=TEXT, relief='flat', wrap='word')
            self._desc.pack(fill='x')

        nums = tk.Frame(w, bg=BG)
        nums.pack(fill='x', pady=(12, 0))
        for attr, lbl, pad in [('_v_price', 'Цена, руб.', (0, 8)),
                                 ('_v_qty',   'Кол-во',    (0, 0))]:
            col = tk.Frame(nums, bg=BG)
            col.pack(side='left', fill='x', expand=True, padx=pad)
            tk.Label(col, text=lbl, font=('Helvetica', 9, 'bold'),
                     fg=TEXT, bg=BG).pack(anchor='w', pady=(0, 2))
            v = tk.StringVar()
            setattr(self, attr, v)
            tk.Entry(col, textvariable=v, font=('Helvetica', 10),
                     fg=TEXT, bg=CARD, insertbackground=TEXT, relief='flat',
                     highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=BLUE).pack(fill='x', ipady=7, ipadx=6)

        tk.Frame(w, bg=BG, height=14).pack()
        btns = tk.Frame(w, bg=BG)
        btns.pack(fill='x')
        tk.Button(btns, text='💾  Сохранить',
                  font=('Helvetica', 11, 'bold'), fg=WHITE, bg=BLUE,
                  activebackground='#1D4ED8', relief='flat', bd=0,
                  padx=20, pady=11, cursor='hand2',
                  command=self._save).pack(side='left', fill='x',
                                            expand=True, padx=(0, 6))
        tk.Button(btns, text='Отмена',
                  font=('Helvetica', 11), fg=SUB, bg=CARD,
                  activebackground=BORDER, relief='flat', bd=0,
                  padx=20, pady=11, cursor='hand2',
                  command=self.destroy).pack(side='left', fill='x', expand=True)

    def _fill(self):
        if self._mode == 'equipment':
            self._v_art.set(self._item.get('article', ''))
            self._v_name.set(self._item.get('name', ''))
            self._desc.insert('1.0', self._item.get('desc', ''))
            self._refresh_photo()
        else:
            self._v_name.set(self._item.get('name', ''))
            self._desc.insert('1.0', self._item.get('desc', ''))
        self._v_price.set(str(self._item.get('price', '')))
        self._v_qty.set(str(self._item.get('qty', 1)))

    def _refresh_photo(self):
        if not HAS_PIL:
            return
        if self._img_bytes:
            try:
                pil = PILImage.open(io.BytesIO(self._img_bytes)).convert('RGB')
                pil.thumbnail((80, 80), PILImage.LANCZOS)
                tk_img = ImageTk.PhotoImage(pil)
                self._img_lbl.configure(image=tk_img, text='')
                self._img_lbl._img = tk_img
            except Exception:
                pass
        else:
            self._img_lbl.configure(image='', text='Нет фото')

    def _pick_photo(self):
        path = filedialog.askopenfilename(
            filetypes=[('Изображения', '*.png *.jpg *.jpeg *.webp *.bmp'),
                       ('Все файлы', '*.*')])
        if path:
            with open(path, 'rb') as f:
                self._img_bytes = f.read()
            self._refresh_photo()

    def _clear_photo(self):
        self._img_bytes = None
        self._refresh_photo()

    def _save(self):
        try:
            price = float(str(self._v_price.get()).replace(',', '.').replace(' ', '') or 0)
            qty   = float(str(self._v_qty.get()).replace(',', '.').replace(' ', '') or 1)
        except ValueError:
            messagebox.showerror('Ошибка', 'Цена и кол-во должны быть числами.')
            return
        self.result = copy.deepcopy(self._item)
        self.result['name']  = self._v_name.get().strip()
        self.result['price'] = price
        self.result['qty']   = qty
        self.result['total'] = price * qty
        if self._mode == 'equipment':
            self.result['article']   = self._v_art.get().strip()
            self.result['img_bytes'] = self._img_bytes
            self.result['desc']      = self._desc.get('1.0', 'end').strip()
        else:
            self.result['desc'] = self._desc.get('1.0', 'end').strip()
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
#  App
# ─────────────────────────────────────────────────────────────────────────────
class App(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title('CTV Document Suite')
        self.geometry('980x720')
        self.minsize(840, 560)
        self.configure(bg=BG)

        self._data    = None
        self._xlsx    = tk.StringVar()
        self._out     = tk.StringVar()
        self._status  = tk.StringVar(value='Загрузите xlsx-файл для начала работы')
        self._prog    = tk.IntVar(value=0)
        self._running = False

        self._build_ui()
        self._auto_updater = updater.AutoUpdater(self, self._on_update_found)
        self._auto_updater.start()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        tk.Frame(self, bg=BLUE, height=5).pack(fill='x')

        hdr = tk.Frame(self, bg=BG, pady=12)
        hdr.pack(fill='x', padx=24)
        tk.Label(hdr, text='CTV Document Suite',
                 font=('Helvetica', 20, 'bold'), fg=WHITE, bg=BG).pack(anchor='w')
        tk.Label(hdr, text='Загрузите Excel → отредактируйте → создайте PDF',
                 font=('Helvetica', 9), fg=SUB, bg=BG).pack(anchor='w')
        tk.Frame(self, bg=BORDER, height=1).pack(fill='x')

        top = tk.Frame(self, bg=BG, padx=24, pady=10)
        top.pack(fill='x')
        self._build_top(top)
        tk.Frame(self, bg=BORDER, height=1).pack(fill='x')

        self._build_editor()
        self._build_bottom()

    def _build_top(self, p):
        for var, lbl, cmd in [(self._xlsx, 'Файл xlsx:', self._pick_xlsx),
                               (self._out,  'Сохранить:', self._pick_out)]:
            row = tk.Frame(p, bg=BG)
            row.pack(fill='x', pady=(0, 5))
            tk.Label(row, text=lbl, font=('Helvetica', 9, 'bold'),
                     fg=TEXT, bg=BG, width=10, anchor='w').pack(side='left')
            tk.Entry(row, textvariable=var, font=('Helvetica', 9),
                     fg=TEXT, bg=CARD, insertbackground=TEXT, relief='flat',
                     highlightthickness=1, highlightbackground=BORDER,
                     highlightcolor=BLUE).pack(side='left', fill='x',
                                                expand=True, ipady=6, ipadx=6)
            tk.Button(row,
                      text='Открыть…' if var is self._xlsx else 'Выбрать…',
                      font=('Helvetica', 9), fg=WHITE, bg=CARD,
                      activebackground=BORDER, relief='flat', bd=0,
                      padx=12, pady=6, cursor='hand2',
                      command=cmd).pack(side='left', padx=(6, 0))

        btns = tk.Frame(p, bg=BG)
        btns.pack(fill='x', pady=(4, 0))
        if HAS_CATALOG:
            self._btn_cat = tk.Button(
                btns, text='📋  Создать Каталог',
                font=('Helvetica', 11, 'bold'), fg=WHITE, bg=BLUE,
                activebackground='#1D4ED8', relief='flat', bd=0,
                padx=18, pady=10, cursor='hand2',
                command=lambda: self._generate('catalog'))
            self._btn_cat.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self._btn_kp = tk.Button(
            btns, text='📄  Создать КП',
            font=('Helvetica', 11, 'bold'), fg=WHITE, bg=BTN_KP,
            activebackground='#92400E', relief='flat', bd=0,
            padx=18, pady=10, cursor='hand2',
            command=lambda: self._generate('kp'))
        self._btn_kp.pack(side='left', fill='x',
                           expand=True, padx=(5 if HAS_CATALOG else 0, 0))

    def _build_editor(self):
        # Нативный PanedWindow — сам управляет sash, не тормозит
        paned = tk.PanedWindow(self, orient='horizontal',
                                bg=BORDER, sashwidth=4,
                                sashrelief='flat', sashpad=0,
                                opaqueresize=True)
        paned.pack(fill='both', expand=True)

        # ── Оборудование ──
        left = tk.Frame(paned, bg=BG)
        left.pack_propagate(False)
        paned.add(left, minsize=300, stretch='always')

        tk.Label(left, text='  Оборудование',
                 font=('Helvetica', 10, 'bold'), fg=GOLD, bg=CARD2,
                 anchor='w', pady=7).pack(fill='x')
        self._equip_list = ScrollList(left)
        self._equip_list.pack(fill='both', expand=True)
        tk.Button(left, text='＋  Добавить товар',
                  font=('Helvetica', 9), fg=SUB, bg=CARD2,
                  activebackground=CARD, relief='flat', bd=0,
                  padx=10, pady=7, cursor='hand2',
                  command=self._add_equipment).pack(fill='x')

        # ── Работы ──
        right = tk.Frame(paned, bg=BG)
        right.pack_propagate(False)
        paned.add(right, minsize=260, stretch='always')

        tk.Label(right, text='  Работы',
                 font=('Helvetica', 10, 'bold'), fg=TEXT, bg=CARD2,
                 anchor='w', pady=7).pack(fill='x')
        self._works_list = ScrollList(right)
        self._works_list.pack(fill='both', expand=True)
        tk.Button(right, text='＋  Добавить работу',
                  font=('Helvetica', 9), fg=SUB, bg=CARD2,
                  activebackground=CARD, relief='flat', bd=0,
                  padx=10, pady=7, cursor='hand2',
                  command=self._add_work).pack(fill='x')

    def _build_bottom(self):
        bot = tk.Frame(self, bg=BG, padx=24, pady=8)
        bot.pack(fill='x', side='bottom')
        sty = ttk.Style(self)
        sty.theme_use('clam')
        sty.configure('CTV.Horizontal.TProgressbar',
                       troughcolor=CARD, background=BLUE,
                       bordercolor=BORDER, lightcolor=BLUE, darkcolor=BLUE)
        ttk.Progressbar(bot, variable=self._prog, maximum=5,
                        style='CTV.Horizontal.TProgressbar'
                        ).pack(fill='x', pady=(0, 4))
        self._stat_lbl = tk.Label(bot, textvariable=self._status,
                                   font=('Helvetica', 9), fg=SUB, bg=BG,
                                   anchor='w')
        self._stat_lbl.pack(fill='x')

    # ── Рендер ────────────────────────────────────────────────────────────────
    def _render_lists(self):
        self._render_equip()
        self._render_works()

    def _render_equip(self):
        self._equip_list.clear()
        if not self._data:
            return
        for idx, item in enumerate(self._data['equipment']):
            self._equip_list.add(self._make_equip_row(item, idx))

    def _render_works(self):
        self._works_list.clear()
        if not self._data:
            return
        for idx, item in enumerate(self._data['works']):
            self._works_list.add(self._make_work_row(item, idx))

    def _make_equip_row(self, item, idx):
        row = tk.Frame(self._equip_list, bg=ITEM_BG, pady=6, padx=10)

        # Фото 46×46
        if HAS_PIL and item.get('img_bytes'):
            try:
                pil = PILImage.open(io.BytesIO(item['img_bytes'])).convert('RGB')
                pil.thumbnail((46, 46), PILImage.LANCZOS)
                tk_img = ImageTk.PhotoImage(pil)
                lbl = tk.Label(row, image=tk_img, bg=ITEM_BG)
                lbl._img = tk_img
                lbl.pack(side='left', padx=(0, 10))
            except Exception:
                pass

        # Текст
        info = tk.Frame(row, bg=ITEM_BG)
        info.pack(side='left', fill='both', expand=True)

        tk.Label(info,
                 text=item.get('name', '—'),
                 font=('Helvetica', 10, 'bold'),
                 fg=TEXT, bg=ITEM_BG,
                 anchor='w', justify='left',
                 wraplength=1).pack(anchor='w', fill='x')   # wraplength обновится через bind

        art   = item.get('article', '')
        price = item.get('price', 0)
        qty   = item.get('qty',   1)
        parts = ([f'Арт. {art}'] if art else []) + [f'{price:,.0f} руб. × {qty:g} шт.']
        tk.Label(info,
                 text='   ·   '.join(parts),
                 font=('Helvetica', 9),
                 fg=SUB, bg=ITEM_BG,
                 anchor='w').pack(anchor='w')

        # Кнопки — всегда справа, всегда видны
        btns = tk.Frame(row, bg=ITEM_BG)
        btns.pack(side='right', padx=(6, 0))
        for sym, clr, cmd in [
            ('✏', GOLD, lambda i=idx: self._edit_equip(i)),
            ('↑', SUB,  lambda i=idx: self._move_equip(i, -1)),
            ('↓', SUB,  lambda i=idx: self._move_equip(i,  1)),
            ('✕', ERR,  lambda i=idx: self._del_equip(i)),
        ]:
            tk.Button(btns, text=sym,
                      font=('Helvetica', 13),
                      fg=clr, bg=ITEM_BG,
                      activebackground=CARD,
                      relief='flat', bd=0,
                      padx=5, pady=2,
                      cursor='hand2',
                      command=cmd).pack(side='left')

        # Разделитель снизу
        tk.Frame(row, bg=BORDER, height=1).pack(
            side='bottom', fill='x', pady=(6, 0))

        return row

    def _make_work_row(self, item, idx):
        row = tk.Frame(self._works_list, bg=ITEM_BG, pady=6, padx=10)

        info = tk.Frame(row, bg=ITEM_BG)
        info.pack(side='left', fill='both', expand=True)

        tk.Label(info,
                 text=item.get('name', '—'),
                 font=('Helvetica', 10, 'bold'),
                 fg=TEXT, bg=ITEM_BG,
                 anchor='w', justify='left').pack(anchor='w', fill='x')

        tk.Label(info,
                 text=f'{item.get("price",0):,.0f} руб. × {item.get("qty",1):g}',
                 font=('Helvetica', 9),
                 fg=SUB, bg=ITEM_BG,
                 anchor='w').pack(anchor='w')

        btns = tk.Frame(row, bg=ITEM_BG)
        btns.pack(side='right', padx=(6, 0))
        for sym, clr, cmd in [
            ('✏', GOLD, lambda i=idx: self._edit_work(i)),
            ('↑', SUB,  lambda i=idx: self._move_work(i, -1)),
            ('↓', SUB,  lambda i=idx: self._move_work(i,  1)),
            ('✕', ERR,  lambda i=idx: self._del_work(i)),
        ]:
            tk.Button(btns, text=sym,
                      font=('Helvetica', 13),
                      fg=clr, bg=ITEM_BG,
                      activebackground=CARD,
                      relief='flat', bd=0,
                      padx=5, pady=2,
                      cursor='hand2',
                      command=cmd).pack(side='left')

        tk.Frame(row, bg=BORDER, height=1).pack(
            side='bottom', fill='x', pady=(6, 0))

        return row

    # ── Файловые диалоги ──────────────────────────────────────────────────────
    def _pick_xlsx(self):
        p = filedialog.askopenfilename(
            title='Выберите xlsx КП',
            filetypes=[('Excel', '*.xlsx'), ('Все', '*.*')])
        if not p:
            return
        self._xlsx.set(p)
        if not self._out.get():
            self._out.set(str(Path(p).with_suffix('.pdf')))
        self._load_xlsx(p)

    def _pick_out(self):
        init = self._out.get() or str(Path.home() / 'output.pdf')
        p = filedialog.asksaveasfilename(
            defaultextension='.pdf',
            initialfile=Path(init).name,
            initialdir=str(Path(init).parent),
            filetypes=[('PDF', '*.pdf'), ('Все', '*.*')])
        if p:
            self._out.set(p)

    def _load_xlsx(self, path):
        self._set_status('Читаю файл…', SUB)
        def _w():
            try:
                data = kp_parser.parse(path)
                self.after(0, self._on_loaded, data)
            except Exception as ex:
                self.after(0, self._set_status, f'Ошибка: {ex}', ERR)
        threading.Thread(target=_w, daemon=True).start()

    def _on_loaded(self, data):
        self._data = data
        self._render_lists()
        self._set_status(
            f'✅  {data["kp_number"]}  |  '
            f'Оборудование: {len(data["equipment"])}  |  '
            f'Работы: {len(data["works"])}', SUCCESS)

    # ── Редактирование ────────────────────────────────────────────────────────
    def _edit_equip(self, idx):
        dlg = ItemEditorDialog(self, self._data['equipment'][idx], 'equipment')
        if dlg.result:
            self._data['equipment'][idx] = dlg.result
            self._render_equip()

    def _edit_work(self, idx):
        dlg = ItemEditorDialog(self, self._data['works'][idx], 'work')
        if dlg.result:
            self._data['works'][idx] = dlg.result
            self._render_works()

    def _add_equipment(self):
        if not self._data:
            messagebox.showinfo('', 'Сначала загрузите xlsx-файл.')
            return
        blank = {'article': '', 'name': '', 'desc': '',
                 'price': 0.0, 'qty': 1.0, 'total': 0.0, 'img_bytes': None}
        dlg = ItemEditorDialog(self, blank, 'equipment')
        if dlg.result:
            self._data['equipment'].append(dlg.result)
            self._render_equip()

    def _add_work(self):
        if not self._data:
            messagebox.showinfo('', 'Сначала загрузите xlsx-файл.')
            return
        blank = {'name': '', 'desc': '', 'price': 0.0, 'qty': 1.0, 'total': 0.0}
        dlg = ItemEditorDialog(self, blank, 'work')
        if dlg.result:
            self._data['works'].append(dlg.result)
            self._render_works()

    def _del_equip(self, idx):
        if messagebox.askyesno('Удалить?', 'Удалить этот товар?'):
            del self._data['equipment'][idx]
            self._render_equip()

    def _del_work(self, idx):
        if messagebox.askyesno('Удалить?', 'Удалить эту работу?'):
            del self._data['works'][idx]
            self._render_works()

    def _move_equip(self, idx, d):
        lst = self._data['equipment']
        n = idx + d
        if 0 <= n < len(lst):
            lst[idx], lst[n] = lst[n], lst[idx]
            self._render_equip()

    def _move_work(self, idx, d):
        lst = self._data['works']
        n = idx + d
        if 0 <= n < len(lst):
            lst[idx], lst[n] = lst[n], lst[idx]
            self._render_works()

    # ── Генерация PDF ─────────────────────────────────────────────────────────
    def _generate(self, mode):
        if self._running:
            return
        if not self._data:
            messagebox.showwarning('', 'Сначала загрузите xlsx-файл.')
            return
        out = self._out.get().strip()
        if not out:
            messagebox.showwarning('', 'Укажите путь для сохранения PDF.')
            return
        self._running = True
        self._prog.set(0)
        if HAS_CATALOG:
            self._btn_cat.configure(state='disabled')
        self._btn_kp.configure(state='disabled')
        label = 'Каталог' if mode == 'catalog' else 'КП'
        self._set_status(f'Генерация {label}…', SUB)
        data = copy.deepcopy(self._data)

        def _w():
            try:
                if mode == 'kp':
                    res = kp_pdf.generate(data, out, progress_cb=self._on_prog)
                else:
                    res = catalog_gen.generate(self._xlsx.get(), out,
                                               progress_cb=self._on_prog)
                self.after(0, self._done, res, label)
            except Exception as ex:
                import traceback; traceback.print_exc()
                self.after(0, self._error, str(ex))

        threading.Thread(target=_w, daemon=True).start()

    def _on_prog(self, step, total, msg):
        self.after(0, self._prog.set, step)
        self.after(0, self._set_status, msg, SUB)

    def _done(self, path, label):
        self._running = False
        self._prog.set(5)
        if HAS_CATALOG:
            self._btn_cat.configure(state='normal')
        self._btn_kp.configure(state='normal')
        self._set_status(f'✅  {label} сохранён: {path}', SUCCESS)
        if messagebox.askyesno('Готово!', f'{label} создан!\n\n{path}\n\nОткрыть?'):
            _open_file(path)

    def _error(self, err):
        self._running = False
        if HAS_CATALOG:
            self._btn_cat.configure(state='normal')
        self._btn_kp.configure(state='normal')
        self._set_status(f'❌  {err}', ERR)
        messagebox.showerror('Ошибка', err)

    def _set_status(self, msg, color=SUB):
        self._status.set(msg)
        self._stat_lbl.configure(fg=color)

    # ── Обновления ────────────────────────────────────────────────────────────
    def _on_update_found(self, info: dict):
        ver      = info.get('version', '?')
        changelog = info.get('changelog', '')
        dl_url   = info.get('download_url', '')
        msg = f'Доступна новая версия {ver}!'
        if changelog:
            msg += f'\n\nЧто нового:\n{changelog}'
        msg += '\n\nУстановить сейчас?'
        if not messagebox.askyesno('🆕 Обновление', msg):
            return
        if not dl_url:
            messagebox.showerror('Ошибка', 'Ссылка не найдена.')
            return
        dlg = tk.Toplevel(self)
        dlg.title('Обновление…')
        dlg.geometry('420x130')
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()
        tk.Label(dlg, text=f'Загрузка версии {ver}…',
                 font=('Helvetica', 11, 'bold'), fg=TEXT, bg=BG
                 ).pack(pady=(16, 4))
        mv = tk.StringVar(value='Подключение…')
        tk.Label(dlg, textvariable=mv,
                 font=('Helvetica', 9), fg=SUB, bg=BG).pack()
        pv = tk.IntVar()
        sty = ttk.Style()
        sty.configure('Upd.Horizontal.TProgressbar',
                       troughcolor=CARD, background=BLUE,
                       bordercolor=BORDER, lightcolor=BLUE, darkcolor=BLUE)
        ttk.Progressbar(dlg, variable=pv, maximum=100,
                        style='Upd.Horizontal.TProgressbar'
                        ).pack(fill='x', padx=20, pady=10)

        def _prog(pct, text):
            if pct >= 0: pv.set(pct)
            mv.set(text); dlg.update()

        def _worker():
            ok = updater.download_and_apply(dl_url, ver, _prog)
            self.after(0, _finish, ok)

        def _finish(ok):
            dlg.destroy()
            if ok:
                messagebox.showinfo('Готово!',
                                    f'Версия {ver} установлена.\nПрограмма перезапустится.')
                updater.quit_for_update()
            else:
                messagebox.showerror('Ошибка', 'Не удалось скачать обновление.')

        threading.Thread(target=_worker, daemon=True).start()


if __name__ == '__main__':
    App().mainloop()
