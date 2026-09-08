# -*- coding: utf-8 -*-
"""
app.py - 跑分绩效汇总自动生成工具(入口 + GUI)

界面/流程参考《安全生产责任制履职清单考评表自动生成工具》:
  tkinter + tkinterdnd2 图形界面; 照片拖拽/点选添加; 后台线程 OCR 不卡界面;
  识别结果预览(可人工核对); 一键生成 xlsx; 完成后一键打开文件/文件夹;
  输出与模板版式完全一致(公式实时计算), 模板只读、生成用副本。

核心逻辑:
  detail_core.py  考核明细照片 OCR 解析(标题/条目/姓名匹配)
  excel_core.py   模板名单读取与按明细填表(日格/合计/金额公式/明细存档)
"""
import json
import logging
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import detail_core
import excel_core

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # 未装 tkinterdnd2 时退化为普通 tkinter(无拖拽)
    DND_FILES = None
    TkinterDnD = None

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(HERE, 'paofen_config.json')
LOG_PATH = os.path.join(HERE, 'paofen.log')
IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.bmp', '.webp')

logger = logging.getLogger('paofen')
logger.setLevel(logging.INFO)
_fh = logging.FileHandler(LOG_PATH, encoding='utf-8')
_fh.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
logger.addHandler(_fh)
logging.getLogger('rapidocr_onnxruntime').setLevel(logging.WARNING)

DEFAULT_CONFIG = {
    'template': '',
    'out_dir': HERE,
    'pool_amount': '',
    'base_score': excel_core.BASE_SCORE_DEFAULT,
}


def load_config():
    cfg = dict(DEFAULT_CONFIG)
    try:
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            cfg.update(json.load(f))
    except (OSError, ValueError):
        pass
    return cfg


def save_config(cfg):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError as e:
        logger.warning('配置保存失败: %s', e)


class PaofenApp:
    def __init__(self, root):
        self.root = root
        root.title('跑分绩效汇总自动生成工具')
        root.geometry('1000x760')
        root.minsize(860, 640)
        self.cfg = load_config()
        self.photos = []      # [{'path': str, 'override': ''|'甲班'...}]
        self.results = {}     # path -> parse_photo 结果
        self.q = queue.Queue()
        self.busy = False
        self.out_path = None
        self._build_ui()
        self._load_initial_config()
        self.root.after(120, self._poll)

    # ---------------- 界面 ----------------
    def _build_ui(self):
        pad = dict(padx=8, pady=4)
        # ---- 顶部: 基本信息 ----
        top = ttk.LabelFrame(self.root, text=' 基本信息 ')
        top.pack(fill='x', **pad)
        ttk.Label(top, text='模板文件:').grid(row=0, column=0, sticky='e', **pad)
        self.var_template = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_template).grid(
            row=0, column=1, sticky='we', **pad)
        ttk.Button(top, text='选择…', width=8,
                   command=self._pick_template).grid(row=0, column=2, **pad)
        ttk.Label(top, text='输出目录:').grid(row=1, column=0, sticky='e', **pad)
        self.var_out = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_out).grid(
            row=1, column=1, sticky='we', **pad)
        ttk.Button(top, text='选择…', width=8,
                   command=self._pick_outdir).grid(row=1, column=2, **pad)
        ttk.Label(top, text='月份(空=自动):').grid(row=0, column=3, sticky='e', **pad)
        self.var_month = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_month, width=8).grid(row=0, column=4, **pad)
        ttk.Label(top, text='基础分:').grid(row=1, column=3, sticky='e', **pad)
        self.var_base = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_base, width=8).grid(row=1, column=4, **pad)
        ttk.Label(top, text='绩效总额(空=模板值):').grid(row=0, column=5, sticky='e', **pad)
        self.var_pool = tk.StringVar()
        ttk.Entry(top, textvariable=self.var_pool, width=12).grid(row=0, column=6, **pad)
        top.columnconfigure(1, weight=1)

        # ---- 中部: 明细照片 ----
        mid = ttk.LabelFrame(self.root, text=' 考核明细照片(每班一张, 可拖入) ')
        mid.pack(fill='both', expand=True, **pad)
        cols = ('file', 'shift', 'count', 'status')
        self.tv_photos = ttk.Treeview(mid, columns=cols, show='headings', height=5)
        for c, w, t in (('file', 380, '照片文件'), ('shift', 90, '班次'),
                        ('count', 70, '条目数'), ('status', 300, '状态')):
            self.tv_photos.heading(c, text=t)
            self.tv_photos.column(c, width=w, anchor='w')
        self.tv_photos.pack(side='left', fill='both', expand=True, padx=(6, 0), pady=6)
        vsb = ttk.Scrollbar(mid, orient='vertical', command=self.tv_photos.yview)
        vsb.pack(side='left', fill='y', pady=6)
        self.tv_photos.configure(yscrollcommand=vsb.set)
        if DND_FILES:
            self.tv_photos.drop_target_register(DND_FILES)
            self.tv_photos.dnd_bind('<<Drop>>', self._on_drop)
        btns = ttk.Frame(mid)
        btns.pack(side='left', fill='y', padx=6, pady=6)
        ttk.Button(btns, text='添加照片…', command=self._pick_photos).pack(fill='x', pady=2)
        ttk.Button(btns, text='移除选中', command=self._remove_photo).pack(fill='x', pady=2)
        ttk.Button(btns, text='清空全部', command=self._clear_photos).pack(fill='x', pady=2)
        ttk.Separator(btns, orient='horizontal').pack(fill='x', pady=6)
        self.var_override = tk.StringVar(value='自动')
        ov = ttk.Combobox(btns, textvariable=self.var_override, width=8,
                          state='readonly', values=('自动',) + detail_core.SHIFT_NAMES)
        ov.pack(fill='x', pady=2)
        ttk.Button(btns, text='指定选中\n照片班次', command=self._apply_override
                   ).pack(fill='x', pady=2)
        ttk.Button(btns, text='识别照片', command=self.start_identify
                   ).pack(fill='x', pady=(10, 2))

        # ---- 中下: 识别结果预览 ----
        prev = ttk.LabelFrame(self.root, text=' 识别结果预览(可人工核对) ')
        prev.pack(fill='both', expand=True, **pad)
        cols = ('photo', 'shift', 'seq', 'date', 'name', 'match', 'reason', 'score')
        self.tv_entries = ttk.Treeview(prev, columns=cols, show='headings', height=9)
        for c, w, t in (('photo', 70, '照片'), ('shift', 50, '班次'),
                        ('seq', 40, '序号'), ('date', 55, '日期'),
                        ('name', 90, '姓名(识别)'), ('match', 90, '匹配名单'),
                        ('reason', 260, '事由'), ('score', 55, '分值')):
            self.tv_entries.heading(c, text=t)
            self.tv_entries.column(c, width=w, anchor='w')
        self.tv_entries.pack(fill='both', expand=True, padx=6, pady=(6, 2))
        self.txt_warn = tk.Text(prev, height=5, wrap='word', fg='#a33')
        self.txt_warn.pack(fill='x', padx=6, pady=(0, 6))
        self.txt_warn.insert('1.0', '警告/提示会显示在这里')
        self.txt_warn.config(state='disabled')

        # ---- 底部: 进度与操作 ----
        bottom = ttk.Frame(self.root)
        bottom.pack(fill='x', **pad)
        self.progress = ttk.Progressbar(bottom, mode='determinate', length=260)
        self.progress.pack(side='left', padx=(8, 4))
        self.var_status = tk.StringVar(value='请添加考核明细照片(支持拖拽)')
        ttk.Label(bottom, textvariable=self.var_status, width=52, anchor='w'
                  ).pack(side='left', padx=4)
        self.btn_open_file = ttk.Button(bottom, text='打开文件', state='disabled',
                                        command=self._open_file)
        self.btn_open_dir = ttk.Button(bottom, text='打开所在文件夹', state='disabled',
                                       command=self._open_folder)
        self.btn_generate = ttk.Button(bottom, text=' 一键生成 ', command=self.start_generate)
        self.btn_open_file.pack(side='right', padx=4, pady=6)
        self.btn_open_dir.pack(side='right', padx=4, pady=6)
        self.btn_generate.pack(side='right', padx=(4, 10), pady=6)

    # ---------------- 配置/文件选择 ----------------
    def _load_initial_config(self):
        self.var_template.set(self.cfg.get('template') or '')
        if not self.var_template.get():
            t = excel_core.find_template(HERE)
            if t:
                self.var_template.set(t)
        self.var_out.set(self.cfg.get('out_dir') or HERE)
        self.var_pool.set(str(self.cfg.get('pool_amount') or ''))
        self.var_base.set(str(self.cfg.get('base_score') or excel_core.BASE_SCORE_DEFAULT))

    def _persist_config(self):
        self.cfg.update({
            'template': self.var_template.get().strip(),
            'out_dir': self.var_out.get().strip(),
            'pool_amount': self.var_pool.get().strip(),
            'base_score': self.var_base.get().strip(),
        })
        save_config(self.cfg)

    def _pick_template(self):
        p = filedialog.askopenfilename(title='选择跑分绩效汇总模板',
                                       filetypes=[('Excel 模板', '*.xlsx')])
        if p:
            self.var_template.set(p)
            self._persist_config()

    def _pick_outdir(self):
        p = filedialog.askdirectory(title='选择输出目录')
        if p:
            self.var_out.set(p)
            self._persist_config()

    # ---------------- 照片管理 ----------------
    def _pick_photos(self):
        ps = filedialog.askopenfilenames(
            title='选择考核明细照片',
            filetypes=[('图片', ' '.join('*' + e for e in IMAGE_EXTS))])
        self._add_photos(list(ps))

    def _on_drop(self, event):
        try:
            paths = [p for p in self.root.tk.splitlist(event.data)]
        except Exception:
            paths = [event.data]
        self._add_photos(paths)

    def _add_photos(self, paths):
        added = 0
        for p in paths:
            p = p.strip('{}')
            if not p or not os.path.isfile(p):
                continue
            if not p.lower().endswith(IMAGE_EXTS):
                continue
            if any(ph['path'] == p for ph in self.photos):
                continue
            self.photos.append({'path': p, 'override': ''})
            added += 1
        if added:
            self._refresh_photo_tree()
            self._set_status(f'已添加 {added} 张照片, 共 {len(self.photos)} 张')

    def _remove_photo(self):
        sel = self.tv_photos.selection()
        if not sel:
            return
        idx = self.tv_photos.index(sel[0])
        if 0 <= idx < len(self.photos):
            removed = self.photos.pop(idx)
            self.results.pop(removed['path'], None)
            self._refresh_photo_tree()
            self._refresh_preview()

    def _clear_photos(self):
        self.photos.clear()
        self.results.clear()
        self._refresh_photo_tree()
        self._refresh_preview()

    def _apply_override(self):
        sel = self.tv_photos.selection()
        if not sel:
            self._set_status('请先在照片列表中选中一行')
            return
        idx = self.tv_photos.index(sel[0])
        v = self.var_override.get()
        self.photos[idx]['override'] = '' if v == '自动' else v
        self.results.pop(self.photos[idx]['path'], None)  # 需重新识别
        self._refresh_photo_tree()
        self._refresh_preview()

    # ---------------- 识别(后台线程) ----------------
    def start_identify(self):
        if self.busy:
            messagebox.showinfo('提示', '正在处理中, 请稍候…')
            return
        if not self.photos:
            messagebox.showinfo('提示', '请先添加考核明细照片')
            return
        template = self.var_template.get().strip()
        if not os.path.isfile(template):
            messagebox.showerror('错误', f'模板文件不存在:\n{template}\n'
                                  '请点击「选择…」指定「煤库作业区跑分绩效汇总」xlsx 模板')
            return
        self._persist_config()
        self._set_busy(True, total=len(self.photos))
        threading.Thread(target=self._identify_worker, daemon=True).start()

    def _identify_worker(self):
        try:
            rosters = excel_core.load_rosters(self.var_template.get().strip())
            for i, ph in enumerate(self.photos):
                path, ov = ph['path'], ph['override']
                self.q.put(('status', f'正在识别({i + 1}/{len(self.photos)}): '
                                      f'{os.path.basename(path)}'))
                try:
                    res = detail_core.parse_photo(path, rosters, shift_override=ov)
                except Exception as e:
                    logger.exception('识别失败 %s', path)
                    res = {'path': path, 'month': None, 'shift': ov or None,
                           'entries': [], 'rotate': 0, 'raw_count': 0,
                           'warnings': [f'识别失败: {e}']}
                self.results[path] = res
                self.q.put(('photo_done', path))
            self.q.put(('identify_done', None))
        except Exception:
            logger.exception('识别线程异常')
            self.q.put(('fatal', traceback.format_exc()))

    # ---------------- 生成(后台线程) ----------------
    def start_generate(self):
        if self.busy:
            messagebox.showinfo('提示', '正在处理中, 请稍候…')
            return
        template = self.var_template.get().strip()
        if not os.path.isfile(template):
            messagebox.showerror('错误', f'模板文件不存在:\n{template}')
            return
        if not self.photos:
            messagebox.showinfo('提示', '请先添加考核明细照片')
            return
        self._persist_config()
        try:
            base = int(str(self.var_base.get()).strip() or excel_core.BASE_SCORE_DEFAULT)
        except ValueError:
            messagebox.showerror('错误', '基础分必须是整数, 如 80')
            return
        pool_txt = self.var_pool.get().strip()
        try:
            pool = float(pool_txt) if pool_txt else None
        except ValueError:
            messagebox.showerror('错误', '绩效总额必须是数字, 如 10800')
            return
        month_txt = self.var_month.get().strip()
        try:
            month = int(month_txt) if month_txt else None
        except ValueError:
            messagebox.showerror('错误', '月份必须是整数, 如 7')
            return
        missing = [ph['path'] for ph in self.photos if ph['path'] not in self.results]
        self._set_busy(True, total=len(self.photos) + 1)
        threading.Thread(target=self._generate_worker,
                         args=(template, missing, base, pool, month),
                         daemon=True).start()

    def _generate_worker(self, template, missing, base, pool, month):
        try:
            rosters = excel_core.load_rosters(template)
            for i, path in enumerate(missing):
                ph = next(p for p in self.photos if p['path'] == path)
                self.q.put(('status', f'正在识别未识别照片({i + 1}/{len(missing)})…'))
                res = detail_core.parse_photo(path, rosters,
                                              shift_override=ph['override'])
                self.results[path] = res
                self.q.put(('photo_done', path))
            self.q.put(('status', '正在生成 Excel…'))
            merged = detail_core.merge_results(
                [self.results[ph['path']] for ph in self.photos], default_month=month)
            m = month or merged['month']
            if not m:
                self.q.put(('fatal', '未能确定月份: 照片标题未识别到「X月」'
                                    '且未在界面填写月份, 请手动填写后重试'))
                return
            out_dir = self.var_out.get().strip() or HERE
            os.makedirs(out_dir, exist_ok=True)
            out_path = excel_core.build_output_path(out_dir, m)
            report = excel_core.fill_from_details(
                template, out_path, merged['shift_entries'], month=m,
                base_score=base, pool_amount=pool)
            self.q.put(('generate_done', (out_path, m, merged, report)))
        except Exception:
            logger.exception('生成失败')
            self.q.put(('fatal', traceback.format_exc()))

    # ---------------- 队列轮询/界面刷新 ----------------
    def _poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind == 'status':
                    self._set_status(msg[1])
                elif kind == 'photo_done':
                    self.progress.step(1)
                    self._refresh_photo_tree()
                    self._refresh_preview()
                elif kind == 'identify_done':
                    self._set_busy(False)
                    self._set_status('识别完成, 请核对预览结果后点击「一键生成」')
                elif kind == 'generate_done':
                    self._set_busy(False)
                    out_path, m, merged, report = msg[1]
                    self.out_path = out_path
                    self.btn_open_file.config(state='normal')
                    self.btn_open_dir.config(state='normal')
                    warns = merged['warnings'] + \
                        [f'[{s}] {w}' for s, r in report.items() for w in r['warnings']]
                    self._show_warnings(warns)
                    self._set_status(f'已生成: {out_path}')
                    if warns:
                        messagebox.showinfo('生成完成',
                                            f'已生成:\n{out_path}\n\n'
                                            f'警告 {len(warns)} 条, 详见预览区下方')
                    else:
                        messagebox.showinfo('生成完成', f'已生成:\n{out_path}')
                elif kind == 'fatal':
                    self._set_busy(False)
                    self._set_status('处理失败, 详情见 paofen.log')
                    messagebox.showerror('错误', str(msg[1]))
        except queue.Empty:
            pass
        self.root.after(120, self._poll)

    def _refresh_photo_tree(self):
        self.tv_photos.delete(*self.tv_photos.get_children())
        for i, ph in enumerate(self.photos):
            res = self.results.get(ph['path'])
            name = os.path.basename(ph['path'])
            count = ''
            if res:
                shift = ph['override'] or res['shift'] or '未知'
                count = str(len(res['entries']))
                status = (f"{res['month'] or '?'}月, {len(res['entries'])} 条明细"
                          if res['entries'] else '未解析到条目')
                if res['warnings']:
                    status += f'(警告{len(res["warnings"])})'
            else:
                shift = ph['override'] or '自动'
                status = '未识别'
            self.tv_photos.insert('', 'end', values=(name, shift, count, status))

    def _refresh_preview(self):
        self.tv_entries.delete(*self.tv_entries.get_children())
        all_warns = []
        for i, ph in enumerate(self.photos):
            res = self.results.get(ph['path'])
            if not res:
                continue
            shift = ph['override'] or res['shift'] or '未知'
            for seq, e in enumerate(res['entries'], 1):
                self.tv_entries.insert('', 'end', values=(
                    f'#{i + 1}', shift, seq, f"{e['month'] or res['month']}.{e['day']}",
                    e['name_raw'], e['name'] or '未匹配', e['reason'],
                    f"{e['delta']:+g}"))
            for w in res['warnings']:
                all_warns.append(f"[{os.path.basename(ph['path'])}] {w}")
        self._show_warnings(all_warns)

    def _show_warnings(self, warns):
        self.txt_warn.config(state='normal')
        self.txt_warn.delete('1.0', 'end')
        if warns:
            self.txt_warn.insert('1.0', '\n'.join(warns))
        else:
            self.txt_warn.insert('1.0', '无警告')
        self.txt_warn.config(state='disabled')

    def _set_busy(self, busy, total=None):
        self.busy = busy
        state = 'disabled' if busy else 'normal'
        self.btn_generate.config(state=state)
        if total is not None:
            self.progress.config(maximum=total, value=0)
        if busy:
            self.btn_open_file.config(state='disabled')
            self.btn_open_dir.config(state='disabled')

    def _set_status(self, text):
        self.var_status.set(text)
        logger.info('%s', text)

    def _open_file(self):
        if self.out_path and os.path.isfile(self.out_path):
            os.startfile(self.out_path)

    def _open_folder(self):
        if self.out_path:
            folder = os.path.dirname(self.out_path) or HERE
            if os.path.isdir(folder):
                os.startfile(folder)


def main():
    root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()
    PaofenApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()

