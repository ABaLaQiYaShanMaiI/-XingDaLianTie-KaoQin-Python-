# -*- coding: utf-8 -*-
"""GUI 冒烟 + 全流程测试: 建界面 -> 添加照片 -> 同步识别 -> 同步生成 -> 校验。"""
import os
import queue
import sys

import app as app_mod

BASE = os.path.dirname(os.path.abspath(__file__))
PHOTO = os.path.join(BASE, '参考文件', '考核明细图片.jpg')
OUT_DIR = os.path.join(BASE, '_out')


def drain(q):
    msgs = []
    try:
        while True:
            msgs.append(q.get_nowait())
    except queue.Empty:
        pass
    return msgs


def main():
    root = app_mod.TkinterDnD.Tk() if app_mod.TkinterDnD else None
    if root is None:
        import tkinter as tk
        root = tk.Tk()
    gui = app_mod.PaofenApp(root)
    root.update()

    # 配置
    gui.var_template.set(os.path.join(BASE, '参考文件',
                                      '煤库作业区跑分绩效汇总X月（表样）.xlsx'))
    gui.var_out.set(OUT_DIR)
    gui.var_month.set('')
    gui.var_pool.set('')
    gui._add_photos([PHOTO])
    assert len(gui.photos) == 1, '照片添加失败'

    # 同步识别(直接调工作函数, 等价于后台线程执行体)
    gui._identify_worker()
    msgs = drain(gui.q)
    res = gui.results[PHOTO]
    print('识别: month=%s shift=%s 条目=%d 警告=%d'
          % (res['month'], res['shift'], len(res['entries']), len(res['warnings'])))
    assert res['month'] == 7 and res['shift'] == '甲班'
    assert len(res['entries']) == 15, f'条目数异常: {len(res["entries"])}'
    root.update()
    gui._refresh_photo_tree()
    gui._refresh_preview()
    assert len(gui.tv_entries.get_children()) == 15, '预览表行数异常'

    # 同步生成(屏蔽弹窗避免测试阻塞), 用 _poll 消化队列并更新界面
    app_mod.messagebox.showinfo = lambda *a, **k: None
    app_mod.messagebox.showerror = lambda *a, **k: None
    gui._generate_worker(gui.var_template.get(), [], 80, None, None)
    gui._poll()
    root.update()
    out_path = gui.out_path
    print('生成:', out_path)
    assert out_path and os.path.isfile(out_path), '输出文件不存在'
    assert '7月' in os.path.basename(out_path), '文件名月份异常'

    assert str(gui.btn_open_file['state']) == 'normal'
    assert str(gui.btn_open_dir['state']) == 'normal'
    root.destroy()
    print('\nGUI 全流程测试通过')


if __name__ == '__main__':
    sys.exit(main())
