# -*- coding: utf-8 -*-
"""手写体《考核明细》照片端到端验证。

标准答案 = 现场提供的甲班 15 条明细(日期/姓名/分值)。
覆盖: 正拍 / 旋转 90·180·270 / 缩放(模拟拍虚) / 模板名单 vs 报名表名单。
判分要点: **日期与分值决定填到哪个格子、填多少**, 是必须全对的字段;
姓名靠名单匹配(含模糊匹配); 事由仅是存档文本, 允许形近字差异。
运行: python _test_handwriting.py
"""
import os
import sys
import time
from collections import Counter

import cv2
import numpy as np

import detail_core
import excel_core

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, '_out')
PHOTO = os.path.join(BASE, '参考文件', '考核明细图片.jpg')          # 手写样例照片
TEMPLATE = os.path.join(BASE, '参考文件', '煤库作业区跑分绩效汇总X月（表样）.xlsx')

# 标准答案: (日期, 分值, 姓名)
TRUTH = [(2, 3, '叶么菊'), (3, 3, '李新明'), (6, 5, '靖豪'), (7, 3, '王雪琴'),
         (10, -5, '李新明'), (11, 3, '邓旭东'), (14, 3, '徐天鹏'),
         (15, 3, '曹启太'), (18, 3, '陈金荣'), (19, 3, '黄双望'),
         (22, 3, '向双梅'), (23, -3, '田忠山'), (26, -3, '田本宏'),
         (27, 5, '张俊峰'), (30, -3, '张正东')]

# 现场 7 月报名表名单(含陈金荣)
ROSTER_JULY = ['叶么菊', '李新明', '靖豪', '王雪琴', '邓旭东', '徐天鹏', '曹启太',
               '陈金荣', '黄双望', '向双梅', '田忠山', '张水生', '田本宏', '张俊峰',
               '张正东', '周应贵', '詹绿霞', '陈朝霞']


def transformed(path, k=0, scale=1.0):
    """生成旋转/缩放后的临时图(模拟横拍、倒拍、拍虚), 返回路径。"""
    img = detail_core.imread_cn(path)
    if k:
        img = np.ascontiguousarray(np.rot90(img, k=-k // 90))
    if scale != 1.0:
        img = cv2.resize(img, None, fx=scale, fy=scale,
                         interpolation=cv2.INTER_AREA)
    out = os.path.join(OUT_DIR, f'_hw_{k}_{int(scale * 100)}.jpg')
    cv2.imencode('.jpg', img, [cv2.IMWRITE_JPEG_QUALITY, 85])[1].tofile(out)
    return out


def measure(tag, path, roster):
    """识别并按多重集比对标准答案, 返回统计 dict。"""
    t0 = time.time()
    res = detail_core.parse_photo(path, {'甲班': list(roster)})
    dt = time.time() - t0
    got = [(e['day'], e['delta'], e['name'] or '') for e in res['entries']]
    exp = [tuple(t) for t in TRUTH]

    def hit(a, b):
        return sum((Counter(a) & Counter(b)).values())

    stats = {
        'tag': tag, 'sec': dt, 'n': len(got),
        'day': hit([g[0] for g in got], [e[0] for e in exp]),
        'delta': hit([g[1] for g in got], [e[1] for e in exp]),
        'name': hit([g[2] for g in got], [e[2] for e in exp]),
        'full': hit(got, exp),
        'miss': [e for e in exp if e not in got],
        'warnings': list(res['warnings']),
        'reasons': [(e['name_raw'], e['reason']) for e in res['entries']],
    }
    print(f"\n--- {tag} (自动方向 rotate={res['rotate']}, {dt:.1f}s) ---")
    print(f"  识别 {stats['n']}/{len(TRUTH)} 条 | 日期对 {stats['day']}/{len(TRUTH)}"
          f" | 分值对 {stats['delta']}/{len(TRUTH)}"
          f" | 姓名对 {stats['name']}/{len(TRUTH)}"
          f" | 三项全对 {stats['full']}/{len(TRUTH)}")
    if stats['miss']:
        print('  未对上的标准条目:', stats['miss'])
    for w in stats['warnings']:
        print('  警告:', w)
    return stats


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    ok = True
    tpl = excel_core.load_rosters(TEMPLATE)['甲班']

    # 1) 正拍 + 模板内名单: 日期/分值必须全对(姓名受模板名单限制)
    a = measure('正拍 · 模板内名单', PHOTO, tpl)
    if a['n'] != 15 or a['day'] != 15 or a['delta'] != 15:
        ok = False
        print('  [X] 手写正拍: 日期/分值应全部正确')
    # 2) 正拍 + 报名表(含陈金荣): 三项全对
    b = measure('正拍 · 报名表(含陈金荣)', PHOTO, ROSTER_JULY)
    if b['full'] != 15:
        ok = False
        print('  [X] 手写正拍+报名表: 应 15/15 全对')
    # 3) 横拍/倒拍: 自动方向识别, 允许极少数行丢失
    for k in (90, 180, 270):
        s = measure(f'旋转{k}度 · 报名表', transformed(PHOTO, k), ROSTER_JULY)
        if s['n'] < 14 or s['delta'] < 14:
            ok = False
            print(f'  [X] 旋转{k}度: 识别条数/分值过少')
    # 4) 拍虚(缩放): 允许极少数行丢失
    for scale in (0.5, 0.35):
        s = measure(f'缩放{int(scale * 100)}% · 报名表',
                    transformed(PHOTO, 0, scale), ROSTER_JULY)
        if s['n'] < 14 or s['delta'] < 14:
            ok = False
            print(f'  [X] 缩放{int(scale * 100)}%: 识别条数/分值过少')

    print('\n手写体识别测试:', '全部通过' if ok else '存在不一致!')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
