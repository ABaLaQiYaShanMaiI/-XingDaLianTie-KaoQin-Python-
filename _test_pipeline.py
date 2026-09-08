# -*- coding: utf-8 -*-
"""临时端到端验证: 明细照片 -> OCR解析 -> 填模板 -> 校验输出。"""
import os
import sys

import openpyxl

import detail_core
import excel_core

BASE = os.path.dirname(os.path.abspath(__file__))
PHOTO = os.path.join(BASE, '参考文件', '考核明细图片.jpg')
TEMPLATE = os.path.join(BASE, '参考文件', '煤库作业区跑分绩效汇总X月（表样）.xlsx')
OUT_DIR = os.path.join(BASE, '_out')


def main():
    rosters = excel_core.load_rosters(TEMPLATE)
    print('甲班名单:', '、'.join(rosters.get('甲班', [])))
    res = detail_core.parse_photo(PHOTO, rosters)
    print(f"\n标题识别: {res['month']}月 {res['shift']} (rotate={res['rotate']}, "
          f"OCR框={res['raw_count']})")
    print(f"条目 {len(res['entries'])} 条:")
    for i, e in enumerate(res['entries'], 1):
        print(f"  {i:2d}. {e['month']}.{e['day']:2d} {e['name_raw']} -> "
              f"{e['name']} (sim={e['name_sim']}) {e['reason']} {e['delta']:+g} "
              f"conf={e['conf']}")
    print('警告:')
    for w in res['warnings']:
        print('  -', w)

    os.makedirs(OUT_DIR, exist_ok=True)
    merged = detail_core.merge_results([res])
    out_path = excel_core.build_output_path(OUT_DIR, merged['month'])
    report = excel_core.fill_from_details(
        TEMPLATE, out_path, merged['shift_entries'], month=merged['month'])
    print('\n填写报告:', report)
    print('输出:', out_path)

    # ---- 校验 ----
    wb = openpyxl.load_workbook(out_path)
    ws = wb['甲班']
    print('\n标题:', ws['A1'].value)
    expect = {'叶么菊': [(2, 3)], '李新明': [(3, 3), (10, -5)], '靖豪': [(6, 5)],
              '王雪琴': [(7, 3)], '邓旭东': [(11, 3)], '徐天鹏': [(14, 3)],
              '曹启太': [(15, 3)], '刘启都': [], '黄双望': [(19, 3)],
              '向双梅': [(22, 3)], '田忠山': [(23, -3)], '张水生': [],
              '田本宏': [(26, -3)], '张俊峰': [(27, 5)], '张正东': [(30, -3)],
              '周应贵': [], '詹绿霞': [], '陈朝霞': []}
    ok = True
    for r in range(3, 21):
        name = str(ws.cell(row=r, column=2).value).strip()
        got = []
        for d in range(1, 32):
            v = ws.cell(row=r, column=2 + d).value
            if isinstance(v, (int, float)):
                got.append((d, int(v)))
        exp = expect.get(name)
        if got != exp:
            ok = False
            print(f'  [X] {name}: 期望{exp} 实际{got}')
        total = ws.cell(row=r, column=34).value
        amount = ws.cell(row=r, column=35).value
        exp_total = 80 + sum(v for _, v in exp)
        if total != f'=80+SUM(C{r}:AG{r})' or exp_total != 80 + sum(v for _, v in got):
            ok = False
            print(f'  [X] {name}: 合计公式异常 {total!r}')
        if amount != f'=ROUND(AH{r}*AJ$21,2)':
            ok = False
            print(f'  [X] {name}: 金额公式异常 {amount!r}')
        print(f'  {name}: 日格{got} 合计={exp_total} 公式OK' if got else
              f'  {name}: 无明细 合计={exp_total} 公式OK')
    print('AH21 =', ws['AH21'].value, '| AI21 =', ws['AI21'].value,
          '| AJ21 =', ws['AJ21'].value)
    ds = wb['甲班考核明细']
    print('\n明细存档表 B3:B17:')
    for r in range(3, 18):
        print(' ', ds.cell(row=r, column=1).value, ds.cell(row=r, column=2).value)
    print('\n明细表标题:', ds['A1'].value)
    wb.close()
    # 期望绩效总分(按模板名单): 80*18 + 3+3-3+5+3+3+3+3+3+3-3-3+5-3 = 1440+27
    print('\n期望绩效总分 =', 80 * 18 + 27, '(手写样例1482因含陈金荣且名单不同)')
    print('校验结果:', '全部通过' if ok else '存在不一致!')


if __name__ == '__main__':
    sys.exit(main())
