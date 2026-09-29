# -*- coding: utf-8 -*-
"""报名表(名单) + 手工明细文本 端到端验证。

  1) 用文字明细(甲班 15 条, 与现场样例一致)解析 -> 校验条数与姓名匹配;
  2) 生成几种「报名表」xlsx(名单变动/人数 18、20、15) -> 校验名单读取;
  3) 按报名表重写考勤表名单并填分 -> 逐格校验(姓名/序号/日格/公式/备注行位置)。
运行: python _test_roster.py
"""
import os
import sys

import openpyxl

import detail_core
import excel_core

BASE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE, '_out')
TEMPLATE = os.path.join(BASE, '参考文件', '煤库作业区跑分绩效汇总X月（表样）.xlsx')

# 甲班 7 月考核明细(现场提供, 带行首序号)
DETAIL_TEXT = '\n'.join([
    '1、7.2叶么菊发现现场安全隐患+3',
    '2、7.3李新明疏通下料斗+3',
    '3、7.6靖豪安排零时性工作+5',
    '4、7.7王雪琴接班发现卫生未做+3',
    '5、7.10李新明两穿一带不规范-5',
    '6、7.11邓旭东更换岗位拉绳+3',
    '7、7.14徐天鹏疏通下料斗+3',
    '8、7.15曹启太发现岗位清扫器故障+3',
    '9、7.18陈金荣接班发现卫生未做+3',
    '10、7.19黄双望发现设备异常+3',
    '11、7.22向双梅更换拉绳+3',
    '12、7.23田忠山两穿一带不规范-3',
    '13、7.26田本宏未按要求点检-3',
    '14、7.27张俊峰安排临时性工作+5',
    '15、7.30张正东班前会迟到-3',
])

# 现场 7 月实际名单(考勤表照片): 陈金荣 替换了模板里的 刘启都
ROSTER_JULY = ['叶么菊', '李新明', '靖豪', '王雪琴', '邓旭东', '徐天鹏', '曹启太',
               '陈金荣', '黄双望', '向双梅', '田忠山', '张水生', '田本宏', '张俊峰',
               '张正东', '周应贵', '詹绿霞', '陈朝霞']

EXPECT_ROWS = {
    '叶么菊': [(2, 3)], '李新明': [(3, 3), (10, -5)], '靖豪': [(6, 5)],
    '王雪琴': [(7, 3)], '邓旭东': [(11, 3)], '徐天鹏': [(14, 3)],
    '曹启太': [(15, 3)], '陈金荣': [(18, 3)], '黄双望': [(19, 3)],
    '向双梅': [(22, 3)], '田忠山': [(23, -3)], '田本宏': [(26, -3)],
    '张俊峰': [(27, 5)], '张正东': [(30, -3)],
}


def make_roster(path, names, sheet, header=True):
    """生成一张「报名表/花名册」(序号 | 姓名 | 岗位)。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    if header:
        ws.cell(row=1, column=1, value='序号')
        ws.cell(row=1, column=2, value='姓名')
        ws.cell(row=1, column=3, value='岗位')
    for i, n in enumerate(names, 1):
        ws.cell(row=i + 1, column=1, value=i)
        ws.cell(row=i + 1, column=2, value=n)
        ws.cell(row=i + 1, column=3, value='皮带工')
    wb.save(path)
    wb.close()
    return path


def day_cells(ws, r):
    return [(d, int(ws.cell(row=r, column=2 + d).value))
            for d in range(1, 32)
            if isinstance(ws.cell(row=r, column=2 + d).value, (int, float))]


def row_of(ws, name):
    for r in range(excel_core.DATA_FIRST_ROW, 60):
        if str(ws.cell(row=r, column=2).value or '').strip() == name:
            return r
    return None


def check_case(tag, names, detail_entries, base_score=80, pool=10800):
    """按报名表 names 填表并逐格校验, 返回是否通过。"""
    ok = True
    out_path = os.path.join(OUT_DIR, f'_roster_{tag}.xlsx')
    report = excel_core.fill_from_details(
        TEMPLATE, out_path, {'甲班': detail_entries}, month=7,
        base_score=base_score, pool_amount=pool,
        roster_override={'甲班': list(names)})
    n = len(names)
    sum_row = excel_core.DATA_FIRST_ROW + n      # 备注行新位置
    print(f'\n===== {tag}: {n} 人 (备注行应为 {sum_row}) =====')
    for w in report['甲班']['warnings']:
        print('  填表提示:', w)
    wb = openpyxl.load_workbook(out_path)
    try:
        ws = wb['甲班']
        for i, want in enumerate(names):         # 序号/姓名
            r = excel_core.DATA_FIRST_ROW + i
            got_name = str(ws.cell(row=r, column=2).value or '').strip()
            got_no = ws.cell(row=r, column=1).value
            if got_name != want or got_no != i + 1:
                ok = False
                print(f'  [X] 第{r}行 序号/姓名异常: {got_no}/{got_name} 期望{i + 1}/{want}')
        for name, exp in EXPECT_ROWS.items():    # 日格(姓名+日期定位)
            r = row_of(ws, name)
            if r is None:
                if name in names:
                    ok = False
                    print(f'  [X] 名单里的 {name} 未找到行')
                continue
            got = day_cells(ws, r)
            if got != exp:
                ok = False
                print(f'  [X] {name} 日格 {got} != 期望 {exp}')
        for i in range(n):                       # 公式
            r = excel_core.DATA_FIRST_ROW + i
            if ws.cell(row=r, column=34).value != f'={base_score}+SUM(C{r}:AG{r})':
                ok = False
                print(f'  [X] 第{r}行 AH 公式异常: {ws.cell(row=r, column=34).value}')
            want = f'=ROUND(AH{r}*AJ${sum_row},2)'
            if ws.cell(row=r, column=35).value != want:
                ok = False
                print(f'  [X] 第{r}行 AI 公式异常: {ws.cell(row=r, column=35).value}')
        last = sum_row - 1                       # 备注行
        for col, want in ((34, f'=SUM(AH3:AH{last})'), (35, pool),
                          (36, f'=AI{sum_row}/AH{sum_row}')):
            got = ws.cell(row=sum_row, column=col).value
            if got != want:
                ok = False
                print(f'  [X] 备注行第{col}列: {got!r} != {want!r}')
        if '备注' not in str(ws.cell(row=sum_row, column=1).value or ''):
            ok = False
            print('  [X] 备注行文本丢失')
        labels = [ws.cell(row=sum_row + 1, column=c).value for c in (34, 35, 36)]
        if labels != ['绩效总分', '绩效总额', '每分价值']:
            ok = False
            print('  [X] 标签行异常:', labels)
        if n > 18:                               # 新增行应有边框样式
            if ws.cell(row=excel_core.DATA_FIRST_ROW + n - 1,
                       column=3).border.left.style != 'thin':
                ok = False
                print('  [X] 新增行缺少边框样式')
        ds = wb['甲班考核明细']                  # 明细存档
        n_archived = sum(1 for r in range(3, 40) if ds.cell(row=r, column=2).value)
        if n_archived != len(detail_entries):
            ok = False
            print(f'  [X] 明细存档 {n_archived} 条 != {len(detail_entries)} 条')
        print(f'  输出: {os.path.basename(out_path)} | 存档 {n_archived} 条 | '
              f'备注行 AH={ws.cell(row=sum_row, column=34).value} '
              f'AI={ws.cell(row=sum_row, column=35).value} '
              f'AJ={ws.cell(row=sum_row, column=36).value}')
    finally:
        wb.close()
    return ok


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    ok = True
    # ---- 1) 模板内名单解析明细文本 ----
    tpl_rosters = excel_core.load_rosters(TEMPLATE)
    print(f"模板甲班名单({len(tpl_rosters['甲班'])}人):", '、'.join(tpl_rosters['甲班']))
    entries_tpl, warns = detail_core.parse_detail_lines(
        DETAIL_TEXT, tuple(tpl_rosters['甲班']), month=7)
    print(f'\n[文字明细·模板名单] {len(entries_tpl)} 条, 警告 {len(warns)} 条')
    for w in warns:
        print('  -', w)
    if len(entries_tpl) != 15 or len(warns) != 1:
        ok = False
        print('  [X] 期望 15 条明细 + 1 条未匹配(陈金荣)警告')
    if next(e for e in entries_tpl if e['name_raw'].startswith('陈金荣'))['name'] is not None:
        ok = False
        print('  [X] 陈金荣 不应匹配到模板名单')

    # ---- 2) 报名表读取(3 种版式/人数) ----
    r18 = make_roster(os.path.join(OUT_DIR, '_roster18.xlsx'), ROSTER_JULY, '甲班名单')
    r20 = make_roster(os.path.join(OUT_DIR, '_roster20.xlsx'),
                      ROSTER_JULY + ['何建军', '吴建兵'], '甲班')
    r15 = make_roster(os.path.join(OUT_DIR, '_roster15.xlsx'),
                      ROSTER_JULY[:15], '花名册', header=False)
    got18, sheet18 = excel_core.load_roster_file(r18, '甲班')
    got20, sheet20 = excel_core.load_roster_file(r20, '甲班')
    got15, sheet15 = excel_core.load_roster_file(r15, '甲班')
    print(f'\n[报名表读取] 18人表->{sheet18}: {len(got18)}人 一致={got18 == ROSTER_JULY}'
          f' | 20人表->{sheet20}: {len(got20)}人'
          f' | 无表头单列->{sheet15}: {len(got15)}人 一致={got15 == ROSTER_JULY[:15]}')
    if got18 != ROSTER_JULY or len(got20) != 20 or got15 != ROSTER_JULY[:15]:
        ok = False
        print('  [X] 报名表姓名读取不一致')

    # ---- 3) 用报名表名单解析 + 填表 ----
    entries_july, warns2 = detail_core.parse_detail_lines(
        DETAIL_TEXT, tuple(ROSTER_JULY), month=7)
    print(f'\n[文字明细·7月报名表] {len(entries_july)} 条, 警告 {len(warns2)} 条')
    for w in warns2:
        print('  -', w)
    if len(entries_july) != 15 or warns2:
        ok = False
        print('  [X] 期望 15 条明细且无警告(陈金荣已在报名表内)')
    if next(e for e in entries_july
            if e['name_raw'].startswith('陈金荣'))['name'] != '陈金荣':
        ok = False
        print('  [X] 陈金荣 应匹配报名表名单')

    ok = check_case('7月18人(含陈金荣)', ROSTER_JULY, entries_july) and ok
    ok = check_case('20人(增2人)', ROSTER_JULY + ['何建军', '吴建兵'], entries_july) and ok
    ok = check_case('15人(减3人)', ROSTER_JULY[:15], entries_july) and ok

    print('\n报名表/明细文本测试:', '全部通过' if ok else '存在不一致!')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
