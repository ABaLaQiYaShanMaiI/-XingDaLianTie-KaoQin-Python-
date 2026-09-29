# -*- coding: utf-8 -*-
"""_test_name_split.py - 姓名槽 / 事由 边界回归测试

实盘反馈: 「7.18陈金荣接班发现卫生未做+3」被切成 姓名=「陈金荣接」+
事由=「班发现卫生未做」。名单里没有「陈金荣」时, 姓名匹配失败, 事由也少一个字。

期望行为:
  * 名单里有该人(如报名表含陈金荣) -> 姓名=陈金荣, 事由=接班发现卫生未做(原有能力不退化);
  * 名单里没有 -> 姓名槽自动退回成「陈金荣」, 事由补回「接班发现卫生未做」,
    只出「未匹配名单」警告(不记分)并提示导入该班报名表;
  * 复姓(欧阳/司马/…)4 字名不被误切;
  * 2 字姓名吞了 2 个字(「张三安排…」)也能退回。

运行: python _test_name_split.py   (退出码 0 = 全部通过)
"""
import sys

import detail_core

# (文本, 名单, 期望 name_raw, 期望 reason, 期望 name)
CASES = [
    ('7.18陈金荣接班发现卫生未做+3', ('叶么菊', '陈金荣'),
     '陈金荣', '接班发现卫生未做', '陈金荣'),
    ('7.18陈金荣接班发现卫生未做+3', ('叶么菊', '李新明'),
     '陈金荣', '接班发现卫生未做', None),
    ('7.18陈金荣接班发现卫生未做+3', (),
     '陈金荣', '接班发现卫生未做', None),
    ('7.9张三安排临时性工作+5', ('叶么菊',), '张三', '安排临时性工作', None),
    ('7.26田本宏未按要求点检-3', ('叶么菊',), '田本宏', '未按要求点检', None),
    ('7.10李新明两穿一带不规范-5', ('李新明',),
     '李新明', '两穿一带不规范', '李新明'),
    ('7.9欧阳未做卫生-3', ('叶么菊',), '欧阳未做', '卫生', None),
    ('7.9欧阳靖修设备+3', ('欧阳靖',), '欧阳靖', '修设备', '欧阳靖'),
]


def main():
    ok = True
    for text, roster, want_raw, want_reason, want_name in CASES:
        e = detail_core.parse_entry(text, roster)
        if e is None:
            ok = False
            print(f'[X] 解析失败: {text}')
            continue
        good = (e['name_raw'] == want_raw and e['reason'] == want_reason
                and e['name'] == want_name)
        ok = ok and good
        print(f"[{'OK' if good else 'X'}] {text}")
        print(f'     姓名槽={e["name_raw"]!r} 事由={e["reason"]!r} '
              f'匹配={e["name"]!r} 相似度={e["name_sim"]}')
        if not good:
            print(f'     期望 姓名槽={want_raw!r} 事由={want_reason!r} '
                  f'匹配={want_name!r}')
    # 归档文本要能原样拼回来(明细存档 = 日期 + 姓名槽 + 事由 + 分值)
    entries, warns = detail_core.parse_detail_lines(
        '7.18陈金荣接班发现卫生未做+3', ('叶么菊',))
    e = entries[0]
    merged = f"{e['month']}.{e['day']}{e['name_raw']}{e['reason']}{e['score_raw']}"
    good = merged == '7.18陈金荣接班发现卫生未做+3'
    ok = ok and good
    print(f"[{'OK' if good else 'X'}] 存档文本拼回: {merged}")
    print('     警告:', warns or '(无)')
    print('\n姓名槽/事由边界测试:', '全部通过' if ok else '存在不一致!')
    return 0 if ok else 1


if __name__ == '__main__':
    sys.exit(main())
