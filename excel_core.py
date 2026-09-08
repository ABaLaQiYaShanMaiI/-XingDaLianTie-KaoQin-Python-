# -*- coding: utf-8 -*-
"""
excel_core.py - 《煤库作业区跑分绩效汇总》模板读取与按考核明细填写

模板约定(与「煤库作业区跑分绩效汇总X月（表样）.xlsx」一致):
  * 工作表: 甲班 / 乙班 / 丙班 / 丁班 + 各班「X班考核明细」共 8 张
  * 考勤表: B2='姓名', C2='1' ... AG2='31'(C~AG 列=1~31日);
    AH列=合计总分, AI列=绩效金额, AJ列=签字确认(手签, 程序不填);
    第3行起为名单, B列出现「备注」即结束; 备注行右侧:
    AH=SUM(AH3:AH20) 绩效总分, AI=绩效总额(模板自带, 如 10200), AJ=AI21/AH21 每分价值
  * 计分规则(与现场确认一致): 每人基础分 80, 考核明细的加减分写入当日格,
    合计总分 = 基础分 + 当月各日加减分之和(Excel 公式实时计算);
    绩效金额 = 合计总分 x 每分价值(Excel 公式), 程序不代算数值
"""
import logging
import os
import re

import openpyxl

logger = logging.getLogger(__name__)

NAME_COL = 2          # B列: 姓名
DAY_COL_START = 3     # C列: 1日
MAX_DAY = 31
TOTAL_COL = 34        # AH列: 合计总分
AMOUNT_COL = 35       # AI列: 绩效金额
DATA_FIRST_ROW = 3
NAME_HEADER = '姓名'
BASE_SCORE_DEFAULT = 80


def load_rosters(template_path):
    """读取模板中所有班次表的名单。返回 {工作表名: [姓名, ...]}。"""
    wb = openpyxl.load_workbook(template_path, data_only=False)
    rosters = {}
    try:
        for ws in wb.worksheets:
            b2 = str(ws.cell(row=2, column=NAME_COL).value or '').strip()
            c2 = str(ws.cell(row=2, column=DAY_COL_START).value or '').strip()
            if b2 != NAME_HEADER or c2 != '1':
                continue
            names = []
            for r in range(DATA_FIRST_ROW, 60):
                v = ws.cell(row=r, column=NAME_COL).value
                if v is None:
                    continue
                s = str(v).strip()
                if not s:
                    continue
                if s.startswith('备注'):
                    break
                names.append(s)
            if names:
                rosters[ws.title] = names
    finally:
        wb.close()
    return rosters


def find_template(start_dir=None):
    """按 程序目录 -> 上级目录 -> 参考文件子目录 找「跑分绩效汇总」xlsx 模板。"""
    here = os.path.dirname(os.path.abspath(__file__))
    dirs = [here, os.path.dirname(here), os.path.join(here, '参考文件')]
    if start_dir:
        dirs.insert(0, start_dir)
    seen = set()
    for d in dirs:
        if not d or not os.path.isdir(d) or d in seen:
            continue
        seen.add(d)
        try:
            names = sorted(os.listdir(d))
        except OSError:
            continue
        for n in names:
            if n.lower().endswith('.xlsx') and '跑分绩效汇总' in n:
                return os.path.join(d, n)
    return None


def build_output_path(out_dir, month):
    """生成输出文件路径, 同名文件自动追加 _1/_2 防止覆盖。"""
    base = f'煤库作业区跑分绩效汇总{month}月.xlsx'
    path = os.path.join(out_dir, base)
    if not os.path.exists(path):
        return path
    stem = base[:-len('.xlsx')]
    i = 1
    while True:
        p = os.path.join(out_dir, f'{stem}_{i}.xlsx')
        if not os.path.exists(p):
            return p
        i += 1


def _fmt_num(v):
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _replace_month(text, month):
    """把标题文本中的「8月」/「X月」替换为目标月份。"""
    t = str(text)
    nt = re.sub(r'\d{1,2}\s*月', f'{month}月', t, count=1)
    if nt == t:
        nt = t.replace('X月', f'{month}月', 1)
    return nt


def _find_sheet(wb, title):
    if title in wb.sheetnames:
        return wb[title]
    for ws in wb.worksheets:  # 兜底: 工作表名包含班次名
        if title in str(ws.title):
            return ws
    return None


def _name_rows(ws):
    """{姓名: 行号} 名单区(B列, 遇「备注」止)。"""
    rows = {}
    for r in range(DATA_FIRST_ROW, 60):
        v = ws.cell(row=r, column=NAME_COL).value
        if v is None:
            continue
        s = str(v).strip()
        if not s:
            continue
        if s.startswith('备注'):
            break
        rows.setdefault(s, r)
    return rows



def _sum_row(ws):
    """定位备注行的「绩效总分」行(其 AH 单元格为 =SUM(...))。"""
    for r in range(DATA_FIRST_ROW, 60):
        v = ws.cell(row=r, column=TOTAL_COL).value
        if isinstance(v, str) and v.strip().upper().startswith('=SUM'):
            return r
    return None


def fill_from_details(template_path, out_path, shift_entries, month=None,
                      base_score=BASE_SCORE_DEFAULT, pool_amount=None,
                      write_detail_sheet=True):
    """按解析好的考核明细填写模板并另存。

    shift_entries: {班次表名(如'甲班'): [entry, ...]};
      entry 为 detail_core.parse_entry 的返回值(day/name/name_raw/reason/delta/month)
    month: 目标月份(用于标题与输出文件名); None=不改标题月份
    base_score: 每人基础分(默认 80), 合计公式 = 基础分+SUM(当月各日)
    pool_amount: 绩效总额; None=保持模板自带值, 由 Excel 公式计算每人金额
    write_detail_sheet: 是否把明细同步写入「X班考核明细」工作表存档

    返回 {班次表名: {'warnings': [str,...], 'cells': int}}
    """
    wb = openpyxl.load_workbook(template_path)
    report = {}
    try:
        for shift, entries in (shift_entries or {}).items():
            warns, cells = [], 0
            ws = _find_sheet(wb, shift)
            if ws is None:
                report[shift] = {'warnings': [f'工作簿中不存在工作表「{shift}」, 已跳过'],
                                 'cells': 0}
                continue
            row_by_name = _name_rows(ws)
            # 1) 当日格写入加减分(同人同日多次自动合并求和)
            dup_seen = {}
            for e in entries or []:
                name = e.get('name')
                r = row_by_name.get(str(name).strip()) if name else None
                day = e.get('day')
                if r is None:
                    if name:
                        warns.append(f'姓名「{e.get("name_raw") or name}」'
                                     f'不在{shift}名单, 未记分(已写入明细存档)')
                    continue
                if not isinstance(day, int) or not (1 <= day <= MAX_DAY):
                    warns.append(f'{name} 日期 {day} 无效, 未记分')
                    continue
                col = DAY_COL_START + day - 1
                old = ws.cell(row=r, column=col).value
                val = e.get('delta') or 0
                if isinstance(old, (int, float)):
                    val = old + val
                    key = (name, day)
                    dup_seen[key] = dup_seen.get(key, 0) + 1
                ws.cell(row=r, column=col,
                        value=int(val) if float(val).is_integer() else float(val))
                cells += 1
            for (name, day), n in dup_seen.items():
                warns.append(f'{name} {day}日有 {n + 1} 条明细, 已合并为单格求和')
            # 2) 全员合计总分/绩效金额走 Excel 公式(包括无明细的人=纯基础分)
            base = _fmt_num(base_score)
            for r in row_by_name.values():
                ws.cell(row=r, column=TOTAL_COL,
                        value=f'={base}+SUM(C{r}:AG{r})')
                ws.cell(row=r, column=AMOUNT_COL,
                        value=f'=ROUND(AH{r}*AJ$21,2)')
                cells += 2
            # 3) 绩效总额(可选): 写入备注行的 AI 单元格
            if pool_amount is not None:
                sr = _sum_row(ws)
                if sr is not None:
                    ws.cell(row=sr, column=AMOUNT_COL, value=pool_amount)
                else:
                    warns.append('未定位到绩效总额单元格, 绩效总额未修改')
            # 4) 标题月份
            if month:
                t = ws.cell(row=1, column=1).value
                if t:
                    ws.cell(row=1, column=1, value=_replace_month(t, month))
            # 5) 明细存档
            if write_detail_sheet:
                ds = _find_sheet(wb, f'{shift}考核明细')
                if ds is not None:
                    if month:
                        t = ds.cell(row=1, column=1).value
                        if t:
                            ds.cell(row=1, column=1, value=_replace_month(t, month))
                    max_rows = 34  # 模板明细表预留行数
                    for i, e in enumerate(entries or []):
                        if i >= max_rows:
                            warns.append(f'明细超过 {max_rows} 条, 存档只保留前 {max_rows} 条')
                            break
                        ds.cell(row=DATA_FIRST_ROW + i, column=1, value=i + 1)
                        ds.cell(row=DATA_FIRST_ROW + i, column=2,
                                value=detail_archive_text(e, month))
                        cells += 1
                else:
                    warns.append(f'未找到「{shift}考核明细」工作表, 跳过明细存档')
            report[shift] = {'warnings': warns, 'cells': cells}
        wb.save(out_path)
        logger.info('已生成: %s', out_path)
    finally:
        wb.close()
    return report


def detail_archive_text(entry, month=None):
    """存档明细文本, 如「7.10李新明两穿一带不规范-5」。"""
    from detail_core import detail_text
    return detail_text(entry, month)
