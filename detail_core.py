# -*- coding: utf-8 -*-
"""
detail_core.py - 《考核明细》照片 OCR 解析

输入: 手机拍摄的《X月X班考核明细》单据照片(参考 参考文件/考核明细图片.jpg)
输出: 结构化考核条目, 供 excel_core.fill_from_details 填写《煤库作业区跑分绩效汇总》

单据约定(与现场手写单一致):
  * 标题行: 「7月甲班考核明细」 -> 月份=7, 班次=甲
  * 条目行: 「7.10李新明两穿一带不规范-5」 -> 日期=7.10, 姓名, 事由, 分值=-5
  * 分值兼容: +3 / 3 / -3 / 3分; 日期兼容: 7.10 / 7月10日 / 22(无月份,取标题月)
  * 姓名与模板名单模糊匹配(处理 OCR 误差), 匹配不上保留原名并给出警告

实现要点:
  * 逐个 OCR 文本框解析(不按整行拼接), 避免左边"序号"数字(10/12/...)混入条目;
  * 自动识别拍照方向(0/90/180/270), 原方向能识别出标题+条目时不再旋转;
  * RapidOCR 引擎惰性加载(首次识别才初始化, GUI 启动不卡)。
"""
import logging
import os
import re
from difflib import SequenceMatcher

import cv2
import numpy as np

logger = logging.getLogger(__name__)

TITLE_KEY = '考核明细'
SHIFTS = ('甲', '乙', '丙', '丁')
SHIFT_NAMES = tuple(s + '班' for s in SHIFTS)

# 标题: 「7月甲班考核明细」(兼容空格/年月混杂)
TITLE_RE = re.compile(r'(?P<month>\d{1,2})\s*月.*?(?P<shift>[甲乙丙丁])\s*班')
# 条目日期: 「7.10」/「7月10」/「22」(无月份), 后接其余文本
DATE_RE = re.compile(
    r'^(?:(?P<month>\d{1,2})[.．,，、月](?P<day>\d{1,2})|(?P<day_only>\d{1,2}))[日号]?(?P<rest>.*)$')
# 姓名槽: 紧跟日期的 2~4 个汉字(可能吞入事由首字, 由 match_name 前缀收敛)
NAME_RE = re.compile(r'^[\u4e00-\u9fa5]{2,4}')
# 行尾分值: 「+3」/「-5」/「3」/「3分」; 兼容 OCR 把符号读到数字右侧(「3-」按 -3 处理)
SCORE_RE = re.compile(
    r'(?:(?P<pre>[+\-])\s*)?(?P<num>\d+(?:\.\d+)?)\s*分?\s*(?P<post>[+\-])?$')
# 打印体表格行内噪声: 最左侧「序号」列碎片(纯 1~3 位数字) 与 日期碎片开头
SEQ_NO_RE = re.compile(r'^\d{1,3}$')
DATE_HEAD_RE = re.compile(r'^\d{1,2}[.．,，、月]\d{1,2}')

# 全角/标点归一化
_NORMAL_TABLE = str.maketrans({
    '　': '', ' ': '', '．': '.', '，': '', '。': '', '、': '',
    '＋': '+', '－': '-', '—': '-', '–': '-', '～': '', '·': '',
})

_ENGINE = None


def get_engine():
    """惰性加载 RapidOCR 引擎(进程内单例)。"""
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR
        logger.info('初始化 RapidOCR 引擎...')
        _ENGINE = RapidOCR()
    return _ENGINE


def imread_cn(path):
    """支持中文路径读图(粗校验文件存在)。"""
    if not os.path.isfile(path):
        raise FileNotFoundError(f'找不到照片: {path}')
    data = np.fromfile(path, dtype=np.uint8)
    img = cv2.imdecode(data, cv2.IMREAD_COLOR)
    if img is None:
        raise ValueError(f'图片无法解码: {path}')
    return img


def _rot(img, k):
    """k 为 90 的倍数(0/90/180/270), 顺时针旋转。"""
    k = int(k) % 360
    if k == 0:
        return img
    return np.ascontiguousarray(np.rot90(img, k=-k // 90))


def normalize(text):
    """OCR 文本归一化: 去空白、全角转半角。"""
    return str(text).translate(_NORMAL_TABLE)



def match_name(slot, roster):
    """把 OCR 得到的姓名槽匹配到模板名单。

    返回 (matched_name 或 None, 相似度)。策略:
      1) 前缀收敛精确匹配: 槽位按 4->2 字截短, 命中名单即成功(处理吞字, 如
         「田忠山两穿一带」-> 槽位「田忠山两」-> 截短为「田忠山」);
      2) difflib 模糊匹配(阈值 0.45, 处理 OCR 形近字, 如 曹启太/曹启大)。
    """
    if not slot or not roster:
        return None, 0.0
    candidates = {slot[:k] for k in range(min(len(slot), 4), 1, -1)}
    for cand in candidates:
        if cand in roster:
            return cand, 1.0
    best, best_ratio = None, 0.0
    for cand in roster:
        for part in candidates:
            ratio = SequenceMatcher(None, part, cand).ratio()
            if ratio > best_ratio:
                best, best_ratio = cand, ratio
    if best_ratio >= 0.45:
        return best, best_ratio
    return None, best_ratio


# 复姓(4 字名多为「复姓+名」, 退回事由首字时要放过)
COMPOUND_SURNAMES = (
    '欧阳', '太史', '端木', '上官', '司马', '东方', '独孤', '南宫', '万俟', '闻人',
    '夏侯', '诸葛', '尉迟', '公羊', '赫连', '澹台', '皇甫', '宗政', '濮阳', '公冶',
    '太叔', '申屠', '公孙', '慕容', '仲孙', '钟离', '长孙', '宇文', '司徒', '鲜于',
    '司空', '闾丘', '子车', '亓官', '司寇', '巫马', '公西', '颛孙', '壤驷', '公良',
    '漆雕', '乐正', '宰父', '谷梁', '拓跋', '夹谷', '轩辕', '令狐', '段干', '百里',
    '呼延', '东郭', '南门', '羊舌', '微生', '梁丘', '左丘', '西门', '第五', '公乘',
    '南荣', '东里', '仲长', '即墨', '达奚', '褚师',
)

# 事由常见开头词: 用来判断「姓名槽的最后 1~2 个字其实属于事由」
REASON_WORDS = (
    '接班', '交班', '上班', '下班', '班前', '班中', '两穿', '一带', '三穿', '两穿一带',
    '发现', '未', '不', '迟到', '早退', '更衣', '清扫', '清理', '疏通', '更换', '安排',
    '岗位', '设备', '安全', '隐患', '记录', '规范', '违规', '漏', '点名', '交接', '卫生',
    '现场', '临时', '主动', '完成', '检查', '点检', '工作', '作业', '区域', '人员', '管理',
    '请假', '离岗', '串岗', '睡岗', '玩手机', '吸烟', '着装', '劳保', '标识', '整改', '考核',
)


def _trim_name_slot(entry):
    """判断姓名槽是否吞了事由首字, 是则退回(最多退 2 个字)。返回是否退过。

    保守策略: 只有「退回后事由以常见事由词开头」且「不是复姓」才退, 避免误切 4 字姓名。
    """
    slot = entry.get('name_raw') or ''
    reason = entry.get('reason') or ''
    if len(slot) <= 2:
        return False
    for k in (1, 2):
        if len(slot) - k < 2:
            break
        name = slot[:len(slot) - k]
        extra = slot[len(slot) - k:]
        if name[:2] in COMPOUND_SURNAMES:
            break                        # 复姓: 不做进退
        merged = extra + reason
        if any(merged.startswith(w) for w in REASON_WORDS):
            entry['name_raw'] = name
            entry['reason'] = merged
            entry['slot_trimmed'] = extra
            return True
    return False


def _apply_roster(entry, roster):
    """用班次名单对 entry 做姓名匹配, 并回收姓名槽吞掉的事由首字(幂等)。

    姓名槽是按「日期后紧跟的 2~4 个汉字」取的, 难免吞进事由首字
    (「陈金荣接」+「班发现卫生未做」)。处理顺序:
      1) 先按原槽匹配名单 —— 名单里真有 4 字名(复姓)或 3 字名都能命中;
      2) 匹配不上时判断是否吞了事由首字(见 _trim_name_slot), 退回后再匹配一次;
      3) 命中则把槽里多出来的字并回事由。
    """
    slot = entry.get('name_raw') or ''
    matched, sim = match_name(slot, roster)
    if not matched and _trim_name_slot(entry):
        slot = entry['name_raw']
        matched, sim = match_name(slot, roster)
    if matched and slot.startswith(matched) and len(slot) > len(matched):
        entry['reason'] = slot[len(matched):] + (entry.get('reason') or '')
        entry['name_raw'] = matched
    entry['name'] = matched
    entry['name_sim'] = round(float(sim), 2)
    return entry


def parse_entry(text, roster):
    """解析单条考核明细文本 -> dict; 不是条目返回 None。

    dict 字段: month/day/name_raw/name/name_sim/reason/delta
    """
    t = normalize(text)
    if not t or TITLE_KEY in t:
        return None
    m = DATE_RE.match(t)
    if not m:
        return None
    if m.group('month') is not None:
        month, day = int(m.group('month')), int(m.group('day'))
    else:
        month, day = None, int(m.group('day_only'))
    rest = m.group('rest') or ''
    nm = NAME_RE.match(rest)
    if not nm:
        return None
    slot = nm.group(0)
    tail = rest[nm.end():]
    sm = SCORE_RE.search(tail)
    if not sm:
        return None
    sign = sm.group('pre') or sm.group('post') or ''
    delta = float(sm.group('num'))
    if sign == '-':
        delta = -delta
    entry = {
        'month': month,
        'day': day,
        'name_raw': slot,
        'name': None,
        'name_sim': 0.0,
        'reason': tail[:sm.start()],
        'score_raw': normalize(sm.group(0)),
        'score_post': bool(sm.group('post') and not sm.group('pre')),
        'delta': int(delta) if float(delta).is_integer() else delta,
    }
    return _apply_roster(entry, roster)


def _entry_key(e):
    return (e.get('day'), e.get('name_raw'), e.get('delta'))


def _drop_seq_no(row):
    """去掉打印体表格行最左侧的「序号」列碎片(纯数字且右侧紧跟日期碎片)。

    分列式打印表格中, OCR 可能把「序号」列单独识别成一个框(如「10」「11」),
    按行拼接时会与日期粘连成「107.19…」导致整行解析失败。
    """
    if len(row) >= 2 and SEQ_NO_RE.match(normalize(row[0]['text'])) \
            and DATE_HEAD_RE.match(normalize(row[1]['text'])):
        return row[1:]
    return row


def _parse_orientation(engine, img, roster):
    """对给定方向的图像做 OCR 并解析, 返回结果 dict。"""
    result, _ = engine(img)
    items = []
    for box, text, score in (result or []):
        ys = [float(p[1]) for p in box]
        xs = [float(p[0]) for p in box]
        items.append({
            'text': str(text),
            'conf': float(score),
            'yc': sum(ys) / 4.0,
            'xc': sum(xs) / 4.0,
            'h': max(ys) - min(ys),
        })
    # 标题(取置信度最高的命中)
    month = shift = None
    best_title = None
    for it in items:
        t = normalize(it['text'])
        if TITLE_KEY in t:
            m = TITLE_RE.search(t)
            if m and (best_title is None or it['conf'] > best_title[0]):
                best_title = (it['conf'], int(m.group('month')), m.group('shift'))
    if best_title:
        month, shift = best_title[1], best_title[2] + '班'
    # 条目: 逐个 OCR 框解析(序号等噪声框不符合条目格式, 自然过滤)
    entries, used_ids = [], set()
    med_h = float(np.median([it['h'] for it in items])) if items else 30.0
    tol = max(14.0, med_h * 0.6)
    for idx, it in enumerate(items):
        e = parse_entry(it['text'], roster)
        if e is None:
            continue
        e['conf'] = round(it['conf'], 2)
        e['_yc'] = it['yc']
        used_ids.add(idx)
        entries.append(e)
    # 兜底: 未解析的碎片按 y 聚成行、按 x 拼接后再试(处理 OCR 拆框)
    left = [it for i, it in enumerate(items) if i not in used_ids]
    left.sort(key=lambda it: it['yc'])
    rows = []
    for it in left:
        if rows and abs(it['yc'] - rows[-1][-1]['yc']) <= tol:
            rows[-1].append(it)
        else:
            rows.append([it])
    for row in rows:
        if any(abs(row[0]['yc'] - e['_yc']) <= tol for e in entries):
            continue  # 该行主体已按整框解析过
        row.sort(key=lambda it: it['xc'])
        joined = ''.join(it['text'] for it in _drop_seq_no(row))
        e = parse_entry(joined, roster)
        if e is None:
            # OCR 把「序号」与日期并成一框(如「107.19…」)时, 去掉行首多余数字重试
            stripped = re.sub(r'^\d{1,3}(?=\d{1,2}[.．,，、月]\d{1,2})', '',
                              normalize(joined))
            if stripped != normalize(joined):
                e = parse_entry(stripped, roster)
        if e is not None:
            e['conf'] = round(min(it['conf'] for it in row), 2)
            e['_yc'] = sum(it['yc'] for it in row) / len(row)
            entries.append(e)
    # 去重(同行同键)并按版面顺序排序
    uniq = []
    for e in sorted(entries, key=lambda x: x['_yc']):
        dup = any(abs(e['_yc'] - u['_yc']) <= tol and _entry_key(e) == _entry_key(u)
                  for u in uniq)
        if not dup:
            uniq.append(e)
    for e in uniq:
        e.pop('_yc', None)
    return {'month': month, 'shift': shift, 'entries': uniq, 'items': items}


def _score_orientation(res):
    return (len(res['entries']), sum(e['conf'] for e in res['entries']))



def parse_photo(path, rosters, shift_override=None):
    """识别一张考核明细照片。

    rosters: {班次表名: [姓名,...]}(excel_core.load_rosters 的返回值)
    shift_override: 用户在界面指定的班次(如 '甲班'); None=按标题自动识别
    返回: {'path', 'month', 'shift', 'entries', 'warnings', 'rotate', 'raw_count'}
    """
    warnings = []
    img = imread_cn(path)
    engine = get_engine()
    best = None
    for k in (0, 90, 180, 270):
        res = _parse_orientation(engine, _rot(img, k), ())
        res['rotate'] = k
        if best is None or _score_orientation(res) > _score_orientation(best):
            best = res
        if res['month'] is not None and res['shift'] is not None \
                and len(res['entries']) > 0:
            break  # 标题+条目齐全, 不必再旋转
        if k == 0 and len(res['entries']) == 0:
            warnings.append('原方向未解析到条目, 已自动尝试旋转识别')
    res = best
    if res['month'] is None:
        warnings.append('未识别到标题「X月X班考核明细」, 请在界面核对/指定班次与月份')
    if not res['entries']:
        warnings.append('未解析到任何考核条目, 请确认照片内容与清晰度')
    # 用最终确定的班次名单匹配姓名
    shift = shift_override or res['shift']
    roster = tuple(rosters.get(shift or '', ()))
    if shift and not roster:
        warnings.append(f'模板中不存在班次「{shift}」的名单, 姓名未能匹配')
    for e in res['entries']:
        _apply_roster(e, roster)
        if e['month'] is not None and res['month'] is not None \
                and e['month'] != res['month']:
            warnings.append(
                f"条目 {e['month']}.{e['day']} 月份与标题({res['month']}月)不一致, 请核对")
        if not (1 <= e['day'] <= 31):
            warnings.append(f"条目日期 {e['day']} 超出 1~31, 请人工核对")
        if e.get('score_post'):
            warnings.append(f"分值「{e.get('score_raw')}」符号在数字右侧(识别倒序), "
                            f"已按 {e['delta']:+g} 处理, 请核对")
        if e['name'] is None:
            tip = ''
            if e.get('slot_trimmed'):
                tip = (f"姓名槽里的「{e['slot_trimmed']}」已并回事由"
                       f"(事由按「{e['reason']}」存档); ")
            warnings.append(f"姓名「{e['name_raw']}」未匹配到{shift or '模板'}名单"
                            f'(相似度 {e["name_sim"]}), 只写入明细存档不记分; '
                            f'{tip}若他确实在{shift or "该班"}, 请在「名单(报名表)」'
                            f'页签导入该班报名表(导入后姓名即可匹配)')
    return {
        'path': path,
        'month': res['month'],
        'shift': shift,
        'entries': res['entries'],
        'warnings': warnings,
        'rotate': res['rotate'],
        'raw_count': len(res['items']),
    }


def merge_results(results, default_month=None):
    """合并多张照片的解析结果 -> {'month', 'shift_entries', 'warnings'}。"""
    warnings = []
    month = None
    for r in results:
        if r['month'] is not None:
            if month is None:
                month = r['month']
            elif month != r['month']:
                warnings.append(
                    f"照片「{os.path.basename(r['path'])}」月份({r['month']}月)"
                    f'与先前({month}月)不一致, 以{month}月为准')
    if month is None:
        month = default_month
    shift_entries = {}
    for r in results:
        shift = r['shift']
        if not shift:
            warnings.append(
                f"照片「{os.path.basename(r['path'])}」班次未知, 未纳入生成"
                f'(请为其指定班次后重新识别)')
            continue
        shift_entries.setdefault(shift, []).extend(r['entries'])
        for w in r['warnings']:
            warnings.append(f"[{os.path.basename(r['path'])}] {w}")
    return {'month': month, 'shift_entries': shift_entries, 'warnings': warnings}


def detail_text(entry, month=None):
    """还原成存档用的一条明细文本, 如「7.10李新明两穿一带不规范-5」。"""
    m = entry.get('month') or month
    name = entry.get('name_raw') or entry.get('name') or ''
    d = entry['delta']
    if float(d).is_integer():
        d = int(d)
    score = ('+' if d >= 0 else '') + str(d)
    return f'{m}.{entry["day"]}{name}{entry.get("reason") or ""}{score}'


# 手工录入/粘贴明细时的行首序号: 「1、」「1.」「1)」等
SEQ_PREFIX_RE = re.compile(r'^\s*\d{1,3}\s*[、.,，:：)）]\s*')


def parse_detail_lines(text, roster=(), month=None):
    """解析手工录入/粘贴的明细文本(一行一条), 返回 (entries, warnings)。

    兼容带序号(「1、7.2叶么菊发现现场安全隐患+3」)与不带序号
    (「7.2叶么菊发现现场安全隐患+3」)两种写法; 行首序号只在
    「去掉后仍以日期开头」时才剥离(避免把「7.2」的月份一起吃掉)。
    """
    entries, warnings = [], []
    for i, line in enumerate(str(text).splitlines(), 1):
        raw = line.strip()
        if not raw:
            continue
        cands = [raw]
        m = SEQ_PREFIX_RE.match(raw)
        if m:
            cands.insert(0, raw[m.end():])
        e = None
        for cand in cands:
            t = normalize(cand)
            if not DATE_HEAD_RE.match(t):
                continue
            e = parse_entry(t, roster)
            if e is not None:
                break
        if e is None:
            warnings.append(f'第 {i} 行无法解析, 已忽略: {raw}')
            continue
        if month is not None and e['month'] is not None and e['month'] != month:
            warnings.append(f"第 {i} 行月份 {e['month']} 与所填月份({month})不一致, 请核对")
        if not (1 <= e['day'] <= 31):
            warnings.append(f"第 {i} 行日期 {e['day']} 超出 1~31, 请核对")
        if e['name'] is None:
            tip = ''
            if e.get('slot_trimmed'):
                tip = (f"(姓名槽里的「{e['slot_trimmed']}」已并回事由: "
                       f"「{e['reason']}」)")
            warnings.append(f"第 {i} 行姓名「{e['name_raw']}」未匹配名单, "
                            f'只写入明细存档不记分{tip}; '
                            f'若该人在本班, 请在「名单(报名表)」页签导入报名表')
        entries.append(e)
    return entries, warnings
