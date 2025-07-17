#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import warnings
warnings.filterwarnings("ignore", message="CropBox missing from /Page.*")

import logging
# 关闭 pdfminer 的冗余日志
for name in ("pdfminer.pdfparser", "pdfminer.pdfdocument", "pdfminer.pdfpage", "pdfminer.layout"):
    logging.getLogger(name).setLevel(logging.ERROR)

import os
import re
import shutil
import pdfplumber
import pandas as pd

# —— 通用正则模式 —— #
PATTERNS = {
    '发票代码': re.compile(r'发票代码[:：]?\s*([0-9A-Za-z]+)'),
    '发票号码': re.compile(r'发票号码[:：]?\s*([0-9A-Za-z]+)'),
    '金额(小写)': re.compile(r'[（(]\s*小写\s*[)）]\s*[¥￥\?]?\s*([\d\.,]+)'),
}

# 中文大写数字解析表
_CN_DIGITS = {'零':0,'壹':1,'贰':2,'叁':3,'肆':4,'伍':5,'陆':6,'柒':7,'捌':8,'玖':9}
_CN_UNITS = {'仟':1000,'佰':100,'拾':10}


def extract_product_candidate(text: str) -> str:
    """
    从类似 *经营租赁*代收通行费 这样的结构中，
    支持 ASCII '*' 和全角 '＊'，准确捕获“代收通行费”。
    """
    pattern = (
        r'[*＊]'                # 第一个星号
        r'[^*＊]*'              # 跳过第一对星号内所有字符
        r'[*＊]'                # 第二个星号
        r'\s*'                  # 可有空白
        r'([\u4e00-\u9fa5]{2,}(?:服务|费|项目)?)'  # 捕获2个以上汉字，可跟 服务/费/项目
    )
    m = re.search(pattern, text)
    return m.group(1).strip() if m else ''


def _parse_cn_integer(s: str) -> int:
    total, num = 0, 0
    for ch in s:
        if ch in _CN_DIGITS:
            num = num * 10 + _CN_DIGITS[ch]
        elif ch in _CN_UNITS:
            num = num or 1
            total += num * _CN_UNITS[ch]
            num = 0
    return total + num


def parse_invoice(text: str) -> dict:
    text = re.sub(r'发\s*票\s*号\s*码\s*[:：]?', '发票号码:', text)
    text = re.sub(r'发\s*票\s*代\s*码\s*[:：]?', '发票代码:', text)
    rec = dict.fromkeys(['开票日期','发票代码','发票号码','商品名称','金额(小写)'], '')
    m = re.search(r'开票日期[:：]?\s*([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日', text)
    if not m:
        m = re.search(r'([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日', text)
    if not m:
        m = re.search(r'(?<!\d)([0-9]{4})\s+([0-9]{1,2})\s+([0-9]{1,2})(?!\d)', text)
    if m:
        y, mo, d = m.group(1), m.group(2), m.group(3)
        rec['开票日期'] = f"{y}年{int(mo):02d}月{int(d)}日"
    if '中国铁路' in text and '电子客票' in text:
        for k in ('发票代码','发票号码'):
            mm = PATTERNS[k].search(text)
            if mm: rec[k] = mm.group(1)
        m_amt = re.search(r'￥\s*([\d\.,]+)', text)
        rec['金额(小写)'] = m_amt.group(1) if m_amt else ''
        m_lbl = re.search(r'(改签费|退票费|票价)', text)
        rec['商品名称'] = f"中国铁路（{m_lbl.group(1)}）" if m_lbl else '中国铁路电子客票'
        return rec
    for k in ('发票代码','发票号码','金额(小写)'):
        if not rec[k]:
            mm = PATTERNS[k].search(text)
            if mm: rec[k] = mm.group(1)
    if not rec['发票号码']:
        mm = re.search(r'发票监[^0-9]*(\d{6,})', text)
        if mm: rec['发票号码'] = mm.group(1)
    if '发票代码' in text and not rec['发票代码']:
        condensed = re.sub(r'\s+', '', text)
        m12 = re.search(r'(?<!\d)(\d{12})(?!\d)', condensed)
        if m12: rec['发票代码'] = m12.group(1)

    cand = extract_product_candidate(text)
    if cand:
        # 去掉数字、小数点和空白，保留前面纯中文部分作为名称
        name = re.sub(r'[\d\.\s].*$', '', cand).strip()
        rec['商品名称'] = name or ''
    else:
        # 兜底逻辑：原来的匹配方式
        if '汽油' in text:
            rec['商品名称'] = '汽油'
        elif '餐饮' in text:
            rec['商品名称'] = '餐饮'
        else:
            # 如果 cand 为空，或者清洗后还是空，再试老的方式
            # 提取两字以上的中文，后面可跟服务/费/项目等后缀
            m_name = re.search(r'[\u4e00-\u9fa5]{2,}(?:服务|费|项目)?', text)
            rec['商品名称'] = m_name.group(0) if m_name else ''
    m_h = re.search(
        r'[¥￥]?\s*([零壹贰叁肆伍陆柒捌玖拾佰仟]+)[圆圓]'
        r'([零壹贰叁肆伍陆柒捌玖])角'
        r'(?:([零壹贰叁肆伍陆柒捌玖])分)?', text
    )
    if m_h:
        amt_val = _parse_cn_integer(m_h.group(1)) + _CN_DIGITS[m_h.group(2)]*0.1 + _CN_DIGITS.get(m_h.group(3),0)*0.01
        try:
            curr = float(rec['金额(小写)'])
        except:
            curr = None
        if curr is None or abs(curr-amt_val)>0.01:
            rec['金额(小写)'] = f"{amt_val:.2f}"
    return rec


def ocr_text(path: str) -> str:
    with pdfplumber.open(path) as pdf:
        return '\n'.join(page.extract_text() or '' for page in pdf.pages)


def main():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    IN = os.path.join(BASE_DIR, 'input')
    XLSX_PATH = os.path.join(BASE_DIR, 'invoicePDF.xlsx')
    PDF_OUT = os.path.join(BASE_DIR, 'output')

    os.makedirs(IN, exist_ok=True)
    os.makedirs(PDF_OUT, exist_ok=True)

    records = []
    for fn in sorted(os.listdir(IN)):
        if not fn.lower().endswith('.pdf'):
            continue
        path = os.path.join(IN, fn)
        txt = ocr_text(path)
        print(f"\n\n=== DEBUG: 处理文件 {fn} ===")
        print("---- 文本抽取 (前500字符) ----")
        print(txt[:500])
        if len(txt) > 500:
            print("... (后续省略)\n")

        cand = extract_product_candidate(txt)
        print(f"商品候选 (cand): {cand!r}")  # 如果是 '' 说明 extract_product_candidate 没匹配到

        rec = parse_invoice(txt)

        # 2. 解析
        rec = parse_invoice(txt)
        rec['_文件名'] = fn

        print("---- parse_invoice 返回字段 ----")
        for field in ['开票日期','发票代码','发票号码','商品名称','金额(小写)']:
            val = rec.get(field, '')
            disp = val if val else "<空>"
            print(f"{field:>8} : {disp}")

        # 3. 拷贝决策
        can_copy = bool(rec['开票日期'] and rec['商品名称'] and rec['金额(小写)'])
        print(f"=> 拷贝条件 (开票日期 & 商品名称 & 金额): {can_copy}\n")

        rec = parse_invoice(txt)
        rec['_文件名'] = fn
        records.append(rec)
        if can_copy:
            m_date = re.match(r'(\d{4})年(\d{2})月(\d{1,2})日', rec['开票日期'])
            if m_date:
                y, mo, da = m_date.groups()
                date_str = f"{y}-{int(mo)}-{da}"
                amt_int = rec['金额(小写)'].split('.')[0]
                new_name = f"{date_str}_{rec['商品名称']}_{amt_int}.pdf"
                shutil.copy(
                    os.path.join(IN, fn),
                    os.path.join(PDF_OUT, new_name)
                )

    df = pd.DataFrame(records)
    mask = df['开票日期'].astype(bool) & df['商品名称'].astype(bool) & df['金额(小写)'].astype(bool)
    failed = df.loc[~mask, '_文件名']
    print("以下文件解析后缺少必要字段，没有拷贝：")
    for fn in failed:
        print(" -", fn)

        if rec['开票日期'] and rec['商品名称'] and rec['金额(小写)']:
            m_date = re.match(r'(\d{4})年(\d{2})月(\d{1,2})日', rec['开票日期'])
            if m_date:
                y, mo, da = m_date.groups()
                date_str = f"{y}-{int(mo)}-{da}"
                amt_int = rec['金额(小写)'].split('.')[0]
                new_name = f"{date_str}_{rec['商品名称']}_{amt_int}.pdf"
                shutil.copy(path, os.path.join(PDF_OUT, new_name))
    df = pd.DataFrame(records)
    df['金额(小写)'] = (
    df['金额(小写)']
    .str.replace(',', '', regex=False)           # 如果有逗号，先去掉
    .astype(float)                               # 转成数值类型
    
)
        # 用整数作为组 ID，自增即可
    next_group_id = 1
    groups = {}        # group_id -> set of filenames
    code_to_gid = {}   # 发票代码 -> group_id
    num_to_gid = {}    # 发票号码 -> group_id

    records = []
    for fn in sorted(os.listdir(IN)):
        if not fn.lower().endswith('.pdf'):
            continue
        path = os.path.join(IN, fn)
        txt = ocr_text(path)
        rec = parse_invoice(txt)
        rec['_文件名'] = fn
        records.append(rec)

        code = rec.get('发票代码','').strip()
        number = rec.get('发票号码','').strip()

        # 先看 code、number 各自有没有对应的组
        gid_code = code_to_gid.get(code)
        gid_num  = num_to_gid.get(number)

        if gid_code and gid_num and gid_code != gid_num:
            # 两个不同组需要合并
            merge_to, merge_from = gid_code, gid_num
            groups[merge_to].update(groups[merge_from])
            # 更新映射
            for f in groups[merge_from]:
                # update both code and number maps
                if records and records: pass
            del groups[merge_from]
            # 把 number_to_gid 指向 merge_to
            num_to_gid[number] = merge_to
            gid = merge_to
        else:
            # 优先用已有的组 ID，否则新建
            gid = gid_code or gid_num
            if not gid:
                gid = next_group_id
                next_group_id += 1
                groups[gid] = set()
        # 将当前文件加入组
        already = fn in groups[gid]
        groups[gid].add(fn)
        # 更新映射
        if code:
            code_to_gid[code] = gid
        if number:
            num_to_gid[number] = gid

        # 只有当这是第二个（或更多）文件加入时，且刚加入时才警示
        if len(groups[gid]) >= 2 and not already:
            names = sorted(groups[gid])
            print(f"⚠️ 重复发票组 {gid}：文件 {names}")
    
        # —— 接着是原有的拷贝逻辑 —— #
        if rec['开票日期'] and rec['商品名称'] and rec['金额(小写)']:
            m_date = re.match(r'(\d{4})年(\d{2})月(\d{1,2})日', rec['开票日期'])
            if m_date:
                y, mo, da = m_date.groups()
                date_str = f"{y}-{int(mo)}-{da}"
                amt_int = rec['金额(小写)'].split('.')[0]
                new_name = f"{date_str}_{rec['商品名称']}_{amt_int}.pdf"
                shutil.copy(
                    os.path.join(IN, fn),
                    os.path.join(PDF_OUT, new_name)
                )
                    
    with pd.ExcelWriter(XLSX_PATH, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # 创建一个“数字保留两位小数”格式
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})

        # 找到“金额(小写)”在 DataFrame 中的列索引（0 开头）
        col_idx = df.columns.get_loc('金额(小写)')
        # set_column 的第四个参数指定了默认单元格格式
        # 这里把该列的宽度留为默认（None），但绑定 money_fmt
        worksheet.set_column(col_idx, col_idx, None, money_fmt)

    print(f"✓ Done. XLSX: {XLSX_PATH}, PDFs: {PDF_OUT}")


if __name__ == '__main__':
    main()
