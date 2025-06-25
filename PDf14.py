#!/usr/bin/env python3
# -*- coding: utf-8 -*-

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
    m = re.search(r'\*[^*]*\*([^\*\n]+)', text)
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
    # 预处理标签
    text = re.sub(r'发\s*票\s*号\s*码\s*[:：]?', '发票号码:', text)
    text = re.sub(r'发\s*票\s*代\s*码\s*[:：]?', '发票代码:', text)

    rec = dict.fromkeys(['开票日期','发票代码','发票号码','商品名称','金额(小写)'], '')

    # 开票日期
    m = re.search(r'开票日期[:：]?\s*([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日', text)
    if not m:
        m = re.search(r'([0-9]{4})\s*年\s*([0-9]{1,2})\s*月\s*([0-9]{1,2})\s*日', text)
    if not m:
        m = re.search(r'(?<!\d)([0-9]{4})\s+([0-9]{1,2})\s+([0-9]{1,2})(?!\d)', text)
    if m:
        y, mo, d = m.group(1), m.group(2), m.group(3)
        rec['开票日期'] = f"{y}年{int(mo):02d}月{int(d)}日"

    # 中国铁路电子客票
    if '中国铁路' in text and '电子客票' in text:
        for k in ('发票代码','发票号码'):
            mm = PATTERNS[k].search(text)
            if mm: rec[k] = mm.group(1)
        m_amt = re.search(r'￥\s*([\d\.,]+)', text)
        rec['金额(小写)'] = m_amt.group(1) if m_amt else ''
        m_lbl = re.search(r'(改签费|退票费|票价)', text)
        rec['商品名称'] = f"中国铁路（{m_lbl.group(1)}）" if m_lbl else '中国铁路电子客票'
        return rec

    # 通用提取
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

    # 商品名称缩略
    cand = extract_product_candidate(text)
    if '汽油' in cand or '汽油' in text:
        rec['商品名称'] = '汽油'
    elif '餐饮' in cand or '餐饮' in text:
        rec['商品名称'] = '餐饮'
    else:
        chars = re.findall(r'[\u4e00-\u9fa5]', cand)
        if len(chars) >= 2:
            m_name = re.search(r'[\u4e00-\u9fa5]{2,}(?:服务|费|项目)?', cand)
            rec['商品名称'] = m_name.group(0) if m_name else ''

    # 大写金额校验
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
    IN = '/Users/zyt/Documents/VScode/Cursor/PDF_plumber/input'
    CSV_DIR = '/Users/zyt/Documents/VScode/Cursor/PDF_plumber/Excel'
    PDF_OUT = '/Users/zyt/Documents/VScode/Cursor/PDF_plumber/output'

    os.makedirs(CSV_DIR, exist_ok=True)
    os.makedirs(PDF_OUT, exist_ok=True)

    records = []
    for fn in sorted(os.listdir(IN)):
        if not fn.lower().endswith('.pdf'):
            continue
        path = os.path.join(IN, fn)
        txt = ocr_text(path)
        rec = parse_invoice(txt)
        rec['_文件名'] = fn
        records.append(rec)
        # 复制并重命名
        if rec['开票日期'] and rec['商品名称'] and rec['金额(小写)']:
            m_date = re.match(r'(\d{4})年(\d{2})月(\d{1,2})日', rec['开票日期'])
            if m_date:
                y, mo, da = m_date.groups()
                date_str = f"{y}-{int(mo)}-{da}"
                amt_int = rec['金额(小写)'].split('.')[0]
                new_name = f"{date_str}_{rec['商品名称']}_{amt_int}.pdf"
                shutil.copy(path, os.path.join(PDF_OUT, new_name))

    # 导出为 XLSX
    df = pd.DataFrame(records)
    xlsx_path = os.path.join(CSV_DIR, 'output.xlsx')
    df.to_excel(xlsx_path, index=False)

    print(f"✓ Done. XLSX: {xlsx_path}, PDFs: {PDF_OUT}")


if __name__ == '__main__':
    main()
