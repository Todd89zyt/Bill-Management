#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import csv
import argparse
from paddleocr import PaddleOCR

# —— 1. 初始化 OCR（关闭结构化） —— #
ocr = PaddleOCR(
    use_angle_cls=False,
    lang='ch',
    table=False,
    layout=False
)

# —— 2. 通用正则模式 —— #
PATTERNS = {
    '开票日期': re.compile(r'开票日期[:：]?\s*([0-9]{4}年[0-9]{1,2}月[0-9]{1,2}日)'),
    '发票代码': re.compile(r'发票代码[:：]?\s*([0-9A-Za-z]+)'),
    '发票号码': re.compile(r'发票号码[:：]?\s*([0-9A-Za-z]+)'),
    # 【改动】金额(小写) 前可带 ¥、￥ 或 OCR 误识别的 '?'，都当作货币符号
    '金额(小写)': re.compile(r'[（(]小写[)）]\s*[¥￥\?]?\s*([\d\.,]+)'),
}

def extract_product_name(text: str) -> str:
    """
    1) 优先用 *xxx*yyy 提取 yyy
    2) 若 OCR 少扫了前一个 * ，出现 xxx*yyy，则取 * 之后的 yyy
    """
    m = re.search(r'\*[^*]+\*([^\*\n]+)', text)
    if m:
        return m.group(1).strip()
    for line in text.splitlines():
        if '*' in line:
            parts = line.split('*')
            if len(parts) >= 2 and parts[-1].strip():
                return parts[-1].strip()
    return ''

def parse_invoice(text: str) -> dict:
    rec = {k: '' for k in ['开票日期','发票代码','发票号码','商品名称','金额(小写)']}

    # —— 1. 通用 PATTERNS —— #
    for key, pat in PATTERNS.items():
        m = pat.search(text)
        if m:
            rec[key] = m.group(1).strip()

    # —— 2. 常规商品名称 —— #
    rec['商品名称'] = extract_product_name(text)

    # —— 3. 高铁电子客票 专用 —— #
    if '电子客票' in text and (not rec['商品名称'] or not rec['金额(小写)']):
        m2 = re.search(
            r'(改签费|退票费|票价)[:：]?\s*[¥￥\?]?\s*([\d\.,]+)',
            text
        )
        if m2:
            fee_label = m2.group(1)
            raw_amount = m2.group(2).strip()
            rec['金额(小写)'] = raw_amount

            mt = re.search(r'[（(]([^）\)]*电子客票[^）\)]*)[）\)]', text)
            ticket_type = mt.group(1).strip() if mt else '铁路电子客票'
            rec['商品名称'] = f"{fee_label}（{ticket_type}）"

    # —— 4. 泛用回退：已有商品名称但金额仍空 —— #
    if rec['商品名称'] and not rec['金额(小写)']:
        label = re.escape(rec['商品名称'])
        m3 = re.search(
            rf'{label}[:：]\s*[¥￥\?]?\s*([\d\.,]+)',
            text
        )
        if m3:
            rec['金额(小写)'] = m3.group(1).strip()

    # —— 5. 金额格式化 —— #
    # 去掉所有非数字和小数点，保留两位小数
    amt = rec['金额(小写)']
    if amt:
        clean = re.sub(r'[^\d\.]', '', amt)
        try:
            rec['金额(小写)'] = f"{float(clean):.2f}"
        except ValueError:
            pass

    return rec

def ocr_text(path: str) -> str:
    raw = ocr.ocr(path, cls=False)
    texts = []
    def extract(obj):
        if isinstance(obj, str):
            texts.append(obj)
        elif isinstance(obj, (list, tuple)):
            for x in obj:
                extract(x)
    extract(raw)
    return '\n'.join(texts)

def main():
    p = argparse.ArgumentParser(description="批量 OCR 发票并输出 CSV")
    p.add_argument('--src-dir', required=True, help="发票图片目录")
    args = p.parse_args()

    IN  = args.src_dir
    OUT = os.path.join(os.getcwd(), 'output')
    os.makedirs(OUT, exist_ok=True)
    OUT_CSV = os.path.join(OUT, 'output.csv')

    with open(OUT_CSV, 'w', newline='', encoding='utf-8-sig') as fp:
        writer = csv.DictWriter(fp,
            fieldnames=['开票日期','发票代码','发票号码','商品名称','金额(小写)','_文件名'])
        writer.writeheader()

        for fn in sorted(os.listdir(IN)):
            if not fn.lower().endswith(('.png','.jpg','.jpeg')):
                continue
            imgp = os.path.join(IN, fn)
            print(f"\n=== 识别 {fn} ===")
            txt = ocr_text(imgp)
            print("OCR 原文:\n", txt)  # 调试用：看清 OCR 实际结果

            rec = parse_invoice(txt)
            rec['_文件名'] = fn
            print("→ 提取:", rec)
            writer.writerow(rec)

    print(f"\n✅ 完成，输出文件：{OUT_CSV}")

if __name__ == '__main__':
    main()
