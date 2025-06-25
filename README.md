# 发票解析脚本简介

`PDf14.py` 是一个使用 [pdfplumber](https://github.com/jsvine/pdfplumber) 的示例脚本，能批量读取指定文件夹中的 PDF 发票文本，利用正则表达式解析常见字段（开票日期、发票代码、发票号码、商品名称及金额等），并将结果汇总到 Excel。

## 依赖
- Python 3.8+
- `pdfplumber`
- `pandas`

安装依赖：
```bash
pip install pdfplumber pandas
```

## 使用方法
1. 打开 `PDf14.py`，根据需要修改脚本内的 `IN`、`CSV_DIR`、`PDF_OUT` 三个目录常量。
2. 将待处理的 PDF 文件放入 `IN` 指定的文件夹。
3. 运行脚本：
```bash
python PDf14.py
```

处理完成后，会在 `CSV_DIR` 中生成 `output.xlsx`，并在 `PDF_OUT` 中复制一份按日期和金额重命名后的 PDF 文件。
