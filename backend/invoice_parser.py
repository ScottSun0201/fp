#!/usr/bin/env python3
"""
发票PDF解析引擎
REQ-004/005/006/013: 增值税电子发票全字段解析
"""
import re
import pdfplumber
from config import DEFAULT_TAX_RATE


def parse_invoice_pdf(pdf_path: str) -> dict:
    """从增值税电子发票PDF中提取全部字段"""
    result = {
        'invoice_number': '',
        'invoice_date': '',
        'invoice_type': 'VAT_SPECIAL',
        'buyer_name': '',
        'buyer_tax_id': '',
        'seller_name': '',
        'seller_tax_id': '',
        'total_amount_excl': 0.0,
        'total_tax': 0.0,
        'total_amount_incl': 0.0,
        'amount_capital': '',
        'items': [],
        'raw_text': '',
        'errors': []
    }

    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ''
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    all_text += t + '\n'
            result['raw_text'] = all_text

            if not all_text.strip():
                result['errors'].append('PDF无可提取文本（可能是扫描件）')
                return result

            # ── 发票号码 ──
            m = re.search(r'发票号码[：:\s]+(\d{8,30})', all_text)
            if m:
                result['invoice_number'] = m.group(1)

            # ── 开票日期 ──
            m = re.search(r'开票日期[：:\s]*(\d{4})年(\d{1,2})月(\d{1,2})日', all_text)
            if m:
                result['invoice_date'] = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
            else:
                m = re.search(r'开票日期[：:\s]+(\d{4}-\d{2}-\d{2})', all_text)
                if m:
                    result['invoice_date'] = m.group(1)

            # ── 购买方 ──
            m = re.search(r'名\s*称[：:]\s*(.+?公司)', all_text)
            if m:
                result['buyer_name'] = m.group(1).strip()

            if result['buyer_name']:
                idx = all_text.find(result['buyer_name'])
                if idx > 0:
                    tail = all_text[idx:idx+200]
                    m2 = re.search(r'(?:纳税人识别号|信用代码)[：:\s]+([0-9A-Za-z]{15,25})', tail)
                    if m2:
                        result['buyer_tax_id'] = m2.group(1)

            # ── 销售方 ──
            companies = re.findall(r'名\s*称[：:]\s*(.+?公司)', all_text)
            if len(companies) >= 2:
                result['seller_name'] = companies[-1].strip()
            tax_ids = re.findall(r'(?:纳税人识别号|信用代码)[：:\s]+([0-9A-Za-z]{15,25})', all_text)
            if len(tax_ids) >= 2:
                result['seller_tax_id'] = tax_ids[-1]

            # ── 价税合计(小写) ──
            m = re.search(r'[（(]小写[）)]\s*[¥￥]?\s*([\d,.]+)', all_text)
            if m:
                result['total_amount_incl'] = float(m.group(1).replace(',', ''))

            # ── 价税合计(大写) ──
            m = re.search(r'[（(]大写[）)]\s*(.{2,20}?)(?:[（(]小写|$)', all_text)
            if m:
                result['amount_capital'] = m.group(1).strip()

            # ── 商品明细行 ──
            lines = all_text.split('\n')
            in_items = False
            item_buf = []
            for line in lines:
                stripped = line.strip()
                if re.search(r'项目名称.*规格型号.*单位.*数量.*单价.*金额', stripped):
                    in_items = True
                    continue
                if not in_items:
                    continue
                if re.match(r'^\s*合\s*计', stripped) or re.search(r'价税合计', stripped):
                    break
                if not stripped:
                    continue
                if item_buf and not re.match(r'^[*\d]', stripped):
                    item_buf[-1] += stripped
                else:
                    item_buf.append(stripped)

            for buf_line in item_buf:
                parsed = _parse_item_line(buf_line)
                if parsed:
                    result['items'].append(parsed)

            if not result['items']:
                result['errors'].append('未能解析商品明细行')

            # ── 价税分离验证 ──
            if result['items']:
                calc_excl = sum(it['amount_excl'] for it in result['items'])
                calc_tax = sum(it['tax_amount'] for it in result['items'])
                calc_incl = calc_excl + calc_tax
                result['total_amount_excl'] = round(calc_excl, 2)
                result['total_tax'] = round(calc_tax, 2)
                if result['total_amount_incl'] == 0:
                    result['total_amount_incl'] = round(calc_incl, 2)
                elif abs(calc_incl - result['total_amount_incl']) > 0.02:
                    result['errors'].append(
                        f'价税合计不平: 计算值{calc_incl:.2f} ≠ 票面值{result["total_amount_incl"]:.2f}'
                    )

        return result

    except Exception as e:
        result['errors'].append(f'PDF解析异常: {str(e)}')
        return result


def _parse_item_line(line: str) -> dict:
    """解析单行商品明细"""
    # 模式1: 完整字段 (*电子元件*压敏电阻  WTR15D050MC3B3.5W  个  7000  0.4867...  3407.08  13%  442.92)
    m = re.match(
        r'([*·][\u4e00-\u9fff（）()A-Za-z0-9·]+?)\s+'
        r'([\w\u4e00-\u9fff/·\-\.]+?)\s+'
        r'(个|台|件|套|只|支|千克|米|卷|升|张|瓶|盒|包|桶|次|组|块|根|PCS)\s+'
        r'(\d+)\s+'
        r'([\d\.]+)\s+'
        r'([\d\.]+)'
        r'(?:\s+(\d+%)?)?'
        r'(?:\s+([\d\.]+))?',
        line
    )
    if m:
        name_raw = m.group(1).strip()
        spec = m.group(2).strip()
        unit = m.group(3)
        qty = int(m.group(4))
        unit_price = float(m.group(5))
        amount = float(m.group(6))
        tax_rate_str = m.group(7) or ''
        tax_amount = float(m.group(8)) if m.group(8) else 0.0

        # 提取分类前缀和物料名
        cat_prefix = ''
        mat_name = name_raw
        cm = re.match(r'[*·]([\u4e00-\u9fff]+)[*·](.+)', name_raw)
        if cm:
            cat_prefix = f"*{cm.group(1)}*"
            mat_name = cm.group(2)

        tax_rate = float(tax_rate_str.replace('%', '')) if tax_rate_str else DEFAULT_TAX_RATE
        if tax_amount == 0 and amount > 0:
            tax_amount = round(amount * tax_rate / 100, 2)

        return {
            'category_prefix': cat_prefix,
            'material_name': mat_name,
            'specification': spec,
            'unit': unit,
            'quantity': qty,
            'unit_price_excl': round(unit_price, 13),
            'amount_excl': round(amount, 2),
            'tax_rate': tax_rate,
            'tax_amount': round(tax_amount, 2),
        }

    return None
