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
        'errors': [],
        'warnings': []
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
            result['items'] = _extract_items(all_text)

            if not result['items']:
                result['warnings'].append('发票无商品明细行（简化发票），仅基础信息可用')

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
            else:
                # 无明细行时，用标准税率反推（默认13%）
                if result['total_amount_incl'] > 0 and result['total_amount_excl'] == 0:
                    tax_rate = DEFAULT_TAX_RATE / 100.0  # 13%
                    result['total_amount_excl'] = round(result['total_amount_incl'] / (1 + tax_rate), 2)
                    result['total_tax'] = round(result['total_amount_incl'] - result['total_amount_excl'], 2)

        return result

    except Exception as e:
        result['errors'].append(f'PDF解析异常: {str(e)}')
        return result


def _parse_item_line(line: str) -> dict:
    """解析单行商品明细"""
    # 模式1: 完整字段 (*电子元件*压敏电阻  WTR15D050MC3B3.5W  个  7000  0.4867...  3407.08  13%  442.92)
    m = re.match(
        r'(\S+)\s+'
        r'(\S+)\s+'
        r'(个|台|件|套|只|支|千克|千个|米|卷|升|张|瓶|盒|包|桶|次|组|块|根|PCS|K)\s+'
        r'(\d+(?:\.\d+)?)\s+'
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
        qty = float(m.group(4))
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
            'quantity': int(qty) if qty.is_integer() else qty,
            'unit_price_excl': round(unit_price, 13),
            'amount_excl': round(amount, 2),
            'tax_rate': tax_rate,
            'tax_amount': round(tax_amount, 2),
        }

    return None


def _extract_items(all_text: str) -> list:
    """Extract invoice item rows, including wrapped specification lines."""
    lines = all_text.split('\n')
    in_items = False
    items = []

    for line in lines:
        stripped = line.strip()
        if re.search(r'项目名称.*规格型号.*单\s*位.*数\s*量.*单\s*价.*金\s*额', stripped):
            in_items = True
            continue
        if not in_items:
            continue
        if re.match(r'^\s*合\s*计', stripped) or re.search(r'价税合计|备\s*注|开票人', stripped):
            break
        if not stripped:
            continue

        parsed = _parse_item_line(stripped)
        if parsed:
            items.append(parsed)
            continue

        if items and _is_spec_continuation(stripped):
            items[-1]['specification'] = f"{items[-1]['specification']}{stripped}"

    return items


def _is_spec_continuation(line: str) -> bool:
    """Return True when a short wrapped line is likely a specification suffix."""
    if re.search(r'[¥￥]|\d+\s*%|^\s*合\s*计', line):
        return False
    return bool(re.match(r'^[A-Za-z0-9.][A-Za-z0-9/·\-\.]*$', line))
