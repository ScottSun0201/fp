#!/usr/bin/env python3
"""
对账单PDF解析引擎
REQ-031: 解析月度对账单PDF -> 提取客户信息+明细行+四项资金
"""
import re
import pdfplumber
from config import AMOUNT_TOLERANCE


def parse_statement_pdf(pdf_path: str) -> dict:
    """从对账单PDF中提取全部字段"""
    result = {
        'customer_name': '',
        'customer_tax_id': '',
        'supplier_name': '',
        'supplier_tax_id': '',
        'statement_month': '',
        'statement_date': '',
        'settlement_days': 30,
        'opening_balance': 0.0,
        'current_payment': 0.0,
        'closing_balance': 0.0,
        'delivered_unpaid': 0.0,
        'total_invoice_amount': 0.0,
        'total_quantity': 0,
        'items': [],
        'balance_check': True,
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
                result['errors'].append('PDF无可提取文本（可能是扫描件，需OCR）')
                return result

            # 客户名称
            m = re.search(r'客户[：:]\s*(.+?)(?:公司|集团)', all_text)
            if m:
                result['customer_name'] = m.group(0).replace('客户：', '').replace('客户:', '').strip()

            # 供应商名称
            m = re.search(r'供应商[：:]\s*(.+?)(?:公司|集团)', all_text)
            if m:
                result['supplier_name'] = m.group(0).replace('供应商：', '').replace('供应商:', '').strip()

            # 对账月份
            m = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月', all_text)
            if m:
                result['statement_month'] = f"{m.group(1)}-{int(m.group(2)):02d}"

            # 月结天数
            m = re.search(r'月结\s*(\d+)\s*天', all_text)
            if m:
                result['settlement_days'] = int(m.group(1))

            # 四项资金
            result['opening_balance'] = _extract_amount(all_text, r'期初[^\d]*?([\d,]+\.\d{2})')
            result['current_payment'] = _extract_amount(all_text, r'本期(?:付款|回款)[^\d]*?[¥￥]?\s*([\d,]+\.?\d*)')
            result['total_invoice_amount'] = _extract_amount(all_text, r'(?:本次|本期)(?:对帐|对账)?开票[^\d]*?[¥￥]?\s*([\d,]+\.?\d*)')
            result['delivered_unpaid'] = _extract_amount(all_text, r'已交货未回款[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)')
            result['closing_balance'] = _extract_amount(all_text, r'期末[^\d]*?([\d,]+\.\d{2})')

            if result['closing_balance'] == 0:
                result['closing_balance'] = result['opening_balance'] + result['total_invoice_amount'] - result['current_payment']

            # 平衡校验
            expected = result['opening_balance'] + result['total_invoice_amount'] - result['current_payment']
            if abs(expected - result['closing_balance']) > AMOUNT_TOLERANCE and result['closing_balance'] > 0:
                result['balance_check'] = False
                diff = round(expected - result['closing_balance'], 2)
                result['errors'].append(f'四项资金不平: 差异={diff:.2f}')

            # 明细行
            result['items'] = _parse_items(all_text)
            if not result['items']:
                result['errors'].append('未能解析对账单明细行')

            if result['items']:
                calc_total = sum(it['amount_incl_tax'] for it in result['items'])
                calc_qty = sum(it['quantity'] for it in result['items'])
                result['total_quantity'] = int(calc_qty)
                if result['total_invoice_amount'] == 0:
                    result['total_invoice_amount'] = round(calc_total, 2)

        return result

    except Exception as e:
        result['errors'].append(f'PDF解析异常: {str(e)}')
        return result


def _extract_amount(text: str, pattern: str) -> float:
    m = re.search(pattern, text)
    if m:
        try:
            return float(m.group(1).replace(',', ''))
        except (ValueError, IndexError):
            pass
    return 0.0


def _parse_items(text: str) -> list:
    items = []
    lines = text.split('\n')
    in_items = False

    for line in lines:
        stripped = line.strip()
        if re.search(r'序号.*?(?:出库日期|发货日期)', stripped) or \
           re.search(r'(?:物料编码|客户料号).*?(?:数量|金额)', stripped):
            in_items = True
            continue
        if not in_items:
            continue
        if re.match(r'^\s*(?:合\s*计|小\s*计|未税|本期|期初|期末|备注)', stripped):
            break
        if re.search(r'(?:未税总额|税额|含税总额|本次.*开票)', stripped):
            break
        if not stripped:
            continue
        item = _parse_item_line(stripped)
        if item:
            items.append(item)

    return items


def _parse_item_line(line: str) -> dict:
    cm = re.search(r'(LW\d{9})', line)
    customer_code = cm.group(1) if cm else ''

    dm = re.search(r'(S\d{7})', line)
    delivery_no = dm.group(1) if dm else ''

    om = re.search(r'(\d{6}-[A-Z]{2,4})', line)
    order_no = om.group(1) if om else ''

    dtm = re.search(r'(\d{4}[-./]\d{1,2}[-./]\d{1,2})', line)
    delivery_date = dtm.group(1).replace('/', '-').replace('.', '-') if dtm else ''

    qty_matches = re.findall(r'\b(\d{3,6})\b', line)
    quantity = 0
    for q in qty_matches:
        v = int(q)
        if 100 <= v <= 999999:
            quantity = v
            break

    price_matches = re.findall(r'\b(\d+\.\d{2,6})\b', line)
    unit_price = 0.0
    amount = 0.0
    if len(price_matches) >= 2:
        unit_price = float(price_matches[-2])
        amount = float(price_matches[-1])
    elif len(price_matches) == 1:
        amount = float(price_matches[0])

    name_parts = re.findall(r'[\u4e00-\u9fff]+', line)
    product_name = ' '.join(name_parts) if name_parts else line[:30]

    if quantity > 0 and (unit_price > 0 or amount > 0):
        if amount == 0 and unit_price > 0:
            amount = round(quantity * unit_price, 2)
        return {
            'customer_order_no': order_no,
            'customer_material_code': customer_code,
            'delivery_no': delivery_no,
            'delivery_date': delivery_date,
            'product_name': product_name,
            'quantity': quantity,
            'unit': 'PCS',
            'unit_price_incl_tax': round(unit_price, 6),
            'amount_incl_tax': round(amount, 2),
        }
    return None
