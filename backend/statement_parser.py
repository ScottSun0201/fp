#!/usr/bin/env python3
"""
对账单PDF解析引擎
REQ-031: 解析月度对账单PDF -> 提取客户信息+明细行+四项资金
"""
import re
import pdfplumber
from openpyxl import load_workbook
from config import AMOUNT_TOLERANCE


def parse_statement_pdf(pdf_path: str) -> dict:
    """从对账单PDF中提取全部字段"""
    result = {
        'customer_name': '',
        'customer_tax_id': '',
        'supplier_name': '',
        'supplier_code': '',
        'supplier_tax_id': '',
        'statement_month': '',
        'statement_date': '',
        'usage_remark': '',
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
        ocr_used = False
        with pdfplumber.open(pdf_path) as pdf:
            all_text = ''
            table_rows = []
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    all_text += t + '\n'
                for table in page.extract_tables() or []:
                    for row in table:
                        if row and any(cell not in (None, '') for cell in row):
                            table_rows.append(tuple(row))
            result['raw_text'] = all_text

            if not all_text.strip() and not table_rows:
                all_text = _ocr_pdf_text(pdf_path)
                result['raw_text'] = all_text
                ocr_used = True
                if not all_text.strip():
                    result['errors'].append('PDF无可提取文本（可能是扫描件，需OCR）')
                    return result

            if table_rows:
                _merge_template_fields(result, table_rows)
                result['items'] = _parse_xlsx_items(table_rows)

            # 客户名称
            m = re.search(r'(?:客户名称|客户)[：:\s]*([^\n]+?(?:公司|集团)[^\n]*)', all_text)
            if m:
                result['customer_name'] = result['customer_name'] or _clean_party_name(m.group(1))

            # 供应商名称
            m = re.search(r'(?:供应商名称|供应商)[：:\s]*([^\n]+?(?:公司|集团)[^\n]*)', all_text)
            if m:
                result['supplier_name'] = result['supplier_name'] or _clean_party_name(m.group(1))

            # 对账月份
            m = re.search(r'(?:对账月份|账期|月份)[：:\s]*(\d{4}[-./年]\s*\d{1,2})', all_text)
            if m:
                result['statement_month'] = result['statement_month'] or _format_month(m.group(1))
            m = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月', all_text)
            if m:
                result['statement_month'] = result['statement_month'] or f"{m.group(1)}-{int(m.group(2)):02d}"

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

            if ocr_used and not result['items']:
                result['items'] = _parse_ocr_column_items(all_text)
            if not ocr_used and not result['items']:
                result['items'] = _parse_template_text_items(all_text)
            if not ocr_used and not result['items']:
                result['items'] = _parse_ocr_column_items(all_text)
            if not ocr_used and not result['items']:
                result['items'] = _parse_items(all_text)
            if not result['items']:
                result['errors'].append('未能解析对账单明细行')

            if result['items']:
                calc_total = sum(it['amount_incl_tax'] for it in result['items'])
                calc_qty = sum(it['quantity'] for it in result['items'])
                result['total_quantity'] = int(calc_qty)
                if result['total_invoice_amount'] == 0 or ocr_used:
                    result['total_invoice_amount'] = round(calc_total, 2)

        return result

    except Exception as e:
        result['errors'].append(f'PDF解析异常: {str(e)}')
        return result


def _ocr_pdf_text(pdf_path: str) -> str:
    """Use macOS Vision OCR for scanned PDFs."""
    try:
        import Foundation
        import Vision
        import pypdfium2 as pdfium
    except Exception:
        return ''

    texts = []
    try:
        pdf = pdfium.PdfDocument(pdf_path)
        for idx, page in enumerate(pdf):
            image_path = f"/tmp/lwfp_ocr_{idx}.png"
            page.render(scale=3).to_pil().save(image_path)
            req = Vision.VNRecognizeTextRequest.alloc().init()
            req.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
            req.setRecognitionLanguages_(['zh-Hans', 'en-US'])
            req.setUsesLanguageCorrection_(True)
            url = Foundation.NSURL.fileURLWithPath_(image_path)
            handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(url, {})
            ok, _ = handler.performRequests_error_([req], None)
            if not ok:
                continue
            for obs in req.results() or []:
                candidates = obs.topCandidates_(1)
                if candidates:
                    texts.append(str(candidates[0].string()))
    except Exception:
        return ''
    return '\n'.join(texts)


def parse_statement_xlsx(xlsx_path: str) -> dict:
    """从对账单 Excel 中提取客户信息、账期、汇总和明细。"""
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
        wb = load_workbook(xlsx_path, data_only=True, read_only=True)
        ws = wb.active
        rows = []
        text_parts = []
        for row in ws.iter_rows(values_only=True):
            values = [v for v in row if v is not None]
            if not values:
                continue
            rows.append(row)
            text_parts.extend(str(v) for v in values)
        all_text = '\n'.join(text_parts)
        result['raw_text'] = all_text

        result['supplier_code'] = _find_adjacent_value(rows, ['供应商编码'])
        result['customer_name'] = _find_adjacent_value(rows, ['客户名称', '客户'])
        result['supplier_name'] = _find_adjacent_value(rows, ['供应商名称', '供应商'])
        result['supplier_tax_id'] = _find_adjacent_value(rows, ['供应商税号', '供应商纳税人识别号', '税号'])
        result['statement_month'] = _format_month(_find_adjacent_value(rows, ['对账月份', '账期', '月份']))
        result['statement_date'] = _format_date(_find_adjacent_value(rows, ['制表日期', '对账日期']))
        result['usage_remark'] = _find_adjacent_value(rows, ['用途备注', '用途', '付款用途'])
        settlement_days = _to_float(_find_adjacent_value(rows, ['结算天数', '月结天数']))
        if settlement_days:
            result['settlement_days'] = int(settlement_days)

        m = re.search(r'客户[：:]\s*([^\n]+)', all_text)
        if m:
            result['customer_name'] = result['customer_name'] or m.group(1).strip()
        m = re.search(r'供应商[：:]\s*([^\n]+)', all_text)
        if m:
            result['supplier_name'] = result['supplier_name'] or m.group(1).strip()
        if not result['supplier_name'] and text_parts:
            result['supplier_name'] = text_parts[0].strip()

        m = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月份?', all_text)
        if m:
            result['statement_month'] = result['statement_month'] or f"{m.group(1)}-{int(m.group(2)):02d}"
        m = re.search(r'制表[：:]\s*(\d{4}[-./]\d{1,2}[-./]\d{1,2})', all_text)
        if m:
            result['statement_date'] = result['statement_date'] or m.group(1).replace('/', '-').replace('.', '-')

        if result['supplier_code'] and result['statement_month']:
            result['statement_key'] = f"{result['supplier_code']}_{result['statement_month']}"

        result['items'] = _parse_xlsx_items(rows)
        if not result['items']:
            result['errors'].append('未能解析对账单明细行')
        else:
            result['total_quantity'] = int(sum(it['quantity'] for it in result['items']))
            result['total_invoice_amount'] = round(sum(it['amount_incl_tax'] for it in result['items']), 2)
            result['closing_balance'] = result['total_invoice_amount']

        return result
    except Exception as e:
        result['errors'].append(f'Excel解析异常: {str(e)}')
        return result


def _parse_xlsx_items(rows: list) -> list:
    items = []
    header = None
    header_idx = -1

    for idx, row in enumerate(rows):
        values = [str(v).strip() if v is not None else '' for v in row]
        normalized_values = [_normalize_header(v) for v in values]
        if '序号' in values and ('数量' in values or '金额' in values):
            header = values
            header_idx = idx
            break
        if '序号' in normalized_values and ('数量' in normalized_values or '金额' in normalized_values):
            header = values
            header_idx = idx
            break
    if header is None:
        return items

    index = {}
    for i, name in enumerate(header):
        normalized = _normalize_header(name)
        if normalized:
            index[normalized] = i
    required = ['序号', '数量', '金额']
    if any(name not in index for name in required):
        return items

    for row in rows[header_idx + 1:]:
        seq = _cell(row, index.get('序号'))
        if '合计' in str(seq or ''):
            break
        qty = _to_float(_cell(row, index.get('数量')))
        amount = _to_float(_cell(row, index.get('金额')))
        if qty <= 0 or amount <= 0:
            continue
        if isinstance(seq, (int, float)) and int(seq) > 0:
            seq_no = int(seq)
        else:
            seq_no = len(items) + 1

        row_values = [str(value or '') for value in row]
        customer_material_code = _normalize_customer_material_code(
            _cell(row, index.get('客户物料编码'))
        ) or _extract_customer_material_code(' '.join(row_values))

        item = {
            'seq': seq_no,
            'customer_order_no': str(_cell(row, index.get('客户订单号')) or '').strip(),
            'customer_material_code': customer_material_code,
            'delivery_no': str(_cell(row, index.get('送货单号')) or '').strip(),
            'delivery_date': _format_date(_cell(row, index.get('交货日期'))),
            'product_name': str(_cell(row, index.get('商品名称')) or '').strip(),
            'quantity': qty,
            'unit': str(_cell(row, index.get('单位')) or 'PCS').strip(),
            'unit_price_incl_tax': round(_to_float(_cell(row, index.get('单价'))), 6),
            'amount_incl_tax': round(amount, 2),
        }
        if not item['product_name']:
            item['product_name'] = str(_cell(row, index.get('商品型号')) or _cell(row, index.get('规格型号')) or '').strip()
        items.append(item)

    return items


def _merge_template_fields(result: dict, rows: list):
    result['supplier_code'] = result.get('supplier_code') or _find_adjacent_value(rows, ['供应商编码'])
    result['customer_name'] = result.get('customer_name') or _find_adjacent_value(rows, ['客户名称', '客户'])
    result['supplier_name'] = result.get('supplier_name') or _find_adjacent_value(rows, ['供应商名称', '供应商'])
    result['supplier_tax_id'] = result.get('supplier_tax_id') or _find_adjacent_value(rows, ['供应商税号', '供应商纳税人识别号', '税号'])
    result['statement_month'] = result.get('statement_month') or _format_month(_find_adjacent_value(rows, ['对账月份', '账期', '月份']))
    result['statement_date'] = result.get('statement_date') or _format_date(_find_adjacent_value(rows, ['制表日期', '对账日期']))
    settlement_days = _to_float(_find_adjacent_value(rows, ['结算天数', '月结天数']))
    if settlement_days:
        result['settlement_days'] = int(settlement_days)
    if result.get('supplier_code') and result.get('statement_month'):
        result['statement_key'] = f"{result['supplier_code']}_{result['statement_month']}"


def _parse_template_text_items(text: str) -> list:
    """Parse text PDFs generated from the unified Excel template."""
    items = []
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    header_seen = False

    for line in lines:
        normalized_line = _normalize_header(line.replace(' ', ''))
        if '序号' in normalized_line and '数量' in normalized_line and '金额' in normalized_line:
            header_seen = True
            continue
        if not header_seen:
            continue
        if re.match(r'^(合计|填写说明|必填项)', line):
            break

        item = _parse_template_text_line(line)
        if item:
            item['seq'] = len(items) + 1
            items.append(item)

    return items


def _parse_template_text_line(line: str) -> dict:
    parts = [part for part in re.split(r'\s+', line.strip()) if part]
    if len(parts) < 8:
        return None
    if not re.match(r'^\d+$', parts[0]):
        return None

    date_idx = None
    for idx, part in enumerate(parts):
        if re.match(r'^\d{4}[-./]\d{1,2}[-./]\d{1,2}$', part):
            date_idx = idx
            break
    if date_idx is None or date_idx < 2:
        return None

    numeric_positions = []
    for idx, part in enumerate(parts):
        if _to_float(part) > 0:
            numeric_positions.append(idx)
    if len(numeric_positions) < 3:
        return None

    amount_idx = numeric_positions[-1]
    price_idx = numeric_positions[-2]
    qty_idx = numeric_positions[-3]
    if qty_idx <= date_idx:
        return None

    unit_idx = qty_idx - 1
    product_parts = parts[date_idx + 1:unit_idx]
    product_name = product_parts[0] if product_parts else ''
    specification = ' '.join(product_parts[1:]) if len(product_parts) > 1 else ''

    customer_material_code = _extract_customer_material_code(' '.join(parts))
    if not customer_material_code:
        return None

    return {
        'customer_order_no': parts[1] if len(parts) > 1 else '',
        'customer_material_code': customer_material_code,
        'supplier_material_code': parts[3] if len(parts) > 3 and date_idx > 4 else '',
        'delivery_no': parts[date_idx - 1] if date_idx >= 4 else '',
        'delivery_date': _format_date(parts[date_idx]),
        'product_name': product_name or specification or line[:30],
        'specification': specification,
        'quantity': _to_float(parts[qty_idx]),
        'unit': parts[unit_idx] if unit_idx >= 0 else 'PCS',
        'unit_price_incl_tax': round(_to_float(parts[price_idx]), 6),
        'amount_incl_tax': round(_to_float(parts[amount_idx]), 2),
    }


def _parse_ocr_column_items(text: str) -> list:
    """Parse scanned statement OCR output where columns are emitted top-to-bottom."""
    lines = [line.strip() for line in str(text or '').splitlines() if line.strip()]
    if not lines:
        return []

    material_codes = [_extract_customer_material_code(line) for line in lines]
    material_codes = [code for code in material_codes if code]
    delivery_nos = []
    for line in lines:
        normalized = line.replace('$', 'S').replace('＄', 'S').upper()
        m = re.search(r'\bS\d{6,8}\b', normalized)
        if m:
            delivery_nos.append(m.group(0))
    dates = [_format_date(m.group(1)) for line in lines for m in [re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', line)] if m]
    order_nos = []
    for line in lines:
        normalized = line.upper().replace('™', 'W').replace('LR', 'LW')
        m = re.search(r'\b(\d{6}-L[WIV]{1,2})\b', normalized)
        if m:
            order_nos.append(m.group(1).replace('LIV', 'LW').replace('LV', 'LW'))

    product_names = [
        line for line in _section_lines(lines, ['产品名称'], ['款项信息', '商品名称'])
        if re.search(r'[\u4e00-\u9fff]', line)
    ]
    specs = _section_lines(lines, ['商品名称'], ['本次对帐开票', '本次对账开票', '出库数量'])
    quantities = [_to_float(v) for v in _numeric_section(lines, ['出库数量'], ['单位'])]
    units = [line for line in _section_lines(lines, ['单位'], ['对帐日期', '对账日期', '含型单价', '含税单价']) if re.match(r'^[A-Za-z]+$', line)]
    prices = [_to_float(v) for v in _numeric_section(lines, ['含型单价', '含税单价', '含税单价'], ['Total', '合计'])]
    amounts = [_to_float(v) for v in _numeric_section(lines, ['Total', '合计'], ['每注', '备注', '路合同', '合同'])]

    if amounts and len(amounts) > len(material_codes):
        amounts = amounts[:len(material_codes)]
    count = max(len(material_codes), len(delivery_nos), len(dates), len(quantities), len(amounts))
    items = []
    for idx in range(count):
        code = _at(material_codes, idx)
        qty = _at(quantities, idx, 0)
        price = _at(prices, idx, 0)
        amount = _at(amounts, idx, 0)
        if not code or qty <= 0 or amount <= 0:
            continue
        if price <= 0 and qty:
            price = round(amount / qty, 6)
        items.append({
            'customer_order_no': _at(order_nos, idx, _at(order_nos, 0, '')),
            'customer_material_code': code,
            'delivery_no': _at(delivery_nos, idx, ''),
            'delivery_date': _at(dates, idx, ''),
            'product_name': _at(product_names, idx, _at(specs, idx, code)),
            'specification': _at(specs, idx, ''),
            'quantity': qty,
            'unit': _at(units, idx, 'PCS').upper(),
            'unit_price_incl_tax': round(price, 6),
            'amount_incl_tax': round(amount, 2),
        })
    return items


def _section_lines(lines, start_markers, end_markers):
    start = None
    for idx, line in enumerate(lines):
        if any(marker in line for marker in start_markers):
            start = idx + 1
            break
    if start is None:
        return []
    end = len(lines)
    for idx in range(start, len(lines)):
        if any(marker in lines[idx] for marker in end_markers):
            end = idx
            break
    return [line for line in lines[start:end] if line]


def _numeric_section(lines, start_markers, end_markers):
    values = []
    for line in _section_lines(lines, start_markers, end_markers):
        cleaned = line.replace('：', '').replace(':', '').replace(' ', '')
        if re.match(r'^\d+(?:\.\d+)?$', cleaned):
            values.append(cleaned)
    return values


def _at(values, index, default=''):
    return values[index] if index < len(values) else default


def _normalize_header(value) -> str:
    text = str(value or '').strip()
    text = re.sub(r'[\s*＊（）()]', '', text)
    aliases = {
        '含税单价': '单价',
        '含税金额': '金额',
        '本次金额': '金额',
        '销售金额': '金额',
        '销售数量': '数量',
        '客户料号': '客户物料编码',
        '物料编码': '客户物料编码',
        '规格': '规格型号',
        '型号': '规格型号',
        '出库单号': '送货单号',
        '发货单号': '送货单号',
        '出库日期': '交货日期',
        '发货日期': '交货日期',
        '供应商纳税人识别号': '供应商税号',
    }
    return aliases.get(text, text)


def _clean_party_name(value: str) -> str:
    text = str(value or '').strip()
    text = re.sub(r'\s+', ' ', text)
    for marker in ('对账月份', '客户名称', '供应商名称', '制表日期', '结算天数'):
        if marker in text:
            text = text.split(marker, 1)[0].strip()
    return text


def _normalize_customer_material_code(value) -> str:
    text = str(value or '').strip().upper()
    if not text:
        return ''
    text = text.replace('Ｌ', 'L').replace('Ｗ', 'W')
    text = re.sub(r'[^A-Z0-9]', '', text)
    replacements = (
        ('1W', 'LW'),
        ('IW', 'LW'),
        ('|W', 'LW'),
        ('LIV', 'LW'),
        ('LVV', 'LW'),
        ('LNW', 'LW'),
        ('LR', 'LW'),
    )
    for wrong, right in replacements:
        if text.startswith(wrong):
            text = right + text[len(wrong):]
            break
    match = re.search(r'LW(\d{5,12})', text)
    if match:
        return f"LW{match.group(1)}"
    return ''


def _extract_customer_material_code(text: str) -> str:
    direct = _normalize_customer_material_code(text)
    if direct:
        return direct
    compact = re.sub(r'\s+', '', str(text or '').upper())
    for match in re.finditer(r'(?:L|1|I|\|)\s*W\s*[\d\s]{5,16}', compact):
        code = _normalize_customer_material_code(match.group(0))
        if code:
            return code
    return ''


def _find_adjacent_value(rows, labels):
    normalized_labels = {_normalize_header(label) for label in labels}
    for row in rows:
        values = [str(v).strip() if v is not None else '' for v in row]
        for idx, value in enumerate(values):
            if not value:
                continue
            inline = re.match(r'^(.+?)[：:]\s*(.+)$', value)
            if inline and _normalize_header(inline.group(1)) in normalized_labels:
                return inline.group(2).strip()
            normalized = _normalize_header(value.rstrip('：:'))
            if normalized in normalized_labels:
                for next_idx in range(idx + 1, len(values)):
                    if values[next_idx]:
                        return values[next_idx].strip()
    return ''


def _cell(row, idx):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def _to_float(value) -> float:
    if value is None or value == '':
        return 0.0
    try:
        return float(str(value).replace(',', ''))
    except ValueError:
        return 0.0


def _format_date(value) -> str:
    if value is None:
        return ''
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    return str(value).replace('/', '-').replace('.', '-')


def _format_month(value) -> str:
    if value is None or value == '':
        return ''
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m')
    text = str(value).strip()
    m = re.search(r'(\d{4})[-./年]\s*(\d{1,2})', text)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    return text


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
    customer_code = _extract_customer_material_code(line)

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

    if customer_code and quantity > 0 and (unit_price > 0 or amount > 0):
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
