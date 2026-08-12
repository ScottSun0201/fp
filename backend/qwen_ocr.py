"""Qwen OCR client for scanned invoices and supplier statements."""
import base64
import json
import logging
import re
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path
from PIL import Image as PILImage

from config import (
    QWEN_OCR_API_KEY,
    QWEN_OCR_API_URL,
    QWEN_OCR_MAX_PAGES,
    QWEN_OCR_MODEL,
)

logger = logging.getLogger(__name__)


STATEMENT_PROMPT = """你是安徽骊威科技有限公司的财务对账单识别程序。读取图片中的全部内容，只输出一个合法 JSON 对象，不要 Markdown。
无法确认的文字填空字符串，无法确认的数字填 null，禁止猜测。
字段角色规则：安徽骊威科技有限公司（以及名称中包含“骊威”的公司）是本系统的客户/采购方，
绝不能填入 supplier_name。supplier_name 必须是与骊威对账的对方供货公司/销售方；
若图片未明确显示对方公司名称，supplier_name 留空，不得把骊威复制进去。
JSON 格式：
{"supplier_name":"","supplier_code":"","statement_month":"YYYY-MM","statement_date":"YYYY-MM-DD",
"column_mapping":{"customer_order_no":"","customer_material_code":"","delivery_no":"","delivery_date":"",
"product_name":"","specification":"","quantity":"","unit":"","unit_price_incl_tax":"","amount_incl_tax":""},
"usage_remark":"","total_invoice_amount":null,"items":[
{"seq":null,"customer_order_no":"","customer_material_code":"","delivery_no":"","delivery_date":"YYYY-MM-DD或YYYY-MM",
"product_name":"","specification":"","quantity":null,"unit":"","unit_price_incl_tax":null,"amount_incl_tax":null}
]}
逐行提取表格明细，不要合并或漏掉明细行；金额、数量和单价只输出数字。
customer_order_no 是采购订单号，customer_material_code 是物料编码，delivery_no 是送货单号，三者不能混填。
客户物料编码重点按值识别：通常是完整单元格以 LW 开头、后跟数字（例如 LW130000031），
无论原表头叫“物料编码、客户货号、编号、料号”都填入 customer_material_code。
AHLW- 开头的是采购订单号，不是物料编码；合作商编码虽然也可能以 LW 开头，但位于公司/合作商信息区，
不能作为明细物料编码。单位、数量、含税单价、含税金额是每条有效物料明细必须重点提取的字段。
字段必须按表头和业务角色提取：日期只能填 delivery_date；品名只能填 product_name；型号、尺寸、参数只能填 specification；
数量、单价、金额分别对应原表的数量列、含税单价列、含税金额列。某列不存在就将对应字段留空或填 null，
不得用其他列内容补齐，不得用数量乘单价反推金额，也不得用金额除以数量反推单价。
日期列只有年月时按 YYYY-MM 输出，不得擅自补成每月1日。
先读取表头，再填写 column_mapping：键是系统字段，值是图片中与其对应的原始列标题；原表没有该列时值为空。
items 中的每个值只能来自 column_mapping 指定的那一列。合计、总计、期初欠款、历史应收、备注行不得作为普通物料明细。
total_invoice_amount 必须是全部明细 amount_incl_tax 的合计。"""


INVOICE_PROMPT = """你是增值税发票识别程序。读取图片中的全部内容，只输出一个合法 JSON 对象，不要 Markdown。
无法确认的文字填空字符串，无法确认的数字填 null，禁止猜测。
JSON 格式：
{"invoice_number":"","invoice_date":"YYYY-MM-DD","invoice_type":"","buyer_name":"","buyer_tax_id":"",
"seller_name":"","seller_tax_id":"","total_amount_excl":null,"total_tax":null,"total_amount_incl":null,
"amount_capital":"","items":[
{"category_prefix":"","material_name":"","specification":"","unit":"","quantity":null,
"unit_price_excl":null,"amount_excl":null,"tax_rate":null,"tax_amount":null}
]}
逐行提取商品明细，金额、数量和税率只输出数字。"""

SUPPLIER_PROMPT = """识别对账单中的供货商公司名称，只输出合法 JSON：
{"supplier_name":""}
名称中包含“安徽骊威”或“骊威”的公司是采购方/收货方，不能作为供货商。
优先读取“供货商、供应商、供货单位、协力厂商、销售方”等标签对应的完整公司名称。
无法确认时 supplier_name 留空，禁止猜测。"""


def is_configured() -> bool:
    return bool(QWEN_OCR_API_KEY and QWEN_OCR_API_URL)


def recognize_pdf(pdf_path: str, document_type: str, progress_callback=None) -> dict:
    """Render a PDF and merge Qwen OCR JSON returned for every page."""
    if not is_configured():
        return {}
    prompt = INVOICE_PROMPT if document_type == 'invoice' else STATEMENT_PROMPT
    pages = _render_pdf(
        pdf_path,
        tile_for_tables=(document_type == 'statement'),
    )
    if not pages:
        return {}
    results = []
    errors = []
    logger.info(
        "开始识别文档 model=%s type=%s pages=%d",
        QWEN_OCR_MODEL, document_type, len(pages),
    )
    for page_number, page_path in enumerate(pages, start=1):
        started_at = time.monotonic()
        if progress_callback:
            progress_callback(page_number - 1, len(pages), None)
        try:
            result = _recognize_image(page_path, prompt)
            if result:
                results.append(result)
                logger.info(
                    "页面识别成功 model=%s page=%d/%d elapsed=%.1fs items=%d",
                    QWEN_OCR_MODEL, page_number, len(pages),
                    time.monotonic() - started_at,
                    len(result.get('items') or []),
                )
                if progress_callback:
                    progress_callback(page_number, len(pages), len(result.get('items') or []))
        except Exception as exc:
            logger.warning(
                "页面识别失败 model=%s page=%d/%d elapsed=%.1fs error=%s",
                QWEN_OCR_MODEL, page_number, len(pages),
                time.monotonic() - started_at, exc,
            )
            errors.append(str(exc))
            if progress_callback:
                progress_callback(page_number, len(pages), 0)
    if not results and errors:
        raise RuntimeError(f"千问 OCR 调用失败: {errors[0]}")
    return _merge_pages(results, document_type)


def recognize_image(image_path: str, document_type: str) -> dict:
    """Recognize one uploaded image with the same schema as a PDF page."""
    if not is_configured():
        return {}
    prompt = INVOICE_PROMPT if document_type == 'invoice' else STATEMENT_PROMPT
    source = Path(image_path)
    if not source.is_file() or source.stat().st_size == 0:
        raise RuntimeError("图片文件为空或不存在")
    logger.info("开始识别图片 model=%s type=%s file=%s", QWEN_OCR_MODEL, document_type, source.name)
    result = _recognize_image(source, prompt)
    return _merge_pages([result] if result else [], document_type)


def recognize_supplier(file_path: str) -> str:
    """Low-token first-stage supplier recognition using only one image/page."""
    if not is_configured():
        return ""
    source = Path(file_path)
    if source.suffix.lower() == '.pdf':
        pages = _render_pdf(str(source), max_pages=1)
        if not pages:
            return ""
        source = pages[0]
    result = _recognize_image(source, SUPPLIER_PROMPT, max_tokens=256)
    return str((result or {}).get("supplier_name") or "").strip()


def _render_pdf(pdf_path: str, max_pages=None, tile_for_tables=False) -> list[Path]:
    try:
        import pypdfium2 as pdfium
    except ImportError:
        logger.warning("pypdfium2 is unavailable; Qwen OCR cannot render PDF")
        return []
    output_dir = Path(tempfile.mkdtemp(prefix='fp_qwen_ocr_'))
    paths = []
    document_header = None
    pdf = pdfium.PdfDocument(pdf_path)
    for index, page in enumerate(pdf):
        if index >= (max_pages or QWEN_OCR_MAX_PAGES):
            break
        image = page.render(scale=2.0).to_pil().convert('RGB')
        if tile_for_tables:
            image = _prepare_table_image(image)
            if document_header is None:
                header_height = min(260, max(120, image.height // 3))
                document_header = image.crop(
                    (0, 0, image.width, header_height)
                )
        if tile_for_tables and image.height > 900:
            # A full A4 page may contain 50+ table rows. Asking one model
            # response to return all rows can truncate JSON. Split vertically
            # into overlapping bands so every response stays comfortably
            # below its output limit, then merge and deduplicate the rows.
            tile_height = 420
            overlap = 60
            top = 0
            part = 1
            while top < image.height:
                bottom = min(image.height, top + tile_height)
                path = output_dir / f'page_{index + 1}_part_{part}.jpg'
                tile = image.crop((0, top, image.width, bottom))
                if paths and document_header is not None:
                    tile = _prepend_table_header(tile, document_header)
                tile.save(
                    path, format='JPEG', quality=88, optimize=True
                )
                paths.append(path)
                if bottom >= image.height:
                    break
                top = bottom - overlap
                part += 1
        else:
            path = output_dir / f'page_{index + 1}.jpg'
            tile = image
            if tile_for_tables and paths and document_header is not None:
                tile = _prepend_table_header(tile, document_header)
            tile.save(path, format='JPEG', quality=85, optimize=True)
            paths.append(path)
    return paths


def _prepend_table_header(tile, header):
    if header.width != tile.width:
        target_height = max(1, round(header.height * tile.width / header.width))
        header = header.resize((tile.width, target_height))
    combined = PILImage.new(
        'RGB', (tile.width, header.height + tile.height), 'white'
    )
    combined.paste(header, (0, 0))
    combined.paste(tile, (0, header.height))
    return combined


def _prepare_table_image(image):
    """Crop page whitespace and rotate sideways dense tables before tiling."""
    grayscale = image.convert('L')
    dark_mask = grayscale.point(lambda value: 255 if value < 185 else 0)
    bbox = dark_mask.getbbox()
    if not bbox:
        return image
    left, top, right, bottom = bbox
    margin = 24
    left = max(0, left - margin)
    top = max(0, top - margin)
    right = min(image.width, right + margin)
    bottom = min(image.height, bottom + margin)
    image = image.crop((left, top, right, bottom))
    mask = dark_mask.crop((left, top, right, bottom))

    sample = mask.copy()
    sample.thumbnail((400, 400))
    width, height = sample.size
    pixels = list(sample.getdata())
    row_hits = [
        sum(1 for value in pixels[y * width:(y + 1) * width] if value)
        for y in range(height)
    ]
    column_hits = [
        sum(1 for y in range(height) if pixels[y * width + x])
        for x in range(width)
    ]

    def grouped_line_count(indices):
        groups = 0
        previous = -2
        for value in indices:
            if value > previous + 1:
                groups += 1
            previous = value
        return groups

    horizontal_lines = grouped_line_count([
        index for index, count in enumerate(row_hits)
        if count >= width * 0.5
    ])
    vertical_lines = grouped_line_count([
        index for index, count in enumerate(column_hits)
        if count >= height * 0.5
    ])
    sideways = (
        image.height > image.width * 1.25
        or (
            vertical_lines >= 8
            and vertical_lines > max(1, horizontal_lines) * 1.5
        )
    )
    if sideways:
        image = image.transpose(PILImage.Transpose.ROTATE_270)
    return image


def _recognize_image(image_path: Path, prompt: str, max_tokens: int = 8192) -> dict:
    encoded = base64.b64encode(image_path.read_bytes()).decode('ascii')
    mime_type = {
        '.png': 'image/png',
        '.webp': 'image/webp',
    }.get(image_path.suffix.lower(), 'image/jpeg')
    content = [
        {
            "type": "image_url",
            "image_url": {"url": f"data:{mime_type};base64,{encoded}"},
        },
        {"type": "text", "text": prompt},
    ]
    if 'chat/completions' in QWEN_OCR_API_URL:
        payload = {
            "model": QWEN_OCR_MODEL,
            "messages": [{
                "role": "user",
                "content": content,
            }],
            "temperature": 0,
            "max_tokens": max_tokens,
            "enable_thinking": False,
        }
    else:
        payload = {
            "model": QWEN_OCR_MODEL,
            "input": {
                "messages": [{
                    "role": "user",
                    "content": [
                        {"image": f"data:{mime_type};base64,{encoded}",
                         "min_pixels": 3072, "max_pixels": 8388608,
                         "enable_rotate": True},
                        {"text": prompt},
                    ],
                }]
            },
            "parameters": {"temperature": 0, "max_tokens": max_tokens},
        }
    request = urllib.request.Request(
        QWEN_OCR_API_URL,
        data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
        headers={
            "Authorization": f"Bearer {QWEN_OCR_API_KEY}",
            "Content-Type": "application/json",
        },
        method='POST',
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            body = json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode('utf-8', errors='replace')[:500]
        raise RuntimeError(f"HTTP {exc.code}: {detail}") from exc
    text = _response_text(body)
    return _parse_json(text)


def _response_text(body: dict) -> str:
    choices = body.get('choices') or body.get('output', {}).get('choices', [])
    if not choices:
        raise RuntimeError(f"千问 OCR 返回格式异常: {str(body)[:300]}")
    content = choices[0].get('message', {}).get('content', [])
    if isinstance(content, str):
        return content
    return '\n'.join(
        str(item.get('text', '')) for item in content if isinstance(item, dict)
    )


def _parse_json(text: str) -> dict:
    cleaned = re.sub(r'^```(?:json)?\s*|\s*```$', '', text.strip(), flags=re.I)
    try:
        value = json.loads(cleaned)
    except json.JSONDecodeError:
        start, end = cleaned.find('{'), cleaned.rfind('}')
        if start < 0 or end <= start:
            raise RuntimeError("千问 OCR 未返回 JSON")
        value = json.loads(cleaned[start:end + 1])
    if not isinstance(value, dict):
        raise RuntimeError("千问 OCR JSON 顶层不是对象")
    return value


def _merge_pages(pages: list[dict], document_type: str) -> dict:
    if not pages:
        return {}
    merged = {}
    items = []
    for source_index, page in enumerate(pages):
        for key, value in page.items():
            if key == 'items':
                if isinstance(value, list):
                    items.extend(
                        (source_index, item)
                        for item in value if isinstance(item, dict)
                    )
            elif key not in merged or merged[key] in ('', None, 0):
                merged[key] = value
    deduplicated = []
    seen = {}
    for source_index, item in items:
        signature = tuple(
            re.sub(r'\s+', '', str(item.get(key) or '')).upper()
            for key in (
                'customer_order_no', 'customer_material_code', 'delivery_no',
                'quantity', 'unit_price_incl_tax', 'amount_incl_tax',
            )
        )
        # Overlapping image bands intentionally repeat boundary rows. Remove
        # a repeated business row only when it came from a different band;
        # preserve genuinely duplicated rows that coexist in the same band.
        if any(signature) and signature in seen and seen[signature] != source_index:
            continue
        if any(signature):
            seen.setdefault(signature, source_index)
        deduplicated.append(item)
    for index, item in enumerate(deduplicated, start=1):
        item['seq'] = index
    merged['items'] = deduplicated
    merged['ocr_provider'] = 'qwen'
    merged['raw_text'] = json.dumps(pages, ensure_ascii=False)
    return merged
