"""Parse a factory proforma invoice (a spreadsheet exported to PDF) into
product rows: style number, photo, colors, sizes, quantity and unit price.

Pure functions only — no database, no storage — so the parser can be tested
against a real invoice on its own. PyMuPDF is imported lazily so the rest of
the app still starts if it's ever unavailable.
"""
import re
from decimal import Decimal, InvalidOperation
from difflib import get_close_matches
from io import BytesIO

from PIL import Image

MAX_PAGES = 80
MAX_ROWS = 600

# Column header aliases (lower-cased) -> logical column.
_HEADER_ALIASES = {
    'style': ('style no.', 'style no', 'style', 'item no.', 'item no', 'art no', 'model'),
    'pic': ('pic', 'picture', 'photo', 'image'),
    'color': ('color', 'colour'),
    'size': ('size', 'sizes'),
    'qty': ('quantity', 'qty', 'pcs'),
    'price': ('price', 'unit price'),
    'amount': ('amount', 'total'),
}

KNOWN_COLORS = [
    'BLACK', 'WHITE', 'GRAY', 'GREY', 'L-GRAY', 'D-GRAY', 'GREEN', 'BLUE', 'NAVY',
    'KHAKI', 'PINK', 'BEIGE', 'RED', 'YELLOW', 'BROWN', 'ORANGE', 'PURPLE', 'MAROON',
    'CREAM', 'OLIVE', 'APRICOT',
]

# Chinese color words seen on factory sheets -> English.
CN_COLORS = {
    '黑': 'BLACK', '黑色': 'BLACK', '白': 'WHITE', '白色': 'WHITE',
    '灰': 'GRAY', '灰色': 'GRAY', '浅灰': 'L-GRAY', '深灰': 'D-GRAY',
    '绿': 'GREEN', '绿色': 'GREEN', '军绿': 'ARMY GREEN', '蓝': 'BLUE', '蓝色': 'BLUE',
    '藏青': 'NAVY', '深蓝': 'NAVY', '红': 'RED', '红色': 'RED', '黄': 'YELLOW',
    '粉': 'PINK', '粉色': 'PINK', '卡其': 'KHAKI', '米': 'BEIGE', '米色': 'BEIGE',
    '米白': 'CREAM', '咖': 'BROWN', '咖啡': 'BROWN', '紫': 'PURPLE', '橙': 'ORANGE',
    '杏': 'APRICOT',
}

# One token is either "4C/50" (N assorted colors, M each — kept as written)
# or "<color name> <quantity>" with an optional space ("黑150", "浅灰 75").
_COLOR_TOKEN = re.compile(
    r'(\d+C/\d+)|([A-Za-z一-鿿][A-Za-z一-鿿\-]*)\s*(\d+)'
)
_CJK = re.compile(r'[一-鿿]')


class InvoiceError(Exception):
    """The file couldn't be read as a recognisable invoice (message is user-facing)."""


def normalize_colors(raw):
    """'BALCK 150\\nGRAY 150' / '黑150 浅灰 75' -> 'BLACK 150, GRAY 150'.
    Fixes obvious typos, translates common Chinese color words, keeps 4C/50
    style tokens as written. Falls back to the cleaned raw text."""
    text = ' '.join((raw or '').split())
    if not text:
        return ''
    parts = []
    for m in _COLOR_TOKEN.finditer(text):
        if m.group(1):
            parts.append(m.group(1))
            continue
        name, qty = m.group(2), m.group(3)
        if _CJK.search(name):
            name = CN_COLORS.get(name, name)
        else:
            name = name.upper()
            if name not in KNOWN_COLORS:
                close = get_close_matches(name, KNOWN_COLORS, n=1, cutoff=0.75)
                if close:
                    name = close[0]
        parts.append(f'{name} {qty}')
    return ', '.join(parts) if parts else text


def normalize_sizes(raw):
    """'M--3XL' -> 'M-3XL'."""
    text = ' '.join((raw or '').split())
    return re.sub(r'\s*-{1,3}\s*', '-', text)


def _to_decimal(text):
    cleaned = re.sub(r'[^\d.]', '', text or '')
    if not cleaned:
        return None
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def _to_int(text):
    cleaned = re.sub(r'[^\d]', '', text or '')
    return int(cleaned) if cleaned else None


def detect_currency(price_text):
    s = price_text or ''
    if re.search(r'[￥¥]', s) or re.search(r'RMB|CNY', s, re.I):
        return 'CNY'
    if '$' in s or re.search(r'USD', s, re.I):
        return 'USD'
    if '€' in s:
        return 'EUR'
    if '£' in s:
        return 'GBP'
    return None


def _column_map(header_cells):
    mapping = {}
    for idx, cell in enumerate(header_cells):
        label = (cell or '').strip().lower()
        for key, aliases in _HEADER_ALIASES.items():
            if key not in mapping and label in aliases:
                mapping[key] = idx
    return mapping


def _to_jpeg_bytes(data):
    """Normalise any embedded image to a plain RGB JPEG (flattens transparency)."""
    with Image.open(BytesIO(data)) as im:
        if im.mode in ('RGBA', 'LA', 'P'):
            im = im.convert('RGBA')
            bg = Image.new('RGB', im.size, 'white')
            bg.paste(im, mask=im.split()[-1])
            im = bg
        else:
            im = im.convert('RGB')
        out = BytesIO()
        im.save(out, 'JPEG', quality=90)
        return out.getvalue()


def _extract_photo(doc, pymupdf, xref):
    info = doc.extract_image(xref)
    if not info:
        return None
    data = info['image']
    if info.get('smask'):  # has a transparency mask: composite it properly
        pix = pymupdf.Pixmap(doc, xref)
        pix = pymupdf.Pixmap(pix, pymupdf.Pixmap(doc, info['smask']))
        data = pix.tobytes('png')
    try:
        return _to_jpeg_bytes(data)
    except Exception:
        return None


def parse_factory_pdf(pdf_bytes):
    """Return {'meta': {...}, 'rows': [...], 'warnings': [...]}.

    Each row: seq, style, colors, colors_raw, sizes, qty, unit_price (Decimal),
    amount (Decimal), photo (JPEG bytes or None)."""
    try:
        import pymupdf
    except ImportError as exc:  # pragma: no cover
        raise InvoiceError('PDF support is not installed on the server.') from exc

    try:
        doc = pymupdf.open(stream=pdf_bytes, filetype='pdf')
    except Exception as exc:
        raise InvoiceError('That file could not be opened as a PDF.') from exc
    if len(doc) > MAX_PAGES:
        raise InvoiceError(f'The PDF has too many pages (max {MAX_PAGES}).')

    meta, warnings, rows = {}, [], []
    colmap = None
    invoice_total_qty = None
    first_text = doc[0].get_text() if len(doc) else ''
    for key, pattern in (('invoice_no', r'INVOICE\s*NO\.?\s*:?\s*([^\s]+)'),
                         ('delivery', r'DELIVERY\s*DATE\s*:?\s*([\d\-/.]+)'),
                         ('customer', r'CUSTOMER\s*:?\s*([^\s]+)'),
                         ('supplier', r'SUPPLIER\s*:?\s*([^\s]+)')):
        m = re.search(pattern, first_text, re.I)
        if m:
            meta[key] = m.group(1).strip()

    for page in doc:
        tables = page.find_tables().tables
        if not tables:
            continue
        table = tables[0]
        data = table.extract()
        infos = page.get_image_info(xrefs=True)
        for rowdata, row in zip(data, table.rows):
            first = (rowdata[0] or '').strip()
            if colmap is None:
                if any((c or '').strip().lower() in _HEADER_ALIASES['style'] for c in rowdata):
                    colmap = _column_map(rowdata)
                continue                       # header (or pre-header) row
            if first.upper().startswith('TOTAL'):
                if 'qty' in colmap:
                    invoice_total_qty = _to_int(rowdata[colmap['qty']])
                continue
            if 'style' not in colmap:
                continue
            style = (rowdata[colmap['style']] or '').strip()
            if not style:
                continue
            if len(rows) >= MAX_ROWS:
                raise InvoiceError(f'The invoice has too many rows (max {MAX_ROWS}).')

            def cell(key):
                idx = colmap.get(key)
                return rowdata[idx] if idx is not None and idx < len(rowdata) else ''

            photo = None
            if 'pic' in colmap and colmap['pic'] < len(row.cells) and row.cells[colmap['pic']]:
                x0, y0, x1, y1 = row.cells[colmap['pic']]
                best = None
                for im in infos:
                    bx0, by0, bx1, by1 = im['bbox']
                    cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2
                    if x0 <= cx <= x1 and y0 <= cy <= y1:
                        area = (bx1 - bx0) * (by1 - by0)
                        if best is None or area > best[0]:
                            best = (area, im['xref'])
                if best:
                    photo = _extract_photo(doc, pymupdf, best[1])

            price_text = cell('price')
            if 'currency' not in meta:
                cur = detect_currency(price_text)
                if cur:
                    meta['currency'] = cur
            rows.append({
                'seq': _to_int(first),
                'style': style,
                'colors_raw': ' '.join((cell('color') or '').split()),
                'colors': normalize_colors(cell('color')),
                'sizes': normalize_sizes(cell('size')),
                'qty': _to_int(cell('qty')),
                'unit_price': _to_decimal(price_text),
                'amount': _to_decimal(cell('amount')),
                'photo': photo,
            })

    if colmap is None or 'style' not in colmap:
        raise InvoiceError('No product table with a STYLE NO. column was found in this PDF.')
    if not rows:
        raise InvoiceError('The table was found but contained no product rows.')

    parsed_total = sum(r['qty'] or 0 for r in rows)
    meta['total_qty_parsed'] = parsed_total
    if invoice_total_qty is not None:
        meta['total_qty_invoice'] = invoice_total_qty
        if invoice_total_qty != parsed_total:
            warnings.append(
                f'Quantities add up to {parsed_total}, but the invoice total says '
                f'{invoice_total_qty} — please review the rows.'
            )
    missing = [r['style'] for r in rows if not r['photo']]
    if missing:
        warnings.append(f'{len(missing)} row(s) have no photo and cannot be imported.')
    return {'meta': meta, 'rows': rows, 'warnings': warnings}


def make_preview_thumb(jpeg_bytes, size=160):
    with Image.open(BytesIO(jpeg_bytes)) as im:
        im = im.convert('RGB')
        im.thumbnail((size, size))
        out = BytesIO()
        im.save(out, 'JPEG', quality=70)
        return out.getvalue()
