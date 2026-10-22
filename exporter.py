"""Excel cell selection and one-page PDF rendering."""
from io import BytesIO
import re
import zipfile
import pymupdf
from openpyxl.utils.cell import range_boundaries, get_column_letter

MAX_CELLS = 5000


def addresses(spec):
    result = []
    seen = set()
    for token in re.split(r'[,，;；\s]+', spec.strip().upper()):
        if not token:
            continue
        if not re.fullmatch(r'[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?', token):
            raise ValueError(f'无效的单元格或区域：{token}')
        left, top, right, bottom = range_boundaries(token)
        if right > 16384 or bottom > 1048576 or right < left or bottom < top:
            raise ValueError(f'区域超出 Excel 范围或顺序错误：{token}')
        if (right-left+1)*(bottom-top+1) > MAX_CELLS:
            raise ValueError(f'每次最多选择 {MAX_CELLS} 个单元格')
        for row in range(top, bottom+1):
            for col in range(left, right+1):
                address = f'{get_column_letter(col)}{row}'
                if address not in seen:
                    seen.add(address)
                    result.append(address)
                    if len(result) > MAX_CELLS:
                        raise ValueError(f'每次最多选择 {MAX_CELLS} 个单元格')
    if not result:
        raise ValueError('请输入至少一个单元格或区域')
    return result


def render_pdf(text, width_mm, height_mm, margin_mm=5, font_size=18,
               align='center', vertical='middle'):
    if not (10 <= width_mm <= 1000 and 10 <= height_mm <= 1000):
        raise ValueError('页面宽高必须在 10–1000 毫米之间')
    if margin_mm < 0 or margin_mm * 2 >= min(width_mm, height_mm):
        raise ValueError('两侧边距之和必须小于页面宽高')
    if not 4 <= font_size <= 200:
        raise ValueError('字号必须在 4–200 之间')
    scale = 72 / 25.4
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=width_mm*scale, height=height_mm*scale)
        rect = pymupdf.Rect(margin_mm*scale, margin_mm*scale,
                            (width_mm-margin_mm)*scale, (height_mm-margin_mm)*scale)
        # Built-in CJK font is embedded: output does not depend on reader fonts.
        font = pymupdf.Font('china-s')
        page.insert_font(fontname='cellfont', fontbuffer=font.buffer)
        normalized = str(text).replace('\r\n', '\n').replace('\r', '\n').replace('\t', '    ')
        def wrap(size):
            lines = []
            for paragraph in normalized.split('\n'):
                current = ''
                for ch in paragraph:
                    if font.text_length(ch, fontsize=size) > rect.width:
                        return None
                    if current and font.text_length(current+ch, fontsize=size) > rect.width:
                        lines.append(current)
                        current = ch
                    else:
                        current += ch
                lines.append(current)
            return lines
        size = float(font_size)
        while True:
            lines = wrap(size)
            line_height = size * 1.45
            if lines is not None and len(lines)*line_height <= rect.height:
                break
            size -= 0.5
            if size < 4:
                raise ValueError('内容过长，缩小至 4pt 后仍无法放入一页；请增大页面或减少内容')
        block_height = len(lines)*line_height
        offset = {'top': 0, 'middle': (rect.height-block_height)/2,
                  'bottom': rect.height-block_height}[vertical]
        for index, line in enumerate(lines):
            length = font.text_length(line, fontsize=size)
            x = rect.x0 + {'left': 0, 'center': (rect.width-length)/2,
                           'right': rect.width-length}[align]
            page.insert_text((x, rect.y0+offset+size+index*line_height), line,
                             fontname='cellfont', fontsize=size)
        return doc.tobytes(garbage=4, deflate=True)
    finally:
        doc.close()


def export_zip(workbook, sheets, spec, skip_empty=True, **options):
    selected = addresses(spec)
    if len(selected)*len(sheets) > MAX_CELLS:
        raise ValueError(f'所有工作表合计最多选择 {MAX_CELLS} 个单元格')
    output = BytesIO()
    count = 0
    skipped = 0
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for sheet_index, name in enumerate(sheets, 1):
            sheet = workbook[name]
            safe = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', name)
            for address in selected:
                value = sheet[address].value
                if value is None or value == '':
                    if skip_empty:
                        skipped += 1
                        continue
                    value = ''
                try:
                    pdf = render_pdf(str(value), **options)
                except ValueError as exc:
                    raise ValueError(f'{name}!{address}：{exc}') from exc
                archive.writestr(f'{sheet_index:02d}_{safe}/{address}.pdf', pdf)
                count += 1
    return output.getvalue(), count, skipped
