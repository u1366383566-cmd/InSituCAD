"""Restore WPS DISPIMG cell pictures as standard, portable spreadsheet drawings."""
import io
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

from openpyxl.drawing.image import Image as SpreadsheetImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU

_OLE_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
_FORMULA = re.compile(r'^=(?:_xlfn\.)?DISPIMG\(\s*"([^"]+)"', re.I)
_NS = {'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
       'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
       'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}


def _picture_data(source_path):
    with open(source_path, 'rb') as f:
        magic = f.read(8)
    if magic == _OLE_MAGIC:
        try:
            import olefile
        except ImportError as exc:
            raise ValueError('缺少 olefile，无法提取旧版 WPS 表格图片，请安装 olefile') from exc
        with olefile.OleFileIO(source_path) as ole:
            if not ole.exists('ETCellImageData'):
                return {}
            source = io.BytesIO(ole.openstream('ETCellImageData').read())
    else:
        source = source_path
    with zipfile.ZipFile(source) as archive:
        try:
            # WPS versions use both cellImages.xml and cellimages.xml.
            # ZIP member names are case-sensitive even on Windows.
            members = {name.casefold(): name for name in archive.namelist()}
            xml = archive.read(members['xl/cellimages.xml'])
            rel_xml = archive.read(members['xl/_rels/cellimages.xml.rels'])
        except KeyError:
            return {}
        rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(rel_xml)
                if r.get('TargetMode') != 'External'}
        result = {}
        for entry in ET.fromstring(xml):
            name = entry.find('.//xdr:cNvPr', _NS)
            blip = entry.find('.//a:blip', _NS)
            if name is None or blip is None:
                continue
            target = rels.get(blip.get('{%s}embed' % _NS['r']))
            if target:
                path = posixpath.normpath(posixpath.join('xl', target.lstrip('/')))
                if target.startswith('/'): path = target.lstrip('/')
                result[name.get('name')] = archive.read(path)
        return result


def restore_wps_cell_images(workbook, source_path):
    """Replace resolvable DISPIMG formulas with images; fail if any are missing."""
    cells = []
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                match = _FORMULA.match(cell.value) if cell.data_type == 'f' and isinstance(cell.value, str) else None
                if match:
                    cells.append((sheet, cell, match.group(1)))
    if not cells:
        return 0
    pictures = _picture_data(source_path)
    missing = [f'{sheet.title}!{cell.coordinate}' for sheet, cell, key in cells if key not in pictures]
    if missing:
        raise ValueError('无法恢复 WPS 单元格图片：' + ', '.join(missing))
    for sheet, cell, key in cells:
        img = SpreadsheetImage(io.BytesIO(pictures[key]))
        col = sheet.column_dimensions[cell.column_letter].width or sheet.sheet_format.defaultColWidth or 8.43
        row = sheet.row_dimensions[cell.row].height or sheet.sheet_format.defaultRowHeight or 15
        width, height = max(1, col * 7 + 5 - 8), max(1, row * 96 / 72 - 8)
        scale = min(width / img.width, height / img.height)
        img.width *= scale
        img.height *= scale
        marker = AnchorMarker(col=cell.column - 1, row=cell.row - 1,
                              colOff=pixels_to_EMU((width + 8 - img.width) / 2),
                              rowOff=pixels_to_EMU((height + 8 - img.height) / 2))
        img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(
            pixels_to_EMU(img.width), pixels_to_EMU(img.height)))
        sheet.add_image(img)
        cell.value = None
    return len(cells)
