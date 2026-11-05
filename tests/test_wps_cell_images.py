import io
import os
import tempfile
import unittest
import zipfile

import openpyxl
from PIL import Image

from wps_cell_images import restore_wps_cell_images


class WpsImagesTest(unittest.TestCase):
    def test_wps_lowercase_image_parts(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, 'source.xlsx')
            image = io.BytesIO()
            Image.new('RGB', (40, 20), 'red').save(image, format='PNG')
            with zipfile.ZipFile(path, 'w') as archive:
                archive.writestr('xl/cellimages.xml', '''<images xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><image><x:cNvPr name="ID_1"/><a:blip r:embed="rId1"/></image></images>''')
                archive.writestr('xl/_rels/cellimages.xml.rels', '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/image1.png"/></Relationships>''')
                archive.writestr('xl/media/image1.png', image.getvalue())
            book = openpyxl.Workbook()
            book.active['A1'] = '=DISPIMG("ID_1",1)'
            self.assertEqual(restore_wps_cell_images(book, path), 1)
            self.assertEqual(len(book.active._images), 1)
            self.assertIsNone(book.active['A1'].value)

    def test_restore_repeated_picture_and_preserve_existing(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, 'source.xlsx')
            image = io.BytesIO()
            Image.new('RGB', (40, 20), 'red').save(image, format='PNG')
            with zipfile.ZipFile(path, 'w') as archive:
                archive.writestr('xl/cellImages.xml', '''<images xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><image><x:cNvPr name="ID_1"/><a:blip r:embed="rId1"/></image></images>''')
                archive.writestr('xl/_rels/cellImages.xml.rels', '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="media/image1.png"/></Relationships>''')
                archive.writestr('xl/media/image1.png', image.getvalue())
            book = openpyxl.Workbook()
            sheet = book.active
            sheet['E4'] = '=DISPIMG("ID_1",1)'
            sheet['E5'] = '=_xlfn.DISPIMG("ID_1",1)'
            sheet.add_image(openpyxl.drawing.image.Image(io.BytesIO(image.getvalue())), 'A1')
            self.assertEqual(restore_wps_cell_images(book, path), 2)
            output = os.path.join(directory, 'output.xlsx')
            book.save(output)
            reloaded = openpyxl.load_workbook(output)
            self.assertEqual(len(reloaded.active._images), 3)
            self.assertIsNone(reloaded.active['E4'].value)
            self.assertIsNone(reloaded.active['E5'].value)

    def test_missing_image_fails_without_erasing_formula(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, 'source.xlsx')
            with zipfile.ZipFile(path, 'w'):
                pass
            book = openpyxl.Workbook()
            book.active['E4'] = '=DISPIMG("missing",1)'
            with self.assertRaisesRegex(ValueError, 'E4'):
                restore_wps_cell_images(book, path)
            self.assertEqual(book.active['E4'].value, '=DISPIMG("missing",1)')


if __name__ == '__main__':
    unittest.main()
