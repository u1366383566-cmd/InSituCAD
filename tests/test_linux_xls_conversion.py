import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

import openpyxl
import cad_pdf_translator as translator


@unittest.skipUnless(os.name != 'nt' and shutil.which('libreoffice'), 'requires Linux LibreOffice')
class LegacyExcelConversionTest(unittest.TestCase):
    def test_converts_binary_workbook_even_with_xlsx_extension(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = openpyxl.Workbook()
            workbook.active['A1'] = '家具报价'
            workbook.active['B1'] = 168
            source = os.path.join(directory, 'original.xlsx')
            workbook.save(source)
            subprocess.run([shutil.which('libreoffice'), '-env:UserInstallation=' + Path(directory, 'profile').as_uri(), '--headless', '--convert-to', 'xls:MS Excel 97', '--outdir', directory, source], check=True, capture_output=True, timeout=60)
            renamed = os.path.join(directory, 'legacy.xlsx')
            shutil.copyfile(os.path.join(directory, 'original.xls'), renamed)
            output = os.path.join(directory, 'converted.xlsx')
            success, message = translator.convert_xls_to_xlsx(renamed, output)
            self.assertTrue(success, message)
            result = openpyxl.load_workbook(output)
            self.assertEqual(result.active['A1'].value, '家具报价')
            self.assertEqual(result.active['B1'].value, 168)
            result.close()


if __name__ == '__main__':
    unittest.main()
