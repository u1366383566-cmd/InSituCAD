import contextlib
import io
import os
import tempfile
import unittest
from unittest.mock import patch

import fitz
import cad_pdf_translator as translator


class ClippedHeaderTest(unittest.TestCase):
    def test_clipped_header_backfill_stays_inside_visible_band(self):
        with tempfile.TemporaryDirectory() as directory:
            source = os.path.join(directory, 'source.pdf')
            output = os.path.join(directory, 'output.pdf')
            doc = fitz.open()
            page = doc.new_page(width=200, height=120)
            page.draw_rect(fitz.Rect(20, 40, 120, 60), color=None, fill=(0.75, 0.75, 0.75))
            page.insert_text((55, 50), 'Original', fontsize=8)
            page.draw_rect(fitz.Rect(0, 0, 200, 50), color=None, fill=(1, 1, 1))
            doc.save(source)
            doc.close()
            rect = fitz.Rect(55, 43, 85, 53)
            boxes = [{'rect': rect, 'text': 'Original'}]
            translated = ([{'id': 0, 'original': 'Original', 'translation': 'Model'}],
                          {'prompt_tokens': 0, 'completion_tokens': 0})
            with patch.object(translator, 'extract_text_lines', return_value=(boxes, 8)), \
                 patch.object(translator, 'remote_ocr', return_value=[]), \
                 patch.object(translator, 'translate_text_blocks', return_value=translated), \
                 patch.object(translator, 'load_cache', return_value={}), \
                 patch.object(translator, 'save_cache'), patch.object(translator, 'bump_cache_counts'), \
                 contextlib.redirect_stdout(io.StringIO()):
                result = translator.process_pdf(source, output, target_lang='en', provider={'name': 'test'})
            self.assertTrue(result['success'])
            with fitz.open(output) as rendered:
                page = rendered[0]
                pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=False)
                offset = ((46 * 4) * pix.width + 60 * 4) * pix.n
                self.assertTrue(all(x > 240 for x in pix.samples[offset:offset + 3]))
                text = page.search_for('Model')
                self.assertEqual(len(text), 1)
                self.assertGreater((text[0].y0 + text[0].y1) / 2, 50)
                self.assertLess((text[0].y0 + text[0].y1) / 2, 60)


if __name__ == '__main__':
    unittest.main()
