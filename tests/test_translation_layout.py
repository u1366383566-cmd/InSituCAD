import unittest
from unittest.mock import Mock

import fitz
import cad_pdf_translator as translator


class TranslationLayoutTest(unittest.TestCase):
    def test_trailing_chinese_character_joins_previous_line(self):
        page = Mock()
        page.get_text.return_value = {'blocks': [{'lines': [
            {'bbox': (10, 10, 60, 16), 'spans': [{'text': '高密度海'}]},
            {'bbox': (32, 17, 37, 23), 'spans': [{'text': '绵'}]},
        ]}]}
        units, _ = translator.extract_text_lines(page, min_chars=0)
        self.assertEqual([u['text'] for u in units], ['高密度海绵'])
        self.assertEqual(units[0]['line_h'], 6)
        self.assertEqual(units[0]['rect'], fitz.Rect(10, 10, 60, 23))

    def test_single_character_without_preceding_line_is_not_discarded(self):
        page = Mock()
        page.get_text.return_value = {'blocks': [{'lines': [
            {'bbox': (10, 10, 16, 16), 'spans': [{'text': '门'}]},
        ]}]}
        units, _ = translator.extract_text_lines(page, min_chars=0)
        self.assertEqual([u['text'] for u in units], ['门'])


if __name__ == '__main__':
    unittest.main()
