import tempfile
import unittest
from pathlib import Path
from scripts.release_notes import generate, read_history


class ReleaseNotesTests(unittest.TestCase):
    def test_current_notes_do_not_include_history_and_preserve_wording(self):
        text = 'v4.0\n\n- 保留我確認過的用詞 & 標點。\n\nv3.2.7\n\n- 舊版內容\n'
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            source = folder / 'notes.md'
            source.write_text(text)
            full = generate('4.0', source, folder / 'dist')
            self.assertEqual(full, text)
            self.assertEqual((folder / 'dist/GITHUB_RELEASE_NOTES.md').read_text(),
                             'v4.0\n\n- 保留我確認過的用詞 & 標點。\n')
            self.assertEqual((folder / 'dist/APP_RELEASE_NOTES.md').read_text(), text)
            with self.assertRaises(ValueError):
                generate('3.2.7', source, folder / 'dist')

    def test_reject_invalid_history_before_publishing(self):
        invalid = ['v4.0\n- 缺空行\n', 'v4.0\n\n沒有條列\n',
                   'v3.2.7\n\n- 舊\n\nv4.0\n\n- 新\n',
                   'v4.0\n\n- 新\n\nv4.0\n\n- 重複\n']
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'notes.md'
            for text in invalid:
                source.write_text(text)
                with self.subTest(text=text), self.assertRaises(ValueError):
                    read_history(source)
