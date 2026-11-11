"""Регрессионные тесты Заполнятора на полностью синтетических ODS/XLSX.

Запуск из корня: python -m unittest discover -s tests -v
Реальные производственные документы не используются.
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from unittest.mock import patch
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import fill_signals_ods as z  # noqa: E402

ODS_T = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
ODS_X = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

LEGACY = {
    "%mobPos%": 3,
    "%rack%": 5,
    "%uso%": 10,
    "%XT%": 16,
    "%chanel%": 17,
    "%ID%": 18,
    "%caption%": 19,
    "%!LinkType?%": 20,
    "%!Link?%": 21,
}
MODERN = {
    "%mobPos%": 3,
    "%rack%": 5,
    "%uso%": 10,
    "%chanel%": 18,
    "%XT%": 19,
    "%ID%": 20,
    "%caption%": 21,
    "%!LinkType?%": 22,
    "%!Link?%": 23,
}
SHIFTED = {marker: index + 4 for marker, index in MODERN.items()}


def make_ods(target: Path, marker_cols: dict[str, int], *, duplicate=False, missing=None):
    """Создать простой ODS вручную как OpenDocument ZIP (только синтетические данные)."""
    maxcol = max(marker_cols.values()) + (1 if duplicate else 0)
    header = {col: marker for marker, col in marker_cols.items() if marker != missing}
    if duplicate:
        header[maxcol] = "%ID%"

    row5 = {
        marker_cols["%mobPos%"]: "3",
        marker_cols["%rack%"]: "A1",
        marker_cols["%uso%"]: "TEST.УСО.4(1)",
        marker_cols["%chanel%"]: "",
        marker_cols["%!LinkType?%"]: "idDiag",
    }
    row6 = {
        marker_cols["%mobPos%"]: "3",
        marker_cols["%rack%"]: "A1",
        marker_cols["%uso%"]: "TEST.УСО.4(1)",
        marker_cols["%chanel%"]: "10",
        marker_cols["%!LinkType?%"]: "idAI",
        marker_cols["%XT%"]: "OLD_XT",
        marker_cols["%ID%"]: "OLD_ID",
        marker_cols["%caption%"]: "OLD_CAPTION",
    }
    if "%!Link?%" in marker_cols:
        row6[marker_cols["%!Link?%"]] = "PRESERVE_LINK"
    row7 = {
        marker_cols["%chanel%"]: "11",
        marker_cols["%!LinkType?%"]: "idDI",
    }

    def row_xml(values):
        return '<table:table-row>' + ''.join(
            '<table:table-cell office:value-type="string"><text:p>'
            + escape(str(values.get(k, "")))
            + '</text:p></table:table-cell>'
            for k in range(1, maxcol + 1)
        ) + '</table:table-row>'

    rows = [row_xml({}), row_xml({}), row_xml(header), row_xml({}),
            row_xml(row5), row_xml(row6), row_xml(row7)]
    content = ('''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2">
<office:body><office:spreadsheet><table:table table:name="Сигналы">'''
               + ''.join(rows) + '</table:table></office:spreadsheet></office:body></office:document-content>')
    manifest = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>'''
    with ZipFile(target, 'w') as zout:
        zout.writestr('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', compress_type=ZIP_STORED)
        zout.writestr('content.xml', content, compress_type=ZIP_DEFLATED)
        zout.writestr('META-INF/manifest.xml', manifest, compress_type=ZIP_DEFLATED)


def make_xlsx(target: Path):
    """Сгенерировать минимальный XLSX; строка с одним совпадающим AI и одним дубликатом."""
    data = [
        (3, {'C': 'Идентификатор', 'D': 'Наименование', 'E': 'Тип схемы',
             'F': 'Станция', 'G': 'Слот', 'H': 'Канал', 'M': 'Клеммник'}),
        (5, {'C': 'TAG_001', 'D': 'SYNTHETIC_CAPTION', 'E': 'AI12(0)',
             'F': 'А1', 'G': '3', 'H': '10', 'M': 'XT_001'}),
        (6, {'C': 'TAG_DUPLICATE', 'D': 'NOT_USED', 'E': 'AI12(0)',
             'F': 'А1', 'G': '3', 'H': '10', 'M': 'BAD_XT'}),
    ]
    rows = []
    for num, cols in data:
        cells = ''.join(
            f'<c r="{col}{num}" t="inlineStr"><is><t>{escape(val)}</t></is></c>'
            for col, val in cols.items()
        )
        rows.append(f'<row r="{num}">{cells}</row>')
    sheet = f'''<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="{MAIN}"><sheetData>{''.join(rows)}</sheetData></worksheet>'''
    workbook = f'''<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="{MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Печать" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    ct = '''<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>'''
    rootrel = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''
    wbrel = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>'''
    with ZipFile(target, 'w') as zout:
        for name, content in (('[Content_Types].xml', ct), ('_rels/.rels', rootrel),
                              ('xl/workbook.xml', workbook), ('xl/_rels/workbook.xml.rels', wbrel),
                              ('xl/worksheets/sheet1.xml', sheet)):
            zout.writestr(name, content, compress_type=ZIP_DEFLATED)


def read_ods_rows(path: Path):
    with ZipFile(path) as zfile:
        root = ET.fromstring(zfile.read('content.xml'))
    out = []
    for row in root.findall('.//{' + ODS_T + '}table-row'):
        vals = []
        for cell in row:
            if cell.tag not in ('{' + ODS_T + '}table-cell', '{' + ODS_T + '}covered-table-cell'):
                continue
            value = ''.join(cell.itertext()).strip()
            repeat = int(cell.get('{' + ODS_T + '}number-columns-repeated', '1'))
            vals.extend([value]*min(repeat, 200))
        out.append(vals)
    return out


class HeaderBasedMappingTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix='zapolnyator_synthetic_')
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        make_xlsx(self.directory / '8. Шкаф TEST.УСО.4(1) - ТБ.xlsx')
        self.ods = self.directory / 'SYNTHETIC.ods'

    def test_both_templates_and_future_column_shift(self):
        for name, cols in (('legacy', LEGACY), ('modern', MODERN), ('shifted', SHIFTED)):
            with self.subTest(template=name):
                make_ods(self.ods, cols)
                stats = z.fill_ods(self.ods)
                self.assertEqual(stats['rows_filled'], 1)
                self.assertEqual(stats['rows_skipped_diag'], 1)
                self.assertEqual(stats['duplicate_keys'], 1)
                self.assertEqual(stats['backup_created'], 1)
                rows = read_ods_rows(self.ods)
                self.assertEqual(rows[5][cols['%XT%']-1], 'XT_001')
                self.assertEqual(rows[5][cols['%ID%']-1], 'TAG_001')
                self.assertEqual(rows[5][cols['%caption%']-1], 'SYNTHETIC_CAPTION')
                self.assertEqual(rows[5][cols['%!LinkType?%']-1], 'idAI')
                self.assertEqual(rows[5][cols['%!Link?%']-1], 'PRESERVE_LINK')
                backups = sorted(self.directory.glob('SYNTHETIC.backup_before_fill_*.ods'))
                self.assertTrue(backups)
                old_rows = read_ods_rows(backups[-1])
                self.assertEqual(old_rows[5][cols['%XT%']-1], 'OLD_XT')
                self.assertEqual(old_rows[5][cols['%ID%']-1], 'OLD_ID')
                self.assertEqual(old_rows[5][cols['%caption%']-1], 'OLD_CAPTION')

    def test_missing_marker_stops_before_any_write(self):
        make_ods(self.ods, MODERN, missing='%XT%')
        before = self.ods.read_bytes()
        with self.assertRaisesRegex(RuntimeError, '%XT%'):
            z.fill_ods(self.ods)
        self.assertEqual(before,self.ods.read_bytes())
        self.assertFalse((self.directory/'SYNTHETIC.backup_before_fill.ods').exists())

    def test_duplicate_marker_stops_before_any_write(self):
        make_ods(self.ods, MODERN, duplicate=True)
        before = self.ods.read_bytes()
        with self.assertRaisesRegex(RuntimeError, '%ID%'):
            z.fill_ods(self.ods)
        self.assertEqual(before,self.ods.read_bytes())
        self.assertFalse((self.directory/'SYNTHETIC.backup_before_fill.ods').exists())


    def test_backup_preserves_previous_versions_and_no_backup_flag(self):
        make_ods(self.ods, MODERN)
        original = self.ods.read_bytes()

        # Принудительно одинаковая секунда для двух последовательных запусков.
        with patch.object(z, 'datetime') as mock_datetime:
            mock_datetime.now.return_value.strftime.return_value = '2026-11-11_16-05-30'
            stats1 = z.fill_ods(self.ods)
            backups1 = set(self.directory.glob('SYNTHETIC.backup_before_fill_*.ods'))
            self.assertEqual(len(backups1), 1)
            first_backup = next(iter(backups1))
            self.assertEqual(first_backup.read_bytes(), original)
            self.assertTrue(first_backup.name.endswith('backup_before_fill_2026-11-11_16-05-30.ods'))
            first_backup_bytes = first_backup.read_bytes()
            after_first_run = self.ods.read_bytes()

            stats2 = z.fill_ods(self.ods)
            backups2 = set(self.directory.glob('SYNTHETIC.backup_before_fill_*.ods'))
        self.assertEqual(len(backups2), 2)
        second_backup = next(iter(backups2 - backups1))
        self.assertEqual(second_backup.read_bytes(), after_first_run)
        self.assertTrue(second_backup.name.endswith('backup_before_fill_2026-11-11_16-05-30_001.ods'))
        self.assertEqual(first_backup.read_bytes(), first_backup_bytes)
        self.assertNotEqual(first_backup, second_backup)

        stats3 = z.fill_ods(self.ods, create_backup=False)
        backups3 = set(self.directory.glob('SYNTHETIC.backup_before_fill_*.ods'))
        self.assertEqual(backups2, backups3)
        self.assertEqual(stats3.get('backup_created', 0), 0)


if __name__ == '__main__':
    unittest.main()
