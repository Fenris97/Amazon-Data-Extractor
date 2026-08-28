import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from amazon_data_extractor.asin_reader import InputWorkbookError, extract_asin, read_asins


class AsinReaderTests(unittest.TestCase):
    def test_extracts_plain_asin_and_amazon_url(self):
        self.assertEqual(extract_asin("b0g38v6mrn"), "B0G38V6MRN")
        self.assertEqual(
            extract_asin("https://www.amazon.com/dp/B0CHHSFMRL?th=1&psc=1"),
            "B0CHHSFMRL",
        )
        self.assertIsNone(extract_asin("not-an-asin"))

    def test_reads_unique_asins_and_audits_bad_rows(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "Products"
            sheet.append(["SKU", "Child ASIN"])
            sheet.append(["A", "B0G38V6MRN"])
            sheet.append(["B", "https://www.amazon.com/dp/B0CHHSFMRL"])
            sheet.append(["C", "B0G38V6MRN"])
            sheet.append(["D", "bad"])
            sheet.append(["E", None])
            workbook.save(path)
            workbook.close()

            result = read_asins(path)

            self.assertEqual(result.asins, ["B0G38V6MRN", "B0CHHSFMRL"])
            self.assertEqual(result.duplicate_count, 1)
            self.assertEqual(result.invalid_count, 2)
            self.assertEqual(result.sheet_name, "Products")
            self.assertEqual(result.column_name, "Child ASIN")

    def test_missing_asin_column_has_clear_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "input.xlsx"
            workbook = Workbook()
            workbook.active.append(["SKU", "Title"])
            workbook.save(path)
            workbook.close()

            with self.assertRaisesRegex(InputWorkbookError, "Could not find column"):
                read_asins(path)


if __name__ == "__main__":
    unittest.main()

