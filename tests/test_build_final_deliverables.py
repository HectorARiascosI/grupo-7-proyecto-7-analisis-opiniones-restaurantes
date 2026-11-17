import csv
import json
import tempfile
import unittest
from pathlib import Path

from build_final_deliverables import build_dashboard, build_pdf, build_summary


class BuildSummaryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.data = self.root / "data"
        self.data.mkdir()
        self._write(
            "dim_business.csv",
            ["business_id", "business_name", "city", "price_range"],
            [
                ["b1", "Cafe One", "City A", "1"],
                ["b2", "Cafe Two", " city a ", ""],
                ["b3", "Cafe Three", "", "1"],
            ],
        )
        self._write(
            "fact_review.csv",
            [
                "business_id",
                "review_day",
                "stars",
                "useful",
                "sentiment_label",
                "category",
                "complaint_terms",
            ],
            [
                ["b1", "2022-01-01", "5", "1", "Positive", "Cafe", "bad"],
                ["b1", "2022-01-02", "3", "2", "Negative", "Cafe", ""],
                ["b2", "2022-02-01", "4", "0", "Positive", "Pet Training", "slow"],
                ["b3", "2022-02-02", "4", "0", "Neutral", "Cafe", ""],
            ],
        )
        self._write(
            "dim_user.csv",
            ["activity_segment", "review_count", "average_stars", "fans"],
            [
                ["Q1", "2", "4.0", "0"],
                ["Q2", "5", "3.5", "2"],
            ],
        )
        self._write(
            "complaint_term_counts.csv",
            ["complaint_term", "frequency"],
            [["bad", "1"], ["slow", "1"]],
        )
        self._write(
            "review_keyword_counts.csv",
            ["keyword", "frequency"],
            [["good", "2"], ["slow", "1"]],
        )

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def _write(self, name: str, fields: list[str], rows: list[list[str]]) -> None:
        with (self.data / name).open("w", encoding="utf-8", newline="") as output:
            writer = csv.writer(output)
            writer.writerow(fields)
            writer.writerows(rows)

    def test_build_summary_aggregates_small_chunks_without_losing_groups(self) -> None:
        result = build_summary(self.data, chunk_size=2)

        self.assertEqual(result["meta"]["reviews"], 4)
        self.assertEqual(result["meta"]["average_stars"], 4.0)
        self.assertEqual(result["meta"]["sentiment"]["Positive"], 2)
        self.assertEqual(result["meta"]["sentiment"]["Negative"], 1)
        self.assertEqual(result["meta"]["sentiment"]["Neutral"], 1)
        self.assertEqual(result["meta"]["cities"], 1)
        self.assertEqual(result["cities_top"][0]["label"], "City A")
        self.assertEqual(result["cities_top"][0]["reviews"], 3)

        price_rows = {row["label"]: row for row in result["price"]}
        self.assertEqual(price_rows["1"]["reviews"], 3)
        self.assertEqual(price_rows["1"]["average_stars"], 4.0)
        self.assertEqual(price_rows["Sin dato"]["reviews"], 1)

        cafe = next(row for row in result["category_volume"] if row["label"] == "Cafe")
        self.assertEqual(cafe["reviews"], 3)
        self.assertEqual(cafe["useful_votes"], 3)
        self.assertEqual(cafe["complaint_reviews"], 1)

    def test_dashboard_contains_aggregates_but_no_review_text(self) -> None:
        result = build_summary(self.data, chunk_size=2)
        output = self.root / "dashboard.html"
        build_dashboard(result, output)

        source = output.read_text(encoding="utf-8")
        payload = source.split(
            '<script id="dashboard-data" type="application/json">', 1
        )[1].split("</script>", 1)[0]
        data = json.loads(payload)
        self.assertEqual(data["meta"]["reviews"], 4)
        self.assertIn("Pet Training", source)
        self.assertNotIn("review_text", payload)
        self.assertNotIn("user_id", payload)


class BuildPdfTests(unittest.TestCase):
    def test_build_pdf_generates_a_readable_pdf_file(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            source = root / "guide.md"
            output = root / "guide.pdf"
            source.write_text(
                "# Guía DAX\n\nTexto breve.\n\n```DAX\nEVALUATE ROW(\"n\", 1)\n```\n",
                encoding="utf-8",
            )

            build_pdf(source, output, title="Guía DAX")

            self.assertTrue(output.is_file())
            self.assertTrue(output.read_bytes().startswith(b"%PDF-"))


if __name__ == "__main__":
    unittest.main()
