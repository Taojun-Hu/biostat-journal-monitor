import csv
import tempfile
import unittest
from pathlib import Path

from scripts.monthly_checker import (
    classify_topic,
    deduplicate,
    new_rows,
    read_history,
    screen_paper,
    update_history,
    write_history,
)


class ScreeningTests(unittest.TestCase):
    def test_core_journal_is_included(self):
        decision, _ = screen_paper("core", "A general method", "", "", "")
        self.assertEqual(decision, "Included")

    def test_biomedical_application_is_included(self):
        decision, _ = screen_paper(
            "broad", "A model for cancer survival", "", "", ""
        )
        self.assertEqual(decision, "Included")

    def test_adjacent_method_needs_review(self):
        decision, _ = screen_paper(
            "broad", "A Bayesian method for networks", "", "", ""
        )
        self.assertEqual(decision, "Needs Review")

    def test_irrelevant_paper_is_excluded(self):
        decision, _ = screen_paper(
            "broad", "Forecasting football scores", "", "", ""
        )
        self.assertEqual(decision, "Excluded")

    def test_topic_uses_title_signal(self):
        topic = classify_topic(
            "Causal effects under interference",
            "A Bayesian model is used.",
            "",
            "",
        )
        self.assertEqual(topic, "Causal inference and treatment effects")


class StateTests(unittest.TestCase):
    def test_deduplicate_prefers_latest_date(self):
        rows = [
            {"DOI": "10.1/test", "OpenAlex ID": "A", "Publication Date": "2026-01-01"},
            {"DOI": "10.1/test", "OpenAlex ID": "B", "Publication Date": "2026-02-01"},
        ]
        result = deduplicate(rows)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["OpenAlex ID"], "B")

    def test_history_round_trip_and_unseen(self):
        row = {
            "DOI": "10.1/test",
            "Publication Date": "2026-02-01",
            "Journal": "Biometrics",
            "Screening Decision": "Included",
            "Title": "Test",
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "history.csv"
            history = {}
            update_history(history, [row], "2026-03-01")
            write_history(history, path)
            loaded = read_history(path)
            self.assertIn("10.1/test", loaded)
            self.assertEqual(new_rows([row], loaded), [])
            with path.open(newline="", encoding="utf-8") as handle:
                self.assertEqual(len(list(csv.DictReader(handle))), 1)


if __name__ == "__main__":
    unittest.main()

