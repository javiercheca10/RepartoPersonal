import tempfile
import unittest
from datetime import date
from pathlib import Path

from reparto.datos import (
    DEFAULT_HOLIDAYS,
    POSITIONS,
    computed_holidays,
    cycle_day_matches_weekday,
    holidays_for_year,
    is_workday,
    load_all_workers_with_report,
    load_skills_and_preferences,
    project_month,
    suggest_cycle_day,
    valid_cycle_days,
)
from tests.fixtures import pattern_sequence, write_config


class HolidayTests(unittest.TestCase):
    def test_computed_matches_verified_tables(self):
        for year, table in DEFAULT_HOLIDAYS.items():
            self.assertEqual(computed_holidays(year), table, year)

    def test_unverified_year_flagged(self):
        self.assertTrue(holidays_for_year(2026)[1])
        self.assertFalse(holidays_for_year(2029)[1])
        # Nunca vacío: el 1 de enero de 2029 es festivo aunque no esté en la tabla
        self.assertFalse(is_workday(date(2029, 1, 1)))

    def test_nov_2026(self):
        self.assertFalse(is_workday(date(2026, 11, 2)))  # 1-nov domingo trasladado
        self.assertTrue(is_workday(date(2026, 11, 3)))


class CycleTests(unittest.TestCase):
    def test_nov_2026_day_63_valid(self):
        self.assertTrue(cycle_day_matches_weekday(2026, 11, 63))
        self.assertIn(63, valid_cycle_days(2026, 11))

    def test_misaligned_detected(self):
        self.assertFalse(cycle_day_matches_weekday(2026, 11, 1))

    def test_suggestion_from_reference(self):
        ref = date(2026, 11, 1)
        self.assertEqual(suggest_cycle_day(ref, 63, 2026, 11), 63)
        self.assertEqual(suggest_cycle_day(ref, 63, 2026, 12), 63 + 30)
        for month in range(1, 13):
            d = suggest_cycle_day(ref, 63, 2027, month)
            self.assertTrue(cycle_day_matches_weekday(2027, month, d), month)
        d = suggest_cycle_day(ref, 63, 2026, 10)
        self.assertTrue(cycle_day_matches_weekday(2026, 10, d))

    def test_project_rejects_misaligned(self):
        with self.assertRaises(ValueError):
            project_month(2026, 11, 1, ())


class LoaderTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.dir = Path(self.tmp.name)

    def _cfg(self, workers):
        return write_config(self.dir, workers)

    def test_empty_skill_means_unknown(self):
        seq, pref = self._cfg([
            {"name": "ANA UNO", "seq": pattern_sequence("M"), "skills": {"EC", "MG"}, "ranking": {"MG": 1, "EC": 2}},
            {"name": "BEA DOS", "seq": pattern_sequence("T"), "skills": set(), "ranking": {}},
        ])
        workers, report = load_all_workers_with_report(seq, pref)
        by_name = {w.name: w for w in workers}
        self.assertEqual(by_name["ANA UNO"].skills, frozenset({"EC", "MG"}))
        self.assertEqual(by_name["BEA DOS"].skills, frozenset())
        self.assertIn("BEA DOS", report.sin_conocimientos)

    def test_missing_file_entry_no_invented_skills(self):
        seq, pref = self._cfg([
            {"name": "ANA UNO", "seq": pattern_sequence("M"), "skills": {"EC"}, "ranking": {"EC": 1}},
        ])
        # Añadir secuencia sin ficha reescribiendo solo secuencias
        seq2, _ = write_config(self.dir / "otra", [
            {"name": "ANA UNO", "seq": pattern_sequence("M")},
            {"name": "SIN FICHA", "seq": pattern_sequence("T")},
        ])
        workers, report = load_all_workers_with_report(seq2, pref)
        sin = [w for w in workers if w.name == "SIN FICHA"][0]
        self.assertEqual(sin.skills, frozenset())
        self.assertEqual(report.sin_ficha, ("SIN FICHA",))

    def test_baja_excluded_and_reported(self):
        seq, pref = self._cfg([
            {"name": "ANA UNO", "seq": pattern_sequence("M"), "skills": {"EC"}, "ranking": {"EC": 1}},
            {"name": "BEA DOS", "seq": pattern_sequence("M"), "skills": {"EC"}, "ranking": {"EC": 1}, "baja": True},
        ])
        workers, report = load_all_workers_with_report(seq, pref)
        self.assertEqual([w.name for w in workers], ["ANA UNO"])
        self.assertEqual(report.bajas, ("BEA DOS",))

    def test_ranking_unranked_known_positions_appended(self):
        _, pref = self._cfg([
            {"name": "ANA UNO", "seq": pattern_sequence("M"), "skills": {"EC", "MG", "UR"}, "ranking": {"UR": 1}},
        ])
        profile = load_skills_and_preferences(pref)["ANA UNO"]
        self.assertEqual(profile["ranking"]["UR"], 1)
        self.assertEqual(sorted(profile["ranking"].values()), [1, 2, 3])
        self.assertEqual(set(profile["ranking"]), {"EC", "MG", "UR"})

    def test_unknown_marks_reported_and_ids_stable(self):
        bad = pattern_sequence("M")
        bad[0] = "N"
        seq, pref = self._cfg([
            {"name": "ANA UNO", "seq": bad, "skills": {"EC"}, "ranking": {"EC": 1}},
            {"name": "BEA DOS", "seq": pattern_sequence("T"), "skills": {"EC"}, "ranking": {"EC": 1}},
        ])
        w1, rep = load_all_workers_with_report(seq, pref)
        self.assertTrue(any("marcas desconocidas" in w for w in rep.warnings))
        # IDs no dependen del orden
        seq_b, pref_b = write_config(self.dir / "inv", [
            {"name": "BEA DOS", "seq": pattern_sequence("T"), "skills": {"EC"}, "ranking": {"EC": 1}},
            {"name": "ANA UNO", "seq": pattern_sequence("M"), "skills": {"EC"}, "ranking": {"EC": 1}},
        ])
        w2, _ = load_all_workers_with_report(seq_b, pref_b)
        self.assertEqual({w.name: w.id for w in w1}, {w.name: w.id for w in w2})

    def test_projection_nov_2026(self):
        seq, pref = self._cfg([
            {"name": "ANA UNO", "seq": pattern_sequence("M"), "skills": {"EC"}, "ranking": {"EC": 1}},
        ])
        workers, _ = load_all_workers_with_report(seq, pref)
        # Con ciclo 63 el 1-nov (domingo) cae en descanso y el 3-nov (martes) es laborable
        schedule = project_month(2026, 11, 63, workers)
        by_day = {s.day.day: s for s in schedule}
        self.assertEqual(len(schedule), 30)
        self.assertFalse(by_day[1].is_workday)
        self.assertFalse(by_day[2].is_workday)   # festivo trasladado
        self.assertEqual(len(by_day[3].workers_m), 1)
        self.assertEqual(len(by_day[7].workers_m), 0)  # sábado


if __name__ == "__main__":
    unittest.main()
