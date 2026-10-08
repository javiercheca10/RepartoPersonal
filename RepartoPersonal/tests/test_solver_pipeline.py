"""Pruebas automáticas del solver, auditoría y exportadores."""

from datetime import date
from pathlib import Path
import unittest

from reparto.auditoria import compute_vacancies
from reparto.datos import (
    DaySchedule,
    WorkerProfile,
    load_all_workers,
    project_month,
)
from reparto.exportadores import export_month_to_excel, export_month_to_pdf
from reparto.solver import solve_month

BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = BASE_DIR / "configuracion"


class TestSolverAndExporters(unittest.TestCase):
    def setUp(self):
        self.seq_path = CONFIG_DIR / "Secuencias_Turnos_Limpio.xlsx"
        self.pref_path = CONFIG_DIR / "Plantilla_Conocimientos_y_Preferencias.xlsx"
        self.workers = load_all_workers(self.seq_path, self.pref_path)
        self.schedules = project_month(2026, 11, 63, self.workers)

    def test_baja_workers_excluded(self):
        """Los trabajadores marcados con BAJA quedan excluidos de los perfiles cargados."""
        self.assertEqual(len(self.workers), 53)

    def test_solve_november_2026(self):
        """El solver genera una solución válida respetando habilidades y vetos."""
        res = solve_month(2026, 11, self.schedules, self.workers, time_limit_seconds=15.0)
        self.assertIn(res.status_name, ("Óptimo", "Factible"))
        self.assertGreater(len(res.assignments), 750)

        # Verificar que no hay violaciones de habilidades ni vetos
        workers_by_id = {w.id: w for w in self.workers}
        for a in res.assignments:
            if a.position == ".":
                continue
            w = workers_by_id[a.worker_id]
            self.assertIn(a.position, w.skills, f"{w.name} asignado a {a.position} sin saberlo")
            self.assertNotIn(a.position, w.vetoes, f"{w.name} asignado a {a.position} que tiene vetado")

    def test_export_excel_and_pdf(self):
        """Se generan los archivos Excel y PDF y contienen las secciones requeridas."""
        res = solve_month(2026, 11, self.schedules, self.workers, time_limit_seconds=15.0)
        out_dir = BASE_DIR / "resultados" / "test_output"
        out_dir.mkdir(parents=True, exist_ok=True)
        xlsx_path = out_dir / "test_nov_2026.xlsx"
        pdf_path = out_dir / "test_nov_2026.pdf"

        export_month_to_excel(res, self.schedules, self.workers, xlsx_path)
        self.assertTrue(xlsx_path.exists())
        self.assertGreater(xlsx_path.stat().st_size, 5000)

        export_month_to_pdf(res, self.schedules, self.workers, pdf_path)
        self.assertTrue(pdf_path.exists())
        self.assertGreater(pdf_path.stat().st_size, 5000)


if __name__ == "__main__":
    unittest.main()
