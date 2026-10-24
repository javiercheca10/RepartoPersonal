"""Pruebas unitarias de regresión para la Fase 1 de la auditoría (Integridad de datos).

Cubre:
- C1: Rangos de preferencia fuera de 1..len(positions) en cargador y en solver CP-SAT.
- C2: Round-trip sin pérdida de reglas de dotación (objetivos, mínimos y prioridades 0/1/2).
- C3: Preservación de espacios de nombres OpenXML (xmlns:mc, xmlns:x14ac, mc:Ignorable).
- C4: Desplazamiento dinámico del pie y celdas combinadas en Hoja 2 con >16 puestos y preservación de estilos.
- A1: Fallback de columnas en load_editable_workers sin NameError ('baja_col').
- A7: Mes sin días laborables en solve_month devuelve MonthResult válido sin TypeError.
- A8: Copias de seguridad rotativas y coincidencia por nombre en save_sequence_records.
"""

from datetime import date
from pathlib import Path
import shutil
import tempfile
import unittest
import xml.etree.ElementTree as ET
import zipfile

from reparto.datos import (
    DaySchedule,
    WorkerProfile,
    load_all_workers_with_report,
    load_sequence_records,
    load_skills_and_preferences,
    parse_position_rules,
    save_sequence_records,
    sync_sequences_roster,
)
from reparto.exportadores import _format_shift_mark
from reparto.gestion_trabajadores import (
    EditableWorker,
    load_editable_workers,
    save_editable_workers,
)
from reparto.solver import solve_month
from tests.fixtures import create_fictitious_dataset, write_config


class TestAuditoriaFase1(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.seq_path, self.pref_path, _ = create_fictitious_dataset(self.dir, total_workers=10)

    def tearDown(self):
        self.tmp.cleanup()

    def test_c1_out_of_range_preference_ranks_do_not_leave_vacancies(self):
        """C1: Rangos de preferencia > len(positions) no fuerzan x == 0 ni dejan vacantes falsas."""
        w1 = WorkerProfile("w1", "TRABAJADOR FICTICIO 1", frozenset({"P1", "P2"}), frozenset(), {"P1": 1, "P2": 15}, ("M",) * 168)
        w2 = WorkerProfile("w2", "TRABAJADOR FICTICIO 2", frozenset({"P1", "P2"}), frozenset(), {"P1": 1, "P2": 15}, ("M",) * 168)
        sched = (
            DaySchedule(
                day=date(2026, 11, 3),
                cycle_day=1,
                weekday=1,
                is_workday=True,
                workers_m=("w1", "w2"),
                workers_t=(),
                workers_r=(),
                targets_m={"P1": (1, 1), "P2": (1, 1)},
                targets_t={"P1": (0, 0), "P2": (0, 0)},
            ),
        )
        res = solve_month(
            2026, 11, sched, (w1, w2),
            positions=("P1", "P2"),
            position_rules={
                "P1": {"target_m": 1, "target_t": 0, "min_m": 1, "min_t": 0, "priority": 0},
                "P2": {"target_m": 1, "target_t": 0, "min_m": 1, "min_t": 0, "priority": 0},
            },
        )
        self.assertEqual(res.vacancies_total, 0)
        self.assertEqual(len(res.vacancies), 0)
        assigned_pos = {a.position for a in res.assignments}
        self.assertEqual(assigned_pos, {"P1", "P2"})

    def test_c1_loader_compacts_sparse_preference_ranks(self):
        """C1: load_skills_and_preferences compacta preferencias dispersas (ej. 5 y 12 -> 1 y 2)."""
        _, pref = write_config(self.dir / "sparse", [
            {"name": "ANA FICTICIA", "seq": ["M"] * 168, "skills": {"EC", "MG"}, "ranking": {"MG": 5, "EC": 12}},
        ])
        prof = load_skills_and_preferences(pref)["ANA FICTICIA"]
        self.assertEqual(prof["ranking"], {"MG": 1, "EC": 2})

    def test_c2_rules_round_trip_lossless(self):
        """C2: Guardar y recargar reglas conserva exactamente target_m, target_t, min_m, min_t y priority."""
        workers = load_editable_workers(self.pref_path)
        positions = ("P_PRIO0_RED", "P_PRIO1_RED", "P_PRIO2_M_T")
        rules = {
            "P_PRIO0_RED": {"name": "Puesto Prioridad 0 con Mínimo", "target_m": 3, "target_t": 2, "min_m": 2, "min_t": 1, "priority": 0},
            "P_PRIO1_RED": {"name": "Puesto Prioridad 1", "target_m": 4, "target_t": 2, "min_m": 2, "min_t": 1, "priority": 1},
            "P_PRIO2_M_T": {"name": "Puesto Prioridad 2 con Mínimo M y T", "target_m": 4, "target_t": 3, "min_m": 2, "min_t": 2, "priority": 2},
        }
        save_editable_workers(self.pref_path, workers[:3], positions=positions, rules=rules)
        reloaded = parse_position_rules(self.pref_path, positions)
        for p in positions:
            self.assertEqual(reloaded[p]["target_m"], rules[p]["target_m"], p)
            self.assertEqual(reloaded[p]["target_t"], rules[p]["target_t"], p)
            self.assertEqual(reloaded[p]["min_m"], rules[p]["min_m"], p)
            self.assertEqual(reloaded[p]["min_t"], rules[p]["min_t"], p)
            self.assertEqual(reloaded[p]["priority"], rules[p]["priority"], p)

    def test_c3_openxml_namespaces_preserved(self):
        """C3: Guardar ficheros conserva xmlns:mc y xmlns:x14ac cuando mc:Ignorable='x14ac' está presente."""
        # Inyectar atributos mc:Ignorable="x14ac" como hace Excel de escritorio al guardar
        with zipfile.ZipFile(self.pref_path, "r") as zin:
            manifest = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        s1 = manifest["xl/worksheets/sheet1.xml"].decode("utf-8")
        s1 = s1.replace(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
            'mc:Ignorable="x14ac">',
        )
        manifest["xl/worksheets/sheet1.xml"] = s1.encode("utf-8")
        with zipfile.ZipFile(self.pref_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for k, v in manifest.items():
                zout.writestr(k, v)

        workers = load_editable_workers(self.pref_path)
        save_editable_workers(self.pref_path, workers)

        with zipfile.ZipFile(self.pref_path, "r") as zin:
            saved_xml = zin.read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn("xmlns:x14ac=", saved_xml)
        self.assertIn("xmlns:mc=", saved_xml)
        # El XML debe ser válido al parsearlo con ElementTree
        ET.fromstring(saved_xml.encode("utf-8"))

    def test_c4_sheet2_more_than_16_positions_no_collision(self):
        """C4: Guardar 20 puestos en Hoja 2 desplaza el bloque informativo inferior sin colisión ni pérdida."""
        src_template = Path("RepartoPersonal/configuracion_ejemplo/Plantilla_Conocimientos_y_Preferencias.xlsx")
        if src_template.exists():
            shutil.copy2(src_template, self.pref_path)
        workers = load_editable_workers(self.pref_path)
        positions_20 = tuple(f"P{i:02d}" for i in range(1, 21))
        rules_20 = {
            p: {"name": f"Puesto {p}", "target_m": 2, "target_t": 1, "min_m": 1, "min_t": 1, "priority": i % 3}
            for i, p in enumerate(positions_20)
        }
        save_editable_workers(self.pref_path, workers[:5], positions=positions_20, rules=rules_20)
        reloaded_rules = parse_position_rules(self.pref_path, positions_20)
        self.assertEqual(len(reloaded_rules), 20)
        for i, p in enumerate(positions_20):
            self.assertEqual(reloaded_rules[p]["priority"], i % 3)

    def test_a1_fallback_header_without_position_matches(self):
        """A1: load_editable_workers no lanza NameError ('baja_col') cuando la fila 3 no contiene los códigos esperados."""
        with zipfile.ZipFile(self.pref_path, "r") as zin:
            manifest = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        tree = ET.fromstring(manifest["xl/worksheets/sheet1.xml"])
        for row_el in tree.findall(f".//{{{MAIN}}}row"):
            if row_el.attrib.get("r") == "3":
                for c_el in list(row_el):
                    if not c_el.attrib.get("r", "").startswith(("A", "B", "C")):
                        row_el.remove(c_el)
        manifest["xl/worksheets/sheet1.xml"] = ET.tostring(tree, encoding="utf-8", xml_declaration=True)
        broken_path = self.dir / "broken_header.xlsx"
        with zipfile.ZipFile(broken_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for k, v in manifest.items():
                zout.writestr(k, v)

        loaded = load_editable_workers(broken_path)
        self.assertGreaterEqual(len(loaded), 1)

    def test_a7_empty_workdays_in_solver(self):
        """A7: solve_month con un mes sin días laborables devuelve MonthResult válido sin TypeError."""
        res = solve_month(2026, 11, (), ())
        self.assertTrue(res.is_optimal)
        self.assertEqual(res.status_name, "Óptimo")
        self.assertEqual(res.vacancies, ())

    def test_a8_save_sequence_records_backup_and_name_match(self):
        """A8: save_sequence_records crea copia en copias/ y empareja por nombre aunque cambie row."""
        records, weeks = load_sequence_records(self.seq_path)
        self.assertGreater(len(records), 0)
        records[0]["sequence"][0] = "R"
        # Simular desfase de índice de fila para verificar que empareja por nombre
        records[0]["row"] = 999
        save_sequence_records(self.seq_path, records, cycle_weeks=weeks)

        reloaded, _ = load_sequence_records(self.seq_path)
        self.assertEqual(reloaded[0]["sequence"][0], "R")
        self.assertTrue((self.dir / "copias").exists())

    def test_c5_sync_sequences_roster_adds_and_removes_workers(self):
        """C5: sync_sequences_roster sincroniza altas y bajas definitivas entre Plantilla y Secuencias."""
        workers = load_editable_workers(self.pref_path)
        removed_name = workers[0].name
        workers.pop(0)
        new_name = "NUEVO PROFESIONAL SINCRONIZADO"
        workers.append(EditableWorker(
            name=new_name,
            row_index=len(workers) + 4,
            is_baja=False,
            skills={"EC": "X"},
            preference_order=["EC"],
        ))
        save_editable_workers(self.pref_path, workers)
        changed = sync_sequences_roster(
            self.seq_path,
            [w.name for w in workers],
            removed_worker_names={removed_name},
        )
        self.assertTrue(changed)

        loaded_profiles, report = load_all_workers_with_report(self.seq_path, self.pref_path)
        names_loaded = {w.name for w in loaded_profiles}
        self.assertIn(new_name, names_loaded)
        self.assertNotIn(removed_name, names_loaded)
        self.assertEqual(report.sin_ficha, ())
        self.assertEqual(report.sin_secuencia, ())

    def test_c6_dirty_tracking_in_tabs(self):
        """C6: TrabajadoresTab y SecuenciasWidget rastrean cambios sin guardar (is_dirty) y los limpian al guardar."""
        from PySide6.QtWidgets import QApplication
        from reparto.ui_secuencias import SecuenciasWidget
        from reparto.ui_trabajadores import TrabajadoresTab

        _app = QApplication.instance() or QApplication([])
        tab_w = TrabajadoresTab(self.pref_path, seq_path=self.seq_path)
        self.assertFalse(tab_w.is_dirty)
        tab_w._on_status_toggled(True)
        self.assertTrue(tab_w.is_dirty)
        self.assertTrue(tab_w.save_data())
        self.assertFalse(tab_w.is_dirty)

        tab_s = SecuenciasWidget(self.seq_path)
        self.assertFalse(tab_s.is_dirty)
        item = tab_s.table.item(0, 2)
        new_mark = "T" if item.text() != "T" else "M"
        item.setText(new_mark)
        self.assertTrue(tab_s.is_dirty)
        self.assertTrue(tab_s.save_data())
        self.assertFalse(tab_s.is_dirty)

    def test_a5_shift_mark_formatting_disambiguates_trailing_t(self):
        """A5: Puestos largos o terminados en 'T' usan sufijo '(T)' en tarde para evitar ambigüedad."""
        self.assertEqual(_format_shift_mark("EC", "M"), "EC")
        self.assertEqual(_format_shift_mark("EC", "T"), "ECt")
        self.assertEqual(_format_shift_mark("RT", "T"), "RT(T)")
        self.assertEqual(_format_shift_mark("EXT", "T"), "EXT(T)")
        self.assertEqual(_format_shift_mark("TRIAJE", "T"), "TRIAJE(T)")
        self.assertEqual(_format_shift_mark(".", "T"), ".t")

    def test_n1_solver_does_not_force_artificial_vacancy_in_prio1_when_prio2_lacks_skills(self):
        """N1: Si DR (prio 2) no tiene suficientes trabajadores con ese conocimiento, EC (prio 1) no se recorta artificialmente."""
        workers = (
            WorkerProfile("w1", "W1", frozenset({"EC"}), frozenset(), {"EC": 1}, ("M",) * 168),
            WorkerProfile("w2", "W2", frozenset({"EC"}), frozenset(), {"EC": 1}, ("M",) * 168),
            WorkerProfile("w3", "W3", frozenset({"EC"}), frozenset(), {"EC": 1}, ("M",) * 168),
            WorkerProfile("w4", "W4", frozenset({"EC"}), frozenset(), {"EC": 1}, ("M",) * 168),
            WorkerProfile("w5", "W5", frozenset({"DR"}), frozenset(), {"DR": 1}, ("M",) * 168),
            WorkerProfile("w6", "W6", frozenset({"DR"}), frozenset(), {"DR": 1}, ("M",) * 168),
        )
        sched = (
            DaySchedule(
                day=date(2026, 11, 3), cycle_day=1, weekday=1, is_workday=True,
                workers_m=("w1", "w2", "w3", "w4", "w5", "w6"), workers_t=(), workers_r=(),
                targets_m={"DR": (3, 2), "EC": (4, 3)},
                targets_t={"DR": (0, 0), "EC": (0, 0)},
            ),
        )
        rules = {
            "DR": {"target_m": 3, "target_t": 0, "min_m": 2, "min_t": 0, "priority": 2},
            "EC": {"target_m": 4, "target_t": 0, "min_m": 3, "min_t": 0, "priority": 1},
        }
        res = solve_month(2026, 11, sched, workers, positions=("DR", "EC"), position_rules=rules)
        self.assertEqual(res.vacancies_total, 1)
        self.assertEqual(res.vacancies[0].position, "DR")
        self.assertEqual(sum(1 for a in res.assignments if a.position == "EC"), 4)
        self.assertEqual(sum(1 for a in res.assignments if a.position == "."), 0)

    def test_n1_solver_reduces_prio1_before_prio2_regardless_of_catalog_order(self):
        """N1: Cuando comparten personal polivalente, prio 1 (EC) se reduce antes que prio 2 (DR) aunque DR vaya antes en el catálogo."""
        workers = tuple(
            WorkerProfile(f"w{i}", f"W{i}", frozenset({"DR", "EC"}), frozenset(), {"DR": 1, "EC": 1}, ("M",) * 168)
            for i in range(1, 7)
        )
        sched = (
            DaySchedule(
                day=date(2026, 11, 3), cycle_day=1, weekday=1, is_workday=True,
                workers_m=tuple(w.id for w in workers), workers_t=(), workers_r=(),
                targets_m={"DR": (3, 2), "EC": (4, 3)},
                targets_t={"DR": (0, 0), "EC": (0, 0)},
            ),
        )
        rules = {
            "DR": {"target_m": 3, "target_t": 0, "min_m": 2, "min_t": 0, "priority": 2},
            "EC": {"target_m": 4, "target_t": 0, "min_m": 3, "min_t": 0, "priority": 1},
        }
        res = solve_month(2026, 11, sched, workers, positions=("DR", "EC"), position_rules=rules)
        self.assertEqual(res.vacancies_total, 1)
        self.assertEqual(res.vacancies[0].position, "EC")
        self.assertEqual(res.vacancies_dr, 0)

    def test_n3_renaming_position_in_dialog_migrates_worker_skills_and_preferences(self):
        """N3: Renombrar el código de un puesto en PuestosDotacionesDialog migra competencias y preferencias."""
        from PySide6.QtWidgets import QApplication
        from reparto.ui_trabajadores import PuestosDotacionesDialog

        _app = QApplication.instance() or QApplication([])
        dlg = PuestosDotacionesDialog(
            positions=("EC", "MG"),
            rules={
                "EC": {"name": "Ecografía", "target_m": 2, "target_t": 1, "min_m": 2, "min_t": 1, "priority": 1},
                "MG": {"name": "Mamografía", "target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
            },
        )
        # Renombrar fila 0 de EC a ECO
        dlg.table.cellWidget(0, 0).setText("ECO")
        dlg._on_save()
        new_pos, new_rules = dlg.get_data()
        renamed = dlg.get_renamed_map()
        self.assertEqual(new_pos, ("ECO", "MG"))
        self.assertEqual(renamed, {"EC": "ECO"})
        self.assertEqual(new_rules["ECO"]["name"], "Ecografía")

    def test_n4_stale_calc_chain_removed_on_save(self):
        """N4: Si un archivo Excel contiene xl/calcChain.xml, atomic_write_zip lo limpia para evitar avisos de Excel."""
        with zipfile.ZipFile(self.pref_path, "r") as zin:
            manifest = {item.filename: zin.read(item.filename) for item in zin.infolist()}
        manifest["xl/calcChain.xml"] = b'<?xml version="1.0" encoding="UTF-8"?><calcChain/>'
        with zipfile.ZipFile(self.pref_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for k, v in manifest.items():
                zout.writestr(k, v)

        workers = load_editable_workers(self.pref_path)
        save_editable_workers(self.pref_path, workers)

        with zipfile.ZipFile(self.pref_path, "r") as zin:
            self.assertNotIn("xl/calcChain.xml", zin.namelist())

    def test_n6_pure_reinforcement_worker_loaded_in_sequences(self):
        """N6: Un profesional con solo marcas 'R' y 'D' se carga en load_sequences y no se descarta."""
        seq_r = ["R" if (d % 7) < 5 else "D" for d in range(168)]
        seq_p, pref_p = write_config(self.dir / "refuerzo", [
            {"name": "CORRETURNOS PURO", "seq": seq_r, "skills": {"EC"}, "ranking": {"EC": 1}},
        ])
        workers, report = load_all_workers_with_report(seq_p, pref_p)
        self.assertEqual(len(workers), 1)
        self.assertEqual(workers[0].name, "CORRETURNOS PURO")
        self.assertEqual(report.sin_secuencia, ())

    def test_r1_sync_sequences_renumbers_rows_contiguously(self):
        """R1: Al eliminar e insertar trabajadores en secuencias, las filas y N.º quedan contiguos (4..3+N y 1..N)."""
        from reparto.datos import load_sequence_records

        sync_sequences_roster(
            self.seq_path,
            active_worker_names=["TECNICO FICTICIO 01", "TECNICO FICTICIO 11"],
            removed_worker_names={"TECNICO FICTICIO 02", "TECNICO FICTICIO 05"},
        )
        recs, _ = load_sequence_records(self.seq_path)
        self.assertEqual(len(recs), 9)
        self.assertEqual([r["num"] for r in recs], [str(i) for i in range(1, 10)])
        self.assertEqual([r["row"] for r in recs], list(range(3, 12)))

    def test_ux_improvements_secuencias_and_preferences(self):
        """Verifica las mejoras UX en Secuencias (columna fija, selector semana, pincel) y flechas ▲/▼ en Preferencias."""
        from PySide6.QtWidgets import QApplication
        from reparto.ui_secuencias import SecuenciasWidget
        from reparto.ui_trabajadores import DraggablePreferenceList

        _app = QApplication.instance() or QApplication([])

        # 1. Flechas ▲ / ▼ en DraggablePreferenceList
        pref_list = DraggablePreferenceList()
        pref_list.set_preferences(["EC", "TAC", "RM"])
        self.assertEqual(pref_list.get_current_order(), ["EC", "TAC", "RM"])
        pref_list.move_item(2, -1)
        self.assertEqual(pref_list.get_current_order(), ["EC", "RM", "TAC"])
        pref_list.move_item(0, 1)
        self.assertEqual(pref_list.get_current_order(), ["RM", "EC", "TAC"])

        # 2. SecuenciasWidget: columna fija, selector de semana y pincel de un clic
        seq_w = SecuenciasWidget(self.seq_path)
        self.assertEqual(seq_w.week_combo.count(), 25)  # 24 semanas + Totales finales
        self.assertFalse(seq_w.frozen_table.isColumnHidden(0))
        self.assertFalse(seq_w.frozen_table.isColumnHidden(1))
        self.assertTrue(seq_w.frozen_table.isColumnHidden(2))

        seq_w._set_brush_mode("R")
        seq_w._on_cell_clicked(0, 2)  # Pintar día 1 de fila 0 con 'R'
        self.assertEqual(seq_w.records[0]["sequence"][0], "R")
        self.assertTrue(seq_w.is_dirty)


if __name__ == "__main__":
    unittest.main()
