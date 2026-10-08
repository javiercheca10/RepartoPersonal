"""Punto de entrada de la aplicación portable RepartoRadiologia.

Funciona de forma idéntica tanto si se ejecuta desde Python como si está
empaquetado en un ejecutable .exe con PyInstaller.
Detecta automáticamente las carpetas 'configuracion' y 'resultados'.
"""

from datetime import date
import os
from pathlib import Path
import platform
import subprocess
import sys

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

# Detectar directorio base tanto en script como en PyInstaller bundle
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

sys.path.insert(0, str(APP_DIR))

from reparto.datos import (
    cycle_day_matches_weekday,
    load_all_workers,
    project_month,
    suggest_cycle_day,
)
from reparto.exportadores import export_month_to_excel, export_month_to_pdf
from reparto.solver import MonthResult, solve_month

MONTHS_ES = [
    ("Enero", 1), ("Febrero", 2), ("Marzo", 3), ("Abril", 4),
    ("Mayo", 5), ("Junio", 6), ("Julio", 7), ("Agosto", 8),
    ("Septiembre", 9), ("Octubre", 10), ("Noviembre", 11), ("Diciembre", 12)
]


class SolverWorkerThread(QThread):
    progress_signal = Signal(str, int)
    finished_signal = Signal(object, object, object, str, str)
    error_signal = Signal(str)

    def __init__(self, year: int, month: int, start_cycle_day: int, seq_path: Path, pref_path: Path, out_dir: Path):
        super().__init__()
        self.year = year
        self.month = month
        self.start_cycle_day = start_cycle_day
        self.seq_path = seq_path
        self.pref_path = pref_path
        self.out_dir = out_dir

    def run(self):
        try:
            self.progress_signal.emit("Cargando archivos de configuración...", 15)
            workers = load_all_workers(self.seq_path, self.pref_path)

            self.progress_signal.emit("Proyectando calendario laboral y turnos...", 35)
            schedules = project_month(self.year, self.month, self.start_cycle_day, workers)

            self.progress_signal.emit("Optimizando asignación y equidad de puestos con CP-SAT...", 60)
            result = solve_month(self.year, self.month, schedules, workers, time_limit_seconds=15.0)

            if not result.is_optimal and len(result.assignments) == 0:
                self.error_signal.emit(
                    "No se pudo encontrar una asignación que cubra los mínimos. Compruebe la plantilla de personal."
                )
                return

            self.progress_signal.emit("Generando cuadrante Excel (.xlsx)...", 80)
            self.out_dir.mkdir(parents=True, exist_ok=True)
            month_name = MONTHS_ES[self.month - 1][0]
            excel_path = self.out_dir / f"Cuadrante_{month_name}_{self.year}.xlsx"
            export_month_to_excel(result, schedules, workers, excel_path)

            self.progress_signal.emit("Generando cuadrante PDF oficial...", 95)
            pdf_str = ""
            try:
                pdf_path = self.out_dir / f"Cuadrante_{month_name}_{self.year}.pdf"
                export_month_to_pdf(result, schedules, workers, pdf_path)
                pdf_str = str(pdf_path)
            except Exception:
                pass  # Si ocurre algún fallo imprevisto, se continúa con el Excel generado

            self.progress_signal.emit("¡Cuadrante completado con éxito!", 100)
            self.finished_signal.emit(result, schedules, workers, str(excel_path), pdf_str)

        except Exception as e:
            self.error_signal.emit(f"Error durante el cálculo: {str(e)}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RepartoPersonal — Planificación y Asignación de Puestos")
        self.resize(760, 680)
        self.setMinimumSize(700, 600)

        self.config_dir = APP_DIR / "configuracion"
        if not self.config_dir.exists() and (APP_DIR / "configuracion_ejemplo").exists():
            self.config_dir = APP_DIR / "configuracion_ejemplo"
        self.out_dir = APP_DIR / "resultados"

        # Buscar nombres de archivo
        self.seq_path = self.config_dir / "Secuencias_Turnos_Limpio.xlsx"
        if not self.seq_path.exists():
            self.seq_path = self.config_dir / "Secuencias_Turnos.xlsx"

        self.pref_path = self.config_dir / "Plantilla_Conocimientos_y_Preferencias.xlsx"

        self.last_excel_path = ""
        self.last_pdf_path = ""

        self._setup_ui()
        self._apply_styles()
        self._check_config_files()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(28, 24, 28, 24)
        main_layout.setSpacing(18)

        # 1. Cabecera
        header_box = QVBoxLayout()
        header_box.setSpacing(4)
        title_lbl = QLabel("Generador de Cuadrantes Mensuales")
        title_lbl.setObjectName("titleLabel")
        subtitle_lbl = QLabel("Sistema de Optimización y Reparto Equitativo de Puestos de Trabajo")
        subtitle_lbl.setObjectName("subtitleLabel")
        header_box.addWidget(title_lbl)
        header_box.addWidget(subtitle_lbl)
        main_layout.addLayout(header_box)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setObjectName("separator")
        main_layout.addWidget(sep)

        # 2. Tarjeta: Archivos de Configuración
        config_card = QFrame()
        config_card.setObjectName("card")
        config_card_layout = QVBoxLayout(config_card)
        config_card_layout.setSpacing(12)

        card1_title = QLabel("1. Archivos de Configuración")
        card1_title.setObjectName("cardTitle")
        config_card_layout.addWidget(card1_title)

        seq_row = QHBoxLayout()
        self.seq_status_lbl = QLabel("● Secuencias de Turnos:")
        self.seq_status_lbl.setObjectName("statusLabel")
        self.seq_path_lbl = QLabel(self.seq_path.name)
        self.seq_path_lbl.setObjectName("pathLabel")
        seq_row.addWidget(self.seq_status_lbl)
        seq_row.addWidget(self.seq_path_lbl, 1)
        config_card_layout.addLayout(seq_row)

        pref_row = QHBoxLayout()
        self.pref_status_lbl = QLabel("● Ficha de Conocimientos:")
        self.pref_status_lbl.setObjectName("statusLabel")
        self.pref_path_lbl = QLabel(self.pref_path.name)
        self.pref_path_lbl.setObjectName("pathLabel")
        pref_row.addWidget(self.pref_status_lbl)
        pref_row.addWidget(self.pref_path_lbl, 1)
        config_card_layout.addLayout(pref_row)

        main_layout.addWidget(config_card)

        # 3. Tarjeta: Período
        params_card = QFrame()
        params_card.setObjectName("card")
        params_card_layout = QVBoxLayout(params_card)
        params_card_layout.setSpacing(14)

        card2_title = QLabel("2. Período a Planificar")
        card2_title.setObjectName("cardTitle")
        params_card_layout.addWidget(card2_title)

        form_layout = QHBoxLayout()
        form_layout.setSpacing(20)

        mes_box = QVBoxLayout()
        mes_box.setSpacing(4)
        mes_lbl = QLabel("Mes:")
        mes_lbl.setObjectName("fieldLabel")
        self.month_combo = QComboBox()
        for name, num in MONTHS_ES:
            self.month_combo.addItem(name, num)
        today = date.today()
        self.month_combo.setCurrentIndex(today.month - 1)
        mes_box.addWidget(mes_lbl)
        mes_box.addWidget(self.month_combo)
        form_layout.addLayout(mes_box)

        year_box = QVBoxLayout()
        year_box.setSpacing(4)
        year_lbl = QLabel("Año:")
        year_lbl.setObjectName("fieldLabel")
        self.year_combo = QComboBox()
        for y in range(2026, 2031):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(today.year if today.year >= 2026 else 2026))
        year_box.addWidget(year_lbl)
        year_box.addWidget(self.year_combo)
        form_layout.addLayout(year_box)

        cycle_box = QVBoxLayout()
        cycle_box.setSpacing(4)
        cycle_lbl = QLabel("Día del Ciclo en Día 1 (1 a 168):")
        cycle_lbl.setObjectName("fieldLabel")
        self.cycle_combo = QComboBox()
        self.cycle_combo.setEditable(True)
        for d in range(1, 169):
            self.cycle_combo.addItem(str(d), d)
        self.cycle_combo.setCurrentIndex(0)
        self.cycle_combo.setToolTip("Indique qué día del ciclo de 24 semanas (1 a 168) corresponde al día 1 del mes a planificar.")
        cycle_box.addWidget(cycle_lbl)
        cycle_box.addWidget(self.cycle_combo)
        form_layout.addLayout(cycle_box)

        params_card_layout.addLayout(form_layout)
        main_layout.addWidget(params_card)

        # 4. Botón Acción
        self.generate_btn = QPushButton("Generar Cuadrante Mensual")
        self.generate_btn.setObjectName("generateButton")
        self.generate_btn.clicked.connect(self._on_generate_clicked)
        main_layout.addWidget(self.generate_btn)

        # 5. Progreso y Estado
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.status_lbl = QLabel("Listo para generar.")
        self.status_lbl.setObjectName("infoLabel")
        self.status_lbl.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.status_lbl)

        # 6. Tarjeta Resultados
        self.results_card = QFrame()
        self.results_card.setObjectName("resultsCard")
        results_layout = QVBoxLayout(self.results_card)
        results_layout.setSpacing(10)

        self.results_title = QLabel("¡Cuadrante Generado Correctamente!")
        self.results_title.setObjectName("resultsTitle")
        self.results_info = QLabel("")
        self.results_info.setObjectName("resultsInfo")

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.open_excel_btn = QPushButton("📊 Abrir Cuadrante Excel")
        self.open_excel_btn.setObjectName("actionBtn")
        self.open_excel_btn.clicked.connect(self._open_excel)

        self.open_pdf_btn = QPushButton("📑 Abrir Cuadrante PDF")
        self.open_pdf_btn.setObjectName("actionBtn")
        self.open_pdf_btn.clicked.connect(self._open_pdf)

        self.open_folder_btn = QPushButton("📁 Abrir Carpeta Resultados")
        self.open_folder_btn.setObjectName("actionSecondaryBtn")
        self.open_folder_btn.clicked.connect(self._open_folder)

        buttons_layout.addWidget(self.open_excel_btn)
        buttons_layout.addWidget(self.open_pdf_btn)
        buttons_layout.addWidget(self.open_folder_btn)

        results_layout.addWidget(self.results_title)
        results_layout.addWidget(self.results_info)
        results_layout.addLayout(buttons_layout)

        self.results_card.setVisible(False)
        main_layout.addWidget(self.results_card)

        main_layout.addStretch()

    def _apply_styles(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #F8FAFC; }
            #titleLabel { font-size: 20px; font-weight: bold; color: #0F172A; }
            #subtitleLabel { font-size: 13px; color: #475569; }
            #separator { color: #E2E8F0; height: 1px; }
            #card { background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px; padding: 14px; }
            #cardTitle { font-size: 14px; font-weight: bold; color: #1E3A8A; }
            #fieldLabel { font-size: 12px; font-weight: 600; color: #334155; }
            #statusLabel { font-size: 13px; font-weight: 600; color: #059669; }
            #pathLabel { font-size: 13px; color: #64748B; }
            QComboBox {
                background-color: #FFFFFF; border: 1px solid #CBD5E1;
                border-radius: 6px; padding: 6px 10px; font-size: 13px; color: #0F172A;
            }
            QComboBox:focus { border: 2px solid #2563EB; }
            #generateButton {
                background-color: #1E3A8A; color: #FFFFFF; font-size: 15px; font-weight: bold;
                border: none; border-radius: 8px; padding: 12px 24px;
            }
            #generateButton:hover { background-color: #1E40AF; }
            #generateButton:disabled { background-color: #94A3B8; }
            #progressBar {
                border: 1px solid #E2E8F0; border-radius: 6px; text-align: center;
                background-color: #F1F5F9; color: #0F172A; font-weight: bold; height: 22px;
            }
            #progressBar::chunk { background-color: #059669; border-radius: 5px; }
            #infoLabel { font-size: 13px; color: #475569; }
            #resultsCard { background-color: #F0FDF4; border: 1px solid #BBF7D0; border-radius: 10px; padding: 16px; }
            #resultsTitle { font-size: 15px; font-weight: bold; color: #166534; }
            #resultsInfo { font-size: 13px; color: #15803D; }
            #actionBtn {
                background-color: #059669; color: #FFFFFF; font-size: 13px; font-weight: 600;
                border: none; border-radius: 6px; padding: 9px 16px;
            }
            #actionBtn:hover { background-color: #047857; }
            #actionSecondaryBtn {
                background-color: #FFFFFF; color: #334155; font-size: 13px; font-weight: 600;
                border: 1px solid #CBD5E1; border-radius: 6px; padding: 9px 16px;
            }
            #actionSecondaryBtn:hover { background-color: #F8FAFC; }
        """)

    def _check_config_files(self):
        seq_ok = self.seq_path.exists()
        pref_ok = self.pref_path.exists()

        if seq_ok:
            self.seq_status_lbl.setText("✔ Secuencias de Turnos:")
            self.seq_status_lbl.setStyleSheet("color: #059669;")
        else:
            self.seq_status_lbl.setText("✖ Secuencias no encontradas:")
            self.seq_status_lbl.setStyleSheet("color: #DC2626;")

        if pref_ok:
            self.pref_status_lbl.setText("✔ Conocimientos y Preferencias:")
            self.pref_status_lbl.setStyleSheet("color: #059669;")
        else:
            self.pref_status_lbl.setText("✖ Ficha no encontrada:")
            self.pref_status_lbl.setStyleSheet("color: #DC2626;")

        if not (seq_ok and pref_ok):
            self.generate_btn.setEnabled(False)
            self.status_lbl.setText("Faltan archivos en la carpeta 'configuracion'.")

    def _on_generate_clicked(self):
        self.generate_btn.setEnabled(False)
        self.results_card.setVisible(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(5)
        self.status_lbl.setText("Iniciando cálculo...")

        month = self.month_combo.currentData()
        year = self.year_combo.currentData()

        raw_val = self.cycle_combo.currentData()
        if raw_val is not None:
            try:
                start_cycle_day = int(raw_val)
            except (ValueError, TypeError):
                start_cycle_day = None
        else:
            start_cycle_day = None

        if start_cycle_day is None:
            txt = self.cycle_combo.currentText().strip()
            digits = "".join(c for c in txt if c.isdigit())
            start_cycle_day = int(digits) if digits else 1

        start_cycle_day = max(1, min(168, start_cycle_day))

        # Validar si el día del ciclo encaja con el día de la semana real del día 1
        if not cycle_day_matches_weekday(year, month, start_cycle_day):
            suggested = suggest_cycle_day(year, month)
            reply = QMessageBox.warning(
                self,
                "Aviso de Desalineación del Ciclo",
                f"El día del ciclo {start_cycle_day} no coincide con el día de la semana del 1/{month}/{year}.\n"
                f"Esto puede descolocar los turnos de fin de semana.\n\n"
                f"¿Desea continuar de todos modos?\n"
                f"(Sugerencia recomendada: día {suggested})",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                self.generate_btn.setEnabled(True)
                self.progress_bar.setVisible(False)
                self.status_lbl.setText("Generación cancelada por desalineación de ciclo.")
                return

        self.thread = SolverWorkerThread(
            year=year,
            month=month,
            start_cycle_day=start_cycle_day,
            seq_path=self.seq_path,
            pref_path=self.pref_path,
            out_dir=self.out_dir,
        )
        self.thread.progress_signal.connect(self._on_progress)
        self.thread.finished_signal.connect(self._on_finished)
        self.thread.error_signal.connect(self._on_error)
        self.thread.start()

    def _on_progress(self, msg: str, val: int):
        self.status_lbl.setText(msg)
        self.progress_bar.setValue(val)

    def _on_finished(self, result: MonthResult, schedules, workers, excel_path: str, pdf_path: str):
        self.generate_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText(f"Generación finalizada ({result.status_name}).")

        self.last_excel_path = excel_path
        self.last_pdf_path = pdf_path

        free_cnt = sum(1 for a in result.assignments if a.position == ".")
        vac_cnt = len(result.vacancies)
        info_txt = (
            f"• Estado del cálculo: {result.status_name}  |  "
            f"Asignaciones: {len(result.assignments)}  |  "
            f"Libres (.): {free_cnt}  |  "
            f"Tiempo: {result.execution_time_seconds}s\n"
            f"• Rango de esfuerzo: {result.effort_spread:.2f} pts/día  "
            f"(Mín: {result.min_mean_effort:.2f}, Máx: {result.max_mean_effort:.2f})"
        )
        if vac_cnt > 0:
            info_txt += f"\n• ⚠️ Puestos descubiertos a revisar a mano: {vac_cnt} (marcados en rojo en el Excel)"
        if result.issues:
            info_txt += "\n• " + "\n• ".join(result.issues)
        self.results_info.setText(info_txt)
        self.open_pdf_btn.setEnabled(bool(pdf_path and Path(pdf_path).exists()))
        self.results_card.setVisible(True)

    def _on_error(self, err_msg: str):
        self.generate_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.status_lbl.setText("Error en el cálculo.")
        QMessageBox.critical(self, "Error al generar cuadrante", err_msg)

    def _open_excel(self):
        if self.last_excel_path and Path(self.last_excel_path).exists():
            self._open_file(self.last_excel_path)

    def _open_pdf(self):
        if self.last_pdf_path and Path(self.last_pdf_path).exists():
            self._open_file(self.last_pdf_path)

    def _open_folder(self):
        folder = str(self.out_dir.resolve())
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.run(["open", folder])
        else:
            subprocess.run(["xdg-open", folder])

    def _open_file(self, filepath: str):
        p = str(Path(filepath).resolve())
        if platform.system() == "Windows":
            os.startfile(p)
        elif platform.system() == "Darwin":
            subprocess.run(["open", p])
        else:
            subprocess.run(["xdg-open", p])


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
