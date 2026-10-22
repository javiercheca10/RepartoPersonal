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
from PySide6.QtGui import QIcon, QPixmap
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
    QSpinBox,
    QTabWidget,
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
from reparto.ui_trabajadores import TrabajadoresTab

MONTHS_ES = [
    ("Enero", 1), ("Febrero", 2), ("Marzo", 3), ("Abril", 4),
    ("Mayo", 5), ("Junio", 6), ("Julio", 7), ("Agosto", 8),
    ("Septiembre", 9), ("Octubre", 10), ("Noviembre", 11), ("Diciembre", 12)
]


class SolverWorkerThread(QThread):
    progress_signal = Signal(str, int)
    finished_signal = Signal(object, object, object, str, str)
    error_signal = Signal(str)

    def __init__(
        self,
        year: int,
        month: int,
        start_cycle_day: int,
        seq_path: Path,
        pref_path: Path,
        out_dir: Path,
        time_limit: float = 15.0,
    ):
        super().__init__()
        self.year = year
        self.month = month
        self.start_cycle_day = start_cycle_day
        self.seq_path = seq_path
        self.pref_path = pref_path
        self.out_dir = out_dir
        self.time_limit = time_limit

    def run(self):
        try:
            self.progress_signal.emit("Cargando archivos de configuración...", 15)
            workers = load_all_workers(self.seq_path, self.pref_path)

            self.progress_signal.emit("Proyectando calendario laboral y turnos...", 35)
            schedules = project_month(self.year, self.month, self.start_cycle_day, workers)

            self.progress_signal.emit("Optimizando asignación y equidad de puestos con CP-SAT...", 60)
            result = solve_month(self.year, self.month, schedules, workers, time_limit_seconds=self.time_limit)

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
        self.setWindowTitle("Planificación de Puestos — Servicio de Radiodiagnóstico")
        self.resize(880, 750)
        self.setMinimumSize(820, 680)

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

        # Identidad visual e icono
        self.icon_path = APP_DIR / "reparto" / "assets" / "icono.png"
        if not self.icon_path.exists():
            self.icon_path = APP_DIR / "reparto" / "assets" / "icono_48.png"
        if self.icon_path.exists():
            self.setWindowIcon(QIcon(str(self.icon_path)))

        self._setup_ui()
        self._apply_styles()
        self._check_config_files()

    def _setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(28, 20, 28, 20)
        main_layout.setSpacing(14)

        # 1. Cabecera Común (Icono de Identidad + Título)
        header_row = QHBoxLayout()
        header_row.setSpacing(14)
        header_row.setAlignment(Qt.AlignVCenter)

        if self.icon_path.exists():
            icon_lbl = QLabel()
            pix = QPixmap(str(self.icon_path)).scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            icon_lbl.setPixmap(pix)
            icon_lbl.setFixedSize(48, 48)
            icon_lbl.setObjectName("headerAppIcon")
            header_row.addWidget(icon_lbl)

        header_box = QVBoxLayout()
        header_box.setSpacing(3)
        title_lbl = QLabel("Planificación de Puestos de Trabajo")
        title_lbl.setObjectName("titleLabel")
        subtitle_lbl = QLabel("Servicio de Radiodiagnóstico · Gestión mensual de turnos y cuadrantes")
        subtitle_lbl.setObjectName("subtitleLabel")
        header_box.addWidget(title_lbl)
        header_box.addWidget(subtitle_lbl)

        header_row.addLayout(header_box, 1)
        main_layout.addLayout(header_row)

        # 2. Selector de Pestañas
        self.tab_widget = QTabWidget()
        self.tab_widget.setObjectName("mainTabWidget")

        # --- PESTAÑA 1: PLANIFICACIÓN Y CUADRANTE ---
        plan_tab = QWidget()
        plan_layout = QVBoxLayout(plan_tab)
        plan_layout.setContentsMargins(0, 14, 0, 0)
        plan_layout.setSpacing(16)

        # Tarjeta: Archivos de Configuración
        config_card = QFrame()
        config_card.setObjectName("card")
        config_card_layout = QVBoxLayout(config_card)
        config_card_layout.setSpacing(12)

        card1_header = QHBoxLayout()
        card1_title = QLabel("Archivos de Configuración")
        card1_title.setObjectName("cardTitle")
        card1_header.addWidget(card1_title)
        card1_header.addStretch()

        self.open_config_btn = QPushButton("Abrir carpeta de datos")
        self.open_config_btn.setObjectName("actionSecondaryBtn")
        self.open_config_btn.setToolTip("Abre la carpeta donde se encuentran las plantillas de turnos y conocimientos para editarlas.")
        self.open_config_btn.clicked.connect(self._open_config_folder)
        card1_header.addWidget(self.open_config_btn)

        config_card_layout.addLayout(card1_header)

        # Fila 1: Secuencias
        seq_row = QHBoxLayout()
        seq_name_box = QVBoxLayout()
        seq_name_box.setSpacing(2)
        seq_title = QLabel("Secuencias de Turnos (Ciclo de 24 semanas)")
        seq_title.setObjectName("itemTitle")
        self.seq_path_lbl = QLabel(self.seq_path.name)
        self.seq_path_lbl.setObjectName("pathLabel")
        seq_name_box.addWidget(seq_title)
        seq_name_box.addWidget(self.seq_path_lbl)

        self.seq_status_lbl = QLabel("Comprobando...")
        self.seq_status_lbl.setObjectName("badgeOk")
        self.seq_status_lbl.setAlignment(Qt.AlignCenter)

        seq_row.addLayout(seq_name_box, 1)
        seq_row.addWidget(self.seq_status_lbl)
        config_card_layout.addLayout(seq_row)

        item_sep = QFrame()
        item_sep.setFrameShape(QFrame.HLine)
        item_sep.setObjectName("itemSeparator")
        config_card_layout.addWidget(item_sep)

        # Fila 2: Conocimientos y Preferencias
        pref_row = QHBoxLayout()
        pref_name_box = QVBoxLayout()
        pref_name_box.setSpacing(2)
        pref_title = QLabel("Ficha de Competencias y Preferencias")
        pref_title.setObjectName("itemTitle")
        self.pref_path_lbl = QLabel(self.pref_path.name)
        self.pref_path_lbl.setObjectName("pathLabel")
        pref_name_box.addWidget(pref_title)
        pref_name_box.addWidget(self.pref_path_lbl)

        self.pref_status_lbl = QLabel("Comprobando...")
        self.pref_status_lbl.setObjectName("badgeOk")
        self.pref_status_lbl.setAlignment(Qt.AlignCenter)

        pref_row.addLayout(pref_name_box, 1)
        pref_row.addWidget(self.pref_status_lbl)
        config_card_layout.addLayout(pref_row)

        plan_layout.addWidget(config_card)

        # Tarjeta: Período a Planificar
        params_card = QFrame()
        params_card.setObjectName("card")
        params_card_layout = QVBoxLayout(params_card)
        params_card_layout.setSpacing(14)

        card2_title = QLabel("Período a Planificar")
        card2_title.setObjectName("cardTitle")
        params_card_layout.addWidget(card2_title)

        form_layout = QHBoxLayout()
        form_layout.setSpacing(16)

        mes_box = QVBoxLayout()
        mes_box.setSpacing(6)
        mes_lbl = QLabel("Mes")
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
        year_box.setSpacing(6)
        year_lbl = QLabel("Año")
        year_lbl.setObjectName("fieldLabel")
        self.year_combo = QComboBox()
        for y in range(2026, 2031):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(today.year if today.year >= 2026 else 2026))
        year_box.addWidget(year_lbl)
        year_box.addWidget(self.year_combo)
        form_layout.addLayout(year_box)

        cycle_box = QVBoxLayout()
        cycle_box.setSpacing(6)
        cycle_lbl = QLabel("Día del Ciclo en Día 1 (1 a 168)")
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

        time_box = QVBoxLayout()
        time_box.setSpacing(6)
        time_lbl = QLabel("Tiempo Máx. Cómputo")
        time_lbl.setObjectName("fieldLabel")
        self.time_spin = QSpinBox()
        self.time_spin.setRange(5, 300)
        self.time_spin.setValue(15)
        self.time_spin.setSingleStep(5)
        self.time_spin.setSuffix(" s")
        self.time_spin.setToolTip("Tiempo máximo de cálculo para el optimizador (segundos). Valores más altos permiten mayor equidad.")
        time_box.addWidget(time_lbl)
        time_box.addWidget(self.time_spin)
        form_layout.addLayout(time_box)

        params_card_layout.addLayout(form_layout)
        plan_layout.addWidget(params_card)

        # Botón Acción
        self.generate_btn = QPushButton("Generar Cuadrante")
        self.generate_btn.setObjectName("generateButton")
        self.generate_btn.clicked.connect(self._on_generate_clicked)
        plan_layout.addWidget(self.generate_btn)

        # Progreso y Estado
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        plan_layout.addWidget(self.progress_bar)

        self.status_lbl = QLabel("Listo para generar.")
        self.status_lbl.setObjectName("infoLabel")
        self.status_lbl.setAlignment(Qt.AlignCenter)
        plan_layout.addWidget(self.status_lbl)

        # Tarjeta Resultados
        self.results_card = QFrame()
        self.results_card.setObjectName("resultsCard")
        results_layout = QVBoxLayout(self.results_card)
        results_layout.setSpacing(12)

        self.results_title = QLabel("Cuadrante generado correctamente")
        self.results_title.setObjectName("resultsTitle")
        self.results_info = QLabel("")
        self.results_info.setObjectName("resultsInfo")

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(12)

        self.open_excel_btn = QPushButton("Abrir Cuadrante Excel")
        self.open_excel_btn.setObjectName("actionPrimaryBtn")
        self.open_excel_btn.clicked.connect(self._open_excel)

        self.open_pdf_btn = QPushButton("Abrir Cuadrante PDF")
        self.open_pdf_btn.setObjectName("actionResultBtn")
        self.open_pdf_btn.clicked.connect(self._open_pdf)

        self.open_folder_btn = QPushButton("Abrir Carpeta de Salida")
        self.open_folder_btn.setObjectName("actionResultBtn")
        self.open_folder_btn.clicked.connect(self._open_folder)

        buttons_layout.addWidget(self.open_excel_btn)
        buttons_layout.addWidget(self.open_pdf_btn)
        buttons_layout.addWidget(self.open_folder_btn)

        results_layout.addWidget(self.results_title)
        results_layout.addWidget(self.results_info)
        results_layout.addLayout(buttons_layout)

        self.results_card.setVisible(False)
        plan_layout.addWidget(self.results_card)
        plan_layout.addStretch()

        self.tab_widget.addTab(plan_tab, "Planificación y Cuadrante")

        # --- PESTAÑA 2: TRABAJADORES ---
        self.tab_trabajadores = TrabajadoresTab(self.pref_path)
        self.tab_trabajadores.data_saved.connect(self._on_worker_data_saved)
        self.tab_widget.addTab(self.tab_trabajadores, "Trabajadores")

        main_layout.addWidget(self.tab_widget, 1)

    def _apply_styles(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #F8FAFC;
                font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            }
            QTabWidget::pane {
                border: none;
                background-color: transparent;
            }
            QTabBar::tab {
                background-color: transparent;
                color: #64748B;
                font-size: 13px;
                font-weight: 600;
                padding: 8px 16px;
                margin-right: 6px;
                border-bottom: 2px solid transparent;
            }
            QTabBar::tab:hover {
                color: #0F172A;
            }
            QTabBar::tab:selected {
                color: #0284C7;
                border-bottom: 2px solid #0284C7;
            }
            #titleLabel {
                font-size: 20px;
                font-weight: 700;
                color: #0F172A;
                letter-spacing: -0.3px;
            }
            #subtitleLabel {
                font-size: 13px;
                color: #64748B;
            }
            #headerAppIcon {
                border-radius: 10px;
            }
            #separator {
                color: #E2E8F0;
                background-color: #E2E8F0;
                height: 1px;
                border: none;
            }
            #card {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 16px 20px;
            }
            #cardTitle {
                font-size: 14px;
                font-weight: 600;
                color: #0F172A;
                letter-spacing: -0.1px;
            }
            #itemTitle {
                font-size: 13px;
                font-weight: 500;
                color: #1E293B;
            }
            #pathLabel {
                font-size: 12px;
                color: #64748B;
            }
            #itemSeparator {
                color: #F1F5F9;
                background-color: #F1F5F9;
                height: 1px;
                border: none;
            }
            #fieldLabel {
                font-size: 12px;
                font-weight: 600;
                color: #475569;
            }
            QComboBox, QSpinBox {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 7px 11px;
                font-size: 13px;
                color: #0F172A;
                min-height: 20px;
            }
            QComboBox:hover, QSpinBox:hover {
                border-color: #94A3B8;
            }
            QComboBox:focus, QSpinBox:focus {
                border: 1.5px solid #0284C7;
            }
            QComboBox QAbstractItemView {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 4px;
                color: #0F172A;
                selection-background-color: #F0F9FF;
                selection-color: #0284C7;
            }
            #badgeOk {
                background-color: #ECFDF5;
                color: #065F46;
                border: 1px solid #A7F3D0;
                border-radius: 4px;
                padding: 3px 9px;
                font-size: 11px;
                font-weight: 600;
            }
            #badgeError {
                background-color: #FEF2F2;
                color: #991B1B;
                border: 1px solid #FECACA;
                border-radius: 4px;
                padding: 3px 9px;
                font-size: 11px;
                font-weight: 600;
            }
            #generateButton {
                background-color: #0284C7;
                color: #FFFFFF;
                font-size: 14px;
                font-weight: 600;
                border: none;
                border-radius: 6px;
                padding: 12px 24px;
                min-height: 20px;
            }
            #generateButton:hover {
                background-color: #0369A1;
            }
            #generateButton:pressed {
                background-color: #075985;
            }
            #generateButton:disabled {
                background-color: #E2E8F0;
                color: #94A3B8;
            }
            #progressBar {
                border: none;
                border-radius: 3px;
                background-color: #E2E8F0;
                height: 6px;
                text-align: center;
            }
            #progressBar::chunk {
                background-color: #0EA5E9;
                border-radius: 3px;
            }
            #infoLabel {
                font-size: 13px;
                color: #64748B;
                font-weight: 500;
            }
            #resultsCard {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 18px 20px;
            }
            #resultsTitle {
                font-size: 15px;
                font-weight: 600;
                color: #0F172A;
            }
            #resultsInfo {
                font-size: 13px;
                color: #334155;
                line-height: 1.5;
            }
            #actionPrimaryBtn {
                background-color: #0284C7;
                color: #FFFFFF;
                font-size: 13px;
                font-weight: 600;
                border: none;
                border-radius: 6px;
                padding: 9px 16px;
            }
            #actionPrimaryBtn:hover {
                background-color: #0369A1;
            }
            #actionResultBtn {
                background-color: #FFFFFF;
                color: #1E293B;
                font-size: 13px;
                font-weight: 500;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 9px 16px;
            }
            #actionResultBtn:hover {
                background-color: #F8FAFC;
                border-color: #94A3B8;
            }
            #actionSecondaryBtn {
                background-color: #FFFFFF;
                color: #334155;
                font-size: 12px;
                font-weight: 500;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 6px 12px;
            }
            #actionSecondaryBtn:hover {
                background-color: #F8FAFC;
                border-color: #94A3B8;
            }
        """)

    def _check_config_files(self):
        if not self.config_dir.exists() and (APP_DIR / "configuracion_ejemplo").exists():
            self.config_dir = APP_DIR / "configuracion_ejemplo"
        self.seq_path = self.config_dir / "Secuencias_Turnos_Limpio.xlsx"
        if not self.seq_path.exists():
            self.seq_path = self.config_dir / "Secuencias_Turnos.xlsx"
        self.pref_path = self.config_dir / "Plantilla_Conocimientos_y_Preferencias.xlsx"

        self.seq_path_lbl.setText(self.seq_path.name)
        self.pref_path_lbl.setText(self.pref_path.name)

        seq_ok = self.seq_path.exists()
        pref_ok = self.pref_path.exists()

        if seq_ok:
            self.seq_status_lbl.setText("Disponible")
            self.seq_status_lbl.setObjectName("badgeOk")
        else:
            self.seq_status_lbl.setText("No encontrado")
            self.seq_status_lbl.setObjectName("badgeError")

        if pref_ok:
            self.pref_status_lbl.setText("Disponible")
            self.pref_status_lbl.setObjectName("badgeOk")
        else:
            self.pref_status_lbl.setText("No encontrado")
            self.pref_status_lbl.setObjectName("badgeError")

        self.seq_status_lbl.style().unpolish(self.seq_status_lbl)
        self.seq_status_lbl.style().polish(self.seq_status_lbl)
        self.pref_status_lbl.style().unpolish(self.pref_status_lbl)
        self.pref_status_lbl.style().polish(self.pref_status_lbl)

        if not (seq_ok and pref_ok):
            self.generate_btn.setEnabled(False)
            self.status_lbl.setText("Faltan archivos en la carpeta de configuración.")
        else:
            self.generate_btn.setEnabled(True)
            self.status_lbl.setText("Listo para generar.")

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

        time_limit = float(self.time_spin.value())

        self.thread = SolverWorkerThread(
            year=year,
            month=month,
            start_cycle_day=start_cycle_day,
            seq_path=self.seq_path,
            pref_path=self.pref_path,
            out_dir=self.out_dir,
            time_limit=time_limit,
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
        self.status_lbl.setText(f"Cálculo completado ({result.status_name}).")

        self.last_excel_path = excel_path
        self.last_pdf_path = pdf_path

        free_cnt = sum(1 for a in result.assignments if a.position == ".")
        vac_cnt = len(result.vacancies)
        info_lines = [
            f"Estado: {result.status_name}   ·   Asignaciones: {len(result.assignments)}   ·   Libres: {free_cnt}   ·   Tiempo: {result.execution_time_seconds}s",
            f"Rango de esfuerzo: {result.effort_spread:.2f} pts/día  (Mín: {result.min_mean_effort:.2f}, Máx: {result.max_mean_effort:.2f})",
        ]
        if vac_cnt > 0:
            info_lines.append(f"Puestos descubiertos para revisión manual: {vac_cnt} (resaltados en el archivo Excel)")
        if result.issues:
            for issue in result.issues:
                info_lines.append(f"Nota: {issue}")
        self.results_info.setText("\n".join(info_lines))
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

    def _open_folder_path(self, folder_path: Path):
        if not folder_path.exists():
            folder_path.mkdir(parents=True, exist_ok=True)
        folder = str(folder_path.resolve())
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.run(["open", folder])
        else:
            subprocess.run(["xdg-open", folder])

    def _on_worker_data_saved(self):
        self._check_config_files()

    def _open_config_folder(self):
        self._open_folder_path(self.config_dir)
        self._check_config_files()
        if hasattr(self, "tab_trabajadores"):
            self.tab_trabajadores.load_data()

    def _open_folder(self):
        self._open_folder_path(self.out_dir)

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
    icon_path = APP_DIR / "reparto" / "assets" / "icono.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
