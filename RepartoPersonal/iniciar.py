"""Punto de entrada de la aplicación portable RepartoRadiologia.

Funciona de forma idéntica tanto si se ejecuta desde Python como si está
empaquetado en un ejecutable .exe con PyInstaller.
Detecta automáticamente las carpetas 'configuracion' y 'resultados'.
"""

from dataclasses import replace
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

from reparto.assets import get_app_icon, get_app_pixmap, setup_windows_app_id
from reparto.datos import (
    cycle_day_matches_weekday,
    detect_cycle_weeks,
    holidays_for_year,
    load_all_workers_with_report,
    parse_positions_from_header,
    parse_position_rules,
    project_month,
    suggest_cycle_day,
)
from reparto.exportadores import export_month_to_excel, export_month_to_pdf
from reparto.solver import MonthResult, solve_month
from reparto.ui_secuencias import SecuenciasWidget
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
        cycle_weeks: int = 24,
    ):
        super().__init__()
        self.year = year
        self.month = month
        self.start_cycle_day = start_cycle_day
        self.seq_path = seq_path
        self.pref_path = pref_path
        self.out_dir = out_dir
        self.time_limit = time_limit
        self.cycle_weeks = cycle_weeks

    def run(self):
        try:
            self.progress_signal.emit("Cargando archivos de configuración y puestos...", 15)
            cycle_length = self.cycle_weeks * 7
            positions = parse_positions_from_header(self.pref_path)
            rules = parse_position_rules(self.pref_path, positions)
            workers, load_report = load_all_workers_with_report(
                self.seq_path,
                self.pref_path,
                positions=positions,
                cycle_weeks=self.cycle_weeks,
            )

            self.progress_signal.emit("Proyectando calendario laboral y turnos...", 35)
            schedules = project_month(
                self.year,
                self.month,
                self.start_cycle_day,
                workers,
                positions=positions,
                position_rules=rules,
                cycle_length=cycle_length,
            )

            self.progress_signal.emit("Optimizando asignación y equidad de puestos con CP-SAT...", 60)
            result = solve_month(
                self.year,
                self.month,
                schedules,
                workers,
                positions=positions,
                position_rules=rules,
                time_limit_seconds=self.time_limit,
            )

            if not result.is_optimal and len(result.assignments) == 0:
                self.error_signal.emit(
                    "No se pudo encontrar una asignación que cubra los mínimos. Compruebe la plantilla de personal."
                )
                return

            extra_issues = list(load_report.warnings)
            _, holidays_verified = holidays_for_year(self.year)
            if not holidays_verified:
                extra_issues.append(
                    f"Los festivos de {self.year} se han calculado por regla general (revise festivos locales)."
                )

            self.progress_signal.emit("Generando cuadrante Excel (.xlsx)...", 80)
            self.out_dir.mkdir(parents=True, exist_ok=True)
            month_name = MONTHS_ES[self.month - 1][0]
            excel_path = self.out_dir / f"Cuadrante_{month_name}_{self.year}.xlsx"
            export_month_to_excel(result, schedules, workers, excel_path, positions=positions)

            self.progress_signal.emit("Generando cuadrante PDF oficial...", 95)
            pdf_str = ""
            try:
                pdf_path = self.out_dir / f"Cuadrante_{month_name}_{self.year}.pdf"
                export_month_to_pdf(result, schedules, workers, pdf_path, positions=positions)
                pdf_str = str(pdf_path)
            except Exception as pdf_err:
                extra_issues.append(
                    f"No se pudo generar el PDF ({pdf_err}); el archivo Excel está completo."
                )

            if extra_issues:
                result = replace(result, issues=tuple(extra_issues) + result.issues)

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

        # Identidad visual e icono de ventana
        self.setWindowIcon(get_app_icon())

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

        icon_pixmap = get_app_pixmap(48)
        if not icon_pixmap.isNull():
            icon_lbl = QLabel()
            icon_lbl.setPixmap(icon_pixmap)
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
        self.seq_title = QLabel("Secuencias de Turnos")
        self.seq_title.setObjectName("itemTitle")
        self.seq_path_lbl = QLabel(self.seq_path.name)
        self.seq_path_lbl.setObjectName("pathLabel")
        seq_name_box.addWidget(self.seq_title)
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
        self.month_combo.currentIndexChanged.connect(self._on_period_changed)
        mes_box.addWidget(mes_lbl)
        mes_box.addWidget(self.month_combo)
        form_layout.addLayout(mes_box)

        year_box = QVBoxLayout()
        year_box.setSpacing(6)
        year_lbl = QLabel("Año")
        year_lbl.setObjectName("fieldLabel")
        self.year_combo = QComboBox()
        for y in range(2025, 2036):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(today.year if 2025 <= today.year <= 2035 else 2026))
        self.year_combo.currentIndexChanged.connect(self._on_period_changed)
        year_box.addWidget(year_lbl)
        year_box.addWidget(self.year_combo)
        form_layout.addLayout(year_box)

        cycle_box = QVBoxLayout()
        cycle_box.setSpacing(6)
        self.cycle_lbl = QLabel("Día del Ciclo en Día 1")
        self.cycle_lbl.setObjectName("fieldLabel")
        self.cycle_combo = QComboBox()
        self.cycle_combo.setEditable(True)
        self.cycle_combo.setToolTip("Indique qué día del ciclo corresponde al día 1 del mes a planificar.")
        cycle_box.addWidget(self.cycle_lbl)
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
        self.tab_trabajadores = TrabajadoresTab(self.pref_path, seq_path=self.seq_path)
        self.tab_trabajadores.data_saved.connect(self._on_worker_data_saved)
        self.tab_trabajadores.dirty_changed.connect(
            lambda dirty: self.tab_widget.setTabText(1, "Trabajadores *" if dirty else "Trabajadores")
        )
        self.tab_widget.addTab(self.tab_trabajadores, "Trabajadores")

        # --- PESTAÑA 3: SECUENCIAS DE TURNOS ---
        self.tab_secuencias = SecuenciasWidget(self.seq_path)
        self.tab_secuencias.data_saved.connect(self._on_sequence_data_saved)
        self.tab_secuencias.dirty_changed.connect(
            lambda dirty: self.tab_widget.setTabText(2, "Secuencias de Turnos *" if dirty else "Secuencias de Turnos")
        )
        self.tab_widget.addTab(self.tab_secuencias, "Secuencias de Turnos")

        self.tab_widget.currentChanged.connect(self._on_tab_changed)

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
            cycle_weeks = detect_cycle_weeks(self.seq_path)
            cycle_days = cycle_weeks * 7
            self.seq_title.setText(f"Secuencias de Turnos (Ciclo de {cycle_weeks} semanas · {cycle_days} días)")
            self.cycle_lbl.setText(f"Día del Ciclo en Día 1 (1 a {cycle_days})")
            self.cycle_combo.setToolTip(
                f"Indique qué día del ciclo de {cycle_weeks} semanas (1 a {cycle_days}) corresponde al día 1 del mes a planificar."
            )
            if self.cycle_combo.count() != cycle_days:
                cur_val = self.cycle_combo.currentData()
                self.cycle_combo.blockSignals(True)
                self.cycle_combo.clear()
                for d in range(1, cycle_days + 1):
                    self.cycle_combo.addItem(str(d), d)
                m = self.month_combo.currentData()
                y = self.year_combo.currentData()
                suggested = suggest_cycle_day(y, m, cycle_length=cycle_days)
                idx = self.cycle_combo.findData(suggested)
                if idx >= 0:
                    self.cycle_combo.setCurrentIndex(idx)
                elif cur_val and cur_val <= cycle_days:
                    self.cycle_combo.setCurrentIndex(cur_val - 1)
                self.cycle_combo.blockSignals(False)
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
            self.status_lbl.setToolTip("")
        else:
            self.generate_btn.setEnabled(True)
            try:
                _, report = load_all_workers_with_report(self.seq_path, self.pref_path)
                if report.warnings:
                    self.status_lbl.setText(
                        f"Listo para generar · {len(report.warnings)} aviso(s) en configuración (pase el ratón para ver)"
                    )
                    self.status_lbl.setToolTip("\n".join(f"• {w}" for w in report.warnings))
                else:
                    self.status_lbl.setText("Listo para generar.")
                    self.status_lbl.setToolTip("")
            except Exception:
                self.status_lbl.setText("Listo para generar.")
                self.status_lbl.setToolTip("")

    def _prompt_save_unsaved_changes(self, action_desc: str) -> bool:
        """Verifica si hay cambios sin guardar en Trabajadores o Secuencias antes de una acción crítica."""
        dirty_tabs = []
        if hasattr(self, "tab_trabajadores") and self.tab_trabajadores.is_dirty:
            dirty_tabs.append("Trabajadores")
        if hasattr(self, "tab_secuencias") and self.tab_secuencias.is_dirty:
            dirty_tabs.append("Secuencias de Turnos")
        if not dirty_tabs:
            return True

        tabs_str = " y ".join(dirty_tabs)
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("Cambios sin guardar")
        msg_box.setText(f"Hay cambios sin guardar en {tabs_str}.")
        msg_box.setInformativeText(f"¿Desea guardarlos en Excel antes de {action_desc}?")
        save_btn = msg_box.addButton("Guardar y continuar", QMessageBox.AcceptRole)
        discard_btn = msg_box.addButton("Continuar sin guardar", QMessageBox.DestructiveRole)
        cancel_btn = msg_box.addButton("Cancelar", QMessageBox.RejectRole)
        msg_box.setDefaultButton(save_btn)
        msg_box.exec()

        clicked = msg_box.clickedButton()
        if clicked == cancel_btn:
            return False
        if clicked == save_btn:
            if hasattr(self, "tab_trabajadores") and self.tab_trabajadores.is_dirty:
                if not self.tab_trabajadores.save_data():
                    return False
            if hasattr(self, "tab_secuencias") and self.tab_secuencias.is_dirty:
                if not self.tab_secuencias.save_data():
                    return False
        elif clicked == discard_btn:
            if hasattr(self, "tab_trabajadores") and self.tab_trabajadores.is_dirty:
                self.tab_trabajadores.load_data()
            if hasattr(self, "tab_secuencias") and self.tab_secuencias.is_dirty:
                self.tab_secuencias.load_data()
        return True

    def _on_period_changed(self):
        """Actualiza la sugerencia de día de ciclo al cambiar de mes o año."""
        cycle_weeks = detect_cycle_weeks(self.seq_path) if self.seq_path.exists() else 24
        cycle_days = cycle_weeks * 7
        m = self.month_combo.currentData()
        y = self.year_combo.currentData()
        if m and y:
            suggested = suggest_cycle_day(y, m, cycle_length=cycle_days)
            idx = self.cycle_combo.findData(suggested)
            if idx >= 0:
                self.cycle_combo.setCurrentIndex(idx)

    def _on_generate_clicked(self):
        if not self._prompt_save_unsaved_changes("generar el cuadrante"):
            return

        self.generate_btn.setEnabled(False)
        self.results_card.setVisible(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(5)
        self.status_lbl.setText("Iniciando cálculo...")

        cycle_weeks = detect_cycle_weeks(self.seq_path) if self.seq_path.exists() else 24
        cycle_length = cycle_weeks * 7

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

        start_cycle_day = max(1, min(cycle_length, start_cycle_day))

        # Validar si el día del ciclo encaja con el día de la semana real del día 1
        if not cycle_day_matches_weekday(year, month, start_cycle_day, cycle_length=cycle_length):
            suggested = suggest_cycle_day(year, month, cycle_length=cycle_length)
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
            cycle_weeks=cycle_weeks,
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
        if hasattr(self, "tab_secuencias") and not self.tab_secuencias.is_dirty:
            self.tab_secuencias.load_data()

    def _on_sequence_data_saved(self):
        self._check_config_files()

    def _on_tab_changed(self, index: int):
        if index == 0:
            self._check_config_files()
        elif index == 1 and hasattr(self, "tab_trabajadores"):
            if not self.tab_trabajadores.is_dirty:
                self.tab_trabajadores.load_data()
        elif index == 2 and hasattr(self, "tab_secuencias"):
            if not self.tab_secuencias.is_dirty:
                self.tab_secuencias.load_data()

    def closeEvent(self, event):
        if self._prompt_save_unsaved_changes("cerrar la aplicación"):
            event.accept()
        else:
            event.ignore()

    def _open_config_folder(self):
        self._open_folder_path(self.config_dir)
        self._check_config_files()
        if hasattr(self, "tab_trabajadores") and not self.tab_trabajadores.is_dirty:
            self.tab_trabajadores.load_data()
        if hasattr(self, "tab_secuencias") and not self.tab_secuencias.is_dirty:
            self.tab_secuencias.load_data()

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
    setup_windows_app_id()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setWindowIcon(get_app_icon())
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
