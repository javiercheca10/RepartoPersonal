"""Módulo de Interfaz Gráfica para Visualización y Ajuste de Secuencias de Turnos.

Permite ver la fila completa de turnos cíclicos de cada trabajador desde el Excel,
modificar marcas (M, T, D, R) en tiempo real con recálculo dinámico de totales,
y persistir los cambios directamente en el archivo de secuencias sin openpyxl.
"""

from pathlib import Path
from PySide6.QtCore import Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QBrush
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QLineEdit,
    QFrame,
    QHeaderView,
    QAbstractItemView,
    QMessageBox,
)

from reparto.datos import (
    detect_cycle_weeks,
    load_sequence_records,
    save_sequence_records,
)

# Paleta de colores para los turnos
SHIFT_COLORS = {
    "M": {"bg": QColor("#E0F2FE"), "fg": QColor("#0369A1"), "desc": "Mañana"},
    "T": {"bg": QColor("#FFEDD5"), "fg": QColor("#C2410C"), "desc": "Tarde"},
    "D": {"bg": QColor("#F8FAFC"), "fg": QColor("#94A3B8"), "desc": "Descanso"},
    "R": {"bg": QColor("#DCFCE7"), "fg": QColor("#15803D"), "desc": "Refuerzo"},
}
WEEKEND_D_BG = QColor("#F1F5F9")
DOW_NAMES = ["L", "M", "X", "J", "V", "S", "D"]
DOW_FULL = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]


class SecuenciasWidget(QWidget):
    """Pestaña interactiva de ajuste y visualización de secuencias de turnos."""

    data_saved = Signal()

    def __init__(self, seq_path: Path, parent=None):
        super().__init__(parent)
        self.seq_path = Path(seq_path)
        self.records: list[dict] = []
        self.cycle_weeks: int = 24
        self.cycle_days: int = 168

        self._init_ui()
        self.load_data()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 14, 0, 0)
        layout.setSpacing(14)

        # 1. Tarjeta de Cabecera y Herramientas
        header_card = QFrame()
        header_card.setObjectName("headerCard")
        header_card.setStyleSheet("""
            #headerCard {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 16px 20px;
            }
        """)
        h_layout = QVBoxLayout(header_card)
        h_layout.setSpacing(12)

        # Fila superior: Título, badges informativos
        top_row = QHBoxLayout()
        top_row.setSpacing(12)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        title_lbl = QLabel("Ajuste de Secuencias de Turnos")
        title_lbl.setStyleSheet("font-size: 17px; font-weight: 700; color: #0F172A;")
        sub_lbl = QLabel(
            "Visualice y edite la secuencia cíclica por trabajador "
            "(M = Mañana, T = Tarde, D = Descanso, R = Refuerzo)."
        )
        sub_lbl.setStyleSheet("font-size: 12px; color: #64748B;")
        title_box.addWidget(title_lbl)
        title_box.addWidget(sub_lbl)
        top_row.addLayout(title_box, 1)

        self.cycle_badge = QLabel("⏱️ Ciclo: Calculando...")
        self.cycle_badge.setStyleSheet("""
            background-color: #F0FDF4;
            color: #15803D;
            font-size: 12px;
            font-weight: 600;
            padding: 5px 12px;
            border-radius: 6px;
            border: 1px solid #BBF7D0;
        """)
        top_row.addWidget(self.cycle_badge)

        self.workers_badge = QLabel("👥 0 profesionales")
        self.workers_badge.setStyleSheet("""
            background-color: #F0F9FF;
            color: #0369A1;
            font-size: 12px;
            font-weight: 600;
            padding: 5px 12px;
            border-radius: 6px;
            border: 1px solid #BAE6FD;
        """)
        top_row.addWidget(self.workers_badge)

        h_layout.addLayout(top_row)

        # Fila de herramientas: Buscador + Botones Guardar y Recargar
        tools_row = QHBoxLayout()
        tools_row.setSpacing(10)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Buscar profesional por nombre...")
        self.search_input.setStyleSheet("""
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 13px;
                color: #0F172A;
            }
            QLineEdit:focus {
                border-color: #0284C7;
            }
        """)
        self.search_input.textChanged.connect(self._on_search_changed)
        tools_row.addWidget(self.search_input, 1)

        self.reload_btn = QPushButton("🔄 Recargar")
        self.reload_btn.setToolTip("Descarta cambios sin guardar y recarga las secuencias desde el archivo Excel.")
        self.reload_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #334155;
                font-size: 13px;
                font-weight: 600;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 7px 14px;
            }
            QPushButton:hover {
                background-color: #F8FAFC;
                border-color: #94A3B8;
            }
        """)
        self.reload_btn.clicked.connect(self.load_data)
        tools_row.addWidget(self.reload_btn)

        self.save_btn = QPushButton("💾 Guardar Cambios en Excel")
        self.save_btn.setToolTip("Guarda las secuencias editadas en el archivo Excel correspondiente.")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: #FFFFFF;
                font-size: 13px;
                font-weight: 600;
                border: none;
                border-radius: 6px;
                padding: 7px 18px;
            }
            QPushButton:hover {
                background-color: #0369A1;
            }
        """)
        self.save_btn.clicked.connect(self._on_save_clicked)
        tools_row.addWidget(self.save_btn)

        h_layout.addLayout(tools_row)
        layout.addWidget(header_card)

        # 2. Notificación temporal auto-desvaneciente
        self.notification_lbl = QLabel()
        self.notification_lbl.setVisible(False)
        self.notification_lbl.setStyleSheet("""
            background-color: #DCFCE7;
            color: #15803D;
            border: 1px solid #86EFAC;
            border-radius: 6px;
            padding: 8px 14px;
            font-size: 13px;
            font-weight: 600;
        """)
        layout.addWidget(self.notification_lbl)

        self._notif_timer = QTimer(self)
        self._notif_timer.setSingleShot(True)
        self._notif_timer.timeout.connect(lambda: self.notification_lbl.setVisible(False))

        # 3. Tabla de Secuencias
        self.table = QTableWidget()
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                gridline-color: #F1F5F9;
                font-size: 12px;
                selection-background-color: #E0F2FE;
                selection-color: #0369A1;
            }
            QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                font-weight: 600;
                font-size: 11px;
                border: 1px solid #E2E8F0;
                padding: 4px;
            }
        """)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.itemChanged.connect(self._on_item_changed)

        layout.addWidget(self.table, 1)

    def _show_notification(self, msg: str, is_error: bool = False):
        if is_error:
            self.notification_lbl.setStyleSheet("""
                background-color: #FEF2F2;
                color: #B91C1C;
                border: 1px solid #FECACA;
                border-radius: 6px;
                padding: 8px 14px;
                font-size: 13px;
                font-weight: 600;
            """)
        else:
            self.notification_lbl.setStyleSheet("""
                background-color: #DCFCE7;
                color: #15803D;
                border: 1px solid #86EFAC;
                border-radius: 6px;
                padding: 8px 14px;
                font-size: 13px;
                font-weight: 600;
            """)
        self.notification_lbl.setText(msg)
        self.notification_lbl.setVisible(True)
        self._notif_timer.start(4000)

    def load_data(self):
        """Carga los datos de secuencias desde el Excel y rellena la tabla."""
        if not self.seq_path.exists():
            self._show_notification(f"No se encuentra el archivo: {self.seq_path.name}", is_error=True)
            return

        try:
            self.cycle_weeks = detect_cycle_weeks(self.seq_path)
            self.cycle_days = self.cycle_weeks * 7
            self.records, _ = load_sequence_records(self.seq_path, cycle_weeks=self.cycle_weeks)

            self.cycle_badge.setText(f"⏱️ Ciclo detectado: {self.cycle_weeks} semanas ({self.cycle_days} días)")
            self.workers_badge.setText(f"👥 {len(self.records)} profesionales")

            self._populate_table()
            self._show_notification("✔ Secuencias cargadas correctamente.")
        except Exception as e:
            self._show_notification(f"Error al cargar secuencias: {str(e)}", is_error=True)

    def _populate_table(self):
        """Reconstruye las columnas y filas de la tabla de secuencias."""
        self.table.blockSignals(True)

        num_days = self.cycle_days
        total_cols = 2 + num_days + 5  # Nº, Nombre + Días + (M, T, D, R, Total)
        self.table.setColumnCount(total_cols)
        self.table.setRowCount(len(self.records))

        # Encabezados
        headers = ["N.º", "Profesional"]
        for d in range(1, num_days + 1):
            dow_idx = (d - 1) % 7
            dow = DOW_NAMES[dow_idx]
            headers.append(f"{dow}\n{d}")
        headers.extend(["Mañanas (M)", "Tardes (T)", "Descansos (D)", "Refuerzos (R)", "Total Días"])

        self.table.setHorizontalHeaderLabels(headers)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        for c in range(2, 2 + num_days):
            self.table.setColumnWidth(c, 36)
        for c in range(2 + num_days, total_cols):
            self.table.horizontalHeader().setSectionResizeMode(c, QHeaderView.ResizeToContents)

        # Tooltips de cabeceras de día
        for d in range(1, num_days + 1):
            w = (d - 1) // 7 + 1
            dow_idx = (d - 1) % 7
            item = self.table.horizontalHeaderItem(1 + d)
            if item:
                item.setToolTip(f"Semana {w} · Día {d} ({DOW_FULL[dow_idx]})")

        bold_font = QFont()
        bold_font.setBold(True)

        # Rellenar filas
        for row_idx, rec in enumerate(self.records):
            # Col 0: N.º
            item_num = QTableWidgetItem(str(rec.get("num", row_idx + 1)))
            item_num.setTextAlignment(Qt.AlignCenter)
            item_num.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(row_idx, 0, item_num)

            # Col 1: Profesional
            item_name = QTableWidgetItem(rec.get("name", ""))
            item_name.setFont(bold_font)
            item_name.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            self.table.setItem(row_idx, 1, item_name)

            # Columnas de días
            seq = rec.get("sequence", [])
            for d_idx in range(num_days):
                col_idx = 2 + d_idx
                mark = seq[d_idx] if d_idx < len(seq) else "D"
                c_item = QTableWidgetItem(mark)
                c_item.setTextAlignment(Qt.AlignCenter)
                c_item.setFont(bold_font)
                self._style_cell_item(c_item, mark, d_idx)
                self.table.setItem(row_idx, col_idx, c_item)

            # Columnas de totales
            self._update_row_totals(row_idx, rec)

        self.table.blockSignals(False)
        self._on_search_changed(self.search_input.text())

    def _style_cell_item(self, item: QTableWidgetItem, mark: str, d_idx: int):
        """Aplica colores y formato a una celda de marca según su valor y día."""
        dow_idx = d_idx % 7
        is_weekend = dow_idx in (5, 6)

        color_info = SHIFT_COLORS.get(mark, SHIFT_COLORS["D"])
        bg = color_info["bg"]
        if mark == "D" and is_weekend:
            bg = WEEKEND_D_BG

        item.setBackground(QBrush(bg))
        item.setForeground(QBrush(color_info["fg"]))
        w = (d_idx // 7) + 1
        item.setToolTip(f"Semana {w} · {DOW_FULL[dow_idx]} (Día {d_idx + 1}): {color_info['desc']} ({mark})")

    def _update_row_totals(self, row_idx: int, rec: dict):
        """Calcula y actualiza las columnas de resumen al final de la fila."""
        seq = rec.get("sequence", [])[:self.cycle_days]
        m = seq.count("M")
        t = seq.count("T")
        d = seq.count("D")
        r = seq.count("R")
        tot = len(seq)

        rec["m"] = m
        rec["t"] = t
        rec["d"] = d
        rec["r"] = r
        rec["total"] = tot

        bold_font = QFont()
        bold_font.setBold(True)

        base_col = 2 + self.cycle_days
        vals = [
            (m, QColor("#F0F9FF"), QColor("#0284C7")),
            (t, QColor("#FFF7ED"), QColor("#EA580C")),
            (d, QColor("#F8FAFC"), QColor("#64748B")),
            (r, QColor("#F0FDF4"), QColor("#16A34A")),
            (tot, QColor("#F8FAFC"), QColor("#0F172A")),
        ]

        for idx, (val, bg, fg) in enumerate(vals):
            col = base_col + idx
            item = self.table.item(row_idx, col)
            if not item:
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                item.setFont(bold_font)
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                self.table.setItem(row_idx, col, item)
            else:
                item.setText(str(val))
            item.setBackground(QBrush(bg))
            item.setForeground(QBrush(fg))

    def _on_item_changed(self, item: QTableWidgetItem):
        """Maneja la edición directa de marcas de turno por el usuario."""
        row = item.row()
        col = item.column()

        # Solo procesar celdas de días
        if 2 <= col < 2 + self.cycle_days:
            d_idx = col - 2
            raw = item.text().strip().upper()
            if raw not in ("M", "T", "D", "R"):
                raw = "D"

            self.table.blockSignals(True)
            item.setText(raw)
            self._style_cell_item(item, raw, d_idx)
            self.table.blockSignals(False)

            if 0 <= row < len(self.records):
                seq = self.records[row]["sequence"]
                if d_idx < len(seq):
                    seq[d_idx] = raw
                self._update_row_totals(row, self.records[row])

    def _on_search_changed(self, text: str):
        """Filtra las filas visibles según el texto de búsqueda."""
        query = text.strip().lower()
        for row in range(self.table.rowCount()):
            if not query:
                self.table.setRowHidden(row, False)
                continue
            name_item = self.table.item(row, 1)
            num_item = self.table.item(row, 0)
            name = name_item.text().lower() if name_item else ""
            num = num_item.text().lower() if num_item else ""
            matches = (query in name) or (query in num)
            self.table.setRowHidden(row, not matches)

    def _on_save_clicked(self):
        """Persiste los cambios de secuencias al archivo Excel."""
        if not self.seq_path.exists():
            QMessageBox.critical(self, "Error", f"No se encontró el archivo de secuencias: {self.seq_path}")
            return

        try:
            save_sequence_records(self.seq_path, self.records, cycle_weeks=self.cycle_weeks)
            self._show_notification(f"✔ Cambios guardados con éxito en {self.seq_path.name}.")
            self.data_saved.emit()
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", f"No se pudieron guardar las secuencias:\n{str(e)}")
