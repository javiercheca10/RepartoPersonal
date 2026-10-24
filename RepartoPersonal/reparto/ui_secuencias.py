"""Módulo de Interfaz Gráfica para Visualización y Ajuste de Secuencias de Turnos.

Permite ver la fila completa de turnos cíclicos de cada trabajador desde el Excel,
modificar marcas (M, T, D, R) en tiempo real con recálculo dinámico de totales,
y persistir los cambios directamente en el archivo de secuencias sin openpyxl.
"""

from pathlib import Path
from PySide6.QtCore import Qt, QTimer, Signal, QEvent
from PySide6.QtGui import QColor, QFont, QBrush
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QComboBox,
    QTableWidget,
    QTableView,
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
    dirty_changed = Signal(bool)

    def __init__(self, seq_path: Path, parent=None):
        super().__init__(parent)
        self.seq_path = Path(seq_path)
        self.records: list[dict] = []
        self.cycle_weeks: int = 24
        self.cycle_days: int = 168
        self.is_dirty: bool = False
        self.active_brush: str | None = None
        self.brush_buttons: dict[str | None, QPushButton] = {}

        self._init_ui()
        self.load_data()

    def _set_dirty(self, dirty: bool):
        dirty = bool(dirty)
        if self.is_dirty != dirty:
            self.is_dirty = dirty
            if hasattr(self, "save_btn"):
                self.save_btn.setText(
                    "💾 Guardar Cambios en Excel ●" if dirty else "💾 Guardar Cambios en Excel"
                )
            self.dirty_changed.emit(dirty)

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

        self.week_combo = QComboBox()
        self.week_combo.setToolTip("Saltar directamente a una semana concreta del ciclo o a los totales finales.")
        self.week_combo.setStyleSheet("""
            QComboBox {
                background-color: #F8FAFC;
                color: #0F172A;
                font-size: 13px;
                font-weight: 600;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 6px 12px;
                min-width: 175px;
            }
            QComboBox:hover {
                border-color: #94A3B8;
                background-color: #F1F5F9;
            }
        """)
        self.week_combo.currentIndexChanged.connect(self._on_week_selected)
        tools_row.addWidget(self.week_combo)

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
        self.reload_btn.clicked.connect(self._on_reload_clicked)
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

        # Fila 3: Modo Pincel / Clic rápido para editar turnos
        brush_row = QHBoxLayout()
        brush_row.setSpacing(8)
        brush_lbl = QLabel("Edición rápida:")
        brush_lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #475569;")
        brush_row.addWidget(brush_lbl)

        brush_specs = [
            (None, "⌨️ Teclado / Espacio", "Doble clic para escribir o Barra espaciadora para alternar M ➔ T ➔ D ➔ R"),
            ("M", "🖌️ Pintar M (Mañana)", "Haga clic en cualquier día para asignarle Mañana (M)"),
            ("T", "🖌️ Pintar T (Tarde)", "Haga clic en cualquier día para asignarle Tarde (T)"),
            ("D", "🖌️ Pintar D (Descanso)", "Haga clic en cualquier día para asignarle Descanso (D)"),
            ("R", "🖌️ Pintar R (Refuerzo)", "Haga clic en cualquier día para asignarle Refuerzo (R)"),
        ]
        for mode_key, btn_text, tip in brush_specs:
            btn = QPushButton(btn_text)
            btn.setToolTip(tip)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _checked=False, mk=mode_key: self._set_brush_mode(mk))
            self.brush_buttons[mode_key] = btn
            brush_row.addWidget(btn)

        brush_row.addStretch()
        h_layout.addLayout(brush_row)
        self._update_brush_buttons_style()

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
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.table.installEventFilter(self)

        # Vista superpuesta fija para mantener inmovilizadas las columnas N.º y Profesional al hacer scroll horizontal
        self.frozen_table = QTableView(self.table)
        self.frozen_table.setObjectName("frozenSeqTable")
        self.frozen_table.setStyleSheet("""
            QTableView#frozenSeqTable {
                background-color: #FFFFFF;
                border: none;
                border-right: 2px solid #CBD5E1;
                gridline-color: #F1F5F9;
                font-size: 12px;
                selection-background-color: #E0F2FE;
                selection-color: #0369A1;
            }
            QTableView#frozenSeqTable QHeaderView::section {
                background-color: #F8FAFC;
                color: #475569;
                font-weight: 600;
                font-size: 11px;
                border: 1px solid #E2E8F0;
                padding: 4px;
            }
        """)
        self.frozen_table.setAlternatingRowColors(True)
        self.frozen_table.setFocusPolicy(Qt.NoFocus)
        self.frozen_table.verticalHeader().hide()
        self.frozen_table.horizontalHeader().setFixedHeight(self.table.horizontalHeader().height())
        self.frozen_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.frozen_table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)
        self.table.verticalScrollBar().valueChanged.connect(self.frozen_table.verticalScrollBar().setValue)
        self.frozen_table.verticalScrollBar().valueChanged.connect(self.table.verticalScrollBar().setValue)
        self.table.horizontalHeader().sectionResized.connect(self._on_main_section_resized)

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

    def _on_reload_clicked(self):
        if self.is_dirty:
            reply = QMessageBox.question(
                self,
                "Cambios sin guardar",
                "Hay modificaciones sin guardar en la pestaña Secuencias de Turnos.\n"
                "¿Desea descartarlas y recargar desde el archivo Excel?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
        self.load_data()

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
            self._set_dirty(False)
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
        self._populate_week_combo()
        self._sync_frozen_table()
        self._on_search_changed(self.search_input.text())

    def _set_brush_mode(self, mode_key: str | None):
        """Activa o desactiva el modo pincel para pintar turnos con un solo clic."""
        self.active_brush = mode_key
        self._update_brush_buttons_style()

    def _update_brush_buttons_style(self):
        """Actualiza el aspecto visual de los botones de edición rápida (pincel)."""
        palette = {
            None: ("#F8FAFC", "#334155", "#CBD5E1", "#0F172A", "#FFFFFF", "#0F172A"),
            "M": ("#F0F9FF", "#0369A1", "#BAE6FD", "#0284C7", "#FFFFFF", "#0369A1"),
            "T": ("#FFF7ED", "#C2410C", "#FED7AA", "#EA580C", "#FFFFFF", "#C2410C"),
            "D": ("#F8FAFC", "#475569", "#CBD5E1", "#475569", "#FFFFFF", "#334155"),
            "R": ("#F0FDF4", "#15803D", "#BBF7D0", "#16A34A", "#FFFFFF", "#15803D"),
        }
        for mk, btn in self.brush_buttons.items():
            bg, fg, border, act_bg, act_fg, act_border = palette[mk]
            if self.active_brush == mk:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {act_bg};
                        color: {act_fg};
                        border: 1.5px solid {act_border};
                        border-radius: 6px;
                        padding: 4px 10px;
                        font-size: 12px;
                        font-weight: 700;
                    }}
                """)
            else:
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {bg};
                        color: {fg};
                        border: 1px solid {border};
                        border-radius: 6px;
                        padding: 4px 10px;
                        font-size: 12px;
                        font-weight: 600;
                    }}
                    QPushButton:hover {{
                        border-color: {act_border};
                    }}
                """)

    def _populate_week_combo(self):
        """Rellena el selector de salto rápido por semana del ciclo."""
        self.week_combo.blockSignals(True)
        self.week_combo.clear()
        for w in range(1, self.cycle_weeks + 1):
            d_start = (w - 1) * 7 + 1
            d_end = w * 7
            self.week_combo.addItem(f"📅 Ir a: Semana {w} (Días {d_start}–{d_end})", w)
        self.week_combo.addItem("📊 Ir a: Totales finales", -1)
        self.week_combo.blockSignals(False)

    def _on_week_selected(self, idx: int):
        """Desplaza horizontalmente la tabla hasta el inicio de la semana seleccionada."""
        if idx < 0 or self.table.rowCount() == 0:
            return
        val = self.week_combo.itemData(idx)
        vis_row = 0
        for r in range(self.table.rowCount()):
            if not self.table.isRowHidden(r):
                vis_row = r
                break
        if val == -1:
            target_col = max(2, self.table.columnCount() - 1)
            item = self.table.item(vis_row, target_col)
            if item:
                self.table.scrollToItem(item, QAbstractItemView.PositionAtCenter)
        elif isinstance(val, int) and val >= 1:
            target_col = 2 + (val - 1) * 7
            # Calcular desplazamiento exacto en píxeles para que el lunes de esa semana quede pegado a la columna fija
            x_offset = sum(self.table.columnWidth(c) for c in range(2, target_col))
            self.table.horizontalScrollBar().setValue(x_offset)

    def _sync_frozen_table(self):
        """Sincroniza el modelo y las columnas visibles de la tabla izquierda inmovilizada (N.º y Profesional)."""
        self.frozen_table.setModel(self.table.model())
        self.frozen_table.setSelectionModel(self.table.selectionModel())
        total_cols = self.table.columnCount()
        for c in range(total_cols):
            self.frozen_table.setColumnHidden(c, c >= 2)
        if total_cols >= 2:
            self.frozen_table.setColumnWidth(0, self.table.columnWidth(0))
            self.frozen_table.setColumnWidth(1, self.table.columnWidth(1))
        for r in range(self.table.rowCount()):
            self.frozen_table.setRowHeight(r, self.table.rowHeight(r))
        self._update_frozen_geometry()
        self.frozen_table.show()

    def _on_main_section_resized(self, logical_index: int, _old_size: int, new_size: int):
        if logical_index in (0, 1):
            self.frozen_table.setColumnWidth(logical_index, new_size)
            self._update_frozen_geometry()

    def _update_frozen_geometry(self):
        if self.table.columnCount() < 2:
            return
        frozen_w = self.table.columnWidth(0) + self.table.columnWidth(1)
        frame_w = self.table.frameWidth()
        vh_w = self.table.verticalHeader().width() if self.table.verticalHeader().isVisible() else 0
        hh_h = self.table.horizontalHeader().height()
        self.frozen_table.horizontalHeader().setFixedHeight(hh_h)
        self.frozen_table.setGeometry(
            vh_w + frame_w,
            frame_w,
            frozen_w,
            self.table.viewport().height() + hh_h,
        )

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_frozen_geometry()

    def _on_cell_clicked(self, row: int, col: int):
        """Si hay un pincel de turno activo, pinta la celda con un solo clic."""
        if self.active_brush in ("M", "T", "D", "R") and 2 <= col < 2 + self.cycle_days:
            item = self.table.item(row, col)
            if item and item.text() != self.active_brush:
                item.setText(self.active_brush)

    def eventFilter(self, obj, event):
        """Permite alternar M ➔ T ➔ D ➔ R pulsando la barra espaciadora sobre celdas de días."""
        if obj is self.table and event.type() == QEvent.KeyPress and event.key() == Qt.Key_Space:
            cycle_order = {"M": "T", "T": "D", "D": "R", "R": "M"}
            selected = self.table.selectedItems()
            day_items = [it for it in selected if 2 <= it.column() < 2 + self.cycle_days]
            if day_items:
                for it in day_items:
                    cur = it.text().strip().upper()
                    it.setText(cycle_order.get(cur, "M"))
                return True
        return super().eventFilter(obj, event)

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
            prev = (
                self.records[row]["sequence"][d_idx]
                if 0 <= row < len(self.records) and d_idx < len(self.records[row]["sequence"])
                else "D"
            )
            if raw == "":
                raw = "D"
            elif raw not in ("M", "T", "D", "R"):
                self.table.blockSignals(True)
                item.setText(prev)
                self._style_cell_item(item, prev, d_idx)
                self.table.blockSignals(False)
                self._show_notification(
                    f"Marca '{raw}' no válida: use M (Mañana), T (Tarde), D (Descanso) o R (Refuerzo).",
                    is_error=True,
                )
                return

            self.table.blockSignals(True)
            item.setText(raw)
            self._style_cell_item(item, raw, d_idx)
            self.table.blockSignals(False)

            if 0 <= row < len(self.records):
                seq = self.records[row]["sequence"]
                if d_idx < len(seq):
                    seq[d_idx] = raw
                    if prev != raw:
                        self._set_dirty(True)
                self._update_row_totals(row, self.records[row])

    def _on_search_changed(self, text: str):
        """Filtra las filas visibles según el texto de búsqueda."""
        query = text.strip().lower()
        for row in range(self.table.rowCount()):
            if not query:
                self.table.setRowHidden(row, False)
                self.frozen_table.setRowHidden(row, False)
                continue
            name_item = self.table.item(row, 1)
            num_item = self.table.item(row, 0)
            name = name_item.text().lower() if name_item else ""
            num = num_item.text().lower() if num_item else ""
            matches = (query in name) or (query in num)
            self.table.setRowHidden(row, not matches)
            self.frozen_table.setRowHidden(row, not matches)

    def save_data(self) -> bool:
        """Persiste los cambios de secuencias al archivo Excel."""
        if not self.seq_path.exists():
            QMessageBox.critical(self, "Error", f"No se encontró el archivo de secuencias: {self.seq_path}")
            return False

        try:
            save_sequence_records(self.seq_path, self.records, cycle_weeks=self.cycle_weeks)
            self._set_dirty(False)
            self._show_notification(f"✔ Cambios guardados con éxito en {self.seq_path.name}.")
            self.data_saved.emit()
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", f"No se pudieron guardar las secuencias:\n{str(e)}")
            return False

    def _on_save_clicked(self):
        self.save_data()
