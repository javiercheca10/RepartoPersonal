"""Componentes y vista de la pestaña 'Trabajadores' (Maestro-Detalle).

Proporciona:
- Deslizador interactivo de estado de baja (Verde Disponible / Rojo De baja).
- Chips interactivos de competencias de 3 estados (No sabe / Sabe / Vetado).
- Lista reordenable mediante Drag & Drop para las preferencias de puestos.
- Navegación maestro-detalle con buscador en tiempo real y persistencia en Excel.
"""

from pathlib import Path

from PySide6.QtCore import (
    QEasingCurve,
    QPropertyAnimation,
    QRectF,
    Qt,
    QTimer,
    Property,
    Signal,
)
from PySide6.QtGui import (
    QBrush,
    QColor,
    QFont,
    QPainter,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from reparto.datos import DEFAULT_POSITION_RULES, POSITIONS, POSITION_RULES, sync_sequences_roster
from reparto.gestion_trabajadores import (
    EditableWorker,
    load_editable_workers,
    save_editable_workers,
)


class StatusSwitch(QWidget):
    """Interruptor deslizante animado para el estado laboral.

    - False (Disponible): Carril verde (#10B981), bola a la derecha, texto 'Disponible'.
    - True (De baja): Carril rojo (#EF4444), bola a la izquierda, texto 'De baja'.
    """
    toggled = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(142, 34)
        self.setCursor(Qt.PointingHandCursor)
        self._is_baja = False
        self._thumb_position = 0.0  # 0.0 = Disponible (derecha), 1.0 = De baja (izquierda)
        self._anim = QPropertyAnimation(self, b"thumb_position", self)
        self._anim.setDuration(160)
        self._anim.setEasingCurve(QEasingCurve.OutCubic)

    def get_thumb_position(self) -> float:
        return self._thumb_position

    def set_thumb_position(self, pos: float):
        self._thumb_position = pos
        self.update()

    thumb_position = Property(float, get_thumb_position, set_thumb_position)

    def is_baja(self) -> bool:
        return self._is_baja

    def set_baja(self, value: bool, animated: bool = True):
        self._is_baja = bool(value)
        target = 1.0 if self._is_baja else 0.0
        if animated and self.isVisible():
            self._anim.stop()
            self._anim.setStartValue(self._thumb_position)
            self._anim.setEndValue(target)
            self._anim.start()
        else:
            self._thumb_position = target
            self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            new_val = not self._is_baja
            self.set_baja(new_val, animated=True)
            self.toggled.emit(self._is_baja)
            event.accept()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        w, h = self.width(), self.height()
        radius = h / 2.0

        # Interpolar color del fondo según _thumb_position
        # 0.0 -> Verde (#10B981)   |   1.0 -> Rojo (#EF4444)
        r = int(16 + (239 - 16) * self._thumb_position)
        g = int(185 + (68 - 185) * self._thumb_position)
        b = int(129 + (68 - 129) * self._thumb_position)
        bg_color = QColor(r, g, b)

        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(bg_color))
        painter.drawRoundedRect(QRectF(0, 0, w, h), radius, radius)

        # Configurar tipografía
        font = QFont(self.font())
        font.setPointSize(10)
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QColor(255, 255, 255))

        thumb_padding = 3.0
        thumb_diameter = h - (2 * thumb_padding)
        min_x = thumb_padding
        max_x = w - thumb_diameter - thumb_padding
        thumb_x = max_x - (max_x - min_x) * self._thumb_position

        # Texto centrado verticalmente
        if self._thumb_position < 0.5:
            text_rect = QRectF(thumb_padding + 8, 0, max_x - thumb_padding, h)
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, "Disponible")
        else:
            text_rect = QRectF(min_x + thumb_diameter + 6, 0, w - (min_x + thumb_diameter + 10), h)
            painter.drawText(text_rect, Qt.AlignVCenter | Qt.AlignLeft, "De baja")

        # Bola deslizante (thumb)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QBrush(QColor(255, 255, 255)))
        painter.drawEllipse(QRectF(thumb_x, thumb_padding, thumb_diameter, thumb_diameter))


class PositionChipButton(QPushButton):
    """Chip interactivo de 3 estados para un puesto:
    - Estado 0 ('') : No sabe  (Fondo gris neutro)
    - Estado 1 ('X'): Sabe     (Fondo verde suave con borde verde)
    - Estado 2 ('V'): Vetado   (Fondo rojo suave con borde rojo)
    """
    state_changed = Signal(str, str)  # (position, new_state)

    def __init__(self, position: str, parent=None):
        super().__init__(parent)
        self.position = position
        self._state = ""
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(34)
        self.clicked.connect(self._on_clicked)
        self._update_appearance()

    def get_state(self) -> str:
        return self._state

    def set_state(self, state: str):
        self._state = state if state in ("X", "V") else ""
        self._update_appearance()

    def _on_clicked(self):
        # Ciclar: '' -> 'X' -> 'V' -> ''
        if self._state == "":
            self._state = "X"
        elif self._state == "X":
            self._state = "V"
        else:
            self._state = ""
        self._update_appearance()
        self.state_changed.emit(self.position, self._state)

    def _update_appearance(self):
        if self._state == "X":
            self.setText(f"✓ {self.position}")
            self.setStyleSheet("""
                QPushButton {
                    background-color: #ECFDF5;
                    color: #065F46;
                    border: 1.5px solid #10B981;
                    border-radius: 6px;
                    font-size: 12px;
                    font-weight: 600;
                    padding: 4px 8px;
                }
                QPushButton:hover { background-color: #D1FAE5; }
            """)
        elif self._state == "V":
            self.setText(f"✕ {self.position} (Veto)")
            self.setStyleSheet("""
                QPushButton {
                    background-color: #FEF2F2;
                    color: #991B1B;
                    border: 1.5px solid #EF4444;
                    border-radius: 6px;
                    font-size: 12px;
                    font-weight: 600;
                    padding: 4px 8px;
                }
                QPushButton:hover { background-color: #FEE2E2; }
            """)
        else:
            self.setText(self.position)
            self.setStyleSheet("""
                QPushButton {
                    background-color: #F8FAFC;
                    color: #64748B;
                    border: 1px solid #CBD5E1;
                    border-radius: 6px;
                    font-size: 12px;
                    font-weight: 500;
                    padding: 4px 8px;
                }
                QPushButton:hover {
                    background-color: #F1F5F9;
                    color: #0F172A;
                    border-color: #94A3B8;
                }
            """)


class DraggablePreferenceList(QListWidget):
    """Lista arrastrable (Drag & Drop) para ordenar las preferencias de puestos."""
    order_changed = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setStyleSheet("""
            QListWidget {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 4px;
            }
            QListWidget::item {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 8px 12px;
                margin-bottom: 4px;
                font-size: 13px;
                font-weight: 500;
                color: #0F172A;
            }
            QListWidget::item:hover {
                background-color: #F1F5F9;
                border-color: #CBD5E1;
            }
            QListWidget::item:selected {
                background-color: #EFF6FF;
                border: 1.5px solid #2563EB;
                color: #1E40AF;
            }
        """)

    def dropEvent(self, event):
        super().dropEvent(event)
        self._refresh_rank_labels()
        self.order_changed.emit(self.get_current_order())

    def set_preferences(self, positions_in_order: list[str]):
        self.clear()
        for idx, pos in enumerate(positions_in_order, start=1):
            item = QListWidgetItem(f"{idx}º — {pos}")
            item.setData(Qt.UserRole, pos)
            self.addItem(item)

    def get_current_order(self) -> list[str]:
        return [self.item(i).data(Qt.UserRole) for i in range(self.count())]

    def _refresh_rank_labels(self):
        for i in range(self.count()):
            item = self.item(i)
            pos = item.data(Qt.UserRole)
            item.setText(f"{i + 1}º — {pos}")


class AddWorkerDialog(QDialog):
    """Diálogo modal para dar de alta a un nuevo profesional."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Añadir Profesional")
        self.setFixedSize(400, 230)
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(22, 20, 22, 20)

        title = QLabel("Alta de Profesional")
        title.setStyleSheet("font-size: 15px; font-weight: 700; color: #0F172A;")
        layout.addWidget(title)

        name_lbl = QLabel("Apellidos y Nombre:")
        name_lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #475569;")
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Ej: GARCIA LOPEZ, MARIA")
        layout.addWidget(name_lbl)
        layout.addWidget(self.name_input)

        status_lbl = QLabel("Estado inicial:")
        status_lbl.setStyleSheet("font-size: 12px; font-weight: 600; color: #475569;")
        self.status_combo = QComboBox()
        self.status_combo.addItem("Disponible (en activo)", False)
        self.status_combo.addItem("De baja temporal", True)
        layout.addWidget(status_lbl)
        layout.addWidget(self.status_combo)

        btn_box = QHBoxLayout()
        btn_box.setSpacing(10)
        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)
        self.accept_btn = QPushButton("Añadir Profesional")
        self.accept_btn.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: white;
                font-weight: 600;
                padding: 7px 16px;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover { background-color: #0369A1; }
        """)
        self.accept_btn.clicked.connect(self._validate_and_accept)
        btn_box.addStretch()
        btn_box.addWidget(self.cancel_btn)
        btn_box.addWidget(self.accept_btn)
        layout.addLayout(btn_box)

    def _validate_and_accept(self):
        txt = self.name_input.text().strip()
        if not txt:
            QMessageBox.warning(self, "Nombre requerido", "Debe indicar el nombre del profesional.")
            return
        self.accept()

    def get_data(self) -> tuple[str, bool]:
        return " ".join(self.name_input.text().strip().split()).upper(), bool(self.status_combo.currentData())


class PuestosDotacionesDialog(QDialog):
    """Diálogo modal para consultar y editar el catálogo de puestos, objetivos (M/T), mínimos y prioridad."""
    def __init__(self, positions: tuple[str, ...], rules: dict[str, dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Catálogo de Puestos y Dotaciones")
        self.resize(720, 500)
        self.positions = list(positions)
        self.rules = {k: dict(v) for k, v in rules.items()}
        self.renamed_map: dict[str, str] = {}

        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(22, 20, 22, 20)

        hdr_box = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("Catálogo de Puestos y Dotaciones del Servicio")
        title.setStyleSheet("font-size: 16px; font-weight: 700; color: #0F172A;")
        sub = QLabel("Configure las dotaciones de mañana y tarde, mínimos requeridos y criterios de reducción ante déficit.")
        sub.setStyleSheet("font-size: 12px; color: #64748B;")
        title_box.addWidget(title)
        title_box.addWidget(sub)
        hdr_box.addLayout(title_box, 1)

        self.add_pos_btn = QPushButton("+ Añadir Puesto")
        self.add_pos_btn.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: white;
                font-weight: 600;
                padding: 6px 12px;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover { background-color: #0369A1; }
        """)
        self.add_pos_btn.clicked.connect(self._on_add_position)

        self.del_pos_btn = QPushButton("🗑️ Eliminar Puesto")
        self.del_pos_btn.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                color: #DC2626;
                border: 1px solid #FECACA;
                font-weight: 600;
                padding: 6px 12px;
                border-radius: 6px;
            }
            QPushButton:hover { background-color: #FEF2F2; }
        """)
        self.del_pos_btn.clicked.connect(self._on_del_position)

        hdr_box.addWidget(self.add_pos_btn)
        hdr_box.addWidget(self.del_pos_btn)
        layout.addLayout(hdr_box)

        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([
            "Código", "Obj. M", "Obj. T", "Mín. M", "Mín. T", "Criterio de Reducción"
        ])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        for i in range(1, 5):
            self.table.horizontalHeader().setSectionResizeMode(i, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        layout.addWidget(self.table, 1)

        self._populate_table()

        btn_box = QHBoxLayout()
        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)
        self.save_btn = QPushButton("Aplicar Configuración")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #0284C7;
                color: white;
                font-weight: 600;
                padding: 8px 20px;
                border-radius: 6px;
                border: none;
                font-size: 13px;
            }
            QPushButton:hover { background-color: #0369A1; }
        """)
        self.save_btn.clicked.connect(self._on_save)
        btn_box.addStretch()
        btn_box.addWidget(self.cancel_btn)
        btn_box.addWidget(self.save_btn)
        layout.addLayout(btn_box)

    def _populate_table(self):
        self.table.setRowCount(len(self.positions))
        for row, pos in enumerate(self.positions):
            r_info = self.rules.get(pos, {})
            tm = r_info.get("target_m", 1)
            tt = r_info.get("target_t", 1)
            mm = r_info.get("min_m", tm)
            mt = r_info.get("min_t", tt)
            prio = r_info.get("priority", 0)

            # Col 0: Código (guarda orig_code para migrar competencias si se renombra)
            c_code = QLineEdit(pos)
            c_code.setMaxLength(8)
            c_code.setProperty("orig_code", pos)
            self.table.setCellWidget(row, 0, c_code)

            # Col 1: Obj M
            s_tm = QSpinBox()
            s_tm.setRange(0, 30)
            s_tm.setValue(tm)
            self.table.setCellWidget(row, 1, s_tm)

            # Col 2: Obj T
            s_tt = QSpinBox()
            s_tt.setRange(0, 30)
            s_tt.setValue(tt)
            self.table.setCellWidget(row, 2, s_tt)

            # Col 3: Min M
            s_mm = QSpinBox()
            s_mm.setRange(0, 30)
            s_mm.setValue(mm)
            self.table.setCellWidget(row, 3, s_mm)

            # Col 4: Min T
            s_mt = QSpinBox()
            s_mt.setRange(0, 30)
            s_mt.setValue(mt)
            self.table.setCellWidget(row, 4, s_mt)

            # Col 5: Prioridad
            cb_prio = QComboBox()
            cb_prio.addItem("Estricto (no se reduce)", 0)
            cb_prio.addItem("Reducible Nivel 1 (se reduce primero)", 1)
            cb_prio.addItem("Reducible Nivel 2 (se reduce después de nivel 1)", 2)
            cb_prio.setCurrentIndex(min(2, max(0, prio)))
            self.table.setCellWidget(row, 5, cb_prio)

    def _on_add_position(self):
        row = self.table.rowCount()
        self.table.insertRow(row)

        c_code = QLineEdit(f"P{row + 1}")
        c_code.setMaxLength(8)
        self.table.setCellWidget(row, 0, c_code)

        s_tm = QSpinBox()
        s_tm.setRange(0, 30)
        s_tm.setValue(1)
        self.table.setCellWidget(row, 1, s_tm)

        s_tt = QSpinBox()
        s_tt.setRange(0, 30)
        s_tt.setValue(1)
        self.table.setCellWidget(row, 2, s_tt)

        s_mm = QSpinBox()
        s_mm.setRange(0, 30)
        s_mm.setValue(1)
        self.table.setCellWidget(row, 3, s_mm)

        s_mt = QSpinBox()
        s_mt.setRange(0, 30)
        s_mt.setValue(1)
        self.table.setCellWidget(row, 4, s_mt)

        cb_prio = QComboBox()
        cb_prio.addItem("Estricto (no se reduce)", 0)
        cb_prio.addItem("Reducible Nivel 1 (se reduce primero)", 1)
        cb_prio.addItem("Reducible Nivel 2 (se reduce después de nivel 1)", 2)
        cb_prio.setCurrentIndex(0)
        self.table.setCellWidget(row, 5, cb_prio)

        self.table.selectRow(row)

    def _on_del_position(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Seleccionar", "Seleccione el puesto que desea eliminar.")
            return
        if self.table.rowCount() <= 1:
            QMessageBox.warning(self, "Aviso", "El catálogo debe contener al menos 1 puesto de trabajo.")
            return
        code_w = self.table.cellWidget(row, 0)
        code = code_w.text().strip().upper() if code_w else f"Fila {row + 1}"
        reply = QMessageBox.question(
            self,
            "Eliminar Puesto",
            f"¿Desea eliminar el puesto '{code}' del catálogo?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.table.removeRow(row)

    def _on_save(self):
        new_positions = []
        new_rules = {}
        renamed_map: dict[str, str] = {}

        for r in range(self.table.rowCount()):
            code_w = self.table.cellWidget(r, 0)
            code = code_w.text().strip().upper() if code_w else ""
            orig_code = str(code_w.property("orig_code") or "").strip().upper() if code_w else ""
            if not code:
                QMessageBox.warning(self, "Código vacío", f"La fila {r + 1} no tiene código de puesto.")
                return
            if code in new_positions:
                QMessageBox.warning(self, "Código duplicado", f"El código '{code}' aparece más de una vez.")
                return
            new_positions.append(code)
            if orig_code and orig_code != code:
                renamed_map[orig_code] = code

            tm = self.table.cellWidget(r, 1).value()
            tt = self.table.cellWidget(r, 2).value()
            mm = self.table.cellWidget(r, 3).value()
            mt = self.table.cellWidget(r, 4).value()
            prio = self.table.cellWidget(r, 5).currentData()

            # Mínimo no puede ser mayor que el objetivo
            if mm > tm:
                mm = tm
            if mt > tt:
                mt = tt

            prev_name = (
                self.rules.get(code, {}).get("name")
                or (self.rules.get(orig_code, {}).get("name") if orig_code else None)
                or code
            )
            new_rules[code] = {
                "name": prev_name,
                "target_m": tm,
                "target_t": tt,
                "min_m": mm,
                "min_t": mt,
                "priority": int(prio),
            }

        self.positions = new_positions
        self.rules = new_rules
        self.renamed_map = renamed_map
        self.accept()

    def get_data(self) -> tuple[tuple[str, ...], dict[str, dict]]:
        return tuple(self.positions), self.rules

    def get_renamed_map(self) -> dict[str, str]:
        return dict(self.renamed_map)


class TrabajadoresTab(QWidget):
    """Pestaña integral de gestión de trabajadores con vista Maestro-Detalle."""
    data_saved = Signal()
    dirty_changed = Signal(bool)

    def __init__(self, pref_path: Path, parent=None, seq_path: Path | None = None):
        super().__init__(parent)
        self.pref_path = Path(pref_path)
        self.seq_path = Path(seq_path) if seq_path is not None else None
        self.workers: list[EditableWorker] = []
        self.positions: tuple[str, ...] = POSITIONS
        self.rules: dict[str, dict] = dict(POSITION_RULES)
        self.current_worker: EditableWorker | None = None
        self.chip_widgets: dict[str, PositionChipButton] = {}
        self.is_dirty: bool = False
        self._removed_worker_names: set[str] = set()

        self._setup_ui()
        self.load_data()

    def _set_dirty(self, dirty: bool):
        dirty = bool(dirty)
        if self.is_dirty != dirty:
            self.is_dirty = dirty
            if hasattr(self, "save_btn"):
                self.save_btn.setText("Guardar Cambios en Excel ●" if dirty else "Guardar Cambios en Excel")
            self.dirty_changed.emit(dirty)

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(14)

        # 1. Barra Superior: Buscador + Contador + Botones de Acción
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar por nombre de trabajador...")
        self.search_input.setObjectName("searchInput")
        self.search_input.textChanged.connect(self._filter_workers)
        top_bar.addWidget(self.search_input, 1)

        self.summary_badge = QLabel("0 trabajadores")
        self.summary_badge.setObjectName("summaryBadge")
        top_bar.addWidget(self.summary_badge)

        self.reload_btn = QPushButton("🔄 Recargar")
        self.reload_btn.setObjectName("secondaryBtn")
        self.reload_btn.setToolTip("Recargar datos desde el archivo Excel para detectar cambios externos.")
        self.reload_btn.clicked.connect(self._on_reload_clicked)
        top_bar.addWidget(self.reload_btn)

        self.positions_btn = QPushButton("⚙️ Puestos y Dotaciones")
        self.positions_btn.setObjectName("secondaryBtn")
        self.positions_btn.setToolTip("Consultar y editar el catálogo de puestos, objetivos (M/T), mínimos y reglas.")
        self.positions_btn.clicked.connect(self._open_positions_dialog)
        top_bar.addWidget(self.positions_btn)

        self.save_btn = QPushButton("Guardar Cambios en Excel")
        self.save_btn.setObjectName("saveExcelButton")
        self.save_btn.clicked.connect(self.save_data)
        top_bar.addWidget(self.save_btn)

        main_layout.addLayout(top_bar)

        # Banner de notificación de guardado
        self.notification_lbl = QLabel("")
        self.notification_lbl.setObjectName("notificationLabel")
        self.notification_lbl.setVisible(False)
        main_layout.addWidget(self.notification_lbl)

        self._notification_timer = QTimer(self)
        self._notification_timer.setSingleShot(True)
        self._notification_timer.timeout.connect(lambda: self.notification_lbl.setVisible(False))

        # 2. Panel Maestro-Detalle (Splitter)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(1)

        # Panel Izquierdo: Lista de Trabajadores
        left_panel = QFrame()
        left_panel.setObjectName("leftPanel")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 12, 0)
        left_layout.setSpacing(8)

        list_title = QLabel("Directorio de Personal")
        list_title.setObjectName("sectionTitle")
        left_layout.addWidget(list_title)

        self.workers_list = QListWidget()
        self.workers_list.setObjectName("workersMasterList")
        self.workers_list.currentRowChanged.connect(self._on_worker_selected)
        left_layout.addWidget(self.workers_list, 1)

        self.add_worker_btn = QPushButton("+ Añadir Profesional")
        self.add_worker_btn.setObjectName("addWorkerBtn")
        self.add_worker_btn.clicked.connect(self._on_add_worker)
        left_layout.addWidget(self.add_worker_btn)

        splitter.addWidget(left_panel)

        # Panel Derecho: Detalle y Edición del Trabajador
        self.right_panel = QFrame()
        self.right_panel.setObjectName("rightPanel")
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(16, 0, 0, 0)
        right_layout.setSpacing(16)

        # Cabecera de la Ficha: Nombre, Botón Eliminar y Deslizador de Baja
        detail_header = QHBoxLayout()
        detail_header.setSpacing(12)
        name_box = QVBoxLayout()
        name_box.setSpacing(2)

        self.worker_name_lbl = QLabel("Seleccione un trabajador")
        self.worker_name_lbl.setObjectName("workerDetailTitle")
        self.worker_sub_lbl = QLabel("Ficha de competencias y preferencias")
        self.worker_sub_lbl.setObjectName("workerDetailSubtitle")
        name_box.addWidget(self.worker_name_lbl)
        name_box.addWidget(self.worker_sub_lbl)
        detail_header.addLayout(name_box, 1)

        self.del_worker_btn = QPushButton("🗑️ Eliminar")
        self.del_worker_btn.setObjectName("deleteBtn")
        self.del_worker_btn.setToolTip("Eliminar a este profesional de la lista de trabajadores.")
        self.del_worker_btn.clicked.connect(self._on_delete_worker)
        detail_header.addWidget(self.del_worker_btn)

        # Deslizador de Estado (Verde Disponible / Rojo De baja)
        switch_box = QVBoxLayout()
        switch_box.setSpacing(4)
        switch_lbl = QLabel("Estado Laboral")
        switch_lbl.setObjectName("fieldLabel")
        self.status_switch = StatusSwitch()
        self.status_switch.toggled.connect(self._on_status_toggled)
        switch_box.addWidget(switch_lbl)
        switch_box.addWidget(self.status_switch)
        detail_header.addLayout(switch_box)

        right_layout.addLayout(detail_header)

        sep1 = QFrame()
        sep1.setFrameShape(QFrame.HLine)
        sep1.setObjectName("itemSeparator")
        right_layout.addWidget(sep1)

        # Sección 1: Competencias y Vetos (Chips de 3 estados)
        skills_sec = QVBoxLayout()
        skills_sec.setSpacing(6)
        sec1_title = QLabel("Competencias y Vetos")
        sec1_title.setObjectName("cardTitle")
        sec1_hint = QLabel("Haga clic en un puesto para alternar: No sabe (gris) ➔ Sabe (verde) ➔ Vetado (rojo)")
        sec1_hint.setObjectName("hintLabel")
        skills_sec.addWidget(sec1_title)
        skills_sec.addWidget(sec1_hint)

        self.grid_widget = QWidget()
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setSpacing(8)
        self.grid_layout.setContentsMargins(0, 6, 0, 6)

        skills_sec.addWidget(self.grid_widget)
        right_layout.addLayout(skills_sec)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        sep2.setObjectName("itemSeparator")
        right_layout.addWidget(sep2)

        # Sección 2: Preferencias de Puestos (Drag & Drop)
        prefs_sec = QVBoxLayout()
        prefs_sec.setSpacing(6)
        sec2_title = QLabel("Orden de Preferencia de Puestos")
        sec2_title.setObjectName("cardTitle")
        sec2_hint = QLabel("Arrastre los puestos sabidos para colocarlos en el orden de mayor a menor preferencia.")
        sec2_hint.setObjectName("hintLabel")
        prefs_sec.addWidget(sec2_title)
        prefs_sec.addWidget(sec2_hint)

        self.pref_list = DraggablePreferenceList()
        self.pref_list.order_changed.connect(self._on_preference_order_changed)
        prefs_sec.addWidget(self.pref_list, 1)

        right_layout.addLayout(prefs_sec, 1)

        splitter.addWidget(self.right_panel)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 6)

        main_layout.addWidget(splitter, 1)
        self._apply_styles()

    def rebuild_position_chips(self, positions: tuple[str, ...]):
        """Reconstruye dinámicamente los chips de competencias según el catálogo de puestos."""
        while self.grid_layout.count():
            item = self.grid_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self.chip_widgets.clear()

        n_cols = 4 if len(positions) <= 8 else (5 if len(positions) >= 15 else 4)
        for idx, pos in enumerate(positions):
            row = idx // n_cols
            col = idx % n_cols
            chip = PositionChipButton(pos)
            chip.state_changed.connect(self._on_chip_state_changed)
            self.chip_widgets[pos] = chip
            self.grid_layout.addWidget(chip, row, col)

        if self.current_worker:
            for pos, chip in self.chip_widgets.items():
                chip.set_state(self.current_worker.skills.get(pos, ""))

    def _show_notification(self, text: str):
        self.notification_lbl.setText(text)
        self.notification_lbl.setVisible(True)
        self._notification_timer.start(4000)

    def _on_reload_clicked(self):
        if self.is_dirty:
            reply = QMessageBox.question(
                self,
                "Cambios sin guardar",
                "Hay modificaciones sin guardar en la pestaña Trabajadores.\n"
                "¿Desea descartarlas y recargar desde el archivo Excel?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return
        self.load_data()
        self._show_notification("✔ Datos sincronizados con el archivo Excel.")

    def _on_add_worker(self):
        dlg = AddWorkerDialog(parent=self)
        if dlg.exec() == QDialog.Accepted:
            name, is_baja = dlg.get_data()
            if any(w.name.upper() == name for w in self.workers):
                QMessageBox.warning(self, "Nombre duplicado", f"Ya existe un profesional con el nombre '{name}'.")
                return
            self._removed_worker_names.discard(name)
            new_worker = EditableWorker(
                name=name,
                row_index=len(self.workers) + 4,
                is_baja=is_baja,
                skills={p: "" for p in self.positions},
                preference_order=[],
            )
            self.workers.append(new_worker)
            if self.search_input.text().strip() and self.search_input.text().strip().upper() not in name.upper():
                self.search_input.blockSignals(True)
                self.search_input.clear()
                self.search_input.blockSignals(False)
            self._render_master_list()
            # Seleccionar nuevo trabajador
            for r in range(self.workers_list.count()):
                w = self.workers_list.item(r).data(Qt.UserRole)
                if w.name == name:
                    self.workers_list.setCurrentRow(r)
                    break
            self._update_summary()
            self._set_dirty(True)
            self._show_notification(f"✔ Profesional '{name}' añadido. Asigne sus competencias y guarde.")

    def _on_delete_worker(self):
        if not self.current_worker:
            return
        w = self.current_worker
        reply = QMessageBox.question(
            self,
            "Eliminar Profesional",
            f"¿Está seguro de que desea eliminar a '{w.name}' del cuadrante?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            name = w.name
            self._removed_worker_names.add(name)
            self.workers.remove(w)
            self.current_worker = None
            self._render_master_list()
            if self.workers_list.count() > 0:
                self.workers_list.setCurrentRow(0)
            else:
                self.worker_name_lbl.setText("Seleccione un profesional")
                self.worker_sub_lbl.setText("Ficha de competencias y preferencias")
                for chip in self.chip_widgets.values():
                    chip.set_state("")
                self.pref_list.clear()
            self._update_summary()
            self._set_dirty(True)
            self._show_notification(f"✔ Profesional '{name}' eliminado de la lista. Recuerde guardar en Excel.")

    def _open_positions_dialog(self):
        dlg = PuestosDotacionesDialog(self.positions, self.rules, parent=self)
        if dlg.exec() == QDialog.Accepted:
            new_positions, new_rules = dlg.get_data()
            renamed_map = dlg.get_renamed_map()
            self.positions = new_positions
            self.rules = new_rules

            # Actualizar competencias de trabajadores actuales para el nuevo catálogo (preservando renombrados N3)
            for w in self.workers:
                for old_code, new_code in renamed_map.items():
                    if old_code in w.skills and new_code not in w.skills:
                        w.skills[new_code] = w.skills.pop(old_code)
                    w.preference_order = [
                        (new_code if p == old_code else p) for p in w.preference_order
                    ]
                for p in new_positions:
                    if p not in w.skills:
                        w.skills[p] = ""
                w.skills = {p: st for p, st in w.skills.items() if p in new_positions}
                seen_prefs: list[str] = []
                for p in w.preference_order:
                    if p in new_positions and w.skills.get(p) == "X" and p not in seen_prefs:
                        seen_prefs.append(p)
                w.preference_order = seen_prefs

            self.rebuild_position_chips(self.positions)
            if self.current_worker:
                for pos, chip in self.chip_widgets.items():
                    chip.set_state(self.current_worker.skills.get(pos, ""))
                self.pref_list.set_preferences(self.current_worker.preference_order)
            self._render_master_list()
            self._set_dirty(True)
            self._show_notification("✔ Puestos actualizados. Haga clic en 'Guardar Cambios en Excel' para persistir.")

    def _apply_styles(self):
        self.setStyleSheet("""
            #searchInput {
                background-color: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 13px;
                color: #0F172A;
            }
            #searchInput:focus { border: 1.5px solid #0284C7; }
            #summaryBadge {
                background-color: #F0F9FF;
                color: #0369A1;
                border: 1px solid #BAE6FD;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 12px;
                font-weight: 600;
            }
            #secondaryBtn {
                background-color: #FFFFFF;
                color: #334155;
                border: 1px solid #CBD5E1;
                border-radius: 6px;
                padding: 8px 14px;
                font-size: 13px;
                font-weight: 500;
            }
            #secondaryBtn:hover {
                background-color: #F8FAFC;
                border-color: #94A3B8;
                color: #0F172A;
            }
            #addWorkerBtn {
                background-color: #F0F9FF;
                color: #0284C7;
                border: 1.5px dashed #BAE6FD;
                border-radius: 6px;
                padding: 8px 12px;
                font-size: 13px;
                font-weight: 600;
            }
            #addWorkerBtn:hover {
                background-color: #E0F2FE;
                border-color: #0284C7;
            }
            #deleteBtn {
                background-color: #FFFFFF;
                color: #DC2626;
                border: 1px solid #FECACA;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 12px;
                font-weight: 600;
            }
            #deleteBtn:hover {
                background-color: #FEF2F2;
                border-color: #EF4444;
            }
            #saveExcelButton {
                background-color: #0284C7;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 600;
            }
            #saveExcelButton:hover { background-color: #0369A1; }
            #saveExcelButton:pressed { background-color: #075985; }
            #sectionTitle {
                font-size: 13px;
                font-weight: 600;
                color: #475569;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }
            #workersMasterList {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 6px;
            }
            #workersMasterList::item {
                border-radius: 6px;
                padding: 8px 10px;
                margin-bottom: 2px;
            }
            #workersMasterList::item:hover { background-color: #F8FAFC; }
            #workersMasterList::item:selected {
                background-color: #EFF6FF;
                color: #1E40AF;
                font-weight: 600;
            }
            #workerDetailTitle {
                font-size: 18px;
                font-weight: 700;
                color: #0F172A;
            }
            #workerDetailSubtitle {
                font-size: 12px;
                color: #64748B;
            }
            #cardTitle {
                font-size: 14px;
                font-weight: 600;
                color: #0F172A;
            }
            #hintLabel {
                font-size: 12px;
                color: #64748B;
            }
            #fieldLabel {
                font-size: 11px;
                font-weight: 600;
                color: #64748B;
            }
            #itemSeparator {
                background-color: #F1F5F9;
                height: 1px;
                border: none;
            }
            #notificationLabel {
                background-color: #ECFDF5;
                color: #065F46;
                border: 1px solid #A7F3D0;
                border-radius: 6px;
                padding: 8px 14px;
                font-size: 12px;
                font-weight: 600;
            }
        """)

    def load_data(self):
        """Carga o recarga los trabajadores desde el Excel de configuración."""
        if not self.pref_path.exists():
            return
        try:
            self.workers = load_editable_workers(self.pref_path)
            self.positions = getattr(self.workers, "positions", POSITIONS)
            self.rules = getattr(self.workers, "rules", dict(POSITION_RULES))
            self._removed_worker_names.clear()
            self.rebuild_position_chips(self.positions)
            self._render_master_list()
            if self.workers:
                self.workers_list.setCurrentRow(0)
            self._update_summary()
            self._set_dirty(False)
        except Exception as e:
            QMessageBox.warning(self, "Error al cargar plantilla", f"No se pudo leer la plantilla: {str(e)}")

    def _render_master_list(self):
        filter_text = self.search_input.text().strip().upper()
        self.workers_list.clear()

        for w in self.workers:
            if filter_text and filter_text not in w.name.upper():
                continue

            status_txt = "● De baja" if w.is_baja else f"{w.count_known_skills()} puestos"
            display_text = f"{w.name}\n   {status_txt}"

            item = QListWidgetItem(display_text)
            item.setData(Qt.UserRole, w)

            # Color suave según estado
            if w.is_baja:
                item.setForeground(QColor("#DC2626"))
            else:
                item.setForeground(QColor("#0F172A"))

            self.workers_list.addItem(item)

    def _filter_workers(self, text: str):
        self._render_master_list()
        if self.workers_list.count() > 0:
            self.workers_list.setCurrentRow(0)
        else:
            self._on_worker_selected(-1)

    def _on_worker_selected(self, row: int):
        if row < 0 or row >= self.workers_list.count():
            self.current_worker = None
            self.right_panel.setEnabled(False)
            self.worker_name_lbl.setText("Sin selección")
            self.worker_sub_lbl.setText("Seleccione un profesional de la lista")
            for chip in self.chip_widgets.values():
                chip.set_state("")
            self.pref_list.clear()
            return
        self.right_panel.setEnabled(True)
        item = self.workers_list.item(row)
        worker: EditableWorker = item.data(Qt.UserRole)
        self.current_worker = worker

        self.worker_name_lbl.setText(worker.name)
        self.worker_sub_lbl.setText(f"Fila {worker.row_index} en plantilla Excel")

        # Configurar switch de baja
        self.status_switch.set_baja(worker.is_baja, animated=False)

        # Configurar chips de los 16 puestos
        for pos, chip in self.chip_widgets.items():
            chip.set_state(worker.skills.get(pos, ""))

        # Configurar lista de preferencias
        self.pref_list.set_preferences(worker.preference_order)

    def _on_status_toggled(self, is_baja: bool):
        if not self.current_worker:
            return
        self.current_worker.is_baja = is_baja
        self._refresh_current_list_item()
        self._update_summary()
        self._set_dirty(True)

    def _on_chip_state_changed(self, pos: str, new_state: str):
        if not self.current_worker:
            return
        self.current_worker.set_skill(pos, new_state)
        # Refrescar lista de preferencias con los puestos sabidos actuales
        self.pref_list.set_preferences(self.current_worker.preference_order)
        self._refresh_current_list_item()
        self._set_dirty(True)

    def _on_preference_order_changed(self, new_order: list[str]):
        if not self.current_worker:
            return
        self.current_worker.preference_order = new_order
        self._set_dirty(True)

    def _refresh_current_list_item(self):
        row = self.workers_list.currentRow()
        if row < 0 or not self.current_worker:
            return
        w = self.current_worker
        status_txt = "● De baja" if w.is_baja else f"{w.count_known_skills()} puestos"
        display_text = f"{w.name}\n   {status_txt}"

        item = self.workers_list.item(row)
        item.setText(display_text)
        if w.is_baja:
            item.setForeground(QColor("#DC2626"))
        else:
            item.setForeground(QColor("#0F172A"))

    def _update_summary(self):
        total = len(self.workers)
        bajas = sum(1 for w in self.workers if w.is_baja)
        self.summary_badge.setText(f"{total} profesionales  ·  {bajas} de baja")

    def save_data(self) -> bool:
        """Guarda las modificaciones en Plantilla_Conocimientos_y_Preferencias.xlsx y sincroniza secuencias."""
        if not self.pref_path.exists():
            QMessageBox.critical(self, "Error", f"No se encuentra la plantilla en {self.pref_path}")
            return False
        try:
            save_editable_workers(self.pref_path, self.workers, positions=self.positions, rules=self.rules)
            if self.seq_path is not None and self.seq_path.exists():
                sync_sequences_roster(
                    self.seq_path,
                    [w.name for w in self.workers],
                    removed_worker_names=self._removed_worker_names,
                )
            self._removed_worker_names.clear()
            self._set_dirty(False)
            self._show_notification("✔ Cambios guardados correctamente en Plantilla_Conocimientos_y_Preferencias.xlsx")
            self.data_saved.emit()
            return True
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", f"No se pudo guardar el archivo: {str(e)}")
            return False
