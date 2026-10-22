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
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from reparto.datos import POSITIONS
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


class TrabajadoresTab(QWidget):
    """Pestaña integral de gestión de trabajadores con vista Maestro-Detalle."""
    data_saved = Signal()

    def __init__(self, pref_path: Path, parent=None):
        super().__init__(parent)
        self.pref_path = Path(pref_path)
        self.workers: list[EditableWorker] = []
        self.current_worker: EditableWorker | None = None
        self.chip_widgets: dict[str, PositionChipButton] = {}

        self._setup_ui()
        self.load_data()

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(14)

        # 1. Barra Superior: Buscador + Contador + Botón Guardar
        top_bar = QHBoxLayout()
        top_bar.setSpacing(14)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar por nombre de trabajador...")
        self.search_input.setObjectName("searchInput")
        self.search_input.textChanged.connect(self._filter_workers)
        top_bar.addWidget(self.search_input, 1)

        self.summary_badge = QLabel("0 trabajadores")
        self.summary_badge.setObjectName("summaryBadge")
        top_bar.addWidget(self.summary_badge)

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
        left_layout.addWidget(self.workers_list)

        splitter.addWidget(left_panel)

        # Panel Derecho: Detalle y Edición del Trabajador
        self.right_panel = QFrame()
        self.right_panel.setObjectName("rightPanel")
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(16, 0, 0, 0)
        right_layout.setSpacing(16)

        # Cabecera de la Ficha: Nombre y Deslizador de Baja
        detail_header = QHBoxLayout()
        name_box = QVBoxLayout()
        name_box.setSpacing(2)

        self.worker_name_lbl = QLabel("Seleccione un trabajador")
        self.worker_name_lbl.setObjectName("workerDetailTitle")
        self.worker_sub_lbl = QLabel("Ficha de competencias y preferencias")
        self.worker_sub_lbl.setObjectName("workerDetailSubtitle")
        name_box.addWidget(self.worker_name_lbl)
        name_box.addWidget(self.worker_sub_lbl)
        detail_header.addLayout(name_box, 1)

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

        # Grid 4x4 de chips
        grid_widget = QWidget()
        grid_layout = QGridLayout(grid_widget)
        grid_layout.setSpacing(8)
        grid_layout.setContentsMargins(0, 6, 0, 6)

        for idx, pos in enumerate(POSITIONS):
            row = idx // 4
            col = idx % 4
            chip = PositionChipButton(pos)
            chip.state_changed.connect(self._on_chip_state_changed)
            self.chip_widgets[pos] = chip
            grid_layout.addWidget(chip, row, col)

        skills_sec.addWidget(grid_widget)
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
            #searchInput:focus { border: 1.5px solid #2563EB; }
            #summaryBadge {
                background-color: #F1F5F9;
                color: #475569;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 6px 12px;
                font-size: 12px;
                font-weight: 600;
            }
            #saveExcelButton {
                background-color: #0F172A;
                color: #FFFFFF;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: 600;
            }
            #saveExcelButton:hover { background-color: #1E293B; }
            #saveExcelButton:pressed { background-color: #020617; }
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
            self._render_master_list()
            if self.workers:
                self.workers_list.setCurrentRow(0)
            self._update_summary()
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

    def _on_worker_selected(self, row: int):
        if row < 0 or row >= self.workers_list.count():
            return
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

    def _on_chip_state_changed(self, pos: str, new_state: str):
        if not self.current_worker:
            return
        self.current_worker.set_skill(pos, new_state)
        # Refrescar lista de preferencias con los puestos sabidos actuales
        self.pref_list.set_preferences(self.current_worker.preference_order)
        self._refresh_current_list_item()

    def _on_preference_order_changed(self, new_order: list[str]):
        if not self.current_worker:
            return
        self.current_worker.preference_order = new_order

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

    def save_data(self):
        """Guarda las modificaciones en Plantilla_Conocimientos_y_Preferencias.xlsx."""
        if not self.pref_path.exists():
            QMessageBox.critical(self, "Error", f"No se encuentra la plantilla en {self.pref_path}")
            return
        try:
            save_editable_workers(self.pref_path, self.workers)
            self.notification_lbl.setText("✔ Cambios guardados correctamente en Plantilla_Conocimientos_y_Preferencias.xlsx")
            self.notification_lbl.setVisible(True)
            self.data_saved.emit()
        except Exception as e:
            QMessageBox.critical(self, "Error al guardar", f"No se pudo guardar el archivo: {str(e)}")
