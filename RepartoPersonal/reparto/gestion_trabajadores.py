"""Gestión y persistencia de trabajadores, competencias, vetos y preferencias.

Permite consultar, modificar y persistir la configuración de trabajadores
directamente sobre el archivo 'Plantilla_Conocimientos_y_Preferencias.xlsx',
garantizando que el archivo de secuencias de turnos NUNCA sea alterado.

Implementado exclusivamente con la biblioteca estándar de Python (zipfile y xml),
sin depender de paquetes externos como openpyxl.
"""

from dataclasses import dataclass, field
from pathlib import Path
import shutil
from xml.etree import ElementTree as ET
import zipfile

from reparto.datos import (
    POSITIONS,
    _parse_xlsx_cells,
    parse_positions_from_header,
    parse_position_rules,
)

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class WorkerList(list):
    """Lista mutable de trabajadores que preserva el catálogo dinámico de puestos y sus reglas."""
    def __init__(self, iterable=(), positions: tuple[str, ...] = (), rules: dict[str, dict] | None = None):
        super().__init__(iterable)
        self.positions: tuple[str, ...] = positions
        self.rules: dict[str, dict] = rules if rules is not None else {}


@dataclass
class EditableWorker:
    """Representación mutable de un trabajador para la interfaz de edición."""
    name: str
    row_index: int
    is_baja: bool = False
    skills: dict[str, str] = field(default_factory=dict)         # pos -> "" (no sabe), "X" (sabe), "V" (veto)
    preference_order: list[str] = field(default_factory=list)     # lista de puestos sabidos ordenados (1º a último)

    def knows_position(self, pos: str) -> bool:
        return self.skills.get(pos) == "X"

    def is_vetoed(self, pos: str) -> bool:
        return self.skills.get(pos) == "V"

    def count_known_skills(self) -> int:
        return sum(1 for s in self.skills.values() if s == "X")

    def count_vetoes(self) -> int:
        return sum(1 for s in self.skills.values() if s == "V")

    def cycle_skill(self, pos: str) -> str:
        """Cicla el estado del puesto: '' -> 'X' -> 'V' -> '' y actualiza el orden de preferencias."""
        current = self.skills.get(pos, "")
        if current == "":
            new_state = "X"
        elif current == "X":
            new_state = "V"
        else:
            new_state = ""
        self.set_skill(pos, new_state)
        return new_state

    def set_skill(self, pos: str, state: str) -> None:
        self.skills[pos] = state
        if state == "X":
            if pos not in self.preference_order:
                self.preference_order.append(pos)
        else:
            if pos in self.preference_order:
                self.preference_order.remove(pos)


def _col_letter(c_idx: int) -> str:
    """Convierte un índice de columna 0-based en letras ('A', 'B', ..., 'AA')."""
    res = ""
    while c_idx >= 0:
        res = chr(c_idx % 26 + ord("A")) + res
        c_idx = c_idx // 26 - 1
    return res


def _col_idx_from_ref(ref: str) -> int:
    """Convierte una referencia tipo 'C4' o 'AA12' en índice 0-based."""
    idx = 0
    for ch in ref:
        if ch.isalpha():
            idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def load_editable_workers(xlsx_path: str | Path) -> WorkerList:
    """Carga los trabajadores desde la hoja 'Trabajadores y Preferencias' del Excel de forma dinámica."""
    path = Path(xlsx_path)
    positions = parse_positions_from_header(path)
    rules = parse_position_rules(path, positions)

    data = path.read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="Trabajadores")

    max_row = max((r for r, c in cells), default=0)
    max_col = max((c for r, c in cells), default=0)

    # Detectar columnas en la cabecera (Fila 3 de Excel -> índice 2)
    header_row = 2
    col_name = 1
    col_baja = next((c for c in range(max_col + 1) if "BAJA" in str(cells.get((header_row, c)) or "").upper()), 2)

    pos_cols = [c for c in range(max_col + 1) if str(cells.get((header_row, c)) or "").strip().upper() in positions]
    k = len(positions)
    if len(pos_cols) >= 2 * k:
        pos_cols_skills = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols[:k]}
        pos_cols_prefs = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols[k:2 * k]}
    else:
        offset = 1 if baja_col == 2 else 0
        pos_cols_skills = {pos: 2 + offset + idx for idx, pos in enumerate(positions)}
        pos_cols_prefs = {pos: 2 + offset + k + idx for idx, pos in enumerate(positions)}

    workers: list[EditableWorker] = []

    for r in range(header_row + 1, max_row + 1):
        name_val = cells.get((r, col_name))
        if not name_val or not str(name_val).strip():
            continue
        name = str(name_val).strip()

        # Comprobar estado de baja
        baja_val = str(cells.get((r, col_baja)) or "").strip().upper()
        is_baja = baja_val in ("X", "S", "SI", "1", "TRUE", "V")

        skills: dict[str, str] = {}
        raw_prefs: dict[str, int] = {}

        for pos in positions:
            # Competencia
            c_skill = pos_cols_skills.get(pos)
            s_val = str(cells.get((r, c_skill)) or "").strip().upper() if c_skill is not None else ""
            if s_val in ("X", "1"):
                skills[pos] = "X"
            elif s_val == "V":
                skills[pos] = "V"
            else:
                skills[pos] = ""

            # Preferencia
            c_pref = pos_cols_prefs.get(pos)
            if c_pref is not None:
                p_val = cells.get((r, c_pref))
                if isinstance(p_val, (int, float)) and not isinstance(p_val, bool):
                    raw_prefs[pos] = int(p_val)

        # Ordenar lista de preferencias para los puestos sabidos ('X')
        known_positions = [pos for pos in positions if skills.get(pos) == "X"]
        known_with_rank = sorted(
            [pos for pos in known_positions if pos in raw_prefs],
            key=lambda p: raw_prefs[p],
        )
        known_without_rank = [pos for pos in known_positions if pos not in raw_prefs]
        preference_order = known_with_rank + known_without_rank

        workers.append(EditableWorker(
            name=name,
            row_index=r + 1,
            is_baja=is_baja,
            skills=skills,
            preference_order=preference_order,
        ))

    return WorkerList(workers, positions=positions, rules=rules)


def save_editable_workers(
    xlsx_path: str | Path,
    workers: list[EditableWorker],
    positions: tuple[str, ...] | None = None,
    rules: dict[str, dict] | None = None,
) -> None:
    """Guarda las modificaciones en 'Plantilla_Conocimientos_y_Preferencias.xlsx'.

    Reconstruye de forma limpia las filas de trabajadores y cabeceras dinámicas
    usando la biblioteca estándar de Python sin requerir openpyxl.
    """
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de plantilla en: {path}")

    if positions is None:
        positions = getattr(workers, "positions", None) or parse_positions_from_header(path)
    if rules is None:
        rules = getattr(workers, "rules", None) or parse_position_rules(path, positions)

    # Copia de seguridad automática previa
    backup_path = path.with_suffix(".xlsx.bak")
    shutil.copy2(path, backup_path)

    # Leer archivo XLSX existente
    with zipfile.ZipFile(path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    wb_root = ET.fromstring(files["xl/workbook.xml"])
    rels_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])

    def _resolve_sheet_path(name_part: str, default_target: str) -> str:
        s_id = None
        for sheet_el in wb_root.findall(f".//{{{MAIN}}}sheet"):
            if name_part in sheet_el.attrib.get("name", "").lower():
                s_id = sheet_el.attrib.get(f"{{{REL}}}id")
                break
        if not s_id:
            return default_target
        for rel_el in rels_root.findall(".//*[@Id]"):
            if rel_el.attrib.get("Id") == s_id:
                tgt = rel_el.attrib.get("Target", default_target).lstrip("/")
                return f"xl/{tgt}" if not tgt.startswith("xl/") else tgt
        return default_target

    sheet1_path = _resolve_sheet_path("trabajador", "xl/worksheets/sheet1.xml")
    sheet2_path = _resolve_sheet_path("guía", "xl/worksheets/sheet2.xml")
    if sheet2_path not in files:
        sheet2_path = _resolve_sheet_path("reglas", "xl/worksheets/sheet2.xml")

    ET.register_namespace("", MAIN)

    def _set_cell(row_el: ET.Element, c_idx: int, r_1_based: int, val: str | int | None) -> None:
        ref = _col_letter(c_idx) + str(r_1_based)
        c = row_el.find(f"{{{MAIN}}}c[@r='{ref}']")
        if c is None:
            if val is None or val == "":
                return
            c = ET.SubElement(row_el, f"{{{MAIN}}}c", {"r": ref})
        for ch in list(c):
            c.remove(ch)
        if val is None or val == "":
            c.attrib.pop("t", None)
            return
        if isinstance(val, int):
            c.set("t", "n")
            v = ET.SubElement(c, f"{{{MAIN}}}v")
            v.text = str(val)
        else:
            c.set("t", "inlineStr")
            is_elem = ET.SubElement(c, f"{{{MAIN}}}is")
            t_elem = ET.SubElement(is_elem, f"{{{MAIN}}}t")
            t_elem.text = str(val)

    # 1. Actualizar Hoja 1 (Trabajadores y Preferencias)
    if sheet1_path in files:
        sheet_root = ET.fromstring(files[sheet1_path])
        sheet_data = sheet_root.find(f"{{{MAIN}}}sheetData")
        if sheet_data is not None:
            # Cabecera Fila 3
            hdr_el = sheet_data.find(f"{{{MAIN}}}row[@r='3']")
            if hdr_el is None:
                hdr_el = ET.SubElement(sheet_data, f"{{{MAIN}}}row", {"r": "3", "ht": "26", "customHeight": "1", "s": "16"})

            _set_cell(hdr_el, 0, 3, "N.º")
            _set_cell(hdr_el, 1, 3, "Apellidos y Nombre")
            _set_cell(hdr_el, 2, 3, "BAJA")
            for i, p in enumerate(positions):
                _set_cell(hdr_el, 3 + i, 3, p)
            for i, p in enumerate(positions):
                _set_cell(hdr_el, 3 + len(positions) + i, 3, p)

            # Limpiar celdas sobrantes en fila 3 si se redujeron puestos
            max_hdr_col = 3 + 2 * len(positions)
            for c in list(hdr_el):
                if _col_idx_from_ref(c.attrib.get("r", "A1")) >= max_hdr_col:
                    hdr_el.remove(c)
            hdr_el[:] = sorted(hdr_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            # Limpiar filas previas de trabajadores (r >= 4)
            for r_el in list(sheet_data):
                if int(r_el.attrib.get("r", 0)) >= 4:
                    sheet_data.remove(r_el)

            # Reconstruir filas de trabajadores de forma ordenada
            for idx, w in enumerate(workers, 1):
                r_num = 3 + idx
                w.row_index = r_num
                r_el = ET.SubElement(sheet_data, f"{{{MAIN}}}row", {"r": str(r_num), "ht": "24", "customHeight": "1", "s": "16"})
                _set_cell(r_el, 0, r_num, idx)
                _set_cell(r_el, 1, r_num, w.name)
                _set_cell(r_el, 2, r_num, "X" if w.is_baja else None)

                for i, p in enumerate(positions):
                    st = w.skills.get(p, "")
                    _set_cell(r_el, 3 + i, r_num, st if st in ("X", "V") else None)

                for i, p in enumerate(positions):
                    if w.knows_position(p) and p in w.preference_order:
                        rank = w.preference_order.index(p) + 1
                        _set_cell(r_el, 3 + len(positions) + i, r_num, rank)

                r_el[:] = sorted(r_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            # Actualizar dimensión
            dim = sheet_root.find(f"{{{MAIN}}}dimension")
            if dim is not None:
                dim.set("ref", f"A1:{_col_letter(2 + 2 * len(positions))}{3 + len(workers)}")

            files[sheet1_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    # 2. Actualizar Hoja 2 (Guía de Puestos y Reglas) si existe
    if sheet2_path in files and rules:
        s2_root = ET.fromstring(files[sheet2_path])
        s2_data = s2_root.find(f"{{{MAIN}}}sheetData")
        if s2_data is not None:
            # Eliminar filas antiguas de catálogo (r entre 4 y 19)
            for r_el in list(s2_data):
                r_n = int(r_el.attrib.get("r", 0))
                if 4 <= r_n <= 19:
                    s2_data.remove(r_el)

            for idx, p in enumerate(positions, 1):
                r_num = 3 + idx
                r_el = ET.SubElement(s2_data, f"{{{MAIN}}}row", {"r": str(r_num), "ht": "24", "customHeight": "1", "s": "16"})
                r_info = rules.get(p, {})
                name = r_info.get("name", p)
                tm = r_info.get("target_m", 1)
                tt = r_info.get("target_t", 1)
                mm = r_info.get("min_m", tm)
                mt = r_info.get("min_t", tt)
                prio = r_info.get("priority", 0)

                if p == "O" and tm == 0 and tt == 0:
                    tgt_str = "1 M (M-J) / 1 T (V)"
                    crit_str = "Martes y Jueves Mañana; Viernes Tarde"
                else:
                    tgt_str = f"{tm} M / {tt} T"
                    if mm == tm and mt == tt and prio == 0:
                        crit_str = "Estricto"
                    elif prio == 2:
                        crit_str = f"Tarde reducible (a {mt} T) después de reducción 1"
                    elif prio == 1:
                        crit_str = f"Puesto reducible si falta personal (hasta {mm} M / {mt} T)"
                    else:
                        crit_str = f"Reducible (hasta {mm} M / {mt} T)"

                _set_cell(r_el, 0, r_num, p)
                _set_cell(r_el, 1, r_num, name)
                _set_cell(r_el, 2, r_num, tgt_str)
                _set_cell(r_el, 3, r_num, crit_str)
                r_el[:] = sorted(r_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            s2_data[:] = sorted(s2_data, key=lambda r: int(r.attrib.get("r", 0)))
            files[sheet2_path] = ET.tostring(s2_root, encoding="utf-8", xml_declaration=True)

    # Guardar en archivo temporal y sustituir atómicamente
    temp_path = path.with_suffix(".tmp")
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    temp_path.replace(path)
