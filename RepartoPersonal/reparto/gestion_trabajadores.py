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

from reparto.datos import POSITIONS, _parse_xlsx_cells

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@dataclass
class EditableWorker:
    """Representación mutable de un trabajador para la interfaz de edición."""
    name: str
    row_index: int
    is_baja: bool
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


def load_editable_workers(xlsx_path: str | Path) -> list[EditableWorker]:
    """Carga los trabajadores desde la hoja 'Trabajadores y Preferencias' del Excel."""
    path = Path(xlsx_path)
    data = path.read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="Trabajadores")

    max_row = max((r for r, c in cells), default=0)
    max_col = max((c for r, c in cells), default=0)

    # Detectar columnas en la cabecera (Fila 3 de Excel -> índice 2)
    header_row = 2
    col_name = 1
    col_baja = next((c for c in range(max_col + 1) if "BAJA" in str(cells.get((header_row, c)) or "").upper()), 2)

    pos_cols = [c for c in range(max_col + 1) if str(cells.get((header_row, c)) or "").strip().upper() in POSITIONS]
    if len(pos_cols) >= 32:
        pos_cols_skills = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols[:16]}
        pos_cols_prefs = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols[16:32]}
    else:
        pos_cols_skills = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols}
        pos_cols_prefs = {}

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

        for pos in POSITIONS:
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
        known_positions = [pos for pos in POSITIONS if skills[pos] == "X"]
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

    return workers


def save_editable_workers(xlsx_path: str | Path, workers: list[EditableWorker]) -> None:
    """Guarda las modificaciones exclusivamente en 'Plantilla_Conocimientos_y_Preferencias.xlsx'.

    Crea una copia de seguridad previa (.bak) y solo modifica las celdas de BAJA,
    competencias y orden de preferencia de los trabajadores usando la biblioteca estándar.
    """
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de plantilla en: {path}")

    # Copia de seguridad automática previa
    backup_path = path.with_suffix(".xlsx.bak")
    shutil.copy2(path, backup_path)

    # Leer archivo XLSX existente
    with zipfile.ZipFile(path, "r") as zin:
        files = {name: zin.read(name) for name in zin.namelist()}

    # Localizar la hoja de trabajadores en workbook.xml y su relación
    wb_root = ET.fromstring(files["xl/workbook.xml"])
    rels_root = ET.fromstring(files["xl/_rels/workbook.xml.rels"])

    sheet_r_id = None
    for sheet_el in wb_root.findall(f".//{{{MAIN}}}sheet"):
        s_name = sheet_el.attrib.get("name", "").lower()
        if "trabajador" in s_name or "preferencia" in s_name:
            sheet_r_id = sheet_el.attrib.get(f"{{{REL}}}id")
            break
    if not sheet_r_id:
        sheet_el = wb_root.find(f".//{{{MAIN}}}sheet")
        if sheet_el is not None:
            sheet_r_id = sheet_el.attrib.get(f"{{{REL}}}id")

    sheet_target = "worksheets/sheet1.xml"
    if sheet_r_id:
        for rel_el in rels_root.findall(".//*[@Id]"):
            if rel_el.attrib.get("Id") == sheet_r_id:
                sheet_target = rel_el.attrib.get("Target", sheet_target).lstrip("/")
                break

    sheet_path = f"xl/{sheet_target}" if not sheet_target.startswith("xl/") else sheet_target
    if sheet_path not in files:
        sheet_path = "xl/worksheets/sheet1.xml"

    ET.register_namespace("", MAIN)
    sheet_root = ET.fromstring(files[sheet_path])
    sheet_data = sheet_root.find(f"{{{MAIN}}}sheetData")
    if sheet_data is None:
        raise ValueError(f"No se encontró el elemento sheetData en {sheet_path}")

    # Identificar columnas en la fila de cabecera (Fila 3 de Excel)
    header_row_el = sheet_data.find(f"{{{MAIN}}}row[@r='3']")
    col_name = 1
    col_baja = 2
    pos_cols_skills: dict[str, int] = {}
    pos_cols_prefs: dict[str, int] = {}

    if header_row_el is not None:
        for c_el in header_row_el.findall(f"{{{MAIN}}}c"):
            ref = c_el.attrib.get("r", "")
            c_idx = _col_idx_from_ref(ref)
            val = "".join(c_el.itertext()).strip().upper()
            if "BAJA" in val:
                col_baja = c_idx
            elif val in POSITIONS:
                if val not in pos_cols_skills:
                    pos_cols_skills[val] = c_idx
                elif val not in pos_cols_prefs:
                    pos_cols_prefs[val] = c_idx

    # Si no se detectaron preferencias como segundo bloque, deducir por offset
    if not pos_cols_prefs and pos_cols_skills:
        pos_cols_prefs = {pos: idx + 16 for pos, idx in pos_cols_skills.items()}

    # Indexar filas por nombre de trabajador
    row_elems_by_name = {}
    for r_el in sheet_data.findall(f"{{{MAIN}}}row"):
        r_num = int(r_el.attrib.get("r", 0))
        if r_num < 4:
            continue
        c_name_el = r_el.find(f"{{{MAIN}}}c[@r='{_col_letter(col_name)}{r_num}']")
        if c_name_el is not None:
            n = "".join(c_name_el.itertext()).strip().upper()
            if n:
                row_elems_by_name[n] = r_el

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

    for w in workers:
        r_el = row_elems_by_name.get(w.name.strip().upper())
        if r_el is None:
            r_el = sheet_data.find(f"{{{MAIN}}}row[@r='{w.row_index}']")
        if r_el is None:
            continue
        r_num = int(r_el.attrib.get("r", w.row_index))

        # 1. Columna BAJA
        _set_cell(r_el, col_baja, r_num, "X" if w.is_baja else None)

        # 2. Competencias
        for pos in POSITIONS:
            c_skill = pos_cols_skills.get(pos)
            if c_skill is not None:
                st = w.skills.get(pos, "")
                _set_cell(r_el, c_skill, r_num, st if st in ("X", "V") else None)

        # 3. Preferencias
        for pos in POSITIONS:
            c_pref = pos_cols_prefs.get(pos)
            if c_pref is not None:
                if w.knows_position(pos) and pos in w.preference_order:
                    rank = w.preference_order.index(pos) + 1
                    _set_cell(r_el, c_pref, r_num, rank)
                else:
                    _set_cell(r_el, c_pref, r_num, None)

        # Ordenar celdas por columna para estricta conformidad OpenXML
        r_el[:] = sorted(r_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

    # Guardar XML actualizado en el mapa de archivos
    files[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    # Escribir de forma atómica el nuevo XLSX
    temp_path = path.with_suffix(".tmp")
    with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    temp_path.replace(path)
