"""Gestión y persistencia de trabajadores, competencias, vetos y preferencias.

Permite consultar, modificar y persistir la configuración de trabajadores
directamente sobre el archivo 'Plantilla_Conocimientos_y_Preferencias.xlsx',
garantizando que el archivo de secuencias de turnos NUNCA sea alterado.

Implementado exclusivamente con la biblioteca estándar de Python (zipfile y xml),
sin depender de paquetes externos como openpyxl.
"""

from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET
import zipfile

from reparto._xml_io import (
    MAIN_NS as MAIN,
    REL_NS as REL,
    atomic_write_zip,
    create_rotating_backup,
    parse_xml_preserving_namespaces,
    serialize_xml_preserving_namespaces,
)
from reparto.datos import (
    POSITIONS,
    _col_idx_from_ref,
    _col_letter,
    _parse_xlsx_cells,
    parse_positions_from_header,
    parse_position_rules,
)


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
    col_baja = next(
        (c for c in range(max_col + 1) if "BAJA" in str(cells.get((header_row, c)) or "").upper()),
        None,
    )

    pos_cols = [c for c in range(max_col + 1) if str(cells.get((header_row, c)) or "").strip().upper() in positions]
    k = len(positions)
    if len(pos_cols) >= 2 * k:
        pos_cols_skills = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols[:k]}
        pos_cols_prefs = {str(cells.get((header_row, c))).strip().upper(): c for c in pos_cols[k:2 * k]}
    else:
        offset = 1 if col_baja == 2 else 0
        pos_cols_skills = {pos: 2 + offset + idx for idx, pos in enumerate(positions)}
        pos_cols_prefs = {pos: 2 + offset + k + idx for idx, pos in enumerate(positions)}

    workers: list[EditableWorker] = []

    for r in range(header_row + 1, max_row + 1):
        name_val = cells.get((r, col_name))
        if not name_val or not str(name_val).strip():
            continue
        name = " ".join(str(name_val).strip().split())
        if any(term in name.upper() for term in ("APELLIDOS", "NOMBRE", "FICHA", "TOTAL")):
            continue

        # Comprobar estado de baja solo si existe columna BAJA explícita
        baja_val = str(cells.get((r, col_baja)) or "").strip().upper() if col_baja is not None else ""
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


def _cell_text(c_el: ET.Element | None) -> str:
    if c_el is None:
        return ""
    return "".join(t.text or "" for t in c_el.iter(f"{{{MAIN}}}t")).strip()


def save_editable_workers(
    xlsx_path: str | Path,
    workers: list[EditableWorker],
    positions: tuple[str, ...] | None = None,
    rules: dict[str, dict] | None = None,
) -> None:
    """Guarda las modificaciones en 'Plantilla_Conocimientos_y_Preferencias.xlsx'.

    Reconstruye de forma limpia las filas de trabajadores, cabeceras dinámicas
    y columnas estructuradas de reglas en Hoja 2, preservando espacios de nombres
    OpenXML, estilos de celda y bloques informativos.
    """
    path = Path(xlsx_path)
    if not path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de plantilla en: {path}")

    if positions is None:
        positions = getattr(workers, "positions", None) or parse_positions_from_header(path)
    if rules is None:
        rules = getattr(workers, "rules", None) or parse_position_rules(path, positions)

    # Copia de seguridad inmediata (.bak) y rotativa con fecha (copias/)
    create_rotating_backup(path)

    # Leer bytes originales y celdas pre-resueltas (incluyendo sharedStrings si Excel guardó el archivo)
    raw_xlsx_bytes = path.read_bytes()
    s2_parsed_cells: dict[tuple[int, int], object] = {}
    try:
        _, s2_parsed_cells = _parse_xlsx_cells(raw_xlsx_bytes, sheet_name_contains="Guía")
        if not s2_parsed_cells:
            _, s2_parsed_cells = _parse_xlsx_cells(raw_xlsx_bytes, sheet_name_contains="Reglas")
    except Exception:
        s2_parsed_cells = {}

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

    def _set_cell(
        row_el: ET.Element,
        c_idx: int,
        r_1_based: int,
        val: str | int | None,
        style: str | None = None,
    ) -> None:
        ref = _col_letter(c_idx) + str(r_1_based)
        c = row_el.find(f"{{{MAIN}}}c[@r='{ref}']")
        if c is None:
            if val is None or val == "":
                if style is None:
                    return
                c = ET.SubElement(row_el, f"{{{MAIN}}}c", {"r": ref, "s": style})
                return
            attrs = {"r": ref}
            if style is not None:
                attrs["s"] = style
            c = ET.SubElement(row_el, f"{{{MAIN}}}c", attrs)
        elif style is not None and "s" not in c.attrib:
            c.set("s", style)
        for ch in list(c):
            c.remove(ch)
        if val is None or val == "":
            c.attrib.pop("t", None)
            return
        if isinstance(val, int) and not isinstance(val, bool):
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
        sheet_root, ns_map1 = parse_xml_preserving_namespaces(files[sheet1_path])
        sheet_data = sheet_root.find(f"{{{MAIN}}}sheetData")
        if sheet_data is not None:
            # Capturar estilos de cabecera (fila 3) y de primera fila de datos (fila 4)
            hdr_el = sheet_data.find(f"{{{MAIN}}}row[@r='3']")
            if hdr_el is None:
                hdr_el = ET.SubElement(
                    sheet_data, f"{{{MAIN}}}row", {"r": "3", "ht": "26", "customHeight": "1", "s": "16"}
                )
            hdr_c0 = hdr_el.find(f"{{{MAIN}}}c[@r='A3']")
            hdr_c3 = hdr_el.find(f"{{{MAIN}}}c[@r='D3']")
            if hdr_c3 is None:
                hdr_c3 = hdr_el.find(f"{{{MAIN}}}c[@r='C3']")
            s_hdr_fixed = hdr_c0.attrib.get("s", "25") if hdr_c0 is not None else "25"
            s_hdr_pos = hdr_c3.attrib.get("s", "19") if hdr_c3 is not None else "19"

            tpl_r4 = sheet_data.find(f"{{{MAIN}}}row[@r='4']")
            s_row = tpl_r4.attrib.get("s", "16") if tpl_r4 is not None else "16"
            s_ht = tpl_r4.attrib.get("ht", "24") if tpl_r4 is not None else "24"
            c_a4 = tpl_r4.find(f"{{{MAIN}}}c[@r='A4']") if tpl_r4 is not None else None
            c_b4 = tpl_r4.find(f"{{{MAIN}}}c[@r='B4']") if tpl_r4 is not None else None
            c_c4 = tpl_r4.find(f"{{{MAIN}}}c[@r='C4']") if tpl_r4 is not None else None
            s_col_a = c_a4.attrib.get("s", "28") if c_a4 is not None else "28"
            s_col_b = c_b4.attrib.get("s", "27") if c_b4 is not None else "27"
            s_col_data = c_c4.attrib.get("s", "28") if c_c4 is not None else "28"

            _set_cell(hdr_el, 0, 3, "N.º", style=s_hdr_fixed)
            _set_cell(hdr_el, 1, 3, "Apellidos y Nombre", style=s_hdr_fixed)
            _set_cell(hdr_el, 2, 3, "BAJA", style=s_hdr_pos)
            for i, p in enumerate(positions):
                _set_cell(hdr_el, 3 + i, 3, p, style=s_hdr_pos)
            for i, p in enumerate(positions):
                _set_cell(hdr_el, 3 + len(positions) + i, 3, p, style=s_hdr_pos)

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

            # Reconstruir filas de trabajadores de forma ordenada y preservando estilos
            for idx, w in enumerate(workers, 1):
                r_num = 3 + idx
                w.row_index = r_num
                r_el = ET.SubElement(
                    sheet_data,
                    f"{{{MAIN}}}row",
                    {"r": str(r_num), "ht": s_ht, "customHeight": "1", "s": s_row},
                )
                _set_cell(r_el, 0, r_num, idx, style=s_col_a)
                _set_cell(r_el, 1, r_num, w.name, style=s_col_b)
                _set_cell(r_el, 2, r_num, "X" if w.is_baja else None, style=s_col_data)

                for i, p in enumerate(positions):
                    st = w.skills.get(p, "")
                    _set_cell(r_el, 3 + i, r_num, st if st in ("X", "V") else None, style=s_col_data)

                for i, p in enumerate(positions):
                    if w.knows_position(p) and p in w.preference_order:
                        rank = w.preference_order.index(p) + 1
                        _set_cell(r_el, 3 + len(positions) + i, r_num, rank, style=s_col_data)
                    else:
                        _set_cell(r_el, 3 + len(positions) + i, r_num, None, style=s_col_data)

                r_el[:] = sorted(r_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            # Actualizar dimensión, anchos de columna (<cols>) y celdas combinadas (<mergeCells>) en Hoja 1 (N5)
            k_pos = len(positions)
            last_col_idx = 2 + 2 * k_pos
            dim = sheet_root.find(f"{{{MAIN}}}dimension")
            if dim is not None:
                dim.set("ref", f"A1:{_col_letter(last_col_idx)}{max(4, 3 + len(workers))}")

            cols_parent = sheet_root.find(f"{{{MAIN}}}cols")
            if cols_parent is not None:
                for col_el in list(cols_parent):
                    if int(col_el.attrib.get("min", "0")) >= 3:
                        col_el.set("max", str(last_col_idx + 1))

            r1_el = sheet_data.find(f"{{{MAIN}}}row[@r='1']")
            r2_el = sheet_data.find(f"{{{MAIN}}}row[@r='2']")
            if r1_el is not None and r2_el is not None and k_pos >= 1:
                c1_cell = r1_el.find(f"{{{MAIN}}}c[@r='C1']")
                c2_cell = r2_el.find(f"{{{MAIN}}}c[@r='C2']")
                s_r1 = c1_cell.attrib.get("s", "18") if c1_cell is not None else "18"
                s_r2 = c2_cell.attrib.get("s", "22") if c2_cell is not None else "22"
                _set_cell(r1_el, 2, 1, "ESTADO", style=s_r1)
                _set_cell(r2_el, 2, 2, "X = De baja", style=s_r2)
                _set_cell(r1_el, 3, 1, "CONOCIMIENTOS Y VETOS POR PUESTO", style=s_r1)
                _set_cell(r2_el, 3, 2, "X = Sabe el puesto  ·  Vacío = No sabe  ·  V = Vetado", style=s_r2)
                _set_cell(r1_el, 3 + k_pos, 1, "ORDEN DE PREFERENCIA (1 = FAVORITO)", style=s_r1)
                _set_cell(r2_el, 3 + k_pos, 2, "Numerar de 1 a N los puestos sabidos según preferencia", style=s_r2)
                for rx_el in (r1_el, r2_el):
                    for c_el in list(rx_el):
                        if _col_idx_from_ref(c_el.attrib.get("r", "A1")) > last_col_idx:
                            rx_el.remove(c_el)
                    rx_el[:] = sorted(rx_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            mc_parent = sheet_root.find(f"{{{MAIN}}}mergeCells")
            if mc_parent is not None and k_pos >= 1:
                for ch in list(mc_parent):
                    mc_parent.remove(ch)
                new_merges = ["A1:B1", "A2:B2"]
                if k_pos > 1:
                    end_sk = _col_letter(2 + k_pos)
                    start_pr = _col_letter(3 + k_pos)
                    end_pr = _col_letter(2 + 2 * k_pos)
                    new_merges.extend([
                        f"D1:{end_sk}1",
                        f"D2:{end_sk}2",
                        f"{start_pr}1:{end_pr}1",
                        f"{start_pr}2:{end_pr}2",
                    ])
                mc_parent.set("count", str(len(new_merges)))
                for m_ref in new_merges:
                    ET.SubElement(mc_parent, f"{{{MAIN}}}mergeCell", {"ref": m_ref})

            files[sheet1_path] = serialize_xml_preserving_namespaces(sheet_root, ns_map1)

    # 2. Actualizar Hoja 2 (Guía de Puestos y Reglas) — crearla si no existe y hay reglas
    if rules and sheet2_path not in files:
        files[sheet2_path] = (
            b'<?xml version="1.0" encoding="utf-8"?>'
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            b'<sheetData>'
            b'<row r="3"><c r="A3" t="inlineStr"><is><t>C\xc3\xb3digo</t></is></c>'
            b'<c r="B3" t="inlineStr"><is><t>Denominaci\xc3\xb3n del Puesto</t></is></c>'
            b'<c r="C3" t="inlineStr"><is><t>Dotaci\xc3\xb3n Objetivo (L-V)</t></is></c>'
            b'<c r="D3" t="inlineStr"><is><t>Criterio / Prioridad de Reducci\xc3\xb3n</t></is></c></row>'
            b'</sheetData></worksheet>'
        )
        if "xl/workbook.xml" in files and "xl/_rels/workbook.xml.rels" in files:
            wb_root, wb_ns = parse_xml_preserving_namespaces(files["xl/workbook.xml"])
            rels_root, rels_ns = parse_xml_preserving_namespaces(files["xl/_rels/workbook.xml.rels"])
            REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
            R_ID_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            wb_ns.setdefault("r", R_ID_NS)
            ET.SubElement(
                rels_root,
                f"{{{REL_NS}}}Relationship",
                {
                    "Id": "rIdSheet2Rules",
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    "Target": "worksheets/sheet2.xml",
                },
            )
            sheets_el = wb_root.find(f"{{{MAIN}}}sheets")
            if sheets_el is not None:
                ET.SubElement(
                    sheets_el,
                    f"{{{MAIN}}}sheet",
                    {"name": "Guía de Puestos y Reglas", "sheetId": "2", f"{{{R_ID_NS}}}id": "rIdSheet2Rules"},
                )
            files["xl/workbook.xml"] = serialize_xml_preserving_namespaces(wb_root, wb_ns)
            files["xl/_rels/workbook.xml.rels"] = serialize_xml_preserving_namespaces(rels_root, rels_ns)
        if "[Content_Types].xml" in files:
            ct_root, ct_ns = parse_xml_preserving_namespaces(files["[Content_Types].xml"])
            CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
            ET.SubElement(
                ct_root,
                f"{{{CT_NS}}}Override",
                {
                    "PartName": "/xl/worksheets/sheet2.xml",
                    "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                },
            )
            files["[Content_Types].xml"] = serialize_xml_preserving_namespaces(ct_root, ct_ns)

    if sheet2_path in files and rules:
        s2_root, ns_map2 = parse_xml_preserving_namespaces(files[sheet2_path])
        s2_data = s2_root.find(f"{{{MAIN}}}sheetData")
        if s2_data is not None:
            # Asegurar cabeceras explícitas en Fila 3 para persistencia sin pérdida (C2)
            s2_hdr = s2_data.find(f"{{{MAIN}}}row[@r='3']")
            if s2_hdr is not None:
                c_a3 = s2_hdr.find(f"{{{MAIN}}}c[@r='A3']")
                s_h2 = c_a3.attrib.get("s", "25") if c_a3 is not None else "25"
                for offset, h_title in enumerate(("Obj. M", "Obj. T", "Mín. M", "Mín. T", "Prioridad")):
                    _set_cell(s2_hdr, 4 + offset, 3, h_title, style=s_h2)
                s2_hdr[:] = sorted(s2_hdr, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            tpl2_r4 = s2_data.find(f"{{{MAIN}}}row[@r='4']")
            s2_c0 = tpl2_r4.find(f"{{{MAIN}}}c[@r='A4']") if tpl2_r4 is not None else None
            s2_c1 = tpl2_r4.find(f"{{{MAIN}}}c[@r='B4']") if tpl2_r4 is not None else None
            s2_center = s2_c0.attrib.get("s", "31") if s2_c0 is not None else "31"
            s2_left = s2_c1.attrib.get("s", "30") if s2_c1 is not None else "30"

            # Separar filas de catálogo respecto al bloque inferior ("CONCEPTOS CLAVE...")
            # Usa s2_parsed_cells para soportar celdas t="s" (sharedStrings) creadas por Microsoft Excel (N2)
            footer_rows: list[ET.Element] = []
            in_footer = False
            old_legend_start = None
            for r_el in list(s2_data):
                r_n = int(r_el.attrib.get("r", 0))
                if r_n < 4:
                    continue
                parsed_val = s2_parsed_cells.get((r_n - 1, 0))
                first_txt = (
                    str(parsed_val).strip()
                    if parsed_val is not None
                    else _cell_text(r_el.find(f"{{{MAIN}}}c[@r='A{r_n}']"))
                )
                if not in_footer and (not first_txt or "CONCEPTOS CLAVE" in first_txt.upper()):
                    in_footer = True
                    old_legend_start = r_n
                s2_data.remove(r_el)
                if in_footer:
                    footer_rows.append(r_el)

            for idx, p in enumerate(positions, 1):
                r_num = 3 + idx
                r_el = ET.SubElement(
                    s2_data, f"{{{MAIN}}}row", {"r": str(r_num), "ht": "24", "customHeight": "1", "s": "16"}
                )
                r_info = rules.get(p, {})
                name = r_info.get("name", p)
                tm = int(r_info.get("target_m", 1))
                tt = int(r_info.get("target_t", 1))
                mm = int(r_info.get("min_m", tm))
                mt = int(r_info.get("min_t", tt))
                prio = int(r_info.get("priority", 0))

                if p == "O" and tm == 0 and tt == 0:
                    tgt_str = "1 M (M-J) / 1 T (V)"
                    crit_str = "Martes y Jueves Mañana; Viernes Tarde"
                else:
                    tgt_str = f"{tm} M / {tt} T"
                    if mm == tm and mt == tt and prio == 0:
                        crit_str = "Estricto"
                    elif prio == 2:
                        crit_str = f"Reducible (hasta {mm} M / {mt} T) después de reducción 1"
                    elif prio == 1:
                        crit_str = f"Puesto reducible si falta personal (hasta {mm} M / {mt} T)"
                    else:
                        crit_str = f"Estricto (mínimo {mm} M / {mt} T)"

                _set_cell(r_el, 0, r_num, p, style=s2_center)
                _set_cell(r_el, 1, r_num, name, style=s2_left)
                _set_cell(r_el, 2, r_num, tgt_str, style=s2_center)
                _set_cell(r_el, 3, r_num, crit_str, style=s2_left)
                # Columnas estructuradas exactas (E..I) para ida y vuelta sin pérdida
                _set_cell(r_el, 4, r_num, tm, style=s2_center)
                _set_cell(r_el, 5, r_num, tt, style=s2_center)
                _set_cell(r_el, 6, r_num, mm, style=s2_center)
                _set_cell(r_el, 7, r_num, mt, style=s2_center)
                _set_cell(r_el, 8, r_num, prio, style=s2_center)
                r_el[:] = sorted(r_el, key=lambda c: _col_idx_from_ref(c.attrib.get("r", "A1")))

            # Reinsertar bloque informativo desplazado tras el último puesto
            new_legend_start = 4 + len(positions)
            delta = (new_legend_start - old_legend_start) if old_legend_start is not None else 0
            for f_el in footer_rows:
                old_r = int(f_el.attrib.get("r", 0))
                new_r = old_r + delta
                f_el.set("r", str(new_r))
                for c_el in f_el.findall(f"{{{MAIN}}}c"):
                    c_idx = _col_idx_from_ref(c_el.attrib.get("r", "A1"))
                    c_el.set("r", f"{_col_letter(c_idx)}{new_r}")
                s2_data.append(f_el)

            # Desplazar mergeCells del bloque informativo si cambió el número de puestos
            if delta != 0 and old_legend_start is not None:
                merge_parent = s2_root.find(f"{{{MAIN}}}mergeCells")
                if merge_parent is not None:
                    import re as _re
                    for mc_el in merge_parent.findall(f"{{{MAIN}}}mergeCell"):
                        ref = mc_el.attrib.get("ref", "")
                        m = _re.fullmatch(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
                        if m and int(m.group(2)) >= old_legend_start:
                            r1 = int(m.group(2)) + delta
                            r2 = int(m.group(4)) + delta
                            mc_el.set("ref", f"{m.group(1)}{r1}:{m.group(3)}{r2}")

            s2_data[:] = sorted(s2_data, key=lambda r: int(r.attrib.get("r", 0)))
            files[sheet2_path] = serialize_xml_preserving_namespaces(s2_root, ns_map2)

    # Escritura atómica con manejo explícito de bloqueo en Excel
    atomic_write_zip(path, files)
