"""Núcleo de Datos y Calendario para el sistema de reparto.

Maneja de forma limpia, robusta e independiente de librerías externas:
1. Catálogo oficial de los 16 puestos y reglas de dotación.
2. Lector del Excel de Secuencias de Turnos (168 días / 24 semanas).
3. Lector del Excel de Conocimientos, Vetos y Preferencias.
4. Calendario oficial de festivos y reglas de exclusión (fines de semana y festivos).
5. Proyector de fechas: mapea cualquier día del año a su turno correspondiente del ciclo.
"""

from dataclasses import dataclass
from datetime import date, timedelta
from hashlib import sha1
from io import BytesIO
import io
from pathlib import Path
import posixpath
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from xml.etree.ElementTree import fromstring
from zipfile import ZipFile

# Catálogo oficial de puestos y dotaciones base en laborables (por defecto)
DEFAULT_POSITIONS = (
    "EC", "MG", "PU", "DR", "UR", "A",
    "RM1", "RM2", "RM3", "TC1", "TC2", "TC3", "TCU",
    "TE", "RT", "O"
)
POSITIONS = DEFAULT_POSITIONS

# Dotaciones objetivo en días laborables estándar
# Formato: (Mañana_Objetivo, Tarde_Objetivo, Mañana_Minimo, Tarde_Minimo)
DEFAULT_POSITION_RULES = {
    "EC":  {"target_m": 4, "target_t": 2, "min_m": 3, "min_t": 1, "priority": 1, "name": "Ecosonografía"},
    "MG":  {"target_m": 2, "target_t": 1, "min_m": 2, "min_t": 1, "priority": 0, "name": "Mamografía"},
    "PU":  {"target_m": 1, "target_t": 0, "min_m": 1, "min_t": 0, "priority": 0, "name": "Punción / Eco Especial"},
    "DR":  {"target_m": 3, "target_t": 3, "min_m": 3, "min_t": 2, "priority": 2, "name": "Radiología Digital General"},
    "UR":  {"target_m": 2, "target_t": 2, "min_m": 2, "min_t": 2, "priority": 0, "name": "Urgencias"},
    "A":   {"target_m": 2, "target_t": 0, "min_m": 2, "min_t": 0, "priority": 0, "name": "Angiografía / Intervencionismo"},
    "RM1": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Resonancia Magnética 1"},
    "RM2": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Resonancia Magnética 2"},
    "RM3": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Resonancia Magnética 3"},
    "TC1": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Tomografía Computarizada 1"},
    "TC2": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Tomografía Computarizada 2"},
    "TC3": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Tomografía Computarizada 3"},
    "TCU": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "TC Urgencias"},
    "TE":  {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Telemando"},
    "RT":  {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0, "name": "Radioterapia / Radiología"},
    # O (Quirófano/Otros): Se modula según día de semana (Martes y Jueves mañana; Viernes tarde)
    "O":   {"target_m": 0, "target_t": 0, "min_m": 0, "min_t": 0, "priority": 0, "name": "Quirófano / Otros"},
}
POSITION_RULES = DEFAULT_POSITION_RULES


def parse_positions_from_header(xlsx_path: str | Path) -> tuple[str, ...]:
    """Detecta dinámicamente el catálogo de puestos desde la cabecera (Fila 3) de la plantilla.

    Lee secuencialmente las columnas tras 'BAJA'. El primer bloque de códigos no vacíos
    hasta que un código se repite (inicio del bloque de preferencias) o se acaben las columnas
    define el catálogo de puestos. Si el archivo no existe o no define puestos, recurre a DEFAULT_POSITIONS.
    """
    path = Path(xlsx_path)
    if not path.exists():
        return DEFAULT_POSITIONS
    try:
        data = path.read_bytes()
        _, cells = _parse_xlsx_cells(data, sheet_name_contains="Trabajadores")
        header_row = 2  # Fila 3 en Excel (0-indexed 2)
        max_col = max((c for r, c in cells if r == header_row), default=0)

        # Localizar columna BAJA o empezar tras columna B (índice 1)
        baja_col = next((c for c in range(max_col + 1) if "BAJA" in str(cells.get((header_row, c)) or "").upper()), 2)

        positions: list[str] = []
        for c in range(baja_col + 1, max_col + 1):
            val = str(cells.get((header_row, c)) or "").strip()
            if not val:
                continue
            code = val.upper()
            if code in positions:
                # Comienza el segundo bloque (preferencias)
                break
            positions.append(code)

        if positions:
            return tuple(positions)
    except Exception:
        pass
    return DEFAULT_POSITIONS


def parse_position_rules(
    xlsx_path: str | Path,
    positions: tuple[str, ...] | None = None
) -> dict[str, dict]:
    """Lee las dotaciones objetivo, mínimos y criterios de reducción desde la Hoja 2.

    Si no existe la Hoja 2 o faltan puestos, aplica valores por defecto inteligentes.
    """
    path = Path(xlsx_path)
    if positions is None:
        positions = parse_positions_from_header(path)

    rules = {}
    if path.exists():
        try:
            data = path.read_bytes()
            sname, cells = _parse_xlsx_cells(data, sheet_name_contains="Guía")
            if not cells:
                sname, cells = _parse_xlsx_cells(data, sheet_name_contains="Reglas")

            max_r = max((r for r, c in cells), default=0)
            for r in range(2, max_r + 1):
                code_raw = cells.get((r, 0))
                if not code_raw:
                    continue
                code = str(code_raw).strip().upper()
                if not code or code in ("CÓDIGO", "CODIGO", "PUESTO", "TOTAL"):
                    continue

                name = str(cells.get((r, 1)) or code).strip()
                target_str = str(cells.get((r, 2)) or "").strip()
                crit_str = str(cells.get((r, 3)) or "").strip()

                # Extraer objetivos
                tm, tt = 1, 1
                m_tgt = re.search(r"(\d+)\s*M", target_str, re.IGNORECASE)
                t_tgt = re.search(r"(\d+)\s*T", target_str, re.IGNORECASE)
                if m_tgt:
                    tm = int(m_tgt.group(1))
                if t_tgt:
                    tt = int(t_tgt.group(1))

                # Extraer mínimos
                mm, mt = tm, tt
                if "hasta" in crit_str.lower():
                    m_min = re.search(r"hasta\s*(\d+)\s*M", crit_str, re.IGNORECASE)
                    t_min = re.search(r"(\d+)\s*T", crit_str, re.IGNORECASE)
                    if m_min:
                        mm = int(m_min.group(1))
                    if t_min:
                        mt = int(t_min.group(1))
                elif "tarde reducible" in crit_str.lower():
                    t_red = re.search(r"a\s*(\d+)\s*T", crit_str, re.IGNORECASE)
                    if t_red:
                        mt = int(t_red.group(1))

                # Extraer prioridad de reducción
                prio = 0
                if "después de" in crit_str.lower() or "despues de" in crit_str.lower():
                    prio = 2
                elif "reducible" in crit_str.lower():
                    prio = 1

                rules[code] = {
                    "name": name,
                    "target_m": tm,
                    "target_t": tt,
                    "min_m": mm,
                    "min_t": mt,
                    "priority": prio,
                }
        except Exception:
            pass

    # Asegurar que todos los puestos solicitados tienen reglas
    final_rules = {}
    for p in positions:
        if p in rules:
            final_rules[p] = rules[p]
        elif p in DEFAULT_POSITION_RULES:
            final_rules[p] = dict(DEFAULT_POSITION_RULES[p])
        else:
            final_rules[p] = {
                "name": p,
                "target_m": 1,
                "target_t": 1,
                "min_m": 1,
                "min_t": 1,
                "priority": 0,
            }
    return final_rules


def get_position_targets(
    pos: str,
    weekday: int,
    position_rules: dict[str, dict] | None = None
) -> tuple[int, int, int, int]:
    """Devuelve (target_m, target_t, min_m, min_t) para un puesto y día de semana (0=Lunes..4=Viernes)."""
    rules = position_rules if position_rules is not None else POSITION_RULES
    if pos == "O" and rules.get("O", {}).get("target_m", 0) == 0:
        # O abre Martes(1) y Jueves(3) por la mañana; Viernes(4) por la tarde
        m = 1 if weekday in (1, 3) else 0
        t = 1 if weekday == 4 else 0
        return m, t, m, t
    r = rules.get(pos)
    if not r:
        return 1, 1, 1, 1
    return r["target_m"], r["target_t"], r["min_m"], r["min_t"]


# --- CALENDARIO OFICIAL DE FESTIVOS LABORALES ---

DEFAULT_HOLIDAYS = {
    2026: {
        date(2026, 1, 1),   # Año Nuevo
        date(2026, 1, 2),   # Festivo Local
        date(2026, 1, 6),   # Epifanía del Señor
        date(2026, 2, 28),  # Día Autonómico
        date(2026, 4, 2),   # Jueves Santo
        date(2026, 4, 3),   # Viernes Santo
        date(2026, 5, 1),   # Fiesta del Trabajo
        date(2026, 6, 4),   # Festivo Local
        date(2026, 8, 15),  # Asunción de la Virgen
        date(2026, 10, 12), # Fiesta Nacional
        date(2026, 11, 2),  # Festivo trasladado a lunes
        date(2026, 12, 7),  # Festivo trasladado a lunes
        date(2026, 12, 8),  # Inmaculada Concepción
        date(2026, 12, 25), # Natividad del Señor
    },
    2027: {
        date(2027, 1, 1),
        date(2027, 1, 2),
        date(2027, 1, 6),
        date(2027, 2, 28),
        date(2027, 3, 25),
        date(2027, 3, 26),
        date(2027, 5, 1),
        date(2027, 5, 27),
        date(2027, 8, 15),
        date(2027, 10, 12),
        date(2027, 11, 1),
        date(2027, 12, 6),
        date(2027, 12, 8),
        date(2027, 12, 25),
    }
}


def _easter_sunday(year: int) -> date:
    """Domingo de Pascua (algoritmo gregoriano anónimo)."""
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month, day = divmod(h + l - 7 * m + 114, 31)
    return date(year, month, day + 1)


def computed_holidays(year: int) -> set[date]:
    """Festivos calculados por regla (Pascua, festivos fijos y locales)."""
    easter = _easter_sunday(year)
    result = {
        date(year, 1, 1), date(year, 1, 2), date(year, 1, 6), date(year, 2, 28),
        date(year, 5, 1), date(year, 8, 15), date(year, 10, 12),
        date(year, 12, 8), date(year, 12, 25),
        easter - timedelta(days=3),    # Jueves Santo
        easter - timedelta(days=2),    # Viernes Santo
        easter + timedelta(days=60),   # Festivo local
    }
    for month, day in ((11, 1), (12, 6)):
        d = date(year, month, day)
        result.add(d + timedelta(days=1) if d.weekday() == 6 else d)
    return result


def holidays_for_year(year: int) -> tuple[set[date], bool]:
    """Devuelve (festivos, verificado)."""
    if year in DEFAULT_HOLIDAYS:
        return set(DEFAULT_HOLIDAYS[year]), True
    return computed_holidays(year), False


def is_workday(d: date) -> bool:
    """Verifica si una fecha es laborable para el reparto (Lunes a Viernes no festivo)."""
    if d.weekday() >= 5:  # 5=Sábado, 6=Domingo
        return False
    holidays, _ = holidays_for_year(d.year)
    return d not in holidays


# --- ESTRUCTURAS DE DATOS ---

@dataclass(frozen=True)
class WorkerProfile:
    id: str
    name: str
    skills: frozenset[str]            # Puestos que sabe desempeñar
    vetoes: frozenset[str]            # Puestos vetados expresamente
    ranking: dict[str, int]           # Puesto -> orden de preferencia (1 = favorito .. 16 = peor)
    sequence: tuple[str, ...]         # 168 marcas ('M', 'T', 'D', 'R')


@dataclass(frozen=True)
class DaySchedule:
    day: date
    cycle_day: int                    # 1 a 168
    weekday: int                      # 0=Lunes .. 6=Domingo
    is_workday: bool                  # True si L..V no festivo
    workers_m: tuple[str, ...]        # IDs de trabajadores disponibles de mañana
    workers_t: tuple[str, ...]        # IDs de trabajadores disponibles de tarde
    workers_r: tuple[str, ...]        # IDs de trabajadores con refuerzo (asignables a M o T)
    targets_m: dict[str, tuple[int, int]]  # Puesto -> (target, minimum)
    targets_t: dict[str, tuple[int, int]]  # Puesto -> (target, minimum)


# --- LECTOR XLSX NATIVO DE CELDAS ---

def _parse_xlsx_cells(data: bytes, sheet_name_contains: str | None = None) -> tuple[str, dict[tuple[int, int], str | float | None]]:
    """Lee un XLSX descomprimiendo el XML directamente, devolviendo {(row, col): valor} (0-indexed)."""
    with ZipFile(BytesIO(data)) as archive:
        workbook_xml = fromstring(archive.read("xl/workbook.xml"))
        ns = workbook_xml.tag.split("}")[0] + "}"
        rels_xml = fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relationships = {r.get("Id"): r for r in rels_xml}

        # Shared strings si existen
        strings = []
        for rel in relationships.values():
            if rel.get("Type", "").endswith("/sharedStrings"):
                val = rel.get("Target", "")
                path = posixpath.normpath(val.lstrip("/") if val.startswith("/") else "xl/" + val)
                ss_xml = fromstring(archive.read(path))
                strings = ["".join(t.text or "" for t in si.iter(ns + "t")) for si in ss_xml]

        # Seleccionar hoja
        target_sheet = None
        for sheet in workbook_xml.findall(".//" + ns + "sheet"):
            sname = sheet.get("name", "")
            if sheet_name_contains is None or sheet_name_contains.lower() in sname.lower():
                target_sheet = sheet
                break
        if target_sheet is None:
            target_sheet = workbook_xml.find(".//" + ns + "sheet")

        sname = target_sheet.get("name", "")
        rid = next((v for k, v in target_sheet.attrib.items() if k.endswith("}id")), None)
        rel_target = relationships[rid].get("Target", "")
        sheet_path = posixpath.normpath(rel_target.lstrip("/") if rel_target.startswith("/") else "xl/" + rel_target)
        ws_xml = fromstring(archive.read(sheet_path))

        cells = {}
        for cell in ws_xml.iter(ns + "c"):
            coord = cell.get("r", "")
            match = re.fullmatch(r"([A-Z]+)([1-9][0-9]*)", coord)
            if not match:
                continue
            col_letters, row_str = match.groups()
            row_idx = int(row_str) - 1
            col_idx = 0
            for char in col_letters:
                col_idx = col_idx * 26 + ord(char) - 64
            col_idx -= 1

            kind = cell.get("t", "n")
            val_elm = cell.find(ns + "v")
            if kind == "inlineStr":
                parsed = "".join(t.text or "" for t in cell.iter(ns + "t"))
            elif kind == "s" and val_elm is not None and val_elm.text:
                parsed = strings[int(val_elm.text)]
            elif kind == "n" and val_elm is not None and val_elm.text:
                try:
                    num = float(val_elm.text)
                    parsed = int(num) if num.is_integer() else num
                except ValueError:
                    parsed = None
            else:
                parsed = val_elm.text if val_elm is not None else None

            cells[(row_idx, col_idx)] = parsed

        return sname, cells


# --- CARGADORES DE CONFIGURACIÓN ---

@dataclass(frozen=True)
class LoadReport:
    """Avisos de la carga de configuración. No bloquean el cálculo: se muestran para revisar."""
    warnings: tuple[str, ...] = ()
    bajas: tuple[str, ...] = ()           # Trabajadores de baja descartados
    sin_ficha: tuple[str, ...] = ()       # En secuencias pero sin ficha de conocimientos
    sin_conocimientos: tuple[str, ...] = ()  # Con ficha pero ningún conocimiento marcado


def _col_letter(c_idx: int) -> str:
    """Convierte índice 0-based a letra de columna (0 -> 'A', 27 -> 'AB')."""
    res = ""
    c_idx += 1
    while c_idx > 0:
        c_idx, rem = divmod(c_idx - 1, 26)
        res = chr(65 + rem) + res
    return res


def _col_idx_from_ref(ref: str) -> int:
    """Extrae el índice de columna (0-based) de una referencia como 'AA12'."""
    col_str = "".join(ch for ch in ref if ch.isalpha()).upper()
    idx = 0
    for ch in col_str:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def _normalize_name(raw: str) -> str:
    return " ".join(raw.split()).upper()


def detect_cycle_weeks(xlsx_path: str | Path, default_weeks: int = 24) -> int:
    """Detecta automáticamente el número de semanas del ciclo a partir del Excel de secuencias."""
    try:
        p = Path(xlsx_path)
        if not p.exists():
            return default_weeks
        data = p.read_bytes()
        _, cells = _parse_xlsx_cells(data, sheet_name_contains="secuencia")

        # 1. Buscar fila con numeración consecutiva 1, 2, 3... a partir de col 2 (Col C)
        for r in range(6):
            c = 2
            consecutive = 0
            while True:
                v = cells.get((r, c))
                try:
                    if int(v) == consecutive + 1:
                        consecutive += 1
                        c += 1
                        continue
                except (ValueError, TypeError):
                    break
            if consecutive >= 7:
                return max(1, consecutive // 7)

        # 2. Alternativa: buscar columnas antes de cabeceras de resumen (Mañanas, Tardes, Total...)
        summary_terms = ("MAÑANA", "TARDE", "DESCANSO", "REFUERZO", "TOTAL", "DÍAS", "DIAS")
        for r in range(6):
            for c in range(2, 400):
                v_str = str(cells.get((r, c)) or "").strip().upper()
                if any(term in v_str for term in summary_terms):
                    days = c - 2
                    if days >= 7:
                        return max(1, days // 7)
                    break

        # 3. Alternativa: inspeccionar longitud de columnas con marcas en filas de trabajadores
        max_row = max((r for r, c in cells), default=0)
        day_counts = []
        for r in range(max_row + 1):
            name_val = cells.get((r, 1))
            if not isinstance(name_val, str) or not name_val.strip():
                continue
            name_clean = _normalize_name(name_val)
            if any(term in name_clean for term in ("APELLIDOS", "NOMBRE", "TOTAL", "PUESTOS", "SEMANA")):
                continue
            c = 2
            while c < 400:
                raw = str(cells.get((r, c)) or "").strip().upper()
                if raw in ("M", "T", "D", "R", ""):
                    c += 1
                else:
                    break
            if (c - 2) >= 7:
                day_counts.append((c - 2) // 7)
        if day_counts:
            from collections import Counter
            return Counter(day_counts).most_common(1)[0][0]
    except Exception:
        pass
    return default_weeks


def load_sequences(
    xlsx_path: str | Path,
    warnings: list[str] | None = None,
    cycle_weeks: int | None = None,
) -> dict[str, tuple[str, ...]]:
    """Carga las secuencias de turnos desde el Excel limpio o original.

    Devuelve: {nombre_normalizado: tuple(marcas 'M', 'T', 'D', 'R')}.
    Celda vacía = descanso. Cualquier otra marca desconocida se trata como descanso y se
    anota en ``warnings`` (si se proporciona la lista).
    """
    if cycle_weeks is None:
        cycle_weeks = detect_cycle_weeks(xlsx_path)
    cycle_days = cycle_weeks * 7

    data = Path(xlsx_path).read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="secuencia")

    sequences = {}
    max_row = max((r for r, c in cells), default=0)

    for r in range(max_row + 1):
        name_val = cells.get((r, 1))  # Col B
        if not isinstance(name_val, str) or not name_val.strip():
            continue
        name_clean = _normalize_name(name_val)
        # Descartar cabeceras y totales
        if any(term in name_clean for term in ("APELLIDOS", "NOMBRE", "SERVICIO", "TOTAL", "PUESTOS", "SEMANA")):
            continue

        # Extraer las marcas del ciclo (Col C en adelante)
        marks = []
        unknown = set()
        for c in range(2, 2 + cycle_days):
            raw = str(cells.get((r, c)) or "").strip().upper()
            if raw in ("M", "T", "D", "R"):
                marks.append(raw)
            else:
                if raw:
                    unknown.add(raw)
                marks.append("D")
        if unknown and warnings is not None:
            warnings.append(
                f"Secuencia de {name_clean}: marcas desconocidas {sorted(unknown)} tratadas como descanso."
            )

        if any(m in ("M", "T") for m in marks):
            if name_clean in sequences and warnings is not None:
                warnings.append(f"Nombre duplicado en secuencias: {name_clean} (se usa la última fila).")
            sequences[name_clean] = tuple(marks)

    return sequences


def load_skills_and_preferences(
    xlsx_path: str | Path,
    positions: tuple[str, ...] | None = None
) -> dict[str, dict]:
    """Carga los conocimientos, vetos y ranking de preferencias desde la plantilla.

    Conocimientos: 'X' = sabe el puesto; celda vacía = NO lo sabe. 'V' = veto expreso.
    Ranking: número 1..k en las columnas de preferencias; los puestos sabidos sin número
    se colocan a continuación del último, conservando el orden del catálogo.
    Soporta columna 'BAJA' (si está marcada con 'X', baja=True).
    Devuelve: {nombre_normalizado: {"skills": set(), "vetoes": set(), "ranking": dict(), "baja": bool}}.
    """
    path = Path(xlsx_path)
    if positions is None:
        positions = parse_positions_from_header(path)

    data = path.read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="Trabajadores")

    profiles = {}
    max_row = max((r for r, c in cells), default=0)
    max_col = max((c for r, c in cells), default=0)

    # Identificar columnas en la fila de cabecera (fila 3 de Excel, índice 2)
    baja_col = next((c for c in range(max_col + 1) if "BAJA" in str(cells.get((2, c)) or "").upper()), None)

    pos_cols = [c for c in range(max_col + 1) if str(cells.get((2, c)) or "").strip().upper() in positions]
    k = len(positions)
    if len(pos_cols) >= 2 * k:
        skills_cols = {str(cells.get((2, c))).strip().upper(): c for c in pos_cols[:k]}
        prefs_cols = {str(cells.get((2, c))).strip().upper(): c for c in pos_cols[k:2 * k]}
    else:
        offset = 1 if baja_col == 2 else 0
        skills_cols = {pos: 2 + offset + idx for idx, pos in enumerate(positions)}
        prefs_cols = {pos: 2 + offset + k + idx for idx, pos in enumerate(positions)}

    for r in range(max_row + 1):
        name_val = cells.get((r, 1))  # Col B
        if not isinstance(name_val, str) or not name_val.strip():
            continue
        name_clean = _normalize_name(name_val)
        if any(term in name_clean for term in ("APELLIDOS", "NOMBRE", "FICHA")):
            continue

        # Verificar si está de BAJA
        is_baja = False
        if baja_col is not None:
            val_baja = str(cells.get((r, baja_col)) or "").strip().upper()
            is_baja = val_baja in ("X", "S", "SI", "1", "TRUE", "V")

        skills = set()
        vetoes = set()
        ranking = {}

        # Columnas de Conocimientos: vacío = no lo sabe
        for pos, c_idx in skills_cols.items():
            val = str(cells.get((r, c_idx)) or "").strip().upper()
            if val == "X":
                skills.add(pos)
            elif val == "V":
                vetoes.add(pos)

        # Columnas de Preferencias
        for pos, c_idx in prefs_cols.items():
            val = cells.get((r, c_idx))
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                ranking[pos] = int(val)

        # Puestos sabidos sin número de preferencia: a continuación del último ranking
        next_rank = max(ranking.values(), default=0)
        for pos in positions:
            if pos in skills and pos not in ranking:
                next_rank += 1
                ranking[pos] = next_rank

        profiles[name_clean] = {
            "skills": frozenset(skills - vetoes),
            "vetoes": frozenset(vetoes),
            "ranking": ranking,
            "baja": is_baja,
        }

    return profiles


def load_all_workers_with_report(
    sequences_path: str | Path,
    preferences_path: str | Path,
    positions: tuple[str, ...] | None = None,
    cycle_weeks: int | None = None,
) -> tuple[tuple[WorkerProfile, ...], LoadReport]:
    """Combina secuencias y ficha en perfiles inmutables e informa de lo que requiere revisión."""
    if positions is None:
        positions = parse_positions_from_header(preferences_path)

    warnings: list[str] = []
    seqs = load_sequences(sequences_path, warnings, cycle_weeks=cycle_weeks)
    prefs = load_skills_and_preferences(preferences_path, positions=positions)

    workers = []
    bajas = []
    sin_ficha = []
    sin_conocimientos = []

    for name, seq in seqs.items():
        p_data = prefs.get(name)
        if not p_data:
            # Coincidencia flexible solo por reordenación de palabras del nombre
            for p_name, data in prefs.items():
                if set(name.split()) == set(p_name.split()):
                    p_data = data
                    break

        if p_data is None:
            # No se inventan conocimientos: sin ficha no sabe ningún puesto
            sin_ficha.append(name)
            p_data = {"skills": frozenset(), "vetoes": frozenset(), "ranking": {}, "baja": False}
        elif not p_data["skills"]:
            sin_conocimientos.append(name)

        if p_data.get("baja", False):
            bajas.append(name)
            continue

        # ID estable derivado del nombre normalizado (no del orden en el Excel)
        wid = "w_" + sha1(name.encode("utf-8")).hexdigest()[:8]
        workers.append(WorkerProfile(
            id=wid,
            name=name,
            skills=p_data["skills"],
            vetoes=p_data["vetoes"],
            ranking=dict(p_data["ranking"]),
            sequence=seq,
        ))

    if sin_ficha:
        warnings.append(
            f"{len(sin_ficha)} trabajador(es) sin ficha de conocimientos (no cubrirán ningún puesto): "
            + ", ".join(sin_ficha)
        )
    if sin_conocimientos:
        warnings.append(
            f"{len(sin_conocimientos)} trabajador(es) sin ningún conocimiento marcado: "
            + ", ".join(sin_conocimientos)
        )
    ids = [w.id for w in workers]
    if len(set(ids)) != len(ids):
        warnings.append("Colisión de identificadores de trabajador: revise nombres duplicados.")

    report = LoadReport(
        warnings=tuple(warnings),
        bajas=tuple(bajas),
        sin_ficha=tuple(sin_ficha),
        sin_conocimientos=tuple(sin_conocimientos),
    )
    return tuple(workers), report


def load_all_workers(
    sequences_path: str | Path,
    preferences_path: str | Path,
    positions: tuple[str, ...] | None = None,
    cycle_weeks: int | None = None,
) -> tuple[WorkerProfile, ...]:
    """Combina secuencias de turnos y ficha de conocimientos en perfiles completos e inmutables."""
    return load_all_workers_with_report(sequences_path, preferences_path, positions=positions, cycle_weeks=cycle_weeks)[0]


# --- CICLO Y PROYECCIÓN MENSUAL ---

CYCLE_LENGTH = 168  # 24 semanas


def cycle_day_matches_weekday(year: int, month: int, start_cycle_day: int, cycle_length: int = CYCLE_LENGTH) -> bool:
    """El ciclo empieza en lunes: el día de ciclo c cae en weekday (c-1) % 7.

    Con esta comprobación se evita proyectar una secuencia desfasada respecto al calendario.
    """
    return 1 <= start_cycle_day <= cycle_length and (start_cycle_day - 1) % 7 == date(year, month, 1).weekday()


def valid_cycle_days(year: int, month: int, cycle_length: int = CYCLE_LENGTH) -> list[int]:
    """Días de ciclo (1..cycle_length) compatibles con el día de la semana real del día 1 del mes."""
    weekday = date(year, month, 1).weekday()
    return [d for d in range(1, cycle_length + 1) if (d - 1) % 7 == weekday]


DEFAULT_REFERENCE_DATE = date(2026, 11, 1)
DEFAULT_REFERENCE_CYCLE_DAY = 63


def suggest_cycle_day(
    reference_or_year: date | int = DEFAULT_REFERENCE_DATE,
    reference_cycle_day_or_month: int = DEFAULT_REFERENCE_CYCLE_DAY,
    year: int | None = None,
    month: int | None = None,
    cycle_length: int = CYCLE_LENGTH,
) -> int:
    """Día de ciclo del día 1 del mes pedido.

    Admite dos formas de llamada:
    - suggest_cycle_day(year, month, cycle_length=...): usa la referencia oficial por defecto (1-nov-2026 = día 63).
    - suggest_cycle_day(reference, reference_cycle_day, year, month, cycle_length=...): calcula desde una referencia específica.
    """
    if isinstance(reference_or_year, date):
        ref_date = reference_or_year
        ref_cycle = reference_cycle_day_or_month
        target_year = year if year is not None else ref_date.year
        target_month = month if month is not None else ref_date.month
    else:
        ref_date = DEFAULT_REFERENCE_DATE
        ref_cycle = DEFAULT_REFERENCE_CYCLE_DAY
        target_year = int(reference_or_year)
        target_month = int(reference_cycle_day_or_month)

    offset = (date(target_year, target_month, 1) - ref_date).days
    ref_c = (ref_cycle - 1) % cycle_length + 1
    return (ref_c - 1 + offset) % cycle_length + 1


def project_month(
    year: int,
    month: int,
    start_cycle_day: int,
    workers: tuple[WorkerProfile, ...],
    positions: tuple[str, ...] | None = None,
    position_rules: dict[str, dict] | None = None,
    cycle_length: int | None = None,
) -> tuple[DaySchedule, ...]:
    """Genera el calendario diario proyectado del mes solicitado a partir del día del ciclo asignado al día 1."""
    if cycle_length is None:
        cycle_length = len(workers[0].sequence) if workers and workers[0].sequence else CYCLE_LENGTH

    if not (1 <= start_cycle_day <= cycle_length):
        raise ValueError(f"El día del ciclo debe estar entre 1 y {cycle_length} (recibido: {start_cycle_day})")
    if not cycle_day_matches_weekday(year, month, start_cycle_day, cycle_length=cycle_length):
        raise ValueError(
            f"El día de ciclo {start_cycle_day} no coincide con el día de la semana del 1/{month}/{year}: "
            f"el ciclo empieza en lunes. Valores compatibles: {valid_cycle_days(year, month, cycle_length)}"
        )

    active_positions = positions if positions is not None else POSITIONS
    active_rules = position_rules if position_rules is not None else POSITION_RULES

    start_date = date(year, month, 1)
    # Último día del mes
    if month == 12:
        end_date = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(year, month + 1, 1) - timedelta(days=1)

    schedules = []
    curr = start_date
    while curr <= end_date:
        # Calcular el día de ciclo (1 a cycle_length) a partir del día 1 del mes
        day_offset = curr.day - 1
        cycle_day = ((start_cycle_day - 1 + day_offset) % cycle_length) + 1  # 1-indexed
        cycle_idx = cycle_day - 1          # 0-indexed en la secuencia

        is_workday_flag = is_workday(curr)

        wm = []
        wt = []
        wr = []

        if is_workday_flag:
            for w in workers:
                seq_idx = cycle_idx % len(w.sequence) if w.sequence else 0
                mark = w.sequence[seq_idx]
                if mark == "M":
                    wm.append(w.id)
                elif mark == "T":
                    wt.append(w.id)
                elif mark == "R":
                    wr.append(w.id)

        # Dotaciones objetivo y mínimas para el día
        targets_m = {}
        targets_t = {}
        if is_workday_flag:
            wd = curr.weekday()
            for pos in active_positions:
                tm, tt, mm, mt = get_position_targets(pos, wd, active_rules)
                targets_m[pos] = (tm, mm)
                targets_t[pos] = (tt, mt)

        schedules.append(DaySchedule(
            day=curr,
            cycle_day=cycle_day,
            weekday=curr.weekday(),
            is_workday=is_workday_flag,
            workers_m=tuple(wm),
            workers_t=tuple(wt),
            workers_r=tuple(wr),
            targets_m=targets_m,
            targets_t=targets_t,
        ))
        curr += timedelta(days=1)

    return tuple(schedules)


def load_sequence_records(xlsx_path: str | Path, cycle_weeks: int | None = None) -> tuple[list[dict], int]:
    """Carga los registros detallados de secuencias de turnos con filas originales y totales."""
    p = Path(xlsx_path)
    if not p.exists():
        return [], 24

    if cycle_weeks is None:
        cycle_weeks = detect_cycle_weeks(p)
    cycle_days = cycle_weeks * 7

    data = p.read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="secuencia")
    max_r = max((r for r, c in cells), default=0)

    records = []
    for r in range(max_r + 1):
        name_val = cells.get((r, 1))
        if not isinstance(name_val, str) or not name_val.strip():
            continue
        name_clean = _normalize_name(name_val)
        if any(term in name_clean for term in ("APELLIDOS", "NOMBRE", "SERVICIO", "TOTAL", "PUESTOS", "SEMANA")):
            continue
        num_val = cells.get((r, 0)) or len(records) + 1
        seq = []
        for c in range(2, 2 + cycle_days):
            raw = str(cells.get((r, c)) or "").strip().upper()
            if raw in ("M", "T", "D", "R"):
                seq.append(raw)
            else:
                seq.append("D")
        m_cnt = seq.count("M")
        t_cnt = seq.count("T")
        d_cnt = seq.count("D")
        r_cnt = seq.count("R")
        records.append({
            "row": r,
            "num": str(num_val),
            "name": name_val.strip(),
            "sequence": seq,
            "m": m_cnt,
            "t": t_cnt,
            "d": d_cnt,
            "r": r_cnt,
            "total": len(seq),
        })

    return records, cycle_weeks


def save_sequence_records(xlsx_path: str | Path, records: list[dict], cycle_weeks: int | None = None) -> None:
    """Guarda modificaciones de marcas de turnos en el archivo Excel de secuencias."""
    p = Path(xlsx_path)
    if not p.exists():
        raise FileNotFoundError(f"No se encuentra el archivo de secuencias: {xlsx_path}")

    # Backup de seguridad antes de escribir
    bak = p.with_suffix(p.suffix + ".bak")
    shutil.copy2(p, bak)

    if cycle_weeks is None:
        cycle_weeks = detect_cycle_weeks(p)
    cycle_days = cycle_weeks * 7

    data = p.read_bytes()
    with zipfile.ZipFile(io.BytesIO(data), "r") as zin:
        manifest = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", MAIN)
    tree = ET.fromstring(manifest["xl/worksheets/sheet1.xml"])

    # Mapa de records por número de fila XML (1-indexed)
    rec_by_row = {r["row"] + 1: r for r in records if "row" in r}

    for row_el in tree.findall(f".//{{{MAIN}}}row"):
        r_num = int(row_el.attrib.get("r", 0))
        if r_num not in rec_by_row:
            continue
        rec = rec_by_row[r_num]
        seq = rec.get("sequence", [])

        # Existentes celdas en la fila indexadas por columna (0-indexed)
        existing_cells = {}
        for c_el in list(row_el.findall(f"{{{MAIN}}}c")):
            ref = c_el.attrib.get("r", "")
            c_idx = _col_idx_from_ref(ref)
            existing_cells[c_idx] = c_el

        # Actualizar celdas de marcas (c_idx = 2 .. 2 + cycle_days - 1)
        for i in range(min(len(seq), cycle_days)):
            c_idx = 2 + i
            mark = seq[i]
            col_let = _col_letter(c_idx)
            cell_ref = f"{col_let}{r_num}"

            if c_idx in existing_cells:
                c_el = existing_cells[c_idx]
            else:
                c_el = ET.SubElement(row_el, f"{{{MAIN}}}c", {"r": cell_ref, "s": "30", "t": "inlineStr"})
                existing_cells[c_idx] = c_el

            for ch in list(c_el):
                c_el.remove(ch)
            c_el.set("t", "inlineStr")
            if "s" not in c_el.attrib:
                c_el.set("s", "30")
            is_el = ET.SubElement(c_el, f"{{{MAIN}}}is")
            t_el = ET.SubElement(is_el, f"{{{MAIN}}}t")
            t_el.text = mark

        # Actualizar columnas de resumen: M, T, D, R, Total
        m_cnt = seq[:cycle_days].count("M")
        t_cnt = seq[:cycle_days].count("T")
        d_cnt = seq[:cycle_days].count("D")
        r_cnt = seq[:cycle_days].count("R")
        tot_cnt = min(len(seq), cycle_days)

        summary_vals = [m_cnt, t_cnt, d_cnt, r_cnt, tot_cnt]
        for s_idx, val in enumerate(summary_vals):
            c_idx = 2 + cycle_days + s_idx
            col_let = _col_letter(c_idx)
            cell_ref = f"{col_let}{r_num}"

            if c_idx in existing_cells:
                c_el = existing_cells[c_idx]
            else:
                c_el = ET.SubElement(row_el, f"{{{MAIN}}}c", {"r": cell_ref, "s": "33", "t": "n"})
                existing_cells[c_idx] = c_el

            for ch in list(c_el):
                c_el.remove(ch)
            c_el.set("t", "n")
            if "s" not in c_el.attrib:
                c_el.set("s", "33")
            v_el = ET.SubElement(c_el, f"{{{MAIN}}}v")
            v_el.text = str(val)

        # Ordenar celdas por índice de columna para conformidad OpenXML estricta
        sorted_cells = sorted(row_el.findall(f"{{{MAIN}}}c"), key=lambda c: _col_idx_from_ref(c.attrib.get("r", "")))
        for c in list(row_el.findall(f"{{{MAIN}}}c")):
            row_el.remove(c)
        for c in sorted_cells:
            row_el.append(c)

    manifest["xl/worksheets/sheet1.xml"] = ET.tostring(tree, encoding="utf-8", xml_declaration=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for fname, fcontent in manifest.items():
            zout.writestr(fname, fcontent)

    p.write_bytes(buf.getvalue())
