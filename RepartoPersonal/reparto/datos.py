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
from pathlib import Path
import posixpath
import re
from xml.etree.ElementTree import fromstring
from zipfile import ZipFile

# Catálogo oficial de puestos y dotaciones base en laborables
POSITIONS = (
    "EC", "MG", "PU", "DR", "UR", "A",
    "RM1", "RM2", "RM3", "TC1", "TC2", "TC3", "TCU",
    "TE", "RT", "O"
)

# Dotaciones objetivo en días laborables estándar
# Formato: (Mañana_Objetivo, Tarde_Objetivo, Mañana_Minimo, Tarde_Minimo)
POSITION_RULES = {
    "EC":  {"target_m": 4, "target_t": 2, "min_m": 3, "min_t": 1, "priority": 1},  # Reducción 1: EC
    "MG":  {"target_m": 2, "target_t": 1, "min_m": 2, "min_t": 1, "priority": 0},  # Estricto
    "PU":  {"target_m": 1, "target_t": 0, "min_m": 1, "min_t": 0, "priority": 0},  # Estricto mañana
    "DR":  {"target_m": 3, "target_t": 3, "min_m": 3, "min_t": 2, "priority": 2},  # Reducción 2: DR tarde
    "UR":  {"target_m": 2, "target_t": 2, "min_m": 2, "min_t": 2, "priority": 0},  # Estricto
    "A":   {"target_m": 2, "target_t": 0, "min_m": 2, "min_t": 0, "priority": 0},  # Estricto mañana
    "RM1": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "RM2": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "RM3": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "TC1": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "TC2": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "TC3": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "TCU": {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "TE":  {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    "RT":  {"target_m": 1, "target_t": 1, "min_m": 1, "min_t": 1, "priority": 0},
    # O (Quirófano/Otros): Se modula según día de semana (Martes y Jueves mañana; Viernes tarde)
    "O":   {"target_m": 0, "target_t": 0, "min_m": 0, "min_t": 0, "priority": 0},
}


def get_position_targets(pos: str, weekday: int) -> tuple[int, int, int, int]:
    """Devuelve (target_m, target_t, min_m, min_t) para un puesto y día de semana (0=Lunes..4=Viernes)."""
    if pos == "O":
        # O abre Martes(1) y Jueves(3) por la mañana; Viernes(4) por la tarde
        m = 1 if weekday in (1, 3) else 0
        t = 1 if weekday == 4 else 0
        return m, t, m, t
    r = POSITION_RULES[pos]
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


def _normalize_name(raw: str) -> str:
    return " ".join(raw.split()).upper()


def load_sequences(xlsx_path: str | Path, warnings: list[str] | None = None) -> dict[str, tuple[str, ...]]:
    """Carga las secuencias de turnos desde el Excel limpio o original.

    Devuelve: {nombre_normalizado: tuple(168 marcas 'M', 'T', 'D', 'R')}.
    Celda vacía = descanso. Cualquier otra marca desconocida se trata como descanso y se
    anota en ``warnings`` (si se proporciona la lista).
    """
    data = Path(xlsx_path).read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="secuencia")

    sequences = {}
    max_row = max(r for r, c in cells)

    for r in range(max_row + 1):
        name_val = cells.get((r, 1))  # Col B
        if not isinstance(name_val, str) or not name_val.strip():
            continue
        name_clean = _normalize_name(name_val)
        # Descartar cabeceras y totales
        if any(term in name_clean for term in ("APELLIDOS", "NOMBRE", "SERVICIO", "TOTAL", "PUESTOS", "SEMANA")):
            continue

        # Extraer las 168 marcas (Col C a FO -> indices 2 a 169)
        marks = []
        unknown = set()
        for c in range(2, 170):
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


def load_skills_and_preferences(xlsx_path: str | Path) -> dict[str, dict]:
    """Carga los conocimientos, vetos y ranking de preferencias desde la plantilla.

    Conocimientos: 'X' = sabe el puesto; celda vacía = NO lo sabe. 'V' = veto expreso.
    Ranking: número 1..k en las columnas de preferencias; los puestos sabidos sin número
    se colocan a continuación del último, conservando el orden del catálogo.
    Soporta columna 'BAJA' (si está marcada con 'X', baja=True).
    Devuelve: {nombre_normalizado: {"skills": set(), "vetoes": set(), "ranking": dict(), "baja": bool}}.
    """
    data = Path(xlsx_path).read_bytes()
    _, cells = _parse_xlsx_cells(data, sheet_name_contains="Trabajadores")

    profiles = {}
    max_row = max(r for r, c in cells)
    max_col = max(c for r, c in cells)

    # Identificar columnas en la fila de cabecera (fila 3 de Excel, índice 2)
    baja_col = next((c for c in range(max_col + 1) if "BAJA" in str(cells.get((2, c)) or "").upper()), None)

    pos_cols = [c for c in range(max_col + 1) if str(cells.get((2, c)) or "").strip().upper() in POSITIONS]
    if len(pos_cols) >= 32:
        skills_cols = {str(cells.get((2, c))).strip().upper(): c for c in pos_cols[:16]}
        prefs_cols = {str(cells.get((2, c))).strip().upper(): c for c in pos_cols[16:32]}
    else:
        offset = 1 if baja_col == 2 else 0
        skills_cols = {pos: 2 + offset + idx for idx, pos in enumerate(POSITIONS)}
        prefs_cols = {pos: 18 + offset + idx for idx, pos in enumerate(POSITIONS)}

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
        for pos in POSITIONS:
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
    sequences_path: str | Path, preferences_path: str | Path
) -> tuple[tuple[WorkerProfile, ...], LoadReport]:
    """Combina secuencias y ficha en perfiles inmutables e informa de lo que requiere revisión."""
    warnings: list[str] = []
    seqs = load_sequences(sequences_path, warnings)
    prefs = load_skills_and_preferences(preferences_path)

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


def load_all_workers(sequences_path: str | Path, preferences_path: str | Path) -> tuple[WorkerProfile, ...]:
    """Combina secuencias de turnos y ficha de conocimientos en perfiles completos e inmutables."""
    return load_all_workers_with_report(sequences_path, preferences_path)[0]


# --- CICLO Y PROYECCIÓN MENSUAL ---

CYCLE_LENGTH = 168  # 24 semanas


def cycle_day_matches_weekday(year: int, month: int, start_cycle_day: int) -> bool:
    """El ciclo empieza en lunes: el día de ciclo c cae en weekday (c-1) % 7.

    Con esta comprobación se evita proyectar una secuencia desfasada respecto al calendario.
    """
    return (start_cycle_day - 1) % 7 == date(year, month, 1).weekday()


def valid_cycle_days(year: int, month: int) -> list[int]:
    """Días de ciclo (1..168) compatibles con el día de la semana real del día 1 del mes."""
    weekday = date(year, month, 1).weekday()
    return [d for d in range(1, CYCLE_LENGTH + 1) if (d - 1) % 7 == weekday]


def suggest_cycle_day(reference: date, reference_cycle_day: int, year: int, month: int) -> int:
    """Día de ciclo del día 1 del mes pedido, a partir de una fecha de referencia conocida."""
    offset = (date(year, month, 1) - reference).days
    return (reference_cycle_day - 1 + offset) % CYCLE_LENGTH + 1


def month_warnings(year: int, month: int) -> list[str]:
    """Avisos del calendario del mes (p. ej. festivos no verificados)."""
    _, verified = holidays_for_year(year)
    if verified:
        return []
    return [f"Los festivos de {year} están calculados por regla y no verificados: revíselos."]


def project_month(
    year: int,
    month: int,
    start_cycle_day: int,
    workers: tuple[WorkerProfile, ...]
) -> tuple[DaySchedule, ...]:
    """Genera el calendario diario proyectado del mes solicitado a partir del día del ciclo asignado al día 1."""
    if not (1 <= start_cycle_day <= CYCLE_LENGTH):
        raise ValueError(f"El día del ciclo debe estar entre 1 y {CYCLE_LENGTH} (recibido: {start_cycle_day})")
    if not cycle_day_matches_weekday(year, month, start_cycle_day):
        raise ValueError(
            f"El día de ciclo {start_cycle_day} no coincide con el día de la semana del 1/{month}/{year}: "
            f"el ciclo empieza en lunes. Valores compatibles: {valid_cycle_days(year, month)}"
        )

    start_date = date(year, month, 1)
    # Último día del mes
    if month == 12:
        end_date = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = date(year, month + 1, 1) - timedelta(days=1)

    schedules = []
    curr = start_date
    while curr <= end_date:
        # Calcular el día de ciclo (1 a 168) a partir del día 1 del mes
        day_offset = curr.day - 1
        cycle_day = ((start_cycle_day - 1 + day_offset) % CYCLE_LENGTH) + 1  # 1-indexed
        cycle_idx = cycle_day - 1          # 0-indexed en la secuencia

        is_workday_flag = is_workday(curr)

        wm = []
        wt = []
        wr = []

        if is_workday_flag:
            for w in workers:
                mark = w.sequence[cycle_idx]
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
            for pos in POSITIONS:
                tm, tt, mm, mt = get_position_targets(pos, wd)
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
