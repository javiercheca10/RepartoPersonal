"""Exportadores oficiales de cuadrante mensual a Excel (.xlsx) y PDF (.pdf).

Genera:
1. Excel (.xlsx):
   - Hoja 1: «Cuadrante Mensual» con trabajadores en filas, días del mes en columnas (1..31).
     Encabezado con número de día y día de la semana (L, M, X, J, V, S, D).
     Fines de semana y festivos oficiales sombreados en gris y protegidos.
     Marcas claras de asignación:
       - 'EC', 'MG', 'DR'... para mañanas.
       - 'EC (T)', 'DR (T)'... para tardes.
       - '.' para Libre.
       - 'D' para descanso laboral del ciclo.
   - Hoja 2: «Resumen y Esfuerzo» con los recuentos exactos por profesional:
     días trabajados, veces en cada puesto, días Libre, puntos totales y media de esfuerzo.
2. PDF (.pdf):
   - Generación directa maquetada mediante LibreOffice o conversor nativo en formato A4/A3 apaisado.
"""

from collections import defaultdict
from datetime import date, timedelta
from io import BytesIO
from pathlib import Path
import re
import subprocess
from xml.etree.ElementTree import Element, SubElement, tostring
from zipfile import ZIP_DEFLATED, ZipFile

from reparto.datos import DaySchedule, POSITIONS, WorkerProfile
from reparto.solver import MonthResult

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
INVALID_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")

MONTH_NAMES_ES = [
    "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]
DAYS_ES = ["L", "Ma", "Mi", "J", "V", "S", "Do"]

# Estilos XML enriquecidos basados exactamente en el formato oficial PUESTOS NUEVOS
STYLES_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="{MAIN}">
 <numFmts count="2">
  <numFmt numFmtId="164" formatCode="0"/>
  <numFmt numFmtId="165" formatCode="0.00"/>
 </numFmts>
 <fonts count="13">
  <!-- 0: Normal 10pt Segoe UI -->
  <font><sz val="10"/><name val="Segoe UI"/><color rgb="FF000000"/></font>
  <!-- 1: Titulo grande negrita 15pt -->
  <font><b/><sz val="15"/><name val="Segoe UI"/><color rgb="FF000000"/></font>
  <!-- 2: Cabecera dia semana negro negrita 13pt -->
  <font><b/><sz val="13"/><name val="Segoe UI"/><color rgb="FF000000"/></font>
  <!-- 3: Cabecera domingo 'Do' rojo negrita 13pt -->
  <font><b/><sz val="13"/><name val="Segoe UI"/><color rgb="FFFF0000"/></font>
  <!-- 4: Cabecera numero dia negro negrita 13pt -->
  <font><b/><sz val="13"/><name val="Segoe UI"/><color rgb="FF000000"/></font>
  <!-- 5: Nombre trabajador negrita 11pt -->
  <font><b/><sz val="11"/><name val="Segoe UI"/><color rgb="FF000000"/></font>
  <!-- 6: Puesto Mañana (Azul Marino Negrita 13pt #191BA4) -->
  <font><b/><sz val="13"/><name val="Segoe UI"/><color rgb="FF191BA4"/></font>
  <!-- 7: Puesto Tarde (Verde Negrita 13pt #00A933) -->
  <font><b/><sz val="13"/><name val="Segoe UI"/><color rgb="FF00A933"/></font>
  <!-- 8: Cabecera Blanca Negrita 9pt (hoja resumen) -->
  <font><b/><sz val="9"/><name val="Segoe UI"/><color rgb="FFFFFFFF"/></font>
  <!-- 9: Resumen negrita 9pt -->
  <font><b/><sz val="9"/><name val="Segoe UI"/><color rgb="FF0F172A"/></font>
  <!-- 10: Subtitulo cursiva 9pt -->
  <font><i/><sz val="9"/><name val="Segoe UI"/><color rgb="FF475569"/></font>
  <!-- 11: Descanso D suave -->
  <font><sz val="8.5"/><name val="Segoe UI"/><color rgb="FF94A3B8"/></font>
  <!-- 12: Vacante / alerta roja negrita 10pt -->
  <font><b/><sz val="10"/><name val="Segoe UI"/><color rgb="FFDC2626"/></font>
 </fonts>
 <fills count="9">
  <!-- 0: None -->
  <fill><patternFill patternType="none"/></fill>
  <!-- 1: Gray125 -->
  <fill><patternFill patternType="gray125"/></fill>
  <!-- 2: Blanco solido #FFFFFF -->
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
  <!-- 3: Cabecera dias lila pastel #CCCCFF -->
  <fill><patternFill patternType="solid"><fgColor rgb="FFCCCCFF"/><bgColor indexed="64"/></patternFill></fill>
  <!-- 4: Cabecera numeros amarillo pastel #FFFFCC -->
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFFFCC"/><bgColor indexed="64"/></patternFill></fill>
  <!-- 5: Resumen azul marino #1E3A8A -->
  <fill><patternFill patternType="solid"><fgColor rgb="FF1E3A8A"/><bgColor indexed="64"/></patternFill></fill>
  <!-- 6: Resumen fila alterna #F8FAFC -->
  <fill><patternFill patternType="solid"><fgColor rgb="FFF8FAFC"/><bgColor indexed="64"/></patternFill></fill>
  <!-- 7: Resumen cabecera verde #059669 -->
  <fill><patternFill patternType="solid"><fgColor rgb="FF059669"/><bgColor indexed="64"/></patternFill></fill>
  <!-- 8: Fondo alerta suave #FEF2F2 -->
  <fill><patternFill patternType="solid"><fgColor rgb="FFFEF2F2"/><bgColor indexed="64"/></patternFill></fill>
 </fills>
 <borders count="3">
  <!-- 0: Sin borde -->
  <border><left/><right/><top/><bottom/></border>
  <!-- 1: Borde fino cuadricula #CBD5E1 -->
  <border>
   <left style="thin"><color rgb="FFCBD5E1"/></left>
   <right style="thin"><color rgb="FFCBD5E1"/></right>
   <top style="thin"><color rgb="FFCBD5E1"/></top>
   <bottom style="thin"><color rgb="FFCBD5E1"/></bottom>
  </border>
  <!-- 2: Borde medio resumen -->
  <border>
   <left style="thin"><color rgb="FF94A3B8"/></left>
   <right style="thin"><color rgb="FF94A3B8"/></right>
   <top style="medium"><color rgb="FF1E3A8A"/></top>
   <bottom style="medium"><color rgb="FF1E3A8A"/></bottom>
  </border>
 </borders>
 <cellStyleXfs count="1">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
 </cellStyleXfs>
 <cellXfs count="21">
  <!-- 0: Texto normal sin borde -->
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 1: Titulo principal (Font 1 negrita 15pt, sin borde) -->
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyFont="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 2: Cabecera dia laborable/sabado (Font 2 negrita 13pt negro, Fill 3 #CCCCFF, Border 1, Center) -->
  <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 3: Cabecera domingo 'Do' (Font 3 negrita 13pt rojo #FF0000, Fill 3 #CCCCFF, Border 1, Center) -->
  <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 4: Cabecera numero de dia (Font 4 negrita 13pt negro, Fill 4 #FFFFCC, Border 1, Center) -->
  <xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 5: Nombre de trabajador (Font 5 negrita 11pt negro, Fill 2 blanco, Border 1, Left) -->
  <xf numFmtId="0" fontId="5" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 6: Puesto Mañana (Font 6 azul #191BA4 negrita 13pt, Fill 2 blanco, Border 1, Center) -->
  <xf numFmtId="0" fontId="6" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 7: Puesto Tarde (Font 7 verde #00A933 negrita 13pt, Fill 2 blanco, Border 1, Center) -->
  <xf numFmtId="0" fontId="7" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 8: Celda vacia / descanso con borde fino (Font 0, Fill 2 blanco, Border 1, Center) -->
  <xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 9: Resumen cabecera azul #1E3A8A -->
  <xf numFmtId="0" fontId="8" fillId="5" borderId="2" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center" wrapText="1"/>
  </xf>
  <!-- 10: Resumen texto izquierda con borde -->
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1" applyBorder="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 11: Resumen centrado con borde -->
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 12: Resumen entero con borde -->
  <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1" applyBorder="1" applyNumberFormat="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 13: Resumen decimal con borde -->
  <xf numFmtId="165" fontId="9" fillId="0" borderId="1" xfId="0" applyAlignment="1" applyBorder="1" applyNumberFormat="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 14: Resumen fila alterna texto izquierda -->
  <xf numFmtId="0" fontId="0" fillId="6" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyBorder="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 15: Resumen fila alterna entero -->
  <xf numFmtId="164" fontId="0" fillId="6" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyBorder="1" applyNumberFormat="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 16: Resumen titulo grande -->
  <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyFont="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 17: Resumen subtitulo cursiva -->
  <xf numFmtId="0" fontId="10" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyFont="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 18: Vacante texto etiqueta izquierda roja suave -->
  <xf numFmtId="0" fontId="12" fillId="8" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment vertical="center"/>
  </xf>
  <!-- 19: Vacante celda detalle centrado roja suave -->
  <xf numFmtId="0" fontId="12" fillId="8" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyFont="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
  <!-- 20: Celda neutra blanca sin vacante para fila de vacantes -->
  <xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyAlignment="1" applyFill="1" applyBorder="1">
   <alignment horizontal="center" vertical="center"/>
  </xf>
 </cellXfs>
</styleSheet>"""


def _col_letter(col_idx: int) -> str:
    name = ""
    while col_idx:
        col_idx, remainder = divmod(col_idx - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _create_sheet_xml(rows_data, widths, header_split_row=0, split_col=0, merge_cells=None):
    sheet = Element("worksheet", xmlns=MAIN)
    props = SubElement(sheet, "sheetPr")
    SubElement(props, "pageSetUpPr", fitToPage="1")

    max_cols = len(widths)
    max_rows = len(rows_data)
    SubElement(sheet, "dimension", ref=f"A1:{_col_letter(max_cols)}{max_rows}")

    if header_split_row > 0 or split_col > 0:
        views = SubElement(sheet, "sheetViews")
        view = SubElement(views, "sheetView", workbookViewId="0", showGridLines="1")
        top_left = f"{_col_letter(split_col + 1)}{header_split_row + 1}"
        pane_attrs = {
            "ySplit": str(header_split_row),
            "topLeftCell": top_left,
            "state": "frozen",
        }
        if split_col > 0:
            pane_attrs["xSplit"] = str(split_col)
            pane_attrs["activePane"] = "bottomRight"
        else:
            pane_attrs["activePane"] = "bottomLeft"
        SubElement(view, "pane", **pane_attrs)

    cols_elm = SubElement(sheet, "cols")
    for idx, width in enumerate(widths, 1):
        SubElement(cols_elm, "col", min=str(idx), max=str(idx), width=str(width), customWidth="1")

    sheet_data = SubElement(sheet, "sheetData")
    for row_idx, (height, cells) in enumerate(rows_data, 1):
        row_elm = SubElement(sheet_data, "row", r=str(row_idx), ht=str(height), customHeight="1")

        for col_idx, (val, style_idx) in enumerate(cells, 1):
            if val is None or val == "":
                if style_idx not in (0, 1, 16, 17):
                    SubElement(row_elm, "c", r=f"{_col_letter(col_idx)}{row_idx}", s=str(style_idx))
                continue

            attrs = {"r": f"{_col_letter(col_idx)}{row_idx}", "s": str(style_idx)}
            if isinstance(val, (int, float)):
                c = SubElement(row_elm, "c", **{**attrs, "t": "n"})
                SubElement(c, "v").text = str(val)
            else:
                text = str(val)
                if INVALID_XML.search(text):
                    text = INVALID_XML.sub("", text)
                c = SubElement(row_elm, "c", **{**attrs, "t": "inlineStr"})
                inline = SubElement(c, "is")
                SubElement(inline, "t", {"{http://www.w3.org/XML/1998/namespace}space": "preserve"}).text = text

    if merge_cells:
        mc_elm = SubElement(sheet, "mergeCells", count=str(len(merge_cells)))
        for mc in merge_cells:
            SubElement(mc_elm, "mergeCell", ref=mc)

    SubElement(sheet, "pageMargins", left="0.3", right="0.3", top="0.4", bottom="0.4", header="0.2", footer="0.2")
    SubElement(sheet, "pageSetup", paperSize="9", orientation="landscape", fitToWidth="1", fitToHeight="0")
    return tostring(sheet, encoding="utf-8", xml_declaration=True)


def export_month_to_excel(
    result: MonthResult,
    schedules: tuple[DaySchedule, ...],
    workers: tuple[WorkerProfile, ...],
    output_path: Path,
) -> None:
    """Genera el libro Excel oficial del mes reproduciendo el formato exacto de PUESTOS NUEVOS."""
    month_name = MONTH_NAMES_ES[result.month]
    year_short = str(result.year)[-2:]
    sheet_title = f"{month_name.upper()} {year_short}"
    num_days = len(schedules)

    # Indexar asignaciones por (worker_id, day) -> list of Assignment
    asgs_by_worker_day = defaultdict(list)
    for a in result.assignments:
        asgs_by_worker_day[a.worker_id, a.day].append(a)

    # --- HOJA 1: FORMATO OFICIAL PUESTOS NUEVOS ---
    # Col A (1): margen (ancho 3)
    # Col B (2): Trabajador (ancho 34)
    # Cols C..: Días del mes (ancho 5.6 cada uno)
    sheet1_widths = [3, 34] + [5.6] * num_days
    sheet1_rows = []

    # Fila 1 (Excel Row 1): Título de Planificación y Servicio
    r1_cells = [
        ("", 0),
        ("PLANIFICACIÓN MENSUAL ", 1),
        (f"       CUADRANTE GENERAL DE TURNOS Y PUESTOS         (  {month_name.upper()} -{year_short} )", 1),
    ]
    for _ in range(num_days - 1):
        r1_cells.append(("", 0))
    sheet1_rows.append((28, r1_cells))

    # Fila 2 (Excel Row 2): DEPARTAMENTO y letras de día (L, Ma, Mi, J, V, S, Do)
    r2_cells = [
        ("", 0),
        ("DEPARTAMENTO", 1),
    ]
    for s in schedules:
        day_initial = DAYS_ES[s.weekday]
        style = 3 if s.weekday == 6 else 2  # 3=Rojo Domingo, 2=Negro
        r2_cells.append((day_initial, style))
    sheet1_rows.append((24, r2_cells))

    # Fila 3 (Excel Row 3): PERSONAL y número de día (1..31)
    r3_cells = [
        ("", 0),
        ("PERSONAL", 1),
    ]
    for s in schedules:
        r3_cells.append((s.day.day, 4))  # 4=Fondo amarillo pastel #FFFFCC
    sheet1_rows.append((24, r3_cells))

    # Filas de trabajadores (Excel Row 4..)
    for idx, w in enumerate(workers, 1):
        row_cells = [
            ("", 0),
            (w.name, 5),  # 5=Negrita negro con borde
        ]

        for s in schedules:
            worker_asgs = asgs_by_worker_day.get((w.id, s.day), [])
            if not worker_asgs:
                # Descanso o no laborable -> celda vacía con borde
                row_cells.append(("", 8))
            else:
                parts = []
                # Si es tarde, verde con 't' (e.g. ECt, .t). Si mañana, azul sin 't' (e.g. EC, .)
                is_afternoon = any(a.shift == "T" for a in worker_asgs)
                for a in worker_asgs:
                    if a.shift == "T":
                        mark = ".t" if a.position == "." else f"{a.position}t"
                    else:
                        mark = "." if a.position == "." else a.position
                    parts.append(mark)

                cell_text = "/".join(parts)
                style_cell = 7 if is_afternoon else 6
                row_cells.append((cell_text, style_cell))

        sheet1_rows.append((22, row_cells))

    # Filas de puestos descubiertos (vacantes a revisar a mano)
    if result.vacancies:
        # Fila separadora visual
        sep_cells = [("", 0), ("PUESTOS A REVISAR A MANO", 18)]
        for _ in range(num_days):
            sep_cells.append(("", 18))
        sheet1_rows.append((20, sep_cells))

        # Vacantes Mañana
        vac_m_by_day = defaultdict(list)
        vac_t_by_day = defaultdict(list)
        for v in result.vacancies:
            if v.shift == "M":
                txt = f"{v.position} ({v.missing})" if v.missing > 1 else v.position
                vac_m_by_day[v.day].append(txt)
            else:
                txt = f"{v.position}t ({v.missing})" if v.missing > 1 else f"{v.position}t"
                vac_t_by_day[v.day].append(txt)

        r_vm = [("", 0), ("  Vacantes Mañana", 18)]
        for s in schedules:
            v_list = vac_m_by_day.get(s.day, [])
            if v_list:
                r_vm.append(("/".join(v_list), 19))
            else:
                r_vm.append(("", 20))
        sheet1_rows.append((22, r_vm))

        r_vt = [("", 0), ("  Vacantes Tarde", 18)]
        for s in schedules:
            v_list = vac_t_by_day.get(s.day, [])
            if v_list:
                r_vt.append(("/".join(v_list), 19))
            else:
                r_vt.append(("", 20))
        sheet1_rows.append((22, r_vt))

    sheet1_xml = _create_sheet_xml(sheet1_rows, sheet1_widths, header_split_row=3, split_col=2)

    # --- HOJA 2: RESUMEN Y ESFUERZO DETALLADO ---
    sheet2_widths = [5, 34, 9, 8, 9, 8] + [7] * len(POSITIONS)
    sheet2_rows = []

    sheet2_rows.append((28, [(f"RESUMEN DE ESFUERZO Y DISTRIBUCIÓN — {month_name.upper()} {result.year}", 16)]))
    sheet2_rows.append((20, [("Métricas individuales acumuladas en el mes (Libre = 0 puntos)", 17)]))

    hdr2 = [
        ("N.º", 9),
        ("Apellidos y Nombre", 9),
        ("Días Trab.", 9),
        ("Puntos", 9),
        ("Media/Día", 9),
        ("Libre (.)", 9),
    ]
    for pos in POSITIONS:
        hdr2.append((pos, 9))
    sheet2_rows.append((24, hdr2))

    for idx, w in enumerate(workers, 1):
        is_alt = (idx % 2 == 0)
        s_txt = 14 if is_alt else 10
        s_num = 15 if is_alt else 12
        we = result.worker_efforts.get(w.id)

        if we:
            r_cells = [
                (idx, s_num),
                (w.name, s_txt),
                (we.worked_days, s_num),
                (we.total_points, s_num),
                (we.mean_effort, 13),
                (we.free_days, s_num),
            ]
            for pos in POSITIONS:
                cnt = we.counts_by_position.get(pos, 0)
                r_cells.append((cnt if cnt > 0 else "", s_num))
        else:
            r_cells = [(idx, s_num), (w.name, s_txt), (0, s_num), (0, s_num), (0.0, 13), (0, s_num)]
            for _ in POSITIONS:
                r_cells.append(("", s_num))

        sheet2_rows.append((22, r_cells))

    sheet2_merges = ["A1:F1", "A2:F2"]

    # Añadir sección de vacantes en Hoja 2 si existen
    if result.vacancies:
        sheet2_rows.append((16, []))
        sheet2_rows.append((26, [("DETALLE DE PUESTOS DESCUBIERTOS (A REVISAR A MANO)", 18)]))
        hdr_vac = [
            ("Fecha", 9),
            ("Día", 9),
            ("Turno", 9),
            ("Puesto", 9),
            ("Objetivo", 9),
            ("Mínimo", 9),
            ("Asignados", 9),
            ("Faltan", 9),
            ("Observación", 9),
        ]
        sheet2_rows.append((22, hdr_vac))
        for v in result.vacancies:
            d_name = DAYS_ES[v.day.weekday()]
            turno_txt = "Mañana" if v.shift == "M" else "Tarde"
            obs = "Por debajo del mínimo estricto" if v.below_minimum else "Reducción permitida"
            r_v = [
                (v.day.strftime("%d/%m/%Y"), 11),
                (d_name, 11),
                (turno_txt, 11),
                (v.position, 11),
                (v.target, 12),
                (v.minimum, 12),
                (v.assigned, 12),
                (v.missing, 19),
                (obs, 18 if v.below_minimum else 10),
            ]
            sheet2_rows.append((20, r_v))

    sheet2_xml = _create_sheet_xml(sheet2_rows, sheet2_widths, header_split_row=3, split_col=0, merge_cells=sheet2_merges)

    # Empaquetar XLSX
    sheets_info = [
        (sheet_title, sheet1_xml),
        ("Resumen y Esfuerzo", sheet2_xml),
    ]

    workbook = Element("workbook", xmlns=MAIN, **{"xmlns:r": REL})
    ws_elm = SubElement(workbook, "sheets")
    rels_elm = Element("Relationships", xmlns=PKG)
    types_elm = Element("Types", xmlns=TYPES)

    SubElement(types_elm, "Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    SubElement(types_elm, "Default", Extension="xml", ContentType="application/xml")
    SubElement(types_elm, "Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    SubElement(types_elm, "Override", PartName="/xl/styles.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")

    for idx, (name, _) in enumerate(sheets_info, 1):
        SubElement(ws_elm, "sheet", name=name, sheetId=str(idx), **{"r:id": f"rId{idx}"})
        SubElement(rels_elm, "Relationship", Id=f"rId{idx}", Type=f"{REL}/worksheet", Target=f"worksheets/sheet{idx}.xml")
        SubElement(types_elm, "Override", PartName=f"/xl/worksheets/sheet{idx}.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")

    SubElement(rels_elm, "Relationship", Id="rIdStyles", Type=f"{REL}/styles", Target="styles.xml")

    root_rels = Element("Relationships", xmlns=PKG)
    SubElement(root_rels, "Relationship", Id="rId1", Type=f"{REL}/officeDocument", Target="xl/workbook.xml")

    buffer = BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", tostring(types_elm, encoding="utf-8", xml_declaration=True))
        archive.writestr("_rels/.rels", tostring(root_rels, encoding="utf-8", xml_declaration=True))
        archive.writestr("xl/workbook.xml", tostring(workbook, encoding="utf-8", xml_declaration=True))
        archive.writestr("xl/_rels/workbook.xml.rels", tostring(rels_elm, encoding="utf-8", xml_declaration=True))
        archive.writestr("xl/styles.xml", STYLES_XML.encode("utf-8"))
        for idx, (_, sheet_bytes) in enumerate(sheets_info, 1):
            archive.writestr(f"xl/worksheets/sheet{idx}.xml", sheet_bytes)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(buffer.getvalue())


def export_month_to_pdf(
    result: MonthResult,
    schedules: tuple[DaySchedule, ...],
    workers: tuple[WorkerProfile, ...],
    output_path: Path,
) -> Path:
    """Genera el cuadrante mensual en PDF maquetado para impresión directa.

    Utiliza ReportLab para generar un PDF nativo, autónomo y sin dependencias externas
    (sin requerir LibreOffice instalado en Windows).
    """
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

        month_name = MONTH_NAMES_ES[result.month]
        year_short = str(result.year)[-2:]
        num_days = len(schedules)

        # Indexar asignaciones
        asgs_by_worker_day = defaultdict(list)
        for a in result.assignments:
            asgs_by_worker_day[a.worker_id, a.day].append(a)

        # Configuración de documento apaisado A4
        # Márgenes estrechos para maximizar espacio
        margin = 12
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=landscape(A4),
            leftMargin=margin,
            rightMargin=margin,
            topMargin=14,
            bottomMargin=14,
        )

        elements = []

        # Título
        title_style = ParagraphStyle(
            "TitleStyle",
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#0F172A"),
        )
        subtitle_style = ParagraphStyle(
            "SubtitleStyle",
            fontName="Helvetica",
            fontSize=7,
            leading=9,
            textColor=colors.HexColor("#475569"),
        )

        title_p = Paragraph(
            f"<b>PLANIFICACIÓN GENERAL DE PERSONAL · CUADRANTE DE TURNOS Y PUESTOS</b> — "
            f"({month_name.upper()} {result.year})",
            title_style,
        )
        elements.append(title_p)
        elements.append(Spacer(1, 4))

        # Tabla del cuadrante
        # Col 0: Trabajador (~120 pt). Resto: 30 o 31 días repartidos en el ancho disponible (~700 pt)
        page_width = landscape(A4)[0] - (2 * margin)  # ~817 pt
        col0_width = 115
        day_col_width = (page_width - col0_width) / num_days

        table_data = []

        # Fila 1: Días de la semana (L, Ma, Mi, J, V, S, Do)
        row_weekdays = ["TRABAJADOR / DÍA"]
        for s in schedules:
            row_weekdays.append(DAYS_ES[s.weekday])
        table_data.append(row_weekdays)

        # Fila 2: Números de día (1..31)
        row_numbers = [""]
        for s in schedules:
            row_numbers.append(str(s.day.day))
        table_data.append(row_numbers)

        # Filas de trabajadores
        for w in workers:
            w_row = [w.name]
            for s in schedules:
                worker_asgs = asgs_by_worker_day.get((w.id, s.day), [])
                if not worker_asgs:
                    w_row.append("")
                else:
                    parts = []
                    for a in worker_asgs:
                        if a.shift == "T":
                            mark = ".t" if a.position == "." else f"{a.position}t"
                        else:
                            mark = "." if a.position == "." else a.position
                        parts.append(mark)
                    w_row.append("/".join(parts))
            table_data.append(w_row)

        # Filas de vacantes
        vac_row_indices = []
        if result.vacancies:
            vac_m_by_day = defaultdict(list)
            vac_t_by_day = defaultdict(list)
            for v in result.vacancies:
                if v.shift == "M":
                    txt = f"{v.position}({v.missing})" if v.missing > 1 else v.position
                    vac_m_by_day[v.day].append(txt)
                else:
                    txt = f"{v.position}t({v.missing})" if v.missing > 1 else f"{v.position}t"
                    vac_t_by_day[v.day].append(txt)

            row_vm = ["VACANTES MAÑANA"]
            for s in schedules:
                v_list = vac_m_by_day.get(s.day, [])
                row_vm.append("/".join(v_list) if v_list else "")
            table_data.append(row_vm)
            vac_row_indices.append(len(table_data) - 1)

            row_vt = ["VACANTES TARDE"]
            for s in schedules:
                v_list = vac_t_by_day.get(s.day, [])
                row_vt.append("/".join(v_list) if v_list else "")
            table_data.append(row_vt)
            vac_row_indices.append(len(table_data) - 1)

        # Estilo de la tabla
        t_style = [
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#CCCCFF")),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FFFFCC")),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 5.5),
            ("LEADING", (0, 0), (-1, -1), 6.5),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 1),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
            ("LEFTPADDING", (0, 0), (-1, -1), 1.5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"),
        ]

        # Colorear texto de Domingo en fila 0
        for idx, s in enumerate(schedules, 1):
            if s.weekday == 6:
                t_style.append(("TEXTCOLOR", (idx, 0), (idx, 0), colors.HexColor("#DC2626")))

        # Colorear puestos: Azul mañana (#191BA4), Verde tarde (#00A933)
        for r_idx, w in enumerate(workers, 2):
            for c_idx, s in enumerate(schedules, 1):
                worker_asgs = asgs_by_worker_day.get((w.id, s.day), [])
                if worker_asgs:
                    is_afternoon = any(a.shift == "T" for a in worker_asgs)
                    color = colors.HexColor("#00A933") if is_afternoon else colors.HexColor("#191BA4")
                    t_style.append(("TEXTCOLOR", (c_idx, r_idx), (c_idx, r_idx), color))
                    t_style.append(("FONTNAME", (c_idx, r_idx), (c_idx, r_idx), "Helvetica-Bold"))

        # Colorear filas de vacantes en rojo suave
        for r_idx in vac_row_indices:
            t_style.append(("BACKGROUND", (0, r_idx), (-1, r_idx), colors.HexColor("#FEF2F2")))
            t_style.append(("TEXTCOLOR", (0, r_idx), (-1, r_idx), colors.HexColor("#DC2626")))
            t_style.append(("FONTNAME", (0, r_idx), (-1, r_idx), "Helvetica-Bold"))

        col_widths = [col0_width] + [day_col_width] * num_days
        table = Table(table_data, colWidths=col_widths, repeatRows=2)
        table.setStyle(TableStyle(t_style))
        elements.append(table)

        doc.build(elements)
        return output_path

    except Exception:
        # Fallback a LibreOffice si ReportLab falla
        excel_path = output_path.with_suffix(".xlsx")
        if excel_path.exists():
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(output_path.parent),
                str(excel_path)
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            return output_path
        raise
