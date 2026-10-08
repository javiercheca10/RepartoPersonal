"""Motor de Optimización CP-SAT para el reparto mensual de radiodiagnóstico.

Prioridades de optimización en cascada lexicográfica:
1. Cobertura obligatoria de puestos (minimizar plazas descubiertas).
2. Minimizar reducciones de DR (después de haber reducido EC).
3. Asignación de trabajadores sobrantes a 'Libre (.)' con 0 puntos de esfuerzo.
4. Máxima equidad: Minimizar la peor media individual (cota superior de esfuerzo diario)
   y secundariamente maximizar la satisfacción global de preferencias (menor suma de puntos).

Manejo de Refuerzo ('R'):
- Si un trabajador tiene 'R', el algoritmo decide de forma óptima si entra en turno de
  Mañana o de Tarde para cubrir los déficits de plazas.

Verificador independiente:
- Post-cálculo audita que cada asignación respete conocimientos y vetos, y que
  nadie tenga asignación duplicada.
"""

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date
from fractions import Fraction
from time import perf_counter

from ortools.sat.python import cp_model

from reparto.auditoria import audit_assignments, compute_vacancies, Vacancy
from reparto.datos import DaySchedule, POSITIONS, WorkerProfile


@dataclass(frozen=True)
class Assignment:
    day: date
    shift: str          # 'M' o 'T'
    worker_id: str
    worker_name: str
    position: str       # Código de puesto o '.' (Libre)
    points: int         # Puntos de esfuerzo (0 si Libre, 1..16 según ranking)


@dataclass(frozen=True)
class WorkerEffort:
    worker_id: str
    worker_name: str
    worked_days: int
    total_points: int
    mean_effort: float
    exact_mean: Fraction | None
    free_days: int
    counts_by_position: dict[str, int]


@dataclass(frozen=True)
class MonthResult:
    year: int
    month: int
    is_optimal: bool
    status_name: str                  # 'Óptimo' o 'Factible'
    vacancies_total: int
    vacancies_dr: int
    effort_spread: float
    max_mean_effort: float
    min_mean_effort: float
    assignments: tuple[Assignment, ...]
    vacancies: tuple[Vacancy, ...]    # Puestos sin cubrir a revisar a mano
    worker_efforts: dict[str, WorkerEffort]
    execution_time_seconds: float
    issues: tuple[str, ...]


def solve_month(
    year: int,
    month: int,
    schedules: tuple[DaySchedule, ...],
    workers: tuple[WorkerProfile, ...],
    time_limit_seconds: float = 15.0
) -> MonthResult:
    """Resuelve la asignación mensual óptima para todos los días laborables proyectados."""
    start_time = perf_counter()
    model = cp_model.CpModel()

    workers_by_id = {w.id: w for w in workers}
    workdays = [s for s in schedules if s.is_workday]

    if not workdays:
        return MonthResult(
            year=year,
            month=month,
            is_optimal=True,
            vacancies_total=0,
            vacancies_dr=0,
            effort_spread=0.0,
            max_mean_effort=0.0,
            min_mean_effort=0.0,
            assignments=(),
            worker_efforts={},
            execution_time_seconds=0.0,
            issues=("Mes sin días laborables para reparto.",),
        )

    # Variables de asignación:
    # x[day, shift, worker_id, pos] = 1 si worker se asigna a pos en (day, shift)
    x = {}

    # Variable de turno para trabajadores con refuerzo 'R':
    # r_shift_is_t[day, worker_id] = 1 si 'R' entra de Tarde, 0 si entra de Mañana
    r_shift_is_t = {}

    # Elecciones por (day, shift, worker_id)
    worker_shift_choices = defaultdict(list)

    # Conteos de cobertura por (day, shift, pos): lista de variables bool
    coverage_choices = defaultdict(list)

    for s in workdays:
        d = s.day

        # 1. Definir presencia y opciones de turno para trabajadores fijos de M
        for wid in s.workers_m:
            w = workers_by_id[wid]
            for pos in POSITIONS:
                if pos in w.skills and pos not in w.vetoes:
                    var = model.new_bool_var(f"x_{d}_M_{wid}_{pos}")
                    x[d, "M", wid, pos] = var
                    worker_shift_choices[d, "M", wid].append(var)
                    coverage_choices[d, "M", pos].append(var)
            # Opción Libre (.)
            var_free = model.new_bool_var(f"x_{d}_M_{wid}_FREE")
            x[d, "M", wid, "."] = var_free
            worker_shift_choices[d, "M", wid].append(var_free)
            model.add_exactly_one(worker_shift_choices[d, "M", wid])

        # 2. Definir presencia y opciones de turno para trabajadores fijos de T
        for wid in s.workers_t:
            w = workers_by_id[wid]
            for pos in POSITIONS:
                if pos in w.skills and pos not in w.vetoes:
                    var = model.new_bool_var(f"x_{d}_T_{wid}_{pos}")
                    x[d, "T", wid, pos] = var
                    worker_shift_choices[d, "T", wid].append(var)
                    coverage_choices[d, "T", pos].append(var)
            # Opción Libre (.)
            var_free = model.new_bool_var(f"x_{d}_T_{wid}_FREE")
            x[d, "T", wid, "."] = var_free
            worker_shift_choices[d, "T", wid].append(var_free)
            model.add_exactly_one(worker_shift_choices[d, "T", wid])

        # 3. Definir presencia flexible para trabajadores con Refuerzo 'R'
        for wid in s.workers_r:
            w = workers_by_id[wid]
            is_t = model.new_bool_var(f"r_is_t_{d}_{wid}")
            r_shift_is_t[d, wid] = is_t

            # Opciones si entra de Mañana
            m_choices = []
            for pos in POSITIONS:
                if pos in w.skills and pos not in w.vetoes:
                    var = model.new_bool_var(f"x_{d}_M_{wid}_{pos}")
                    x[d, "M", wid, pos] = var
                    m_choices.append(var)
                    coverage_choices[d, "M", pos].append(var)
            var_free_m = model.new_bool_var(f"x_{d}_M_{wid}_FREE")
            x[d, "M", wid, "."] = var_free_m
            m_choices.append(var_free_m)

            # Opciones si entra de Tarde
            t_choices = []
            for pos in POSITIONS:
                if pos in w.skills and pos not in w.vetoes:
                    var = model.new_bool_var(f"x_{d}_T_{wid}_{pos}")
                    x[d, "T", wid, pos] = var
                    t_choices.append(var)
                    coverage_choices[d, "T", pos].append(var)
            var_free_t = model.new_bool_var(f"x_{d}_T_{wid}_FREE")
            x[d, "T", wid, "."] = var_free_t
            t_choices.append(var_free_t)

            # Si is_t == 0 -> Mañana
            model.add(sum(m_choices) == 1).only_enforce_if(is_t.negated())
            model.add(sum(t_choices) == 0).only_enforce_if(is_t.negated())

            # Si is_t == 1 -> Tarde
            model.add(sum(t_choices) == 1).only_enforce_if(is_t)
            model.add(sum(m_choices) == 0).only_enforce_if(is_t)

    # 4. Restricciones de Cobertura y Dotaciones por Puesto
    coverage_counts = {}
    vacancy_vars = []
    dr_vacancy_vars = []
    min_deficit_vars = []

    for s in workdays:
        d = s.day
        for shift, targets in (("M", s.targets_m), ("T", s.targets_t)):
            for pos, (tgt, mn) in targets.items():
                if tgt == 0:
                    if (d, shift, pos) in coverage_choices:
                        model.add(sum(coverage_choices[d, shift, pos]) == 0)
                    continue

                actual_count = model.new_int_var(0, tgt, f"cnt_{d}_{shift}_{pos}")
                coverage_counts[d, shift, pos] = actual_count
                model.add(actual_count == sum(coverage_choices[d, shift, pos]))

                # Vacante respecto al objetivo (0 .. tgt)
                vac = model.new_int_var(0, tgt, f"vac_{d}_{shift}_{pos}")
                model.add(vac == tgt - actual_count)
                vacancy_vars.append(vac)

                # Déficit respecto al mínimo (si actual_count < mn, def_min = mn - actual_count)
                if mn > 0:
                    def_min = model.new_int_var(0, mn, f"def_min_{d}_{shift}_{pos}")
                    model.add(def_min >= mn - actual_count)
                    min_deficit_vars.append(def_min)

                if pos == "DR":
                    dr_vacancy_vars.append(vac)

                # Regla de reducción ordenada: DR tarde solo se reduce si EC tarde ya está en su mínimo o menor
                if pos == "DR" and shift == "T" and (d, "T", "EC") in coverage_counts:
                    ec_cnt = coverage_counts[d, "T", "EC"]
                    ec_min = s.targets_t["EC"][1]
                    reduced_dr = model.new_bool_var(f"red_dr_{d}")
                    model.add(actual_count < tgt).only_enforce_if(reduced_dr)
                    model.add(actual_count == tgt).only_enforce_if(reduced_dr.negated())
                    model.add(ec_cnt <= ec_min).only_enforce_if(reduced_dr)

    total_min_deficits = sum(min_deficit_vars) if min_deficit_vars else 0
    total_vacancies = sum(vacancy_vars) if vacancy_vars else 0
    total_dr_vacancies = sum(dr_vacancy_vars) if dr_vacancy_vars else 0

    # 5. Cálculo de Esfuerzo Individual y Equidad
    worker_days_count = {}
    for w in workers:
        d_cnt = sum(1 for s in workdays if w.sequence[s.cycle_day - 1] in ("M", "T", "R"))
        worker_days_count[w.id] = d_cnt

    worker_pts = {}
    for w in workers:
        d_cnt = worker_days_count[w.id]
        if d_cnt > 0:
            pts_var = model.new_int_var(0, 16 * d_cnt, f"pts_{w.id}")
            terms = [
                x[d, shift, w.id, pos] * w.ranking.get(pos, 8)
                for s in workdays
                for d in (s.day,)
                for shift in ("M", "T")
                for pos in POSITIONS
                if (d, shift, w.id, pos) in x
            ]
            model.add(pts_var == sum(terms))
            worker_pts[w.id] = pts_var

    # Equidad Min-Max (D016): minimizar la dispersión entre la mayor y la menor media
    # Para evaluar con precisión decimal usando enteros en CP-SAT, escalamos por 100:
    # 100 * pts_i <= M_max_scaled * d_cnt
    # 100 * pts_i >= M_min_scaled * d_cnt
    m_max_scaled = model.new_int_var(0, 1600, "M_max_scaled")
    m_min_scaled = model.new_int_var(0, 1600, "M_min_scaled")
    spread_scaled = model.new_int_var(0, 1600, "spread_scaled")
    model.add(spread_scaled == m_max_scaled - m_min_scaled)

    for wid, pts_var in worker_pts.items():
        d_cnt = worker_days_count[wid]
        model.add(100 * pts_var <= m_max_scaled * d_cnt)
        model.add(100 * pts_var >= m_min_scaled * d_cnt)

    # Suma total de puntos (maximizar satisfacción general de preferencias)
    total_points_sum = sum(worker_pts.values()) if worker_pts else 0

    # Penalización suave: evitar más de 2 días consecutivos en el mismo puesto (si es posible)
    consec_penalty_vars = []
    for i in range(len(workdays) - 2):
        d1, d2, d3 = workdays[i].day, workdays[i + 1].day, workdays[i + 2].day
        for w in workers:
            if len(w.skills - w.vetoes) <= 1:
                continue
            for pos in POSITIONS:
                vars_d1 = [x[d1, sh, w.id, pos] for sh in ("M", "T") if (d1, sh, w.id, pos) in x]
                vars_d2 = [x[d2, sh, w.id, pos] for sh in ("M", "T") if (d2, sh, w.id, pos) in x]
                vars_d3 = [x[d3, sh, w.id, pos] for sh in ("M", "T") if (d3, sh, w.id, pos) in x]
                if vars_d1 and vars_d2 and vars_d3:
                    viol = model.new_bool_var(f"consec_{w.id}_{pos}_{d1}")
                    model.add(viol >= sum(vars_d1) + sum(vars_d2) + sum(vars_d3) - 2)
                    consec_penalty_vars.append(viol)

    total_consec_penalty = sum(consec_penalty_vars) if consec_penalty_vars else 0

    # 6. Optimización Lexicográfica en Fases:
    solver = cp_model.CpSolver()
    solver.parameters.random_seed = 42
    solver.parameters.num_search_workers = 4

    opt_min_deficits = 0
    opt_vacancies = 0
    opt_dr_vacancies = 0
    is_optimal = True

    # Fase 1: Minimizar déficit de mínimos y vacantes totales
    coverage_penalty = total_min_deficits * 1000 + total_vacancies
    model.minimize(coverage_penalty)
    solver.parameters.max_time_in_seconds = max(1.0, time_limit_seconds * 0.35)
    st1 = solver.solve(model)
    if st1 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        opt_min_deficits = int(solver.value(total_min_deficits))
        opt_vacancies = int(solver.value(total_vacancies))
        if st1 != cp_model.OPTIMAL:
            is_optimal = False
        model.add(total_min_deficits == opt_min_deficits)
        model.add(total_vacancies == opt_vacancies)
    else:
        return MonthResult(
            year=year,
            month=month,
            is_optimal=False,
            status_name="Infactible",
            vacancies_total=999,
            vacancies_dr=999,
            effort_spread=999.0,
            max_mean_effort=999.0,
            min_mean_effort=999.0,
            assignments=(),
            vacancies=(),
            worker_efforts={},
            execution_time_seconds=round(perf_counter() - start_time, 2),
            issues=("Tiempo agotado sin encontrar solución factible.",),
        )

    # Fase 2: Vacantes DR
    model.minimize(total_dr_vacancies)
    solver.parameters.max_time_in_seconds = max(1.0, time_limit_seconds * 0.25)
    st2 = solver.solve(model)
    if st2 in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        opt_dr_vacancies = int(solver.value(total_dr_vacancies))
        if st2 != cp_model.OPTIMAL:
            is_optimal = False
        model.add(total_dr_vacancies == opt_dr_vacancies)

    model.clear_hints()
    for v in x.values():
        model.add_hint(v, solver.value(v))

    # Fase 3: Equidad D016 (Minimizar dispersión de medias escaladas), no repetir >2 días y suma de puntos
    # spread_scaled * 10000 asegura prioridad sobre total_consec_penalty * 500 y total_points_sum
    model.minimize(spread_scaled * 10000 + total_consec_penalty * 500 + total_points_sum)
    rem_time = max(1.0, time_limit_seconds - (perf_counter() - start_time))
    solver.parameters.max_time_in_seconds = rem_time
    st3 = solver.solve(model)
    if st3 not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        # Fallback de seguridad: si la fase 3 agota el tiempo sin solución factible,
        # resolver con objetivo simple para recuperar la solución garantizada de Fase 2
        model.minimize(total_points_sum)
        solver.parameters.max_time_in_seconds = 2.0
        st3 = solver.solve(model)
        is_optimal = False
    elif st3 != cp_model.OPTIMAL:
        is_optimal = False

    status_name = "Óptimo" if is_optimal and st3 == cp_model.OPTIMAL else "Factible"

    # 7. Reconstrucción y Auditoría del Resultado
    assignments_list = []
    w_assignments = defaultdict(list)

    for (d, shift, wid, pos), var in x.items():
        if solver.value(var) == 1:
            w = workers_by_id[wid]
            pts = 0 if pos == "." else w.ranking.get(pos, 8)
            asg = Assignment(
                day=d,
                shift=shift,
                worker_id=wid,
                worker_name=w.name,
                position=pos,
                points=pts,
            )
            assignments_list.append(asg)
            w_assignments[wid].append(asg)

    assignments_list.sort(key=lambda a: (a.day, 0 if a.shift == "M" else 1, a.worker_name))

    # Auditoría independiente con reparto/auditoria.py
    issues = []
    audit_errors = audit_assignments(assignments_list, schedules, workers)
    issues.extend(audit_errors)

    vacancies = compute_vacancies(assignments_list, schedules)
    min_deficits_count = sum(v.missing_to_minimum for v in vacancies)
    if min_deficits_count > 0:
        issues.append(f"Aviso: {min_deficits_count} puestos mínimos quedaron sin cubrir por falta de personal disponible (quedan en blanco para revisar a mano).")
    elif len(vacancies) > 0:
        issues.append(f"Aviso: {len(vacancies)} vacantes respecto al objetivo cubiertas según reglas de dotación.")

    worker_efforts_dict = {}
    calculated_means = []

    for w in workers:
        asgs = w_assignments[w.id]
        d_count = worker_days_count[w.id]
        tot_pts = sum(a.points for a in asgs)
        free_days = sum(1 for a in asgs if a.position == ".")
        counts_pos = Counter(a.position for a in asgs if a.position != ".")

        mean_flt = (tot_pts / d_count) if d_count > 0 else 0.0
        exact_frac = Fraction(tot_pts, d_count) if d_count > 0 else None
        if d_count > 0:
            calculated_means.append(mean_flt)

        worker_efforts_dict[w.id] = WorkerEffort(
            worker_id=w.id,
            worker_name=w.name,
            worked_days=d_count,
            total_points=tot_pts,
            mean_effort=round(mean_flt, 2),
            exact_mean=exact_frac,
            free_days=free_days,
            counts_by_position=dict(counts_pos),
        )

    min_mean = min(calculated_means) if calculated_means else 0.0
    max_mean = max(calculated_means) if calculated_means else 0.0
    spread = max_mean - min_mean

    return MonthResult(
        year=year,
        month=month,
        is_optimal=is_optimal,
        status_name=status_name,
        vacancies_total=len(vacancies),
        vacancies_dr=opt_dr_vacancies,
        effort_spread=round(spread, 2),
        max_mean_effort=round(max_mean, 2),
        min_mean_effort=round(min_mean, 2),
        assignments=tuple(assignments_list),
        vacancies=vacancies,
        worker_efforts=worker_efforts_dict,
        execution_time_seconds=round(perf_counter() - start_time, 2),
        issues=tuple(issues),
    )
