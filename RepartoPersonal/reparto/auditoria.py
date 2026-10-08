"""Auditoría independiente del cuadrante calculado.

No usa variables del solver: parte únicamente de las asignaciones, el calendario y los
perfiles. Comprueba conocimientos, vetos, presencia, duplicados y dotaciones, y calcula
las vacantes (huecos que quedan en blanco para revisar a mano).
"""

from collections import Counter
from dataclasses import dataclass
from datetime import date

from reparto.datos import DaySchedule, WorkerProfile


@dataclass(frozen=True)
class Vacancy:
    """Puesto sin cubrir en un turno de un día.

    - ``below_minimum`` True: quedó por debajo del mínimo permitido (hueco a revisar a mano).
    - ``below_minimum`` False: reducción permitida respecto al objetivo (EC/DR, D012/D013).
    """
    day: date
    shift: str
    position: str
    target: int
    minimum: int
    assigned: int

    @property
    def missing(self) -> int:
        return self.target - self.assigned

    @property
    def below_minimum(self) -> bool:
        return self.assigned < self.minimum

    @property
    def missing_to_minimum(self) -> int:
        return max(0, self.minimum - self.assigned)


def compute_vacancies(assignments, schedules: tuple[DaySchedule, ...]) -> tuple[Vacancy, ...]:
    """Recuenta puestos asignados por día/turno y devuelve los que no alcanzan el objetivo."""
    counts = Counter(
        (a.day, a.shift, a.position) for a in assignments if a.position != "."
    )
    result = []
    for s in schedules:
        if not s.is_workday:
            continue
        for shift, targets in (("M", s.targets_m), ("T", s.targets_t)):
            for pos, (target, minimum) in targets.items():
                if target <= 0:
                    continue
                assigned = counts.get((s.day, shift, pos), 0)
                if assigned < target:
                    result.append(Vacancy(s.day, shift, pos, target, minimum, assigned))
    return tuple(result)


def audit_assignments(
    assignments,
    schedules: tuple[DaySchedule, ...],
    workers: tuple[WorkerProfile, ...],
) -> list[str]:
    """Devuelve la lista de violaciones de reglas (vacía si todo es correcto)."""
    issues: list[str] = []
    workers_by_id = {w.id: w for w in workers}
    day_info = {s.day: s for s in schedules}

    per_worker_day = Counter()
    worker_shifts: dict[tuple[str, date], set[str]] = {}
    pos_counts = Counter()

    for a in assignments:
        w = workers_by_id.get(a.worker_id)
        s = day_info.get(a.day)
        if w is None or s is None:
            issues.append(f"AUDITORÍA: asignación con trabajador/día desconocido ({a.worker_id}, {a.day})")
            continue
        if not s.is_workday:
            issues.append(f"AUDITORÍA: {w.name} asignado en día no laborable {a.day}")
        per_worker_day[(a.worker_id, a.day)] += 1
        worker_shifts.setdefault((a.worker_id, a.day), set()).add(a.shift)

        mark = w.sequence[s.cycle_day - 1]
        if mark == "M" and a.shift != "M" or mark == "T" and a.shift != "T" or mark not in ("M", "T", "R"):
            issues.append(f"AUDITORÍA: {w.name} asignado el {a.day} en turno {a.shift} pero su marca es {mark}")

        if a.position != ".":
            pos_counts[(a.day, a.shift, a.position)] += 1
            if a.position in w.vetoes:
                issues.append(f"VIOLACIÓN DE VETO: {w.name} asignado a {a.position} el {a.day}")
            if a.position not in w.skills:
                issues.append(f"FALTA DE CONOCIMIENTO: {w.name} no sabe {a.position} (asignado el {a.day})")

    for (wid, day), n in per_worker_day.items():
        if n != 1:
            issues.append(f"AUDITORÍA: {workers_by_id[wid].name} tiene {n} asignaciones el {day}")

    # Todos los presentes deben tener exactamente una asignación (puesto o Libre)
    for s in schedules:
        if not s.is_workday:
            continue
        for wid in (*s.workers_m, *s.workers_t, *s.workers_r):
            if per_worker_day.get((wid, s.day), 0) != 1:
                issues.append(f"AUDITORÍA: {workers_by_id[wid].name} sin asignación el {s.day}")

    # Dotaciones: nunca por encima del objetivo; puestos con objetivo 0 permanecen cerrados
    for s in schedules:
        if not s.is_workday:
            continue
        for shift, targets in (("M", s.targets_m), ("T", s.targets_t)):
            for pos, (target, _minimum) in targets.items():
                if pos_counts.get((s.day, shift, pos), 0) > target:
                    issues.append(
                        f"AUDITORÍA: {pos} {shift} el {s.day} supera la dotación ({pos_counts[(s.day, shift, pos)]}>{target})"
                    )
    return issues
