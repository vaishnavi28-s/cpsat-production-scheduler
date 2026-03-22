"""
CP-SAT optimizer for print job combination.

Groups print jobs by (PRODUCTTYPE, PRESS_LOCATION, SEND_TO_LOCATION, PAPER)
and solves a bin-packing problem to find optimal job combinations that can
run together on the same press, subject to color, finishing, and quantity constraints.
"""

import re
import time
from collections import defaultdict

from ortools.sat.python import cp_model

from .finishing_rules import (
    type_label, signature, special_ink_set, finishing_conflict, SPECIAL_TYPES
)

SHEET_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*[xX]\s*(\d+(?:\.\d+)?(?:\s+\d+\s*/\s*\d+)?)"
)
MAX_QTY_DELTA = 3500
MAX_SPECIAL_PER_JOB = 2


def _parse_dim(s: str):
    s = (s or "").strip()
    if not s:
        return None
    m = re.match(r"^(\d+)\s+(\d+)\s*/\s*(\d+)$", s)
    if m:
        return float(m.group(1)) + float(m.group(2)) / float(m.group(3))
    try:
        return float(s)
    except Exception:
        return None


def is_28x20_sheet(job) -> bool:
    paper = str(getattr(job, "PAPER", "") or "")
    m = SHEET_RE.search(paper)
    if not m:
        return False
    d1 = _parse_dim(m.group(1))
    d2 = _parse_dim(m.group(2))
    if d1 is None or d2 is None:
        return False
    dims = sorted([round(d1, 2), round(d2, 2)])
    return abs(dims[1] - 28.0) <= 0.1 and (abs(dims[0] - 20.0) <= 0.1 or abs(dims[0] - 20.5) <= 0.1)


def solve_group(group_jobs: list) -> list:
    """
    Solve optimal job combinations for a single group using CP-SAT.

    Returns a list of combos, where each combo is a list of JOB IDs.
    """
    n = len(group_jobs)
    if n == 0:
        return []

    t_start = time.time()
    ptype_lc = (group_jobs[0].PRODUCTTYPE or "").strip().lower()

    if "cover" in ptype_lc:
        capacity = 4
    elif "jacket" in ptype_lc:
        capacity = 2
    else:
        return [[j.JOB] for j in group_jobs]

    if is_28x20_sheet(group_jobs[0]):
        capacity = 2

    TL = [type_label(j) for j in group_jobs]
    SIG = [signature(j) for j in group_jobs]
    IS40 = [1 if tl == "4/0" else 0 for tl in TL]
    ISSP = [1 if tl in SPECIAL_TYPES else 0 for tl in TL]
    QORD = [int(getattr(j, "QUANTITYORDERED", 0)) for j in group_jobs]
    SPSET = [special_ink_set(j) for j in group_jobs]
    NSP = [len(s) for s in SPSET]

    model = cp_model.CpModel()

    x = {(i, g): model.NewBoolVar(f"x_{i}_{g}") for i in range(n) for g in range(n)}
    y = [model.NewBoolVar(f"y_{g}") for g in range(n)]

    for i in range(n):
        model.AddHint(x[(i, i)], 1)
        for g in range(n):
            if g != i:
                model.AddHint(x[(i, g)], 0)

    for i in range(n):
        model.Add(sum(x[(i, g)] for g in range(n)) == 1)

    for g in range(n):
        s = sum(x[(i, g)] for i in range(n))
        model.Add(s >= 1).OnlyEnforceIf(y[g])
        model.Add(s == 0).OnlyEnforceIf(y[g].Not())

    for g in range(n - 1):
        model.Add(y[g] >= y[g + 1])

    for g in range(n):
        model.Add(sum(x[(i, g)] for i in range(n)) <= capacity).OnlyEnforceIf(y[g])

    is_cover = "cover" in ptype_lc
    is_jacket = "jacket" in ptype_lc
    full_bin_bonus = []

    if is_cover and capacity == 4:
        for g in range(n):
            bin_sz = model.NewIntVar(0, capacity, f"binsz_{g}")
            model.Add(bin_sz == sum(x[(i, g)] for i in range(n)))
            model.Add(bin_sz != 3).OnlyEnforceIf(y[g])
            fb = model.NewBoolVar(f"fullbin_{g}")
            model.Add(fb <= y[g])
            model.Add(bin_sz == 4).OnlyEnforceIf(fb)
            model.Add(bin_sz <= 3).OnlyEnforceIf(fb.Not())
            full_bin_bonus.append(fb)

    # Color compatibility pairs
    color_incompat = []
    for i in range(n):
        for j in range(i + 1, n):
            heavy_i = NSP[i] > MAX_SPECIAL_PER_JOB
            heavy_j = NSP[j] > MAX_SPECIAL_PER_JOB
            if heavy_i or heavy_j:
                if SPSET[i] != SPSET[j]:
                    color_incompat.append((i, j))
                continue
            if TL[i] in SPECIAL_TYPES and TL[j] in SPECIAL_TYPES:
                if SIG[i] != SIG[j]:
                    color_incompat.append((i, j))
            elif TL[i] == "OTHER" or TL[j] == "OTHER":
                if not (TL[i] == "OTHER" and TL[j] == "OTHER" and SIG[i] == SIG[j]):
                    color_incompat.append((i, j))

    # Finishing conflict pairs
    finish_incompat = [
        (i, j)
        for i in range(n)
        for j in range(i + 1, n)
        if finishing_conflict(group_jobs[i], group_jobs[j])
    ]

    # Quantity conflict pairs
    qty_incompat = [
        (i, j)
        for i in range(n)
        for j in range(i + 1, n)
        if abs(QORD[i] - QORD[j]) > MAX_QTY_DELTA
    ]

    all_conflicts = set(color_incompat + finish_incompat + qty_incompat)
    for g in range(n):
        for (i, j) in all_conflicts:
            model.Add(x[(i, g)] + x[(j, g)] <= 1).OnlyEnforceIf(y[g])

    # At most one special if 4/0 is in bin
    for g in range(n):
        s40 = sum(x[(i, g)] for i in range(n) if IS40[i])
        has40 = model.NewBoolVar(f"has40_{g}")
        model.Add(s40 >= 1).OnlyEnforceIf(has40)
        model.Add(s40 == 0).OnlyEnforceIf(has40.Not())
        ssp = sum(x[(i, g)] for i in range(n) if ISSP[i])
        model.Add(ssp <= 1 + capacity * (1 - has40))

    # Soft: minimize quantity spread
    MAX_RUN_QTY = max(max(QORD) if QORD else 0, MAX_QTY_DELTA)
    BIG_M = MAX_RUN_QTY
    maxqty_g = []
    for g in range(n):
        mq = model.NewIntVar(0, MAX_RUN_QTY, f"maxqty_{g}")
        maxqty_g.append(mq)
        for i in range(n):
            model.Add(mq >= QORD[i]).OnlyEnforceIf(x[(i, g)])

    diff = {}
    for g in range(n):
        for i in range(n):
            d = model.NewIntVar(0, BIG_M, f"diff_{i}_{g}")
            diff[(i, g)] = d
            model.Add(d <= BIG_M * x[(i, g)])
            model.Add(d >= maxqty_g[g] - QORD[i] - BIG_M * (1 - x[(i, g)]))
            model.Add(d <= maxqty_g[g] - QORD[i] + BIG_M * (1 - x[(i, g)]))

    total_spread = sum(diff[(i, g)] for g in range(n) for i in range(n))

    # Soft: 2x quantity ratio bonus for covers
    ratio_bonus = []
    if is_cover and capacity == 4:
        for g in range(n):
            for i in range(n):
                for j in range(i + 1, n):
                    qi, qj = QORD[i], QORD[j]
                    if qi <= 0 or qj <= 0:
                        continue
                    r = max(qi, qj) / min(qi, qj)
                    if 1.7 <= r <= 2.3:
                        is2x = model.NewBoolVar(f"is2x_{i}_{j}_{g}")
                        model.Add(is2x <= x[(i, g)])
                        model.Add(is2x <= x[(j, g)])
                        model.Add(is2x >= x[(i, g)] + x[(j, g)] - 1)
                        ratio_bonus.append(is2x)

    # Soft: jacket 4/4 pairing
    jacket_44_bonus = []
    if is_jacket and capacity == 2:
        idx_44 = [i for i, tl in enumerate(TL) if tl == "4/4"]
        if len(idx_44) >= 2:
            for g in range(n):
                for a in range(len(idx_44)):
                    for b in range(a + 1, len(idx_44)):
                        i, j = idx_44[a], idx_44[b]
                        pair44 = model.NewBoolVar(f"pair44_{i}_{j}_{g}")
                        model.Add(pair44 <= x[(i, g)])
                        model.Add(pair44 <= x[(j, g)])
                        model.Add(pair44 >= x[(i, g)] + x[(j, g)] - 1)
                        jacket_44_bonus.append(pair44)

    bonus_expr = sum(full_bin_bonus) + sum(ratio_bonus) + sum(jacket_44_bonus)
    max_spread = BIG_M * n * n
    max_bonus = n * n + n
    MID_W = max_bonus + 1
    BIG_W = MID_W * max_spread + max_bonus + 1

    model.Minimize(BIG_W * sum(y) + MID_W * total_spread - bonus_expr)

    model.AddDecisionStrategy(
        [x[(i, g)] for g in range(n) for i in range(n)],
        cp_model.CHOOSE_FIRST, cp_model.SELECT_MAX_VALUE
    )

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 90
    solver.parameters.num_search_workers = 8
    solver.parameters.random_seed = 123
    solver.parameters.log_search_progress = False
    solver.parameters.cp_model_presolve = True

    status = solver.Solve(model)
    elapsed = time.time() - t_start
    print(f"  [n={n:3d} | {elapsed:.2f}s | status={status}]")

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return [
            [group_jobs[i].JOB for i in range(n) if solver.Value(x[(i, g)]) == 1]
            for g in range(n)
            if any(solver.Value(x[(i, g)]) == 1 for i in range(n))
        ]

    return [[j.JOB] for j in group_jobs]


def run_optimizer(jobs: list) -> dict:
    """
    Run the optimizer across all jobs, grouped by location.

    Returns:
        dict: {press_location: {combo_id: [job_ids]}}
    """
    jobs_by_loc = defaultdict(list)
    for j in jobs:
        jobs_by_loc[j.PRESS_LOCATION or "Unknown"].append(j)

    total_runs = {}

    for loc, loc_jobs in jobs_by_loc.items():
        print(f"\n{'='*50}")
        print(f"Location: {loc} ({len(loc_jobs)} jobs)")
        print('=' * 50)

        groups = defaultdict(list)
        for job in loc_jobs:
            groups[job.group_key()].append(job)

        runs = {}
        cid = 1

        for key, group_jobs in groups.items():
            group_jobs = sorted(group_jobs, key=lambda j: j.JOB)
            combos = solve_group(group_jobs)
            for combo in combos:
                runs[cid] = combo
                cid += 1

        total_runs[loc] = runs

    return total_runs
