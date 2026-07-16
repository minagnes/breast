\
# -*- coding: utf-8 -*-
"""
유방영상의학과 근무 스케줄 자동 생성기 (schedule_generator.py)
================================================================
기준 문서: 스케줄_배정_규정.docx (K3 퇴사→K5 대체, F4 신규 반영 버전)

* K4·K5 고정 배정 규칙 (2026-09 갱신):
    - K5: 오전 Mammo 화·금 고정 / 오후 Breast US 화·금 고정 / 오후 Mammo2 월·수·목 고정
    - K4: 오후 Mammo2 월·화·금 고정

YEAR/MONTH/HOLIDAYS/VACATIONS 는 이 파일 상단 "설정" 구역에서 직접 채워서 쓴다.
연/월을 매번 터미널에서 입력받고 vacation 엑셀과 자동으로 연동하려면
generate_schedule.py 를 사용할 것.
"""

import os
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# 1. 설정 (매달 갱신)
# ============================================================

YEAR = 2026
MONTH = 9

# 공휴일 (2026년 추석: 9/24(목)~9/26(토))
HOLIDAYS = {
    date(2026, 9, 24): "추석",
    date(2026, 9, 25): "추석",
    date(2026, 9, 26): "추석",
}

# 휴가/반휴/학회/출장/교육 등 근무 불가 일정.
# 형식: (인원, date(년,월,일), 구분, 라벨)   구분: "AM"/"PM"/"ALL"
VACATIONS = [
    # ("K2", date(2026, 9, 7), "ALL", "휴가"),
]

# 전공의(R) 표기 로테이션 기준
ROTATION_ANCHOR_MONDAY = date(2026, 8, 17)
ROTATION_ANCHOR_LABEL = 1

# 수동 보정 (14장)
MANUAL_ADJUSTMENTS = [
    # ("F2", date(2026, 9, 10), "PM", "breast", "mammo2"),
]

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# 2. 인력 구성 (규정 2장)
# ============================================================

K_STAFF = ["K1", "K2", "K4", "K5"]
F_STAFF = ["F1", "F2", "F3", "F4"]
J_STAFF = ["J1", "J2"]
R_NAME = "R"

ALL_NAMES = K_STAFF + F_STAFF + J_STAFF + [R_NAME]

# 오전 Breast US 주 담당 K (요일별 1명): 월/목/금=K4, 화/수=K2 (6.1)
BREAST_AM_LEAD = {0: "K4", 1: "K2", 2: "K2", 3: "K4", 4: "K4"}
# K5 추가 고정: 월/수/목 오전 Breast US 에 추가로 고정 (6.1)
BREAST_AM_EXTRA = {0: "K5", 2: "K5", 3: "K5"}
# K5: 오전 Mammo 화/금 고정
K5_AM_MAMMO_DAYS = {1, 4}

# 오후 Mammo2 요일별 고정 K (7.1) — K5: 월·수·목 / K4: 월·화·금
MAMMO2_PM_FIX = {
    0: ["K1", "K5", "K4"],  # 월
    1: ["K1", "K2", "K4"],  # 화
    2: ["K1", "K2", "K5"],  # 수
    3: ["K5"],              # 목
    4: ["K1", "K2", "K4"],  # 금
}
# 오후 Breast US 지정자 (Mammo2 비고정 K 중) — K5: 화/금 고정
BREAST_PM_FIX = {1: "K5", 2: "K4", 4: "K5"}   # K5: 오후 Breast US 화·금 고정 / K4: 오후 Breast US 수 고정

WEEKDAY_NAMES = ["월", "화", "수", "목", "금", "토"]

if MONTH % 2 == 1:
    MONWED_ITEM, TUETHU_ITEM = "VAB", "ST-MMT"
else:
    MONWED_ITEM, TUETHU_ITEM = "ST-MMT", "VAB"

# ============================================================
# 3. 달력 / 휴가 유틸
# ============================================================


def month_weeks(year, month):
    probe = date(year, month, 15) - timedelta(days=40)
    mondays = []
    d = probe - timedelta(days=probe.weekday())
    for _ in range(16):
        wed = d + timedelta(days=2)
        if wed.year == year and wed.month == month:
            mondays.append(d)
        d += timedelta(days=7)
    return sorted(set(mondays))


def is_holiday(d):
    return d in HOLIDAYS


def is_workday(d):
    return d.weekday() != 5 and not is_holiday(d)


VAC_INDEX = {}
for (nm, dt_, period, label) in VACATIONS:
    VAC_INDEX.setdefault((nm, dt_), []).append((period, label))


def away(name, d, session):
    for period, _label in VAC_INDEX.get((name, d), []):
        if period == "ALL" or period == session:
            return True
    return False


def rotation_label_for_monday(monday):
    weeks_diff = (monday - ROTATION_ANCHOR_MONDAY).days // 7
    num = (ROTATION_ANCHOR_LABEL - 1 + weeks_diff) % 4 + 1
    return f"R{num}"


# ============================================================
# 4. 상태(월간 누적 카운터)
# ============================================================


def new_counter(keys):
    return {k: 0 for k in keys}


class State:
    def __init__(self):
        self.f_am = {f: new_counter(["breast", "thyroid", "mammo"]) for f in F_STAFF}
        self.r_am = new_counter(["breast", "thyroid", "mammo"])
        self.r_am_week = new_counter(["breast", "thyroid", "mammo"])
        self.j_am = {j: new_counter(["breast", "thyroid"]) for j in J_STAFF}
        self.f_pm = {f: new_counter(["breast", "thyroid"]) for f in F_STAFF}
        self.r_pm = new_counter(["breast", "thyroid"])
        self.f_locali_am = new_counter(F_STAFF)
        self.f_locali_am_first = new_counter(F_STAFF)
        self.f_locali_pm = new_counter(F_STAFF)
        self.k_vab = new_counter(K_STAFF)
        self.k_stmmt = new_counter(K_STAFF)
        self.r_mammo_partners = set()
        self.f_mammo2_rotation_idx = 0


STATE = State()

# ============================================================
# 5. 오전(AM) 배정 (6장)
# ============================================================


def assign_am(d, week_index, is_first_workday):
    wd = d.weekday()
    notes = []

    avail_K = [k for k in K_STAFF if not away(k, d, "AM")]
    avail_F = [f for f in F_STAFF if not away(f, d, "AM")]
    avail_R = [] if away(R_NAME, d, "AM") else [R_NAME]
    avail_J = [j for j in J_STAFF if not away(j, d, "AM")]

    duties = {"breast": [], "mammo": [], "thyroid": [], "abus": []}

    used_K = set()
    if "K1" in avail_K:
        duties["thyroid"].append("K1")
        used_K.add("K1")

    lead = BREAST_AM_LEAD[wd]
    extra = BREAST_AM_EXTRA.get(wd)
    if lead in avail_K:
        duties["breast"].append(lead)
        used_K.add(lead)
    else:
        pool = [
            k for k in avail_K
            if k not in used_K and k != extra and k != lead
            and not (wd in K5_AM_MAMMO_DAYS and k == "K5")  # 화/금 K5는 Mammo 고정, 대체 후보 제외
        ]
        if pool:
            sub = pool[0]
            duties["breast"].append(sub)
            used_K.add(sub)
            notes.append(f"{d.month}/{d.day} {lead} 오전 휴가 → {sub}(으)로 Breast US 대체")
    if extra and extra in avail_K and extra not in used_K:
        duties["breast"].append(extra)
        used_K.add(extra)

    # 남은 K: 화/금은 K5가 Mammo 고정, 그 외에는 1명 Mammo·나머지 ABUS
    remaining_K = [k for k in avail_K if k not in used_K]
    if wd in K5_AM_MAMMO_DAYS and "K5" in remaining_K:
        duties["mammo"].append("K5")
        for k in remaining_K:
            if k != "K5":
                duties["abus"].append(k)
    elif remaining_K:
        duties["mammo"].append(remaining_K[0])
        for k in remaining_K[1:]:
            duties["abus"].append(k)

    breast_cap = (4 if wd in (0, 2, 3) else 5) + (1 if extra and extra in duties["breast"] else 0)
    thyroid_cap = 3 if wd in (0, 2, 3) else 4
    mammo_cap = 2

    def slot_open(cat):
        if cat == "breast":
            return len(duties["breast"]) < breast_cap
        if cat == "thyroid":
            return len(duties["thyroid"]) < thyroid_cap
        if cat == "mammo":
            return len(duties["mammo"]) < mammo_cap and len(duties["mammo"]) >= 1
        return False

    day_ordinal = (d - date(d.year, 1, 1)).days
    start = day_ordinal % max(len(avail_F), 1)
    order_F = avail_F[start:] + avail_F[:start]

    weight = {"breast": 1.3, "thyroid": 1.0, "mammo": 1.0}
    unassigned_F = []
    for f in order_F:
        open_cats = [c for c in ("breast", "thyroid", "mammo") if slot_open(c)]
        if not open_cats:
            unassigned_F.append(f)
            continue
        cat = min(open_cats, key=lambda c: STATE.f_am[f][c] / weight[c])
        duties[cat].append(f)
        STATE.f_am[f][cat] += 1

    for r in avail_R:
        open_cats = [c for c in ("breast", "thyroid", "mammo") if slot_open(c)]
        chosen = None
        if "mammo" in open_cats and STATE.r_am_week["mammo"] < 1:
            chosen = "mammo"
        elif open_cats:
            chosen = min(open_cats, key=lambda c: STATE.r_am_week[c])
        elif thyroid_cap and len(duties["thyroid"]) < thyroid_cap + 1:
            chosen = "thyroid"
        if chosen:
            duties[chosen].append(R_NAME)
            STATE.r_am[chosen] += 1
            STATE.r_am_week[chosen] += 1
            if chosen == "mammo" and duties["mammo"]:
                STATE.r_mammo_partners.add(duties["mammo"][0])
        else:
            notes.append(f"{d.month}/{d.day} R 오전 자리 부족 → 배정 실패(정원 초과 확인 필요)")

    if len(avail_J) == 2:
        j_a, j_b = avail_J
        cat_a = min(("breast", "thyroid"), key=lambda c: STATE.j_am[j_a][c])
        cat_b = "thyroid" if cat_a == "breast" else "breast"
        duties[cat_a].append(j_a)
        duties[cat_b].append(j_b)
        STATE.j_am[j_a][cat_a] += 1
        STATE.j_am[j_b][cat_b] += 1
        for c in (cat_a, cat_b):
            if len(duties[c]) > (breast_cap if c == "breast" else thyroid_cap):
                notes.append(f"{d.month}/{d.day} J 배정으로 {c} 정원 초과(규정상 정상)")
    elif len(avail_J) == 1:
        j = avail_J[0]
        cat = min(("breast", "thyroid"), key=lambda c: STATE.j_am[j][c])
        duties[cat].append(j)
        STATE.j_am[j][cat] += 1

    r_cat = None
    for c in ("breast", "thyroid"):
        if R_NAME in duties[c]:
            r_cat = c
            break
    for f in unassigned_F:
        target = r_cat if r_cat else "breast"
        duties[target].append(f)
        STATE.f_am[f][target] += 1
        tag = "정원 여유분" if len(duties[target]) <= (breast_cap if target == "breast" else thyroid_cap) else "정원 초과"
        notes.append(f"{d.month}/{d.day} {f} 오전 자리 부족 → {'Thyroid US' if target=='thyroid' else 'Breast US'}에 배정({tag})")

    breast_F = [p for p in duties["breast"] if p in F_STAFF]
    us_locali = None
    if breast_F:
        if is_first_workday:
            us_locali = min(breast_F, key=lambda f: (STATE.f_locali_am_first[f], STATE.f_locali_am[f]))
            STATE.f_locali_am_first[us_locali] += 1
        else:
            us_locali = min(breast_F, key=lambda f: STATE.f_locali_am[f])
        STATE.f_locali_am[us_locali] += 1

    return duties, us_locali, notes


# ============================================================
# 6. 오후(PM) 배정 (7장)
# ============================================================


def assign_pm(d, week_dates, notes_out):
    wd = d.weekday()
    notes = []

    avail_K = [k for k in K_STAFF if not away(k, d, "PM")]
    avail_F = [f for f in F_STAFF if not away(f, d, "PM")]
    avail_R = [] if away(R_NAME, d, "PM") else [R_NAME]
    avail_J1 = "J1" in J_STAFF and not away("J1", d, "PM")

    duties = {"breast": [], "mammo": [], "thyroid": [], "abus": [], "mammo2": []}

    fixed = [k for k in MAMMO2_PM_FIX.get(wd, []) if k in avail_K]
    duties["mammo2"].extend(fixed)

    remaining_K = [k for k in avail_K if k not in fixed]
    for k in remaining_K:
        if k == "K1":
            duties["thyroid"].append(k)
        elif k == "K2":
            duties["breast"].append(k)
        elif BREAST_PM_FIX.get(wd) == k:
            duties["breast"].append(k)
        else:
            duties["thyroid"].append(k)

    if wd in (1, 2, 3) and not duties["breast"]:
        # K2는 화/수/금 오후 Mammo2 고정이라(신규 규칙) 대체 후보에서 제외 - K1처럼 대체 없이 유지
        candidates = [k for k in duties["mammo2"] if k in ("K5", "K4")]
        if candidates:
            sub = candidates[0]
            duties["mammo2"].remove(sub)
            duties["breast"].append(sub)
            extra_note = f"(후보: {', '.join(candidates)} 중 자동 배정, 다른 인원으로 변경 가능)" if len(candidates) > 1 else ""
            notes.append(f"{d.month}/{d.day} 오후 Breast US 공백 → Mammo2 담당 K 중 {sub}(으)로 대체 {extra_note}".strip())

    if wd == 4 and "K4" not in duties["thyroid"] and "K4" not in avail_K:
        rec = [k for k in ("K2", "K5") if k in duties["mammo2"]]
        if rec:
            notes.append(f"{d.month}/{d.day} 금요일 오후 Thyroid US(K4) 부재 → 대체 후보: {', '.join(rec)} (자동 대체 없음, 검토 필요)")

    if wd == 0 and avail_J1:
        if not duties["breast"]:
            duties["breast"].append("J1")
        else:
            duties["thyroid"].append("J1")
    elif wd == 4 and avail_J1:
        duties["breast"].append("J1")

    return duties, notes, avail_F, avail_R


def finish_pm_f_and_r(duties, avail_F, avail_R, f_mammo2_today, notes):
    for f in avail_F:
        if f in f_mammo2_today:
            duties["mammo2"].append(f)
            continue
        cat = min(("breast", "thyroid"), key=lambda c: STATE.f_pm[f][c])
        duties[cat].append(f)
        STATE.f_pm[f][cat] += 1

    for r in avail_R:
        cat = min(("breast", "thyroid"), key=lambda c: STATE.r_pm[c])
        duties[cat].append(r)
        STATE.r_pm[cat] += 1

    return duties


def balance_pm(duties, wd, notes):
    def movable(cat):
        return [p for p in duties[cat] if p in F_STAFF or p == R_NAME]

    while len(duties["thyroid"]) < 2 and movable("breast"):
        p = movable("breast")[0]
        duties["breast"].remove(p)
        duties["thyroid"].append(p)

    if wd != 4:
        while len(duties["breast"]) < 2 and movable("thyroid"):
            p = movable("thyroid")[0]
            duties["thyroid"].remove(p)
            duties["breast"].append(p)
        guard = 0
        while abs(len(duties["breast"]) - len(duties["thyroid"])) > 1 and guard < 10:
            guard += 1
            if len(duties["breast"]) > len(duties["thyroid"]) and movable("breast"):
                p = movable("breast")[0]
                duties["breast"].remove(p)
                duties["thyroid"].append(p)
            elif len(duties["thyroid"]) > len(duties["breast"]) and movable("thyroid") and len(duties["thyroid"]) > 2:
                p = movable("thyroid")[0]
                duties["thyroid"].remove(p)
                duties["breast"].append(p)
            else:
                break
    return duties


# ============================================================
# 7. VAB / ST-MMT / Locali (8장)
# ============================================================


def assign_extras_pm(d, pm_duties):
    wd = d.weekday()
    breast_K = [p for p in pm_duties["breast"] if p in K_STAFF]
    breast_F = [p for p in pm_duties["breast"] if p in F_STAFF]

    vab = stmmt = locali = None

    if wd in (0, 2) and breast_K:
        item = MONWED_ITEM
    elif wd in (1, 3) and breast_K:
        item = TUETHU_ITEM
    else:
        item = None

    if item and breast_K:
        counter = STATE.k_vab if item == "VAB" else STATE.k_stmmt
        chosen = min(breast_K, key=lambda k: counter[k])
        counter[chosen] += 1
        if item == "VAB":
            vab = chosen
        else:
            stmmt = chosen

    if wd in (0, 1, 2, 3) and breast_F:
        locali = min(breast_F, key=lambda f: STATE.f_locali_pm[f])
        STATE.f_locali_pm[locali] += 1

    return vab, stmmt, locali


# ============================================================
# 8. 주 단위 F Mammo2 순환 계획 (7.3)
# ============================================================


def plan_week_f_mammo2(week_dates):
    workdays = [d for d in week_dates if is_workday(d)]
    plan = {d: [] for d in workdays}
    if not workdays:
        return plan

    eligible = []
    for f in F_STAFF:
        vac_days = sum(1 for d in workdays if away(f, d, "PM") or away(f, d, "ALL"))
        if vac_days < 2:
            eligible.append(f)

    n = len(workdays)
    for i, f in enumerate(eligible):
        idx = (i + STATE.f_mammo2_rotation_idx) % n
        target_day = workdays[idx]
        if not away(f, target_day, "PM"):
            plan[target_day].append(f)
    STATE.f_mammo2_rotation_idx = (STATE.f_mammo2_rotation_idx + 1) % max(n, 1)
    return plan


# ============================================================
# 9. 하루 전체 조립 + 수동 보정 적용
# ============================================================


def apply_manual_adjustments(day_data):
    for (name, d, session, src, dst) in MANUAL_ADJUSTMENTS:
        if d not in day_data:
            continue
        duties = day_data[d]["am"] if session == "AM" else day_data[d]["pm"]
        if src in duties and name in duties[src]:
            duties[src].remove(name)
            duties.setdefault(dst, []).append(name)
            day_data[d]["notes"].append(
                f"{d.month}/{d.day} {name} {'오전' if session=='AM' else '오후'} 수동 보정: {src} → {dst}"
            )


def build_schedule():
    weeks = month_weeks(YEAR, MONTH)
    day_data = {}

    for week_index, monday in enumerate(weeks):
        week_dates = [monday + timedelta(days=i) for i in range(6)]
        workdays = [d for d in week_dates if is_workday(d)]
        first_workday = workdays[0] if workdays else None

        f_mammo2_plan = plan_week_f_mammo2(week_dates)
        STATE.r_am_week = new_counter(["breast", "thyroid", "mammo"])

        for d in week_dates:
            entry = {
                "am": {"breast": [], "mammo": [], "thyroid": [], "abus": []},
                "pm": {"breast": [], "mammo": [], "thyroid": [], "abus": [], "mammo2": []},
                "us_locali": None,
                "vab": None,
                "stmmt": None,
                "locali": None,
                "notes": [],
                "week_index": week_index,
                "rotation_label": rotation_label_for_monday(monday),
            }
            if is_workday(d):
                am_duties, us_locali, am_notes = assign_am(d, week_index, d == first_workday)
                pm_duties, pm_notes, avail_F, avail_R = assign_pm(d, week_dates, entry["notes"])
                pm_duties = finish_pm_f_and_r(pm_duties, avail_F, avail_R, f_mammo2_plan.get(d, []), pm_notes)
                pm_duties = balance_pm(pm_duties, d.weekday(), pm_notes)
                vab, stmmt, locali = assign_extras_pm(d, pm_duties)

                entry["am"] = am_duties
                entry["pm"] = pm_duties
                entry["us_locali"] = us_locali
                entry["vab"] = vab
                entry["stmmt"] = stmmt
                entry["locali"] = locali
                entry["notes"] = am_notes + pm_notes
            day_data[d] = entry

    apply_manual_adjustments(day_data)
    return weeks, day_data


# ============================================================
# 10. 통계 / 형평성 (10장)
# ============================================================

STAT_COLS = ["오전 Breast US", "오전 Mammo", "오전 Thyroid US", "오전 ABUS", "오후 Breast US", "오후 Thyroid US"]
STAT_KEYS = [("am", "breast"), ("am", "mammo"), ("am", "thyroid"), ("am", "abus"), ("pm", "breast"), ("pm", "thyroid")]


def compute_stats(day_data):
    people = K_STAFF + F_STAFF + ["J1", "J2", "R"]
    stat = {p: {"counts": [0] * 6, "mammo2": 0, "workdays": 0} for p in people}

    for d, entry in day_data.items():
        if not is_workday(d):
            continue
        present_today = set()
        for i, (sess, key) in enumerate(STAT_KEYS):
            for p in entry[sess][key]:
                stat[p]["counts"][i] += 1
                present_today.add(p)
        for p in entry["pm"]["mammo2"]:
            stat[p]["mammo2"] += 1
            present_today.add(p)
        for p in present_today:
            stat[p]["workdays"] += 1

    rows = []
    for p in K_STAFF + F_STAFF + ["J1", "J2"]:
        s = stat[p]
        total = sum(s["counts"]) + s["mammo2"]
        actual = sum(s["counts"])
        ratios = [round(c / actual, 6) if actual and p not in ("J1", "J2") else None for c in s["counts"]]
        rows.append([p] + s["counts"] + [s["mammo2"], total, s["workdays"], actual if p not in ("J1", "J2") else None] + ratios)

    s = stat["R"]
    total = sum(s["counts"]) + s["mammo2"]
    actual = sum(s["counts"])
    ratios = [round(c / actual, 6) if actual else None for c in s["counts"]]
    rows.append(["R (R1~R4)"] + s["counts"] + [s["mammo2"], total, s["workdays"], actual] + ratios)

    return rows, stat


def equity_report(stat):
    reports = []
    groups = [("K (K2·K4)", ["K2", "K4"], 0.15), ("F (F1~F4)", F_STAFF, 0.35)]
    warnings = []
    for label, members, tol in groups:
        table = []
        for i, colname in enumerate(STAT_COLS):
            ratios = {}
            for m in members:
                actual = sum(stat[m]["counts"])
                ratios[m] = stat[m]["counts"][i] / actual if actual else 0.0
            diff = max(ratios.values()) - min(ratios.values()) if ratios else 0.0
            table.append((colname, ratios, diff))
            if diff >= tol:
                hi = max(ratios, key=ratios.get)
                lo = min(ratios, key=ratios.get)
                warnings.append(
                    f"※ {label} 간 {colname} 비율 차이가 큽니다(기준 {tol}) "
                    f"({hi} {ratios[hi]:.2f} 최다 / {lo} {ratios[lo]:.2f} 최소) - "
                    f"다음 달 배정 시 {lo}에게 {colname}를 더 배정하는 것을 고려해 주세요."
                )
        reports.append((label, members, table))
    return reports, warnings


# ============================================================
# 11. 배정 자동 검증 (12장)
# ============================================================


def pm_expected(name, d):
    if name == "J2":
        return False
    if name == "J1":
        return d.weekday() in (0, 4)
    return True


def validate(day_data):
    problems = []
    for d in sorted(day_data):
        if not is_workday(d):
            continue
        entry = day_data[d]
        for session, keys in (("am", ["breast", "mammo", "thyroid", "abus"]), ("pm", ["breast", "mammo", "thyroid", "abus", "mammo2"])):
            seen = {}
            for key in keys:
                for p in entry[session][key]:
                    seen[p] = seen.get(p, 0) + 1
            dups = [p for p, c in seen.items() if c > 1]
            expected = [p for p in ALL_NAMES if not away(p, d, "AM" if session == "am" else "PM")]
            if session == "pm":
                expected = [p for p in expected if pm_expected(p, d)]
            missing = [p for p in expected if p not in seen]
            if dups:
                problems.append(f"{d.month}/{d.day} {'오전' if session=='am' else '오후'} 중복자: {', '.join(dups)}")
            if missing:
                problems.append(f"{d.month}/{d.day} {'오전' if session=='am' else '오후'} 누락자: {', '.join(missing)}")
    return problems


# ============================================================
# 12. 엑셀 출력 (13장)
# ============================================================

FONT_NAME = "맑은 고딕"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
LABEL_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DUTY_ROWS = [
    ("US-locali", "extra_us_locali", None),
    ("Breast US", "am", "breast"),
    ("Mammo", "am", "mammo"),
    ("Thyroid US", "am", "thyroid"),
    ("ABUS", "am", "abus"),
    ("Mammo 2", "extra_blank", None),
    ("VAB", "extra_vab", None),
    ("ST-MMT", "extra_stmmt", None),
    ("Locali", "extra_locali", None),
    ("US", "pm", "breast"),
    ("Mammo", "extra_blank", None),
    ("ABUS", "extra_blank", None),
    ("Thyroid US", "pm", "thyroid"),
    ("Mammo 2", "pm", "mammo2"),
]


def cell_list_for(entry, source, key, day_offset):
    if source == "am" or source == "pm":
        people = list(entry[source][key])
    elif source == "extra_us_locali":
        people = [entry["us_locali"]] if entry["us_locali"] else []
    elif source == "extra_vab":
        people = [entry["vab"]] if entry["vab"] else []
    elif source == "extra_stmmt":
        people = [entry["stmmt"]] if entry["stmmt"] else []
    elif source == "extra_locali":
        people = [entry["locali"]] if entry["locali"] else []
    else:
        people = []
    return [entry["rotation_label"] if p == R_NAME else p for p in people]


def style_header_cell(cell, bold=True, size=10, fill=HEADER_FILL, align="center"):
    cell.font = Font(name=FONT_NAME, bold=bold, size=size)
    cell.fill = fill
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    cell.border = BORDER


def write_main_sheet(ws, weeks, day_data):
    ws.column_dimensions["A"].width = 12
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 9
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 46

    ws["A1"] = f"{YEAR}년 {MONTH}월 유방영상의학과 스케줄 (자동 생성)"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=13)

    row = 3
    for monday in weeks:
        week_dates = [monday + timedelta(days=i) for i in range(6)]
        for i, d in enumerate(week_dates):
            c = ws.cell(row=row, column=2 + i, value=d.day)
            style_header_cell(c)
        style_header_cell(ws.cell(row=row, column=8, value="휴가/출장"), align="left")
        style_header_cell(ws.cell(row=row, column=9, value="문의/주의사항"), align="left")
        style_header_cell(ws.cell(row=row, column=1))
        row += 1
        style_header_cell(ws.cell(row=row, column=1))
        for i in range(6):
            style_header_cell(ws.cell(row=row, column=2 + i, value=WEEKDAY_NAMES[i]))
        row += 1

        vac_lines = []
        note_lines = []
        for d in week_dates:
            for (nm, dt_, period, label) in VACATIONS:
                if dt_ == d:
                    period_kr = {"AM": "오전", "PM": "오후", "ALL": "종일"}[period]
                    vac_lines.append(f"{d.month}/{d.day}({WEEKDAY_NAMES[d.weekday()]}) {nm}: {label}({period_kr})")
            if d in day_data:
                for n in day_data[d]["notes"]:
                    note_lines.append(n)
            if is_holiday(d):
                vac_lines.append(f"{d.month}/{d.day} 공휴일({HOLIDAYS[d]})")

        block_start_row = row
        for label, source, key in DUTY_ROWS:
            counts = []
            for d in week_dates:
                if d not in day_data or not is_workday(d):
                    counts.append(0)
                    continue
                counts.append(len(cell_list_for(day_data[d], source, key, 0)))
            height = max(1, max(counts) if counts else 1)
            for sub in range(height):
                if sub == 0:
                    c = ws.cell(row=row + sub, column=1, value=label)
                    c.font = Font(name=FONT_NAME, bold=True, size=10)
                    c.fill = LABEL_FILL
                    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                for i, d in enumerate(week_dates):
                    val = None
                    if d in day_data and is_workday(d):
                        people = cell_list_for(day_data[d], source, key, 0)
                        if sub < len(people):
                            val = people[sub]
                    cc = ws.cell(row=row + sub, column=2 + i, value=val)
                    cc.font = Font(name=FONT_NAME, size=10)
                    cc.alignment = Alignment(horizontal="center", vertical="center")
                    cc.border = BORDER
            row += height

        max_notes = max(len(vac_lines), len(note_lines), 1)
        needed_rows = row - block_start_row
        for idx in range(max(needed_rows, max_notes)):
            r = block_start_row + idx
            cc = ws.cell(row=r, column=8, value=vac_lines[idx] if idx < len(vac_lines) else None)
            cc.font = Font(name=FONT_NAME, size=10)
            cc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cc2 = ws.cell(row=r, column=9, value=note_lines[idx] if idx < len(note_lines) else None)
            cc2.font = Font(name=FONT_NAME, size=10)
            cc2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        row = max(row, block_start_row + max_notes)
        row += 2

    return row


def write_stats_section(ws, start_row, day_data):
    row = start_row
    ws.cell(row=row, column=1, value="■ 근무 횟수 통계 (자동 집계, duty별 · 월간)").font = Font(name=FONT_NAME, bold=True, size=11)
    row += 1

    headers = ["인원"] + [c.replace(" ", "\n", 1) for c in STAT_COLS] + ["Mammo2\n(오전+오후)", "합계", "근무일수", "실제근무\n(합계-Mammo2)"] + [c + " 비율" for c in STAT_COLS]
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=1 + i, value=h)
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    row += 1

    stat_rows, stat = compute_stats(day_data)
    for r_data in stat_rows:
        for i, v in enumerate(r_data):
            c = ws.cell(row=row, column=1 + i, value=v)
            c.font = Font(name=FONT_NAME, size=10)
            c.alignment = Alignment(horizontal="center", vertical="center")
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="■ 같은 직급 내 형평성 비교 - 실제 근무 횟수(Mammo2 제외) 대비 duty별 비율, J·R 제외").font = Font(name=FONT_NAME, bold=True, size=11)
    row += 1

    reports, warnings = equity_report(stat)
    for label, members, table in reports:
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT_NAME, bold=True, size=10)
        row += 1
        hdr = ["duty"] + members + ["최대-최소 차이"]
        for i, h in enumerate(hdr):
            ws.cell(row=row, column=1 + i, value=h).font = Font(name=FONT_NAME, bold=True, size=10)
        row += 1
        for colname, ratios, diff in table:
            ws.cell(row=row, column=1, value=colname)
            for i, m in enumerate(members):
                ws.cell(row=row, column=2 + i, value=round(ratios[m], 6))
            ws.cell(row=row, column=2 + len(members), value=round(diff, 6))
            row += 1
        row += 1

    for w in warnings:
        ws.cell(row=row, column=1, value=w).font = Font(name=FONT_NAME, size=10, color="C00000")
        row += 1
    row += 1

    ws.cell(row=row, column=1, value="■ US-locali(오전) · Locali(오후) · VAB · ST-MMT 현황 (다른 duty와 중복 배정, 참고용 집계)").font = Font(name=FONT_NAME, bold=True, size=11)
    row += 1
    for i, h in enumerate(["인원", "오전 US-locali", "오후 Locali", "VAB", "ST-MMT"]):
        ws.cell(row=row, column=1 + i, value=h).font = Font(name=FONT_NAME, bold=True, size=10)
    row += 1
    for f in F_STAFF:
        ws.cell(row=row, column=1, value=f)
        ws.cell(row=row, column=2, value=STATE.f_locali_am[f])
        ws.cell(row=row, column=3, value=STATE.f_locali_pm[f])
        row += 1
    for k in K_STAFF:
        ws.cell(row=row, column=1, value=k)
        ws.cell(row=row, column=4, value=STATE.k_vab[k])
        ws.cell(row=row, column=5, value=STATE.k_stmmt[k])
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="■ 배정 자동 검증 결과").font = Font(name=FONT_NAME, bold=True, size=11)
    row += 1
    problems = validate(day_data)
    if not problems:
        ws.cell(row=row, column=1, value="중복자/누락자 없음")
        row += 1
    else:
        for p in problems:
            ws.cell(row=row, column=1, value=p).font = Font(name=FONT_NAME, size=10, color="C00000")
            row += 1

    return row, problems


def write_vip_sheet(ws, weeks, day_data):
    ws["A1"] = f"{YEAR}년 {MONTH}월 VIP 스케줄"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=13)

    row = 3
    for pair_start in range(0, len(weeks), 2):
        pair = weeks[pair_start:pair_start + 2]
        base_cols = [1, 8]
        for wi, monday in enumerate(pair):
            week_dates = [monday + timedelta(days=i) for i in range(5)]
            col0 = base_cols[wi]
            label = f"{week_dates[0].day}~{week_dates[-1].day}"
            style_header_cell(ws.cell(row=row, column=col0, value=label))
            for i, d in enumerate(week_dates):
                style_header_cell(ws.cell(row=row, column=col0 + 1 + i, value=d.day))
            style_header_cell(ws.cell(row=row + 1, column=col0))
            for i in range(5):
                style_header_cell(ws.cell(row=row + 1, column=col0 + 1 + i, value=WEEKDAY_NAMES[i]))

            rows_def = [
                ("오전", lambda e: next((p for p in e["am"]["breast"] if p in K_STAFF), (e["am"]["breast"][0] if e["am"]["breast"] else None))),
                ("오후", lambda e: (e["pm"]["breast"][0] if e["pm"]["breast"] else None)),
                ("갑상선", lambda e: "K1" if "K1" in e["am"]["thyroid"] else None),
            ]
            for ri, (rlabel, getter) in enumerate(rows_def):
                c = ws.cell(row=row + 2 + ri, column=col0, value=rlabel)
                c.font = Font(name=FONT_NAME, bold=True, size=10)
                c.fill = LABEL_FILL
                for i, d in enumerate(week_dates):
                    val = None
                    if d in day_data and is_workday(d):
                        val = getter(day_data[d])
                        if val == R_NAME:
                            val = day_data[d]["rotation_label"]
                    elif d in day_data and is_holiday(d):
                        val = "공휴일"
                    cc = ws.cell(row=row + 2 + ri, column=col0 + 1 + i, value=val)
                    cc.font = Font(name=FONT_NAME, size=10)
                    cc.alignment = Alignment(horizontal="center", vertical="center")
        row += 5


def save_workbook(day_data, weeks):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{YEAR}-{MONTH:02d}"
    end_row = write_main_sheet(ws, weeks, day_data)
    _, problems = write_stats_section(ws, end_row, day_data)

    ws2 = wb.create_sheet("VIP")
    write_vip_sheet(ws2, weeks, day_data)

    base_name = f"{YEAR}-{MONTH:02d}"
    path = os.path.join(OUTPUT_DIR, base_name + ".xlsx")
    n = 1
    while os.path.exists(path):
        path = os.path.join(OUTPUT_DIR, f"{base_name}-{n}.xlsx")
        n += 1
    wb.save(path)
    return path, problems


# ============================================================
# 13. 메인
# ============================================================


def main():
    weeks, day_data = build_schedule()
    path, problems = save_workbook(day_data, weeks)
    print(f"저장 완료: {path}")
    if problems:
        print(f"[검증] 문제 {len(problems)}건 발견:")
        for p in problems:
            print(" -", p)
    else:
        print("[검증] 중복자/누락자 없음")


if __name__ == "__main__":
    main()
