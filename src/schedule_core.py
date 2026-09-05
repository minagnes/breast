# -*- coding: utf-8 -*-
"""
유방영상의학과 근무 스케줄 - 공통 엔진 (schedule_core.py)
================================================================
배정 규칙·스케줄 생성 로직·엑셀 시트 작성을 모두 담은 "단일 진실 공급원".
터미널판(generate_schedule.py)과 웹판(streamlit_app.py)이 이 모듈을 함께 import 한다.
→ 규칙을 바꿀 때는 이 파일 한 곳만 고치면 두 프로그램에 동시에 반영된다.

이 파일은 직접 실행하지 않는다(진입점 아님). 실행은 아래 두 파일로 한다.
    - 터미널: python src/generate_schedule.py
    - 웹    : python -m streamlit run src/streamlit_app.py

기준 문서(SoT): docs/스케줄_배정_규정.docx  ← 규칙을 바꾸면 이 문서도 함께 갱신한다.
    (SoT 문서가 말하는 schedule_generator.py 의 배정 로직이 지금은 이 파일에 들어 있다.)
K4·K5 고정 배정 규칙:
    - K5: 오전 Mammo 화·금 고정 / 오후 Breast US 화·금 고정 / 오후 Mammo2 월·수·목 고정
    - K4: 오후 Mammo2 월·화·금 고정

의존 패키지: openpyxl, holidays (둘 다 없으면 `pip install openpyxl holidays` 로 설치)

파일 저장 위치(어디에 xlsx 를 만들지)는 이 모듈이 정하지 않는다.
이 모듈의 build_vacation_template()·save_workbook() 은 Workbook 객체만 돌려주고,
실제 저장(터미널=디스크, 웹=다운로드)은 각 진입점이 담당한다.
"""

import io
from datetime import date, timedelta

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

# ============================================================
# 0. 고정 설정 (인력 구성 / 배정 규칙 - 사람이 바뀌지 않는 한 그대로 유지)
# ============================================================

K_STAFF = ["K1", "K2", "K3", "K4", "K5"]
F_STAFF = ["F1", "F2", "F3", "F4"]
J_STAFF = ["J1", "J2"]
R_NAME = "R"
ALL_NAMES = K_STAFF + F_STAFF + J_STAFF + [R_NAME]

BREAST_AM_LEAD = {0: "K4", 1: "K2", 2: "K2", 3: "K4", 4: "K4"}
# 특정 날짜에만 오전 Breast US 대표 K를 교체(그 날 원래 대표 K는 자동으로 mammo/abus로 밀림).
# 2026-08-31(월): K3가 9/1부터 출장이라 이 하루만 근무 → K3를 대표로, K4는 Breast US에서 빠짐.
BREAST_AM_LEAD_OVERRIDE = {date(2026, 8, 31): "K3"}
# 특정 날짜에만 '남는 K 중 오전 mammo로 보낼 사람'을 지정(나머지 남는 K는 abus).
# 2026-08-31(월): K3가 Breast 대표로 들어가며 밀려난 K4를 abus 대신 mammo로.
AM_MAMMO_PREFER = {date(2026, 8, 31): "K4"}
BREAST_AM_EXTRA = {0: "K5", 2: "K5", 3: "K5"}
K5_AM_MAMMO_DAYS = {1, 4}            # K5: 오전 Mammo 화·금 고정

MAMMO2_PM_FIX = {
    0: ["K1", "K5", "K4"],  # 월 (K5·K4 고정)
    1: ["K1", "K2", "K5"],  # 화 (K5 고정)
    2: ["K1", "K2"],        # 수 (K4는 목으로 이동 — K4 Mammo2는 월·목·금)
    3: ["K5", "K4"],        # 목 (K5·K4 고정)
    4: ["K1", "K2", "K4"],  # 금 (K4 고정)
}
# 오후 Breast US 지정자 (Mammo2 비고정 K 중) — 위 Mammo2 고정에서 빠지는 K가 자동으로 배정됨
BREAST_PM_FIX = {1: "K4", 2: "K5", 4: "K5"}

WEEKDAY_NAMES = ["월", "화", "수", "목", "금", "토"]

# ------------------------------------------------------------
# 전공의(R) 로테이션: 28일(4주) 단위 '텀'마다 다른 전공의가 근무한다.
# 각 텀은 일요일 시작·토요일 끝이라 스케줄 주(월~토) 경계와 정확히 맞는다.
# (텀 시작일, 라벨) 표를 정답 소스로 쓴다. 라벨은 R# 대신 이니셜 등으로 바꿔도 된다.
# 새 로테이션표가 나오면 이 표 13줄을 갱신/추가하면 된다(첨부 연간표 = 2026 사이클).
# ------------------------------------------------------------
R_TERMS = [
    (date(2026, 3, 1),   "R1"),    # 1텀  3.1~3.28
    (date(2026, 3, 29),  "R2"),    # 2텀  3.29~4.25
    (date(2026, 4, 26),  "R3"),    # 3텀  4.26~5.23
    (date(2026, 5, 24),  "R4"),    # 4텀  5.24~6.20
    (date(2026, 6, 21),  "R5"),    # 5텀  6.21~7.18
    (date(2026, 7, 19),  "R6"),    # 6텀  7.19~8.15
    (date(2026, 8, 16),  "R7"),    # 7텀  8.16~9.12
    (date(2026, 9, 13),  "R8"),    # 8텀  9.13~10.10
    (date(2026, 10, 11), "R9"),    # 9텀  10.11~11.7
    (date(2026, 11, 8),  "R10"),   # 10텀 11.8~12.5
    (date(2026, 12, 6),  "R11"),   # 11텀 12.6~1.2
    (date(2027, 1, 3),   "R12"),   # 12텀 1.3~1.30
    (date(2027, 1, 31),  "R13"),   # 13텀 1.31~2.27
]
R_TERM_DAYS = 28                   # 한 텀 길이(일)
R_TERM_COUNT = 13                  # 한 사이클의 텀 수
R_CYCLE_ANCHOR = R_TERMS[0][0]     # 표 밖 날짜용 폴백 기준(1텀 시작)

MANUAL_ADJUSTMENTS = []

FONT_NAME = "맑은 고딕"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
LABEL_FILL = PatternFill("solid", fgColor="F2F2F2")
SAT_FILL = PatternFill("solid", fgColor="F2F2F2")
HOLIDAY_FILL = PatternFill("solid", fgColor="FCE4D6")
MANUAL_FILL = PatternFill("solid", fgColor="FFF2CC")   # 수동 배정일 표시색(연노랑)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 휴가 입력 드롭다운 선택지 (자유 텍스트 입력도 가능)
VAC_CHOICES = ["", "휴가(종일)", "휴가(오전)", "휴가(오후)", "학회(종일)", "학회(오전)", "학회(오후)",
               "출장(종일)", "출장(오전)", "출장(오후)", "교육(종일)", "반차(오전)", "반차(오후)"]

# ------------------------------------------------------------
# 특일(수동/임시공휴일) 지정용 '제어행'
# vacation 엑셀 맨 아래 '◆특일지정' 행에 날짜별로 입력한다.
#   - '수동배정'  → 그 날은 규칙 배정을 건너뛰고 전부 빈칸으로 둔다
#                   (학회 등 대부분이 자리를 비워 규칙대로 짤 수 없는 날).
#   - '임시공휴일' → 그 날을 공휴일로 취급한다
#                   (holidays 라이브러리에 없던 갑작스런 임시공휴일).
# ------------------------------------------------------------
CONTROL_ROW_LABEL = "◆특일지정"
CONTROL_CHOICES = ["", "수동배정", "임시공휴일"]


def classify_control(text):
    """제어행 셀 텍스트 → 'MANUAL' / 'HOLIDAY' / None."""
    t = str(text).strip().lower()
    if not t:
        return None
    if ("수동" in t) or ("manual" in t):
        return "MANUAL"
    if ("공휴일" in t) or ("휴일" in t) or ("holiday" in t):
        return "HOLIDAY"
    return None

# ------------------------------------------------------------
# 엔진이 읽는 전역 상태 (모듈 로드시엔 빈 값).
# 진입점(generate_schedule.py / streamlit_app.py)이 스케줄 생성 직전에
# schedule_core.HOLIDAYS / .VACATIONS / .VAC_INDEX 로 직접 대입해 채운다.
# ※ from-import 로 가져와 재대입하면 엔진이 못 보므로 반드시 모듈 속성으로 설정할 것.
# ------------------------------------------------------------
YEAR = None
MONTH = None
HOLIDAYS = {}
VACATIONS = []
VAC_INDEX = {}
MANUAL_DAYS = set()   # 규칙 배정을 건너뛰고 빈칸으로 둘 '수동 배정일' (제어행에서 채움)


def vab_stmmt_items(month):
    if month % 2 == 1:
        return "VAB", "ST-MMT"
    return "ST-MMT", "VAB"


# ============================================================
# 1. 달력 / 휴가 유틸
# ============================================================


def month_weeks(year, month):
    """그 주의 수요일이 해당 월에 속하는 모든 주의 월요일 리스트."""
    probe = date(year, month, 15) - timedelta(days=40)
    mondays = []
    d = probe - timedelta(days=probe.weekday())
    for _ in range(16):
        wed = d + timedelta(days=2)
        if wed.year == year and wed.month == month:
            mondays.append(d)
        d += timedelta(days=7)
    return sorted(set(mondays))


def all_dates(year, month):
    dates = []
    for monday in month_weeks(year, month):
        dates.extend(monday + timedelta(days=i) for i in range(6))  # 월~토
    return dates


def is_holiday(d):
    return d in HOLIDAYS


def is_manual(d):
    return d in MANUAL_DAYS


def is_workday(d):
    return d.weekday() != 5 and not is_holiday(d) and not is_manual(d)


def away(name, d, session):
    for period, _label in VAC_INDEX.get((name, d), []):
        if period == "ALL" or period == session:
            return True
    return False


def rotation_label_for_monday(monday):
    """그 주(월요일)가 속한 전공의 로테이션 텀의 라벨을 돌려준다.
    먼저 R_TERMS 표에서 찾고, 표 범위를 벗어나면 28일×13텀 순환으로 근사한다."""
    first_start = R_TERMS[0][0]
    last_end = R_TERMS[-1][0] + timedelta(days=R_TERM_DAYS)
    if first_start <= monday < last_end:
        label = R_TERMS[0][1]
        for start, lbl in R_TERMS:
            if start <= monday:
                label = lbl
            else:
                break
        return label
    # 표 밖: 1텀 시작 기준으로 28일마다 R1~R13 순환(라벨만 근사, 필요시 표 갱신).
    idx = ((monday - R_CYCLE_ANCHOR).days // R_TERM_DAYS) % R_TERM_COUNT
    return f"R{idx + 1}"


def get_holidays_for_month(year, month):
    """holidays 라이브러리로 대한민국 공휴일(대체공휴일 포함)을 자동 조회해
    {date: 라벨} 딕셔너리로 돌려준다. 이 달력에 표시되는 전/다음달 며칠도 함께 포함한다."""
    try:
        import holidays as holidays_lib
    except ImportError:
        # 표시 방법(터미널 종료 / 웹 오류창)은 진입점이 정하도록 RuntimeError 로 알린다.
        raise RuntimeError(
            "공휴일 자동 조회에 필요한 'holidays' 패키지가 설치되어 있지 않습니다.\n"
            "터미널에서 아래 명령을 실행한 뒤 다시 실행하세요:\n"
            "    pip install holidays"
        )

    dates = all_dates(year, month)
    years_needed = sorted({d.year for d in dates})
    try:
        kr = holidays_lib.KR(years=years_needed, language="ko")
    except TypeError:
        kr = holidays_lib.KR(years=years_needed)

    return {d: kr[d] for d in dates if d in kr}


def parse_period_label(text):
    """'휴가(오전)' → ('AM', '휴가'). 구분 표기가 없으면 종일(ALL)로 처리."""
    text = str(text).strip()
    for kr, code in (("오전", "AM"), ("오후", "PM"), ("종일", "ALL")):
        marker = f"({kr})"
        if marker in text:
            label = text.replace(marker, "").strip()
            return code, (label or text)
    return "ALL", text


# ============================================================
# 2. 휴가 입력 엑셀(vacation_YYYY-MM.xlsx) 생성 / 파싱
# ============================================================


def build_vacation_template(year, month, holidays):
    """휴가 입력표 Workbook 을 만들어 돌려준다 (파일로 저장하지 않음).
    저장은 진입점이 담당한다(터미널=디스크, 웹=다운로드)."""
    dates = all_dates(year, month)

    wb = Workbook()
    ws = wb.active
    ws.title = "vacation"

    ws["A1"] = f"{year}년 {month}월 휴가/부재 입력표 (vacation)"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=13)
    ws["A2"] = ("칸에 직접 입력하거나 드롭다운에서 선택 (예: 휴가(오전)). 빈칸 = 정상 근무  |  "
                "맨 아래 ◆특일지정 행: 수동배정 / 임시공휴일")
    ws["A2"].font = Font(name=FONT_NAME, size=9, italic=True)

    header_row1, header_row2, first_data_row = 4, 5, 6

    ws.cell(row=header_row1, column=1, value="근무자")
    for r in (header_row1, header_row2):
        c = ws.cell(row=r, column=1)
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER

    for i, d in enumerate(dates):
        col = 2 + i
        is_sat = d.weekday() == 5
        is_hol = d in holidays
        fill = HOLIDAY_FILL if is_hol else (SAT_FILL if is_sat else HEADER_FILL)

        c1 = ws.cell(row=header_row1, column=col, value=d)   # 실제 date 객체를 저장 (되읽기용)
        c1.number_format = "m/d"
        c1.font = Font(name=FONT_NAME, bold=True, size=10)
        c1.fill = fill
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = BORDER

        wd_label = WEEKDAY_NAMES[d.weekday()]
        if is_hol:
            wd_label += f"·{holidays[d]}"
        c2 = ws.cell(row=header_row2, column=col, value=wd_label)
        c2.font = Font(name=FONT_NAME, bold=True, size=9)
        c2.fill = fill
        c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c2.border = BORDER

        ws.column_dimensions[c1.column_letter].width = 9

    for r, name in enumerate(ALL_NAMES, start=first_data_row):
        c = ws.cell(row=r, column=1, value=name)
        c.font = Font(name=FONT_NAME, bold=True, size=10)
        c.fill = LABEL_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
        for i, d in enumerate(dates):
            col = 2 + i
            cc = ws.cell(row=r, column=col)
            cc.font = Font(name=FONT_NAME, size=10)
            cc.alignment = Alignment(horizontal="center", vertical="center")
            cc.border = BORDER
            if d.weekday() == 5:
                cc.fill = SAT_FILL
            elif d in holidays:
                cc.fill = HOLIDAY_FILL

    dv = DataValidation(type="list", formula1='"' + ",".join(VAC_CHOICES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    last_row = first_data_row + len(ALL_NAMES) - 1
    last_col_letter = ws.cell(row=header_row1, column=1 + len(dates)).column_letter
    dv.add(f"B{first_data_row}:{last_col_letter}{last_row}")

    # --- 특일 지정(제어) 행: '수동배정' / '임시공휴일' --------------------
    control_row = last_row + 1
    cc = ws.cell(row=control_row, column=1, value=CONTROL_ROW_LABEL)
    cc.font = Font(name=FONT_NAME, bold=True, size=10)
    cc.fill = MANUAL_FILL
    cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cc.border = BORDER
    for i, d in enumerate(dates):
        c = ws.cell(row=control_row, column=2 + i)
        c.font = Font(name=FONT_NAME, size=10)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
        if d.weekday() == 5:
            c.fill = SAT_FILL
        elif d in holidays:
            c.fill = HOLIDAY_FILL

    dv2 = DataValidation(type="list", formula1='"' + ",".join(CONTROL_CHOICES) + '"', allow_blank=True)
    ws.add_data_validation(dv2)
    dv2.add(f"B{control_row}:{last_col_letter}{control_row}")

    ws.cell(
        row=control_row + 2, column=1,
        value="◆특일지정 사용법:  '수동배정'=그 날은 자동배정을 건너뛰고 전부 빈칸(학회 등 대부분 부재일, 직접 작성) · "
              "'임시공휴일'=그 날을 공휴일로 처리",
    ).font = Font(name=FONT_NAME, size=9, italic=True)

    ws.column_dimensions["A"].width = 10
    ws.freeze_panes = ws.cell(row=first_data_row, column=2)
    return wb


def parse_vacation_file(src, year, month, warn=print):
    """(구버전 호환) 사람별 부재 리스트만 돌려준다.
    manual_days/extra_holidays 까지 필요하면 parse_vacation_workbook() 을 쓴다."""
    vacations, _manual, _extra = parse_vacation_workbook(src, year, month, warn=warn)
    return vacations


def parse_vacation_workbook(src, year, month, warn=print):
    """작성된 vacation 엑셀을 '한 번만' 읽어 다음 3가지를 돌려준다.
        vacations      : [(이름, date, 구분, 라벨), ...]      ← 사람별 부재
        manual_days    : set[date]                           ← ◆특일지정=수동배정
        extra_holidays : dict[date, str]                     ← ◆특일지정=임시공휴일
    src 는 파일 경로 문자열도, 업로드된 파일 객체도 모두 받는다.
    (업로드 스트림을 두 번 열면 소진될 수 있으므로 반드시 한 번만 읽는다.)
    경고는 warn 콜백으로 낸다(터미널=print, 웹=st.warning)."""
    wb = load_workbook(src, data_only=True)
    ws = wb["vacation"] if "vacation" in wb.sheetnames else wb.active
    dates = all_dates(year, month)
    first_data_row = 6

    if ws.max_column < 1 + len(dates):
        warn(f"[경고] 파일의 열 개수가 예상({len(dates)}개)보다 적습니다. "
             f"연/월이 휴가표를 만들 때와 같은지, 날짜 헤더가 바뀌지 않았는지 확인하세요.")

    scan_rows = range(first_data_row, first_data_row + len(ALL_NAMES) + 12)
    name_to_row = {}
    control_row = None
    for r in scan_rows:
        label = ws.cell(row=r, column=1).value
        if label in ALL_NAMES:
            name_to_row[label] = r
        elif label == CONTROL_ROW_LABEL:
            control_row = r

    missing_names = [n for n in ALL_NAMES if n not in name_to_row]
    if missing_names:
        warn(f"[경고] vacation 엑셀에서 다음 근무자를 찾지 못했습니다: {', '.join(missing_names)}")

    vacations = []
    for name, r in name_to_row.items():
        for i, d in enumerate(dates):
            val = ws.cell(row=r, column=2 + i).value
            if val is None:
                continue
            text = str(val).strip()
            if not text:
                continue
            period, label = parse_period_label(text)
            vacations.append((name, d, period, label))

    manual_days, extra_holidays = set(), {}
    if control_row is not None:
        for i, d in enumerate(dates):
            kind = classify_control(ws.cell(row=control_row, column=2 + i).value)
            if kind == "MANUAL":
                manual_days.add(d)
            elif kind == "HOLIDAY":
                extra_holidays[d] = "임시공휴일"
    else:
        warn("[안내] ◆특일지정 행을 찾지 못했습니다(구버전 휴가표일 수 있음). "
             "수동배정/임시공휴일 기능 없이 진행합니다.")

    return vacations, manual_days, extra_holidays


# ============================================================
# 3. 상태(월간 누적 카운터)
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
# 4. 오전(AM) 배정 (6장)
# ============================================================


def assign_am(d, is_first_workday):
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

    lead = BREAST_AM_LEAD_OVERRIDE.get(d, BREAST_AM_LEAD[wd])
    extra = BREAST_AM_EXTRA.get(wd)
    if lead in avail_K:
        duties["breast"].append(lead)
        used_K.add(lead)
    else:
        # 주 담당 K가 없을 때의 단독 대체 후보.
        # K4·K5는 다른 K 없이도 오전 Breast US를 단독으로 맡을 수 있다(자격자).
        # → 예전엔 K5를 자기 Mammo 요일(화·금)에 제외했으나, 이제 제외하지 않아
        #   그 요일에 주 담당이 비면 K5가 혼자 Breast US를 커버할 수 있다.
        pool = [
            k for k in avail_K
            if k not in used_K and k != extra and k != lead
        ]
        if pool:
            sub = pool[0]
            duties["breast"].append(sub)
            used_K.add(sub)
            notes.append(f"{d.month}/{d.day} {lead} 오전 휴가 → {sub}(으)로 Breast US 대체")
    if extra and extra in avail_K and extra not in used_K:
        duties["breast"].append(extra)
        used_K.add(extra)

    remaining_K = [k for k in avail_K if k not in used_K]
    # 특정 날짜에만 '남는 K 중 mammo로 보낼 사람'을 지정(나머지는 abus).
    prefer_mammo = AM_MAMMO_PREFER.get(d)
    if prefer_mammo in remaining_K:
        remaining_K = [prefer_mammo] + [k for k in remaining_K if k != prefer_mammo]
    if wd in K5_AM_MAMMO_DAYS and "K5" in remaining_K:
        duties["mammo"].append("K5")
        for k in remaining_K:
            if k != "K5":
                duties["abus"].append(k)
    elif remaining_K:
        duties["mammo"].append(remaining_K[0])
        for k in remaining_K[1:]:
            duties["abus"].append(k)

    # AM 정원(최대 인원): Breast US 최대 5명, Thyroid US 최대 4명 (요일 무관)
    breast_cap = 5
    thyroid_cap = 4
    mammo_cap = 2

    def slot_open(cat):
        if cat == "breast":
            return len(duties["breast"]) < breast_cap
        if cat == "thyroid":
            return len(duties["thyroid"]) < thyroid_cap
        if cat == "mammo":
            return len(duties["mammo"]) < mammo_cap and len(duties["mammo"]) >= 1
        return False

    # J 배정을 F·R보다 '먼저' 처리해 자리를 선점한다.
    # → 이후 F·R 채움이 정원 안에서만 이뤄져 J를 포함해도 Breast US 5명 / Thyroid US 4명을
    #   넘지 않는다(정원은 최대값이며, 인원이 적어 그보다 작아지는 것은 정상).
    # 규정: J1·J2가 모두 있으면 서로 다른 duty에, 한 명만 있으면 자기 누적이 적은 쪽에 배정.
    if len(avail_J) == 2:
        j_a, j_b = avail_J
        cat_a = min(("breast", "thyroid"), key=lambda c: STATE.j_am[j_a][c])
        cat_b = "thyroid" if cat_a == "breast" else "breast"
        duties[cat_a].append(j_a)
        duties[cat_b].append(j_b)
        STATE.j_am[j_a][cat_a] += 1
        STATE.j_am[j_b][cat_b] += 1
    elif len(avail_J) == 1:
        j = avail_J[0]
        cat = min(("breast", "thyroid"), key=lambda c: STATE.j_am[j][c])
        duties[cat].append(j)
        STATE.j_am[j][cat] += 1

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
        if chosen:
            duties[chosen].append(R_NAME)
            STATE.r_am[chosen] += 1
            STATE.r_am_week[chosen] += 1
            if chosen == "mammo" and duties["mammo"]:
                STATE.r_mammo_partners.add(duties["mammo"][0])
        else:
            notes.append(f"{d.month}/{d.day} R 오전 자리 부족 → 배정 실패(정원 초과 확인 필요)")

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
# 5. 오후(PM) 배정 (7장)
# ============================================================


def assign_pm(d):
    wd = d.weekday()
    notes = []

    avail_K = [k for k in K_STAFF if not away(k, d, "PM")]
    avail_F = [f for f in F_STAFF if not away(f, d, "PM")]
    avail_R = [] if away(R_NAME, d, "PM") else [R_NAME]
    avail_J1 = not away("J1", d, "PM")

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
    # 금요일 오후 J 근무 없음(삭제됨). 월요일 오후만 근무.

    return duties, notes, avail_F, avail_R


def finish_pm_f_and_r(duties, avail_F, avail_R, f_mammo2_today):
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


def balance_pm(duties, wd):
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
# 6. VAB / ST-MMT / Locali (8장)
# ============================================================


def assign_extras_pm(d, pm_duties, monwed_item, tuethu_item):
    wd = d.weekday()
    breast_K = [p for p in pm_duties["breast"] if p in K_STAFF]
    breast_F = [p for p in pm_duties["breast"] if p in F_STAFF]

    vab = stmmt = locali = None
    item = None
    if wd in (0, 2) and breast_K:
        item = monwed_item
    elif wd in (1, 3) and breast_K:
        item = tuethu_item

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
# 7. 주 단위 F Mammo2 순환 계획 (7.3)
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
# 8. 하루 전체 조립 + 수동 보정 적용
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


def build_schedule(year, month):
    monwed_item, tuethu_item = vab_stmmt_items(month)
    weeks = month_weeks(year, month)
    day_data = {}

    for monday in weeks:
        week_dates = [monday + timedelta(days=i) for i in range(6)]
        workdays = [d for d in week_dates if is_workday(d)]
        first_workday = workdays[0] if workdays else None

        f_mammo2_plan = plan_week_f_mammo2(week_dates)
        STATE.r_am_week = new_counter(["breast", "thyroid", "mammo"])

        for d in week_dates:
            entry = {
                "am": {"breast": [], "mammo": [], "thyroid": [], "abus": []},
                "pm": {"breast": [], "mammo": [], "thyroid": [], "abus": [], "mammo2": []},
                "us_locali": None, "vab": None, "stmmt": None, "locali": None,
                "notes": [],
                "rotation_label": rotation_label_for_monday(monday),
            }
            if is_workday(d):
                am_duties, us_locali, am_notes = assign_am(d, d == first_workday)
                pm_duties, pm_notes, avail_F, avail_R = assign_pm(d)
                pm_duties = finish_pm_f_and_r(pm_duties, avail_F, avail_R, f_mammo2_plan.get(d, []))
                pm_duties = balance_pm(pm_duties, d.weekday())
                vab, stmmt, locali = assign_extras_pm(d, pm_duties, monwed_item, tuethu_item)

                entry["am"] = am_duties
                entry["pm"] = pm_duties
                entry["us_locali"] = us_locali
                entry["vab"] = vab
                entry["stmmt"] = stmmt
                entry["locali"] = locali
                entry["notes"] = am_notes + pm_notes
            elif is_manual(d):
                entry["notes"].append("※ 수동 배정일: 자동배정을 건너뜁니다. 직접 채워 주세요.")
            day_data[d] = entry

    apply_manual_adjustments(day_data)
    return weeks, day_data


# ============================================================
# 9. 통계 / 형평성 (10장)
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
    rows.append(["R (R1~R13)"] + s["counts"] + [s["mammo2"], total, s["workdays"], actual] + ratios)

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
# 10. 배정 자동 검증 (12장)
# ============================================================


def pm_expected(name, d):
    if name == "J2":
        return False
    if name == "J1":
        return d.weekday() == 0   # J1 오후 근무는 월요일만(금요일 오후 삭제됨)
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
# 11. 본 스케줄 / VIP 엑셀 출력 (13장)
# ============================================================

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


def cell_list_for(entry, source, key):
    if source in ("am", "pm"):
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


def write_main_sheet(ws, year, month, weeks, day_data):
    ws.column_dimensions["A"].width = 12
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 9
    ws.column_dimensions["H"].width = 34
    ws.column_dimensions["I"].width = 46

    ws["A1"] = f"{year}년 {month}월 유방영상의학과 스케줄 (자동 생성)"
    ws["A1"].font = Font(name=FONT_NAME, bold=True, size=13)

    row = 3
    for monday in weeks:
        week_dates = [monday + timedelta(days=i) for i in range(6)]
        for i, d in enumerate(week_dates):
            style_header_cell(ws.cell(row=row, column=2 + i, value=d.day))
        style_header_cell(ws.cell(row=row, column=8, value="휴가/출장"), align="left")
        style_header_cell(ws.cell(row=row, column=9, value="문의/주의사항"), align="left")
        style_header_cell(ws.cell(row=row, column=1))
        row += 1
        style_header_cell(ws.cell(row=row, column=1))
        for i in range(6):
            d = week_dates[i]
            hc = ws.cell(row=row, column=2 + i, value=WEEKDAY_NAMES[i])
            if is_manual(d):
                hc.value = WEEKDAY_NAMES[i] + "(수동)"
                style_header_cell(hc, fill=MANUAL_FILL)
            elif is_holiday(d):
                hc.value = WEEKDAY_NAMES[i] + "(휴일)"
                style_header_cell(hc, fill=HOLIDAY_FILL)
            else:
                style_header_cell(hc)
        row += 1

        vac_lines, note_lines = [], []
        for d in week_dates:
            for (nm, dt_, period, label) in VACATIONS:
                if dt_ == d:
                    period_kr = {"AM": "오전", "PM": "오후", "ALL": "종일"}[period]
                    vac_lines.append(f"{d.month}/{d.day}({WEEKDAY_NAMES[d.weekday()]}) {nm}: {label}({period_kr})")
            if d in day_data:
                note_lines.extend(day_data[d]["notes"])
            if is_holiday(d):
                vac_lines.append(f"{d.month}/{d.day} 공휴일({HOLIDAYS[d]})")

        block_start_row = row
        for label, source, key in DUTY_ROWS:
            counts = []
            for d in week_dates:
                if d not in day_data or not is_workday(d):
                    counts.append(0)
                    continue
                counts.append(len(cell_list_for(day_data[d], source, key)))
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
                        people = cell_list_for(day_data[d], source, key)
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


def write_vip_sheet(ws, year, month, weeks, day_data):
    ws["A1"] = f"{year}년 {month}월 VIP 스케줄"
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
                    elif d in day_data and is_manual(d):
                        val = "수동"
                    cc = ws.cell(row=row + 2 + ri, column=col0 + 1 + i, value=val)
                    cc.font = Font(name=FONT_NAME, size=10)
                    cc.alignment = Alignment(horizontal="center", vertical="center")
        row += 5


def save_workbook(year, month, day_data, weeks):
    """본 스케줄 + VIP 시트가 담긴 Workbook 을 만들어 (wb, problems) 로 돌려준다.
    실제 저장 위치는 진입점이 정한다(터미널=outputs 디스크, 웹=브라우저 다운로드)."""
    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}-{month:02d}"
    end_row = write_main_sheet(ws, year, month, weeks, day_data)
    _, problems = write_stats_section(ws, end_row, day_data)

    ws2 = wb.create_sheet("VIP")
    write_vip_sheet(ws2, year, month, weeks, day_data)

    return wb, problems


def to_bytes(wb):
    """Workbook 을 바이트로 직렬화한다(웹 다운로드 버튼용)."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def set_vacations(vacations):
    """엔진이 읽는 전역 VACATIONS / VAC_INDEX 를 한 번에 설정하는 도우미.
    진입점은 반드시 이 함수(또는 schedule_core.VACATIONS 직접 대입)로 채운다."""
    VAC_INDEX_local = {}
    for (nm, dt_, period, label) in vacations:
        VAC_INDEX_local.setdefault((nm, dt_), []).append((period, label))
    # 모듈 전역에 직접 대입 (엔진 함수들이 이 이름을 참조한다)
    globals()["VACATIONS"] = list(vacations)
    globals()["VAC_INDEX"] = VAC_INDEX_local
