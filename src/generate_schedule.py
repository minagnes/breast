# -*- coding: utf-8 -*-
"""
유방영상의학과 근무 스케줄 생성기 - 터미널판 (generate_schedule.py)
================================================================
배정 규칙·스케줄 로직은 모두 schedule_core.py 에 있고, 이 파일은
터미널(검은 화면)에서 사용자와 대화하며 그 엔진을 실행하는 "얇은 껍데기"다.

실행:  python src/generate_schedule.py

한 번 실행으로 아래 4단계가 순서대로 진행된다.
    1) 연도/월 입력 → 공휴일 자동 조회 → outputs/vacation_YYYY-MM.xlsx 생성(있으면 재사용)
    2) 그 파일을 열어 각자의 휴가·학회·출장을 입력하고 "같은 위치에 그대로" 저장
    3) 터미널에 Y 입력
    4) 휴가를 반영한 outputs/YYYY-MM.xlsx(본 스케줄 + VIP 시트) 생성 후 종료

결과물은 모두 이 프로젝트의 outputs/ 폴더(레포 루트)에 저장된다.
"""

import os

import schedule_core as core
from schedule_core import (
    get_holidays_for_month,
    build_vacation_template,
    parse_vacation_file,
    build_schedule,
    save_workbook,
)

# 결과물 저장 폴더: 이 스크립트(src/)의 부모 = 레포 루트, 그 아래 outputs/
OUTPUT_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "outputs"
)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _unique_path(base_name):
    """outputs 안에서 겹치지 않는 파일 경로를 만든다(있으면 -1, -2 … 를 붙임)."""
    path = os.path.join(OUTPUT_DIR, base_name + ".xlsx")
    n = 1
    while os.path.exists(path):
        path = os.path.join(OUTPUT_DIR, f"{base_name}-{n}.xlsx")
        n += 1
    return path


def main():
    print("=== 유방영상의학과 스케줄 생성기 ===")

    # 1) 연도/월 입력 → 공휴일 자동 조회 → 휴가 입력표 즉시 생성
    year = int(input("연도 (예: 2026): ").strip())
    month = int(input("월 (예: 9): ").strip())

    try:
        core.HOLIDAYS = get_holidays_for_month(year, month)
    except RuntimeError as e:
        raise SystemExit(str(e))

    if core.HOLIDAYS:
        print("자동 조회된 공휴일: " + ", ".join(
            f"{d.month}/{d.day}({name})" for d, name in sorted(core.HOLIDAYS.items())
        ))
    else:
        print("자동 조회 결과 해당 기간에 공휴일이 없습니다.")

    vac_path = os.path.join(OUTPUT_DIR, f"vacation_{year}-{month:02d}.xlsx")
    vac_full = os.path.abspath(vac_path)

    if os.path.exists(vac_path):
        print("\n휴가 입력표가 이미 있습니다(기존 파일을 그대로 재사용합니다).")
    else:
        wb = build_vacation_template(year, month, core.HOLIDAYS)
        wb.save(vac_path)
        print("\n휴가 입력표를 새로 만들었습니다.")

    print(f"[휴가 입력표] {vac_full}")

    # 2) 사용자가 엑셀을 열어 휴가/출장 일정을 작성하고 저장하도록 안내
    print("위 경로의 파일을 열어 K1~K5(K3 제외)/F1~F4/J1·J2/R 각자의 휴가·학회·출장 일정을")
    print("해당 날짜 칸에 입력한 뒤 저장하세요.")
    print("※ 위 경로의 파일을 같은 위치에 그대로 저장하세요(다른 이름/폴더로 저장하면 못 읽음).")

    # 3) 저장 완료를 터미널에서 Y로 확인받을 때까지 대기
    while True:
        ans = input("작성 후 저장하셨으면 Y를 누르세요 (Y 입력 시 스케줄 생성 진행): ").strip().lower()
        if ans == "y":
            break

    # 4) vacation 엑셀을 다시 읽어 스케줄 생성
    vacations = parse_vacation_file(vac_path, year, month)
    core.set_vacations(vacations)
    print(f"\n{os.path.basename(vac_path)} 에서 휴가/부재 {len(vacations)}건을 읽었습니다.")

    weeks, day_data = build_schedule(year, month)
    wb, problems = save_workbook(year, month, day_data, weeks)

    out_path = _unique_path(f"{year}-{month:02d}")
    wb.save(out_path)

    print("\n저장 완료.")
    print(f"[완성 스케줄] {os.path.abspath(out_path)}")
    print(f"[outputs 폴더] {os.path.abspath(OUTPUT_DIR)}")
    if problems:
        print(f"[검증] 문제 {len(problems)}건 발견:")
        for p in problems:
            print(" -", p)
    else:
        print("[검증] 중복자/누락자 없음")


if __name__ == "__main__":
    main()
