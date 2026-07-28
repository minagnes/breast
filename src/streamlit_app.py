# -*- coding: utf-8 -*-
"""
유방영상의학과 근무 스케줄 생성기 - 웹판 (streamlit_app.py)
================================================================
배정 규칙·스케줄 로직은 모두 schedule_core.py 에 있고, 이 파일은
터미널판의 4단계를 Streamlit 화면(버튼·업로드·다운로드)으로 바꾼 "얇은 껍데기"다.

실행:  python -m streamlit run src/streamlit_app.py

화면 흐름 (터미널판의 4단계와 동일)
    1) 연도/월 입력 → [휴가 입력표 생성] → [휴가 입력표 다운로드] 로 vacation_YYYY-MM.xlsx 받기
    2) 받은 엑셀을 열어 각자의 휴가·학회·출장을 입력하고 저장
    3) 저장한 파일을 업로드
    4) [스케줄 생성] → 완성 스케줄 다운로드(브라우저 다운로드 폴더에 저장) + 검증 결과 표시
"""

import streamlit as st

import schedule_core as core
from schedule_core import (
    get_holidays_for_month,
    build_vacation_template,
    parse_vacation_workbook,
    build_schedule,
    save_workbook,
    to_bytes,
)


def render_app():
    st.set_page_config(page_title="유방영상의학과 스케줄 생성기", layout="wide")
    st.title("유방영상의학과 근무 스케줄 생성기")
    st.caption("기준 문서: 스케줄_배정_규정.docx · K4/K5 고정 배정 규칙 반영")

    col1, col2 = st.columns(2)
    with col1:
        year_input = st.number_input("연도", min_value=2020, max_value=2100, value=2026, step=1)
    with col2:
        month_input = st.number_input("월", min_value=1, max_value=12, value=9, step=1)

    st.markdown("### 1단계 · 휴가 입력표 만들기")
    if st.button("휴가 입력표 생성", type="primary"):
        year_i, month_i = int(year_input), int(month_input)
        try:
            holidays_found = get_holidays_for_month(year_i, month_i)
        except RuntimeError as e:
            st.error(str(e))
            st.stop()
        wb = build_vacation_template(year_i, month_i, holidays_found)

        st.session_state["year"] = year_i
        st.session_state["month"] = month_i
        st.session_state["holidays"] = holidays_found
        st.session_state["vacation_bytes"] = to_bytes(wb)

        if holidays_found:
            st.success("자동 조회된 공휴일: " + ", ".join(
                f"{d.month}/{d.day}({name})" for d, name in sorted(holidays_found.items())
            ))
        else:
            st.info("자동 조회 결과 해당 기간에 공휴일이 없습니다.")

    if "vacation_bytes" in st.session_state:
        y_, m_ = st.session_state["year"], st.session_state["month"]
        st.download_button(
            f"휴가 입력표 다운로드 (vacation_{y_}-{m_:02d}.xlsx)",
            data=st.session_state["vacation_bytes"],
            file_name=f"vacation_{y_}-{m_:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            "다운로드한 파일을 열어 K1·K2·K4·K5 / F1~F4 / J1·J2 / R 각자의 휴가·학회·출장을 "
            "해당 날짜 칸에 입력(드롭다운 선택 가능)하고 저장한 뒤, 아래 2단계에 업로드하세요."
        )

    st.markdown("---")
    st.markdown("### 2단계 · 작성한 휴가표 업로드 후 스케줄 생성")
    uploaded = st.file_uploader("작성을 마친 vacation 엑셀을 업로드하세요", type=["xlsx"])

    generate_clicked = st.button("스케줄 생성", type="primary", disabled=uploaded is None)

    if generate_clicked:
        if "year" not in st.session_state:
            st.error("먼저 1단계에서 휴가 입력표를 생성하세요(연도/월이 확정되어야 합니다).")
        else:
            year_i = st.session_state["year"]
            month_i = st.session_state["month"]

            # 휴가표를 한 번만 읽어 사람별 부재 + 특일지정(수동/임시공휴일)을 함께 파싱
            vacations, manual_days, extra_holidays = parse_vacation_workbook(
                uploaded, year_i, month_i, warn=st.warning
            )

            # 엔진이 읽는 전역은 반드시 core 모듈에 직접 설정한다.
            # 자동 조회 공휴일 + 엑셀에서 추가한 임시공휴일을 합친다.
            core.HOLIDAYS = {**st.session_state["holidays"], **extra_holidays}
            core.MANUAL_DAYS = set(manual_days)   # 규칙 배정 건너뛰고 빈칸으로 둘 날
            core.set_vacations(vacations)

            weeks, day_data = build_schedule(year_i, month_i)
            wb, problems = save_workbook(year_i, month_i, day_data, weeks)

            st.session_state["schedule_bytes"] = to_bytes(wb)
            st.session_state["problems"] = problems
            st.session_state["vac_count"] = len(vacations)

            msg = f"휴가/부재 {len(vacations)}건을 반영해 스케줄을 생성했습니다."
            if manual_days:
                days_txt = ", ".join(f"{d.month}/{d.day}" for d in sorted(manual_days))
                msg += f"  · 수동 배정일(빈칸): {days_txt}"
            if extra_holidays:
                hol_txt = ", ".join(f"{d.month}/{d.day}" for d in sorted(extra_holidays))
                msg += f"  · 임시공휴일: {hol_txt}"
            st.success(msg)

    if "schedule_bytes" in st.session_state:
        y_, m_ = st.session_state["year"], st.session_state["month"]
        st.download_button(
            f"완성된 스케줄 다운로드 ({y_}-{m_:02d}.xlsx)",
            data=st.session_state["schedule_bytes"],
            file_name=f"{y_}-{m_:02d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("⬇ 위 버튼을 누르면 브라우저의 다운로드 폴더에 저장됩니다.")
        problems = st.session_state.get("problems", [])
        if problems:
            st.warning(f"[검증] 문제 {len(problems)}건 발견")
            for p in problems:
                st.write("- " + p)
        else:
            st.success("[검증] 중복자/누락자 없음")


if __name__ == "__main__":
    render_app()
