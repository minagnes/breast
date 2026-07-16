# -*- coding: utf-8 -*-
"""
[배포용 진입 다리 파일] streamlit_app.py  (레포 루트)
================================================================
⚠ 이 파일은 수정할 필요가 없습니다. 진짜 앱 코드는 src/streamlit_app.py 에 있습니다.

Streamlit Community Cloud 의 "메인 파일 경로"가 streamlit_app.py(루트)로
설정돼 있어도 그대로 배포되도록, src/ 안의 진짜 앱을 대신 실행해 주는
얇은 다리(shim)일 뿐입니다.

로컬에서도 다음 중 아무거나로 실행됩니다.
    python -m streamlit run streamlit_app.py         # (이 다리 파일)
    python -m streamlit run src/streamlit_app.py     # (진짜 앱 직접)
"""
import os
import sys
import runpy

_SRC = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src")
if _SRC not in sys.path:
    sys.path.insert(0, _SRC)

# src/streamlit_app.py 를 메인 스크립트처럼 실행 (그 안의 render_app() 이 호출됨)
runpy.run_path(os.path.join(_SRC, "streamlit_app.py"), run_name="__main__")
