# CLAUDE.md — 프로젝트 지도 (Claude Code용)

유방영상의학과 월간 근무 스케줄을 자동 생성하는 프로그램. 터미널판과 웹판 두 가지 실행 방법이 있고, 둘 다 같은 엔진을 쓴다.

## 파일 지도

| 경로 | 역할 |
|------|------|
| `src/schedule_core.py` | ★ **배정 규칙·스케줄 엔진·엑셀 시트 작성** 단일 소스. 규칙은 여기 한 곳에서만 고친다. |
| `src/generate_schedule.py` | 터미널 진입점(얇은 껍데기). 사용자 입력·파일 저장만 담당. |
| `src/streamlit_app.py` | 웹 진입점(얇은 껍데기). 화면·업로드·다운로드만 담당. |
| `docs/스케줄_배정_규정.docx` | **SoT(원천 규칙 문서)**. ⚠ 개인정보 포함 — GitHub 미배포(로컬 전용). |
| `outputs/` | 생성된 결과물(완성 스케줄·휴가 입력표) 저장 위치. ⚠ 개인정보 — GitHub 미배포. |
| `requirements.txt` | 의존 패키지: streamlit / openpyxl / holidays |

## ★ 가장 중요한 규칙

1. **배정 규칙은 `src/schedule_core.py` 한 곳에서만 고친다.** 터미널판·웹판은 이 파일을 import 하므로 여기만 고치면 둘 다 반영된다. (예전엔 두 파일에 규칙이 복사돼 있었다 — 다시 그렇게 만들지 말 것.)
2. **규칙을 바꾸면 `docs/스케줄_배정_규정.docx`(SoT)도 함께 갱신한다.** SoT 문서 자체가 "규칙 변경 시 코드와 문서를 함께 갱신한다"고 명시한다. 코드와 문서가 어긋나면 안 된다.
3. **개인정보 보호**: `docs/`(SoT)와 `outputs/`(실제 스케줄)는 근무자 정보라 **절대 GitHub에 올리지 않는다**(`.gitignore` 처리됨). 커밋·푸시 전 이 파일들이 스테이징되지 않았는지 확인할 것.

> 참고: SoT 문서가 언급하는 `schedule_generator.py`(과거 파일, 현재 삭제됨)의 배정 로직이 지금은 `src/schedule_core.py`에 들어 있다.

## 자주 바꾸는 곳 (src/schedule_core.py 상단)

- **인력 구성**: `K_STAFF`, `F_STAFF`, `J_STAFF`, `R_NAME` (사람이 바뀔 때)
- **요일별 고정 배정**: `BREAST_AM_LEAD`, `MAMMO2_PM_FIX`, `BREAST_PM_FIX` 등
- **전공의 로테이션 기준일**: `ROTATION_ANCHOR_MONDAY`, `ROTATION_ANCHOR_LABEL`
- **휴가 드롭다운 선택지**: `VAC_CHOICES`

## 실행

```bash
# 준비(최초 1회): 가상환경 + 패키지
python -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

# 터미널판
python src/generate_schedule.py            # 결과물 → outputs/

# 웹판
python -m streamlit run src/streamlit_app.py   # 결과물 → 브라우저 다운로드 폴더
```

## 결과물 저장 위치

- 터미널판: 프로젝트 루트의 **`outputs/`** 폴더 (`YYYY-MM.xlsx`, 겹치면 `-1`, `-2` … 자동). 실행 중 전체 경로를 화면에 출력한다.
- 웹판: 브라우저의 **다운로드 폴더**.

## 안전한 수정 방법 (엔진을 건드릴 때)

규칙/엔진을 고친 뒤에는 **결과가 의도대로 바뀌었는지, 의도치 않게 깨지지 않았는지** 확인한다. 규칙 변경이 아닌 리팩터링(동작 보존)이라면, 변경 전 스케줄 결과를 저장해두고 변경 후와 비교해 동일한지 본다. `.venv/bin/python`에 openpyxl·holidays가 설치돼 있다.

## 커밋 메시지 규칙

`추가:` (새 기능) / `수정:` (기존 변경) / `버그:` (오작동 수정) / `문서:` (문서만).
예) `수정: K4 오후 Mammo2 고정 요일 변경` · `버그: 금요일 Thyroid US 최소 인원 미적용 수정`
