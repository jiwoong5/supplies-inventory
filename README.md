# 소모품 재고관리 프로그램

엑셀 파일(`data/소모품_재고.xlsx`)을 데이터 원본으로 쓰는 재고 조회·입출고 웹 프로그램입니다.
Docker만 있으면 실행됩니다.

- 시연 영상과 사용 가이드: https://jiwoong5.github.io/supplies-inventory/
- 같은 가이드의 PDF: `사용가이드.pdf`

## 실행

```bash
docker compose up -d        # 시작 → 브라우저에서 http://localhost:8000
docker compose down         # 종료
```

- 처음 실행하면 과제 예시 데이터(무선마우스 30, 키보드 20, USB 메모리 50)로 엑셀 파일이 만들어집니다.
- 엑셀 파일은 `data` 폴더에 있어 Excel로 그대로 열어 볼 수 있습니다. 입출고 기록은 `입출고이력` 시트에 쌓입니다.
- Excel에서 파일을 열어 둔 채로 입·출고하면 저장이 거부되고 안내가 뜹니다. 파일을 닫고 다시 누르면 됩니다.
- 예시 데이터로 되돌리려면 `data/소모품_재고.xlsx`를 지우고 `docker compose restart web`을 실행합니다.

## 테스트

```bash
docker compose run --rm web python -m unittest -v
```

## 제출용 PDF와 사용 가이드 만들기

프로그램은 Docker에서 띄우고, 화면 조작과 캡처는 PC에 설치된 Edge(또는 Chrome)를 Playwright로 구동합니다.

```bash
# 준비 (한 번만)
py -3.11 -m venv --system-site-packages report/.venv
report/.venv/Scripts/pip install playwright

# 생성
report/.venv/Scripts/python report/build_pdf.py
```

하는 일: 컨테이너에서 자동 테스트 실행 → 예시 데이터로 프로그램을 따로 띄워 사용자 흐름대로 조작·확인·캡처 → PDF 저장

| 결과물 | 내용 |
|---|---|
| `AI활용능력테스트_제출.pdf` | ① 보고서(1페이지) + ② 소스코드 + ③ 프롬프트 + 부록(사용 가이드) |
| `사용가이드.pdf` | 사용 가이드만 |
| `docs/index.html`, `docs/demo.webm` | GitHub Pages로 배포되는 웹 페이지(시연 영상 + 사용 가이드) |

PDF 머리말의 링크는 `git remote`의 GitHub 주소에서 자동으로 만들어집니다.

- **프롬프트 캡처**: 캡처 이미지(png, jpg)를 `report/prompts` 폴더에 넣고 다시 생성하면 ③에 들어갑니다.
  파일 이름 순서대로 배치됩니다. 이미지가 없으면 `report/prompts.txt`의 글자로 대신합니다.
- **지원자 이름 표시**: 환경변수 `APPLICANT`에 이름을 넣고 실행합니다.
- **부록 빼기**: 환경변수 `WITH_GUIDE=0`으로 실행하면 제출용 PDF에서 사용 가이드를 뺍니다.
- **문구 수정**: `report/templates`의 `submission.html`(보고서), `_guide.html`(사용 가이드)을 고칩니다.
  보고서가 1페이지를 넘으면 PDF를 만들지 않고 알려 줍니다.

## 파일 구성

| 파일 | 역할 |
|---|---|
| `inventory.py` | 재고 로직과 엑셀 저장 (조회, 입고, 출고, 품목 추가, 이력 기록) |
| `app.py` | 웹 화면의 요청 처리 (Flask) |
| `templates/index.html` | 화면 |
| `test_inventory.py`, `test_app.py` | 자동 테스트 |
| `Dockerfile`, `docker-compose.yml` | 실행 환경 |
| `report/` | 문서 생성 스크립트, 보고서·사용 가이드 템플릿, 캡처 이미지 |
