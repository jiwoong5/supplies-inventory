"""제출용 PDF와 사용 가이드 생성기.

프로그램은 Docker 컨테이너에서 띄우고, 화면 조작과 캡처는 PC에 설치된 Edge(또는 Chrome)를 Playwright로 구동한다.

  준비(한 번)   py -3.11 -m venv --system-site-packages report/.venv
                report/.venv/Scripts/pip install playwright
  실행          report/.venv/Scripts/python report/build_pdf.py

하는 일
  1) 컨테이너 안에서 자동 테스트를 돌려 전부 통과하는지 확인한다.
  2) 예시 데이터만 든 엑셀로 프로그램을 따로 띄우고, 사용자 흐름대로 조작하면서 결과를 확인하고 화면을 캡처한다.
     (접속 → 조회 → 검색 → 입고 → 출고 → 재고 부족 → 품목 추가 → 이력)
  3) PDF 두 개를 프로젝트 루트에 저장한다.
       AI활용능력테스트_제출.pdf : ① 보고서 + ② 소스코드 + ③ 프롬프트 + 부록(사용 가이드)
       사용가이드.pdf            : 사용 가이드만

프롬프트 캡처 이미지는 report/prompts 폴더에 넣는다.
환경변수:  APPLICANT=지원자 이름,  WITH_GUIDE=0 (제출용 PDF에서 부록을 뺀다)
"""
import base64
import json
import os
import re
import subprocess
import sys
import time
import urllib.request
from contextlib import contextmanager
from datetime import date
from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from markupsafe import Markup
from playwright.sync_api import expect, sync_playwright
from pygments import highlight
from pygments.formatters import HtmlFormatter
from pygments.lexers import get_lexer_by_name

ROOT = Path(__file__).resolve().parent.parent
REPORT = ROOT / "report"
SHOTS = REPORT / "screenshots"
OUT = REPORT / "out"
DOCS = ROOT / "docs"  # GitHub Pages로 배포되는 폴더 (시연 영상 + 웹용 사용 가이드)
SUBMISSION_PDF = ROOT / "AI활용능력테스트_제출.pdf"
GUIDE_PDF = ROOT / "사용가이드.pdf"

DEMO = "swhigh-demo"  # 캡처용으로 잠깐 띄우는 컨테이너 이름
DEMO_PORT = 8001
DEMO_FILE = "/tmp/demo/소모품_재고.xlsx"  # 컨테이너 안의 임시 엑셀 → 실제 data 폴더는 건드리지 않는다
SOURCES = [  # (파일, 역할, 문법 강조용 언어 이름)
    ("inventory.py", "재고 로직 · 엑셀 저장", "python"),
    ("app.py", "웹 화면 (요청 처리)", "python"),
    ("templates/index.html", "화면 템플릿", "html+jinja"),
    ("test_inventory.py", "재고 로직 자동 테스트", "python"),
    ("test_app.py", "웹 화면 자동 테스트", "python"),
    ("Dockerfile", "실행 환경 정의", "docker"),
    ("docker-compose.yml", "실행 구성", "yaml"),
    ("requirements.txt", "사용 라이브러리", "text"),
]
IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".webp": "image/webp"}
WIDTH = 660  # 캡처할 화면 너비(px)
FOOTER = (
    "<div style=\"width:100%; text-align:center; font:7pt 'Malgun Gothic',sans-serif; color:#98a2b3;\">"
    "<span class='pageNumber'></span> / <span class='totalPages'></span></div>"
)
MARK_SCRIPT = """(boxes) => {
  const layer = document.createElement('div');
  layer.id = 'guide-marks';
  boxes.forEach((box, index) => {
    const frame = document.createElement('div');
    frame.style.cssText = `position:absolute; left:${box.x - 5}px; top:${box.y - 5}px;
      width:${box.width + 10}px; height:${box.height + 10}px; border:2.5px solid #e11d48; border-radius:9px;
      box-shadow:0 0 0 3px rgba(225,29,72,.16); pointer-events:none; z-index:9999;`;
    const badge = document.createElement('span');
    badge.textContent = index + 1;
    badge.style.cssText = `position:absolute; left:-12px; top:-12px; width:22px; height:22px; border-radius:50%;
      background:#e11d48; color:#fff; font:700 13px/22px sans-serif; text-align:center; box-shadow:0 0 0 2px #fff;`;
    frame.append(badge);
    layer.append(frame);
  });
  document.body.append(layer);
}"""


def main():
    for stream in (sys.stdout, sys.stderr):
        if not stream.isatty():
            stream.reconfigure(encoding="utf-8")
    with_guide = os.environ.get("WITH_GUIDE", "1") != "0"
    docker("compose", "build", "web")  # 지금 소스 그대로 이미지에 반영
    test_count = run_tests()
    DOCS.mkdir(exist_ok=True)
    with sync_playwright() as playwright:
        with demo_app() as url:
            record_demo(playwright, url, DOCS / "demo.webm")
        with demo_app() as url:
            browser = launch_browser(playwright)
            try:
                shots = walk_through(browser, url)
                content = document_content(test_count, shots, read_sheets(), with_guide)
                submission_pdf, guide_pdf = build_documents(browser, content)
            finally:
                browser.close()
    for path, pdf in ((SUBMISSION_PDF, submission_pdf), (GUIDE_PDF, guide_pdf)):
        try:
            path.write_bytes(pdf)
        except OSError as error:
            sys.exit(f"{path.name} 을 저장하지 못했습니다. 파일이 열려 있다면 닫고 다시 실행하세요. ({error})")
        print(f"완료: {path.name}  ({count_pages(pdf)}페이지, {len(pdf) / 1024:.0f}KB)")
    print(f"완료: docs/index.html, docs/demo.webm  (GitHub Pages용 · 링크: {content['links'].get('site', '원격 저장소 없음')})")
    print(f"확인: 자동 테스트 {test_count}건 통과 · 화면 조작 확인 통과 · 보고서 1페이지"
          f" · 부록(사용 가이드) {'포함' if with_guide else '제외'}")


def build_documents(browser, content: dict) -> tuple[bytes, bytes]:
    """웹 페이지(docs/index.html)를 쓰고, 제출용 PDF와 사용 가이드 PDF를 만들어 돌려준다."""
    render("site.html", content, DOCS / "index.html")
    guide_pdf = print_pdf(browser, render("guide.html", content))
    submission = render("submission.html", content)
    report_pdf = print_pdf(browser, submission, report_only=True)
    (OUT / "report_only.pdf").write_bytes(report_pdf)  # 보고서만 따로 인쇄해 1페이지인지 확인
    if count_pages(report_pdf) != 1:
        sys.exit(f"보고서가 {count_pages(report_pdf)}페이지입니다. report/out/report_only.pdf 에서 넘친 부분을 확인하고"
                 " report/templates/submission.html 을 줄이세요.")
    return print_pdf(browser, submission), guide_pdf


def links() -> dict:
    """git 원격 주소에서 저장소와 GitHub Pages 주소를 알아낸다. 원격 저장소가 없으면 링크를 넣지 않는다."""
    remote = subprocess.run(["git", "remote", "get-url", "origin"], cwd=ROOT, capture_output=True, text=True).stdout
    found = re.search(r"github\.com[:/]([^/]+)/(.+?)(?:\.git)?\s*$", remote)
    if not found:
        return {}
    user, repository = found.groups()
    return {"repo": f"https://github.com/{user}/{repository}", "site": f"https://{user.lower()}.github.io/{repository}/"}



# ── Docker: 테스트 실행 · 캡처용 프로그램 띄우기 ──────────────────────
def docker(*arguments, check=True) -> str:
    result = subprocess.run(["docker", *arguments], cwd=ROOT, capture_output=True, encoding="utf-8", errors="replace")
    if check and result.returncode != 0:
        sys.exit(f"docker {' '.join(arguments)} 실패\n{result.stdout}{result.stderr}")
    return result.stdout + result.stderr


def run_tests() -> int:
    """컨테이너 안에서 자동 테스트를 실행하고 통과한 건수를 돌려준다. 하나라도 실패하면 문서를 만들지 않는다."""
    output = docker("compose", "run", "--rm", "web", "python", "-m", "unittest")
    ran = re.search(r"Ran (\d+) tests", output)
    if not ran or not re.search(r"^OK\s*$", output, re.MULTILINE):
        sys.exit(f"{output}\n자동 테스트가 실패해 문서를 만들지 않았습니다.")
    return int(ran.group(1))


@contextmanager
def demo_app():
    """예시 데이터만 든 새 엑셀 파일로 프로그램을 따로 띄우고, 끝나면 지운다."""
    docker("rm", "-f", DEMO, check=False)
    docker("compose", "run", "-d", "--name", DEMO, "-p", f"{DEMO_PORT}:8000", "-e", f"INVENTORY_FILE={DEMO_FILE}", "web")
    url = f"http://localhost:{DEMO_PORT}/"
    try:
        for _ in range(60):
            try:
                urllib.request.urlopen(url, timeout=2).close()
                break
            except OSError:
                time.sleep(0.5)
        else:
            sys.exit(f"프로그램이 켜지지 않았습니다.\n{docker('logs', DEMO, check=False)}")
        yield url
    finally:
        docker("rm", "-f", DEMO, check=False)


def read_sheets() -> dict:
    """캡처를 마친 뒤 엑셀 파일에 실제로 저장된 내용을 {시트 이름: 줄 목록}으로 읽어 온다."""
    code = ("import json, os; from openpyxl import load_workbook; "
            "book = load_workbook(os.environ['INVENTORY_FILE']); "
            "print(json.dumps({s.title: list(s.iter_rows(values_only=True)) for s in book.worksheets}))")
    return json.loads(docker("exec", DEMO, "python", "-c", code))


# ── Playwright로 사용자 흐름대로 조작하며 확인 · 캡처 ─────────────────
def launch_browser(playwright, **options):
    for channel in ("msedge", "chrome"):  # PC에 설치된 브라우저를 그대로 쓴다
        try:
            return playwright.chromium.launch(channel=channel, **options)
        except Exception:
            continue
    sys.exit("Microsoft Edge 또는 Chrome 을 찾지 못했습니다.")


CAPTION_SCRIPT = """addEventListener('DOMContentLoaded', () => {
  const bar = document.createElement('div');
  bar.id = 'demo-caption';
  bar.style.cssText = `position:fixed; left:0; right:0; bottom:0; padding:13px 20px; background:rgba(17,24,39,.93);
    color:#fff; font:600 17px "Malgun Gothic",sans-serif; z-index:9999;`;
  bar.textContent = sessionStorage.getItem('caption') || '';
  document.body.append(bar);
});"""


def record_demo(playwright, url, target: Path) -> None:
    """시연 영상: 사람이 쓰는 속도로 조회 → 입고 → 출고 → 재고 부족 → 품목 추가 → 이력을 조작하며 녹화한다."""
    size = {"width": 760, "height": 940}
    browser = launch_browser(playwright, slow_mo=250)
    context = browser.new_context(viewport=size, locale="ko-KR", record_video_dir=str(OUT / "video"), record_video_size=size)
    context.add_init_script(CAPTION_SCRIPT)  # 화면이 바뀌어도 아래쪽 자막이 유지되게 한다
    page = context.new_page()
    page.set_default_timeout(10_000)
    notice = page.get_by_role("alert")

    def caption(text):
        page.evaluate("""text => { sessionStorage.setItem('caption', text);
                                   document.getElementById('demo-caption').textContent = text; }""", text)
        page.wait_for_timeout(1500)

    def row(name):
        return page.locator("#stock tbody tr").filter(has=page.get_by_role("cell", name=name, exact=True))

    def press(button):  # 누르는 버튼이 눈에 띄도록 테두리를 잠깐 보여 준 뒤 누른다
        button.evaluate("element => element.style.outline = '4px solid #facc15'")
        page.wait_for_timeout(600)
        button.click()

    def process(name, quantity, kind):
        row(name).get_by_placeholder("수량").press_sequentially(str(quantity), delay=180)
        press(row(name).get_by_role("button", name=kind))

    page.goto(url)
    caption("1. 재고 조회 · 품목별 현재 재고를 확인합니다 (예시 데이터 30 / 20 / 50)")
    page.wait_for_timeout(1500)
    caption("2. 입고 · 무선마우스 10개를 입력하고 [입고]를 누릅니다")
    process("무선마우스", 10, "입고")
    expect(notice).to_contain_text("입고 완료")
    caption("재고가 30개 → 40개로 바뀌고 엑셀 파일에 저장됩니다")
    caption("3. 출고 · USB 메모리 5개를 입력하고 [출고]를 누릅니다")
    process("USB 메모리", 5, "출고")
    expect(notice).to_contain_text("출고 완료")
    caption("재고가 50개 → 45개로 줄어듭니다")
    caption("4. 재고 부족 · 재고 20개인 키보드를 25개 출고해 봅니다")
    process("키보드", 25, "출고")
    expect(notice).to_contain_text("재고가 부족해 출고할 수 없습니다.")
    caption("처리되지 않고 이유를 알려 줍니다. 재고는 20개 그대로입니다")
    caption("5. 품목 추가 · 'HDMI 케이블'을 등록합니다")
    page.locator("#add").get_by_placeholder("새 품목 이름").press_sequentially("HDMI 케이블", delay=120)
    press(page.locator("#add").get_by_role("button", name="추가"))
    expect(notice).to_contain_text("품목을 등록했습니다.")
    caption("재고 0개로 등록됩니다. 입고로 15개를 넣습니다")
    process("HDMI 케이블", 15, "입고")
    expect(row("HDMI 케이블").locator("td.stock")).to_have_text("15개")
    caption("6. 이력 · 모든 입출고가 아래 표와 엑셀 '입출고이력' 시트에 자동으로 남습니다")
    page.wait_for_timeout(2500)
    context.close()  # 닫아야 영상 파일이 완성된다
    Path(page.video.path()).replace(target)
    browser.close()


def walk_through(browser, url) -> dict:
    """사용자가 쓰는 순서대로 화면을 조작한다. 기대와 다른 결과가 나오면 예외가 나서 문서를 만들지 않는다."""
    SHOTS.mkdir(exist_ok=True)
    page = browser.new_page(viewport={"width": WIDTH, "height": 1400}, device_scale_factor=2, locale="ko-KR")
    page.set_default_timeout(10_000)
    stock_card, add_card, history_card = page.locator("#stock"), page.locator("#add"), page.locator("#history")
    notice = page.get_by_role("alert")
    shots = {}

    def row(name):  # 재고 표에서 해당 품목의 줄
        return stock_card.locator("tbody tr").filter(has=page.get_by_role("cell", name=name, exact=True))

    def stock(name):
        return row(name).locator("td.stock")

    def column(card, number):  # 표의 한 열 전체
        return union(card.locator(f"th:nth-child({number})"), card.locator(f"tbody tr:last-child td:nth-child({number})"))

    def process(name, quantity, kind):  # 수량을 입력하고 [입고] 또는 [출고] 버튼을 누른다
        row(name).get_by_placeholder("수량").fill(str(quantity))
        row(name).get_by_role("button", name=kind).click()

    # 1. 접속 · 2. 재고 조회
    page.goto(url)
    for name, expected in (("무선마우스", "30개"), ("키보드", "20개"), ("USB 메모리", "50개")):
        expect(stock(name)).to_have_text(expected)
    expect(page.locator("header .sub")).to_have_text("품목 3종 · 총 재고 100개")
    shots["report_view"] = capture(page, "report_1_view.png", bottom=stock_card)
    shots["overview"] = capture(page, "guide_1_overview.png", marks=[
        stock_card, row("무선마우스").locator("form"), add_card, history_card])
    shots["view"] = capture(page, "guide_2_view.png", bottom=stock_card, marks=[
        column(stock_card, 2), page.locator("header .sub")])

    # 3. 품목 검색
    stock_card.get_by_role("searchbox").fill("usb")
    stock_card.get_by_role("button", name="검색").click()
    expect(stock_card.locator("tbody tr")).to_have_count(1)
    expect(stock("USB 메모리")).to_have_text("50개")
    shots["search"] = capture(page, "guide_3_search.png", top=stock_card, bottom=stock_card, marks=[
        stock_card.locator("form[method=get]")])
    stock_card.get_by_role("searchbox").fill("")
    stock_card.get_by_role("button", name="검색").click()
    expect(stock_card.locator("tbody tr")).to_have_count(3)

    # 잘못된 입력: Enter 키만 누르거나, 1 이상의 정수가 아닌 수량이면 처리되지 않아야 한다
    field = row("키보드").get_by_placeholder("수량")
    field.fill("3")
    field.press("Enter")
    for wrong in ("", "0", "-3", "1.5"):
        field.fill(wrong)
        assert not field.evaluate("input => input.checkValidity()"), f"수량 {wrong!r} 이(가) 허용됨"
        row("키보드").get_by_role("button", name="입고").click()
    page.wait_for_timeout(500)
    expect(notice).to_have_count(0)
    expect(stock("키보드")).to_have_text("20개")
    page.reload()  # 브라우저가 띄운 입력 오류 말풍선이 다음 캡처에 남지 않게 한다

    # 4. 입고
    row("무선마우스").get_by_placeholder("수량").fill("10")
    shots["in_before"] = capture(page, "guide_4_in_before.png", top=stock_card, bottom=stock_card, marks=[
        row("무선마우스").get_by_placeholder("수량"), row("무선마우스").get_by_role("button", name="입고")])
    row("무선마우스").get_by_role("button", name="입고").click()
    expect(notice).to_have_text("무선마우스 10개 입고 완료 (재고 30개 → 40개)")
    expect(stock("무선마우스")).to_have_text("40개")
    shots["in_after"] = capture(page, "guide_4_in_after.png", bottom=stock_card, marks=[notice, stock("무선마우스")])

    # 5. 출고
    process("USB 메모리", 5, "출고")
    expect(notice).to_have_text("USB 메모리 5개 출고 완료 (재고 50개 → 45개)")
    expect(stock("USB 메모리")).to_have_text("45개")
    expect(history_card.locator("tbody tr")).to_have_count(2)
    shots["report_change"] = capture(page, "report_2_change.png")
    shots["out_after"] = capture(page, "guide_5_out_after.png", bottom=stock_card, marks=[notice, stock("USB 메모리")])

    # 6. 재고보다 많은 출고
    process("키보드", 25, "출고")
    expect(notice).to_contain_text("재고가 부족해 출고할 수 없습니다.")
    expect(stock("키보드")).to_have_text("20개")
    expect(history_card.locator("tbody tr")).to_have_count(2)
    shots["report_error"] = capture(page, "report_3_error.png", top=notice, bottom=stock_card)
    shots["error"] = capture(page, "guide_6_error.png", bottom=stock_card, marks=[notice, stock("키보드")])

    # 7. 품목 추가 → 입고로 수량 넣기
    add_card.get_by_placeholder("새 품목 이름").fill("HDMI 케이블")
    shots["add_before"] = capture(page, "guide_7_add_before.png", top=add_card, bottom=add_card, marks=[
        add_card.get_by_placeholder("새 품목 이름"), add_card.get_by_role("button", name="추가")])
    add_card.get_by_role("button", name="추가").click()
    expect(notice).to_contain_text("'HDMI 케이블' 품목을 등록했습니다.")
    expect(stock("HDMI 케이블")).to_have_text("0개")
    expect(stock("HDMI 케이블")).to_have_class(re.compile(r"\bzero\b"))  # 재고 0개는 빨간색
    shots["add_after"] = capture(page, "guide_7_add_after.png", bottom=stock_card, marks=[notice, row("HDMI 케이블")])
    process("HDMI 케이블", 15, "입고")
    expect(stock("HDMI 케이블")).to_have_text("15개")

    # 8. 이력
    expect(history_card.locator("tbody tr")).to_have_count(4)
    expect(page.locator("header .sub")).to_have_text("품목 4종 · 총 재고 120개")
    shots["history"] = capture(page, "guide_8_history.png", top=history_card, bottom=history_card, marks=[
        column(history_card, 2), column(history_card, 5)])
    page.close()
    return shots


def capture(page, name, *, top=None, bottom=None, marks=()) -> Path:
    """top 위쪽부터 bottom 아래쪽까지(생략하면 화면 전체)를 PNG로 저장한다. marks 의 요소에는 번호 테두리를 덧그린다."""
    if marks:
        page.evaluate(MARK_SCRIPT, [box_of(mark) for mark in marks])
    start = max(0, box_of(top)["y"] - 16) if top else 0
    last = box_of(bottom or page.locator("main"))
    end = last["y"] + last["height"] + (16 if bottom else 0)
    path = SHOTS / name
    page.screenshot(path=path, clip={"x": 0, "y": start, "width": WIDTH, "height": end - start})
    page.evaluate("document.getElementById('guide-marks')?.remove()")
    return path


def box_of(target) -> dict:
    return target if isinstance(target, dict) else target.bounding_box()


def union(*locators) -> dict:
    """여러 요소를 모두 감싸는 사각형."""
    boxes = [locator.bounding_box() for locator in locators]
    left, top = min(box["x"] for box in boxes), min(box["y"] for box in boxes)
    right = max(box["x"] + box["width"] for box in boxes)
    bottom = max(box["y"] + box["height"] for box in boxes)
    return {"x": left, "y": top, "width": right - left, "height": bottom - top}


# ── 문서 조립 · PDF 인쇄 ──────────────────────────────────────────────
def document_content(test_count: int, shots: dict, sheets: dict, with_guide: bool) -> dict:
    captures = sorted(path for path in (REPORT / "prompts").glob("*") if path.suffix.lower() in IMAGE_TYPES)
    if not captures:
        print("알림: report/prompts 폴더에 캡처 이미지가 없어 ③ 프롬프트를 글자로 대신 넣었습니다.")
    prompts = (REPORT / "prompts.txt").read_text(encoding="utf-8").splitlines()
    return dict(
        today=date.today().strftime("%Y. %m. %d."),
        applicant=os.environ.get("APPLICANT", "").strip(),
        test_count=test_count,
        with_guide=with_guide,
        links=links(),
        shots={key: data_uri(path) for key, path in shots.items()},
        sheets=sheets,
        sources=[source(*entry) for entry in SOURCES],
        prompt_images=[data_uri(path) for path in captures],
        prompt_texts=[  # "프롬프트 || 설명" → (프롬프트, 설명)
            [part.strip() for part in line.partition("||")[::2]]
            for line in prompts if line.strip() and not line.startswith("#")
        ],
    )


def render(template: str, content: dict, target: Path | None = None) -> Path:
    environment = Environment(loader=FileSystemLoader(REPORT / "templates"), autoescape=select_autoescape(["html"]))
    OUT.mkdir(exist_ok=True)
    page = target or OUT / template
    page.write_text(environment.get_template(template).render(**content), encoding="utf-8")
    return page


def source(file, role, language) -> dict:
    code = (ROOT / file).read_text(encoding="utf-8")
    html = highlight(code, get_lexer_by_name(language), HtmlFormatter(nowrap=True, noclasses=True, style="xcode"))
    return {"file": file, "role": role, "lines": [Markup(line) or " " for line in html.rstrip("\n").split("\n")]}


def data_uri(path: Path) -> str:
    return f"data:{IMAGE_TYPES[path.suffix.lower()]};base64,{base64.b64encode(path.read_bytes()).decode()}"


def print_pdf(browser, document: Path, report_only=False) -> bytes:
    page = browser.new_page()
    page.goto(document.as_uri())
    if report_only:
        page.evaluate("document.body.classList.add('report-only')")
    pdf = page.pdf(print_background=True, prefer_css_page_size=True,
                   display_header_footer=True, header_template="<span></span>", footer_template=FOOTER)
    page.close()
    return pdf


def count_pages(pdf: bytes) -> int:
    return len(re.findall(rb"/Type\s*/Page\b", pdf))


if __name__ == "__main__":
    main()
