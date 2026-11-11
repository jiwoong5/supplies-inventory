"""웹 화면(app.py) 자동 테스트.   실행: docker compose run --rm web python -m unittest -v"""
import tempfile
import unittest
from pathlib import Path

from app import create_app
from inventory import Inventory


class WebTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name) / "소모품_재고.xlsx"
        self.client = create_app(self.path).test_client()

    def submit(self, url, **form):
        """화면에서 버튼을 누른 것처럼 폼을 보내고, 이어서 표시되는 화면(HTML)을 돌려준다."""
        return self.client.post(url, data=form, follow_redirects=True).get_data(as_text=True)

    def stock(self, name):
        return dict(Inventory(self.path).list_items())[name]

    def test_첫_화면에_품목별_현재_재고가_보인다(self):
        page = self.client.get("/").get_data(as_text=True)
        for text in ("무선마우스", "30개", "키보드", "20개", "USB 메모리", "50개", "품목 3종 · 총 재고 100개"):
            self.assertIn(text, page)

    def test_입고_버튼을_누르면_재고가_늘고_완료_알림이_뜬다(self):
        page = self.submit("/change", item="무선마우스", quantity="10", kind="입고")
        self.assertIn("무선마우스 10개 입고 완료 (재고 30개 → 40개)", page)
        self.assertEqual(self.stock("무선마우스"), 40)

    def test_출고_버튼을_누르면_재고가_줄고_이력에_남는다(self):
        page = self.submit("/change", item="USB 메모리", quantity="5", kind="출고")
        self.assertIn("USB 메모리 5개 출고 완료 (재고 50개 → 45개)", page)
        self.assertIn("50 → 45", page)  # 이력 표
        self.assertEqual(self.stock("USB 메모리"), 45)

    def test_재고보다_많은_출고는_오류_알림만_뜨고_재고는_그대로다(self):
        page = self.submit("/change", item="키보드", quantity="25", kind="출고")
        self.assertIn("재고가 부족해 출고할 수 없습니다.", page)
        self.assertEqual(self.stock("키보드"), 20)

    def test_잘못된_요청은_재고를_바꾸지_않는다(self):
        for form in (
            dict(item="키보드", quantity="abc", kind="입고"),
            dict(item="키보드", quantity="-5", kind="입고"),
            dict(item="키보드", quantity="5", kind="폐기"),
            dict(item="모니터", quantity="5", kind="입고"),
            dict(),
        ):
            with self.subTest(form=form):
                self.assertIn('class="notice error"', self.submit("/change", **form))
        self.assertEqual(self.stock("키보드"), 20)

    def test_처리_후에는_목록_화면으로_이동해_새로고침해도_중복_처리되지_않는다(self):
        response = self.client.post("/change", data=dict(item="키보드", quantity="1", kind="입고", q="키"))
        self.assertEqual(response.status_code, 302)
        self.assertEqual(self.client.get(response.location).status_code, 200)  # 새로고침 = 조회(GET)만 다시 실행
        self.assertEqual(self.stock("키보드"), 21)

    def test_품목_이름으로_검색할_수_있다(self):
        page = self.client.get("/?q=usb").get_data(as_text=True)
        self.assertIn("50개", page)
        self.assertNotIn("30개", page)
        self.assertIn("검색 결과가 없습니다.", self.client.get("/?q=모니터").get_data(as_text=True))

    def test_품목을_추가하면_재고_0개로_목록에_나타난다(self):
        page = self.submit("/items", name="HDMI 케이블")
        self.assertIn("HDMI 케이블", page)
        self.assertEqual(self.stock("HDMI 케이블"), 0)
        self.assertIn("이미 등록된 품목", self.submit("/items", name="HDMI 케이블"))

    def test_품목_이름에_들어간_HTML_태그는_실행되지_않고_글자로만_보인다(self):
        page = self.submit("/items", name="<script>alert(1)</script>")
        self.assertNotIn("<script>alert(1)</script>", page)
        self.assertIn("&lt;script&gt;", page)


if __name__ == "__main__":
    unittest.main()
