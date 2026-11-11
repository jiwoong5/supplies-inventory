"""재고 로직(inventory.py) 자동 테스트.   실행: docker compose run --rm web python -m unittest -v"""
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

from openpyxl import Workbook, load_workbook

from inventory import HISTORY_SHEET, STOCK_SHEET, Inventory, InventoryError


class InventoryTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.path = Path(folder.name) / "소모품_재고.xlsx"
        self.inventory = Inventory(self.path)

    def stock(self, name):
        """엑셀 파일을 새로 읽어 재고를 확인한다 (메모리가 아니라 파일에 저장됐는지 검증)."""
        return dict(Inventory(self.path).list_items())[name]

    # ── 재고 조회 ──
    def test_처음_실행하면_예시_데이터로_엑셀_파일을_만든다(self):
        self.assertEqual(self.inventory.list_items(), [("무선마우스", 30), ("키보드", 20), ("USB 메모리", 50)])
        self.assertEqual(self.inventory.history(), [])

    def test_엑셀에서_직접_고친_수량이_그대로_조회된다(self):
        book = load_workbook(self.path)
        book[STOCK_SHEET]["B2"] = 7
        book.save(self.path)
        self.assertEqual(self.inventory.list_items()[0], ("무선마우스", 7))

    def test_이력_시트가_없는_기존_엑셀_파일도_그대로_쓸_수_있다(self):
        path = self.path.with_name("기존장부.xlsx")
        book = Workbook()
        book.active.append(["품명", "수량"])  # 시트 이름도 '재고'가 아닌 기본값(Sheet)
        book.active.append(["A4용지", 12])
        book.save(path)
        inventory = Inventory(path)
        self.assertEqual(inventory.issue("A4용지", 2), (12, 10))
        self.assertEqual(inventory.list_items(), [("A4용지", 10)])
        self.assertEqual(len(inventory.history()), 1)

    def test_엑셀에_정수가_아닌_재고가_적혀_있으면_알려_준다(self):
        book = load_workbook(self.path)
        book[STOCK_SHEET]["B3"] = "많음"
        book.save(self.path)
        with self.assertRaisesRegex(InventoryError, "키보드"):
            self.inventory.list_items()

    # ── 재고 변경 ──
    def test_입고하면_재고가_늘고_엑셀에_저장된다(self):
        self.assertEqual(self.inventory.receive("무선마우스", 10), (30, 40))
        self.assertEqual(self.stock("무선마우스"), 40)

    def test_출고하면_재고가_줄어든다(self):
        self.assertEqual(self.inventory.issue("USB 메모리", "5"), (50, 45))  # 화면 입력값은 문자열로 들어온다
        self.assertEqual(self.stock("USB 메모리"), 45)

    def test_재고_전부를_출고하면_0개가_된다(self):
        self.inventory.issue("키보드", 20)
        self.assertEqual(self.stock("키보드"), 0)

    def test_재고보다_많은_출고는_거부하고_재고를_바꾸지_않는다(self):
        with self.assertRaisesRegex(InventoryError, "재고가 부족"):
            self.inventory.issue("키보드", 21)
        self.assertEqual(self.stock("키보드"), 20)

    def test_1_이상의_정수가_아닌_수량은_거부한다(self):
        for wrong in (0, -3, "abc", "1.5", "", None):
            with self.subTest(수량=wrong), self.assertRaises(InventoryError):
                self.inventory.receive("키보드", wrong)
        self.assertEqual(self.stock("키보드"), 20)

    def test_없는_품목은_거부한다(self):
        with self.assertRaisesRegex(InventoryError, "찾을 수 없습니다"):
            self.inventory.receive("모니터", 1)

    def test_입출고는_이력_시트에_최신순으로_남고_실패한_요청은_남지_않는다(self):
        self.inventory.receive("무선마우스", 10)
        self.inventory.issue("USB 메모리", 5)
        with self.assertRaises(InventoryError):
            self.inventory.issue("키보드", 999)
        records = [row[1:] for row in self.inventory.history()]  # 일시(row[0])를 뺀 나머지
        self.assertEqual(records, [("출고", "USB 메모리", 5, 50, 45), ("입고", "무선마우스", 10, 30, 40)])
        self.assertEqual(load_workbook(self.path)[HISTORY_SHEET].max_row, 3)  # 제목 줄 + 2건

    def test_동시에_여러_요청이_와도_수량이_어긋나지_않는다(self):
        threads = [threading.Thread(target=self.inventory.receive, args=("키보드", 1)) for _ in range(10)]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        self.assertEqual(self.stock("키보드"), 30)

    def test_저장에_실패하면_안내하고_원본_엑셀은_그대로_둔다(self):
        with mock.patch("inventory.os.replace", side_effect=PermissionError("파일 사용 중")):  # 엑셀에서 열어 둔 상황
            with self.assertRaisesRegex(InventoryError, "닫고 다시 시도"):
                self.inventory.receive("키보드", 5)
        self.assertEqual(self.stock("키보드"), 20)
        self.assertEqual([file.name for file in self.path.parent.iterdir()], [self.path.name])  # 임시 파일 없음

    # ── 품목 추가 ──
    def test_새_품목은_재고_0개로_등록된다(self):
        self.assertEqual(self.inventory.add_item("  HDMI 케이블 "), "HDMI 케이블")
        self.assertEqual(self.stock("HDMI 케이블"), 0)

    def test_이름이_비었거나_이미_있는_품목은_등록할_수_없다(self):
        for wrong in ("", "   ", "키보드"):
            with self.subTest(이름=wrong), self.assertRaises(InventoryError):
                self.inventory.add_item(wrong)
        self.assertEqual(len(self.inventory.list_items()), 3)


if __name__ == "__main__":
    unittest.main()
