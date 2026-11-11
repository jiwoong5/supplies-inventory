"""소모품 재고 관리 - 핵심 로직 (화면과 분리되어 있어 이 파일만으로 테스트할 수 있다).

회사에서 쓰던 엑셀 파일을 그대로 데이터 원본으로 사용한다.
  - '재고' 시트       : A열 품목, B열 현재재고 (이 이름의 시트가 없으면 첫 번째 시트를 재고표로 본다)
  - '입출고이력' 시트 : 입고·출고가 일어날 때마다 한 줄씩 자동 기록 (없으면 새로 만든다)
모든 작업은 '엑셀 읽기 → 변경 → 즉시 저장' 순서로 처리해, 엑셀에서 직접 고친 내용과 어긋나지 않는다.
"""
from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

STOCK_SHEET = "재고"
HISTORY_SHEET = "입출고이력"
STOCK_HEADER = ("품목", "현재재고")
HISTORY_HEADER = ("일시", "구분", "품목", "수량", "변경 전", "변경 후")
SAMPLE_ITEMS = (("무선마우스", 30), ("키보드", 20), ("USB 메모리", 50))  # 과제 예시 데이터


class InventoryError(Exception):
    """사용자에게 메시지를 그대로 보여 줄 수 있는 업무 오류."""


def parse_quantity(text) -> int:
    """입력값을 '1 이상의 정수' 수량으로 바꾼다. 그렇게 볼 수 없으면 InventoryError."""
    try:
        quantity = int(str(text).strip())
    except ValueError:
        raise InventoryError("수량은 숫자(정수)로 입력하세요.") from None
    if quantity < 1:
        raise InventoryError("수량은 1 이상이어야 합니다.")
    return quantity


class Inventory:
    """엑셀 파일 하나를 재고 장부로 다루는 클래스."""

    def __init__(self, path):
        self.path = Path(path)
        self._lock = threading.Lock()  # 여러 사람이 동시에 요청해도 '읽기→변경→저장'이 섞이지 않게 한다
        if not self.path.exists():  # 처음 실행이면 예시 데이터로 엑셀 파일을 만든다
            self._create_file()

    # ── 재고 조회 ─────────────────────────────────────────────────────
    def list_items(self) -> list[tuple[str, int]]:
        """[(품목, 현재재고), ...] 를 엑셀에 적힌 순서대로 돌려준다."""
        sheet = _stock_sheet(self._open())
        return [(name, _to_stock(cell.value, name)) for name, cell in _stock_rows(sheet)]

    def history(self, limit: int = 50) -> list[tuple]:
        """입출고 이력을 최신순으로 최대 limit건 돌려준다."""
        sheet = self._open()[HISTORY_SHEET]
        rows = sheet.iter_rows(min_row=2, max_col=len(HISTORY_HEADER), values_only=True)
        return [row for row in rows if row[0] is not None][::-1][:limit]

    # ── 재고 변경 ─────────────────────────────────────────────────────
    def receive(self, name: str, quantity) -> tuple[int, int]:
        """입고: 재고를 quantity만큼 늘리고 (변경 전, 변경 후) 를 돌려준다."""
        return self._change("입고", name, quantity)

    def issue(self, name: str, quantity) -> tuple[int, int]:
        """출고: 재고를 quantity만큼 줄인다. 현재 재고보다 많이 출고할 수는 없다."""
        return self._change("출고", name, quantity)

    def add_item(self, name: str) -> str:
        """새 품목을 재고 0개로 등록한다. 수량은 이후 입고로 채운다."""
        name = str(name or "").strip()
        if not name:
            raise InventoryError("품목 이름을 입력하세요.")
        with self._lock:
            book = self._open()
            sheet = _stock_sheet(book)
            if any(item == name for item, _ in _stock_rows(sheet)):
                raise InventoryError(f"'{name}'은(는) 이미 등록된 품목입니다.")
            sheet.append([name, 0])
            _log(book, "등록", name, 0, 0, 0)
            self._save(book)
        return name

    # ── 내부 구현 ─────────────────────────────────────────────────────
    def _change(self, kind: str, name: str, quantity) -> tuple[int, int]:
        quantity = parse_quantity(quantity)
        with self._lock:
            book = self._open()
            for item, cell in _stock_rows(_stock_sheet(book)):
                if item == name:
                    break
            else:
                raise InventoryError(f"'{name}' 품목을 찾을 수 없습니다.")
            before = _to_stock(cell.value, name)
            after = before + quantity if kind == "입고" else before - quantity
            if after < 0:
                raise InventoryError(
                    f"재고가 부족해 출고할 수 없습니다.\n{name}: 현재 재고 {before}개 / 출고 요청 {quantity}개"
                )
            cell.value = after
            _log(book, kind, name, quantity, before, after)
            self._save(book)
        return before, after

    def _open(self):
        try:
            book = load_workbook(self.path)
        except Exception as error:  # 파일 없음·손상·형식 오류를 모두 같은 안내로 처리
            raise InventoryError(f"엑셀 파일({self.path.name})을 읽을 수 없습니다.\n({error})") from None
        if HISTORY_SHEET not in book.sheetnames:
            sheet = book.create_sheet(HISTORY_SHEET)
            sheet.append(HISTORY_HEADER)
            sheet.column_dimensions["A"].width = 20
        return book

    def _save(self, book) -> None:
        # 임시 파일에 먼저 쓰고 교체한다 → 저장 도중 문제가 생겨도 원본 엑셀은 깨지지 않는다
        temp = self.path.with_name(self.path.stem + ".저장중" + self.path.suffix)
        try:
            book.save(temp)
            os.replace(temp, self.path)
        except OSError as error:
            raise InventoryError(
                "엑셀 파일에 저장하지 못했습니다.\n"
                f"Excel에서 '{self.path.name}' 파일이 열려 있다면 닫고 다시 시도하세요.\n(원인: {error.strerror or error})"
            ) from None
        finally:
            temp.unlink(missing_ok=True)

    def _create_file(self) -> None:
        book = Workbook()
        sheet = book.active
        sheet.title = STOCK_SHEET
        sheet.append(STOCK_HEADER)
        for row in SAMPLE_ITEMS:
            sheet.append(row)
        sheet.column_dimensions["A"].width = 20
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._save(book)


def _stock_sheet(book):
    return book[STOCK_SHEET] if STOCK_SHEET in book.sheetnames else book.worksheets[0]


def _stock_rows(sheet):
    """재고 시트에서 (품목, 수량 셀) 을 차례로 꺼낸다. 품목이 비어 있는 줄은 건너뛴다."""
    for name_cell, quantity_cell in sheet.iter_rows(min_row=2, max_col=2):
        name = str(name_cell.value or "").strip()
        if name:
            yield name, quantity_cell


def _to_stock(value, name: str) -> int:
    """엑셀 셀 값을 재고 수량(정수)으로 읽는다. 빈 칸은 0, 정수가 아니면 InventoryError."""
    if value is None or value == "":
        return 0
    try:
        number = float(value)
    except (TypeError, ValueError):
        number = float("nan")
    if not number.is_integer():
        raise InventoryError(f"엑셀에 적힌 '{name}'의 재고({value!r})가 정수가 아닙니다. 엑셀 파일을 확인하세요.")
    return int(number)


def _log(book, kind: str, name: str, quantity: int, before: int, after: int) -> None:
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    book[HISTORY_SHEET].append([now, kind, name, quantity, before, after])
