"""소모품 재고 관리 - 웹 화면 (Flask).

실행:  docker compose up -d   →   브라우저에서 http://localhost:8000
이 파일은 요청을 받아 inventory.py 에 넘기고 결과를 화면에 보여 주는 일만 한다.
"""
import os

from flask import Flask, flash, redirect, render_template, request, url_for

from inventory import Inventory, InventoryError

DEFAULT_FILE = "data/소모품_재고.xlsx"


def create_app(excel_path=None) -> Flask:
    app = Flask(__name__)
    app.secret_key = os.urandom(16)  # 처리 결과 알림(flash)을 다음 화면까지 전달하는 데만 쓴다
    inventory = Inventory(excel_path or os.environ.get("INVENTORY_FILE", DEFAULT_FILE))

    @app.get("/")
    def index():
        """재고 조회: 품목별 현재 재고와 최근 입출고 이력을 보여 준다."""
        keyword = request.args.get("q", "").strip()
        try:
            items, history = inventory.list_items(), inventory.history(limit=10)
        except InventoryError as error:
            items, history = [], []
            flash(str(error), "error")
        return render_template(
            "index.html",
            items=[(name, stock) for name, stock in items if keyword.lower() in name.lower()],
            item_count=len(items),
            total=sum(stock for _, stock in items),
            history=history,
            keyword=keyword,
            file_name=inventory.path.name,
        )

    @app.post("/change")
    def change():
        """재고 변경: 선택한 품목을 입력한 수량만큼 입고(+) 또는 출고(-) 처리한다."""
        name, kind = request.form.get("item", ""), request.form.get("kind", "")
        try:
            if kind not in ("입고", "출고"):
                raise InventoryError("입고 또는 출고를 선택하세요.")
            action = inventory.receive if kind == "입고" else inventory.issue
            before, after = action(name, request.form.get("quantity", ""))
            flash(f"{name} {abs(after - before)}개 {kind} 완료 (재고 {before}개 → {after}개)", "ok")
        except InventoryError as error:
            flash(str(error), "error")
        return back_to_list()

    @app.post("/items")
    def add_item():
        """품목 추가: 새 품목을 재고 0개로 등록한다."""
        try:
            name = inventory.add_item(request.form.get("name", ""))
            flash(f"'{name}' 품목을 등록했습니다. 입고로 수량을 넣으세요.", "ok")
        except InventoryError as error:
            flash(str(error), "error")
        return back_to_list()

    def back_to_list():
        # 처리 후 목록 화면으로 이동시킨다 → 새로고침(F5)을 눌러도 같은 입·출고가 두 번 처리되지 않는다
        return redirect(url_for("index", q=request.form.get("q") or None))

    return app


if __name__ == "__main__":  # Docker 없이 바로 실행할 때:  python app.py
    create_app().run(host="0.0.0.0", port=8000)
