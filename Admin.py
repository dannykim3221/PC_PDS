"""
피시방 주문 시스템 - 관리자 화면 (Admin.py)

실행 방법:
    python Admin.py

의존 라이브러리:
    pip install openpyxl
"""

import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db_manager as db

db.init_db()

# ── 색상 테마 ──────────────────────────────────────────────
BG_MAIN    = "#0D1117"
BG_CARD    = "#161B22"
BG_ACCENT  = "#21262D"
FG_MAIN    = "#C9D1D9"
FG_GREEN   = "#3FB950"
FG_RED     = "#F85149"
FG_BLUE    = "#58A6FF"
FG_YELLOW  = "#D29922"

POLL_MS = 3000   # 주문 목록 자동 갱신 주기 (ms)


# ══════════════════════════════════════════════════════════
#  관리자 앱
# ══════════════════════════════════════════════════════════

class AdminApp:
    """
    레이아웃:
        ┌──────────────────────────────────────────────┐
        │  헤더                                        │
        ├──────────────────────┬───────────────────────┤
        │  주문 대기열          │  재고 관리            │
        │  (처리 / 취소 버튼)   │  (수량 수정 + 저장)   │
        ├──────────────────────┴───────────────────────┤
        │  매출 현황                                    │
        └──────────────────────────────────────────────┘
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🛠️ 피시방 관리자 화면")
        self.root.geometry("1300x750")
        self.root.configure(bg=BG_MAIN)
        self.root.resizable(True, True)
        self.root.minsize(1100, 650)

        self._build_ui()
        self._load_inventory()
        self._poll_orders()   # 자동 갱신 시작

    # ══════════════ UI 빌드 ══════════════

    def _build_ui(self):
        self._build_header()
        body = tk.Frame(self.root, bg=BG_MAIN)
        body.pack(fill="both", expand=True, padx=10, pady=6)
        self._build_order_panel(body)
        self._build_inventory_panel(body)
        self._build_sales_panel()

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BG_ACCENT, pady=10)
        hdr.pack(fill="x")

        tk.Label(
            hdr, text="🛠️  관리자 대시보드",
            font=("맑은 고딕", 16, "bold"),
            bg=BG_ACCENT, fg=FG_BLUE,
        ).pack(side="left", padx=20)

        tk.Label(
            hdr, text="⚙️ 관리자 전용 화면",
            font=("맑은 고딕", 10),
            bg=BG_ACCENT, fg=FG_RED,
        ).pack(side="right", padx=20)

    # ── 주문 대기열 패널 ──
    def _build_order_panel(self, parent):
        frame = tk.LabelFrame(
            parent, text=" 주문 대기열 (자동 갱신) ",
            font=("맑은 고딕", 11, "bold"),
            bg=BG_CARD, fg=FG_GREEN, bd=2, relief="groove",
        )
        frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        cols = ("번호", "좌석", "메뉴", "금액", "상태", "시간")
        self.order_tree = ttk.Treeview(
            frame, columns=cols, show="headings", height=14,
        )
        widths = (55, 55, 220, 80, 65, 85)
        for col, w in zip(cols, widths):
            self.order_tree.heading(col, text=col)
            self.order_tree.column(col, width=w, anchor="center")

        self.order_tree.tag_configure("대기", foreground=FG_YELLOW)
        self.order_tree.tag_configure("완료", foreground=FG_GREEN)
        self.order_tree.tag_configure("취소", foreground=FG_RED)

        sb = ttk.Scrollbar(frame, orient="vertical",
                           command=self.order_tree.yview)
        self.order_tree.configure(yscrollcommand=sb.set)
        self.order_tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        sb.pack(side="left", fill="y", pady=6, padx=(0, 6))

        # 버튼 행
        btn_row = tk.Frame(frame, bg=BG_CARD)
        btn_row.pack(fill="x", padx=6, pady=(0, 6))

        tk.Button(
            btn_row, text="✅  주문 처리 (완료)",
            command=self._process_order,
            bg=FG_GREEN, fg="black",
            font=("맑은 고딕", 10, "bold"),
            relief="flat", cursor="hand2", pady=6,
        ).pack(side="left", expand=True, fill="x", padx=(0, 3))

        tk.Button(
            btn_row, text="❌  주문 취소",
            command=self._cancel_order,
            bg=FG_RED, fg="white",
            font=("맑은 고딕", 10, "bold"),
            relief="flat", cursor="hand2", pady=6,
        ).pack(side="left", expand=True, fill="x", padx=(3, 0))

        # 필터 라디오
        filter_row = tk.Frame(frame, bg=BG_CARD)
        filter_row.pack(fill="x", padx=6, pady=(0, 6))

        self.filter_var = tk.StringVar(value="대기")
        for label, val in [("대기만", "대기"), ("전체", ""), ("완료", "완료"), ("취소", "취소")]:
            tk.Radiobutton(
                filter_row, text=label, variable=self.filter_var, value=val,
                command=self._load_orders,
                bg=BG_CARD, fg=FG_MAIN, selectcolor=BG_ACCENT,
                activebackground=BG_CARD, activeforeground=FG_MAIN,
                font=("맑은 고딕", 9),
            ).pack(side="left", padx=6)

    # ── 재고 관리 패널 ──
    def _build_inventory_panel(self, parent):
        frame = tk.LabelFrame(
            parent, text=" 재고 관리 ",
            font=("맑은 고딕", 11, "bold"),
            bg=BG_CARD, fg=FG_GREEN, bd=2, relief="groove", width=420,
        )
        frame.pack(side="right", fill="both", padx=(5, 0))
        frame.pack_propagate(False)

        cols = ("메뉴", "카테고리", "가격", "재고", "판매")
        self.inv_tree = ttk.Treeview(
            frame, columns=cols, show="headings", height=14,
        )
        widths = (120, 85, 80, 60, 60)
        for col, w in zip(cols, widths):
            self.inv_tree.heading(col, text=col)
            self.inv_tree.column(col, width=w, anchor="center")

        self.inv_tree.tag_configure("low", foreground=FG_RED)

        sb2 = ttk.Scrollbar(frame, orient="vertical",
                            command=self.inv_tree.yview)
        self.inv_tree.configure(yscrollcommand=sb2.set)
        self.inv_tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        sb2.pack(side="left", fill="y", pady=6, padx=(0, 6))

        # 재고 수정 영역
        edit_row = tk.Frame(frame, bg=BG_CARD)
        edit_row.pack(fill="x", padx=6, pady=2)

        tk.Label(edit_row, text="재고 변경 →",
                 bg=BG_CARD, fg=FG_MAIN, font=("맑은 고딕", 9)).pack(side="left")

        self.inv_edit_var = tk.StringVar()
        tk.Entry(
            edit_row, textvariable=self.inv_edit_var, width=6,
            bg=BG_ACCENT, fg=FG_MAIN, insertbackground=FG_MAIN,
        ).pack(side="left", padx=4)

        tk.Button(
            edit_row, text="저장",
            command=self._save_stock,
            bg=FG_BLUE, fg="black",
            font=("맑은 고딕", 9, "bold"),
            relief="flat", cursor="hand2",
        ).pack(side="left")

        tk.Button(
            frame, text="🔄 재고 새로고침",
            command=self._load_inventory,
            bg=BG_ACCENT, fg=FG_MAIN,
            font=("맑은 고딕", 9), relief="flat", cursor="hand2",
        ).pack(pady=(2, 6))

    # ── 매출 현황 패널 ──
    def _build_sales_panel(self):
        frame = tk.LabelFrame(
            self.root, text=" 매출 현황 ",
            font=("맑은 고딕", 10, "bold"),
            bg=BG_CARD, fg=FG_GREEN, bd=2, relief="groove",
        )
        frame.pack(fill="x", padx=10, pady=(0, 8))

        self.sales_label = tk.Label(
            frame, text="총 매출 (완료 주문): 0 원",
            font=("맑은 고딕", 12, "bold"),
            bg=BG_CARD, fg=FG_BLUE,
        )
        self.sales_label.pack(pady=8)

    # ══════════════ 데이터 로드 ══════════════

    def _load_orders(self):
        """DB에서 주문 목록을 읽어 트리뷰를 갱신합니다."""
        for row in self.order_tree.get_children():
            self.order_tree.delete(row)

        status_filter = self.filter_var.get() or None
        orders = db.get_orders(status_filter)

        for o in orders:
            time_str = o["created_at"][11:16] if o["created_at"] else ""
            self.order_tree.insert(
                "", tk.END,
                iid=str(o["id"]),
                values=(
                    f"#{o['id']}",
                    f"{o['seat']}석",
                    o["items"] or "",
                    f"{o['total']:,}원",
                    o["status"],
                    time_str,
                ),
                tags=(o["status"],),
            )

        # 매출 갱신
        sales = db.get_total_sales()
        self.sales_label.config(text=f"총 매출 (완료 주문): {sales:,} 원")

    def _load_inventory(self):
        """DB에서 재고 목록을 읽어 트리뷰를 갱신합니다."""
        for row in self.inv_tree.get_children():
            self.inv_tree.delete(row)

        for m in db.get_menu():
            tag = "low" if m["stock"] < 5 else ""
            self.inv_tree.insert(
                "", tk.END,
                values=(m["name"], m["category"], f"{m['price']:,}원",
                        m["stock"], m["sold"]),
                tags=(tag,),
            )

    # ══════════════ 주문 처리 / 취소 ══════════════

    def _get_selected_order_id(self):
        """트리뷰에서 선택된 주문 id 반환. 없으면 None."""
        sel = self.order_tree.selection()
        if not sel:
            messagebox.showwarning("경고", "처리할 주문을 선택하세요.")
            return None
        return int(sel[0])

    def _process_order(self):
        """선택한 주문을 '완료' 처리합니다."""
        order_id = self._get_selected_order_id()
        if order_id is None:
            return

        # 현재 상태 확인
        orders = {o["id"]: o for o in db.get_orders()}
        order = orders.get(order_id)
        if not order:
            return

        if order["status"] != "대기":
            messagebox.showinfo("알림", f"이미 '{order['status']}' 상태인 주문입니다.")
            return

        if messagebox.askyesno(
            "주문 처리",
            f"#{order_id}번 주문을 완료 처리하시겠습니까?\n"
            f"좌석: {order['seat']}  |  {order['items']}"
        ):
            db.update_order_status(order_id, "완료")
            self._load_orders()
            self._load_inventory()
            messagebox.showinfo("완료", f"#{order_id}번 주문이 처리되었습니다. ✅")

    def _cancel_order(self):
        """선택한 주문을 '취소' 처리하고 재고를 복구합니다."""
        order_id = self._get_selected_order_id()
        if order_id is None:
            return

        orders = {o["id"]: o for o in db.get_orders()}
        order = orders.get(order_id)
        if not order:
            return

        if order["status"] != "대기":
            messagebox.showinfo("알림", f"이미 '{order['status']}' 상태인 주문은 취소할 수 없습니다.")
            return

        if messagebox.askyesno(
            "주문 취소",
            f"#{order_id}번 주문을 취소하시겠습니까?\n"
            f"재고가 복구됩니다.\n\n"
            f"좌석: {order['seat']}  |  {order['items']}"
        ):
            # 주문 항목 조회 후 재고 복구
            items = db.get_order_items(order_id)
            for item in items:
                db.restore_stock(item["name"], item["qty"])

            db.update_order_status(order_id, "취소")
            self._load_orders()
            self._load_inventory()
            messagebox.showinfo("취소 완료", f"#{order_id}번 주문이 취소되고 재고가 복구되었습니다. ❌")

    # ══════════════ 재고 직접 수정 ══════════════

    def _save_stock(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showwarning("경고", "수정할 메뉴를 선택하세요.")
            return
        try:
            new_stock = int(self.inv_edit_var.get())
            if new_stock < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("경고", "재고 수량은 0 이상의 숫자로 입력하세요.")
            return

        name = self.inv_tree.item(sel[0])["values"][0]
        db.update_stock(name, new_stock)
        self.inv_edit_var.set("")
        self._load_inventory()
        messagebox.showinfo("저장 완료", f"[{name}] 재고 → {new_stock}개로 저장되었습니다.")

    # ══════════════ 자동 갱신 ══════════════

    def _poll_orders(self):
        """POLL_MS 마다 주문 목록을 자동 갱신합니다."""
        self._load_orders()
        self.root.after(POLL_MS, self._poll_orders)


# ══════════════════════════════════════════════════════════
#  진입점
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = AdminApp(root)
    root.mainloop()