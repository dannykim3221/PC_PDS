"""
피시방 주문 시스템 - 주문자 화면 (Customer.py)

실행 방법:
    python Customer.py

의존 라이브러리:
    pip install openpyxl
"""

import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys
from datetime import datetime

# db_manager 를 같은 폴더에서 import
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db_manager as db

# ── DB 초기화 ──────────────────────────────────────────────
db.init_db()

# ── 색상 테마 ──────────────────────────────────────────────
BG_MAIN     = "#1A1A2E"
BG_CARD     = "#16213E"
BG_ACCENT   = "#0F3460"
FG_MAIN     = "#E0E0E0"
FG_ACCENT   = "#E94560"
FG_SUCCESS  = "#4ECDC4"
BTN_ORDER   = "#E94560"
BTN_CANCEL  = "#555577"

CATEGORY_COLORS = {
    "음료": "#4A90D9",
    "식사": "#E67E22",
    "스낵": "#27AE60",
}

POLL_MS = 3000   # 주문 상태 자동 갱신 주기 (ms)


# ══════════════════════════════════════════════════════════
#  주문자 앱
# ══════════════════════════════════════════════════════════

class CustomerApp:
    """
    레이아웃:
        ┌─────────────────────────────────────────────┐
        │  헤더 (제목 + 좌석번호 + 새로고침)            │
        ├───────────────────────┬─────────────────────┤
        │  메뉴 패널             │  장바구니 패널       │
        │  (카테고리 탭 + 버튼)  │  (항목 + 합계)      │
        ├───────────────────────┴─────────────────────┤
        │  내 주문 현황 (DB 폴링으로 상태 자동 갱신)    │
        └─────────────────────────────────────────────┘
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("🎮 피시방 주문 시스템 - 고객 화면")
        self.root.geometry("1100x750")
        self.root.configure(bg=BG_MAIN)
        self.root.resizable(True, True)
        self.root.minsize(900, 650)

        self.cart: dict = {}        # {메뉴명: 수량}
        self.menu_map: dict = {}    # {메뉴명: {price, stock, category}}
        self.my_order_ids: list = []  # 이 세션에서 접수한 주문 id 목록
        self.seat_var = tk.StringVar(value="1")

        self._build_ui()
        self._load_menu()
        self._poll_order_status()   # 자동 갱신 시작

    # ══════════════ UI 빌드 ══════════════

    def _build_ui(self):
        self._build_header()
        self._build_body()
        self._build_status_panel()

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BG_ACCENT, pady=10)
        hdr.pack(fill="x")

        tk.Label(
            hdr, text="🎮  PC방 주문 시스템",
            font=("맑은 고딕", 18, "bold"),
            bg=BG_ACCENT, fg=FG_ACCENT,
        ).pack(side="left", padx=20)

        right = tk.Frame(hdr, bg=BG_ACCENT)
        right.pack(side="right", padx=20)

        tk.Label(right, text="좌석번호", font=("맑은 고딕", 11),
                 bg=BG_ACCENT, fg=FG_MAIN).pack(side="left")

        tk.Spinbox(
            right, from_=1, to=50, textvariable=self.seat_var,
            width=5, font=("맑은 고딕", 12, "bold"),
            bg=BG_CARD, fg=FG_ACCENT, buttonbackground=BG_ACCENT,
        ).pack(side="left", padx=8)

        tk.Button(
            right, text="🔄 메뉴 새로고침",
            command=self._load_menu,
            bg=BG_CARD, fg=FG_MAIN, font=("맑은 고딕", 9),
            relief="flat", cursor="hand2",
        ).pack(side="left")

    def _build_body(self):
        body = tk.Frame(self.root, bg=BG_MAIN)
        body.pack(fill="both", expand=True, padx=10, pady=8)
        self._build_menu_panel(body)
        self._build_cart_panel(body)

    # ── 메뉴 패널 ──
    def _build_menu_panel(self, parent):
        frame = tk.LabelFrame(
            parent, text=" 메뉴 ",
            font=("맑은 고딕", 11, "bold"),
            bg=BG_CARD, fg=FG_SUCCESS, bd=2, relief="groove",
        )
        frame.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TNotebook", background=BG_CARD, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background=BG_ACCENT, foreground=FG_MAIN,
            padding=[10, 5], font=("맑은 고딕", 10),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", BTN_ORDER)],
            foreground=[("selected", "white")],
        )
        self.tab_frames: dict = {}

    # ── 장바구니 패널 ──
    def _build_cart_panel(self, parent):
        frame = tk.LabelFrame(
            parent, text=" 장바구니 ",
            font=("맑은 고딕", 11, "bold"),
            bg=BG_CARD, fg=FG_SUCCESS, bd=2, relief="groove", width=320,
        )
        frame.pack(side="right", fill="y", padx=(6, 0))
        frame.pack_propagate(False)

        self.cart_listbox = tk.Listbox(
            frame, bg=BG_MAIN, fg=FG_MAIN,
            font=("맑은 고딕", 10),
            selectbackground=BG_ACCENT, selectforeground=FG_ACCENT,
            relief="flat", bd=0, activestyle="none",
        )
        self.cart_listbox.pack(fill="both", expand=True, padx=6, pady=6)

        self.total_label = tk.Label(
            frame, text="합계: 0 원",
            font=("맑은 고딕", 13, "bold"),
            bg=BG_CARD, fg=FG_ACCENT,
        )
        self.total_label.pack(pady=4)

        btn_f = tk.Frame(frame, bg=BG_CARD)
        btn_f.pack(fill="x", padx=6, pady=6)

        for text, cmd, color in [
            ("선택 항목 제거",  self._remove_cart_item, BTN_CANCEL),
            ("장바구니 비우기", self._clear_cart,        BTN_CANCEL),
        ]:
            tk.Button(
                btn_f, text=text, command=cmd,
                bg=color, fg=FG_MAIN, font=("맑은 고딕", 9),
                relief="flat", cursor="hand2", pady=5,
            ).pack(fill="x", pady=2)

        tk.Button(
            btn_f, text="✅  주문하기",
            command=self._submit_order,
            bg=BTN_ORDER, fg="white",
            font=("맑은 고딕", 12, "bold"),
            relief="flat", cursor="hand2", pady=8,
        ).pack(fill="x", pady=(8, 2))

    # ── 내 주문 현황 패널 ──
    def _build_status_panel(self):
        frame = tk.LabelFrame(
            self.root, text=" 내 주문 현황 (자동 갱신) ",
            font=("맑은 고딕", 10, "bold"),
            bg=BG_CARD, fg=FG_SUCCESS, bd=2, relief="groove",
        )
        frame.pack(fill="x", padx=10, pady=(0, 8))

        cols = ("주문번호", "좌석", "메뉴", "금액", "상태", "시간")
        self.status_tree = ttk.Treeview(
            frame, columns=cols, show="headings", height=4,
        )
        widths = (70, 60, 260, 80, 70, 90)
        for col, w in zip(cols, widths):
            self.status_tree.heading(col, text=col)
            self.status_tree.column(col, width=w, anchor="center")

        # 상태별 색상 태그
        self.status_tree.tag_configure("대기", foreground="#FFD700")
        self.status_tree.tag_configure("완료", foreground="#4ECDC4")
        self.status_tree.tag_configure("취소", foreground="#FF6B6B")

        sb = ttk.Scrollbar(frame, orient="vertical",
                           command=self.status_tree.yview)
        self.status_tree.configure(yscrollcommand=sb.set)
        self.status_tree.pack(side="left", fill="x", expand=True, padx=6, pady=6)
        sb.pack(side="right", fill="y", pady=6)

    # ══════════════ 메뉴 렌더링 ══════════════

    def _load_menu(self):
        """DB에서 메뉴를 읽어 탭 + 버튼을 다시 그립니다."""
        self.menu_map = {m["name"]: m for m in db.get_menu()}

        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.tab_frames.clear()

        categories = sorted(set(v["category"] for v in self.menu_map.values()))
        for cat in ["전체"] + categories:
            tab_frame = tk.Frame(self.notebook, bg=BG_CARD)
            self.notebook.add(tab_frame, text=f"  {cat}  ")
            self.tab_frames[cat] = tab_frame

            items = (
                list(self.menu_map.items()) if cat == "전체"
                else [(n, v) for n, v in self.menu_map.items()
                      if v["category"] == cat]
            )
            for idx, (name, info) in enumerate(items):
                self._make_menu_btn(tab_frame, name, info, idx // 3, idx % 3)

    def _make_menu_btn(self, parent, name, info, row, col):
        sold_out  = info["stock"] <= 0
        cat_color = CATEGORY_COLORS.get(info["category"], "#777799")
        stock_text = "품절" if sold_out else f"재고 {info['stock']}개"
        btn_text   = f"{name}\n{info['price']:,}원\n{stock_text}"

        frame = tk.Frame(parent, bg=BG_CARD, padx=4, pady=4)
        frame.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
        parent.columnconfigure(col, weight=1)

        btn = tk.Button(
            frame, text=btn_text,
            command=(lambda n=name, p=info["price"]: self._add_to_cart(n, p))
                     if not sold_out else None,
            bg="#444455" if sold_out else cat_color,
            fg="#888888" if sold_out else "white",
            font=("맑은 고딕", 10, "bold"), relief="flat",
            cursor="arrow" if sold_out else "hand2",
            width=12, height=4, wraplength=100,
            state="disabled" if sold_out else "normal",
        )
        btn.pack(fill="both", expand=True)

    # ══════════════ 장바구니 ══════════════

    def _add_to_cart(self, name, price):
        self.cart[name] = self.cart.get(name, 0) + 1
        self._refresh_cart_ui()

    def _remove_cart_item(self):
        sel = self.cart_listbox.curselection()
        if not sel:
            messagebox.showwarning("경고", "제거할 항목을 선택하세요.")
            return
        name = self.cart_listbox.get(sel[0]).split("  ")[0]
        if name in self.cart:
            self.cart[name] -= 1
            if self.cart[name] <= 0:
                del self.cart[name]
        self._refresh_cart_ui()

    def _clear_cart(self):
        if not self.cart:
            return
        if messagebox.askyesno("확인", "장바구니를 비우시겠습니까?"):
            self.cart.clear()
            self._refresh_cart_ui()

    def _refresh_cart_ui(self):
        self.cart_listbox.delete(0, tk.END)
        total = 0
        for name, qty in self.cart.items():
            price = self.menu_map[name]["price"]
            subtotal = price * qty
            total += subtotal
            self.cart_listbox.insert(tk.END, f"{name}  x{qty}  {subtotal:,}원")
        self.total_label.config(text=f"합계: {total:,} 원")

    # ══════════════ 주문 접수 ══════════════

    def _submit_order(self):
        if not self.cart:
            messagebox.showwarning("알림", "장바구니가 비어 있습니다.")
            return

        seat = self.seat_var.get().strip()
        if not seat:
            messagebox.showwarning("알림", "좌석번호를 입력하세요.")
            return

        # 재고 재확인
        out_of_stock = [
            n for n, q in self.cart.items()
            if self.menu_map[n]["stock"] < q
        ]
        if out_of_stock:
            messagebox.showerror(
                "재고 부족",
                f"재고 부족: {', '.join(out_of_stock)}\n해당 항목을 제거 후 다시 주문하세요."
            )
            return

        # DB에 주문 저장
        order_id = db.create_order(seat, self.cart, self.menu_map)

        # 재고 차감 (DB + 엑셀 동기화)
        for name, qty in self.cart.items():
            db.reduce_stock(name, qty)

        # 이번 세션 주문 목록에 추가
        self.my_order_ids.append(order_id)

        total = sum(self.menu_map[n]["price"] * q for n, q in self.cart.items())
        self.cart.clear()
        self._refresh_cart_ui()
        self._load_menu()          # 재고 반영된 메뉴 다시 렌더링
        self._refresh_status_ui()  # 주문 현황 즉시 갱신

        messagebox.showinfo(
            "주문 완료 ✅",
            f"주문번호 [{order_id}]번이 접수되었습니다!\n"
            f"좌석: {seat}번  |  합계: {total:,}원\n"
            "잠시 후 준비해 드리겠습니다 😊"
        )

    # ══════════════ 주문 현황 자동 갱신 ══════════════

    def _poll_order_status(self):
        """POLL_MS 마다 내 주문 상태를 DB에서 읽어 갱신합니다."""
        self._refresh_status_ui()
        self.root.after(POLL_MS, self._poll_order_status)

    def _refresh_status_ui(self):
        """내 주문 현황 트리뷰를 DB 최신 상태로 갱신합니다."""
        if not self.my_order_ids:
            return

        # 기존 행 전체 삭제 후 재삽입
        for row in self.status_tree.get_children():
            self.status_tree.delete(row)

        all_orders = {o["id"]: o for o in db.get_orders()}

        for oid in self.my_order_ids:
            o = all_orders.get(oid)
            if not o:
                continue
            # 시간에서 시:분만 추출
            time_str = o["created_at"][11:16] if o["created_at"] else ""
            self.status_tree.insert(
                "", tk.END,
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


# ══════════════════════════════════════════════════════════
#  진입점
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    root = tk.Tk()
    app = CustomerApp(root)
    root.mainloop()