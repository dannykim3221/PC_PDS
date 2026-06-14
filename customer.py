"""
피시방 주문 시스템 - 주문자 화면 (customer.py)

[자료구조 개선 사항]
  - 장바구니: dict → OrderedDict (삽입 순서 보장 + 수량 상한 검증)
  - 재고 확인: menu_map 조회 → db.get_stock() (인메모리 dict 캐시, O(1))
  - 주문 상태 폴링 → 옵저버 패턴으로 교체
  - 메뉴 자동 갱신 → 옵저버 패턴으로 교체

[색상 테마 — 웜 다크 앰버]
  배경: 따뜻한 다크 브라운 계열 3단 계층
  강조: 앰버(황금색) 포인트 — 버튼·선택 강조에만 사용
  글자: 흰색(밝은 배경 위) / 검정(흰 버튼 위) 두 가지만
  메뉴 버튼: 카테고리별 채도 낮은 단색 면 + 흰 글자
"""

import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db_manager as db

db.init_db()

# ── 색상 테마 (웜 다크 앰버) ────────────────────────────────
BG_MAIN    = "#13100D"   # 최상위 배경 — 짙은 에스프레소
BG_CARD    = "#1D1812"   # 카드/패널 배경
BG_SURFACE = "#28211A"   # 헤더·입력칸 배경
BG_BORDER  = "#3D3228"   # 구분선

BTN_AMBER  = "#C07A00"   # 주문 버튼 — 진한 앰버
BTN_MUTED  = "#3D3228"   # 보조 버튼 — 뮤트 브라운
BTN_SOLD   = "#2A2520"   # 품절 버튼

FG_WHITE = "#FFFFFF"
FG_BLACK = "#13100D"

# 카테고리별 메뉴 버튼 배경 (채도 낮은 단색)
CATEGORY_COLORS = {
    "라면류":   "#2A1A0A",   # 딥 번트 오렌지
    "밥류":     "#1A280A",   # 딥 올리브 그린
    "분식·스낵": "#2A0A1A",  # 딥 버건디
    "음료":     "#0A1A2A",   # 딥 네이비
}

MAX_QTY_PER_ITEM = 10
FALLBACK_POLL_MS = 10000


class CustomerApp:
    """
    레이아웃:
        ┌─────────────────────────────────────────────┐
        │  헤더 (제목 + 좌석번호)                      │
        ├───────────────────────┬─────────────────────┤
        │  메뉴 패널             │  장바구니 패널       │
        │  (카테고리 탭 + 버튼)  │  (항목 + 합계)      │
        ├───────────────────────┴─────────────────────┤
        │  내 주문 현황 (옵저버로 실시간 갱신)          │
        └─────────────────────────────────────────────┘
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PC방 주문 시스템")
        self.root.geometry("1100x750")
        self.root.configure(bg=BG_MAIN)
        self.root.resizable(True, True)
        self.root.minsize(900, 650)

        # ── [자료구조 2] 장바구니 — OrderedDict ──────────────────
        self.cart: OrderedDict[str, int] = OrderedDict()
        self.menu_map: dict = {}
        self.my_order_ids: list = []
        self.btn_map: dict[str, list] = {}   # { 메뉴명: [버튼위젯, ...] }
        self.seat_var = tk.StringVar(value="1")

        self._apply_style()
        self._build_ui()
        self._load_menu()

        db.subscribe(self._on_db_event)
        self._fallback_poll()

    def __del__(self):
        db.unsubscribe(self._on_db_event)

    # ══════════════ 전역 스타일 ══════════════

    def _apply_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure("Warm.Treeview",
                         background=BG_CARD,
                         foreground=FG_WHITE,
                         fieldbackground=BG_CARD,
                         rowheight=26,
                         font=("맑은 고딕", 9))
        style.map("Warm.Treeview",
                  background=[("selected", BG_SURFACE)],
                  foreground=[("selected", FG_WHITE)])
        style.configure("Warm.Treeview.Heading",
                         background=BG_SURFACE,
                         foreground=FG_WHITE,
                         font=("맑은 고딕", 9, "bold"),
                         relief="flat")
        style.map("Warm.Treeview.Heading",
                  background=[("active", BG_BORDER)])

        # 탭 — 기본: 뮤트 브라운 / 선택: 앰버 배경 + 검정 글자
        style.configure("TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("TNotebook.Tab",
                         background=BG_SURFACE,
                         foreground=FG_WHITE,
                         padding=[12, 6],
                         font=("맑은 고딕", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", BTN_AMBER)],
                  foreground=[("selected", FG_BLACK)])

    # ══════════════ 옵저버 콜백 ══════════════

    def _on_db_event(self, event: str, payload: dict):
        if event == "reset":
            self.root.after(0, self._on_reset)
        elif event in ("new_order", "order_updated"):
            self.root.after(0, self._refresh_status_ui)
        elif event == "menu_changed":
            self.root.after(0, self._load_menu)

    def _on_reset(self):
        self.cart.clear()
        self.my_order_ids.clear()
        self._refresh_cart_ui()
        self._load_menu()   # 초기화 시에는 탭 구조까지 완전 갱신
        for row in self.status_tree.get_children():
            self.status_tree.delete(row)

    def _fallback_poll(self):
        self._update_btn_states()   # 탭 유지, 버튼만 갱신 → 폴링 깜빡임 없음
        self._refresh_status_ui()
        self.root.after(FALLBACK_POLL_MS, self._fallback_poll)

    # ══════════════ UI 빌드 ══════════════

    def _build_ui(self):
        self._build_header()
        self._build_body()
        self._build_status_panel()

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BG_SURFACE, pady=12)
        hdr.pack(fill="x")

        tk.Label(hdr, text="PC방 주문 시스템",
                 font=("맑은 고딕", 18, "bold"),
                 bg=BG_SURFACE, fg=FG_WHITE).pack(side="left", padx=20)

        right = tk.Frame(hdr, bg=BG_SURFACE)
        right.pack(side="right", padx=20)

        tk.Label(right, text="좌석번호",
                 font=("맑은 고딕", 11),
                 bg=BG_SURFACE, fg=FG_WHITE).pack(side="left")

        tk.Spinbox(right, from_=1, to=50, textvariable=self.seat_var,
                   width=5, font=("맑은 고딕", 12, "bold"),
                   bg=BG_CARD, fg=FG_WHITE,
                   buttonbackground=BG_BORDER,
                   relief="flat").pack(side="left", padx=8)

    def _build_body(self):
        body = tk.Frame(self.root, bg=BG_MAIN)
        body.pack(fill="both", expand=True, padx=10, pady=8)
        self._build_menu_panel(body)
        self._build_cart_panel(body)

    def _build_menu_panel(self, parent):
        frame = tk.LabelFrame(parent, text="  메뉴  ",
                              font=("맑은 고딕", 10, "bold"),
                              bg=BG_CARD, fg=FG_WHITE,
                              bd=1, relief="solid")
        frame.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.notebook = ttk.Notebook(frame)
        self.notebook.pack(fill="both", expand=True, padx=6, pady=6)
        self.tab_frames: dict  = {}   # { 카테고리: 탭 최상위 Frame }
        self.canvas_map: dict  = {}   # { 카테고리: Canvas }  — 스크롤 갱신용

    def _build_cart_panel(self, parent):
        frame = tk.LabelFrame(parent, text="  장바구니  ",
                              font=("맑은 고딕", 10, "bold"),
                              bg=BG_CARD, fg=FG_WHITE,
                              bd=1, relief="solid", width=320)
        frame.pack(side="right", fill="y", padx=(6, 0))
        frame.pack_propagate(False)

        self.cart_listbox = tk.Listbox(
            frame, bg=BG_SURFACE, fg=FG_WHITE,
            font=("맑은 고딕", 10),
            selectbackground=BTN_AMBER, selectforeground=FG_BLACK,
            relief="flat", bd=0, activestyle="none")
        self.cart_listbox.pack(fill="both", expand=True, padx=6, pady=6)

        # 구분선
        tk.Frame(frame, bg=BG_BORDER, height=1).pack(fill="x", padx=6)

        self.total_label = tk.Label(frame, text="합계: 0 원",
                                    font=("맑은 고딕", 14, "bold"),
                                    bg=BG_CARD, fg=FG_WHITE)
        self.total_label.pack(pady=8)

        btn_f = tk.Frame(frame, bg=BG_CARD)
        btn_f.pack(fill="x", padx=6, pady=(0, 6))

        for text, cmd in [("선택 항목 제거", self._remove_cart_item),
                           ("장바구니 비우기", self._clear_cart)]:
            tk.Button(btn_f, text=text, command=cmd,
                      bg=BTN_MUTED, fg=FG_WHITE,
                      font=("맑은 고딕", 9),
                      relief="flat", cursor="hand2", pady=6,
                      ).pack(fill="x", pady=2)

        tk.Button(btn_f, text="주문하기",
                  command=self._submit_order,
                  bg=BTN_AMBER, fg=FG_BLACK,
                  font=("맑은 고딕", 13, "bold"),
                  relief="flat", cursor="hand2", pady=10,
                  ).pack(fill="x", pady=(10, 2))

    def _build_status_panel(self):
        frame = tk.LabelFrame(self.root, text="  내 주문 현황  ",
                              font=("맑은 고딕", 10, "bold"),
                              bg=BG_CARD, fg=FG_WHITE,
                              bd=1, relief="solid")
        frame.pack(fill="x", padx=10, pady=(0, 10))

        cols = ("주문번호", "좌석", "메뉴", "금액", "상태", "시간")
        self.status_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            height=4, style="Warm.Treeview",
        )
        widths = (70, 60, 260, 80, 70, 90)
        for col, w in zip(cols, widths):
            self.status_tree.heading(col, text=col)
            self.status_tree.column(col, width=w, anchor="center")

        # 상태 행 배경으로 구분 — 글자는 모두 흰색
        self.status_tree.tag_configure("대기",
                                       background="#2A2510", foreground=FG_WHITE)
        self.status_tree.tag_configure("완료",
                                       background="#162010", foreground=FG_WHITE)
        self.status_tree.tag_configure("취소",
                                       background="#2A1510", foreground=FG_WHITE)

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.status_tree.yview)
        self.status_tree.configure(yscrollcommand=sb.set)
        self.status_tree.pack(side="left", fill="x", expand=True, padx=6, pady=6)
        sb.pack(side="right", fill="y", pady=6)

    # ══════════════ 메뉴 렌더링 ══════════════

    def _make_scroll_tab(self, cat: str) -> tk.Frame:
        """
        탭 하나에 Canvas + 세로 Scrollbar를 만들고,
        버튼을 담을 btn_frame을 Canvas 위에 올려 반환합니다.
        마우스 휠 스크롤도 바인딩합니다.
        """
        outer = tk.Frame(self.notebook, bg=BG_CARD)
        self.notebook.add(outer, text=f"  {cat}  ")
        self.tab_frames[cat] = outer

        canvas = tk.Canvas(outer, bg=BG_CARD, highlightthickness=0)
        sb     = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)

        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        btn_frame = tk.Frame(canvas, bg=BG_CARD)
        win_id    = canvas.create_window((0, 0), window=btn_frame, anchor="nw")
        self.canvas_map[cat] = canvas

        # btn_frame 크기가 바뀌면 스크롤 영역 갱신
        def _on_frame_configure(e, c=canvas):
            c.configure(scrollregion=c.bbox("all"))

        # canvas 너비에 맞춰 btn_frame 너비 동기화
        def _on_canvas_configure(e, c=canvas, w=win_id):
            c.itemconfig(w, width=e.width)

        btn_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        # 마우스 휠 스크롤 — 탭 진입/이탈 시 바인딩 관리
        def _bind_wheel(e, c=canvas):
            c.bind_all("<MouseWheel>",
                       lambda ev, cv=c: cv.yview_scroll(-1*(ev.delta//120), "units"))
            c.bind_all("<Button-4>",
                       lambda ev, cv=c: cv.yview_scroll(-1, "units"))
            c.bind_all("<Button-5>",
                       lambda ev, cv=c: cv.yview_scroll(1, "units"))

        def _unbind_wheel(e, c=canvas):
            c.unbind_all("<MouseWheel>")
            c.unbind_all("<Button-4>")
            c.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)
        btn_frame.bind("<Enter>", _bind_wheel)
        btn_frame.bind("<Leave>", _unbind_wheel)

        return btn_frame

    def _load_menu(self):
        """
        DB에서 메뉴를 읽어 탭·버튼 구조를 처음 만들거나
        카테고리 목록이 바뀐 경우에 완전 재생성합니다.
        그 외에는 _update_btn_states()로 버튼 속성만 변경합니다.
        """
        self.menu_map = {m["name"]: m for m in db.get_menu()}

        db_cats = sorted(set(v["category"] for v in self.menu_map.values()))
        existing_tabs = [self.notebook.tab(t, "text").strip()
                         for t in self.notebook.tabs()]
        expected_tabs = ["전체"] + db_cats

        if existing_tabs == expected_tabs and self.btn_map:
            # 카테고리 동일 + 버튼이 이미 존재 → 속성만 업데이트
            self._update_btn_states()
            return

        # 탭 구조가 바뀌었거나 처음 실행 → 완전 재생성
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.tab_frames.clear()
        self.canvas_map.clear()
        self.btn_map.clear()

        for cat in ["전체"] + db_cats:
            btn_frame = self._make_scroll_tab(cat)

            items = (list(self.menu_map.items()) if cat == "전체"
                     else [(n, v) for n, v in self.menu_map.items()
                           if v["category"] == cat])
            for idx, (name, info) in enumerate(items):
                self._make_menu_btn(btn_frame, name, info, idx // 3, idx % 3)

    def _update_btn_states(self):
        """
        버튼 위젯을 destroy하지 않고 config()로 텍스트·색·상태만 바꿉니다.
        누른 버튼 하나만 시각적으로 바뀌고 나머지는 그대로여서 깜빡임이 없습니다.
        btn_map 구조: { 메뉴명: [tk.Button, ...] }  — 전체탭·카테고리탭 버튼 모두 포함
        """
        for name, btns in self.btn_map.items():
            info     = self.menu_map.get(name)
            if info is None:
                continue
            in_cart  = self.cart.get(name, 0)
            sold_out = info["stock"] <= 0
            bg_color = BTN_SOLD if sold_out else CATEGORY_COLORS.get(info["category"], BG_SURFACE)

            lines = [name, f"{info['price']:,}원"]
            if sold_out:
                lines.append("품절")
            elif in_cart > 0:
                lines.append(f"담김  {in_cart}개")

            for btn in btns:
                btn.config(
                    text="\n".join(lines),
                    bg=bg_color,
                    activebackground=bg_color,
                    state="disabled" if sold_out else "normal",
                    cursor="arrow" if sold_out else "hand2",
                    command=(lambda n=name, p=info["price"]: self._add_to_cart(n, p))
                             if not sold_out else None,
                )

    def _make_menu_btn(self, parent, name, info, row, col):
        in_cart  = self.cart.get(name, 0)
        sold_out = info["stock"] <= 0
        bg_color = BTN_SOLD if sold_out else CATEGORY_COLORS.get(info["category"], BG_SURFACE)

        lines = [name, f"{info['price']:,}원"]
        if sold_out:
            lines.append("품절")
        elif in_cart > 0:
            lines.append(f"담김  {in_cart}개")

        frame = tk.Frame(parent, bg=BG_CARD, padx=4, pady=4)
        frame.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
        parent.columnconfigure(col, weight=1)

        btn = tk.Button(
            frame,
            text="\n".join(lines),
            command=(lambda n=name, p=info["price"]: self._add_to_cart(n, p))
                     if not sold_out else None,
            bg=bg_color,
            fg=FG_WHITE,
            font=("맑은 고딕", 10, "bold"),
            relief="flat",
            cursor="arrow" if sold_out else "hand2",
            width=12, height=4, wraplength=100,
            state="disabled" if sold_out else "normal",
            disabledforeground="#555555",
            activebackground=bg_color,   # 클릭 시 배경색 유지 → 흰색 깜빡임 방지
            activeforeground=FG_WHITE,
        )
        btn.pack(fill="both", expand=True)

        # btn_map에 추가 — 전체탭·카테고리탭 모두에 등록되므로 list로 관리
        if name not in self.btn_map:
            self.btn_map[name] = []
        self.btn_map[name].append(btn)

    # ══════════════ 장바구니 ══════════════

    def _add_to_cart(self, name: str, price: int):
        current_qty = self.cart.get(name, 0)
        if current_qty >= MAX_QTY_PER_ITEM:
            messagebox.showwarning("수량 초과",
                                   f"'{name}'은(는) 최대 {MAX_QTY_PER_ITEM}개까지만 담을 수 있습니다.")
            return
        if db.get_stock(name) <= current_qty:
            messagebox.showwarning("재고 부족", f"'{name}'의 재고가 부족합니다.")
            return
        self.cart[name] = current_qty + 1
        self._refresh_cart_ui()
        self._update_btn_states()   # 탭 유지, 버튼만 갱신 → 깜빡임 없음

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
        self._update_btn_states()   # 탭 유지, 버튼만 갱신 → 깜빡임 없음

    def _clear_cart(self):
        if not self.cart:
            return
        if messagebox.askyesno("확인", "장바구니를 비우시겠습니까?"):
            self.cart.clear()
            self._refresh_cart_ui()
            self._update_btn_states()   # 탭 유지, 버튼만 갱신 → 깜빡임 없음

    def _refresh_cart_ui(self):
        self.cart_listbox.delete(0, tk.END)
        total = 0
        for name, qty in self.cart.items():
            price    = self.menu_map[name]["price"]
            subtotal = price * qty
            total   += subtotal
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

        out_of_stock = [n for n, q in self.cart.items() if db.get_stock(n) < q]
        if out_of_stock:
            messagebox.showerror("재고 부족",
                                 f"재고 부족: {', '.join(out_of_stock)}\n"
                                 f"해당 항목을 제거 후 다시 주문하세요.")
            return

        order_id = db.create_order(seat, self.cart, self.menu_map)
        for name, qty in self.cart.items():
            db.reduce_stock(name, qty)

        self.my_order_ids.append(order_id)
        total = sum(self.menu_map[n]["price"] * q for n, q in self.cart.items())

        self.cart.clear()
        self._refresh_cart_ui()
        self._update_btn_states()   # 재고 반영, 탭 유지
        self._refresh_status_ui()

        messagebox.showinfo("주문 완료",
                            f"주문번호 [{order_id}]번이 접수되었습니다!\n"
                            f"좌석: {seat}번  |  합계: {total:,}원\n"
                            "잠시 후 준비해 드리겠습니다.")

    # ══════════════ 주문 현황 갱신 ══════════════

    def _refresh_status_ui(self):
        if not self.my_order_ids:
            return
        for row in self.status_tree.get_children():
            self.status_tree.delete(row)
        all_orders = {o["id"]: o for o in db.get_orders()}
        for oid in self.my_order_ids:
            o = all_orders.get(oid)
            if not o:
                continue
            time_str = o["created_at"][11:16] if o["created_at"] else ""
            self.status_tree.insert(
                "", tk.END,
                values=(f"#{o['id']}", f"{o['seat']}석",
                        o["items"] or "", f"{o['total']:,}원",
                        o["status"], time_str),
                tags=(o["status"],),
            )


if __name__ == "__main__":
    root = tk.Tk()
    app  = CustomerApp(root)
    root.mainloop()