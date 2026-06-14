"""
피시방 주문 시스템 - 관리자 화면 (admin.py)

[자료구조 개선 사항]
  - 폴링(3초 DB 반복 조회) → 옵저버 패턴으로 교체
    · subscribe()로 콜백 등록 → 변화 발생 시에만 UI 갱신
  - 메뉴 중복 확인 → db.menu_exists() (set, O(1))
  - 대기 주문 목록 → DB 직접 조회 (항상 최신 상태 보장)
  - 인기 메뉴 → db.get_top_menu() (Counter.most_common)

[색상 테마 — 슬레이트 모노크롬]
  배경: 짙은 슬레이트 3단 계층 (MAIN → CARD → SURFACE)
  강조: 순수 흰색 텍스트 / 순수 검정 텍스트 두 가지만 사용
  버튼:
    · 주요 동작  — 불투명 흰색 면 + 검정 글자
    · 위험 동작  — 짙은 적색 면 + 흰색 글자
    · 보조 동작  — 중간 슬레이트 면 + 흰색 글자
"""

import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import db_manager as db

db.init_db()

# ── 색상 테마 (슬레이트 모노크롬) ──────────────────────────
BG_MAIN    = "#111318"   # 최상위 배경 — 거의 검정에 가까운 슬레이트
BG_CARD    = "#1C2028"   # 카드/패널 배경
BG_SURFACE = "#252B35"   # 입력칸·헤더 배경
BG_BORDER  = "#323844"   # 구분선·테두리

# 버튼 배경색
BTN_PRIMARY = "#FFFFFF"  # 주요 동작 (흰색 면)
BTN_DANGER  = "#8B1A1A"  # 위험 동작 (딥 레드)
BTN_NEUTRAL = "#323844"  # 보조 동작 (슬레이트)
BTN_BLUE    = "#1A3A5C"  # 정보성 동작 (딥 네이비)

# 글자색 — 흰색·검정 두 가지만
FG_WHITE = "#FFFFFF"
FG_BLACK = "#111318"

# Treeview 행 상태 색 (배경으로 구분, 글자는 흰색 유지)
ROW_PENDING  = "#2A3040"  # 대기 — 슬레이트 블루
ROW_DONE     = "#1A2E20"  # 완료 — 딥 그린
ROW_CANCEL   = "#2E1A1A"  # 취소 — 딥 레드

FALLBACK_POLL_MS = 10000


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
        self.root.title("피시방 관리자 화면")
        self.root.geometry("1300x750")
        self.root.configure(bg=BG_MAIN)
        self.root.resizable(True, True)
        self.root.minsize(1100, 650)

        self._apply_treeview_style()
        self._build_ui()
        self._load_inventory()
        self._load_orders()

        db.subscribe(self._on_db_event)
        self._fallback_poll()

    def __del__(self):
        db.unsubscribe(self._on_db_event)

    # ══════════════ Treeview 전역 스타일 ══════════════

    def _apply_treeview_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        # Treeview 본문
        style.configure("Mono.Treeview",
                         background=BG_CARD,
                         foreground=FG_WHITE,
                         fieldbackground=BG_CARD,
                         rowheight=28,
                         font=("맑은 고딕", 9))
        style.map("Mono.Treeview",
                  background=[("selected", BG_SURFACE)],
                  foreground=[("selected", FG_WHITE)])

        # 헤더
        style.configure("Mono.Treeview.Heading",
                         background=BG_SURFACE,
                         foreground=FG_WHITE,
                         font=("맑은 고딕", 9, "bold"),
                         relief="flat",
                         borderwidth=0)
        style.map("Mono.Treeview.Heading",
                  background=[("active", BG_BORDER)])

        # Notebook (고객 탭)
        style.configure("TNotebook", background=BG_MAIN, borderwidth=0)
        style.configure("TNotebook.Tab",
                         background=BG_SURFACE,
                         foreground=FG_WHITE,
                         padding=[12, 6],
                         font=("맑은 고딕", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", BTN_PRIMARY)],
                  foreground=[("selected", FG_BLACK)])

    # ══════════════ 옵저버 콜백 ══════════════

    def _on_db_event(self, event: str, payload: dict):
        self.root.after(0, self._refresh_all)

    def _refresh_all(self):
        self._load_orders()
        self._load_inventory()

    def _fallback_poll(self):
        self._refresh_all()
        self.root.after(FALLBACK_POLL_MS, self._fallback_poll)

    # ══════════════ UI 빌드 ══════════════

    def _build_ui(self):
        self._build_header()
        body = tk.Frame(self.root, bg=BG_MAIN)
        body.pack(fill="both", expand=True, padx=10, pady=6)
        self._build_order_panel(body)
        self._build_inventory_panel(body)
        self._build_sales_panel()

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=BG_SURFACE, pady=12)
        hdr.pack(fill="x")

        # 왼쪽 — 타이틀
        left = tk.Frame(hdr, bg=BG_SURFACE)
        left.pack(side="left", padx=20)
        tk.Label(left, text="관리자 대시보드",
                 font=("맑은 고딕", 17, "bold"),
                 bg=BG_SURFACE, fg=FG_WHITE).pack(side="left")

        # 오른쪽 — 구분 레이블
        tk.Label(hdr, text="관리자 전용",
                 font=("맑은 고딕", 9),
                 bg=BG_SURFACE, fg=BG_BORDER).pack(side="right", padx=24)

    # ── 주문 대기열 패널 ──
    def _build_order_panel(self, parent):
        frame = tk.LabelFrame(
            parent, text="  주문 대기열  ",
            font=("맑은 고딕", 10, "bold"),
            bg=BG_CARD, fg=FG_WHITE,
            bd=1, relief="solid", highlightbackground=BG_BORDER,
        )
        frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        cols = ("번호", "좌석", "메뉴", "금액", "상태", "시간")
        self.order_tree = ttk.Treeview(
            frame, columns=cols, show="headings",
            height=14, style="Mono.Treeview",
        )
        widths = (55, 55, 220, 80, 65, 85)
        for col, w in zip(cols, widths):
            self.order_tree.heading(col, text=col)
            self.order_tree.column(col, width=w, anchor="center")

        # 상태별 행 배경색 — 글자는 모두 흰색
        self.order_tree.tag_configure("대기",
                                      background=ROW_PENDING, foreground=FG_WHITE)
        self.order_tree.tag_configure("완료",
                                      background=ROW_DONE,    foreground=FG_WHITE)
        self.order_tree.tag_configure("취소",
                                      background=ROW_CANCEL,  foreground=FG_WHITE)

        sb = ttk.Scrollbar(frame, orient="vertical", command=self.order_tree.yview)
        self.order_tree.configure(yscrollcommand=sb.set)
        self.order_tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
        sb.pack(side="left", fill="y", pady=6, padx=(0, 6))

        btn_row = tk.Frame(frame, bg=BG_CARD)
        btn_row.pack(fill="x", padx=6, pady=(0, 6))

        tk.Button(btn_row, text="주문 처리 (완료)",
                  command=self._process_order,
                  bg=BTN_PRIMARY, fg=FG_BLACK,
                  font=("맑은 고딕", 10, "bold"),
                  relief="flat", cursor="hand2", pady=7,
                  ).pack(side="left", expand=True, fill="x", padx=(0, 3))

        tk.Button(btn_row, text="주문 취소",
                  command=self._cancel_order,
                  bg=BTN_DANGER, fg=FG_WHITE,
                  font=("맑은 고딕", 10, "bold"),
                  relief="flat", cursor="hand2", pady=7,
                  ).pack(side="left", expand=True, fill="x", padx=(3, 0))

        filter_row = tk.Frame(frame, bg=BG_CARD)
        filter_row.pack(fill="x", padx=6, pady=(0, 8))

        self.filter_var = tk.StringVar(value="대기")
        for label, val in [("대기", "대기"), ("전체", ""), ("완료", "완료"), ("취소", "취소")]:
            tk.Radiobutton(
                filter_row, text=label, variable=self.filter_var, value=val,
                command=self._load_orders,
                bg=BG_CARD, fg=FG_WHITE,
                selectcolor=BG_SURFACE,
                activebackground=BG_CARD, activeforeground=FG_WHITE,
                font=("맑은 고딕", 9),
            ).pack(side="left", padx=6)

    # ── 재고 관리 패널 ──
    def _build_inventory_panel(self, parent):
        frame = tk.LabelFrame(
            parent, text="  메뉴 관리  ",
            font=("맑은 고딕", 10, "bold"),
            bg=BG_CARD, fg=FG_WHITE,
            bd=1, relief="solid", width=440,
        )
        frame.pack(side="right", fill="both", padx=(5, 0))
        frame.pack_propagate(False)

        tree_frame = tk.Frame(frame, bg=BG_CARD)
        tree_frame.pack(fill="both", expand=True, padx=6, pady=(6, 2))

        cols = ("메뉴", "카테고리", "가격", "재고", "판매")
        self.inv_tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings",
            height=9, style="Mono.Treeview",
        )
        widths = (110, 80, 75, 55, 55)
        for col, w in zip(cols, widths):
            self.inv_tree.heading(col, text=col)
            self.inv_tree.column(col, width=w, anchor="center")

        # 재고 부족 행 — 딥 레드 배경, 흰 글자
        self.inv_tree.tag_configure("low",
                                    background=ROW_CANCEL, foreground=FG_WHITE)
        self.inv_tree.bind("<<TreeviewSelect>>", self._on_menu_select)

        sb2 = ttk.Scrollbar(tree_frame, orient="vertical", command=self.inv_tree.yview)
        self.inv_tree.configure(yscrollcommand=sb2.set)
        self.inv_tree.pack(side="left", fill="both", expand=True)
        sb2.pack(side="left", fill="y")

        # ── 입력 폼 ──
        form = tk.Frame(frame, bg=BG_CARD)
        form.pack(fill="x", padx=8, pady=(0, 6))
        tk.Frame(form, bg=BG_BORDER, height=1).pack(fill="x", pady=(2, 6))

        fields = [("메뉴 이름", "name"), ("카테고리", "category"),
                  ("가격 (원)", "price"), ("재고 수량", "stock")]
        self._form_vars = {}
        for label, key in fields:
            row = tk.Frame(form, bg=BG_CARD)
            row.pack(fill="x", pady=2)
            tk.Label(row, text=label, width=9, anchor="w",
                     bg=BG_CARD, fg=FG_WHITE,
                     font=("맑은 고딕", 9)).pack(side="left")
            var = tk.StringVar()
            self._form_vars[key] = var
            if key == "category":
                combo = ttk.Combobox(row, textvariable=var, width=14,
                                     values=["음료", "식사", "스낵", "기타"],
                                     state="normal", font=("맑은 고딕", 9))
                combo.pack(side="left", padx=4)
            else:
                tk.Entry(row, textvariable=var, width=16,
                         bg=BG_SURFACE, fg=FG_WHITE,
                         insertbackground=FG_WHITE,
                         relief="flat",
                         font=("맑은 고딕", 9)).pack(side="left", padx=4)

        tk.Frame(form, bg=BG_BORDER, height=1).pack(fill="x", pady=6)

        btn_row1 = tk.Frame(form, bg=BG_CARD)
        btn_row1.pack(fill="x", pady=2)
        tk.Button(btn_row1, text="메뉴 추가", command=self._add_menu,
                  bg=BTN_BLUE, fg=FG_WHITE,
                  font=("맑은 고딕", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6,
                  ).pack(side="left", expand=True, fill="x", padx=(0, 2))
        tk.Button(btn_row1, text="메뉴 수정", command=self._edit_menu,
                  bg=BTN_PRIMARY, fg=FG_BLACK,
                  font=("맑은 고딕", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6,
                  ).pack(side="left", expand=True, fill="x", padx=(2, 0))

        btn_row2 = tk.Frame(form, bg=BG_CARD)
        btn_row2.pack(fill="x", pady=2)
        tk.Button(btn_row2, text="메뉴 삭제", command=self._delete_menu,
                  bg=BTN_DANGER, fg=FG_WHITE,
                  font=("맑은 고딕", 9, "bold"),
                  relief="flat", cursor="hand2", pady=6,
                  ).pack(side="left", expand=True, fill="x", padx=(0, 2))
        tk.Button(btn_row2, text="폼 초기화", command=self._clear_form,
                  bg=BTN_NEUTRAL, fg=FG_WHITE,
                  font=("맑은 고딕", 9),
                  relief="flat", cursor="hand2", pady=6,
                  ).pack(side="left", expand=True, fill="x", padx=(2, 0))

    # ── 매출 현황 패널 ──
    def _build_sales_panel(self):
        frame = tk.Frame(self.root, bg=BG_SURFACE, pady=10)
        frame.pack(fill="x", padx=10, pady=(0, 10))

        self.sales_label = tk.Label(
            frame,
            text="총 매출 (완료 주문): 0 원   |   인기 메뉴: -",
            font=("맑은 고딕", 11, "bold"),
            bg=BG_SURFACE, fg=FG_WHITE,
        )
        self.sales_label.pack(side="left", padx=16)

        tk.Button(frame, text="전체 초기화",
                  command=self._reset_all,
                  bg=BTN_DANGER, fg=FG_WHITE,
                  font=("맑은 고딕", 10, "bold"),
                  relief="flat", cursor="hand2", padx=14, pady=4,
                  ).pack(side="right", padx=(4, 16))

        tk.Button(frame, text="Excel 내보내기",
                  command=self._export_excel,
                  bg=BTN_NEUTRAL, fg=FG_WHITE,
                  font=("맑은 고딕", 10, "bold"),
                  relief="flat", cursor="hand2", padx=14, pady=4,
                  ).pack(side="right", padx=4)

    # ══════════════ 데이터 로드 ══════════════

    def _load_orders(self):
        for row in self.order_tree.get_children():
            self.order_tree.delete(row)

        status_filter = self.filter_var.get() or None

        # DB에서 직접 조회 — deque는 프로그램 재시작 시 비어있을 수 있으므로
        # 항상 DB를 기준으로 표시하고, deque는 인메모리 연산에만 활용
        for o in db.get_orders(status_filter):
            time_str = o["created_at"][11:16] if o["created_at"] else ""
            self.order_tree.insert(
                "", tk.END, iid=str(o["id"]),
                values=(f"#{o['id']}", f"{o['seat']}석",
                        o["items"] or "", f"{o['total']:,}원",
                        o["status"], time_str),
                tags=(o["status"],),
            )

        sales   = db.get_total_sales()
        top     = db.get_top_menu(3)
        top_str = ", ".join(f"{m['name']}({m['sold']})" for m in top) or "-"
        self.sales_label.config(
            text=f"총 매출 (완료 주문): {sales:,} 원   |   인기 메뉴: {top_str}"
        )

    def _load_inventory(self):
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
        sel = self.order_tree.selection()
        if not sel:
            messagebox.showwarning("경고", "처리할 주문을 선택하세요.")
            return None
        return int(sel[0])

    def _process_order(self):
        order_id = self._get_selected_order_id()
        if order_id is None:
            return
        orders = {o["id"]: o for o in db.get_orders()}
        order  = orders.get(order_id)
        if not order:
            return
        if order["status"] != "대기":
            messagebox.showinfo("알림", f"이미 '{order['status']}' 상태인 주문입니다.")
            return
        if messagebox.askyesno("주문 처리",
                               f"#{order_id}번 주문을 완료 처리하시겠습니까?\n"
                               f"좌석: {order['seat']}  |  {order['items']}"):
            db.update_order_status(order_id, "완료")
            messagebox.showinfo("완료", f"#{order_id}번 주문이 처리되었습니다.")

    def _cancel_order(self):
        order_id = self._get_selected_order_id()
        if order_id is None:
            return
        orders = {o["id"]: o for o in db.get_orders()}
        order  = orders.get(order_id)
        if not order:
            return
        if order["status"] != "대기":
            messagebox.showinfo("알림", f"이미 '{order['status']}' 상태인 주문은 취소할 수 없습니다.")
            return
        if messagebox.askyesno("주문 취소",
                               f"#{order_id}번 주문을 취소하시겠습니까?\n"
                               f"재고가 복구됩니다.\n\n"
                               f"좌석: {order['seat']}  |  {order['items']}"):
            items = db.get_order_items(order_id)
            for item in items:
                db.restore_stock(item["name"], item["qty"])
            db.update_order_status(order_id, "취소")
            messagebox.showinfo("취소 완료", f"#{order_id}번 주문이 취소되고 재고가 복구되었습니다.")

    # ══════════════ 메뉴 관리 ══════════════

    def _on_menu_select(self, event=None):
        sel = self.inv_tree.selection()
        if not sel:
            return
        vals = self.inv_tree.item(sel[0])["values"]
        self._form_vars["name"].set(vals[0])
        self._form_vars["category"].set(vals[1])
        price_clean = str(vals[2]).replace(",", "").replace("원", "").strip()
        self._form_vars["price"].set(price_clean)
        self._form_vars["stock"].set(vals[3])

    def _get_form_values(self):
        name     = self._form_vars["name"].get().strip()
        category = self._form_vars["category"].get().strip()
        price    = self._form_vars["price"].get().strip()
        stock    = self._form_vars["stock"].get().strip()
        if not name:
            messagebox.showwarning("경고", "메뉴 이름을 입력하세요."); return None
        if not category:
            messagebox.showwarning("경고", "카테고리를 입력하세요."); return None
        try:
            price = int(price)
            if price < 0: raise ValueError
        except ValueError:
            messagebox.showwarning("경고", "가격은 0 이상의 숫자로 입력하세요."); return None
        try:
            stock = int(stock)
            if stock < 0: raise ValueError
        except ValueError:
            messagebox.showwarning("경고", "재고 수량은 0 이상의 숫자로 입력하세요."); return None
        return name, category, price, stock

    def _add_menu(self):
        vals = self._get_form_values()
        if not vals:
            return
        name, category, price, stock = vals
        if db.menu_exists(name):
            messagebox.showwarning("경고",
                                   f"'{name}' 메뉴가 이미 존재합니다.\n"
                                   f"수정하려면 '메뉴 수정' 버튼을 사용하세요.")
            return
        db.add_menu(name, category, price, stock)
        self._clear_form()
        messagebox.showinfo("추가 완료", f"'{name}' 메뉴가 추가되었습니다.")

    def _edit_menu(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showwarning("경고", "수정할 메뉴를 목록에서 선택하세요.")
            return
        vals = self._get_form_values()
        if not vals:
            return
        new_name, category, price, stock = vals
        original_name = self.inv_tree.item(sel[0])["values"][0]

        if not messagebox.askyesno("메뉴 수정",
                                   f"'{original_name}' 메뉴를 아래와 같이 수정하시겠습니까?\n\n"
                                   f"  이름     : {new_name}\n"
                                   f"  카테고리 : {category}\n"
                                   f"  가격     : {price:,}원\n"
                                   f"  재고     : {stock}개"):
            return

        if original_name != new_name:
            sold = self.inv_tree.item(sel[0])["values"][4]
            db.delete_menu(original_name)
            db.add_menu(new_name, category, price, stock)
            with db.get_conn() as conn:
                conn.execute("UPDATE menu SET sold=? WHERE name=?", (sold, new_name))
            db.sales_counter[new_name] = sold
        else:
            db.update_menu(original_name, category, price)
            db.update_stock(original_name, stock)

        self._clear_form()
        messagebox.showinfo("수정 완료", f"'{new_name}' 메뉴가 수정되었습니다.")

    def _delete_menu(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showwarning("경고", "삭제할 메뉴를 목록에서 선택하세요.")
            return
        name = self.inv_tree.item(sel[0])["values"][0]

        pending_orders = db.get_orders("대기")
        blocking = []
        for o in pending_orders:
            items = db.get_order_items(o["id"])
            if any(item["name"] == name for item in items):
                blocking.append(f"#{o['id']}번 ({o['seat']}석)")
        if blocking:
            messagebox.showerror("삭제 불가",
                                 f"'{name}' 메뉴는 현재 대기 중인 주문에 포함되어 있어 삭제할 수 없습니다.\n\n"
                                 f"관련 주문: {', '.join(blocking)}\n\n"
                                 f"해당 주문을 먼저 처리하거나 취소한 후 삭제해 주세요.")
            return

        if not messagebox.askyesno("메뉴 삭제",
                                   f"'{name}' 메뉴를 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다."):
            return

        db.delete_menu(name)
        self._clear_form()
        messagebox.showinfo("삭제 완료", f"'{name}' 메뉴가 삭제되었습니다.")

    def _clear_form(self):
        for var in self._form_vars.values():
            var.set("")
        if self.inv_tree.selection():
            self.inv_tree.selection_remove(self.inv_tree.selection())

    # ══════════════ 전체 초기화 ══════════════

    def _reset_all(self):
        if not messagebox.askyesno(
            "전체 초기화",
            "재고·주문·매출 데이터를 모두 초기화하시겠습니까?\n\n"
            "기본 메뉴 목록은 다시 삽입됩니다.",
        ):
            return
        if not messagebox.askyesno(
            "최종 확인",
            "이 작업은 되돌릴 수 없습니다.\n정말 초기화하시겠습니까?",
            icon="warning",
        ):
            return
        db.reset_db()
        self._clear_form()
        messagebox.showinfo("초기화 완료", "모든 데이터가 초기화되었습니다.")

    # ══════════════ Excel 내보내기 ══════════════

    def _export_excel(self):
        try:
            path = db.export_to_excel()
            messagebox.showinfo("내보내기 완료",
                                f"Excel 파일이 저장되었습니다.\n\n"
                                f"경로: {path}\n\n"
                                f"포함 시트:\n"
                                f"  재고현황 / 주문내역 / 매출요약")
        except Exception as e:
            messagebox.showerror("오류", f"내보내기 실패:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app  = AdminApp(root)
    root.mainloop()