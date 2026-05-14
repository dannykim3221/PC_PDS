"""
피시방 주문 시스템 - 공유 DB 관리자 (db_manager.py)

Customer.py 와 Admin.py 가 함께 사용하는 SQLite DB 레이어.
직접 실행하지 말고 import 해서 사용하세요.

DB 구조:
    menu    : 메뉴 정보 + 재고
    orders  : 주문 헤더 (주문번호, 좌석, 합계, 상태)
    order_items : 주문 상세 (메뉴명, 수량, 단가)
"""

import sqlite3
import os
from datetime import datetime
import openpyxl

# ── 경로 설정 ──────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
DB_PATH     = os.path.join(BASE_DIR, "pcroom.db")
EXCEL_PATH  = os.path.join(BASE_DIR, "inventory.xlsx")
SHEET_NAME  = "재고현황"


# ══════════════════════════════════════════════════════════
#  DB 초기화
# ══════════════════════════════════════════════════════════

def get_conn() -> sqlite3.Connection:
    """WAL 모드 커넥션 반환 (동시 읽기/쓰기 안전)."""
    conn = sqlite3.connect(DB_PATH, timeout=5)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """DB와 테이블이 없으면 생성하고, 엑셀에서 메뉴를 임포트합니다."""
    with get_conn() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS menu (
                name      TEXT PRIMARY KEY,
                category  TEXT NOT NULL,
                price     INTEGER NOT NULL,
                stock     INTEGER NOT NULL DEFAULT 0,
                sold      INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS orders (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                seat        TEXT NOT NULL,
                total       INTEGER NOT NULL,
                status      TEXT NOT NULL DEFAULT '대기',
                created_at  TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS order_items (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id  INTEGER NOT NULL REFERENCES orders(id),
                name      TEXT NOT NULL,
                qty       INTEGER NOT NULL,
                price     INTEGER NOT NULL
            );
        """)

    # 메뉴 테이블이 비어있으면 엑셀에서 가져오기
    with get_conn() as conn:
        count = conn.execute("SELECT COUNT(*) FROM menu").fetchone()[0]
        if count == 0:
            _import_menu_from_excel()


def _import_menu_from_excel():
    """inventory.xlsx → menu 테이블 초기 적재."""
    if not os.path.exists(EXCEL_PATH):
        return
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, category, price, stock, sold, *_ = row
        if name:
            rows.append((
                str(name),
                str(category) if category else "기타",
                int(price)    if price    else 0,
                int(stock)    if stock    else 0,
                int(sold)     if sold     else 0,
            ))
    wb.close()
    with get_conn() as conn:
        conn.executemany(
            "INSERT OR IGNORE INTO menu(name,category,price,stock,sold) VALUES(?,?,?,?,?)",
            rows,
        )


# ══════════════════════════════════════════════════════════
#  메뉴 / 재고
# ══════════════════════════════════════════════════════════

def get_menu() -> list[dict]:
    """전체 메뉴 목록 반환."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT name, category, price, stock, sold FROM menu ORDER BY category, name"
        ).fetchall()
    return [dict(r) for r in rows]


def reduce_stock(name: str, qty: int = 1) -> bool:
    """재고 차감. 성공 True / 재고 부족 False."""
    with get_conn() as conn:
        cur = conn.execute(
            "SELECT stock FROM menu WHERE name=?", (name,)
        )
        row = cur.fetchone()
        if not row or row["stock"] < qty:
            return False
        conn.execute(
            "UPDATE menu SET stock=stock-?, sold=sold+? WHERE name=?",
            (qty, qty, name),
        )
    _sync_excel_stock(name)
    return True


def restore_stock(name: str, qty: int = 1):
    """재고 복구 (주문 취소 시)."""
    with get_conn() as conn:
        conn.execute(
            "UPDATE menu SET stock=stock+?, sold=MAX(0,sold-?) WHERE name=?",
            (qty, qty, name),
        )
    _sync_excel_stock(name)


def update_stock(name: str, new_stock: int):
    """관리자 직접 재고 수정."""
    with get_conn() as conn:
        conn.execute(
            "UPDATE menu SET stock=? WHERE name=?", (new_stock, name)
        )
    _sync_excel_stock(name)


def _sync_excel_stock(name: str):
    """DB → 엑셀 단방향 동기화 (해당 메뉴 행만 갱신)."""
    if not os.path.exists(EXCEL_PATH):
        return
    with get_conn() as conn:
        row = conn.execute(
            "SELECT stock, sold FROM menu WHERE name=?", (name,)
        ).fetchone()
    if not row:
        return
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb[SHEET_NAME]
        for r in ws.iter_rows(min_row=2):
            if str(r[0].value) == name:
                r[3].value = row["stock"]   # D열
                r[4].value = row["sold"]    # E열
                break
        wb.save(EXCEL_PATH)
        wb.close()
    except Exception:
        pass   # 엑셀이 열려있거나 잠긴 경우 조용히 무시


# ══════════════════════════════════════════════════════════
#  주문
# ══════════════════════════════════════════════════════════

def create_order(seat: str, cart: dict, menu_map: dict) -> int:
    """
    장바구니를 DB에 저장하고 새 order_id 반환.
    cart   : {메뉴명: 수량}
    menu_map: {메뉴명: {price, ...}}
    """
    total = sum(menu_map[n]["price"] * q for n, q in cart.items())
    now   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO orders(seat, total, status, created_at) VALUES(?,?,?,?)",
            (seat, total, "대기", now),
        )
        order_id = cur.lastrowid
        items = [
            (order_id, n, q, menu_map[n]["price"])
            for n, q in cart.items()
        ]
        conn.executemany(
            "INSERT INTO order_items(order_id,name,qty,price) VALUES(?,?,?,?)",
            items,
        )
    return order_id


def get_orders(status_filter: str = None) -> list[dict]:
    """
    주문 목록 반환.
    status_filter: '대기' | '완료' | '취소' | None(전체)
    """
    sql = """
        SELECT o.id, o.seat, o.total, o.status, o.created_at,
               GROUP_CONCAT(oi.name || ' x' || oi.qty, ', ') AS items
        FROM orders o
        LEFT JOIN order_items oi ON o.id = oi.order_id
    """
    params = ()
    if status_filter:
        sql += " WHERE o.status=?"
        params = (status_filter,)
    sql += " GROUP BY o.id ORDER BY o.id DESC"

    with get_conn() as conn:
        rows = conn.execute(sql, params).fetchall()
    return [dict(r) for r in rows]


def get_order_items(order_id: int) -> list[dict]:
    """특정 주문의 상세 항목 반환."""
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT name, qty, price FROM order_items WHERE order_id=?",
            (order_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def update_order_status(order_id: int, status: str):
    """주문 상태 변경: '대기' → '완료' 또는 '취소'."""
    with get_conn() as conn:
        conn.execute(
            "UPDATE orders SET status=? WHERE id=?", (status, order_id)
        )


def get_total_sales() -> int:
    """완료된 주문의 총 매출 반환."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT COALESCE(SUM(total),0) FROM orders WHERE status='완료'"
        ).fetchone()
    return row[0]