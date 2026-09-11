import json
import os
import re
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[2]
CATALOG_XLSX_PATH = Path(os.getenv("CATALOG_XLSX_PATH", str(ROOT / "AI穿搭导购_商品更新源.xlsx")))
DATABASE_PATH = Path(os.getenv("DATABASE_PATH", str(ROOT / "agent-backend" / "data" / "agent.db")))
DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
ALLOWED_ORIGINS = [item.strip() for item in os.getenv("ALLOWED_ORIGINS", "http://localhost:3000").split(",") if item.strip()]

app = FastAPI(title="AI 穿搭导购 Agent API", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type", "X-Request-ID"],
)


class ChatRequest(BaseModel):
    message: str = Field(min_length=1, max_length=500)
    session_id: str | None = None


class ImportRequest(BaseModel):
    source_path: str | None = None


@contextmanager
def db_connection():
    connection = sqlite3.connect(DATABASE_PATH)
    connection.row_factory = sqlite3.Row
    try:
        yield connection
        connection.commit()
    finally:
        connection.close()


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db() -> None:
    with db_connection() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS catalog_items (
                product_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                category TEXT NOT NULL,
                subcategory TEXT,
                styles TEXT,
                scenarios TEXT,
                color TEXT,
                fit TEXT,
                slim_tag TEXT,
                size_range TEXT,
                height_range TEXT,
                price REAL NOT NULL,
                inventory INTEGER NOT NULL,
                sale_status TEXT NOT NULL,
                recommendable TEXT NOT NULL,
                image_url TEXT,
                source TEXT,
                updated_at TEXT,
                note TEXT
            );
            CREATE TABLE IF NOT EXISTS sessions (
                session_id TEXT PRIMARY KEY,
                state_json TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS traces (
                trace_id TEXT PRIMARY KEY,
                session_id TEXT NOT NULL,
                user_message TEXT NOT NULL,
                state_json TEXT NOT NULL,
                tool_summary TEXT NOT NULL,
                response_json TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """
        )


def value(row: tuple[Any, ...], index: int, default: Any = "") -> Any:
    item = row[index] if len(row) > index else default
    return default if item is None else item


def import_catalog(source_path: Path) -> dict[str, Any]:
    if not source_path.exists() or source_path.suffix.lower() != ".xlsx":
        raise HTTPException(status_code=422, detail="商品更新源必须是可访问的 .xlsx 文件")
    workbook = load_workbook(source_path, data_only=False, read_only=True)
    if "商品主数据" not in workbook.sheetnames:
        raise HTTPException(status_code=422, detail="工作簿缺少“商品主数据”工作表")
    sheet = workbook["商品主数据"]
    rows = list(sheet.iter_rows(min_row=2, values_only=True))
    products: list[tuple[Any, ...]] = []
    for row in rows:
        product_id = str(value(row, 1)).strip()
        if not product_id:
            continue
        try:
            price = float(value(row, 12, 0))
            inventory = int(value(row, 13, 0))
        except (TypeError, ValueError) as error:
            raise HTTPException(status_code=422, detail=f"商品 {product_id} 的价格或库存无效") from error
        products.append((
            product_id, str(value(row, 2)), str(value(row, 3)), str(value(row, 4)),
            str(value(row, 5)), str(value(row, 6)), str(value(row, 7)), str(value(row, 8)),
            str(value(row, 9)), str(value(row, 10)), str(value(row, 11)), price, inventory,
            str(value(row, 14)), "是" if inventory > 0 and str(value(row, 14)) == "在售" else "否", str(value(row, 16)), str(value(row, 17)),
            str(value(row, 18)), str(value(row, 19)),
        ))
    if len(products) != 70:
        raise HTTPException(status_code=422, detail=f"商品主数据应为 70 条，当前读取到 {len(products)} 条")
    if len({item[0] for item in products}) != len(products):
        raise HTTPException(status_code=422, detail="商品 ID 存在重复，拒绝导入")
    with db_connection() as db:
        db.execute("DELETE FROM catalog_items")
        db.executemany(
            """
            INSERT INTO catalog_items (
                product_id,name,category,subcategory,styles,scenarios,color,fit,slim_tag,
                size_range,height_range,price,inventory,sale_status,recommendable,image_url,
                source,updated_at,note
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            products,
        )
    return {"source": str(source_path), "imported": len(products), "updated_at": now()}


def get_session(session_id: str) -> dict[str, Any]:
    with db_connection() as db:
        record = db.execute("SELECT state_json FROM sessions WHERE session_id = ?", (session_id,)).fetchone()
    default_state = {
        "scenario": None, "budget": None, "style": None, "height": None, "size": None,
        "slim": False, "excluded_categories": [], "question_count": 0,
    }
    return json.loads(record["state_json"]) if record else default_state


def save_session(session_id: str, state: dict[str, Any]) -> None:
    with db_connection() as db:
        db.execute(
            "INSERT INTO sessions (session_id,state_json,updated_at) VALUES (?,?,?) "
            "ON CONFLICT(session_id) DO UPDATE SET state_json=excluded.state_json,updated_at=excluded.updated_at",
            (session_id, json.dumps(state, ensure_ascii=False), now()),
        )


def parse_conditions(message: str, state: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
    changes: list[str] = []
    normalized = message.replace("￥", "¥").replace("以内", "内")
    for scenario in ["通勤", "约会", "旅行", "面试", "日常"]:
        if scenario in normalized and state.get("scenario") != scenario:
            state["scenario"] = scenario
            changes.append(f"场景更新为{scenario}")
    for style in ["简约", "法式", "运动", "休闲", "复古"]:
        if style in normalized and state.get("style") != style:
            state["style"] = style
            changes.append(f"风格更新为{style}")
    budget_match = re.search(r"(?:¥|预算|不超过|控制在)?\s*(\d{2,4})\s*(?:元|块|¥|以内|内)", normalized)
    if budget_match:
        state["budget"] = int(budget_match.group(1))
        changes.append(f"预算更新为¥{state['budget']}")
    elif any(token in normalized for token in ["再便宜", "便宜一点", "更便宜"]):
        previous = state.get("budget") or 500
        state["budget"] = max(200, previous - 100)
        changes.append(f"预算下调至¥{state['budget']}")
    height_match = re.search(r"(1[4-8]\d)\s*(?:cm|厘米|公分)?", normalized, flags=re.I)
    if height_match:
        state["height"] = int(height_match.group(1))
        changes.append(f"身高更新为{state['height']}cm")
    size_match = re.search(r"(?<![A-Z])(S|M|L|XL)(?:码|号)?", normalized, flags=re.I)
    if size_match:
        state["size"] = size_match.group(1).upper()
        changes.append(f"尺码更新为{state['size']}码")
    if "更显瘦" in normalized or "显瘦" in normalized:
        state["slim"] = True
        changes.append("已优先筛选显瘦版型")
    if "不要裙子" in normalized or "不穿裙子" in normalized:
        state["excluded_categories"] = list({*state["excluded_categories"], "半身裙", "连衣裙"})
        changes.append("已排除半身裙和连衣裙")
    return state, changes


def missing_fields(state: dict[str, Any]) -> list[str]:
    return [field for field in ["scenario", "budget", "style", "height", "size"] if not state.get(field)]


def ask_question(state: dict[str, Any]) -> str:
    missing = missing_fields(state)
    if "scenario" in missing:
        return "这套穿搭主要用于什么场景？例如通勤、约会、旅行或面试。"
    if "budget" in missing:
        return "你的总预算大约是多少？例如 ¥300 内或 ¥500 内。"
    if "style" in missing:
        return "你更偏好简约、法式、运动、休闲还是复古？"
    if "height" in missing or "size" in missing:
        return "请告诉我身高和常穿尺码，例如“165cm，M码”，我会把版型和尺码建议一起考虑。"
    return ""


def query_products(state: dict[str, Any]) -> list[dict[str, Any]]:
    clauses = ["sale_status = '在售'", "recommendable = '是'", "inventory > 0"]
    params: list[Any] = []
    if state.get("budget"):
        clauses.append("price <= ?")
        params.append(state["budget"])
    if state.get("size"):
        clauses.append("size_range LIKE ?")
        params.append(f"%{state['size']}%")
    for category in state.get("excluded_categories", []):
        clauses.append("category != ?")
        params.append(category)
    if state.get("scenario"):
        clauses.append("scenarios LIKE ?")
        params.append(f"%{state['scenario']}%")
    if state.get("style"):
        clauses.append("styles LIKE ?")
        params.append(f"%{state['style']}%")
    where = " AND ".join(clauses)
    order = "CASE WHEN slim_tag = '显瘦' THEN 0 ELSE 1 END, price ASC" if state.get("slim") else "price ASC"
    with db_connection() as db:
        records = db.execute(f"SELECT * FROM catalog_items WHERE {where} ORDER BY {order}", params).fetchall()
    return [dict(record) for record in records]


def choose_first(items: list[dict[str, Any]], category: str, used: set[str]) -> dict[str, Any] | None:
    return next((item for item in items if item["category"] == category and item["product_id"] not in used), None)


def build_outfits(items: list[dict[str, Any]], state: dict[str, Any]) -> list[dict[str, Any]]:
    budget = state.get("budget") or 9999
    outfits: list[dict[str, Any]] = []
    patterns = [
        ["上装", "裤装", "鞋履", "配饰"],
        ["连衣裙", "鞋履", "配饰"],
        ["上装", "半身裙", "鞋履", "配饰"],
    ]
    used: set[str] = set()
    for pattern in patterns:
        if any(category in state.get("excluded_categories", []) for category in pattern):
            continue
        picks: list[dict[str, Any]] = []
        local_used = set(used)
        for category in pattern:
            candidate = choose_first(items, category, local_used)
            if not candidate:
                picks = []
                break
            picks.append(candidate)
            local_used.add(candidate["product_id"])
        total = sum(item["price"] for item in picks)
        if picks and total <= budget:
            used.update(local_used)
            outfits.append({
                "name": f"{state.get('scenario') or '日常'} {state.get('style') or '轻松'}搭配 {len(outfits) + 1}",
                "items": [{key: product[key] for key in ["product_id", "name", "category", "price", "image_url", "fit", "slim_tag", "size_range"]} for product in picks],
                "total_price": total,
                "reason": "优先满足当前场景、预算、尺码和已选择的风格条件。" + (" 已优先选择显瘦版型。" if state.get("slim") else ""),
                "size_advice": f"建议优先试穿 {state.get('size') or '常穿'} 码；具体以商品尺码表为准。",
            })
        if len(outfits) == 3:
            break
    return outfits


def state_summary(state: dict[str, Any]) -> dict[str, Any]:
    return {
        "场景": state.get("scenario"), "预算": f"¥{state['budget']}" if state.get("budget") else None,
        "风格": state.get("style"), "身高": f"{state['height']}cm" if state.get("height") else None,
        "尺码": state.get("size"), "显瘦优先": "是" if state.get("slim") else "否",
        "排除品类": state.get("excluded_categories"), "已追问轮次": state.get("question_count"),
    }


@app.on_event("startup")
def startup() -> None:
    init_db()
    with db_connection() as db:
        count = db.execute("SELECT COUNT(*) AS count FROM catalog_items").fetchone()["count"]
    if count == 0 and CATALOG_XLSX_PATH.exists():
        import_catalog(CATALOG_XLSX_PATH)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/ready")
def ready() -> dict[str, Any]:
    with db_connection() as db:
        count = db.execute("SELECT COUNT(*) AS count FROM catalog_items").fetchone()["count"]
    return {"status": "ok" if count == 70 else "degraded", "catalog_items": count, "catalog_source": str(CATALOG_XLSX_PATH)}


@app.post("/api/catalog/import-xlsx")
def refresh_catalog(request: ImportRequest) -> dict[str, Any]:
    source = Path(request.source_path) if request.source_path else CATALOG_XLSX_PATH
    return import_catalog(source)


@app.post("/api/chat")
def chat(request: ChatRequest) -> dict[str, Any]:
    session_id = request.session_id or str(uuid.uuid4())
    trace_id = str(uuid.uuid4())
    state = get_session(session_id)
    state, changes = parse_conditions(request.message, state)
    missing = missing_fields(state)
    response_type = "recommendation"
    tool_summary = ""
    if missing and state["question_count"] < 3:
        state["question_count"] += 1
        response_type = "follow_up"
        assistant_message = ask_question(state)
        outfits: list[dict[str, Any]] = []
        tool_summary = "未调用商品检索：仍在收集必要条件"
    else:
        products = query_products(state)
        outfits = build_outfits(products, state)
        tool_summary = f"catalog.search 返回 {len(products)} 条候选；outfit.compose 生成 {len(outfits)} 套"
        if outfits:
            assistant_message = f"已按当前条件为你生成 {len(outfits)} 套完整穿搭。你还可以继续说“更显瘦”“再便宜一点”或“不要裙子”。"
        else:
            response_type = "no_match"
            assistant_message = "当前条件下没有足够的完整搭配。可以试试提高预算、放宽风格，或告诉我“换成通勤”。"
    save_session(session_id, state)
    payload = {
        "trace_id": trace_id,
        "session_id": session_id,
        "type": response_type,
        "message": assistant_message,
        "changes": changes,
        "state": state_summary(state),
        "outfits": outfits,
    }
    with db_connection() as db:
        db.execute(
            "INSERT INTO traces (trace_id,session_id,user_message,state_json,tool_summary,response_json,created_at) VALUES (?,?,?,?,?,?,?)",
            (trace_id, session_id, request.message, json.dumps(state, ensure_ascii=False), tool_summary, json.dumps(payload, ensure_ascii=False), now()),
        )
    return payload


@app.get("/api/traces/{trace_id}")
def get_trace(trace_id: str) -> dict[str, Any]:
    with db_connection() as db:
        record = db.execute("SELECT * FROM traces WHERE trace_id = ?", (trace_id,)).fetchone()
    if not record:
        raise HTTPException(status_code=404, detail="trace 未找到")
    return dict(record)
