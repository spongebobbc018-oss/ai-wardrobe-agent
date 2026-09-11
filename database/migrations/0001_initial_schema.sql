-- 衣橱顾问核心数据结构。
-- 本文件用于审计和迁移追踪；本地开发环境由 FastAPI 启动时幂等执行同等结构。

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

CREATE TABLE IF NOT EXISTS user_profiles (
  profile_id TEXT PRIMARY KEY,
  profile_json TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS wardrobe_items (
  item_id TEXT PRIMARY KEY,
  profile_id TEXT NOT NULL,
  name TEXT NOT NULL,
  category TEXT NOT NULL,
  color TEXT,
  styles_json TEXT NOT NULL,
  scenarios_json TEXT NOT NULL,
  fit TEXT,
  size TEXT,
  ownership_status TEXT NOT NULL,
  image_url TEXT,
  idle INTEGER NOT NULL DEFAULT 0,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_wardrobe_profile ON wardrobe_items(profile_id);

CREATE TABLE IF NOT EXISTS outfit_calendar_entries (
  entry_id TEXT PRIMARY KEY,
  profile_id TEXT NOT NULL,
  wear_date TEXT NOT NULL,
  wardrobe_item_ids_json TEXT NOT NULL,
  outfit_title TEXT,
  scene TEXT,
  weather_feeling TEXT,
  comfort_rating INTEGER,
  satisfaction_rating INTEGER,
  note TEXT,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL,
  UNIQUE(profile_id, wear_date)
);

CREATE INDEX IF NOT EXISTS idx_calendar_profile_month ON outfit_calendar_entries(profile_id, wear_date);
