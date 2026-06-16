CREATE TABLE IF NOT EXISTS projects (
  project_id TEXT PRIMARY KEY,
  project_name TEXT,
  data TEXT NOT NULL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS bim_files (
  id TEXT PRIMARY KEY,
  project_id TEXT,
  data TEXT NOT NULL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS material_mappings (
  id TEXT PRIMARY KEY,
  project_id TEXT,
  bim_material_name TEXT,
  item_name TEXT,
  data TEXT NOT NULL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS material_schemes (
  id TEXT PRIMARY KEY,
  project_id TEXT,
  scheme_name TEXT,
  data TEXT NOT NULL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS analysis_results (
  analysis_id TEXT PRIMARY KEY,
  project_id TEXT,
  project_name TEXT,
  analysis_date TEXT,
  total_volume REAL DEFAULT 0,
  total_carbon REAL DEFAULT 0,
  data TEXT NOT NULL,
  created_at TEXT DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS app_state (
  id TEXT PRIMARY KEY,
  active_project_id TEXT,
  calculation TEXT,
  data TEXT,
  updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
