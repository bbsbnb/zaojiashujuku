from __future__ import annotations

import json
import sqlite3
from pathlib import Path
from typing import Any, Iterable


SCHEMA_VERSION = 1


def connect(library: Path) -> sqlite3.Connection:
    db_path = library / "database" / "cost_database.sqlite"
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_library(library: Path) -> None:
    for rel in [
        "database",
        "database/migrations",
        "projects",
        "inquiries",
        "new_projects",
        "exports/pricing_results",
        "exports/query_results",
        "exports/reports",
        "backups/database",
        "backups/full_library",
        "templates/import_templates",
        "templates/export_templates",
        "logs/import",
        "logs/pricing",
        "logs/ai",
        "logs/app",
        "temp",
        "config",
    ]:
        (library / rel).mkdir(parents=True, exist_ok=True)
    with connect(library) as conn:
        create_schema(conn)
        seed_settings(conn, library)


def create_schema(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS projects (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            short_name TEXT,
            region TEXT NOT NULL,
            city TEXT,
            district TEXT,
            year INTEGER NOT NULL,
            project_type TEXT NOT NULL,
            building_area REAL,
            structure_type TEXT,
            cost_stage TEXT NOT NULL,
            price_scope TEXT NOT NULL,
            payment_terms TEXT,
            duration TEXT,
            currency TEXT DEFAULT 'CNY',
            original_folder_path TEXT NOT NULL,
            archive_folder_path TEXT NOT NULL,
            notes TEXT,
            status TEXT DEFAULT 'active',
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS files (
            id TEXT PRIMARY KEY,
            project_id TEXT,
            original_name TEXT NOT NULL,
            original_path TEXT NOT NULL,
            archive_path TEXT NOT NULL,
            archive_relative_path TEXT NOT NULL,
            file_type TEXT NOT NULL,
            extension TEXT NOT NULL,
            file_size INTEGER,
            file_hash TEXT,
            import_status TEXT NOT NULL,
            parsed_rows INTEGER DEFAULT 0,
            error_rows INTEGER DEFAULT 0,
            parse_message TEXT,
            file_status TEXT DEFAULT 'active',
            version_no INTEGER DEFAULT 1,
            replaced_file_id TEXT,
            is_archived_only INTEGER DEFAULT 0,
            imported_at TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id)
        );

        CREATE TABLE IF NOT EXISTS standard_items (
            id TEXT PRIMARY KEY,
            item_type TEXT NOT NULL,
            standard_name TEXT NOT NULL,
            standard_spec TEXT,
            standard_unit TEXT NOT NULL,
            category TEXT,
            keywords TEXT,
            aliases TEXT,
            brand_requirement TEXT,
            notes TEXT,
            created_by TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS supplier_quotes (
            id TEXT PRIMARY KEY,
            project_id TEXT,
            file_id TEXT NOT NULL,
            supplier_name TEXT NOT NULL,
            supplier_alias TEXT,
            quote_date TEXT NOT NULL,
            valid_until TEXT,
            region TEXT,
            payment_terms TEXT,
            is_tax_included TEXT DEFAULT 'unknown',
            tax_rate REAL,
            is_freight_included TEXT DEFAULT 'unknown',
            is_installation_included TEXT DEFAULT 'unknown',
            currency TEXT DEFAULT 'CNY',
            contact_name TEXT,
            contact_phone TEXT,
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id),
            FOREIGN KEY(file_id) REFERENCES files(id)
        );

        CREATE TABLE IF NOT EXISTS quote_items (
            id TEXT PRIMARY KEY,
            quote_id TEXT NOT NULL,
            project_id TEXT,
            file_id TEXT NOT NULL,
            standard_item_id TEXT,
            item_type TEXT,
            original_name TEXT NOT NULL,
            original_spec TEXT,
            original_unit TEXT NOT NULL,
            standard_name TEXT,
            standard_spec TEXT,
            standard_unit TEXT,
            brand TEXT,
            quantity REAL,
            unit_price REAL NOT NULL,
            total_price REAL,
            quote_date TEXT NOT NULL,
            region TEXT,
            is_tax_included TEXT DEFAULT 'unknown',
            tax_rate REAL,
            is_freight_included TEXT DEFAULT 'unknown',
            is_installation_included TEXT DEFAULT 'unknown',
            payment_terms TEXT,
            source_sheet TEXT NOT NULL,
            source_row INTEGER NOT NULL,
            is_outlier INTEGER DEFAULT 0,
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(quote_id) REFERENCES supplier_quotes(id),
            FOREIGN KEY(file_id) REFERENCES files(id),
            FOREIGN KEY(standard_item_id) REFERENCES standard_items(id)
        );

        CREATE TABLE IF NOT EXISTS prices (
            id TEXT PRIMARY KEY,
            project_id TEXT,
            file_id TEXT NOT NULL,
            quote_item_id TEXT,
            standard_item_id TEXT,
            item_type TEXT NOT NULL,
            original_name TEXT NOT NULL,
            original_spec TEXT,
            original_unit TEXT NOT NULL,
            standard_name TEXT,
            standard_spec TEXT,
            standard_unit TEXT,
            brand TEXT,
            quantity REAL,
            unit_price REAL NOT NULL,
            total_price REAL,
            source_type TEXT NOT NULL,
            cost_stage TEXT,
            region TEXT NOT NULL,
            project_type TEXT,
            price_year INTEGER,
            price_date TEXT,
            price_scope TEXT,
            is_tax_included TEXT DEFAULT 'unknown',
            tax_rate REAL,
            is_freight_included TEXT DEFAULT 'unknown',
            is_installation_included TEXT DEFAULT 'unknown',
            payment_terms TEXT,
            duration TEXT,
            source_sheet TEXT NOT NULL,
            source_row INTEGER NOT NULL,
            is_outlier INTEGER DEFAULT 0,
            outlier_reason TEXT,
            confidence_note TEXT,
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(project_id) REFERENCES projects(id),
            FOREIGN KEY(file_id) REFERENCES files(id),
            FOREIGN KEY(quote_item_id) REFERENCES quote_items(id),
            FOREIGN KEY(standard_item_id) REFERENCES standard_items(id)
        );

        CREATE TABLE IF NOT EXISTS name_match_history (
            id TEXT PRIMARY KEY,
            original_name TEXT NOT NULL,
            original_spec TEXT,
            original_unit TEXT NOT NULL,
            item_type TEXT NOT NULL,
            standard_item_id TEXT NOT NULL,
            match_source TEXT NOT NULL,
            use_count INTEGER DEFAULT 1,
            last_used_at TEXT NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(standard_item_id) REFERENCES standard_items(id)
        );

        CREATE TABLE IF NOT EXISTS pricing_tasks (
            id TEXT PRIMARY KEY,
            project_name TEXT NOT NULL,
            region TEXT NOT NULL,
            project_type TEXT NOT NULL,
            quote_date TEXT NOT NULL,
            payment_terms TEXT,
            duration TEXT,
            price_scope TEXT NOT NULL,
            is_tax_included TEXT DEFAULT 'unknown',
            is_freight_included TEXT DEFAULT 'unknown',
            is_installation_included TEXT DEFAULT 'unknown',
            prefer_recent_inquiry INTEGER DEFAULT 1,
            input_file_path TEXT NOT NULL,
            archive_input_path TEXT NOT NULL,
            archive_input_relative_path TEXT NOT NULL,
            output_file_path TEXT,
            output_relative_path TEXT,
            status TEXT NOT NULL,
            summary TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS pricing_task_items (
            id TEXT PRIMARY KEY,
            task_id TEXT NOT NULL,
            source_row INTEGER NOT NULL,
            item_type TEXT,
            original_name TEXT NOT NULL,
            original_spec TEXT,
            original_unit TEXT NOT NULL,
            selected_standard_item_id TEXT,
            standard_name TEXT,
            standard_spec TEXT,
            standard_unit TEXT,
            original_unit_price REAL,
            suggested_price REAL,
            reference_low REAL,
            reference_high REAL,
            main_source_text TEXT,
            main_source_price_id TEXT,
            source_type TEXT,
            confidence TEXT,
            risk_tags TEXT,
            risk_text TEXT,
            ai_explanation TEXT,
            is_accepted INTEGER,
            confirmed_price REAL,
            status TEXT NOT NULL,
            notes TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            FOREIGN KEY(task_id) REFERENCES pricing_tasks(id),
            FOREIGN KEY(selected_standard_item_id) REFERENCES standard_items(id)
        );

        CREATE TABLE IF NOT EXISTS match_candidates (
            id TEXT PRIMARY KEY,
            task_item_id TEXT NOT NULL,
            rank_no INTEGER NOT NULL,
            source_price_id TEXT,
            source_quote_item_id TEXT,
            standard_item_id TEXT,
            candidate_name TEXT NOT NULL,
            candidate_spec TEXT,
            candidate_unit TEXT NOT NULL,
            candidate_price REAL NOT NULL,
            reference_low REAL,
            reference_high REAL,
            source_type TEXT NOT NULL,
            source_text TEXT NOT NULL,
            source_region TEXT,
            source_date_or_year TEXT,
            match_score REAL,
            confidence TEXT NOT NULL,
            match_reason TEXT,
            risk_tags TEXT,
            risk_text TEXT,
            is_selected INTEGER DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(task_item_id) REFERENCES pricing_task_items(id)
        );

        CREATE TABLE IF NOT EXISTS query_logs (
            id TEXT PRIMARY KEY,
            query_text TEXT NOT NULL,
            parsed_item_type TEXT,
            parsed_name TEXT,
            parsed_spec TEXT,
            parsed_unit TEXT,
            filter_region TEXT,
            filter_project_type TEXT,
            filter_year_range TEXT,
            suggested_price REAL,
            reference_low REAL,
            reference_high REAL,
            confidence TEXT,
            risk_tags TEXT,
            ai_explanation TEXT,
            result_summary TEXT,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY,
            value TEXT,
            value_type TEXT,
            description TEXT,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS ai_call_logs (
            id TEXT PRIMARY KEY,
            provider TEXT NOT NULL,
            base_url TEXT,
            model TEXT NOT NULL,
            task_type TEXT NOT NULL,
            related_query_id TEXT,
            related_pricing_task_id TEXT,
            related_task_item_id TEXT,
            prompt_summary TEXT,
            response_summary TEXT,
            token_usage TEXT,
            status TEXT NOT NULL,
            error_message TEXT,
            created_at TEXT NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_prices_name ON prices(standard_name, original_name);
        CREATE INDEX IF NOT EXISTS idx_prices_filter ON prices(item_type, region, price_year, source_type, is_outlier);
        CREATE INDEX IF NOT EXISTS idx_files_project ON files(project_id, file_type);
        """
    )
    conn.execute(
        "INSERT OR REPLACE INTO app_settings(key,value,value_type,description,updated_at) VALUES(?,?,?,?,datetime('now'))",
        ("schema_version", str(SCHEMA_VERSION), "number", "Database schema version"),
    )
    conn.commit()


def seed_settings(conn: sqlite3.Connection, library: Path) -> None:
    defaults = {
        "library_path": str(library.resolve()),
        "database_path": str((library / "database" / "cost_database.sqlite").resolve()),
        "export_path": str((library / "exports").resolve()),
        "backup_path": str((library / "backups" / "database").resolve()),
        "region_priority_for_shenzhen": json.dumps(["深圳", "广州", "东莞", "佛山", "广东其他", "外省"], ensure_ascii=False),
        "ai_provider": "openai_compatible",
        "ai_provider_name": "DeepSeek",
        "ai_base_url": "https://api.deepseek.com",
        "ai_model": "deepseek-chat",
        "ai_temperature": "0.2",
        "ai_timeout_seconds": "60",
        "ai_enabled": "false",
        "ai_send_raw_excel": "false",
        "ai_send_project_name": "false",
        "ai_send_supplier_name": "false",
        "ai_send_file_path": "false",
        "ai_send_price_summary": "true",
        "ai_anonymize_entities": "true",
    }
    for key, value in defaults.items():
        conn.execute(
            "INSERT OR IGNORE INTO app_settings(key,value,value_type,description,updated_at) VALUES(?,?,?,?,datetime('now'))",
            (key, value, "json" if value.startswith("[") or value.startswith("{") else "string", "default setting"),
        )
    conn.commit()


def insert_many(conn: sqlite3.Connection, table: str, rows: Iterable[dict[str, Any]]) -> int:
    rows = list(rows)
    if not rows:
        return 0
    keys = list(rows[0].keys())
    placeholders = ",".join("?" for _ in keys)
    sql = f"INSERT INTO {table} ({','.join(keys)}) VALUES ({placeholders})"
    conn.executemany(sql, [[row.get(k) for k in keys] for row in rows])
    return len(rows)

