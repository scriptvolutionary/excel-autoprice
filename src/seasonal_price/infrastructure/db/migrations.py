from __future__ import annotations

MIGRATIONS: list[tuple[int, str]] = [
    (
        1,
        """
        CREATE TABLE IF NOT EXISTS schema_version (
            version INTEGER PRIMARY KEY,
            applied_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS seasons (
            season_id TEXT PRIMARY KEY,
            created_at TEXT NOT NULL,
            status TEXT NOT NULL,
            closed_at TEXT
        );

        CREATE TABLE IF NOT EXISTS sku_registry (
            sku TEXT PRIMARY KEY,
            category_code TEXT NOT NULL,
            category_name TEXT NOT NULL,
            sort_name TEXT NOT NULL,
            container TEXT NOT NULL,
            size TEXT NOT NULL,
            created_at TEXT NOT NULL,
            is_active INTEGER NOT NULL DEFAULT 1,
            UNIQUE(category_name, sort_name, container, size)
        );

        CREATE TABLE IF NOT EXISTS stock_items (
            season_id TEXT NOT NULL,
            sku TEXT NOT NULL,
            category_name TEXT NOT NULL,
            category_code TEXT NOT NULL,
            sort_name TEXT NOT NULL,
            container TEXT NOT NULL,
            size TEXT NOT NULL,
            price REAL NOT NULL,
            multiplicity INTEGER NOT NULL,
            stock_total INTEGER NOT NULL,
            source_file TEXT NOT NULL,
            imported_at TEXT NOT NULL,
            PRIMARY KEY(season_id, sku)
        );

        CREATE TABLE IF NOT EXISTS import_runs (
            run_id TEXT PRIMARY KEY,
            season_id TEXT NOT NULL,
            profile_id TEXT NOT NULL,
            started_at TEXT NOT NULL,
            finished_at TEXT,
            input_dir TEXT NOT NULL,
            duplicate_strategy TEXT NOT NULL,
            total_files INTEGER NOT NULL,
            success_files INTEGER NOT NULL,
            error_files INTEGER NOT NULL
        );

        CREATE TABLE IF NOT EXISTS import_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            file_path TEXT NOT NULL,
            client_name TEXT NOT NULL,
            mtime TEXT NOT NULL,
            status TEXT NOT NULL,
            message TEXT
        );

        CREATE TABLE IF NOT EXISTS order_lines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            season_id TEXT NOT NULL,
            profile_id TEXT NOT NULL,
            client_name TEXT NOT NULL,
            source_file TEXT NOT NULL,
            order_mtime TEXT NOT NULL,
            sku TEXT NOT NULL,
            requested_qty INTEGER NOT NULL,
            rounded_qty INTEGER NOT NULL,
            was_rounded INTEGER NOT NULL,
            unit_price REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS allocation_lines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            season_id TEXT NOT NULL,
            profile_id TEXT NOT NULL,
            client_name TEXT NOT NULL,
            source_file TEXT NOT NULL,
            order_mtime TEXT NOT NULL,
            sku TEXT NOT NULL,
            requested_qty INTEGER NOT NULL,
            confirmed_qty INTEGER NOT NULL,
            was_rounded INTEGER NOT NULL,
            allocation_mode TEXT NOT NULL,
            allocated_at TEXT NOT NULL
        );
        """,
    )
]
