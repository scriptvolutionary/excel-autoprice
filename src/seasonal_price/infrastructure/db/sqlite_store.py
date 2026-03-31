from __future__ import annotations

import logging
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterator

from seasonal_price.domain.models import AllocationLine, OrderLine, StockItem
from seasonal_price.infrastructure.db.migrations import MIGRATIONS

ISO = "%Y-%m-%dT%H:%M:%S%z"


def utc_now() -> str:
    return datetime.now(timezone.utc).strftime(ISO)


def parse_datetime(value: str) -> datetime:
    """Парсит дату из БД в tolerant-режиме.

    Исторические записи могут быть как timezone-aware, так и naive.
    """

    try:
        return datetime.strptime(value, ISO)
    except ValueError:
        parsed = datetime.fromisoformat(value)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed


class SQLiteStore:
    def __init__(self, db_path: Path, logger: logging.Logger) -> None:
        self._db_path = db_path
        self._logger = logger
        self._db_path.parent.mkdir(parents=True, exist_ok=True)
        self.ensure_schema()

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        con = sqlite3.connect(self._db_path)
        con.row_factory = sqlite3.Row
        try:
            yield con
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()

    def ensure_schema(self) -> None:
        with self._connect() as con:
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA synchronous=NORMAL")
            con.execute("PRAGMA temp_store=MEMORY")
            con.execute("PRAGMA cache_size=-20000")
            con.execute("PRAGMA foreign_keys=OFF")
            table_exists = (
                con.execute(
                    "SELECT 1 FROM sqlite_master WHERE type='table' AND name='schema_version' LIMIT 1"
                ).fetchone()
                is not None
            )
            if not table_exists:
                con.execute(
                    "CREATE TABLE IF NOT EXISTS schema_version (version INTEGER PRIMARY KEY, applied_at TEXT NOT NULL)"
                )
            existing_versions = {row["version"] for row in con.execute(
                "SELECT version FROM schema_version")}
            for version, sql in MIGRATIONS:
                if version in existing_versions:
                    continue
                con.executescript(sql)
                con.execute(
                    "INSERT INTO schema_version(version, applied_at) VALUES (?, ?)",
                    (version, utc_now()),
                )
                self._logger.info("Применена миграция БД v%s", version)

    def init_season(self, season_id: str) -> None:
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO seasons(season_id, created_at, status, closed_at)
                VALUES (?, ?, 'active', NULL)
                ON CONFLICT(season_id) DO NOTHING
                """,
                (season_id, utc_now()),
            )

    def close_season(self, season_id: str) -> None:
        with self._connect() as con:
            con.execute(
                """
                UPDATE seasons
                SET status='closed', closed_at=?
                WHERE season_id=?
                """,
                (utc_now(), season_id),
            )

    def upsert_sku(
        self,
        sku: str,
        category_code: str,
        category_name: str,
        sort_name: str,
        container: str,
        size: str,
    ) -> None:
        self.upsert_skus(
            [
                (
                    sku,
                    category_code,
                    category_name,
                    sort_name,
                    container,
                    size,
                )
            ]
        )

    def upsert_skus(
        self,
        entries: list[tuple[str, str, str, str, str, str]],
    ) -> None:
        if not entries:
            return
        created_at = utc_now()
        rows = [
            (
                sku,
                category_code,
                category_name,
                sort_name,
                container,
                size,
                created_at,
            )
            for sku, category_code, category_name, sort_name, container, size in entries
        ]
        with self._connect() as con:
            con.executemany(
                """
                INSERT INTO sku_registry(
                    sku, category_code, category_name, sort_name, container, size, created_at, is_active
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 1)
                ON CONFLICT(category_name, sort_name, container, size)
                DO UPDATE SET
                    sku=excluded.sku,
                    category_code=excluded.category_code,
                    is_active=1
                """,
                rows,
            )

    def find_sku_by_key(self, category_name: str, sort_name: str, container: str, size: str) -> str | None:
        with self._connect() as con:
            row = con.execute(
                """
                SELECT sku
                FROM sku_registry
                WHERE category_name=? AND sort_name=? AND container=? AND size=?
                LIMIT 1
                """,
                (category_name, sort_name, container, size),
            ).fetchone()
            return None if row is None else str(row["sku"])

    def find_skus_by_keys(
        self, keys: list[tuple[str, str, str, str]]
    ) -> dict[tuple[str, str, str, str], str]:
        if not keys:
            return {}

        out: dict[tuple[str, str, str, str], str] = {}
        chunk_size = 200  # 200 * 4 params = 800 < sqlite default 999
        with self._connect() as con:
            for start in range(0, len(keys), chunk_size):
                chunk = keys[start: start + chunk_size]
                placeholders = ",".join(["(?, ?, ?, ?)"] * len(chunk))
                params: list[str] = []
                for category_name, sort_name, container, size in chunk:
                    params.extend([category_name, sort_name, container, size])
                rows = con.execute(
                    f"""
                    SELECT sku, category_name, sort_name, container, size
                    FROM sku_registry
                    WHERE (category_name, sort_name, container, size) IN ({placeholders})
                    """,
                    params,
                ).fetchall()
                for row in rows:
                    key = (
                        str(row["category_name"]),
                        str(row["sort_name"]),
                        str(row["container"]),
                        str(row["size"]),
                    )
                    out[key] = str(row["sku"])
        return out

    def all_skus(self) -> set[str]:
        with self._connect() as con:
            return {str(row["sku"]) for row in con.execute("SELECT sku FROM sku_registry")}

    def replace_stock_items(self, season_id: str, stock_items: list[StockItem], source_file: Path) -> None:
        imported_at = utc_now()
        rows = [
            (
                season_id,
                item.sku,
                item.category_name,
                item.category_code,
                item.sort_name,
                item.container,
                item.size,
                item.price,
                item.multiplicity,
                item.stock_total,
                str(source_file),
                imported_at,
            )
            for item in stock_items
        ]
        with self._connect() as con:
            con.execute(
                "DELETE FROM stock_items WHERE season_id=?", (season_id,))
            con.executemany(
                """
                INSERT INTO stock_items(
                    season_id, sku, category_name, category_code, sort_name, container, size,
                    price, multiplicity, stock_total, source_file, imported_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )

    def list_stock_items(self, season_id: str) -> list[StockItem]:
        with self._connect() as con:
            rows = con.execute(
                """
                SELECT sku, category_name, category_code, sort_name, container, size, price, multiplicity, stock_total
                FROM stock_items
                WHERE season_id=?
                ORDER BY category_name, sort_name
                """,
                (season_id,),
            ).fetchall()
        return [
            StockItem(
                sku=str(row["sku"]),
                category_name=str(row["category_name"]),
                category_code=str(row["category_code"]),
                sort_name=str(row["sort_name"]),
                container=str(row["container"]),
                size=str(row["size"]),
                price=float(row["price"]),
                multiplicity=int(row["multiplicity"]),
                stock_total=int(row["stock_total"]),
            )
            for row in rows
        ]

    def start_import_run(
        self,
        run_id: str,
        season_id: str,
        profile_id: str,
        input_dir: Path,
        duplicate_strategy: str,
        total_files: int,
    ) -> None:
        with self._connect() as con:
            con.execute(
                """
                INSERT INTO import_runs(
                    run_id, season_id, profile_id, started_at, input_dir, duplicate_strategy,
                    total_files, success_files, error_files
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, 0, 0)
                """,
                (
                    run_id,
                    season_id,
                    profile_id,
                    utc_now(),
                    str(input_dir),
                    duplicate_strategy,
                    total_files,
                ),
            )

    def finish_import_run(self, run_id: str, success_files: int, error_files: int) -> None:
        with self._connect() as con:
            con.execute(
                """
                UPDATE import_runs
                SET finished_at=?, success_files=?, error_files=?
                WHERE run_id=?
                """,
                (utc_now(), success_files, error_files, run_id),
            )

    def insert_import_file_log(
        self,
        run_id: str,
        file_path: Path,
        client_name: str,
        mtime: datetime,
        status: str,
        message: str | None,
    ) -> None:
        self.insert_import_file_logs(
            [
                (
                    run_id,
                    file_path,
                    client_name,
                    mtime,
                    status,
                    message,
                )
            ]
        )

    def insert_import_file_logs(
        self,
        logs: list[tuple[str, Path, str, datetime, str, str | None]],
    ) -> None:
        if not logs:
            return
        rows = [
            (
                run_id,
                str(file_path),
                client_name,
                mtime.strftime(ISO),
                status,
                message,
            )
            for run_id, file_path, client_name, mtime, status, message in logs
        ]
        with self._connect() as con:
            con.executemany(
                """
                INSERT INTO import_files(run_id, file_path, client_name, mtime, status, message)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                rows,
            )

    def replace_order_lines(
        self, run_id: str, season_id: str, profile_id: str, order_lines: list[OrderLine]
    ) -> None:
        rows = [
            (
                run_id,
                line.season_id,
                line.profile_id,
                line.client_name,
                str(line.source_file),
                line.order_mtime.strftime(ISO),
                line.sku,
                line.requested_qty,
                line.rounded_qty,
                int(line.was_rounded),
                line.unit_price,
            )
            for line in order_lines
        ]
        with self._connect() as con:
            con.execute(
                "DELETE FROM order_lines WHERE season_id=? AND profile_id=?", (season_id, profile_id))
            con.executemany(
                """
                INSERT INTO order_lines(
                    run_id, season_id, profile_id, client_name, source_file, order_mtime,
                    sku, requested_qty, rounded_qty, was_rounded, unit_price
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )

    def list_order_lines(self, season_id: str, profile_id: str) -> list[OrderLine]:
        with self._connect() as con:
            rows = con.execute(
                """
                SELECT season_id, profile_id, client_name, source_file, order_mtime, sku,
                       requested_qty, rounded_qty, was_rounded, unit_price
                FROM order_lines
                WHERE season_id=? AND profile_id=?
                """,
                (season_id, profile_id),
            ).fetchall()
        out: list[OrderLine] = []
        for row in rows:
            out.append(
                OrderLine(
                    season_id=str(row["season_id"]),
                    profile_id=str(row["profile_id"]),
                    client_name=str(row["client_name"]),
                    source_file=Path(str(row["source_file"])),
                    order_mtime=parse_datetime(str(row["order_mtime"])),
                    sku=str(row["sku"]),
                    requested_qty=int(row["requested_qty"]),
                    rounded_qty=int(row["rounded_qty"]),
                    was_rounded=bool(row["was_rounded"]),
                    unit_price=float(row["unit_price"]),
                )
            )
        return out

    def replace_allocation_lines(
        self, season_id: str, profile_id: str, allocation_lines: list[AllocationLine]
    ) -> None:
        allocated_at = utc_now()
        rows = [
            (
                line.season_id,
                line.profile_id,
                line.client_name,
                str(line.source_file),
                line.order_mtime.strftime(ISO),
                line.sku,
                line.requested_qty,
                line.confirmed_qty,
                int(line.was_rounded),
                line.allocation_mode,
                allocated_at,
            )
            for line in allocation_lines
        ]
        with self._connect() as con:
            con.execute(
                "DELETE FROM allocation_lines WHERE season_id=? AND profile_id=?",
                (season_id, profile_id),
            )
            con.executemany(
                """
                INSERT INTO allocation_lines(
                    season_id, profile_id, client_name, source_file, order_mtime, sku,
                    requested_qty, confirmed_qty, was_rounded, allocation_mode, allocated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                rows,
            )

    def list_allocation_lines(self, season_id: str, profile_id: str) -> list[AllocationLine]:
        with self._connect() as con:
            rows = con.execute(
                """
                SELECT season_id, profile_id, client_name, source_file, order_mtime, sku,
                       requested_qty, confirmed_qty, was_rounded, allocation_mode
                FROM allocation_lines
                WHERE season_id=? AND profile_id=?
                """,
                (season_id, profile_id),
            ).fetchall()
        out: list[AllocationLine] = []
        for row in rows:
            out.append(
                AllocationLine(
                    season_id=str(row["season_id"]),
                    profile_id=str(row["profile_id"]),
                    client_name=str(row["client_name"]),
                    source_file=Path(str(row["source_file"])),
                    order_mtime=parse_datetime(str(row["order_mtime"])),
                    sku=str(row["sku"]),
                    requested_qty=int(row["requested_qty"]),
                    confirmed_qty=int(row["confirmed_qty"]),
                    was_rounded=bool(row["was_rounded"]),
                    allocation_mode=str(row["allocation_mode"]),
                )
            )
        return out

    def list_profile_ids_with_allocations(self, season_id: str) -> list[str]:
        with self._connect() as con:
            rows = con.execute(
                """
                SELECT DISTINCT profile_id
                FROM allocation_lines
                WHERE season_id=?
                ORDER BY profile_id
                """,
                (season_id,),
            ).fetchall()
        return [str(row["profile_id"]) for row in rows]

    def confirmed_by_sku(self, season_id: str, profile_id: str | None) -> dict[str, int]:
        with self._connect() as con:
            if profile_id is None:
                rows = con.execute(
                    """
                    SELECT sku, SUM(confirmed_qty) AS total
                    FROM allocation_lines
                    WHERE season_id=?
                    GROUP BY sku
                    """,
                    (season_id,),
                ).fetchall()
            else:
                rows = con.execute(
                    """
                    SELECT sku, SUM(confirmed_qty) AS total
                    FROM allocation_lines
                    WHERE season_id=? AND profile_id=?
                    GROUP BY sku
                    """,
                    (season_id, profile_id),
                ).fetchall()
        return {str(row["sku"]): int(row["total"]) for row in rows}
