from __future__ import annotations

from collections import defaultdict
from collections.abc import Callable, Iterable
from dataclasses import replace
from datetime import datetime
import json
import logging
from pathlib import Path
import shutil
import sys
from typing import Any
from uuid import uuid4

from tqdm import tqdm

from seasonal_price.config import AppConfig
from seasonal_price.domain.models import AllocationLine, ImportIssue, OrderLine, StockItem
from seasonal_price.domain.services.allocation import allocate_fifo, allocate_proportional
from seasonal_price.domain.services.rounding import round_to_multiplicity
from seasonal_price.domain.services.sku_registry import SkuRegistryService, normalize_category_code
from seasonal_price.exceptions import DuplicateResolutionError, ValidationError
from seasonal_price.infrastructure.db.sqlite_store import SQLiteStore
from seasonal_price.infrastructure.excel.excel_processor import ExcelProcessor, StockInputRow
from seasonal_price.infrastructure.pdf.reportlab_renderer import ReportLabRenderer


class SeasonalPriceEngine:
    """Основной orchestration-слой приложения.

    Слой связывает доменные правила и инфраструктурные адаптеры, но не содержит
    UI-специфики.
    """

    def __init__(
        self,
        config: AppConfig,
        logger: logging.Logger,
        store: SQLiteStore,
        excel_processor: ExcelProcessor,
        pdf_renderer: ReportLabRenderer,
    ) -> None:
        self._config = config
        self._logger = logger
        self._store = store
        self._excel = excel_processor
        self._pdf = pdf_renderer

    def init_season(self, season_id: str) -> dict[str, Any]:
        self._store.init_season(season_id)
        return {"season_id": season_id, "status": "initialized"}

    def generate_price(self, stock_file: Path, season_id: str, output_dir: Path) -> dict[str, Any]:
        self._store.init_season(season_id)
        stock_rows = self._excel.read_stock_sheet(
            stock_file=stock_file, sheet_name="Склад")
        stock_items = self._materialize_stock_with_sku(stock_rows)
        self._store.replace_stock_items(
            season_id=season_id, stock_items=stock_items, source_file=stock_file)

        output_path = output_dir / f"{season_id}_Прайс_для_клиентов.xlsx"
        self._excel.write_client_price(output_path, stock_items)
        self._logger.info("Сформирован клиентский прайс: %s", output_path)
        return {
            "season_id": season_id,
            "items": len(stock_items),
            "output_file": str(output_path),
        }

    def import_orders(
        self,
        input_dir: Path,
        season_id: str,
        profile_id: str,
        duplicate_strategy: str = "manual",
        duplicate_map_path: Path | None = None,
        rounding_mode: str = "down",
        progress_callback: Callable[[dict[str, Any]], None] | None = None,
    ) -> dict[str, Any]:
        stock_items = self._store.list_stock_items(season_id)
        if not stock_items:
            raise ValidationError(
                "Для сезона не загружен склад. Сначала выполните generate-price с листом 'Склад'."
            )

        stock_by_sku = {item.sku: item for item in stock_items}
        legacy_index = self._build_legacy_index(stock_items)

        files = self._excel.discover_order_files(input_dir)
        total_files = len(files)
        total_steps = max(total_files * 2, 1)
        self._emit_progress(
            progress_callback=progress_callback,
            phase="start",
            current=0,
            total=max(total_files, 1),
            percent=0,
            message=f"Найдено файлов для импорта: {total_files}",
        )
        run_id = str(uuid4())
        self._store.start_import_run(
            run_id=run_id,
            season_id=season_id,
            profile_id=profile_id,
            input_dir=input_dir,
            duplicate_strategy=duplicate_strategy,
            total_files=len(files),
        )

        parsed_files = []
        parse_issues_by_file: dict[Path, list[ImportIssue]] = {}
        scan_issues: list[ImportIssue] = []
        for scan_index, file_path in enumerate(
            self._iter_with_tqdm(
                files, desc="Сканирование файлов", unit="file"),
            start=1,
        ):
            try:
                parsed = self._excel.read_order_file(file_path)
                parsed_files.append(parsed)
                parse_issues_by_file[file_path] = list(parsed.issues)
            except Exception as exc:
                issue = ImportIssue(
                    file_path=file_path,
                    issue_code="parse_error",
                    message=f"Не удалось прочитать файл: {exc}",
                )
                scan_issues.append(issue)
            scan_percent = int((scan_index / total_steps) * 100)
            self._emit_progress(
                progress_callback=progress_callback,
                phase="scan",
                current=scan_index,
                total=max(total_files, 1),
                percent=scan_percent,
                message=f"Сканирование файлов: {scan_index}/{max(total_files, 1)}",
            )

        selected_paths = self._resolve_duplicate_files(
            parsed_files=parsed_files,
            strategy=duplicate_strategy,
            duplicate_map_path=duplicate_map_path,
        )
        self._emit_progress(
            progress_callback=progress_callback,
            phase="resolve_duplicates",
            current=len(selected_paths),
            total=max(len(parsed_files), 1),
            percent=50 if total_files > 0 else 100,
            message=(
                f"Разрешение дублей завершено: выбрано файлов {len(selected_paths)}"
                f" из {len(parsed_files)}."
            ),
        )

        order_lines: list[OrderLine] = []
        import_issues = list(scan_issues)
        for parsed in parsed_files:
            if parsed.meta.file_path in selected_paths:
                import_issues.extend(
                    parse_issues_by_file.get(parsed.meta.file_path, []))
        file_logs: list[tuple[str, Path, str, datetime, str, str | None]] = []
        success_files = 0
        error_files = 0

        for import_index, parsed in enumerate(
            self._iter_with_tqdm(
                parsed_files, desc="Импорт заказов", unit="file"),
            start=1,
        ):
            file_path = parsed.meta.file_path
            client_name = parsed.meta.client_name
            mtime = parsed.meta.mtime

            if file_path not in selected_paths:
                file_logs.append(
                    (
                        run_id,
                        file_path,
                        client_name,
                        mtime,
                        "skipped_duplicate",
                        "Файл пропущен по стратегии разрешения дублей.",
                    )
                )
                current_step = total_files + import_index
                import_percent = min(
                    int((current_step / total_steps) * 100), 99)
                self._emit_progress(
                    progress_callback=progress_callback,
                    phase="import",
                    current=import_index,
                    total=max(len(parsed_files), 1),
                    percent=import_percent,
                    message=f"Импорт заказов: {import_index}/{max(len(parsed_files), 1)}",
                )
                continue

            file_line_count = 0
            for row in parsed.rows:
                sku = row.sku.strip().upper() if row.sku else None
                if not sku:
                    sku = self._match_legacy_sku(
                        file_path=file_path,
                        row=row,
                        legacy_index=legacy_index,
                        issues=import_issues,
                    )
                    if sku is None:
                        continue
                if sku not in stock_by_sku:
                    import_issues.append(
                        ImportIssue(
                            file_path=file_path,
                            issue_code="unknown_sku",
                            message=f"SKU '{sku}' не найден в эталонном складе.",
                            row_index=row.row_index,
                            sku=sku,
                        )
                    )
                    continue

                stock = stock_by_sku[sku]
                rounded_qty, was_rounded = round_to_multiplicity(
                    row.quantity, stock.multiplicity, mode=rounding_mode)
                line = OrderLine(
                    season_id=season_id,
                    profile_id=profile_id,
                    client_name=client_name,
                    source_file=file_path,
                    order_mtime=mtime,
                    sku=sku,
                    requested_qty=row.quantity,
                    rounded_qty=rounded_qty,
                    was_rounded=was_rounded,
                    unit_price=stock.price,
                )
                order_lines.append(line)
                file_line_count += 1

            if file_line_count == 0:
                error_files += 1
                file_logs.append(
                    (
                        run_id,
                        file_path,
                        client_name,
                        mtime,
                        "error",
                        "Не найдено валидных строк заказа.",
                    )
                )
            else:
                success_files += 1
                file_logs.append(
                    (
                        run_id,
                        file_path,
                        client_name,
                        mtime,
                        "success",
                        f"Импортировано строк: {file_line_count}",
                    )
                )
            current_step = total_files + import_index
            import_percent = min(int((current_step / total_steps) * 100), 99)
            self._emit_progress(
                progress_callback=progress_callback,
                phase="import",
                current=import_index,
                total=max(len(parsed_files), 1),
                percent=import_percent,
                message=f"Импорт заказов: {import_index}/{max(len(parsed_files), 1)}",
            )

        self._store.insert_import_file_logs(file_logs)
        self._store.replace_order_lines(
            run_id=run_id,
            season_id=season_id,
            profile_id=profile_id,
            order_lines=order_lines,
        )
        self._store.finish_import_run(
            run_id=run_id, success_files=success_files, error_files=error_files)

        out_dir = self._config.output_dir / season_id / profile_id
        out_dir.mkdir(parents=True, exist_ok=True)
        report_path = out_dir / f"{run_id}_Отчет_импорта.xlsx"
        summary_path = out_dir / "Сводка_заказов.xlsx"

        self._excel.write_import_report(
            output_path=report_path,
            success_files=success_files,
            error_files=error_files,
            issues=import_issues,
        )
        self._excel.write_order_summary(
            summary_path, stock_items, order_lines, allocation_lines=None)
        self._emit_progress(
            progress_callback=progress_callback,
            phase="finalize",
            current=1,
            total=1,
            percent=100,
            message="Импорт завершен, формирование отчетов окончено.",
        )

        self._logger.info(
            "Импорт завершен. Успешно: %s, с ошибками: %s, run_id=%s",
            success_files,
            error_files,
            run_id,
        )
        return {
            "run_id": run_id,
            "season_id": season_id,
            "profile_id": profile_id,
            "success_files": success_files,
            "error_files": error_files,
            "imported_lines": len(order_lines),
            "rounding_mode": rounding_mode,
            "report_file": str(report_path),
            "summary_file": str(summary_path),
        }

    def allocate(self, season_id: str, mode: str, profile_id: str) -> dict[str, Any]:
        stock_items = self._store.list_stock_items(season_id)
        stock_by_sku = {item.sku: item for item in stock_items}
        order_lines = self._store.list_order_lines(
            season_id=season_id, profile_id=profile_id)
        if not order_lines:
            raise ValidationError("Нет импортированных заказов для аллокации.")

        if mode == "fifo":
            result = allocate_fifo(order_lines, stock_by_sku)
        elif mode == "proportional":
            result = allocate_proportional(order_lines, stock_by_sku)
        else:
            raise ValidationError(f"Неизвестный режим аллокации: {mode}")

        self._store.replace_allocation_lines(
            season_id=season_id,
            profile_id=profile_id,
            allocation_lines=result.lines,
        )

        out_dir = self._config.output_dir / season_id / profile_id
        out_dir.mkdir(parents=True, exist_ok=True)
        summary_path = out_dir / "Сводка_заказов.xlsx"
        report_path = out_dir / f"{mode}_Отчет_аллокации.xlsx"
        self._excel.write_order_summary(
            summary_path, stock_items, order_lines, allocation_lines=result.lines)
        self._excel.write_allocation_report(report_path, result.lines)

        return {
            "season_id": season_id,
            "profile_id": profile_id,
            "mode": mode,
            "allocation_lines": len(result.lines),
            "rounded_lines": result.corrected_count,
            "summary_file": str(summary_path),
            "allocation_report": str(report_path),
        }

    def export_confirmations(
        self,
        season_id: str,
        output_dir: Path,
        pdf_mode: str = "builtin",
    ) -> dict[str, Any]:
        profile_ids = self._store.list_profile_ids_with_allocations(season_id)
        if not profile_ids:
            raise ValidationError(
                "Для сезона не найдено профилей с импортированными данными.")

        stock_items = self._store.list_stock_items(season_id)
        stock_by_sku = {item.sku: item for item in stock_items}

        date_folder = datetime.now().strftime("%Y-%m-%d") + "_Подтверждения"
        target_dir = output_dir / date_folder
        target_dir.mkdir(parents=True, exist_ok=True)

        generated_total = 0
        clients_by_profile: dict[str, int] = {}

        for profile_id in profile_ids:
            allocation_lines = self._store.list_allocation_lines(
                season_id=season_id, profile_id=profile_id)
            clients = self._build_confirmation_lines(
                allocation_lines=allocation_lines,
                stock_by_sku=stock_by_sku,
            )

            profile_target_dir = target_dir / \
                profile_id if len(profile_ids) > 1 else target_dir
            profile_target_dir.mkdir(parents=True, exist_ok=True)
            generated_profile = 0

            for client, pairs in clients.items():
                safe_name = self._safe_file_name(client)
                xlsx_path = profile_target_dir / f"{safe_name}.xlsx"
                pdf_path = profile_target_dir / f"{safe_name}.pdf"
                self._excel.write_client_confirmation_xlsx(
                    xlsx_path, client, pairs)
                self._pdf.render_confirmation_pdf(
                    pdf_path, client, pairs, mode=pdf_mode)
                generated_profile += 1

            clients_by_profile[profile_id] = generated_profile
            generated_total += generated_profile

        return {
            "season_id": season_id,
            "profile_id": profile_ids[0] if len(profile_ids) == 1 else None,
            "profiles": profile_ids,
            "clients": generated_total,
            "clients_by_profile": clients_by_profile,
            "output_dir": str(target_dir),
        }

    def build_residual_price(self, season_id: str, output_dir: Path) -> dict[str, Any]:
        stock_items = self._store.list_stock_items(season_id)
        confirmed = self._store.confirmed_by_sku(
            season_id=season_id, profile_id=None)

        residual_items: list[StockItem] = []
        for item in stock_items:
            free = max(item.stock_total - confirmed.get(item.sku, 0), 0)
            if free == 0:
                continue
            residual_items.append(replace(item, stock_total=free))

        output_path = output_dir / f"{season_id}_Остаточный_прайс.xlsx"
        self._excel.write_residual_price(output_path, residual_items)
        return {
            "season_id": season_id,
            "positions": len(residual_items),
            "output_file": str(output_path),
        }

    def close_season(self, season_id: str, archive_dir: Path) -> dict[str, Any]:
        self._store.close_season(season_id)
        archive_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target_dir = archive_dir / f"{season_id}_{timestamp}"

        closed_at = datetime.now().isoformat(timespec="seconds")
        copied_outputs = False
        copied_logs = False
        copied_db_files: list[str] = []

        output_source_dir = self._config.output_dir / season_id
        output_target_dir = target_dir / "outputs" / season_id
        if output_source_dir.exists():
            output_target_dir.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(output_source_dir,
                            output_target_dir, dirs_exist_ok=True)
            copied_outputs = True

        if self._config.log_dir.exists():
            logs_target_dir = target_dir / "logs"
            shutil.copytree(self._config.log_dir,
                            logs_target_dir, dirs_exist_ok=True)
            copied_logs = True

        db_target_dir = target_dir / "data"
        db_target_dir.mkdir(parents=True, exist_ok=True)
        db_candidates = [
            self._config.db_path,
            Path(f"{self._config.db_path}-wal"),
            Path(f"{self._config.db_path}-shm"),
        ]
        for db_file in db_candidates:
            if not db_file.exists():
                continue
            shutil.copy2(db_file, db_target_dir / db_file.name)
            copied_db_files.append(db_file.name)

        archive_manifest = {
            "season_id": season_id,
            "closed_at": closed_at,
            "outputs_copied": copied_outputs,
            "logs_copied": copied_logs,
            "db_files": copied_db_files,
        }
        (target_dir / "manifest.json").write_text(
            json.dumps(archive_manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        new_season_id = f"season_{timestamp}"
        self._store.init_season(new_season_id)
        return {
            "closed_season": season_id,
            "archive_dir": str(target_dir),
            "new_season": new_season_id,
            "manifest_file": str(target_dir / "manifest.json"),
        }

    def generate_price(self, stock_file: Path, season_id: str, output_dir: Path) -> dict[str, Any]:
        self._store.init_season(season_id)
        stock_rows = self._excel.read_stock_sheet(stock_file=stock_file, sheet_name="Склад")
        stock_items = self._materialize_stock_with_sku(stock_rows)
        self._store.replace_stock_items(
            season_id=season_id,
            stock_items=stock_items,
            source_file=stock_file,
        )

        output_path = output_dir / f"{season_id}_Прайс_для_клиентов.xlsx"
        pdf_path = output_dir / f"{season_id}_Прайс_для_клиентов.pdf"
        self._excel.write_client_price(output_path, stock_items)
        self._pdf.render_price_pdf(pdf_path, season_id, stock_items)
        self._logger.info("Сформирован клиентский прайс: %s", output_path)
        return {
            "season_id": season_id,
            "items": len(stock_items),
            "output_file": str(output_path),
            "pdf_file": str(pdf_path),
        }

    def _materialize_stock_with_sku(self, rows: list[StockInputRow]) -> list[StockItem]:
        existing_skus = self._store.all_skus()
        registry = SkuRegistryService(existing_skus=existing_skus)

        keys = [
            (row.category_name, row.sort_name, row.container, row.size)
            for row in rows
        ]
        existing_by_key = self._store.find_skus_by_keys(keys)

        result: list[StockItem] = []
        sku_registry_rows: list[tuple[str, str, str, str, str, str]] = []
        for row in rows:
            category_code = normalize_category_code(row.category_name)
            key = (row.category_name, row.sort_name, row.container, row.size)
            sku = existing_by_key.get(key)
            if sku is None:
                sku = registry.next_sku(category_code)
                existing_by_key[key] = sku

            sku_registry_rows.append(
                (
                    sku,
                    category_code,
                    row.category_name,
                    row.sort_name,
                    row.container,
                    row.size,
                )
            )
            result.append(
                StockItem(
                    sku=sku,
                    category_name=row.category_name,
                    category_code=category_code,
                    sort_name=row.sort_name,
                    container=row.container,
                    size=row.size,
                    price=row.price,
                    multiplicity=row.multiplicity,
                    stock_total=row.stock_total,
                )
            )

        self._store.upsert_skus(sku_registry_rows)
        return result

    @staticmethod
    def _build_legacy_index(stock_items: list[StockItem]) -> dict[tuple[str, str, str], list[str]]:
        legacy: dict[tuple[str, str, str], list[str]] = defaultdict(list)
        for item in stock_items:
            legacy[item.legacy_key].append(item.sku)
        return legacy

    def _match_legacy_sku(
        self,
        file_path: Path,
        row: Any,
        legacy_index: dict[tuple[str, str, str], list[str]],
        issues: list[ImportIssue],
    ) -> str | None:
        if not row.sort_name or not row.container or not row.size:
            issues.append(
                ImportIssue(
                    file_path=file_path,
                    issue_code="legacy_missing_columns",
                    message="Не удалось сопоставить строку без SKU: нет сорт/контейнер/размер.",
                    row_index=row.row_index,
                )
            )
            return None
        key = (
            row.sort_name.strip().casefold(),
            row.container.strip().casefold(),
            row.size.strip().casefold(),
        )
        candidates = legacy_index.get(key, [])
        if not candidates:
            issues.append(
                ImportIssue(
                    file_path=file_path,
                    issue_code="legacy_not_found",
                    message="Legacy-сопоставление не найдено в эталонном складе.",
                    row_index=row.row_index,
                )
            )
            return None
        if len(candidates) > 1:
            issues.append(
                ImportIssue(
                    file_path=file_path,
                    issue_code="legacy_ambiguous",
                    message="Legacy-сопоставление неоднозначно: найдено несколько SKU.",
                    row_index=row.row_index,
                )
            )
            return None
        self._logger.warning(
            "Использовано legacy-сопоставление без SKU: файл=%s, строка=%s, SKU=%s",
            file_path.name,
            row.row_index,
            candidates[0],
        )
        return candidates[0]

    def _resolve_duplicate_files(
        self,
        parsed_files: list[Any],
        strategy: str,
        duplicate_map_path: Path | None,
    ) -> set[Path]:
        by_client: dict[str, list[Any]] = defaultdict(list)
        for parsed in parsed_files:
            by_client[parsed.meta.client_name].append(parsed)

        selected: set[Path] = set()
        unresolved: dict[str, list[str]] = {}
        duplicate_map: dict[str, str] = {}
        if duplicate_map_path and duplicate_map_path.exists():
            duplicate_map = json.loads(
                duplicate_map_path.read_text(encoding="utf-8"))

        for client, files in by_client.items():
            if len(files) == 1:
                selected.add(files[0].meta.file_path)
                continue

            if strategy == "latest":
                chosen = max(files, key=lambda item: item.meta.mtime)
                selected.add(chosen.meta.file_path)
                continue

            if strategy == "sum":
                for file in files:
                    selected.add(file.meta.file_path)
                continue

            if strategy != "manual":
                raise DuplicateResolutionError(
                    f"Неизвестная стратегия дублей: {strategy}")

            if client not in duplicate_map:
                unresolved[client] = [str(file.meta.file_path)
                                      for file in files]
                continue

            mapped = Path(duplicate_map[client]).resolve()
            matched = [
                item for item in files if item.meta.file_path.resolve() == mapped]
            if not matched:
                unresolved[client] = [str(file.meta.file_path)
                                      for file in files]
                continue
            selected.add(matched[0].meta.file_path)

        if unresolved:
            message = (
                "Обнаружены дубли файлов клиентов. Передайте JSON-карту через --duplicate-map. "
                f"Конфликты: {unresolved}"
            )
            raise DuplicateResolutionError(message)
        return selected

    def _emit_progress(
        self,
        progress_callback: Callable[[dict[str, Any]], None] | None,
        phase: str,
        current: int,
        total: int,
        percent: int,
        message: str,
    ) -> None:
        if progress_callback is None:
            return
        payload = {
            "phase": phase,
            "current": current,
            "total": total,
            "percent": max(0, min(percent, 100)),
            "message": message,
        }
        try:
            progress_callback(payload)
        except Exception as exc:
            self._logger.debug("Игнорируем ошибку progress_callback: %s", exc)

    def _iter_with_tqdm(self, items: Iterable[Any], desc: str, unit: str) -> Iterable[Any]:
        if not self._has_writable_stream(getattr(sys, "stderr", None)):
            return items
        try:
            return tqdm(items, desc=desc, unit=unit)
        except Exception as exc:
            self._logger.debug("Не удалось инициализировать tqdm: %s", exc)
            return items

    @staticmethod
    def _has_writable_stream(stream: Any) -> bool:
        return stream is not None and hasattr(stream, "write")

    def _build_confirmation_lines(
        self,
        allocation_lines: list[AllocationLine],
        stock_by_sku: dict[str, StockItem],
    ) -> dict[str, list[tuple[StockItem, AllocationLine]]]:
        clients: dict[str, dict[str, AllocationLine]] = defaultdict(dict)
        for line in allocation_lines:
            if line.sku not in stock_by_sku:
                continue
            if line.requested_qty <= 0 and line.rounded_qty <= 0 and line.confirmed_qty <= 0:
                continue

            by_sku = clients[line.client_name]
            existing = by_sku.get(line.sku)
            if existing is None:
                by_sku[line.sku] = line
                continue

            if line.order_mtime < existing.order_mtime:
                source_file = line.source_file
                order_mtime = line.order_mtime
            else:
                source_file = existing.source_file
                order_mtime = existing.order_mtime

            by_sku[line.sku] = AllocationLine(
                season_id=existing.season_id,
                profile_id=existing.profile_id,
                client_name=existing.client_name,
                source_file=source_file,
                order_mtime=order_mtime,
                sku=existing.sku,
                requested_qty=existing.requested_qty + line.requested_qty,
                rounded_qty=existing.rounded_qty + line.rounded_qty,
                confirmed_qty=existing.confirmed_qty + line.confirmed_qty,
                was_rounded=existing.was_rounded or line.was_rounded,
                allocation_mode=existing.allocation_mode,
            )

        result: dict[str, list[tuple[StockItem, AllocationLine]]] = {}
        for client_name, by_sku in clients.items():
            lines = sorted(
                by_sku.values(),
                key=lambda line: (
                    stock_by_sku[line.sku].category_name,
                    stock_by_sku[line.sku].sort_name,
                    stock_by_sku[line.sku].container,
                    stock_by_sku[line.sku].size,
                    line.sku,
                ),
            )
            result[client_name] = [(stock_by_sku[line.sku], line) for line in lines]
        return result

    @staticmethod
    def _safe_file_name(name: str) -> str:
        allowed = []
        for ch in name:
            if ch.isalnum() or ch in (" ", "-", "_"):
                allowed.append(ch)
        return "".join(allowed).strip() or "client"
