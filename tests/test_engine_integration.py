from __future__ import annotations

from pathlib import Path
import shutil
from unittest.mock import patch
from uuid import uuid4
import unittest

from openpyxl import Workbook, load_workbook

from seasonal_price.application.api import (
    allocate,
    build_residual_price,
    close_season,
    export_confirmations,
    generate_price,
    import_orders,
    init_season,
)
from seasonal_price.exceptions import ValidationError


class EngineIntegrationTests(unittest.TestCase):
    def test_generate_price_accepts_case_insensitive_stock_sheet_name(self) -> None:
        tmp_root = Path.cwd() / ".tmp_tests"
        tmp_root.mkdir(parents=True, exist_ok=True)
        base_dir = tmp_root / f"seasonal_price_{uuid4().hex}"
        base_dir.mkdir(parents=True, exist_ok=True)
        try:
            stock_file = base_dir / "stock_case.xlsx"
            self._create_stock_file(stock_file, sheet_name="склад")

            result = generate_price(
                stock_file=stock_file,
                season_id="spring_2026",
                output_dir=base_dir / "out",
                base_dir=base_dir,
            )
            self.assertTrue(Path(result["output_file"]).exists())
        finally:
            shutil.rmtree(base_dir, ignore_errors=True)

    def test_generate_price_reports_available_sheet_names(self) -> None:
        tmp_root = Path.cwd() / ".tmp_tests"
        tmp_root.mkdir(parents=True, exist_ok=True)
        base_dir = tmp_root / f"seasonal_price_{uuid4().hex}"
        base_dir.mkdir(parents=True, exist_ok=True)
        try:
            stock_file = base_dir / "stock_missing_sheet.xlsx"
            self._create_stock_file(stock_file, sheet_name="Ост")

            with self.assertRaises(ValidationError) as exc:
                generate_price(
                    stock_file=stock_file,
                    season_id="spring_2026",
                    output_dir=base_dir / "out",
                    base_dir=base_dir,
                )

            message = str(exc.exception)
            self.assertIn("не найден лист 'Склад'", message)
            self.assertIn("'Ост'", message)
            self.assertIn(stock_file.name, message)
        finally:
            shutil.rmtree(base_dir, ignore_errors=True)

    def test_generate_price_supports_legacy_stock_layout_without_headers(self) -> None:
        tmp_root = Path.cwd() / ".tmp_tests"
        tmp_root.mkdir(parents=True, exist_ok=True)
        base_dir = tmp_root / f"seasonal_price_{uuid4().hex}"
        base_dir.mkdir(parents=True, exist_ok=True)
        try:
            stock_file = base_dir / "stock_legacy.xlsx"
            self._create_legacy_stock_file(stock_file, sheet_name="склад")

            result = generate_price(
                stock_file=stock_file,
                season_id="spring_2026",
                output_dir=base_dir / "out",
                base_dir=base_dir,
            )
            self.assertEqual(result["items"], 2)
            self.assertTrue(Path(result["output_file"]).exists())
        finally:
            shutil.rmtree(base_dir, ignore_errors=True)

    def test_import_orders_works_when_stderr_is_unavailable(self) -> None:
        tmp_root = Path.cwd() / ".tmp_tests"
        tmp_root.mkdir(parents=True, exist_ok=True)
        base_dir = tmp_root / f"seasonal_price_{uuid4().hex}"
        base_dir.mkdir(parents=True, exist_ok=True)
        try:
            season_id = "spring_2026"
            profile_id = "default"

            stock_file = base_dir / "stock.xlsx"
            self._create_stock_file(stock_file)
            init_season(season_id=season_id, base_dir=base_dir)
            price_result = generate_price(
                stock_file=stock_file,
                season_id=season_id,
                output_dir=base_dir / "out",
                base_dir=base_dir,
            )
            sku_map = self._read_generated_skus(
                Path(price_result["output_file"]))

            orders_dir = base_dir / "orders"
            orders_dir.mkdir(parents=True, exist_ok=True)
            self._create_order_file(
                orders_dir / "order_client_1.xlsx",
                client="Клиент 1",
                sku=sku_map["Анабель"],
                sort_name="Анабель",
                qty=30,
            )

            with patch("seasonal_price.application.engine.sys.stderr", None):
                import_result = import_orders(
                    input_dir=orders_dir,
                    season_id=season_id,
                    profile_id=profile_id,
                    base_dir=base_dir,
                    duplicate_strategy="latest",
                )

            self.assertEqual(import_result["error_files"], 0)
            self.assertEqual(import_result["success_files"], 1)
        finally:
            shutil.rmtree(base_dir, ignore_errors=True)

    def test_end_to_end_flow(self) -> None:
        tmp_root = Path.cwd() / ".tmp_tests"
        tmp_root.mkdir(parents=True, exist_ok=True)
        base_dir = tmp_root / f"seasonal_price_{uuid4().hex}"
        base_dir.mkdir(parents=True, exist_ok=True)
        try:
            season_id = "spring_2026"
            profile_id = "default"

            stock_file = base_dir / "stock.xlsx"
            self._create_stock_file(stock_file)

            init_season(season_id=season_id, base_dir=base_dir)
            price_result = generate_price(
                stock_file=stock_file,
                season_id=season_id,
                output_dir=base_dir / "out",
                base_dir=base_dir,
            )
            self.assertTrue(Path(price_result["output_file"]).exists())

            sku_map = self._read_generated_skus(
                Path(price_result["output_file"]))
            orders_dir = base_dir / "orders"
            orders_dir.mkdir(parents=True, exist_ok=True)
            self._create_order_file(
                orders_dir / "order_client_1.xlsx",
                client="Клиент 1",
                sku=sku_map["Анабель"],
                sort_name="Анабель",
                qty=30,
            )
            self._create_order_file(
                orders_dir / "order_client_2.xlsx",
                client="Клиент 2",
                sku="",
                sort_name="Блю",
                qty=48,
            )

            import_result = import_orders(
                input_dir=orders_dir,
                season_id=season_id,
                profile_id=profile_id,
                base_dir=base_dir,
                duplicate_strategy="latest",
            )
            self.assertEqual(import_result["error_files"], 0)
            self.assertEqual(import_result["success_files"], 2)
            self.assertTrue(Path(import_result["summary_file"]).exists())

            alloc_result = allocate(
                season_id=season_id,
                profile_id=profile_id,
                mode="fifo",
                base_dir=base_dir,
            )
            self.assertGreater(alloc_result["allocation_lines"], 0)
            self.assertTrue(Path(alloc_result["allocation_report"]).exists())

            conf_result = export_confirmations(
                season_id=season_id,
                output_dir=base_dir / "confirmations",
                pdf_mode="builtin",
                base_dir=base_dir,
            )
            self.assertGreater(conf_result["clients"], 0)
            self.assertTrue(Path(conf_result["output_dir"]).exists())

            residual = build_residual_price(
                season_id=season_id,
                output_dir=base_dir / "out",
                base_dir=base_dir,
            )
            self.assertTrue(Path(residual["output_file"]).exists())

            close = close_season(
                season_id=season_id,
                archive_dir=base_dir / "archive",
                base_dir=base_dir,
            )
            self.assertTrue(Path(close["archive_dir"]).exists())
            self.assertTrue(str(close["new_season"]).startswith("season_"))
        finally:
            shutil.rmtree(base_dir, ignore_errors=True)

    @staticmethod
    def _create_stock_file(path: Path, sheet_name: str = "Склад") -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        # Проверяем устойчивость к перестановке колонок через заголовки.
        ws.append(["Размер", "Категория", "Сорт", "Остаток",
                  "Контейнер", "Кратность", "Цена"])
        ws.append(["2-3 листа", "Гортензия", "Анабель", 100, "40 шт", 24, 42])
        ws.append(["2-3 листа", "Гортензия", "Блю", 50, "40 шт", 24, 45])
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)

    @staticmethod
    def _create_legacy_stock_file(path: Path, sheet_name: str) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append(["", "", "Закрытая корневая система", "", "", "", ""])
        ws.append(["", "Земляника садовая", "", "", "", "", ""])
        ws.append(["", "Заказ кратен 40 (кассета 40 шт)", "", "", "", "", ""])
        ws.append(["", "1", "Анабель", "40 шт", "2-3 листа", 42, 5000])
        ws.append(["", "2", "Блю", "40 шт", "2-3 листа", 45, 1500])
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)

    @staticmethod
    def _create_order_file(path: Path, client: str, sku: str, sort_name: str, qty: int) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = "Заказ"
        ws.append(["Фирма заказчика", client])
        ws.append([])
        ws.append(["SKU", "Культура/сорт", "Контейнер", "Размер", "Заказ"])
        ws.append([sku, sort_name, "40 шт", "2-3 листа", qty])
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)

    @staticmethod
    def _read_generated_skus(price_file: Path) -> dict[str, str]:
        wb = load_workbook(price_file)
        ws = wb["Прайс"]
        out: dict[str, str] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            sku = str(row[0])
            sort_name = str(row[2])
            out[sort_name] = sku
        return out


if __name__ == "__main__":
    unittest.main()
