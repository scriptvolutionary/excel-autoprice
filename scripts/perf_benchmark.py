from __future__ import annotations

from datetime import datetime
from pathlib import Path
import time

from openpyxl import Workbook

from seasonal_price.application.api import allocate, generate_price, import_orders, init_season


def create_stock(path: Path, positions: int = 300) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Склад"
    ws.append(["Категория", "Сорт", "Контейнер", "Размер", "Цена", "Кратность", "Остаток"])
    for idx in range(1, positions + 1):
        ws.append(
            [
                "Гортензии",
                f"Сорт {idx}",
                "40 шт",
                "2-3 листа",
                40 + (idx % 10),
                24,
                500,
            ]
        )
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def create_orders(dir_path: Path, files_count: int = 300, rows_per_file: int = 40) -> None:
    dir_path.mkdir(parents=True, exist_ok=True)
    for file_idx in range(1, files_count + 1):
        wb = Workbook()
        ws = wb.active
        ws.title = "Заказ"
        ws.append(["Фирма заказчика", f"Клиент {file_idx:03d}"])
        ws.append([])
        ws.append(["SKU", "Культура/сорт", "Контейнер", "Размер", "Заказ"])
        for row_idx in range(1, rows_per_file + 1):
            # SKU создаются после generate_price, поэтому для нагрузочного прогона
            # используется legacy-сопоставление.
            ws.append(["", f"Сорт {row_idx}", "40 шт", "2-3 листа", 30])
        wb.save(dir_path / f"order_{file_idx:03d}.xlsx")


def main() -> None:
    base = Path.cwd() / ".benchmark"
    stock_file = base / "stock.xlsx"
    orders_dir = base / "orders"
    out_dir = base / "out"
    season_id = "benchmark_season"
    profile_id = "default"

    create_stock(stock_file)
    create_orders(orders_dir)

    started = time.perf_counter()
    init_season(season_id=season_id, base_dir=base)
    generate_price(stock_file=stock_file, season_id=season_id, output_dir=out_dir, base_dir=base)
    import_orders(
        input_dir=orders_dir,
        season_id=season_id,
        profile_id=profile_id,
        base_dir=base,
        duplicate_strategy="latest",
    )
    allocate(season_id=season_id, mode="fifo", profile_id=profile_id, base_dir=base)
    elapsed = time.perf_counter() - started
    print(f"Benchmark completed in {elapsed:.2f}s at {datetime.now().isoformat()}")


if __name__ == "__main__":
    main()
