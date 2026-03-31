from __future__ import annotations

from datetime import datetime
from pathlib import Path
import unittest

from seasonal_price.domain.models import OrderLine, StockItem
from seasonal_price.domain.services.allocation import allocate_fifo, allocate_proportional


class AllocationServiceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.stock = StockItem(
            sku="ROS-000001",
            category_name="Розы",
            category_code="ROS",
            sort_name="Флорибунда",
            container="P9",
            size="20-30",
            price=100.0,
            multiplicity=24,
            stock_total=100,
        )

    def test_fifo_allocation_respects_order_and_multiplicity(self) -> None:
        lines = [
            OrderLine(
                season_id="s1",
                profile_id="p1",
                client_name="Клиент 1",
                source_file=Path("c1.xlsx"),
                order_mtime=datetime(2026, 3, 10, 10, 0, 0),
                sku="ROS-000001",
                requested_qty=50,
                rounded_qty=48,
                was_rounded=True,
                unit_price=100.0,
            ),
            OrderLine(
                season_id="s1",
                profile_id="p1",
                client_name="Клиент 2",
                source_file=Path("c2.xlsx"),
                order_mtime=datetime(2026, 3, 10, 11, 0, 0),
                sku="ROS-000001",
                requested_qty=70,
                rounded_qty=48,
                was_rounded=True,
                unit_price=100.0,
            ),
        ]
        result = allocate_fifo(lines, {"ROS-000001": self.stock})
        by_client = {
            line.client_name: line.confirmed_qty for line in result.lines}
        self.assertEqual(by_client["Клиент 1"], 48)
        self.assertEqual(by_client["Клиент 2"], 48)

    def test_proportional_allocation_does_not_exceed_stock(self) -> None:
        lines = [
            OrderLine(
                season_id="s1",
                profile_id="p1",
                client_name="Клиент 1",
                source_file=Path("c1.xlsx"),
                order_mtime=datetime(2026, 3, 10, 10, 0, 0),
                sku="ROS-000001",
                requested_qty=96,
                rounded_qty=96,
                was_rounded=False,
                unit_price=100.0,
            ),
            OrderLine(
                season_id="s1",
                profile_id="p1",
                client_name="Клиент 2",
                source_file=Path("c2.xlsx"),
                order_mtime=datetime(2026, 3, 10, 11, 0, 0),
                sku="ROS-000001",
                requested_qty=96,
                rounded_qty=96,
                was_rounded=False,
                unit_price=100.0,
            ),
        ]
        result = allocate_proportional(lines, {"ROS-000001": self.stock})
        total = sum(line.confirmed_qty for line in result.lines)
        self.assertLessEqual(total, 100)
        # При кратности 24 подтверждение всегда кратно.
        self.assertTrue(all(line.confirmed_qty %
                        24 == 0 for line in result.lines))


if __name__ == "__main__":
    unittest.main()
