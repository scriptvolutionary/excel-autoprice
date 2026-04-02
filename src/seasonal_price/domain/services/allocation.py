from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from math import floor

from seasonal_price.domain.models import AllocationLine, OrderLine, StockItem
from seasonal_price.domain.services.rounding import round_down_to_multiplicity


@dataclass(frozen=True)
class AllocationResult:
    lines: list[AllocationLine]
    corrected_count: int


def _sort_fifo(lines: list[OrderLine]) -> list[OrderLine]:
    return sorted(lines, key=lambda line: (line.order_mtime, str(line.source_file)))


def allocate_fifo(order_lines: list[OrderLine], stock_map: dict[str, StockItem]) -> AllocationResult:
    grouped: dict[str, list[OrderLine]] = defaultdict(list)
    for line in order_lines:
        grouped[line.sku].append(line)

    out: list[AllocationLine] = []
    corrected = 0
    for sku, sku_lines in grouped.items():
        stock = stock_map[sku]
        remaining = stock.stock_total
        for line in _sort_fifo(sku_lines):
            raw_confirmed = min(line.rounded_qty, remaining)
            confirmed, was_rounded = round_down_to_multiplicity(
                raw_confirmed, stock.multiplicity)
            remaining -= confirmed
            if was_rounded:
                corrected += 1
            out.append(
                AllocationLine(
                    season_id=line.season_id,
                    profile_id=line.profile_id,
                    client_name=line.client_name,
                    source_file=line.source_file,
                    order_mtime=line.order_mtime,
                    sku=sku,
                    requested_qty=line.requested_qty,
                    rounded_qty=line.rounded_qty,
                    confirmed_qty=confirmed,
                    was_rounded=was_rounded,
                    allocation_mode="fifo",
                )
            )
    return AllocationResult(lines=out, corrected_count=corrected)


def allocate_proportional(
    order_lines: list[OrderLine], stock_map: dict[str, StockItem]
) -> AllocationResult:
    grouped: dict[str, list[OrderLine]] = defaultdict(list)
    for line in order_lines:
        grouped[line.sku].append(line)

    out: list[AllocationLine] = []
    corrected = 0

    for sku, sku_lines in grouped.items():
        stock = stock_map[sku]
        lines = _sort_fifo(sku_lines)
        total_requested = sum(line.rounded_qty for line in lines)
        if total_requested <= stock.stock_total:
            raw_allocations = [line.rounded_qty for line in lines]
        else:
            # Шаг 1: базовое пропорциональное деление целыми числами.
            factors = [
                (line.rounded_qty * stock.stock_total) /
                total_requested if total_requested else 0.0
                for line in lines
            ]
            raw_allocations = [floor(value) for value in factors]
            remainder = stock.stock_total - sum(raw_allocations)
            order = sorted(
                range(len(lines)),
                key=lambda idx: (
                    factors[idx] - raw_allocations[idx],
                    -lines[idx].order_mtime.timestamp(),
                ),
                reverse=True,
            )
            for idx in order:
                if remainder <= 0:
                    break
                if raw_allocations[idx] < lines[idx].rounded_qty:
                    raw_allocations[idx] += 1
                    remainder -= 1

        # Шаг 2: приведение к кратности по каждой строке.
        confirmed_allocations: list[int] = []
        for raw in raw_allocations:
            confirmed, _ = round_down_to_multiplicity(raw, stock.multiplicity)
            confirmed_allocations.append(confirmed)

        # Шаг 3: повторная раздача оставшихся штук с учетом кратности.
        remaining = stock.stock_total - sum(confirmed_allocations)
        chunk = max(stock.multiplicity, 1)
        if chunk > 1 and remaining >= chunk:
            order = sorted(
                range(len(lines)),
                key=lambda idx: (
                    (raw_allocations[idx] - confirmed_allocations[idx]),
                    -lines[idx].order_mtime.timestamp(),
                ),
                reverse=True,
            )
            for idx in order:
                need = lines[idx].rounded_qty - confirmed_allocations[idx]
                while need >= chunk and remaining >= chunk:
                    confirmed_allocations[idx] += chunk
                    need -= chunk
                    remaining -= chunk

        for line, raw, confirmed in zip(lines, raw_allocations, confirmed_allocations, strict=True):
            was_rounded = confirmed != raw
            if was_rounded:
                corrected += 1
            out.append(
                AllocationLine(
                    season_id=line.season_id,
                    profile_id=line.profile_id,
                    client_name=line.client_name,
                    source_file=line.source_file,
                    order_mtime=line.order_mtime,
                    sku=sku,
                    requested_qty=line.requested_qty,
                    rounded_qty=line.rounded_qty,
                    confirmed_qty=confirmed,
                    was_rounded=was_rounded,
                    allocation_mode="proportional",
                )
            )
    return AllocationResult(lines=out, corrected_count=corrected)
