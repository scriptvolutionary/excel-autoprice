from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class StockItem:
    sku: str
    category_name: str
    category_code: str
    sort_name: str
    container: str
    size: str
    price: float
    multiplicity: int
    stock_total: int

    @property
    def legacy_key(self) -> tuple[str, str, str]:
        """Ключ для legacy-сопоставления без SKU."""

        return (
            self.sort_name.strip().casefold(),
            self.container.strip().casefold(),
            self.size.strip().casefold(),
        )


@dataclass(frozen=True)
class OrderLine:
    season_id: str
    profile_id: str
    client_name: str
    source_file: Path
    order_mtime: datetime
    sku: str
    requested_qty: int
    rounded_qty: int
    was_rounded: bool
    unit_price: float


@dataclass(frozen=True)
class AllocationLine:
    season_id: str
    profile_id: str
    client_name: str
    source_file: Path
    order_mtime: datetime
    sku: str
    requested_qty: int
    confirmed_qty: int
    was_rounded: bool
    allocation_mode: str


@dataclass(frozen=True)
class ImportIssue:
    file_path: Path
    issue_code: str
    message: str
    row_index: int | None = None
    sku: str | None = None


@dataclass(frozen=True)
class ImportFileMeta:
    file_path: Path
    client_name: str
    mtime: datetime
