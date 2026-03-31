from __future__ import annotations

from dataclasses import dataclass
import re


SKU_PATTERN = re.compile(r"^[A-Z0-9]{2,8}-\d{6}$")


def normalize_category_code(category_name: str) -> str:
    """Формирует стабильный префикс категории для SKU.

    Логика:
    - оставляем только буквы/цифры;
    - берем до 3 символов;
    - fallback: CAT.
    """

    cleaned = "".join(ch for ch in category_name.upper() if ch.isalnum())
    if not cleaned:
        return "CAT"
    return cleaned[:3]


@dataclass(frozen=True)
class SkuCandidate:
    category_code: str
    sequence: int

    @property
    def sku(self) -> str:
        return f"{self.category_code}-{self.sequence:06d}"


class SkuRegistryService:
    """Генератор новых SKU формата CAT-000001."""

    def __init__(self, existing_skus: set[str]) -> None:
        self._existing_skus = set(existing_skus)

    def next_sku(self, category_code: str) -> str:
        normalized = category_code.upper().strip() or "CAT"
        sequence = 1
        while True:
            candidate = SkuCandidate(
                category_code=normalized, sequence=sequence).sku
            if candidate not in self._existing_skus:
                self._existing_skus.add(candidate)
                return candidate
            sequence += 1

    @staticmethod
    def validate(sku: str) -> bool:
        return bool(SKU_PATTERN.match(sku.strip().upper()))
