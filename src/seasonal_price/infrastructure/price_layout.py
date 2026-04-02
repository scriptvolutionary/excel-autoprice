from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from seasonal_price.domain.models import StockItem


@dataclass(frozen=True)
class PriceSection:
    title: str
    subtitle: str
    multiplicity_note: str
    items: tuple[StockItem, ...]

    @property
    def row_count(self) -> int:
        return 4 + len(self.items)


def build_price_sections(stock_items: list[StockItem]) -> list[PriceSection]:
    grouped: dict[tuple[str, str, str, int], list[StockItem]] = defaultdict(list)
    for item in sorted(
        stock_items,
        key=lambda value: (
            value.category_name.strip().casefold(),
            value.container.strip().casefold(),
            value.size.strip().casefold(),
            value.multiplicity,
            value.sort_name.strip().casefold(),
            value.sku,
        ),
    ):
        key = (
            item.category_name.strip(),
            item.container.strip(),
            item.size.strip(),
            item.multiplicity,
        )
        grouped[key].append(item)

    sections: list[PriceSection] = []
    for key in sorted(
        grouped,
        key=lambda value: (
            value[0].casefold(),
            value[1].casefold(),
            value[2].casefold(),
            value[3],
        ),
    ):
        category_name, container, size, multiplicity = key
        parts = [part for part in (container, size) if part]
        subtitle = ". ".join(
            filter(
                None,
                [
                    f"Контейнер: {container}" if container else "",
                    f"Размер: {size}" if size else "",
                ],
            )
        )
        if multiplicity > 1:
            multiplicity_note = f"Заказ кратен {multiplicity} шт."
        else:
            multiplicity_note = "Заказ без ограничения по кратности."
        if not subtitle and parts:
            subtitle = ". ".join(parts)
        sections.append(
            PriceSection(
                title=category_name or "Без категории",
                subtitle=subtitle,
                multiplicity_note=multiplicity_note,
                items=tuple(grouped[key]),
            )
        )
    return sections


def split_price_sections(
    sections: list[PriceSection],
    *,
    first_sheet_max_rows: int = 180,
    next_sheet_max_rows: int = 210,
) -> list[list[PriceSection]]:
    if not sections:
        return [[]]

    pages: list[list[PriceSection]] = []
    current_page: list[PriceSection] = []
    current_rows = 0
    current_limit = first_sheet_max_rows

    for section in sections:
        needed_rows = section.row_count
        if current_page and current_rows + needed_rows > current_limit:
            pages.append(current_page)
            current_page = []
            current_rows = 0
            current_limit = next_sheet_max_rows
        current_page.append(section)
        current_rows += needed_rows

    if current_page:
        pages.append(current_page)
    return pages


def format_season_title(season_id: str) -> str:
    normalized = season_id.replace("_", " ").replace("-", " ").strip()
    if not normalized:
        return "ПРАЙС"

    mapping = {
        "spring": "ВЕСНА",
        "summer": "ЛЕТО",
        "autumn": "ОСЕНЬ",
        "fall": "ОСЕНЬ",
        "winter": "ЗИМА",
    }
    words = []
    for word in normalized.split():
        words.append(mapping.get(word.casefold(), word.upper()))
    return f"ПРАЙС {' '.join(words)}".strip()
