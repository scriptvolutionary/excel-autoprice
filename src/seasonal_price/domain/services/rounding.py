from __future__ import annotations


def round_down_to_multiplicity(quantity: int, multiplicity: int) -> tuple[int, bool]:
    """Округляет количество вниз до кратности.

    Возвращает (скорректированное_значение, был_ли_пересчет).
    """

    if quantity < 0:
        raise ValueError("Количество не может быть отрицательным.")
    if multiplicity <= 1:
        return quantity, False
    rounded = (quantity // multiplicity) * multiplicity
    return rounded, rounded != quantity
