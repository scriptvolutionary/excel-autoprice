from __future__ import annotations


def round_to_multiplicity(
    quantity: int,
    multiplicity: int,
    mode: str = "down",
) -> tuple[int, bool]:
    """Округляет количество до кратности по выбранному правилу.

    Поддерживаемые режимы:
    - ``down``: округление вниз;
    - ``up``: округление вверх.
    """

    if quantity < 0:
        raise ValueError("Количество не может быть отрицательным.")
    if multiplicity <= 1:
        return quantity, False

    if mode == "down":
        rounded = (quantity // multiplicity) * multiplicity
    elif mode == "up":
        rounded = 0 if quantity == 0 else ((quantity + multiplicity - 1) // multiplicity) * multiplicity
    else:
        raise ValueError(f"Неизвестный режим округления по кратности: {mode}")

    return rounded, rounded != quantity


def round_down_to_multiplicity(quantity: int, multiplicity: int) -> tuple[int, bool]:
    """Округляет количество вниз до кратности.

    Возвращает (скорректированное_значение, был_ли_пересчет).
    """

    return round_to_multiplicity(quantity, multiplicity, mode="down")
