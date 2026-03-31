from __future__ import annotations

from datetime import datetime
import logging
from pathlib import Path
from typing import Iterable

from seasonal_price.domain.models import AllocationLine, StockItem
from seasonal_price.exceptions import PdfExportError


class ReportLabRenderer:
    """Рендер подтверждений в PDF.

    Режим `builtin` работает без внешних офисных зависимостей.
    """

    def __init__(self, logger: logging.Logger) -> None:
        self._logger = logger

    def render_confirmation_pdf(
        self,
        output_path: Path,
        client_name: str,
        lines: list[tuple[StockItem, AllocationLine]],
        mode: str = "builtin",
    ) -> None:
        if mode != "builtin":
            raise PdfExportError(
                f"Режим PDF '{mode}' не поддерживается. Доступен только встроенный режим 'builtin'."
            )
        self._render_builtin(output_path, client_name, lines)

    def _render_builtin(
        self, output_path: Path, client_name: str, lines: list[tuple[StockItem, AllocationLine]]
    ) -> None:
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.units import mm
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
        except Exception as exc:  # pragma: no cover - зависит от окружения
            raise PdfExportError(
                "Не удалось импортировать reportlab для PDF-экспорта.") from exc

        # Пытаемся зарегистрировать шрифт с кириллицей. Если не найден, используем fallback.
        font_registered = False
        candidate_fonts = [
            Path("C:/Windows/Fonts/arial.ttf"),
            Path("C:/Windows/Fonts/calibri.ttf"),
        ]
        for font_path in candidate_fonts:
            if font_path.exists():
                pdfmetrics.registerFont(TTFont("CyrillicFont", str(font_path)))
                font_registered = True
                break
        font_name = "CyrillicFont" if font_registered else "Helvetica"

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc = SimpleDocTemplate(str(output_path), pagesize=landscape(
            A4), leftMargin=12 * mm, rightMargin=12 * mm)
        styles = getSampleStyleSheet()
        title_style = styles["Heading2"]
        title_style.fontName = font_name
        body_style = styles["Normal"]
        body_style.fontName = font_name

        elements = [
            Paragraph(f"Подтверждение заказа: {client_name}", title_style),
            Spacer(1, 8),
            Paragraph(
                f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M')}", body_style),
            Spacer(1, 8),
        ]

        data = [["SKU", "Категория", "Сорт", "Контейнер",
                 "Размер", "Цена", "Заказ", "Подтверждено", "Сумма"]]
        total_sum = 0.0
        for stock, alloc in lines:
            line_sum = stock.price * alloc.confirmed_qty
            total_sum += line_sum
            data.append(
                [
                    stock.sku,
                    stock.category_name,
                    stock.sort_name,
                    stock.container,
                    stock.size,
                    f"{stock.price:.2f}",
                    str(alloc.requested_qty),
                    str(alloc.confirmed_qty),
                    f"{line_sum:.2f}",
                ]
            )
        data.append(["", "", "", "", "", "", "", "Итого", f"{total_sum:.2f}"])

        table = Table(data, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), font_name),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("ALIGN", (5, 1), (-1, -1), "RIGHT"),
                ]
            )
        )
        elements.append(table)
        doc.build(elements)
