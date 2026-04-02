from __future__ import annotations

from datetime import datetime
import logging
from pathlib import Path

from seasonal_price.domain.models import AllocationLine, StockItem
from seasonal_price.exceptions import PdfExportError
from seasonal_price.infrastructure.price_layout import build_price_sections, format_season_title


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

    def render_price_pdf(
        self,
        output_path: Path,
        season_id: str,
        stock_items: list[StockItem],
        mode: str = "builtin",
    ) -> None:
        if mode != "builtin":
            raise PdfExportError(
                f"Режим PDF '{mode}' не поддерживается. Доступен только встроенный режим 'builtin'."
            )
        self._render_price_builtin(output_path, season_id, stock_items)

    def _render_builtin(
        self, output_path: Path, client_name: str, lines: list[tuple[StockItem, AllocationLine]]
    ) -> None:
        try:
            from reportlab.lib import colors
            from reportlab.lib.enums import TA_CENTER, TA_RIGHT
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
            from reportlab.lib.units import mm
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle
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
        page_size = landscape(A4)
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=page_size,
            leftMargin=8 * mm,
            rightMargin=8 * mm,
            topMargin=8 * mm,
            bottomMargin=8 * mm,
        )
        styles = getSampleStyleSheet()
        title_style = styles["Heading2"]
        title_style.fontName = font_name
        body_style = styles["Normal"]
        body_style.fontName = font_name
        body_style.fontSize = 9
        body_style.leading = 11
        cell_style = ParagraphStyle(
            "Cell",
            parent=body_style,
            fontName=font_name,
            fontSize=8,
            leading=9,
            wordWrap="CJK",
        )
        header_style = ParagraphStyle(
            "Header",
            parent=cell_style,
            alignment=TA_CENTER,
        )
        numeric_style = ParagraphStyle(
            "Numeric",
            parent=cell_style,
            alignment=TA_RIGHT,
        )

        elements = [
            Paragraph(f"Подтверждение заказа: {client_name}", title_style),
            Spacer(1, 8),
            Paragraph(
                f"Дата: {datetime.now().strftime('%Y-%m-%d %H:%M')}", body_style),
            Spacer(1, 8),
        ]

        data: list[list[Paragraph]] = [[
            Paragraph("SKU", header_style),
            Paragraph("Категория", header_style),
            Paragraph("Сорт", header_style),
            Paragraph("Контейнер", header_style),
            Paragraph("Размер", header_style),
            Paragraph("Цена", header_style),
            Paragraph("Заказано", header_style),
            Paragraph("К расчету", header_style),
            Paragraph("Подтверждено", header_style),
            Paragraph("Сумма", header_style),
        ]]
        total_sum = 0.0
        for stock, alloc in lines:
            line_sum = stock.price * alloc.confirmed_qty
            total_sum += line_sum
            data.append(
                [
                    Paragraph(stock.sku, cell_style),
                    Paragraph(stock.category_name, cell_style),
                    Paragraph(stock.sort_name, cell_style),
                    Paragraph(stock.container, cell_style),
                    Paragraph(stock.size, cell_style),
                    Paragraph(f"{stock.price:.2f}", numeric_style),
                    Paragraph(str(alloc.requested_qty), numeric_style),
                    Paragraph(str(alloc.rounded_qty), numeric_style),
                    Paragraph(str(alloc.confirmed_qty), numeric_style),
                    Paragraph(f"{line_sum:.2f}", numeric_style),
                ]
            )
        data.append(
            [
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph("Итого", numeric_style),
                Paragraph(f"{total_sum:.2f}", numeric_style),
            ]
        )

        available_width = page_size[0] - doc.leftMargin - doc.rightMargin
        base_widths = [24, 30, 42, 22, 22, 18, 18, 18, 20, 22]
        total_base = sum(base_widths)
        col_widths = [(available_width * width / total_base) for width in base_widths]

        table = LongTable(data, colWidths=col_widths, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        elements.append(table)
        doc.build(elements)

    def _render_price_builtin(
        self,
        output_path: Path,
        season_id: str,
        stock_items: list[StockItem],
    ) -> None:
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
            from reportlab.lib.units import mm
            from reportlab.pdfbase import pdfmetrics
            from reportlab.pdfbase.ttfonts import TTFont
            from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle
        except Exception as exc:  # pragma: no cover - зависит от окружения
            raise PdfExportError(
                "Не удалось импортировать reportlab для PDF-экспорта."
            ) from exc

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
        page_size = landscape(A4)
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=page_size,
            leftMargin=8 * mm,
            rightMargin=8 * mm,
            topMargin=8 * mm,
            bottomMargin=8 * mm,
        )
        styles = getSampleStyleSheet()
        title_style = styles["Heading2"]
        title_style.fontName = font_name
        title_style.fontSize = 16
        title_style.leading = 19

        body_style = styles["Normal"]
        body_style.fontName = font_name
        body_style.fontSize = 9
        body_style.leading = 11

        section_style = ParagraphStyle(
            "PriceSection",
            parent=body_style,
            fontName=font_name,
            fontSize=11,
            leading=13,
            textColor=colors.HexColor("#355E3B"),
        )
        note_style = ParagraphStyle(
            "PriceNote",
            parent=body_style,
            fontName=font_name,
            fontSize=8,
            leading=10,
        )
        header_style = ParagraphStyle(
            "PriceHeader",
            parent=body_style,
            fontName=font_name,
            fontSize=8,
            leading=9,
            alignment=1,
        )
        cell_style = ParagraphStyle(
            "PriceCell",
            parent=body_style,
            fontName=font_name,
            fontSize=8,
            leading=9,
            wordWrap="CJK",
        )
        numeric_style = ParagraphStyle(
            "PriceNumeric",
            parent=cell_style,
            alignment=2,
        )

        elements = [
            Paragraph(format_season_title(season_id), title_style),
            Spacer(1, 6),
            Paragraph("Оптовый прайс. Для заполнения заказа используйте Excel-версию шаблона.", body_style),
            Paragraph("Группировка сохранена по категориям, контейнеру, размеру и кратности.", body_style),
            Spacer(1, 8),
        ]

        available_width = page_size[0] - doc.leftMargin - doc.rightMargin
        base_widths = [18, 56, 20, 20, 16, 16]
        total_base = sum(base_widths)
        col_widths = [(available_width * width / total_base) for width in base_widths]

        for section in build_price_sections(stock_items):
            elements.append(Paragraph(section.title, section_style))
            if section.subtitle:
                elements.append(Paragraph(section.subtitle, note_style))
            elements.append(Paragraph(section.multiplicity_note, note_style))
            elements.append(Spacer(1, 4))

            data = [[
                Paragraph("SKU", header_style),
                Paragraph("Культура/сорт", header_style),
                Paragraph("Контейнер", header_style),
                Paragraph("Размер", header_style),
                Paragraph("Цена", header_style),
                Paragraph("Кратность", header_style),
            ]]
            for item in section.items:
                data.append(
                    [
                        Paragraph(item.sku, cell_style),
                        Paragraph(item.sort_name, cell_style),
                        Paragraph(item.container, cell_style),
                        Paragraph(item.size, cell_style),
                        Paragraph(f"{item.price:.0f}", numeric_style),
                        Paragraph(str(item.multiplicity), numeric_style),
                    ]
                )

            table = LongTable(data, colWidths=col_widths, repeatRows=1)
            table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EFE6C8")),
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 3),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                        ("BACKGROUND", (0, 1), (0, -1), colors.HexColor("#E8F4FD")),
                        ("BACKGROUND", (1, 1), (1, -1), colors.HexColor("#FFF59D")),
                    ]
                )
            )
            elements.append(table)
            elements.append(Spacer(1, 8))

        doc.build(elements)
