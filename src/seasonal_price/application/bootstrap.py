from __future__ import annotations

from pathlib import Path

from seasonal_price.application.engine import SeasonalPriceEngine
from seasonal_price.config import AppConfig
from seasonal_price.infrastructure.db.sqlite_store import SQLiteStore
from seasonal_price.infrastructure.excel.excel_processor import ExcelProcessor
from seasonal_price.infrastructure.pdf.reportlab_renderer import ReportLabRenderer
from seasonal_price.logging_config import setup_logging


def build_engine(base_dir: Path | None = None) -> SeasonalPriceEngine:
    """Собирает production-конфигурацию приложения."""

    project_dir = base_dir or Path.cwd()
    config = AppConfig.from_base_dir(project_dir)
    config.ensure_directories()
    logger = setup_logging(config.log_dir)
    store = SQLiteStore(config.db_path, logger)
    excel_processor = ExcelProcessor(logger)
    pdf_renderer = ReportLabRenderer(logger)
    return SeasonalPriceEngine(
        config=config,
        logger=logger,
        store=store,
        excel_processor=excel_processor,
        pdf_renderer=pdf_renderer,
    )
