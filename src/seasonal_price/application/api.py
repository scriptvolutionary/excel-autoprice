from __future__ import annotations

from pathlib import Path
from typing import Any

from seasonal_price.application.bootstrap import build_engine


def init_season(season_id: str, base_dir: Path | None = None) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.init_season(season_id=season_id)


def generate_price(
    stock_file: Path,
    season_id: str,
    output_dir: Path,
    base_dir: Path | None = None,
) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.generate_price(stock_file=stock_file, season_id=season_id, output_dir=output_dir)


def import_orders(
    input_dir: Path,
    season_id: str,
    profile_id: str,
    base_dir: Path | None = None,
    duplicate_strategy: str = "manual",
    duplicate_map_path: Path | None = None,
) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.import_orders(
        input_dir=input_dir,
        season_id=season_id,
        profile_id=profile_id,
        duplicate_strategy=duplicate_strategy,
        duplicate_map_path=duplicate_map_path,
    )


def allocate(
    season_id: str,
    mode: str,
    profile_id: str,
    base_dir: Path | None = None,
) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.allocate(season_id=season_id, mode=mode, profile_id=profile_id)


def export_confirmations(
    season_id: str,
    output_dir: Path,
    pdf_mode: str = "builtin",
    base_dir: Path | None = None,
) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.export_confirmations(season_id=season_id, output_dir=output_dir, pdf_mode=pdf_mode)


def build_residual_price(
    season_id: str,
    output_dir: Path,
    base_dir: Path | None = None,
) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.build_residual_price(season_id=season_id, output_dir=output_dir)


def close_season(
    season_id: str,
    archive_dir: Path,
    base_dir: Path | None = None,
) -> dict[str, Any]:
    engine = build_engine(base_dir=base_dir)
    return engine.close_season(season_id=season_id, archive_dir=archive_dir)
