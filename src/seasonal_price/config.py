from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class AppConfig:
    """Конфигурация путей приложения.

    Все директории создаются рядом с рабочим проектом, чтобы систему можно было
    переносить между машинами без внешних зависимостей.
    """

    base_dir: Path
    data_dir: Path
    output_dir: Path
    log_dir: Path
    archive_dir: Path
    db_path: Path

    @classmethod
    def from_base_dir(cls, base_dir: Path) -> AppConfig:
        data_dir = base_dir / "data"
        output_dir = base_dir / "outputs"
        log_dir = base_dir / "logs"
        archive_dir = base_dir / "archive"
        db_path = data_dir / "seasonal_price.db"
        return cls(
            base_dir=base_dir,
            data_dir=data_dir,
            output_dir=output_dir,
            log_dir=log_dir,
            archive_dir=archive_dir,
            db_path=db_path,
        )

    def ensure_directories(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.archive_dir.mkdir(parents=True, exist_ok=True)
