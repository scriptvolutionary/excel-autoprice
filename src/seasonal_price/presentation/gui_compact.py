from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
import traceback
from typing import Any

from seasonal_price.application.bootstrap import build_engine

APP_TITLE = "Сезонный Прайс"
APP_EXE_NAME = "SeasonalPrice"
APP_VERSION = "0.1.0"

try:
    from PySide6.QtCore import QObject, QRunnable, QThreadPool, QUrl, Signal, Slot
    from PySide6.QtGui import QDesktopServices, QFont
    from PySide6.QtWidgets import (
        QApplication,
        QComboBox,
        QDialog,
        QDialogButtonBox,
        QFileDialog,
        QFormLayout,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QProgressBar,
        QPushButton,
        QVBoxLayout,
        QWidget,
    )
except ImportError:  # pragma: no cover - GUI опционален для окружений без Qt
    QApplication = None  # type: ignore[assignment]
    QComboBox = object  # type: ignore[assignment]
    QDesktopServices = object  # type: ignore[assignment]
    QDialog = object  # type: ignore[assignment]
    QDialogButtonBox = object  # type: ignore[assignment]
    QFileDialog = object  # type: ignore[assignment]
    QFont = object  # type: ignore[assignment]
    QFormLayout = object  # type: ignore[assignment]
    QGridLayout = object  # type: ignore[assignment]
    QGroupBox = object  # type: ignore[assignment]
    QHBoxLayout = object  # type: ignore[assignment]
    QLabel = object  # type: ignore[assignment]
    QLineEdit = object  # type: ignore[assignment]
    QMainWindow = object  # type: ignore[assignment]
    QMessageBox = object  # type: ignore[assignment]
    QObject = object  # type: ignore[assignment]
    QProgressBar = object  # type: ignore[assignment]
    QPushButton = object  # type: ignore[assignment]
    QRunnable = object  # type: ignore[assignment]
    QThreadPool = object  # type: ignore[assignment]
    QUrl = object  # type: ignore[assignment]
    QVBoxLayout = object  # type: ignore[assignment]
    QWidget = object  # type: ignore[assignment]

    def Signal(*_args: Any, **_kwargs: Any) -> Any:  # type: ignore[misc]
        return None

    def Slot(*_args: Any, **_kwargs: Any) -> Any:  # type: ignore[misc]
        def _decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
            return fn

        return _decorator


class WorkerSignals(QObject):
    finished = Signal(object)
    failed = Signal(str)
    progress = Signal(object)


class Worker(QRunnable):
    def __init__(
        self,
        fn: Callable[[Callable[[dict[str, Any]], None] | None], Any],
        with_progress: bool = False,
    ) -> None:
        super().__init__()
        self._fn = fn
        self._with_progress = with_progress
        self.signals = WorkerSignals()

    @Slot()
    def run(self) -> None:
        try:
            progress_callback = self.signals.progress.emit if self._with_progress else None
            result = self._fn(progress_callback)
            self.signals.finished.emit(result)
        except Exception:
            self.signals.failed.emit(traceback.format_exc())


class PathField(QWidget):
    def __init__(
        self,
        title: str,
        default_text: str = "",
        pick_mode: str = "dir",
        file_filter: str = "Все файлы (*.*)",
    ) -> None:
        super().__init__()
        self._pick_mode = pick_mode
        self._file_filter = file_filter

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        layout.addWidget(QLabel(title))

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)

        self.edit = QLineEdit(default_text)
        self.edit.setClearButtonEnabled(True)

        btn = QPushButton("Обзор...")
        btn.setAutoDefault(False)
        btn.clicked.connect(self._pick_path)  # type: ignore[arg-type]

        row.addWidget(self.edit, 1)
        row.addWidget(btn)
        layout.addLayout(row)

    def text(self) -> str:
        return self.edit.text().strip()

    def set_text(self, value: str) -> None:
        self.edit.setText(value)

    def _pick_path(self) -> None:
        if self._pick_mode == "dir":
            folder = QFileDialog.getExistingDirectory(self, "Выберите папку")
            if folder:
                self.edit.setText(folder)
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            filter=self._file_filter,
        )
        if file_path:
            self.edit.setText(file_path)


@dataclass
class GuiSettings:
    base_dir: Path
    season_id: str
    profile_id: str
    stock_file: Path
    orders_dir: Path
    price_output_dir: Path
    confirmations_output_dir: Path
    residual_output_dir: Path
    archive_dir: Path
    duplicate_strategy: str
    duplicate_map_path: Path | None
    allocation_mode: str
    rounding_mode: str

    @classmethod
    def default(cls, base_dir: Path | None = None) -> GuiSettings:
        root = (base_dir or Path.cwd()).expanduser()
        out_dir = root / "outputs"
        return cls(
            base_dir=root,
            season_id="spring_2026",
            profile_id="default",
            stock_file=root / "data" / "stock.xlsx",
            orders_dir=root / "orders",
            price_output_dir=out_dir,
            confirmations_output_dir=out_dir,
            residual_output_dir=out_dir,
            archive_dir=root / "archive",
            duplicate_strategy="latest",
            duplicate_map_path=root / "duplicate_map.json",
            allocation_mode="fifo",
            rounding_mode="down",
        )

    def normalized(self) -> GuiSettings:
        self.base_dir = self.base_dir.expanduser()
        self.stock_file = self.stock_file.expanduser()
        self.orders_dir = self.orders_dir.expanduser()
        self.price_output_dir = self.price_output_dir.expanduser()
        self.confirmations_output_dir = self.confirmations_output_dir.expanduser()
        self.residual_output_dir = self.residual_output_dir.expanduser()
        self.archive_dir = self.archive_dir.expanduser()
        if self.duplicate_map_path is not None:
            self.duplicate_map_path = self.duplicate_map_path.expanduser()
        return self

    def price_file(self) -> Path:
        return self.price_output_dir / f"{self.season_id}_Прайс_для_клиентов.xlsx"

    def summary_file(self) -> Path:
        return self.base_dir / "outputs" / self.season_id / self.profile_id / "Сводка_заказов.xlsx"

    def residual_file(self) -> Path:
        return self.residual_output_dir / f"{self.season_id}_Остаточный_прайс.xlsx"


class SettingsDialog(QDialog):
    def __init__(self, settings: GuiSettings, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Настройки")
        self.resize(760, 640)
        self.close_season_requested = False

        self._settings = GuiSettings(**settings.__dict__).normalized()

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        root.addWidget(self._build_general_group())
        root.addWidget(self._build_paths_group())
        root.addWidget(self._build_logic_group())

        footer = QHBoxLayout()
        footer.setContentsMargins(0, 0, 0, 0)
        footer.setSpacing(8)

        btn_defaults = QPushButton("Стандартные пути")
        btn_defaults.setAutoDefault(False)
        btn_defaults.clicked.connect(self._sync_paths_from_base)  # type: ignore[arg-type]

        btn_close_season = QPushButton("Закрыть сезон")
        btn_close_season.setAutoDefault(False)
        btn_close_season.clicked.connect(self._request_close_season)  # type: ignore[arg-type]

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)  # type: ignore[arg-type]
        buttons.rejected.connect(self.reject)  # type: ignore[arg-type]

        footer.addWidget(btn_defaults)
        footer.addWidget(btn_close_season)
        footer.addStretch(1)
        footer.addWidget(buttons)
        root.addLayout(footer)

        self._update_duplicate_map_visibility()

    def _build_general_group(self) -> QGroupBox:
        box = QGroupBox("Основное")
        form = QFormLayout(box)
        form.setSpacing(10)

        self.base_dir_input = PathField("Рабочая папка проекта", str(self._settings.base_dir), pick_mode="dir")
        self.season_edit = QLineEdit(self._settings.season_id)
        self.profile_edit = QLineEdit(self._settings.profile_id)

        form.addRow(self.base_dir_input)
        form.addRow("Сезон", self.season_edit)
        form.addRow("Профиль", self.profile_edit)
        return box

    def _build_paths_group(self) -> QGroupBox:
        box = QGroupBox("Файлы и папки")
        layout = QVBoxLayout(box)
        layout.setSpacing(10)

        self.stock_file_input = PathField(
            "Файл склада (.xlsx/.xls)",
            str(self._settings.stock_file),
            pick_mode="file",
            file_filter="Excel (*.xlsx *.xls);;Все файлы (*.*)",
        )
        self.orders_dir_input = PathField("Папка с заказами", str(self._settings.orders_dir), pick_mode="dir")
        self.price_output_input = PathField("Папка для прайса", str(self._settings.price_output_dir), pick_mode="dir")
        self.confirm_output_input = PathField(
            "Папка для подтверждений",
            str(self._settings.confirmations_output_dir),
            pick_mode="dir",
        )
        self.residual_output_input = PathField(
            "Папка для остаточного прайса",
            str(self._settings.residual_output_dir),
            pick_mode="dir",
        )
        self.archive_dir_input = PathField("Папка архива", str(self._settings.archive_dir), pick_mode="dir")

        layout.addWidget(self.stock_file_input)
        layout.addWidget(self.orders_dir_input)
        layout.addWidget(self.price_output_input)
        layout.addWidget(self.confirm_output_input)
        layout.addWidget(self.residual_output_input)
        layout.addWidget(self.archive_dir_input)
        return box

    def _build_logic_group(self) -> QGroupBox:
        box = QGroupBox("Логика")
        form = QFormLayout(box)
        form.setSpacing(10)

        self.duplicate_strategy_combo = QComboBox()
        self.duplicate_strategy_combo.addItem("Последний файл клиента", "latest")
        self.duplicate_strategy_combo.addItem("Суммировать дубли", "sum")
        self.duplicate_strategy_combo.addItem("Выбор по JSON", "manual")
        index = self.duplicate_strategy_combo.findData(self._settings.duplicate_strategy)
        if index >= 0:
            self.duplicate_strategy_combo.setCurrentIndex(index)
        self.duplicate_strategy_combo.currentIndexChanged.connect(
            self._update_duplicate_map_visibility)  # type: ignore[arg-type]

        duplicate_map_value = ""
        if self._settings.duplicate_map_path is not None:
            duplicate_map_value = str(self._settings.duplicate_map_path)
        self.duplicate_map_input = PathField(
            "JSON-карта дублей",
            duplicate_map_value,
            pick_mode="file",
            file_filter="JSON (*.json);;Все файлы (*.*)",
        )

        self.allocation_mode_combo = QComboBox()
        self.allocation_mode_combo.addItem("FIFO", "fifo")
        self.allocation_mode_combo.addItem("Пропорционально", "proportional")
        index = self.allocation_mode_combo.findData(self._settings.allocation_mode)
        if index >= 0:
            self.allocation_mode_combo.setCurrentIndex(index)

        self.rounding_mode_combo = QComboBox()
        self.rounding_mode_combo.addItem("Вниз до кратности", "down")
        self.rounding_mode_combo.addItem("Вверх до кратности", "up")
        index = self.rounding_mode_combo.findData(self._settings.rounding_mode)
        if index >= 0:
            self.rounding_mode_combo.setCurrentIndex(index)

        form.addRow("Дубли заказов", self.duplicate_strategy_combo)
        form.addRow(self.duplicate_map_input)
        form.addRow("Аллокация", self.allocation_mode_combo)
        form.addRow("Кратность", self.rounding_mode_combo)
        return box

    def _sync_paths_from_base(self) -> None:
        base = Path(self.base_dir_input.text() or str(Path.cwd())).expanduser()
        defaults = GuiSettings.default(base)
        self.stock_file_input.set_text(str(defaults.stock_file))
        self.orders_dir_input.set_text(str(defaults.orders_dir))
        self.price_output_input.set_text(str(defaults.price_output_dir))
        self.confirm_output_input.set_text(str(defaults.confirmations_output_dir))
        self.residual_output_input.set_text(str(defaults.residual_output_dir))
        self.archive_dir_input.set_text(str(defaults.archive_dir))
        self.duplicate_map_input.set_text(str(defaults.duplicate_map_path))

    def _update_duplicate_map_visibility(self) -> None:
        is_manual = str(self.duplicate_strategy_combo.currentData()) == "manual"
        self.duplicate_map_input.setVisible(is_manual)
        self.duplicate_map_input.setEnabled(is_manual)

    def _request_close_season(self) -> None:
        answer = QMessageBox.question(
            self,
            "Подтвердите действие",
            "Закрыть текущий сезон и открыть новый?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer == QMessageBox.Yes:
            self.close_season_requested = True
            self.accept()

    def values(self) -> GuiSettings:
        duplicate_map_path: Path | None = None
        duplicate_map_text = self.duplicate_map_input.text()
        if duplicate_map_text:
            duplicate_map_path = Path(duplicate_map_text)

        return GuiSettings(
            base_dir=Path(self.base_dir_input.text() or str(Path.cwd())),
            season_id=self.season_edit.text().strip() or "spring_2026",
            profile_id=self.profile_edit.text().strip() or "default",
            stock_file=Path(self.stock_file_input.text()),
            orders_dir=Path(self.orders_dir_input.text()),
            price_output_dir=Path(self.price_output_input.text()),
            confirmations_output_dir=Path(self.confirm_output_input.text()),
            residual_output_dir=Path(self.residual_output_input.text()),
            archive_dir=Path(self.archive_dir_input.text()),
            duplicate_strategy=str(self.duplicate_strategy_combo.currentData()),
            duplicate_map_path=duplicate_map_path,
            allocation_mode=str(self.allocation_mode_combo.currentData()),
            rounding_mode=str(self.rounding_mode_combo.currentData()),
        ).normalized()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(540, 470)
        self.setMinimumSize(500, 430)
        self.setMaximumWidth(580)
        self.setMaximumHeight(560)

        self._settings = GuiSettings.default(Path.cwd()).normalized()
        self._pool = QThreadPool.globalInstance()
        self._busy = False
        self._workers: list[Worker] = []
        self._action_buttons: list[QPushButton] = []
        self._engine_instance: Any | None = None
        self._engine_base_dir: Path | None = None

        self._build_ui()
        self._refresh_context()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        top_row = QHBoxLayout()
        top_row.setContentsMargins(0, 0, 0, 0)
        top_row.setSpacing(8)

        title = QLabel(APP_TITLE)
        font = QFont(title.font())
        font.setBold(True)
        font.setPointSize(font.pointSize() + 4)
        title.setFont(font)

        self.btn_settings = QPushButton("Настройки")
        self.btn_settings.setAutoDefault(False)
        self.btn_settings.clicked.connect(self._show_settings)  # type: ignore[arg-type]

        top_row.addWidget(title, 1)
        top_row.addWidget(self.btn_settings)
        layout.addLayout(top_row)

        self.context_label = QLabel()
        self.context_label.setWordWrap(True)
        layout.addWidget(self.context_label)

        self.state_label = QLabel("Готово к работе.")
        layout.addWidget(self.state_label)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        layout.addWidget(self.progress)

        button_grid = QGridLayout()
        button_grid.setContentsMargins(0, 0, 0, 0)
        button_grid.setHorizontalSpacing(10)
        button_grid.setVerticalSpacing(10)

        button_grid.addWidget(self._make_button("Сформировать прайс", self._action_generate_price), 0, 0)
        button_grid.addWidget(self._make_button("Открыть прайс", self._action_open_price), 0, 1)
        button_grid.addWidget(self._make_button("Собрать заказы", self._action_import_orders), 1, 0)
        button_grid.addWidget(self._make_button("Открыть сводную", self._action_open_summary), 1, 1)
        button_grid.addWidget(self._make_button("Создать подтверждения", self._action_create_confirmations), 2, 0)
        button_grid.addWidget(self._make_button("Сформировать остаточный прайс", self._action_build_residual), 2, 1)

        layout.addLayout(button_grid)
        layout.addStretch(1)
        self.setCentralWidget(root)

    def _make_button(self, text: str, slot: Callable[[], None]) -> QPushButton:
        btn = QPushButton(text)
        btn.setMinimumHeight(46)
        btn.setAutoDefault(False)
        btn.clicked.connect(slot)  # type: ignore[arg-type]
        self._action_buttons.append(btn)
        return btn

    def _refresh_context(self) -> None:
        round_label = "вниз" if self._settings.rounding_mode == "down" else "вверх"
        allocation_label = "FIFO" if self._settings.allocation_mode == "fifo" else "пропорционально"
        self.context_label.setText(
            f"Папка: {self._settings.base_dir}\n"
            f"Сезон: {self._settings.season_id} | Профиль: {self._settings.profile_id}\n"
            f"Кратность: {round_label} | Аллокация: {allocation_label}"
        )

    def _show_settings(self) -> None:
        if self._busy:
            return

        dialog = SettingsDialog(self._settings, self)
        if dialog.exec() != QDialog.Accepted:
            return

        self._settings = dialog.values()
        self._refresh_context()
        if dialog.close_season_requested:
            self._action_close_season()

    def _set_busy(self, busy: bool, state_text: str, determinate_progress: bool = False) -> None:
        self._busy = busy
        self.state_label.setText(state_text)
        self.btn_settings.setEnabled(not busy)
        for btn in self._action_buttons:
            btn.setEnabled(not busy)

        if busy:
            if determinate_progress:
                self.progress.setRange(0, 100)
                self.progress.setValue(0)
            else:
                self.progress.setRange(0, 0)
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(100)

    def _run_task(
        self,
        title: str,
        task: Callable[[Callable[[dict[str, Any]], None] | None], Any],
        with_progress: bool = False,
    ) -> None:
        if self._busy:
            QMessageBox.information(self, "Операция уже выполняется", "Дождитесь завершения текущей операции.")
            return

        self._set_busy(True, f"Выполняется: {title}", determinate_progress=with_progress)
        worker = Worker(task, with_progress=with_progress)
        self._workers.append(worker)
        worker.signals.finished.connect(lambda result: self._handle_success(title, result, worker))  # type: ignore[arg-type]
        worker.signals.failed.connect(lambda error: self._handle_error(title, error, worker))  # type: ignore[arg-type]
        if with_progress:
            worker.signals.progress.connect(self._handle_progress)  # type: ignore[arg-type]
        self._pool.start(worker)

    def _handle_success(self, title: str, result: Any, worker: Worker) -> None:
        self._release_worker(worker)
        summary = self._summarize_result(title, result)
        self._set_busy(False, summary)
        self._apply_result_side_effects(result)

    def _handle_error(self, title: str, error_text: str, worker: Worker) -> None:
        self._release_worker(worker)
        self._set_busy(False, f"Ошибка: {title}")
        self.progress.setValue(0)
        short_text = error_text.strip().splitlines()[-1] if error_text.strip() else "Неизвестная ошибка"
        QMessageBox.critical(self, title, short_text)

    def _handle_progress(self, payload: Any) -> None:
        if not isinstance(payload, dict):
            return
        percent_raw = payload.get("percent")
        if isinstance(percent_raw, (int, float)):
            percent = max(0, min(int(percent_raw), 100))
            if self.progress.minimum() == 0 and self.progress.maximum() == 0:
                self.progress.setRange(0, 100)
            self.progress.setValue(percent)
        message = payload.get("message")
        if isinstance(message, str) and message.strip():
            self.state_label.setText(message.strip())

    def _release_worker(self, worker: Worker) -> None:
        try:
            self._workers.remove(worker)
        except ValueError:
            return

    def _apply_result_side_effects(self, result: Any) -> None:
        if not isinstance(result, dict):
            return
        new_season = result.get("new_season")
        if isinstance(new_season, str) and new_season:
            self._settings.season_id = new_season
            self._refresh_context()

    def _summarize_result(self, title: str, result: Any) -> str:
        if not isinstance(result, dict):
            return f"Готово: {title}"
        if title == "Формирование прайса":
            return f"Прайс сформирован: {result.get('items', 0)} позиций."
        if title == "Сбор заказов":
            return (
                f"Заказы собраны: файлов {result.get('success_files', 0)}, "
                f"ошибок {result.get('error_files', 0)}, строк {result.get('imported_lines', 0)}."
            )
        if title == "Подтверждения":
            allocate_result = result.get("allocate", {})
            confirm_result = result.get("confirmations", {})
            if isinstance(allocate_result, dict) and isinstance(confirm_result, dict):
                return (
                    f"Подтверждения готовы: клиентов {confirm_result.get('clients', 0)}, "
                    f"режим {allocate_result.get('mode', '—')}."
                )
        if title == "Остаточный прайс":
            return f"Остаточный прайс сформирован: {result.get('positions', 0)} позиций."
        if title == "Закрытие сезона":
            return f"Сезон закрыт. Новый сезон: {result.get('new_season', '—')}."
        return f"Готово: {title}"

    def _engine_for_base(self, base_dir: Path) -> Any:
        if self._engine_instance is None or self._engine_base_dir != base_dir:
            self._engine_instance = build_engine(base_dir)
            self._engine_base_dir = base_dir
        return self._engine_instance

    @staticmethod
    def _ensure_file_exists(path_text: str, title: str, allowed_suffixes: set[str] | None = None) -> Path:
        if not path_text.strip():
            raise ValueError(f"Не заполнено поле: {title}.")
        path = Path(path_text).expanduser()
        if not path.exists() or not path.is_file():
            raise ValueError(f"Файл не найден: {path}")
        if allowed_suffixes is not None and path.suffix.lower() not in allowed_suffixes:
            suffixes = ", ".join(sorted(allowed_suffixes))
            raise ValueError(f"Неверный формат файла ({path.name}). Ожидается: {suffixes}")
        return path

    @staticmethod
    def _ensure_dir_exists(path_text: str, title: str) -> Path:
        if not path_text.strip():
            raise ValueError(f"Не заполнено поле: {title}.")
        path = Path(path_text).expanduser()
        if not path.exists() or not path.is_dir():
            raise ValueError(f"Папка не найдена: {path}")
        return path

    @staticmethod
    def _ensure_dir(path_text: str, title: str) -> Path:
        if not path_text.strip():
            raise ValueError(f"Не заполнено поле: {title}.")
        path = Path(path_text).expanduser()
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _optional_duplicate_map(self) -> Path | None:
        if self._settings.duplicate_strategy != "manual":
            return None
        duplicate_map = self._settings.duplicate_map_path
        if duplicate_map is None:
            raise ValueError("Для ручного режима дублей укажите JSON-карту.")
        return self._ensure_file_exists(str(duplicate_map), "JSON-карта дублей", {".json"})

    def _show_validation_error(self, message: str) -> None:
        QMessageBox.warning(self, "Проверьте входные данные", message)

    def _action_generate_price(self) -> None:
        try:
            settings = self._settings.normalized()
            base_dir = self._ensure_dir(str(settings.base_dir), "Рабочая папка проекта")
            stock_file = self._ensure_file_exists(str(settings.stock_file), "Файл склада", {".xlsx", ".xls"})
            output_dir = self._ensure_dir(str(settings.price_output_dir), "Папка для прайса")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.generate_price(stock_file=stock_file, season_id=settings.season_id, output_dir=output_dir)

        self._run_task("Формирование прайса", task)

    def _action_open_price(self) -> None:
        self._open_path(self._settings.price_file(), "Клиентский прайс")

    def _action_import_orders(self) -> None:
        try:
            settings = self._settings.normalized()
            base_dir = self._ensure_dir(str(settings.base_dir), "Рабочая папка проекта")
            orders_dir = self._ensure_dir_exists(str(settings.orders_dir), "Папка с заказами")
            duplicate_map = self._optional_duplicate_map()
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.import_orders(
                input_dir=orders_dir,
                season_id=settings.season_id,
                profile_id=settings.profile_id,
                duplicate_strategy=settings.duplicate_strategy,
                duplicate_map_path=duplicate_map,
                rounding_mode=settings.rounding_mode,
                progress_callback=_progress,
            )

        self._run_task("Сбор заказов", task, with_progress=True)

    def _action_open_summary(self) -> None:
        self._open_path(self._settings.summary_file(), "Сводка заказов")

    def _action_create_confirmations(self) -> None:
        try:
            settings = self._settings.normalized()
            base_dir = self._ensure_dir(str(settings.base_dir), "Рабочая папка проекта")
            output_dir = self._ensure_dir(str(settings.confirmations_output_dir), "Папка для подтверждений")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            if _progress is not None:
                _progress({"percent": 20, "message": "Аллокация заказов"})
            allocate_result = engine.allocate(
                season_id=settings.season_id,
                mode=settings.allocation_mode,
                profile_id=settings.profile_id,
            )
            if _progress is not None:
                _progress({"percent": 60, "message": "Создание подтверждений"})
            confirmations_result = engine.export_confirmations(
                season_id=settings.season_id,
                output_dir=output_dir,
            )
            if _progress is not None:
                _progress({"percent": 100, "message": "Подтверждения сформированы"})
            return {
                "allocate": allocate_result,
                "confirmations": confirmations_result,
            }

        self._run_task("Подтверждения", task, with_progress=True)

    def _action_build_residual(self) -> None:
        try:
            settings = self._settings.normalized()
            base_dir = self._ensure_dir(str(settings.base_dir), "Рабочая папка проекта")
            output_dir = self._ensure_dir(str(settings.residual_output_dir), "Папка для остаточного прайса")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.build_residual_price(season_id=settings.season_id, output_dir=output_dir)

        self._run_task("Остаточный прайс", task)

    def _action_close_season(self) -> None:
        try:
            settings = self._settings.normalized()
            base_dir = self._ensure_dir(str(settings.base_dir), "Рабочая папка проекта")
            archive_dir = self._ensure_dir(str(settings.archive_dir), "Папка архива")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.close_season(season_id=settings.season_id, archive_dir=archive_dir)

        self._run_task("Закрытие сезона", task)

    def _open_path(self, path: Path, title: str) -> None:
        if not path.exists():
            QMessageBox.warning(self, title, f"Путь пока не найден:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

def _main_window_summarize_result(self: MainWindow, title: str, result: Any) -> str:
    if not isinstance(result, dict):
        return f"Готово: {title}"
    if title == "Формирование прайса":
        pdf_created = " PDF готов." if result.get("pdf_file") else ""
        return f"Прайс сформирован: {result.get('items', 0)} позиций.{pdf_created}"
    if title == "Сбор заказов":
        return (
            f"Заказы собраны: файлов {result.get('success_files', 0)}, "
            f"ошибок {result.get('error_files', 0)}, строк {result.get('imported_lines', 0)}."
        )
    if title == "Подтверждения":
        allocate_result = result.get("allocate", {})
        confirm_result = result.get("confirmations", {})
        if isinstance(allocate_result, dict) and isinstance(confirm_result, dict):
            return (
                f"Подтверждения готовы: клиентов {confirm_result.get('clients', 0)}, "
                f"режим {allocate_result.get('mode', '—')}."
            )
    if title == "Остаточный прайс":
        return f"Остаточный прайс сформирован: {result.get('positions', 0)} позиций."
    if title == "Закрытие сезона":
        return f"Сезон закрыт. Новый сезон: {result.get('new_season', '—')}."
    return f"Готово: {title}"


MainWindow._summarize_result = _main_window_summarize_result


def run_gui() -> None:
    if QApplication is None:
        raise RuntimeError(
            "PySide6 не установлен. Установите зависимости: python -m pip install -e .[gui]")

    app = QApplication([])
    app.setApplicationName(APP_TITLE)
    app.setApplicationDisplayName(APP_TITLE)
    app.setOrganizationName(APP_EXE_NAME)

    try:
        window = MainWindow()
        window.show()
        app.exec()
    except Exception:
        error_text = traceback.format_exc()
        error_log = Path.cwd() / "gui_startup_error.log"
        error_log.write_text(error_text, encoding="utf-8")
        QMessageBox.critical(
            None,
            "Ошибка запуска",
            (
                "Не удалось запустить приложение.\n\n"
                f"Детали записаны в файл:\n{error_log}"
            ),
        )
        raise
