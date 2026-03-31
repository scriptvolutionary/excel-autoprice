from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from pathlib import Path
import traceback
from typing import Any

from seasonal_price.application.bootstrap import build_engine

APP_TITLE = "Сезонный Прайс"
APP_EXE_NAME = "SeasonalPrice"
APP_VERSION = "0.1.0"

try:
    from PySide6.QtCore import QObject, QRunnable, Qt, QThreadPool, QUrl, Signal, Slot
    from PySide6.QtGui import QDesktopServices
    from PySide6.QtWidgets import (
        QApplication,
        QComboBox,
        QDialog,
        QFileDialog,
        QFrame,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPlainTextEdit,
        QProgressBar,
        QPushButton,
        QScrollArea,
        QSplitter,
        QTabWidget,
        QTextBrowser,
        QVBoxLayout,
        QWidget,
    )
except ImportError:  # pragma: no cover - GUI опционален для окружений без Qt
    QApplication = None  # type: ignore[assignment]
    QComboBox = object  # type: ignore[assignment]
    QDesktopServices = object  # type: ignore[assignment]
    QDialog = object  # type: ignore[assignment]
    QFileDialog = object  # type: ignore[assignment]
    QFrame = object  # type: ignore[assignment]
    QHBoxLayout = object  # type: ignore[assignment]
    QLabel = object  # type: ignore[assignment]
    QLineEdit = object  # type: ignore[assignment]
    QListWidget = object  # type: ignore[assignment]
    QListWidgetItem = object  # type: ignore[assignment]
    QMainWindow = object  # type: ignore[assignment]
    QMessageBox = object  # type: ignore[assignment]
    QObject = object  # type: ignore[assignment]
    QPlainTextEdit = object  # type: ignore[assignment]
    QProgressBar = object  # type: ignore[assignment]
    QPushButton = object  # type: ignore[assignment]
    QRunnable = object  # type: ignore[assignment]
    QScrollArea = object  # type: ignore[assignment]
    QSplitter = object  # type: ignore[assignment]
    QTabWidget = object  # type: ignore[assignment]
    QTextBrowser = object  # type: ignore[assignment]
    QThreadPool = object  # type: ignore[assignment]
    Qt = object  # type: ignore[assignment]
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
    """Фоновый worker для запуска операций ядра без блокировки GUI-потока."""

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
        hint: str = "",
        placeholder: str = "",
    ) -> None:
        super().__init__()
        self._pick_mode = pick_mode
        self._file_filter = file_filter

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        label = QLabel(title)
        label.setObjectName("FieldLabel")
        layout.addWidget(label)

        if hint:
            hint_label = QLabel(hint)
            hint_label.setObjectName("HintLabel")
            hint_label.setWordWrap(True)
            layout.addWidget(hint_label)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)

        self.edit = QLineEdit(default_text)
        self.edit.setMinimumHeight(40)
        self.edit.setClearButtonEnabled(True)
        if placeholder:
            self.edit.setPlaceholderText(placeholder)

        btn = QPushButton("Выбрать")
        btn.setObjectName("SecondaryButton")
        btn.setMinimumHeight(40)
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


class HelpDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Справка - {APP_TITLE}")
        self.resize(920, 720)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setMarkdown(
            """
# Как работать в программе

## Что подготовить заранее

1. Файл склада Excel с листом **Склад**.
2. Папку с заказами клиентов (`.xlsx`/`.xls`).
3. Рабочую папку проекта (в ней будут `data`, `outputs`, `logs`).

## Рекомендуемая последовательность

1. Нажмите **Подготовить сезон**.
2. На вкладке **Прайс** сформируйте клиентский прайс из файла склада.
3. На вкладке **Заказы** импортируйте папку с заказами.
4. На вкладке **Аллокация** выберите режим распределения и запустите расчет.
5. На вкладке **Подтверждения** сформируйте файлы клиентам.
6. На вкладке **Остатки** сформируйте остаточный прайс.

## Дубли заказов

- `latest` - берется самый новый файл клиента.
- `sum` - суммируются все дубли клиента.
- `manual` - выбор файла по JSON-карте.

## Где искать результаты

- `outputs/` - итоговые Excel/PDF.
- `logs/` - лог работы приложения.
- `data/` - база SQLite и служебные данные.
- `archive/` - архив закрытых сезонов.

## Если возникла ошибка

1. Посмотрите последнюю строку в журнале внизу окна.
2. Откройте `logs/seasonal_price.log`.
3. Проверьте формат входного файла и названия листов.
"""
        )
        layout.addWidget(browser, 1)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("PrimaryButton")
        btn_close.setMinimumHeight(38)
        btn_close.clicked.connect(self.accept)  # type: ignore[arg-type]
        layout.addWidget(btn_close)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1320, 860)
        self.setMinimumSize(980, 680)

        self._pool = QThreadPool.globalInstance()
        self._busy = False
        self._workers: list[Worker] = []
        self._seen_output_paths: set[str] = set()
        self._action_buttons: list[QPushButton] = []
        self._engine_instance: Any | None = None
        self._engine_base_dir: Path | None = None

        self._build_ui()
        self._apply_styles()
        self._sync_paths_from_base(force=True)
        self._update_duplicate_map_visibility()
        self._log("Интерфейс готов к работе.")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")

        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(18, 18, 18, 18)
        root_layout.setSpacing(12)

        root_layout.addWidget(self._build_header())

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(8)
        splitter.addWidget(self._build_sidebar())
        splitter.addWidget(self._build_workspace())
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 5)
        splitter.setSizes([360, 940])
        root_layout.addWidget(splitter, 1)

        self.setCentralWidget(root)

    def _build_header(self) -> QWidget:
        card = QFrame()
        card.setObjectName("HeaderCard")

        layout = QHBoxLayout(card)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        text_box = QVBoxLayout()
        text_box.setSpacing(2)

        title = QLabel(APP_TITLE)
        title.setObjectName("AppTitle")
        subtitle = QLabel(
            "Пошаговая обработка склада, заказов и подтверждений без ручной рутины.")
        subtitle.setObjectName("AppSubtitle")

        text_box.addWidget(title)
        text_box.addWidget(subtitle)

        btn_help = QPushButton("Как работать")
        btn_help.setObjectName("SecondaryButton")
        btn_help.setMinimumHeight(36)
        btn_help.clicked.connect(self._show_help)  # type: ignore[arg-type]

        btn_about = QPushButton("О программе")
        btn_about.setObjectName("SecondaryButton")
        btn_about.setMinimumHeight(36)
        btn_about.clicked.connect(self._show_about)  # type: ignore[arg-type]

        layout.addLayout(text_box, 1)
        layout.addWidget(btn_help)
        layout.addWidget(btn_about)
        return card

    def _build_sidebar(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        settings_card = self._create_card("1. Параметры")
        settings_layout = settings_card.layout()
        assert isinstance(settings_layout, QVBoxLayout)

        self.base_dir_input = PathField(
            "Рабочая папка проекта",
            str(Path.cwd()),
            pick_mode="dir",
            hint="Здесь программа хранит базу, логи и все результаты.",
            placeholder="Например: C:\\Work\\seasonal-price",
        )
        settings_layout.addWidget(self.base_dir_input)

        season_label = QLabel("Сезон")
        season_label.setObjectName("FieldLabel")
        self.season_edit = QLineEdit("spring_2026")
        self.season_edit.setMinimumHeight(40)
        self.season_edit.setPlaceholderText("Например: spring_2026")

        profile_label = QLabel("Профиль")
        profile_label.setObjectName("FieldLabel")
        self.profile_edit = QLineEdit("default")
        self.profile_edit.setMinimumHeight(40)
        self.profile_edit.setPlaceholderText("Например: retail или wholesale")

        settings_layout.addWidget(season_label)
        settings_layout.addWidget(self.season_edit)
        settings_layout.addWidget(profile_label)
        settings_layout.addWidget(self.profile_edit)

        ids_hint = QLabel(
            "Сезон и профиль используются в именах отчетов и в базе.")
        ids_hint.setObjectName("HintLabel")
        ids_hint.setWordWrap(True)
        settings_layout.addWidget(ids_hint)

        btn_sync = QPushButton("Подставить стандартные пути")
        btn_sync.setObjectName("SecondaryButton")
        btn_sync.setMinimumHeight(38)
        btn_sync.clicked.connect(lambda: self._sync_paths_from_base(
            force=True))  # type: ignore[arg-type]
        settings_layout.addWidget(btn_sync)

        scenario_card = self._create_card("2. Сценарий")
        scenario_layout = scenario_card.layout()
        assert isinstance(scenario_layout, QVBoxLayout)
        scenario_text = QLabel(
            "1) Подготовить сезон\n"
            "2) Прайс из склада\n"
            "3) Импорт заказов\n"
            "4) Аллокация\n"
            "5) Подтверждения\n"
            "6) Остаточный прайс"
        )
        scenario_text.setObjectName("ScenarioLabel")
        scenario_text.setWordWrap(True)
        scenario_layout.addWidget(scenario_text)

        quick_card = self._create_card("3. Быстрый запуск")
        quick_layout = quick_card.layout()
        assert isinstance(quick_layout, QVBoxLayout)

        self.btn_init_season = QPushButton("Подготовить сезон")
        self.btn_init_season.setObjectName("PrimaryButton")
        self.btn_init_season.setMinimumHeight(40)
        self.btn_init_season.clicked.connect(
            self._action_init_season)  # type: ignore[arg-type]
        self._action_buttons.append(self.btn_init_season)

        self.btn_full_cycle = QPushButton("Запустить полный цикл")
        self.btn_full_cycle.setObjectName("PrimaryButton")
        self.btn_full_cycle.setMinimumHeight(40)
        self.btn_full_cycle.clicked.connect(
            self._action_full_cycle)  # type: ignore[arg-type]
        self._action_buttons.append(self.btn_full_cycle)

        hint = QLabel(
            "Полный цикл выполнит шаги: прайс -> импорт -> аллокация -> подтверждения -> остатки.")
        hint.setObjectName("HintLabel")
        hint.setWordWrap(True)

        quick_layout.addWidget(self.btn_init_season)
        quick_layout.addWidget(self.btn_full_cycle)
        quick_layout.addWidget(hint)

        layout.addWidget(settings_card)
        layout.addWidget(scenario_card)
        layout.addWidget(quick_card)
        layout.addStretch(1)
        return panel

    def _build_workspace(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        operations_card = self._create_card("Операции")
        operations_layout = operations_card.layout()
        assert isinstance(operations_layout, QVBoxLayout)

        intro = QLabel(
            "Работайте по вкладкам слева направо. На каждой вкладке есть обязательные поля и одна кнопка запуска."
        )
        intro.setObjectName("IntroLabel")
        intro.setWordWrap(True)
        operations_layout.addWidget(intro)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_price_tab(), "1. Прайс")
        self.tabs.addTab(self._build_import_tab(), "2. Заказы")
        self.tabs.addTab(self._build_allocation_tab(), "3. Аллокация")
        self.tabs.addTab(self._build_confirmations_tab(), "4. Подтверждения")
        self.tabs.addTab(self._build_residual_tab(), "5. Остатки")
        self.tabs.addTab(self._build_archive_tab(), "6. Сезон")
        operations_layout.addWidget(self.tabs, 1)

        monitor_card = self._create_card("Мониторинг")
        monitor_layout = monitor_card.layout()
        assert isinstance(monitor_layout, QVBoxLayout)

        self.state_label = QLabel("Ожидание запуска")
        self.state_label.setObjectName("StateLabel")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)

        outputs_title = QLabel("Результаты")
        outputs_title.setObjectName("FieldLabel")
        self.outputs_list = QListWidget()
        self.outputs_list.setMinimumHeight(120)
        self.outputs_list.itemDoubleClicked.connect(
            self._open_selected_output)  # type: ignore[arg-type]

        btn_open_output = QPushButton("Открыть выбранный путь")
        btn_open_output.setObjectName("SecondaryButton")
        btn_open_output.setMinimumHeight(36)
        btn_open_output.clicked.connect(
            self._open_selected_output)  # type: ignore[arg-type]

        log_title = QLabel("Журнал")
        log_title.setObjectName("FieldLabel")
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(170)

        monitor_layout.addWidget(self.state_label)
        monitor_layout.addWidget(self.progress)
        monitor_layout.addWidget(outputs_title)
        monitor_layout.addWidget(self.outputs_list)
        monitor_layout.addWidget(btn_open_output)
        monitor_layout.addWidget(log_title)
        monitor_layout.addWidget(self.log_text, 1)

        layout.addWidget(operations_card, 3)
        layout.addWidget(monitor_card, 2)
        return panel

    def _build_price_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        tip = QLabel(
            "Шаг 1: выберите файл склада и папку, куда сохранить клиентский прайс.")
        tip.setObjectName("HintLabel")
        tip.setWordWrap(True)

        self.stock_file_input = PathField(
            "Файл склада (.xlsx/.xls)",
            pick_mode="file",
            file_filter="Excel (*.xlsx *.xls);;Все файлы (*.*)",
            hint="В файле должен быть лист 'Склад' или лист с эквивалентной структурой.",
            placeholder="Например: C:\\Data\\остатки.xlsx",
        )
        self.price_output_input = PathField(
            "Папка для клиентского прайса",
            pick_mode="dir",
            placeholder="Например: C:\\Data\\outputs",
        )

        btn = QPushButton("Сформировать клиентский прайс")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(40)
        # type: ignore[arg-type]
        btn.clicked.connect(self._action_generate_price)
        self._action_buttons.append(btn)

        layout.addWidget(tip)
        layout.addWidget(self.stock_file_input)
        layout.addWidget(self.price_output_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return self._wrap_tab_scroll(tab)

    def _build_import_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        tip = QLabel(
            "Шаг 2: укажите папку с заказами и как обрабатывать дубли клиентов.")
        tip.setObjectName("HintLabel")
        tip.setWordWrap(True)

        self.orders_dir_input = PathField(
            "Папка с заказами",
            pick_mode="dir",
            hint="Программа прочитает все .xls/.xlsx файлы в папке.",
            placeholder="Например: C:\\Data\\orders",
        )

        dup_label = QLabel("Стратегия дублей")
        dup_label.setObjectName("FieldLabel")
        self.duplicate_strategy_combo = QComboBox()
        self.duplicate_strategy_combo.addItem(
            "Последний файл клиента", "latest")
        self.duplicate_strategy_combo.addItem("Суммировать все дубли", "sum")
        self.duplicate_strategy_combo.addItem(
            "Ручной выбор через JSON", "manual")
        self.duplicate_strategy_combo.setMinimumHeight(40)
        self.duplicate_strategy_combo.currentIndexChanged.connect(
            self._update_duplicate_map_visibility)  # type: ignore[arg-type]

        self.duplicate_map_input = PathField(
            "JSON карта дублей",
            pick_mode="file",
            file_filter="JSON (*.json);;Все файлы (*.*)",
            hint="Нужно только для режима 'Ручной выбор через JSON'.",
            placeholder="Например: C:\\Data\\duplicate_map.json",
        )

        btn = QPushButton("Импортировать заказы")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(40)
        # type: ignore[arg-type]
        btn.clicked.connect(self._action_import_orders)
        self._action_buttons.append(btn)

        layout.addWidget(tip)
        layout.addWidget(self.orders_dir_input)
        layout.addWidget(dup_label)
        layout.addWidget(self.duplicate_strategy_combo)
        layout.addWidget(self.duplicate_map_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return self._wrap_tab_scroll(tab)

    def _build_allocation_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        tip = QLabel(
            "Шаг 3: выберите режим распределения подтверждений по остаткам.")
        tip.setObjectName("HintLabel")
        tip.setWordWrap(True)

        mode_label = QLabel("Режим аллокации")
        mode_label.setObjectName("FieldLabel")

        self.allocation_mode_combo = QComboBox()
        self.allocation_mode_combo.addItem("FIFO (по времени заказа)", "fifo")
        self.allocation_mode_combo.addItem("Пропорционально", "proportional")
        self.allocation_mode_combo.setMinimumHeight(40)

        mode_hint = QLabel(
            "FIFO приоритетнее для строгой очереди заказов, пропорционально - для равномерного распределения."
        )
        mode_hint.setObjectName("HintLabel")
        mode_hint.setWordWrap(True)

        btn = QPushButton("Выполнить аллокацию")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(40)
        btn.clicked.connect(self._action_allocate)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(tip)
        layout.addWidget(mode_label)
        layout.addWidget(self.allocation_mode_combo)
        layout.addWidget(mode_hint)
        layout.addStretch(1)
        layout.addWidget(btn)
        return self._wrap_tab_scroll(tab)

    def _build_confirmations_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        tip = QLabel(
            "Шаг 4: сформируйте клиентские подтверждения (XLSX + PDF).")
        tip.setObjectName("HintLabel")
        tip.setWordWrap(True)

        self.confirm_output_input = PathField(
            "Папка подтверждений",
            pick_mode="dir",
            placeholder="Например: C:\\Data\\outputs",
        )
        pdf_hint = QLabel("PDF формируется встроенным рендером.")
        pdf_hint.setObjectName("HintLabel")
        pdf_hint.setWordWrap(True)

        btn = QPushButton("Сформировать подтверждения")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(40)
        # type: ignore[arg-type]
        btn.clicked.connect(self._action_export_confirmations)
        self._action_buttons.append(btn)

        layout.addWidget(tip)
        layout.addWidget(self.confirm_output_input)
        layout.addWidget(pdf_hint)
        layout.addStretch(1)
        layout.addWidget(btn)
        return self._wrap_tab_scroll(tab)

    def _build_residual_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        tip = QLabel(
            "Шаг 5: сформируйте остаточный прайс после подтверждений.")
        tip.setObjectName("HintLabel")
        tip.setWordWrap(True)

        self.residual_output_input = PathField(
            "Папка остаточного прайса",
            pick_mode="dir",
            placeholder="Например: C:\\Data\\outputs",
        )

        btn = QPushButton("Сформировать остаточный прайс")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(40)
        # type: ignore[arg-type]
        btn.clicked.connect(self._action_build_residual)
        self._action_buttons.append(btn)

        layout.addWidget(tip)
        layout.addWidget(self.residual_output_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return self._wrap_tab_scroll(tab)

    def _build_archive_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        warning = QLabel(
            "Закрытие сезона переносит артефакты в архив и открывает новый сезон.")
        warning.setObjectName("WarningLabel")
        warning.setWordWrap(True)

        self.archive_dir_input = PathField(
            "Папка архива",
            pick_mode="dir",
            placeholder="Например: C:\\Data\\archive",
        )

        btn = QPushButton("Закрыть текущий сезон")
        btn.setObjectName("DangerButton")
        btn.setMinimumHeight(40)
        # type: ignore[arg-type]
        btn.clicked.connect(self._action_close_season)
        self._action_buttons.append(btn)

        layout.addWidget(warning)
        layout.addWidget(self.archive_dir_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return self._wrap_tab_scroll(tab)

    def _wrap_tab_scroll(self, inner: QWidget) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setWidget(inner)
        return scroll

    def _create_card(self, title: str) -> QFrame:
        card = QFrame()
        card.setObjectName("Card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(10)

        label = QLabel(title)
        label.setObjectName("CardTitle")
        layout.addWidget(label)
        return card

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            * {
                font-family: "Segoe UI Variable Text", "Segoe UI", sans-serif;
                font-size: 13px;
                color: #0f172a;
            }
            QWidget#AppRoot, QMainWindow {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #f4f7fb,
                    stop:1 #eef2f9
                );
            }
            QFrame#HeaderCard, QFrame#Card {
                background: #ffffff;
                border: 1px solid #dbe3ef;
                border-radius: 16px;
            }
            QLabel#AppTitle {
                font-size: 24px;
                font-weight: 700;
                color: #0f172a;
            }
            QLabel#AppSubtitle {
                color: #5b6a80;
                font-size: 13px;
            }
            QLabel#CardTitle {
                font-size: 14px;
                font-weight: 700;
                color: #0f172a;
            }
            QLabel#FieldLabel {
                color: #334155;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#HintLabel {
                color: #64748b;
                font-size: 12px;
            }
            QLabel#StateLabel {
                color: #0f172a;
                font-size: 13px;
                font-weight: 600;
            }
            QLabel#IntroLabel {
                color: #1e293b;
                background: #eff6ff;
                border: 1px solid #bfdbfe;
                border-radius: 12px;
                padding: 10px;
            }
            QLabel#ScenarioLabel {
                color: #334155;
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                padding: 10px;
                line-height: 1.35em;
            }
            QLabel#WarningLabel {
                color: #854d0e;
                background: #fffbeb;
                border: 1px solid #fde68a;
                border-radius: 12px;
                padding: 10px;
                font-weight: 600;
                font-size: 12px;
            }
            QLineEdit, QComboBox, QListWidget, QPlainTextEdit, QTextBrowser {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #d1d9e6;
                border-radius: 11px;
                selection-background-color: #dbeafe;
                selection-color: #0f172a;
            }
            QLineEdit, QComboBox {
                padding: 8px 10px;
            }
            QLineEdit:focus, QComboBox:focus, QListWidget:focus, QPlainTextEdit:focus, QTextBrowser:focus {
                border: 1px solid #93c5fd;
            }
            QLineEdit::placeholder {
                color: #94a3b8;
            }
            QComboBox::drop-down {
                border: none;
                width: 26px;
            }
            QComboBox QAbstractItemView {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #d1d9e6;
                selection-background-color: #dbeafe;
                selection-color: #0f172a;
            }
            QListWidget {
                padding: 4px;
            }
            QListWidget::item {
                color: #0f172a;
                padding: 6px 8px;
                border-radius: 8px;
            }
            QListWidget::item:selected {
                background: #e2e8f0;
                color: #0f172a;
            }
            QPlainTextEdit, QTextBrowser {
                background: #f8fafc;
                color: #0f172a;
                padding: 8px;
                font-family: "Cascadia Mono", "Consolas", monospace;
            }
            QTabWidget::pane {
                border: 1px solid #dbe3ef;
                border-radius: 13px;
                background: #ffffff;
                top: -1px;
                margin-top: 8px;
            }
            QTabBar::tab {
                background: transparent;
                border: 1px solid transparent;
                border-radius: 9px;
                color: #64748b;
                padding: 8px 12px;
                margin-right: 4px;
                font-weight: 600;
            }
            QTabBar::tab:hover {
                background: #f8fafc;
                color: #334155;
            }
            QTabBar::tab:selected {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #dbe3ef;
            }
            QPushButton {
                min-height: 36px;
            }
            QPushButton#PrimaryButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                border-radius: 11px;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton#PrimaryButton:hover {
                background: #1d4ed8;
                border: 1px solid #1d4ed8;
            }
            QPushButton#PrimaryButton:pressed {
                background: #1e40af;
                border: 1px solid #1e40af;
            }
            QPushButton#SecondaryButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #cfd8e6;
                border-radius: 11px;
                padding: 8px 12px;
                font-weight: 600;
            }
            QPushButton#SecondaryButton:hover {
                background: #f8fafc;
                border: 1px solid #93a6c2;
            }
            QPushButton#DangerButton {
                background: #dc2626;
                color: #ffffff;
                border: 1px solid #dc2626;
                border-radius: 11px;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton#DangerButton:hover {
                background: #b91c1c;
                border: 1px solid #b91c1c;
            }
            QPushButton:disabled {
                background: #e5e7eb;
                color: #94a3b8;
                border: 1px solid #e5e7eb;
            }
            QProgressBar {
                background: #e5e7eb;
                border: none;
                border-radius: 7px;
                min-height: 14px;
                max-height: 14px;
            }
            QProgressBar::chunk {
                background: #2563eb;
                border-radius: 7px;
            }
            """
        )

    def _show_help(self) -> None:
        HelpDialog(self).exec()

    def _show_about(self) -> None:
        QMessageBox.information(
            self,
            "О программе",
            (
                f"{APP_TITLE}\n\n"
                "GUI-приложение для пакетной обработки прайсов и заказов.\n"
                f"Версия: {APP_VERSION}"
            ),
        )

    def _sync_paths_from_base(self, force: bool = False) -> None:
        try:
            base = self._base_dir(create=False)
        except Exception:
            return

        data_dir = base / "data"
        out_dir = base / "outputs"
        archive_dir = base / "archive"
        orders_dir = base / "orders"

        self._set_path(self.stock_file_input, str(
            data_dir / "stock.xlsx"), force=force)
        self._set_path(self.price_output_input, str(out_dir), force=force)
        self._set_path(self.orders_dir_input, str(orders_dir), force=force)
        self._set_path(self.duplicate_map_input, str(
            base / "duplicate_map.json"), force=force)
        self._set_path(self.confirm_output_input, str(out_dir), force=force)
        self._set_path(self.residual_output_input, str(out_dir), force=force)
        self._set_path(self.archive_dir_input, str(archive_dir), force=force)

    @staticmethod
    def _set_path(field: PathField, value: str, force: bool = False) -> None:
        if force or not field.text():
            field.set_text(value)

    def _set_busy(self, busy: bool, state_text: str, determinate_progress: bool = False) -> None:
        self._busy = busy
        self.state_label.setText(state_text)
        self.tabs.setEnabled(not busy)
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
            QMessageBox.information(
                self, "Операция уже выполняется", "Дождитесь завершения текущей операции.")
            return

        self._set_busy(
            True, f"Выполняется: {title}", determinate_progress=with_progress)
        self._log(f"> {title}")

        worker = Worker(task, with_progress=with_progress)
        self._workers.append(worker)
        worker.signals.finished.connect(lambda result: self._handle_success(
            title, result, worker))  # type: ignore[arg-type]
        worker.signals.failed.connect(lambda error: self._handle_error(
            title, error, worker))  # type: ignore[arg-type]
        if with_progress:
            worker.signals.progress.connect(
                self._handle_progress)  # type: ignore[arg-type]
        self._pool.start(worker)

    def _handle_success(self, title: str, result: Any, worker: Worker) -> None:
        self._release_worker(worker)
        self._set_busy(False, f"Готово: {title}")
        self._log(f"+ {title}: успешно")
        if isinstance(result, dict):
            self._register_outputs(result)

    def _handle_error(self, title: str, error_text: str, worker: Worker) -> None:
        self._release_worker(worker)
        self._set_busy(False, f"Ошибка: {title}")
        self._log(f"! {title}: ошибка")
        self._log(error_text.strip())
        short_text = error_text.strip().splitlines(
        )[-1] if error_text.strip() else "Неизвестная ошибка"
        QMessageBox.critical(self, "Ошибка операции", short_text)

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

    def _register_outputs(self, payload: Any, prefix: str = "") -> None:
        if isinstance(payload, dict):
            for key, value in payload.items():
                self._register_outputs(value, prefix=f"{prefix}{key}.")
            return
        if not isinstance(payload, str):
            return

        path = Path(payload)
        if not path.exists():
            return

        resolved = str(path.resolve())
        if resolved in self._seen_output_paths:
            return
        self._seen_output_paths.add(resolved)

        label = prefix[:-1] if prefix else "Файл"
        item = QListWidgetItem(f"{label}: {resolved}")
        item.setData(Qt.UserRole, resolved)
        self.outputs_list.addItem(item)

    def _open_selected_output(self) -> None:
        item = self.outputs_list.currentItem()
        if item is None:
            return
        path_text = item.data(Qt.UserRole)
        if not path_text:
            return

        path = Path(str(path_text))
        if not path.exists():
            QMessageBox.warning(self, "Путь не найден",
                                f"Путь отсутствует:\n{path}")
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(path)))

    def _log(self, message: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.appendPlainText(f"[{ts}] {message}")

    def _base_dir(self, create: bool = True) -> Path:
        text = self.base_dir_input.text() or str(Path.cwd())
        path = Path(text).expanduser()
        if create:
            path.mkdir(parents=True, exist_ok=True)
        return path

    def _season_id(self) -> str:
        value = self.season_edit.text().strip()
        if not value:
            raise ValueError("Поле «Сезон» не может быть пустым.")
        return value

    def _profile_id(self) -> str:
        value = self.profile_edit.text().strip()
        if not value:
            raise ValueError("Поле «Профиль» не может быть пустым.")
        return value

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
            raise ValueError(
                f"Неверный формат файла ({path.name}). Ожидается: {suffixes}")
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

    def _optional_duplicate_map(self, strategy: str) -> Path | None:
        text = self.duplicate_map_input.text()
        if strategy != "manual":
            return None
        if not text:
            raise ValueError("Для ручного режима дублей укажите JSON карту.")
        return self._ensure_file_exists(text, "JSON карта дублей", {".json"})

    def _show_validation_error(self, message: str) -> None:
        QMessageBox.warning(self, "Проверьте входные данные", message)

    def _update_duplicate_map_visibility(self) -> None:
        strategy = str(self.duplicate_strategy_combo.currentData())
        self.duplicate_map_input.setVisible(strategy == "manual")

    def _action_init_season(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.init_season(season)

        self._run_task("Инициализация сезона", task)

    def _action_generate_price(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
            stock_file = self._ensure_file_exists(
                self.stock_file_input.text(),
                "Файл склада",
                {".xlsx", ".xls"},
            )
            output_dir = self._ensure_dir(
                self.price_output_input.text(), "Папка для прайса")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.generate_price(
                stock_file=stock_file,
                season_id=season,
                output_dir=output_dir,
            )

        self._run_task("Формирование прайса", task)

    def _action_import_orders(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
            profile = self._profile_id()
            orders_dir = self._ensure_dir_exists(
                self.orders_dir_input.text(), "Папка с заказами")
            duplicate_strategy = str(
                self.duplicate_strategy_combo.currentData())
            duplicate_map = self._optional_duplicate_map(duplicate_strategy)
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.import_orders(
                input_dir=orders_dir,
                season_id=season,
                profile_id=profile,
                duplicate_strategy=duplicate_strategy,
                duplicate_map_path=duplicate_map,
                progress_callback=_progress,
            )

        self._run_task("Импорт заказов", task, with_progress=True)

    def _action_allocate(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
            profile = self._profile_id()
            mode = str(self.allocation_mode_combo.currentData())
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.allocate(
                season_id=season,
                mode=mode,
                profile_id=profile,
            )

        self._run_task("Аллокация", task)

    def _action_export_confirmations(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
            output_dir = self._ensure_dir(
                self.confirm_output_input.text(), "Папка подтверждений")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.export_confirmations(
                season_id=season,
                output_dir=output_dir,
            )

        self._run_task("Экспорт подтверждений", task)

    def _action_build_residual(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
            output_dir = self._ensure_dir(
                self.residual_output_input.text(), "Папка остаточного прайса")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.build_residual_price(
                season_id=season,
                output_dir=output_dir,
            )

        self._run_task("Формирование остаточного прайса", task)

    def _action_close_season(self) -> None:
        answer = QMessageBox.question(
            self,
            "Подтвердите действие",
            "Закрыть текущий сезон и создать новый?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if answer != QMessageBox.Yes:
            return

        try:
            base_dir = self._base_dir()
            season = self._season_id()
            archive_dir = self._ensure_dir(
                self.archive_dir_input.text(), "Папка архива")
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)
            return engine.close_season(
                season_id=season,
                archive_dir=archive_dir,
            )

        self._run_task("Закрытие сезона", task)

    def _action_full_cycle(self) -> None:
        try:
            base_dir = self._base_dir()
            season = self._season_id()
            profile = self._profile_id()
            stock_file = self._ensure_file_exists(
                self.stock_file_input.text(),
                "Файл склада",
                {".xlsx", ".xls"},
            )
            price_output = self._ensure_dir(
                self.price_output_input.text(), "Папка для прайса")
            orders_dir = self._ensure_dir_exists(
                self.orders_dir_input.text(), "Папка с заказами")
            duplicate_strategy = str(
                self.duplicate_strategy_combo.currentData())
            duplicate_map = self._optional_duplicate_map(duplicate_strategy)
            confirm_output = self._ensure_dir(
                self.confirm_output_input.text(), "Папка подтверждений")
            residual_output = self._ensure_dir(
                self.residual_output_input.text(), "Папка остаточного прайса")
            allocation_mode = str(self.allocation_mode_combo.currentData())
        except ValueError as exc:
            self._show_validation_error(str(exc))
            return

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = self._engine_for_base(base_dir)

            result: dict[str, Any] = {}
            if _progress is not None:
                _progress(
                    {"percent": 5, "message": "Полный цикл: инициализация сезона"})
            result["init"] = engine.init_season(season)
            if _progress is not None:
                _progress(
                    {"percent": 15, "message": "Полный цикл: формирование клиентского прайса"})
            result["price"] = engine.generate_price(
                stock_file=stock_file,
                season_id=season,
                output_dir=price_output,
            )

            def _import_progress(payload: dict[str, Any]) -> None:
                if _progress is None:
                    return
                import_percent = int(payload.get("percent", 0))
                mapped_percent = 20 + int(import_percent * 0.5)
                _progress(
                    {
                        "percent": mapped_percent,
                        "message": f"Полный цикл: {payload.get('message', 'импорт заказов')}",
                    }
                )

            result["import"] = engine.import_orders(
                input_dir=orders_dir,
                season_id=season,
                profile_id=profile,
                duplicate_strategy=duplicate_strategy,
                duplicate_map_path=duplicate_map,
                progress_callback=_import_progress,
            )
            if _progress is not None:
                _progress({"percent": 75, "message": "Полный цикл: аллокация"})
            result["allocate"] = engine.allocate(
                season_id=season,
                mode=allocation_mode,
                profile_id=profile,
            )
            if _progress is not None:
                _progress(
                    {"percent": 88, "message": "Полный цикл: экспорт подтверждений"})
            result["confirmations"] = engine.export_confirmations(
                season_id=season,
                output_dir=confirm_output,
            )
            if _progress is not None:
                _progress(
                    {"percent": 96, "message": "Полный цикл: построение остаточного прайса"})
            result["residual"] = engine.build_residual_price(
                season_id=season,
                output_dir=residual_output,
            )
            if _progress is not None:
                _progress({"percent": 100, "message": "Полный цикл завершен"})
            return result

        self._run_task("Полный цикл", task, with_progress=True)


def run_gui() -> None:
    if QApplication is None:
        raise RuntimeError(
            "PySide6 не установлен. Установите зависимости: python -m pip install -e .[gui]")

    app = QApplication([])
    app.setStyle("Fusion")
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


if __name__ == "__main__":
    run_gui()
