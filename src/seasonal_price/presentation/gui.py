from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from pathlib import Path
import traceback
from typing import Any

from seasonal_price.application.bootstrap import build_engine

APP_TITLE = "Сезонный Прайс"
APP_EXE_NAME = "SeasonalPrice"

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

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(8)

        self.edit = QLineEdit(default_text)
        self.edit.setMinimumHeight(38)

        btn = QPushButton("Обзор")
        btn.setObjectName("SecondaryButton")
        btn.setMinimumHeight(38)
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
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите файл", filter=self._file_filter)
        if file_path:
            self.edit.setText(file_path)


class HelpDialog(QDialog):
    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Справка - {APP_TITLE}")
        self.resize(880, 680)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setMarkdown(
            """
# Справка оператора

## Базовый сценарий

1. Инициализируйте сезон.
2. Сформируйте клиентский прайс из листа **Склад**.
3. Импортируйте папку заказов.
4. Выполните аллокацию.
5. Сформируйте подтверждения клиентам.
6. Постройте остаточный прайс.
7. Закройте сезон в архив.

## Дубли заказов

- `latest` - берется последний файл клиента по дате изменения.
- `sum` - заказы из всех дублей клиента суммируются.
- `manual` - используется JSON-карта выбора версии файла.

## Кратность

Кратность применяется автоматически в ядре. Скорректированные позиции помечаются в отчетах.

## Где результаты

- База и служебные данные: `data/`
- Логи: `logs/`
- Выходные файлы: `outputs/`
- Архивы сезонов: `archive/`
"""
        )
        layout.addWidget(browser, 1)

        btn_close = QPushButton("Закрыть")
        btn_close.setObjectName("PrimaryButton")
        btn_close.setMinimumHeight(36)
        btn_close.clicked.connect(self.accept)  # type: ignore[arg-type]
        layout.addWidget(btn_close)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1280, 840)
        self.setMinimumSize(1120, 760)

        self._pool = QThreadPool.globalInstance()
        self._busy = False
        self._workers: list[Worker] = []
        self._seen_output_paths: set[str] = set()
        self._action_buttons: list[QPushButton] = []

        self._build_ui()
        self._apply_styles()
        self._sync_paths_from_base(force=True)
        self._log("Интерфейс готов к работе.")

    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("AppRoot")

        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(16, 16, 16, 16)
        root_layout.setSpacing(12)

        root_layout.addWidget(self._build_header())

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._build_sidebar())
        splitter.addWidget(self._build_workspace())
        splitter.setSizes([360, 920])
        root_layout.addWidget(splitter, 1)

        self.setCentralWidget(root)

    def _build_header(self) -> QWidget:
        card = QFrame()
        card.setObjectName("HeaderCard")

        layout = QHBoxLayout(card)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(10)

        text_box = QVBoxLayout()
        text_box.setSpacing(2)

        title = QLabel(APP_TITLE)
        title.setObjectName("AppTitle")
        subtitle = QLabel("Пакетная обработка прайсов, заказов, подтверждений и остатков.")
        subtitle.setObjectName("AppSubtitle")

        text_box.addWidget(title)
        text_box.addWidget(subtitle)

        btn_help = QPushButton("Справка")
        btn_help.setObjectName("SecondaryButton")
        btn_help.setMinimumHeight(34)
        btn_help.clicked.connect(self._show_help)  # type: ignore[arg-type]

        btn_about = QPushButton("О программе")
        btn_about.setObjectName("SecondaryButton")
        btn_about.setMinimumHeight(34)
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

        settings_card = self._create_card("Параметры")
        settings_layout = settings_card.layout()
        assert isinstance(settings_layout, QVBoxLayout)

        self.base_dir_input = PathField("Рабочая папка проекта", str(Path.cwd()), pick_mode="dir")
        settings_layout.addWidget(self.base_dir_input)

        season_label = QLabel("Сезон")
        season_label.setObjectName("FieldLabel")
        self.season_edit = QLineEdit("spring_2026")
        self.season_edit.setMinimumHeight(38)

        profile_label = QLabel("Профиль")
        profile_label.setObjectName("FieldLabel")
        self.profile_edit = QLineEdit("default")
        self.profile_edit.setMinimumHeight(38)

        settings_layout.addWidget(season_label)
        settings_layout.addWidget(self.season_edit)
        settings_layout.addWidget(profile_label)
        settings_layout.addWidget(self.profile_edit)

        btn_sync = QPushButton("Синхронизировать пути")
        btn_sync.setObjectName("SecondaryButton")
        btn_sync.setMinimumHeight(36)
        btn_sync.clicked.connect(lambda: self._sync_paths_from_base(force=True))  # type: ignore[arg-type]
        settings_layout.addWidget(btn_sync)

        quick_card = self._create_card("Быстрые действия")
        quick_layout = quick_card.layout()
        assert isinstance(quick_layout, QVBoxLayout)

        self.btn_init_season = QPushButton("Инициализировать сезон")
        self.btn_init_season.setObjectName("PrimaryButton")
        self.btn_init_season.setMinimumHeight(38)
        self.btn_init_season.clicked.connect(self._action_init_season)  # type: ignore[arg-type]
        self._action_buttons.append(self.btn_init_season)

        self.btn_full_cycle = QPushButton("Запустить полный цикл")
        self.btn_full_cycle.setObjectName("PrimaryButton")
        self.btn_full_cycle.setMinimumHeight(38)
        self.btn_full_cycle.clicked.connect(self._action_full_cycle)  # type: ignore[arg-type]
        self._action_buttons.append(self.btn_full_cycle)

        hint = QLabel("Полный цикл: прайс -> импорт -> аллокация -> подтверждения -> остатки.")
        hint.setObjectName("MutedLabel")
        hint.setWordWrap(True)

        quick_layout.addWidget(self.btn_init_season)
        quick_layout.addWidget(self.btn_full_cycle)
        quick_layout.addWidget(hint)

        layout.addWidget(settings_card)
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

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_price_tab(), "Прайс")
        self.tabs.addTab(self._build_import_tab(), "Заказы")
        self.tabs.addTab(self._build_allocation_tab(), "Аллокация")
        self.tabs.addTab(self._build_confirmations_tab(), "Подтверждения")
        self.tabs.addTab(self._build_residual_tab(), "Остатки")
        self.tabs.addTab(self._build_archive_tab(), "Архив")
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

        outputs_title = QLabel("Сформированные пути")
        outputs_title.setObjectName("FieldLabel")
        self.outputs_list = QListWidget()
        self.outputs_list.setMinimumHeight(120)
        self.outputs_list.itemDoubleClicked.connect(self._open_selected_output)  # type: ignore[arg-type]

        btn_open_output = QPushButton("Открыть выбранный путь")
        btn_open_output.setObjectName("SecondaryButton")
        btn_open_output.setMinimumHeight(34)
        btn_open_output.clicked.connect(self._open_selected_output)  # type: ignore[arg-type]

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

        self.stock_file_input = PathField(
            "Файл склада (.xlsx/.xls)",
            pick_mode="file",
            file_filter="Excel (*.xlsx *.xls);;Все файлы (*.*)",
        )
        self.price_output_input = PathField("Папка для прайса", pick_mode="dir")

        btn = QPushButton("Сформировать прайс")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self._action_generate_price)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(self.stock_file_input)
        layout.addWidget(self.price_output_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return tab

    def _build_import_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        self.orders_dir_input = PathField("Папка с заказами", pick_mode="dir")

        dup_label = QLabel("Стратегия дублей")
        dup_label.setObjectName("FieldLabel")
        self.duplicate_strategy_combo = QComboBox()
        self.duplicate_strategy_combo.addItem("Последний файл (latest)", "latest")
        self.duplicate_strategy_combo.addItem("Суммировать дубли (sum)", "sum")
        self.duplicate_strategy_combo.addItem("Ручной выбор (manual)", "manual")
        self.duplicate_strategy_combo.setMinimumHeight(38)

        self.duplicate_map_input = PathField(
            "JSON карта дублей (для manual)",
            pick_mode="file",
            file_filter="JSON (*.json);;Все файлы (*.*)",
        )

        btn = QPushButton("Импортировать заказы")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self._action_import_orders)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(self.orders_dir_input)
        layout.addWidget(dup_label)
        layout.addWidget(self.duplicate_strategy_combo)
        layout.addWidget(self.duplicate_map_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return tab

    def _build_allocation_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        mode_label = QLabel("Режим аллокации")
        mode_label.setObjectName("FieldLabel")

        self.allocation_mode_combo = QComboBox()
        self.allocation_mode_combo.addItem("FIFO (по времени заказа)", "fifo")
        self.allocation_mode_combo.addItem("Пропорционально", "proportional")
        self.allocation_mode_combo.setMinimumHeight(38)

        btn = QPushButton("Выполнить аллокацию")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self._action_allocate)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(mode_label)
        layout.addWidget(self.allocation_mode_combo)
        layout.addStretch(1)
        layout.addWidget(btn)
        return tab

    def _build_confirmations_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        self.confirm_output_input = PathField("Папка подтверждений", pick_mode="dir")
        pdf_hint = QLabel("PDF формируется встроенным рендером (backend-режим отключен).")
        pdf_hint.setObjectName("MutedLabel")
        pdf_hint.setWordWrap(True)

        btn = QPushButton("Сформировать подтверждения")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self._action_export_confirmations)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(self.confirm_output_input)
        layout.addWidget(pdf_hint)
        layout.addStretch(1)
        layout.addWidget(btn)
        return tab

    def _build_residual_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        self.residual_output_input = PathField("Папка остаточного прайса", pick_mode="dir")

        btn = QPushButton("Сформировать остаточный прайс")
        btn.setObjectName("PrimaryButton")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self._action_build_residual)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(self.residual_output_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return tab

    def _build_archive_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        warning = QLabel("Закрытие сезона переносит артефакты в архив и открывает новый сезон.")
        warning.setObjectName("WarningLabel")
        warning.setWordWrap(True)

        self.archive_dir_input = PathField("Папка архива", pick_mode="dir")

        btn = QPushButton("Закрыть сезон")
        btn.setObjectName("DangerButton")
        btn.setMinimumHeight(38)
        btn.clicked.connect(self._action_close_season)  # type: ignore[arg-type]
        self._action_buttons.append(btn)

        layout.addWidget(warning)
        layout.addWidget(self.archive_dir_input)
        layout.addStretch(1)
        layout.addWidget(btn)
        return tab

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
            QWidget#AppRoot {
                background: #f6f7f9;
            }
            QMainWindow {
                background: #f6f7f9;
            }
            QFrame#HeaderCard, QFrame#Card {
                background: #ffffff;
                border: 1px solid #e5e7eb;
                border-radius: 14px;
            }
            QLabel#AppTitle {
                font-size: 22px;
                font-weight: 700;
                color: #0f172a;
            }
            QLabel#AppSubtitle {
                color: #64748b;
                font-size: 13px;
            }
            QLabel#CardTitle {
                font-size: 14px;
                font-weight: 700;
                color: #0f172a;
            }
            QLabel#FieldLabel {
                color: #475569;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#MutedLabel {
                color: #64748b;
                font-size: 12px;
            }
            QLabel#StateLabel {
                color: #0f172a;
                font-size: 13px;
                font-weight: 600;
            }
            QLabel#WarningLabel {
                color: #854d0e;
                background: #fffbeb;
                border: 1px solid #fde68a;
                border-radius: 10px;
                padding: 10px;
                font-weight: 600;
                font-size: 12px;
            }
            QLineEdit, QComboBox, QListWidget, QPlainTextEdit, QTextBrowser {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #d1d5db;
                border-radius: 10px;
                selection-background-color: #dbeafe;
                selection-color: #0f172a;
            }
            QLineEdit, QComboBox {
                padding: 7px 10px;
            }
            QLineEdit:focus, QComboBox:focus, QListWidget:focus, QPlainTextEdit:focus, QTextBrowser:focus {
                border: 1px solid #94a3b8;
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
                border: 1px solid #d1d5db;
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
                border: 1px solid #e5e7eb;
                border-radius: 12px;
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
                border: 1px solid #e5e7eb;
            }
            QPushButton {
                min-height: 34px;
            }
            QPushButton#PrimaryButton {
                background: #111827;
                color: #ffffff;
                border: 1px solid #111827;
                border-radius: 10px;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton#PrimaryButton:hover {
                background: #1f2937;
                border: 1px solid #1f2937;
            }
            QPushButton#PrimaryButton:pressed {
                background: #0f172a;
                border: 1px solid #0f172a;
            }
            QPushButton#SecondaryButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #d1d5db;
                border-radius: 10px;
                padding: 8px 12px;
                font-weight: 600;
            }
            QPushButton#SecondaryButton:hover {
                background: #f8fafc;
                border: 1px solid #94a3b8;
            }
            QPushButton#DangerButton {
                background: #b91c1c;
                color: #ffffff;
                border: 1px solid #b91c1c;
                border-radius: 10px;
                padding: 8px 14px;
                font-weight: 600;
            }
            QPushButton#DangerButton:hover {
                background: #991b1b;
                border: 1px solid #991b1b;
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
                background: #111827;
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
                "Версия: 0.1.0"
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

        self._set_path(self.stock_file_input, str(data_dir / "stock.xlsx"), force=force)
        self._set_path(self.price_output_input, str(out_dir), force=force)
        self._set_path(self.orders_dir_input, str(orders_dir), force=force)
        self._set_path(self.duplicate_map_input, str(base / "duplicate_map.json"), force=force)
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
            QMessageBox.information(self, "Операция уже выполняется", "Дождитесь завершения текущей операции.")
            return

        self._set_busy(True, f"Выполняется: {title}", determinate_progress=with_progress)
        self._log(f"> {title}")

        worker = Worker(task, with_progress=with_progress)
        self._workers.append(worker)
        worker.signals.finished.connect(lambda result: self._handle_success(title, result, worker))  # type: ignore[arg-type]
        worker.signals.failed.connect(lambda error: self._handle_error(title, error, worker))  # type: ignore[arg-type]
        if with_progress:
            worker.signals.progress.connect(self._handle_progress)  # type: ignore[arg-type]
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
        short_text = error_text.strip().splitlines()[-1] if error_text.strip() else "Неизвестная ошибка"
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
            QMessageBox.warning(self, "Путь не найден", f"Путь отсутствует:\n{path}")
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

    def _action_init_season(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            return engine.init_season(self._season_id())

        self._run_task("Инициализация сезона", task)

    def _action_generate_price(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            return engine.generate_price(
                stock_file=Path(self.stock_file_input.text()),
                season_id=self._season_id(),
                output_dir=Path(self.price_output_input.text()),
            )

        self._run_task("Формирование прайса", task)

    def _action_import_orders(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            duplicate_map = Path(self.duplicate_map_input.text()) if self.duplicate_map_input.text() else None
            return engine.import_orders(
                input_dir=Path(self.orders_dir_input.text()),
                season_id=self._season_id(),
                profile_id=self._profile_id(),
                duplicate_strategy=str(self.duplicate_strategy_combo.currentData()),
                duplicate_map_path=duplicate_map,
                progress_callback=_progress,
            )

        self._run_task("Импорт заказов", task, with_progress=True)

    def _action_allocate(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            return engine.allocate(
                season_id=self._season_id(),
                mode=str(self.allocation_mode_combo.currentData()),
                profile_id=self._profile_id(),
            )

        self._run_task("Аллокация", task)

    def _action_export_confirmations(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            return engine.export_confirmations(
                season_id=self._season_id(),
                output_dir=Path(self.confirm_output_input.text()),
            )

        self._run_task("Экспорт подтверждений", task)

    def _action_build_residual(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            return engine.build_residual_price(
                season_id=self._season_id(),
                output_dir=Path(self.residual_output_input.text()),
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

        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            return engine.close_season(
                season_id=self._season_id(),
                archive_dir=Path(self.archive_dir_input.text()),
            )

        self._run_task("Закрытие сезона", task)

    def _action_full_cycle(self) -> None:
        def task(_progress: Callable[[dict[str, Any]], None] | None = None) -> dict[str, Any]:
            engine = build_engine(self._base_dir())
            season = self._season_id()
            profile = self._profile_id()

            result: dict[str, Any] = {}
            if _progress is not None:
                _progress({"percent": 5, "message": "Полный цикл: инициализация сезона"})
            result["init"] = engine.init_season(season)
            if _progress is not None:
                _progress({"percent": 15, "message": "Полный цикл: формирование клиентского прайса"})
            result["price"] = engine.generate_price(
                stock_file=Path(self.stock_file_input.text()),
                season_id=season,
                output_dir=Path(self.price_output_input.text()),
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
                input_dir=Path(self.orders_dir_input.text()),
                season_id=season,
                profile_id=profile,
                duplicate_strategy=str(self.duplicate_strategy_combo.currentData()),
                duplicate_map_path=(Path(self.duplicate_map_input.text()) if self.duplicate_map_input.text() else None),
                progress_callback=_import_progress,
            )
            if _progress is not None:
                _progress({"percent": 75, "message": "Полный цикл: аллокация"})
            result["allocate"] = engine.allocate(
                season_id=season,
                mode=str(self.allocation_mode_combo.currentData()),
                profile_id=profile,
            )
            if _progress is not None:
                _progress({"percent": 88, "message": "Полный цикл: экспорт подтверждений"})
            result["confirmations"] = engine.export_confirmations(
                season_id=season,
                output_dir=Path(self.confirm_output_input.text()),
            )
            if _progress is not None:
                _progress({"percent": 96, "message": "Полный цикл: построение остаточного прайса"})
            result["residual"] = engine.build_residual_price(
                season_id=season,
                output_dir=Path(self.residual_output_input.text()),
            )
            if _progress is not None:
                _progress({"percent": 100, "message": "Полный цикл завершен"})
            return result

        self._run_task("Полный цикл", task, with_progress=True)


def run_gui() -> None:
    if QApplication is None:
        raise RuntimeError("PySide6 не установлен. Установите зависимости: python -m pip install -e .[gui]")

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
