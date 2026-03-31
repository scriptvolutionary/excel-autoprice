from __future__ import annotations

from contextlib import suppress
from pathlib import Path
import sys

from cx_Freeze import Executable, setup
from cx_Freeze.command.bdist_msi import bdist_msi as BaseBdistMsi


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

APP_NAME = "Сезонный Прайс"
APP_EXE = "SeasonalPrice.exe"
APP_VENDOR = "SeasonalPrice"
APP_VERSION = "0.1.0"
APP_DESCRIPTION = "GUI-приложение для сезонной обработки прайсов и заказов."
APP_UPGRADE_CODE = "{2C0C3A5E-3A22-4E2D-95AA-5F43E8DFA67A}"


class RussianBdistMsi(BaseBdistMsi):
    """Кастомизация MSI: русские тексты мастера + удобные значения по умолчанию."""

    def add_ui(self) -> None:
        super().add_ui()
        self._localize_controls()
        self._localize_radiobuttons()

    def add_properties(self) -> None:
        super().add_properties()
        self._update_property("Progress1", "Установка")
        self._update_property("Progress2", "устанавливает")

    def _update_sql(self, sql: str) -> None:
        with suppress(Exception):
            view = self.db.OpenView(sql)
            view.Execute(None)
            view.Close()

    @staticmethod
    def _escape(value: str) -> str:
        return value.replace("'", "''")

    def _update_control_text(self, dialog: str, control: str, text: str) -> None:
        dialog_escaped = self._escape(dialog)
        control_escaped = self._escape(control)
        text_escaped = self._escape(text)
        self._update_sql(
            "UPDATE `Control` "
            f"SET `Text`='{text_escaped}' "
            f"WHERE `Dialog_`='{dialog_escaped}' AND `Control`='{control_escaped}'"
        )

    def _update_property(self, key: str, value: str) -> None:
        key_escaped = self._escape(key)
        value_escaped = self._escape(value)
        self._update_sql(
            "UPDATE `Property` "
            f"SET `Value`='{value_escaped}' "
            f"WHERE `Property`='{key_escaped}'"
        )

    def _update_radiobutton_text(self, property_name: str, value: str, text: str) -> None:
        property_escaped = self._escape(property_name)
        value_escaped = self._escape(value)
        text_escaped = self._escape(text)
        self._update_sql(
            "UPDATE `RadioButton` "
            f"SET `Text`='{text_escaped}' "
            f"WHERE `Property`='{property_escaped}' AND `Value`='{value_escaped}'"
        )

    def _localize_controls(self) -> None:
        # Большая часть английского текста создается внутри bdist_msi;
        # здесь централизованно заменяем его на русский после построения UI-таблиц.
        translations = [
            ("CancelDlg", "Text", "Вы действительно хотите прервать установку [ProductName]?"),
            ("CancelDlg", "Yes", "Да"),
            ("CancelDlg", "No", "Нет"),
            ("ErrorDlg", "N", "Нет"),
            ("ErrorDlg", "Y", "Да"),
            ("ErrorDlg", "A", "Прервать"),
            ("ErrorDlg", "C", "Отмена"),
            ("ErrorDlg", "I", "Игнорировать"),
            ("ErrorDlg", "O", "ОК"),
            ("ErrorDlg", "R", "Повтор"),
            ("ExitDialog", "Title", "Завершение установки [ProductName]"),
            ("ExitDialog", "Back", "< Назад"),
            ("ExitDialog", "Cancel", "Отмена"),
            ("ExitDialog", "Description", "Нажмите «Готово», чтобы закрыть мастер установки."),
            ("ExitDialog", "Finish", "Готово"),
            ("ExitDialog", "LaunchOnFinish", "Запустить приложение после завершения установки"),
            ("FatalError", "Title", "Установка [ProductName] завершилась с ошибкой"),
            ("FatalError", "Back", "< Назад"),
            ("FatalError", "Cancel", "Отмена"),
            (
                "FatalError",
                "Description1",
                "[ProductName] не удалось установить из-за ошибки. "
                "Система не была изменена. Запустите установку повторно позже.",
            ),
            ("FatalError", "Description2", "Нажмите «Готово», чтобы закрыть мастер установки."),
            ("FatalError", "Finish", "Готово"),
            ("FilesInUse", "Title", r"{\DlgFontBold8}Файлы используются"),
            ("FilesInUse", "Description", "Некоторые файлы, которые нужно обновить, сейчас используются."),
            (
                "FilesInUse",
                "Text",
                "Следующие приложения используют файлы, которые должны быть обновлены. "
                "Закройте их и нажмите «Повторить», либо «Выход» для отмены установки.",
            ),
            ("FilesInUse", "Exit", "Выход"),
            ("FilesInUse", "Ignore", "Игнорировать"),
            ("FilesInUse", "Retry", "Повторить"),
            ("MaintenanceTypeDlg", "Title", "Добро пожаловать в мастер установки [ProductName]"),
            (
                "MaintenanceTypeDlg",
                "BodyText",
                "Выберите, что нужно сделать с [ProductName]: восстановить или удалить.",
            ),
            ("MaintenanceTypeDlg", "Back", "< Назад"),
            ("MaintenanceTypeDlg", "Finish", "Далее"),
            ("MaintenanceTypeDlg", "Cancel", "Отмена"),
            (
                "PrepareDlg",
                "Description",
                "Подождите, пока мастер установки подготовит все необходимые данные.",
            ),
            ("PrepareDlg", "Title", "Добро пожаловать в установщик [ProductName]"),
            ("PrepareDlg", "ActionText", "Подготовка..."),
            ("PrepareDlg", "Back", "Назад"),
            ("PrepareDlg", "Next", "Далее"),
            ("PrepareDlg", "Cancel", "Отмена"),
            ("ProgressDlg", "Title", r"{\DlgFontBold8}[Progress1] [ProductName]"),
            ("ProgressDlg", "Text", "Пожалуйста, подождите, установщик [Progress2] [ProductName]."),
            ("ProgressDlg", "StatusLabel", "Статус:"),
            ("ProgressDlg", "ActionText", "Выполняется..."),
            ("ProgressDlg", "Back", "< Назад"),
            ("ProgressDlg", "Next", "Далее >"),
            ("ProgressDlg", "Cancel", "Отмена"),
            ("SelectDirectoryDlg", "Title", "Выбор папки установки"),
            ("SelectDirectoryDlg", "Back", "< Назад"),
            ("SelectDirectoryDlg", "Next", "Далее >"),
            ("SelectDirectoryDlg", "Cancel", "Отмена"),
            ("SelectDirectoryDlg", "Up", "Вверх"),
            ("SelectDirectoryDlg", "NewDir", "Новая"),
            ("UserExit", "Title", "Установка [ProductName] была прервана"),
            ("UserExit", "Back", "< Назад"),
            ("UserExit", "Cancel", "Отмена"),
            (
                "UserExit",
                "Description1",
                "Установка [ProductName] была прервана. Система не была изменена. "
                "Запустите установку снова, когда будете готовы.",
            ),
            ("UserExit", "Description2", "Нажмите «Готово», чтобы закрыть мастер установки."),
            ("UserExit", "Finish", "Готово"),
            ("WaitForCostingDlg", "Text", "Подождите, идет расчет требуемого места на диске."),
            ("WaitForCostingDlg", "Return", "ОК"),
            ("LicenseAgreementDlg", "Title", "Лицензионное соглашение"),
            ("LicenseAgreementDlg", "Back", "< Назад"),
            ("LicenseAgreementDlg", "Next", "Принять"),
            ("LicenseAgreementDlg", "Cancel", "Отмена"),
            (
                "LicenseAgreementDlg",
                "LicenseAcceptedCheckbox",
                "Я принимаю условия лицензионного соглашения",
            ),
        ]
        for dialog, control, text in translations:
            self._update_control_text(dialog, control, text)

    def _localize_radiobuttons(self) -> None:
        self._update_radiobutton_text(
            "MaintenanceForm_Action",
            "Repair",
            "&Восстановить [ProductName]",
        )
        self._update_radiobutton_text(
            "MaintenanceForm_Action",
            "Remove",
            "&Удалить [ProductName]",
        )


build_exe_options = {
    "packages": [
        "seasonal_price",
        "pandas",
        "openpyxl",
        "xlrd",
        "reportlab",
        "sqlite3",
    ],
    "includes": [],
    "excludes": [
        "tkinter",
        "test",
        "unittest",
        "pytest",
        "_pytest",
        "mypy",
        "tests",
        "rich",
        "pygments",
        "IPython",
        "matplotlib",
        "jupyter",
        "jupyterlab",
        "notebook",
        "distutils.tests",
        "setuptools.tests",
    ],
    # Super-optimized MSI profile:
    # 1) Сильно сокращаем количество файлов (главный фактор времени установки MSI).
    # 2) Оставляем бинарные/Qt-пакеты в файловой системе для стабильности импорта.
    "zip_include_packages": ["*"],
    "zip_exclude_packages": [],
    "include_msvcr": True,
    "optimize": 2,
}

bdist_msi_options = {
    "add_to_path": False,
    "all_users": False,
    "launch_on_finish": True,
    "upgrade_code": APP_UPGRADE_CODE,
    "initial_target_dir": r"[LocalAppDataFolder]\Programs\SeasonalPrice",
    "summary_data": {
        "author": APP_VENDOR,
        "comments": "Installer package for SeasonalPrice.",
        "keywords": "SeasonalPrice, Excel, Orders, MSI",
    },
    "data": {
        "Shortcut": [
            (
                "S_DESKTOP_APP",
                "DesktopFolder",
                APP_NAME,
                "TARGETDIR",
                f"[TARGETDIR]{APP_EXE}",
                None,
                None,
                None,
                None,
                None,
                None,
                "TARGETDIR",
            )
        ]
    },
}

executables = [
    Executable(
        script=str(SRC / "seasonal_price" / "presentation" / "gui.py"),
        base="gui",
        target_name=APP_EXE,
        shortcut_name=APP_NAME,
        shortcut_dir="ProgramMenuFolder",
    )
]

setup(
    name=APP_NAME,
    version=APP_VERSION,
    description=APP_DESCRIPTION,
    options={
        "build_exe": build_exe_options,
        "bdist_msi": bdist_msi_options,
    },
    executables=executables,
    cmdclass={"bdist_msi": RussianBdistMsi},
)
