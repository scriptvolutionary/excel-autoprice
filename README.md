# Сезонный Прайс

GUI-приложение для массовой обработки прайсов и заказов.

## Что умеет

- Формировать клиентский прайс из листа `Склад`.
- Пакетно импортировать заказы `.xls/.xlsx`.
- Выполнять аллокацию (FIFO и пропорционально).
- Генерировать подтверждения `XLSX + PDF`.
- Строить остаточный прайс.
- Архивировать сезон.

## Быстрый запуск (GUI)

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m ensurepip --upgrade
.\.venv\Scripts\python.exe -m pip install -e .[dev,gui]
seasonal-price
```

Альтернатива без `pip install -e`:

```powershell
$env:PYTHONPATH="src"
.\.venv\Scripts\python.exe -m seasonal_price.presentation.gui
```

## Тесты (с автоочисткой кэшей)

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_tests.ps1
```

## Сборка onefile EXE (без установщика)

Подготовка окружения (один раз):

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
```

Сборка:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_pyinstaller.ps1
```

Режимы сборки:

- Быстрый инкрементальный (по умолчанию): без `--clean`, для ежедневной разработки.
- Чистый: `powershell -ExecutionPolicy Bypass -File .\scripts\build_pyinstaller.ps1 -Clean` (для релизной проверки).

Файл для передачи:

- `dist\SeasonalPrice.exe`

## Структура данных

- `data/` - база SQLite и служебные данные.
- `outputs/` - сводки и артефакты обработки.
- `logs/` - лог-файлы.
- `archive/` - архив закрытых сезонов.

## Кодировка и правила

- Все файлы проекта: `UTF-8`.
- Пользовательские тексты, логи и ошибки: русский язык.
- Подробные правила разработки: [AGENTS.md](./AGENTS.md).
