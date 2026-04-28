# JSON2XLSX

Утилита с графическим интерфейсом на **PyQt6**: конвертация JSON с маркировкой (`TaskMarks`) в Excel и формирование **отчёта агрегации** по JSON Schema плюс выгрузка кодов уровня 0 в CSV.

## Требования

- **Python 3.10 или новее** (используется конструкция `match/case`).
- Windows, macOS или Linux (для GUI нужна поддержка Qt).

## Установка

```powershell
cd путь\к\JSON2XLSX
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

На Linux/macOS активация окружения: `source .venv/bin/activate`.

## Запуск приложения

```powershell
python ConvertorJ2X.py
```

После выбора JSON-файла доступны:

- **Конвертировать в XLSX** — таблица уровней маркировки; при указании лимита строк возможно разбиение на несколько файлов и общий файл `*_big.xlsx`.
- **Отчёт агрегации + CSV (ур. 0)** — создаются рядом с исходным JSON файлы `*_agg_report.json` и `*_lv0.csv`. Нужны поля **productGroup** (по умолчанию `bio`) и **participantId** (обязательно с формы). В JSON отчёта поле `sntins` содержит коды формата **01 + GTIN (14 цифр) + 21 + серия (13 символов)** без криптохвоста; в CSV — **полные** штрихкоды из исходного JSON.

Проверка сгенерированного отчёта выполняется по файлу **`schemas/aggregation_report.schema.json`** (папку `schemas` удалять не следует, если нужна валидация).

## Командная строка (отчёт агрегации без GUI)

```powershell
python taskmarks_aggregation.py путь\к\файлу.json --participant-id "ваш_идентификатор"
```

Дополнительные аргументы:

| Аргумент | Описание |
|----------|----------|
| `--product-group STR` | Значение `productGroup` (если не задано — `bio`) |
| `--participant-id STR` | Значение `participantId` в корне JSON |
| `--schema ПУТЬ` | Другой файл JSON Schema (по умолчанию — `schemas/aggregation_report.schema.json`) |
| `--no-validate` | Не проверять результат по схеме |

## Сборка Windows (исполняемый файл)

Сборка в **один файл** (onefile): зависимости и папка `schemas` вшиваются в один `JSON2XLSX.exe`. При первом запуске PyInstaller распаковывает содержимое во временный каталог (`sys._MEIPASS`); код уже учитывает это при поиске JSON Schema.

```powershell
pip install -r requirements-build.txt
pyinstaller JSON2XLSX.spec
```

Готовый файл: **`dist/JSON2XLSX.exe`**. Достаточно переносить только его (отдельные DLL рядом не нужны).

Повторная сборка после правок кода: снова `pyinstaller JSON2XLSX.spec`. Первый запуск после сборки может быть чуть дольше из‑за распаковки; антивирус иногда дольше проверяет большие onefile-EXE.

## Редактирование формы

Разметка окна задаётся в **`design.ui`** (Qt Designer). После правок пересоберите модуль интерфейса:

```powershell
pyuic6 design.ui -o design.py
```

Приложение импортирует модуль **`design`** (`design.py`), а не `design_ui.py`.

## Зависимости (`requirements.txt`)

| Пакет | Назначение |
|-------|------------|
| `xlsxwriter` | Запись `.xlsx` |
| `PyQt6` | Окно и диалоги |
| `jsonschema` | Проверка отчёта агрегации по схеме |
