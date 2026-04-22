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
- **Отчёт агрегации + CSV (ур. 0)** — создаются рядом с исходным JSON файлы `*_agg_report.json` и `*_lv0.csv`. Нужны поля **productGroup** (по умолчанию `dietary_supplements`) и **participantId** (обязательно с формы). В JSON отчёта поле `sntins` содержит компактные коды **GTIN (14 цифр) + серия (13 символов)** без криптохвоста; в CSV — **полные** штрихкоды из исходного JSON.

Проверка сгенерированного отчёта выполняется по файлу **`schemas/aggregation_report.schema.json`** (папку `schemas` удалять не следует, если нужна валидация).

## Командная строка (отчёт агрегации без GUI)

```powershell
python taskmarks_aggregation.py путь\к\файлу.json --participant-id "ваш_идентификатор"
```

Дополнительные аргументы:

| Аргумент | Описание |
|----------|----------|
| `--product-group STR` | Значение `productGroup` (если не задано — `dietary_supplements`) |
| `--participant-id STR` | Значение `participantId` в корне JSON |
| `--schema ПУТЬ` | Другой файл JSON Schema (по умолчанию — `schemas/aggregation_report.schema.json`) |
| `--no-validate` | Не проверять результат по схеме |

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
