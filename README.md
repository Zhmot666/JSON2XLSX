# JSON2XLSX

Утилита с графическим интерфейсом на **PyQt6**: конвертация JSON с маркировкой (`TaskMarks`) в Excel, формирование **отчёта агрегации** по JSON Schema, выгрузка кодов уровня 0 в CSV и XML **расформирования упаковки** (Честный ЗНАК, `DISAGGREGATION_DOCUMENT_XML`).

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

| Кнопка | Результат | Поля на форме |
|--------|-----------|---------------|
| **Конвертировать в XLSX** | Таблица уровней маркировки; при лимите строк — несколько `.xlsx` и общий `*_big.xlsx` | Лимит строк (необязательно) |
| **Раздельные TXT** | По одному `.txt` на коробку (уровень 1): полные штрихкоды изделий (уровень 0), имя файла — штрихкод коробки | Не нужны |
| **Разагрегация** | `*_disaggregation.xml` — расформирование упаковки (КИТУ, уровень 1) | **ИНН** в поле участника |
| **Отчёт агрегации + CSV (ур. 0)** | `*_agg_report.json` и `*_lv0.csv` | **productGroup** (по умолчанию `bio`), **participantId** |

В отчёте агрегации поле `sntins` — формат **01 + GTIN (14) + 21 + серия (13)** без криптохвоста; в CSV — **полные** штрихкоды из JSON.

Проверка отчёта агрегации — по **`schemas/aggregation_report.schema.json`** (папку `schemas` не удаляйте, если нужна валидация).

### Разагрегация (XML)

Формат соответствует документу **«Расформирование упаковки»** (`action_id="31"`, `version="2"`):

- `trade_participant_inn` — ИНН из поля «Участник»;
- `packings_list` / `packing` / `kitu` — коды коробок уровня 1 (КИТУ), с той же нормализацией, что `unitSerialNumber` в агрегации (снятие двух ведущих нулей при наличии).

Пример имени файла: `задание_disaggregation.xml` рядом с исходным JSON.

## Командная строка

### Отчёт агрегации

```powershell
python taskmarks_aggregation.py путь\к\файлу.json --participant-id "ваш_идентификатор"
```

| Аргумент | Описание |
|----------|----------|
| `--product-group STR` | `productGroup` (по умолчанию — `bio`) |
| `--participant-id STR` | `participantId` в корне JSON |
| `--schema ПУТЬ` | JSON Schema (по умолчанию — `schemas/aggregation_report.schema.json`) |
| `--no-validate` | Не проверять по схеме |

### Разагрегация (XML)

```powershell
python taskmarks_disaggregation.py путь\к\файлу.json --participant-inn 7707083893
```

## Сборка Windows (исполняемый файл)

Сборка **onefile**: в один `JSON2XLSX.exe` попадают зависимости, `schemas/` и `assets/`. При старте PyInstaller распаковывает содержимое во временный каталог (`sys._MEIPASS`).

```powershell
pip install -r requirements-build.txt
pyinstaller JSON2XLSX.spec --noconfirm
```

Готовый файл: **`dist/JSON2XLSX.exe`**. Достаточно переносить только его.

Повторная сборка после правок: `pyinstaller JSON2XLSX.spec --noconfirm`. Первый запуск может быть дольше; антивирус иногда дольше проверяет большие onefile-EXE.

Каталоги **`build/`** и **`dist/`** в репозиторий не попадают (см. `.gitignore`).

## Иконка приложения

Иконка окна и EXE — **`assets/app_icon.ico`** (исходник — `assets/app_icon.png`). После замены PNG пересоберите ICO, например:

```powershell
.\.venv\Scripts\python.exe -c "from pathlib import Path; from PIL import Image; img=Image.open('assets/app_icon.png').convert('RGBA'); img.save('assets/app_icon.ico', format='ICO', sizes=[(s,s) for s in (16,24,32,48,64,128,256)])"
```

Затем: `pyinstaller JSON2XLSX.spec --noconfirm`.

**Иконка в Проводнике не обновилась?** Windows кэширует значки. Переименуйте exe или выполните `ie4uinit.exe -show` после закрытия окон Проводника.

## Редактирование формы

Разметка — в **`design.ui`** (Qt Designer). После правок:

```powershell
pyuic6 design.ui -o design.py
```

Приложение импортирует **`design`** (`design.py`).

## Зависимости

| Файл | Пакеты |
|------|--------|
| `requirements.txt` | `xlsxwriter`, `PyQt6`, `jsonschema` |
| `requirements-build.txt` | то же + `pyinstaller` |

## Что не коммитить

Сгенерированные рядом с JSON файлы (`*_agg_report.json`, `*_lv0.csv`, `*_disaggregation.xml`, `*.xlsx`, раздельные TXT), каталоги `build/`, `dist/`, виртуальное окружение — перечислены в **`.gitignore`**.
