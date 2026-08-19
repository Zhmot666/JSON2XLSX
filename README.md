# JSON2XLSX

Утилита с графическим интерфейсом на **PyQt6**: конвертация JSON с маркировкой (`TaskMarks`) в Excel, **отчёт агрегации** (JSON Schema + XML `unit_pack`), CSV кодов уровня 0, раздельные TXT по коробкам и XML **расформирования упаковки** (Честный ЗНАК).

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
| **Выгрузка агрегации** | `*_agg_report.json` — JSON-отчёт по схеме (как в `Example/Пример итогового файла.json`) | **productGroup** (по умолчанию `bio`), **participantId** |
| **Отчёт агрегации + CSV (ур. 0)** | `*_agg_report.json`, `*_unit_pack.xml` и `*_lv0.csv` | **productGroup** (по умолчанию `bio`), **participantId** |

Два режима JSON-отчёта агрегации:

| Кнопка | Что агрегируется | `sntins` |
|--------|------------------|----------|
| **Выгрузка агрегации** | верхний уровень (напр. паллета) | коды дочерних упаковок (напр. коробки), **без level 0** |
| **Отчёт агрегации + CSV (ур. 0)** | коробка (уровень 1) | коды изделий: **01 + GTIN (14) + 21 + серия (13)** без криптохвоста |

В обоих отчётах `unitSerialNumber` формируется **без двух ведущих нулей** (`00…` → `…`); в `*_unit_pack.xml` и при разагрегации используется **исходный** `Barcode` коробки из JSON. В CSV (`*_lv0.csv`) — **полные** штрихкоды изделий из JSON.

Примеры входного и итогового JSON можно положить в папку **`Example/`** (`Выгрузка станции агрегации.json` → `Пример итогового файла.json`).

Проверка JSON-отчёта — по **`schemas/aggregation_report.schema.json`** (папку `schemas` не удаляйте, если нужна валидация).

### Выгрузка агрегации (JSON)

Кнопка **«Выгрузка агрегации»** формирует только `*_agg_report.json` из JSON выгрузки станции (`TaskMarks`):

- отчёт строится **только для уровней выше 0** (например, паллета → коробки); коды изделий (level 0) в `sntins` **не попадают**;
- `productGroup`, `participantId`, `aggregationUnits` — по JSON Schema;
- для каждого агрегата верхнего уровня: `sntins` — коды непосредственных дочерних упаковок, `unitSerialNumber` — код агрегата.

### Отчёт агрегации (полный)

Кнопка **«Отчёт агрегации + CSV (ур. 0)»** создаёт три файла:

- `*_agg_report.json` — отчёт агрегации коробок с кодами изделий (level 0) в `sntins`;
- `*_unit_pack.xml` — XML формирования упаковки для ГИС МТ;
- `*_lv0.csv` — полные штрихкоды всех изделий (level 0), по одному на строку.

### Агрегация (XML unit_pack)

Файл `*_unit_pack.xml` — документ формирования упаковки для загрузки в ГИС МТ (создаётся кнопкой **«Отчёт агрегации + CSV (ур. 0)»**):

- `organisation` / `id_info` / `LP_info@LP_TIN` — ИНН из **participantId**;
- для каждой коробки (уровень 1) — блок `pack_content`: `pack_code` = `Barcode` коробки из JSON (без изменений), дочерние `cis` = коды изделий (01+GTIN+21+серия без криптохвоста).
- значения `pack_code` и `cis` записываются в XML через `<![CDATA[...]]>`, чтобы спецсимволы в КИ не требовали дополнительного экранирования.

Пример: `задание_unit_pack.xml` рядом с исходным JSON.

### Разагрегация (XML)

Формат соответствует документу **«Расформирование упаковки»** (`action_id="31"`, `version="2"`):

- `trade_participant_inn` — ИНН из поля «Участник»;
- `packings_list` / `packing` / `kitu` — коды коробок уровня 1 (КИТУ) как в JSON (`Barcode` без изменений).
- файл формируется стандартной XML-сериализацией (без принудительного CDATA для `kitu`).

Пример имени файла: `задание_disaggregation.xml` рядом с исходным JSON.

## Командная строка

### Отчёт агрегации

```powershell
python taskmarks_aggregation.py путь\к\файлу.json --participant-id "ваш_идентификатор"
```

Создаёт рядом с JSON: `*_agg_report.json`, `*_unit_pack.xml`, `*_lv0.csv`.

| Аргумент | Описание |
|----------|----------|
| `--product-group STR` | `productGroup` (по умолчанию — `bio`) |
| `--participant-id STR` | `participantId` в JSON и `LP_TIN` в XML |
| `--schema ПУТЬ` | JSON Schema (по умолчанию — `schemas/aggregation_report.schema.json`) |
| `--no-validate` | Не проверять JSON по схеме |

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

Сгенерированные рядом с JSON: `*_agg_report.json`, `*_unit_pack.xml`, `*_lv0.csv`, `*_disaggregation.xml`, `*_big.xlsx`, `*.xlsx` — в **`.gitignore`**. Раздельные TXT именуются штрихкодом коробки (маска в gitignore не задана). Также не коммитятся `build/`, `dist/`, `__pycache__/`, виртуальное окружение.
