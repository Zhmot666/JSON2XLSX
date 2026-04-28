"""
Преобразование JSON с TaskMarks в отчёт агрегации по JSON Schema
и выгрузка всех Barcode с level 0 в CSV (один полный код на строку).

В JSON отчёта поле sntins — формат 01+GTIN(14)+21+Serial(13) без криптохвоста;
в CSV — исходные полные штрихкоды из поля Barcode.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


def _app_dir() -> Path:
    """Корень приложения: исходники, onefile (_MEIPASS), onedir (_internal или рядом с .exe)."""
    if not getattr(sys, "frozen", False):
        return Path(__file__).resolve().parent
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    exe_dir = Path(sys.executable).resolve().parent
    internal = exe_dir / "_internal"
    if internal.is_dir() and (internal / "schemas").is_dir():
        return internal
    return exe_dir


SCHEMA_PATH = _app_dir() / "schemas" / "aggregation_report.schema.json"

# Значение productGroup по умолчанию, если на форме / в CLI не задано.
PRODUCT_GROUP_DEFAULT = "bio"

_GS = "\x1d"
# Полный штрихкод в JSON: 01 + GTIN(14) + 21 + serial(13) + GS + криптохвост (91/92…)
_SNTIN_HEAD = re.compile(r"^01(\d{14})21(.{13})")


def barcode_to_sntin(barcode: str) -> str:
    """
    Код для поля sntins: 01 + GTIN(14) + 21 + Serial(13), без криптохвоста после GS.
    """
    if not isinstance(barcode, str):
        barcode = str(barcode)
    head = barcode.split(_GS, 1)[0]
    m = _SNTIN_HEAD.match(head)
    if not m:
        raise ValueError(
            "Ожидается префикс GS1 01+14 цифр GTIN+21+13 символов серии до первого GS; "
            f"фрагмент: {head[:48]!r}…"
        )
    return f"01{m.group(1)}21{m.group(2)}"


def _is_level0_leaf(d: Any) -> bool:
    return isinstance(d, dict) and d.get("level") == 0 and "Barcode" in d


def normalize_unit_serial_number(value: Any) -> str:
    """Для unitSerialNumber удаляем ровно два лидирующих нуля (если есть)."""
    s = str(value)
    return s[2:] if s.startswith("00") else s


def iter_level1_boxes(task_root: dict):
    """Все узлы уровня 1 с дочерними только листьями уровня 0 (коробки)."""

    def walk(obj: Any):
        if isinstance(obj, dict):
            ch = obj.get("ChildBarcodes")
            if (
                obj.get("level") == 1
                and isinstance(ch, list)
                and ch
                and all(_is_level0_leaf(c) for c in ch)
            ):
                yield obj
            for v in obj.values():
                yield from walk(v)
        elif isinstance(obj, list):
            for item in obj:
                yield from walk(item)

    yield from walk(task_root)


def boxes_to_aggregation_units(task_mark: dict) -> list[dict[str, Any]]:
    units: list[dict[str, Any]] = []
    for box in iter_level1_boxes(task_mark):
        sntins = [barcode_to_sntin(str(c["Barcode"])) for c in box["ChildBarcodes"]]
        n = len(sntins)
        units.append(
            {
                "sntins": sntins,
                "unitSerialNumber": normalize_unit_serial_number(box.get("Barcode", "")),
                "aggregationUnitCapacity": n,
                "aggregationType": "AGGREGATION",
                "aggregatedItemsCount": n,
            }
        )
    return units


def collect_level0_barcodes_ordered(data: dict) -> list[str]:
    """Полные штрихкоды уровня 0 (как в JSON), порядок обхода — по TaskMarks."""
    out: list[str] = []

    def walk(o: Any):
        if isinstance(o, dict):
            if _is_level0_leaf(o):
                out.append(str(o["Barcode"]))
            for v in o.values():
                walk(v)
        elif isinstance(o, list):
            for x in o:
                walk(x)

    for tm in data.get("TaskMarks") or []:
        if isinstance(tm, dict):
            walk(tm)
    return out


def resolve_product_group(product_group: str | None) -> str:
    if product_group is None:
        return PRODUCT_GROUP_DEFAULT
    s = str(product_group).strip()
    return s if s else PRODUCT_GROUP_DEFAULT


def resolve_participant_id(participant_id: str | None) -> str:
    """Значение participantId с формы / CLI (пустое допустимо как '')."""
    if participant_id is None:
        return ""
    return str(participant_id).strip()


def build_aggregation_report(
    data: dict,
    product_group: str | None = None,
    participant_id: str | None = None,
) -> dict[str, Any]:
    tasks = data.get("TaskMarks")
    if not isinstance(tasks, list) or not tasks:
        raise ValueError("В JSON отсутствует непустой массив TaskMarks.")
    all_units: list[dict[str, Any]] = []
    for tm in tasks:
        if not isinstance(tm, dict):
            continue
        all_units.extend(boxes_to_aggregation_units(tm))
    if not all_units:
        raise ValueError(
            "Не найдено агрегационных единиц уровня 1 с кодами уровня 0 "
            "(ожидается структура ChildBarcodes → коробки → изделия)."
        )
    return {
        "productGroup": resolve_product_group(product_group),
        "aggregationUnits": all_units,
        "participantId": resolve_participant_id(participant_id),
    }


def validate_report(instance: dict[str, Any], schema_path: Path) -> None:
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    import jsonschema

    jsonschema.validate(instance=instance, schema=schema)


def export_paths(json_path: Path) -> tuple[Path, Path]:
    parent = json_path.parent
    stem = json_path.stem
    return parent / f"{stem}_agg_report.json", parent / f"{stem}_lv0.csv"


def process_file(
    input_path: Path,
    schema_path: Path | None = None,
    *,
    validate: bool = True,
    product_group: str | None = None,
    participant_id: str | None = None,
) -> tuple[Path, Path]:
    data = json.loads(input_path.read_text(encoding="utf-8"))
    report = build_aggregation_report(
        data,
        product_group=product_group,
        participant_id=participant_id,
    )
    sch = schema_path if schema_path is not None else SCHEMA_PATH
    if validate and sch.is_file():
        validate_report(report, sch)
    out_json, out_csv = export_paths(input_path)
    out_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = collect_level0_barcodes_ordered(data)
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        for code in lines:
            f.write(f"{code}\n")
    return out_json, out_csv


def _main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="TaskMarks JSON → отчёт агрегации (_agg_report.json) + CSV кодов уровня 0 (_lv0.csv)."
    )
    p.add_argument("input_json", type=Path, help="Входной JSON с TaskMarks")
    p.add_argument(
        "--schema",
        type=Path,
        default=SCHEMA_PATH,
        help=f"JSON Schema (по умолчанию: {SCHEMA_PATH})",
    )
    p.add_argument(
        "--no-validate",
        action="store_true",
        help="Не проверять результат по схеме",
    )
    p.add_argument(
        "--product-group",
        default=None,
        metavar="STR",
        help=f"Поле productGroup в JSON отчёта (по умолчанию: {PRODUCT_GROUP_DEFAULT})",
    )
    p.add_argument(
        "--participant-id",
        default=None,
        metavar="STR",
        help="Поле participantId в корне JSON отчёта (после aggregationUnits)",
    )
    args = p.parse_args(argv)
    inp = args.input_json
    if not inp.is_file():
        print(f"Файл не найден: {inp}", file=sys.stderr)
        return 1
    try:
        out_j, out_c = process_file(
            inp,
            args.schema,
            validate=not args.no_validate,
            product_group=args.product_group,
            participant_id=args.participant_id,
        )
    except Exception as e:
        print(str(e), file=sys.stderr)
        return 1
    print(f"JSON: {out_j}")
    print(f"CSV:  {out_c}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
