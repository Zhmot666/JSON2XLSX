"""
TaskMarks JSON → документ расформирования упаковки (DISAGGREGATION_DOCUMENT_XML).

Коды КИТУ берутся из узлов уровня 1 (коробки), как unitSerialNumber в отчёте агрегации.
"""
from __future__ import annotations

import argparse
import json
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from taskmarks_aggregation import iter_level1_boxes, normalize_unit_serial_number

DISAGGREGATION_ACTION_ID = "31"
DISAGGREGATION_VERSION = "2"


def collect_kitu_codes(data: dict) -> list[str]:
    """Коды КИТУ (уровень 1) для расформирования, порядок обхода TaskMarks."""
    codes: list[str] = []
    tasks = data.get("TaskMarks")
    if not isinstance(tasks, list) or not tasks:
        raise ValueError("В JSON отсутствует непустой массив TaskMarks.")
    for tm in tasks:
        if not isinstance(tm, dict):
            continue
        for box in iter_level1_boxes(tm):
            bc = box.get("Barcode", "")
            if not bc:
                raise ValueError("У коробки уровня 1 отсутствует Barcode.")
            codes.append(normalize_unit_serial_number(bc))
    if not codes:
        raise ValueError(
            "Не найдено коробок уровня 1 с кодами уровня 0 "
            "(ожидается структура ChildBarcodes → коробки → изделия)."
        )
    return codes


def build_disaggregation_xml(participant_inn: str, kitu_codes: list[str]) -> ET.Element:
    root = ET.Element(
        "disaggregation",
        {"action_id": DISAGGREGATION_ACTION_ID, "version": DISAGGREGATION_VERSION},
    )
    ET.SubElement(root, "trade_participant_inn").text = participant_inn
    packings_list = ET.SubElement(root, "packings_list")
    for kitu in kitu_codes:
        packing = ET.SubElement(packings_list, "packing")
        ET.SubElement(packing, "kitu").text = kitu
    return root


def disaggregation_xml_bytes(participant_inn: str, kitu_codes: list[str]) -> bytes:
    root = build_disaggregation_xml(participant_inn, kitu_codes)
    ET.indent(root, space="    ")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def export_disaggregation_xml_path(json_path: Path) -> Path:
    return json_path.parent / f"{json_path.stem}_disaggregation.xml"


def process_file(input_path: Path, participant_inn: str) -> Path:
    inn = str(participant_inn).strip()
    if not inn:
        raise ValueError("Не указан ИНН участника (trade_participant_inn).")
    data = json.loads(input_path.read_text(encoding="utf-8"))
    kitu_codes = collect_kitu_codes(data)
    out_path = export_disaggregation_xml_path(input_path)
    out_path.write_bytes(disaggregation_xml_bytes(inn, kitu_codes))
    return out_path


def _main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="TaskMarks JSON → XML расформирования упаковки (_disaggregation.xml)."
    )
    p.add_argument("input_json", type=Path, help="Входной JSON с TaskMarks")
    p.add_argument(
        "--participant-inn",
        required=True,
        metavar="INN",
        help="ИНН участника оборота (trade_participant_inn)",
    )
    args = p.parse_args(argv)
    inp = args.input_json
    if not inp.is_file():
        print(f"Файл не найден: {inp}", file=sys.stderr)
        return 1
    try:
        out = process_file(inp, args.participant_inn)
    except Exception as e:
        print(str(e), file=sys.stderr)
        return 1
    print(f"XML: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
