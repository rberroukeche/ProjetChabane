from __future__ import annotations

import posixpath
import re
import unicodedata
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Flowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


NAMESPACE = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "relationships": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "package": "http://schemas.openxmlformats.org/package/2006/relationships",
}
STORE_FIELD = "Responsible Store"
STATUS_FIELD = "Status"
ZONE_FIELD = "Zone"
REPORT_ZONE = "Zone 4"
FINISHED_STATUS = "Finished"
TARGET_PROGRESS = 85
OUTPUT_DIRECTORY = "rapports"
PAGE_WIDTH, PAGE_HEIGHT = landscape(A4)
PAGE_MARGIN = 15 * mm
NAVY = colors.HexColor("#132B46")
TEAL = colors.HexColor("#007F86")
GREEN = colors.HexColor("#198754")
AMBER = colors.HexColor("#D48A00")
RED = colors.HexColor("#C64545")
INK = colors.HexColor("#243447")
MUTED = colors.HexColor("#68788A")
PALE = colors.HexColor("#F3F6FA")
PALE_BLUE = colors.HexColor("#EAF1F7")
PALE_TEAL = colors.HexColor("#E8F5F4")
PALE_RED = colors.HexColor("#FBEDEE")
GRID = colors.HexColor("#DFE6ED")
WHITE = colors.white


def _column_index(cell_reference: str) -> int:
    letters = re.match(r"[A-Z]+", cell_reference)
    if letters is None:
        raise ValueError(f"Référence de cellule Excel invalide : {cell_reference}")

    index = 0
    for letter in letters.group():
        index = index * 26 + ord(letter) - ord("A") + 1
    return index - 1


def _read_cell_value(cell: ElementTree.Element, shared_strings: list[str]) -> str:
    cell_type = cell.get("t")
    if cell_type == "inlineStr":
        return "".join(
            text.text or "" for text in cell.findall(".//main:t", NAMESPACE)
        ).strip()

    value = cell.find("main:v", NAMESPACE)
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)].strip()
        except (ValueError, IndexError) as error:
            raise ValueError("Le classeur contient une chaîne Excel invalide.") from error
    return value.text.strip()


def load_database(
    workbook_path: Path,
    sheet_name: str = "Database",
    required_fields: set[str] | None = None,
) -> list[dict[str, str]]:
    """Read a named worksheet from an .xlsx workbook without requiring Excel."""
    try:
        with zipfile.ZipFile(workbook_path) as workbook:
            workbook_xml = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
            relationships_xml = ElementTree.fromstring(
                workbook.read("xl/_rels/workbook.xml.rels")
            )
            relationships = {
                relationship.get("Id", ""): relationship.get("Target", "")
                for relationship in relationships_xml.findall(
                    "package:Relationship", NAMESPACE
                )
            }
            sheets = workbook_xml.findall("main:sheets/main:sheet", NAMESPACE)
            database_sheet = next(
                (
                    sheet
                    for sheet in sheets
                    if sheet.get("name", "").casefold() == sheet_name.casefold()
                ),
                None,
            )
            if database_sheet is None:
                raise ValueError(
                    f"The '{sheet_name}' sheet is missing from the workbook."
                )

            relationship_id = database_sheet.get(
                f"{{{NAMESPACE['relationships']}}}id", ""
            )
            target = relationships.get(relationship_id)
            if not target:
                raise ValueError(f"Could not locate the '{sheet_name}' sheet.")
            sheet_path = posixpath.normpath(
                target.lstrip("/") if target.startswith("/") else f"xl/{target}"
            )

            shared_strings: list[str] = []
            if "xl/sharedStrings.xml" in workbook.namelist():
                strings_xml = ElementTree.fromstring(
                    workbook.read("xl/sharedStrings.xml")
                )
                shared_strings = [
                    "".join(
                        text.text or ""
                        for text in item.findall(".//main:t", NAMESPACE)
                    )
                    for item in strings_xml.findall("main:si", NAMESPACE)
                ]

            sheet_xml = ElementTree.fromstring(workbook.read(sheet_path))
    except (OSError, zipfile.BadZipFile, KeyError, ElementTree.ParseError) as error:
        raise ValueError(
            f"Could not read Excel workbook '{workbook_path.name}'."
        ) from error

    sheet_rows = sheet_xml.findall(".//main:sheetData/main:row", NAMESPACE)
    if not sheet_rows:
        raise ValueError(f"The '{sheet_name}' sheet is empty.")

    header_cells = {
        _column_index(cell.get("r", "")): _read_cell_value(cell, shared_strings)
        for cell in sheet_rows[0].findall("main:c", NAMESPACE)
    }
    headers = {
        column: name for column, name in header_cells.items() if name
    }
    fields_to_check = required_fields or {
        STORE_FIELD,
        STATUS_FIELD,
        ZONE_FIELD,
        "Campaign Code",
        "Campaign Name",
    }
    missing_fields = fields_to_check.difference(headers.values())
    if missing_fields:
        raise ValueError(
            f"Missing columns in '{sheet_name}': "
            + ", ".join(sorted(missing_fields))
        )

    rows: list[dict[str, str]] = []
    for row in sheet_rows[1:]:
        values = {
            _column_index(cell.get("r", "")): _read_cell_value(cell, shared_strings)
            for cell in row.findall("main:c", NAMESPACE)
        }
        record = {
            header: values.get(column, "") for column, header in headers.items()
        }
        if any(record.values()):
            rows.append(record)
    if not rows:
        raise ValueError("No data found in the 'Database' sheet.")
    return rows


def find_workbook() -> Path:
    workbooks = sorted(
        path
        for path in Path(__file__).resolve().parent.glob("*.xlsx")
        if not path.name.startswith("~$")
    )
    if not workbooks:
        raise FileNotFoundError(
            "No .xlsx file found next to app.py. "
            "Place the Excel workbook in the same folder."
        )
    if len(workbooks) == 1:
        return workbooks[0]

    print("Several Excel files are available:")
    for index, workbook in enumerate(workbooks, start=1):
        print(f"  {index}. {workbook.name}")
    while True:
        choice = input("Choose a file number: ").strip()
        if choice.isdigit() and 1 <= int(choice) <= len(workbooks):
            return workbooks[int(choice) - 1]
        print("Invalid choice. Enter a file number.")


def _record_summary(records: list[dict[str, str]]) -> tuple[int, int, float]:
    finished = sum(
        record.get(STATUS_FIELD, "").casefold() == FINISHED_STATUS.casefold()
        for record in records
    )
    total = len(records)
    return total, finished, (finished / total * 100) if total else 0.0


def _paragraph_text(value: Any) -> str:
    return escape(str(value)).replace("\n", "<br/>")


def _progress_color(progress: float) -> colors.Color:
    if progress >= TARGET_PROGRESS:
        return GREEN
    if progress >= 60:
        return AMBER
    return RED


class ProgressCell(Flowable):
    def __init__(self, progress: float, width: float = 34 * mm):
        super().__init__()
        self.progress = max(0, min(100, progress))
        self.width = width
        self.height = 7 * mm

    def draw(self) -> None:
        color = _progress_color(self.progress)
        self.canv.setFont("Helvetica-Bold", 8)
        self.canv.setFillColor(color)
        self.canv.drawCentredString(
            self.width / 2, self.height - 2.5 * mm, f"{self.progress:.1f}%"
        )
        bar_width = self.width - 4 * mm
        bar_height = 2 * mm
        bar_x = 2 * mm
        bar_y = 1 * mm
        self.canv.setFillColor(PALE_BLUE)
        self.canv.roundRect(bar_x, bar_y, bar_width, bar_height, 1 * mm, fill=1, stroke=0)
        self.canv.setFillColor(color)
        self.canv.roundRect(
            bar_x,
            bar_y,
            max(bar_height, bar_width * self.progress / 100),
            bar_height,
            1 * mm,
            fill=1,
            stroke=0,
        )


class StatusBarChart(Flowable):
    def __init__(self, counts: Counter[str], width: float = 180 * mm):
        super().__init__()
        self.counts = counts
        self.width = width
        self.row_height = 8 * mm
        self.height = max(1, len(counts)) * self.row_height

    def draw(self) -> None:
        maximum = max(self.counts.values(), default=1)
        label_width = 50 * mm
        count_width = 14 * mm
        bar_width = self.width - label_width - count_width
        for index, (label, count) in enumerate(self.counts.most_common()):
            y = self.height - (index + 1) * self.row_height + 2 * mm
            self.canv.setFont("Helvetica", 8)
            self.canv.setFillColor(INK)
            self.canv.drawString(0, y, label[:34])
            self.canv.setFillColor(PALE_BLUE)
            self.canv.roundRect(
                label_width,
                y - 1 * mm,
                bar_width,
                4 * mm,
                1.5 * mm,
                fill=1,
                stroke=0,
            )
            self.canv.setFillColor(
                GREEN if label.casefold() == FINISHED_STATUS.casefold() else TEAL
            )
            self.canv.roundRect(
                label_width,
                y - 1 * mm,
                max(2 * mm, bar_width * count / maximum),
                4 * mm,
                1.5 * mm,
                fill=1,
                stroke=0,
            )
            self.canv.setFillColor(NAVY)
            self.canv.drawRightString(
                self.width, y, f"{count:,} ({count / sum(self.counts.values()) * 100:.1f}%)"
            )


def _make_table(
    rows: list[list[Any]],
    widths: list[float],
    font_size: int = 8,
    highlight_top: bool = False,
) -> Table:
    header_style = ParagraphStyle(
        name="TableHeader",
        fontName="Helvetica-Bold",
        fontSize=font_size,
        leading=font_size + 2,
        textColor=WHITE,
        alignment=TA_CENTER,
    )
    table_rows = [
        [
            Paragraph(_paragraph_text(value), header_style)
            if isinstance(value, str)
            else value
            for value in rows[0]
        ],
        *rows[1:],
    ]
    table = Table(table_rows, colWidths=widths, repeatRows=1, hAlign="LEFT")
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("TEXTCOLOR", (0, 1), (-1, -1), INK),
        ("FONTSIZE", (0, 0), (-1, -1), font_size),
        ("LEADING", (0, 0), (-1, -1), font_size + 2),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("LINEBELOW", (0, 0), (-1, 0), 1, TEAL),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, GRID),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, PALE]),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, 0), 3),
        ("RIGHTPADDING", (0, 0), (-1, 0), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]
    if highlight_top:
        for row_index, color in enumerate(
            (
                colors.HexColor("#FFF1C2"),
                colors.HexColor("#E8EDF2"),
                colors.HexColor("#F6E2D3"),
            ),
            start=1,
        ):
            if row_index < len(rows):
                table_styles.append(
                    ("BACKGROUND", (0, row_index), (-1, row_index), color)
                )
    table.setStyle(
        TableStyle(table_styles)
    )
    return table


def _kpi_cards(
    cards: list[tuple[str, str, str, colors.Color]],
    styles: Any,
    total_width: float,
) -> Table:
    gap = 3 * mm
    card_width = (total_width - gap * (len(cards) - 1)) / len(cards)
    cells: list[Any] = []
    column_widths: list[float] = []
    for index, (label, value, detail, accent) in enumerate(cards):
        card = Table(
            [
                [Paragraph(_paragraph_text(label.upper()), styles["KpiLabel"])],
                [Paragraph(_paragraph_text(value), styles["KpiValue"])],
                [Paragraph(_paragraph_text(detail), styles["KpiDetail"])],
            ],
            colWidths=[card_width],
        )
        card.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), WHITE),
                    ("BOX", (0, 0), (-1, -1), 0.6, GRID),
                    ("LINEABOVE", (0, 0), (-1, 0), 3, accent),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, 0), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 3),
                    ("TOPPADDING", (0, 1), (-1, 1), 2),
                    ("BOTTOMPADDING", (0, 1), (-1, 1), 1),
                    ("TOPPADDING", (0, 2), (-1, 2), 1),
                    ("BOTTOMPADDING", (0, 2), (-1, 2), 8),
                ]
            )
        )
        cells.append(card)
        column_widths.append(card_width)
        if index < len(cards) - 1:
            cells.append("")
            column_widths.append(gap)
    outer = Table([cells], colWidths=column_widths)
    outer.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return outer


def _insight_panel(lines: list[str], styles: Any) -> Table:
    content = [
        Paragraph(_paragraph_text(line), styles["InsightText"]) for line in lines
    ]
    panel = Table([[content]], colWidths=[PAGE_WIDTH - 2 * PAGE_MARGIN])
    panel.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), PALE_TEAL),
                ("BOX", (0, 0), (-1, -1), 0.5, GRID),
                ("LINEBEFORE", (0, 0), (0, 0), 3, TEAL),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
            ]
        )
    )
    return panel


def _group_campaigns(
    records: list[dict[str, str]],
) -> dict[str, list[dict[str, str]]]:
    campaigns: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in records:
        code = record.get("Campaign Code", "").strip() or "No code"
        campaigns[code].append(record)
    return campaigns


def _open_records(records: list[dict[str, str]]) -> list[dict[str, str]]:
    return [
        record
        for record in records
        if record.get(STATUS_FIELD, "").strip().casefold()
        != FINISHED_STATUS.casefold()
    ]


def _dimension_rows(
    records: list[dict[str, str]],
    field: str,
    styles: Any,
    display_names: dict[str, str] | None = None,
    use_dealership_vocabulary: bool = False,
) -> list[list[Any]]:
    groups: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in records:
        value = record.get(field, "").strip() or "Not recorded"
        groups[value].append(record)

    rows: list[list[Any]] = [
        [
            "Segment",
            "Total campaigns",
            "Finished",
            "To be finished",
            "Completion rate",
        ]
        if use_dealership_vocabulary
        else ["Segment", "Total", "Completed", "Open", "Completion"]
    ]
    for value, items in sorted(
        groups.items(),
        key=lambda item: (-len(_open_records(item[1])), item[0].casefold()),
    ):
        total, completed, progress = _record_summary(items)
        remaining = total - completed
        if not remaining:
            continue
        display_value = (display_names or {}).get(value, value)
        rows.append(
            [
                Paragraph(_paragraph_text(display_value), styles["SmallBody"]),
                total,
                completed,
                remaining,
                ProgressCell(progress),
            ]
        )
    return rows


def _campaign_dimension_rows(
    records: list[dict[str, str]],
    styles: Any,
    use_dealership_vocabulary: bool = False,
) -> list[list[Any]]:
    headers = (
        [
            "Campaign code / name",
            "Total campaigns",
            "Finished",
            "To be finished",
            "Completion rate",
            "Extra to 85%",
        ]
        if use_dealership_vocabulary
        else [
            "Campaign code / name",
            "Total",
            "Completed",
            "Open",
            "Progress",
            "Extra to 85%",
        ]
    )
    rows: list[list[Any]] = [
        headers
    ]
    groups = _group_campaigns(records)
    for code, items in sorted(
        groups.items(),
        key=lambda item: (-len(_open_records(item[1])), item[0].casefold()),
    ):
        total, completed, progress = _record_summary(items)
        if total == completed:
            continue
        name = items[0].get("Campaign Name", "").strip() or "Unnamed campaign"
        rows.append(
            [
                Paragraph(
                    f"<b>{_paragraph_text(code)}</b><br/>{_paragraph_text(name)}",
                    styles["SmallBody"],
                ),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
                max(0, (TARGET_PROGRESS * total + 99) // 100 - completed),
            ]
        )
    return rows


def _completion_period_rows(records: list[dict[str, str]]) -> list[list[Any]]:
    completed = [
        record
        for record in records
        if record.get(STATUS_FIELD, "").strip().casefold()
        == FINISHED_STATUS.casefold()
    ]
    periods: Counter[tuple[str, str]] = Counter()
    for record in completed:
        year = record.get("Completion Year", "").strip() or "Not recorded"
        quarter = record.get("Completion Quarter", "").strip() or "Not recorded"
        periods[(year, quarter)] += 1
    rows: list[list[Any]] = [["Year", "Quarter", "Completed", "Share of completed"]]
    for (year, quarter), count in sorted(periods.items()):
        rows.append(
            [
                year,
                quarter,
                count,
                f"{count / len(completed) * 100:.1f}%" if completed else "0%",
            ]
        )
    return rows


def _dealer_campaign_hotspot_rows(
    store_records: dict[str, list[dict[str, str]]],
    total_open: int,
    styles: Any,
    limit: int = 20,
) -> list[list[Any]]:
    hotspots = _dealer_campaign_hotspots(store_records)
    rows: list[list[Any]] = [
        ["Dealership", "Campaign", "Total", "Open", "Completion", "Zone open share"]
    ]
    for open_count, store, code, name, items in hotspots[:limit]:
        total, completed, progress = _record_summary(items)
        rows.append(
            [
                Paragraph(_paragraph_text(store), styles["SmallBody"]),
                Paragraph(
                    f"<b>{_paragraph_text(code)}</b><br/>{_paragraph_text(name or 'Unnamed campaign')}",
                    styles["SmallBody"],
                ),
                total,
                open_count,
                ProgressCell(progress),
                f"{open_count / total_open * 100:.1f}%" if total_open else "0%",
            ]
        )
    return rows


def _dealer_campaign_hotspots(
    store_records: dict[str, list[dict[str, str]]],
) -> list[tuple[int, str, str, str, list[dict[str, str]]]]:
    hotspots: list[
        tuple[int, str, str, str, list[dict[str, str]]]
    ] = []
    for store, records in store_records.items():
        for code, campaign_records in _group_campaigns(records).items():
            open_count = len(_open_records(campaign_records))
            if open_count:
                name = campaign_records[0].get("Campaign Name", "").strip()
                hotspots.append((open_count, store, code, name, campaign_records))
    return sorted(
        hotspots,
        key=lambda item: (-item[0], item[1].casefold(), item[2].casefold()),
    )


def _dealer_group_name(store_name: str) -> str:
    match = re.search(r"\s+x\s+(.+)$", store_name, flags=re.IGNORECASE)
    return match.group(1).strip() if match else "Group not specified"


def _rank_dealership_groups(
    zone_records: dict[str, list[dict[str, str]]],
) -> list[tuple[str, int, int, int, float]]:
    records_by_group: dict[str, list[dict[str, str]]] = defaultdict(list)
    names_by_group: dict[str, str] = {}
    dealerships_by_group: dict[str, set[str]] = defaultdict(set)
    for records in zone_records.values():
        for record in records:
            dealership = record.get(STORE_FIELD, "").strip()
            if not dealership:
                continue
            group = _dealer_group_name(dealership)
            group_key = group.casefold()
            names_by_group.setdefault(group_key, group)
            records_by_group[group_key].append(record)
            dealerships_by_group[group_key].add(dealership)

    ranked = []
    for group_key, records in records_by_group.items():
        total, completed, progress = _record_summary(records)
        ranked.append(
            (
                names_by_group[group_key],
                len(dealerships_by_group[group_key]),
                total,
                completed,
                progress,
            )
        )
    return sorted(
        ranked,
        key=lambda item: (-item[4], -item[2], item[0].casefold()),
    )


def _top_dealership_group_rows(
    zone_records: dict[str, list[dict[str, str]]],
    styles: Any,
    limit: int = 10,
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["Rank", "Group", "Dealerships", "Records", "Completed", "Open", "Completion"]
    ]
    for rank, (group, dealerships, total, completed, progress) in enumerate(
        _rank_dealership_groups(zone_records)[:limit], start=1
    ):
        rows.append(
            [
                rank,
                Paragraph(_paragraph_text(group), styles["SmallBody"]),
                dealerships,
                total,
                completed,
                total - completed,
                ProgressCell(progress),
            ]
        )
    return rows


def _rank_dealerships(
    zone_records: dict[str, list[dict[str, str]]],
) -> list[tuple[str, str, str, int, int, float]]:
    ranked: list[tuple[str, str, str, int, int, float]] = []
    for zone, records in zone_records.items():
        store_records: dict[str, list[dict[str, str]]] = defaultdict(list)
        for record in records:
            store = record.get(STORE_FIELD, "").strip()
            if store:
                store_records[store].append(record)
        for store, items in store_records.items():
            total, completed, progress = _record_summary(items)
            ranked.append(
                (
                    zone,
                    store,
                    _dealer_group_name(store),
                    total,
                    completed,
                    progress,
                )
            )
    ranked.sort(
        key=lambda item: (-item[5], -item[3], item[1].casefold(), item[0].casefold())
    )
    return ranked


def _rank_dealerships_by_finished(
    zone_records: dict[str, list[dict[str, str]]],
) -> list[tuple[str, str, str, int, int, float]]:
    ranked = _rank_dealerships(zone_records)
    ranked.sort(
        key=lambda item: (-item[4], -item[3], item[1].casefold(), item[0].casefold())
    )
    return ranked


def _rank_zones_by_finished(
    zone_records: dict[str, list[dict[str, str]]],
) -> list[tuple[str, int, int, int, float, str]]:
    ranked: list[tuple[str, int, int, int, float, str]] = []
    for zone, records in zone_records.items():
        if not records:
            continue
        total, finished, completion = _record_summary(records)
        zone_leaders = _rank_dealerships_by_finished({zone: records})
        top_dealership = zone_leaders[0][1] if zone_leaders else "No dealership"
        ranked.append((zone, total, finished, total - finished, completion, top_dealership))
    ranked.sort(
        key=lambda item: (
            -item[2],
            int(re.search(r"\d+$", item[0]).group())
            if re.search(r"\d+$", item[0])
            else 0,
        )
    )
    return ranked


def _top_dealership_rows(
    zone_records: dict[str, list[dict[str, str]]],
    styles: Any,
    limit: int = 10,
    include_zone: bool = False,
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        (
            ["Rank", "Zone", "Dealership", "Group", "Cases", "Finished", "Open", "Rate"]
            if include_zone
            else ["Rank", "Dealership", "Group", "Records", "Completed", "Open", "Completion"]
        )
    ]
    for rank, (zone, store, group, total, completed, progress) in enumerate(
        _rank_dealerships(zone_records)[:limit], start=1
    ):
        values: list[Any] = [
            rank,
            Paragraph(_paragraph_text(zone), styles["SmallBody"]),
        ] if include_zone else [rank]
        values.extend(
            [
                Paragraph(_paragraph_text(store), styles["SmallBody"]),
                Paragraph(_paragraph_text(group), styles["SmallBody"]),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
            ]
        )
        rows.append(values)
    return rows


def _top_finished_dealership_rows(
    zone_records: dict[str, list[dict[str, str]]],
    styles: Any,
    limit: int = 10,
    include_zone: bool = True,
) -> list[list[Any]]:
    rows: list[list[Any]] = [["Rank"]]
    if include_zone:
        rows[0].append("Zone")
    rows[0].extend(
        [
            "Dealership",
            "Group",
            "Total campaigns",
            "Finished",
            "To be finished",
            "Completion rate",
        ]
    )
    for rank, (zone, store, group, total, completed, progress) in enumerate(
        _rank_dealerships_by_finished(zone_records)[:limit], start=1
    ):
        row: list[Any] = [rank]
        if include_zone:
            row.append(Paragraph(_paragraph_text(zone), styles["SmallBody"]))
        row.extend(
            [
                Paragraph(_paragraph_text(store), styles["SmallBody"]),
                Paragraph(_paragraph_text(group), styles["SmallBody"]),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
            ]
        )
        rows.append(row)
    return rows


def _top_zone_finished_rows(
    zone_records: dict[str, list[dict[str, str]]],
    styles: Any,
    dealership_counts: dict[str, int] | None = None,
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        [
            "Rank",
            "Zone",
            "Dealerships",
            "Top dealership",
            "Total campaigns",
            "Finished",
            "To be finished",
            "Completion rate",
        ]
    ]
    for rank, (
        zone,
        total,
        finished,
        to_be_finished,
        completion,
        top_dealership,
    ) in enumerate(_rank_zones_by_finished(zone_records), start=1):
        rows.append(
            [
                rank,
                Paragraph(_paragraph_text(zone), styles["SmallBody"]),
                (dealership_counts or {}).get(zone, 0),
                Paragraph(_paragraph_text(top_dealership), styles["SmallBody"]),
                total,
                finished,
                to_be_finished,
                ProgressCell(completion),
            ]
        )
    return rows


def _vin_integrity_summary(
    records: list[dict[str, str]],
) -> tuple[int, int, int, int]:
    campaigns_by_vin: dict[str, set[str]] = defaultdict(set)
    rows_by_vin_campaign: Counter[tuple[str, str]] = Counter()
    statuses_by_vin_campaign: dict[tuple[str, str], set[str]] = defaultdict(set)
    for record in records:
        vin = record.get("VIN", "").strip()
        if not vin:
            continue
        campaign = record.get("Campaign Code", "").strip() or "No code"
        key = (vin, campaign)
        campaigns_by_vin[vin].add(campaign)
        rows_by_vin_campaign[key] += 1
        statuses_by_vin_campaign[key].add(record.get(STATUS_FIELD, "").strip())
    multi_campaign_vins = sum(
        len(campaigns) > 1 for campaigns in campaigns_by_vin.values()
    )
    duplicate_keys = [
        key for key, count in rows_by_vin_campaign.items() if count > 1
    ]
    extra_rows = sum(rows_by_vin_campaign[key] - 1 for key in duplicate_keys)
    status_conflicts = sum(
        len(statuses_by_vin_campaign[key]) > 1 for key in duplicate_keys
    )
    return multi_campaign_vins, len(duplicate_keys), extra_rows, status_conflicts


def _campaign_model_rows(
    records: list[dict[str, str]], styles: Any, limit: int = 12
) -> list[list[Any]]:
    groups: dict[
        tuple[str, str], list[dict[str, str]]
    ] = defaultdict(list)
    for record in records:
        campaign = record.get("Campaign Code", "").strip() or "No code"
        model = record.get("Model", "").strip() or "Not recorded"
        groups[(campaign, model)].append(record)

    prioritized = [
        (key, items) for key, items in groups.items() if _open_records(items)
    ]
    rows: list[list[Any]] = [
        ["Campaign / vehicle model", "Total", "Open", "Completion", "Zone open share"]
    ]
    for (code, model), items in sorted(
        prioritized,
        key=lambda item: (
            -len(_open_records(item[1])),
            item[0][0].casefold(),
            item[0][1].casefold(),
        ),
    )[:limit]:
        total, completed, progress = _record_summary(items)
        campaign_name = items[0].get("Campaign Name", "").strip() or "Unnamed campaign"
        rows.append(
            [
                Paragraph(
                    f"<b>{_paragraph_text(code)} — {_paragraph_text(campaign_name)}</b>"
                    f"<br/>{_paragraph_text(model)}",
                    styles["SmallBody"],
                ),
                total,
                total - completed,
                ProgressCell(progress),
                _open_case_share(items, len(_open_records(records))),
            ]
        )
    return rows


def _data_quality_rows(
    records: list[dict[str, str]], styles: Any
) -> list[list[Any]]:
    fields = [
        ("Open records without a job order", "Job Order No.", True),
        ("Open records without a completion date", "Completion Date", True),
        ("Records without an actual operating store", "Actual Operating Store", False),
        ("Records without in-store time", "In-store Time", False),
        ("Records without repair status", "Repair Status", False),
        ("Records without settlement time", "Settlement Time", False),
    ]
    opened = _open_records(records)
    rows: list[list[Any]] = [["Data check", "Missing", "Records checked", "Missing rate"]]
    for label, field, open_only in fields:
        checked = opened if open_only else records
        missing = sum(not record.get(field, "").strip() for record in checked)
        rows.append(
            [
                Paragraph(_paragraph_text(label), styles["SmallBody"]),
                missing,
                len(checked),
                f"{missing / len(checked) * 100:.1f}%" if checked else "0%",
            ]
        )
    return rows


def _open_case_share(records: list[dict[str, str]], total_open: int) -> str:
    open_count = len(_open_records(records))
    return f"{open_count / total_open * 100:.1f}%" if total_open else "0%"


def _campaign_rows(
    records: list[dict[str, str]], styles: Any
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        [
            "Campaign code",
            "Campaign",
            "Total campaigns",
            "Finished",
            "To be finished",
            "Completion rate",
        ]
    ]
    campaigns = _group_campaigns(records)
    for code, items in sorted(
        campaigns.items(),
        key=lambda item: (-_record_summary(item[1])[0], item[0].casefold()),
    ):
        total, completed, progress = _record_summary(items)
        campaign_name = items[0].get("Campaign Name", "").strip() or "Unnamed campaign"
        rows.append(
            [
                _paragraph_text(code),
                Paragraph(_paragraph_text(campaign_name), styles["SmallBody"]),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
            ]
        )
    return rows


def _store_rows(
    store_records: dict[str, list[dict[str, str]]],
    styles: Any,
    total_open: int,
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["Dealership", "Total", "Completed", "Open", "Completion", "Zone open share"]
    ]
    for store, items in sorted(
        store_records.items(),
        key=lambda item: (-_record_summary(item[1])[0], item[0].casefold()),
    ):
        total, completed, progress = _record_summary(items)
        rows.append(
            [
                Paragraph(_paragraph_text(store), styles["SmallBody"]),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
                _open_case_share(items, total_open),
            ]
        )
    return rows


def _assigned_dealership_rows(
    assigned_dealership_records: dict[str, list[dict[str, str]]],
    styles: Any,
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        [
            "Dealership site",
            "Total campaigns",
            "Finished",
            "To be finished",
            "Completion rate",
            "Campaign data",
        ]
    ]
    for site, items in sorted(
        assigned_dealership_records.items(), key=lambda item: item[0].casefold()
    ):
        total, finished, progress = _record_summary(items)
        rows.append(
            [
                Paragraph(_paragraph_text(site), styles["SmallBody"]),
                total,
                finished,
                total - finished,
                ProgressCell(progress) if total else "—",
                "Campaign records found" if total else "No campaign records",
            ]
        )
    return rows


def _priority_campaign_rows(
    records: list[dict[str, str]], styles: Any
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["Campaign code", "Campaign", "Total", "Completed", "Remaining", "Progress", "Extra to 85%"]
    ]
    below_target = [
        (code, items)
        for code, items in _group_campaigns(records).items()
        if _record_summary(items)[2] < TARGET_PROGRESS
    ]
    for code, items in sorted(
        below_target,
        key=lambda item: (
            -(_record_summary(item[1])[0] - _record_summary(item[1])[1]),
            _record_summary(item[1])[2],
        ),
    ):
        total, completed, progress = _record_summary(items)
        target_gap = max(0, (TARGET_PROGRESS * total + 99) // 100 - completed)
        campaign_name = items[0].get("Campaign Name", "").strip() or "Unnamed campaign"
        rows.append(
            [
                _paragraph_text(code),
                Paragraph(_paragraph_text(campaign_name), styles["SmallBody"]),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
                target_gap,
            ]
        )
    return rows


def _priority_store_rows(
    store_records: dict[str, list[dict[str, str]]],
    styles: Any,
    total_open: int,
) -> list[list[Any]]:
    rows: list[list[Any]] = [
        ["Dealership", "Total", "Completed", "Open", "Completion", "Extra to 85%", "Zone open share"]
    ]
    below_target = [
        (store, items)
        for store, items in store_records.items()
        if _record_summary(items)[2] < TARGET_PROGRESS
    ]
    for store, items in sorted(
        below_target,
        key=lambda item: (
            -(_record_summary(item[1])[0] - _record_summary(item[1])[1]),
            _record_summary(item[1])[2],
        ),
    ):
        total, completed, progress = _record_summary(items)
        target_gap = max(0, (TARGET_PROGRESS * total + 99) // 100 - completed)
        rows.append(
            [
                Paragraph(_paragraph_text(store), styles["SmallBody"]),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
                target_gap,
                _open_case_share(items, total_open),
            ]
        )
    return rows


def _global_story(
    records: list[dict[str, str]],
    styles: Any,
    report_date: date,
    assigned_dealership_count: int | None = None,
    assigned_dealership_records: dict[str, list[dict[str, str]]] | None = None,
    report_zone: str = REPORT_ZONE,
) -> list[Any]:
    total, finished, progress = _record_summary(records)
    store_records: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in records:
        store = record.get(STORE_FIELD, "").strip()
        if store:
            store_records[store].append(record)
    unassigned = sum(not record.get(STORE_FIELD, "").strip() for record in records)
    campaign_count = len(_group_campaigns(records))
    open_records = _open_records(records)
    total_open = len(open_records)
    status_counts = Counter(
        record.get(STATUS_FIELD, "").strip() or "Not specified" for record in records
    )
    campaigns = _group_campaigns(records)
    stores_below_target = [
        (name, items)
        for name, items in store_records.items()
        if _record_summary(items)[2] < TARGET_PROGRESS
    ]
    campaigns_below_target = [
        (code, items)
        for code, items in campaigns.items()
        if _record_summary(items)[2] < TARGET_PROGRESS
    ]
    top_store = max(
        stores_below_target,
        key=lambda item: len(item[1]) - _record_summary(item[1])[1],
        default=None,
    )
    top_campaign = max(
        campaigns.items(),
        key=lambda item: len(item[1]) - _record_summary(item[1])[1],
        default=None,
    )
    top_dealer_campaign = _dealer_campaign_hotspots(store_records)
    campaign_type_groups: dict[str, list[dict[str, str]]] = defaultdict(list)
    model_groups: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in records:
        campaign_type_groups[
            record.get("Campaign Type", "").strip() or "Not recorded"
        ].append(record)
        model_groups[record.get("Model", "").strip() or "Not recorded"].append(record)
    top_campaign_type = max(
        campaign_type_groups.items(),
        key=lambda item: len(_open_records(item[1])),
        default=None,
    )
    top_model = max(
        model_groups.items(),
        key=lambda item: len(_open_records(item[1])),
        default=None,
    )
    top_campaign_model_rows = _campaign_model_rows(records, styles, limit=1)
    top_campaign_model = (
        top_campaign_model_rows[1] if len(top_campaign_model_rows) > 1 else None
    )
    (
        multi_campaign_vins,
        duplicate_vin_campaigns,
        duplicate_vin_rows,
        duplicate_status_conflicts,
    ) = _vin_integrity_summary(records)
    unique_open_vins = {
        record.get("VIN", "").strip()
        for record in open_records
        if record.get("VIN", "").strip()
    }

    story: list[Any] = [
        Paragraph(
            f"Global Service Campaign Report — {_paragraph_text(report_zone)}",
            styles["ReportTitle"],
        ),
        Paragraph("Executive summary and review priorities", styles["Subtitle"]),
        Paragraph(
            f"ZONE PERFORMANCE  /  {_paragraph_text(report_zone)}  /  "
            f"{report_date:%d %B %Y}",
            styles["Eyebrow"],
        ),
        Spacer(1, 3 * mm),
        _kpi_cards(
            [
                (
                    "Overall progress",
                    f"{progress:.1f}%",
                    f"{finished:,} completed out of {total:,}",
                    TEAL,
                ),
                (
                    "Remaining workload",
                    f"{total - finished:,}",
                    "records still to complete",
                    RED,
                ),
                (
                    "Dealerships with records below target",
                    str(len(stores_below_target)),
                    (
                        f"of {assigned_dealership_count} assigned dealerships; "
                        f"{max(0, assigned_dealership_count - len(store_records))} "
                        "have no campaign records"
                        if assigned_dealership_count is not None
                        else f"of {len(store_records)} dealerships"
                    ),
                    AMBER if stores_below_target else GREEN,
                ),
                (
                    "Campaigns below target",
                    str(len(campaigns_below_target)),
                    f"of {campaign_count} campaigns",
                    AMBER if campaigns_below_target else GREEN,
                ),
            ],
            styles,
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        Spacer(1, 3 * mm),
        Paragraph("Finished campaign records by dealership", styles["SectionTitle"]),
        StatusBarChart(
            Counter(
                {
                    name: _record_summary(items)[1]
                    for name, items in sorted(
                        store_records.items(),
                        key=lambda item: (
                            -_record_summary(item[1])[1],
                            item[0].casefold(),
                        ),
                    )[:10]
                }
            ),
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        Spacer(1, 3 * mm),
        Paragraph("Meeting takeaways", styles["SectionTitle"]),
    ]
    insight_lines = [
        f"Overall: {finished:,} of {total:,} records completed ({progress:.1f}%); "
        f"{total - finished:,} remain.",
        (
            f"Dealerships to review first: {top_store[0]} has "
            f"{len(top_store[1]) - _record_summary(top_store[1])[1]:,} open "
            f"({_record_summary(top_store[1])[2]:.1f}% complete)."
            if top_store
            else "All dealerships have met the reference target."
        ),
        (
            f"Largest campaign backlog: {top_campaign[0]} — "
            f"{top_campaign[1][0].get('Campaign Name', '').strip() or top_campaign[0]} has "
            f"{len(top_campaign[1]) - _record_summary(top_campaign[1])[1]:,} open "
            f"of {len(top_campaign[1])} cases."
            if top_campaign
            else "No campaign records are available."
        ),
        (
            f"Campaign/model overlap: {top_campaign_model[2]} open of "
            f"{top_campaign_model[1]} for "
            f"{top_campaign_model[0].text.replace('<br/>', ' / ').replace('<b>', '').replace('</b>', '')}."
            if top_campaign_model
            else "No campaign/model cluster is available."
        ),
        (
            f"Campaign type {top_campaign_type[0]}: "
            f"{len(_open_records(top_campaign_type[1]))} open of "
            f"{len(top_campaign_type[1])}; leading model {top_model[0]}: "
            f"{len(_open_records(top_model[1]))} open of {len(top_model[1])}."
            if top_campaign_type and top_model
            else "Campaign type and vehicle-model breakdowns are included below."
        ),
        (
            f"VIN check: {multi_campaign_vins:,} VINs appear in multiple campaigns; "
            f"{duplicate_vin_rows:,} extra row(s) repeat the same VIN and campaign "
            f"({duplicate_status_conflicts:,} with conflicting statuses)."
        ),
        "All findings are workload signals, not proof of dealership fault. "
        "Open cases lack job-order numbers and completion dates; verify source "
        "records before attributing cause or estimating delay.",
    ]
    story.extend(
        [
            _insight_panel(insight_lines, styles),
            Spacer(1, 7 * mm),
            Paragraph("Workload by status", styles["SectionTitle"]),
        ]
    )
    status_rows: list[list[Any]] = [["Status", "Records", "Share of total"]]
    for status, count in status_counts.most_common():
        status_rows.append(
            [
                _paragraph_text(status),
                count,
                f"{count / total * 100:.1f}%" if total else "0%",
            ]
        )
    story.extend(
        [
            _make_table(status_rows, [100 * mm, 35 * mm, 35 * mm]),
            PageBreak(),
            Paragraph(
                "Dealership progress — campaign records present",
                styles["ReportTitle"],
            ),
            Paragraph(
                f"Dealerships with campaign records. Progress bars are color-coded "
                f"against the {TARGET_PROGRESS}% reference target. "
                + (
                    f"{len(assigned_dealership_records)} assigned sites are listed "
                    "in the coverage table."
                    if assigned_dealership_records is not None
                    else ""
                ),
                styles["Subtitle"],
            ),
            _make_table(
                _store_rows(store_records, styles, total_open),
                [83 * mm, 20 * mm, 22 * mm, 20 * mm, 38 * mm, 27 * mm],
            ),
            Paragraph(
                f"Zone total: {finished:,} completed of {total:,} records "
                f"({progress:.1f}%). Records without a dealership: {unassigned:,}.",
                styles["BodyText"],
            ),
            *(
                [
                    PageBreak(),
                    Paragraph("Assigned dealership coverage", styles["ReportTitle"]),
                    Paragraph(
                        "This list includes every site from the After-Sales zone "
                        "file. Sites without matching campaign rows show zero; no "
                        "campaign records are inferred for them.",
                        styles["Subtitle"],
                    ),
                    _make_table(
                        _assigned_dealership_rows(assigned_dealership_records, styles),
                        [74 * mm, 24 * mm, 24 * mm, 27 * mm, 36 * mm, 45 * mm],
                    ),
                ]
                if assigned_dealership_records is not None
                else []
            ),
            PageBreak(),
            Paragraph("Dealerships to review", styles["ReportTitle"]),
            Paragraph(
                f"{len(stores_below_target)} of {len(store_records)} dealerships "
                f"with campaign records are below the {TARGET_PROGRESS}% target. "
                "Highest workloads are listed first.",
                styles["Subtitle"],
            ),
            _make_table(
                _priority_store_rows(store_records, styles, total_open),
                [78 * mm, 18 * mm, 20 * mm, 17 * mm, 35 * mm, 22 * mm, 20 * mm],
            ),
            PageBreak(),
            Paragraph("Campaigns to review", styles["ReportTitle"]),
            Paragraph(
                "Campaigns below target, ranked by remaining records.",
                styles["Subtitle"],
            ),
            _make_table(
                _priority_campaign_rows(records, styles),
                [27 * mm, 103 * mm, 17 * mm, 18 * mm, 18 * mm, 35 * mm, 25 * mm],
            ),
            PageBreak(),
            Paragraph("Campaign and process drivers", styles["ReportTitle"]),
            Paragraph(
                "Open cases grouped by campaign and campaign type. Completion bars show "
                "the share of records already finished.",
                styles["Subtitle"],
            ),
            _make_table(
                _campaign_dimension_rows(records, styles),
                [105 * mm, 20 * mm, 24 * mm, 20 * mm, 35 * mm, 25 * mm],
            ),
            PageBreak(),
            Paragraph("Workload by campaign type", styles["ReportTitle"]),
            Paragraph(
                "Campaign-type groupings show where unresolved cases are concentrated.",
                styles["Subtitle"],
            ),
            Paragraph("By campaign type", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(records, "Campaign Type", styles),
                [120 * mm, 25 * mm, 28 * mm, 25 * mm, 40 * mm],
            ),
            Spacer(1, 4 * mm),
            Paragraph("Completed records by year and quarter", styles["SectionTitle"]),
            _make_table(
                _completion_period_rows(records),
                [50 * mm, 50 * mm, 40 * mm, 55 * mm],
                font_size=7,
            ),
            PageBreak(),
            Paragraph("Vehicle and record attributes", styles["ReportTitle"]),
            Paragraph(
                "Model, tier and delivery-status breakdowns of all open records. "
                "Small groups should be interpreted cautiously.",
                styles["Subtitle"],
            ),
            Paragraph("By vehicle model", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(records, "Model", styles),
                [120 * mm, 25 * mm, 28 * mm, 25 * mm, 40 * mm],
            ),
            PageBreak(),
            Paragraph("Delivery status and data completeness", styles["ReportTitle"]),
            Paragraph(
                "Operational data gaps limit how confidently the cause or waiting time "
                "of open cases can be assessed.",
                styles["Subtitle"],
            ),
            Paragraph("By delivery status", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(records, "Delivery Status", styles),
                [120 * mm, 25 * mm, 28 * mm, 25 * mm, 40 * mm],
            ),
            Spacer(1, 3 * mm),
            Paragraph("By campaign tier", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(
                    records, "Tier", styles, {"Not recorded": "Tier not recorded"}
                ),
                [120 * mm, 25 * mm, 28 * mm, 25 * mm, 40 * mm],
            ),
            Spacer(1, 3 * mm),
            Paragraph("Missing operational fields", styles["SectionTitle"]),
            _make_table(
                _data_quality_rows(records, styles),
                [105 * mm, 27 * mm, 35 * mm, 30 * mm],
                font_size=7,
            ),
            Paragraph(
                (
                    f"Open cases have no job-order number or completion date; "
                    f"operating store, in-store time, repair status and settlement "
                    f"time are blank zone-wide. {total_open:,} open campaign records "
                    f"cover {len(unique_open_vins):,} distinct VINs. "
                    f"{multi_campaign_vins:,} VINs appear in multiple campaigns; "
                    f"{duplicate_vin_campaigns:,} same-VIN/same-campaign duplicate "
                    f"key(s) add {duplicate_vin_rows:,} extra row(s). "
                    "Repeated VINs across different campaigns are not automatically "
                    "errors. Check source records before assigning cause or age."
                ),
                styles["SmallBody"],
            ),
            PageBreak(),
            Paragraph("Dealership and campaign hotspots", styles["ReportTitle"]),
            Paragraph(
                f"Top dealership/campaign combinations by open records. "
                f"Each combination shows its own completion rate and its share "
                f"of all {total_open} open zone records.",
                styles["Subtitle"],
            ),
            _make_table(
                _dealer_campaign_hotspot_rows(store_records, total_open, styles, limit=15),
                [76 * mm, 89 * mm, 18 * mm, 18 * mm, 35 * mm, 27 * mm],
                font_size=7,
            ),
            PageBreak(),
            Paragraph("Campaign and vehicle overlap", styles["ReportTitle"]),
            Paragraph(
                "Cross-tabulation of campaign and vehicle model. This isolates "
                "specific technical populations for eligibility and repair checks.",
                styles["Subtitle"],
            ),
            _make_table(
                _campaign_model_rows(records, styles, limit=15),
                [140 * mm, 25 * mm, 20 * mm, 35 * mm, 27 * mm],
                font_size=7,
            ),
            PageBreak(),
            Paragraph("Top 10 dealerships by completion rate", styles["ReportTitle"]),
            Paragraph(
                "Ranked by completion percentage, then by record volume. "
                "Case counts are shown so lower-volume results can be interpreted "
                "in context; ranking is descriptive, not a measure of repair quality.",
                styles["Subtitle"],
            ),
            _make_table(
                _top_dealership_rows(
                    {report_zone: records}, styles, limit=10
                ),
                [17 * mm, 80 * mm, 48 * mm, 24 * mm, 27 * mm, 22 * mm, 35 * mm],
                font_size=7,
            ),
            PageBreak(),
            Paragraph("Top 10 dealerships by finished cases", styles["ReportTitle"]),
            Paragraph(
                "Ranked by the number of records with Status = Finished. "
                "Record totals and completion rates are shown for context.",
                styles["Subtitle"],
            ),
            _make_table(
                _top_finished_dealership_rows(
                    {report_zone: records},
                    styles,
                    limit=10,
                    include_zone=False,
                ),
                [
                    14 * mm,
                    22 * mm,
                    80 * mm,
                    40 * mm,
                    27 * mm,
                    27 * mm,
                    35 * mm,
                ],
                font_size=7,
            ),
        ]
    )
    return story


def _store_story(
    records: list[dict[str, str]],
    store_name: str,
    styles: Any,
    report_date: date,
    report_zone: str = REPORT_ZONE,
) -> list[Any]:
    store_records = [
        record for record in records if record.get(STORE_FIELD, "").strip() == store_name
    ]
    if not store_records:
        raise ValueError(f"No records found for dealership '{store_name}'.")
    total, finished, progress = _record_summary(store_records)
    statuses = Counter(
        record.get(STATUS_FIELD, "").strip() or "Not specified"
        for record in store_records
    )
    campaigns = _group_campaigns(store_records)
    campaigns_below_target = sum(
        _record_summary(items)[2] < TARGET_PROGRESS for items in campaigns.values()
    )
    open_records = _open_records(store_records)
    models: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in store_records:
        models[record.get("Model", "").strip() or "Not recorded"].append(record)
    top_campaign = max(
        campaigns.items(),
        key=lambda item: len(_open_records(item[1])),
        default=None,
    )
    top_model = max(
        models.items(),
        key=lambda item: len(_open_records(item[1])),
        default=None,
    )
    unique_open_vins = {
        record.get("VIN", "").strip()
        for record in open_records
        if record.get("VIN", "").strip()
    }
    (
        multi_campaign_vins,
        duplicate_vin_campaigns,
        duplicate_vin_rows,
        duplicate_status_conflicts,
    ) = _vin_integrity_summary(store_records)
    story: list[Any] = [
        Paragraph("Dealership Progress Report", styles["ReportTitle"]),
        Paragraph(_paragraph_text(store_name), styles["Subtitle"]),
        Paragraph(
            f"ZONE PERFORMANCE  /  {_paragraph_text(report_zone)}  /  "
            f"{report_date:%d %B %Y}",
            styles["Eyebrow"],
        ),
        Spacer(1, 2 * mm),
        _kpi_cards(
            [
                (
                    "Total campaigns",
                    f"{total:,}",
                    "campaign records",
                    TEAL,
                ),
                (
                    "Finished",
                    f"{finished:,}",
                    "campaign records",
                    GREEN,
                ),
                (
                    "To be finished",
                    f"{total - finished:,}",
                    "campaign records",
                    RED,
                ),
                (
                    "Completion rate",
                    f"{progress:.1f}%",
                    f"{finished:,} of {total:,} finished",
                    _progress_color(progress),
                ),
            ],
            styles,
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        Spacer(1, 3 * mm),
        Paragraph("Campaign status", styles["SectionTitle"]),
        StatusBarChart(statuses, PAGE_WIDTH - 2 * PAGE_MARGIN),
        Spacer(1, 3 * mm),
        Paragraph("Status breakdown", styles["SectionTitle"]),
    ]
    status_rows: list[list[Any]] = [
        ["Status", "Total campaigns", "Share of dealership total"]
    ]
    for status, count in statuses.most_common():
        status_rows.append(
            [
                _paragraph_text(status),
                count,
                f"{count / total * 100:.1f}%",
            ]
        )
    story.extend(
        [
            _make_table(status_rows, [100 * mm, 35 * mm, 35 * mm]),
            Spacer(1, 3 * mm),
            Paragraph("Where to investigate first", styles["SectionTitle"]),
            _insight_panel(
                [
                    (
                        f"Campaign hotspot: {top_campaign[0]} — "
                        f"{top_campaign[1][0].get('Campaign Name', '').strip() or top_campaign[0]} "
                        f"has {len(_open_records(top_campaign[1]))} open of "
                        f"{len(top_campaign[1])} records."
                        if top_campaign and _open_records(top_campaign[1])
                        else "No open campaign records."
                    ),
                    (
                        f"Vehicle signal: {top_model[0]} has "
                        f"{len(_open_records(top_model[1]))} open records."
                        if top_model and _open_records(top_model[1])
                        else "No open vehicle-model records."
                    ),
                    f"{len(open_records)} open campaign records cover "
                    f"{len(unique_open_vins)} distinct VINs.",
                    (
                        f"VIN check: {multi_campaign_vins} VIN(s) appear in multiple "
                        f"campaigns; {duplicate_vin_rows} extra row(s) repeat the "
                        f"same VIN and campaign ({duplicate_status_conflicts} with "
                        "conflicting statuses). Repeated VINs across different "
                        "campaigns are not automatically errors."
                    ),
                    "These are observed workload patterns, not proof that this "
                    "dealership caused the issue.",
                    f"Data gap: {sum(not r.get('Job Order No.', '').strip() for r in open_records)} "
                    f"of {len(open_records)} open cases lack a job-order number, and "
                    f"{sum(not r.get('Completion Date', '').strip() for r in open_records)} "
                    "lack a completion date. Check source-system data before "
                    "attributing responsibility or estimating delay.",
                ],
                styles,
            ),
            PageBreak(),
            Paragraph("Campaigns driving open workload", styles["ReportTitle"]),
            Paragraph(
                "Campaign case counts, completion and open workload for this dealership.",
                styles["Subtitle"],
            ),
            _make_table(
                _campaign_dimension_rows(
                    store_records, styles, use_dealership_vocabulary=True
                ),
                [100 * mm, 30 * mm, 25 * mm, 30 * mm, 38 * mm, 28 * mm],
            ),
            Spacer(1, 4 * mm),
            Paragraph("By campaign type", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(
                    store_records,
                    "Campaign Type",
                    styles,
                    use_dealership_vocabulary=True,
                ),
                [105 * mm, 32 * mm, 25 * mm, 32 * mm, 39 * mm],
            ),
            PageBreak(),
            Paragraph("Vehicle and record detail", styles["ReportTitle"]),
            Paragraph(
                "Use vehicle and tier concentrations to narrow a technical review; "
                "confirm against campaign eligibility and workshop evidence.",
                styles["Subtitle"],
            ),
            Paragraph("By vehicle model", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(
                    store_records,
                    "Model",
                    styles,
                    use_dealership_vocabulary=True,
                ),
                [105 * mm, 32 * mm, 25 * mm, 32 * mm, 39 * mm],
            ),
            Spacer(1, 4 * mm),
            Paragraph("By campaign tier", styles["SectionTitle"]),
            _make_table(
                _dimension_rows(
                    store_records,
                    "Tier",
                    styles,
                    {"Not recorded": "Tier not recorded"},
                    use_dealership_vocabulary=True,
                ),
                [105 * mm, 32 * mm, 25 * mm, 32 * mm, 39 * mm],
            ),
            PageBreak(),
            Paragraph("Campaign-by-campaign progress", styles["ReportTitle"]),
            Paragraph(
                "Use this table to track completion and remaining workload.",
                styles["Subtitle"],
            ),
            _make_table(
                _campaign_rows(store_records, styles),
                [25 * mm, 115 * mm, 28 * mm, 25 * mm, 30 * mm, 38 * mm],
            ),
        ]
    )
    return story


def _group_story(
    records: list[dict[str, str]],
    group_name: str,
    styles: Any,
    report_date: date,
    report_zone: str = REPORT_ZONE,
) -> list[Any]:
    def matches_group(record: dict[str, str]) -> bool:
        record_group = _dealer_group_name(record.get(STORE_FIELD, "").strip())
        if record_group == "Group not specified":
            record_group = "Group not specified"
        return record_group.casefold() == group_name.casefold()

    group_records = [record for record in records if matches_group(record)]
    if not group_records:
        raise ValueError(f"No records found for dealership group '{group_name}'.")

    records_by_store: dict[str, list[dict[str, str]]] = defaultdict(list)
    for record in group_records:
        store = record.get(STORE_FIELD, "").strip()
        if store:
            records_by_store[store].append(record)

    rankings = []
    for store, store_records in records_by_store.items():
        total, completed, progress = _record_summary(store_records)
        rankings.append((store, store_records, total, completed, progress))
    rankings.sort(
        key=lambda item: (-item[4], -item[2], item[0].casefold())
    )

    total, completed, progress = _record_summary(group_records)
    campaigns = _group_campaigns(group_records)
    open_by_campaign = [
        (len(_open_records(items)), code, items)
        for code, items in campaigns.items()
        if _open_records(items)
    ]
    top_campaign = max(open_by_campaign, default=None)
    top_dealership = rankings[0] if rankings else None
    rows: list[list[Any]] = [
        [
            "Rank",
            "Dealership",
            "Records",
            "Completed",
            "Open",
            "Campaigns",
            "Completion",
        ]
    ]
    for rank, (store, store_records, store_total, store_completed, store_progress) in enumerate(
        rankings, start=1
    ):
        rows.append(
            [
                rank,
                Paragraph(_paragraph_text(store), styles["SmallBody"]),
                store_total,
                store_completed,
                store_total - store_completed,
                len(_group_campaigns(store_records)),
                ProgressCell(store_progress),
            ]
        )

    return [
        Paragraph("Dealership Group Performance Report", styles["ReportTitle"]),
        Paragraph(_paragraph_text(group_name), styles["Subtitle"]),
        Paragraph(
            f"GROUP PERFORMANCE  /  {_paragraph_text(report_zone)}  /  "
            f"{report_date:%d %B %Y}",
            styles["Eyebrow"],
        ),
        Spacer(1, 2 * mm),
        _kpi_cards(
            [
                (
                    "Group progress",
                    f"{progress:.1f}%",
                    f"{completed:,} completed of {total:,} records",
                    _progress_color(progress),
                ),
                (
                    "Dealerships",
                    f"{len(rankings):,}",
                    "ranked in this group",
                    TEAL,
                ),
                ("Completed", f"{completed:,}", "records closed", GREEN),
                ("Remaining", f"{total - completed:,}", "records to complete", RED),
            ],
            styles,
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        Spacer(1, 3 * mm),
        Paragraph("Finished campaign records by dealership", styles["SectionTitle"]),
        StatusBarChart(
            Counter(
                {
                    store: _record_summary(items)[1]
                    for store, items in sorted(
                        records_by_store.items(),
                        key=lambda item: (
                            -_record_summary(item[1])[1],
                            item[0].casefold(),
                        ),
                    )[:10]
                }
            ),
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        Spacer(1, 3 * mm),
        Paragraph("Dealership ranking", styles["SectionTitle"]),
        Paragraph(
            "All dealerships in this group are ranked by completion percentage "
            "from highest to lowest. Ties are ordered by record volume, then name. "
            "The full ranking includes every dealership, starting at rank 1.",
            styles["Subtitle"],
        ),
        _make_table(
            rows,
            [14 * mm, 104 * mm, 23 * mm, 25 * mm, 22 * mm, 25 * mm, 37 * mm],
            font_size=7,
        ),
        Spacer(1, 4 * mm),
        Paragraph("Group review priorities", styles["SectionTitle"]),
        _insight_panel(
            [
                (
                    f"Top dealership: {top_dealership[0]} leads with "
                    f"{top_dealership[4]:.1f}% completion "
                    f"({top_dealership[3]:,}/{top_dealership[2]:,} records)."
                    if top_dealership
                    else "No dealership records are available for ranking."
                ),
                (
                    f"Largest open campaign workload: {top_campaign[1]} has "
                    f"{top_campaign[0]:,} open of {len(top_campaign[2]):,} records."
                    if top_campaign
                    else "No open campaign records remain in this group."
                ),
                (
                    f"The ranking covers all {len(rankings):,} dealerships and "
                    f"{total:,} campaign records in the {group_name} group."
                ),
            ],
            styles,
        ),
    ]


def _france_overview_story(
    zones: dict[str, list[dict[str, str]]],
    unmapped_dealerships: dict[str, int],
    dealership_counts: dict[str, int],
    styles: Any,
    report_date: date,
) -> list[Any]:
    zone_summaries = {
        zone: _record_summary(records) for zone, records in zones.items() if records
    }
    french_records = [
        record for records in zones.values() for record in records
    ]
    if not french_records:
        raise ValueError("No assigned campaign records are available for France.")
    french_total, french_completed, french_progress = _record_summary(french_records)
    mapped_total = sum(summary[0] for summary in zone_summaries.values())
    unmapped_total = sum(unmapped_dealerships.values())
    (
        multi_campaign_vins,
        duplicate_vin_campaigns,
        duplicate_vin_rows,
        duplicate_status_conflicts,
    ) = _vin_integrity_summary(
        [record for zone_records in zones.values() for record in zone_records]
    )
    rows: list[list[Any]] = [
        ["Zone", "Dealerships", "Records", "Completed", "Open", "Completion"]
    ]
    for zone, records in zones.items():
        if not records:
            continue
        total, completed, progress = zone_summaries[zone]
        rows.append(
            [
                Paragraph(_paragraph_text(zone), styles["SmallBody"]),
                dealership_counts.get(zone, 0),
                total,
                completed,
                total - completed,
                ProgressCell(progress),
            ]
        )
    rows.append(
        [
            Paragraph("France total — Zones 1–4", styles["SmallBody"]),
            sum(dealership_counts.values()),
            french_total,
            french_completed,
            french_total - french_completed,
            ProgressCell(french_progress),
        ]
    )

    unmapped_detail = ""
    if unmapped_dealerships:
        examples = sorted(
            unmapped_dealerships.items(),
            key=lambda item: (-item[1], item[0].casefold()),
        )[:8]
        names = ", ".join(
            f"{name} ({count})" for name, count in examples
        )
        remaining = len(unmapped_dealerships) - len(examples)
        if remaining:
            names += f", and {remaining} other dealerships"
        unmapped_detail = (
            f" Not mapped to Zones 1–4: {unmapped_total:,} campaign records across "
            f"{len(unmapped_dealerships)} dealerships ({names})."
        )
    unmapped_rows: list[list[Any]] = [
        ["Dealership name in campaign file", "Records"]
    ]
    unmapped_rows.extend(
        [
            Paragraph(_paragraph_text(name), styles["SmallBody"]),
            count,
        ]
        for name, count in sorted(
            unmapped_dealerships.items(),
            key=lambda item: (-item[1], item[0].casefold()),
        )
    )

    return [
        Paragraph("France Service Campaign Overview", styles["ReportTitle"]),
        Paragraph(
            "Campaign completion across French dealerships assigned to Zones 1–4.",
            styles["Subtitle"],
        ),
        Paragraph(
            f"FRANCE OVERVIEW  /  ZONES 1–4  /  {report_date:%d %B %Y}",
            styles["Eyebrow"],
        ),
        Spacer(1, 3 * mm),
        _kpi_cards(
            [
                (
                    "France completion rate",
                    f"{french_progress:.1f}%",
                    f"{french_completed:,} Finished of {french_total:,}",
                    _progress_color(french_progress),
                ),
                (
                    "French campaign records",
                    f"{french_total:,}",
                    "across Zones 1–4",
                    TEAL,
                ),
                (
                    "Finished",
                    f"{french_completed:,}",
                    f"{french_total - french_completed:,} to be finished",
                    GREEN,
                ),
                (
                    "Assigned dealerships",
                    f"{sum(dealership_counts.values()):,}",
                    "across four French zones",
                    TEAL,
                ),
            ],
            styles,
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        Spacer(1, 4 * mm),
        _insight_panel(
            [
                (
                    f"Comparison covers {len(zones)} assigned zones and "
                    f"{mapped_total:,} campaign records. The zone labels and "
                    "dealership assignments come from the zone-management workbook."
                ),
                (
                    "Campaign records are counted as rows; the same vehicle can "
                    "appear in more than one campaign."
                ),
                (
                    "This is a descriptive comparison, not an adjusted performance "
                    "ranking. Campaign mix, workload, and reporting completeness "
                    "can differ between zones."
                )
                + unmapped_detail,
                (
                    f"France total includes records assigned to Zones 1–4 only. "
                    f"{unmapped_total:,} records without a French zone assignment "
                    "are excluded."
                ),
                (
                    f"VIN check: {multi_campaign_vins:,} VINs appear in multiple "
                    f"campaigns. There are {duplicate_vin_campaigns:,} "
                    f"same-VIN/same-campaign duplicate key(s), adding "
                    f"{duplicate_vin_rows:,} extra row(s); repeated VINs across "
                    f"different campaigns are not automatically errors. "
                    f"Status conflicts on duplicate campaign rows: "
                    f"{duplicate_status_conflicts:,}."
                ),
            ],
            styles,
        ),
        Spacer(1, 5 * mm),
        Paragraph("Completion by zone", styles["SectionTitle"]),
        _make_table(rows, [76 * mm, 30 * mm, 31 * mm, 34 * mm, 27 * mm, 38 * mm]),
        Spacer(1, 4 * mm),
        Paragraph("Finished campaign records by zone", styles["SectionTitle"]),
        StatusBarChart(
            Counter(
                {
                    zone: summary[1]
                    for zone, summary in zone_summaries.items()
                }
            ),
            PAGE_WIDTH - 2 * PAGE_MARGIN,
        ),
        *(
            [
                PageBreak(),
                Paragraph(
                    "Campaign records not mapped to Zones 1–4",
                    styles["ReportTitle"],
                ),
                Paragraph(
                    "These campaign rows do not match a dealership assignment in "
                    "the Zone 1–4 workbook. This includes virtual stores, locations "
                    "outside the assigned zones, and rows with no dealership name.",
                    styles["Subtitle"],
                ),
                _make_table(unmapped_rows, [150 * mm, 35 * mm]),
            ]
            if unmapped_dealerships
            else []
        ),
        PageBreak(),
        Paragraph("Zone ranking by finished cases", styles["ReportTitle"]),
        Paragraph(
            "Zones are ranked by total Finished records. The leading dealership "
            "in each zone is shown; higher Finished counts rank higher.",
            styles["Subtitle"],
        ),
        _make_table(
            _top_zone_finished_rows(zones, styles, dealership_counts),
            [14 * mm, 22 * mm, 25 * mm, 60 * mm, 27 * mm, 25 * mm, 30 * mm, 37 * mm],
            font_size=7,
            highlight_top=True,
        ),
        PageBreak(),
        Paragraph("Top 10 dealerships across all assigned zones", styles["ReportTitle"]),
        Paragraph(
            "Ranked by completion percentage, then by record volume. "
            "Case counts and zone names are shown for context; lower-volume "
            "dealerships may rank highly with fewer cases. This is a descriptive "
            "ranking, not a measure of repair quality.",
            styles["Subtitle"],
        ),
        _make_table(
            _top_dealership_rows(zones, styles, limit=10, include_zone=True),
            [
                14 * mm,
                20 * mm,
                82 * mm,
                42 * mm,
                20 * mm,
                23 * mm,
                18 * mm,
                35 * mm,
            ],
            font_size=7,
            highlight_top=True,
        ),
        PageBreak(),
        Paragraph("Top 10 dealerships by finished cases", styles["ReportTitle"]),
        Paragraph(
            "Ranked by the number of records with Status = Finished. "
            "Record totals and completion rates are shown for context.",
            styles["Subtitle"],
        ),
        _make_table(
            _top_finished_dealership_rows(zones, styles, limit=10),
            [
                14 * mm,
                22 * mm,
                64 * mm,
                44 * mm,
                22 * mm,
                22 * mm,
                22 * mm,
                34 * mm,
            ],
            font_size=7,
            highlight_top=True,
        ),
        Paragraph(
            "Group is taken from the dealer-group suffix in the concession name. "
            "“Group not specified” means the source name does not contain that "
            "group suffix.",
            styles["SmallBody"],
        ),
        PageBreak(),
        Paragraph("Top 10 dealership groups across all assigned zones", styles["ReportTitle"]),
        Paragraph(
            "Groups are ranked by combined completion percentage across their "
            "assigned dealerships, then by total record volume. Open workload and "
            "dealership counts are shown for context.",
            styles["Subtitle"],
        ),
        _make_table(
            _top_dealership_group_rows(zones, styles, limit=10),
            [18 * mm, 73 * mm, 29 * mm, 25 * mm, 27 * mm, 22 * mm, 39 * mm],
            font_size=7,
            highlight_top=True,
        ),
    ]


def _page_decorator(
    report_label: str, report_date: date, report_zone: str = REPORT_ZONE
) -> Any:
    report_date_label = report_date.strftime("%d %b %Y")

    def draw(canvas: Any, document: Any) -> None:
        canvas.saveState()
        canvas.setFillColor(TEAL)
        canvas.rect(0, PAGE_HEIGHT - 3 * mm, PAGE_WIDTH, 3 * mm, fill=1, stroke=0)
        canvas.setFillColor(NAVY)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawString(PAGE_MARGIN, PAGE_HEIGHT - 11 * mm, "SERVICE CAMPAIGN PERFORMANCE")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(MUTED)
        canvas.drawRightString(
            PAGE_WIDTH - PAGE_MARGIN,
            PAGE_HEIGHT - 11 * mm,
            (
                f"ZONES 1-4  |  {report_label.upper()}"
                if report_label == "France overview"
                else f"{report_zone.upper()}  |  {report_label.upper()}"
            ),
        )
        canvas.setStrokeColor(GRID)
        canvas.setLineWidth(0.6)
        canvas.line(PAGE_MARGIN, 13 * mm, PAGE_WIDTH - PAGE_MARGIN, 13 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(MUTED)
        canvas.drawString(
            PAGE_MARGIN, 8 * mm, f"Management report  |  {report_date_label}"
        )
        canvas.drawRightString(
            PAGE_WIDTH - PAGE_MARGIN,
            8 * mm,
            f"PAGE {document.page}",
        )
        canvas.restoreState()

    return draw


def create_pdf(
    records: list[dict[str, str]],
    output_path: Path,
    store_name: str | None = None,
    report_date: date | None = None,
    comparison_zones: dict[str, list[dict[str, str]]] | None = None,
    unmapped_dealerships: dict[str, int] | None = None,
    group_name: str | None = None,
    dealership_counts: dict[str, int] | None = None,
    assigned_dealership_records: dict[str, list[dict[str, str]]] | None = None,
    report_zone: str = REPORT_ZONE,
) -> None:
    report_date = report_date or date.today()
    page_size = landscape(A4)
    report_label = (
        f"Dealership - {store_name}"
        if store_name
        else f"Group - {group_name}"
        if group_name
        else "France overview"
        if comparison_zones is not None
        else "Global overview"
    )
    document = SimpleDocTemplate(
        str(output_path),
        pagesize=page_size,
        rightMargin=PAGE_MARGIN,
        leftMargin=PAGE_MARGIN,
        topMargin=20 * mm,
        bottomMargin=18 * mm,
        title=(
            f"Dealership Report - {store_name} - {report_zone}"
            if store_name
            else f"Dealership Group Report - {group_name} - {report_zone}"
            if group_name
            else "France Overview - Zones 1-4"
            if comparison_zones is not None
            else f"Global Report - {report_zone}"
        ),
        author="Dealership Report Generator",
    )
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="ReportTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=23,
            leading=27,
            textColor=NAVY,
            alignment=TA_LEFT,
            spaceAfter=3 * mm,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Subtitle",
            parent=styles["Normal"],
            alignment=TA_LEFT,
            fontSize=10,
            leading=14,
            textColor=MUTED,
            spaceAfter=3 * mm,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Eyebrow",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=10,
            textColor=TEAL,
            spaceAfter=2 * mm,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionTitle",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=16,
            textColor=NAVY,
            spaceBefore=2 * mm,
            spaceAfter=3 * mm,
        )
    )
    styles.add(
        ParagraphStyle(
            name="KpiLabel",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=9,
            textColor=MUTED,
        )
    )
    styles.add(
        ParagraphStyle(
            name="KpiValue",
            parent=styles["Normal"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=25,
            textColor=NAVY,
        )
    )
    styles.add(
        ParagraphStyle(
            name="KpiDetail",
            parent=styles["Normal"],
            fontSize=7,
            leading=9,
            textColor=MUTED,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InsightText",
            parent=styles["Normal"],
            fontSize=9,
            leading=14,
            textColor=INK,
            spaceAfter=2 * mm,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SmallBody",
            parent=styles["BodyText"],
            fontSize=7,
            leading=9,
        )
    )
    if comparison_zones is not None:
        story = _france_overview_story(
            comparison_zones,
            unmapped_dealerships or {},
            dealership_counts or {},
            styles,
            report_date,
        )
    elif store_name:
        story = _store_story(records, store_name, styles, report_date, report_zone)
    elif group_name:
        story = _group_story(records, group_name, styles, report_date, report_zone)
    else:
        story = _global_story(
            records,
            styles,
            report_date,
            assigned_dealership_count=(
                dealership_counts.get(report_zone)
                if dealership_counts is not None
                else None
            ),
            assigned_dealership_records=assigned_dealership_records,
            report_zone=report_zone,
        )
    decorator = _page_decorator(report_label, report_date, report_zone)
    document.build(story, onFirstPage=decorator, onLaterPages=decorator)


def _slugify(value: str) -> str:
    ascii_value = (
        unicodedata.normalize("NFKD", value)
        .encode("ascii", "ignore")
        .decode("ascii")
        .lower()
    )
    slug = re.sub(r"[^a-z0-9]+", "_", ascii_value).strip("_")
    return slug or "dealership"


def main() -> None:
    workbook_path = find_workbook()
    all_records = load_database(workbook_path)
    records = [
        record
        for record in all_records
        if record.get(ZONE_FIELD, "").strip().casefold() == REPORT_ZONE.casefold()
    ]
    if not records:
        available_zones = sorted(
            {
                record.get(ZONE_FIELD, "").strip()
                for record in all_records
                if record.get(ZONE_FIELD, "").strip()
            },
            key=str.casefold,
        )
        raise ValueError(
            f"No records found for zone '{REPORT_ZONE}'. "
            f"Zones present in the workbook: {', '.join(available_zones)}."
        )
    print(
        f"\nWorkbook loaded: {workbook_path.name} "
        f"({len(records)} records in {REPORT_ZONE}; "
        f"{len(all_records)} records total)"
    )
    print("\nWhich report would you like to generate?")
    print("  1. Global report")
    print("  2. Dealership report")

    while True:
        choice = input("Your choice (1 or 2): ").strip()
        if choice in {"1", "2"}:
            break
        print("Invalid choice. Enter 1 or 2.")

    store_name: str | None = None
    if choice == "2":
        stores = sorted(
            {
                record.get(STORE_FIELD, "").strip()
                for record in records
                if record.get(STORE_FIELD, "").strip()
            },
            key=str.casefold,
        )
        print("\nDealerships in your zone:")
        for index, store in enumerate(stores, start=1):
            print(f"  {index}. {store}")
        while True:
            selected = input("Dealership number: ").strip()
            if selected.isdigit() and 1 <= int(selected) <= len(stores):
                store_name = stores[int(selected) - 1]
                break
            print(f"Invalid choice. Enter a number from 1 to {len(stores)}.")

    output_directory = Path(__file__).resolve().parent / OUTPUT_DIRECTORY
    output_directory.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    zone_slug = _slugify(REPORT_ZONE)
    report_type = (
        f"{zone_slug}_concession_{_slugify(store_name)}"
        if store_name
        else f"{zone_slug}_global"
    )
    output_path = output_directory / f"report_{report_type}_{timestamp}.pdf"
    create_pdf(records, output_path, store_name)
    print(f"\nPDF created successfully: {output_path}")


if __name__ == "__main__":
    try:
        main()
    except (FileNotFoundError, ValueError) as error:
        raise SystemExit(f"Error: {error}") from error
