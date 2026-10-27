from __future__ import annotations

from collections import Counter
from datetime import datetime
from html import escape as escape_html
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import re
import zipfile
from urllib.parse import parse_qs, urlsplit
from xml.etree import ElementTree

import app as report_app


HOST = "127.0.0.1"
PORT = 8765
MAX_FORM_BYTES = 16_384
LIVE_SERVER_ORIGINS = {
    "http://127.0.0.1:5500",
    "http://localhost:5500",
}
OWN_ZONE_ID = "4"
CAMPAIGN_FIELDS = {
    report_app.STORE_FIELD,
    report_app.STATUS_FIELD,
    "Campaign Code",
    "Campaign Name",
}
ZONE_FIELDS = {"Zone", "Dealer Site"}
DEALERSHIP_MAP_LOCATIONS = {
    "Amadeus Cannes": (7.0152, 43.5454, "06"),
    "Amadeus Nice": (7.2528, 43.7032, "06"),
    "De Boussac Toulouse": (1.4328, 43.6007, "31"),
    "Deffeuille Aix-en-Provence": (5.3879, 43.5360, "13"),
    "Delenne Alès": (4.0905, 44.1250, "30"),
    "Etoile 83 Bandol": (5.7488, 43.1470, "83"),
    "Etoile 83 Toulon-Est": (5.9334, 43.1364, "83"),
    "GGA Arles": (4.6513, 43.5441, "13"),
    "Grim Avignon": (4.8333, 43.9416, "84"),
    "Grim Béziers": (3.2342, 43.3481, "34"),
    "Grim Montpellier": (3.8742, 43.6100, "34"),
    "Grim Nîmes": (4.3429, 43.8322, "30"),
    "Jean Lain Valence": (4.9164, 44.9234, "26"),
    "LG Brive-la-Gaillarde": (1.5144, 45.1450, "19"),
    "LG Carcassonne": (2.3491, 43.2078, "11"),
    "LG Castres": (2.2448, 43.6132, "81"),
    "LG Limoges": (1.2260, 45.8567, "87"),
    "LG Périgueux": (0.7105, 45.1939, "24"),
    "LG Perpignan": (2.9045, 42.6990, "66"),
    "Mapauto Fréjus": (6.7848, 43.4553, "83"),
    "Mazet Gap": (6.0616, 44.5797, "05"),
    "Milanesio Salon-de-Provence": (5.0490, 43.6430, "13"),
    "Negrier Rodez": (2.5699, 44.3591, "12"),
    "Parascandola Marseille": (5.3806, 43.2803, "13"),
    "Parascandola Vitrolles": (5.2656, 43.4509, "13"),
    "RVO Clermont-Ferrand": (3.1127, 45.7870, "63"),
    "SIA Biguglia": (9.4401, 42.6133, "2B"),
}
DEALER_SITE_ALIASES = {
    "lyon": "Jean Lain Vénissieux",
    "le cannet": "Amadeus Cannes",
    "bordeaux": "Edenauto Bruges",
    "nantes": "BPM Rezé",
    "paris 17": "Horizon Saint-Gratien",
    "roissy": "Bauer Saint-Witz",
    "bastia": "SIA Biguglia",
    "le mans": "Thibault Le Mans",
    "rouen": "Hecquet Rouen",
    "troyes": "Froment Troyes",
    "chateauroux": "Périer Châteauroux",
    "rodez": "Negrier Rodez",
    "arcachon": "Edenauto Bruges",
    "colmar": "Car Avenue Mulhouse",
    "strasbourg": "Passion Strasbourg",
    "dijon": "Passion Dijon",
    "villefranche sur saone": "Thivolle Villefranche-sur-Saône",
}
ZoneData = tuple[
    dict[str, list[dict[str, str]]],
    dict[str, int],
    str,
    dict[str, int],
    dict[str, dict[str, list[dict[str, str]]]],
]
_ZONE_DATA_CACHE: ZoneData | None = None
_ZONE_DATA_CACHE_SIGNATURE: tuple[tuple[str, int, int], ...] | None = None


def _page() -> bytes:
    template_path = Path(__file__).resolve().parent / "index.html"
    return template_path.read_bytes()


def _stylesheet() -> bytes:
    stylesheet_path = Path(__file__).resolve().parent / "style.css"
    return stylesheet_path.read_bytes()


def _tokens(value: str) -> set[str]:
    return set(report_app._slugify(value).split("_")) - {
        "xpeng",
        "service",
        "center",
        "car",
        "avenue",
        "automobile",
        "automobiles",
        "auto",
        "autos",
        "france",
        "group",
        "x",
    }


def _zone_id_for_store(
    store_name: str,
    assignments: list[dict[str, str]],
) -> str | None:
    normalized_store = report_app._slugify(store_name).replace("_", " ")
    assignment_by_site = {
        report_app._slugify(assignment["Dealer Site"]): assignment["Zone"].strip()
        for assignment in assignments
    }
    for location, dealer_site in DEALER_SITE_ALIASES.items():
        if location in normalized_store:
            return assignment_by_site.get(report_app._slugify(dealer_site))

    store_tokens = _tokens(store_name)
    matching_zones = {
        assignment["Zone"].strip()
        for assignment in assignments
        if len(_tokens(assignment["Dealer Site"])) > 1
        and _tokens(assignment["Dealer Site"]).issubset(store_tokens)
    }
    if len(matching_zones) == 1:
        return next(iter(matching_zones))

    location = re.sub(
        r"^xpeng_(service_center|service|car_avenue|motors)_?",
        "",
        report_app._slugify(store_name),
    ).replace("_", " ")
    location = re.split(r"\s+x\s+", location, maxsplit=1)[0]
    location_tokens = _tokens(location)
    location_zones = {
        assignment["Zone"].strip()
        for assignment in assignments
        if location_tokens
        and location_tokens.issubset(_tokens(assignment["Dealer Site"]))
    }
    return next(iter(location_zones)) if len(location_zones) == 1 else None


def _dealer_site_for_store(
    store_name: str,
    assignments: list[dict[str, str]],
) -> str | None:
    normalized_store = report_app._slugify(store_name).replace("_", " ")
    for location, dealer_site in DEALER_SITE_ALIASES.items():
        if location in normalized_store and any(
            assignment["Dealer Site"].strip() == dealer_site
            for assignment in assignments
        ):
            return dealer_site

    store_tokens = _tokens(store_name)
    matching_sites = {
        assignment["Dealer Site"].strip()
        for assignment in assignments
        if len(_tokens(assignment["Dealer Site"])) > 1
        and _tokens(assignment["Dealer Site"]).issubset(store_tokens)
    }
    if len(matching_sites) == 1:
        return next(iter(matching_sites))

    location = re.sub(
        r"^xpeng_(service_center|service|car_avenue|motors)_?",
        "",
        report_app._slugify(store_name),
    ).replace("_", " ")
    location = re.split(r"\s+x\s+", location, maxsplit=1)[0]
    location_tokens = _tokens(location)
    location_sites = {
        assignment["Dealer Site"].strip()
        for assignment in assignments
        if location_tokens
        and location_tokens.issubset(_tokens(assignment["Dealer Site"]))
    }
    return next(iter(location_sites)) if len(location_sites) == 1 else None


def _zone_label(zone_id: str) -> str:
    return f"Zone {zone_id}"


def _workbook_sheet_names(workbook_path: Path) -> list[str]:
    with zipfile.ZipFile(workbook_path) as workbook:
        workbook_xml = ElementTree.fromstring(workbook.read("xl/workbook.xml"))
    return [
        sheet.get("name", "")
        for sheet in workbook_xml.findall(
            "main:sheets/main:sheet", report_app.NAMESPACE
        )
    ]


def _find_workbook_data(
    project_directory: Path,
    required_fields: set[str],
    excluded_sheet_names: set[str] | None = None,
) -> tuple[Path, list[dict[str, str]]]:
    candidates: list[tuple[float, Path, list[dict[str, str]]]] = []
    excluded = {name.casefold() for name in excluded_sheet_names or set()}
    for workbook_path in project_directory.glob("*.xlsx"):
        if workbook_path.name.startswith("~$"):
            continue
        try:
            sheet_names = _workbook_sheet_names(workbook_path)
        except (
            OSError,
            KeyError,
            ValueError,
            zipfile.BadZipFile,
            ElementTree.ParseError,
        ):
            continue
        for sheet_name in sheet_names:
            if sheet_name.casefold() in excluded:
                continue
            try:
                records = report_app.load_database(
                    workbook_path,
                    sheet_name=sheet_name,
                    required_fields=required_fields,
                )
            except (OSError, ValueError):
                continue
            candidates.append((workbook_path.stat().st_mtime, workbook_path, records))
            break
    if not candidates:
        required = ", ".join(sorted(required_fields))
        raise FileNotFoundError(
            f"No Excel workbook with the required columns was found in "
            f"{project_directory}: {required}."
        )
    _, workbook_path, records = max(candidates, key=lambda candidate: candidate[0])
    return workbook_path, records


def _load_zone_data(
) -> tuple[
    dict[str, list[dict[str, str]]],
    dict[str, int],
    str,
    dict[str, int],
    dict[str, dict[str, list[dict[str, str]]]],
]:
    project_directory = Path(__file__).resolve().parent
    _, campaign_records = _find_workbook_data(
        project_directory,
        CAMPAIGN_FIELDS,
        excluded_sheet_names={"DASHBOARD"},
    )
    _, assignments = _find_workbook_data(
        project_directory,
        ZONE_FIELDS,
    )
    zone_ids = {
        assignment["Zone"].strip()
        for assignment in assignments
        if assignment["Zone"].strip()
    }
    dealership_counts = {
        _zone_label(zone_id): len(
            {
                assignment["Dealer Site"].strip()
                for assignment in assignments
                if assignment["Zone"].strip() == zone_id
                and assignment["Dealer Site"].strip()
            }
        )
        for zone_id in zone_ids
    }
    dealership_records_by_zone = {
        _zone_label(zone_id): {
            assignment["Dealer Site"].strip(): []
            for assignment in assignments
            if assignment["Zone"].strip() == zone_id
            and assignment["Dealer Site"].strip()
        }
        for zone_id in zone_ids
    }
    zone_by_site = {
        assignment["Dealer Site"].strip(): assignment["Zone"].strip()
        for assignment in assignments
        if assignment["Dealer Site"].strip()
    }

    if OWN_ZONE_ID not in zone_ids:
        raise ValueError(
            f"Zone {OWN_ZONE_ID} was not found in the dealership assignment workbook."
        )

    records_by_id: dict[str, list[dict[str, str]]] = {}
    unmapped: Counter[str] = Counter()
    for record in campaign_records:
        store_name = record.get(report_app.STORE_FIELD, "").strip()
        zone_id = _zone_id_for_store(store_name, assignments) if store_name else None
        if zone_id not in zone_ids:
            unmapped[store_name or "No dealership recorded"] += 1
            continue
        records_by_id.setdefault(zone_id, []).append(record)
        dealer_site = _dealer_site_for_store(store_name, assignments) if store_name else None
        if dealer_site and zone_by_site.get(dealer_site) == zone_id:
            dealership_records_by_zone[_zone_label(zone_id)][dealer_site].append(record)

    zones = {
        _zone_label(zone_id): records_by_id.get(zone_id, [])
        for zone_id in sorted(
            zone_ids,
            key=lambda value: (
                not value.isdigit(),
                int(value) if value.isdigit() else value,
            ),
        )
    }
    own_zone = _zone_label(OWN_ZONE_ID)
    if not zones.get(own_zone):
        raise ValueError(f"No campaign records were assigned to {own_zone}.")
    return zones, dict(unmapped), own_zone, dealership_counts, dealership_records_by_zone


def _workbook_data_signature() -> tuple[tuple[str, int, int], ...]:
    project_directory = Path(__file__).resolve().parent
    return tuple(
        sorted(
            (
                workbook.name,
                workbook.stat().st_mtime_ns,
                workbook.stat().st_size,
            )
            for workbook in project_directory.glob("*.xlsx")
            if not workbook.name.startswith("~$")
        )
    )


def _load_zone_data_cached() -> ZoneData:
    global _ZONE_DATA_CACHE, _ZONE_DATA_CACHE_SIGNATURE
    signature = _workbook_data_signature()
    if _ZONE_DATA_CACHE is None or signature != _ZONE_DATA_CACHE_SIGNATURE:
        data = _load_zone_data()
        _ZONE_DATA_CACHE = data
        _ZONE_DATA_CACHE_SIGNATURE = signature
    return _ZONE_DATA_CACHE


def _dealerships(records: list[dict[str, str]]) -> list[str]:
    return sorted(
        {
            record.get(report_app.STORE_FIELD, "").strip()
            for record in records
            if record.get(report_app.STORE_FIELD, "").strip()
        },
        key=str.casefold,
    )


def _dealership_options(records: list[dict[str, str]]) -> list[dict[str, str]]:
    groups_by_key: dict[str, str] = {}
    options = []
    for name in _dealerships(records):
        match = re.search(r"\s+x\s+(.+)$", name, flags=re.IGNORECASE)
        group = match.group(1).strip() if match else "Group not specified"
        group_key = group.casefold()
        display_group = groups_by_key.setdefault(group_key, group)
        options.append({"name": name, "group": display_group})
    return options


def _report_timestamp(filename: str) -> datetime | None:
    match = re.search(
        r"_(\d{8}_\d{6}(?:_\d{6})?)\.pdf$",
        filename,
    )
    if match is None:
        return None
    timestamp = match.group(1)
    for timestamp_format in ("%Y%m%d_%H%M%S_%f", "%Y%m%d_%H%M%S"):
        try:
            return datetime.strptime(timestamp, timestamp_format)
        except ValueError:
            continue
    return None


def _report_details(filename: str) -> tuple[str, str]:
    stem = filename.removesuffix(".pdf")
    timestamp = re.search(r"_\d{8}_\d{6}(?:_\d{6})?$", stem)
    if timestamp is None:
        return "Report", stem
    slug = stem[len("report_") : timestamp.start()].strip("_")
    if slug == "zones_comparison":
        return "France overview", "Zones 1–4"
    if slug.endswith("_global"):
        return "Zone report", slug[: -len("_global")].replace("_", " ").title()
    if "_concession_" in slug:
        zone, dealership = slug.split("_concession_", 1)
        return (
            "Dealership report",
            f"{zone.replace('_', ' ').title()} - "
            f"{dealership.replace('_', ' ').title()}",
        )
    if "_group_" in slug:
        zone, group = slug.split("_group_", 1)
        return (
            "Group report",
            f"{zone.replace('_', ' ').title()} - {group.replace('_', ' ').title()}",
        )
    return "Report", slug.replace("_", " ").title()


def _report_history_page(reports: list[dict[str, str]]) -> bytes:
    if reports:
        rows = "".join(
            "<tr>"
            f"<td>{escape_html(report['date'])}</td>"
            f"<td>{escape_html(report['type'])}</td>"
            f"<td>{escape_html(report['details'])}</td>"
            f"<td>{escape_html(report['filename'])}</td>"
            "<td class=\"report-actions\">"
            f"<a href=\"/reports/{escape_html(report['filename'])}\">Open</a>"
            f"<a href=\"/reports/{escape_html(report['filename'])}?download=1\">"
            "Download</a></td></tr>"
            for report in reports
        )
        table = (
            '<div class="table-wrap"><table class="dealership-table report-table">'
            "<thead><tr><th>Date</th><th>Type</th><th>Details</th>"
            "<th>File</th><th>Actions</th></tr></thead>"
            f"<tbody>{rows}</tbody></table></div>"
        )
    else:
        table = (
            '<section class="card"><p class="hint">No PDF reports have been saved '
            "yet. Reports will appear here after they are generated."
            "</p></section>"
        )
    template_path = Path(__file__).resolve().parent / "rapport.html"
    template = template_path.read_text(encoding="utf-8")
    return (
        template.replace(
            "{{REPORT_COUNT_TEXT}}",
            f"{len(reports)} saved report"
            f"{'s' if len(reports) != 1 else ''}",
        )
        .replace("{{REPORT_CONTENT}}", table)
        .encode("utf-8")
    )


class ReportHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True


class ReportRequestHandler(BaseHTTPRequestHandler):
    records: list[dict[str, str]]
    dealerships: list[str]
    dealership_options: list[dict[str, str]]
    comparison_zones: dict[str, list[dict[str, str]]]
    unmapped_dealerships: dict[str, int]
    own_zone: str
    dealership_counts: dict[str, int]
    dealership_records_by_zone: dict[str, dict[str, list[dict[str, str]]]]

    def _refresh_data(self) -> None:
        (
            zones,
            unmapped,
            own_zone,
            dealership_counts,
            dealership_records_by_zone,
        ) = _load_zone_data_cached()
        self.comparison_zones = zones
        self.unmapped_dealerships = unmapped
        self.own_zone = own_zone
        self.dealership_counts = dealership_counts
        self.dealership_records_by_zone = dealership_records_by_zone
        self.records = zones[own_zone]
        self.dealerships = _dealerships(self.records)
        self.dealership_options = _dealership_options(self.records)

    def do_GET(self) -> None:
        parsed_path = urlsplit(self.path)
        if parsed_path.path in {"/", "/index.html"}:
            self._send_html(_page())
            return
        if parsed_path.path in {"/ma-zone", "/ma-zone.html"}:
            page_path = Path(__file__).resolve().parent / "ma-zone.html"
            self._send_html(page_path.read_bytes())
            return
        if self.path == "/style.css":
            self._send_css(_stylesheet())
            return
        static_assets = {
            "/app.js": ("app.js", "text/javascript; charset=utf-8"),
            "/ma-zone.js": ("ma-zone.js", "text/javascript; charset=utf-8"),
            "/france-departments.json": (
                "france-departments.json",
                "application/json; charset=utf-8",
            ),
        }
        if parsed_path.path in static_assets:
            filename, content_type = static_assets[parsed_path.path]
            asset_path = Path(__file__).resolve().parent / filename
            try:
                content = asset_path.read_bytes()
            except OSError as error:
                self.log_error("Could not read site asset %s: %s", filename, error)
                self.send_error(500, "Could not load a site asset.")
                return
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(content)))
            self.send_header("X-Content-Type-Options", "nosniff")
            self._send_live_server_cors_headers()
            self.end_headers()
            self.wfile.write(content)
            return
        if parsed_path.path in {"/reports", "/rapport.html"}:
            report_directory = (
                Path(__file__).resolve().parent / report_app.OUTPUT_DIRECTORY
            )
            try:
                report_files = [
                    path
                    for path in report_directory.iterdir()
                    if path.is_file()
                    and re.fullmatch(
                        r"report_[a-z0-9_]+_\d{8}_\d{6}(?:_\d{6})?\.pdf",
                        path.name,
                    )
                ]
            except FileNotFoundError:
                report_files = []
            except OSError as error:
                self.log_error("Could not list saved reports: %s", error)
                self.send_error(500, "Could not read the saved reports.")
                return

            reports = []
            for path in report_files:
                generated_at = _report_timestamp(path.name)
                if generated_at is None:
                    continue
                report_type, details = _report_details(path.name)
                reports.append(
                    {
                        "filename": path.name,
                        "date": generated_at.strftime("%d/%m/%Y %H:%M"),
                        "type": report_type,
                        "details": details,
                    }
                )
            reports.sort(
                key=lambda report: _report_timestamp(report["filename"])
                or datetime.min,
                reverse=True,
            )
            try:
                page = _report_history_page(reports)
            except OSError as error:
                self.log_error("Could not read report history page: %s", error)
                self.send_error(500, "Could not read the report history page.")
                return
            self._send_html(page)
            return
        if parsed_path.path.startswith("/reports/"):
            filename = parsed_path.path.removeprefix("/reports/")
            if not re.fullmatch(
                r"report_[a-z0-9_]+_\d{8}_\d{6}(?:_\d{6})?\.pdf",
                filename,
            ):
                self.send_error(404, "Report not found.")
                return
            report_path = (
                Path(__file__).resolve().parent
                / report_app.OUTPUT_DIRECTORY
                / filename
            )
            if not report_path.is_file():
                self.send_error(404, "Report not found.")
                return
            query = parse_qs(parsed_path.query)
            action = "download" if query.get("download") == ["1"] else "view"
            self._send_pdf(report_path, action)
            return
        if parsed_path.path == "/zones":
            self._refresh_data()
            self._send_json(list(self.comparison_zones))
            return
        if parsed_path.path == "/my-zone-dealerships":
            self._refresh_data()
            assigned_sites = self.dealership_records_by_zone.get(self.own_zone, {})
            dealerships = []
            department_codes = set()
            for name, records in sorted(
                assigned_sites.items(), key=lambda item: item[0].casefold()
            ):
                total, finished, completion = report_app._record_summary(records)
                location = DEALERSHIP_MAP_LOCATIONS.get(name)
                if location:
                    longitude, latitude, department_code = location
                    department_codes.add(department_code)
                else:
                    longitude = latitude = department_code = None
                dealerships.append(
                    {
                        "name": name,
                        "records": total,
                        "finished": finished,
                        "open": total - finished,
                        "completion_rate": completion,
                        "longitude": longitude,
                        "latitude": latitude,
                        "department_code": department_code,
                    }
                )
            self._send_json(
                {
                    "zone": self.own_zone,
                    "dealerships": dealerships,
                    "department_codes": sorted(department_codes),
                }
            )
            return
        if parsed_path.path == "/dealerships":
            self._refresh_data()
            query = parse_qs(parsed_path.query)
            selected_zone = query.get("zone", [self.own_zone])[0]
            if selected_zone not in self.comparison_zones:
                self.send_error(400, "Choose a valid French zone.")
                return
            self.dealerships = _dealerships(self.comparison_zones[selected_zone])
            self.dealership_options = _dealership_options(
                self.comparison_zones[selected_zone]
            )
            self._send_json(self.dealership_options)
            return
        self.send_error(404, "Page not found")

    def do_POST(self) -> None:
        if self.path != "/report":
            self.send_error(404, "Page not found")
            return
        form: dict[str, list[str]] = {}
        try:
            self._refresh_data()
            length = int(self.headers.get("Content-Length", "0"))
            if length <= 0 or length > MAX_FORM_BYTES:
                raise ValueError("Invalid form submission.")
            body = self.rfile.read(length).decode("utf-8")
            form = parse_qs(body, keep_blank_values=True, strict_parsing=True)
            action = self._form_value(form, "action")
            report_type = self._form_value(form, "report_type")
            if action not in {"view", "download"}:
                raise ValueError("Choose whether to display or download the PDF.")
            if report_type not in {"global", "all_zones", "dealership", "group"}:
                raise ValueError("Choose a valid report type.")

            selected_zone = self.own_zone
            if report_type != "all_zones":
                selected_zone = self._form_value(form, "zone")
                if selected_zone not in self.comparison_zones:
                    raise ValueError("Choose a valid French zone.")
            selected_records = self.comparison_zones[selected_zone]
            selected_dealerships = _dealerships(selected_records)
            selected_options = _dealership_options(selected_records)
            store_name: str | None = None
            group_name: str | None = None
            if report_type == "dealership":
                store_name = self._form_value(form, "dealership")
                if store_name not in selected_dealerships:
                    raise ValueError("Choose a dealership from the list.")
            elif report_type == "group":
                submitted_group = self._form_value(form, "group_filter")
                groups_by_key = {
                    option["group"].casefold(): option["group"]
                    for option in selected_options
                }
                group_name = groups_by_key.get(submitted_group.casefold())
                if group_name is None:
                    raise ValueError("Choose a dealership group from the list.")

            if action == "view":
                self._send_html(
                    self._generate_preview(
                        report_type, store_name, group_name, selected_zone
                    )
                )
            else:
                output_path = self._generate_report(
                    report_type,
                    store_name,
                    group_name,
                    selected_zone,
                    selected_records,
                )
                self._send_pdf(output_path, action)
        except (ValueError, UnicodeDecodeError, KeyError) as error:
            self._send_json({"error": str(error)}, status=400)
        except (FileNotFoundError, OSError) as error:
            self.log_error("Report generation failed: %s", error)
            self._send_json({"error": f"Could not generate the report: {error}"}, status=500)
        except Exception as error:
            self.log_error("Unexpected report generation failure: %s", error)
            self._send_json(
                {
                    "error": (
                        "An unexpected error occurred while generating the report. "
                        "Check the terminal for details."
                    )
                },
                status=500,
            )

    def _generate_report(
        self,
        report_type: str,
        store_name: str | None,
        group_name: str | None = None,
        selected_zone: str | None = None,
        records: list[dict[str, str]] | None = None,
    ) -> Path:
        output_directory = Path(__file__).resolve().parent / report_app.OUTPUT_DIRECTORY
        output_directory.mkdir(exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        zone_slug = report_app._slugify(selected_zone or self.own_zone)
        report_slug = {
            "global": f"{zone_slug}_global",
            "all_zones": "zones_comparison",
            "dealership": (
                f"{zone_slug}_concession_{report_app._slugify(store_name or '')}"
            ),
            "group": (
                f"{zone_slug}_group_{report_app._slugify(group_name or '')}"
            ),
        }[report_type]
        filename = f"report_{report_slug}_{timestamp}.pdf"
        output_path = output_directory / filename
        report_app.create_pdf(
            records if records is not None else self.records,
            output_path,
            store_name=store_name if report_type == "dealership" else None,
            comparison_zones=(
                self.comparison_zones if report_type == "all_zones" else None
            ),
            unmapped_dealerships={},
            group_name=group_name if report_type == "group" else None,
            dealership_counts=self.dealership_counts,
            assigned_dealership_records=self.dealership_records_by_zone.get(
                selected_zone or self.own_zone, {}
            ),
            report_zone=selected_zone or self.own_zone,
        )
        return output_path

    def _generate_preview(
        self,
        report_type: str,
        store_name: str | None,
        group_name: str | None,
        selected_zone: str | None = None,
    ) -> bytes:
        selected_zone = selected_zone or self.own_zone
        records = self.comparison_zones[selected_zone]
        title = f"{selected_zone} Global Overview"
        subtitle = (
            f"Campaign progress for dealerships with records in {selected_zone}. "
            "The coverage table lists all assigned sites, including those with no "
            "campaign rows."
        )
        table_headers: list[str] = [
            "Rank",
            "Dealership",
            "Group",
            "Records",
            "Completed",
            "Open",
            "Completion",
        ]
        rows: list[list[str | int]] = []
        additional_tables: list[
            tuple[str, list[str], list[list[str | int]]]
        ] = []
        status_chart: list[tuple[str, int]] | None = None
        chart_title = "Campaign status"
        if report_type == "dealership":
            title = f"Dealership Report — {store_name or ''} — {selected_zone}"
            subtitle = (
                f"Campaign progress and remaining workload in {selected_zone}."
            )
            records = [
                record
                for record in records
                if record.get(report_app.STORE_FIELD, "").strip() == store_name
            ]
            status_counts = Counter(
                record.get(report_app.STATUS_FIELD, "").strip() or "Not specified"
                for record in records
            )
            status_chart = status_counts.most_common()
            campaigns = []
            for code, items in report_app._group_campaigns(records).items():
                total, completed, progress = report_app._record_summary(items)
                campaigns.append((code, items, total, completed, progress))
            campaigns.sort(
                key=lambda item: (-(item[2] - item[3]), item[0].casefold())
            )
            table_headers = [
                "Campaign",
                "Campaign name",
                "Total campaigns",
                "Finished",
                "To be finished",
                "Completion rate",
            ]
            rows = [
                [
                    code,
                    items[0].get("Campaign Name", "").strip() or code,
                    total,
                    completed,
                    total - completed,
                    f"{progress:.1f}%",
                ]
                for code, items, total, completed, progress in campaigns
            ]
        elif report_type == "all_zones":
            title = "France Overview — Zones 1–4"
            subtitle = (
                "Campaign progress for French dealerships assigned to Zones 1–4."
            )
            records = [
                record
                for zone_records in self.comparison_zones.values()
                for record in zone_records
            ]
            table_headers = [
                "Zone",
                "Dealerships",
                "Records",
                "Completed",
                "Open",
                "Completion",
            ]
            for zone, zone_records in self.comparison_zones.items():
                if not zone_records:
                    continue
                total, completed, progress = report_app._record_summary(zone_records)
                rows.append(
                    [
                        zone,
                        self.dealership_counts.get(
                            zone, len(_dealerships(zone_records))
                        ),
                        total,
                        completed,
                        total - completed,
                        f"{progress:.1f}%",
                    ]
                )
            status_chart = [
                (zone, report_app._record_summary(zone_records)[1])
                for zone, zone_records in self.comparison_zones.items()
                if zone_records
            ]
            chart_title = "Finished campaign records by zone"
            ranked_dealerships = report_app._rank_dealerships(
                self.comparison_zones
            )[:10]
            dealership_rows = [
                [
                    rank,
                    zone,
                    name,
                    group,
                    total,
                    completed,
                    total - completed,
                    f"{completion:.1f}%",
                ]
                for rank, (
                    zone,
                    name,
                    group,
                    total,
                    completed,
                    completion,
                ) in enumerate(ranked_dealerships, start=1)
            ]
            ranked_groups = report_app._rank_dealership_groups(
                self.comparison_zones
            )
            group_rows = [
                [
                    rank,
                    group,
                    dealerships_count,
                    total,
                    completed,
                    total - completed,
                    f"{group_progress:.1f}%",
                ]
                for rank, (
                    group,
                    dealerships_count,
                    total,
                    completed,
                    group_progress,
                ) in enumerate(ranked_groups[:10], start=1)
            ]
            ranked_finished = report_app._rank_dealerships_by_finished(
                self.comparison_zones
            )[:10]
            finished_rows = [
                [
                    rank,
                    zone,
                    name,
                    group,
                    total,
                    finished,
                    total - finished,
                    f"{completion:.1f}%",
                ]
                for rank, (
                    zone,
                    name,
                    group,
                    total,
                    finished,
                    completion,
                ) in enumerate(ranked_finished, start=1)
            ]
            ranked_zones = report_app._rank_zones_by_finished(
                self.comparison_zones
            )
            zone_finished_rows = [
                [
                    rank,
                    zone,
                    self.dealership_counts.get(zone, 0),
                    top_dealership,
                    total,
                    finished,
                    to_be_finished,
                    f"{completion:.1f}%",
                ]
                for rank, (
                    zone,
                    total,
                    finished,
                    to_be_finished,
                    completion,
                    top_dealership,
                ) in enumerate(ranked_zones, start=1)
            ]
            status_chart = [
                (zone, finished)
                for zone, _total, finished, _open, _rate, _top in ranked_zones
            ]
            additional_tables = [
                (
                    "Zone ranking by finished cases",
                    [
                        "Rank",
                        "Zone",
                        "Dealerships",
                        "Top dealership",
                        "Total campaigns",
                        "Finished",
                        "To be finished",
                        "Completion rate",
                    ],
                    zone_finished_rows,
                ),
                (
                    "Top 10 dealerships by finished cases",
                    [
                        "Rank",
                        "Zone",
                        "Dealership",
                        "Group",
                        "Total campaigns",
                        "Finished",
                        "To be finished",
                        "Completion rate",
                    ],
                    finished_rows,
                ),
                (
                    "Top 10 dealerships across all assigned zones",
                    [
                        "Rank",
                        "Zone",
                        "Dealership",
                        "Group",
                        "Records",
                        "Completed",
                        "Open",
                        "Completion",
                    ],
                    dealership_rows,
                ),
                (
                    "Top 10 dealership groups across all assigned zones",
                    [
                        "Rank",
                        "Group",
                        "Dealerships",
                        "Records",
                        "Completed",
                        "Open",
                        "Completion",
                    ],
                    group_rows,
                ),
            ]
        else:
            if report_type == "group":
                selected_group = group_name or ""
                title = (
                    f"Dealership Group Report — {selected_group} — {selected_zone}"
                )
                subtitle = (
                    f"Complete dealership ranking in {selected_zone}, by "
                    "completion percentage and then record volume."
                )
                group_records = [
                    record
                    for record in records
                    if _dealership_group(record.get(report_app.STORE_FIELD, ""))
                    .casefold()
                    == selected_group.casefold()
                ]
                if not group_records:
                    raise ValueError(
                        f"No records found for dealership group '{selected_group}'."
                    )
                records = group_records
            else:
                title = f"{selected_zone} Global Overview"
                subtitle = (
                    f"Campaign progress across dealerships assigned to {selected_zone}."
                )

            records_by_dealership: dict[str, list[dict[str, str]]] = {}
            for record in records:
                name = record.get(report_app.STORE_FIELD, "").strip()
                if name:
                    records_by_dealership.setdefault(name, []).append(record)
            ranking = []
            for name, items in records_by_dealership.items():
                total, completed, progress = report_app._record_summary(items)
                ranking.append((name, items, total, completed, progress))
            ranking.sort(
                key=lambda item: (-item[4], -item[2], item[0].casefold())
            )
            status_chart = [
                (name, finished)
                for name, _items, _total, finished, _progress in sorted(
                    ranking, key=lambda item: (-item[3], item[0].casefold())
                )[:10]
            ]
            chart_title = (
                "Finished campaign records by dealership"
                if report_type == "global"
                else "Finished campaign records by dealership in this group"
            )
            rows = [
                [
                    rank,
                    name,
                    _dealership_group(name),
                    total,
                    completed,
                    total - completed,
                    f"{progress:.1f}%",
                ]
                for rank, (name, _items, total, completed, progress) in enumerate(
                    ranking, start=1
                )
            ]
            finished_ranking = report_app._rank_dealerships_by_finished(
                {selected_zone: records}
            )[:10]
            additional_tables = [
                (
                    "Top 10 dealerships by finished cases",
                    [
                        "Rank",
                        "Dealership",
                        "Group",
                        "Total campaigns",
                        "Finished",
                        "To be finished",
                        "Completion rate",
                    ],
                    [
                        [
                            rank,
                            name,
                            group,
                            total,
                            finished,
                            total - finished,
                            f"{completion:.1f}%",
                        ]
                        for rank, (
                            _zone,
                            name,
                            group,
                            total,
                            finished,
                            completion,
                        ) in enumerate(finished_ranking, start=1)
                    ],
                )
            ]
            if report_type == "global":
                coverage_rows = []
                for site, site_records in sorted(
                    self.dealership_records_by_zone.get(selected_zone, {}).items(),
                    key=lambda item: item[0].casefold(),
                ):
                    site_total, site_finished, site_progress = (
                        report_app._record_summary(site_records)
                    )
                    coverage_rows.append(
                        [
                            site,
                            site_total,
                            site_finished,
                            site_total - site_finished,
                            f"{site_progress:.1f}%" if site_total else "—",
                            (
                                "Campaign records found"
                                if site_total
                                else "No campaign records"
                            ),
                        ]
                    )
                additional_tables.append(
                    (
                        f"All {len(coverage_rows)} assigned dealerships — campaign record coverage",
                        [
                            "Dealership site",
                            "Total campaigns",
                            "Finished",
                            "To be finished",
                            "Completion rate",
                            "Campaign data",
                        ],
                        coverage_rows,
                    )
                )

        total, completed, progress = report_app._record_summary(records)
        if report_type == "dealership":
            cards = [
                ("Total campaigns", f"{total:,}", "campaign records"),
                ("Finished", f"{completed:,}", "campaign records"),
                ("To be finished", f"{total - completed:,}", "campaign records"),
                ("Completion rate", f"{progress:.1f}%", "finished records"),
            ]
        else:
            cards = [
                ("Progress", f"{progress:.1f}%", f"{completed:,} of {total:,} records"),
                ("Completed", f"{completed:,}", "records closed"),
                ("Open", f"{total - completed:,}", "records to complete"),
            ]
        if report_type == "global":
            assigned_count = self.dealership_counts.get(selected_zone, 0)
            site_count = len(self.dealership_records_by_zone.get(selected_zone, {}))
            with_campaigns = sum(
                bool(site_records)
                for site_records in self.dealership_records_by_zone.get(
                    selected_zone, {}
                ).values()
            )
            cards.append(
                (
                    "Assigned dealerships",
                    str(assigned_count),
                    "from the After-Sales zone file",
                )
            )
            cards.extend(
                [
                    (
                        "With campaign records",
                        str(with_campaigns),
                        "matched to an assigned site",
                    ),
                    (
                        "Without campaign records",
                        str(max(0, site_count - with_campaigns)),
                        "assigned sites with no matching rows",
                    ),
                ]
            )
        if report_type == "group":
            cards.append(("Dealerships", str(len(rows)), "ranked in this group"))
        elif report_type == "dealership":
            cards.append(
                (
                    "Distinct campaigns",
                    str(len(rows)),
                    "different campaign codes",
                )
            )
        elif report_type == "all_zones":
            cards.append(("Scope", "France", "assigned Zones 1–4 only"))
        return _preview_html(
            title,
            subtitle,
            cards,
            table_headers,
            rows,
            report_type == "group",
            additional_tables,
            status_chart,
            chart_title,
        ).encode("utf-8")

    def _send_pdf(self, path: Path, action: str) -> None:
        disposition = "inline" if action == "view" else "attachment"
        filename = path.name
        content = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/pdf")
        self.send_header(
            "Content-Disposition",
            f'{disposition}; filename="{filename}"',
        )
        self.send_header("Content-Length", str(len(content)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self._send_live_server_cors_headers()
        self.end_headers()
        self.wfile.write(content)

    def _send_html(self, content: bytes, status: int = 200) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self._send_live_server_cors_headers()
        self.end_headers()
        self.wfile.write(content)

    def _send_css(self, content: bytes) -> None:
        self.send_response(200)
        self.send_header("Content-Type", "text/css; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self._send_live_server_cors_headers()
        self.end_headers()
        self.wfile.write(content)

    def _send_json(self, value: object, status: int = 200) -> None:
        content = json.dumps(value, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self._send_live_server_cors_headers()
        self.end_headers()
        self.wfile.write(content)

    def _send_live_server_cors_headers(self) -> None:
        origin = self.headers.get("Origin", "")
        if origin in LIVE_SERVER_ORIGINS:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Access-Control-Expose-Headers", "Content-Disposition")
            self.send_header("Vary", "Origin")

    def _form_value(self, form: dict[str, list[str]], key: str) -> str:
        values = form.get(key, [])
        if len(values) != 1:
            raise ValueError("Invalid form submission.")
        return values[0].strip()


def _dealership_group(store_name: str) -> str:
    match = re.search(r"\s+x\s+(.+)$", store_name, flags=re.IGNORECASE)
    return match.group(1).strip() if match else "Group not specified"


def _preview_html(
    title: str,
    subtitle: str,
    cards: list[tuple[str, str, str]],
    headers: list[str],
    rows: list[list[str | int]],
    highlight_ranking: bool,
    additional_tables: list[
        tuple[str, list[str], list[list[str | int]]]
    ] | None = None,
    status_chart: list[tuple[str, int]] | None = None,
    chart_title: str = "Campaign status",
) -> str:
    card_markup = "".join(
        (
            '<article class="metric">'
            f"<span>{escape_html(label)}</span>"
            f"<strong>{escape_html(value)}</strong>"
            f"<small>{escape_html(detail)}</small>"
            "</article>"
        )
        for label, value, detail in cards
    )
    header_markup = "".join(
        f"<th scope=\"col\">{escape_html(header)}</th>" for header in headers
    )
    def render_rows(
        table_rows: list[list[str | int]],
        highlight_top: bool,
    ) -> str:
        markup = []
        for row in table_rows:
            rank = row[0] if row and isinstance(row[0], int) else None
            row_class = (
                f' class="rank-{rank}"'
                if highlight_top and rank in {1, 2, 3}
                else ""
            )
            cells = "".join(
                f"<td>{escape_html(str(value))}</td>" for value in row
            )
            markup.append(f"<tr{row_class}>{cells}</tr>")
        return "".join(markup)

    row_markup = render_rows(rows, highlight_ranking)
    table_class = f"report-table columns-{len(headers)}"
    ranking_note = (
        f"<p class=\"note\">Full ranking: {len(rows)} dealership(s), "
        "from rank 1 to last, by completion rate.</p>"
        if highlight_ranking
        else ""
    )
    additional_markup = ""
    chart_markup = ""
    if status_chart:
        maximum = max(count for _, count in status_chart)
        total_statuses = sum(count for _, count in status_chart)
        chart_rows = "".join(
            (
                '<div class="chart-row">'
                f'<span class="chart-label">{escape_html(label)}</span>'
                '<div class="chart-track"><div class="chart-bar'
                + (' finished' if label.casefold() == "finished" else "")
                + f'" style="width:{count / maximum * 100:.1f}%"></div></div>'
                f'<strong>{count:,} ({count / total_statuses * 100:.1f}%)</strong>'
                "</div>"
            )
            for label, count in status_chart
        )
        chart_markup = (
            f'<section><h2>{escape_html(chart_title)}</h2>'
            f'<div class="status-chart">{chart_rows}</div></section>'
        )
    for section_title, section_headers, section_rows in additional_tables or []:
        section_header_markup = "".join(
            f"<th scope=\"col\">{escape_html(header)}</th>"
            for header in section_headers
        )
        section_row_markup = render_rows(section_rows, True)
        additional_markup += (
            '<section><h2>'
            + escape_html(section_title)
            + f'</h2><div class="table-wrap"><table class="report-table columns-{len(section_headers)}"><thead><tr>'
            + section_header_markup
            + "</tr></thead><tbody>"
            + section_row_markup
            + "</tbody></table></div></section>"
        )
    report_date = datetime.now().strftime("%d %B %Y")
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{escape_html(title)}</title>
  <style>
    :root {{ color-scheme: light; font-family: "Segoe UI", Arial, sans-serif; color: #192d43; }}
    body {{ margin: 0; padding: 24px; background: #f4f7fa; }}
    header {{ border-top: 5px solid #078b8d; padding: 20px 22px; background: #fff; border-radius: 8px; }}
    h1 {{ margin: 0 0 8px; font-size: 25px; }} p {{ margin: 6px 0; color: #536477; }}
    .date {{ color: #078b8d; font-size: 12px; font-weight: 700; text-transform: uppercase; }}
    .metrics {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin: 16px 0; }}
    .metric {{ display: flex; flex-direction: column; gap: 6px; padding: 16px; background: #fff; border: 1px solid #dce4eb; border-top: 3px solid #078b8d; border-radius: 7px; }}
    .metric span {{ color: #536477; font-size: 11px; font-weight: 700; text-transform: uppercase; }}
    .metric strong {{ color: #192d43; font-size: 25px; }} .metric small {{ color: #536477; }}
    section {{ margin-top: 16px; padding: 18px; background: #fff; border: 1px solid #dce4eb; border-radius: 8px; }}
    h2 {{ margin: 0 0 12px; font-size: 18px; }}
    .table-wrap {{ width: 100%; overflow-x: auto; }}
    table.report-table {{ width: 100%; min-width: 720px; table-layout: fixed; border-collapse: collapse; font-size: 13px; }}
    table.columns-7, table.columns-8 {{ min-width: 900px; }}
    table.columns-8 {{ min-width: 1040px; }}
    th {{ padding: 10px; text-align: left; color: #fff; background: #192d43; overflow-wrap: anywhere; }}
    td {{ padding: 9px 10px; border-bottom: 1px solid #e1e7ed; overflow-wrap: anywhere; vertical-align: top; }} tbody tr:nth-child(even) {{ background: #f4f7fa; }}
    tbody tr:first-child td:first-child {{ font-weight: 700; color: #078b8d; }}
    tbody tr.rank-1 {{ background: #fff1c2; font-weight: 700; }}
    tbody tr.rank-2 {{ background: #e8edf2; font-weight: 700; }}
    tbody tr.rank-3 {{ background: #f6e2d3; font-weight: 700; }}
    .note {{ margin-top: 12px; font-size: 12px; }}
    .status-chart {{ display: grid; gap: 12px; }}
    .chart-row {{ display: grid; grid-template-columns: minmax(120px, 190px) 1fr 110px; align-items: center; gap: 12px; }}
    .chart-label {{ font-size: 13px; }}
    .chart-track {{ height: 14px; overflow: hidden; border-radius: 7px; background: #eaf1f7; }}
    .chart-bar {{ height: 100%; border-radius: inherit; background: #007f86; }}
    .chart-bar.finished {{ background: #198754; }}
    .chart-row strong {{ font-size: 12px; text-align: right; }}
    @media (max-width: 600px) {{ body {{ padding: 12px; }} header, section {{ padding: 14px; }} table.report-table {{ min-width: 680px; }} }}
  </style>
</head>
<body>
  <header>
    <p class="date">Service campaign performance · {escape_html(report_date)}</p>
    <h1>{escape_html(title)}</h1>
    <p>{escape_html(subtitle)}</p>
  </header>
  <div class="metrics">{card_markup}</div>
  {chart_markup}
  <section>
    <h2>{"Full dealership ranking" if highlight_ranking else "Report details"}</h2>
    <div class="table-wrap"><table class="{table_class}"><thead><tr>{header_markup}</tr></thead>
      <tbody>{row_markup}</tbody>
    </table></div>
    {ranking_note}
  </section>
  {additional_markup}
</body>
</html>"""


def main() -> None:
    zones, unmapped, own_zone, dealership_counts, site_records = _load_zone_data_cached()
    records = zones[own_zone]
    dealerships = _dealerships(records)
    dealership_options = _dealership_options(records)
    handler = type(
        "ConfiguredReportRequestHandler",
        (ReportRequestHandler,),
        {
            "records": records,
            "dealerships": dealerships,
            "dealership_options": dealership_options,
            "comparison_zones": zones,
            "unmapped_dealerships": unmapped,
            "own_zone": own_zone,
            "dealership_counts": dealership_counts,
            "dealership_records_by_zone": site_records,
        },
    )
    try:
        server = ReportHTTPServer((HOST, PORT), handler)
    except OSError as error:
        if isinstance(error, OSError) and error.errno in {98, 10048, 10013}:
            raise SystemExit(
                f"Port {PORT} is already in use. "
                "Close the previous web_app.py server or stop the active Python "
                "process before restarting the report generator."
            ) from error
        raise

    print(
        f"Reports available for {report_app.REPORT_ZONE} "
        "and the assigned zones."
    )
    print(f"Open http://{HOST}:{PORT} in your browser.")
    print("To use Go Live, start this server before opening index.html.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
