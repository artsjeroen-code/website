import io
import re
import zipfile
from html import escape

FULL_REGISTRATION_FROM = "2027-01-01"
MONTHS = [
    ("01", "Januari"), ("02", "Februari"), ("03", "Maart"), ("04", "April"),
    ("05", "Mei"), ("06", "Juni"), ("07", "Juli"), ("08", "Augustus"),
    ("09", "September"), ("10", "Oktober"), ("11", "November"), ("12", "December"),
]


def xml_escape(value):
    return escape(str(value if value is not None else ""), quote=True)


def col_name(index):
    value = index + 1
    result = ""
    while value > 0:
        value -= 1
        result = chr(65 + (value % 26)) + result
        value //= 26
    return result


def text_cell(row, col, value, style=0):
    if value in ("", None):
        return ""
    ref = f"{col_name(col)}{row}"
    text = str(value)
    preserve = ' xml:space="preserve"' if re.search(r"^\s|\s$", text) else ""
    return f'<c r="{ref}" t="inlineStr" s="{style}"><is><t{preserve}>{xml_escape(text)}</t></is></c>'


def number_cell(row, col, value, style=0):
    if value in ("", None):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return ""
    if number.is_integer():
        number = int(number)
    return f'<c r="{col_name(col)}{row}" s="{style}"><v>{number}</v></c>'


def blank_cell(row, col, style=0):
    return f'<c r="{col_name(col)}{row}" s="{style}"/>'


def row_xml(row, cells, height=None):
    height_attrs = f' ht="{height}" customHeight="1"' if height else ""
    return f'<row r="{row}"{height_attrs}>{"".join(cells)}</row>'


def format_date(value):
    match = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", str(value or ""))
    return f"{match.group(3)}-{match.group(2)}-{match.group(1)}" if match else str(value or "")


def odometers(rides):
    values = []
    for ride in rides:
        for key in ("startOdometer", "endOdometer"):
            try:
                values.append(float(ride[key]))
            except (KeyError, TypeError, ValueError):
                pass
    return values


def month_stats(all_vehicle_rides, month_rides, vehicle, year, month):
    business_rides = [ride for ride in month_rides if ride.get("type") == "business"]
    business = sum(float(ride.get("distance") or 0) for ride in business_rides)
    month_odometers = odometers(month_rides)
    end_odometer = max(month_odometers) if month_odometers else None

    if f"{year}-{month}-01" >= FULL_REGISTRATION_FROM:
        private_km = sum(
            float(ride.get("distance") or 0)
            for ride in month_rides
            if ride.get("type") == "private"
        )
        return {
            "business": business,
            "privateKm": private_km,
            "endOdometer": end_odometer,
            "businessRides": business_rides,
        }

    if end_odometer is None:
        return {
            "business": business,
            "privateKm": 0,
            "endOdometer": None,
            "businessRides": business_rides,
        }

    month_start = f"{year}-{month}-01"
    previous_rides = [
        ride for ride in all_vehicle_rides
        if str(ride.get("date") or "") < month_start
    ]
    previous_odometers = odometers(previous_rides)
    if previous_odometers:
        baseline = max(previous_odometers)
    else:
        try:
            baseline = float(vehicle.get("initialOdometer"))
        except (TypeError, ValueError):
            baseline = None

    private_km = max(0, end_odometer - baseline - business) if baseline is not None else 0
    return {
        "business": business,
        "privateKm": private_km,
        "endOdometer": end_odometer,
        "businessRides": business_rides,
    }


def sheet_xml(vehicle, all_vehicle_rides, year, month):
    month_rides = sorted(
        [
            ride for ride in all_vehicle_rides
            if str(ride.get("date") or "").startswith(f"{year}-{month}-")
        ],
        key=lambda ride: (str(ride.get("date") or ""), int(ride.get("id") or 0)),
    )
    stats = month_stats(all_vehicle_rides, month_rides, vehicle, year, month)
    rows = []

    rows.append(row_xml(1, [text_cell(1, 0, "Rittenregistratie", 1)], 22))
    rows.append(row_xml(2, [
        text_cell(2, 0, "Jaartal", 2),
        text_cell(2, 2, year, 5),
        text_cell(2, 9, "Eind kilometerstand", 3),
    ]))
    rows.append(row_xml(3, [
        number_cell(
            3, 9,
            stats["endOdometer"] if stats["endOdometer"] is not None else 0,
            9,
        )
    ]))
    rows.append(row_xml(4, [
        text_cell(4, 0, "Merk", 2),
        text_cell(4, 2, "Type", 2),
        text_cell(4, 5, "Kenteken", 2),
    ]))
    rows.append(row_xml(5, [
        text_cell(5, 0, vehicle.get("make") or "", 5),
        text_cell(5, 2, vehicle.get("model") or "", 5),
        text_cell(5, 5, vehicle.get("plate") or "", 5),
        text_cell(5, 9, "Totaal zakelijk gereden", 3),
        text_cell(5, 11, "Prive", 3),
    ]))
    rows.append(row_xml(6, [
        text_cell(6, 0, "Personeelsnummer", 2),
        text_cell(6, 2, "Naam", 2),
        text_cell(6, 5, "Team manager", 2),
        number_cell(6, 9, stats["business"], 9),
        number_cell(6, 11, stats["privateKm"], 9),
    ]))
    rows.append(row_xml(7, [
        blank_cell(7, 0, 5),
        blank_cell(7, 2, 5),
        blank_cell(7, 5, 5),
    ]))
    rows.append(row_xml(8, [
        text_cell(8, 0, "Datum", 4),
        text_cell(8, 1, "Tijd dienstrit", 4),
        text_cell(8, 2, "Begin kilometerstand", 4),
        text_cell(8, 3, "Eind kilometerstand", 4),
        text_cell(8, 4, "Gereden zakelijke kilometers", 4),
        text_cell(8, 5, "Adres van vertrek", 4),
        text_cell(8, 6, "Adres van aankomst", 4),
        text_cell(8, 7, "Opmerking", 4),
    ], 58))

    row = 9
    for ride in stats["businessRides"]:
        rows.append(row_xml(row, [
            text_cell(row, 0, format_date(ride.get("date")), 5),
            text_cell(row, 1, str(ride.get("departureTime") or "")[:5], 8),
            number_cell(row, 2, ride.get("startOdometer"), 6),
            number_cell(row, 3, ride.get("endOdometer"), 6),
            number_cell(row, 4, ride.get("distance"), 6),
            text_cell(row, 5, ride.get("departureAddress") or "", 5),
            text_cell(row, 6, ride.get("arrivalAddress") or "", 5),
            text_cell(row, 7, ride.get("notes") or "", 5),
        ], 22))
        row += 1

    minimum_last_row = 58
    while row <= minimum_last_row:
        rows.append(row_xml(
            row,
            [blank_cell(row, col, 5) for col in range(8)],
            22,
        ))
        row += 1

    dimension_end = f"L{max(minimum_last_row, row - 1)}"
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:{dimension_end}"/><sheetViews><sheetView workbookViewId="0" showGridLines="1"><pane ySplit="8" topLeftCell="A9" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="2" width="17" customWidth="1"/><col min="3" max="3" width="18" customWidth="1"/><col min="4" max="4" width="17" customWidth="1"/><col min="5" max="5" width="18" customWidth="1"/><col min="6" max="6" width="14" customWidth="1"/><col min="7" max="7" width="32" customWidth="1"/><col min="8" max="8" width="18" customWidth="1"/><col min="9" max="9" width="4" customWidth="1"/><col min="10" max="10" width="20" customWidth="1"/><col min="11" max="11" width="4" customWidth="1"/><col min="12" max="12" width="13" customWidth="1"/></cols><sheetData>{''.join(rows)}</sheetData><mergeCells count="18"><mergeCell ref="A1:H1"/><mergeCell ref="A2:B2"/><mergeCell ref="C2:D2"/><mergeCell ref="A4:B4"/><mergeCell ref="C4:E4"/><mergeCell ref="F4:G4"/><mergeCell ref="A5:B5"/><mergeCell ref="C5:E5"/><mergeCell ref="F5:G5"/><mergeCell ref="A6:B6"/><mergeCell ref="C6:E6"/><mergeCell ref="F6:G6"/><mergeCell ref="A7:B7"/><mergeCell ref="C7:E7"/><mergeCell ref="F7:G7"/><mergeCell ref="J2:K2"/><mergeCell ref="J3:K3"/><mergeCell ref="J5:K5"/><mergeCell ref="J6:K6"/></mergeCells><pageMargins left="0.25" right="0.25" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0" paperSize="9"/></worksheet>'''


def styles_xml():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Aptos Narrow"/><family val="2"/></font>
    <font><b/><sz val="14"/><name val="Aptos Narrow"/><family val="2"/></font>
    <font><b/><sz val="11"/><name val="Aptos Narrow"/><family val="2"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Aptos Narrow"/><family val="2"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE7E6E6"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF000000"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="10">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="bottom"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def workbook_xml():
    sheets = "".join(
        f'<sheet name="{name}" sheetId="{index}" r:id="rId{index}"/>'
        for index, (_, name) in enumerate(MONTHS, 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" '
        'windowHeight="12000"/></bookViews>'
        f'<sheets>{sheets}</sheets><calcPr calcId="191029"/></workbook>'
    )


def workbook_rels_xml():
    sheets = "".join(
        f'<Relationship Id="rId{index}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, 13)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{sheets}<Relationship Id="rId13" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/></Relationships>'
    )


def content_types_xml():
    sheets = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, 13)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f'{sheets}</Types>'
    )


def root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )


def build_workbook(vehicle, rides, year):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml())
        workbook.writestr("_rels/.rels", root_rels_xml())
        workbook.writestr("xl/workbook.xml", workbook_xml())
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        workbook.writestr("xl/styles.xml", styles_xml())
        for index, (month, _) in enumerate(MONTHS, 1):
            workbook.writestr(
                f"xl/worksheets/sheet{index}.xml",
                sheet_xml(vehicle, rides, str(year), month),
            )
    return output.getvalue()


def cumulative_totals(vehicle, rides, year, through_month):
    business = 0
    private_km = 0
    ride_count = 0
    for month_number in range(1, through_month + 1):
        month = f"{month_number:02d}"
        month_rides = [
            ride for ride in rides
            if str(ride.get("date") or "").startswith(f"{year}-{month}-")
        ]
        stats = month_stats(rides, month_rides, vehicle, str(year), month)
        business += stats["business"]
        private_km += stats["privateKm"]
        ride_count += len(month_rides)
    return ride_count, business, private_km
