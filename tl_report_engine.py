from __future__ import annotations

import csv
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter


ROOT = Path(__file__).resolve().parents[2]
OUT_DIR = ROOT / "outputs" / "opex_profile_master_2026-04-17"
OUT_FILE = OUT_DIR / "TL_OPEX_Profile_Master_v2026-04-17.xlsx"

FINANCE_LINK = ROOT / "00_Company" / "Property_Registry" / "tl_finance_bei_linkage_initial_2026-04-17.csv"
PROPERTY_MASTER = ROOT / "00_Company" / "Property_Registry" / "tl_property_master_initial_2026-04-16.csv"
CATEGORY_DICT = ROOT / "02_Finance" / "Reports_and_Registers" / "opex_category_dictionary_v2_2026-04-17.csv"

NAVY = "1F3864"
TEAL = "1F6E6E"
GOLD = "C9A84C"
LIGHT = "EBF0F7"
PALE_YELLOW = "FFF2CC"
PALE_GREEN = "E2F0D9"
PALE_RED = "FCE4D6"
WHITE = "FFFFFF"
GREY = "D9E1F2"


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def as_int(value: str | None) -> int | None:
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except ValueError:
        return None


def profile_family(row: dict[str, str]) -> str:
    prop_type = (row.get("property_type") or "").strip() or "Property"
    bedrooms = as_int(row.get("bedrooms"))
    zone = (row.get("zone") or "").strip() or "Zone"
    if bedrooms:
        return f"{prop_type}_{bedrooms}BR_{zone}"
    return f"{prop_type}_{zone}"


def style_title(ws, title: str, subtitle: str | None = None) -> None:
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=16, color=WHITE)
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    if subtitle:
        ws["A2"] = subtitle
        ws["A2"].font = Font(italic=True, size=10, color="666666")
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)


def header_style(ws, row: int, fill: str = TEAL) -> None:
    for cell in ws[row]:
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="808080"))


def add_table(ws, name: str, start_row: int, end_row: int, end_col: int) -> None:
    ref = f"A{start_row}:{get_column_letter(end_col)}{end_row}"
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(tab)


def set_widths(ws, widths: dict[str, int]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def main() -> None:
    finance_rows = read_csv(FINANCE_LINK)
    pm_rows = {r["tlp_id"]: r for r in read_csv(PROPERTY_MASTER)}
    categories = read_csv(CATEGORY_DICT)

    properties = []
    for f in finance_rows:
        pm = pm_rows.get(f["tlp_id"], {})
        combined = {**pm, **f}
        profile_code = combined.get("pipeline_property_code") or combined.get("tlp_id") or f"ROW{len(properties) + 1:04d}"
        combined["profile_id"] = f"{profile_code}_OPEX_v2026-04"
        combined["profile_version"] = "v2026-04"
        combined["effective_from"] = "2026-04-01"
        combined["profile_status"] = "draft"
        combined["profile_family"] = profile_family(combined)
        if not combined.get("pipeline_property_code"):
            existing_flags = combined.get("qa_flags") or ""
            combined["qa_flags"] = "|".join([x for x in [existing_flags, "MISSING_PIPELINE_PROPERTY_CODE"] if x])
        properties.append(combined)

    core_categories = [c for c in categories if c["opex_passport_core_flag"].upper() == "TRUE"]
    wide_category_codes = [
        "FIX-POOL",
        "FIX-GARDEN",
        "MNT-MAIN",
        "FIX-INET",
        "FIX-COM",
        "FIX-PEST",
        "FIX-SEC",
        "FIX-INS",
        "CLN-REG",
        "CLN-DEEP",
        "CLN-DRY",
        "CLN-LNDRY",
        "CLN-CHEM",
        "UTL-ELEC",
        "UTL-WAT",
        "MNT-AC",
        "MNT-SEPTIC",
        "WASTE",
        "TAXES-PRP",
        "MNT-REPAIR",
        "EMRG",
        "ADJ",
    ]
    cat_by_code = {c["category_code"]: c for c in categories}

    wb = Workbook()
    ws = wb.active
    ws.title = "README_START_HERE"
    style_title(ws, "TL OPEX Profile Master", "Use Weekend_Fill_View first. Automation reads OPEX_Profile_Lines.")
    readme_rows = [
        ("What this workbook is", "One central source for OPEX Profile data. It is not 130 separate files."),
        ("Where to fill this weekend", "Use Weekend_Fill_View. One row = one property. Category columns = monthly THB budget/reserve."),
        ("What automation reads", "OPEX_Profile_Header and OPEX_Profile_Lines."),
        ("Important", "Do not use MISC as a budget line. Use a real category or ADJ for approved corrections."),
        ("Status flow", "draft -> finance_review -> approved_for_reporting -> owner_approved / superseded."),
        ("Scheduled services", "For annual/quarterly/3x-year services, enter the monthly reserve amount unless finance decides to book only in service month."),
        ("Regular cleaning", "Use CLN-REG. Deep cleaning is CLN-DEEP. Dry cleaning is CLN-DRY."),
        ("Legacy", "Old codes like VAR-CLEAN and FIX-CLEAN are aliases only. New OPEX Passports should use v2 codes."),
    ]
    ws.append(["Topic", "Instruction"])
    for row in readme_rows:
        ws.append(list(row))
    header_style(ws, 3)
    set_widths(ws, {"A": 28, "B": 110})
    for row in ws.iter_rows(min_row=4, max_col=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Weekend fill view
    ws = wb.create_sheet("Weekend_Fill_View")
    style_title(ws, "Weekend Fill View", "Fill monthly THB amounts in yellow category columns. One row = one property.")
    base_headers = [
        "profile_id",
        "tlp_id",
        "pipeline_property_code",
        "property_name",
        "owner_id",
        "owner_name",
        "owner_report_enabled",
        "profile_status",
        "profile_family",
        "property_type",
        "bedrooms",
        "zone",
    ]
    headers = base_headers + wide_category_codes + ["fill_notes", "qa_flags"]
    ws.append(headers)
    header_row = 3
    header_style(ws, header_row, NAVY)
    for p in properties:
        property_name = p.get("public_name") or p.get("internal_name") or p.get("property_name") or p.get("pipeline_property_code")
        row = [
            p.get("profile_id"),
            p.get("tlp_id"),
            p.get("pipeline_property_code"),
            property_name,
            p.get("owner_id"),
            p.get("owner_name"),
            p.get("owner_report_enabled"),
            p.get("profile_status"),
            p.get("profile_family"),
            p.get("property_type"),
            p.get("bedrooms"),
            p.get("zone"),
        ]
        row.extend(["" for _ in wide_category_codes])
        row.extend(["", p.get("qa_flags") or ""])
        ws.append(row)
    for col_idx, code in enumerate(wide_category_codes, start=len(base_headers) + 1):
        cell = ws.cell(header_row, col_idx)
        c = cat_by_code[code]
        cell.comment = Comment(
            f"{c['category_name_ru']} / {c['category_name_en']}\n"
            f"Default type: {c['default_cost_type']}\n"
            f"Minimum: {c['default_minimum_frequency']}\n"
            f"Enter monthly THB budget/reserve.",
            "Codex",
        )
        for row_idx in range(header_row + 1, header_row + 1 + len(properties)):
            ws.cell(row_idx, col_idx).fill = PatternFill("solid", fgColor=PALE_YELLOW)
            ws.cell(row_idx, col_idx).number_format = '#,##0'
    add_table(ws, "WeekendFillView", header_row, header_row + len(properties), len(headers))
    ws.freeze_panes = "M4"
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(len(headers))}{header_row + len(properties)}"
    set_widths(
        ws,
        {
            "A": 24,
            "B": 12,
            "C": 14,
            "D": 24,
            "E": 12,
            "F": 22,
            "G": 14,
            "H": 18,
            "I": 22,
            "J": 14,
            "K": 10,
            "L": 12,
        },
    )
    for idx in range(len(base_headers) + 1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(idx)].width = 14
    dv_status = DataValidation(type="list", formula1='"draft,finance_review,approved_for_reporting,owner_approved,superseded,blocked"', allow_blank=False)
    ws.add_data_validation(dv_status)
    dv_status.add(f"H4:H{header_row + len(properties)}")
    ws.conditional_formatting.add(
        f"A4:{get_column_letter(len(headers))}{header_row + len(properties)}",
        FormulaRule(formula=[f'$H4="approved_for_reporting"'], fill=PatternFill("solid", fgColor=PALE_GREEN)),
    )
    ws.conditional_formatting.add(
        f"A4:{get_column_letter(len(headers))}{header_row + len(properties)}",
        FormulaRule(formula=[f'$H4="blocked"'], fill=PatternFill("solid", fgColor=PALE_RED)),
    )

    # Header source table
    ws = wb.create_sheet("OPEX_Profile_Header")
    style_title(ws, "OPEX Profile Header", "One row per property profile version.")
    header_cols = [
        "profile_id",
        "tlp_id",
        "pipeline_property_code",
        "property_name",
        "owner_id",
        "owner_name",
        "profile_version",
        "effective_from",
        "effective_to",
        "profile_status",
        "profile_family",
        "property_type",
        "bedrooms",
        "zone",
        "prepared_by",
        "prepared_date",
        "reviewed_by",
        "reviewed_date",
        "approved_by",
        "approved_date",
        "owner_approval_status",
        "owner_approval_ref",
        "snapshot_link",
        "source_notes",
        "qa_flags",
    ]
    ws.append(header_cols)
    header_style(ws, 3, TEAL)
    for p in properties:
        property_name = p.get("public_name") or p.get("internal_name") or p.get("property_name") or p.get("pipeline_property_code")
        ws.append(
            [
                p.get("profile_id"),
                p.get("tlp_id"),
                p.get("pipeline_property_code"),
                property_name,
                p.get("owner_id"),
                p.get("owner_name"),
                p.get("profile_version"),
                p.get("effective_from"),
                "",
                p.get("profile_status"),
                p.get("profile_family"),
                p.get("property_type"),
                p.get("bedrooms"),
                p.get("zone"),
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                p.get("qa_flags") or "",
            ]
        )
    add_table(ws, "OPEXProfileHeader", 3, 3 + len(properties), len(header_cols))
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 24, "B": 12, "C": 14, "D": 24, "F": 22, "J": 18, "K": 22, "Y": 40})

    # Lines table, formula-driven from Weekend_Fill_View
    ws = wb.create_sheet("OPEX_Profile_Lines")
    style_title(ws, "OPEX Profile Lines", "Automation source. Budget amount formulas read Weekend_Fill_View.")
    line_cols = [
        "line_id",
        "profile_id",
        "tlp_id",
        "pipeline_property_code",
        "category_code",
        "category_name",
        "cost_type",
        "frequency",
        "unit",
        "quantity_basis",
        "unit_rate_thb",
        "budget_amount_monthly_thb",
        "budget_amount_annual_thb",
        "calculation_basis",
        "supplier_or_source",
        "evidence_link",
        "owner_chargeable_flag",
        "approval_required_flag",
        "approval_threshold_thb",
        "capex_candidate_threshold_thb",
        "effective_from",
        "effective_to",
        "line_status",
        "notes",
    ]
    ws.append(line_cols)
    header_style(ws, 3, TEAL)
    line_no = 1
    weekend_last_row = 3 + len(properties)
    for prop_idx, p in enumerate(properties, start=4):
        for code in wide_category_codes:
            c = cat_by_code[code]
            row_idx = ws.max_row + 1
            category_col = get_column_letter(headers.index(code) + 1)
            formula = f'=IFERROR(INDEX(Weekend_Fill_View!${category_col}$4:${category_col}${weekend_last_row},MATCH(B{row_idx},Weekend_Fill_View!$A$4:$A${weekend_last_row},0)),"")'
            annual_formula = f'=IF(L{row_idx}="","",L{row_idx}*12)'
            status_formula = f'=IF(L{row_idx}="","not_filled","draft")'
            ws.append(
                [
                    f"OPXL-{line_no:05d}",
                    p.get("profile_id"),
                    p.get("tlp_id"),
                    p.get("pipeline_property_code"),
                    code,
                    c["category_name_ru"],
                    c["default_cost_type"],
                    c["default_frequency"],
                    "",
                    c["default_minimum_frequency"],
                    "",
                    formula,
                    annual_formula,
                    "Fill monthly THB amount in Weekend_Fill_View",
                    "",
                    "",
                    c["owner_chargeable_default"],
                    c["approval_required_default"],
                    "",
                    "",
                    "2026-04-01",
                    "",
                    status_formula,
                    c["notes"],
                ]
            )
            line_no += 1
    add_table(ws, "OPEXProfileLines", 3, ws.max_row, len(line_cols))
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 14, "B": 24, "C": 12, "D": 14, "E": 13, "F": 28, "G": 20, "H": 14, "J": 18, "L": 18, "M": 18, "N": 34, "X": 52})
    for row in ws.iter_rows(min_row=4, min_col=12, max_col=13):
        for cell in row:
            cell.number_format = '#,##0'

    # Category dictionary
    ws = wb.create_sheet("Category_Dictionary")
    style_title(ws, "OPEX Category Dictionary v2", "Approved OPEX Passport categories and legacy non-core aliases.")
    cat_cols = list(categories[0].keys())
    ws.append(cat_cols)
    header_style(ws, 3, GOLD)
    for c in categories:
        ws.append([c.get(col, "") for col in cat_cols])
    add_table(ws, "OPEXCategoryDictionary", 3, 3 + len(categories), len(cat_cols))
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 16, "B": 28, "C": 28, "D": 24, "F": 22, "G": 18, "H": 22, "K": 18, "L": 70})
    for row in ws.iter_rows(min_row=4, max_col=len(cat_cols)):
        if str(row[4].value).upper() == "FALSE":
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=GREY)

    # Default profiles
    ws = wb.create_sheet("Default_Profiles")
    style_title(ws, "Default Profiles", "Seed profile families. Amounts are intentionally blank until finance approves defaults.")
    default_cols = [
        "default_profile_id",
        "profile_family",
        "property_type",
        "bedrooms",
        "zone",
        "residence",
        "category_code",
        "category_name",
        "cost_type",
        "frequency",
        "unit",
        "quantity_basis",
        "unit_rate_thb",
        "default_budget_amount_monthly_thb",
        "calculation_basis",
        "owner_chargeable_flag",
        "approval_required_flag",
        "approval_threshold_thb",
        "capex_candidate_threshold_thb",
        "notes",
    ]
    ws.append(default_cols)
    header_style(ws, 3, TEAL)
    families = [
        ("VILLA_2BR", "Villa", "2", ""),
        ("VILLA_3BR", "Villa", "3", ""),
        ("VILLA_4BR", "Villa", "4", ""),
        ("APARTMENT_1BR", "Apartment", "1", ""),
        ("APARTMENT_2BR", "Apartment", "2", ""),
        ("APARTMENT_3BR", "Apartment", "3", ""),
        ("CONDO", "Apartment", "", ""),
    ]
    default_codes = ["FIX-INET", "FIX-COM", "FIX-POOL", "FIX-GARDEN", "MNT-MAIN", "FIX-PEST", "FIX-SEC", "FIX-INS", "CLN-REG", "CLN-DEEP", "CLN-DRY", "CLN-LNDRY", "UTL-ELEC", "UTL-WAT", "MNT-AC", "MNT-SEPTIC", "WASTE", "TAXES-PRP"]
    default_no = 1
    for fam, prop_type, bedrooms, zone in families:
        for code in default_codes:
            c = cat_by_code[code]
            ws.append(
                [
                    f"DEF-{default_no:04d}",
                    fam,
                    prop_type,
                    bedrooms,
                    zone,
                    "",
                    code,
                    c["category_name_ru"],
                    c["default_cost_type"],
                    c["default_frequency"],
                    "",
                    c["default_minimum_frequency"],
                    "",
                    "",
                    "Fill after finance approves default rates",
                    c["owner_chargeable_default"],
                    c["approval_required_default"],
                    "",
                    "",
                    c["notes"],
                ]
            )
            default_no += 1
    add_table(ws, "OPEXDefaultProfiles", 3, ws.max_row, len(default_cols))
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 16, "B": 18, "C": 14, "G": 13, "H": 28, "I": 20, "J": 14, "L": 22, "N": 18, "O": 36, "T": 60})

    # Vendor rate cards
    ws = wb.create_sheet("Vendor_Rate_Cards")
    style_title(ws, "Vendor Rate Cards", "Optional source references for supplier rates.")
    vendor_cols = [
        "rate_card_id",
        "vendor_name",
        "service_category_code",
        "service_description",
        "unit",
        "unit_rate_thb",
        "effective_from",
        "effective_to",
        "evidence_link",
        "notes",
    ]
    ws.append(vendor_cols)
    header_style(ws, 3, GOLD)
    for i in range(1, 21):
        ws.append([f"RC-{i:04d}", "", "", "", "", "", "", "", "", ""])
    add_table(ws, "VendorRateCards", 3, ws.max_row, len(vendor_cols))
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 14, "B": 26, "C": 18, "D": 36, "E": 12, "F": 16, "I": 50, "J": 50})

    # QA dashboard
    ws = wb.create_sheet("QA_Dashboard")
    style_title(ws, "QA Dashboard", "Quick checks for weekend filling.")
    qa_rows = [
        ("Total properties in workbook", f"=COUNTA(OPEX_Profile_Header!A4:A{3 + len(properties)})"),
        ("Owner-report enabled TRUE", '=COUNTIF(OPEX_Profile_Header!J:J,"draft")'),
        ("Approved for reporting", '=COUNTIF(Weekend_Fill_View!H:H,"approved_for_reporting")'),
        ("Finance review", '=COUNTIF(Weekend_Fill_View!H:H,"finance_review")'),
        ("Draft", '=COUNTIF(Weekend_Fill_View!H:H,"draft")'),
        ("Blocked", '=COUNTIF(Weekend_Fill_View!H:H,"blocked")'),
        ("Profile line rows", f"=COUNTA(OPEX_Profile_Lines!A4:A{3 + len(properties) * len(wide_category_codes)})"),
        ("Filled budget line rows", f'=COUNTIF(OPEX_Profile_Lines!L4:L{3 + len(properties) * len(wide_category_codes)},">0")'),
        ("Blank budget line rows", f'=COUNTBLANK(OPEX_Profile_Lines!L4:L{3 + len(properties) * len(wide_category_codes)})'),
    ]
    ws.append(["Check", "Value"])
    header_style(ws, 3, NAVY)
    for row in qa_rows:
        ws.append(list(row))
    set_widths(ws, {"A": 34, "B": 18})
    for row in range(4, 4 + len(qa_rows)):
        ws.cell(row, 2).fill = PatternFill("solid", fgColor=PALE_YELLOW)

    # Change log
    ws = wb.create_sheet("Change_Log")
    style_title(ws, "Change Log", "Record meaningful changes and approvals.")
    change_cols = [
        "change_id",
        "change_date",
        "changed_by",
        "profile_id",
        "pipeline_property_code",
        "field_changed",
        "old_value",
        "new_value",
        "reason",
        "approved_by",
        "evidence_link",
        "notes",
    ]
    ws.append(change_cols)
    header_style(ws, 3, TEAL)
    for i in range(1, 21):
        ws.append([f"CHG-{i:04d}", "", "", "", "", "", "", "", "", "", "", ""])
    add_table(ws, "OPEXChangeLog", 3, ws.max_row, len(change_cols))
    ws.freeze_panes = "A4"
    set_widths(ws, {"A": 14, "B": 14, "C": 18, "D": 24, "E": 16, "F": 20, "G": 24, "H": 24, "I": 36, "J": 18, "K": 48, "L": 48})

    # Overall workbook formatting
    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = False

    # Friendly tab colors
    wb["README_START_HERE"].sheet_properties.tabColor = NAVY
    wb["Weekend_Fill_View"].sheet_properties.tabColor = GOLD
    wb["OPEX_Profile_Header"].sheet_properties.tabColor = TEAL
    wb["OPEX_Profile_Lines"].sheet_properties.tabColor = TEAL
    wb["Category_Dictionary"].sheet_properties.tabColor = GOLD
    wb["Default_Profiles"].sheet_properties.tabColor = GOLD
    wb["QA_Dashboard"].sheet_properties.tabColor = NAVY

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_FILE)
    print(OUT_FILE)


if __name__ == "__main__":
    main()
