"""
TropicLook Owner Report Engine — Universal Generator
FIN-REG-OWN-RPT-001 v1.0

USAGE:
    python3 tl_report_engine.py <input_template.xlsx> [output_dir]

    input_template.xlsx — заполненный бухгалтером INPUT TEMPLATE
    output_dir          — папка для сохранения (по умолчанию: текущая папка)

OUTPUT:
    TL_[CODE]_OwnerReport_[YYYY-MM]_v1.xlsx

ПРИНЦИПЫ (из регламента FIN-REG-OWN-RPT-001):
    - Доход признаётся на дату ВЫЕЗДА гостя (checkout date)
    - Комиссия TL берётся из таблицы (не пересчитывается)
    - OTA комиссии — только информационно, на баланс не влияют
    - 5 автоматических валидаций перед генерацией
"""

import sys
import os
import re
from datetime import datetime, date
from collections import defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ─── BRAND COLORS ─────────────────────────────────────────────────────────────
NAVY   = "1F3864"
TEAL   = "1F6E6E"
GOLD   = "C9A84C"
WHITE  = "FFFFFF"
LGRAY  = "F5F5F5"
DGRAY  = "E8E8E8"
YELLOW = "FFF9C4"
RED_BG = "FFEBEE"
RED_FG = "C62828"
GRN_BG = "E8F5E9"
GRN_FG = "1B5E20"
AMB_BG = "FFF8E1"
AMB_FG = "E65100"

# Category color map (bg, fg)
CAT_COLOR = {
    "REVENUE":    (NAVY,    WHITE),
    "UTL-RCHG":   ("37474F", WHITE),
    "OWNER-PAYIN":("1B5E20", WHITE),
    "MGMT-FEE":   (TEAL,    WHITE),
    "PAYOUT":     ("BF360C", WHITE),
    "UTL-ELEC":   (YELLOW,  "333333"),
    "UTL":        (YELLOW,  "333333"),
}
DEFAULT_CAT_COLOR = (LGRAY, "333333")

# ─── STYLE HELPERS ────────────────────────────────────────────────────────────
def _fill(h):
    return PatternFill("solid", start_color=h, fgColor=h)

def _bdr(color="CCCCCC"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def _align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def _font(bold=False, color="000000", size=9, italic=False, name="Arial"):
    return Font(name=name, bold=bold, color=color, size=size, italic=italic)

def _fmt_thb(val):
    """Format number as Thai Baht with comma separator."""
    if val is None:
        return ""
    try:
        return f"{float(val):,.0f}"
    except (ValueError, TypeError):
        return str(val)

def _to_float(val, default=0.0):
    """Safe float conversion."""
    if val is None:
        return default
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return default

def _to_int(val, default=0):
    try:
        return int(float(str(val).replace(",", "").strip()))
    except (ValueError, TypeError):
        return default

def _parse_date(val):
    """Parse date from various formats → date object or None."""
    if val is None:
        return None
    if isinstance(val, (date, datetime)):
        return val.date() if isinstance(val, datetime) else val
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

def _date_display(d):
    """date → 'DD.MM' string for Transaction Ledger."""
    if d is None:
        return ""
    return f"{d.day:02d}.{d.month:02d}"

def _apply_header_row(ws, row, cols_widths, bg=NAVY, fg=WHITE, height=20):
    """Write a styled header row. cols_widths = [(label, width), ...]"""
    ws.row_dimensions[row].height = height
    for ci, (label, w) in enumerate(cols_widths, 1):
        c = ws.cell(row=row, column=ci, value=label)
        c.font = _font(bold=True, color=fg, size=9)
        c.fill = _fill(bg)
        c.alignment = _align("center")
        c.border = _bdr()
        ws.column_dimensions[get_column_letter(ci)].width = w

def _cell(ws, row, col, value=None, bold=False, color="000000", size=9,
          bg=WHITE, align_h="left", num_fmt=None, border=True, italic=False):
    c = ws.cell(row=row, column=col, value=value)
    c.font = _font(bold=bold, color=color, size=size, italic=italic)
    c.fill = _fill(bg)
    c.alignment = _align(align_h)
    if border:
        c.border = _bdr()
    if num_fmt:
        c.number_format = num_fmt
    return c

def _section_hdr(ws, row, text, ncols, bg=TEAL, start_col=1):
    end_col = start_col + ncols - 1
    try:
        ws.merge_cells(start_row=row, start_column=start_col,
                       end_row=row, end_column=end_col)
    except Exception:
        pass
    c = ws.cell(row=row, column=start_col, value=text)
    c.font = _font(bold=True, color=WHITE, size=9)
    c.fill = _fill(bg)
    c.alignment = _align("left")
    ws.row_dimensions[row].height = 16

# ─── INPUT READER ─────────────────────────────────────────────────────────────
def read_kv_sheet(ws, key_col=1, val_col=3, start_row=1):
    """Read a key-value sheet (Property_Info, Prior_Period, etc.) → dict.
    Auto-detects start row by finding first row with a plain text key."""
    data = {}
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        k = row[key_col - 1] if len(row) >= key_col else None
        v = row[val_col - 1] if len(row) >= val_col else None
        if k and isinstance(k, str) and not k.startswith("▸") and \
                not k.startswith("КОНТРОЛЬ") and not k.startswith("ФОРМУЛА") and \
                not k.startswith("─") and not k.endswith("данные"):
            # Only store if key looks like a field name (no spaces or short)
            k_clean = k.strip()
            if k_clean and not any(x in k_clean for x in ["─","═","ЗАПОЛНИТЬ","Статичные"]):
                data[k_clean] = v
    return data

def read_row_sheet(ws, start_row=2, max_cols=10):
    """Read a row-based sheet (Reservations, Expenses, etc.) → list of tuples.
    start_row: first DATA row (after headers)."""
    rows = []
    for row in ws.iter_rows(min_row=start_row, values_only=True):
        if not any(c is not None for c in row[:max_cols]):
            continue
        first = str(row[0] or "").strip()
        if first.startswith("⚠") or first.startswith("ТИПЫ") or \
           first.startswith("✓") or first.startswith("booking_id") or \
           first.startswith("date") or first.startswith("Дата") or \
           first.startswith("ID"):
            continue  # skip header-like rows
        rows.append(row)
    return rows

class InputData:
    """Parsed and validated INPUT TEMPLATE data."""

    def __init__(self, path):
        self.path = path
        self.errors = []
        self.warnings = []
        self._load()

    def _load(self):
        wb = load_workbook(self.path, data_only=True)
        sheets = wb.sheetnames

        # ── Property_Info ──
        ws = wb["Property_Info"]
        pi = read_kv_sheet(ws, key_col=1, val_col=3, start_row=1)
        self.property_name    = str(pi.get("property_name", "Unknown Property")).strip()
        self.property_code    = str(pi.get("property_code", "PROP")).strip().upper()
        self.owner_name       = str(pi.get("owner_name", "Owner")).strip()
        self.owner_email      = str(pi.get("owner_email", "")).strip()
        self.report_period    = str(pi.get("report_period", "")).strip()   # YYYY-MM
        self.report_month_name= str(pi.get("report_month_name", "")).strip()
        self.property_type    = str(pi.get("property_type", "Villa")).strip()
        self.bedrooms         = _to_int(pi.get("bedrooms"), 0)
        self.location         = str(pi.get("location", "Phuket, Thailand")).strip()
        self.legal_entity     = str(pi.get("legal_entity", "TropicLook Group")).strip()
        self.commission_rate  = _to_float(pi.get("commission_rate"), 25.0)
        self.contract_start   = str(pi.get("contract_start", "")).strip()
        self.days_in_month    = _to_int(pi.get("days_in_month"), 30)
        self.available_nights = _to_int(pi.get("available_rooms"), self.bedrooms * self.days_in_month)
        self.opening_balance  = _to_float(pi.get("opening_balance"), 0.0)

        # Derive period label for display
        try:
            yr, mo = self.report_period.split("-")
            self.period_label = self.report_month_name or f"{yr}-{mo}"
        except Exception:
            self.period_label = self.report_period

        # ── Reservations ──
        ws_res = wb["Reservations"]
        raw_res = read_row_sheet(ws_res, start_row=2, max_cols=10)
        self.reservations = []
        for row in raw_res:
            if row[0] is None:
                continue
            def _r(i, d=None): return row[i] if len(row) > i else d
            r = {
                "booking_id":   str(_r(0) or "").strip(),
                "channel":      str(_r(1) or "").strip(),
                "checkin":      _parse_date(_r(2)),
                "checkout":     _parse_date(_r(3)),
                "nights":       _to_int(_r(4), 0),
                "gross":        _to_float(_r(5), 0),
                "ota_comm":     _to_float(_r(6), 0),
                "tl_comm":      _to_float(_r(7), 0),
                "net_to_owner": _to_float(_r(8), 0),
                "status":       str(_r(9) or "confirmed").strip(),
            }
            if r["booking_id"]:
                self.reservations.append(r)

        # ── Expenses ──
        ws_exp = wb["Expenses"]
        raw_exp = read_row_sheet(ws_exp, start_row=2, max_cols=8)
        self.expenses = []
        for row in raw_exp:
            # Skip fully empty rows (all yellow input rows without data)
            if row[0] is None and row[1] is None and row[4] is None:
                continue
            # Skip rows with no amount
            amt = _to_float(row[4], 0)
            if amt == 0:
                continue
            code = str(row[2] or "").strip().upper()
            if not code:
                self.errors.append(f"Расход без кода категории: '{row[1]}' / {row[4]} THB")
                continue
            e = {
                "date":          _parse_date(row[0]),
                "description":   str(row[1] or "").strip(),
                "category_code": code,
                "category_name": str(row[3] or code).strip(),
                "amount":        _to_float(row[4], 0),
                "booking_ref":   str(row[5] or "").strip() if len(row) > 5 else "",
                "receipt_ref":   str(row[6] or "").strip() if len(row) > 6 else "",
                "notes":         str(row[7] or "").strip() if len(row) > 7 else "",
            }
            if e["amount"] > 0:
                self.expenses.append(e)

        # ── Owner_Payouts ──
        ws_pay = wb["Owner_Payouts"]
        raw_pay = read_row_sheet(ws_pay, start_row=2, max_cols=6)
        self.payouts = []
        for row in raw_pay:
            if row[0] is None and row[1] is None:
                continue
            p = {
                "date":        _parse_date(row[0]),
                "amount":      _to_float(row[1], 0),
                "type":        str(row[2] or "Bank Transfer").strip(),
                "reference":   str(row[3] or "").strip(),
                "description": str(row[4] or "").strip(),
                "notes":       str(row[5] or "").strip() if len(row) > 5 else "",
            }
            if p["amount"] > 0:
                self.payouts.append(p)

        # ── OPEX_Budget ──
        ws_bud = wb["OPEX_Budget"]
        raw_bud = read_row_sheet(ws_bud, start_row=2, max_cols=5)
        self.opex_budget = {}
        for row in raw_bud:
            if row[0] is None:
                continue
            code = str(row[0] or "").strip().upper()
            def _rb(i, d=None): return row[i] if len(row) > i else d
            self.opex_budget[code] = {
                "name":   str(_rb(1) or code).strip(),
                "budget": _to_float(_rb(2), 0),
                "type":   str(_rb(3) or "").strip(),
                "notes":  str(_rb(4) or "").strip(),
            }

        # ── Prior_Period ──
        ws_prior = wb["Prior_Period"]
        pp = read_kv_sheet(ws_prior)
        self.prior = {
            "period":       str(pp.get("prior_period", "")).strip(),
            "gross":        _to_float(pp.get("gross_revenue"), 0),
            "mgmt_fee":     _to_float(pp.get("management_fee"), 0),
            "opex":         _to_float(pp.get("total_opex"), 0),
            "net_income":   _to_float(pp.get("net_income"), 0),
            "payouts":      _to_float(pp.get("total_payouts"), 0),
            "balance":      _to_float(pp.get("closing_balance"), 0),
            "occupancy":    _to_float(pp.get("occupancy_pct"), 0),
            "adr":          _to_float(pp.get("adr_thb"), 0),
            "bookings":     _to_int(pp.get("bookings_count"), 0),
            "nights":       _to_int(pp.get("nights_occupied"), 0),
        }

        # ── Cumulative ──
        ws_cum = wb["Cumulative"]
        cu = read_kv_sheet(ws_cum)
        self.cumulative = {
            "period":          str(cu.get("ytd_period", "")).strip(),
            "gross":           _to_float(cu.get("ytd_gross_revenue"), 0),
            "mgmt_fee":        _to_float(cu.get("ytd_management_fee"), 0),
            "opex":            _to_float(cu.get("ytd_total_opex"), 0),
            "net_income":      _to_float(cu.get("ytd_net_income"), 0),
            "payouts":         _to_float(cu.get("ytd_total_payouts"), 0),
            "bookings":        _to_int(cu.get("ytd_bookings"), 0),
            "nights":          _to_int(cu.get("ytd_nights_occupied"), 0),
            "available":       _to_int(cu.get("ytd_available_nights"), 0),
            "occupancy":       _to_float(cu.get("ytd_avg_occupancy"), 0),
            "adr":             _to_float(cu.get("ytd_avg_adr"), 0),
            "purchase_price":  _to_float(cu.get("purchase_price"), 0),
            "yield_pct":       _to_float(cu.get("ytd_yield_pct"), 0),
        }

        # ── Cash_Balance ──
        ws_cb = wb["Cash_Balance"]
        cb = read_kv_sheet(ws_cb)
        self.cash = {
            "opening":  _to_float(cb.get("opening_balance"), self.opening_balance),
            "income":   _to_float(cb.get("total_income"), 0),
            "expenses": _to_float(cb.get("total_expenses"), 0),
            "payouts":  _to_float(cb.get("total_payouts"), 0),
            "closing":  _to_float(cb.get("closing_balance"), 0),
        }

        wb.close()

    # ── COMPUTED AGGREGATES ────────────────────────────────────────────────────
    @property
    def total_gross(self):
        return sum(r["gross"] for r in self.reservations)

    @property
    def total_ota_comm(self):
        return sum(r["ota_comm"] for r in self.reservations)

    @property
    def total_tl_comm(self):
        return sum(r["tl_comm"] for r in self.reservations)

    @property
    def total_net_to_owner(self):
        return sum(r["net_to_owner"] for r in self.reservations)

    @property
    def total_opex(self):
        return sum(e["amount"] for e in self.expenses)

    @property
    def total_payouts(self):
        return sum(p["amount"] for p in self.payouts)

    @property
    def net_income(self):
        """Net Income = Gross - TL Commission - OPEX"""
        return self.total_gross - self.total_tl_comm - self.total_opex

    @property
    def nights_occupied(self):
        return sum(r["nights"] for r in self.reservations)

    @property
    def occupancy_pct(self):
        if self.available_nights == 0:
            return 0.0
        return round(self.nights_occupied / self.available_nights * 100, 1)

    @property
    def adr(self):
        if self.nights_occupied == 0:
            return 0.0
        return round(self.total_gross / self.nights_occupied, 0)

    @property
    def expense_ratio(self):
        if self.total_gross == 0:
            return 0.0
        return round(self.total_opex / self.total_gross * 100, 1)

    @property
    def closing_balance(self):
        return round(self.opening_balance + self.total_gross
                     - self.total_tl_comm - self.total_opex - self.total_payouts, 2)

    def opex_by_category(self):
        """Group expenses by category code → {code: {'name': ..., 'total': ...}}"""
        cats = defaultdict(lambda: {"name": "", "total": 0.0})
        for e in self.expenses:
            code = e["category_code"]
            cats[code]["name"] = e["category_name"] or code
            cats[code]["total"] += e["amount"]
        # Fill names from budget if missing
        for code, bud in self.opex_budget.items():
            if code in cats and not cats[code]["name"]:
                cats[code]["name"] = bud["name"]
        return dict(cats)

    # ── VALIDATION ────────────────────────────────────────────────────────────
    def validate(self):
        """Run all 5 regulatory validation rules. Returns (ok, errors, warnings)."""
        errors = list(self.errors)
        warnings = list(self.warnings)

        # RULE 1: Cash balance arithmetic
        calc_closing = round(self.opening_balance + self.total_gross
                             - self.total_tl_comm - self.total_opex
                             - self.total_payouts, 2)
        diff = abs(calc_closing - self.cash["closing"])
        if diff > 1.0:
            errors.append(
                f"ПРАВИЛО 1 — ОШИБКА БАЛАНСА: расчётный остаток {calc_closing:,.2f} ≠ "
                f"указанный {self.cash['closing']:,.2f} (расхождение {diff:,.2f} THB)"
            )

        # RULE 2: Closing balance ≥ 0
        if calc_closing < -0.01:
            errors.append(
                f"ПРАВИЛО 2 — ОТРИЦАТЕЛЬНЫЙ БАЛАНС: {calc_closing:,.2f} THB. "
                f"Выплаты превысили доступные средства."
            )

        # RULE 3: All expenses categorized (already checked during load)
        # (errors already added above)

        # RULE 4: Commission sanity (±5% tolerance)
        expected_comm = round(self.total_gross * self.commission_rate / 100, 2)
        if expected_comm > 0:
            comm_diff_pct = abs(self.total_tl_comm - expected_comm) / expected_comm * 100
            if comm_diff_pct > 5:
                warnings.append(
                    f"ПРАВИЛО 4 — Комиссия TL ({self.total_tl_comm:,.0f}) отличается "
                    f"от ожидаемой ({expected_comm:,.0f}) на {comm_diff_pct:.1f}%. "
                    f"Если ручная корректировка — добавьте примечание."
                )

        # RULE 5: At least some reservations
        if not self.reservations:
            warnings.append(
                "ПРАВИЛО 5 — В листе Reservations нет данных. "
                "Если месяц простойный — это допустимо."
            )

        ok = len(errors) == 0
        return ok, errors, warnings


# ─── REPORT BUILDER ───────────────────────────────────────────────────────────
class ReportBuilder:

    def __init__(self, data: InputData):
        self.d = data
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    # ── TAB 1: DASHBOARD ──────────────────────────────────────────────────────
    def build_dashboard(self):
        ws = self.wb.create_sheet("Dashboard")
        ws.sheet_properties.tabColor = NAVY
        ws.sheet_view.showGridLines = False
        d = self.d

        # Title block
        ws.merge_cells("B2:H2")
        c = ws["B2"]
        c.value = "TROPICLOOK — OWNER FINANCIAL REPORT"
        c.font = _font(bold=True, color=WHITE, size=14)
        c.fill = _fill(NAVY)
        c.alignment = _align("center")
        ws.row_dimensions[2].height = 28

        ws.merge_cells("B3:H3")
        c = ws["B3"]
        c.value = f"{d.property_name}  |  {d.period_label}  |  TL-{d.property_code}"
        c.font = _font(color=WHITE, size=10)
        c.fill = _fill(TEAL)
        c.alignment = _align("center")
        ws.row_dimensions[3].height = 18

        ws.merge_cells("B4:H4")
        c = ws["B4"]
        c.value = (f"Owner: {d.owner_name}  |  Commission: {d.commission_rate}%  |  "
                   f"{d.property_type} {d.bedrooms}BR  |  {d.location}")
        c.font = _font(color="DDDDDD", size=9, italic=True)
        c.fill = _fill(NAVY)
        c.alignment = _align("center")
        ws.row_dimensions[4].height = 16

        # KPI cards (row 6-8)
        ws.row_dimensions[6].height = 14
        kpis = [
            ("GROSS REVENUE",    f"{d.total_gross:,.0f} ฿",    "Доход от бронирований"),
            ("NET OWNER INCOME", f"{d.net_income:,.0f} ฿",     "После комиссии и OPEX"),
            ("CASH BALANCE",     f"{d.closing_balance:,.0f} ฿","Остаток на счёте"),
        ]
        kpi_cols = [2, 4, 6]
        for i, (title, val, sub) in enumerate(kpis):
            col = kpi_cols[i]
            ws.merge_cells(start_row=7, start_column=col, end_row=7, end_column=col+1)
            ws.merge_cells(start_row=8, start_column=col, end_row=8, end_column=col+1)
            ws.merge_cells(start_row=9, start_column=col, end_row=9, end_column=col+1)
            h = ws.cell(row=7, column=col, value=title)
            h.font = _font(bold=True, color=WHITE, size=9)
            h.fill = _fill(NAVY)
            h.alignment = _align("center")
            v = ws.cell(row=8, column=col, value=val)
            v.font = _font(bold=True, color=GOLD, size=16)
            v.fill = _fill(TEAL)
            v.alignment = _align("center")
            s = ws.cell(row=9, column=col, value=sub)
            s.font = _font(color="AAAAAA", size=8)
            s.fill = _fill("1A5050")
            s.alignment = _align("center")
            ws.row_dimensions[7].height = 16
            ws.row_dimensions[8].height = 26
            ws.row_dimensions[9].height = 14

        kpis2 = [
            ("OCCUPANCY",     f"{d.occupancy_pct:.0f}%",
                              f"{d.nights_occupied} из {d.days_in_month} ночей"),
            ("ADR",           f"{d.adr:,.0f} ฿",       "Средняя стоимость/ночь"),
            ("EXPENSE RATIO", f"{d.expense_ratio:.1f}%", "OPEX / Revenue"),
        ]
        for i, (title, val, sub) in enumerate(kpis2):
            col = kpi_cols[i]
            for rn in [11, 12, 13]:
                ws.merge_cells(start_row=rn, start_column=col, end_row=rn, end_column=col+1)
            h = ws.cell(row=11, column=col, value=title)
            h.font = _font(bold=True, color=WHITE, size=9)
            h.fill = _fill(NAVY)
            h.alignment = _align("center")
            v = ws.cell(row=12, column=col, value=val)
            v.font = _font(bold=True, color=GOLD, size=14)
            v.fill = _fill(TEAL)
            v.alignment = _align("center")
            s = ws.cell(row=13, column=col, value=sub)
            s.font = _font(color="AAAAAA", size=8)
            s.fill = _fill("1A5050")
            s.alignment = _align("center")
            ws.row_dimensions[11].height = 16
            ws.row_dimensions[12].height = 22
            ws.row_dimensions[13].height = 14

        # Bookings table (row 16+)
        ws.row_dimensions[15].height = 8
        ws.merge_cells("B16:H16")
        hdr = ws["B16"]
        hdr.value = f"BOOKINGS — {d.period_label.upper()}"
        hdr.font = _font(bold=True, color=WHITE, size=10)
        hdr.fill = _fill(NAVY)
        hdr.alignment = _align("left")
        ws.row_dimensions[16].height = 18

        bk_headers = [("Бронирование",16),("Канал",14),("Гость/Описание",22),
                      ("Даты",16),("Ночей",8),("Gross (฿)",14),("Net Owner (฿)",14)]
        _apply_header_row(ws, 17, bk_headers, bg=TEAL)

        if not d.reservations:
            ws.merge_cells("B18:H18")
            c = ws.cell(row=18, column=2, value="Бронирований нет")
            c.font = _font(color="888888", italic=True)
            bk_end_row = 18
        else:
            for ri, res in enumerate(d.reservations, 18):
                ws.row_dimensions[ri].height = 15
                bg = LGRAY if ri % 2 == 0 else WHITE
                ci_d = res["checkin"]
                co_d = res["checkout"]
                dates_str = ""
                if ci_d and co_d:
                    dates_str = f"{ci_d.strftime('%d.%m')}—{co_d.strftime('%d.%m')}"
                vals = [
                    res["booking_id"], res["channel"],
                    "", dates_str,
                    res["nights"],
                    res["gross"] if res["gross"] else "",
                    res["net_to_owner"] if res["net_to_owner"] else "",
                ]
                for ci, val in enumerate(vals, 2):
                    c = ws.cell(row=ri, column=ci, value=val)
                    c.font = _font(size=9)
                    c.fill = _fill(bg)
                    c.border = _bdr()
                    if ci in [7, 8]:
                        c.number_format = "#,##0"
                        c.alignment = _align("right")
            bk_end_row = 17 + len(d.reservations)

        # MoM comparison (2 rows below bookings)
        mom_start = bk_end_row + 2
        ws.row_dimensions[mom_start].height = 8
        ws.merge_cells(start_row=mom_start+1, start_column=2,
                       end_row=mom_start+1, end_column=8)
        hdr = ws.cell(row=mom_start+1, column=2,
                      value="СРАВНЕНИЕ — ТЕКУЩИЙ МЕСЯЦ vs ПРЕДЫДУЩИЙ")
        hdr.font = _font(bold=True, color=WHITE, size=10)
        hdr.fill = _fill(NAVY)
        hdr.alignment = _align("left")
        ws.row_dimensions[mom_start+1].height = 18

        mom_hdr = [("Показатель",22),("Текущий месяц",18),("Предыдущий месяц",18),
                   ("Δ THB / %",14)]
        _apply_header_row(ws, mom_start+2, mom_hdr, bg=TEAL)

        mom_data = [
            ("Gross Revenue (฿)",    d.total_gross,     d.prior["gross"]),
            ("TL Management Fee (฿)",d.total_tl_comm,   d.prior["mgmt_fee"]),
            ("Total OPEX (฿)",       d.total_opex,      d.prior["opex"]),
            ("Net Income (฿)",       d.net_income,      d.prior["net_income"]),
            ("Occupancy %",          d.occupancy_pct,   d.prior["occupancy"]),
            ("ADR (฿)",              d.adr,             d.prior["adr"]),
        ]
        for mi, (label, curr, prev) in enumerate(mom_data, mom_start+3):
            ws.row_dimensions[mi].height = 15
            bg = LGRAY if mi % 2 == 0 else WHITE
            delta = curr - prev if prev else 0
            delta_pct = (delta / prev * 100) if prev else 0
            delta_str = f"{delta:+,.0f} / {delta_pct:+.1f}%"
            delta_color = GRN_FG if delta >= 0 else RED_FG
            for ci, (val, fmt, al) in enumerate([
                (label, None, "left"),
                (curr, "#,##0.0", "right"),
                (prev, "#,##0.0", "right"),
                (delta_str, None, "center"),
            ], 2):
                c = ws.cell(row=mi, column=ci, value=val)
                c.font = _font(size=9, color=(delta_color if ci == 5 else "000000"),
                               bold=(ci == 5))
                c.fill = _fill(bg)
                c.border = _bdr()
                if fmt:
                    c.number_format = fmt
                c.alignment = _align(al)

        # Balance movement row
        bal_row = mom_start + 3 + len(mom_data) + 1
        _section_hdr(ws, bal_row, "ДВИЖЕНИЕ БАЛАНСА", 7, bg=NAVY)
        ws.row_dimensions[bal_row+1].height = 18
        bal_labels = [
            f"Начальный баланс\n{d.opening_balance:,.0f} ฿",
            f"+ Доходы\n{d.total_gross:,.0f} ฿",
            f"− Комиссия TL\n{d.total_tl_comm:,.0f} ฿",
            f"− OPEX\n{d.total_opex:,.0f} ฿",
            f"− Выплаты\n{d.total_payouts:,.0f} ฿",
            "=",
            f"Конечный баланс\n{d.closing_balance:,.0f} ฿",
        ]
        bal_colors = [LGRAY, GRN_BG, RED_BG, RED_BG, RED_BG, WHITE, TEAL]
        bal_fcolors = ["333333", GRN_FG, RED_FG, RED_FG, RED_FG, "000000", WHITE]
        for ci, (lbl, bg, fg) in enumerate(zip(bal_labels, bal_colors, bal_fcolors), 2):
            c = ws.cell(row=bal_row+1, column=ci, value=lbl)
            c.font = _font(bold=(ci in [2, 8]), color=fg, size=8)
            c.fill = _fill(bg)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = _bdr()
        ws.row_dimensions[bal_row+1].height = 30

        # Column widths
        for col, w in [("A",3),("B",24),("C",16),("D",24),("E",16),("F",10),("G",16),("H",16)]:
            ws.column_dimensions[col].width = w

        return ws

    # ── TAB 2: P&L MONTHLY ────────────────────────────────────────────────────
    def build_pl_monthly(self):
        ws = self.wb.create_sheet("P&L Monthly")
        ws.sheet_properties.tabColor = TEAL
        ws.sheet_view.showGridLines = False
        d = self.d

        # Title
        ws.merge_cells("A1:D1")
        t = ws["A1"]
        t.value = f"P&L — {d.property_name} — {d.period_label}"
        t.font = _font(bold=True, color=WHITE, size=12)
        t.fill = _fill(NAVY)
        t.alignment = _align("center")
        ws.row_dimensions[1].height = 24

        ws.merge_cells("A2:D2")
        t2 = ws["A2"]
        t2.value = f"Начальный баланс: {d.opening_balance:,.2f} ฿"
        t2.font = _font(color="CCCCCC", size=9)
        t2.fill = _fill(TEAL)
        t2.alignment = _align("left")

        # Column headers
        ws.cell(row=4, column=2, value=d.period_label).font = _font(bold=True, size=9)
        ws.cell(row=4, column=3, value=d.prior["period"]).font = _font(bold=True, size=9)
        ws.cell(row=4, column=4, value="Примечание").font = _font(bold=True, size=9, italic=True)
        ws.row_dimensions[4].height = 16
        for col in [2,3,4]:
            ws.cell(row=4, column=col).border = _bdr()
            ws.cell(row=4, column=col).fill = _fill(LGRAY)

        row = 6

        def pl_section(label, bg=NAVY):
            nonlocal row
            _section_hdr(ws, row, label, 4)
            row += 1

        def pl_line(label, curr, prior=None, note="", bold=False,
                    color="000000", indent=3):
            nonlocal row
            ws.row_dimensions[row].height = 15
            bg = LGRAY if bold else WHITE
            fg = GOLD if bold else color
            c1 = ws.cell(row=row, column=1, value=(" " * indent) + label)
            c1.font = _font(bold=bold, color=fg, size=9)
            c1.fill = _fill(bg if not bold else NAVY)
            c1.alignment = _align("left")
            c1.border = _bdr()
            c2 = ws.cell(row=row, column=2, value=curr if curr != 0 else "")
            c2.font = _font(bold=bold, color=(WHITE if bold else color), size=9)
            c2.fill = _fill(bg if not bold else NAVY)
            c2.number_format = "#,##0"
            c2.alignment = _align("right")
            c2.border = _bdr()
            if prior is not None:
                c3 = ws.cell(row=row, column=3, value=prior if prior != 0 else "")
                c3.font = _font(size=9, color="666666")
                c3.fill = _fill(bg if not bold else NAVY)
                c3.number_format = "#,##0"
                c3.alignment = _align("right")
                c3.border = _bdr()
            c4 = ws.cell(row=row, column=4, value=note)
            c4.font = _font(size=8, italic=True, color="888888")
            c4.fill = _fill(WHITE)
            c4.alignment = _align("left", wrap=True)
            c4.border = _bdr()
            row += 1

        def pl_separator():
            nonlocal row
            ws.row_dimensions[row].height = 5
            row += 1

        # ─ REVENUE ─
        pl_section("REVENUE / ДОХОДЫ")
        for res in d.reservations:
            co = res["checkout"]
            note_str = (f"{res['channel']}, {res['nights']}н"
                        + (f", выезд {co.strftime('%d.%m')}" if co else ""))
            pl_line(res["booking_id"], res["gross"], note=note_str)

        # Utility recharge (if any)
        utl_rchg = sum(e["amount"] for e in d.expenses
                       if e["category_code"] in ("UTL-RCHG","UTL_RCHG"))
        if utl_rchg > 0:
            pl_line("Возмещение электричества", utl_rchg, note="Utility recharge")

        total_income = d.total_gross + utl_rchg
        pl_line("ИТОГО ДОХОД (A)", total_income, d.prior["gross"] + 0,
                bold=True, indent=0)
        pl_separator()

        # ─ TL COMMISSION ─
        pl_section("КОМИССИЯ TROPICLOOK")
        for res in d.reservations:
            note_c = f"{d.commission_rate}% от {res['gross']:,.0f}"
            pl_line(res["booking_id"], -res["tl_comm"], note=note_c)
        if d.total_ota_comm > 0:
            pl_line("Справочно: комиссия OTA (не влияет на баланс)",
                    -d.total_ota_comm, note="Удержано OTA", color="888888")
        pl_line("ИТОГО КОМИССИЯ TL (B)", -d.total_tl_comm, -d.prior["mgmt_fee"],
                bold=True, indent=0)
        pl_separator()

        # ─ OPEX ─
        pl_section("ОПЕРАЦИОННЫЕ РАСХОДЫ / OPEX")
        opex_cats = d.opex_by_category()
        for code, cat in sorted(opex_cats.items(), key=lambda x: -x[1]["total"]):
            bud = d.opex_budget.get(code, {}).get("budget", 0)
            note_str = f"бюджет:{bud:,.0f}" if bud else ""
            pl_line(cat["name"] or code, -cat["total"], note=note_str)
        pl_line(f"ИТОГО OPEX (C) — {len(d.expenses)} операций",
                -d.total_opex, -d.prior["opex"], bold=True, indent=0)
        pl_separator()

        # ─ NET INCOME ─
        pl_section("ЧИСТЫЙ ДОХОД", bg=TEAL)
        pl_line("ЧИСТЫЙ ДОХОД (A+B+C)", d.net_income, d.prior["net_income"],
                bold=True, color=GOLD if d.net_income >= 0 else RED_FG, indent=0)
        pl_separator()

        # ─ PAYOUTS ─
        pl_section("ВЫПЛАТЫ СОБСТВЕННИКУ")
        for pay in d.payouts:
            note_p = pay["date"].strftime("%d.%m.%Y") if pay["date"] else ""
            pl_line(pay["description"] or pay["type"], -pay["amount"], note=note_p)
        pl_line(f"ИТОГО ВЫПЛАТЫ (D) — {len(d.payouts)} выплат",
                -d.total_payouts, -d.prior["payouts"], bold=True, indent=0)
        pl_separator()

        # ─ BALANCE ─
        pl_section("ДВИЖЕНИЕ БАЛАНСА", bg=NAVY)
        pl_line("Баланс на начало",     d.opening_balance,   note="")
        pl_line("+ Доходы (A)",          total_income,        note="")
        pl_line("+ Utility Recharge",    utl_rchg if utl_rchg else "",
                note="уже включено в A") if utl_rchg > 0 else None
        pl_line("− Комиссия TL (B)",     -d.total_tl_comm,
                note="OTA не влияет на баланс")
        pl_line("− OPEX (C)",            -d.total_opex,       note="")
        pl_line("− Выплаты (D)",         -d.total_payouts,    note="")
        pl_line("БАЛАНС НА КОНЕЦ",       d.closing_balance,
                bold=True, color=(GRN_FG if d.closing_balance >= 0 else RED_FG), indent=0)

        for col, w in [("A",28),("B",18),("C",18),("D",30)]:
            ws.column_dimensions[col].width = w

        return ws

    # ── TAB 3: OPEX PASSPORT ──────────────────────────────────────────────────
    def build_opex_passport(self):
        ws = self.wb.create_sheet("OPEX Passport")
        ws.sheet_properties.tabColor = GOLD
        ws.sheet_view.showGridLines = False
        d = self.d

        ws.merge_cells("A1:G1")
        t = ws["A1"]
        t.value = f"OPEX PASSPORT — {d.property_name} — {d.period_label}"
        t.font = _font(bold=True, color=WHITE, size=11)
        t.fill = _fill(NAVY)
        t.alignment = _align("center")
        ws.row_dimensions[1].height = 22

        _apply_header_row(ws, 4,
            [("Код",10),("Категория",24),("Бюджет ฿",14),("Факт ฿",14),
             ("Δ ฿",14),("Δ %",10),("Статус",12)], bg=TEAL)

        opex_cats = d.opex_by_category()
        # Merge budget and actual categories
        all_codes = sorted(
            set(list(d.opex_budget.keys()) + list(opex_cats.keys()))
        )

        total_bud = 0.0
        total_act = 0.0

        def status_flag(bud, act):
            if bud == 0 and act == 0:
                return ("➖ Zero",    LGRAY,  "555555")
            if bud == 0 and act > 0:
                return ("📌 Unbudgeted", "E3F2FD", "0D47A1")
            if act == 0 and bud > 0:
                return ("✅ In Budget", GRN_BG, GRN_FG)
            diff = (act - bud) / bud * 100
            if diff <= 0:
                return ("✅ In Budget", GRN_BG, GRN_FG)
            elif diff <= 10:
                return ("🟡 Minor",    AMB_BG, AMB_FG)
            elif diff <= 30:
                return ("⚠️ Significant", "FFF3E0", "E65100")
            else:
                return ("🚨 Critical",  RED_BG, RED_FG)

        for ri, code in enumerate(all_codes, 5):
            ws.row_dimensions[ri].height = 15
            bg_row = LGRAY if ri % 2 == 0 else WHITE
            bud_info = d.opex_budget.get(code, {"name": code, "budget": 0})
            act_info = opex_cats.get(code, {"name": bud_info["name"], "total": 0})
            bud = bud_info["budget"]
            act = act_info["total"]
            total_bud += bud
            total_act += act
            delta = act - bud
            delta_pct = (delta / bud * 100) if bud > 0 else 0
            flag, flag_bg, flag_fg = status_flag(bud, act)

            vals = [code, act_info.get("name") or bud_info.get("name") or code,
                    bud, act, delta, f"{delta_pct:.0f}%", flag]
            colors_  = ["555555","000000","000000","000000",
                        (RED_FG if delta > 0 else GRN_FG), (RED_FG if delta > 0 else GRN_FG),
                        flag_fg]
            bgs_     = [bg_row]*6 + [flag_bg]

            for ci, (val, color, bg_c) in enumerate(zip(vals, colors_, bgs_), 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = _font(size=9, color=color,
                               bold=(ci in [1, 7]))
                c.fill = _fill(bg_c)
                c.border = _bdr()
                if ci in [3,4,5]:
                    c.number_format = "#,##0"
                    c.alignment = _align("right")

        # Totals row
        tot_row = 5 + len(all_codes)
        ws.row_dimensions[tot_row].height = 18
        total_delta = total_act - total_bud
        tot_vals = ["", "ИТОГО", total_bud, total_act, total_delta,
                    f"{(total_delta/total_bud*100):.0f}%" if total_bud else "",
                    ("✅" if total_delta <= 0 else ("⚠️" if total_delta/total_bud < 0.3 else "🚨")) if total_bud else ""]
        for ci, val in enumerate(tot_vals, 1):
            c = ws.cell(row=tot_row, column=ci, value=val)
            c.font = _font(bold=True, size=9, color=WHITE)
            c.fill = _fill(NAVY)
            c.border = _bdr()
            if ci in [3,4,5]:
                c.number_format = "#,##0"
                c.alignment = _align("right")

        # Anomaly block
        anomalies = []
        for code in all_codes:
            bud_info = d.opex_budget.get(code, {"budget": 0})
            act_info = opex_cats.get(code, {"total": 0})
            bud = bud_info["budget"]
            act = act_info["total"]
            flag, _, _ = status_flag(bud, act)
            if "✅" not in flag and "➖" not in flag:
                name = act_info.get("name") or bud_info.get("name") or code
                anomalies.append(f"{flag} {code} ({name}): бюджет {bud:,.0f} / факт {act:,.0f} ฿")

        if anomalies:
            anom_start = tot_row + 2
            _section_hdr(ws, anom_start, "⚠ ANOMALY FLAGS — Требуют внимания", 7, bg="BF360C")
            for ai, txt in enumerate(anomalies, anom_start+1):
                ws.row_dimensions[ai].height = 14
                ws.merge_cells(start_row=ai, start_column=1, end_row=ai, end_column=7)
                c = ws.cell(row=ai, column=1, value=txt)
                c.font = _font(size=9, color=RED_FG)
                c.fill = _fill(RED_BG)
                c.border = _bdr()

        for col, w in [("A",10),("B",26),("C",14),("D",14),("E",14),("F",10),("G",16)]:
            ws.column_dimensions[col].width = w

        ws.freeze_panes = "A5"
        ws.auto_filter.ref = f"A4:G{4+len(all_codes)}"
        return ws

    # ── TAB 4: 12-MONTH SUMMARY ───────────────────────────────────────────────
    def build_12month(self):
        ws = self.wb.create_sheet("12-Month Summary")
        ws.sheet_properties.tabColor = "444444"
        ws.sheet_view.showGridLines = False
        d = self.d

        ws.merge_cells("A1:N1")
        t = ws["A1"]
        t.value = f"PERFORMANCE OVERVIEW — {d.property_name}"
        t.font = _font(bold=True, color=WHITE, size=11)
        t.fill = _fill(NAVY)
        t.alignment = _align("center")
        ws.row_dimensions[1].height = 22

        ws.merge_cells("A2:N2")
        t2 = ws["A2"]
        t2.value = f"Период управления | YTD: {d.cumulative['period']}"
        t2.font = _font(color="CCCCCC", size=9)
        t2.fill = _fill(TEAL)
        t2.alignment = _align("center")

        # Derive month columns based on report_period
        try:
            yr, mo = d.report_period.split("-")
            yr, mo = int(yr), int(mo)
        except Exception:
            yr, mo = 2025, 1

        months = []
        for i in range(11, -1, -1):
            m = mo - i
            y = yr
            while m <= 0:
                m += 12
                y -= 1
            months.append((y, m))

        month_labels = []
        for y, m in months:
            mon_abbr = ["Jan","Feb","Mar","Apr","May","Jun",
                        "Jul","Aug","Sep","Oct","Nov","Dec"][m-1]
            month_labels.append(f"{mon_abbr}'{str(y)[2:]}")

        # Headers: row 4
        ws.cell(row=4, column=1, value="Показатель")
        ws.cell(row=4, column=1).font = _font(bold=True, size=9, color=WHITE)
        ws.cell(row=4, column=1).fill = _fill(NAVY)
        ws.column_dimensions["A"].width = 26

        for ci, lbl in enumerate(month_labels, 2):
            is_current = (months[ci-2] == (yr, mo))
            c = ws.cell(row=4, column=ci, value=lbl)
            c.font = _font(bold=True, size=9,
                           color=(WHITE if not is_current else GOLD))
            c.fill = _fill(TEAL if not is_current else NAVY)
            c.alignment = _align("center")
            c.border = _bdr()
            ws.column_dimensions[get_column_letter(ci)].width = 12

        ws.cell(row=4, column=14, value="TOTAL/YTD")
        ws.cell(row=4, column=14).font = _font(bold=True, size=9, color=WHITE)
        ws.cell(row=4, column=14).fill = _fill(NAVY)
        ws.cell(row=4, column=14).alignment = _align("center")
        ws.column_dimensions["N"].width = 14
        ws.row_dimensions[4].height = 18

        # Current month index
        curr_idx = 11  # last column (index from 0)

        # Data rows
        sections = [
            ("REVENUE", None),
            ("Gross Revenue (฿)",    d.total_gross,    d.prior["gross"],   d.cumulative["gross"]),
            ("TL Commission (฿)",    d.total_tl_comm,  d.prior["mgmt_fee"],d.cumulative["mgmt_fee"]),
            ("OPEX", None),
            ("Total OPEX (฿)",       d.total_opex,     d.prior["opex"],    d.cumulative["opex"]),
            ("RESULT", None),
            ("Net Income (฿)",       d.net_income,     d.prior["net_income"],d.cumulative["net_income"]),
            ("Payouts (฿)",          d.total_payouts,  d.prior["payouts"], d.cumulative["payouts"]),
            ("Closing Balance (฿)",  d.closing_balance,d.prior["balance"], d.closing_balance),
            ("KPIs", None),
            ("Bookings (#)",         len(d.reservations), d.prior["bookings"], d.cumulative["bookings"]),
            ("Nights occupied",      d.nights_occupied,   d.prior["nights"],   d.cumulative["nights"]),
            ("Occupancy %",          d.occupancy_pct,     d.prior["occupancy"],d.cumulative["occupancy"]),
            ("ADR (฿)",              d.adr,               d.prior["adr"],      d.cumulative["adr"]),
        ]

        for ri, row_data in enumerate(sections, 5):
            ws.row_dimensions[ri].height = 15
            label = row_data[0]
            # Section header
            if row_data[1] is None:
                _section_hdr(ws, ri, label, 14)
                continue
            curr_val, prior_val, ytd_val = row_data[1], row_data[2], row_data[3]

            c = ws.cell(row=ri, column=1, value=label)
            c.font = _font(size=9)
            c.border = _bdr()

            for ci in range(2, 14):
                col_idx = ci - 2  # 0-based
                is_current = (col_idx == curr_idx)
                is_prior = (col_idx == curr_idx - 1)
                if is_current:
                    val = curr_val
                elif is_prior:
                    val = prior_val
                else:
                    val = "—"
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.font = _font(size=9,
                                  bold=is_current,
                                  color=(GOLD if is_current else "444444"))
                cell.fill = _fill(TEAL if is_current else (LGRAY if ri % 2 == 0 else WHITE))
                cell.border = _bdr()
                cell.alignment = _align("right")
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0.0"

            ytd_cell = ws.cell(row=ri, column=14, value=ytd_val)
            ytd_cell.font = _font(bold=True, size=9, color=WHITE)
            ytd_cell.fill = _fill(NAVY)
            ytd_cell.border = _bdr()
            ytd_cell.alignment = _align("right")
            ytd_cell.number_format = "#,##0.0"

        ws.freeze_panes = "B5"
        return ws

    # ── TAB 5: TRANSACTION LEDGER ─────────────────────────────────────────────
    def build_ledger(self):
        ws = self.wb.create_sheet("Transaction Ledger")
        ws.sheet_properties.tabColor = NAVY
        ws.sheet_view.showGridLines = False
        d = self.d

        ws.merge_cells("B2:G2")
        t = ws["B2"]
        t.value = f"LEDGER — {d.property_name} — {d.period_label}"
        t.font = _font(bold=True, color=WHITE, size=11)
        t.fill = _fill(NAVY)
        t.alignment = _align("center")
        ws.row_dimensions[2].height = 22

        _apply_header_row(ws, 5,
            [("",1),("Дата",10),("Описание",34),("Кат.",14),
             ("Приход ฿",14),("Расход ฿",14),("Баланс ฿",16)],
            bg=TEAL, height=18)

        # Build transaction list
        # Key rule: revenue recognized on CHECKOUT DATE
        transactions = []

        # Add expense transactions (date of payment)
        for e in d.expenses:
            if e["category_code"] in ("UTL-RCHG", "UTL_RCHG"):
                continue  # utility recharge goes with its booking
            transactions.append({
                "date": e["date"],
                "desc": e["description"],
                "cat":  e["category_code"],
                "inc":  0,
                "exp":  e["amount"],
            })

        # Add revenue on CHECKOUT date (principle from FIN-REG-OWN-RPT-001)
        for res in d.reservations:
            checkout = res["checkout"]
            # Revenue
            transactions.append({
                "date": checkout,
                "desc": f"Приход: {res['booking_id']} (выезд — checkout)",
                "cat":  "REVENUE",
                "inc":  res["gross"],
                "exp":  0,
            })
            # TL Commission (same date as revenue)
            if res["tl_comm"] > 0:
                transactions.append({
                    "date": checkout,
                    "desc": f"Комиссия TL: {res['booking_id']}",
                    "cat":  "MGMT-FEE",
                    "inc":  0,
                    "exp":  res["tl_comm"],
                })
            # Utility recharge (same date as checkout)
            utl = sum(e["amount"] for e in d.expenses
                      if e["category_code"] in ("UTL-RCHG", "UTL_RCHG")
                      and (e["booking_ref"] == res["booking_id"] or
                           e["booking_ref"] == ""))
            if utl > 0:
                transactions.append({
                    "date": checkout,
                    "desc": f"Электричество: {res['booking_id']} (возмещение)",
                    "cat":  "UTL-RCHG",
                    "inc":  utl,
                    "exp":  0,
                })

        # Add payouts
        for pay in d.payouts:
            transactions.append({
                "date": pay["date"],
                "desc": pay["description"] or pay["type"],
                "cat":  "PAYOUT",
                "inc":  0,
                "exp":  pay["amount"],
            })

        # Sort chronologically (None dates go last)
        def sort_key(t):
            if t["date"] is None:
                return date(9999, 12, 31)
            return t["date"]

        transactions.sort(key=sort_key)

        # Write opening balance
        ws.row_dimensions[6].height = 15
        ws.cell(row=6, column=3, value="НАЧАЛЬНЫЙ БАЛАНС").font = _font(bold=True, color=NAVY, size=9)
        ws.cell(row=6, column=7, value=d.opening_balance).font = _font(bold=True, color=NAVY, size=9)
        ws.cell(row=6, column=7).number_format = "#,##0.00"
        ws.cell(row=6, column=7).border = _bdr()
        ws.cell(row=6, column=3).border = _bdr()

        # Write transactions
        running = d.opening_balance
        for ti, txn in enumerate(transactions, 7):
            ws.row_dimensions[ti].height = 15
            cat = txn["cat"]
            inc = txn["inc"]
            exp = txn["exp"]
            if inc > 0:
                running += inc
            if exp > 0:
                running -= exp
            running = round(running, 2)

            bg, fg = CAT_COLOR.get(cat, DEFAULT_CAT_COLOR)
            bold_row = cat in ("REVENUE", "MGMT-FEE", "PAYOUT")

            date_str = _date_display(txn["date"])
            vals = [None, date_str, txn["desc"], cat,
                    inc if inc > 0 else None,
                    exp if exp > 0 else None,
                    running]
            for ci, val in enumerate(vals, 1):
                c = ws.cell(row=ti, column=ci, value=val)
                c.fill = _fill(bg)
                c.font = _font(size=9, color=fg, bold=bold_row)
                c.border = _bdr()
                if ci in [5, 6, 7] and val is not None:
                    c.number_format = "#,##0.00"
                c.alignment = _align("center" if ci in [2, 4] else "left")

        # Closing balance
        closing_row = 7 + len(transactions)
        ws.row_dimensions[closing_row].height = 15
        ws.cell(row=closing_row, column=3, value="КОНЕЧНЫЙ БАЛАНС")
        ws.cell(row=closing_row, column=3).font = _font(bold=True, color=NAVY, size=9)
        ws.cell(row=closing_row, column=3).border = _bdr()
        ws.cell(row=closing_row, column=7, value=round(running, 2))
        ws.cell(row=closing_row, column=7).font = _font(bold=True, color=NAVY, size=9)
        ws.cell(row=closing_row, column=7).number_format = "#,##0.00"
        ws.cell(row=closing_row, column=7).border = _bdr()

        # Note
        note_row = closing_row + 2
        try:
            ws.merge_cells(start_row=note_row, start_column=2,
                           end_row=note_row, end_column=7)
        except Exception:
            pass
        note_txt = (f"⚠ Принцип признания дохода: дата ВЫЕЗДА гостя (checkout date). "
                    f"Регламент FIN-REG-OWN-RPT-001, Раздел 2.")
        nc = ws.cell(row=note_row, column=2, value=note_txt)
        nc.font = _font(size=8, italic=True, color="666666")
        nc.alignment = _align("left")

        # Totals
        total_in = sum(t["inc"] for t in transactions if t["inc"] > 0)
        total_out = sum(t["exp"] for t in transactions if t["exp"] > 0)
        tot_row = note_row + 1
        ws.row_dimensions[tot_row].height = 16
        for ci, (lbl, val) in enumerate([
            ("", None), ("", None),
            ("ИТОГО", None),
            ("", None),
            (total_in, total_in),
            (total_out, total_out),
            ("", None),
        ], 1):
            c = ws.cell(row=tot_row, column=ci,
                        value=val if isinstance(val, float) else lbl)
            c.font = _font(bold=True, size=9, color=WHITE)
            c.fill = _fill(TEAL)
            c.border = _bdr()
            if ci == 3:
                c.value = "ИТОГО ПРИХОД / РАСХОД"
            if isinstance(val, float):
                c.number_format = "#,##0.00"
                c.alignment = _align("right")

        for col, w in [("A",3),("B",10),("C",36),("D",14),
                       ("E",14),("F",14),("G",16)]:
            ws.column_dimensions[col].width = w

        ws.freeze_panes = "B6"
        ws.auto_filter.ref = f"B5:G{closing_row}"
        return ws

    # ── TAB 6: DAP SNAPSHOT ───────────────────────────────────────────────────
    def build_dap_snapshot(self):
        ws = self.wb.create_sheet("DAP Snapshot")
        ws.sheet_properties.tabColor = TEAL
        ws.sheet_view.showGridLines = False
        d = self.d

        ws.merge_cells("A1:F1")
        t = ws["A1"]
        t.value = f"DIGITAL ASSET PASSPORT — {d.property_name}"
        t.font = _font(bold=True, color=WHITE, size=12)
        t.fill = _fill(NAVY)
        t.alignment = _align("center")
        ws.row_dimensions[1].height = 24

        ws.merge_cells("A2:F2")
        t2 = ws["A2"]
        t2.value = (f"TL-{d.property_code}-001 | ACTIVE | {d.period_label}")
        t2.font = _font(color="CCCCCC", size=9)
        t2.fill = _fill(TEAL)
        t2.alignment = _align("center")

        row = 4

        def dap_section(label):
            nonlocal row
            _section_hdr(ws, row, label, 6, bg=NAVY)
            row += 1

        def dap_row(label, value, label2=None, value2=None):
            nonlocal row
            ws.row_dimensions[row].height = 15
            ws.cell(row=row, column=1, value=label).font = _font(bold=True, size=9, color=NAVY)
            c = ws.cell(row=row, column=2, value=value)
            c.font = _font(size=9)
            c.border = _bdr()
            c.fill = _fill(LGRAY)
            if label2:
                ws.cell(row=row, column=4, value=label2).font = _font(bold=True, size=9, color=NAVY)
                c2 = ws.cell(row=row, column=5, value=value2)
                c2.font = _font(size=9)
                c2.border = _bdr()
                c2.fill = _fill(LGRAY)
            row += 1

        # L1
        dap_section("L1 — IDENTITY")
        dap_row("Property", d.property_name, "Owner", d.owner_name)
        dap_row("Code", f"TL-{d.property_code}-001", "Email", d.owner_email)
        dap_row("Type", d.property_type, "Bedrooms", d.bedrooms)
        dap_row("Location", d.location, "Legal Entity", d.legal_entity)
        dap_row("Commission", f"{d.commission_rate}%", "Contract Start", d.contract_start)
        dap_row("Status", "ACTIVE", "Period", d.period_label)
        row += 1

        # L2
        dap_section("L2 — TECHNICAL")
        dap_row("Property Type", d.property_type, "Bedrooms", d.bedrooms)
        dap_row("Available Nights/mo", d.available_nights, "Days in Month", d.days_in_month)
        dap_row("Last Inspection", "— (pending)", "Condition Score", "— (pending)")
        row += 1

        # L3
        channels = list({r["channel"] for r in d.reservations if r["channel"]})
        dap_section("L3 — OPERATIONAL (current month)")
        dap_row("Bookings", len(d.reservations), "Nights Occupied", d.nights_occupied)
        dap_row("Occupancy", f"{d.occupancy_pct:.0f}%", "ADR", f"{d.adr:,.0f} ฿")
        dap_row("Channels", ", ".join(channels) if channels else "—", "Top Channel",
                max(channels, key=lambda ch: sum(r["gross"] for r in d.reservations
                                                  if r["channel"]==ch)) if channels else "—")
        row += 1

        # L4
        opex_by_cat = d.opex_by_category()
        top3 = sorted(opex_by_cat.items(), key=lambda x: -x[1]["total"])[:3]
        dap_section("L4 — OPEX SUMMARY (current month)")
        dap_row("Total OPEX", f"{d.total_opex:,.0f} ฿",
                "Expense Ratio", f"{d.expense_ratio:.1f}%")
        for i, (code, cat) in enumerate(top3):
            bud = d.opex_budget.get(code, {}).get("budget", 0)
            dap_row(f"Top OPEX #{i+1}", f"{code}: {cat['total']:,.0f} ฿",
                    "vs Budget", f"{bud:,.0f} ฿")
        row += 1

        # L5
        dap_section("L5 — FINANCIAL (current month)")
        dap_row("Gross Revenue", f"{d.total_gross:,.0f} ฿",
                "TL Commission", f"{d.total_tl_comm:,.0f} ฿")
        dap_row("Total OPEX", f"{d.total_opex:,.0f} ฿",
                "Net Income", f"{d.net_income:,.0f} ฿")
        dap_row("Payouts", f"{d.total_payouts:,.0f} ฿",
                "Opening Balance", f"{d.opening_balance:,.0f} ฿")
        dap_row("Closing Balance", f"{d.closing_balance:,.0f} ฿", "", "")
        row += 1

        # L5 Cumulative
        dap_section("L5 — FINANCIAL (YTD cumulative)")
        dap_row("YTD Gross Revenue", f"{d.cumulative['gross']:,.0f} ฿",
                "YTD TL Commission", f"{d.cumulative['mgmt_fee']:,.0f} ฿")
        dap_row("YTD OPEX", f"{d.cumulative['opex']:,.0f} ฿",
                "YTD Net Income", f"{d.cumulative['net_income']:,.0f} ฿")
        dap_row("YTD Payouts", f"{d.cumulative['payouts']:,.0f} ฿",
                "Purchase Price", f"{d.cumulative['purchase_price']:,.0f} ฿"
                if d.cumulative['purchase_price'] > 0 else "Not disclosed")
        if d.cumulative['purchase_price'] > 0:
            dap_row("Gross Yield (annualized)", f"{d.cumulative['yield_pct']:.1f}%", "", "")
        row += 1

        # L6 — Owner narrative
        dap_section("L6 — OWNER COMMUNICATION NOTE")
        inc_flag = "▲" if d.net_income > d.prior["net_income"] else "▼"
        note = (
            f"За {d.period_label} объект {d.property_name} сгенерировал "
            f"{d.total_gross:,.0f} ฿ валовой выручки ({len(d.reservations)} бронирований, "
            f"{d.nights_occupied} ночей). Чистый доход собственника составил "
            f"{d.net_income:,.0f} ฿ {inc_flag} "
            f"{'(выше' if d.net_income >= d.prior['net_income'] else '(ниже'} прошлого месяца "
            f"на {abs(d.net_income - d.prior['net_income']):,.0f} ฿). "
            f"Занятость {d.occupancy_pct:.0f}%. "
            f"Выплачено собственнику: {d.total_payouts:,.0f} ฿. "
            f"Остаток на счёте: {d.closing_balance:,.0f} ฿."
        )
        ws.merge_cells(start_row=row, start_column=1, end_row=row+2, end_column=6)
        nc = ws.cell(row=row, column=1, value=note)
        nc.font = _font(size=9, color="222222", italic=True)
        nc.fill = _fill("EFF8F8")
        nc.border = _bdr()
        nc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[row].height = 60

        for col, w in [("A",22),("B",24),("C",4),("D",22),("E",24),("F",4)]:
            ws.column_dimensions[col].width = w

        return ws

    # ── MAIN BUILD ────────────────────────────────────────────────────────────
    def build(self, output_dir="."):
        print(f"  Building Dashboard...")
        self.build_dashboard()
        print(f"  Building P&L Monthly...")
        self.build_pl_monthly()
        print(f"  Building OPEX Passport...")
        self.build_opex_passport()
        print(f"  Building 12-Month Summary...")
        self.build_12month()
        print(f"  Building Transaction Ledger...")
        self.build_ledger()
        print(f"  Building DAP Snapshot...")
        self.build_dap_snapshot()

        filename = (f"TL_{self.d.property_code}_OwnerReport_"
                    f"{self.d.report_period}_v1.xlsx")
        out_path = os.path.join(output_dir, filename)
        self.wb.save(out_path)
        return out_path


# ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        print("USAGE: python3 tl_report_engine.py <input_template.xlsx> [output_dir]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(input_path) or "."

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"  TropicLook Owner Report Engine")
    print(f"  FIN-REG-OWN-RPT-001 v1.0")
    print(f"{'='*60}")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_dir}")
    print()

    # Step 1: Parse input
    print("[ 1/3 ] Reading input template...")
    data = InputData(input_path)
    print(f"        Property:  {data.property_name} ({data.property_code})")
    print(f"        Period:    {data.period_label}")
    print(f"        Owner:     {data.owner_name}")
    print(f"        Bookings:  {len(data.reservations)}")
    print(f"        Expenses:  {len(data.expenses)}")
    print(f"        Payouts:   {len(data.payouts)}")

    # Step 2: Validate
    print("\n[ 2/3 ] Running validation (5 rules)...")
    ok, errors, warnings = data.validate()

    if warnings:
        print(f"        ⚠  {len(warnings)} warning(s):")
        for w in warnings:
            print(f"           {w}")

    if errors:
        print(f"\n        ✗  VALIDATION FAILED — {len(errors)} error(s):")
        for e in errors:
            print(f"           {e}")
        print("\n  Report NOT generated. Fix errors and re-run.")
        sys.exit(2)

    print(f"        ✓  All validation rules passed")
    print(f"\n        Gross Revenue:    {data.total_gross:>14,.2f} THB")
    print(f"        TL Commission:    {data.total_tl_comm:>14,.2f} THB")
    print(f"        Total OPEX:       {data.total_opex:>14,.2f} THB")
    print(f"        Net Income:       {data.net_income:>14,.2f} THB")
    print(f"        Total Payouts:    {data.total_payouts:>14,.2f} THB")
    print(f"        Opening Balance:  {data.opening_balance:>14,.2f} THB")
    print(f"        Closing Balance:  {data.closing_balance:>14,.2f} THB")

    # Step 3: Generate report
    print("\n[ 3/3 ] Generating 6-tab Excel report...")
    builder = ReportBuilder(data)
    out_path = builder.build(output_dir)

    print(f"\n{'='*60}")
    print(f"  ✓  Report generated successfully!")
    print(f"     {out_path}")
    print(f"{'='*60}\n")
    return out_path


if __name__ == "__main__":
    main()
