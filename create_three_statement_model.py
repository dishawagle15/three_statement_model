from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Tuple
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

OUTPUT_FILE = "three_statement_model.xlsx"

YEARS_HIST = ["FY2023A", "FY2024A", "FY2025A"]
YEARS_FCST = ["FY2026E", "FY2027E", "FY2028E", "FY2029E", "FY2030E"]

STYLE_DEFAULT = 0
STYLE_TITLE = 1
STYLE_HEADER = 2
STYLE_SECTION = 3
STYLE_CURRENCY = 4
STYLE_PCT = 5
STYLE_INT = 6
STYLE_BOLD = 7


@dataclass
class Cell:
    kind: str
    value: str
    style: int = STYLE_DEFAULT


@dataclass
class Sheet:
    name: str
    cells: Dict[Tuple[int, int], Cell] = field(default_factory=dict)
    col_widths: Dict[int, float] = field(default_factory=dict)
    data_validations: List[Tuple[str, str]] = field(default_factory=list)

    def set_str(self, r: int, c: int, v: str, style: int = STYLE_DEFAULT) -> None:
        self.cells[(r, c)] = Cell("str", v, style)

    def set_num(self, r: int, c: int, v: float, style: int = STYLE_DEFAULT) -> None:
        self.cells[(r, c)] = Cell("num", str(v), style)

    def set_formula(self, r: int, c: int, f: str, style: int = STYLE_DEFAULT) -> None:
        self.cells[(r, c)] = Cell("formula", f[1:] if f.startswith("=") else f, style)

    def set_width(self, c: int, w: float) -> None:
        self.col_widths[c] = w

    def add_validation(self, sqref: str, formula1: str) -> None:
        self.data_validations.append((sqref, formula1))


def col_letter(c: int) -> str:
    out = ""
    while c:
        c, rem = divmod(c - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_ref(r: int, c: int) -> str:
    return f"{col_letter(c)}{r}"


def build_sheet_xml(sheet: Sheet) -> str:
    max_row = max((r for r, _ in sheet.cells), default=1)
    by_row: Dict[int, List[Tuple[int, Cell]]] = {}
    for (r, c), cell in sheet.cells.items():
        by_row.setdefault(r, []).append((c, cell))

    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    ]

    if sheet.col_widths:
        parts.append("<cols>")
        for c, w in sorted(sheet.col_widths.items()):
            parts.append(f'<col min="{c}" max="{c}" width="{w}" customWidth="1"/>')
        parts.append("</cols>")

    parts.append("<sheetData>")
    for r in range(1, max_row + 1):
        cells = sorted(by_row.get(r, []), key=lambda x: x[0])
        if not cells:
            continue
        parts.append(f'<row r="{r}">')
        for c, cell in cells:
            ref = cell_ref(r, c)
            s_attr = f' s="{cell.style}"' if cell.style else ""
            if cell.kind == "str":
                v = escape(cell.value)
                parts.append(f'<c r="{ref}" t="inlineStr"{s_attr}><is><t>{v}</t></is></c>')
            elif cell.kind == "num":
                parts.append(f'<c r="{ref}"{s_attr}><v>{cell.value}</v></c>')
            else:
                f = escape(cell.value)
                parts.append(f'<c r="{ref}"{s_attr}><f>{f}</f><v>0</v></c>')
        parts.append("</row>")
    parts.append("</sheetData>")

    if sheet.data_validations:
        parts.append(f'<dataValidations count="{len(sheet.data_validations)}">')
        for sqref, formula1 in sheet.data_validations:
            parts.append(
                f'<dataValidation type="list" allowBlank="0" sqref="{sqref}"><formula1>{escape(formula1)}</formula1></dataValidation>'
            )
        parts.append("</dataValidations>")

    parts.append("</worksheet>")
    return "".join(parts)


def styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="14"/><color rgb="FF1F4E78"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FF1F4E78"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def workbook_xml(sheet_names: List[str]) -> str:
    sheets = []
    for i, name in enumerate(sheet_names, start=1):
        sheets.append(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>')
    sheets_str = "".join(sheets)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets_str}</sheets>"
        "</workbook>"
    )


def workbook_rels_xml(sheet_count: int) -> str:
    rels = []
    for i in range(1, sheet_count + 1):
        rels.append(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        )
    rels.append(
        f'<Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + "</Relationships>"
    )


def root_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )


def content_types_xml(sheet_count: int) -> str:
    overrides = [
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for i in range(1, sheet_count + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + "".join(overrides)
        + "</Types>"
    )


def scenario_lookup(range_ref: str, idx: int) -> str:
    return (
        f"INDEX({range_ref},MATCH('Control Panel'!$B$3,Assumptions!$A$5:$A$7,0),{idx})"
    )


def global_lookup(row: int, idx: int) -> str:
    return f"INDEX(Assumptions!$B${row}:$F${row},1,{idx})"


def build_control_panel() -> Sheet:
    ws = Sheet("Control Panel")
    ws.set_str(1, 1, "Three Statement Model - Control Panel", STYLE_TITLE)

    ws.set_str(3, 1, "Scenario", STYLE_BOLD)
    ws.set_str(3, 2, "Base")
    ws.add_validation("B3", '"Base,Bull,Bear"')

    ws.set_str(5, 1, "Revenue Growth (FY2026E)")
    ws.set_formula(5, 2, "=INDEX(Assumptions!$B$5:$F$7,MATCH(B3,Assumptions!$A$5:$A$7,0),1)", STYLE_PCT)
    ws.set_str(6, 1, "EBITDA Margin (FY2026E)")
    ws.set_formula(6, 2, "=INDEX(Assumptions!$B$11:$F$13,MATCH(B3,Assumptions!$A$11:$A$13,0),1)", STYLE_PCT)
    ws.set_str(7, 1, "Capex % Revenue (FY2026E)")
    ws.set_formula(7, 2, "=INDEX(Assumptions!$B$17:$F$19,MATCH(B3,Assumptions!$A$17:$A$19,0),1)", STYLE_PCT)
    ws.set_str(8, 1, "Receivables Days")
    ws.set_formula(8, 2, "=Assumptions!B23", STYLE_INT)
    ws.set_str(9, 1, "Inventory Days")
    ws.set_formula(9, 2, "=Assumptions!B24", STYLE_INT)
    ws.set_str(10, 1, "Payables Days")
    ws.set_formula(10, 2, "=Assumptions!B25", STYLE_INT)
    ws.set_str(11, 1, "Tax Rate")
    ws.set_formula(11, 2, "=Assumptions!B26", STYLE_PCT)
    ws.set_str(12, 1, "Interest Rate")
    ws.set_formula(12, 2, "=Assumptions!B27", STYLE_PCT)
    ws.set_str(13, 1, "Minimum Cash % Revenue")
    ws.set_formula(13, 2, "=Assumptions!B28", STYLE_PCT)

    ws.set_str(15, 1, "Model checks", STYLE_BOLD)
    ws.set_str(16, 1, "Latest Balance Check (FY2030E)")
    ws.set_formula(16, 2, "='Projection Model'!F28", STYLE_CURRENCY)

    ws.set_width(1, 42)
    ws.set_width(2, 18)
    return ws


def build_historical() -> Sheet:
    ws = Sheet("Historical Financials")
    ws.set_str(1, 1, "Historical Financials (INR bn)", STYLE_TITLE)

    for c, y in enumerate(YEARS_HIST, start=2):
        ws.set_str(3, c, y, STYLE_HEADER)

    ws.set_str(4, 1, "Income Statement", STYLE_SECTION)
    is_rows = [
        ("Revenue", [974864, 1000122, 1071174]),
        ("COGS", [601959, 598751, 645132]),
        ("Gross Profit", [372905, 401371, 426042]),
        ("Operating Expenses", [147190, 153488, 169560]),
        ("EBITDA", [225715, 247883, 256482]),
        ("Depreciation", [40303, 50832, 53136]),
        ("EBIT", [185412, 197051, 203346]),
        ("Interest", [19571, 23118, 24269]),
        ("Tax", [20376, 25707, 25230]),
        ("Net Income", [73670, 79020, 81309]),
    ]
    r = 5
    for label, vals in is_rows:
        ws.set_str(r, 1, label)
        for c, v in enumerate(vals, start=2):
            ws.set_num(r, c, v, STYLE_CURRENCY)
        r += 1

    ws.set_str(13, 1, "Balance Sheet", STYLE_SECTION)
    bs_rows = [
        ("Cash", [68664, 97225, 106502]),
        ("Accounts Receivable", [28448, 31628, 42121]),
        ("Inventory", [140008, 152770, 146062]),
        ("PPE", [570503, 606084, 683102]),
        ("Total Assets", [1607431, 1755986, 1950121]),
        ("Accounts Payable", [147172, 178377, 186789]),
        ("Debt", [313966, 324622, 347530]),
        ("Equity", [828881, 925788, 1009626]),
        ("Total Liabilities + Equity", [1607431, 1755986, 1950121]),
    ]
    r = 14
    for label, vals in bs_rows:
        ws.set_str(r, 1, label)
        for c, v in enumerate(vals, start=2):
            ws.set_num(r, c, v, STYLE_CURRENCY)
        r += 1

    ws.set_width(1, 32)
    for c in [2, 3, 4]:
        ws.set_width(c, 14)
    return ws


def build_assumptions() -> Sheet:
    ws = Sheet("Assumptions")
    ws.set_str(1, 1, "Assumptions (Scenario-Driven)", STYLE_TITLE)

    # Revenue growth
    ws.set_str(3, 1, "Revenue Growth", STYLE_SECTION)
    ws.set_str(4, 1, "Scenario", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(4, c, y, STYLE_HEADER)
    ws.set_str(5, 1, "Base")
    ws.set_str(6, 1, "Bull")
    ws.set_str(7, 1, "Bear")
    for c, v in enumerate([0.09, 0.085, 0.08, 0.075, 0.07], start=2):
        ws.set_num(5, c, v, STYLE_PCT)
    for c, v in enumerate([0.12, 0.11, 0.10, 0.095, 0.09], start=2):
        ws.set_num(6, c, v, STYLE_PCT)
    for c, v in enumerate([0.05, 0.045, 0.04, 0.04, 0.035], start=2):
        ws.set_num(7, c, v, STYLE_PCT)

    # EBITDA margin
    ws.set_str(9, 1, "EBITDA Margin", STYLE_SECTION)
    ws.set_str(10, 1, "Scenario", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(10, c, y, STYLE_HEADER)
    ws.set_str(11, 1, "Base")
    ws.set_str(12, 1, "Bull")
    ws.set_str(13, 1, "Bear")
    for c, v in enumerate([0.27, 0.272, 0.275, 0.277, 0.28], start=2):
        ws.set_num(11, c, v, STYLE_PCT)
    for c, v in enumerate([0.29, 0.292, 0.295, 0.297, 0.30], start=2):
        ws.set_num(12, c, v, STYLE_PCT)
    for c, v in enumerate([0.24, 0.238, 0.235, 0.233, 0.23], start=2):
        ws.set_num(13, c, v, STYLE_PCT)

    # Capex
    ws.set_str(15, 1, "Capex % Revenue", STYLE_SECTION)
    ws.set_str(16, 1, "Scenario", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(16, c, y, STYLE_HEADER)
    ws.set_str(17, 1, "Base")
    ws.set_str(18, 1, "Bull")
    ws.set_str(19, 1, "Bear")
    for c, v in enumerate([0.11, 0.108, 0.105, 0.102, 0.10], start=2):
        ws.set_num(17, c, v, STYLE_PCT)
    for c, v in enumerate([0.125, 0.122, 0.12, 0.118, 0.115], start=2):
        ws.set_num(18, c, v, STYLE_PCT)
    for c, v in enumerate([0.09, 0.088, 0.085, 0.082, 0.08], start=2):
        ws.set_num(19, c, v, STYLE_PCT)

    # Global assumptions
    ws.set_str(21, 1, "Global Assumptions", STYLE_SECTION)
    ws.set_str(22, 1, "Metric", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(22, c, y, STYLE_HEADER)

    globals_rows = [
        (23, "AR Days", [50, 50, 50, 50, 50], STYLE_INT),
        (24, "Inventory Days", [47, 47, 47, 47, 47], STYLE_INT),
        (25, "AP Days", [42, 42, 42, 42, 42], STYLE_INT),
        (26, "Tax Rate", [0.25, 0.25, 0.25, 0.25, 0.25], STYLE_PCT),
        (27, "Interest Rate", [0.095, 0.095, 0.095, 0.095, 0.095], STYLE_PCT),
        (28, "Minimum Cash % Revenue", [0.03, 0.03, 0.03, 0.03, 0.03], STYLE_PCT),
        (29, "Depreciation % Opening PPE", [0.09, 0.09, 0.09, 0.09, 0.09], STYLE_PCT),
        (30, "Operating Expenses % Revenue", [0.065, 0.064, 0.063, 0.062, 0.061], STYLE_PCT),
        (31, "Other Assets % Revenue", [0.907, 0.907, 0.907, 0.907, 0.907], STYLE_PCT),
        (32, "Other Liabilities % Revenue", [0.379, 0.379, 0.379, 0.379, 0.379], STYLE_PCT),
    ]
    for r, label, vals, style in globals_rows:
        ws.set_str(r, 1, label)
        for c, v in enumerate(vals, start=2):
            ws.set_num(r, c, v, style)

    ws.set_width(1, 36)
    for c in [2, 3, 4, 5, 6]:
        ws.set_width(c, 12.5)
    return ws


def build_working_capital() -> Sheet:
    ws = Sheet("Working Capital Schedule")
    ws.set_str(1, 1, "Working Capital Schedule", STYLE_TITLE)
    ws.set_str(3, 1, "Line Item", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(3, c, y, STYLE_HEADER)

    labels = [
        "Revenue",
        "COGS",
        "AR Days",
        "Inventory Days",
        "AP Days",
        "Accounts Receivable",
        "Inventory",
        "Accounts Payable",
        "Net Working Capital",
        "Change in NWC",
    ]
    for r, label in enumerate(labels, start=4):
        ws.set_str(r, 1, label)

    for c in range(2, 7):
        idx = c - 1
        col = col_letter(c)
        prev = col_letter(c - 1)
        ws.set_formula(4, c, f"='Projection Model'!{col}5", STYLE_CURRENCY)
        ws.set_formula(5, c, f"='Projection Model'!{col}6", STYLE_CURRENCY)
        ws.set_formula(6, c, f"={global_lookup(23, idx)}", STYLE_INT)
        ws.set_formula(7, c, f"={global_lookup(24, idx)}", STYLE_INT)
        ws.set_formula(8, c, f"={global_lookup(25, idx)}", STYLE_INT)
        ws.set_formula(9, c, f"={col}4*{col}6/365", STYLE_CURRENCY)
        ws.set_formula(10, c, f"={col}5*{col}7/365", STYLE_CURRENCY)
        ws.set_formula(11, c, f"={col}5*{col}8/365", STYLE_CURRENCY)
        ws.set_formula(12, c, f"={col}9+{col}10-{col}11", STYLE_CURRENCY)
        if c == 2:
            ws.set_formula(13, c, f"={col}12-('Historical Financials'!D15+'Historical Financials'!D16-'Historical Financials'!D19)", STYLE_CURRENCY)
        else:
            ws.set_formula(13, c, f"={col}12-{prev}12", STYLE_CURRENCY)

    ws.set_width(1, 28)
    for c in [2, 3, 4, 5, 6]:
        ws.set_width(c, 12)
    return ws


def build_ppe_schedule() -> Sheet:
    ws = Sheet("PPE Schedule")
    ws.set_str(1, 1, "PPE Roll-Forward Schedule", STYLE_TITLE)
    ws.set_str(3, 1, "Line Item", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(3, c, y, STYLE_HEADER)

    labels = [
        "Opening PPE",
        "Capex % Revenue",
        "Capex",
        "Depreciation % Opening PPE",
        "Depreciation",
        "Closing PPE",
    ]
    for r, label in enumerate(labels, start=4):
        ws.set_str(r, 1, label)

    for c in range(2, 7):
        idx = c - 1
        col = col_letter(c)
        prev = col_letter(c - 1)
        if c == 2:
            ws.set_formula(4, c, "='Historical Financials'!D17", STYLE_CURRENCY)
        else:
            ws.set_formula(4, c, f"={prev}9", STYLE_CURRENCY)
        ws.set_formula(5, c, f"={scenario_lookup('Assumptions!$B$17:$F$19', idx)}", STYLE_PCT)
        ws.set_formula(6, c, f"='Projection Model'!{col}5*{col}5", STYLE_CURRENCY)
        ws.set_formula(7, c, f"={global_lookup(29, idx)}", STYLE_PCT)
        ws.set_formula(8, c, f"={col}4*{col}7", STYLE_CURRENCY)
        ws.set_formula(9, c, f"={col}4+{col}6-{col}8", STYLE_CURRENCY)

    ws.set_width(1, 34)
    for c in [2, 3, 4, 5, 6]:
        ws.set_width(c, 12)
    return ws


def build_debt_schedule() -> Sheet:
    ws = Sheet("Debt Schedule")
    ws.set_str(1, 1, "Debt Schedule", STYLE_TITLE)
    ws.set_str(3, 1, "Line Item", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(3, c, y, STYLE_HEADER)

    labels = [
        "Opening Debt",
        "Interest Rate",
        "Interest Expense",
        "Beginning Cash",
        "Pre-Financing Cash",
        "Minimum Cash Target",
        "Debt Draw / (Repay)",
        "Closing Debt",
    ]
    for r, label in enumerate(labels, start=4):
        ws.set_str(r, 1, label)

    for c in range(2, 7):
        idx = c - 1
        col = col_letter(c)
        prev = col_letter(c - 1)
        if c == 2:
            ws.set_formula(4, c, "='Historical Financials'!D20", STYLE_CURRENCY)
        else:
            ws.set_formula(4, c, f"={prev}11", STYLE_CURRENCY)
        ws.set_formula(5, c, f"={global_lookup(27, idx)}", STYLE_PCT)
        ws.set_formula(6, c, f"={col}4*{col}5", STYLE_CURRENCY)
        ws.set_formula(7, c, f"='Projection Model'!{col}39", STYLE_CURRENCY)
        ws.set_formula(8, c, f"={col}7+'Projection Model'!{col}33+'Projection Model'!{col}35", STYLE_CURRENCY)
        ws.set_formula(9, c, f"='Projection Model'!{col}5*{global_lookup(28, idx)}", STYLE_CURRENCY)
        ws.set_formula(10, c, f"=MAX(0,{col}9-{col}8)-MIN({col}4,MAX(0,{col}8-{col}9))", STYLE_CURRENCY)
        ws.set_formula(11, c, f"={col}4+{col}10", STYLE_CURRENCY)

    ws.set_width(1, 30)
    for c in [2, 3, 4, 5, 6]:
        ws.set_width(c, 12)
    return ws


def build_projection_model() -> Sheet:
    ws = Sheet("Projection Model")
    ws.set_str(1, 1, "Integrated Projection Model (5-Year Forecast)", STYLE_TITLE)
    ws.set_str(2, 1, "Selected Scenario", STYLE_BOLD)
    ws.set_formula(2, 2, "='Control Panel'!B3", STYLE_BOLD)

    ws.set_str(3, 1, "Line Item", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(3, c, y, STYLE_HEADER)

    labels = [
        "Income Statement",
        "Revenue",
        "COGS",
        "Gross Profit",
        "Operating Expenses",
        "EBITDA",
        "Depreciation",
        "EBIT",
        "Interest",
        "EBT",
        "Tax",
        "Net Income",
        "",
        "Balance Sheet",
        "Cash",
        "Accounts Receivable",
        "Inventory",
        "PPE",
        "Total Assets",
        "Accounts Payable",
        "Debt",
        "Equity",
        "Total Liabilities + Equity",
        "Balance Check",
        "",
        "Cash Flow Statement",
        "Net Income",
        "Add: Depreciation",
        "Less: Change in NWC",
        "Cash from Operations",
        "Capex",
        "Cash from Investing",
        "Debt Issued / (Repaid)",
        "Cash from Financing",
        "Net Change in Cash",
        "Beginning Cash",
        "Ending Cash",
        "",
        "Model Support",
        "Other Assets (Support)",
        "Other Liabilities (Support)",
    ]
    for r, label in enumerate(labels, start=4):
        style = STYLE_SECTION if label in {"Income Statement", "Balance Sheet", "Cash Flow Statement"} else STYLE_DEFAULT
        ws.set_str(r, 1, label, style)

    for c in range(2, 7):
        idx = c - 1
        col = col_letter(c)
        prev = col_letter(c - 1)

        growth = scenario_lookup("Assumptions!$B$5:$F$7", idx)
        margin = scenario_lookup("Assumptions!$B$11:$F$13", idx)
        opex_pct = global_lookup(30, idx)
        tax = global_lookup(26, idx)
        other_assets_pct = global_lookup(31, idx)
        other_liab_pct = global_lookup(32, idx)

        # Income Statement
        if c == 2:
            ws.set_formula(5, c, f"='Historical Financials'!D5*(1+{growth})", STYLE_CURRENCY)
        else:
            ws.set_formula(5, c, f"={prev}5*(1+{growth})", STYLE_CURRENCY)
        ws.set_formula(9, c, f"={col}5*{margin}", STYLE_CURRENCY)
        ws.set_formula(8, c, f"={col}5*{opex_pct}", STYLE_CURRENCY)
        ws.set_formula(6, c, f"={col}5-{col}8-{col}9", STYLE_CURRENCY)
        ws.set_formula(7, c, f"={col}5-{col}6", STYLE_CURRENCY)
        ws.set_formula(10, c, f"='PPE Schedule'!{col}8", STYLE_CURRENCY)
        ws.set_formula(11, c, f"={col}9-{col}10", STYLE_CURRENCY)
        ws.set_formula(12, c, f"='Debt Schedule'!{col}6", STYLE_CURRENCY)
        ws.set_formula(13, c, f"={col}11-{col}12", STYLE_CURRENCY)
        ws.set_formula(14, c, f"=MAX(0,{col}13*{tax})", STYLE_CURRENCY)
        ws.set_formula(15, c, f"={col}13-{col}14", STYLE_CURRENCY)

        # Cash flow
        ws.set_formula(31, c, f"={col}15", STYLE_CURRENCY)
        ws.set_formula(32, c, f"={col}10", STYLE_CURRENCY)
        ws.set_formula(33, c, f"='Working Capital Schedule'!{col}13", STYLE_CURRENCY)
        ws.set_formula(34, c, f"={col}31+{col}32-{col}33", STYLE_CURRENCY)
        ws.set_formula(35, c, f"=-'PPE Schedule'!{col}6", STYLE_CURRENCY)
        ws.set_formula(36, c, f"={col}35", STYLE_CURRENCY)
        ws.set_formula(37, c, f"='Debt Schedule'!{col}10", STYLE_CURRENCY)
        ws.set_formula(38, c, f"={col}37", STYLE_CURRENCY)
        ws.set_formula(39, c, f"={col}34+{col}36+{col}38", STYLE_CURRENCY)
        if c == 2:
            ws.set_formula(40, c, "='Historical Financials'!D14", STYLE_CURRENCY)
        else:
            ws.set_formula(40, c, f"={prev}41", STYLE_CURRENCY)
        ws.set_formula(41, c, f"={col}40+{col}39", STYLE_CURRENCY)

        # Balance Sheet
        ws.set_formula(19, c, f"={col}41", STYLE_CURRENCY)
        ws.set_formula(20, c, f"='Working Capital Schedule'!{col}9", STYLE_CURRENCY)
        ws.set_formula(21, c, f"='Working Capital Schedule'!{col}10", STYLE_CURRENCY)
        ws.set_formula(22, c, f"='PPE Schedule'!{col}9", STYLE_CURRENCY)
        ws.set_formula(44, c, f"={col}5*{other_assets_pct}", STYLE_CURRENCY)
        ws.set_formula(45, c, f"={col}5*{other_liab_pct}", STYLE_CURRENCY)
        ws.set_formula(23, c, f"={col}19+{col}20+{col}21+{col}22+{col}44", STYLE_CURRENCY)
        ws.set_formula(24, c, f"='Working Capital Schedule'!{col}11", STYLE_CURRENCY)
        ws.set_formula(25, c, f"='Debt Schedule'!{col}11", STYLE_CURRENCY)
        if c == 2:
            ws.set_formula(26, c, f"='Historical Financials'!D21+{col}15", STYLE_CURRENCY)
        else:
            ws.set_formula(26, c, f"={prev}26+{col}15", STYLE_CURRENCY)
        ws.set_formula(27, c, f"={col}24+{col}25+{col}26+{col}45", STYLE_CURRENCY)
        ws.set_formula(28, c, f"={col}23-{col}27", STYLE_CURRENCY)

    ws.set_width(1, 32)
    for c in [2, 3, 4, 5, 6]:
        ws.set_width(c, 13)
    return ws


def build_ratio_dashboard() -> Sheet:
    ws = Sheet("Ratio Dashboard")
    ws.set_str(1, 1, "Ratio Dashboard", STYLE_TITLE)
    ws.set_str(2, 1, "Scenario", STYLE_BOLD)
    ws.set_formula(2, 2, "='Control Panel'!B3", STYLE_BOLD)

    ws.set_str(4, 1, "Metric", STYLE_HEADER)
    for c, y in enumerate(YEARS_FCST, start=2):
        ws.set_str(4, c, y, STYLE_HEADER)

    metrics = [
        "EBITDA Margin",
        "ROE",
        "ROCE",
        "Net Debt / EBITDA",
        "Free Cash Flow",
    ]
    for r, m in enumerate(metrics, start=5):
        ws.set_str(r, 1, m)

    for c in range(2, 7):
        col = col_letter(c)
        prev = col_letter(c - 1)
        tax = global_lookup(26, c - 1)

        ws.set_formula(5, c, f"='Projection Model'!{col}9/'Projection Model'!{col}5", STYLE_PCT)
        if c == 2:
            ws.set_formula(6, c, f"='Projection Model'!{col}15/AVERAGE('Historical Financials'!D21,'Projection Model'!{col}26)", STYLE_PCT)
            ws.set_formula(7, c, f"='Projection Model'!{col}11*(1-{tax})/AVERAGE(('Historical Financials'!D21+'Historical Financials'!D20-'Historical Financials'!D14),('Projection Model'!{col}26+'Projection Model'!{col}25-'Projection Model'!{col}19))", STYLE_PCT)
        else:
            ws.set_formula(6, c, f"='Projection Model'!{col}15/AVERAGE('Projection Model'!{prev}26,'Projection Model'!{col}26)", STYLE_PCT)
            ws.set_formula(7, c, f"='Projection Model'!{col}11*(1-{tax})/AVERAGE(('Projection Model'!{prev}26+'Projection Model'!{prev}25-'Projection Model'!{prev}19),('Projection Model'!{col}26+'Projection Model'!{col}25-'Projection Model'!{col}19))", STYLE_PCT)
        ws.set_formula(8, c, f"=('Projection Model'!{col}25-'Projection Model'!{col}19)/'Projection Model'!{col}9", STYLE_PCT)
        ws.set_formula(9, c, f"='Projection Model'!{col}34-'PPE Schedule'!{col}6", STYLE_CURRENCY)

    ws.set_width(1, 28)
    for c in [2, 3, 4, 5, 6]:
        ws.set_width(c, 12)
    return ws


def write_workbook(path: Path, sheets: List[Sheet]) -> None:
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml(len(sheets)))
        zf.writestr("_rels/.rels", root_rels_xml())
        zf.writestr("xl/workbook.xml", workbook_xml([s.name for s in sheets]))
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml(len(sheets)))
        zf.writestr("xl/styles.xml", styles_xml())
        for i, sheet in enumerate(sheets, start=1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", build_sheet_xml(sheet))


def main() -> None:
    sheets = [
        build_control_panel(),
        build_historical(),
        build_assumptions(),
        build_working_capital(),
        build_ppe_schedule(),
        build_debt_schedule(),
        build_projection_model(),
        build_ratio_dashboard(),
    ]
    out = Path(OUTPUT_FILE)
    write_workbook(out, sheets)
    print(f"Created {out.resolve()}")


if __name__ == "__main__":
    main()
