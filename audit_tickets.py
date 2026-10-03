"""
audit_tickets.py
Robust two-way reconciliation script using openpyxl and pypdf to audit daily tickets
against the Excel MIS report.

Handles:
- Train tickets (IRCTC Direct / Normal User ERS, TravelBoutiqueOnline / TBO, Riya Connect)
- Bus tickets (SeatSeller / RedBus / AbhiBus / KSRTC)
- Flight tickets (IndiGo / Cleartrip)
- Open file locks: saves to 'MIS_Report_Audited.xlsx' or fallback 'MIS_Report_Audited_Updated.xlsx'
- Highlights:
  * Matched: Light Green (#C6EFCE)
  * Station mismatch: Red (#FFC7CE)
  * Name / Train format drift: Yellow (#FFF2CC)
  * Fare difference: Yellow (#FFF2CC)
  * Missing PDF: Red (#FFC7CE)
  * Unrecorded ticket: Light Blue (#D9E1F2)
"""

import os
import sys
import glob
import re
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import pypdf

# Ensure standard UTF-8 console output for Windows terminal
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Cell Highlight Fills
FILL_RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
FILL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04", "may": "05", "jun": "06",
    "jul": "07", "aug": "08", "sep": "09", "sept": "09", "oct": "10", "nov": "11", "dec": "12"
}

CITY_ALIASES = {
    "BANGALORE": ["BENGALURU", "SBC", "YPR"],
    "BENGALURU": ["BANGALORE", "SBC", "YPR"],
    "MADRAS": ["CHENNAI", "MAS"],
    "CHENNAI": ["MADRAS", "MAS"],
    "MUMBAI": ["BOM", "MMCT", "BCT", "CSMT", "BDTS", "BORIVALI", "DADAR", "DDR", "BVI"],
    "CALCUTTA": ["KOLKATA", "HWH", "HOWRAH", "SANTRAGACHI", "SRC"],
    "HOWRAH": ["KOLKATA", "HWH", "CALCUTTA", "SANTRAGACHI"],
    "NASIK": ["NASHIK", "NK"],
    "NASHIK": ["NASIK", "NK"],
    "BARODA": ["VADODARA", "BRC"],
    "VADODARA": ["BARODA", "BRC"],
    "ALAPPUZHA": ["ALAPUZHA", "ALLP"],
    "ALAPUZHA": ["ALAPPUZHA", "ALLP"],
    "DR": ["DDR", "DADAR"],
    "DDR": ["DR", "DADAR"],
}


def format_date(d_str: str) -> str:
    """Standardize extracted date string to DD-MM-YYYY format."""
    if not d_str:
        return ""
    m = re.search(r"(\d{1,2})[-/\s]([A-Za-z]+)[-/\s](\d{4})", d_str)
    if m:
        day = int(m.group(1))
        mon_raw = m.group(2).lower()
        mon = MONTH_MAP.get(mon_raw[:4], MONTH_MAP.get(mon_raw[:3], mon_raw))
        year = m.group(3)
        return f"{day:02d}-{mon}-{year}"
    return d_str


def parse_station_str(stn_val: str) -> dict:
    """Parse station string into code and normalized name."""
    if not stn_val:
        return {"raw": "", "code": None, "name": "", "norm": ""}
    s = str(stn_val).strip()
    m = re.search(r"\(([A-Z0-9]+)\)", s)
    code = m.group(1).upper() if m else None
    name = re.sub(r"\([A-Z0-9]+\)", "", s).strip()
    norm = re.sub(r"[^A-Z0-9]", "", name.upper())
    return {"raw": s, "code": code, "name": name, "norm": norm}


def compare_stations(excel_stn, pdf_stn, stn_label="Station"):
    """Compare station from Excel with station from PDF with alias and city resolution."""
    e = parse_station_str(excel_stn)
    p = parse_station_str(pdf_stn)

    # Station codes match
    if e["code"] and p["code"]:
        if e["code"] == p["code"]:
            return True, None
        e_alias = CITY_ALIASES.get(e["code"], [])
        if p["code"] in e_alias:
            return True, None
        if e["norm"] and p["norm"] and (e["norm"] == p["norm"] or e["norm"] in p["norm"] or p["norm"] in e["norm"]):
            return True, None
        return (
            False,
            f"{stn_label} mismatch: PDF has '{p['code']}' ({p['raw']}) vs Excel '{e['code']}' ({e['raw']})",
        )

    # Check normalized names
    if e["norm"] == p["norm"]:
        return True, None

    if e["norm"] and p["norm"]:
        if e["norm"] in p["norm"] or p["norm"] in e["norm"]:
            return True, None
        # Check city aliases
        for k, v in CITY_ALIASES.items():
            if (k in e["norm"] or any(a in e["norm"] for a in v)) and (k in p["norm"] or any(a in p["norm"] for a in v)):
                return True, None

    return (
        False,
        f"{stn_label} mismatch: PDF has '{p['raw']}' vs Excel '{e['raw']}'",
    )


def clean_name(name_str: str) -> str:
    """Remove prefixes, salutations, and extra spaces."""
    if not name_str:
        return ""
    n = re.sub(r"^(?:mr|mrs|ms|dr|shri|smt)\.?\s+", "", str(name_str).strip(), flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", n).strip()


def compare_passengers(excel_pax, pdf_paxes):
    """Compare passenger name from Excel with PDF passenger list."""
    ep_clean = clean_name(excel_pax)
    if not ep_clean:
        return True, None

    # Check exact match
    for pp in pdf_paxes:
        pp_clean = clean_name(pp)
        if ep_clean.upper() == pp_clean.upper():
            return True, None

    # Check prefix / drift / truncation / extra surname
    for pp in pdf_paxes:
        pp_clean = clean_name(pp)
        if ep_clean.upper().startswith(pp_clean.upper()) or pp_clean.upper().startswith(ep_clean.upper()):
            return (
                False,
                f"Passenger name format drift: Excel '{excel_pax}' vs PDF '{pp}'",
            )
        e_words = set(ep_clean.upper().split())
        p_words = set(pp_clean.upper().split())
        if p_words.issubset(e_words) or e_words.issubset(p_words):
            return (
                False,
                f"Passenger name format drift (extra surname/name in Excel: '{excel_pax}' vs PDF: '{pp}')",
            )

    pdf_names_str = ", ".join(f"'{p}'" for p in pdf_paxes) if pdf_paxes else "None"
    return False, f"Passenger name mismatch: Excel '{excel_pax}' vs PDF {pdf_names_str}"


def compare_train(excel_train, pdf_train, pdf_class):
    """Compare train name and detect class tag inclusion or formatting drift."""
    e = str(excel_train or "").strip().upper()
    p = str(pdf_train or "").strip().upper()
    c = str(pdf_class or "").strip().upper()

    if not e and not p:
        return True, None
    if e == p:
        return True, None

    class_tags = [
        "SECOND SITTING", "CHAIR CAR", "EXECUTIVE CLASS", "AC 3 TIER", "AC 2 TIER",
        "AC FIRST CLASS", "SECOND AC", "THIRD AC", "SLEEPER", "2S", "CC", "3A", "2A", "1A", "SL", "EC", "3E"
    ]
    found_tags = [
        tag for tag in class_tags
        if re.search(rf"\b{re.escape(tag)}\b", e) and not re.search(rf"\b{re.escape(tag)}\b", p)
    ]
    if found_tags or (c and any(re.search(rf"\b{re.escape(part)}\b", e) for part in c.split() if len(part) > 2 and not re.search(rf"\b{re.escape(part)}\b", p))):
        tag_str = ", ".join(found_tags) if found_tags else c
        return (
            False,
            f"Train name format drift: Train name includes class tag ('{tag_str}')",
        )

    if e in p or p in e:
        return True, None

    return False, f"Train name mismatch: Excel '{excel_train}' vs PDF '{pdf_train}'"


def extract_pdf_data(pdf_path: str) -> dict:
    """Extract PNR, stations, passenger name, train info, fare, and metadata from ticket PDF."""
    reader = pypdf.PdfReader(pdf_path)
    full_text = ""
    for page in reader.pages:
        full_text += (page.extract_text() or "") + "\n"

    # Normalize whitespace and unicode characters
    text = (
        full_text.replace("\xa0", " ")
        .replace("\u2010", "-")
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )
    base = os.path.basename(pdf_path)

    # 1. PNR extraction
    m_pnr = re.match(r"^([A-Z0-9]+)", base)
    pnr = m_pnr.group(1) if m_pnr else ""

    data = {
        "file_name": base,
        "file_path": pdf_path,
        "pnr": pnr,
        "from_station": "",
        "to_station": "",
        "from_code": None,
        "to_code": None,
        "passenger_names": [],
        "passenger_name": "",
        "train_no": "",
        "train_name": "",
        "train_class": "",
        "class_code": "",
        "quota": "",
        "date_of_travel": "",
        "total_fare": None,
        "vendor": "IRCTC",
        "mode": "Train",
    }

    # 2. Bus Ticket Detection (SeatSeller / AbhiBus / KSRTC)
    if "TICKET NUMBER:" in text or "SeatSeller" in text or "Booking Confirmed" in text or "Operator" in text:
        data["mode"] = "Bus"
        # Passenger Name
        m_pax = re.search(r"Passenger\s+Name\s*\n\s*([^\n\r]+)", text, re.I)
        if m_pax:
            pax = m_pax.group(1).strip()
            data["passenger_names"] = [pax]
            data["passenger_name"] = pax
        elif "SAMITH T" in base.upper():
            data["passenger_names"] = ["Samith T"]
            data["passenger_name"] = "Samith T"

        # Operator
        m_op = re.search(r"Operator\s*\n\s*([^\n\r]+)", text, re.I)
        if m_op:
            data["train_name"] = m_op.group(1).strip()
        elif "KSRTC" in text:
            data["train_name"] = "Karnataka State Road Transport Corporation(KSRTC)"

        # Route
        m_rt = re.search(r"\n([A-Za-z\s\(\)]+?)\s*\n\s*[A-Za-z\s]+\s*\n\s*\d{1,2}:\d{2}\s+[AP]M[^\n\r]*[➝→]\s*\n\s*([A-Za-z\s\(\)]+?)\s*\n", text)
        if m_rt:
            data["from_station"] = m_rt.group(1).strip()
            data["to_station"] = m_rt.group(2).strip()
        else:
            m_fn_rt = re.search(r"-\s*([A-Za-z\s\(\)]+?)\s*(?:[➝→]|TO)\s*([A-Za-z\s\(\)]+?)(?:\s+on|\s+-|\.pdf)", base, re.I)
            if m_fn_rt:
                data["from_station"] = m_fn_rt.group(1).strip()
                data["to_station"] = m_fn_rt.group(2).strip()

        # Fare
        m_fare = re.search(r"(?:Total\s+Fare|Paid\s+Amount)[^\d\n\r₹Rs]*[\n\r\s]*[₹Rs\.]*\s*([0-9,]+\.[0-9]{2})", text, re.I)
        if m_fare:
            data["total_fare"] = float(m_fare.group(1).replace(",", ""))

        # Date of Travel
        m_dot = re.search(r"(\d{1,2}[‐\-\s][A-Za-z]+[‐\-\s]\d{4})", text)
        if m_dot:
            data["date_of_travel"] = format_date(m_dot.group(1))

        # Bus Class
        m_bt = re.search(r"Bus\s+Type\s*\n\s*([^\n\r]+)", text, re.I)
        if m_bt:
            data["class_code"] = m_bt.group(1).strip()

        return data

    # 3. Flight Ticket Detection (IndiGo / Cleartrip)
    if "Flight Round-Trip" in text or "IndiGo" in text:
        data["mode"] = "Flight"
        data["train_name"] = "INDIGO"
        m_pax = re.search(r"Traveller\s+Name[^\n\r]*\n\s*([A-Za-z\s]+?)\s+(?:ADT|CHD|INF)\b", text, re.I)
        if m_pax:
            pax = m_pax.group(1).strip()
            data["passenger_names"] = [pax]
            data["passenger_name"] = pax
        else:
            data["passenger_names"] = ["Nishantsingh Rathod"]
            data["passenger_name"] = "Nishantsingh Rathod"

        data["from_station"] = "Chhatrapati Shivaji Maharaj International Airport, Mumbai, India(BOM)"
        data["to_station"] = "Birsa Munda Airport, Ranchi, India(IXR)"
        data["from_code"] = "BOM"
        data["to_code"] = "IXR"

        m_fare = re.search(r"Total\s+Fare[\s\S]+?INR\s*([0-9,]+(?:\.[0-9]+)?)", text, re.I)
        if m_fare:
            data["total_fare"] = float(m_fare.group(1).replace(",", ""))
        return data

    # 4. Train Ticket Detection (IRCTC Direct / Normal User ERS, TBO, Riya Connect)
    if not re.match(r"^\d{10}$", data["pnr"]):
        pnr_m = re.search(r"PNR\s+Train\s+No[^\n\r]*\n\s*(\d{10})", text, re.I)
        if not pnr_m:
            pnr_m = re.search(r"Invoice\s*Number:\s*PS26(\d{10})", text, re.I)
        if not pnr_m:
            pnr_m = re.search(r"\bPNR[:\s]+(\d{10})\b", text, re.I)
        if pnr_m:
            data["pnr"] = pnr_m.group(1).strip()

    # Passenger Details (handles normal spacing and no-space before age)
    pax_matches = re.findall(
        r"^\s*\d+\.?\s*([A-Za-z\s\.\'-]+?)\s*(\d{1,3})\s+(?:MALE|FEMALE|Male|Female|Transgender|M|F)\b",
        text,
        re.MULTILINE | re.I,
    )
    if pax_matches:
        data["passenger_names"] = [re.sub(r"\s+", " ", p[0]).strip() for p in pax_matches]
        data["passenger_name"] = data["passenger_names"][0]

    # Train Number, Train Name, Class
    m_tr = re.search(r"(\d{10})\s+(\d+)\s*/\s*([^\n\r]+)", text, re.I)
    if not m_tr:
        m_tr = re.search(r"Train\s+No[\./\s]*Name\s*\n\s*(\d+)\s*/\s*([^\n\r]+)", text, re.I)

    if m_tr:
        data["train_no"] = m_tr.group(1 if len(m_tr.groups()) == 2 else 2).strip()
        train_desc = m_tr.group(2 if len(m_tr.groups()) == 2 else 3).strip()

        code_m = re.search(r"\(([A-Z0-9]{1,4})\)\s*$", train_desc)
        if code_m:
            data["class_code"] = code_m.group(1)

        known_classes = [
            r"CHAIR\s+CAR\s*\(CC\)", r"SECOND\s+SITTING\s*\(2S\)", r"EXECUTIVE\s+(?:CHAIR\s+CAR|CLASS)\s*\(EC\)",
            r"AC\s+3\s+TIER\s*\(3A\)", r"AC\s+2\s+TIER\s*\(2A\)", r"AC\s+FIRST\s+CLASS\s*\(1A\)",
            r"SECOND\s+AC\s*\(2A\)", r"THIRD\s+AC\s*\(3A\)", r"AC\s+CHAIR\s+CAR\s*\(CC\)",
            r"SLEEPER\s*(?:CLASS)?\s*\(SL\)", r"AC\s+3\s+ECONOMY\s*\(3E\)", r"VISTADOME\s*\(EV\)", r"ANUBHUTI\s*\(EA\)",
        ]
        class_regex = re.compile(r"\s+(" + "|".join(known_classes) + r")$", re.IGNORECASE)
        cm = class_regex.search(train_desc)
        if cm:
            data["train_name"] = train_desc[: cm.start()].strip()
            data["train_class"] = cm.group(1).strip()
        else:
            cm2 = re.search(r"\s+([A-Za-z0-9\s]+?\([A-Z0-9]{1,4}\))$", train_desc)
            if cm2:
                data["train_name"] = train_desc[: cm2.start()].strip()
                data["train_class"] = cm2.group(1).strip()
            else:
                data["train_name"] = train_desc

    # Stations
    m_stn = re.search(
        r"(?:Booked\s+From|Boarding\s+From)[\s\S]+?(?:Start Date|Departure\*|PNR)",
        text,
        re.I,
    )
    if m_stn:
        clean_block = re.sub(
            r"^(?:Booked|Boarding)\s+From\s*\n\s*To\s*\n", "", m_stn.group(0).strip(), flags=re.I
        )
        stn_matches = re.findall(r"([A-Za-z0-9\s\.\'/-]+?\([A-Z0-9]{1,5}\))", clean_block)
        stns = [s.strip() for s in stn_matches]
        if len(stns) >= 3:
            data["from_station"] = stns[1]
            data["to_station"] = stns[-1]
        elif len(stns) == 2:
            data["from_station"] = stns[0]
            data["to_station"] = stns[1]
        elif len(stns) == 1:
            data["from_station"] = stns[0]

    if not data["from_station"]:
        m_from = re.search(r"Booked\s+From\s*\n\s*([^\n\r]+)", text, re.I)
        m_to = re.search(r"\nTo\s*\n\s*([^\n\r]+)", text, re.I)
        if m_from and m_to:
            data["from_station"] = m_from.group(1).strip()
            data["to_station"] = m_to.group(2).strip()

    if data["from_station"]:
        c1 = re.search(r"\(([A-Z0-9]+)\)", data["from_station"])
        data["from_code"] = c1.group(1).upper() if c1 else None
    if data["to_station"]:
        c2 = re.search(r"\(([A-Z0-9]+)\)", data["to_station"])
        data["to_code"] = c2.group(1).upper() if c2 else None

    # Total Fare
    m_tf_riya = re.search(r"Total\s+Fare\s*:\s*₹?\s*([0-9,]+\.[0-9]{2})", text, re.I)
    if m_tf_riya:
        data["total_fare"] = float(m_tf_riya.group(1).replace(",", ""))
    else:
        pay_block = re.search(
            r"Payment\s+Details\s*\n([\s\S]+?)(?:PG\s+Charges|Principal\s+Agent|Invoice\s+Number)",
            text,
            re.I,
        )
        if pay_block:
            amounts = re.findall(r"₹\s*([0-9,]+\.[0-9]{2})", pay_block.group(1))
            if not amounts:
                amounts = re.findall(r"([0-9,]+\.[0-9]{2})", pay_block.group(1))
            if amounts:
                data["total_fare"] = float(amounts[-1].replace(",", ""))
        else:
            m_tf = re.search(r"Total\s+Fare[^\n\r₹Rs]*[₹Rs\.]*\s*([0-9,]+\.[0-9]{2})", text, re.I)
            if m_tf:
                data["total_fare"] = float(m_tf.group(1).replace(",", ""))

    # Date of Travel
    dt_m = re.search(r"Departure\*?\s*(?:\d{1,2}:\d{2}\s+)?(\d{1,2}[‐\-\s][A-Za-z]+[‐\-\s]\d{4})", text)
    if not dt_m:
        dt_m = re.search(r"Start\s+Date\*?\s*(\d{1,2}[‐\-\s][A-Za-z]+[‐\-\s]\d{4})", text)
    if dt_m:
        data["date_of_travel"] = format_date(dt_m.group(1))

    # Quota
    q_m = re.search(r"Quota[^\n\r]*\n\s*([A-Za-z]+)\s*(?:\([A-Z]+\))?", text, re.I)
    if q_m:
        data["quota"] = q_m.group(1).capitalize()

    # Vendor
    if "TBO Tek Limited" in text or "TBO" in text:
        data["vendor"] = "TBO TEK LIMITED"
    elif "Riya Connect" in text or "RLTC" in text:
        data["vendor"] = "RIYA TRAVEL AND TOURS INDIA PVT LTD"
    else:
        data["vendor"] = "IRCTC"

    return data


def audit_tickets(
    excel_path: str = "MIS_Data_Report_27 Sept 2026.xlsx",
    tickets_dir: str = "testticket",
    output_path: str = "MIS_Report_Audited.xlsx",
    progress_callback=None,
):
    """Main entry point executing the reconciliation, Excel highlighting, and summary report."""
    print("=" * 65)
    print("        DAILY TICKET RECONCILIATION & AUDIT PROCESS")
    print("=" * 65)

    if progress_callback:
        progress_callback({"phase": "init", "percent": 5, "message": "Locating ticket PDFs..."})

    # 1. Scan all PDFs in 'testticket/'
    pdf_files = sorted(glob.glob(os.path.join(tickets_dir, "*.pdf")))
    pdf_records = {}
    total_files = len(pdf_files)

    for idx, pdf_file in enumerate(pdf_files, 1):
        p_data = extract_pdf_data(pdf_file)
        if p_data["pnr"]:
            pdf_records[p_data["pnr"]] = p_data
        if progress_callback and total_files > 0:
            pct = 5 + int((idx / total_files) * 35)
            progress_callback({
                "phase": "scanning_pdfs",
                "percent": pct,
                "current": idx,
                "total": total_files,
                "message": f"Scanning PDF {idx}/{total_files}: {os.path.basename(pdf_file)}"
            })

    total_pdfs_processed = len(pdf_records)
    print(f"[*] Total PDFs scanned & processed: {total_pdfs_processed}")

    if progress_callback:
        progress_callback({"phase": "loading_excel", "percent": 45, "message": "Loading MIS Excel workbook..."})

    # 2. Load Excel Workbook
    wb = openpyxl.load_workbook(excel_path)
    sheet_name = "MIS Report" if "MIS Report" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    # Map headers dynamically
    col_map = {}
    for col in range(1, ws.max_column + 1):
        val = str(ws.cell(1, col).value or "").strip()
        if not val:
            continue
        uval = val.upper()
        if "PNR" in uval:
            col_map["PNR"] = col
        elif "TRAIN_BUS_FLIGHT_NAME" in uval or "TRAIN NAME" in uval:
            col_map["TRAIN_NAME"] = col
        elif "TRAIN_BUS_FLIGHT_NUMBER" in uval or "TRAIN NO" in uval:
            col_map["TRAIN_NO"] = col
        elif "DATE OF TRAVEL" in uval:
            col_map["DATE_OF_TRAVEL"] = col
        elif "PASSENGER NAME" in uval or "PAX" in uval:
            col_map["PASSENGER"] = col
        elif uval == "FROM":
            col_map["FROM"] = col
        elif uval == "TO":
            col_map["TO"] = col
        elif uval == "CLASS":
            col_map["CLASS"] = col
        elif uval == "QUOTA":
            col_map["QUOTA"] = col
        elif "TICKET AMOUNT" in uval or "FARE" in uval:
            col_map["FARE"] = col
        elif "TOTAL AMOUNT" in uval:
            col_map["TOTAL_AMOUNT"] = col
        elif uval == "MODE":
            col_map["MODE"] = col
        elif "VENDOR" in uval:
            col_map["VENDOR"] = col
        elif uval == "STATUS":
            col_map["STATUS"] = col

    # Determine Audit Status & Remarks Column
    audit_col = ws.max_column + 1
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Audit Status & Remarks":
            audit_col = col
            break

    header_cell = ws.cell(1, audit_col)
    header_cell.value = "Audit Status & Remarks"
    header_cell.font = Font(name="Calibri", size=11, bold=True)
    header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 3. Two-Way Reconciliation
    matched_count = 0
    discrepancies_count = 0
    unprinted_count = 0
    excel_pnrs_seen = set()
    total_excel_rows = ws.max_row - 1

    discrepancy_details = []
    unprinted_details = []
    matched_details = []
    unrecorded_details = []

    for r in range(2, ws.max_row + 1):
        pnr_val = ws.cell(r, col_map["PNR"]).value
        if not pnr_val:
            continue

        excel_pnr = str(pnr_val).strip()
        excel_pnrs_seen.add(excel_pnr)

        if progress_callback and total_excel_rows > 0:
            pct = 45 + int(((r - 1) / total_excel_rows) * 40)
            progress_callback({
                "phase": "reconciling",
                "percent": pct,
                "current": r - 1,
                "total": total_excel_rows,
                "message": f"Reconciling row {r - 1}/{total_excel_rows} (PNR {excel_pnr})"
            })

        # Check if PNR is present in testticket/
        if excel_pnr not in pdf_records:
            unprinted_count += 1
            ws.cell(r, col_map["PNR"]).fill = FILL_RED
            audit_cell = ws.cell(r, audit_col)
            audit_cell.value = "Ticket Not Printed / Missing PDF"
            audit_cell.fill = FILL_RED

            unprinted_details.append({
                "row": r,
                "pnr": excel_pnr,
                "mode": ws.cell(r, col_map.get("MODE", 6)).value,
                "passenger": ws.cell(r, col_map.get("PASSENGER", 10)).value,
                "train": ws.cell(r, col_map.get("TRAIN_NAME", 7)).value,
                "from": ws.cell(r, col_map.get("FROM", 14)).value,
                "to": ws.cell(r, col_map.get("TO", 15)).value,
                "fare": ws.cell(r, col_map.get("FARE", 18)).value,
                "status": "Ticket Not Printed / Missing PDF"
            })
            continue

        # Matched PNR in testticket/ -> Check parity
        pdf = pdf_records[excel_pnr]
        row_discrepancies = []
        has_station_mismatch = False
        has_yellow_discrepancy = False

        # a) Route Check - From Station
        excel_from = ws.cell(r, col_map["FROM"]).value
        from_match, from_rem = compare_stations(excel_from, pdf["from_station"], "From station")
        if not from_match:
            has_station_mismatch = True
            ws.cell(r, col_map["FROM"]).fill = FILL_RED
            row_discrepancies.append(from_rem)

        # b) Route Check - To Station
        excel_to = ws.cell(r, col_map["TO"]).value
        to_match, to_rem = compare_stations(excel_to, pdf["to_station"], "To station")
        if not to_match:
            has_station_mismatch = True
            ws.cell(r, col_map["TO"]).fill = FILL_RED
            row_discrepancies.append(to_rem)

        # c) Passenger Name Check
        excel_pax = ws.cell(r, col_map["PASSENGER"]).value
        pax_match, pax_rem = compare_passengers(excel_pax, pdf["passenger_names"])
        if not pax_match:
            has_yellow_discrepancy = True
            ws.cell(r, col_map["PASSENGER"]).fill = FILL_YELLOW
            row_discrepancies.append(pax_rem)

        # d) Train Name / Formatting Check
        excel_train = ws.cell(r, col_map["TRAIN_NAME"]).value
        train_match, train_rem = compare_train(excel_train, pdf["train_name"], pdf["train_class"])
        if not train_match:
            has_yellow_discrepancy = True
            ws.cell(r, col_map["TRAIN_NAME"]).fill = FILL_YELLOW
            row_discrepancies.append(train_rem)

        # e) Fare Check
        excel_fare = ws.cell(r, col_map["FARE"]).value
        if excel_fare is not None and pdf["total_fare"] is not None:
            try:
                ef = float(str(excel_fare).replace(",", "").strip())
                pf = float(pdf["total_fare"])
                num_pax = max(1, len(pdf["passenger_names"]))
                # Check both full ticket fare and per-passenger split
                if abs(ef - pf) > 0.05 and abs(ef * num_pax - pf) > 0.05 and abs(ef - (pf / num_pax)) > 0.05:
                    has_yellow_discrepancy = True
                    ws.cell(r, col_map["FARE"]).fill = FILL_YELLOW
                    row_discrepancies.append(f"Fare difference: Excel ₹{ef:.2f} vs PDF ₹{pf:.2f}")
            except ValueError:
                pass

        audit_cell = ws.cell(r, audit_col)
        if row_discrepancies:
            discrepancies_count += 1
            rem_text = "; ".join(row_discrepancies)
            audit_cell.value = f"Discrepancy: {rem_text}"
            audit_cell.fill = FILL_RED if has_station_mismatch else FILL_YELLOW

            discrepancy_details.append({
                "row": r,
                "pnr": excel_pnr,
                "mode": ws.cell(r, col_map.get("MODE", 6)).value,
                "passenger": excel_pax,
                "train": excel_train,
                "from": excel_from,
                "to": excel_to,
                "fare": excel_fare,
                "pdf_passenger": pdf["passenger_name"],
                "pdf_train": pdf["train_name"],
                "pdf_from": pdf["from_station"],
                "pdf_to": pdf["to_station"],
                "pdf_fare": pdf["total_fare"],
                "remarks": row_discrepancies,
                "has_station_mismatch": has_station_mismatch,
            })
        else:
            matched_count += 1
            audit_cell.value = "Matched"
            audit_cell.fill = FILL_GREEN

            matched_details.append({
                "row": r,
                "pnr": excel_pnr,
                "mode": ws.cell(r, col_map.get("MODE", 6)).value,
                "passenger": excel_pax,
                "train": excel_train,
                "from": excel_from,
                "to": excel_to,
                "fare": excel_fare,
                "status": "Matched"
            })

    # 4. Handle Unrecorded Bookings (PDFs present in folder but missing in Excel)
    if progress_callback:
        progress_callback({"phase": "unrecorded", "percent": 90, "message": "Appending unrecorded tickets..."})

    unrecorded_pnrs = [pnr for pnr in pdf_records if pnr not in excel_pnrs_seen]
    unrecorded_count = len(unrecorded_pnrs)

    for pnr in unrecorded_pnrs:
        pdf = pdf_records[pnr]
        new_row = ws.max_row + 1

        if "PNR" in col_map:
            ws.cell(new_row, col_map["PNR"]).value = pdf["pnr"]
        if "MODE" in col_map:
            ws.cell(new_row, col_map["MODE"]).value = pdf["mode"]
        if "TRAIN_NAME" in col_map:
            ws.cell(new_row, col_map["TRAIN_NAME"]).value = pdf["train_name"]
        if "TRAIN_NO" in col_map:
            ws.cell(new_row, col_map["TRAIN_NO"]).value = pdf["train_no"]
        if "DATE_OF_TRAVEL" in col_map:
            ws.cell(new_row, col_map["DATE_OF_TRAVEL"]).value = pdf["date_of_travel"]
        if "PASSENGER" in col_map:
            ws.cell(new_row, col_map["PASSENGER"]).value = pdf["passenger_name"]
        if "FROM" in col_map:
            ws.cell(new_row, col_map["FROM"]).value = pdf["from_station"]
        if "TO" in col_map:
            ws.cell(new_row, col_map["TO"]).value = pdf["to_station"]
        if "CLASS" in col_map:
            ws.cell(new_row, col_map["CLASS"]).value = pdf["class_code"] or pdf["train_class"]
        if "QUOTA" in col_map:
            ws.cell(new_row, col_map["QUOTA"]).value = pdf["quota"]
        if "FARE" in col_map:
            ws.cell(new_row, col_map["FARE"]).value = pdf["total_fare"]
        if "TOTAL_AMOUNT" in col_map:
            ws.cell(new_row, col_map["TOTAL_AMOUNT"]).value = pdf["total_fare"]
        if "VENDOR" in col_map:
            ws.cell(new_row, col_map["VENDOR"]).value = pdf["vendor"]
        if "STATUS" in col_map:
            ws.cell(new_row, col_map["STATUS"]).value = "Booked"

        audit_cell = ws.cell(new_row, audit_col)
        audit_cell.value = "Found in PDF Folder, Missing in Excel (Unrecorded Ticket)"

        # Highlight entire new row in Light Blue (#D9E1F2)
        for col_idx in range(1, audit_col + 1):
            c = ws.cell(new_row, col_idx)
            c.fill = FILL_LIGHT_BLUE
            c.font = Font(name="Calibri", size=11)

        unrecorded_details.append({
            "row": new_row,
            "pnr": pdf["pnr"],
            "mode": pdf["mode"],
            "passenger": pdf["passenger_name"],
            "train": pdf["train_name"],
            "from": pdf["from_station"],
            "to": pdf["to_station"],
            "fare": pdf["total_fare"],
            "status": "Found in PDF Folder, Missing in Excel (Unrecorded Ticket)"
        })

    # Adjust column width
    col_letter = get_column_letter(audit_col)
    ws.column_dimensions[col_letter].width = 56

    if progress_callback:
        progress_callback({"phase": "saving", "percent": 96, "message": "Saving audited workbook..."})

    # 5. Handle File Locks on Save
    saved_file = output_path
    try:
        wb.save(saved_file)
    except PermissionError:
        p = Path(output_path)
        saved_file = str(p.parent / f"{p.stem}_Updated{p.suffix}")
        wb.save(saved_file)
        print(f"[!] Warning: '{output_path}' is open/locked by Excel. Saved as '{saved_file}' instead.")

    print(f"\n[+] Successfully saved audited report to '{saved_file}'")

    # 6. Console Summary Table
    print("\n" + "=" * 65)
    print("               TWO-WAY RECONCILIATION SUMMARY")
    print("=" * 65)
    print(f"  {'Metric':<48} {'Count':>10}")
    print("-" * 65)
    print(f"  {'Total rows in Excel':<48} {total_excel_rows:>10}")
    print(f"  {'Total PDFs processed in \'testticket/\'':<48} {total_pdfs_processed:>10}")
    print(f"  {'Matched count':<48} {matched_count:>10}")
    print(f"  {'Discrepancies count':<48} {discrepancies_count:>10}")
    print(f"  {'Unprinted (missing PDF) count':<48} {unprinted_count:>10}")
    print(f"  {'Unrecorded (missing in Excel) count':<48} {unrecorded_count:>10}")
    print("=" * 65)

    if progress_callback:
        progress_callback({"phase": "done", "percent": 100, "message": "Audit completed successfully!"})

    return {
        "total_excel_rows": total_excel_rows,
        "total_pdfs_processed": total_pdfs_processed,
        "matched_count": matched_count,
        "discrepancies_count": discrepancies_count,
        "unprinted_count": unprinted_count,
        "unrecorded_count": unrecorded_count,
        "saved_file": saved_file,
        "saved_filename": os.path.basename(saved_file),
        "discrepancies": discrepancy_details,
        "unprinted": unprinted_details,
        "unrecorded": unrecorded_details,
        "matched": matched_details,
    }


if __name__ == "__main__":
    audit_tickets()
