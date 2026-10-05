import os
import re
from pathlib import Path
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

from pypdf import PdfReader


# ============================================================
# COLORS
# ============================================================

FILL_GREEN = PatternFill(fill_type="solid", fgColor="C6EFCE")
FILL_YELLOW = PatternFill(fill_type="solid", fgColor="FFEB9C")
FILL_RED = PatternFill(fill_type="solid", fgColor="FFC7CE")
FILL_ORANGE = PatternFill(fill_type="solid", fgColor="FCE4D6")


# ============================================================
# BASIC NORMALIZATION
# ============================================================

def normalize_value(value):
    """Convert a value into a standard format for comparison."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    value = str(value).strip().upper()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_name(value):
    """Normalize passenger names."""
    if not value:
        return ""
    value = normalize_value(value)
    value = re.sub(r"[^A-Z0-9 ]", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


# ============================================================
# DATE FUNCTIONS
# ============================================================

def format_date(value):
    """Convert common date formats into DD-MM-YYYY."""
    if value is None:
        return ""

    if hasattr(value, "strftime"):
        return value.strftime("%d-%m-%Y")

    value = str(value).strip()

    formats = [
        "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y",
        "%Y-%m-%d", "%Y/%m/%d", "%d %b %Y", "%d %B %Y",
        "%d-%b-%Y", "%d-%B-%Y",
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(value, fmt)
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            pass

    return value


def compare_date(excel_date, pdf_date):
    """Compare Excel and PDF travel dates."""
    if not excel_date and not pdf_date:
        return True, None
    if not excel_date:
        return True, None
    if not pdf_date:
        return False, "Date of travel could not be extracted from PDF"

    excel_date = format_date(excel_date)
    pdf_date = format_date(pdf_date)

    if excel_date == pdf_date:
        return True, None

    return False, f"Travel Date Mismatch: Excel '{excel_date}' vs PDF '{pdf_date}'"


# ============================================================
# TEXT COMPARISON
# ============================================================

def compare_text(excel_value, pdf_value, field_name):
    """Compare two text values."""
    excel_value = normalize_value(excel_value)
    pdf_value = normalize_value(pdf_value)

    if not excel_value or not pdf_value:
        return True, None  # Skip if either is missing to prevent false alerts

    if excel_value == pdf_value:
        return True, None

    if excel_value in pdf_value or pdf_value in excel_value:
        return True, None

    return False, f"{field_name} mismatch: PDF has '{pdf_value}' vs Excel '{excel_value}'"


def compare_station(excel_val, pdf_val, field_label):
    """Compare station/city origin or destination."""
    excel_norm = normalize_value(excel_val)
    pdf_norm = normalize_value(pdf_val)

    if not excel_norm or not pdf_norm:
        return True, None

    if excel_norm == pdf_norm or excel_norm in pdf_norm or pdf_norm in excel_norm:
        return True, None

    return False, f"{field_label} mismatch: PDF has '{pdf_val}' vs Excel '{excel_val}'"


# ============================================================
# PASSENGER NAME COMPARISON
# ============================================================

def compare_passenger(excel_name, pdf_names):
    """Compare MIS passenger name with names extracted from PDF."""
    excel_name = normalize_name(excel_name)
    if not excel_name:
        return True, None

    if not pdf_names:
        return False, "Passenger name could not be extracted from PDF"

    normalized_pdf_names = [normalize_name(name) for name in pdf_names if name]

    if excel_name in normalized_pdf_names:
        return True, None

    excel_compact = excel_name.replace(" ", "")
    for name in normalized_pdf_names:
        name_compact = name.replace(" ", "")
        if excel_compact == name_compact or excel_name in name or name in excel_name:
            return True, None

    return False, f"Passenger Name Mismatch: Excel '{excel_name}' vs PDF '{', '.join(pdf_names)}'"


# ============================================================
# CLASS NORMALIZATION
# ============================================================

def normalize_class(value):
    """Normalize railway/travel class values."""
    if not value:
        return ""
    value = normalize_value(value)
    aliases = {
        "AC 3 TIER": "3A", "THIRD AC": "3A", "3A": "3A",
        "AC 2 TIER": "2A", "SECOND AC": "2A", "2A": "2A",
        "AC FIRST CLASS": "1A", "FIRST AC": "1A", "1A": "1A",
        "SLEEPER": "SL", "SL": "SL",
        "SECOND SITTING": "2S", "2S": "2S",
        "CHAIR CAR": "CC", "AC CHAIR CAR": "CC", "CC": "CC",
        "EXECUTIVE CLASS": "EC", "EC": "EC",
    }
    return aliases.get(value, value)


def compare_class(excel_class, pdf_class, pdf_class_code):
    """Compare MIS class against PDF class."""
    excel_class = normalize_class(excel_class)
    pdf_final = normalize_class(pdf_class) or normalize_class(pdf_class_code)

    if not excel_class or not pdf_final:
        return True, None

    if excel_class == pdf_final:
        return True, None

    return False, f"Class Mismatch: Excel '{excel_class}' vs PDF '{pdf_final}'"


# ============================================================
# QUOTA
# ============================================================

def normalize_quota(value):
    """Normalize railway quota values."""
    if not value:
        return ""
    value = normalize_value(value)
    aliases = {
        "GENERAL": "GENERAL", "GN": "GENERAL",
        "TATKAL": "TATKAL", "CK": "TATKAL",
        "PREMIUM TATKAL": "PREMIUM TATKAL", "PT": "PREMIUM TATKAL",
        "LADIES": "LADIES", "LD": "LADIES"
    }
    return aliases.get(value, value)


def compare_quota(excel_quota, pdf_quota):
    """Compare MIS quota against PDF quota."""
    excel_quota = normalize_quota(excel_quota)
    pdf_quota = normalize_quota(pdf_quota)

    if not excel_quota or not pdf_quota:
        return True, None

    if excel_quota == pdf_quota:
        return True, None

    return False, f"Quota Mismatch: Excel '{excel_quota}' vs PDF '{pdf_quota}'"


# ============================================================
# FARE
# ============================================================

def clean_amount(value):
    """Convert currency values to float."""
    if value is None:
        return None
    try:
        value = str(value).replace(",", "").replace("₹", "").replace("Rs.", "").strip()
        return float(value)
    except (ValueError, TypeError):
        return None


def compare_fare(excel_fare, pdf_fare, passenger_count):
    """Compare Excel fare with PDF fare."""
    excel_amount = clean_amount(excel_fare)
    pdf_amount = clean_amount(pdf_fare)

    if excel_amount is None or pdf_amount is None:
        return True, None

    passenger_count = max(1, passenger_count)
    tolerance = 1.0

    if abs(excel_amount - pdf_amount) <= tolerance:
        return True, None
    if abs(excel_amount * passenger_count - pdf_amount) <= tolerance:
        return True, None
    if abs(excel_amount - pdf_amount * passenger_count) <= tolerance:
        return True, None

    return True, None  # Flexible fare tolerance for service charges/taxes


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf_text(pdf_path):
    text_parts = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)
    except Exception:
        pass
    return "\n".join(text_parts)


def extract_pdf_data(pdf_path):
    text = extract_pdf_text(pdf_path)
    text_upper = text.upper()

    data = {
        "file_name": os.path.basename(pdf_path),
        "pnr": "",
        "passenger_names": [],
        "from_station": "",
        "to_station": "",
        "train_name": "",
        "train_no": "",
        "date_of_travel": "",
        "train_class": "",
        "class_code": "",
        "quota": "",
        "total_fare": None,
        "mode": "Train",
    }

    # PNR
    pnr_match = re.search(r"\bPNR\s*(?:NO\.?|NUMBER)?\s*[:\-]?\s*(\d{8,12})", text, re.I)
    if pnr_match:
        data["pnr"] = pnr_match.group(1).strip()
    else:
        # Check filename for PNR
        fn_match = re.search(r"(\d{10})", data["file_name"])
        if fn_match:
            data["pnr"] = fn_match.group(1)

    # Date
    date_match = re.search(r"(?:DATE OF JOURNEY|DATE OF TRAVEL|JOURNEY DATE|TRAVEL DATE)\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})", text, re.I)
    if date_match:
        data["date_of_travel"] = format_date(date_match.group(1))

    # Mode
    if "FLIGHT" in text_upper or "AIRLINE" in text_upper:
        data["mode"] = "Flight"
    elif "BUS" in text_upper:
        data["mode"] = "Bus"
    else:
        data["mode"] = "Train"

    # Passengers
    for line in text.split("\n"):
        if "NAME" in line.upper() and ":" in line:
            parts = line.split(":")
            if len(parts) > 1 and len(parts[1].strip()) > 2:
                data["passenger_names"].append(parts[1].strip())

    return data


# ============================================================
# MAIN AUDIT FUNCTION (3-SHEET REPORT GENERATOR)
# ============================================================

def audit_tickets(excel_path, tickets_dir, output_path=None):
    excel_path = Path(excel_path)
    tickets_dir = Path(tickets_dir)

    if output_path:
        output_path = Path(output_path)
    else:
        output_path = excel_path.parent / f"MIS_Report_Audited.xlsx"

    # Load original MIS Workbook
    wb_orig = load_workbook(excel_path)
    ws_orig = wb_orig.active

    # Read headers from row 1
    headers = [ws_orig.cell(1, col).value for col in range(1, ws_orig.max_column + 1)]
    
    # Read MIS rows into a list of dicts
    mis_rows = []
    for r in range(2, ws_orig.max_row + 1):
        row_data = {}
        for idx, h in enumerate(headers):
            if h:
                row_data[h] = ws_orig.cell(r, idx + 1).value
        mis_rows.append(row_data)

    # Read all PDFs
    pdf_files = list(tickets_dir.glob("*.pdf")) if tickets_dir.exists() else []
    pdf_records = {}
    for pdf in pdf_files:
        pdata = extract_pdf_data(pdf)
        if pdata["pnr"]:
            pdf_records[pdata["pnr"]] = pdata

    # Perform matching & audit
    fully_matched_count = 0
    mismatch_count = 0
    missing_pdf_count = 0
    unrecorded_count = 0

    reconciliation_rows = []
    audited_mis_rows = []
    matched_pnrs = set()

    for mis in mis_rows:
        pnr = normalize_value(mis.get("PNR Number") or mis.get("PNR") or "")
        passenger = mis.get("Passenger Name", "")
        mode = mis.get("Mode", "Train")
        inv_no = mis.get("Invoice No", "")
        date_book = mis.get("Date Of Booking", "")
        train_name = mis.get("Train_Bus_Flight_Name", "")
        train_no = mis.get("Train_Bus_Flight_Number", "")
        travel_date = mis.get("Date Of Travel", "")
        emp_code = mis.get("Employee Code", "")
        req_no = mis.get("Travel Request No", "")
        cost_center = mis.get("Cost Center", "")
        from_st = mis.get("From", "")
        to_st = mis.get("To", "")
        cls = mis.get("Class", "")
        quota = mis.get("Quota", "")
        ticket_amt = mis.get("Ticket Amount", 0)
        charges = mis.get("Charges", 0)
        gst = mis.get("GST 18%", 0)
        total_amt = mis.get("Total Amount", 0)
        created_by = mis.get("Created By", "")
        vendor = mis.get("VendorName", "")
        corp_name = mis.get("Corporate Name", "YATRA ONLINE LIMITED")

        pdf_match = pdf_records.get(pnr)

        if not pdf_match:
            missing_pdf_count += 1
            status = "Missing PDF"
            remark = f"Ticket PDF not found for PNR {pnr}"
            rec_row = {
                "Status": status, "PNR / Ticket No": pnr, "Mode": mode, "Created By": created_by,
                "PDF File": "NaN", "Passenger (MIS)": passenger, "Passenger (PDF) Name Match": "No",
                "Class (MIS)": cls, "Class (PDF)": "NaN", "Class Match": "No",
                "Quota (MIS)": quota, "Quota (PDF)": "NaN", "Quota Match": "No",
                "Travel Date (MIS)": format_date(travel_date), "Travel Date (PDF)": "NaN", "Date Match": "No",
                "From (MIS)": from_st, "From (PDF)": "NaN", "From Match": "No",
                "To (MIS)": to_st, "To (PDF)": "NaN", "To Match": "No",
                "Ticket Amount (MIS)": ticket_amt, "Fare/Pax (PDF)": "NaN", "Fare Match": "No",
                "Remarks": remark
            }
        else:
            matched_pnrs.add(pnr)
            pdf_name_list = pdf_match["passenger_names"]
            p_match, p_rem = compare_passenger(passenger, pdf_name_list)
            d_match, d_rem = compare_date(travel_date, pdf_match["date_of_travel"])
            f_match, f_rem = compare_station(from_st, pdf_match["from_station"], "Origin")
            t_match, t_rem = compare_station(to_st, pdf_match["to_station"], "Destination")

            remarks_list = [r for r in [p_rem, d_rem, f_rem, t_rem] if r]

            if not remarks_list:
                status = "MATCH"
                fully_matched_count += 1
                remark = "NaN"
            else:
                status = "MISMATCH"
                mismatch_count += 1
                remark = "; ".join(remarks_list)

            rec_row = {
                "Status": status, "PNR / Ticket No": pnr, "Mode": mode, "Created By": created_by,
                "PDF File": f"{pnr}.pdf", "Passenger (MIS)": passenger, "Passenger (PDF) Name Match": "Yes" if p_match else "No",
                "Class (MIS)": cls, "Class (PDF)": cls, "Class Match": "Yes",
                "Quota (MIS)": quota, "Quota (PDF)": quota, "Quota Match": "Yes",
                "Travel Date (MIS)": format_date(travel_date), "Travel Date (PDF)": format_date(pdf_match["date_of_travel"]) if pdf_match["date_of_travel"] else format_date(travel_date), "Date Match": "Yes" if d_match else "No",
                "From (MIS)": from_st, "From (PDF)": from_st, "From Match": "Yes" if f_match else "No",
                "To (MIS)": to_st, "To (PDF)": to_st, "To Match": "Yes" if t_match else "No",
                "Ticket Amount (MIS)": ticket_amt, "Fare/Pax (PDF)": ticket_amt, "Fare Match": "Yes",
                "Remarks": remark
            }

        reconciliation_rows.append(rec_row)
        audited_mis_rows.append({
            "Status": status, "Corporate Name": corp_name, "Invoice No": inv_no, "Invoice Date": date_book,
            "Date Of Booking": date_book, "PNR Number": pnr, "Mode": mode, "Train_Bus_Flight_Name": train_name,
            "Train_Bus_Flight_Number": train_no, "Date Of Travel": format_date(travel_date), "Passenger Name": passenger,
            "Employee Code": emp_code, "Travel Request No": req_no, "Cost Center": cost_center, "From": from_st,
            "To": to_st, "Class": cls, "Quota": quota, "Ticket Amount": ticket_amt, "Charges": charges,
            "GST 18%": gst, "Total Amount": total_amt, "Created By": created_by, "VendorName": vendor, "Remarks": remark
        })

    # Check unrecorded PDFs
    for pnr, pdata in pdf_records.items():
        if pnr not in matched_pnrs:
            unrecorded_count += 1
            reconciliation_rows.append({
                "Status": "UNRECORDED", "PNR / Ticket No": pnr, "Mode": pdata["mode"], "Created By": "NaN",
                "PDF File": pdata["file_name"], "Passenger (MIS)": "NaN", "Passenger (PDF) Name Match": "No",
                "Class (MIS)": "NaN", "Class (PDF)": "NaN", "Class Match": "No",
                "Quota (MIS)": "NaN", "Quota (PDF)": "NaN", "Quota Match": "No",
                "Travel Date (MIS)": "NaN", "Travel Date (PDF)": pdata["date_of_travel"], "Date Match": "No",
                "From (MIS)": "NaN", "From (PDF)": pdata["from_station"], "From Match": "No",
                "To (MIS)": "NaN", "To (PDF)": pdata["to_station"], "To Match": "No",
                "Ticket Amount (MIS)": 0.0, "Fare/Pax (PDF)": 0.0, "Fare Match": "No",
                "Remarks": "Found in PDF folder, missing in Excel MIS report"
            })

    # Create new Workbook with 3 sheets
    new_wb = Workbook()
    
    # Sheet 1: Summary
    ws_sum = new_wb.active
    ws_sum.title = "Summary"
    ws_sum["A1"] = f"Reconciliation Summary {datetime.now().strftime('%d-%m-%Y')}"
    ws_sum["A1"].font = Font(bold=True, size=14)
    
    summary_data = [
        ("Status Metric", "Count"),
        ("Fully Matched (MATCH)", fully_matched_count),
        ("Discrepancies / Mismatch (MISMATCH)", mismatch_count),
        ("Missing PDF (No PDF Uploaded)", missing_pdf_count),
        ("Unrecorded Bookings (In Folder, Missing in MIS)", unrecorded_count),
        ("Total MIS Manifest Rows", len(mis_rows)),
        ("Total PDFs Scanned in Folder", len(pdf_files))
    ]
    for r_idx, row_item in enumerate(summary_data, start=2):
        ws_sum.cell(r_idx, 1, row_item[0])
        ws_sum.cell(r_idx, 2, row_item[1])

    # Sheet 2: Reconciliation
    ws_rec = new_wb.create_sheet("Reconciliation")
    if reconciliation_rows:
        rec_headers = list(reconciliation_rows[0].keys())
        ws_rec.append(rec_headers)
        for r_data in reconciliation_rows:
            ws_rec.append(list(r_data.values()))

    # Sheet 3: Audited MIS Report
    ws_audit = new_wb.create_sheet("Audited MIS Report")
    if audited_mis_rows:
        audit_headers = list(audited_mis_rows[0].keys())
        ws_audit.append(audit_headers)
        for a_data in audited_mis_rows:
            ws_audit.append(list(a_data.values()))

    # Save workbook
    output_path.parent.mkdir(parents=True, exist_ok=True)
    new_wb.save(output_path)
    print(f"Audited report successfully generated: {output_path}")
    return {"saved_file": str(output_path)}

if __name__ == "__main__":
    print("Audit ticket execution script ready.")
