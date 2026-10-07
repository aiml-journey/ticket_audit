
import os
import re
import glob
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from pypdf import PdfReader


# ============================================================
# EXACT DESIRED EXCEL STRUCTURE
# ============================================================

RECONCILIATION_HEADERS = [
    "Status",
    "PNR / Ticket No",
    "Mode",
    "Created By",
    "PDF File",
    "Passenger (MIS)",
    "Passenger (PDF)",
    "Name Match",
    "Class (MIS)",
    "Class (PDF)",
    "Class Match",
    "Quota (MIS)",
    "Quota (PDF)",
    "Quota Match",
    "Travel Date (MIS)",
    "Travel Date (PDF)",
    "Date Match",
    "From (MIS)",
    "From (PDF)",
    "From Match",
    "To (MIS)",
    "To (PDF)",
    "To Match",
    "Ticket Amount (MIS)",
    "Fare/Pax (PDF)",
    "Fare Match",
    "Remarks"
]


AUDITED_HEADERS = [
    "Status",
    "Corporate Name",
    "Invoice No",
    "Invoice Date",
    "Date Of Booking",
    "PNR Number",
    "Mode",
    "Train_Bus_Flight_Name",
    "Train_Bus_Flight_Number",
    "Date Of Travel",
    "Passenger Name",
    "Employee Code",
    "Travel Request No",
    "Cost Center",
    "From",
    "To",
    "Class",
    "Quota",
    "Ticket Amount",
    "Charges",
    "GST 18%",
    "Total Amount",
    "Created By",
    "VendorName",
    "Remarks"
]


# ============================================================
# EXCEL STYLES
# ============================================================

HEADER_FILL = PatternFill(
    "solid",
    fgColor="305496"
)

HEADER_FONT = Font(
    bold=True,
    color="FFFFFF"
)

MATCH_FILL = PatternFill(
    "solid",
    fgColor="C6EFCE"
)

MISMATCH_FILL = PatternFill(
    "solid",
    fgColor="FFC7CE"
)

WARNING_FILL = PatternFill(
    "solid",
    fgColor="FFEB9C"
)

INFO_FILL = PatternFill(
    "solid",
    fgColor="D9EAF7"
)

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean(value):
    """Convert a value into a clean string."""
    if value is None:
        return ""

    return str(value).strip()


def normalize_text(value):
    """
    Normalize text for comparison.

    Example:
    '  Rahul Kumar  ' -> 'RAHUL KUMAR'
    'Rahul-Kumar'     -> 'RAHUL KUMAR'
    """
    value = clean(value).upper()

    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = value.replace("-", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_name(value):
    """
    Normalize passenger names.

    Allows comparison despite:
    - different spacing
    - punctuation
    - comma placement
    - upper/lower case
    """
    value = normalize_text(value)

    value = re.sub(r"[^A-Z0-9 ]", "", value)

    words = value.split()

    return " ".join(words)


def normalize_number(value):
    """
    Normalize numerical values such as:

    1250
    1250.00
    ₹1,250.00
    INR 1250
    """
    if value is None:
        return None

    value = clean(value)

    value = value.replace(",", "")
    value = value.replace("₹", "")
    value = value.replace("INR", "")
    value = value.replace("Rs.", "")
    value = value.replace("Rs", "")

    value = value.strip()

    match = re.search(r"-?\d+(?:\.\d+)?", value)

    if not match:
        return None

    try:
        return round(float(match.group()), 2)
    except ValueError:
        return None


def numbers_match(value1, value2):
    """Compare two numerical values."""
    n1 = normalize_number(value1)
    n2 = normalize_number(value2)

    if n1 is None or n2 is None:
        return False

    return abs(n1 - n2) < 0.01


def normalize_date(value):
    """
    Normalize common date formats to DD-MM-YYYY.
    """

    if value is None or value == "":
        return ""

    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")

    value = clean(value)

    # Remove time if present
    value = value.split(" ")[0]

    formats = [
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d.%m.%Y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d %b %Y",
        "%d %B %Y",
        "%d-%b-%Y",
        "%d-%B-%Y"
    ]

    for fmt in formats:
        try:
            dt = datetime.strptime(value, fmt)
            return dt.strftime("%d-%m-%Y")
        except ValueError:
            pass

    return normalize_text(value)


def values_match(value1, value2):
    """General text comparison."""
    return normalize_text(value1) == normalize_text(value2)


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(pdf_path):
    """
    Extract all text from a PDF.
    """

    try:
        reader = PdfReader(pdf_path)

        pages = []

        for page in reader.pages:
            try:
                text = page.extract_text()

                if text:
                    pages.append(text)

            except Exception:
                continue

        return "\n".join(pages)

    except Exception:
        return ""


# ============================================================
# FIELD EXTRACTION FROM PDF
# ============================================================

def extract_after_label(text, labels):
    """
    Extract text appearing after labels such as:

    Passenger Name:
    Passenger:
    Class:
    Quota:
    """

    if not text:
        return ""

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    for i, line in enumerate(lines):

        normalized_line = normalize_text(line)

        for label in labels:

            normalized_label = normalize_text(label)

            if normalized_line.startswith(normalized_label):

                # Case 1:
                # Passenger Name: Rahul Kumar
                remaining = line[len(label):].strip()

                remaining = remaining.lstrip(":|- ")

                if remaining:
                    return remaining

                # Case 2:
                # Passenger Name
                # Rahul Kumar
                if i + 1 < len(lines):
                    return lines[i + 1]

    return ""


def extract_pnr(text):
    """
    Extract PNR / ticket number.

    Handles common formats such as:
    PNR: 1234567890
    PNR Number: 1234567890
    PNR 1234567890
    """

    patterns = [
        r"\bPNR\s*(?:NUMBER|NO|NUM|#)?\s*[:\-]?\s*([A-Z0-9]{5,20})",
        r"\bBOOKING\s*(?:ID|NO|NUMBER)?\s*[:\-]?\s*([A-Z0-9]{5,20})",
        r"\bTICKET\s*(?:NO|NUMBER|ID)?\s*[:\-]?\s*([A-Z0-9]{5,20})"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return ""


def extract_passenger(text):
    """
    Extract passenger name from PDF.
    """

    labels = [
        "Passenger Name",
        "Passenger",
        "Passenger Name:",
        "Name"
    ]

    value = extract_after_label(
        text,
        labels
    )

    if value:
        return value

    # Common railway format:
    # 1. NAME
    # 1 NAME
    patterns = [
        r"\b1[\.\)]\s*([A-Za-z][A-Za-z .]+)",
        r"\b1\s+([A-Za-z][A-Za-z .]+)"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return ""


def extract_class(text):
    """
    Extract travel class.
    """

    labels = [
        "Class",
        "Travel Class",
        "Class Type",
        "Booking Class"
    ]

    value = extract_after_label(
        text,
        labels
    )

    if value:
        return value

    # Common railway classes
    railway_classes = [
        "1A",
        "2A",
        "3A",
        "SL",
        "CC",
        "EC",
        "2S",
        "3E",
        "FC"
    ]

    upper_text = normalize_text(text)

    for travel_class in railway_classes:

        if re.search(
            r"\b" + re.escape(travel_class) + r"\b",
            upper_text
        ):
            return travel_class

    return ""


def extract_quota(text):
    """
    Extract railway quota.
    """

    labels = [
        "Quota",
        "Booking Quota",
        "Reservation Quota"
    ]

    value = extract_after_label(
        text,
        labels
    )

    if value:
        return value

    # Common railway quotas
    quotas = [
        "GENERAL",
        "GN",
        "LADIES",
        "LD",
        "TATKAL",
        "CK",
        "PREMIUM TATKAL",
        "PT",
        "SENIOR CITIZEN",
        "LOWER BERTH",
        "DUTY PASS"
    ]

    upper_text = normalize_text(text)

    for quota in quotas:

        if re.search(
            r"\b" + re.escape(quota) + r"\b",
            upper_text
        ):
            return quota

    return ""


def extract_travel_date(text):
    """
    Extract travel date from PDF.
    """

    labels = [
        "Date Of Travel",
        "Date of Travel",
        "Travel Date",
        "Journey Date",
        "Date of Journey",
        "Departure Date"
    ]

    value = extract_after_label(
        text,
        labels
    )

    if value:
        return value

    # Generic date search
    patterns = [
        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
        r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}\b"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text
        )

        if match:
            return match.group()

    return ""


def extract_station(text, labels):
    """
    Extract From / To station.
    """

    value = extract_after_label(
        text,
        labels
    )

    return value


def extract_fare(text):
    """
    Extract ticket fare / fare per passenger.

    Tries common ticket wording.
    """

    patterns = [
        r"(?:Fare\s*(?:/Pax|Per Passenger|Per Pax)?)\s*[:\-]?\s*(?:₹|INR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)",
        r"(?:Ticket Amount|Ticket Fare|Total Fare|Basic Fare)\s*[:\-]?\s*(?:₹|INR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)",
        r"(?:Amount Payable|Amount Paid|Total Amount)\s*[:\-]?\s*(?:₹|INR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)"
    ]

    for pattern in patterns:

        matches = re.findall(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        if matches:
            return matches[-1]

    # Generic INR/₹ amount as last resort
    amounts = re.findall(
        r"(?:₹|INR|Rs\.?)\s*([\d,]+(?:\.\d+)?)",
        text,
        flags=re.IGNORECASE
    )

    if amounts:
        return amounts[-1]

    return ""


def extract_pdf_fields(pdf_path):
    """
    Extract all required fields from one PDF.
    """

    text = extract_pdf_text(pdf_path)

    fields = {
        "pnr": extract_pnr(text),
        "passenger": extract_passenger(text),
        "class": extract_class(text),
        "quota": extract_quota(text),
        "travel_date": extract_travel_date(text),
        "from": extract_station(
            text,
            [
                "From",
                "Origin",
                "Boarding From",
                "Source"
            ]
        ),
        "to": extract_station(
            text,
            [
                "To",
                "Destination",
                "Boarding To",
                "Drop",
                "Arrival"
            ]
        ),
        "fare": extract_fare(text)
    }

    return fields


# ============================================================
# MIS HELPERS
# ============================================================

def get_header_map(ws):
    """
    Create:
        column name -> column number
    """

    header_map = {}

    for cell in ws[1]:

        if cell.value is not None:

            header_map[
                normalize_text(cell.value)
            ] = cell.column

    return header_map


def find_column(header_map, possible_names):
    """
    Find a column using multiple possible names.
    """

    for name in possible_names:

        key = normalize_text(name)

        if key in header_map:
            return header_map[key]

    return None


def get_cell_value(ws, row_number, header_map, possible_names):
    """
    Safely retrieve a value from MIS.
    """

    column = find_column(
        header_map,
        possible_names
    )

    if column is None:
        return ""

    return ws.cell(
        row=row_number,
        column=column
    ).value


def detect_mode(row_data):
    """
    Determine whether ticket is Railway.

    Quota comparison is performed ONLY when this returns True.
    """

    mode = normalize_text(
        row_data.get("Mode", "")
    )

    railway_words = [
        "RAIL",
        "RAILWAY",
        "TRAIN",
        "IRCTC"
    ]

    for word in railway_words:

        if word in mode:
            return True

    return False


# ============================================================
# PDF INDEX
# ============================================================

def build_pdf_index(tickets_dir):
    """
    Scan all PDFs and index them by PNR.

    Returns:
        {
            PNR: {
                "file": filename,
                "fields": {...}
            }
        }

    Also returns PDFs whose PNR could not be extracted.
    """

    pdf_index = {}
    unindexed_pdfs = []

    pdf_files = glob.glob(
        os.path.join(
            tickets_dir,
            "**",
            "*.pdf"
        ),
        recursive=True
    )

    for pdf_path in pdf_files:

        fields = extract_pdf_fields(
            pdf_path
        )

        pnr = normalize_text(
            fields.get("pnr", "")
        )

        filename = os.path.basename(
            pdf_path
        )

        data = {
            "file": filename,
            "path": pdf_path,
            "fields": fields
        }

        if pnr:
            pdf_index[pnr] = data
        else:
            unindexed_pdfs.append(data)

    return pdf_index, unindexed_pdfs, pdf_files


# ============================================================
# MATCHING LOGIC
# ============================================================

def compare_passenger(mis_value, pdf_value):
    if not mis_value or not pdf_value:
        return False

    return (
        normalize_name(mis_value)
        == normalize_name(pdf_value)
    )


def compare_class(mis_value, pdf_value):
    if not mis_value or not pdf_value:
        return False

    return (
        normalize_text(mis_value)
        == normalize_text(pdf_value)
    )


def compare_quota(mis_value, pdf_value):
    if not mis_value or not pdf_value:
        return False

    return (
        normalize_text(mis_value)
        == normalize_text(pdf_value)
    )


def compare_fare(mis_value, pdf_value):
    return numbers_match(
        mis_value,
        pdf_value
    )


def compare_date(mis_value, pdf_value):
    if not mis_value or not pdf_value:
        return False

    return (
        normalize_date(mis_value)
        == normalize_date(pdf_value)
    )


def compare_location(mis_value, pdf_value):
    if not mis_value or not pdf_value:
        return False

    return (
        normalize_text(mis_value)
        == normalize_text(pdf_value)
    )


# ============================================================
# REMARKS
# ============================================================

def create_remarks(
    name_match,
    class_match,
    quota_match,
    fare_match,
    is_railway
):

    problems = []

    if not name_match:
        problems.append(
            "Passenger Name Mismatch"
        )

    if not class_match:
        problems.append(
            "Class Mismatch"
        )

    if is_railway and not quota_match:
        problems.append(
            "Quota Mismatch"
        )

    if not fare_match:
        problems.append(
            "Fare Mismatch"
        )

    if not problems:
        return "All required fields matched"

    return "; ".join(problems)


# ============================================================
# WORKBOOK FORMATTING
# ============================================================

def style_header(ws):
    for cell in ws[1]:

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )
        cell.border = THIN_BORDER

    ws.row_dimensions[1].height = 35


def apply_status_style(cell):
    status = normalize_text(
        cell.value
    )

    if status == "MATCH":

        cell.fill = MATCH_FILL
        cell.font = Font(
            bold=True,
            color="006100"
        )

    elif (
        status == "MISMATCH"
        or status == "MISSING PDF"
        or status == "UNRECORDED"
    ):

        cell.fill = MISMATCH_FILL
        cell.font = Font(
            bold=True,
            color="9C0006"
        )

    else:

        cell.fill = WARNING_FILL


def auto_width(ws, maximum=35):

    for column_cells in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            try:
                value_length = len(
                    str(cell.value)
                )

                if value_length > max_length:
                    max_length = value_length

            except Exception:
                pass

        ws.column_dimensions[
            column_letter
        ].width = min(
            max(max_length + 2, 12),
            maximum
        )


def style_data_sheet(ws):

    style_header(ws)

    for row in ws.iter_rows(
        min_row=2
    ):

        for cell in row:

            cell.border = THIN_BORDER
            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        # Status is always column A
        apply_status_style(
            row[0]
        )

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    auto_width(ws)


# ============================================================
# MAIN AUDIT FUNCTION
# ============================================================

def audit_tickets(
    excel_path,
    tickets_dir,
    output_path
):
    """
    Main function used by app.py.

    Parameters:
        excel_path:
            Uploaded MIS Excel file.

        tickets_dir:
            Folder containing PDF tickets.

        output_path:
            Path where audited Excel is saved.

    Returns:
        output_path
    """

    # --------------------------------------------------------
    # LOAD MIS
    # --------------------------------------------------------

    source_wb = load_workbook(
        excel_path,
        data_only=False
    )

    source_ws = source_wb.active

    header_map = get_header_map(
        source_ws
    )

    # --------------------------------------------------------
    # BUILD PDF INDEX
    # --------------------------------------------------------

    pdf_index, unindexed_pdfs, pdf_files = build_pdf_index(
        tickets_dir
    )

    # --------------------------------------------------------
    # CREATE OUTPUT WORKBOOK
    # --------------------------------------------------------

    wb = Workbook()

    # Remove default sheet
    default_ws = wb.active
    wb.remove(default_ws)

    summary_ws = wb.create_sheet(
        "Summary"
    )

    reconciliation_ws = wb.create_sheet(
        "Reconciliation"
    )

    audited_ws = wb.create_sheet(
        "Audited MIS Report"
    )

    # --------------------------------------------------------
    # EXACT RECONCILIATION HEADERS
    # --------------------------------------------------------

    reconciliation_ws.append(
        RECONCILIATION_HEADERS
    )

    # --------------------------------------------------------
    # EXACT AUDITED MIS HEADERS
    # --------------------------------------------------------

    audited_ws.append(
        AUDITED_HEADERS
    )

    # --------------------------------------------------------
    # COUNTERS
    # --------------------------------------------------------

    matched_count = 0
    mismatch_count = 0
    missing_pdf_count = 0
    unrecorded_count = 0
    total_mis_rows = 0

    # Keep track of PDFs matched to MIS
    matched_pdf_files = set()

    # --------------------------------------------------------
    # PROCESS EVERY MIS ROW
    # --------------------------------------------------------

    for row_number in range(
        2,
        source_ws.max_row + 1
    ):

        # Ignore completely empty rows
        row_values = [
            source_ws.cell(
                row=row_number,
                column=col
            ).value
            for col in range(
                1,
                source_ws.max_column + 1
            )
        ]

        if not any(
            value is not None
            and str(value).strip() != ""
            for value in row_values
        ):
            continue

        total_mis_rows += 1

        # ----------------------------------------------------
        # READ MIS VALUES
        # ----------------------------------------------------

        corporate_name = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Corporate Name",
                "CorporateName"
            ]
        )

        invoice_no = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Invoice No",
                "Invoice Number"
            ]
        )

        invoice_date = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Invoice Date"
            ]
        )

        date_of_booking = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Date Of Booking",
                "Date of Booking"
            ]
        )

        pnr = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "PNR Number",
                "PNR",
                "Ticket No",
                "Ticket Number"
            ]
        )

        mode = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Mode"
            ]
        )

        train_name = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Train_Bus_Flight_Name",
                "Train/Bus/Flight Name",
                "Train Name",
                "Flight Name",
                "Bus Name"
            ]
        )

        train_number = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Train_Bus_Flight_Number",
                "Train/Bus/Flight Number",
                "Train Number",
                "Flight Number",
                "Bus Number"
            ]
        )

        travel_date = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Date Of Travel",
                "Travel Date",
                "Date of Travel"
            ]
        )

        passenger = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Passenger Name",
                "Passenger",
                "Name"
            ]
        )

        employee_code = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Employee Code",
                "EmployeeCode"
            ]
        )

        travel_request_no = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Travel Request No",
                "Travel Request Number"
            ]
        )

        cost_center = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Cost Center",
                "CostCentre"
            ]
        )

        from_location = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "From",
                "Origin"
            ]
        )

        to_location = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "To",
                "Destination"
            ]
        )

        travel_class = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Class",
                "Travel Class"
            ]
        )

        quota = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Quota"
            ]
        )

        ticket_amount = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Ticket Amount",
                "Ticket Fare",
                "Fare",
                "Amount"
            ]
        )

        charges = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Charges"
            ]
        )

        gst = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "GST 18%",
                "GST",
                "GST Amount"
            ]
        )

        total_amount = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Total Amount"
            ]
        )

        created_by = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "Created By",
                "CreatedBy"
            ]
        )

        vendor_name = get_cell_value(
            source_ws,
            row_number,
            header_map,
            [
                "VendorName",
                "Vendor Name"
            ]
        )

        # ----------------------------------------------------
        # NORMALIZE PNR
        # ----------------------------------------------------

        normalized_pnr = normalize_text(
            pnr
        )

        # ----------------------------------------------------
        # DETERMINE RAILWAY
        # ----------------------------------------------------

        mis_row_data = {
            "Mode": mode
        }

        is_railway = detect_mode(
            mis_row_data
        )

        # ----------------------------------------------------
        # FIND PDF
        # ----------------------------------------------------

        pdf_data = pdf_index.get(
            normalized_pnr
        )

        # ----------------------------------------------------
        # CASE 1: PDF NOT FOUND
        # ----------------------------------------------------

        if pdf_data is None:

            status = "MISSING PDF"

            missing_pdf_count += 1

            remarks = (
                "No PDF found for PNR / Ticket No"
            )

            passenger_pdf = ""
            class_pdf = ""
            quota_pdf = ""
            fare_pdf = ""
            travel_date_pdf = ""
            from_pdf = ""
            to_pdf = ""

            name_match = "NOT CHECKED"
            class_match = "NOT CHECKED"
            quota_match = (
                "N/A"
                if not is_railway
                else "NOT CHECKED"
            )
            date_match = "NOT CHECKED"
            from_match = "NOT CHECKED"
            to_match = "NOT CHECKED"
            fare_match = "NOT CHECKED"

        # ----------------------------------------------------
        # CASE 2: PDF FOUND
        # ----------------------------------------------------

        else:

            matched_pdf_files.add(
                pdf_data["path"]
            )

            pdf_fields = pdf_data[
                "fields"
            ]

            passenger_pdf = pdf_fields.get(
                "passenger",
                ""
            )

            class_pdf = pdf_fields.get(
                "class",
                ""
            )

            quota_pdf = pdf_fields.get(
                "quota",
                ""
            )

            fare_pdf = pdf_fields.get(
                "fare",
                ""
            )

            travel_date_pdf = pdf_fields.get(
                "travel_date",
                ""
            )

            from_pdf = pdf_fields.get(
                "from",
                ""
            )

            to_pdf = pdf_fields.get(
                "to",
                ""
            )

            # ------------------------------------------------
            # REQUIRED STATUS CHECKS
            # ------------------------------------------------

            name_ok = compare_passenger(
                passenger,
                passenger_pdf
            )

            class_ok = compare_class(
                travel_class,
                class_pdf
            )

            fare_ok = compare_fare(
                ticket_amount,
                fare_pdf
            )

            # Quota is checked ONLY for railway
            if is_railway:

                quota_ok = compare_quota(
                    quota,
                    quota_pdf
                )

            else:

                quota_ok = True

            # ------------------------------------------------
            # INFORMATIONAL CHECKS
            # These DO NOT affect STATUS.
            # ------------------------------------------------

            date_ok = compare_date(
                travel_date,
                travel_date_pdf
            )

            from_ok = compare_location(
                from_location,
                from_pdf
            )

            to_ok = compare_location(
                to_location,
                to_pdf
            )

            # ------------------------------------------------
            # DISPLAY VALUES
            # ------------------------------------------------

            name_match = (
                "MATCH"
                if name_ok
                else "MISMATCH"
            )

            class_match = (
                "MATCH"
                if class_ok
                else "MISMATCH"
            )

            if is_railway:

                quota_match = (
                    "MATCH"
                    if quota_ok
                    else "MISMATCH"
                )

            else:

                quota_match = "N/A"

            date_match = (
                "MATCH"
                if date_ok
                else "MISMATCH"
            )

            from_match = (
                "MATCH"
                if from_ok
                else "MISMATCH"
            )

            to_match = (
                "MATCH"
                if to_ok
                else "MISMATCH"
            )

            fare_match = (
                "MATCH"
                if fare_ok
                else "MISMATCH"
            )

            # ------------------------------------------------
            # FINAL STATUS
            #
            # IMPORTANT:
            # Only these affect final status:
            # Passenger
            # Class
            # Quota for Railway
            # Fare
            #
            # Date / From / To are NOT included.
            # ------------------------------------------------

            if (
                name_ok
                and class_ok
                and quota_ok
                and fare_ok
            ):

                status = "MATCH"

                matched_count += 1

            else:

                status = "MISMATCH"

                mismatch_count += 1

            remarks = create_remarks(
                name_ok,
                class_ok,
                quota_ok,
                fare_ok,
                is_railway
            )

        # ----------------------------------------------------
        # RECONCILIATION ROW
        # ----------------------------------------------------

        reconciliation_row = [
            status,
            pnr,
            mode,
            created_by,
            (
                pdf_data["file"]
                if pdf_data
                else ""
            ),
            passenger,
            passenger_pdf,
            name_match,
            travel_class,
            class_pdf,
            class_match,
            quota,
            quota_pdf,
            quota_match,
            travel_date,
            travel_date_pdf,
            date_match,
            from_location,
            from_pdf,
            from_match,
            to_location,
            to_pdf,
            to_match,
            ticket_amount,
            fare_pdf,
            fare_match,
            remarks
        ]

        reconciliation_ws.append(
            reconciliation_row
        )

        # ----------------------------------------------------
        # AUDITED MIS REPORT ROW
        #
        # IMPORTANT:
        # This preserves the MIS structure and adds the
        # audit status + remarks.
        # ----------------------------------------------------

        audited_row = [
            status,
            corporate_name,
            invoice_no,
            invoice_date,
            date_of_booking,
            pnr,
            mode,
            train_name,
            train_number,
            travel_date,
            passenger,
            employee_code,
            travel_request_no,
            cost_center,
            from_location,
            to_location,
            travel_class,
            quota,
            ticket_amount,
            charges,
            gst,
            total_amount,
            created_by,
            vendor_name,
            remarks
        ]

        audited_ws.append(
            audited_row
        )

    # ========================================================
    # UNRECORDED PDFs
    #
    # PDFs found in folder but whose PNR does not exist
    # in MIS.
    # ========================================================

    unrecorded_pdf_paths = []

    for pdf_path in pdf_files:

        if pdf_path not in matched_pdf_files:

            # If the PDF has a PNR, check whether that PNR
            # existed in MIS.
            filename = os.path.basename(
                pdf_path
            )

            pdf_data = None

            for data in unindexed_pdfs:

                if data["path"] == pdf_path:

                    pdf_data = data
                    break

            if pdf_data is None:

                # Find indexed PDF by path
                for data in pdf_index.values():

                    if data["path"] == pdf_path:

                        pdf_data = data
                        break

            pnr_pdf = ""

            if pdf_data:

                pnr_pdf = pdf_data[
                    "fields"
                ].get(
                    "pnr",
                    ""
                )

            # If PNR is absent from MIS, it is unrecorded.
            if normalize_text(pnr_pdf) not in [
                normalize_text(
                    get_cell_value(
                        source_ws,
                        r,
                        header_map,
                        [
                            "PNR Number",
                            "PNR",
                            "Ticket No",
                            "Ticket Number"
                        ]
                    )
                )
                for r in range(
                    2,
                    source_ws.max_row + 1
                )
            ]:

                unrecorded_count += 1

                unrecorded_pdf_paths.append(
                    pdf_path
                )

                # Add an UNRECORDED row to Reconciliation
                reconciliation_ws.append([
                    "UNRECORDED",
                    pnr_pdf,
                    "",
                    "",
                    filename,
                    "",
                    (
                        pdf_data["fields"].get(
                            "passenger",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "",
                    (
                        pdf_data["fields"].get(
                            "class",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "",
                    (
                        pdf_data["fields"].get(
                            "quota",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "",
                    (
                        pdf_data["fields"].get(
                            "travel_date",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "",
                    (
                        pdf_data["fields"].get(
                            "from",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "",
                    (
                        pdf_data["fields"].get(
                            "to",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "",
                    (
                        pdf_data["fields"].get(
                            "fare",
                            ""
                        )
                        if pdf_data
                        else ""
                    ),
                    "",
                    "PDF found in folder but PNR is not present in MIS"
                ])

    # ========================================================
    # SUMMARY
    # ========================================================

    summary_ws["A1"] = "Reconciliation Summary"

    summary_ws["B1"] = datetime.now().strftime(
        "%d-%m-%Y"
    )

    summary_ws["A2"] = "Status Metric"
    summary_ws["B2"] = "Count"

    summary_ws["A3"] = "Fully Matched (MATCH)"
    summary_ws["B3"] = matched_count

    summary_ws["A4"] = "Discrepancies / Mismatch (MISMATCH)"
    summary_ws["B4"] = mismatch_count

    summary_ws["A5"] = "Missing PDF (No PDF Uploaded)"
    summary_ws["B5"] = missing_pdf_count

    summary_ws["A6"] = (
        "Unrecorded Bookings "
        "(In Folder, Missing in MIS)"
    )
    summary_ws["B6"] = unrecorded_count

    summary_ws["A7"] = "Total MIS Manifest Rows"
    summary_ws["B7"] = total_mis_rows

    summary_ws["A8"] = "Total PDFs Scanned in Folder"
    summary_ws["B8"] = len(pdf_files)

    # Summary formatting
    summary_ws["A1"].font = Font(
        bold=True,
        size=14
    )

    summary_ws["B1"].font = Font(
        bold=True
    )

    for cell in summary_ws[2]:

        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center"
        )
        cell.border = THIN_BORDER

    for row in summary_ws.iter_rows(
        min_row=3,
        max_row=8,
        min_col=1,
        max_col=2
    ):

        for cell in row:

            cell.border = THIN_BORDER

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

    summary_ws["B3"].fill = MATCH_FILL
    summary_ws["B3"].font = Font(
        bold=True,
        color="006100"
    )

    summary_ws["B4"].fill = MISMATCH_FILL
    summary_ws["B4"].font = Font(
        bold=True,
        color="9C0006"
    )

    summary_ws["B5"].fill = MISMATCH_FILL
    summary_ws["B6"].fill = MISMATCH_FILL

    summary_ws.column_dimensions["A"].width = 48
    summary_ws.column_dimensions["B"].width = 20

    # ========================================================
    # STYLE RECONCILIATION
    # ========================================================

    style_data_sheet(
        reconciliation_ws
    )

    # ========================================================
    # STYLE AUDITED MIS REPORT
    # ========================================================

    style_data_sheet(
        audited_ws
    )

    # ========================================================
    # ADD COLOUR TO INDIVIDUAL MATCH COLUMNS
    # ========================================================

    # Reconciliation:
    # H  = Name Match
    # K  = Class Match
    # N  = Quota Match
    # Q  = Date Match
    # T  = From Match
    # W  = To Match
    # Z  = Fare Match

    match_columns = [
        8,
        11,
        14,
        17,
        20,
        23,
        26
    ]

    for row in reconciliation_ws.iter_rows(
        min_row=2
    ):

        for column_number in match_columns:

            cell = row[
                column_number - 1
            ]

            value = normalize_text(
                cell.value
            )

            if value == "MATCH":

                cell.fill = MATCH_FILL
                cell.font = Font(
                    bold=True,
                    color="006100"
                )

            elif value == "MISMATCH":

                cell.fill = MISMATCH_FILL
                cell.font = Font(
                    bold=True,
                    color="9C0006"
                )

            elif value in [
                "N/A",
                "NOT CHECKED"
            ]:

                cell.fill = WARNING_FILL

    # ========================================================
    # FREEZE PANES
    # ========================================================

    summary_ws.freeze_panes = None

    reconciliation_ws.freeze_panes = "B2"

    audited_ws.freeze_panes = "B2"

    # ========================================================
    # SAVE
    # ========================================================

    wb.save(
        output_path
    )

    return output_path


# ============================================================
# COMMAND-LINE SUPPORT
# ============================================================

if __name__ == "__main__":

    import argparse

    parser = argparse.ArgumentParser(
        description="Daily Ticket Auditing"
    )

    parser.add_argument(
        "--excel",
        required=True,
        help="Path to MIS Excel file"
    )

    parser.add_argument(
        "--tickets",
        required=True,
        help="Folder containing PDF tickets"
    )

    parser.add_argument(
        "--output",
        required=True,
        help="Path for output Excel report"
    )

    args = parser.parse_args()

    result = audit_tickets(
        excel_path=args.excel,
        tickets_dir=args.tickets,
        output_path=args.output
    )

    print(
        f"Audit completed successfully: {result}"
    )

