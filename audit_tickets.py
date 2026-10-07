import os
import re
import glob
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


# ============================================================
# EXACT REQUIRED EXCEL STRUCTURE
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
# STYLES
# ============================================================

HEADER_FILL = PatternFill(
    "solid",
    fgColor="305496"
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

HEADER_FONT = Font(
    bold=True,
    color="FFFFFF"
)

THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)


# ============================================================
# BASIC HELPERS
# ============================================================

def clean(value):
    if value is None:
        return ""

    return str(value).strip()


def normalize_text(value):
    value = clean(value)

    value = value.upper()
    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = value.replace("-", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def normalize_name(value):
    value = normalize_text(value)

    value = re.sub(
        r"[^A-Z0-9 ]",
        "",
        value
    )

    return " ".join(
        value.split()
    )


def normalize_number(value):
    if value is None:
        return None

    value = clean(value)

    value = value.replace(",", "")
    value = value.replace("₹", "")
    value = value.replace("INR", "")
    value = value.replace("Rs.", "")
    value = value.replace("Rs", "")

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        value
    )

    if not match:
        return None

    try:
        return round(
            float(match.group()),
            2
        )
    except ValueError:
        return None


def numbers_match(value1, value2):
    n1 = normalize_number(value1)
    n2 = normalize_number(value2)

    if n1 is None or n2 is None:
        return False

    return abs(n1 - n2) < 0.01


def normalize_date(value):

    if value is None or value == "":
        return ""

    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")

    value = clean(value)

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
            dt = datetime.strptime(
                value,
                fmt
            )

            return dt.strftime(
                "%d-%m-%Y"
            )

        except ValueError:
            pass

    return normalize_text(value)


# ============================================================
# EXCEL HELPERS
# ============================================================

def get_header_map(ws):

    result = {}

    for cell in ws[1]:

        if cell.value is not None:

            result[
                normalize_text(cell.value)
            ] = cell.column

    return result


def find_column(
    header_map,
    possible_names
):

    for name in possible_names:

        key = normalize_text(name)

        if key in header_map:
            return header_map[key]

    return None


def get_value(
    ws,
    row,
    header_map,
    possible_names
):

    column = find_column(
        header_map,
        possible_names
    )

    if column is None:
        return ""

    return ws.cell(
        row=row,
        column=column
    ).value


# ============================================================
# PDF TEXT
# ============================================================

def extract_pdf_text(pdf_path):

    try:

        reader = PdfReader(
            pdf_path
        )

        all_text = []

        for page in reader.pages:

            try:

                text = page.extract_text()

                if text:
                    all_text.append(text)

            except Exception:
                pass

        return "\n".join(all_text)

    except Exception:

        return ""


# ============================================================
# GENERIC PDF FIELD EXTRACTION
# ============================================================

def extract_after_label(
    text,
    labels
):

    if not text:
        return ""

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    for i, line in enumerate(lines):

        upper_line = normalize_text(
            line
        )

        for label in labels:

            upper_label = normalize_text(
                label
            )

            if upper_line.startswith(
                upper_label
            ):

                remaining = line[
                    len(label):
                ].strip()

                remaining = remaining.lstrip(
                    ":|- "
                )

                if remaining:
                    return remaining

                if i + 1 < len(lines):
                    return lines[i + 1]

    return ""


def extract_pnr(text):

    patterns = [
        r"\bPNR\s*(?:NUMBER|NO|NUM|#)?\s*[:\-]?\s*([A-Z0-9]{5,20})",
        r"\bBOOKING\s*(?:ID|NO|NUMBER)?\s*[:\-]?\s*([A-Z0-9]{5,20})",
        r"\bTICKET\s*(?:NO|NUMBER|ID)?\s*[:\-]?\s*([A-Z0-9]{5,20})"
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if match:
            return match.group(1).strip()

    return ""


def extract_passenger(text):

    value = extract_after_label(
        text,
        [
            "Passenger Name",
            "Passenger",
            "Passenger Name:",
            "Name"
        ]
    )

    if value:
        return value

    return ""


def extract_class(text):

    value = extract_after_label(
        text,
        [
            "Class",
            "Travel Class",
            "Class Type",
            "Booking Class"
        ]
    )

    if value:
        return value

    classes = [
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

    upper_text = normalize_text(
        text
    )

    for travel_class in classes:

        if re.search(
            r"\b" + re.escape(
                travel_class
            ) + r"\b",
            upper_text
        ):
            return travel_class

    return ""


def extract_quota(text):

    value = extract_after_label(
        text,
        [
            "Quota",
            "Booking Quota",
            "Reservation Quota"
        ]
    )

    if value:
        return value

    quotas = [
        "GENERAL",
        "GN",
        "LADIES",
        "LD",
        "TATKAL",
        "CK",
        "PREMIUM TATKAL",
        "PT",
        "SENIOR CITIZEN"
    ]

    upper_text = normalize_text(
        text
    )

    for quota in quotas:

        if re.search(
            r"\b" + re.escape(
                quota
            ) + r"\b",
            upper_text
        ):
            return quota

    return ""


def extract_date(text):

    value = extract_after_label(
        text,
        [
            "Date Of Travel",
            "Date of Travel",
            "Travel Date",
            "Journey Date",
            "Date of Journey",
            "Departure Date"
        ]
    )

    if value:
        return value

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


def extract_location(
    text,
    labels
):

    return extract_after_label(
        text,
        labels
    )


def extract_fare(text):

    patterns = [
        r"(?:Fare\s*(?:/Pax|Per Passenger|Per Pax)?)\s*[:\-]?\s*(?:₹|INR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)",
        r"(?:Ticket Amount|Ticket Fare|Total Fare|Basic Fare)\s*[:\-]?\s*(?:₹|INR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)",
        r"(?:Amount Payable|Amount Paid|Total Amount)\s*[:\-]?\s*(?:₹|INR|Rs\.?)?\s*([\d,]+(?:\.\d+)?)"
    ]

    for pattern in patterns:

        matches = re.findall(
            pattern,
            text,
            re.IGNORECASE
        )

        if matches:
            return matches[-1]

    amounts = re.findall(
        r"(?:₹|INR|Rs\.?)\s*([\d,]+(?:\.\d+)?)",
        text,
        re.IGNORECASE
    )

    if amounts:
        return amounts[-1]

    return ""


def extract_pdf_fields(
    pdf_path
):

    text = extract_pdf_text(
        pdf_path
    )

    return {
        "pnr": extract_pnr(text),

        "passenger": extract_passenger(
            text
        ),

        "class": extract_class(
            text
        ),

        "quota": extract_quota(
            text
        ),

        "travel_date": extract_date(
            text
        ),

        "from": extract_location(
            text,
            [
                "From",
                "Origin",
                "Boarding From",
                "Source"
            ]
        ),

        "to": extract_location(
            text,
            [
                "To",
                "Destination",
                "Boarding To",
                "Drop",
                "Arrival"
            ]
        ),

        "fare": extract_fare(
            text
        )
    }


# ============================================================
# PDF INDEX
# ============================================================

def build_pdf_index(
    tickets_dir
):

    pdf_index = {}

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
            fields.get(
                "pnr",
                ""
            )
        )

        data = {
            "file": os.path.basename(
                pdf_path
            ),
            "path": pdf_path,
            "fields": fields
        }

        # Only index PDFs where a PNR was found
        if pnr:

            pdf_index[pnr] = data

    return pdf_index, pdf_files


# ============================================================
# COMPARISON FUNCTIONS
# ============================================================

def compare_name(
    mis,
    pdf
):

    if not mis or not pdf:
        return False

    return (
        normalize_name(mis)
        == normalize_name(pdf)
    )


def compare_text(
    mis,
    pdf
):

    if not mis or not pdf:
        return False

    return (
        normalize_text(mis)
        == normalize_text(pdf)
    )


def compare_fare(
    mis,
    pdf
):

    return numbers_match(
        mis,
        pdf
    )


# ============================================================
# RAILWAY DETECTION
# ============================================================

def is_railway(mode):

    mode = normalize_text(
        mode
    )

    return (
        "RAIL" in mode
        or "TRAIN" in mode
        or "IRCTC" in mode
    )


# ============================================================
# REMARKS
# ============================================================

def make_remarks(
    name_ok,
    class_ok,
    quota_ok,
    fare_ok,
    railway
):

    problems = []

    if not name_ok:
        problems.append(
            "Passenger Name Mismatch"
        )

    if not class_ok:
        problems.append(
            "Class Mismatch"
        )

    if railway and not quota_ok:
        problems.append(
            "Quota Mismatch"
        )

    if not fare_ok:
        problems.append(
            "Fare Mismatch"
        )

    if not problems:
        return "All required fields matched"

    return "; ".join(
        problems
    )


# ============================================================
# FORMATTING
# ============================================================

def format_sheet(ws):

    # Header
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

    # Data
    for row in ws.iter_rows(
        min_row=2
    ):

        for cell in row:

            cell.border = THIN_BORDER

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        status = normalize_text(
            row[0].value
        )

        if status == "MATCH":

            row[0].fill = MATCH_FILL
            row[0].font = Font(
                bold=True,
                color="006100"
            )

        elif status in [
            "MISMATCH",
            "MISSING PDF",
            "UNRECORDED"
        ]:

            row[0].fill = MISMATCH_FILL
            row[0].font = Font(
                bold=True,
                color="9C0006"
            )

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    # Column width
    for column_cells in ws.columns:

        maximum = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            if cell.value is not None:

                maximum = max(
                    maximum,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max(maximum + 2, 12),
            35
        )


def format_match_columns(ws):

    # H, K, N, Q, T, W, Z
    columns = [
        8,
        11,
        14,
        17,
        20,
        23,
        26
    ]

    for row in ws.iter_rows(
        min_row=2
    ):

        for column in columns:

            cell = row[
                column - 1
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


# ============================================================
# MAIN FUNCTION
# ============================================================

def audit_tickets(
    excel_path,
    tickets_dir,
    output_path
):

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
    # INDEX PDFs
    # --------------------------------------------------------

    pdf_index, pdf_files = build_pdf_index(
        tickets_dir
    )

    # --------------------------------------------------------
    # CREATE OUTPUT
    # --------------------------------------------------------

    wb = Workbook()

    wb.remove(
        wb.active
    )

    summary_ws = wb.create_sheet(
        "Summary"
    )

    reconciliation_ws = wb.create_sheet(
        "Reconciliation"
    )

    audited_ws = wb.create_sheet(
        "Audited MIS Report"
    )

    reconciliation_ws.append(
        RECONCILIATION_HEADERS
    )

    audited_ws.append(
        AUDITED_HEADERS
    )

    # --------------------------------------------------------
    # COUNTERS
    # --------------------------------------------------------

    matched = 0
    mismatched = 0
    missing_pdf = 0
    total_rows = 0

    # PDFs which were actually matched to MIS
    used_pdfs = set()

    # PNRs present in MIS
    mis_pnrs = set()

    # --------------------------------------------------------
    # PROCESS MIS
    # --------------------------------------------------------

    for row in range(
        2,
        source_ws.max_row + 1
    ):

        # Skip completely empty row
        values = [
            source_ws.cell(
                row=row,
                column=col
            ).value

            for col in range(
                1,
                source_ws.max_column + 1
            )
        ]

        if not any(
            v is not None
            and str(v).strip() != ""
            for v in values
        ):
            continue

        total_rows += 1

        # ----------------------------------------------------
        # MIS VALUES
        # ----------------------------------------------------

        corporate_name = get_value(
            source_ws,
            row,
            header_map,
            [
                "Corporate Name"
            ]
        )

        invoice_no = get_value(
            source_ws,
            row,
            header_map,
            [
                "Invoice No",
                "Invoice Number"
            ]
        )

        invoice_date = get_value(
            source_ws,
            row,
            header_map,
            [
                "Invoice Date"
            ]
        )

        booking_date = get_value(
            source_ws,
            row,
            header_map,
            [
                "Date Of Booking",
                "Date of Booking"
            ]
        )

        pnr = get_value(
            source_ws,
            row,
            header_map,
            [
                "PNR Number",
                "PNR",
                "Ticket No",
                "Ticket Number"
            ]
        )

        mode = get_value(
            source_ws,
            row,
            header_map,
            [
                "Mode"
            ]
        )

        train_name = get_value(
            source_ws,
            row,
            header_map,
            [
                "Train_Bus_Flight_Name",
                "Train/Bus/Flight Name",
                "Train Name",
                "Flight Name",
                "Bus Name"
            ]
        )

        train_number = get_value(
            source_ws,
            row,
            header_map,
            [
                "Train_Bus_Flight_Number",
                "Train/Bus/Flight Number",
                "Train Number",
                "Flight Number",
                "Bus Number"
            ]
        )

        travel_date = get_value(
            source_ws,
            row,
            header_map,
            [
                "Date Of Travel",
                "Travel Date"
            ]
        )

        passenger = get_value(
            source_ws,
            row,
            header_map,
            [
                "Passenger Name",
                "Passenger",
                "Name"
            ]
        )

        employee_code = get_value(
            source_ws,
            row,
            header_map,
            [
                "Employee Code"
            ]
        )

        travel_request = get_value(
            source_ws,
            row,
            header_map,
            [
                "Travel Request No",
                "Travel Request Number"
            ]
        )

        cost_center = get_value(
            source_ws,
            row,
            header_map,
            [
                "Cost Center"
            ]
        )

        from_location = get_value(
            source_ws,
            row,
            header_map,
            [
                "From",
                "Origin"
            ]
        )

        to_location = get_value(
            source_ws,
            row,
            header_map,
            [
                "To",
                "Destination"
            ]
        )

        travel_class = get_value(
            source_ws,
            row,
            header_map,
            [
                "Class",
                "Travel Class"
            ]
        )

        quota = get_value(
            source_ws,
            row,
            header_map,
            [
                "Quota"
            ]
        )

        ticket_amount = get_value(
            source_ws,
            row,
            header_map,
            [
                "Ticket Amount",
                "Ticket Fare",
                "Fare",
                "Amount"
            ]
        )

        charges = get_value(
            source_ws,
            row,
            header_map,
            [
                "Charges"
            ]
        )

        gst = get_value(
            source_ws,
            row,
            header_map,
            [
                "GST 18%",
                "GST"
            ]
        )

        total_amount = get_value(
            source_ws,
            row,
            header_map,
            [
                "Total Amount"
            ]
        )

        created_by = get_value(
            source_ws,
            row,
            header_map,
            [
                "Created By"
            ]
        )

        vendor = get_value(
            source_ws,
            row,
            header_map,
            [
                "VendorName",
                "Vendor Name"
            ]
        )

        normalized_pnr = normalize_text(
            pnr
        )

        if normalized_pnr:
            mis_pnrs.add(
                normalized_pnr
            )

        railway = is_railway(
            mode
        )

        # ----------------------------------------------------
        # FIND PDF
        # ----------------------------------------------------

        pdf_data = pdf_index.get(
            normalized_pnr
        )

        # ----------------------------------------------------
        # MISSING PDF
        # ----------------------------------------------------

        if pdf_data is None:

            status = "MISSING PDF"

            missing_pdf += 1

            passenger_pdf = ""
            class_pdf = ""
            quota_pdf = ""
            travel_date_pdf = ""
            from_pdf = ""
            to_pdf = ""
            fare_pdf = ""

            name_match = "NOT CHECKED"
            class_match = "NOT CHECKED"

            if railway:
                quota_match = "NOT CHECKED"
            else:
                quota_match = "N/A"

            date_match = "NOT CHECKED"
            from_match = "NOT CHECKED"
            to_match = "NOT CHECKED"
            fare_match = "NOT CHECKED"

            remarks = (
                "No PDF found for PNR / Ticket No"
            )

        # ----------------------------------------------------
        # PDF FOUND
        # ----------------------------------------------------

        else:

            used_pdfs.add(
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

            fare_pdf = pdf_fields.get(
                "fare",
                ""
            )

            # ------------------------------------------------
            # REQUIRED CHECKS
            # ------------------------------------------------

            name_ok = compare_name(
                passenger,
                passenger_pdf
            )

            class_ok = compare_text(
                travel_class,
                class_pdf
            )

            fare_ok = compare_fare(
                ticket_amount,
                fare_pdf
            )

            # Quota ONLY for Railway
            if railway:

                quota_ok = compare_text(
                    quota,
                    quota_pdf
                )

                quota_match = (
                    "MATCH"
                    if quota_ok
                    else "MISMATCH"
                )

            else:

                quota_ok = True
                quota_match = "N/A"

            # ------------------------------------------------
            # INFORMATIONAL ONLY
            #
            # These do NOT affect status.
            # ------------------------------------------------

            if travel_date and travel_date_pdf:

                date_match = (
                    "MATCH"
                    if normalize_date(
                        travel_date
                    )
                    ==
                    normalize_date(
                        travel_date_pdf
                    )
                    else "MISMATCH"
                )

            else:

                date_match = "NOT CHECKED"

            if from_location and from_pdf:

                from_match = (
                    "MATCH"
                    if normalize_text(
                        from_location
                    )
                    ==
                    normalize_text(
                        from_pdf
                    )
                    else "MISMATCH"
                )

            else:

                from_match = "NOT CHECKED"

            if to_location and to_pdf:

                to_match = (
                    "MATCH"
                    if normalize_text(
                        to_location
                    )
                    ==
                    normalize_text(
                        to_pdf
                    )
                    else "MISMATCH"
                )

            else:

                to_match = "NOT CHECKED"

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

            fare_match = (
                "MATCH"
                if fare_ok
                else "MISMATCH"
            )

            # ------------------------------------------------
            # FINAL STATUS
            #
            # ONLY:
            # Passenger
            # Class
            # Railway Quota
            # Fare
            # ------------------------------------------------

            if (
                name_ok
                and class_ok
                and quota_ok
                and fare_ok
            ):

                status = "MATCH"

                matched += 1

            else:

                status = "MISMATCH"

                mismatched += 1

            remarks = make_remarks(
                name_ok,
                class_ok,
                quota_ok,
                fare_ok,
                railway
            )

        # ----------------------------------------------------
        # RECONCILIATION
        # ----------------------------------------------------

        reconciliation_ws.append([
            status,
            pnr,
            mode,
            created_by,

            (
                pdf_data["file"]
                if pdf_data is not None
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
        ])

        # ----------------------------------------------------
        # AUDITED MIS REPORT
        # ----------------------------------------------------

        audited_ws.append([
            status,
            corporate_name,
            invoice_no,
            invoice_date,
            booking_date,
            pnr,
            mode,
            train_name,
            train_number,
            travel_date,
            passenger,
            employee_code,
            travel_request,
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
            vendor,
            remarks
        ])

    # ========================================================
    # UNRECORDED PDFs
    #
    # PDF exists but its PNR is NOT in MIS.
    # ========================================================

    unrecorded = 0

    for pdf_path in pdf_files:

        if pdf_path in used_pdfs:
            continue

        fields = extract_pdf_fields(
            pdf_path
        )

        pdf_pnr = normalize_text(
            fields.get(
                "pnr",
                ""
            )
        )

        # Only mark as unrecorded when we have a PNR
        # that isn't present in MIS.
        if pdf_pnr and pdf_pnr not in mis_pnrs:

            unrecorded += 1

            reconciliation_ws.append([
                "UNRECORDED",
                fields.get(
                    "pnr",
                    ""
                ),
                "",
                "",
                os.path.basename(
                    pdf_path
                ),
                "",
                fields.get(
                    "passenger",
                    ""
                ),
                "",
                "",
                fields.get(
                    "class",
                    ""
                ),
                "",
                "",
                fields.get(
                    "quota",
                    ""
                ),
                "",
                "",
                fields.get(
                    "travel_date",
                    ""
                ),
                "",
                "",
                fields.get(
                    "from",
                    ""
                ),
                "",
                "",
                fields.get(
                    "to",
                    ""
                ),
                "",
                "",
                "",
                fields.get(
                    "fare",
                    ""
                ),
                "",
                "PDF found in ticket folder but PNR is not present in MIS"
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
    summary_ws["B3"] = matched

    summary_ws["A4"] = "Discrepancies / Mismatch (MISMATCH)"
    summary_ws["B4"] = mismatched

    summary_ws["A5"] = "Missing PDF"
    summary_ws["B5"] = missing_pdf

    summary_ws["A6"] = "Unrecorded Bookings"
    summary_ws["B6"] = unrecorded

    summary_ws["A7"] = "Total MIS Manifest Rows"
    summary_ws["B7"] = total_rows

    summary_ws["A8"] = "Total PDFs Scanned"
    summary_ws["B8"] = len(pdf_files)

    # ========================================================
    # SUMMARY FORMATTING
    # ========================================================

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
    summary_ws["B4"].fill = MISMATCH_FILL
    summary_ws["B5"].fill = MISMATCH_FILL
    summary_ws["B6"].fill = MISMATCH_FILL

    summary_ws.column_dimensions[
        "A"
    ].width = 48

    summary_ws.column_dimensions[
        "B"
    ].width = 20

    # ========================================================
    # FORMAT DETAILED SHEETS
    # ========================================================

    format_sheet(
        reconciliation_ws
    )

    format_sheet(
        audited_ws
    )

    format_match_columns(
        reconciliation_ws
    )

    # ========================================================
    # SAVE
    # ========================================================

    wb.save(
        output_path
    )

    return output_path


# ============================================================
# COMMAND LINE
# ============================================================

if __name__ == "__main__":

    import argparse

    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--excel",
        required=True
    )

    parser.add_argument(
        "--tickets",
        required=True
    )

    parser.add_argument(
        "--output",
        required=True
    )

    args = parser.parse_args()

    output = audit_tickets(
        excel_path=args.excel,
        tickets_dir=args.tickets,
        output_path=args.output
    )

    print(
        "Audit completed successfully:"
    )

    print(output)

