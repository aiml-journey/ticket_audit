
import os
import re
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

from pypdf import PdfReader


# ============================================================
# COLORS
# ============================================================

FILL_GREEN = PatternFill(
    fill_type="solid",
    fgColor="C6EFCE"
)

FILL_YELLOW = PatternFill(
    fill_type="solid",
    fgColor="FFEB9C"
)

FILL_RED = PatternFill(
    fill_type="solid",
    fgColor="FFC7CE"
)

FILL_ORANGE = PatternFill(
    fill_type="solid",
    fgColor="FCE4D6"
)


# ============================================================
# BASIC NORMALIZATION
# ============================================================

def normalize_value(value):
    """
    Convert a value into a standard format for comparison.
    """

    if value is None:
        return ""

    value = str(value).strip().upper()

    # Replace multiple spaces
    value = re.sub(r"\s+", " ", value)

    return value


def normalize_name(value):
    """
    Normalize passenger names.
    """

    if not value:
        return ""

    value = normalize_value(value)

    # Remove punctuation
    value = re.sub(r"[^A-Z0-9 ]", "", value)

    # Remove extra spaces
    value = re.sub(r"\s+", " ", value)

    return value.strip()


# ============================================================
# DATE FUNCTIONS
# ============================================================

def format_date(value):
    """
    Convert common date formats into DD-MM-YYYY.
    """

    if value is None:
        return ""

    if hasattr(value, "strftime"):
        return value.strftime("%d-%m-%Y")

    value = str(value).strip()

    formats = [
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d-%m-%y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d %b %Y",
        "%d %B %Y",
        "%d-%b-%Y",
        "%d-%B-%Y",
    ]

    for fmt in formats:

        try:
            dt = datetime.strptime(value, fmt)
            return dt.strftime("%d-%m-%Y")

        except ValueError:
            pass

    return value


def compare_date(excel_date, pdf_date):
    """
    Compare Excel and PDF travel dates.
    """

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

    return (
        False,
        f"Date mismatch: Excel '{excel_date}' vs PDF '{pdf_date}'"
    )


# ============================================================
# TEXT COMPARISON
# ============================================================

def compare_text(excel_value, pdf_value, field_name):
    """
    Compare two text values.
    """

    excel_value = normalize_value(excel_value)
    pdf_value = normalize_value(pdf_value)

    if not excel_value and not pdf_value:
        return True, None

    if not excel_value:
        return True, None

    if not pdf_value:
        return (
            False,
            f"{field_name} could not be extracted from PDF"
        )

    if excel_value == pdf_value:
        return True, None

    return (
        False,
        f"{field_name} mismatch: "
        f"Excel '{excel_value}' vs PDF '{pdf_value}'"
    )


# ============================================================
# PASSENGER NAME COMPARISON
# ============================================================

def compare_passenger(excel_name, pdf_names):
    """
    Compare MIS passenger name with names extracted from PDF.
    """

    excel_name = normalize_name(excel_name)

    if not excel_name:
        return True, None

    if not pdf_names:
        return (
            False,
            "Passenger name could not be extracted from PDF"
        )

    normalized_pdf_names = [
        normalize_name(name)
        for name in pdf_names
        if name
    ]

    if not normalized_pdf_names:
        return (
            False,
            "Passenger name could not be extracted from PDF"
        )

    # Exact match
    if excel_name in normalized_pdf_names:
        return True, None

    # Compare without spaces
    excel_compact = excel_name.replace(" ", "")

    for name in normalized_pdf_names:

        name_compact = name.replace(" ", "")

        if excel_compact == name_compact:
            return True, None

    # Check partial names
    for name in normalized_pdf_names:

        if (
            excel_name in name
            or name in excel_name
        ):
            return True, None

    return (
        False,
        f"Passenger mismatch: "
        f"Excel '{excel_name}' vs PDF '{', '.join(pdf_names)}'"
    )


# ============================================================
# CLASS NORMALIZATION
# ============================================================

def normalize_class(value):
    """
    Normalize railway/travel class values.
    """

    if not value:
        return ""

    value = normalize_value(value)

    aliases = {

        "AC 3 TIER": "3A",
        "THIRD AC": "3A",
        "THIRD AC TIER": "3A",
        "3A": "3A",

        "AC 2 TIER": "2A",
        "SECOND AC": "2A",
        "SECOND AC TIER": "2A",
        "2A": "2A",

        "AC FIRST CLASS": "1A",
        "FIRST AC": "1A",
        "FIRST CLASS": "1A",
        "1A": "1A",

        "SLEEPER": "SL",
        "SLEEPER CLASS": "SL",
        "SL": "SL",

        "SECOND SITTING": "2S",
        "SECOND SEATING": "2S",
        "2S": "2S",

        "CHAIR CAR": "CC",
        "AC CHAIR CAR": "CC",
        "CC": "CC",

        "EXECUTIVE CLASS": "EC",
        "EXECUTIVE CHAIR CAR": "EC",
        "EC": "EC",

        "AC 3 ECONOMY": "3E",
        "3E": "3E",

        "VISTADOME": "EV",
        "EV": "EV",

        "ANUBHUTI": "EA",
        "EA": "EA",
    }

    return aliases.get(value, value)


def compare_class(excel_class, pdf_class, pdf_class_code):
    """
    Compare MIS class against PDF class.
    """

    excel_class = normalize_class(excel_class)

    pdf_class = normalize_class(pdf_class)
    pdf_class_code = normalize_class(pdf_class_code)

    pdf_final = pdf_class or pdf_class_code

    if not excel_class and not pdf_final:
        return True, None

    if not excel_class:
        return True, None

    if not pdf_final:
        return (
            False,
            "Class could not be extracted from PDF"
        )

    if excel_class == pdf_final:
        return True, None

    return (
        False,
        f"Class mismatch: "
        f"Excel '{excel_class}' vs PDF '{pdf_final}'"
    )


# ============================================================
# QUOTA
# ============================================================

def normalize_quota(value):
    """
    Normalize railway quota values.
    """

    if not value:
        return ""

    value = normalize_value(value)

    aliases = {

        "GENERAL": "GENERAL",
        "GN": "GENERAL",

        "LADIES": "LADIES",
        "LD": "LADIES",

        "TATKAL": "TATKAL",
        "CK": "TATKAL",

        "PREMIUM TATKAL": "PREMIUM TATKAL",
        "PT": "PREMIUM TATKAL",

        "SENIOR CITIZEN": "SENIOR CITIZEN",
        "SS": "SENIOR CITIZEN",

        "LOWER BERTH": "LOWER BERTH",
        "LB": "LOWER BERTH",
    }

    return aliases.get(value, value)


def compare_quota(excel_quota, pdf_quota):
    """
    Compare MIS quota against PDF quota.
    """

    excel_quota = normalize_quota(excel_quota)
    pdf_quota = normalize_quota(pdf_quota)

    if not excel_quota and not pdf_quota:
        return True, None

    if not excel_quota:
        return True, None

    if not pdf_quota:
        return (
            False,
            "Quota could not be extracted from PDF"
        )

    if excel_quota == pdf_quota:
        return True, None

    return (
        False,
        f"Quota mismatch: "
        f"Excel '{excel_quota}' vs PDF '{pdf_quota}'"
    )


# ============================================================
# FARE
# ============================================================

def clean_amount(value):
    """
    Convert currency values to float.
    """

    if value is None:
        return None

    try:

        value = str(value)

        value = (
            value
            .replace(",", "")
            .replace("₹", "")
            .replace("Rs.", "")
            .replace("Rs", "")
            .strip()
        )

        return float(value)

    except (ValueError, TypeError):

        return None


def compare_fare(excel_fare, pdf_fare, passenger_count):
    """
    Compare Excel fare with PDF fare.

    Allows:
    - Full ticket amount
    - Per-passenger amount
    """

    excel_amount = clean_amount(excel_fare)
    pdf_amount = clean_amount(pdf_fare)

    if excel_amount is None:
        return True, None

    if pdf_amount is None:
        return (
            False,
            "Fare could not be extracted from PDF"
        )

    passenger_count = max(
        1,
        passenger_count
    )

    tolerance = 0.05

    # Direct comparison
    if abs(excel_amount - pdf_amount) <= tolerance:
        return True, None

    # Excel = per passenger
    if abs(
        excel_amount * passenger_count
        - pdf_amount
    ) <= tolerance:
        return True, None

    # PDF = per passenger
    if abs(
        excel_amount
        - pdf_amount * passenger_count
    ) <= tolerance:
        return True, None

    return (
        False,
        f"Fare mismatch: "
        f"Excel ₹{excel_amount:.2f} "
        f"vs PDF ₹{pdf_amount:.2f}"
    )


# ============================================================
# TRAIN / FLIGHT / BUS NAME
# ============================================================

def compare_train(excel_train, pdf_train, pdf_class=""):
    """
    Compare train/flight/bus name.

    Allows minor formatting differences.
    """

    excel_train = normalize_value(excel_train)
    pdf_train = normalize_value(pdf_train)

    if not excel_train:
        return True, None

    if not pdf_train:
        return (
            False,
            "Train/Flight/Bus name could not be extracted from PDF"
        )

    if excel_train == pdf_train:
        return True, None

    excel_compact = excel_train.replace(" ", "")
    pdf_compact = pdf_train.replace(" ", "")

    if excel_compact == pdf_compact:
        return True, None

    if (
        excel_train in pdf_train
        or pdf_train in excel_train
    ):
        return True, None

    return (
        False,
        f"Train/Flight/Bus name mismatch: "
        f"Excel '{excel_train}' vs PDF '{pdf_train}'"
    )


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def extract_pdf_text(pdf_path):
    """
    Extract all text from a PDF.
    """

    text_parts = []

    try:

        reader = PdfReader(pdf_path)

        for page in reader.pages:

            try:

                text = page.extract_text()

                if text:
                    text_parts.append(text)

            except Exception:
                pass

    except Exception as e:

        print(
            f"Could not read PDF {pdf_path}: {e}"
        )

    return "\n".join(text_parts)


# ============================================================
# PDF DATA EXTRACTION
# ============================================================

def extract_pdf_data(pdf_path):
    """
    Extract ticket information from PDF.

    IMPORTANT:
    This function never inserts fake passenger/route values.
    If something cannot be extracted, it remains blank.
    """

    text = extract_pdf_text(pdf_path)

    text_upper = text.upper()

    data = {

        "file_name": os.path.basename(pdf_path),

        "pnr": "",

        "passenger_names": [],

        "passenger_name": "",

        "from_station": "",

        "to_station": "",

        "train_name": "",

        "train_no": "",

        "date_of_travel": "",

        "train_class": "",

        "class_code": "",

        "quota": "",

        "total_fare": None,

        "mode": "",

        "extraction_error": False,
    }


    if not text.strip():

        data["extraction_error"] = True

        return data


    # ========================================================
    # PNR
    # ========================================================

    pnr_patterns = [

        r"\bPNR\s*(?:NO\.?|NUMBER)?\s*[:\-]?\s*(\d{8,12})",

        r"\bPNR\s*[:\-]?\s*(\d{8,12})",

    ]

    for pattern in pnr_patterns:

        match = re.search(
            pattern,
            text,
            re.I
        )

        if match:

            data["pnr"] = match.group(1).strip()

            break


    # ========================================================
    # DATE
    # ========================================================

    date_patterns = [

        r"(?:DATE OF JOURNEY|DATE OF TRAVEL|JOURNEY DATE|TRAVEL DATE)"
        r"\s*[:\-]?\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",

        r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})\b",

        r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2})\b",

    ]

    for pattern in date_patterns:

        match = re.search(
            pattern,
            text,
            re.I
        )

        if match:

            data["date_of_travel"] = format_date(
                match.group(1)
            )

            break


    # ========================================================
    # CLASS
    # ========================================================

    class_patterns = [

        r"\b(1A|2A|3A|3E|SL|2S|CC|EC|EV|EA)\b",

        r"(?:CLASS|TRAVEL CLASS)\s*[:\-]?\s*"
        r"(1A|2A|3A|3E|SL|2S|CC|EC|EV|EA)",

    ]

    for pattern in class_patterns:

        match = re.search(
            pattern,
            text_upper,
            re.I
        )

        if match:

            data["class_code"] = (
                match.group(1).upper()
            )

            data["train_class"] = (
                match.group(1).upper()
            )

            break


    # ========================================================
    # QUOTA
    # ========================================================

    quota_patterns = [

        r"QUOTA\s*[:\-]?\s*"
        r"([A-Z][A-Z ]+?)(?:\s*\([A-Z]+\))?(?:\n|$)",

        r"\b(GENERAL|GN|LADIES|LD|TATKAL|CK|"
        r"PREMIUM TATKAL|PT|SENIOR CITIZEN)\b",

    ]

    for pattern in quota_patterns:

        match = re.search(
            pattern,
            text_upper,
            re.I
        )

        if match:

            quota = match.group(1).strip()

            data["quota"] = quota

            break


    # ========================================================
    # FARE
    # ========================================================

    fare_patterns = [

        r"(?:TOTAL FARE|TOTAL AMOUNT|TOTAL PRICE)"
        r"\s*[:\-]?\s*(?:₹|RS\.?|INR)?\s*([\d,]+(?:\.\d{1,2})?)",

        r"(?:FARE|AMOUNT)"
        r"\s*[:\-]?\s*(?:₹|RS\.?|INR)?\s*([\d,]+(?:\.\d{1,2})?)",

    ]

    for pattern in fare_patterns:

        matches = re.findall(
            pattern,
            text_upper,
            re.I
        )

        if matches:

            try:

                data["total_fare"] = float(
                    matches[-1].replace(",", "")
                )

                break

            except ValueError:
                pass


    # ========================================================
    # TRAIN / FLIGHT / BUS NUMBER
    # ========================================================

    number_patterns = [

        r"(?:TRAIN NO|TRAIN NUMBER|TRAIN)\s*[:\-]?\s*(\d{4,6})",

        r"(?:FLIGHT NO|FLIGHT NUMBER|FLIGHT)\s*[:\-]?\s*"
        r"([A-Z]{1,3}\s*\d{2,5})",

        r"(?:BUS NO|BUS NUMBER|BUS)\s*[:\-]?\s*"
        r"([A-Z0-9\-]{2,15})",

    ]

    for pattern in number_patterns:

        match = re.search(
            pattern,
            text_upper,
            re.I
        )

        if match:

            data["train_no"] = (
                match.group(1).strip()
            )

            break


    # ========================================================
    # TRAIN NAME
    # ========================================================

    train_patterns = [

        r"(?:TRAIN NAME|TRAIN)\s*[:\-]\s*"
        r"([A-Za-z0-9 .&()\-]+)",

    ]

    for pattern in train_patterns:

        match = re.search(
            pattern,
            text,
            re.I
        )

        if match:

            name = match.group(1).strip()

            # Avoid taking a very long unrelated line
            if len(name) <= 100:

                data["train_name"] = name

                break


    # ========================================================
    # FROM / TO
    # ========================================================

    route_patterns = [

        r"(?:FROM)\s*[:\-]?\s*"
        r"([A-Za-z .,'()\-]+?)"
        r"\s+(?:TO)\s*[:\-]?\s*"
        r"([A-Za-z .,'()\-]+?)(?:\n|$)",

    ]

    for pattern in route_patterns:

        match = re.search(
            pattern,
            text,
            re.I
        )

        if match:

            data["from_station"] = (
                match.group(1).strip()
            )

            data["to_station"] = (
                match.group(2).strip()
            )

            break


    # ========================================================
    # PASSENGER NAMES
    # ========================================================

    passenger_patterns = [

        r"(?:PASSENGER NAME|PASSENGER|NAME)"
        r"\s*[:\-]\s*([A-Za-z .]+)",

    ]

    for pattern in passenger_patterns:

        matches = re.findall(
            pattern,
            text,
            re.I
        )

        for name in matches:

            name = name.strip()

            if (
                len(name) >= 3
                and len(name) <= 80
                and not re.search(
                    r"PNR|DATE|TRAIN|CLASS|QUOTA|FARE|"
                    r"FROM|TO|AGE|GENDER",
                    name,
                    re.I
                )
            ):

                data["passenger_names"].append(
                    name
                )


    # Remove duplicates
    data["passenger_names"] = list(
        dict.fromkeys(
            data["passenger_names"]
        )
    )


    if data["passenger_names"]:

        data["passenger_name"] = (
            data["passenger_names"][0]
        )


    # ========================================================
    # MODE
    # ========================================================

    if (
        "FLIGHT" in text_upper
        or "AIRLINE" in text_upper
        or "BOARDING PASS" in text_upper
    ):

        data["mode"] = "FLIGHT"

    elif (
        "TRAIN" in text_upper
        or "RAILWAY" in text_upper
        or "IRCTC" in text_upper
    ):

        data["mode"] = "TRAIN"

    elif "BUS" in text_upper:

        data["mode"] = "BUS"


    return data


# ============================================================
# EXCEL COLUMN MAPPING
# ============================================================

def get_column_map(ws):
    """
    Map Excel headers to column numbers.
    """

    column_map = {}

    for col in range(
        1,
        ws.max_column + 1
    ):

        value = ws.cell(
            1,
            col
        ).value

        if value is None:
            continue

        header = normalize_value(
            value
        )

        column_map[header] = col

    return column_map


# ============================================================
# FIND HEADER
# ============================================================

def find_column(column_map, *names):
    """
    Find a column using multiple possible names.
    """

    for name in names:

        normalized = normalize_value(
            name
        )

        if normalized in column_map:

            return column_map[
                normalized
            ]

    return None


# ============================================================
# MAIN AUDIT FUNCTION
# ============================================================

def audit_tickets(
    excel_path,
    tickets_dir,
    output_path=None
):
    """
    Main ticket auditing function.

    Parameters:
        excel_path:
            Path to MIS Excel file.

        tickets_dir:
            Folder containing ticket PDFs.

        output_path:
            Where audited Excel should be saved.

    Returns:
        Dictionary containing saved file details.
    """

    excel_path = Path(
        excel_path
    )

    tickets_dir = Path(
        tickets_dir
    )


    if output_path:

        output_path = Path(
            output_path
        )

    else:

        output_path = (
            excel_path.parent
            / f"Audited_{excel_path.name}"
        )


    # ========================================================
    # LOAD EXCEL
    # ========================================================

    wb = load_workbook(
        excel_path
    )

    ws = wb.active


    column_map = get_column_map(
        ws
    )


    # ========================================================
    # FIND IMPORTANT COLUMNS
    # ========================================================

    col_pnr = find_column(
        column_map,
        "PNR",
        "PNR NO",
        "PNR NUMBER"
    )

    col_passenger = find_column(
        column_map,
        "PASSENGER NAME",
        "PASSENGER",
        "NAME"
    )

    col_from = find_column(
        column_map,
        "FROM",
        "FROM STATION",
        "SOURCE"
    )

    col_to = find_column(
        column_map,
        "TO",
        "TO STATION",
        "DESTINATION"
    )

    col_mode = find_column(
        column_map,
        "MODE",
        "TRAVEL MODE"
    )

    col_train_name = find_column(
        column_map,
        "TRAIN_NAME",
        "TRAIN NAME",
        "FLIGHT NAME",
        "BUS NAME"
    )

    col_train_no = find_column(
        column_map,
        "TRAIN_NO",
        "TRAIN NO",
        "TRAIN NUMBER",
        "FLIGHT NO",
        "FLIGHT NUMBER",
        "BUS NO"
    )

    col_date = find_column(
        column_map,
        "DATE_OF_TRAVEL",
        "DATE OF TRAVEL",
        "TRAVEL DATE",
        "JOURNEY DATE"
    )

    col_class = find_column(
        column_map,
        "CLASS",
        "TRAIN CLASS",
        "TRAVEL CLASS"
    )

    col_quota = find_column(
        column_map,
        "QUOTA"
    )

    col_fare = find_column(
        column_map,
        "FARE",
        "TICKET FARE",
        "AMOUNT",
        "TICKET AMOUNT",
        "TOTAL AMOUNT"
    )


    # ========================================================
    # REMARK COLUMN
    # ========================================================

    remark_col = find_column(
        column_map,
        "REMARKS",
        "REMARK",
        "AUDIT REMARKS",
        "AUDIT REMARK"
    )

    if remark_col is None:

        remark_col = (
            ws.max_column + 1
        )

        ws.cell(
            1,
            remark_col
        ).value = "Audit Remarks"

        ws.cell(
            1,
            remark_col
        ).font = Font(
            bold=True
        )


    # ========================================================
    # READ ALL PDF FILES
    # ========================================================

    pdf_files = list(
        tickets_dir.glob("*.pdf")
    )

    pdf_records = {}

    pdf_without_pnr = []

    duplicate_pnrs = []


    for pdf_path in pdf_files:

        print(
            f"Reading PDF: {pdf_path.name}"
        )

        pdf_data = extract_pdf_data(
            pdf_path
        )

        pnr = pdf_data["pnr"]


        if not pnr:

            pdf_without_pnr.append(
                pdf_data
            )

            continue


        if pnr in pdf_records:

            duplicate_pnrs.append(
                pnr
            )

            continue


        pdf_records[pnr] = pdf_data


    # ========================================================
    # AUDIT EACH MIS ROW
    # ========================================================

    matched_mis_pnrs = set()


    for row in range(
        2,
        ws.max_row + 1
    ):

        row_discrepancies = []

        has_red_error = False

        has_yellow_discrepancy = False


        # ----------------------------------------------------
        # GET MIS PNR
        # ----------------------------------------------------

        excel_pnr = ""

        if col_pnr:

            excel_pnr = normalize_value(
                ws.cell(
                    row,
                    col_pnr
                ).value
            )


        if not excel_pnr:

            ws.cell(
                row,
                remark_col
            ).value = (
                "Manual Review: PNR missing from MIS"
            )

            ws.cell(
                row,
                remark_col
            ).fill = FILL_RED

            continue


        # ----------------------------------------------------
        # FIND PDF
        # ----------------------------------------------------

        if excel_pnr not in pdf_records:

            ws.cell(
                row,
                remark_col
            ).value = (
                "Ticket PDF not found for PNR "
                + excel_pnr
            )

            ws.cell(
                row,
                remark_col
            ).fill = FILL_RED

            continue


        matched_mis_pnrs.add(
            excel_pnr
        )


        pdf = pdf_records[
            excel_pnr
        ]


        # ----------------------------------------------------
        # PASSENGER
        # ----------------------------------------------------

        if col_passenger:

            excel_passenger = (
                ws.cell(
                    row,
                    col_passenger
                ).value
            )

            match, remark = compare_passenger(
                excel_passenger,
                pdf["passenger_names"]
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_passenger
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # FROM
        # ----------------------------------------------------

        if col_from:

            excel_from = (
                ws.cell(
                    row,
                    col_from
                ).value
            )

            match, remark = compare_text(
                excel_from,
                pdf["from_station"],
                "From"
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_from
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # TO
        # ----------------------------------------------------

        if col_to:

            excel_to = (
                ws.cell(
                    row,
                    col_to
                ).value
            )

            match, remark = compare_text(
                excel_to,
                pdf["to_station"],
                "To"
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_to
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # MODE
        # ----------------------------------------------------

        if col_mode:

            excel_mode = (
                ws.cell(
                    row,
                    col_mode
                ).value
            )

            match, remark = compare_text(
                excel_mode,
                pdf["mode"],
                "Mode"
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_mode
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # TRAIN / FLIGHT / BUS NAME
        # ----------------------------------------------------

        if col_train_name:

            excel_train = (
                ws.cell(
                    row,
                    col_train_name
                ).value
            )

            match, remark = compare_train(
                excel_train,
                pdf["train_name"],
                pdf["train_class"]
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_train_name
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # TRAIN / FLIGHT / BUS NUMBER
        # ----------------------------------------------------

        if col_train_no:

            excel_train_no = (
                ws.cell(
                    row,
                    col_train_no
                ).value
            )

            match, remark = compare_text(
                excel_train_no,
                pdf["train_no"],
                "Train/Flight/Bus number"
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_train_no
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # DATE
        # ----------------------------------------------------

        if col_date:

            excel_date = (
                ws.cell(
                    row,
                    col_date
                ).value
            )

            match, remark = compare_date(
                excel_date,
                pdf["date_of_travel"]
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_date
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # CLASS
        # ----------------------------------------------------

        if col_class:

            excel_class = (
                ws.cell(
                    row,
                    col_class
                ).value
            )

            match, remark = compare_class(
                excel_class,
                pdf["train_class"],
                pdf["class_code"]
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_class
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # QUOTA
        # ----------------------------------------------------

        if col_quota:

            excel_quota = (
                ws.cell(
                    row,
                    col_quota
                ).value
            )

            match, remark = compare_quota(
                excel_quota,
                pdf["quota"]
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_quota
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # FARE
        # ----------------------------------------------------

        if col_fare:

            excel_fare = (
                ws.cell(
                    row,
                    col_fare
                ).value
            )

            passenger_count = max(
                1,
                len(
                    pdf["passenger_names"]
                )
            )

            match, remark = compare_fare(
                excel_fare,
                pdf["total_fare"],
                passenger_count
            )

            if not match:

                has_yellow_discrepancy = True

                row_discrepancies.append(
                    remark
                )

                ws.cell(
                    row,
                    col_fare
                ).fill = FILL_YELLOW


        # ----------------------------------------------------
        # FINAL REMARK
        # ----------------------------------------------------

        if has_red_error:

            final_remark = (
                "Critical audit error"
            )

            ws.cell(
                row,
                remark_col
            ).fill = FILL_RED

        elif has_yellow_discrepancy:

            final_remark = (
                "Discrepancy: "
                + "; ".join(
                    row_discrepancies
                )
            )

            ws.cell(
                row,
                remark_col
            ).fill = FILL_YELLOW

        else:

            final_remark = "Matched"

            ws.cell(
                row,
                remark_col
            ).fill = FILL_GREEN


        ws.cell(
            row,
            remark_col
        ).value = final_remark


    # ========================================================
    # FIND PDFs NOT PRESENT IN MIS
    # ========================================================

    extra_pdfs = []

    for pnr, pdf_data in pdf_records.items():

        if pnr not in matched_mis_pnrs:

            extra_pdfs.append(
                pdf_data
            )


    # ========================================================
    # ADD AUDIT SUMMARY SHEET
    # ========================================================

    if "Audit Summary" in wb.sheetnames:

        del wb["Audit Summary"]


    summary = wb.create_sheet(
        "Audit Summary"
    )


    summary["A1"] = "Ticket Audit Summary"

    summary["A1"].font = Font(
        bold=True,
        size=16
    )


    summary_data = [

        (
            "MIS Tickets",
            ws.max_row - 1
        ),

        (
            "PDF Tickets",
            len(pdf_files)
        ),

        (
            "Matched PNRs",
            len(matched_mis_pnrs)
        ),

        (
            "Missing PDF Tickets",
            (ws.max_row - 1)
            - len(matched_mis_pnrs)
        ),

        (
            "PDFs Not Present in MIS",
            len(extra_pdfs)
        ),

        (
            "PDFs Without PNR",
            len(pdf_without_pnr)
        ),

        (
            "Duplicate PNRs",
            len(duplicate_pnrs)
        ),

    ]


    start_row = 3


    for index, (label, value) in enumerate(
        summary_data,
        start=start_row
    ):

        summary.cell(
            index,
            1
        ).value = label

        summary.cell(
            index,
            2
        ).value = value


    # ========================================================
    # EXTRA PDF SECTION
    # ========================================================

    row = start_row + len(
        summary_data
    ) + 2


    summary.cell(
        row,
        1
    ).value = "PDFs Not Present in MIS"

    summary.cell(
        row,
        1
    ).font = Font(
        bold=True
    )


    row += 1


    for pdf_data in extra_pdfs:

        summary.cell(
            row,
            1
        ).value = pdf_data["pnr"]

        summary.cell(
            row,
            2
        ).value = pdf_data["file_name"]

        row += 1


    # ========================================================
    # PDFs WITHOUT PNR
    # ========================================================

    row += 1


    summary.cell(
        row,
        1
    ).value = "PDFs Where PNR Could Not Be Extracted"

    summary.cell(
        row,
        1
    ).font = Font(
        bold=True
    )


    row += 1


    for pdf_data in pdf_without_pnr:

        summary.cell(
            row,
            1
        ).value = pdf_data["file_name"]

        summary.cell(
            row,
            2
        ).value = (
            "Manual Review Required"
        )

        row += 1


    # ========================================================
    # DUPLICATE PNRS
    # ========================================================

    row += 1


    summary.cell(
        row,
        1
    ).value = "Duplicate PNRs"

    summary.cell(
        row,
        1
    ).font = Font(
        bold=True
    )


    row += 1


    for pnr in duplicate_pnrs:

        summary.cell(
            row,
            1
        ).value = pnr

        summary.cell(
            row,
            2
        ).value = (
            "Duplicate PNR found in PDFs"
        )

        row += 1


    # ========================================================
    # FORMAT SUMMARY
    # ========================================================

    for column in summary.columns:

        max_length = 0

        column_letter = (
            column[0].column_letter
        )

        for cell in column:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        summary.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            60
        )


    # ========================================================
    # SAVE
    # ========================================================

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True
    )


    wb.save(
        output_path
    )


    print(
        f"\nAudit completed successfully."
    )

    print(
        f"Saved file: {output_path}"
    )


    return {

        "saved_file": str(
            output_path
        ),

        "saved_filename": (
            output_path.name
        ),

        "mis_tickets": (
            ws.max_row - 1
        ),

        "pdf_tickets": len(
            pdf_files
        ),

        "matched": len(
            matched_mis_pnrs
        ),

        "missing_pdf": (
            (ws.max_row - 1)
            - len(matched_mis_pnrs)
        ),

        "extra_pdf": len(
            extra_pdfs
        ),

        "pdf_without_pnr": len(
            pdf_without_pnr
        ),

        "duplicate_pnr": len(
            duplicate_pnrs
        ),
    }


# ============================================================
# RUN DIRECTLY
# ============================================================

if __name__ == "__main__":

    print(
        "audit_tickets.py"
    )

    print(
        "This file contains the ticket auditing engine."
    )

    print(
        "Use audit_tickets() from app.py to run the audit."
    )
