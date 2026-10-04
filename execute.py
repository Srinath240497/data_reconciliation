from pathlib import Path
import pandas as pd

BASE = Path.cwd()
DATA_CANDIDATES = (BASE / "data", BASE / "candidate_pack" / "data")
DATA = next(
    path for path in DATA_CANDIDATES
    if (path / "bank_transactions_may_2026.csv").exists()
)
REMITTANCES = DATA / "remittances"
BANK_PATH = DATA / "bank_transactions_may_2026.csv"
BDX_PATH = DATA / "premium_bdx_april_2026.csv"

bank = pd.read_csv(BANK_PATH, dtype=str); 
bdx = pd.read_csv(BDX_PATH, dtype=str); 
remittance_files = sorted(REMITTANCES.iterdir()); 
bank.head(), bdx.head(), [p.name for p in remittance_files]

import csv
import warnings
# These workbooks have no default style. Ignore that warning so the preview stays readable.
warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style",
    module="openpyxl",
)

PREVIEW_ROW_COUNT = 3
# A cover sheet is a short list of labels, not a column table. Show the whole list
# when it is small so totals and payment notes stay visible.
SHORT_LABEL_SHEET = 8


def cell_text(value):
    """Return the cell as text, including leading zeros in identifiers."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return str(value).strip()


def is_column_label(value):
    """A column label is a word such as 'Policy No', not an amount, date, or reference."""
    if value == "":
        return False
    has_letter = any(character.isalpha() for character in value)
    has_digit = any(character.isdigit() for character in value)
    return has_letter and not has_digit


def find_header_index(rows):
    """Find the row of column names.

    Some brokers put a short banner above the table. The column row is the
    widest row that contains only words.
    """
    header_index = None
    header_width = 1
    for index, row in enumerate(rows):
        labels = [cell for cell in row if is_column_label(cell)]
        other_values = [cell for cell in row if cell != "" and cell not in labels]
        if other_values:
            continue
        if len(labels) > header_width:
            header_width = len(labels)
            header_index = index
    return header_index


def show_preview(filename, location, rows):
    """Print the source, the field names, and the first three data rows."""
    print("=" * 72)
    print("File:", filename)
    if location:
        print("Sheet or table:", location)

    stored_rows = [row for row in rows if any(cell != "" for cell in row)]
    if not stored_rows:
        print("Column or field names: none")
        print("First three data rows: this sheet or table is empty.")
        return

    header_index = find_header_index(rows)
    if header_index is None:
        print("Column or field names: no column header on this sheet or table.")
        rows_to_show = stored_rows
        if len(stored_rows) > SHORT_LABEL_SHEET:
            rows_to_show = stored_rows[:PREVIEW_ROW_COUNT]
        print("Stored rows:")
        for row in rows_to_show:
            visible = [cell for cell in row if cell != ""]
            print(" -", " | ".join(visible))
        print(f"Non-empty rows: {len(stored_rows)}. Showing {len(rows_to_show)}.")
        return

    banner_rows = [row for row in rows[:header_index] if any(cell != "" for cell in row)]
    if banner_rows:
        print("Rows above the column header:")
        for row in banner_rows:
            visible = [cell for cell in row if cell != ""]
            print(" -", " | ".join(visible))

    field_names = list(rows[header_index])
    while field_names and field_names[-1] == "":
        field_names.pop()

    print("Column or field names:")
    print(field_names)

    data_rows = []
    for row in rows[header_index + 1 :]:
        if not any(cell != "" for cell in row):
            continue
        values = list(row[: len(field_names)])
        if len(values) < len(field_names):
            values.extend([""] * (len(field_names) - len(values)))
        data_rows.append(values)
        if len(data_rows) == PREVIEW_ROW_COUNT:
            break

    print("First three data rows:")
    if not data_rows:
        print(" (no data rows under this header)")
        return

    preview = pd.DataFrame(data_rows, columns=field_names)
    print(preview.to_string(index=False, max_cols=None, max_colwidth=None))


def read_csv_rows(path):
    """Read a CSV as text. Some files have a banner, so rows are not all the same width."""
    with path.open(newline="", encoding="utf-8-sig") as handle:
        return [[cell_text(value) for value in row] for row in csv.reader(handle)]


def read_excel_sheets(path):
    """Read every sheet as text. dtype=str keeps identifiers unchanged."""
    workbook = pd.read_excel(
        path,
        sheet_name=None,
        header=None,
        dtype=str,
        engine="openpyxl",
    )
    sheets = []
    for sheet_name, frame in workbook.items():
        rows = []
        for record in frame.itertuples(index=False, name=None):
            rows.append([cell_text(value) for value in record])
        sheets.append((sheet_name, rows))
    return sheets


def preview_pdf(path):
    """Extract PDF tables when a PDF reader is installed. Otherwise name the missing package."""
    try:
        import pdfplumber # type: ignore
    except ImportError:
        print("=" * 72)
        print("File:", path.name)
        print(
            "Missing package: pdfplumber. "
            "pandas and openpyxl do not extract PDF tables, "
            "and pdfplumber is not installed in this project."
        )
        return

    found_table = False
    with pdfplumber.open(path) as document:
        for page_number, page in enumerate(document.pages, start=1):
            for table_number, table in enumerate(page.extract_tables() or [], start=1):
                if not table:
                    continue
                found_table = True
                rows = [
                    [cell_text(value) for value in (raw_row or [])]
                    for raw_row in table
                ]
                location = f"page {page_number}, table {table_number}"
                show_preview(path.name, location, rows)

    if not found_table:
        print("=" * 72)
        print("File:", path.name)
        print("No extractable table in this PDF.")


# Build the file list again on each run. This cell only reads the files.
remittance_files = sorted(
    path
    for path in REMITTANCES.iterdir()
    if path.is_file() and not path.name.startswith(".")
)

if not remittance_files:
    print("No files found in", REMITTANCES)

for remittance_path in remittance_files:
    suffix = remittance_path.suffix.lower()
    try:
        if suffix == ".csv":
            show_preview(remittance_path.name, None, read_csv_rows(remittance_path))
        elif suffix in {".xlsx", ".xlsm"}:
            check = read_excel_sheets(remittance_path)
            for sheet_name, sheet_rows in read_excel_sheets(remittance_path):
                show_preview(remittance_path.name, sheet_name, sheet_rows)
        elif suffix == ".xls":
            print("=" * 72)
            print("File:", remittance_path.name)
            print("Missing package: xlrd. openpyxl reads .xlsx files, but not older .xls files.")
#         elif suffix == ".pdf":
#             preview_pdf(remittance_path)
        else:
            print("=" * 72)
            print("File:", remittance_path.name)
            print(f"Unsupported file type '{suffix or '(none)'}' for this inspection.")
    except ImportError as error:
        print("=" * 72)
        print("File:", remittance_path.name)
        print("Missing package:", error)
    except Exception as error:
        print("=" * 72)
        print("File:", remittance_path.name)
        print("Could not read this file:", error)

from IPython.display import display

PROVENANCE = ("source_file", "source_sheet", "source_row")

KEY_FIELDS = {
    "bank_transactions_may_2026.csv": [
        "transaction_id", "booking_date", "currency", "credit_amount", "counterparty_name", "narrative",
    ],
    "premium_bdx_april_2026.csv": [
        "bdx_record_id", "broker_name", "policy_reference", "original_currency",
        "net_due_to_mga_original_ccy", "settlement_amount_gbp",
    ],
    "apex_risk_partners_remittance_apr_2026.csv": [
        "Policy No", "Insured", "Gross Premium", "Net Settlement",
    ],
    "pioneer_wholesale_statement_may_2026.csv": [
        "Policy Reference", "Client", "Currency", "Amount Due",
    ],
    "continental_coverholders_apr_2026.xlsx": [
        "Policy", "Original CCY", "Net Due Original CCY", "GBP Settlement Amount",
    ],
    "harbour_specialty_remittance_apr_2026.xlsx": [
        "Policy / Risk Ref", "Insured Name", "Gross", "Net Due",
    ],
    "northshore_broking_remittance_apr_2026.xlsx": [
        "Contract No", "UW Ref", "Currency", "Amount Due MGA",
    ],
}


def trim_field_names(row):
    names = list(row)
    while names and names[-1] == "":
        names.pop()
    used = {}
    cleaned = []
    for index, name in enumerate(names):
        label = name or f"unnamed_{index + 1}"
        used[label] = used.get(label, 0) + 1
        if used[label] > 1:
            label = f"{label}_{used[label]}"
        cleaned.append(label)
    return cleaned


def control_record(source_file, source_sheet, source_row, control_type, label, value):
    return {
        "source_file": source_file,
        "source_sheet": source_sheet,
        "source_row": source_row,
        "control_type": control_type,
        "label": label,
        "value": value,
    }


def split_detail_and_controls(rows, source_file, source_sheet=""):
    """Separate a policy table from banners, cover sheets, and statement totals."""
    controls = []
    header_index = find_header_index(rows)
    if header_index is None:
        for row_number, row in enumerate(rows, start=1):
            visible = [cell for cell in row if cell != ""]
            if not visible:
                continue
            value = " | ".join(visible[1:]) if len(visible) > 1 else ""
            controls.append(control_record(
                source_file, source_sheet, row_number, "cover_sheet", visible[0], value,
            ))
        return pd.DataFrame(), controls

    for row_number, row in enumerate(rows[:header_index], start=1):
        visible = [cell for cell in row if cell != ""]
        if not visible:
            continue
        value = " | ".join(visible[1:]) if len(visible) > 1 else ""
        controls.append(control_record(
            source_file, source_sheet, row_number, "banner", visible[0], value,
        ))

    field_names = trim_field_names(rows[header_index])
    detail_rows = []
    for row_number, row in enumerate(rows[header_index + 1:], start=header_index + 2):
        if not any(cell != "" for cell in row):
            continue
        values = list(row[:len(field_names)])
        if len(values) < len(field_names):
            values.extend([""] * (len(field_names) - len(values)))
        first_value = next((cell for cell in values if cell != ""), "")
        if first_value.casefold() == "total":
            amount = next((cell for cell in reversed(values) if cell and cell.casefold() != "total"), "")
            controls.append(control_record(
                source_file, source_sheet, row_number, "statement_total", "TOTAL", amount,
            ))
            continue
        record = dict(zip(field_names, values))
        record["source_file"] = source_file
        record["source_sheet"] = source_sheet
        record["source_row"] = row_number
        detail_rows.append(record)

    return pd.DataFrame(detail_rows), controls


def load_table(path, sheet_name="", rows=None):
    if rows is None:
        rows = read_csv_rows(path)
    return split_detail_and_controls(rows, path.name, sheet_name)


bank_transactions, bank_controls = load_table(BANK_PATH)
premium_bdx, bdx_controls = load_table(BDX_PATH)

remittance_details = {}
remittance_controls = []
for remittance_path in remittance_files:
    suffix = remittance_path.suffix.lower()
    if suffix == ".csv":
        detail, controls = load_table(remittance_path)
        remittance_details[remittance_path.name] = detail
        remittance_controls.extend(controls)
    elif suffix in {".xlsx", ".xlsm"}:
        for sheet_name, sheet_rows in read_excel_sheets(remittance_path):
            detail, controls = load_table(remittance_path, sheet_name, sheet_rows)
            if not detail.empty:
                remittance_details[f"{remittance_path.name} | {sheet_name}"] = detail
            remittance_controls.extend(controls)

source_controls = pd.DataFrame(bank_controls + bdx_controls + remittance_controls)
loaded_sources = {
    "bank_transactions": bank_transactions,
    "premium_bdx": premium_bdx,
    "remittance_details": remittance_details,
    "source_controls": source_controls,
}


def missing_key_fields(frame, filename):
    notes = []
    for column in KEY_FIELDS.get(filename, []):
        if column not in frame.columns:
            notes.append(f"{column}: column absent")
            continue
        missing = int(frame[column].eq("").sum())
        if missing:
            notes.append(f"{column}: {missing}")
    return "; ".join(notes) if notes else "none"


CURRENCY_CODES = ("GBP", "EUR", "USD")


def currency_codes(value):
    """Currency codes stored as the whole value or as an amount prefix such as 'GBP 1,200.00'."""
    upper = value.upper()
    found = []
    for code in CURRENCY_CODES:
        if upper == code or upper.startswith(code + " "):
            found.append(code)
    return found


def currencies_in(frame):
    found = set()
    for column in frame.columns:
        if column in PROVENANCE:
            continue
        for value in frame[column]:
            if value:
                found.update(currency_codes(value))
    return found


def currencies_text(frame, control_text=""):
    found = set(currencies_in(frame))
    for code in CURRENCY_CODES:
        if code in control_text:
            found.add(code)
    return ", ".join(sorted(found))


def quality_notes(frame):
    notes = []
    business_columns = [column for column in frame.columns if column not in PROVENANCE]
    duplicate_count = int(frame.duplicated(subset=business_columns, keep=False).sum())
    if duplicate_count:
        notes.append(f"{duplicate_count} detail rows are exact duplicates of another row")
    for column in business_columns:
        values = [value for value in frame[column] if value]
        has_currency_prefix = any(value[:3].upper() in {"GBP", "EUR", "USD"} for value in values)
        has_bare_number = any(value[:1].isdigit() or value[:1] in "-(" for value in values)
        if has_currency_prefix and has_bare_number:
            notes.append(f"{column} mixes currency prefixes and bare numbers")
        date_like = [value for value in values if "/" in value or (len(value) >= 10 and value[4:5] == "-")]
        slash_dates = sum("/" in value for value in date_like)
        iso_dates = sum(len(value) >= 10 and value[4:5] == "-" for value in date_like)
        if slash_dates and iso_dates:
            notes.append(f"{column} mixes slash dates and ISO dates")
    blank_columns = [
        column for column in business_columns
        if frame[column].eq("").any()
    ]
    if blank_columns:
        notes.append("blank values in " + ", ".join(blank_columns))
    return "; ".join(notes) if notes else "none noted"


def controls_for(filename, sheet_name=""):
    if source_controls.empty:
        return ""
    matched = source_controls[source_controls["source_file"].eq(filename)]
    if sheet_name:
        matched = matched[matched["source_sheet"].eq(sheet_name)]
    elif filename.endswith(".csv"):
        matched = matched[matched["source_sheet"].eq("")]
    if matched.empty:
        return ""
    return "; ".join(
        f"{row.control_type}: {row.label}={row.value} (row {row.source_row})"
        for row in matched.itertuples(index=False)
    )


profile_rows = []
for label, frame in (
    ("bank_transactions", bank_transactions),
    ("premium_bdx", premium_bdx),
):
    filename = frame["source_file"].iloc[0]
    profile_rows.append({
        "dataset": label,
        "source_file": filename,
        "source_sheet": "",
        "detail_rows": len(frame),
        "columns": ", ".join(column for column in frame.columns if column not in PROVENANCE),
        "missing_key_fields": missing_key_fields(frame, filename),
        "currencies": currencies_text(frame, controls_for(filename)),
        "controls": controls_for(filename),
        "observations": quality_notes(frame),
    })

for dataset_name, frame in remittance_details.items():
    filename = frame["source_file"].iloc[0]
    sheet_name = frame["source_sheet"].iloc[0]
    control_text = controls_for(filename, sheet_name)
    profile_rows.append({
        "dataset": dataset_name,
        "source_file": filename,
        "source_sheet": sheet_name,
        "detail_rows": len(frame),
        "columns": ", ".join(column for column in frame.columns if column not in PROVENANCE),
        "missing_key_fields": missing_key_fields(frame, filename),
        "currencies": currencies_text(frame, control_text),
        "controls": control_text,
        "observations": quality_notes(frame),
    })

# Cover sheets have no detail rows, so add them from the control table.
cover_sheets = source_controls[source_controls["control_type"].eq("cover_sheet")]
for (filename, sheet_name), group in cover_sheets.groupby(["source_file", "source_sheet"], sort=False):
    profile_rows.append({
        "dataset": f"{filename} | {sheet_name}",
        "source_file": filename,
        "source_sheet": sheet_name,
        "detail_rows": 0,
        "columns": "",
        "missing_key_fields": "",
        "currencies": ", ".join(
            code for code in CURRENCY_CODES
            if code in " ".join(group["value"].astype(str))
        ),
        "controls": "; ".join(
            f"cover_sheet: {row.label}={row.value} (row {row.source_row})"
            for row in group.itertuples(index=False)
        ),
        "observations": "cover or summary sheet; no policy-level rows",
    })

source_profile = pd.DataFrame(profile_rows)
pd.set_option("display.max_colwidth", 120)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", None)
# display(source_profile)
print(source_profile.to_string(index=False))
print(source_controls.to_string(index=False))
# source_controls

# import re

# ROUNDING_TOLERANCE = 0.01  # GBP; differences above one penny are reported

# NAME_STOPWORDS = {
#     "LTD", "LIMITED", "PLC", "SA", "LLP", "INC",
#     "BROKER", "BROKERS", "BROKING", "WHOLESALE", "PARTNERS",
#     "SPECIALTY", "COVERHOLDERS", "RISK", "SERVICES", "AND", "THE",
# }
# REF_LABELS = {"statement id", "statement ref", "remittance id"}
# TOTAL_LABELS = {"stated total", "statement total"}
# BROKER_LABELS = {"broker"}
# CCY_LABELS = {"settlement ccy", "currency"}


# def parse_amount(text):
#     """Parse a stated amount, including a currency prefix and thousands separators."""
#     if text is None or str(text).strip() == "":
#         return None
#     raw = str(text).strip()
#     negative = raw.startswith("(") and raw.endswith(")")
#     cleaned = raw.strip("()")
#     cleaned = re.sub(r"\b[A-Z]{3}\b", "", cleaned)
#     cleaned = cleaned.replace(",", "").strip()
#     if cleaned in {"", "-"}:
#         return None
#     value = float(cleaned)
#     return -value if negative else value


# def currency_of(text):
#     if text is None:
#         return ""
#     match = re.search(r"\b(GBP|EUR|USD)\b", str(text).upper())
#     return match.group(1) if match else ""


# def name_tokens(broker):
#     tokens = re.findall(r"[A-Z0-9]+", str(broker).upper())
#     return [token for token in tokens if token not in NAME_STOPWORDS and len(token) >= 4]


# def text_has_token(text, token):
#     return re.search(rf"\b{re.escape(token)}\b", text) is not None


# def file_controls(filename):
#     return source_controls[source_controls["source_file"].eq(filename)]


# def statement_for(filename, frames):
#     controls = file_controls(filename)
#     detail = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

#     broker = ""
#     statement_ref = ""
#     currency = ""
#     stated_total = None
#     stated_total_text = ""
#     total_file = filename
#     total_sheet = ""
#     total_row = ""

#     for row in controls.itertuples(index=False):
#         label = str(row.label).strip()
#         label_key = label.casefold()
#         value = str(row.value).strip()
#         if label_key in BROKER_LABELS and value:
#             broker = value
#         elif label_key in REF_LABELS and value and not statement_ref:
#             statement_ref = value
#         elif label_key in CCY_LABELS and value:
#             currency = currency_of(value) or value.upper()
#         elif label_key in TOTAL_LABELS and value:
#             stated_total = parse_amount(value)
#             stated_total_text = value
#             currency = currency or currency_of(value)
#             total_sheet = row.source_sheet
#             total_row = row.source_row
#         elif row.control_type == "statement_total" and stated_total is None:
#             stated_total = parse_amount(value)
#             stated_total_text = value
#             currency = currency or currency_of(value)
#             total_sheet = row.source_sheet
#             total_row = row.source_row
#         elif row.control_type == "cover_sheet" and not value and not broker:
#             broker = label

#     if not detail.empty:
#         columns = {column.casefold(): column for column in detail.columns}
#         if not broker and "broker" in columns:
#             broker = next(value for value in detail[columns["broker"]] if value)
#         if not statement_ref:
#             for key in ("statement id", "statement ref", "remittance id"):
#                 if key in columns and detail[columns[key]].ne("").any():
#                     statement_ref = next(value for value in detail[columns[key]] if value)
#                     break
#         if not currency and "currency" in columns:
#             codes = sorted({value for value in detail[columns["currency"]] if currency_of(value) or len(value) == 3})
#             if len(codes) == 1:
#                 currency = currency_of(codes[0]) or codes[0].upper()

#     refs = []
#     for ref in [statement_ref]:
#         if ref and ref not in refs:
#             refs.append(ref)
#     if not detail.empty:
#         for key in ("payment ref", "statement id", "statement ref", "remittance id"):
#             column = {name.casefold(): name for name in detail.columns}.get(key)
#             if column:
#                 for value in detail[column]:
#                     if value and value not in refs:
#                         refs.append(value)

#     return {
#         "broker": broker,
#         "statement_reference": statement_ref,
#         "match_refs": refs,
#         "name_tokens": name_tokens(broker),
#         "currency": currency or "GBP",
#         "statement_total": stated_total,
#         "statement_total_text": stated_total_text,
#         "source_file": total_file,
#         "source_sheet": total_sheet,
#         "source_row": total_row,
#     }


# frames_by_file = {}
# for dataset_name, frame in remittance_details.items():
#     filename = dataset_name.split(" | ", 1)[0]
#     frames_by_file.setdefault(filename, []).append(frame)

# statements = [statement_for(filename, frames) for filename, frames in frames_by_file.items()]
# statements = [row for row in statements if row["statement_reference"] or row["statement_total"] is not None]

# bank = bank_transactions.copy()
# bank["credit"] = bank["credit_amount"].map(parse_amount)
# bank["debit"] = bank["debit_amount"].map(parse_amount)
# credits = bank[bank["credit"].notna() & (bank["credit"] > 0)].copy()
# excluded_debits = bank[bank["debit"].notna() & (bank["debit"] > 0) & bank["credit"].isna()]


# def bank_blob(row):
#     return " ".join(str(getattr(row, column)) for column in ("counterparty_name", "bank_reference", "narrative")).upper()


# def matching_statements(row):
#     blob = bank_blob(row)
#     hits = []
#     for statement in statements:
#         ref_hits = [ref for ref in statement["match_refs"] if ref.upper() in blob]
#         token_hits = [token for token in statement["name_tokens"] if text_has_token(blob, token)]
#         if ref_hits or token_hits:
#             hits.append((statement, ref_hits, token_hits))
#     return hits


# assignments = []
# unresolved_credits = []
# for row in credits.itertuples(index=False):
#     hits = matching_statements(row)
#     receipt = {
#         "transaction_id": row.transaction_id,
#         "receipt_date": row.value_date,
#         "currency": row.currency,
#         "amount": row.credit,
#         "counterparty_name": row.counterparty_name,
#         "bank_reference": row.bank_reference,
#         "narrative": row.narrative,
#     }
#     if len(hits) == 1:
#         statement, ref_hits, token_hits = hits[0]
#         receipt["statement_reference"] = statement["statement_reference"]
#         receipt["link"] = "; ".join(
#             ([f"reference {', '.join(ref_hits)}"] if ref_hits else [])
#             + ([f"name {', '.join(token_hits)}"] if token_hits else [])
#         )
#         if row.currency != statement["currency"]:
#             receipt["link"] += f"; currency {row.currency} vs statement {statement['currency']}"
#         assignments.append(receipt)
#     else:
#         receipt["statement_reference"] = ""
#         if len(hits) > 1:
#             receipt["link"] = "more than one statement shares this name or reference; left unmatched"
#         else:
#             receipt["link"] = "no broker name or statement reference in the bank fields"
#         unresolved_credits.append(receipt)

# assigned = pd.DataFrame(assignments)
# results = []

# for statement in statements:
#     linked = (
#         assigned[assigned["statement_reference"].eq(statement["statement_reference"])]
#         if not assigned.empty
#         else assigned
#     )
#     received = float(linked["amount"].sum()) if not linked.empty else 0.0
#     total = statement["statement_total"]
#     variance = None if total is None else round(received - total, 2)
#     ids = ", ".join(linked["transaction_id"]) if not linked.empty else ""
#     dates = ", ".join(linked["receipt_date"]) if not linked.empty else ""
#     narratives = " | ".join(linked["narrative"]) if not linked.empty else ""
#     links = "; ".join(dict.fromkeys(linked["link"])) if not linked.empty else ""

#     if linked.empty:
#         status = "no identified bank receipt"
#         reason = "No bank credit names this broker or quotes this statement reference."
#     elif variance is not None and abs(variance) <= ROUNDING_TOLERANCE:
#         status = "matched"
#         reason = f"{links}. Combined receipts {received:,.2f} equal the statement total."
#     else:
#         status = "amount difference"
#         reason = (
#             f"{links}. Bank received {received:,.2f} against statement total {total:,.2f} "
#             f"(variance {variance:,.2f})."
#         )
#         if narratives:
#             reason += f" Bank narrative: {narratives}."

#     source = statement["source_file"]
#     if statement["source_sheet"]:
#         source = f"{source} | {statement['source_sheet']}"

#     results.append({
#         "broker": statement["broker"],
#         "statement_reference": statement["statement_reference"],
#         "bank_transaction_ids": ids,
#         "receipt_dates": dates,
#         "statement_total": total,
#         "currency": statement["currency"],
#         "bank_amount_received": round(received, 2),
#         "variance": variance,
#         "status": status,
#         "reason": reason,
#         "source_file_sheet": source,
#         "source_row": statement["source_row"],
#         "bank_narratives": narratives,
#     })

# for receipt in unresolved_credits:
#     results.append({
#         "broker": receipt["counterparty_name"],
#         "statement_reference": "",
#         "bank_transaction_ids": receipt["transaction_id"],
#         "receipt_dates": receipt["receipt_date"],
#         "statement_total": None,
#         "currency": receipt["currency"],
#         "bank_amount_received": round(receipt["amount"], 2),
#         "variance": None,
#         "status": "unmatched bank credit",
#         "reason": f"{receipt['link']}. Narrative: {receipt['narrative']}.",
#         "source_file_sheet": "bank_transactions_may_2026.csv",
#         "source_row": "",
#         "bank_narratives": receipt["narrative"],
#     })

# bank_to_remittance = pd.DataFrame(results)
# bank_receipt_lines = pd.DataFrame(assignments + unresolved_credits)

# pd.set_option("display.max_colwidth", 160)
# pd.set_option("display.max_rows", 50)
# # display(bank_to_remittance)
# # display(bank_receipt_lines)
# print(bank_to_remittance.to_string(index=False))
# print(bank_receipt_lines.to_string(index=False))
# print(
#     f"Rounding tolerance: {ROUNDING_TOLERANCE:.2f}. "
#     f"Debits excluded from premium receipts: "
#     + ", ".join(
#         f"{row.transaction_id} {row.currency} {row.debit:.2f} ({row.narrative})"
#         for row in excluded_debits.itertuples(index=False)
#     )
# )

# import re

# PENNY_TOLERANCE = 0.01

# TXN_GROUPS = {
#     "NB": "new business",
#     "NEW BUSINESS": "new business",
#     "CAN": "cancellation",
#     "CANCELLATION": "cancellation",
#     "MTA": "adjustment",
#     "END": "adjustment",
#     "ENDORSEMENT": "adjustment",
# }


# def parse_money(text):
#     if text is None or str(text).strip() == "":
#         return None
#     raw = str(text).strip()
#     negative = raw.startswith("(") and raw.endswith(")")
#     cleaned = re.sub(r"\b[A-Z]{3}\b", "", raw.strip("()"))
#     cleaned = cleaned.replace(",", "").strip()
#     if cleaned in {"", "-"}:
#         return None
#     value = float(cleaned)
#     return -value if negative else value


# def column(frame, *names):
#     lookup = {name.casefold(): name for name in frame.columns}
#     for name in names:
#         if name.casefold() in lookup:
#             return lookup[name.casefold()]
#     return None


# def first_value(frame, *names):
#     found = column(frame, *names)
#     if found is None:
#         return ""
#     for value in frame[found]:
#         if str(value).strip():
#             return str(value).strip()
#     return ""


# def broker_for(frame):
#     on_row = first_value(frame, "Broker")
#     if on_row:
#         return on_row
#     filename = frame["source_file"].iloc[0]
#     controls = source_controls[source_controls["source_file"].eq(filename)]
#     labelled = controls[controls["label"].str.casefold().eq("broker")]
#     if not labelled.empty and str(labelled.iloc[0]["value"]).strip():
#         return str(labelled.iloc[0]["value"]).strip()
#     cover = controls[controls["control_type"].eq("cover_sheet") & controls["value"].eq("")]
#     if not cover.empty:
#         return str(cover.iloc[0]["label"]).strip()
#     return ""


# def broker_key(name):
#     text = re.sub(r"[^A-Z0-9 ]", " ", str(name).upper())
#     text = re.sub(r"\b(LTD|LIMITED|PLC|SA|LLP)\b", " ", text)
#     return re.sub(r"\s+", " ", text).strip()


# def policy_key(reference):
#     """Comparable policy key. Does not replace the reference stored on the row."""
#     text = str(reference).upper().strip()
#     if not text:
#         return ""
#     text = re.sub(r"[-/\s]*(CAN|ENDT\d*|MTA)$", "", text)
#     text = re.sub(r"\s+", "", text)
#     year_only = re.fullmatch(r"(\d{4})[/-](\d+)", text)
#     if year_only:
#         text = f"POL-{year_only.group(1)}-{int(year_only.group(2)):04d}"
#     short_pol = re.fullmatch(r"POL(\d{1,4})", text)
#     if short_pol:
#         text = f"POL-2026-{int(short_pol.group(1)):04d}"
#     return re.sub(r"[^A-Z0-9]", "", text)


# def name_key(name):
#     text = str(name).upper().replace("&", " AND ")
#     text = re.sub(r"\b(LTD|LIMITED|PLC|SA|LLP|GROUP)\b", " ", text)
#     return re.sub(r"[^A-Z0-9]", "", text)


# def txn_group(label):
#     return TXN_GROUPS.get(str(label).upper().strip(), "")


# def money_fields(frame):
#     if column(frame, "GBP Settlement Amount"):
#         return column(frame, "GBP Settlement Amount"), "GBP"
#     for name in ("Net Settlement", "Amount Due MGA", "Amount Due", "Net Due"):
#         found = column(frame, name)
#         if found:
#             return found, ""
#     return None, ""


# lines = []
# for dataset_name, frame in remittance_details.items():
#     amount_column, forced_currency = money_fields(frame)
#     policy_column = column(frame, "Policy No", "Policy Reference", "UW Ref", "Policy", "Policy / Risk Ref")
#     insured_column = column(frame, "Insured", "Insured Name", "Client", "Assured")
#     txn_column = column(frame, "Transaction Type", "Txn Type")
#     currency_column = column(frame, "Currency")
#     broker = broker_for(frame)
#     for row in frame.to_dict("records"):
#         amount_text = row.get(amount_column, "") if amount_column else ""
#         currency = forced_currency or (row.get(currency_column, "") if currency_column else "") or "GBP"
#         policy_reference = row.get(policy_column, "") if policy_column else ""
#         lines.append({
#             "broker": broker,
#             "broker_key": broker_key(broker),
#             "source_file": row["source_file"],
#             "source_sheet": row["source_sheet"],
#             "source_row": row["source_row"],
#             "policy_reference": policy_reference,
#             "policy_key": policy_key(policy_reference),
#             "insured_name": row.get(insured_column, "") if insured_column else "",
#             "txn_type": row.get(txn_column, "") if txn_column else "",
#             "amount": parse_money(amount_text),
#             "currency": currency,
#         })

# bdx_rows = []
# for row in premium_bdx.itertuples(index=False):
#     bdx_rows.append({
#         "bdx_record_id": row.bdx_record_id,
#         "broker": row.broker_name,
#         "broker_key": broker_key(row.broker_name),
#         "policy_reference": row.policy_reference,
#         "policy_key": policy_key(row.policy_reference),
#         "insured_name": row.insured_name,
#         "txn_type": row.transaction_type,
#         "amount": parse_money(row.settlement_amount_gbp),
#         "currency": row.settlement_currency,
#         "used": False,
#     })


# def describe_checks(line, match):
#     notes = []
#     if name_key(line["insured_name"]) == name_key(match["insured_name"]):
#         notes.append("insured name agrees")
#     else:
#         notes.append(f"insured name differs ({line['insured_name']} vs {match['insured_name']})")
#     if line["txn_type"] and match["txn_type"] and line["txn_type"].upper() != match["txn_type"].upper():
#         if txn_group(line["txn_type"]) and txn_group(line["txn_type"]) == txn_group(match["txn_type"]):
#             notes.append(f"transaction label {line['txn_type']} vs {match['txn_type']}")
#         else:
#             notes.append(f"transaction type {line['txn_type']} vs {match['txn_type']}")
#     if line["currency"] != match["currency"]:
#         notes.append(f"currency {line['currency']} vs {match['currency']}")
#     return "; ".join(notes)


# results = []
# seen_policy = set()
# for line in lines:
#     same_broker = [row for row in bdx_rows if row["broker_key"] == line["broker_key"]]
#     policy_hits = [row for row in same_broker if line["policy_key"] and row["policy_key"] == line["policy_key"]]
#     status = ""
#     reason = ""
#     match = None

#     duplicate = bool(line["policy_key"]) and (line["broker_key"], line["policy_key"]) in seen_policy
#     if line["policy_key"]:
#         seen_policy.add((line["broker_key"], line["policy_key"]))

#     if duplicate:
#         status = "duplicate remittance line"
#         reason = "This policy reference is repeated on the remittance. The BDX record was not used a second time."
#     elif len(policy_hits) == 1:
#         match = policy_hits[0]
#         if match["used"]:
#             status = "duplicate remittance line"
#             reason = f"{match['bdx_record_id']} is already linked to another remittance line."
#             match = None
#         else:
#             match["used"] = True
#     elif len(policy_hits) > 1:
#         status = "unresolved"
#         ids = ", ".join(row["bdx_record_id"] for row in policy_hits)
#         reason = f"More than one BDX record shares this policy reference ({ids})."
#     elif not line["policy_key"]:
#         name_hits = [
#             row for row in same_broker
#             if not row["used"]
#             and name_key(line["insured_name"])
#             and name_key(row["insured_name"]) == name_key(line["insured_name"])
#         ]
#         loose_hits = [
#             row for row in same_broker
#             if not row["used"]
#             and name_key(line["insured_name"])
#             and name_key(line["insured_name"]) in name_key(row["insured_name"])
#         ]
#         candidates = name_hits or loose_hits
#         status = "unresolved"
#         if len(candidates) > 1:
#             ids = ", ".join(
#                 f"{row['bdx_record_id']} {row['policy_reference']} {row['insured_name']}" for row in candidates
#             )
#             reason = f"No policy reference. More than one BDX record could fit ({ids})."
#         elif len(candidates) == 1:
#             reason = (
#                 f"No policy reference. Only {candidates[0]['bdx_record_id']} has a similar insured name, "
#                 "which is not enough to assign it."
#             )
#         else:
#             reason = "No policy reference and no single BDX insured name to compare."
#     else:
#         status = "unresolved"
#         reason = f"No BDX policy reference matches {line['policy_reference']} for this broker."

#     variance = None
#     if match is not None:
#         variance = (
#             None if line["amount"] is None or match["amount"] is None
#             else round(line["amount"] - match["amount"], 2)
#         )
#         checks = describe_checks(line, match)
#         if variance is not None and abs(variance) <= PENNY_TOLERANCE and line["currency"] == match["currency"]:
#             status = "matched"
#             reason = f"Policy reference matches {match['bdx_record_id']}. {checks}."
#         else:
#             status = "amount difference"
#             reason = (
#                 f"Policy reference matches {match['bdx_record_id']}. {checks}. "
#                 f"Remittance {line['amount']:,.2f} vs BDX {match['amount']:,.2f} ({line['currency']})."
#             )

#     results.append({
#         "broker": line["broker"],
#         "source_file": line["source_file"],
#         "source_sheet": line["source_sheet"],
#         "source_row": line["source_row"],
#         "remittance_policy_reference": line["policy_reference"],
#         "insured_name": line["insured_name"],
#         "bdx_record_id": match["bdx_record_id"] if match else "",
#         "bdx_policy_reference": match["policy_reference"] if match else "",
#         "remittance_amount": line["amount"],
#         "bdx_amount": match["amount"] if match else None,
#         "currency": line["currency"],
#         "amount_variance": variance,
#         "status": status,
#         "reason": reason,
#     })

# remittance_to_bdx = pd.DataFrame(results)
# bdx_without_remittance = pd.DataFrame([
#     {
#         "bdx_record_id": row["bdx_record_id"],
#         "broker": row["broker"],
#         "bdx_policy_reference": row["policy_reference"],
#         "insured_name": row["insured_name"],
#         "transaction_type": row["txn_type"],
#         "settlement_amount": row["amount"],
#         "currency": row["currency"],
#     }
#     for row in bdx_rows
#     if not row["used"]
# ])

# pd.set_option("display.max_colwidth", 140)
# pd.set_option("display.max_rows", 40)
# # display(remittance_to_bdx)
# print(f"Rounding tolerance: {PENNY_TOLERANCE:.2f}. Variance is remittance amount minus BDX settlement amount.")
# print("BDX records with no remittance match:")
# # display(bdx_without_remittance)
# print(remittance_to_bdx.to_string(index=False))
# print(bdx_without_remittance.to_string(index=False))