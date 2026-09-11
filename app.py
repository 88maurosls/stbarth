import io
import os
import re
from collections import Counter

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(page_title="St.Barth PDF to Excel", layout="wide")
st.title("St.Barth PDF to Excel")
st.write("v1.0 by MM")


STATIC_COLS = ["CODICE", "COLORE", "DESCRIZIONE", "PREZZO WHS", "PREZZO RTL"]

ALPHA_SIZE_ORDER = {
    "XXS": 0,
    "XS": 1,
    "S": 2,
    "M": 3,
    "L": 4,
    "XL": 5,
    "XXL": 6,
    "3XL": 7,
    "4XL": 8,
    "5XL": 9,
    "XS-S": 10,
    "S-M": 11,
    "M-L": 12,
    "L-XL": 13,
    "XL-XXL": 14,
}

SKU_RE = re.compile(r"^[A-Z0-9]+-[A-Z0-9]+$")
MONEY_RE = re.compile(r"^\d{1,4}(?:\.\d{3})*,\d{2}$")


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_size(size: str) -> str:
    size = normalize_text(size).upper()

    if size in {"0", "ONE SIZE", "ONESIZE"}:
        return "UNICA"

    size = size.replace("–", "-").replace("—", "-")
    size = size.replace("/", "-")
    size = re.sub(r"\s*-\s*", "-", size)

    return size


def parse_money(value: str):
    value = normalize_text(value)
    if not value:
        return None
    try:
        return float(value.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def group_words_into_lines(words, y_tol=1.8):
    """Raggruppa le parole in righe usando la coordinata verticale."""
    rows = []

    for word in sorted(words, key=lambda w: (w["top"], w["x0"])):
        chosen = None

        # Basta cercare tra le ultime righe create: le parole sono già ordinate per Y.
        for row in reversed(rows[-8:]):
            if abs(row["top"] - word["top"]) <= y_tol:
                chosen = row
                break

        if chosen is None:
            chosen = {"top": word["top"], "words": []}
            rows.append(chosen)

        chosen["words"].append(word)

    normalized = []
    for row in rows:
        row_words = sorted(row["words"], key=lambda w: w["x0"])
        normalized.append(
            {
                "top": row["top"],
                "words": row_words,
                "text": normalize_text(" ".join(w["text"] for w in row_words)),
            }
        )

    return sorted(normalized, key=lambda row: row["top"])


def find_product_headers(words):
    """
    Individua le righe prodotto del formato St.Barth, ad esempio:
    40 HER0001-03664M - HERON
    """
    headers = []
    lines = group_words_into_lines(words)

    for line in lines:
        sku_words = [
            w
            for w in line["words"]
            if 44 <= w["x0"] <= 130 and SKU_RE.fullmatch(w["text"])
        ]
        if not sku_words:
            continue

        index_words = [
            w
            for w in line["words"]
            if w["x0"] < 44 and re.fullmatch(r"\d+", w["text"])
        ]
        if not index_words:
            continue

        sku_word = sku_words[0]
        headers.append(
            {
                "top": sku_word["top"],
                "sku": sku_word["text"],
                "index": int(index_words[-1]["text"]),
            }
        )

    return headers


def extract_header_description(block_words, header_top, sku):
    sku_matches = [
        w
        for w in block_words
        if w["text"] == sku and abs(w["top"] - header_top) <= 1.8
    ]
    if not sku_matches:
        return ""

    sku_word = sku_matches[0]
    desc_words = [
        w
        for w in block_words
        if abs(w["top"] - header_top) <= 1.8
        and sku_word["x1"] < w["x0"] < 180
        and w["text"] != "-"
    ]

    return normalize_text(" ".join(w["text"] for w in sorted(desc_words, key=lambda x: x["x0"])))


def extract_variant_detail(block_words, header_top):
    """
    Nel layout St.Barth il dettaglio grafica/colore è normalmente l'ultima riga
    della colonna sinistra del blocco, ad es. 'APRES SKI TARTAN DRAW 1845 EMB'.
    """
    left_words = [
        w
        for w in block_words
        if 44 <= w["x0"] < 180 and w["top"] > header_top + 3
    ]
    lines = group_words_into_lines(left_words)
    texts = [line["text"] for line in lines if line["text"]]
    return texts[-1] if texts else ""


def extract_sizes(block_words, header_top):
    """Restituisce una lista di tuple (taglia, posizione_x)."""
    size_words = [
        w
        for w in block_words
        if header_top - 2.2 <= w["top"] <= header_top + 1.8
        and 180 <= w["x0"] < 400
    ]
    size_words = sorted(size_words, key=lambda w: w["x0"])

    upper_tokens = [w["text"].upper() for w in size_words]
    if "ONE" in upper_tokens and "SIZE" in upper_tokens:
        one_word = next(w for w in size_words if w["text"].upper() == "ONE")
        return [("UNICA", one_word["x0"] + 1)]

    sizes = []
    for word in size_words:
        token = word["text"].upper()
        if re.fullmatch(
            r"(?:XXS|XS|S|M|L|XL|XXL|3XL|4XL|5XL|\d{1,3}(?:/\d{1,3})?)",
            token,
        ):
            sizes.append((normalize_size(token), word["x0"]))

    return sizes


def extract_quantities(block_words, header_top, sizes):
    """
    Legge la riga quantità sotto le taglie e associa ogni numero alla taglia
    più vicina orizzontalmente.
    """
    result = {size: 0 for size, _ in sizes}
    if not sizes:
        return result

    qty_words = [
        w
        for w in block_words
        if header_top + 4 <= w["top"] <= header_top + 11
        and 180 <= w["x0"] < 400
        and re.fullmatch(r"\d+", w["text"])
    ]

    for word in qty_words:
        qty = int(word["text"])
        nearest_size, nearest_x = min(sizes, key=lambda item: abs(item[1] - word["x0"]))

        if abs(nearest_x - word["x0"]) <= 15:
            result[nearest_size] += qty

    return result


def extract_wholesale_price(block_words):
    """Prezzo unitario WHS, presente sotto le quantità nella griglia taglie."""
    prices = [
        w
        for w in block_words
        if 180 <= w["x0"] < 450 and MONEY_RE.fullmatch(w["text"])
    ]
    if not prices:
        return ""

    # Lo stesso prezzo può apparire più volte per le varie taglie.
    # Usiamo il valore più frequente, con fallback al primo.
    values = [w["text"] for w in prices]
    counts = Counter(values)
    return counts.most_common(1)[0][0]


def extract_retail_price(block_words):
    """Prezzo Listino SellOUT, nella colonna destra del blocco."""
    prices = sorted(
        [
            w
            for w in block_words
            if w["x0"] > 530 and MONEY_RE.fullmatch(w["text"])
        ],
        key=lambda w: w["top"],
    )
    if not prices:
        return ""

    # Il primo è normalmente l'importo totale della riga, l'ultimo il SellOUT.
    return prices[-1]["text"]


def extract_row_total_qty(block_words, header_top):
    qty_words = [
        w
        for w in block_words
        if 495 <= w["x0"] < 530
        and header_top - 1 <= w["top"] <= header_top + 5
        and re.fullmatch(r"\d+", w["text"])
    ]
    return int(qty_words[0]["text"]) if qty_words else None


def parse_product_block(block_words, header):
    sku = header["sku"]
    header_top = header["top"]

    if "-" not in sku:
        return None

    codice, colore = sku.split("-", 1)

    header_description = extract_header_description(block_words, header_top, sku)
    variant_detail = extract_variant_detail(block_words, header_top)

    descrizione = header_description
    if variant_detail and variant_detail.upper() != header_description.upper():
        descrizione = normalize_text(f"{header_description} {variant_detail}")

    sizes = extract_sizes(block_words, header_top)
    mapped_qty = extract_quantities(block_words, header_top, sizes)

    record = {
        "CODICE": codice,
        "COLORE": colore,
        "DESCRIZIONE": descrizione,
        "PREZZO WHS": extract_wholesale_price(block_words),
        "PREZZO RTL": extract_retail_price(block_words),
    }

    for size, qty in mapped_qty.items():
        if qty:
            record[size] = qty

    expected_total = extract_row_total_qty(block_words, header_top)
    extracted_total = sum(mapped_qty.values())

    diagnostics = {
        "row_number": header["index"],
        "sku": sku,
        "expected_qty": expected_total,
        "extracted_qty": extracted_total,
    }

    return record, diagnostics


def parse_pdf(file_obj):
    records = []
    diagnostics = []

    file_obj.seek(0)
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False,
                use_text_flow=False,
            )

            headers = find_product_headers(words)
            if not headers:
                continue

            for i, header in enumerate(headers):
                block_top = header["top"] - 2.5
                block_bottom = (
                    headers[i + 1]["top"] - 2.5
                    if i + 1 < len(headers)
                    else page.height
                )

                block_words = [
                    w for w in words if block_top <= w["top"] < block_bottom
                ]

                parsed = parse_product_block(block_words, header)
                if parsed:
                    record, info = parsed
                    records.append(record)
                    diagnostics.append(info)

    file_obj.seek(0)
    return records, diagnostics


def extract_order_number(file_obj):
    """Estrae N°Ordine, es. 170.843."""
    file_obj.seek(0)

    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages[:2]:
            text = normalize_text(page.extract_text() or "")
            match = re.search(r"N[°º]?Ordine\s+([\d.]+)", text, re.IGNORECASE)
            if match:
                file_obj.seek(0)
                return match.group(1).strip()

    file_obj.seek(0)
    return ""


def extract_document_summary(file_obj):
    """
    Cerca il riepilogo finale del documento.
    Nel PDF di esempio: 19.259,10 EUR - Qta 496.
    """
    file_obj.seek(0)

    with pdfplumber.open(file_obj) as pdf:
        all_text = " ".join(page.extract_text() or "" for page in pdf.pages)

    text = normalize_text(all_text)

    match = re.search(
        r"Totale\s+Merce.*?(\d{1,3}(?:\.\d{3})*,\d{2})\s+EUR\s+(\d+)",
        text,
        re.IGNORECASE,
    )

    file_obj.seek(0)

    if not match:
        return None, None

    total_value = parse_money(match.group(1))
    total_qty = int(match.group(2))
    return total_qty, total_value


def size_sort_key(value: str):
    value = normalize_size(value)

    if value == "UNICA":
        return (0, 0)

    if re.fullmatch(r"\d+", value):
        return (1, int(value))

    # Range numerici, es. 36-39
    range_match = re.fullmatch(r"(\d+)-(\d+)", value)
    if range_match:
        return (2, int(range_match.group(1)), int(range_match.group(2)))

    if value in ALPHA_SIZE_ORDER:
        return (3, ALPHA_SIZE_ORDER[value])

    return (4, value)


def build_dataframe(all_records):
    if not all_records:
        return pd.DataFrame()

    size_cols = set()
    for record in all_records:
        for key in record.keys():
            if key not in STATIC_COLS:
                size_cols.add(key)

    ordered_size_cols = sorted(size_cols, key=size_sort_key)

    rows = []
    for record in all_records:
        row = {col: record.get(col, "") for col in STATIC_COLS}
        for size_col in ordered_size_cols:
            row[size_col] = record.get(size_col, "")
        rows.append(row)

    return pd.DataFrame(rows)


def calculate_total_qty(df: pd.DataFrame) -> int:
    if df.empty:
        return 0

    total_qty = 0
    for col in df.columns:
        if col not in STATIC_COLS:
            total_qty += pd.to_numeric(df[col], errors="coerce").fillna(0).sum()

    return int(total_qty)


def calculate_wholesale_total(df: pd.DataFrame) -> float:
    if df.empty:
        return 0.0

    total = 0.0
    size_cols = [col for col in df.columns if col not in STATIC_COLS]

    for _, row in df.iterrows():
        price = parse_money(row.get("PREZZO WHS", ""))
        if price is None:
            continue

        qty = 0
        for col in size_cols:
            value = pd.to_numeric(row.get(col, 0), errors="coerce")
            if pd.notna(value):
                qty += value

        total += price * qty

    return round(float(total), 2)


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ORDINE")
        ws = writer.book["ORDINE"]

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter

            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))

            ws.column_dimensions[col_letter].width = min(max_len + 2, 65)

    output.seek(0)
    return output.getvalue()


def build_output_filename(uploaded_files, order_number):
    if len(uploaded_files) == 1:
        base = os.path.splitext(uploaded_files[0].name)[0]
        if order_number:
            clean_order = order_number.replace(".", "")
            return f"{base}_{clean_order}.xlsx"
        return f"{base}.xlsx"

    if order_number:
        clean_order = order_number.replace(".", "")
        return f"STBARTH_export_{clean_order}.xlsx"

    return "STBARTH_export.xlsx"


uploaded_files = st.file_uploader(
    "Carica PDF St.Barth",
    type=["pdf"],
    accept_multiple_files=True,
)


if uploaded_files:
    all_records = []
    all_diagnostics = []
    first_order_number = ""
    expected_total_qty = None
    expected_total_value = None

    with st.spinner("Sto leggendo il PDF St.Barth e creando l'Excel..."):
        for uploaded_file in uploaded_files:
            try:
                if not first_order_number:
                    first_order_number = extract_order_number(uploaded_file)

                if len(uploaded_files) == 1:
                    expected_total_qty, expected_total_value = extract_document_summary(uploaded_file)

                uploaded_file.seek(0)
                records, diagnostics = parse_pdf(uploaded_file)
                all_records.extend(records)
                all_diagnostics.extend(diagnostics)

            except Exception as exc:
                st.error(f"Errore su {uploaded_file.name}: {exc}")

    df = build_dataframe(all_records)

    if df.empty:
        st.warning("Non sono riuscito a trovare prodotti nel PDF.")
    else:
        total_qty = calculate_total_qty(df)
        wholesale_total = calculate_wholesale_total(df)
        output_filename = build_output_filename(uploaded_files, first_order_number)

        bad_rows = [
            row
            for row in all_diagnostics
            if row["expected_qty"] is not None
            and row["expected_qty"] != row["extracted_qty"]
        ]

        col1, col2, col3 = st.columns(3)
        col1.metric("Prodotti estratti", len(df))
        col2.metric("Quantità estratta", total_qty)
        col3.metric("Totale WHS estratto", f"€ {wholesale_total:,.2f}")

        if first_order_number:
            st.info(f"N° ordine rilevato: {first_order_number}")

        if expected_total_qty is not None:
            if total_qty == expected_total_qty:
                st.success(
                    f"Controllo quantità OK: il PDF riporta {expected_total_qty} pezzi "
                    "e l'estrazione coincide."
                )
            else:
                st.error(
                    f"ATTENZIONE: il PDF riporta {expected_total_qty} pezzi, "
                    f"ma ne sono stati estratti {total_qty}."
                )

        if expected_total_value is not None:
            if abs(wholesale_total - expected_total_value) < 0.01:
                st.success(
                    f"Controllo valore WHS OK: € {expected_total_value:,.2f}."
                )
            else:
                st.warning(
                    f"Il totale WHS del PDF è € {expected_total_value:,.2f}, "
                    f"mentre l'estrazione calcola € {wholesale_total:,.2f}."
                )

        if bad_rows:
            st.warning(
                f"Ci sono {len(bad_rows)} righe in cui la somma taglie non coincide "
                "con la quantità totale del PDF."
            )
            st.dataframe(pd.DataFrame(bad_rows), use_container_width=True)

        st.dataframe(df, use_container_width=True)

        excel_bytes = dataframe_to_excel_bytes(df)

        st.download_button(
            label="Scarica Excel",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
