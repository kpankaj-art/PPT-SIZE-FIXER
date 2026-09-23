import streamlit as st
import pandas as pd
import re
import io

from pptx import Presentation


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V6.1",
    page_icon="📐",
    layout="wide"
)


# =========================================================
# TITLE
# =========================================================

st.title("📐 PPT SIZE FIXER")

st.caption(
    "Excel Width × Height ko PPT ke existing Size field mein replace karega."
)


# =========================================================
# EXCEL COLUMN ALIASES
# =========================================================

WIDTH_ALIASES = [
    "width",
    "width inches",
    "width inch",
    "width (inches)",
    "width (inch)",
    "w",
    "w inches",
    "w inch",
    "w (inches)",
    "w (inch)",
    "board width",
    "size width"
]

HEIGHT_ALIASES = [
    "height",
    "height inches",
    "height inch",
    "height (inches)",
    "height (inch)",
    "h",
    "h inches",
    "h inch",
    "h (inches)",
    "h (inch)",
    "board height",
    "size height"
]


# =========================================================
# PROTECTED WORDS
# =========================================================

PROTECTED_WORDS = {
    "qty",
    "quantity",
    "media",
    "media type",
    "remarks",
    "remark",
    "outlet",
    "outlet name",
    "address",
    "contact",
    "contact no",
    "contact number",
    "mobile",
    "phone",
    "sap",
    "sap code",
    "outlet code",
    "district"
}


# =========================================================
# SESSION STATE
# =========================================================

if "fixed_ppt_bytes" not in st.session_state:
    st.session_state.fixed_ppt_bytes = None

if "result_df" not in st.session_state:
    st.session_state.result_df = None

if "output_filename" not in st.session_state:
    st.session_state.output_filename = "PPT_SIZE_FIXED_V6.1.pptx"


# =========================================================
# TEXT NORMALIZATION
# =========================================================

def normalize_text(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\n", " ")
    value = value.replace("\r", " ")
    value = value.replace("\t", " ")

    value = value.strip().lower()

    value = re.sub(r"\s+", " ", value)

    return value


def normalize_column_name(value):

    value = normalize_text(value)

    value = value.replace("_", " ")
    value = value.replace("-", " ")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


# =========================================================
# NUMBER CLEANING
# =========================================================

def clean_number(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    if not text:
        return ""

    text = text.replace(",", "")

    try:

        number = float(text)

        if number.is_integer():
            return str(int(number))

        return str(number).rstrip("0").rstrip(".")

    except Exception:

        return text


# =========================================================
# PROTECTED FIELD
# =========================================================

def is_protected_field(text):

    n = normalize_text(text)

    if not n:
        return False

    if n in PROTECTED_WORDS:
        return True

    for word in PROTECTED_WORDS:

        if n.startswith(word + ":"):
            return True

        if n.startswith(word + " :"):
            return True

        if n.startswith(word + " :-"):
            return True

        if n.startswith(word + " -"):
            return True

    return False


# =========================================================
# SHAPE TEXT
# =========================================================

def get_shape_text(shape):

    try:

        if not shape.has_text_frame:
            return ""

        return shape.text or ""

    except Exception:

        return ""


# =========================================================
# SHAPE GEOMETRY
# =========================================================

def get_geometry(shape):

    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height

    return {
        "left": left,
        "top": top,
        "right": left + width,
        "bottom": top + height,
        "width": width,
        "height": height,
        "cx": left + width / 2,
        "cy": top + height / 2
    }


# =========================================================
# COLLECT TEXT SHAPES
# =========================================================

def collect_text_shapes(slide):

    items = []

    for index, shape in enumerate(slide.shapes):

        text = get_shape_text(shape)

        if not text.strip():
            continue

        items.append({
            "index": index,
            "shape": shape,
            "text": text,
            "norm": normalize_text(text),
            "geo": get_geometry(shape)
        })

    return items


# =========================================================
# SIZE PATTERN
# =========================================================

SIZE_PATTERN = re.compile(
    r"\d+(?:\.\d+)?\s*[x×*]\s*\d+(?:\.\d+)?",
    re.IGNORECASE
)


# =========================================================
# SIZE VALUE
# =========================================================

def contains_size_value(text):

    if not text:
        return False

    return bool(
        SIZE_PATTERN.search(str(text))
    )


def is_standalone_size(text):

    if not text:
        return False

    raw = str(text).strip()

    pattern = (
        r"^\d+(?:\.\d+)?"
        r"\s*[x×*]\s*"
        r"\d+(?:\.\d+)?$"
    )

    return bool(
        re.match(
            pattern,
            raw,
            re.IGNORECASE
        )
    )


# =========================================================
# X SHAPE
# =========================================================

def is_x_shape(text):

    if not text:
        return False

    return normalize_text(text) in {
        "x",
        "×",
        "*"
    }


# =========================================================
# NUMBER
# =========================================================

def is_number(text):

    if not text:
        return False

    return bool(
        re.match(
            r"^\d+(?:\.\d+)?$",
            str(text).strip()
        )
    )


# =========================================================
# SIZE LABEL
# =========================================================

def find_size_labels(items):

    labels = []

    for item in items:

        n = item["norm"]

        if n in {
            "size",
            "size:",
            "size :",
            "size :-",
            "size -",
            "size=",
            "size ="
        }:

            labels.append(item)
            continue

        if re.match(
            r"^size\s*[:=\-]",
            n
        ):

            if not contains_size_value(
                item["text"]
            ):

                labels.append(item)

    return labels


# =========================================================
# REPLACE INLINE SIZE
# =========================================================

def replace_inline_size(
    text,
    width,
    height
):

    if not text:
        return text, False

    raw = str(text)

    if not contains_size_value(raw):
        return raw, False

    new_text, count = SIZE_PATTERN.subn(
        f"{width} X {height}",
        raw,
        count=1
    )

    return new_text, count > 0


# =========================================================
# SET EXISTING TEXT
# =========================================================

def set_shape_text(
    shape,
    new_text
):

    try:

        if not shape.has_text_frame:
            return False

        tf = shape.text_frame

        if tf.paragraphs:

            paragraph = tf.paragraphs[0]

            if paragraph.runs:

                paragraph.runs[0].text = str(
                    new_text
                )

                for run in paragraph.runs[1:]:
                    run.text = ""

                for p in tf.paragraphs[1:]:

                    for run in p.runs:
                        run.text = ""

                return True

        tf.text = str(new_text)

        return True

    except Exception:

        try:
            shape.text = str(new_text)
            return True

        except Exception:
            return False


# =========================================================
# FIND X NEAR SIZE LABEL
# =========================================================

def find_x_near_label(
    items,
    label
):

    candidates = []

    label_geo = label["geo"]

    for item in items:

        if item is label:
            continue

        if not is_x_shape(
            item["text"]
        ):
            continue

        geo = item["geo"]

        if abs(
            geo["cy"] -
            label_geo["cy"]
        ) > Inches(0.55):

            continue

        if abs(
            geo["cx"] -
            label_geo["cx"]
        ) > Inches(2.5):

            continue

        candidates.append(item)

    candidates.sort(
        key=lambda x: abs(
            x["geo"]["cx"] -
            label_geo["cx"]
        )
    )

    return candidates


# =========================================================
# FIND NUMBERS AROUND X
# =========================================================

def find_numbers_around_x(
    items,
    x_item,
    label
):

    x_geo = x_item["geo"]

    left_candidates = []
    right_candidates = []

    for item in items:

        if item is x_item:
            continue

        if item is label:
            continue

        if not is_number(
            item["text"]
        ):
            continue

        if is_protected_field(
            item["text"]
        ):
            continue

        geo = item["geo"]

        if abs(
            geo["cy"] -
            x_geo["cy"]
        ) > Inches(0.40):

            continue

        if abs(
            geo["cx"] -
            x_geo["cx"]
        ) > Inches(2.0):

            continue

        if geo["cx"] < x_geo["cx"]:

            left_candidates.append(item)

        elif geo["cx"] > x_geo["cx"]:

            right_candidates.append(item)

    left_candidates.sort(
        key=lambda item:
        abs(
            item["geo"]["cx"] -
            x_geo["cx"]
        )
    )

    right_candidates.sort(
        key=lambda item:
        abs(
            item["geo"]["cx"] -
            x_geo["cx"]
        )
    )

    if not left_candidates:
        return None

    if not right_candidates:
        return None

    return {
        "width": left_candidates[0],
        "x": x_item,
        "height": right_candidates[0]
    }


# =========================================================
# FIND SEPARATE SIZE STRUCTURE
# =========================================================

def find_separate_size_structure(
    items,
    labels
):

    candidates = []

    for label in labels:

        x_candidates = find_x_near_label(
            items,
            label
        )

        for x_item in x_candidates:

            components = find_numbers_around_x(
                items,
                x_item,
                label
            )

            if components is None:
                continue

            score = 0

            score += max(
                0,
                100 -
                int(
                    abs(
                        x_item["geo"]["cx"] -
                        label["geo"]["cx"]
                    )
                    /
                    Inches(0.05)
                )
            )

            candidates.append({
                "label": label,
                "components": components,
                "score": score
            })

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return candidates[0]


# =========================================================
# FIND EXISTING STANDALONE SIZE
# =========================================================

def find_standalone_size_box(
    items,
    labels
):

    candidates = []

    for item in items:

        if not is_standalone_size(
            item["text"]
        ):
            continue

        geo = item["geo"]

        for label in labels:

            label_geo = label["geo"]

            vertical = abs(
                geo["cy"] -
                label_geo["cy"]
            )

            horizontal = abs(
                geo["cx"] -
                label_geo["cx"]
            )

            if vertical > Inches(0.60):
                continue

            if horizontal > Inches(4.0):
                continue

            candidates.append({
                "item": item,
                "label": label,
                "distance": (
                    vertical +
                    horizontal
                )
            })

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x["distance"]
    )

    return candidates[0]


# =========================================================
# DELETE DUPLICATE SIZE BOX
# =========================================================

def delete_shape(
    shape
):

    try:

        element = shape._element

        parent = element.getparent()

        parent.remove(element)

        return True

    except Exception:

        return False


# =========================================================
# PROCESS ONE SLIDE
# =========================================================

def process_slide(
    slide,
    width,
    height
):

    items = collect_text_shapes(
        slide
    )

    # -----------------------------------------------------
    # 1. INLINE SIZE
    # -----------------------------------------------------

    for item in items:

        text = item["text"]

        if not text:
            continue

        if not item["norm"].startswith("size"):
            continue

        if not contains_size_value(
            text
        ):
            continue

        old_text = text

        new_text, changed = replace_inline_size(
            old_text,
            width,
            height
        )

        if changed:

            success = set_shape_text(
                item["shape"],
                new_text
            )

            if success:

                return {
                    "status": "Updated Inline Size",
                    "action": (
                        f"{old_text} -> {new_text}"
                    )
                }

    # -----------------------------------------------------
    # 2. SIZE LABEL
    # -----------------------------------------------------

    labels = find_size_labels(
        items
    )

    # -----------------------------------------------------
    # 3. SEPARATE WIDTH X HEIGHT
    # -----------------------------------------------------

    structure = find_separate_size_structure(
        items,
        labels
    )

    if structure is not None:

        components = structure[
            "components"
        ]

        old_width = components[
            "width"
        ]["text"]

        old_height = components[
            "height"
        ]["text"]

        ok1 = set_shape_text(
            components["width"]["shape"],
            str(width)
        )

        ok2 = set_shape_text(
            components["x"]["shape"],
            "X"
        )

        ok3 = set_shape_text(
            components["height"]["shape"],
            str(height)
        )

        # -----------------------------------------------
        # Remove duplicate combined Size box
        # -----------------------------------------------

        deleted = 0

        for item in items:

            shape = item["shape"]

            if shape in {
                components["width"]["shape"],
                components["x"]["shape"],
                components["height"]["shape"]
            }:
                continue

            if not is_standalone_size(
                item["text"]
            ):
                continue

            # Only remove if it is near same Size area
            label_geo = structure[
                "label"
            ]["geo"]

            item_geo = item["geo"]

            if abs(
                item_geo["cy"] -
                label_geo["cy"]
            ) > Inches(0.55):

                continue

            if abs(
                item_geo["cx"] -
                label_geo["cx"]
            ) > Inches(3.0):

                continue

            if delete_shape(
                shape
            ):

                deleted += 1

        if ok1 and ok2 and ok3:

            action = (
                f"{old_width} X {old_height}"
                f" -> "
                f"{width} X {height}"
            )

            if deleted:
                action += (
                    f" | Deleted duplicate box: {deleted}"
                )

            return {
                "status": "Updated Existing W-X-H Fields",
                "action": action
            }

    # -----------------------------------------------------
    # 4. EXISTING COMBINED SIZE BOX
    # -----------------------------------------------------

    standalone = find_standalone_size_box(
        items,
        labels
    )

    if standalone is not None:

        item = standalone["item"]

        old_text = item["text"]

        new_text = (
            f"{width} X {height}"
        )

        success = set_shape_text(
            item["shape"],
            new_text
        )

        if success:

            return {
                "status": "Updated Existing Size Box",
                "action": (
                    f"{old_text} -> {new_text}"
                )
            }

    # -----------------------------------------------------
    # 5. NOTHING FOUND
    # -----------------------------------------------------

    return {
        "status": "Size Not Found",
        "action": (
            "Existing Size field nahi mila. "
            "New box create nahi kiya."
        )
    }


# =========================================================
# FIND EXCEL COLUMN
# =========================================================

def find_column(
    columns,
    aliases
):

    normalized = {}

    for col in columns:

        normalized[
            normalize_column_name(col)
        ] = col

    # Exact
    for alias in aliases:

        alias_norm = normalize_column_name(
            alias
        )

        if alias_norm in normalized:

            return normalized[
                alias_norm
            ]

    # Flexible
    for col_norm, original in normalized.items():

        for alias in aliases:

            alias_norm = normalize_column_name(
                alias
            )

            if (
                alias_norm in col_norm
                or
                col_norm in alias_norm
            ):

                return original

    return None


# =========================================================
# READ EXCEL
# =========================================================

def read_excel_file(
    uploaded_file
):

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):

        return {
            "CSV": pd.read_csv(
                uploaded_file
            )
        }

    if filename.endswith(".xls"):

        return pd.read_excel(
            uploaded_file,
            sheet_name=None,
            engine="xlrd"
        )

    return pd.read_excel(
        uploaded_file,
        sheet_name=None,
        engine="openpyxl"
    )


# =========================================================
# GET DATAFRAME
# =========================================================

def get_dataframe(
    uploaded_file
):

    sheets = read_excel_file(
        uploaded_file
    )

    if isinstance(
        sheets,
        pd.DataFrame
    ):

        return sheets, "CSV"

    if "Merged_Result" in sheets:

        return (
            sheets["Merged_Result"],
            "Merged_Result"
        )

    for name, df in sheets.items():

        if normalize_text(
            name
        ) == "merged_result":

            return df, name

    for name, df in sheets.items():

        if (
            isinstance(df, pd.DataFrame)
            and
            not df.empty
        ):

            return df, name

    first = list(
        sheets.keys()
    )[0]

    return sheets[first], first


# =========================================================
# PROCESS PPT
# =========================================================

def process_ppt(
    ppt_bytes,
    df,
    width_col,
    height_col
):

    prs = Presentation(
        io.BytesIO(
            ppt_bytes
        )
    )

    total_slides = len(
        prs.slides
    )

    rows = []

    for index, row in df.iterrows():

        rows.append({
            "excel_row": index + 2,
            "width": clean_number(
                row.get(
                    width_col,
                    ""
                )
            ),
            "height": clean_number(
                row.get(
                    height_col,
                    ""
                )
            )
        })

    results = []

    progress = st.progress(
        0
    )

    status = st.empty()

    # =====================================================
    # SLIDE → EXCEL ROW
    # =====================================================

    for slide_number, slide in enumerate(
        prs.slides,
        start=1
    ):

        status.write(
            f"Processing Slide {slide_number} / {total_slides}"
        )

        excel_index = slide_number - 1

        if excel_index >= len(rows):

            results.append({
                "Slide": slide_number,
                "Excel Row": "",
                "Width": "",
                "Height": "",
                "Status": "Excel Row Not Available",
                "Action": ""
            })

            progress.progress(
                slide_number /
                total_slides
            )

            continue

        row = rows[
            excel_index
        ]

        width = row["width"]
        height = row["height"]

        if (
            width == ""
            or
            height == ""
        ):

            results.append({
                "Slide": slide_number,
                "Excel Row": row["excel_row"],
                "Width": width,
                "Height": height,
                "Status": "Width/Height Missing",
                "Action": ""
            })

            progress.progress(
                slide_number /
                total_slides
            )

            continue

        result = process_slide(
            slide,
            width,
            height
        )

        results.append({
            "Slide": slide_number,
            "Excel Row": row["excel_row"],
            "Width": width,
            "Height": height,
            "Status": result["status"],
            "Action": result["action"]
        })

        progress.progress(
            slide_number /
            total_slides
        )

    progress.progress(
        1.0
    )

    status.success(
        f"Completed {total_slides} slides."
    )

    return (
        prs,
        pd.DataFrame(
            results
        )
    )


# =========================================================
# FILE UPLOAD
# =========================================================

st.markdown("### 📊 Excel Master File")

excel_file = st.file_uploader(
    "Upload Excel",
    type=[
        "xlsx",
        "xls",
        "xlsm",
        "csv"
    ],
    key="excel_file"
)


st.markdown("### 📄 PowerPoint File")

ppt_file = st.file_uploader(
    "Upload PPTX",
    type=["pptx"],
    key="ppt_file"
)


# =========================================================
# MAIN
# =========================================================

if excel_file is not None:

    try:

        df, sheet_name = get_dataframe(
            excel_file
        )

        # -------------------------------------------------
        # Automatic detection - NO UI BOX
        # -------------------------------------------------

        width_col = find_column(
            df.columns,
            WIDTH_ALIASES
        )

        height_col = find_column(
            df.columns,
            HEIGHT_ALIASES
        )

        if width_col is None:

            st.error(
                "❌ Excel mein Width column nahi mila."
            )

            st.stop()

        if height_col is None:

            st.error(
                "❌ Excel mein Height column nahi mila."
            )

            st.stop()

        # -------------------------------------------------
        # PPT uploaded
        # -------------------------------------------------

        if ppt_file is not None:

            if st.button(
                "🚀 FIX PPT SIZE",
                type="primary",
                use_container_width=True
            ):

                try:

                    ppt_bytes = (
                        ppt_file.getvalue()
                    )

                    prs, result_df = process_ppt(
                        ppt_bytes,
                        df,
                        width_col,
                        height_col
                    )

                    # -------------------------------------
                    # Save to memory
                    # -------------------------------------

                    output_buffer = io.BytesIO()

                    prs.save(
                        output_buffer
                    )

                    output_buffer.seek(0)

                    st.session_state.fixed_ppt_bytes = (
                        output_buffer.getvalue()
                    )

                    st.session_state.result_df = (
                        result_df
                    )

                    st.session_state.output_filename = (
                        "PPT_SIZE_FIXED_V6.1.pptx"
                    )

                    st.success(
                        "✅ PPT successfully processed."
                    )

                except Exception as e:

                    st.error(
                        f"❌ Processing Error: {str(e)}"
                    )

                    st.exception(e)

        else:

            st.info(
                "PPTX file upload karo."
            )

    except Exception as e:

        st.error(
            f"❌ Excel Error: {str(e)}"
        )

        st.exception(e)


else:

    st.info(
        "Pehle Excel Master File upload karo."
    )


# =========================================================
# DOWNLOAD SECTION
# IMPORTANT:
# Ye FIX PPT button ke bahar hai.
# Isliye download ke baad bhi dikhega.
# =========================================================

if st.session_state.fixed_ppt_bytes is not None:

    st.markdown("---")

    st.markdown(
        "### ✅ Fixed PPT Ready"
    )

    st.download_button(
        label="⬇️ DOWNLOAD FIXED PPT",
        data=st.session_state.fixed_ppt_bytes,
        file_name=st.session_state.output_filename,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "presentationml.presentation"
        ),
        type="primary",
        use_container_width=True,
        on_click="ignore"
    )


# =========================================================
# RESULT REPORT
# =========================================================

if st.session_state.result_df is not None:

    st.markdown(
        "### 📋 Processing Report"
    )

    result_df = (
        st.session_state.result_df
    )

    updated_count = len(
        result_df[
            result_df["Status"].str.contains(
                "Updated",
                na=False
            )
        ]
    )

    not_found_count = len(
        result_df[
            result_df["Status"].str.contains(
                "Not Found",
                na=False
            )
        ]
    )

    failed_count = len(
        result_df[
            result_df["Status"].str.contains(
                "Failed",
                na=False
            )
        ]
    )

    c1, c2, c3 = st.columns(3)

    with c1:

        st.metric(
            "Updated",
            updated_count
        )

    with c2:

        st.metric(
            "Size Not Found",
            not_found_count
        )

    with c3:

        st.metric(
            "Failed",
            failed_count
        )

    st.dataframe(
        result_df,
        use_container_width=True,
        height=500
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown("---")

st.caption(
    "PPT SIZE FIXER V6.1 | "
    "Excel Row 2 → Slide 1 | "
    "Excel Row 3 → Slide 2 | "
    "Existing Size field only"
)
