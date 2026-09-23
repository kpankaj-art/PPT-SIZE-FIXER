import streamlit as st
import pandas as pd
import re
import io

from pptx import Presentation


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V7",
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
# SESSION STATE
# =========================================================

if "fixed_ppt_bytes" not in st.session_state:
    st.session_state.fixed_ppt_bytes = None

if "result_df" not in st.session_state:
    st.session_state.result_df = None

if "output_filename" not in st.session_state:
    st.session_state.output_filename = "PPT_SIZE_FIXED_V7.pptx"


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
    "district",
    "type"
}


# =========================================================
# NORMALIZATION
# =========================================================

def normalize_text(value):

    if value is None:
        return ""

    text = str(value)

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = text.replace("\t", " ")

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip().lower()


def normalize_column_name(value):

    text = normalize_text(value)

    text = text.replace("_", " ")
    text = text.replace("-", " ")

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================================================
# CLEAN EXCEL NUMBER
# =========================================================

def clean_number(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    if not text:
        return ""

    text = text.replace(",", "")

    # Remove common units
    text = re.sub(
        r"\s*(inch|inches|in)\s*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    try:

        number = float(text)

        if number.is_integer():
            return str(int(number))

        return str(
            number
        ).rstrip("0").rstrip(".")

    except Exception:

        return text


# =========================================================
# PROTECTED FIELD CHECK
# =========================================================

def is_protected_field(text):

    n = normalize_text(text)

    if not n:
        return False

    if n in PROTECTED_WORDS:
        return True

    for word in PROTECTED_WORDS:

        if n.startswith(
            word + ":"
        ):
            return True

        if n.startswith(
            word + " :"
        ):
            return True

        if n.startswith(
            word + " :-"
        ):
            return True

        if n.startswith(
            word + " -"
        ):
            return True

    return False


# =========================================================
# GET SHAPE TEXT
# =========================================================

def get_shape_text(shape):

    try:

        if not shape.has_text_frame:
            return ""

        return shape.text or ""

    except Exception:

        return ""


# =========================================================
# SET SHAPE TEXT WHILE PRESERVING FORMAT
# =========================================================

def set_shape_text(
    shape,
    new_text
):

    try:

        if not shape.has_text_frame:
            return False

        tf = shape.text_frame

        # ---------------------------------------------
        # Preserve existing first run formatting
        # ---------------------------------------------

        if tf.paragraphs:

            first_paragraph = tf.paragraphs[0]

            if first_paragraph.runs:

                first_run = (
                    first_paragraph.runs[0]
                )

                first_run.text = str(
                    new_text
                )

                # Clear extra runs
                for run in first_paragraph.runs[1:]:

                    run.text = ""

                # Clear extra paragraphs
                for paragraph in tf.paragraphs[1:]:

                    for run in paragraph.runs:

                        run.text = ""

                return True

        tf.text = str(
            new_text
        )

        return True

    except Exception:

        try:

            shape.text = str(
                new_text
            )

            return True

        except Exception:

            return False


# =========================================================
# COLLECT ALL TEXT SHAPES
# =========================================================

def collect_text_shapes(slide):

    items = []

    for index, shape in enumerate(
        slide.shapes
    ):

        text = get_shape_text(
            shape
        )

        if not text.strip():
            continue

        items.append({
            "index": index,
            "shape": shape,
            "text": text,
            "norm": normalize_text(text),
        })

    return items


# =========================================================
# SIZE PATTERNS
#
# Handles:
#
# 300X48
# 300 X 48
# 300×48
# 300 * 48
# 300x48 Inch
#
# Also incomplete:
#
# 300X
# X48
#
# =========================================================

FULL_SIZE_PATTERN = re.compile(
    r"""
    (?P<width>\d+(?:\.\d+)?)
    \s*
    [x×*]
    \s*
    (?P<height>\d+(?:\.\d+)?)
    """,
    re.IGNORECASE | re.VERBOSE
)


PARTIAL_SIZE_PATTERN = re.compile(
    r"""
    (?:
        (?P<width>\d+(?:\.\d+)?)
        \s*
        [x×*]
        \s*
    )
    |
    (?:
        \s*
        [x×*]
        \s*
        (?P<height>\d+(?:\.\d+)?)
    )
    """,
    re.IGNORECASE | re.VERBOSE
)


# =========================================================
# CHECK SIZE LABEL
# =========================================================

def contains_size_word(text):

    if not text:
        return False

    return bool(
        re.search(
            r"\bsize\b",
            str(text),
            flags=re.IGNORECASE
        )
    )


# =========================================================
# CHECK SIZE VALUE
# =========================================================

def contains_full_size(text):

    if not text:
        return False

    return bool(
        FULL_SIZE_PATTERN.search(
            str(text)
        )
    )


# =========================================================
# REPLACE SIZE INSIDE SAME TEXTBOX
#
# IMPORTANT:
# Only the Size value is replaced.
# Shop Name / Address / Mobile / Type etc.
# remain untouched.
# =========================================================

def replace_size_inside_text(
    text,
    width,
    height
):

    if not text:
        return text, False

    original = str(text)

    # -----------------------------------------------------
    # First: Full Width X Height
    # -----------------------------------------------------

    match = FULL_SIZE_PATTERN.search(
        original
    )

    if match:

        start = match.start()
        end = match.end()

        old_value = (
            original[start:end]
        )

        new_value = (
            f"{width}X{height}"
        )

        new_text = (
            original[:start]
            +
            new_value
            +
            original[end:]
        )

        return new_text, True

    # -----------------------------------------------------
    # Second: partial Size
    #
    # Example:
    # Size:- 16X
    # Size:- X27
    # -----------------------------------------------------

    # Only attempt this when the textbox
    # actually contains "Size"
    if not contains_size_word(
        original
    ):

        return original, False

    # Find Size word
    size_match = re.search(
        r"\bsize\b",
        original,
        flags=re.IGNORECASE
    )

    if not size_match:

        return original, False

    after_size = original[
        size_match.end():
    ]

    partial = PARTIAL_SIZE_PATTERN.search(
        after_size
    )

    if partial:

        start = (
            size_match.end()
            +
            partial.start()
        )

        end = (
            size_match.end()
            +
            partial.end()
        )

        new_value = (
            f"{width}X{height}"
        )

        new_text = (
            original[:start]
            +
            new_value
            +
            original[end:]
        )

        return new_text, True

    return original, False


# =========================================================
# FIND INLINE SIZE SHAPE
# =========================================================

def find_inline_size_shapes(
    items
):

    candidates = []

    for item in items:

        text = item["text"]

        # Must contain Size
        if not contains_size_word(
            text
        ):
            continue

        # Must contain dimension pattern
        if not contains_full_size(
            text
        ):

            # Partial format allowed
            if not re.search(
                r"\bsize\b.*?[x×*]",
                text,
                flags=re.IGNORECASE
                | re.DOTALL
            ):

                continue

        candidates.append(
            item
        )

    return candidates


# =========================================================
# FIND SIZE LABELS
# =========================================================

def find_size_labels(
    items
):

    labels = []

    for item in items:

        text = item["text"]

        if not contains_size_word(
            text
        ):
            continue

        # If same shape already contains size value,
        # it will be handled as inline.
        if contains_full_size(
            text
        ):
            continue

        # Partial size also considered
        if re.search(
            r"\bsize\b.*?[x×*]",
            text,
            flags=re.IGNORECASE
            | re.DOTALL
        ):
            continue

        labels.append(
            item
        )

    return labels


# =========================================================
# CHECK NUMBER SHAPE
# =========================================================

def is_number_text(text):

    if not text:
        return False

    return bool(
        re.fullmatch(
            r"\s*\d+(?:\.\d+)?\s*",
            str(text)
        )
    )


# =========================================================
# CHECK X SHAPE
# =========================================================

def is_x_text(text):

    if not text:
        return False

    return normalize_text(
        text
    ) in {
        "x",
        "×",
        "*"
    }


# =========================================================
# FIND SEPARATE W X H
#
# Example:
#
# Size:
# 15
# X
# 3
#
# =========================================================

def find_separate_size(
    items,
    labels
):

    candidates = []

    for label in labels:

        label_shape = label[
            "shape"
        ]

        label_left = (
            label_shape.left
        )

        label_right = (
            label_shape.left
            +
            label_shape.width
        )

        label_top = (
            label_shape.top
        )

        label_bottom = (
            label_shape.top
            +
            label_shape.height
        )

        x_candidates = []

        for item in items:

            if item["shape"] is label_shape:
                continue

            if not is_x_text(
                item["text"]
            ):
                continue

            shape = item[
                "shape"
            ]

            center_x = (
                shape.left
                +
                shape.width / 2
            )

            center_y = (
                shape.top
                +
                shape.height / 2
            )

            # X should be near Size label
            if (
                center_x <
                label_left - Inches(0.3)
            ):
                continue

            if (
                center_x >
                label_right + Inches(0.3)
            ):
                continue

            if abs(
                center_y -
                (
                    label_top
                    +
                    label_shape.height / 2
                )
            ) > Inches(0.6):

                continue

            x_candidates.append(
                item
            )

        for x_item in x_candidates:

            x_shape = x_item[
                "shape"
            ]

            x_center_x = (
                x_shape.left
                +
                x_shape.width / 2
            )

            x_center_y = (
                x_shape.top
                +
                x_shape.height / 2
            )

            left_numbers = []
            right_numbers = []

            for item in items:

                shape = item[
                    "shape"
                ]

                if shape is label_shape:
                    continue

                if shape is x_shape:
                    continue

                if not is_number_text(
                    item["text"]
                ):
                    continue

                # Do not touch protected field
                if is_protected_field(
                    item["text"]
                ):
                    continue

                center_x = (
                    shape.left
                    +
                    shape.width / 2
                )

                center_y = (
                    shape.top
                    +
                    shape.height / 2
                )

                if abs(
                    center_y -
                    x_center_y
                ) > Inches(0.5):

                    continue

                # Must remain inside/near Size label
                if (
                    center_x <
                    label_left - Inches(0.15)
                ):
                    continue

                if (
                    center_x >
                    label_right + Inches(0.15)
                ):
                    continue

                if center_x < x_center_x:

                    left_numbers.append(
                        item
                    )

                elif center_x > x_center_x:

                    right_numbers.append(
                        item
                    )

            if (
                not left_numbers
                or
                not right_numbers
            ):
                continue

            left_numbers.sort(
                key=lambda x:
                abs(
                    (
                        x["shape"].left
                        +
                        x["shape"].width / 2
                    )
                    -
                    x_center_x
                )
            )

            right_numbers.sort(
                key=lambda x:
                abs(
                    (
                        x["shape"].left
                        +
                        x["shape"].width / 2
                    )
                    -
                    x_center_x
                )
            )

            candidates.append({
                "label": label,
                "width": left_numbers[0],
                "x": x_item,
                "height": right_numbers[0]
            })

    if not candidates:
        return None

    return candidates[0]


# =========================================================
# FIND STANDALONE SIZE BOX
#
# Example:
#
# Size:
# 180 X 36
#
# =========================================================

def find_standalone_size(
    items,
    labels
):

    for item in items:

        text = item[
            "text"
        ].strip()

        if not re.fullmatch(
            r"\d+(?:\.\d+)?\s*[x×*]\s*\d+(?:\.\d+)?",
            text,
            flags=re.IGNORECASE
        ):
            continue

        shape = item[
            "shape"
        ]

        center_x = (
            shape.left
            +
            shape.width / 2
        )

        center_y = (
            shape.top
            +
            shape.height / 2
        )

        for label in labels:

            label_shape = label[
                "shape"
            ]

            label_center_x = (
                label_shape.left
                +
                label_shape.width / 2
            )

            label_center_y = (
                label_shape.top
                +
                label_shape.height / 2
            )

            if abs(
                center_y -
                label_center_y
            ) > Inches(0.7):

                continue

            if abs(
                center_x -
                label_center_x
            ) > Inches(4):

                continue

            return item

    return None


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

    # =====================================================
    # PRIORITY 1
    # SIZE INSIDE SAME TEXTBOX
    # =====================================================

    inline_items = find_inline_size_shapes(
        items
    )

    if inline_items:

        # If multiple Size-containing textboxes,
        # process the first valid one.
        for item in inline_items:

            old_text = item[
                "text"
            ]

            new_text, changed = (
                replace_size_inside_text(
                    old_text,
                    width,
                    height
                )
            )

            if changed:

                success = set_shape_text(
                    item["shape"],
                    new_text
                )

                if success:

                    return {
                        "status": "Updated Size Inside Textbox",
                        "action": (
                            f"{old_text} -> {new_text}"
                        )
                    }

    # =====================================================
    # PRIORITY 2
    # SEPARATE SIZE LABEL + W X H
    # =====================================================

    labels = find_size_labels(
        items
    )

    separate = find_separate_size(
        items,
        labels
    )

    if separate is not None:

        old_width = separate[
            "width"
        ]["text"]

        old_height = separate[
            "height"
        ]["text"]

        ok_width = set_shape_text(
            separate["width"]["shape"],
            str(width)
        )

        ok_x = set_shape_text(
            separate["x"]["shape"],
            "X"
        )

        ok_height = set_shape_text(
            separate["height"]["shape"],
            str(height)
        )

        if (
            ok_width
            and
            ok_x
            and
            ok_height
        ):

            return {
                "status": "Updated Separate Size Fields",
                "action": (
                    f"{old_width} X {old_height}"
                    f" -> "
                    f"{width} X {height}"
                )
            }

    # =====================================================
    # PRIORITY 3
    # STANDALONE COMBINED SIZE BOX
    # =====================================================

    standalone = find_standalone_size(
        items,
        labels
    )

    if standalone is not None:

        old_text = standalone[
            "text"
        ]

        new_text = (
            f"{width} X {height}"
        )

        success = set_shape_text(
            standalone["shape"],
            new_text
        )

        if success:

            return {
                "status": "Updated Existing Size Box",
                "action": (
                    f"{old_text} -> {new_text}"
                )
            }

    # =====================================================
    # NOT FOUND
    # =====================================================

    return {
        "status": "Size Not Found",
        "action": (
            "Size field/value detect nahi hua. "
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

    normalized_columns = {}

    for column in columns:

        normalized_columns[
            normalize_column_name(
                column
            )
        ] = column

    # Exact
    for alias in aliases:

        alias_normalized = (
            normalize_column_name(
                alias
            )
        )

        if (
            alias_normalized
            in
            normalized_columns
        ):

            return normalized_columns[
                alias_normalized
            ]

    # Flexible
    for column_normalized, original in (
        normalized_columns.items()
    ):

        for alias in aliases:

            alias_normalized = (
                normalize_column_name(
                    alias
                )
            )

            if (
                alias_normalized
                in
                column_normalized
                or
                column_normalized
                in
                alias_normalized
            ):

                return original

    return None


# =========================================================
# READ EXCEL
# =========================================================

def read_excel_file(
    uploaded_file
):

    filename = (
        uploaded_file.name.lower()
    )

    if filename.endswith(
        ".csv"
    ):

        return {
            "CSV": pd.read_csv(
                uploaded_file
            )
        }

    if filename.endswith(
        ".xls"
    ):

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

    # Preferred sheet
    if "Merged_Result" in sheets:

        return (
            sheets["Merged_Result"],
            "Merged_Result"
        )

    # Case insensitive
    for name, df in sheets.items():

        if normalize_text(
            name
        ) == "merged_result":

            return df, name

    # First non-empty
    for name, df in sheets.items():

        if (
            isinstance(
                df,
                pd.DataFrame
            )
            and
            not df.empty
        ):

            return df, name

    first_name = list(
        sheets.keys()
    )[0]

    return (
        sheets[first_name],
        first_name
    )


# =========================================================
# PROCESS PPT
#
# Excel Row 2 -> Slide 1
# Excel Row 3 -> Slide 2
# ...
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

    results = []

    # -----------------------------------------------------
    # Excel rows
    # -----------------------------------------------------

    excel_rows = []

    for index, row in df.iterrows():

        excel_rows.append({
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

    progress = st.progress(
        0
    )

    status_box = st.empty()

    # =====================================================
    # SLIDES
    # =====================================================

    for slide_number, slide in enumerate(
        prs.slides,
        start=1
    ):

        status_box.write(
            f"Processing Slide {slide_number} / {total_slides}"
        )

        excel_index = (
            slide_number - 1
        )

        # -------------------------------------------------
        # No Excel row
        # -------------------------------------------------

        if excel_index >= len(
            excel_rows
        ):

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

        row = excel_rows[
            excel_index
        ]

        width = row[
            "width"
        ]

        height = row[
            "height"
        ]

        # -------------------------------------------------
        # Missing size
        # -------------------------------------------------

        if (
            width == ""
            or
            height == ""
        ):

            results.append({
                "Slide": slide_number,
                "Excel Row": row[
                    "excel_row"
                ],
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

        # -------------------------------------------------
        # Process slide
        # -------------------------------------------------

        result = process_slide(
            slide,
            width,
            height
        )

        results.append({
            "Slide": slide_number,
            "Excel Row": row[
                "excel_row"
            ],
            "Width": width,
            "Height": height,
            "Status": result[
                "status"
            ],
            "Action": result[
                "action"
            ]
        })

        progress.progress(
            slide_number /
            total_slides
        )

    progress.progress(
        1.0
    )

    status_box.success(
        f"Completed {total_slides} slides."
    )

    return (
        prs,
        pd.DataFrame(
            results
        )
    )


# =========================================================
# UI
# =========================================================

st.markdown(
    "### 📊 Excel Master File"
)

excel_file = st.file_uploader(
    "Upload Excel",
    type=[
        "xlsx",
        "xls",
        "xlsm",
        "csv"
    ],
    key="excel_upload_v7"
)


st.markdown(
    "### 📄 PowerPoint File"
)

ppt_file = st.file_uploader(
    "Upload PPTX",
    type=[
        "pptx"
    ],
    key="ppt_upload_v7"
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
        # Automatic Width/Height detection
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
        # PPT
        # -------------------------------------------------

        if ppt_file is not None:

            if st.button(
                "🚀 FIX PPT SIZE",
                type="primary",
                use_container_width=True
            ):

                try:

                    # Clear previous result
                    st.session_state.fixed_ppt_bytes = None
                    st.session_state.result_df = None

                    ppt_bytes = (
                        ppt_file.getvalue()
                    )

                    prs, result_df = process_ppt(
                        ppt_bytes,
                        df,
                        width_col,
                        height_col
                    )

                    # -------------------------------------------------
                    # Save in memory
                    # -------------------------------------------------

                    output_buffer = io.BytesIO()

                    prs.save(
                        output_buffer
                    )

                    output_buffer.seek(
                        0
                    )

                    st.session_state.fixed_ppt_bytes = (
                        output_buffer.getvalue()
                    )

                    st.session_state.result_df = (
                        result_df
                    )

                    st.session_state.output_filename = (
                        "PPT_SIZE_FIXED_V7.pptx"
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
# DOWNLOAD
#
# Download button processing ke baad bhi visible rahega.
# =========================================================

if (
    st.session_state.fixed_ppt_bytes
    is not None
):

    st.markdown(
        "---"
    )

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
# REPORT
# =========================================================

if (
    st.session_state.result_df
    is not None
):

    result_df = (
        st.session_state.result_df
    )

    st.markdown(
        "### 📋 Processing Report"
    )

    updated = len(
        result_df[
            result_df[
                "Status"
            ].str.contains(
                "Updated",
                na=False
            )
        ]
    )

    not_found = len(
        result_df[
            result_df[
                "Status"
            ].str.contains(
                "Not Found",
                na=False
            )
        ]
    )

    failed = len(
        result_df[
            result_df[
                "Status"
            ].str.contains(
                "Failed",
                na=False
            )
        ]
    )

    c1, c2, c3 = st.columns(
        3
    )

    with c1:

        st.metric(
            "Updated",
            updated
        )

    with c2:

        st.metric(
            "Size Not Found",
            not_found
        )

    with c3:

        st.metric(
            "Failed",
            failed
        )

    st.dataframe(
        result_df,
        use_container_width=True,
        height=500
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    "---"
)

st.caption(
    "PPT SIZE FIXER V7 | "
    "Excel Row 2 → Slide 1 | "
    "Existing Size field only | "
    "No new Size box created"
)
