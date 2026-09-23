import streamlit as st
import pandas as pd
import re
import io

from pptx import Presentation


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V8",
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
    st.session_state.output_filename = "PPT_SIZE_FIXED_V8.pptx"


# =========================================================
# WIDTH / HEIGHT ALIASES
# IMPORTANT:
# Short aliases like "H" are NOT used for automatic
# detection unless no proper Height heading exists.
# =========================================================

WIDTH_EXACT_ALIASES = [
    "width",
    "width inches",
    "width inch",
    "width (inches)",
    "width (inch)",
    "width in inches",
    "board width",
    "size width",
    "width size"
]

HEIGHT_EXACT_ALIASES = [
    "height",
    "height inches",
    "height inch",
    "height (inches)",
    "height (inch)",
    "height in inches",
    "board height",
    "size height",
    "height size"
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
# NORMALIZE TEXT
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


# =========================================================
# NORMALIZE COLUMN NAME
# =========================================================

def normalize_column_name(value):

    text = normalize_text(value)

    # Replace common separators
    text = text.replace("_", " ")
    text = text.replace("-", " ")

    # Remove brackets
    text = text.replace("(", " ")
    text = text.replace(")", " ")

    # Remove punctuation
    text = re.sub(
        r"[^a-z0-9\s]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


# =========================================================
# COLUMN TYPE DETECTION
# =========================================================

def column_contains_word(
    column_name,
    word
):

    normalized = normalize_column_name(
        column_name
    )

    words = normalized.split()

    return word in words


# =========================================================
# FIND WIDTH COLUMN
#
# STRICT:
# First exact proper Width heading.
# Then token based Width heading.
# NEVER uses "W" alone.
# =========================================================

def find_width_column(
    columns
):

    # -----------------------------------------------------
    # 1. Exact aliases
    # -----------------------------------------------------

    normalized_map = {}

    for col in columns:

        normalized_map[
            normalize_column_name(col)
        ] = col

    for alias in WIDTH_EXACT_ALIASES:

        alias_norm = normalize_column_name(
            alias
        )

        if alias_norm in normalized_map:

            return normalized_map[
                alias_norm
            ]

    # -----------------------------------------------------
    # 2. Token based
    # -----------------------------------------------------

    candidates = []

    for col in columns:

        normalized = normalize_column_name(
            col
        )

        words = normalized.split()

        if "width" in words:

            score = 0

            if normalized.startswith(
                "width"
            ):
                score += 100

            if "inch" in words:
                score += 20

            if "inches" in words:
                score += 20

            if "size" in words:
                score += 5

            candidates.append(
                (
                    score,
                    col
                )
            )

    if candidates:

        candidates.sort(
            key=lambda x: x[0],
            reverse=True
        )

        return candidates[0][1]

    return None


# =========================================================
# FIND HEIGHT COLUMN
#
# STRICT:
# First exact proper Height heading.
# Then token based Height heading.
# NEVER uses "H" alone unless explicitly enabled
# as a final fallback.
# =========================================================

def find_height_column(
    columns,
    width_column=None
):

    # -----------------------------------------------------
    # 1. Exact aliases
    # -----------------------------------------------------

    normalized_map = {}

    for col in columns:

        # Never use Width column as Height
        if (
            width_column is not None
            and
            col == width_column
        ):
            continue

        normalized_map[
            normalize_column_name(col)
        ] = col

    for alias in HEIGHT_EXACT_ALIASES:

        alias_norm = normalize_column_name(
            alias
        )

        if alias_norm in normalized_map:

            return normalized_map[
                alias_norm
            ]

    # -----------------------------------------------------
    # 2. Token based
    # -----------------------------------------------------

    candidates = []

    for col in columns:

        if (
            width_column is not None
            and
            col == width_column
        ):
            continue

        normalized = normalize_column_name(
            col
        )

        words = normalized.split()

        if "height" in words:

            score = 0

            if normalized.startswith(
                "height"
            ):
                score += 100

            if "inch" in words:
                score += 20

            if "inches" in words:
                score += 20

            if "size" in words:
                score += 5

            candidates.append(
                (
                    score,
                    col
                )
            )

    if candidates:

        candidates.sort(
            key=lambda x: x[0],
            reverse=True
        )

        return candidates[0][1]

    # -----------------------------------------------------
    # 3. LAST RESORT:
    # Only use H if there is literally a column named H.
    #
    # But NEVER use arbitrary fuzzy matching.
    # -----------------------------------------------------

    for col in columns:

        if (
            width_column is not None
            and
            col == width_column
        ):
            continue

        normalized = normalize_column_name(
            col
        )

        if normalized == "h":

            return col

    return None


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

    # Remove units
    text = re.sub(
        r"\s*(inch|inches|in)\s*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    # Remove spaces
    text = text.strip()

    try:

        number = float(text)

        if number.is_integer():

            return str(
                int(number)
            )

        return (
            str(number)
            .rstrip("0")
            .rstrip(".")
        )

    except Exception:

        return text


# =========================================================
# PROTECTED FIELD
# =========================================================

def is_protected_field(text):

    normalized = normalize_text(
        text
    )

    if not normalized:
        return False

    if normalized in PROTECTED_WORDS:
        return True

    for word in PROTECTED_WORDS:

        if normalized.startswith(
            word + ":"
        ):
            return True

        if normalized.startswith(
            word + " :"
        ):
            return True

        if normalized.startswith(
            word + " :-"
        ):
            return True

        if normalized.startswith(
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
# SET SHAPE TEXT
# =========================================================

def set_shape_text(
    shape,
    new_text
):

    try:

        if not shape.has_text_frame:
            return False

        text_frame = shape.text_frame

        # ---------------------------------------------
        # Preserve first run formatting
        # ---------------------------------------------

        if text_frame.paragraphs:

            paragraph = (
                text_frame.paragraphs[0]
            )

            if paragraph.runs:

                first_run = (
                    paragraph.runs[0]
                )

                first_run.text = str(
                    new_text
                )

                # Clear other runs
                for run in paragraph.runs[1:]:

                    run.text = ""

                # Clear other paragraphs
                for p in text_frame.paragraphs[1:]:

                    for run in p.runs:

                        run.text = ""

                return True

        text_frame.text = str(
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
# COLLECT TEXT SHAPES
# =========================================================

def collect_text_shapes(
    slide
):

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
            "norm": normalize_text(text)
        })

    return items


# =========================================================
# SIZE PATTERN
#
# Handles:
#
# 300X48
# 300 X 48
# 300x48
# 300×48
# 300 * 48
# 300 X 48 Inch
#
# =========================================================

FULL_SIZE_PATTERN = re.compile(
    r"""
    (?P<width>
        \d+(?:\.\d+)?
    )
    \s*
    [x×*]
    \s*
    (?P<height>
        \d+(?:\.\d+)?
    )
    """,
    re.IGNORECASE |
    re.VERBOSE
)


# =========================================================
# SIZE WORD
# =========================================================

def contains_size_word(
    text
):

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
# FULL SIZE VALUE
# =========================================================

def contains_full_size(
    text
):

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
# Example:
#
# Shop Name : ABC
# Address : XYZ
# Type : GSB
# Size:- 300X48 Inch
#
# becomes:
#
# Shop Name : ABC
# Address : XYZ
# Type : GSB
# Size:- 180X36 Inch
#
# Only 300X48 is changed.
# =========================================================

def replace_size_inside_text(
    text,
    width,
    height
):

    if not text:
        return text, False

    original = str(
        text
    )

    # -----------------------------------------------------
    # Find "Size"
    # -----------------------------------------------------

    size_match = re.search(
        r"\bsize\b",
        original,
        flags=re.IGNORECASE
    )

    if not size_match:

        return original, False

    # -----------------------------------------------------
    # IMPORTANT:
    # Search size value ONLY AFTER "Size".
    # This prevents another number earlier in the
    # textbox from being replaced.
    # -----------------------------------------------------

    after_size_start = (
        size_match.end()
    )

    after_size = original[
        after_size_start:
    ]

    # -----------------------------------------------------
    # Full Width X Height
    # -----------------------------------------------------

    size_value_match = (
        FULL_SIZE_PATTERN.search(
            after_size
        )
    )

    if size_value_match:

        start = (
            after_size_start
            +
            size_value_match.start()
        )

        end = (
            after_size_start
            +
            size_value_match.end()
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
# FIND INLINE SIZE
# =========================================================

def find_inline_size_shapes(
    items
):

    candidates = []

    for item in items:

        text = item[
            "text"
        ]

        if not contains_size_word(
            text
        ):
            continue

        if not contains_full_size(
            text
        ):
            continue

        candidates.append(
            item
        )

    return candidates


# =========================================================
# SIZE LABEL
# =========================================================

def find_size_labels(
    items
):

    labels = []

    for item in items:

        text = item[
            "text"
        ]

        if not contains_size_word(
            text
        ):
            continue

        if contains_full_size(
            text
        ):
            continue

        labels.append(
            item
        )

    return labels


# =========================================================
# NUMBER TEXT
# =========================================================

def is_number_text(
    text
):

    if not text:
        return False

    return bool(
        re.fullmatch(
            r"\s*\d+(?:\.\d+)?\s*",
            str(text)
        )
    )


# =========================================================
# X TEXT
# =========================================================

def is_x_text(
    text
):

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
# FIND SEPARATE SIZE
#
# Handles:
#
# Size:
# 180
# X
# 36
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

        label_center_y = (
            label_shape.top
            +
            label_shape.height / 2
        )

        # -------------------------------------------------
        # Find X
        # -------------------------------------------------

        x_candidates = []

        for item in items:

            shape = item[
                "shape"
            ]

            if shape is label_shape:
                continue

            if not is_x_text(
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

            if (
                center_x <
                label_left - Inches(0.5)
            ):
                continue

            if (
                center_x >
                label_right + Inches(0.5)
            ):
                continue

            if abs(
                center_y -
                label_center_y
            ) > Inches(0.7):

                continue

            x_candidates.append(
                item
            )

        # -------------------------------------------------
        # For every X find left/right numbers
        # -------------------------------------------------

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
                ) > Inches(0.55):

                    continue

                # Keep numbers in Size area
                if (
                    center_x <
                    label_left - Inches(0.25)
                ):
                    continue

                if (
                    center_x >
                    label_right + Inches(0.25)
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

            # Closest left number
            left_numbers.sort(
                key=lambda item:
                abs(
                    (
                        item["shape"].left
                        +
                        item["shape"].width / 2
                    )
                    -
                    x_center_x
                )
            )

            # Closest right number
            right_numbers.sort(
                key=lambda item:
                abs(
                    (
                        item["shape"].left
                        +
                        item["shape"].width / 2
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
# STANDALONE SIZE BOX
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
    # SEPARATE SIZE FIELDS
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
    # EXISTING STANDALONE SIZE
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

    # Merged_Result priority
    if "Merged_Result" in sheets:

        return (
            sheets["Merged_Result"],
            "Merged_Result"
        )

    # Case-insensitive Merged_Result
    for name, df in sheets.items():

        if normalize_text(
            name
        ) == "merged_result":

            return df, name

    # First non-empty sheet
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
# Excel Row 4 -> Slide 3
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
    # Prepare Excel rows
    # -----------------------------------------------------

    excel_rows = []

    for index, row in df.iterrows():

        width = clean_number(
            row.get(
                width_col,
                ""
            )
        )

        height = clean_number(
            row.get(
                height_col,
                ""
            )
        )

        excel_rows.append({
            "excel_row": index + 2,
            "width": width,
            "height": height
        })

    progress = st.progress(
        0
    )

    status_box = st.empty()

    # =====================================================
    # SLIDE LOOP
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
        # Excel row not available
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
        # Missing Width / Height
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
        # Process
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
# UPLOAD EXCEL
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
    key="excel_upload_v8"
)


# =========================================================
# UPLOAD PPT
# =========================================================

st.markdown(
    "### 📄 PowerPoint File"
)

ppt_file = st.file_uploader(
    "Upload PPTX",
    type=[
        "pptx"
    ],
    key="ppt_upload_v8"
)


# =========================================================
# MAIN
# =========================================================

if excel_file is not None:

    try:

        df, sheet_name = get_dataframe(
            excel_file
        )

        # =================================================
        # STRICT AUTOMATIC DETECTION
        # =================================================

        width_col = find_width_column(
            df.columns
        )

        height_col = find_height_column(
            df.columns,
            width_col
        )

        # -------------------------------------------------
        # Width not found
        # -------------------------------------------------

        if width_col is None:

            st.error(
                "❌ Width column automatically detect nahi hua."
            )

            st.write(
                "Available Excel headings:"
            )

            st.write(
                list(df.columns)
            )

            st.stop()

        # -------------------------------------------------
        # Height not found
        # -------------------------------------------------

        if height_col is None:

            st.error(
                "❌ Height column automatically detect nahi hua."
            )

            st.write(
                "Available Excel headings:"
            )

            st.write(
                list(df.columns)
            )

            st.stop()

        # -------------------------------------------------
        # SAFETY CHECK
        # -------------------------------------------------

        if width_col == height_col:

            st.error(
                "❌ Width aur Height same column detect ho gaye. "
                "Processing stop kar di gayi hai."
            )

            st.write(
                "Available headings:"
            )

            st.write(
                list(df.columns)
            )

            st.stop()

        # =================================================
        # PPT
        # =================================================

        if ppt_file is not None:

            if st.button(
                "🚀 FIX PPT SIZE",
                type="primary",
                use_container_width=True
            ):

                try:

                    # Clear previous
                    st.session_state.fixed_ppt_bytes = None
                    st.session_state.result_df = None

                    ppt_bytes = (
                        ppt_file.getvalue()
                    )

                    # -------------------------------------
                    # Process
                    # -------------------------------------

                    prs, result_df = process_ppt(
                        ppt_bytes,
                        df,
                        width_col,
                        height_col
                    )

                    # -------------------------------------
                    # Save in memory
                    # -------------------------------------

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
                        "PPT_SIZE_FIXED_V8.pptx"
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
    "PPT SIZE FIXER V8 | "
    "Excel Row 2 → Slide 1 | "
    "Excel Row 3 → Slide 2 | "
    "Existing Size field only"
)
