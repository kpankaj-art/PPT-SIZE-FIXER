import streamlit as st
import pandas as pd
import re
import io
import os

from pptx import Presentation
from pptx.util import Inches, Pt


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V6",
    page_icon="📐",
    layout="wide"
)


# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>

.main-title {
    font-size: 32px;
    font-weight: 800;
    margin-bottom: 5px;
}

.sub-title {
    color: #777;
    margin-bottom: 25px;
}

.detected-box {
    padding: 12px;
    border-radius: 8px;
    background: #e8f5e9;
    border: 1px solid #81c784;
    margin-bottom: 15px;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# TITLE
# =========================================================

st.markdown(
    '<div class="main-title">📐 PPT SIZE FIXER V6</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">'
    'Excel Width × Height ko PowerPoint ke existing Size field mein replace karega.'
    '</div>',
    unsafe_allow_html=True
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
# INLINE SIZE PATTERN
# =========================================================

SIZE_PATTERN = re.compile(
    r"\d+(?:\.\d+)?\s*[x×*]\s*\d+(?:\.\d+)?",
    re.IGNORECASE
)


# =========================================================
# DETECT SIZE VALUE
# =========================================================

def contains_size_value(text):

    if not text:
        return False

    return bool(
        SIZE_PATTERN.search(
            str(text)
        )
    )


# =========================================================
# STANDALONE SIZE
# =========================================================

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
# IS X SHAPE
# =========================================================

def is_x_shape(text):

    if not text:
        return False

    n = normalize_text(text)

    return n in {
        "x",
        "×",
        "*"
    }


# =========================================================
# IS NUMBER
# =========================================================

def is_number(text):

    if not text:
        return False

    raw = str(text).strip()

    return bool(
        re.match(
            r"^\d+(?:\.\d+)?$",
            raw
        )
    )


# =========================================================
# FIND SIZE LABEL
# =========================================================

def find_size_labels(items):

    labels = []

    for item in items:

        n = item["norm"]

        # Exact Size
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

        # Size label with extra text,
        # but NOT inline numeric size
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

    replacement = f"{width} X {height}"

    new_text, count = SIZE_PATTERN.subn(
        replacement,
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

        # Preserve first run formatting
        if tf.paragraphs:

            paragraph = tf.paragraphs[0]

            if paragraph.runs:

                paragraph.runs[0].text = str(
                    new_text
                )

                # Remove extra runs
                for run in paragraph.runs[1:]:

                    run.text = ""

                # Clear extra paragraphs
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
# DISTANCE HELPERS
# =========================================================

def y_distance(a, b):

    return abs(
        a["geo"]["cy"] -
        b["geo"]["cy"]
    )


def x_distance(a, b):

    return abs(
        a["geo"]["cx"] -
        b["geo"]["cx"]
    )


# =========================================================
# FIND INLINE SIZE SHAPES
# =========================================================

def find_inline_size_shapes(items):

    result = []

    for item in items:

        text = item["text"]

        if contains_size_value(text):

            result.append(item)

    return result


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

        # Same horizontal row
        if abs(
            geo["cy"] -
            label_geo["cy"]
        ) > Inches(0.55):

            continue

        # X should be reasonably close
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

        # Same row
        if abs(
            geo["cy"] -
            x_geo["cy"]
        ) > Inches(0.40):

            continue

        # Only reasonably close
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

    all_candidates = []

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

            # Score
            score = 0

            # X close to label
            score += max(
                0,
                100 -
                int(
                    x_distance(
                        x_item,
                        label
                    ) /
                    Inches(0.05)
                )
            )

            all_candidates.append({
                "label": label,
                "components": components,
                "score": score
            })

    if not all_candidates:
        return None

    all_candidates.sort(
        key=lambda x: x["score"],
        reverse=True
    )

    return all_candidates[0]


# =========================================================
# FIND STANDALONE SIZE BOX
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

        # Find nearest Size label
        nearest_label = None
        nearest_distance = None

        for label in labels:

            distance = abs(
                geo["cy"] -
                label["geo"]["cy"]
            )

            horizontal = abs(
                geo["cx"] -
                label["geo"]["cx"]
            )

            if distance > Inches(0.60):
                continue

            if horizontal > Inches(4.0):
                continue

            total = distance + horizontal

            if (
                nearest_distance is None
                or
                total < nearest_distance
            ):

                nearest_distance = total
                nearest_label = label

        if nearest_label is not None:

            candidates.append({
                "item": item,
                "label": nearest_label,
                "distance": nearest_distance
            })

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x["distance"]
    )

    return candidates[0]


# =========================================================
# REMOVE DUPLICATE STANDALONE SIZE
# =========================================================

def delete_duplicate_size_boxes(
    items,
    selected_structure,
    selected_standalone
):

    if selected_structure is None:
        return 0

    components = selected_structure["components"]

    protected_shape_ids = {
        id(
            components["width"]["shape"]
        ),
        id(
            components["x"]["shape"]
        ),
        id(
            components["height"]["shape"]
        )
    }

    deleted = 0

    for item in items:

        shape = item["shape"]

        if id(shape) in protected_shape_ids:
            continue

        if not is_standalone_size(
            item["text"]
        ):
            continue

        # Selected standalone is definitely duplicate
        if (
            selected_standalone is not None
            and
            shape is selected_standalone["item"]["shape"]
        ):

            try:

                element = shape._element

                parent = element.getparent()

                parent.remove(element)

                deleted += 1

            except Exception:
                pass

    return deleted


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
    # 1. Inline Size
    # -----------------------------------------------------

    inline_candidates = []

    for item in items:

        text = item["text"]

        if not contains_size_value(
            text
        ):
            continue

        # Don't treat a random text containing
        # numbers as Size unless it has Size nearby.
        n = item["norm"]

        if n.startswith("size"):

            inline_candidates.append(
                item
            )

    # Direct inline Size
    if inline_candidates:

        item = inline_candidates[0]

        old_text = item["text"]

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
    # 2. Size Labels
    # -----------------------------------------------------

    labels = find_size_labels(
        items
    )

    # -----------------------------------------------------
    # 3. Separate W X H
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

        old_x = components[
            "x"
        ]["text"]

        old_height = components[
            "height"
        ]["text"]

        ok_width = set_shape_text(
            components["width"]["shape"],
            str(width)
        )

        ok_x = set_shape_text(
            components["x"]["shape"],
            "X"
        )

        ok_height = set_shape_text(
            components["height"]["shape"],
            str(height)
        )

        # Find standalone duplicate
        standalone = find_standalone_size_box(
            items,
            labels
        )

        deleted = 0

        if standalone is not None:

            deleted = delete_duplicate_size_boxes(
                items,
                structure,
                standalone
            )

        if (
            ok_width
            and
            ok_x
            and
            ok_height
        ):

            action = (
                f"{old_width} X {old_height}"
                f" -> "
                f"{width} X {height}"
            )

            if deleted:

                action += (
                    f" | Deleted duplicate size box: "
                    f"{deleted}"
                )

            return {
                "status": "Updated Existing W-X-H Fields",
                "action": action
            }

        return {
            "status": "Size Fields Found But Update Failed",
            "action": ""
        }

    # -----------------------------------------------------
    # 4. Existing Standalone Size
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

        return {
            "status": "Existing Size Box Found But Update Failed",
            "action": old_text
        }

    # -----------------------------------------------------
    # 5. Size Not Found
    # -----------------------------------------------------

    return {
        "status": "Size Not Found",
        "action": (
            "No existing Size field detected. "
            "No new box created."
        )
    }


# =========================================================
# EXCEL COLUMN FINDER
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

    # Exact match
    for alias in aliases:

        alias_norm = normalize_column_name(
            alias
        )

        if alias_norm in normalized:

            return normalized[
                alias_norm
            ]

    # Flexible match
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

def read_excel(
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
# GET PREFERRED SHEET
# =========================================================

def get_dataframe(
    uploaded_file
):

    sheets = read_excel(
        uploaded_file
    )

    # CSV
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

    # Case insensitive
    for name, df in sheets.items():

        if normalize_text(
            name
        ) == "merged_result":

            return df, name

    # First non-empty
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

    results = []

    progress = st.progress(
        0
    )

    status = st.empty()

    # -----------------------------------------------------
    # Excel rows
    # -----------------------------------------------------

    rows = []

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

        rows.append({
            "excel_row": index + 2,
            "width": width,
            "height": height
        })

    # -----------------------------------------------------
    # Slides
    # -----------------------------------------------------

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

        # -----------------------------------------------
        # Process
        # -----------------------------------------------

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
# UI
# =========================================================

st.markdown(
    "### 📊 Step 1 — Upload Excel"
)

excel_file = st.file_uploader(
    "Upload Excel Master File",
    type=[
        "xlsx",
        "xls",
        "xlsm",
        "csv"
    ]
)


st.markdown(
    "### 📄 Step 2 — Upload PowerPoint"
)

ppt_file = st.file_uploader(
    "Upload PPTX File",
    type=["pptx"]
)


# =========================================================
# EXCEL AUTO DETECTION
# =========================================================

if excel_file is not None:

    try:

        df, sheet_name = get_dataframe(
            excel_file
        )

        st.success(
            f"Excel Loaded: {excel_file.name} | Sheet: {sheet_name}"
        )

        # -------------------------------------------------
        # AUTOMATIC COLUMN DETECTION
        # -------------------------------------------------

        width_col = find_column(
            df.columns,
            WIDTH_ALIASES
        )

        height_col = find_column(
            df.columns,
            HEIGHT_ALIASES
        )

        if (
            width_col is None
            or
            height_col is None
        ):

            st.error(
                "❌ Width / Height column automatically detect nahi ho paya."
            )

            st.write(
                "Excel mein available headings:"
            )

            st.write(
                list(df.columns)
            )

            st.stop()

        # -------------------------------------------------
        # DETECTED INFO
        # -------------------------------------------------

        st.markdown(
            f"""
            <div class="detected-box">

            ✅ <b>Automatic Column Detection</b><br><br>

            Width Column: <b>{width_col}</b><br>
            Height Column: <b>{height_col}</b>

            </div>
            """,
            unsafe_allow_html=True
        )

        # -------------------------------------------------
        # Preview
        # -------------------------------------------------

        st.markdown(
            "### 🔎 Width × Height Preview"
        )

        st.dataframe(
            df[
                [
                    width_col,
                    height_col
                ]
            ].head(10),
            use_container_width=True
        )

        # =================================================
        # PPT
        # =================================================

        if ppt_file is not None:

            st.markdown("---")

            st.markdown(
                f"""
                **Excel:** `{excel_file.name}`  
                **PPT:** `{ppt_file.name}`  
                **Width:** `{width_col}`  
                **Height:** `{height_col}`
                """
            )

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
                    # Save
                    # -------------------------------------

                    output_path = (
                        "/tmp/PPT_SIZE_FIXED_V6.pptx"
                    )

                    prs.save(
                        output_path
                    )

                    # -------------------------------------
                    # Stats
                    # -------------------------------------

                    updated_count = len(
                        result_df[
                            result_df[
                                "Status"
                            ].str.contains(
                                "Updated",
                                na=False
                            )
                        ]
                    )

                    not_found_count = len(
                        result_df[
                            result_df[
                                "Status"
                            ].str.contains(
                                "Not Found",
                                na=False
                            )
                        ]
                    )

                    failed_count = len(
                        result_df[
                            result_df[
                                "Status"
                            ].str.contains(
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

                    # -------------------------------------
                    # Download
                    # -------------------------------------

                    with open(
                        output_path,
                        "rb"
                    ) as f:

                        output_data = f.read()

                    st.download_button(
                        label="⬇️ DOWNLOAD FIXED PPT",
                        data=output_data,
                        file_name="PPT_SIZE_FIXED_V6.pptx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "presentationml.presentation"
                        ),
                        type="primary",
                        use_container_width=True
                    )

                    # -------------------------------------
                    # Report
                    # -------------------------------------

                    st.markdown(
                        "### 📋 Processing Report"
                    )

                    st.dataframe(
                        result_df,
                        use_container_width=True,
                        height=550
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
# FOOTER
# =========================================================

st.markdown("---")

st.caption(
    "PPT SIZE FIXER V6 | Existing Size fields only | "
    "No automatic new Size box creation"
)
