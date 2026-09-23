import streamlit as st
import pandas as pd
import re
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from copy import deepcopy


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V5",
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
    color: #666;
    margin-bottom: 25px;
}

.success-box {
    padding: 12px;
    border-radius: 8px;
    background: #e8f5e9;
    border: 1px solid #81c784;
}

.warning-box {
    padding: 12px;
    border-radius: 8px;
    background: #fff8e1;
    border: 1px solid #ffcc80;
}

.error-box {
    padding: 12px;
    border-radius: 8px;
    background: #ffebee;
    border: 1px solid #ef9a9a;
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# TITLE
# =========================================================

st.markdown(
    '<div class="main-title">📐 PPT SIZE FIXER V5</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="sub-title">Excel Width × Height ko existing PowerPoint Size field mein update karega.</div>',
    unsafe_allow_html=True
)


# =========================================================
# CONSTANTS
# =========================================================

PROTECTED_WORDS = {
    "media",
    "media type",
    "qty",
    "quantity",
    "remarks",
    "remark",
    "outlet",
    "outlet name",
    "address",
    "contact",
    "contact no",
    "contact number",
    "phone",
    "mobile",
    "sap",
    "sap code",
    "outlet code",
    "district",
}


WIDTH_ALIASES = [
    "width",
    "width inches",
    "width inch",
    "width inches",
    "width (inches)",
    "width (inch)",
    "w",
    "w inches",
    "w inch",
    "w (inches)",
    "w (inch)",
    "board width",
    "size width",
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
    "size height",
]


# =========================================================
# NORMALIZATION
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


def clean_number(value):
    """
    Excel values ko clean karke string number return karta hai.
    Examples:
        180       -> 180
        180.0     -> 180
        "180"     -> 180
        "180.00"  -> 180
        " 180 "   -> 180
    """

    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text == "":
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
# PROTECTED FIELD DETECTION
# =========================================================

def is_protected_text(text):
    """
    Ye decide karta hai ki koi shape protected field hai ya nahi.
    """

    n = normalize_text(text)

    if not n:
        return False

    # Exact protected values
    if n in PROTECTED_WORDS:
        return True

    # Starts with protected field
    for word in PROTECTED_WORDS:

        if n.startswith(word + ":"):
            return True

        if n.startswith(word + " :-"):
            return True

        if n.startswith(word + " -"):
            return True

        if n.startswith(word + " :"):
            return True

    return False


# =========================================================
# STANDALONE SIZE DETECTION
# =========================================================

def is_standalone_size(text):
    """
    Detect:
        180 X 36
        180x36
        180 X 48
        180 × 36

    But does NOT detect:
        Size :- 180 X 48

    because inline size ko separately handle karna hai.
    """

    if text is None:
        return False

    raw = str(text).strip()

    if not raw:
        return False

    # Inline Size field ko standalone nahi maanenge
    n = normalize_text(raw)

    if n.startswith("size"):
        return False

    pattern = r"^\s*\d+(?:\.\d+)?\s*[x×*]\s*\d+(?:\.\d+)?\s*$"

    return bool(re.match(pattern, raw, flags=re.IGNORECASE))


# =========================================================
# INLINE SIZE DETECTION
# =========================================================

def replace_inline_size(text, width, height):
    """
    Examples:

    Size :- 180 X 48
    ->
    Size :- 200 X 36

    Size: 180X48
    ->
    Size: 200 X 36
    """

    if text is None:
        return None, False

    raw = str(text)

    pattern = (
        r"(\bsize\b"
        r"(?:\s*[:\-]"
        r"|"
        r"\s*:-"
        r")"
        r"\s*)"
        r"(\d+(?:\.\d+)?\s*[x×*]\s*\d+(?:\.\d+)?)"
    )

    replacement = rf"\g<1>{width} X {height}"

    new_text, count = re.subn(
        pattern,
        replacement,
        raw,
        flags=re.IGNORECASE
    )

    return new_text, count > 0


def is_inline_size_field(text):
    if text is None:
        return False

    n = normalize_text(text)

    if not n.startswith("size"):
        return False

    return bool(
        re.search(
            r"\d+(?:\.\d+)?\s*[x×*]\s*\d+(?:\.\d+)?",
            n,
            flags=re.IGNORECASE
        )
    )


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


def set_shape_text_preserve_style(shape, new_text):
    """
    Existing shape ke andar text replace karta hai.
    Shape ko delete/create nahi karta.
    """

    try:
        if not shape.has_text_frame:
            return False

        tf = shape.text_frame

        # Existing paragraph/run formatting ko preserve karne ki koshish
        if len(tf.paragraphs) > 0:

            paragraph = tf.paragraphs[0]

            if len(paragraph.runs) > 0:

                # First run mein text set
                paragraph.runs[0].text = str(new_text)

                # Extra runs clear
                for run in paragraph.runs[1:]:
                    run.text = ""

                # Extra paragraphs clear
                for p in tf.paragraphs[1:]:
                    for run in p.runs:
                        run.text = ""

                return True

        # Fallback
        tf.text = str(new_text)

        return True

    except Exception:

        try:
            shape.text = str(new_text)
            return True

        except Exception:
            return False


# =========================================================
# SHAPE GEOMETRY
# =========================================================

def shape_geometry(shape):

    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height

    right = left + width
    bottom = top + height

    center_x = left + width / 2
    center_y = top + height / 2

    return {
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "right": right,
        "bottom": bottom,
        "center_x": center_x,
        "center_y": center_y,
    }


def vertical_distance(a, b):
    return abs(a["center_y"] - b["center_y"])


def horizontal_distance(a, b):
    return abs(a["center_x"] - b["center_x"])


def boxes_overlap(a, b):

    if a["right"] <= b["left"]:
        return False

    if b["right"] <= a["left"]:
        return False

    if a["bottom"] <= b["top"]:
        return False

    if b["bottom"] <= a["top"]:
        return False

    return True


# =========================================================
# COLLECT TEXT SHAPES
# =========================================================

def collect_text_shapes(slide):

    result = []

    for index, shape in enumerate(slide.shapes):

        text = get_shape_text(shape)

        if not text.strip():
            continue

        geo = shape_geometry(shape)

        result.append({
            "index": index,
            "shape": shape,
            "text": text,
            "norm": normalize_text(text),
            "geo": geo,
        })

    return result


# =========================================================
# FIND SIZE LABEL
# =========================================================

def find_size_label(items):

    exact_candidates = []

    for item in items:

        n = item["norm"]

        if n in {
            "size",
            "size:",
            "size :",
            "size :-",
            "size -",
        }:

            exact_candidates.append(item)

    if exact_candidates:
        return exact_candidates[0]

    # fallback
    for item in items:

        n = item["norm"]

        if n.startswith("size"):

            if not is_inline_size_field(item["text"]):

                return item

    return None


# =========================================================
# FIND QTY
# =========================================================

def find_qty_shape(items):

    candidates = []

    for item in items:

        n = item["norm"]

        if n in {
            "qty",
            "qty:",
            "qty :",
            "quantity",
            "quantity:",
            "quantity :",
        }:

            candidates.append(item)

    if candidates:
        return candidates[0]

    return None


# =========================================================
# FIND INLINE SIZE
# =========================================================

def find_inline_size_field(items):

    for item in items:

        if is_inline_size_field(item["text"]):

            return item

    return None


# =========================================================
# FIND X SHAPE
# =========================================================

def is_x_shape(text):

    if text is None:
        return False

    n = normalize_text(text)

    return n in {
        "x",
        "×",
        "*",
    }


def find_x_candidates(items, size_label):

    label_geo = size_label["geo"]

    candidates = []

    for item in items:

        if item is size_label:
            continue

        if not is_x_shape(item["text"]):
            continue

        geo = item["geo"]

        # Same row
        if vertical_distance(geo, label_geo) > Inches(0.45):
            continue

        # X should be inside/near Size label area
        horizontal_near = (
            geo["center_x"] >= label_geo["left"] - Inches(0.20)
            and
            geo["center_x"] <= label_geo["right"] + Inches(0.20)
        )

        if not horizontal_near:
            continue

        candidates.append(item)

    candidates.sort(
        key=lambda x: horizontal_distance(
            x["geo"],
            label_geo
        )
    )

    return candidates


# =========================================================
# NUMERIC SHAPE
# =========================================================

def is_numeric_text(text):

    if text is None:
        return False

    raw = str(text).strip()

    if not raw:
        return False

    return bool(
        re.match(
            r"^\d+(?:\.\d+)?$",
            raw
        )
    )


# =========================================================
# FIND SIZE COMPONENTS
# =========================================================

def find_size_components(items, size_label):

    """
    Actual template mein:

        Size:   15   x   3       Qty: 1

    Size label ke area ke andar:
        15
        x
        3

    Ye function exactly in 3 shapes ko find karega.
    """

    label_geo = size_label["geo"]

    x_candidates = find_x_candidates(
        items,
        size_label
    )

    if not x_candidates:
        return None

    # Sabse suitable X
    x_item = x_candidates[0]

    x_geo = x_item["geo"]

    left_candidates = []
    right_candidates = []

    for item in items:

        if item is size_label:
            continue

        if item is x_item:
            continue

        text = item["text"]

        if not is_numeric_text(text):
            continue

        # Protected numeric fields avoid
        if is_protected_text(text):
            continue

        geo = item["geo"]

        # Same row
        if vertical_distance(geo, x_geo) > Inches(0.35):
            continue

        # IMPORTANT:
        # Numeric field Size label ke horizontal area ke andar hona chahiye.
        #
        # Isse Qty value "1" accidentally select nahi hogi.

        if geo["center_x"] < label_geo["left"] - Inches(0.10):
            continue

        if geo["center_x"] > label_geo["right"] + Inches(0.10):
            continue

        if geo["center_x"] < x_geo["center_x"]:

            left_candidates.append(item)

        elif geo["center_x"] > x_geo["center_x"]:

            right_candidates.append(item)

    if not left_candidates or not right_candidates:
        return None

    # X ke nearest numeric shape
    left_candidates.sort(
        key=lambda x: abs(
            x["geo"]["center_x"] -
            x_geo["center_x"]
        )
    )

    right_candidates.sort(
        key=lambda x: abs(
            x["geo"]["center_x"] -
            x_geo["center_x"]
        )
    )

    width_item = left_candidates[0]
    height_item = right_candidates[0]

    return {
        "width": width_item,
        "x": x_item,
        "height": height_item,
    }


# =========================================================
# DELETE SHAPE SAFELY
# =========================================================

def delete_shape(shape):

    try:

        sp = shape._element

        parent = sp.getparent()

        parent.remove(sp)

        return True

    except Exception:

        return False


# =========================================================
# FIND DUPLICATE STANDALONE SIZE BOXES
# =========================================================

def find_duplicate_size_boxes(
    items,
    size_label,
    size_components
):

    """
    Agar actual template mein:

        15 x 3

    ke saath galat:

        180 X 36

    Qty ke paas pada hai,

    to us duplicate box ko remove karega.
    """

    if not size_components:
        return []

    label_geo = size_label["geo"]

    qty_item = find_qty_shape(items)

    qty_geo = None

    if qty_item:
        qty_geo = qty_item["geo"]

    component_shapes = {
        id(size_components["width"]["shape"]),
        id(size_components["x"]["shape"]),
        id(size_components["height"]["shape"]),
    }

    duplicates = []

    for item in items:

        shape = item["shape"]

        if id(shape) in component_shapes:
            continue

        text = item["text"]

        if not is_standalone_size(text):
            continue

        geo = item["geo"]

        # Same horizontal row
        if vertical_distance(
            geo,
            label_geo
        ) > Inches(0.45):

            continue

        should_delete = False

        # -------------------------------------------------
        # CASE 1:
        # Size label ke right side mein stray box
        # -------------------------------------------------

        if geo["left"] >= label_geo["right"] - Inches(0.15):

            should_delete = True

        # -------------------------------------------------
        # CASE 2:
        # Qty box ke saath overlap
        # -------------------------------------------------

        if qty_geo is not None:

            if boxes_overlap(
                geo,
                qty_geo
            ):

                should_delete = True

        # -------------------------------------------------
        # CASE 3:
        # Size label aur Qty ke beech
        # -------------------------------------------------

        if qty_geo is not None:

            if (
                geo["center_x"] >
                label_geo["center_x"]
                and
                geo["center_x"] <
                qty_geo["left"] + Inches(0.25)
            ):

                should_delete = True

        if should_delete:

            duplicates.append(item)

    return duplicates


# =========================================================
# UPDATE INLINE SIZE
# =========================================================

def update_inline_size(
    item,
    width,
    height
):

    old_text = item["text"]

    new_text, changed = replace_inline_size(
        old_text,
        width,
        height
    )

    if not changed:
        return False, old_text, old_text

    success = set_shape_text_preserve_style(
        item["shape"],
        new_text
    )

    return success, old_text, new_text


# =========================================================
# UPDATE SEPARATE SIZE COMPONENTS
# =========================================================

def update_separate_components(
    components,
    width,
    height
):

    width_item = components["width"]
    x_item = components["x"]
    height_item = components["height"]

    old_width = width_item["text"]
    old_x = x_item["text"]
    old_height = height_item["text"]

    ok1 = set_shape_text_preserve_style(
        width_item["shape"],
        str(width)
    )

    ok2 = set_shape_text_preserve_style(
        x_item["shape"],
        "X"
    )

    ok3 = set_shape_text_preserve_style(
        height_item["shape"],
        str(height)
    )

    return (
        ok1 and ok2 and ok3,
        old_width,
        old_x,
        old_height
    )


# =========================================================
# FIND EXISTING COMBINED SIZE BOX
# =========================================================

def find_existing_combined_size_box(
    items,
    size_label
):

    label_geo = size_label["geo"]

    candidates = []

    for item in items:

        if not is_standalone_size(
            item["text"]
        ):
            continue

        geo = item["geo"]

        if vertical_distance(
            geo,
            label_geo
        ) > Inches(0.50):

            continue

        # Size label ke aas-paas
        distance = abs(
            geo["center_x"] -
            label_geo["center_x"]
        )

        if distance <= Inches(3.0):

            candidates.append(item)

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: abs(
            x["geo"]["center_x"] -
            label_geo["center_x"]
        )
    )

    return candidates[0]


# =========================================================
# CREATE SIZE BOX
# =========================================================

def create_size_box(
    slide,
    size_label,
    qty_item,
    width,
    height
):

    label_geo = size_label["geo"]

    text = f"{width} X {height}"

    # Default position
    left = label_geo["right"] + Inches(0.05)

    top = label_geo["top"]

    box_width = Inches(1.15)
    box_height = Inches(0.35)

    # Qty ke paas nahi jana
    if qty_item is not None:

        qty_geo = qty_item["geo"]

        # Agar right side mein enough space nahi hai,
        # label ke andar lower safe position try karein.
        if left + box_width > qty_geo["left"] - Inches(0.05):

            left = label_geo["left"]

            top = label_geo["bottom"] + Inches(0.02)

    # Slide boundaries
    slide_width = slide.part.presentation.slide_width
    slide_height = slide.part.presentation.slide_height

    if left + box_width > slide_width:

        left = max(
            Inches(0.1),
            slide_width - box_width - Inches(0.1)
        )

    if top + box_height > slide_height:

        top = max(
            Inches(0.1),
            slide_height - box_height - Inches(0.1)
        )

    new_shape = slide.shapes.add_textbox(
        left,
        top,
        box_width,
        box_height
    )

    new_shape.text = text

    try:

        tf = new_shape.text_frame

        for paragraph in tf.paragraphs:

            for run in paragraph.runs:

                run.font.size = Pt(12)

    except Exception:
        pass

    return new_shape


# =========================================================
# EXCEL READER
# =========================================================

def read_excel_file(uploaded_file):

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):

        return pd.read_csv(
            uploaded_file
        )

    if filename.endswith(".xls"):

        return pd.read_excel(
            uploaded_file,
            engine="xlrd"
        )

    if filename.endswith(".xlsm"):

        return pd.read_excel(
            uploaded_file,
            sheet_name=None,
            engine="openpyxl"
        )

    # XLSX
    return pd.read_excel(
        uploaded_file,
        sheet_name=None,
        engine="openpyxl"
    )


# =========================================================
# SELECT SHEET
# =========================================================

def get_excel_dataframe(uploaded_file):

    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):

        df = pd.read_csv(
            uploaded_file
        )

        return df, "CSV"

    sheets = read_excel_file(
        uploaded_file
    )

    if not isinstance(sheets, dict):

        return sheets, "Sheet1"

    # Preferred sheet
    if "Merged_Result" in sheets:

        return (
            sheets["Merged_Result"],
            "Merged_Result"
        )

    # Case-insensitive search
    for sheet_name, df in sheets.items():

        if normalize_text(sheet_name) == "merged_result":

            return df, sheet_name

    # First non-empty sheet
    for sheet_name, df in sheets.items():

        if df is not None and not df.empty:

            return df, sheet_name

    # fallback
    first_sheet = list(sheets.keys())[0]

    return (
        sheets[first_sheet],
        first_sheet
    )


# =========================================================
# FIND EXCEL COLUMN
# =========================================================

def find_matching_column(
    columns,
    aliases
):

    normalized_columns = {}

    for col in columns:

        normalized_columns[
            normalize_column_name(col)
        ] = col

    # Exact
    for alias in aliases:

        alias_norm = normalize_column_name(
            alias
        )

        if alias_norm in normalized_columns:

            return normalized_columns[
                alias_norm
            ]

    # Slight flexible matching
    for col_norm, original_col in normalized_columns.items():

        for alias in aliases:

            alias_norm = normalize_column_name(
                alias
            )

            if (
                alias_norm in col_norm
                or
                col_norm in alias_norm
            ):

                return original_col

    return None


# =========================================================
# PROCESS PPT
# =========================================================

def process_ppt(
    ppt_bytes,
    excel_df,
    width_col,
    height_col
):

    import io

    prs = Presentation(
        io.BytesIO(ppt_bytes)
    )

    results = []

    total_slides = len(prs.slides)

    progress = st.progress(0)

    status_text = st.empty()

    # -----------------------------------------------------
    # Excel rows
    # -----------------------------------------------------

    excel_rows = []

    for idx, row in excel_df.iterrows():

        width = clean_number(
            row.get(width_col, "")
        )

        height = clean_number(
            row.get(height_col, "")
        )

        excel_rows.append({
            "row": idx + 2,
            "width": width,
            "height": height,
        })

    # -----------------------------------------------------
    # Process slides
    # -----------------------------------------------------

    for slide_no, slide in enumerate(
        prs.slides,
        start=1
    ):

        status_text.write(
            f"Processing slide {slide_no}/{total_slides}"
        )

        # -------------------------------------------------
        # IMPORTANT:
        # Each slide ka text ek hi baar collect
        # -------------------------------------------------

        items = collect_text_shapes(
            slide
        )

        # -------------------------------------------------
        # Excel row
        #
        # Slide 1 -> Excel row 1
        # Slide 2 -> Excel row 2
        #
        # Agar rows kam hain -> skip
        # -------------------------------------------------

        excel_index = slide_no - 1

        if excel_index >= len(excel_rows):

            results.append({
                "Slide": slide_no,
                "Status": "Skipped - Excel row not available",
                "Width": "",
                "Height": "",
                "Action": "",
            })

            progress.progress(
                slide_no / total_slides
            )

            continue

        row_data = excel_rows[
            excel_index
        ]

        width = row_data["width"]
        height = row_data["height"]

        if width == "" or height == "":

            results.append({
                "Slide": slide_no,
                "Status": "Skipped - Width/Height missing",
                "Width": width,
                "Height": height,
                "Action": "",
            })

            progress.progress(
                slide_no / total_slides
            )

            continue

        # =================================================
        # STEP 1
        # INLINE SIZE
        # =================================================

        inline_item = find_inline_size_field(
            items
        )

        if inline_item is not None:

            success, old_text, new_text = update_inline_size(
                inline_item,
                width,
                height
            )

            if success:

                results.append({
                    "Slide": slide_no,
                    "Status": "Updated Inline Size",
                    "Width": width,
                    "Height": height,
                    "Action": (
                        f"{old_text} -> {new_text}"
                    ),
                })

            else:

                results.append({
                    "Slide": slide_no,
                    "Status": "Inline Size Found - Update Failed",
                    "Width": width,
                    "Height": height,
                    "Action": old_text,
                })

            progress.progress(
                slide_no / total_slides
            )

            continue

        # =================================================
        # STEP 2
        # FIND SIZE LABEL
        # =================================================

        size_label = find_size_label(
            items
        )

        if size_label is None:

            results.append({
                "Slide": slide_no,
                "Status": "Skipped - Size label not found",
                "Width": width,
                "Height": height,
                "Action": "",
            })

            progress.progress(
                slide_no / total_slides
            )

            continue

        # =================================================
        # STEP 3
        # FIND SEPARATE COMPONENTS
        # =================================================

        components = find_size_components(
            items,
            size_label
        )

        if components is not None:

            # ---------------------------------------------
            # Update actual existing 3 fields
            # ---------------------------------------------

            success, old_w, old_x, old_h = (
                update_separate_components(
                    components,
                    width,
                    height
                )
            )

            # ---------------------------------------------
            # Remove wrong duplicate combined box
            # ---------------------------------------------

            duplicate_boxes = (
                find_duplicate_size_boxes(
                    items,
                    size_label,
                    components
                )
            )

            deleted_count = 0

            for duplicate in duplicate_boxes:

                if delete_shape(
                    duplicate["shape"]
                ):

                    deleted_count += 1

            if success:

                action = (
                    f"Existing Size fields: "
                    f"{old_w} X {old_h} -> "
                    f"{width} X {height}"
                )

                if deleted_count:

                    action += (
                        f"; Removed {deleted_count} duplicate size box"
                    )

                results.append({
                    "Slide": slide_no,
                    "Status": "Updated Existing Size Fields",
                    "Width": width,
                    "Height": height,
                    "Action": action,
                })

            else:

                results.append({
                    "Slide": slide_no,
                    "Status": "Size fields found but update failed",
                    "Width": width,
                    "Height": height,
                    "Action": "",
                })

            progress.progress(
                slide_no / total_slides
            )

            continue

        # =================================================
        # STEP 4
        # EXISTING COMBINED SIZE BOX
        # =================================================

        combined_item = find_existing_combined_size_box(
            items,
            size_label
        )

        if combined_item is not None:

            old_text = combined_item["text"]

            new_text = (
                f"{width} X {height}"
            )

            success = set_shape_text_preserve_style(
                combined_item["shape"],
                new_text
            )

            if success:

                results.append({
                    "Slide": slide_no,
                    "Status": "Updated Existing Size Box",
                    "Width": width,
                    "Height": height,
                    "Action": (
                        f"{old_text} -> {new_text}"
                    ),
                })

            else:

                results.append({
                    "Slide": slide_no,
                    "Status": "Existing Size Box Update Failed",
                    "Width": width,
                    "Height": height,
                    "Action": old_text,
                })

            progress.progress(
                slide_no / total_slides
            )

            continue

        # =================================================
        # STEP 5
        # ONLY IF NOTHING EXISTS -> CREATE
        # =================================================

        qty_item = find_qty_shape(
            items
        )

        new_shape = create_size_box(
            slide,
            size_label,
            qty_item,
            width,
            height
        )

        results.append({
            "Slide": slide_no,
            "Status": "Created New Size Box",
            "Width": width,
            "Height": height,
            "Action": (
                f"Created {width} X {height}"
            ),
        })

        progress.progress(
            slide_no / total_slides
        )

    progress.progress(1.0)

    status_text.success(
        f"Completed {total_slides} slides."
    )

    return prs, pd.DataFrame(results)


# =========================================================
# UI - FILE UPLOAD
# =========================================================

st.markdown("### 📊 Step 1 — Excel Master File")

excel_file = st.file_uploader(
    "Upload Excel",
    type=[
        "xlsx",
        "xls",
        "xlsm",
        "csv"
    ],
    key="excel_upload"
)


st.markdown("### 📄 Step 2 — PowerPoint File")

ppt_file = st.file_uploader(
    "Upload PPTX",
    type=["pptx"],
    key="ppt_upload"
)


# =========================================================
# MAIN
# =========================================================

if excel_file is not None:

    try:

        df, sheet_name = get_excel_dataframe(
            excel_file
        )

        st.success(
            f"Excel loaded: {excel_file.name} | Sheet: {sheet_name}"
        )

        st.write(
            f"Rows: **{len(df)}** | Columns: **{len(df.columns)}**"
        )

        # -------------------------------------------------
        # Column detection
        # -------------------------------------------------

        width_col = find_matching_column(
            df.columns,
            WIDTH_ALIASES
        )

        height_col = find_matching_column(
            df.columns,
            HEIGHT_ALIASES
        )

        col1, col2 = st.columns(2)

        with col1:

            st.markdown("**Width Column**")

            width_options = list(
                df.columns
            )

            default_width_index = 0

            if width_col in width_options:

                default_width_index = (
                    width_options.index(
                        width_col
                    )
                )

            selected_width_col = st.selectbox(
                "Select Width column",
                width_options,
                index=default_width_index,
                key="width_column"
            )

        with col2:

            st.markdown("**Height Column**")

            height_options = list(
                df.columns
            )

            default_height_index = 0

            if height_col in height_options:

                default_height_index = (
                    height_options.index(
                        height_col
                    )
                )

            selected_height_col = st.selectbox(
                "Select Height column",
                height_options,
                index=default_height_index,
                key="height_column"
            )

        # -------------------------------------------------
        # Preview
        # -------------------------------------------------

        st.markdown("### 🔎 Excel Preview")

        preview_cols = [
            selected_width_col,
            selected_height_col
        ]

        st.dataframe(
            df[preview_cols].head(10),
            use_container_width=True
        )

        # =================================================
        # PROCESS BUTTON
        # =================================================

        if ppt_file is not None:

            st.markdown("---")

            if st.button(
                "🚀 FIX PPT SIZE",
                type="primary",
                use_container_width=True
            ):

                try:

                    ppt_bytes = ppt_file.getvalue()

                    prs, result_df = process_ppt(
                        ppt_bytes,
                        df,
                        selected_width_col,
                        selected_height_col
                    )

                    # -------------------------------------
                    # Save output
                    # -------------------------------------

                    output_path = (
                        "/tmp/PPT_SIZE_FIXED_V5.pptx"
                    )

                    prs.save(
                        output_path
                    )

                    # -------------------------------------
                    # Result summary
                    # -------------------------------------

                    st.markdown("---")

                    st.markdown(
                        "### ✅ Processing Complete"
                    )

                    total = len(result_df)

                    updated = len(
                        result_df[
                            result_df["Status"].str.contains(
                                "Updated",
                                na=False
                            )
                        ]
                    )

                    created = len(
                        result_df[
                            result_df["Status"].str.contains(
                                "Created",
                                na=False
                            )
                        ]
                    )

                    skipped = len(
                        result_df[
                            result_df["Status"].str.contains(
                                "Skipped",
                                na=False
                            )
                        ]
                    )

                    c1, c2, c3, c4 = st.columns(4)

                    with c1:
                        st.metric(
                            "Total Slides",
                            total
                        )

                    with c2:
                        st.metric(
                            "Updated",
                            updated
                        )

                    with c3:
                        st.metric(
                            "Created",
                            created
                        )

                    with c4:
                        st.metric(
                            "Skipped",
                            skipped
                        )

                    # -------------------------------------
                    # Download PPT
                    # -------------------------------------

                    with open(
                        output_path,
                        "rb"
                    ) as f:

                        output_bytes = f.read()

                    st.download_button(
                        label="⬇️ DOWNLOAD FIXED PPT",
                        data=output_bytes,
                        file_name="PPT_SIZE_FIXED_V5.pptx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "presentationml.presentation"
                        ),
                        type="primary",
                        use_container_width=True
                    )

                    # -------------------------------------
                    # Result table
                    # -------------------------------------

                    st.markdown(
                        "### 📋 Processing Report"
                    )

                    st.dataframe(
                        result_df,
                        use_container_width=True,
                        height=500
                    )

                except Exception as e:

                    st.error(
                        f"Processing Error: {str(e)}"
                    )

                    st.exception(e)

        else:

            st.info(
                "PPTX file upload karo."
            )

    except Exception as e:

        st.error(
            f"Excel Error: {str(e)}"
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
    "PPT SIZE FIXER V5 — Existing Size fields ko update karta hai; "
    "Qty/Media/Remarks fields protected hain."
)
