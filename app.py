from io import BytesIO
import re

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt
import streamlit as st


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V2",
    page_icon="📊",
    layout="centered"
)

st.title("📊 PPT SIZE FIXER V2")

st.write(
    "Excel Width + Height → PPT Size"
)

st.info(
    "SAFE MODE: Existing Media Type, Qty, Remarks and "
    "other PPT elements are protected."
)

st.markdown("---")


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_excel = st.file_uploader(
    "1. Upload Master Excel File",
    type=["xlsx", "xls", "csv", "xlsm"],
    key="excel_uploader_v2"
)

uploaded_ppt = st.file_uploader(
    "2. Upload PowerPoint Presentation",
    type=["pptx"],
    key="ppt_uploader_v2"
)


# ============================================================
# NORMALIZE COLUMN NAME
# ============================================================

def normalize_column_name(value):

    value = str(value).strip().lower()

    value = value.replace("_", " ")
    value = value.replace("-", " ")

    value = re.sub(r"\s+", " ", value)

    return value


# ============================================================
# FIND COLUMN
# ============================================================

def find_column(df, possible_names):

    columns = list(df.columns)

    normalized_columns = {}

    for col in columns:

        normalized_columns[
            normalize_column_name(col)
        ] = col

    # --------------------------------------------------------
    # 1. EXACT MATCH
    # --------------------------------------------------------

    for name in possible_names:

        key = normalize_column_name(name)

        if key in normalized_columns:

            return normalized_columns[key]

    # --------------------------------------------------------
    # 2. NORMALIZED EXACT MATCH
    # --------------------------------------------------------

    for col in columns:

        col_norm = normalize_column_name(col)

        for name in possible_names:

            name_norm = normalize_column_name(name)

            if col_norm == name_norm:

                return col

    # --------------------------------------------------------
    # 3. SAFE PARTIAL MATCH
    # --------------------------------------------------------

    for col in columns:

        col_norm = normalize_column_name(col)

        for name in possible_names:

            name_norm = normalize_column_name(name)

            if (
                name_norm
                and name_norm in col_norm
            ):

                return col

    return None


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_safely(file_obj):

    name = file_obj.name.lower()

    if name.endswith(".csv"):

        return pd.read_csv(file_obj)

    xl_file = pd.ExcelFile(file_obj)

    if "Merged_Result" in xl_file.sheet_names:

        sheet_to_use = "Merged_Result"

    else:

        sheet_to_use = xl_file.sheet_names[0]

    return pd.read_excel(
        file_obj,
        sheet_name=sheet_to_use
    )


# ============================================================
# FIND WIDTH / HEIGHT COLUMNS
# ============================================================

def get_size_columns(df):

    width_names = [

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

    ]

    height_names = [

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

    ]

    width_column = find_column(
        df,
        width_names
    )

    height_column = find_column(
        df,
        height_names
    )

    return width_column, height_column


# ============================================================
# CLEAN NUMBER
# ============================================================

def clean_number(value):

    if value is None:

        return ""

    try:

        if pd.isna(value):

            return ""

    except Exception:

        pass

    text = str(value).strip()

    if not text:

        return ""

    if text.lower() == "nan":

        return ""

    text = text.replace(",", "")

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text
    )

    if not match:

        return ""

    number = float(
        match.group()
    )

    if number.is_integer():

        return str(
            int(number)
        )

    return (
        str(number)
        .rstrip("0")
        .rstrip(".")
    )


# ============================================================
# MAKE PPT SIZE
# IMPORTANT:
# PPT FORMAT = WIDTH X HEIGHT
# ============================================================

def make_size(width, height):

    width = clean_number(width)

    height = clean_number(height)

    if not width or not height:

        return ""

    return f"{width} X {height}"


# ============================================================
# GET SHAPE TEXT
# ============================================================

def get_shape_text(shape):

    try:

        if not shape.has_text_frame:

            return ""

        return (
            shape.text_frame.text
            .strip()
        )

    except Exception:

        return ""


# ============================================================
# GET ALL TEXT SHAPES
# ============================================================

def get_all_text_shapes(slide):

    result = []

    for shape in slide.shapes:

        try:

            text = get_shape_text(
                shape
            )

            if text:

                result.append(
                    (
                        shape,
                        text
                    )
                )

        except Exception:

            pass

    return result


# ============================================================
# FIND SIZE LABEL
# ============================================================

def find_size_label(slide):

    candidates = []

    for shape, text in get_all_text_shapes(
        slide
    ):

        normalized = (
            normalize_column_name(
                text
            )
        )

        if normalized in [

            "size",
            "size:",
            "size :-",
            "size -",
            "size :",

        ]:

            candidates.append(
                shape
            )

    if not candidates:

        return None

    # Prefer lower part of slide
    candidates.sort(
        key=lambda s: s.top,
        reverse=True
    )

    return candidates[0]


# ============================================================
# CHECK IF TEXT LOOKS LIKE SIZE
# ============================================================

def looks_like_size(text):

    if not text:

        return False

    text = str(
        text
    ).strip().lower()

    pattern = (
        r"^\s*"
        r"\d+(?:\.\d+)?"
        r"\s*"
        r"(x|×|\*)"
        r"\s*"
        r"\d+(?:\.\d+)?"
        r"(?:\s*(in|inch|inches|ft|feet))?"
        r"\s*$"
    )

    return bool(
        re.match(
            pattern,
            text
        )
    )


# ============================================================
# PROTECTED PPT FIELD
# ============================================================

def is_protected_shape_text(text):

    if not text:

        return False

    text_lower = (
        str(text)
        .strip()
        .lower()
    )

    protected_exact = [

        "media",
        "media:",
        "media type",
        "media type:",
        "qty",
        "qty:",
        "quantity",
        "quantity:",
        "remarks",
        "remarks:",
        "remark",
        "remark:",
        "outlet",
        "outlet name",
        "outlet name:",
        "address",
        "address:",
        "contact",
        "contact no",
        "contact no:",
        "sap",
        "sap code",
        "sap code:",

    ]

    if text_lower in protected_exact:

        return True

    protected_contains = [

        "media type",
        "contact no",
        "sap code",
        "outlet name",

    ]

    for word in protected_contains:

        if word in text_lower:

            return True

    return False


# ============================================================
# FIND EXISTING SIZE VALUE
# ============================================================

def find_size_value_near_label(
    slide,
    label_shape
):

    if label_shape is None:

        return None

    candidates = []

    label_right = (
        label_shape.left
        + label_shape.width
    )

    label_center_y = (
        label_shape.top
        + (
            label_shape.height
            / 2
        )
    )

    for shape, text in get_all_text_shapes(
        slide
    ):

        if shape == label_shape:

            continue

        # ----------------------------------------------------
        # NEVER TOUCH PROTECTED FIELDS
        # ----------------------------------------------------

        if is_protected_shape_text(
            text
        ):

            continue

        # ----------------------------------------------------
        # ONLY EXISTING SIZE-LIKE TEXT
        # ----------------------------------------------------

        if not looks_like_size(
            text
        ):

            continue

        # ----------------------------------------------------
        # POSITION CHECK
        # ----------------------------------------------------

        shape_center_y = (
            shape.top
            + (
                shape.height
                / 2
            )
        )

        vertical_distance = abs(
            shape_center_y
            - label_center_y
        )

        if vertical_distance > Pt(80):

            continue

        horizontal_distance = (
            shape.left
            - label_right
        )

        # Allow slightly overlapping boxes
        if horizontal_distance < -Pt(40):

            continue

        # Don't search too far
        if horizontal_distance > Pt(220):

            continue

        score = (
            abs(horizontal_distance)
            + vertical_distance
        )

        candidates.append(
            (
                score,
                shape
            )
        )

    if not candidates:

        return None

    candidates.sort(
        key=lambda x: x[0]
    )

    return candidates[0][1]


# ============================================================
# GET SHAPE STYLE
# ============================================================

def get_style_from_shape(shape):

    border_color = RGBColor(
        227,
        108,
        10
    )

    line_width = Pt(2)

    font_name = "Calibri"

    font_color = RGBColor(
        0,
        0,
        0
    )

    font_size = Pt(16)

    # --------------------------------------------------------
    # LINE
    # --------------------------------------------------------

    try:

        if shape.line:

            try:

                if (
                    shape.line.fill.type == 1
                    and shape.line.color.rgb
                ):

                    border_color = (
                        shape.line.color.rgb
                    )

            except Exception:

                pass

            try:

                if shape.line.width:

                    line_width = (
                        shape.line.width
                    )

            except Exception:

                pass

    except Exception:

        pass

    # --------------------------------------------------------
    # FONT
    # --------------------------------------------------------

    try:

        if shape.has_text_frame:

            for paragraph in (
                shape.text_frame.paragraphs
            ):

                for run in paragraph.runs:

                    try:

                        if run.font.name:

                            font_name = (
                                run.font.name
                            )

                    except Exception:

                        pass

                    try:

                        if run.font.size:

                            font_size = (
                                run.font.size
                            )

                    except Exception:

                        pass

                    try:

                        if (
                            run.font.color
                            and run.font.color.rgb
                        ):

                            font_color = (
                                run.font.color.rgb
                            )

                    except Exception:

                        pass

                    break

                break

    except Exception:

        pass

    return (
        border_color,
        line_width,
        font_name,
        font_color,
        font_size,
    )


# ============================================================
# UPDATE EXISTING SIZE BOX
# ============================================================

def update_existing_size_box(
    shape,
    final_text,
    style
):

    (
        border_color,
        line_width,
        font_name,
        font_color,
        font_size,
    ) = style

    try:

        # IMPORTANT:
        # We are NOT deleting the shape.
        # Only changing its text.

        tf = shape.text_frame

        tf.clear()

        tf.word_wrap = False

        tf.vertical_anchor = (
            MSO_ANCHOR.MIDDLE
        )

        tf.margin_top = Pt(1)
        tf.margin_bottom = Pt(1)
        tf.margin_left = Pt(1)
        tf.margin_right = Pt(1)

        paragraph = (
            tf.paragraphs[0]
        )

        paragraph.alignment = (
            PP_ALIGN.CENTER
        )

        run = (
            paragraph.add_run()
        )

        run.text = final_text

        run.font.name = font_name

        run.font.bold = True

        run.font.size = font_size

        try:

            run.font.color.rgb = (
                font_color
            )

        except Exception:

            pass

        try:

            shape.line.color.rgb = (
                border_color
            )

            shape.line.width = (
                line_width
            )

        except Exception:

            pass

        return True

    except Exception:

        return False


# ============================================================
# CREATE NEW SIZE BOX
# ============================================================

def create_size_box(
    slide,
    size_label_shape,
    final_text,
    style
):

    (
        border_color,
        line_width,
        font_name,
        font_color,
        font_size,
    ) = style

    # --------------------------------------------------------
    # LOCATION
    # --------------------------------------------------------

    if size_label_shape:

        box_left = (
            size_label_shape.left
            + size_label_shape.width
            + Pt(5)
        )

        box_top = (
            size_label_shape.top
        )

        box_height = (
            size_label_shape.height
        )

        box_width = Pt(120)

    else:

        # SAFE FALLBACK
        box_left = Pt(300)
        box_top = Pt(430)

        box_width = Pt(120)
        box_height = Pt(30)

    # --------------------------------------------------------
    # CREATE
    # --------------------------------------------------------

    new_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        box_left,
        box_top,
        box_width,
        box_height
    )

    new_box.fill.background()

    new_box.line.color.rgb = (
        border_color
    )

    new_box.line.width = (
        line_width
    )

    # --------------------------------------------------------
    # TEXT
    # --------------------------------------------------------

    tf = new_box.text_frame

    tf.word_wrap = False

    tf.vertical_anchor = (
        MSO_ANCHOR.MIDDLE
    )

    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    tf.margin_left = Pt(1)
    tf.margin_right = Pt(1)

    paragraph = (
        tf.paragraphs[0]
    )

    paragraph.alignment = (
        PP_ALIGN.CENTER
    )

    run = (
        paragraph.add_run()
    )

    run.text = final_text

    run.font.name = font_name

    run.font.bold = True

    run.font.size = font_size

    try:

        run.font.color.rgb = (
            font_color
        )

    except Exception:

        pass

    return new_box


# ============================================================
# PROCESS ONE SLIDE
# ============================================================

def process_slide(
    slide,
    size_text
):

    result = {

        "status": "Skipped",

        "reason": "",

    }

    # --------------------------------------------------------
    # NO SIZE
    # --------------------------------------------------------

    if not size_text:

        result["reason"] = (
            "Excel Width/Height is blank"
        )

        return result

    # --------------------------------------------------------
    # FIND SIZE LABEL
    # --------------------------------------------------------

    size_label = (
        find_size_label(
            slide
        )
    )

    # --------------------------------------------------------
    # FIND EXISTING SIZE VALUE
    # --------------------------------------------------------

    size_value = (
        find_size_value_near_label(
            slide,
            size_label
        )
    )

    # --------------------------------------------------------
    # UPDATE EXISTING VALUE
    # --------------------------------------------------------

    if size_value:

        style = (
            get_style_from_shape(
                size_value
            )
        )

        success = (
            update_existing_size_box(
                size_value,
                size_text,
                style
            )
        )

        if success:

            result["status"] = (
                "Updated Existing Size"
            )

            result["reason"] = (
                "Existing Size value updated"
            )

            return result

    # --------------------------------------------------------
    # CREATE SIZE BOX
    # --------------------------------------------------------

    if size_label:

        style = (
            get_style_from_shape(
                size_label
            )
        )

        create_size_box(
            slide,
            size_label,
            size_text,
            style
        )

        result["status"] = (
            "Created Size Box"
        )

        result["reason"] = (
            "Size value not found; "
            "new Size box created"
        )

        return result

    # --------------------------------------------------------
    # SAFE SKIP
    # --------------------------------------------------------

    result["status"] = (
        "Skipped - Safe Mode"
    )

    result["reason"] = (
        "Size label could not be safely identified"
    )

    return result


# ============================================================
# MAIN APPLICATION
# ============================================================

if uploaded_excel and uploaded_ppt:

    if st.button(
        "🚀 Process & Sync Files",
        type="primary"
    ):

        with st.spinner(
            "Processing PPT safely... Please wait..."
        ):

            try:

                # ====================================================
                # READ EXCEL
                # ====================================================

                df = (
                    read_excel_safely(
                        uploaded_excel
                    )
                )

                if df.empty:

                    st.error(
                        "❌ Excel file is empty."
                    )

                    st.stop()

                # ====================================================
                # FIND WIDTH / HEIGHT
                # ====================================================

                (
                    width_column,
                    height_column
                ) = get_size_columns(
                    df
                )

                # ====================================================
                # WIDTH CHECK
                # ====================================================

                if not width_column:

                    st.error(
                        "❌ Width column not found."
                    )

                    st.write(
                        "Available Excel columns:"
                    )

                    st.write(
                        list(df.columns)
                    )

                    st.stop()

                # ====================================================
                # HEIGHT CHECK
                # ====================================================

                if not height_column:

                    st.error(
                        "❌ Height column not found."
                    )

                    st.write(
                        "Available Excel columns:"
                    )

                    st.write(
                        list(df.columns)
                    )

                    st.stop()

                # ====================================================
                # SHOW DETECTED COLUMNS
                # ====================================================

                st.success(
                    f"✅ Width Column: {width_column}"
                )

                st.success(
                    f"✅ Height Column: {height_column}"
                )

                # ====================================================
                # LOAD PPT
                # ====================================================

                ppt_bytes = (
                    uploaded_ppt.read()
                )

                prs = Presentation(
                    BytesIO(
                        ppt_bytes
                    )
                )

                total_slides = len(
                    prs.slides
                )

                total_excel_rows = len(
                    df
                )

                process_count = min(
                    total_slides,
                    total_excel_rows
                )

                # ====================================================
                # COUNTERS
                # ====================================================

                updated_count = 0

                created_count = 0

                skipped_count = 0

                # ====================================================
                # REPORT
                # ====================================================

                report = []

                # ====================================================
                # PROCESS SLIDES
                # ====================================================

                progress_bar = st.progress(
                    0
                )

                for i in range(
                    process_count
                ):

                    slide = (
                        prs.slides[i]
                    )

                    # ----------------------------------------------
                    # Excel row
                    # ----------------------------------------------

                    width = (
                        df.iloc[i][
                            width_column
                        ]
                    )

                    height = (
                        df.iloc[i][
                            height_column
                        ]
                    )

                    # ----------------------------------------------
                    # Create Width X Height
                    # ----------------------------------------------

                    size_text = (
                        make_size(
                            width,
                            height
                        )
                    )

                    # ----------------------------------------------
                    # Process slide
                    # ----------------------------------------------

                    result = (
                        process_slide(
                            slide,
                            size_text
                        )
                    )

                    status = (
                        result["status"]
                    )

                    reason = (
                        result["reason"]
                    )

                    # ----------------------------------------------
                    # Counter
                    # ----------------------------------------------

                    if (
                        status
                        == "Updated Existing Size"
                    ):

                        updated_count += 1

                    elif (
                        status
                        == "Created Size Box"
                    ):

                        created_count += 1

                    else:

                        skipped_count += 1

                    # ----------------------------------------------
                    # Report
                    # ----------------------------------------------

                    report.append(
                        {

                            "Slide":
                                i + 1,

                            "Excel Row":
                                i + 2,

                            "Width":
                                clean_number(
                                    width
                                ),

                            "Height":
                                clean_number(
                                    height
                                ),

                            "PPT Size":
                                size_text,

                            "Status":
                                status,

                            "Reason":
                                reason,

                        }
                    )

                    # ----------------------------------------------
                    # Progress
                    # ----------------------------------------------

                    progress_bar.progress(
                        int(
                            (
                                (i + 1)
                                / process_count
                            )
                            * 100
                        )
                    )

                # ====================================================
                # EXTRA PPT SLIDES
                # ====================================================

                if (
                    total_slides
                    > total_excel_rows
                ):

                    for i in range(
                        total_excel_rows,
                        total_slides
                    ):

                        skipped_count += 1

                        report.append(
                            {

                                "Slide":
                                    i + 1,

                                "Excel Row":
                                    "",

                                "Width":
                                    "",

                                "Height":
                                    "",

                                "PPT Size":
                                    "",

                                "Status":
                                    "Skipped",

                                "Reason":
                                    "No matching Excel row",

                            }
                        )

                # ====================================================
                # SAVE PPT
                # ====================================================

                output_ppt_buffer = (
                    BytesIO()
                )

                prs.save(
                    output_ppt_buffer
                )

                output_ppt_buffer.seek(
                    0
                )

                # ====================================================
                # SUCCESS
                # ====================================================

                st.success(
                    "🎉 PPT processing completed successfully."
                )

                # ====================================================
                # METRICS
                # ====================================================

                col1, col2, col3 = (
                    st.columns(3)
                )

                with col1:

                    st.metric(
                        "Updated",
                        updated_count
                    )

                with col2:

                    st.metric(
                        "Created",
                        created_count
                    )

                with col3:

                    st.metric(
                        "Skipped",
                        skipped_count
                    )

                # ====================================================
                # REPORT DATAFRAME
                # ====================================================

                report_df = pd.DataFrame(
                    report
                )

                st.subheader(
                    "📋 Processing Report"
                )

                st.dataframe(
                    report_df,
                    use_container_width=True
                )

                # ====================================================
                # DOWNLOAD PPT
                # ====================================================

                st.download_button(
                    label=(
                        "📥 Download Updated PPT"
                    ),

                    data=(
                        output_ppt_buffer
                    ),

                    file_name=(
                        "Updated_Presentation_V2.pptx"
                    ),

                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    )
                )

                # ====================================================
                # DOWNLOAD REPORT
                # ====================================================

                report_buffer = (
                    BytesIO()
                )

                with pd.ExcelWriter(
                    report_buffer,
                    engine="openpyxl"
                ) as writer:

                    report_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Processing Report"
                    )

                report_buffer.seek(
                    0
                )

                st.download_button(
                    label=(
                        "📊 Download Processing Report"
                    ),

                    data=(
                        report_buffer
                    ),

                    file_name=(
                        "PPT_Size_Processing_Report.xlsx"
                    ),

                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

            except Exception as e:

                st.error(
                    f"❌ Error Occurred: {e}"
                )

                st.exception(e)
