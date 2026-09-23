from io import BytesIO
import re

import pandas as pd
import streamlit as st

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Pt


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="PPT SIZE FIXER V4",
    page_icon="📊",
    layout="centered"
)

st.title("📊 PPT SIZE FIXER V4")

st.write(
    "Excel Width + Height → PowerPoint Size"
)

st.info(
    "SAFE MODE: Existing Media Type, Qty, Remarks, SAP, "
    "Outlet, Address and other protected PPT elements are preserved."
)

st.markdown("---")


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_excel = st.file_uploader(
    "1. Upload Master Excel File",
    type=["xlsx", "xls", "csv", "xlsm"],
    key="excel_uploader_v4"
)

uploaded_ppt = st.file_uploader(
    "2. Upload PowerPoint Presentation",
    type=["pptx"],
    key="ppt_uploader_v4"
)


# ============================================================
# NORMALIZE
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip().lower()

    text = text.replace("_", " ")
    text = text.replace("-", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# FIND EXCEL COLUMN
# ============================================================

def find_column(df, possible_names):

    columns = list(df.columns)

    normalized_columns = {}

    for col in columns:
        normalized_columns[
            normalize_text(col)
        ] = col

    # Exact
    for name in possible_names:

        key = normalize_text(name)

        if key in normalized_columns:
            return normalized_columns[key]

    # Partial
    for col in columns:

        col_norm = normalize_text(col)

        for name in possible_names:

            name_norm = normalize_text(name)

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

    file_name = file_obj.name.lower()

    if file_name.endswith(".csv"):

        return pd.read_csv(file_obj)

    excel_file = pd.ExcelFile(file_obj)

    # Prefer Merged_Result
    if "Merged_Result" in excel_file.sheet_names:
        sheet_name = "Merged_Result"
    else:
        sheet_name = excel_file.sheet_names[0]

    return pd.read_excel(
        file_obj,
        sheet_name=sheet_name
    )


# ============================================================
# FIND WIDTH / HEIGHT
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
        "size width",

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
        "size height",

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

    try:
        number = float(match.group())
    except Exception:
        return ""

    if number.is_integer():
        return str(int(number))

    return (
        str(number)
        .rstrip("0")
        .rstrip(".")
    )


# ============================================================
# FINAL PPT SIZE
#
# IMPORTANT:
# WIDTH X HEIGHT
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

        return shape.text_frame.text.strip()

    except Exception:

        return ""


# ============================================================
# COLLECT TEXT SHAPES ONCE
#
# IMPORTANT FOR PERFORMANCE
# ============================================================

def collect_text_shapes(slide):

    result = []

    for shape in slide.shapes:

        try:

            text = get_shape_text(shape)

            if text:

                result.append({
                    "shape": shape,
                    "text": text,
                    "norm": normalize_text(text),
                    "left": shape.left,
                    "top": shape.top,
                    "width": shape.width,
                    "height": shape.height
                })

        except Exception:
            continue

    return result


# ============================================================
# PROTECTED FIELD
# ============================================================

def is_protected_text(text):

    if not text:
        return False

    t = normalize_text(text)

    protected_exact = {

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
        "outlet:",
        "outlet name",
        "outlet name:",

        "address",
        "address:",

        "contact",
        "contact:",
        "contact no",
        "contact no:",

        "sap",
        "sap:",
        "sap code",
        "sap code:",

    }

    if t in protected_exact:
        return True

    protected_words = [
        "media type",
        "contact no",
        "sap code",
        "outlet name"
    ]

    for word in protected_words:

        if word in t:
            return True

    return False


# ============================================================
# SIZE PATTERN
# ============================================================

SIZE_PATTERN = re.compile(
    r"""
    (?P<w>\d+(?:\.\d+)?)
    \s*
    [xX×*]
    \s*
    (?P<h>\d+(?:\.\d+)?)
    """,
    re.VERBOSE
)


# ============================================================
# STANDALONE SIZE
#
# Example:
# 180 X 36
# 240X36
# 96 × 48
# ============================================================

def is_standalone_size(text):

    if not text:
        return False

    text = str(text).strip()

    pattern = re.compile(
        r"^\s*"
        r"\d+(?:\.\d+)?"
        r"\s*[xX×*]\s*"
        r"\d+(?:\.\d+)?"
        r"(?:\s*(?:in|inch|inches|ft|feet))?"
        r"\s*$",
        re.IGNORECASE
    )

    return bool(
        pattern.match(text)
    )


# ============================================================
# INLINE SIZE FIELD
#
# Examples:
# Size :- 180X48
# Size: 180 X 48
# Size - 180X48
# ============================================================

def is_inline_size_field(text):

    if not text:
        return False

    t = normalize_text(text)

    if "size" not in t:
        return False

    return bool(
        SIZE_PATTERN.search(text)
    )


# ============================================================
# FIND SIZE LABEL
# ============================================================

def find_size_label(text_shapes):

    candidates = []

    for item in text_shapes:

        text = item["text"]
        norm = item["norm"]

        if norm in [
            "size",
            "size:",
            "size :",
            "size :-",
            "size -"
        ]:

            candidates.append(item)

    if not candidates:
        return None

    # Prefer lower size field
    candidates.sort(
        key=lambda x: x["top"],
        reverse=True
    )

    return candidates[0]


# ============================================================
# FIND INLINE SIZE FIELD
# ============================================================

def find_inline_size_field(text_shapes):

    candidates = []

    for item in text_shapes:

        text = item["text"]

        if is_inline_size_field(text):

            # Never touch protected fields
            if is_protected_text(text):
                continue

            candidates.append(item)

    if not candidates:
        return None

    # Prefer the lower size-related field
    candidates.sort(
        key=lambda x: x["top"],
        reverse=True
    )

    return candidates[0]


# ============================================================
# FIND STANDALONE SIZE BOX
#
# We search near the Size label.
# ============================================================

def find_standalone_size_box(
    text_shapes,
    size_label
):

    if size_label is None:
        return None

    label = size_label["shape"]

    label_left = label.left
    label_right = (
        label.left
        + label.width
    )

    label_center_y = (
        label.top
        + label.height / 2
    )

    candidates = []

    for item in text_shapes:

        shape = item["shape"]
        text = item["text"]

        if shape == label:
            continue

        if is_protected_text(text):
            continue

        if not is_standalone_size(text):
            continue

        shape_center_y = (
            shape.top
            + shape.height / 2
        )

        vertical_distance = abs(
            shape_center_y
            - label_center_y
        )

        # Same row tolerance
        if vertical_distance > Pt(100):
            continue

        # Distance from label
        if shape.left >= label_right:

            horizontal_distance = (
                shape.left
                - label_right
            )

        else:

            horizontal_distance = (
                label_left
                - (
                    shape.left
                    + shape.width
                )
            )

        # Too far away
        if abs(horizontal_distance) > Pt(250):
            continue

        score = (
            abs(horizontal_distance)
            + vertical_distance
        )

        candidates.append(
            (
                score,
                item
            )
        )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x[0]
    )

    return candidates[0][1]


# ============================================================
# GET STYLE
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

                if shape.line.color.rgb:

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
        font_size
    )


# ============================================================
# SET SHAPE TEXT
# ============================================================

def set_shape_text(
    shape,
    text,
    style=None,
    bold=True
):

    try:

        if style is None:

            style = (
                get_style_from_shape(
                    shape
                )
            )

        (
            border_color,
            line_width,
            font_name,
            font_color,
            font_size
        ) = style

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

        run.text = str(text)

        run.font.name = font_name

        run.font.size = font_size

        run.font.bold = bold

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
# UPDATE INLINE SIZE
#
# Keeps:
# Size :-
#
# Only replaces:
# 180X48
#
# with:
# 240 X 36
# ============================================================

def update_inline_size(
    shape,
    final_size
):

    try:

        old_text = get_shape_text(
            shape
        )

        match = SIZE_PATTERN.search(
            old_text
        )

        if not match:
            return False

        start = match.start()
        end = match.end()

        new_text = (
            old_text[:start]
            + final_size
            + old_text[end:]
        )

        style = (
            get_style_from_shape(
                shape
            )
        )

        return set_shape_text(
            shape,
            new_text,
            style,
            bold=True
        )

    except Exception:

        return False


# ============================================================
# UPDATE EXISTING SIZE BOX
# ============================================================

def update_existing_size_box(
    shape,
    final_size
):

    try:

        style = (
            get_style_from_shape(
                shape
            )
        )

        return set_shape_text(
            shape,
            final_size,
            style,
            bold=True
        )

    except Exception:

        return False


# ============================================================
# CREATE NEW SIZE BOX
# ============================================================

def create_size_box(
    slide,
    size_label_item,
    final_size
):

    if size_label_item:

        label_shape = (
            size_label_item["shape"]
        )

        box_left = (
            label_shape.left
            + label_shape.width
            + Pt(5)
        )

        box_top = (
            label_shape.top
        )

        box_height = (
            label_shape.height
        )

        box_width = Pt(120)

        style = (
            get_style_from_shape(
                label_shape
            )
        )

    else:

        box_left = Pt(300)
        box_top = Pt(430)

        box_width = Pt(120)
        box_height = Pt(30)

        style = (
            RGBColor(227, 108, 10),
            Pt(2),
            "Calibri",
            RGBColor(0, 0, 0),
            Pt(16)
        )

    (
        border_color,
        line_width,
        font_name,
        font_color,
        font_size
    ) = style

    new_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        box_left,
        box_top,
        box_width,
        box_height
    )

    try:
        new_box.fill.background()
    except Exception:
        pass

    try:

        new_box.line.color.rgb = (
            border_color
        )

        new_box.line.width = (
            line_width
        )

    except Exception:
        pass

    set_shape_text(
        new_box,
        final_size,
        style,
        bold=True
    )

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
        "reason": ""
    }

    if not size_text:

        result["reason"] = (
            "Excel Width/Height is blank"
        )

        return result

    # --------------------------------------------------------
    # COLLECT ALL TEXT SHAPES ONLY ONCE
    # --------------------------------------------------------

    text_shapes = (
        collect_text_shapes(
            slide
        )
    )

    # --------------------------------------------------------
    # 1. SPECIAL FORMAT
    #
    # Size :- 180X48
    #
    # This must be checked FIRST.
    # --------------------------------------------------------

    inline_size = (
        find_inline_size_field(
            text_shapes
        )
    )

    if inline_size:

        success = (
            update_inline_size(
                inline_size["shape"],
                size_text
            )
        )

        if success:

            result["status"] = (
                "Updated Inline Size"
            )

            result["reason"] = (
                "Updated existing Size :- "
                "Width X Height field"
            )

            return result

    # --------------------------------------------------------
    # 2. FIND SIZE LABEL
    # --------------------------------------------------------

    size_label = (
        find_size_label(
            text_shapes
        )
    )

    # --------------------------------------------------------
    # 3. FIND EXISTING COMBINED SIZE BOX
    #
    # Example:
    # 180 X 36
    #
    # IMPORTANT:
    # Do this BEFORE creating anything.
    # --------------------------------------------------------

    existing_size = (
        find_standalone_size_box(
            text_shapes,
            size_label
        )
    )

    if existing_size:

        success = (
            update_existing_size_box(
                existing_size["shape"],
                size_text
            )
        )

        if success:

            result["status"] = (
                "Updated Existing Size"
            )

            result["reason"] = (
                "Existing combined Size box "
                "updated to Width X Height"
            )

            return result

    # --------------------------------------------------------
    # 4. NO EXISTING SIZE FOUND
    #
    # Only now create a new box.
    # --------------------------------------------------------

    if size_label:

        create_size_box(
            slide,
            size_label,
            size_text
        )

        result["status"] = (
            "Created Size Box"
        )

        result["reason"] = (
            "No safe existing Size box "
            "was found"
        )

        return result

    # --------------------------------------------------------
    # SAFE SKIP
    # --------------------------------------------------------

    result["status"] = (
        "Skipped - Safe Mode"
    )

    result["reason"] = (
        "Size field could not be safely identified"
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

        try:

            # =================================================
            # READ EXCEL
            # =================================================

            with st.status(
                "Preparing files...",
                expanded=False
            ):

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

                # ---------------------------------------------
                # FIND WIDTH / HEIGHT
                # ---------------------------------------------

                (
                    width_column,
                    height_column
                ) = get_size_columns(
                    df
                )

                if not width_column:

                    st.error(
                        "❌ Width column not found."
                    )

                    st.write(
                        "Available columns:"
                    )

                    st.write(
                        list(df.columns)
                    )

                    st.stop()

                if not height_column:

                    st.error(
                        "❌ Height column not found."
                    )

                    st.write(
                        "Available columns:"
                    )

                    st.write(
                        list(df.columns)
                    )

                    st.stop()

            # =================================================
            # SHOW DETECTED COLUMNS
            # =================================================

            col1, col2 = st.columns(2)

            with col1:

                st.success(
                    f"Width: {width_column}"
                )

            with col2:

                st.success(
                    f"Height: {height_column}"
                )

            # =================================================
            # LOAD PPT
            # =================================================

            with st.status(
                "Loading PowerPoint...",
                expanded=False
            ):

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

            # =================================================
            # INFORMATION
            # =================================================

            st.write(
                f"📊 PPT Slides: **{total_slides}**"
            )

            st.write(
                f"📑 Excel Rows: **{total_excel_rows}**"
            )

            st.write(
                f"🔄 Slides to process: **{process_count}**"
            )

            st.markdown("---")

            # =================================================
            # COUNTERS
            # =================================================

            updated_count = 0
            inline_count = 0
            created_count = 0
            skipped_count = 0

            report = []

            # =================================================
            # PROGRESS
            # =================================================

            progress_bar = st.progress(
                0
            )

            progress_text = st.empty()

            # =================================================
            # PROCESS SLIDES
            # =================================================

            for i in range(
                process_count
            ):

                slide = (
                    prs.slides[i]
                )

                # ---------------------------------------------
                # EXCEL DATA
                # ---------------------------------------------

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

                # ---------------------------------------------
                # WIDTH X HEIGHT
                # ---------------------------------------------

                size_text = (
                    make_size(
                        width,
                        height
                    )
                )

                # ---------------------------------------------
                # PROCESS
                # ---------------------------------------------

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

                # ---------------------------------------------
                # COUNTERS
                # ---------------------------------------------

                if status == (
                    "Updated Existing Size"
                ):

                    updated_count += 1

                elif status == (
                    "Updated Inline Size"
                ):

                    inline_count += 1

                elif status == (
                    "Created Size Box"
                ):

                    created_count += 1

                else:

                    skipped_count += 1

                # ---------------------------------------------
                # REPORT
                # ---------------------------------------------

                report.append({

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

                    "Final PPT Size":
                        size_text,

                    "Status":
                        status,

                    "Reason":
                        reason,

                })

                # ---------------------------------------------
                # PROGRESS
                # ---------------------------------------------

                percent = int(
                    (
                        (i + 1)
                        / max(
                            process_count,
                            1
                        )
                    )
                    * 100
                )

                progress_bar.progress(
                    percent
                )

                progress_text.write(
                    f"Processing slide "
                    f"{i + 1} / {process_count} "
                    f"({percent}%)"
                )

            # =================================================
            # EXTRA PPT SLIDES
            # =================================================

            if total_slides > total_excel_rows:

                for i in range(
                    total_excel_rows,
                    total_slides
                ):

                    skipped_count += 1

                    report.append({

                        "Slide":
                            i + 1,

                        "Excel Row":
                            "",

                        "Width":
                            "",

                        "Height":
                            "",

                        "Final PPT Size":
                            "",

                        "Status":
                            "Skipped",

                        "Reason":
                            "No matching Excel row"

                    })

            # =================================================
            # SAVE PPT
            # =================================================

            progress_text.write(
                "Saving updated PowerPoint..."
            )

            output_buffer = BytesIO()

            prs.save(
                output_buffer
            )

            output_buffer.seek(
                0
            )

            progress_bar.progress(
                100
            )

            progress_text.write(
                "✅ Processing completed."
            )

            # =================================================
            # SUCCESS
            # =================================================

            st.success(
                "🎉 PPT processing completed successfully."
            )

            # =================================================
            # METRICS
            # =================================================

            c1, c2, c3, c4 = st.columns(4)

            with c1:

                st.metric(
                    "Existing Size Updated",
                    updated_count
                )

            with c2:

                st.metric(
                    "Inline Size Updated",
                    inline_count
                )

            with c3:

                st.metric(
                    "New Box Created",
                    created_count
                )

            with c4:

                st.metric(
                    "Skipped",
                    skipped_count
                )

            # =================================================
            # REPORT
            # =================================================

            report_df = pd.DataFrame(
                report
            )

            st.subheader(
                "📋 Processing Report"
            )

            st.dataframe(
                report_df,
                use_container_width=True,
                height=450
            )

            # =================================================
            # DOWNLOAD PPT
            # =================================================

            st.download_button(

                label=(
                    "📥 Download Updated PPT"
                ),

                data=(
                    output_buffer.getvalue()
                ),

                file_name=(
                    "Updated_Presentation_V4.pptx"
                ),

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),

                type="primary"

            )

            # =================================================
            # DOWNLOAD REPORT
            # =================================================

            report_buffer = BytesIO()

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
                    report_buffer.getvalue()
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
