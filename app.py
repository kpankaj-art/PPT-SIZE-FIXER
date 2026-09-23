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
    page_title="PPT SIZE FIXER V3",
    page_icon="📊",
    layout="centered",
)

st.title("📊 PPT SIZE FIXER V3")

st.write("Excel Width + Height → PPT Size")

st.info(
    "SAFE MODE: Media Type, Qty, Remarks, Outlet, Address, "
    "Contact and SAP Code are protected."
)

st.markdown("---")


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_excel = st.file_uploader(
    "1. Upload Master Excel File",
    type=["xlsx", "xls", "csv", "xlsm"],
    key="excel_uploader_v3",
)

uploaded_ppt = st.file_uploader(
    "2. Upload PowerPoint Presentation",
    type=["pptx"],
    key="ppt_uploader_v3",
)


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(value):
    if value is None:
        return ""

    text = str(value).strip().lower()

    text = text.replace("_", " ")
    text = text.replace("-", " ")
    text = text.replace(":", " ")
    text = text.replace("(", " ")
    text = text.replace(")", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# FIND EXCEL COLUMN
# ============================================================

def find_column(df, possible_names):

    normalized_columns = {}

    for col in df.columns:
        normalized_columns[normalize_text(col)] = col

    # Exact match
    for name in possible_names:

        key = normalize_text(name)

        if key in normalized_columns:
            return normalized_columns[key]

    # Safe partial match
    for col in df.columns:

        col_norm = normalize_text(col)

        for name in possible_names:

            name_norm = normalize_text(name)

            if not name_norm:
                continue

            if col_norm == name_norm:
                return col

    # Partial only for longer names
    for col in df.columns:

        col_norm = normalize_text(col)

        for name in possible_names:

            name_norm = normalize_text(name)

            if len(name_norm) >= 4 and name_norm in col_norm:
                return col

    return None


# ============================================================
# READ EXCEL
# ============================================================

def read_excel_safely(file_obj):

    name = file_obj.name.lower()

    if name.endswith(".csv"):

        file_obj.seek(0)

        return pd.read_csv(file_obj)

    file_obj.seek(0)

    excel_file = pd.ExcelFile(file_obj)

    if not excel_file.sheet_names:
        raise ValueError("Excel file has no worksheets.")

    if "Merged_Result" in excel_file.sheet_names:

        sheet_name = "Merged_Result"

    else:

        sheet_name = excel_file.sheet_names[0]

    file_obj.seek(0)

    return pd.read_excel(
        file_obj,
        sheet_name=sheet_name,
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
        width_names,
    )

    height_column = find_column(
        df,
        height_names,
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

    if text.lower() in ("nan", "none", "null"):
        return ""

    text = text.replace(",", "")

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not match:
        return ""

    number = float(match.group())

    if number.is_integer():
        return str(int(number))

    return (
        str(number)
        .rstrip("0")
        .rstrip(".")
    )


# ============================================================
# CREATE FINAL SIZE
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
# GET TEXT SHAPES ONLY ONCE
# ============================================================

def collect_text_shapes(slide):

    result = []

    try:

        for shape in slide.shapes:

            try:

                if not shape.has_text_frame:
                    continue

                text = shape.text_frame.text.strip()

                if not text:
                    continue

                result.append(
                    {
                        "shape": shape,
                        "text": text,
                        "norm": normalize_text(text),
                        "left": shape.left,
                        "top": shape.top,
                        "width": shape.width,
                        "height": shape.height,
                    }
                )

            except Exception:
                continue

    except Exception:
        pass

    return result


# ============================================================
# SIZE LABEL CHECK
# ============================================================

def is_size_label(text):

    norm = normalize_text(text)

    return norm in {
        "size",
        "size :",
        "size -",
        "size :-",
    }


# ============================================================
# PROTECTED FIELD
# ============================================================

def is_protected(text):

    norm = normalize_text(text)

    protected = [
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
        "sap",
        "sap code",
    ]

    for word in protected:

        if norm == word:
            return True

        if word in norm and len(word) >= 6:
            return True

    return False


# ============================================================
# SIZE TEXT CHECK
# ============================================================

def looks_like_combined_size(text):

    if not text:
        return False

    value = str(text).strip()

    pattern = (
        r"^\s*"
        r"\d+(?:\.\d+)?"
        r"\s*[xX×*]"
        r"\s*"
        r"\d+(?:\.\d+)?"
        r"\s*"
        r"(?:in|inch|inches|ft|feet)?"
        r"\s*$"
    )

    return bool(
        re.match(
            pattern,
            value,
        )
    )


# ============================================================
# NUMBER ONLY
# ============================================================

def is_number_only(text):

    if not text:
        return False

    value = str(text).strip()

    return bool(
        re.fullmatch(
            r"\d+(?:\.\d+)?",
            value,
        )
    )


# ============================================================
# X SEPARATOR
# ============================================================

def is_x_separator(text):

    if not text:
        return False

    value = str(text).strip().lower()

    return value in {
        "x",
        "×",
        "*",
    }


# ============================================================
# GET STYLE
# ============================================================

def get_style(shape):

    border_color = RGBColor(
        227,
        108,
        10,
    )

    line_width = Pt(2)

    font_name = "Calibri"

    font_color = RGBColor(
        0,
        0,
        0,
    )

    font_size = Pt(16)

    try:

        if shape.line:

            try:

                if (
                    shape.line.fill.type == 1
                    and shape.line.color.rgb
                ):

                    border_color = shape.line.color.rgb

            except Exception:
                pass

            try:

                if shape.line.width:
                    line_width = shape.line.width

            except Exception:
                pass

    except Exception:
        pass

    try:

        if shape.has_text_frame:

            for paragraph in shape.text_frame.paragraphs:

                for run in paragraph.runs:

                    try:

                        if run.font.name:
                            font_name = run.font.name

                    except Exception:
                        pass

                    try:

                        if run.font.size:
                            font_size = run.font.size

                    except Exception:
                        pass

                    try:

                        if (
                            run.font.color
                            and run.font.color.rgb
                        ):

                            font_color = run.font.color.rgb

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
# SET TEXT IN EXISTING SHAPE
# ============================================================

def set_shape_text(
    shape,
    text,
    keep_style=True,
):

    try:

        old_style = get_style(shape)

        (
            border_color,
            line_width,
            font_name,
            font_color,
            font_size,
        ) = old_style

        tf = shape.text_frame

        tf.clear()

        tf.word_wrap = False

        tf.vertical_anchor = MSO_ANCHOR.MIDDLE

        tf.margin_top = Pt(1)
        tf.margin_bottom = Pt(1)
        tf.margin_left = Pt(1)
        tf.margin_right = Pt(1)

        paragraph = tf.paragraphs[0]

        paragraph.alignment = PP_ALIGN.CENTER

        run = paragraph.add_run()

        run.text = text

        run.font.name = font_name

        run.font.bold = True

        run.font.size = font_size

        try:
            run.font.color.rgb = font_color
        except Exception:
            pass

        if keep_style:

            try:
                shape.line.color.rgb = border_color
                shape.line.width = line_width
            except Exception:
                pass

        return True

    except Exception:
        return False


# ============================================================
# UPDATE COMBINED SIZE
# ============================================================

def update_combined_size(
    text_shapes,
    final_size,
):

    candidates = []

    for item in text_shapes:

        shape = item["shape"]

        text = item["text"]

        if is_protected(text):
            continue

        if looks_like_combined_size(text):

            candidates.append(item)

    if not candidates:
        return None

    # Prefer lower portion
    candidates.sort(
        key=lambda x: x["top"],
        reverse=True,
    )

    target = candidates[0]["shape"]

    if set_shape_text(
        target,
        final_size,
    ):

        return target

    return None


# ============================================================
# UPDATE SEPARATE WIDTH / HEIGHT / X
# ============================================================

def update_separate_size(
    text_shapes,
    size_label,
    width,
    height,
):

    if size_label is None:
        return False

    width_text = clean_number(width)
    height_text = clean_number(height)

    if not width_text or not height_text:
        return False

    label_center_y = (
        size_label.top
        + size_label.height / 2
    )

    label_right = (
        size_label.left
        + size_label.width
    )

    candidates = []

    for item in text_shapes:

        shape = item["shape"]

        text = item["text"]

        if shape == size_label:
            continue

        if is_protected(text):
            continue

        if not is_number_only(text):
            continue

        center_y = (
            shape.top
            + shape.height / 2
        )

        vertical_distance = abs(
            center_y - label_center_y
        )

        if vertical_distance > Pt(80):
            continue

        horizontal_distance = (
            shape.left - label_right
        )

        if horizontal_distance < -Pt(30):
            continue

        if horizontal_distance > Pt(300):
            continue

        candidates.append(
            (
                horizontal_distance
                + vertical_distance,
                item,
            )
        )

    if len(candidates) < 2:
        return False

    candidates.sort(
        key=lambda x: x[0]
    )

    nearby = [
        item
        for _, item in candidates[:6]
    ]

    # Sort from left to right
    nearby.sort(
        key=lambda x: x["left"]
    )

    # Find X separator
    x_item = None

    for item in nearby:

        if is_x_separator(
            item["text"]
        ):

            x_item = item
            break

    if x_item:

        left_numbers = []

        right_numbers = []

        x_left = x_item["left"]

        for item in nearby:

            if item is x_item:
                continue

            if not is_number_only(
                item["text"]
            ):
                continue

            if item["left"] < x_left:
                left_numbers.append(item)

            else:
                right_numbers.append(item)

        if left_numbers and right_numbers:

            left_numbers.sort(
                key=lambda x: abs(
                    x["left"]
                    - x_item["left"]
                )
            )

            right_numbers.sort(
                key=lambda x: abs(
                    x["left"]
                    - x_item["left"]
                )
            )

            width_shape = left_numbers[0]["shape"]

            height_shape = right_numbers[0]["shape"]

            ok1 = set_shape_text(
                width_shape,
                width_text,
                keep_style=False,
            )

            ok2 = set_shape_text(
                height_shape,
                height_text,
                keep_style=False,
            )

            return ok1 and ok2

    # --------------------------------------------------------
    # If no X separator:
    # choose two closest numeric shapes
    # --------------------------------------------------------

    numeric_shapes = [
        item
        for item in nearby
        if is_number_only(
            item["text"]
        )
    ]

    if len(numeric_shapes) >= 2:

        numeric_shapes.sort(
            key=lambda x: x["left"]
        )

        width_shape = numeric_shapes[0]["shape"]

        height_shape = numeric_shapes[1]["shape"]

        ok1 = set_shape_text(
            width_shape,
            width_text,
            keep_style=False,
        )

        ok2 = set_shape_text(
            height_shape,
            height_text,
            keep_style=False,
        )

        return ok1 and ok2

    return False


# ============================================================
# CREATE SIZE BOX
# ============================================================

def create_size_box(
    slide,
    size_label,
    final_size,
):

    if size_label is not None:

        box_left = (
            size_label.left
            + size_label.width
            + Pt(5)
        )

        box_top = size_label.top

        box_height = size_label.height

    else:

        box_left = Pt(300)

        box_top = Pt(430)

        box_height = Pt(30)

    box_width = Pt(130)

    new_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        box_left,
        box_top,
        box_width,
        box_height,
    )

    new_box.fill.background()

    new_box.line.color.rgb = RGBColor(
        227,
        108,
        10,
    )

    new_box.line.width = Pt(2)

    tf = new_box.text_frame

    tf.clear()

    tf.word_wrap = False

    tf.vertical_anchor = (
        MSO_ANCHOR.MIDDLE
    )

    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    tf.margin_left = Pt(1)
    tf.margin_right = Pt(1)

    paragraph = tf.paragraphs[0]

    paragraph.alignment = PP_ALIGN.CENTER

    run = paragraph.add_run()

    run.text = final_size

    run.font.name = "Calibri"

    run.font.bold = True

    run.font.size = Pt(16)

    run.font.color.rgb = RGBColor(
        0,
        0,
        0,
    )

    return new_box


# ============================================================
# PROCESS ONE SLIDE
# ============================================================

def process_slide(
    slide,
    width,
    height,
):

    result = {
        "status": "Skipped",
        "reason": "",
    }

    final_size = make_size(
        width,
        height,
    )

    if not final_size:

        result["reason"] = (
            "Width or Height is blank"
        )

        return result

    # --------------------------------------------------------
    # Collect ALL text shapes only once
    # --------------------------------------------------------

    text_shapes = collect_text_shapes(
        slide
    )

    # --------------------------------------------------------
    # Find Size label
    # --------------------------------------------------------

    size_label = None

    for item in text_shapes:

        if is_size_label(
            item["text"]
        ):

            size_label = item["shape"]

            break

    # --------------------------------------------------------
    # FIRST:
    # Combined format 180X48
    # --------------------------------------------------------

    combined = update_combined_size(
        text_shapes,
        final_size,
    )

    if combined is not None:

        result["status"] = (
            "Updated Existing Size"
        )

        result["reason"] = (
            "Combined Size value updated"
        )

        return result

    # --------------------------------------------------------
    # SECOND:
    # Separate 36 / 180 / x format
    # --------------------------------------------------------

    separate_updated = (
        update_separate_size(
            text_shapes,
            size_label,
            width,
            height,
        )
    )

    if separate_updated:

        result["status"] = (
            "Updated Separate Size"
        )

        result["reason"] = (
            "Width and Height values updated"
        )

        return result

    # --------------------------------------------------------
    # THIRD:
    # Create new Size box
    # --------------------------------------------------------

    if size_label is not None:

        create_size_box(
            slide,
            size_label,
            final_size,
        )

        result["status"] = (
            "Created Size Box"
        )

        result["reason"] = (
            "Existing Size value not safely detected"
        )

        return result

    # --------------------------------------------------------
    # SAFE SKIP
    # --------------------------------------------------------

    result["reason"] = (
        "Size field could not be safely detected"
    )

    return result


# ============================================================
# MAIN PROCESS
# ============================================================

if uploaded_excel and uploaded_ppt:

    st.markdown("---")

    if st.button(
        "🚀 Process & Sync Files",
        type="primary",
        use_container_width=True,
    ):

        try:

            # ==================================================
            # STEP 1 - READ EXCEL
            # ==================================================

            status_box = st.empty()

            status_box.info(
                "📖 Reading Excel file..."
            )

            df = read_excel_safely(
                uploaded_excel
            )

            if df.empty:

                st.error(
                    "❌ Excel file is empty."
                )

                st.stop()

            # ==================================================
            # STEP 2 - FIND WIDTH / HEIGHT
            # ==================================================

            status_box.info(
                "🔎 Detecting Width and Height columns..."
            )

            (
                width_column,
                height_column,
            ) = get_size_columns(df)

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

            st.success(
                f"✅ Width: `{width_column}`"
            )

            st.success(
                f"✅ Height: `{height_column}`"
            )

            # ==================================================
            # STEP 3 - READ PPT
            # ==================================================

            status_box.info(
                "📂 Loading PowerPoint..."
            )

            ppt_bytes = uploaded_ppt.getvalue()

            prs = Presentation(
                BytesIO(ppt_bytes)
            )

            total_slides = len(
                prs.slides
            )

            total_rows = len(df)

            process_count = min(
                total_slides,
                total_rows,
            )

            st.info(
                f"📊 PPT Slides: {total_slides} | "
                f"Excel Rows: {total_rows} | "
                f"Slides to process: {process_count}"
            )

            # ==================================================
            # PROGRESS
            # ==================================================

            progress_bar = st.progress(
                0
            )

            progress_text = st.empty()

            # ==================================================
            # COUNTERS
            # ==================================================

            updated_count = 0
            separate_count = 0
            created_count = 0
            skipped_count = 0

            report = []

            # ==================================================
            # PROCESS SLIDES
            # ==================================================

            for i in range(
                process_count
            ):

                slide = prs.slides[i]

                width = df.iloc[i][
                    width_column
                ]

                height = df.iloc[i][
                    height_column
                ]

                result = process_slide(
                    slide,
                    width,
                    height,
                )

                status = result[
                    "status"
                ]

                reason = result[
                    "reason"
                ]

                final_size = make_size(
                    width,
                    height,
                )

                # ------------------------------------------------
                # COUNTERS
                # ------------------------------------------------

                if status == (
                    "Updated Existing Size"
                ):

                    updated_count += 1

                elif status == (
                    "Updated Separate Size"
                ):

                    separate_count += 1

                elif status == (
                    "Created Size Box"
                ):

                    created_count += 1

                else:

                    skipped_count += 1

                # ------------------------------------------------
                # REPORT
                # ------------------------------------------------

                report.append(
                    {
                        "Slide": i + 1,
                        "Excel Row": i + 2,
                        "Width": clean_number(
                            width
                        ),
                        "Height": clean_number(
                            height
                        ),
                        "PPT Size": final_size,
                        "Status": status,
                        "Reason": reason,
                    }
                )

                # ------------------------------------------------
                # PROGRESS
                # ------------------------------------------------

                percent = int(
                    (
                        (i + 1)
                        / process_count
                    )
                    * 100
                )

                progress_bar.progress(
                    percent
                )

                progress_text.write(
                    f"⚙️ Processing slide "
                    f"{i + 1} / {process_count} "
                    f"({percent}%)"
                )

            # ==================================================
            # EXTRA PPT SLIDES
            # ==================================================

            if total_slides > total_rows:

                for i in range(
                    total_rows,
                    total_slides,
                ):

                    skipped_count += 1

                    report.append(
                        {
                            "Slide": i + 1,
                            "Excel Row": "",
                            "Width": "",
                            "Height": "",
                            "PPT Size": "",
                            "Status": "Skipped",
                            "Reason": (
                                "No matching Excel row"
                            ),
                        }
                    )

            # ==================================================
            # SAVE PPT
            # ==================================================

            status_box.info(
                "💾 Saving updated PowerPoint..."
            )

            output_ppt_buffer = BytesIO()

            prs.save(
                output_ppt_buffer
            )

            output_ppt_buffer.seek(0)

            # ==================================================
            # REPORT DATAFRAME
            # ==================================================

            report_df = pd.DataFrame(
                report
            )

            # ==================================================
            # COMPLETE
            # ==================================================

            progress_bar.progress(
                100
            )

            progress_text.success(
                "✅ Processing completed!"
            )

            status_box.empty()

            st.success(
                "🎉 PPT Size processing completed successfully."
            )

            # ==================================================
            # METRICS
            # ==================================================

            col1, col2, col3, col4 = (
                st.columns(4)
            )

            with col1:

                st.metric(
                    "Combined Updated",
                    updated_count,
                )

            with col2:

                st.metric(
                    "Separate Updated",
                    separate_count,
                )

            with col3:

                st.metric(
                    "Created",
                    created_count,
                )

            with col4:

                st.metric(
                    "Skipped",
                    skipped_count,
                )

            # ==================================================
            # DOWNLOAD PPT
            # ==================================================

            st.download_button(
                label="📥 Download Updated PPT",
                data=output_ppt_buffer.getvalue(),
                file_name=(
                    "Updated_Presentation_V3.pptx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                ),
                use_container_width=True,
            )

            # ==================================================
            # REPORT
            # ==================================================

            st.markdown("---")

            st.subheader(
                "📋 Processing Report"
            )

            st.dataframe(
                report_df,
                use_container_width=True,
                height=500,
            )

            # ==================================================
            # DOWNLOAD REPORT
            # ==================================================

            report_buffer = BytesIO()

            with pd.ExcelWriter(
                report_buffer,
                engine="openpyxl",
            ) as writer:

                report_df.to_excel(
                    writer,
                    index=False,
                    sheet_name=(
                        "Processing Report"
                    ),
                )

            report_buffer.seek(0)

            st.download_button(
                label="📊 Download Processing Report",
                data=report_buffer.getvalue(),
                file_name=(
                    "PPT_Size_Processing_Report.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        except Exception as e:

            st.error(
                f"❌ Error Occurred: {e}"
            )

            st.exception(e)
