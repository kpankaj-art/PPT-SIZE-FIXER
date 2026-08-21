import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt
import streamlit as st

st.set_page_config(
    page_title="Universal PPT Size Box Automator",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Universal Multi-Format PPT Automator")
st.write(
    "Yeh script dono me se kisi bhi template format ke sath **100% perfectly work** karegi."
)

st.markdown("---")

uploaded_excel = st.file_uploader(
    "1. Upload Excel File (.xlsx)", type=["xlsx"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

if uploaded_excel and uploaded_ppt:
    if st.button("🚀 Process & Regenerate PPT Size Box", type="primary"):
        with st.spinner("Processing Presentation and Excel Data..."):
            try:
                # 1. Excel Read safely
                xl_file = pd.ExcelFile(uploaded_excel)
                sheet_to_use = (
                    "Merged_Result"
                    if "Merged_Result" in xl_file.sheet_names
                    else xl_file.sheet_names[0]
                )
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_to_use)

                # 2. Size Column Find
                size_column_name = None
                for col in df.columns:
                    if "size" in str(col).lower():
                        size_column_name = col
                        break

                prs = Presentation(uploaded_ppt)

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    try:
                        if size_column_name:
                            size_val = str(df[size_column_name].iloc[i]).strip()
                        else:
                            size_val = str(df.iloc[i, 7]).strip()
                    except Exception:
                        size_val = ""

                    # Default fallback style properties
                    border_color = RGBColor(227, 108, 10)
                    line_width = Pt(2.0)
                    font_name = "Calibri"
                    font_color = RGBColor(0, 0, 0)

                    left_anchor = None
                    right_anchor = None

                    # STEP 1: DETECT LEFT & RIGHT ANCHOR SHAPES (Supports both Formats)
                    for shape in slide.shapes:
                        txt = ""
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                        txt_lower = txt.lower()

                        # Detect Left Neighbor (Media Type / Type)
                        if any(k in txt_lower for k in ["media", "type:"]) and not left_anchor:
                            left_anchor = shape
                        # Detect Right Neighbor (Qty / Board_Qty)
                        elif any(k in txt_lower for k in ["qty", "board_qty"]) and not right_anchor:
                            right_anchor = shape

                    # Extract Style Properties from detected anchor shapes
                    ref_shape = left_anchor or right_anchor
                    if ref_shape:
                        try:
                            if ref_shape.line and ref_shape.line.fill.type == 1:
                                border_color = ref_shape.line.color.rgb
                            if ref_shape.line and ref_shape.line.width:
                                line_width = ref_shape.line.width
                        except Exception:
                            pass

                        try:
                            if ref_shape.has_text_frame and ref_shape.text_frame.paragraphs:
                                p_ref = ref_shape.text_frame.paragraphs[0]
                                if p_ref.runs:
                                    r_ref = p_ref.runs[0]
                                    if r_ref.font.name:
                                        font_name = r_ref.font.name
                                    if r_ref.font.color and r_ref.font.color.rgb:
                                        font_color = r_ref.font.color.rgb
                        except Exception:
                            pass

                    # STEP 2: CALCULATE BOUNDARIES
                    if left_anchor and right_anchor:
                        left_bound = left_anchor.left + left_anchor.width
                        right_bound = right_anchor.left
                        top_bound = min(left_anchor.top, right_anchor.top) - Pt(20)
                        bottom_bound = max(
                            left_anchor.top + left_anchor.height,
                            right_anchor.top + right_anchor.height,
                        ) + Pt(20)
                        target_top = left_anchor.top
                        target_height = left_anchor.height
                    elif left_anchor:
                        left_bound = left_anchor.left + left_anchor.width
                        right_bound = left_anchor.left + left_anchor.width + Pt(150)
                        top_bound = left_anchor.top - Pt(20)
                        bottom_bound = left_anchor.top + left_anchor.height + Pt(20)
                        target_top = left_anchor.top
                        target_height = left_anchor.height
                    else:
                        left_bound = Pt(220)
                        right_bound = Pt(410)
                        top_bound = Pt(380)
                        bottom_bound = Pt(500)
                        target_top = Pt(432)
                        target_height = Pt(28)

                    shapes_to_delete = []

                    # STEP 3: SWEEP & PURGE GHOST ELEMENTS
                    for shape in slide.shapes:
                        if shape == left_anchor or shape == right_anchor:
                            continue

                        txt = ""
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                        txt_lower = txt.lower()

                        # Exclude main titles / header buttons
                        if any(k in txt_lower for k in ["close view", "far view", "remarks"]):
                            continue

                        # Check if shape is in the middle gap zone
                        is_in_gap_x = (shape.left >= (left_bound - Pt(15))) and (
                            (shape.left + shape.width) <= (right_bound + Pt(15))
                        )
                        is_in_gap_y = (shape.top >= top_bound) and (
                            (shape.top + shape.height) <= bottom_bound
                        )

                        is_size_label = "size" in txt_lower or (
                            "x" in txt_lower and len(txt) < 30
                        )

                        if (is_in_gap_x and is_in_gap_y) or is_size_label:
                            shapes_to_delete.append(shape)

                    # HARD XML DELETION
                    for shape in shapes_to_delete:
                        try:
                            if shape.has_text_frame:
                                shape.text_frame.text = ""
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 4: RE-CREATE SINGLE CLEAN SIZE BOX
                    if size_val and size_val.lower() != "nan":
                        final_text = (
                            size_val
                            if size_val.lower().startswith("size")
                            else f"Size: {size_val}"
                        )

                        box_left = left_bound + Pt(8)
                        box_width = (right_bound - left_bound) - Pt(16)

                        if box_width < Pt(70):
                            box_width = Pt(130)

                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            box_left,
                            target_top,
                            box_width,
                            target_height,
                        )

                        new_box.fill.background()
                        new_box.line.color.rgb = border_color
                        new_box.line.width = line_width

                        tf = new_box.text_frame
                        tf.word_wrap = False
                        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
                        tf.margin_top = Pt(1)
                        tf.margin_bottom = Pt(1)
                        tf.margin_left = Pt(1)
                        tf.margin_right = Pt(1)

                        p = tf.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER

                        run = p.add_run()
                        run.text = final_text
                        run.font.name = font_name
                        run.font.bold = True
                        run.font.color.rgb = font_color

                        # Font Sizing Logic
                        text_len = len(final_text)
                        if text_len > 16:
                            run.font.size = Pt(12)
                        elif text_len > 13:
                            run.font.size = Pt(14)
                        else:
                            run.font.size = Pt(16)

                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 Success! Dono formats ke liye universal processing successfully complete ho gayi hai."
                )

                with open(output_ppt_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Updated PPT",
                        data=file,
                        file_name="Updated_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            except Exception as e:
                st.error(f"❌ Error Aaya: {e}")
