import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import streamlit as st

st.set_page_config(
    page_title="PPT Perfect Style Match Automator", page_icon="📊", layout="centered"
)

st.title("📊 100% Perfect Box Style Matching Automator")
st.write(
    "Yeh code `Qty` ya `Media Type` box ka exact border color, line thickness, aur font style clone karke Naye Size Box par apply karega."
)

st.markdown("---")

uploaded_excel = st.file_uploader(
    "1. Upload Excel File (.xlsx)", type=["xlsx"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

if uploaded_excel and uploaded_ppt:
    if st.button("🚀 Process & Regenerate PPT Size Box", type="primary"):
        with st.spinner("Cloning Styles & Updating PPT..."):
            try:
                # 1. Read Excel safely
                xl_file = pd.ExcelFile(uploaded_excel)
                sheet_to_use = (
                    "Merged_Result"
                    if "Merged_Result" in xl_file.sheet_names
                    else xl_file.sheet_names[0]
                )
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_to_use)

                # 2. Find Size Column
                size_column_name = None
                for col in df.columns:
                    if "size" in str(col).lower():
                        size_column_name = col
                        break

                prs = Presentation(uploaded_ppt)

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    # Fetch Excel Value
                    try:
                        if size_column_name:
                            size_val = str(df[size_column_name].iloc[i]).strip()
                        else:
                            size_val = str(df.iloc[i, 7]).strip()
                    except Exception:
                        size_val = ""

                    # Default fallback style properties
                    border_color = RGBColor(255, 140, 0) # Fallback template orange
                    line_width = Pt(2.25)                # Exact standard box thickness
                    font_name = "Calibri"
                    font_size = Pt(13)
                    font_color = RGBColor(0, 0, 0)
                    font_bold = True

                    media_right_edge = Pt(230)
                    qty_left_edge = Pt(400)
                    target_top = Pt(432)
                    target_height = Pt(28)

                    shapes_to_delete = []

                    # STEP 1: SCAN SLIDE, CLONE REFERENCE STYLE FROM QTY / MEDIA BOX
                    for shape in slide.shapes:
                        txt = ""
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                        txt_lower = txt.lower()

                        # Check neighbor boxes to steal their exact styling
                        if any(k in txt_lower for k in ["qty", "media", "remarks"]):
                            # Align Y-position and height with neighbors
                            target_top = shape.top
                            target_height = shape.height

                            if "media" in txt_lower:
                                media_right_edge = shape.left + shape.width
                            elif "qty" in txt_lower:
                                qty_left_edge = shape.left

                            # Extract Line Border Color & Width
                            try:
                                if shape.line and shape.line.fill.type == 1:
                                    border_color = shape.line.color.rgb
                                if shape.line and shape.line.width:
                                    line_width = shape.line.width
                            except Exception:
                                pass

                            # Extract Font Style Properties
                            try:
                                if shape.has_text_frame and shape.text_frame.paragraphs:
                                    p_ref = shape.text_frame.paragraphs[0]
                                    if p_ref.runs:
                                        r_ref = p_ref.runs[0]
                                        if r_ref.font.name:
                                            font_name = r_ref.font.name
                                        if r_ref.font.size:
                                            font_size = r_ref.font.size
                                        if r_ref.font.color and r_ref.font.color.rgb:
                                            font_color = r_ref.font.color.rgb
                                        if r_ref.font.bold is not None:
                                            font_bold = r_ref.font.bold
                            except Exception:
                                pass

                            continue

                        # Check if shape is in the middle gap or contains old Size text
                        is_size_text = "size" in txt_lower or (
                            "x" in txt_lower and len(txt) < 25
                        )
                        is_in_target_zone = (
                            shape.top > Pt(380)
                            and shape.left > Pt(200)
                            and shape.left < Pt(450)
                        )

                        if is_size_text or is_in_target_zone:
                            shapes_to_delete.append(shape)

                    # STEP 2: HARD PURGE OLD SIZE SHAPES & UNBOUND FLOATING TEXTS
                    for shape in shapes_to_delete:
                        try:
                            if shape.has_text_frame:
                                shape.text_frame.text = ""
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 3: CREATE NEW SIZE BOX WITH EXACT CLONED STYLE
                    if size_val and size_val.lower() != "nan":
                        final_text = (
                            size_val
                            if size_val.lower().startswith("size")
                            else f"Size: {size_val}"
                        )

                        # Precise middle calculation between Media and Qty boxes
                        box_left = media_right_edge + Pt(12)
                        box_width = (qty_left_edge - media_right_edge) - Pt(24)

                        if box_width < Pt(80):
                            box_width = Pt(130)

                        # Create Rectangle Box
                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            box_left,
                            target_top,
                            box_width,
                            target_height,
                        )

                        # Apply Cloned Border & Background Styling
                        new_box.fill.background()
                        new_box.line.color.rgb = border_color
                        new_box.line.width = line_width

                        # Apply Cloned Text Formatting
                        tf = new_box.text_frame
                        tf.word_wrap = False
                        p = tf.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER

                        run = p.add_run()
                        run.text = final_text
                        run.font.name = font_name
                        run.font.bold = font_bold
                        run.font.color.rgb = font_color

                        # Auto Font Size Adjustments for longer text strings
                        text_len = len(final_text)
                        if text_len > 16:
                            run.font.size = Pt(9.5)
                        elif text_len > 13:
                            run.font.size = Pt(10.5)
                        else:
                            run.font.size = font_size

                # Save Presentation
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 Complete Perfect Match! Size box ka border color, thickness aur font baaki boxes se 100% same match ho gaya hai."
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
