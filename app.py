import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import streamlit as st

st.set_page_config(
    page_title="PPT Spatial Clean Automator", page_icon="📊", layout="centered"
)

st.title("📊 Spatial Coordinate Clean & Re-Create Automator")
st.write(
    "Yeh code `Media Type` aur `Qty` ke beech ke pure area ko 100% clean wiping karke fresh Size box add karega."
)

st.markdown("---")

uploaded_excel = st.file_uploader(
    "1. Upload Excel File (.xlsx)", type=["xlsx"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

if uploaded_excel and uploaded_ppt:
    if st.button("🚀 Process & Regenerate PPT Size Box", type="primary"):
        with st.spinner("Clearing Zone & Updating PPT..."):
            try:
                # 1. Read Excel safely
                xl_file = pd.ExcelFile(uploaded_excel)
                sheet_to_use = (
                    "Merged_Result"
                    if "Merged_Result" in xl_file.sheet_names
                    else xl_file.sheet_names[0]
                )
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_to_use)

                # 2. Auto Find Size Column
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

                    # Coordinates memory for new box
                    media_right_edge = Pt(240)
                    qty_left_edge = Pt(400)
                    target_top = Pt(432)
                    target_height = Pt(28)

                    shapes_to_delete = []

                    # STEP 1: SCAN SPATIAL ZONE & KEYWORDS
                    for shape in slide.shapes:
                        txt = ""
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                        txt_lower = txt.lower()

                        # Protected keywords check
                        is_protected = any(
                            k in txt_lower
                            for k in [
                                "media",
                                "qty",
                                "remark",
                                "outlet",
                                "address",
                                "mobile",
                            ]
                        )

                        if is_protected:
                            if "media" in txt_lower:
                                media_right_edge = shape.left + shape.width
                                target_top = shape.top
                                target_height = shape.height
                            elif "qty" in txt_lower:
                                qty_left_edge = shape.left
                            continue

                        # Condition A: Text explicitly contains "size" or dimension pattern
                        is_size_text = "size" in txt_lower or (
                            "x" in txt_lower and len(txt) < 25
                        )

                        # Condition B: Position based (Anything sitting in the middle gap)
                        # Scanning horizontal center area between Media and Qty
                        is_in_target_zone = (
                            shape.top > Pt(380)
                            and shape.left > Pt(200)
                            and shape.left < Pt(450)
                        )

                        if is_size_text or is_in_target_zone:
                            shapes_to_delete.append(shape)

                    # STEP 2: HARD PURGE (Delete Text Frames + XML Shape Elements)
                    for shape in shapes_to_delete:
                        try:
                            if shape.has_text_frame:
                                shape.text_frame.text = ""
                                for p in shape.text_frame.paragraphs:
                                    p.text = ""

                            # Remove element directly from PowerPoint XML Tree
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 3: CREATE CLEAN NEW SIZE BOX
                    if size_val and size_val.lower() != "nan":
                        final_text = (
                            size_val
                            if size_val.lower().startswith("size")
                            else f"Size: {size_val}"
                        )

                        # Calculate precise width and position in the middle
                        box_left = media_right_edge + Pt(10)
                        box_width = (qty_left_edge - media_right_edge) - Pt(20)

                        if box_width < Pt(80):
                            box_width = Pt(130)

                        # Draw Box
                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            box_left,
                            target_top,
                            box_width,
                            target_height,
                        )

                        # Transparent background with standard orange border
                        new_box.fill.background()
                        new_box.line.color.rgb = RGBColor(227, 108, 10)
                        new_box.line.width = Pt(1.5)

                        # Insert Clean Text
                        tf = new_box.text_frame
                        tf.word_wrap = False
                        p = tf.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER

                        run = p.add_run()
                        run.text = final_text
                        run.font.name = "Calibri"
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0, 0, 0)

                        # Font size scaling according to text length
                        text_len = len(final_text)
                        if text_len > 16:
                            run.font.size = Pt(9.5)
                        elif text_len > 12:
                            run.font.size = Pt(10.5)
                        else:
                            run.font.size = Pt(11.5)

                # Save PPT File
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 Complete Zone Cleaned! Peeche pada float text `12 3` poori tarah erase ho gaya hai aur Naya Size Box perfectly placement me aa gaya hai."
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
