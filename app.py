import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import streamlit as st

st.set_page_config(
    page_title="PPT Size Box Automator", page_icon="📊", layout="centered"
)

st.title("📊 PPT Size Box Auto-Fixer")
st.write(
    "Apni **Excel File** aur **PPT File** upload karein. Code slide ke saare purane text ko clean karke naya aligned Orange Size Box bana dega."
)

st.markdown("---")

uploaded_excel = st.file_uploader(
    "1. Upload Excel File (.xlsx)", type=["xlsx"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

if uploaded_excel and uploaded_ppt:
    if st.button("🚀 Process & Generate Updated PPT", type="primary"):
        with st.spinner("Processing Presentation and Excel Data..."):
            try:
                # Read Excel
                df = pd.read_excel(uploaded_excel, sheet_name="Merged_Result")

                # Auto Find Size Column in Excel
                size_column_name = None
                for col in df.columns:
                    if "size" in str(col).lower():
                        size_column_name = col
                        break

                prs = Presentation(uploaded_ppt)
                manual_check_slides = []

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    slide_no = i + 1

                    # Get Size value from Excel
                    try:
                        if size_column_name:
                            size_val = str(df[size_column_name].iloc[i]).strip()
                        else:
                            size_val = str(df.iloc[i, 7]).strip()
                    except Exception:
                        size_val = ""

                    type_box = None
                    qty_box = None
                    shapes_to_remove = []

                    # STEP 1: Slide ke sabhi Shapes ko scan karein
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                            txt_lower = txt.lower()

                            # Detect Type Box & Qty Box (Position matching ke liye)
                            if "type" in txt_lower and "size" not in txt_lower:
                                type_box = shape
                            elif "qty" in txt_lower and "size" not in txt_lower:
                                qty_box = shape

                            # Clean/Delete: Size word ho, 'x' pattern ho ya bottom-middle space me floating text ho
                            if (
                                "size" in txt_lower
                                or ("x" in txt_lower and len(txt) < 20)
                                or (
                                    shape.top > Pt(380)
                                    and shape.left > Pt(350)
                                    and shape.left < Pt(650)
                                    and "type" not in txt_lower
                                    and "qty" not in txt_lower
                                )
                            ):
                                shapes_to_remove.append(shape)

                    # STEP 2: Purane Shapes aur Background Plain Text Complete Wipe Out
                    for shape in shapes_to_remove:
                        try:
                            shape.text_frame.text = ""
                            for p in shape.text_frame.paragraphs:
                                p.text = ""
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 3: Alignment Calculate Karein (Type aur Qty ke beech)
                    if type_box and qty_box:
                        target_left = int(
                            (
                                type_box.left
                                + type_box.width
                                + qty_box.left
                                - Pt(130)
                            )
                            / 2
                        )
                        target_top = type_box.top
                        target_width = Pt(130)
                        target_height = type_box.height
                    else:
                        # Fallback position agar Type/Qty na mile
                        target_left = Pt(450)
                        target_top = Pt(437)
                        target_width = Pt(130)
                        target_height = Pt(30)

                    # STEP 4: Perfect Single Clean Orange Box Create Karein
                    if size_val and size_val.lower() != "nan":
                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            target_left,
                            target_top,
                            target_width,
                            target_height,
                        )

                        new_box.fill.background()
                        new_box.line.color.rgb = RGBColor(
                            227, 108, 10
                        )  # Orange Border
                        new_box.line.width = Pt(2.0)  # Clean Border

                        tf = new_box.text_frame
                        tf.word_wrap = False
                        p = tf.paragraphs[0]
                        p.text = f"Size: {size_val}"
                        p.alignment = PP_ALIGN.CENTER

                        run = p.runs[0]
                        run.font.name = "Calibri"
                        run.font.size = Pt(16)
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(
                            0, 80, 160
                        )  # PPT Blue Font

                # Save Output
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 PPT Successfully Clean & Align Ho Gayi Hai! Old Overlapping Text Bilkul Clear Ho Gaya."
                )

                # Download Button
                with open(output_ppt_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Updated PPT",
                        data=file,
                        file_name="Updated_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            except Exception as e:
                st.error(f"❌ Error aaya: {e}")
