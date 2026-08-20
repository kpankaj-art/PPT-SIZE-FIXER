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

st.title("📊 Smart Exact-Clone PPT Automator")
st.write(
    "Yeh code purane box ka sab kuch clone karega: **Position, Box Size, Font Style, Font Size, aur Exact Font Color!**"
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
                # 1. Read Excel safely
                xl_file = pd.ExcelFile(uploaded_excel)
                sheet_to_use = (
                    "Merged_Result"
                    if "Merged_Result" in xl_file.sheet_names
                    else xl_file.sheet_names[0]
                )
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_to_use)

                # 2. Auto Find Size Column in Excel
                size_column_name = None
                for col in df.columns:
                    if "size" in str(col).lower():
                        size_column_name = col
                        break

                prs = Presentation(uploaded_ppt)

                # Default Fallback Memory (Agar shuruat me purana box bilkul na mile)
                last_left = Pt(410)
                last_top = Pt(435)
                last_width = Pt(150)
                last_height = Pt(32)
                last_border_color = RGBColor(227, 108, 10) # Orange default
                last_border_width = Pt(1.5)
                last_font_color = RGBColor(0, 80, 160)     # Blue default
                last_font_size = Pt(18)
                last_font_name = "Calibri"

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    # Excel se Size read karna
                    try:
                        if size_column_name:
                            size_val = str(df[size_column_name].iloc[i]).strip()
                        else:
                            size_val = str(df.iloc[i, 7]).strip()
                    except Exception:
                        size_val = ""

                    shapes_to_remove = []

                    # STEP 1: Purane Size Box ko scan karke Position + Font + COLOR Details Capture Karna
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                            txt_lower = txt.lower()

                            # Size box / Dimension pattern matching
                            if (
                                "size" in txt_lower
                                or ("x" in txt_lower and len(txt) < 25)
                                or (
                                    shape.top > Pt(380)
                                    and shape.left > Pt(320)
                                    and shape.left < Pt(600)
                                    and "type" not in txt_lower
                                    and "qty" not in txt_lower
                                )
                            ):
                                shapes_to_remove.append(shape)

                                # 1. Position & Size Capture
                                last_left = shape.left
                                last_top = shape.top
                                last_width = shape.width
                                last_height = shape.height

                                # 2. Border Color & Width Capture
                                try:
                                    if hasattr(shape.line, "color") and hasattr(shape.line.color, "rgb") and shape.line.color.rgb:
                                        last_border_color = shape.line.color.rgb
                                    if shape.line.width:
                                        last_border_width = shape.line.width
                                except Exception:
                                    pass

                                # 3. Exact Font Name, Size, & ADVANCED COLOR Capture
                                try:
                                    for p in shape.text_frame.paragraphs:
                                        for run in p.runs:
                                            # Font Name Capture
                                            if run.font.name:
                                                last_font_name = run.font.name
                                            # Font Size Capture
                                            if run.font.size:
                                                last_font_size = run.font.size
                                            # Font EXACT RGB Color Capture
                                            if run.font.color and hasattr(run.font.color, "rgb") and run.font.color.rgb is not None:
                                                last_font_color = run.font.color.rgb
                                except Exception:
                                    pass

                    # STEP 2: Clear Old Elements Safely
                    for shape in shapes_to_remove:
                        try:
                            shape.text_frame.text = ""
                            for p in shape.text_frame.paragraphs:
                                p.text = ""
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 3: Re-create Box at EXACT Position with EXACT Formatting
                    if size_val and size_val.lower() != "nan":
                        # Font size ke hisab se width adjust karein taaki text overlap na ho
                        box_width = max(last_width, Pt(140))

                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            last_left,
                            last_top,
                            box_width,
                            last_height,
                        )

                        # Border & Fill Styling
                        new_box.fill.background()
                        new_box.line.color.rgb = last_border_color
                        new_box.line.width = last_border_width

                        # Text Frame Formatting
                        tf = new_box.text_frame
                        tf.word_wrap = False
                        p = tf.paragraphs[0]
                        p.text = f"Size: {size_val}"
                        p.alignment = PP_ALIGN.CENTER

                        # Apply EXACT Inherited Properties
                        run = p.runs[0]
                        run.font.name = last_font_name        # Clone Font Family
                        run.font.size = last_font_size        # Clone Font Size
                        run.font.bold = True
                        run.font.color.rgb = last_font_color  # Clone EXACT Font Color

                # Save Output Presentation
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 PPT Successfully Updated! Purane box ki **Position, Font Name, Font Size aur FONT COLOR** exact clone ho chuki hai."
                )

                # Download Link
                with open(output_ppt_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Updated PPT",
                        data=file,
                        file_name="Updated_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            except Exception as e:
                st.error(f"❌ Error Aaya: {e}")
