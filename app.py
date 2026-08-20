import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import streamlit as st

# Page Title & Configuration
st.set_page_config(
    page_title="PPT Size Box Automator", page_icon="📊", layout="centered"
)

st.title("📊 PPT Size Box Auto-Fixer")
st.write(
    "Apni **Excel File** aur **PPT File** upload karein. System automatically size boxes ko update/fix karke aapko download link de dega."
)

st.markdown("---")

# File Uploaders
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

                # Auto Find Size Column
                size_column_name = None
                for col in df.columns:
                    if "size" in str(col).lower():
                        size_column_name = col
                        break

                if size_column_name:
                    st.info(
                        f"✅ Excel me Automatic Size Column mil gaya: **'{size_column_name}'**"
                    )
                else:
                    st.warning(
                        "⚠️ Excel me 'Size' column nahi mila. Fallback Index 7 (Column H) use ho raha hai."
                    )

                # Read PPT
                prs = Presentation(uploaded_ppt)

                last_left = None
                last_top = None
                last_width = None
                last_height = None
                last_color = None

                manual_check_slides = []

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    slide_no = i + 1

                    try:
                        if size_column_name:
                            size_val = str(df[size_column_name].iloc[i]).strip()
                        else:
                            size_val = str(df.iloc[i, 7]).strip()
                    except Exception:
                        size_val = ""

                    box_found = False
                    shapes_to_remove = []

                    # Search existing Size Box
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            if "Size" in shape.text_frame.text:
                                shapes_to_remove.append(shape)
                                box_found = True

                                last_left = shape.left
                                last_top = shape.top
                                last_width = shape.width
                                last_height = shape.height

                                try:
                                    if shape.line.fill.type == 1:
                                        last_color = shape.line.color.rgb
                                except Exception:
                                    pass

                    # Delete old boxes
                    for shape in shapes_to_remove:
                        sp = shape._element
                        sp.getparent().remove(sp)

                    # Re-create box
                    if size_val and size_val.lower() != "nan":
                        if last_left is not None:
                            new_box = slide.shapes.add_shape(
                                MSO_SHAPE.RECTANGLE,
                                last_left,
                                last_top,
                                last_width,
                                last_height,
                            )

                            new_box.fill.background()

                            if last_color is None:
                                last_color = RGBColor(227, 108, 10)

                            new_box.line.color.rgb = last_color
                            new_box.line.width = Pt(3.0)

                            tf = new_box.text_frame
                            tf.word_wrap = False
                            p = tf.paragraphs[0]
                            p.text = f"Size: {size_val}"
                            p.alignment = PP_ALIGN.CENTER

                            run = p.runs[0]
                            run.font.name = "Calibri"
                            run.font.size = Pt(18)
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0, 80, 160)
                        else:
                            manual_check_slides.append(slide_no)

                # Save output to memory buffer for download
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 PPT Successfully Update ho gayi hai! Niche button se download karein."
                )

                if manual_check_slides:
                    st.warning(
                        f"⚠️ **In Slide Numbers par pehle se koi box/reference nahi mila:** {manual_check_slides}\n\n"
                        "👉 In slides par Size Box ko kripya manually check karein."
                    )

                # Provide Download Button
                with open(output_ppt_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Updated PPT",
                        data=file,
                        file_name="Updated_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            except Exception as e:
                st.error(f"❌ Kuch error aaya: {e}")
