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
    "Apni koi bhi **Excel File** aur **PPT File** upload karein. Automatic font size, color aur alignment match ho jayega."
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
                # 1. Read Excel Safely (Peheli sheet automatically load karega)
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

                if size_column_name:
                    st.info(
                        f"✅ Excel me Automatic Size Column mil gaya: **'{size_column_name}'** (Sheet: '{sheet_to_use}')"
                    )
                else:
                    st.warning(
                        f"⚠️ 'Size' column nahi mila. Fallback Column Index 7 (Column H) use ho raha hai (Sheet: '{sheet_to_use}')."
                    )

                # 3. Read PPT
                prs = Presentation(uploaded_ppt)

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    slide_no = i + 1

                    # Get Size value safely
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

                    # Default Formatting Properties
                    target_font_size = Pt(20)
                    target_font_name = "Calibri"
                    target_font_color = RGBColor(0, 80, 160)
                    target_border_color = RGBColor(227, 108, 10)
                    target_border_width = Pt(1.5)

                    # STEP 1: Scan Shapes
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                            txt_lower = txt.lower()

                            # Detect Type Box
                            if "type" in txt_lower and "size" not in txt_lower:
                                type_box = shape
                                try:
                                    if shape.line.fill.type == 1:
                                        target_border_color = (
                                            shape.line.color.rgb
                                        )
                                    if shape.line.width:
                                        target_border_width = shape.line.width

                                    first_p = shape.text_frame.paragraphs[0]
                                    if first_p.runs:
                                        first_run = first_p.runs[0]
                                        if first_run.font.size:
                                            target_font_size = (
                                                first_run.font.size
                                            )
                                        if first_run.font.name:
                                            target_font_name = (
                                                first_run.font.name
                                            )
                                        if first_run.font.color.rgb:
                                            target_font_color = (
                                                first_run.font.color.rgb
                                            )
                                except Exception:
                                    pass

                            elif "qty" in txt_lower and "size" not in txt_lower:
                                qty_box = shape

                            # Identify Old Size Box / Floating Text
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

                    # STEP 3: Alignment Calculation
                    if type_box and qty_box:
                        target_width = Pt(160)
                        target_height = type_box.height
                        target_top = type_box.top
                        target_left = int(
                            (
                                type_box.left
                                + type_box.width
                                + qty_box.left
                                - target_width
                            )
                            / 2
                        )
                    else:
                        target_left = Pt(450)
                        target_top = Pt(437)
                        target_width = Pt(160)
                        target_height = Pt(32)

                    # STEP 4: Create New Size Box
                    if size_val and size_val.lower() != "nan":
                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            target_left,
                            target_top,
                            target_width,
                            target_height,
                        )

                        new_box.fill.background()
                        new_box.line.color.rgb = target_border_color
                        new_box.line.width = target_border_width

                        tf = new_box.text_frame
                        tf.word_wrap = False
                        p = tf.paragraphs[0]
                        p.text = f"Size: {size_val}"
                        p.alignment = PP_ALIGN.CENTER

                        run = p.runs[0]
                        run.font.name = target_font_name
                        run.font.size = target_font_size
                        run.font.bold = True
                        run.font.color.rgb = target_font_color

                # Save Output PPT
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 PPT Successfully Updated with Perfect Formatting!"
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
                st.error(
                    f"❌ Error Aaya: {e}\n\nKripya check karein ki Excel me row count PPT ki slides se match ho raha hai ya nahi."
                )
