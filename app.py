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

st.title("📊 Clean Delete & Re-Create PPT Automator")
st.write(
    "Pehle purana text/box bilkul delete hoga, fir Excel se Naya Size Box exact position par create hoga."
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
                # 1. Read Excel safely
                xl_file = pd.ExcelFile(uploaded_excel)
                sheet_to_use = (
                    "Merged_Result"
                    if "Merged_Result" in xl_file.sheet_names
                    else xl_file.sheet_names[0]
                )
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_to_use)

                # 2. Find Size Column in Excel
                size_column_name = None
                for col in df.columns:
                    if "size" in str(col).lower():
                        size_column_name = col
                        break

                prs = Presentation(uploaded_ppt)

                # Default Position & Style Memory (if ppt doesn't have existing coordinates)
                target_left = Pt(280)
                target_top = Pt(435)
                target_width = Pt(120)
                target_height = Pt(28)
                target_border_color = RGBColor(227, 108, 10)
                target_font_color = RGBColor(0, 0, 0)
                target_font_size = Pt(11)
                target_font_name = "Calibri"

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

                    shapes_to_delete = []

                    # STEP 1: SCAN & MATCH all Old Size Elements (Plain Text + Shape Box)
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                            txt_lower = txt.lower()

                            is_excluded = any(
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

                            # Match Size Box or size-related floating texts
                            if (
                                "size" in txt_lower
                                or ("x" in txt_lower and len(txt) < 25)
                            ) and not is_excluded:
                                shapes_to_delete.append(shape)

                                # Save exact layout coordinates from the first matched shape
                                target_left = shape.left
                                target_top = shape.top
                                target_width = shape.width
                                target_height = shape.height

                                # Save font properties if available
                                try:
                                    if shape.line.fill.type == 1:
                                        target_border_color = (
                                            shape.line.color.rgb
                                        )
                                    p = shape.text_frame.paragraphs[0]
                                    if p.runs:
                                        r = p.runs[0]
                                        if r.font.name:
                                            target_font_name = r.font.name
                                        if r.font.size:
                                            target_font_size = r.font.size
                                        if r.font.color.rgb:
                                            target_font_color = (
                                                r.font.color.rgb
                                            )
                                except Exception:
                                    pass

                    # STEP 2: COMPLETE HARD DELETE (Clear XML + Erase Frame)
                    for shape in shapes_to_delete:
                        try:
                            # Sub-text clear
                            shape.text_frame.text = ""
                            for p in shape.text_frame.paragraphs:
                                p.text = ""

                            # XML Element deletion (Permanent PPT purge)
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 3: CREATE EXACT SINGLE NEW BOX
                    if size_val and size_val.lower() != "nan":
                        final_text = (
                            size_val
                            if size_val.lower().startswith("size")
                            else f"Size: {size_val}"
                        )

                        # Create Clean Shape Box
                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            target_left,
                            target_top,
                            target_width,
                            target_height,
                        )

                        # Styling Outer Box
                        new_box.fill.background()
                        new_box.line.color.rgb = target_border_color
                        new_box.line.width = Pt(1.5)

                        # Text Formatting
                        tf = new_box.text_frame
                        tf.word_wrap = False

                        p = tf.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER

                        run = p.add_run()
                        run.text = final_text
                        run.font.name = target_font_name
                        run.font.bold = True
                        run.font.color.rgb = target_font_color

                        # Auto Font Scaling for Tight Fits
                        text_len = len(final_text)
                        if text_len > 15:
                            run.font.size = Pt(9.5)
                        elif text_len > 12:
                            run.font.size = Pt(10.5)
                        else:
                            run.font.size = target_font_size

                # Save Updated Presentation
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 Purana Saara Size Content Delete Ho Gaya Aur Excel Se Sirf 1 Naya Size Box Ready Hai!"
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
