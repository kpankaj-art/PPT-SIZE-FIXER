import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt
import streamlit as st

st.set_page_config(
    page_title="PPT Precise Size Target Automator",
    page_icon="📊",
    layout="centered",
)

st.title("📊 Protected Neighbor Boxes & Size Target Automator")
st.write(
    "Is code me `Media Type:`, `Qty:` aur `Remarks:` ka text bilkul **safe** rahega. Deletion sirf aur sirf **Size Box** zone par hoga."
)

st.markdown("---")

uploaded_excel = st.file_uploader(
    "1. Upload Excel File (.xlsx)", type=["xlsx"]
)
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

if uploaded_excel and uploaded_ppt:
    if st.button("🚀 Process & Regenerate PPT Size Box", type="primary"):
        with st.spinner("Processing PPT safely..."):
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

                    # Default fallback styling properties
                    border_color = RGBColor(255, 140, 0)
                    line_width = Pt(2.0)
                    font_name = "Calibri"
                    font_color = RGBColor(0, 0, 0)

                    media_right_edge = Pt(230)
                    qty_left_edge = Pt(400)

                    ref_top = None
                    ref_height = None

                    shapes_to_delete = []

                    # STEP 1: READ NEIGHBOR STYLES (STRICT DO NOT DELETE NEIGHBORS)
                    for shape in slide.shapes:
                        txt = ""
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                        txt_lower = txt.lower()

                        # PROTECTED KEYWORDS: Media Type, Qty, Remarks
                        is_media = "media" in txt_lower
                        is_qty = "qty" in txt_lower
                        is_remarks = "remark" in txt_lower

                        if is_media or is_qty or is_remarks:
                            # Capture Top & Height alignment from neighboring boxes
                            ref_top = shape.top
                            ref_height = shape.height

                            if is_media:
                                media_right_edge = shape.left + shape.width
                            elif is_qty:
                                qty_left_edge = shape.left

                            # Copy Border Properties safely
                            try:
                                if shape.line and shape.line.fill.type == 1:
                                    border_color = shape.line.color.rgb
                                if shape.line and shape.line.width:
                                    line_width = shape.line.width
                            except Exception:
                                pass

                            # Copy Font Name & Color safely
                            try:
                                if (
                                    shape.has_text_frame
                                    and shape.text_frame.paragraphs
                                ):
                                    p_ref = shape.text_frame.paragraphs[0]
                                    if p_ref.runs:
                                        r_ref = p_ref.runs[0]
                                        if r_ref.font.name:
                                            font_name = r_ref.font.name
                                        if (
                                            r_ref.font.color
                                            and r_ref.font.color.rgb
                                        ):
                                            font_color = r_ref.font.color.rgb
                            except Exception:
                                pass

                            # ABSOLUTELY SKIP FURTHER DELETION LOGIC FOR THESE BOXES
                            continue

                        # CONDITIONAL SCAN: Identify ONLY old Size Box or unbound ghost text in middle gap
                        is_size_text = "size" in txt_lower or (
                            "x" in txt_lower and len(txt) < 25
                        )
                        is_in_middle_gap = (
                            shape.top > Pt(380)
                            and shape.left > Pt(200)
                            and shape.left < Pt(450)
                        )

                        if is_size_text or is_in_middle_gap:
                            shapes_to_delete.append(shape)

                    # STEP 2: DELETE ONLY SIZE SHAPES / GHOST TEXTS
                    for shape in shapes_to_delete:
                        try:
                            if shape.has_text_frame:
                                shape.text_frame.text = ""
                            sp = shape._element
                            sp.getparent().remove(sp)
                        except Exception:
                            pass

                    # STEP 3: CREATE FRESH BOLD 16PT SIZE BOX
                    if size_val and size_val.lower() != "nan":
                        final_text = (
                            size_val
                            if size_val.lower().startswith("size")
                            else f"Size: {size_val}"
                        )

                        # Middle Gap & Position Calculation
                        box_left = media_right_edge + Pt(12)
                        box_width = (qty_left_edge - media_right_edge) - Pt(24)

                        if box_width < Pt(80):
                            box_width = Pt(130)

                        final_top = (
                            ref_top if ref_top is not None else Pt(432)
                        )
                        final_height = (
                            ref_height if ref_height is not None else Pt(25)
                        )

                        # Add Shape Box
                        new_box = slide.shapes.add_shape(
                            MSO_SHAPE.RECTANGLE,
                            box_left,
                            final_top,
                            box_width,
                            final_height,
                        )

                        # Apply Outer Border Line & Transparent Fill
                        new_box.fill.background()
                        new_box.line.color.rgb = border_color
                        new_box.line.width = line_width

                        # Text Frame Setup & Vertical Middle Alignment
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
                        run.font.bold = True  # BOLD ENABLED
                        run.font.color.rgb = font_color

                        # Set Font Size 15 / 16
                        text_len = len(final_text)
                        if text_len > 16:
                            run.font.size = Pt(12)
                        elif text_len > 13:
                            run.font.size = Pt(14)
                        else:
                            run.font.size = Pt(16)  # Default 16pt Font Size

                # Save Presentation
                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success(
                    "🎉 Success! `Media Type:` aur `Qty:` ka text bilkul safe hai. Size Box Bold + 16pt font size ke saath perfectly place ho gaya hai!"
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
