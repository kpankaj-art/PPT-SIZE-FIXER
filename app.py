import pandas as pd
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
import streamlit as st

st.set_page_config(
    page_title="PPT Size & Remarks Automator", page_icon="📊", layout="centered"
)

st.title("📊 Auto-Fit Text PPT Automator")
st.write(
    "Yeh code text ko single line me rakhega aur long values ke liye font size auto-adjust karega."
)

st.markdown("---")

uploaded_excel = st.file_uploader("1. Upload Excel File (.xlsx)", type=["xlsx"])
uploaded_ppt = st.file_uploader("2. Upload PPT File (.pptx)", type=["pptx"])

if uploaded_excel and uploaded_ppt:
    if st.button("🚀 Process & Update PPT", type="primary"):
        with st.spinner("Updating Presentation Data..."):
            try:
                # 1. Read Excel safely
                xl_file = pd.ExcelFile(uploaded_excel)
                sheet_to_use = (
                    "Merged_Result"
                    if "Merged_Result" in xl_file.sheet_names
                    else xl_file.sheet_names[0]
                )
                df = pd.read_excel(uploaded_excel, sheet_name=sheet_to_use)

                # 2. Find Size & Remarks Columns
                size_column_name = None
                remarks_column_name = None

                for col in df.columns:
                    col_lower = str(col).lower()
                    if "size" in col_lower and not size_column_name:
                        size_column_name = col
                    if any(k in col_lower for k in ["remark", "remarks", "comment"]) and not remarks_column_name:
                        remarks_column_name = col

                prs = Presentation(uploaded_ppt)

                # Process Slides
                for i, slide in enumerate(prs.slides):
                    # Fetch Excel Values
                    try:
                        if size_column_name:
                            size_val = str(df[size_column_name].iloc[i]).strip()
                        else:
                            size_val = str(df.iloc[i, 7]).strip()
                    except Exception:
                        size_val = ""

                    try:
                        if remarks_column_name:
                            remarks_val = str(df[remarks_column_name].iloc[i]).strip()
                        else:
                            remarks_val = ""
                    except Exception:
                        remarks_val = ""

                    # Clean newlines or trailing spaces
                    size_val = " ".join(size_val.split())
                    remarks_val = " ".join(remarks_val.split())

                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            txt = shape.text_frame.text.strip()
                            txt_lower = txt.lower()

                            is_excluded = any(
                                k in txt_lower
                                for k in ["outlet", "address", "mobile"]
                            )

                            # Identify Size Box
                            is_size_box = (
                                (
                                    "size" in txt_lower
                                    or ("x" in txt_lower and len(txt) < 25)
                                )
                                and not any(
                                    k in txt_lower
                                    for k in ["media", "qty", "remark"]
                                )
                                and not is_excluded
                            )

                            # Identify Remarks Box
                            is_remarks_box = (
                                any(k in txt_lower for k in ["remark", "remarks"])
                                and not is_excluded
                            )

                            # UPDATE SIZE BOX
                            if is_size_box and size_val and size_val.lower() != "nan":
                                tf = shape.text_frame
                                tf.word_wrap = False  # Stacking hone se roke

                                final_size_text = (
                                    size_val
                                    if size_val.lower().startswith("size")
                                    else f"Size: {size_val}"
                                )

                                # Clear extra paragraphs
                                while len(tf.paragraphs) > 1:
                                    p_elem = tf.paragraphs[-1]._p
                                    p_elem.getparent().remove(p_elem)

                                p = tf.paragraphs[0]
                                p.text = final_size_text
                                p.alignment = PP_ALIGN.CENTER

                                # Dynamic Auto-Font Sizing based on length
                                text_len = len(final_size_text)
                                font_size = Pt(14) if text_len < 12 else (Pt(11) if text_len < 16 else Pt(9.5))

                                for run in p.runs:
                                    run.font.size = font_size
                                    run.font.bold = True

                            # UPDATE REMARKS BOX
                            if is_remarks_box and remarks_val and remarks_val.lower() != "nan":
                                tf = shape.text_frame
                                tf.word_wrap = False

                                final_remarks_text = (
                                    remarks_val
                                    if remarks_val.lower().startswith("remark")
                                    else f"Remarks: {remarks_val}"
                                )

                                while len(tf.paragraphs) > 1:
                                    p_elem = tf.paragraphs[-1]._p
                                    p_elem.getparent().remove(p_elem)

                                p = tf.paragraphs[0]
                                p.text = final_remarks_text

                                for run in p.runs:
                                    run.font.size = Pt(11)

                output_ppt_path = "Updated_Presentation.pptx"
                prs.save(output_ppt_path)

                st.success("🎉 PPT Successfully Updated! Text bilkul perfectly fit ho gaya hai.")

                with open(output_ppt_path, "rb") as file:
                    st.download_button(
                        label="📥 Download Updated PPT",
                        data=file,
                        file_name="Updated_Presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

            except Exception as e:
                st.error(f"❌ Error Aaya: {e}")
