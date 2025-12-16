"""
Streamlit web interface for Carta Cap Table Transformer.
"""

import os
import tempfile
import streamlit as st

from carta_to_cap_table import run_transformation


st.set_page_config(page_title="Carta Cap Table Transformer", page_icon="📊")

st.title("📊 Carta Cap Table Transformer")
st.markdown("Transform your Carta export into a formatted cap table.")

# Check if default template exists (in templates/ folder relative to project root)
default_template_path = os.path.join(os.path.dirname(__file__), "..", "templates", "Cap Table Template.xlsx")
default_template_exists = os.path.exists(default_template_path)

# File uploaders
st.subheader("1. Upload Files")

carta_file = st.file_uploader(
    "Upload Carta Export",
    type=["xlsx"],
    help="Upload your Carta cap table export (.xlsx file)"
)

# Template selection
use_default_template = False
template_file = None

if default_template_exists:
    use_default_template = st.checkbox("Use default template", value=True)

if not use_default_template:
    template_file = st.file_uploader(
        "Upload Template",
        type=["xlsx"],
        help="Upload your cap table template (.xlsx file)"
    )

# Transform button
st.subheader("2. Transform")

can_transform = carta_file is not None and (use_default_template or template_file is not None)

if st.button("🔄 Transform", disabled=not can_transform, type="primary"):
    if not can_transform:
        st.warning("Please upload all required files.")
    else:
        try:
            with tempfile.TemporaryDirectory() as temp_dir:
                # Save uploaded Carta file
                carta_path = os.path.join(temp_dir, "carta_export.xlsx")
                with open(carta_path, "wb") as f:
                    f.write(carta_file.getvalue())

                # Determine template path
                if use_default_template:
                    template_path = default_template_path
                else:
                    template_path = os.path.join(temp_dir, "template.xlsx")
                    with open(template_path, "wb") as f:
                        f.write(template_file.getvalue())

                # Run transformation
                with st.spinner("Transforming cap table..."):
                    result = run_transformation(carta_path, template_path, temp_dir)

                # Store result in session state
                st.session_state.result = result
                with open(result["output_path"], "rb") as f:
                    st.session_state.output_data = f.read()
                st.session_state.output_filename = os.path.basename(result["output_path"])

                # Success message
                st.success("✅ Transformation complete!")

                # Display stats
                st.subheader("📈 Results")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Investors Processed", result["investors_processed"])
                with col2:
                    st.metric("Top Investors", result["top_investors"])
                with col3:
                    st.metric("Rolled Up (Other)", result["other_investors_count"])

                # Share classes mapped
                if result.get("share_classes_mapped"):
                    st.markdown("**Share Classes Mapped:** " + ", ".join(result["share_classes_mapped"]))

                # Validation warnings
                validation = result.get("validation", {})
                warnings = validation.get("warnings", [])
                if warnings:
                    st.subheader("⚠️ Validation Warnings")
                    for warning in warnings:
                        st.warning(warning)

        except Exception as e:
            st.error(f"❌ Error during transformation: {str(e)}")

# Download button (outside the transform block to persist)
if "output_data" in st.session_state and st.session_state.output_data:
    st.subheader("3. Download Result")
    st.download_button(
        label="📥 Download Transformed Cap Table",
        data=st.session_state.output_data,
        file_name=st.session_state.output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
