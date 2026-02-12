import streamlit as st
from PyPDF2 import PdfMerger
import io

st.set_page_config(page_title="Financial Report Aggregator", page_icon="📊", layout="centered")

st.title("📊 Financial Report Aggregator")
st.markdown("Upload your financial reports and merge them into a single PDF package.")

st.markdown("---")

# Define the expected report order
report_types = [
    ("Balance Sheet (With Period Change)", "balance_sheet"),
    ("12 Month Statement", "statement"),
    ("Budget Comparison", "budget"),
    ("Rent Roll with Lease Charges", "rent_roll"),
    ("Aged Receivables / Aging Summary", "receivables"),
    ("Payables Aging Report", "payables"),
]

st.subheader("Upload Reports")
st.markdown("Upload each report in the slots below. They will be merged in the correct order.")

uploaded_files = {}

for i, (report_name, key) in enumerate(report_types, 1):
    uploaded_files[key] = st.file_uploader(
        f"{i}. {report_name}",
        type=["pdf"],
        key=key
    )

st.markdown("---")

# Property name for the output file
property_name = st.text_input("Property Name (for output filename)", value="Property")
month_year = st.text_input("Month/Year", value="02 2026")

# Merge button
if st.button("Merge PDFs", type="primary", use_container_width=True):
    # Check which files are uploaded
    files_to_merge = [(name, uploaded_files[key]) for (name, key) in report_types if uploaded_files[key] is not None]

    if len(files_to_merge) == 0:
        st.error("Please upload at least one PDF file.")
    else:
        try:
            merger = PdfMerger()

            for report_name, file in files_to_merge:
                merger.append(file)
                st.write(f"✅ Added: {report_name}")

            # Create output
            output = io.BytesIO()
            merger.write(output)
            merger.close()
            output.seek(0)

            # Generate filename
            filename = f"{property_name} Financials {month_year}.pdf".replace(" ", "_")

            st.success(f"Successfully merged {len(files_to_merge)} reports!")

            st.download_button(
                label="📥 Download Merged PDF",
                data=output,
                file_name=filename,
                mime="application/pdf",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Error merging PDFs: {str(e)}")

st.markdown("---")
st.markdown("*Reports will be merged in the order shown above.*")
