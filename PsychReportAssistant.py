import os
import tempfile
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
from pathlib import Path

def clean_filename(filename):
    # Remove characters that are not allowed in Windows filenames
    return "".join(c for c in filename if c.isalnum() or c in (' ', '.', '_')).rstrip()

def generate_report(data_row, template):
    # Render the template with data
    report = template.render(data_row)
    return report

def main():
    # Set page configuration
    st.set_page_config(
        page_title="Report Assistant",
        page_icon="✅",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.title("Psych Report Assistant")

    uploaded_file = st.file_uploader("Select CSV or Excel file:", type=["csv", "xlsx"])

    if uploaded_file is not None:
        # Read data from CSV or Excel file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        else:
            st.warning("Unsupported file format. Please upload a CSV or Excel file.")
            st.stop()

        st.write("CSV or Excel file uploaded successfully.")

        uploaded_docx_file = st.file_uploader("Upload template.docx:", type=["docx"])

        if uploaded_docx_file is not None:
            # Load template
            template = DocxTemplate(uploaded_docx_file)

            if st.button("Generate Reports"):
                st.info("Generating reports...")

                # Define the output directory
                temp_dir = Path(tempfile.mkdtemp()) / "reports"
                temp_dir.mkdir(parents=True, exist_ok=True)

                # Generate reports
                for index, row in df.iterrows():
                    # Generate report
                    report = generate_report(row, template)

                    # Clean the filename to remove invalid characters
                    filename = clean_filename(f"{row['FirstName']}_{row['LastName']}_Report.docx")

                    # Save the report to the predefined subdirectory
                    output_path = temp_dir / filename
                    template.save(output_path)

                st.success(f"Reports generated successfully in '{temp_dir}'")


if __name__ == "__main__":
    main()
