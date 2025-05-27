import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import msoffcrypto

# Session state to manage comparison workflow
if 'workflow_state' not in st.session_state:
    st.session_state['workflow_state'] = 'waiting_for_compare'
if 'password' not in st.session_state:
    st.session_state['password'] = ''
if 'compared' not in st.session_state:
    st.session_state['compared'] = False

# Custom CSS to reduce padding and center header
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1rem;
    }
    h1 {
        color: #000000;
        font-weight: bold;
        text-align: center;
        margin-top: 0;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Header
st.markdown('<h1>DISCOVERY MEDICAL CONTRIBUTION CHANGE REPORT</h1>', unsafe_allow_html=True)

# Upload section layout
col_upload1, col_upload2 = st.columns([1, 1])
with col_upload1:
    st.markdown("📂 **Current Month**")
    may_file = st.file_uploader("Upload Current Month's Medical Excel File", type=["xlsx"], key="may_file")
with col_upload2:
    st.markdown("📂 **Prior Month**")
    april_file = st.file_uploader("Upload Prior Month's Medical Excel File", type=["xlsx"], key="april_file")

# Compare button
st.markdown("")
compare_button = st.button("Compare")

# Handle Compare button click
if compare_button and may_file and april_file:
    st.session_state['workflow_state'] = 'waiting_for_password'
    st.session_state['password'] = ''  # Reset password input

# Password prompt
if st.session_state['workflow_state'] == 'waiting_for_password':
    password = st.text_input("Please enter the password to decrypt the Excel files:", type="password", value=st.session_state['password'])
    if password:
        st.session_state['password'] = password
        st.session_state['workflow_state'] = 'comparison_complete'

# Function to load and clean data
def load_data(uploaded_file, password):
    if uploaded_file is None:
        return None, False
    try:
        decrypted = io.BytesIO()
        file_buffer = io.BytesIO(uploaded_file.read())
        office_file = msoffcrypto.OfficeFile(file_buffer)
        try:
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        except msoffcrypto.exceptions.InvalidKeyError:
            return None, True  # Indicates password failure
        else:
            decrypted.seek(0)

        df_raw = pd.read_excel(decrypted, sheet_name=0, skiprows=5)
        df_raw.columns = df_raw.columns.str.strip()

        expected_cols = ['CARD NUMBER', 'EMPLOYEE NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT']
        missing = [col for col in expected_cols if col not in df_raw.columns]
        if missing:
            st.error(f"Missing expected columns: {missing}")
            st.write("Available columns:", df_raw.columns.tolist())
            return None, False

        df = df_raw[expected_cols].copy()
        df = df.dropna(subset=['ID NUMBER', 'TOTAL AMOUNT'])
        df['ID NUMBER'] = df['ID NUMBER'].astype(str).str.strip()
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)
        return df, False
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None, False

# Function to merge and compare data
def compare_data(may_df, april_df):
    if may_df is None or april_df is None:
        return None
    merged = pd.merge(april_df, may_df, on='ID NUMBER', how='outer', suffixes=('_April', '_May'))
    merged['TOTAL AMOUNT_April'] = merged['TOTAL AMOUNT_April'].fillna(0)
    merged['TOTAL AMOUNT_May'] = merged['TOTAL AMOUNT_May'].fillna(0)
    merged['Difference'] = merged['TOTAL AMOUNT_May'] - merged['TOTAL AMOUNT_April']
    merged['Changed'] = merged['Difference'] != 0
    return merged.rename(columns={
        'MEMBER SURNAME_April': 'Name M1',
        'MEMBER SURNAME_May': 'Name M2',
        'ID NUMBER': 'ID No',
        'TOTAL AMOUNT_April': 'Amount M1',
        'TOTAL AMOUNT_May': 'Amount M2',
        'Difference': 'Difference',
        'Changed': 'Change'
    })

# Function to create PDF
def create_pdf(df):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=16)

    # Title
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(0, 10)
    pdf.cell(0, 10, txt="Rham Discovery Medical Contribution Change Report", ln=True, align='C')
    pdf.ln(2)

    col_names = ['Name M1', 'Name M2', 'ID No', 'Amount M1', 'Amount M2', 'Difference', 'Change']
    col_widths = [35, 35, 35, 19, 19, 19, 25]

    def print_header():
        pdf.set_font('Arial', style='B', size=8)
        pdf.set_fill_color(200, 200, 200)
        pdf.set_text_color(0, 0, 0)
        for i, col_name in enumerate(col_names):
            pdf.cell(col_widths[i], 10, col_name.upper(), border=1, align='C', fill=True)
        pdf.ln(10)
        pdf.set_font('Arial', style='', size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)

    print_header()

    page_height_limit = 275
    for _, row in df.iterrows():
        if pdf.get_y() + 10 > page_height_limit:
            pdf.add_page()
            pdf.set_font("Arial", style='B', size=16)
            pdf.set_text_color(0, 0, 0)
            pdf.set_xy(0, 10)
            pdf.cell(0, 10, txt="Rham Discovery Medical Contribution Change Report", ln=True, align='C')
            pdf.ln(2)
            print_header()

        pdf.cell(col_widths[0], 8, str(row.get('Name M1', '')), border=1)
        pdf.cell(col_widths[1], 8, str(row.get('Name M2', '')), border=1)
        pdf.cell(col_widths[2], 8, str(row.get('ID No', '')), border=1)
        pdf.cell(col_widths[3], 8, f"{row.get('Amount M1', 0):.2f}", border=1, align='R')
        pdf.cell(col_widths[4], 8, f"{row.get('Amount M2', 0):.2f}", border=1, align='R')
        pdf.cell(col_widths[5], 8, f"{row.get('Difference', 0):.2f}", border=1, align='R')
        changed_value = row.get('Change', False)
        if changed_value:
            pdf.set_text_color(255, 0, 0)
            pdf.cell(col_widths[6], 8, "Changed", border=1, align='C')
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.cell(col_widths[6], 8, "", border=1, align='C')
        pdf.ln()

    pdf_bytes = pdf.output(dest='S').encode('latin1')
    return io.BytesIO(pdf_bytes)

# Main logic with workflow state
if may_file and april_file:
    if st.session_state['workflow_state'] == 'comparison_complete':
        with st.spinner("Comparing files..."):
            password = st.session_state['password']
            may_df, password_failed_may = load_data(may_file, password)
            april_df, password_failed_april = load_data(april_file, password)

            if password_failed_may or password_failed_april:
                st.error("Incorrect password, please re-enter password.")
                st.session_state['compared'] = False
                st.session_state['workflow_state'] = 'waiting_for_password'
            elif may_df is not None and april_df is not None:
                merged_df = compare_data(may_df, april_df)
                if merged_df is not None:
                    st.session_state['compared'] = True
                    st.session_state['merged_df'] = merged_df
                    st.success("Comparison completed successfully!")
                else:
                    st.error("Comparison failed due to data issues.")
                    st.session_state['compared'] = False
                    st.session_state['workflow_state'] = 'waiting_for_password'
            else:
                st.error("Failed to load one or both files.")
                st.session_state['compared'] = False
                st.session_state['workflow_state'] = 'waiting_for_password'
    elif st.session_state['workflow_state'] == 'waiting_for_compare':
        st.warning("Please click 'Compare' to start the process.")
else:
    st.warning("Please upload both Excel files to proceed.")

# Display results
if st.session_state['compared']:
    st.subheader("Comparison Report")
    st.dataframe(st.session_state['merged_df'][['Name M1', 'Name M2', 'ID No', 'Amount M1', 'Amount M2', 'Difference', 'Change']])

    csv = st.session_state['merged_df'].to_csv(index=False)
    st.download_button("Download Report as CSV", csv, "change_report.csv", "text/csv")

    pdf_buffer = create_pdf(st.session_state['merged_df'])
    st.download_button("Download Report as PDF", pdf_buffer, "Rham Medical Change report.pdf", "application/pdf")
