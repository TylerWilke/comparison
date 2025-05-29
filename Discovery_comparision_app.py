

import streamlit as st
import pandas as pd
import io
import re

# Define medical_data with leading spaces
medical_data = [
    {"CARD NUMBER": " 32893940", "number_format": "32893940", "id_number": "5211085134185", "member_surname": "ALCARAZ", "member_initial": "RJ", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 782418254", "number_format": "782418254", "id_number": "7311115139085", "member_surname": "HORN", "member_initial": "B", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 245834291", "number_format": "245834291", "id_number": "7903260226087", "member_surname": "LANGE", "member_initial": "M", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 559536990", "number_format": "559536990", "id_number": "8407241371086", "member_surname": "MALGAS", "member_initial": "B", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 651616480", "number_format": "651616480", "id_number": "8906235056082", "member_surname": "NAUDE", "member_initial": "SJE", "employee_number": "Related party Journal"},
]

# Custom CSS
st.markdown(
    """
    <style>
    .block-container { padding-top: 1rem; }
    h1 { color: #000000; font-weight: bold; text-align: center; margin-top: 0; margin-bottom: 1rem; }
    </style>
    """,
    unsafe_allow_html=True
)

# Header
st.markdown('<h1>RELATED PARTY JOURNAL COMPARISON</h1>', unsafe_allow_html=True)

# File uploads
st.markdown("📂 **Upload Excel Files**")
col1, col2 = st.columns(2)
with col1:
    current_file = st.file_uploader("Current Month Excel", type=["xlsx"], key="current_file")
with col2:
    prior_file = st.file_uploader("Prior Month Excel", type=["xlsx"], key="prior_file")

# Process button
st.markdown("")
process_button = st.button("Compare Files")

def find_column(df, keywords):
    """Find column name containing keywords (case-insensitive, space-ignoring)."""
    df_cols = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True).str.lower()
    for col, orig_col in zip(df_cols, df.columns):
        if any(re.search(keyword.lower(), col) for keyword in keywords):
            return orig_col
    return None

def process_excel_file(uploaded_file, file_label):
    if not uploaded_file or uploaded_file.size == 0:
        st.error(f"No {file_label} file uploaded or file is empty.")
        return None

    try:
        # Read Excel with skiprows=5 (headers on row 6)
        df = pd.read_excel(uploaded_file, sheet_name='Sheet1', skiprows=5)
        df.columns = df.columns.str.strip().str.replace(r'\s+', ' ', regex=True)

        # Map required columns
        required_cols = {
            'CARD NUMBER': ['CARD NUMBER', 'CARD NO', 'CARDNUMBER'],
            'MEMBER SURNAME': ['MEMBER SURNAME', 'SURNAME'],
            'MEMBER INITIAL': ['MEMBER INITIAL', 'INITIAL', 'INITIALS'],
            'ID NUMBER': ['ID NUMBER', 'ID NO', 'IDNUMBER'],
            'TOTAL AMOUNT': ['TOTAL AMOUNT', 'AMOUNT', 'TOTAL']
        }
        col_mapping = {}
        for key, keywords in required_cols.items():
            col = find_column(df, keywords)
            if col:
                col_mapping[key] = col
            else:
                st.error(f"Cannot find column for {key} in {file_label} file. Expected names like: {keywords}")
                return None

        # Check for missing columns
        missing_cols = [key for key in required_cols if key not in col_mapping]
        if missing_cols:
            st.error(f"Missing columns in {file_label} file: {missing_cols}")
            st.write(f"Available columns:", df.columns.tolist())
            return None

        # Select and rename columns
        df = df[list(col_mapping.values())].copy()
        df.columns = col_mapping.keys()

        # Normalize CARD NUMBER
        df['CARD NUMBER'] = df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0')
        df['ID NUMBER'] = df['ID NUMBER'].astype(str).str.strip()
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)

        # Convert medical_data
        medical_df = pd.DataFrame(medical_data)
        medical_df['CARD NUMBER'] = medical_df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0')

        # Merge
        merged_df = pd.merge(
            df,
            medical_df[['CARD NUMBER', 'employee_number']],
            on='CARD NUMBER',
            how='left'
        )

        # Filter
        related_party_df = merged_df[merged_df['employee_number'] == 'Related party Journal']
        related_party_df = related_party_df.drop(columns=['employee_number'], errors='ignore')

        return related_party_df

    except Exception as e:
        st.error(f"Error processing {file_label} file: {e}")
        st.write("Ensure the file is a valid .xlsx, not password-protected, and has headers on row 6.")
        return None

def compare_related_parties(current_df, prior_df):
    if current_df is None or prior_df is None:
        return None

    # Select relevant columns
    current_df = current_df[['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT']].copy()
    prior_df = prior_df[['CARD NUMBER', 'TOTAL AMOUNT']].copy()

    # Rename prior month column
    prior_df = prior_df.rename(columns={'TOTAL AMOUNT': 'TOTAL AMOUNT PRIOR'})

    # Merge on CARD NUMBER
    comparison_df = pd.merge(
        current_df,
        prior_df,
        on='CARD NUMBER',
        how='outer'
    )

    # Fill NaN with 0
    comparison_df['TOTAL AMOUNT'] = comparison_df['TOTAL AMOUNT'].fillna(0)
    comparison_df['TOTAL AMOUNT PRIOR'] = comparison_df['TOTAL AMOUNT PRIOR'].fillna(0)

    # Calculate difference
    comparison_df['DIFFERENCE'] = comparison_df['TOTAL AMOUNT'] - comparison_df['TOTAL AMOUNT PRIOR']

    # Fill missing info from medical_data
    medical_df = pd.DataFrame(medical_data)
    medical_df['CARD NUMBER'] = medical_df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0')
    comparison_df = pd.merge(
        comparison_df,
        medical_df[['CARD NUMBER', 'member_surname', 'member_initial', 'id_number']],
        on='CARD NUMBER',
        how='left'
    )

    # Update columns with medical_data info
    comparison_df['MEMBER SURNAME'] = comparison_df['member_surname'].combine_first(comparison_df['MEMBER SURNAME'])
    comparison_df['MEMBER INITIAL'] = comparison_df['member_initial'].combine_first(comparison_df['MEMBER INITIAL'])
    comparison_df['ID NUMBER'] = comparison_df['id_number'].combine_first(comparison_df['ID NUMBER'])

    # Drop extra columns
    comparison_df = comparison_df.drop(columns=['member_surname', 'member_initial', 'id_number'], errors='ignore')

    # Reorder columns
    comparison_df = comparison_df[['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT', 'TOTAL AMOUNT PRIOR', 'DIFFERENCE']]

    return comparison_df

if process_button:
    if not current_file or not prior_file:
        st.error("Please upload both current and prior month Excel files.")
    else:
        with st.spinner("Processing files..."):
            current_df = process_excel_file(current_file, "current month")
            prior_df = process_excel_file(prior_file, "prior month")

            if current_df is not None or prior_df is not None:
                comparison_df = compare_related_parties(current_df, prior_df)

                if comparison_df is not None and not comparison_df.empty:
                    st.markdown("### Related Party Journal Comparison")
                    st.dataframe(comparison_df)

                    # Download CSV
                    csv_buffer = io.StringIO()
                    comparison_df.to_csv(csv_buffer, index=False)
                    csv_buffer.seek(0)
                    st.download_button(
                        label="Download as CSV",
                        data=csv_buffer.getvalue(),
                        file_name="related_party_comparison.csv",
                        mime="text/csv"
                    )
                else:
                    st.error("No Related party Journal entries found or error occurred. Check column names and file format.")
            else:
                st.error("Error processing one or both files. Check error messages above.")
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