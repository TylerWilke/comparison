import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import msoffcrypto
import uuid
print("Hello")
# Session state initialization
if 'password' not in st.session_state:
    st.session_state['password'] = ''
if 'password_valid' not in st.session_state:
    st.session_state['password_valid'] = False
if 'compared_all' not in st.session_state:
    st.session_state['compared_all'] = False
if 'compared_related' not in st.session_state:
    st.session_state['compared_related'] = False
if 'merged_df_all' not in st.session_state:
    st.session_state['merged_df_all'] = None
if 'merged_df_related' not in st.session_state:
    st.session_state['merged_df_related'] = None
if 'may_df_related' not in st.session_state:
    st.session_state['may_df_related'] = None

# Medical data for related party accounts
medical_data = [
    {"CARD NUMBER": " 32893940", "number_format": "32893940", "id_number": "5211085134185", "member_surname": "ALCARAZ", "member_initial": "RJ", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 782418254", "number_format": "782418254", "id_number": "7311115139085", "member_surname": "HORN", "member_initial": "B", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 245834291", "number_format": "245834291", "id_number": "7903260226087", "member_surname": "LANGE", "member_initial": "M", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 559536990", "number_format": "559536990", "id_number": "8407241371086", "member_surname": "MALGAS", "member_initial": "B", "employee_number": "Related party Journal"},
    {"CARD NUMBER": " 651616480", "number_format": "651616480", "id_number": "8906235056082", "member_surname": "NAUDE", "member_initial": "SJE", "employee_number": "Related party Journal"},
]

# CARD NUMBER to Account mapping
account_mapping = {
    "32893940": {"credit": "Loans Receivable>Alcaraz Family Trust", "debit": "Medical Aid"},
    "782418254": {"credit": "Loans Receivable>Alcaraz Family Trust", "debit": "Medical Aid"},
    "245834291": {"credit": "Loans Receivable>Alcaraz Family Trust", "debit": "Medical Aid"},
    "559536990": {"credit": "Noboscope", "debit": "Medical Aid"},
    "651616480": {"credit": "Noboscope", "debit": "Medical Aid"},
}

# Function to load and clean data (all data)
def load_data_all(uploaded_file, password):
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
            return None, True
        except Exception as e:
            st.error(f"Error decrypting file: {e}")
            return None, False

        try:
            df_raw = pd.read_excel(decrypted, sheet_name=0, skiprows=5)
        except ValueError as e:
            st.error(f"Error loading data: Unsupported file format - {e}")
            return None, False

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
        df['CARD NUMBER'] = df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0')
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)
        return df, False
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return None, False

# Function to load and clean data (related parties)
def load_data_related(uploaded_file, password):
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
            return None, True
        except Exception as e:
            st.error(f"Error decrypting file: {e}")
            return None, False

        try:
            df_raw = pd.read_excel(decrypted, sheet_name=0, skiprows=5)
        except ValueError as e:
            st.error(f"Error loading data: Unsupported file format - {e}")
            return None, False

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
        df['CARD NUMBER'] = df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0')
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)

        # Filter for related party accounts
        related_card_numbers = [item['number_format'] for item in medical_data]
        df = df[df['CARD NUMBER'].isin(related_card_numbers)]

        if df.empty:
            st.warning("No related party accounts found in the uploaded file based on the provided CARD NUMBERs.")
            return None, False

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
def create_pdf(df, title, is_related=False, may_df=None):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=16)

    # Title
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(0, 10)
    pdf.cell(0, 10, txt=title, ln=True, align='C')
    pdf.ln(2)

    # Comparison table
    comp_col_names = ['Name M1', 'Name M2', 'ID No', 'Current', 'Prior', 'Difference', 'Change']
    comp_col_widths = [35, 35, 35, 19, 19, 19, 25]

    def print_comparison_header():
        pdf.set_font('Arial', style='B', size=8)
        pdf.set_fill_color(200, 200, 200)
        pdf.set_text_color(0, 0, 0)
        for i, col_name in enumerate(comp_col_names):
            pdf.cell(comp_col_widths[i], 10, col_name.upper(), border=1, align='C', fill=True)
        pdf.ln(10)
        pdf.set_font('Arial', style='', size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)

    print_comparison_header()

    page_height_limit = 275
    for _, row in df.iterrows():
        if pdf.get_y() + 10 > page_height_limit:
            pdf.add_page()
            pdf.set_font("Arial", style='B', size=16)
            pdf.set_xy(0, 10)
            pdf.cell(0, 10, txt=title, ln=True, align='C')
            pdf.ln(2)
            print_comparison_header()

        pdf.cell(comp_col_widths[0], 8, str(row.get('Name M1', '')), border=1)
        pdf.cell(comp_col_widths[1], 8, str(row.get('Name M2', '')), border=1)
        pdf.cell(comp_col_widths[2], 8, str(row.get('ID No', '')), border=1)
        pdf.cell(comp_col_widths[3], 8, f"{row.get('Amount M2', 0):.2f}", border=1, align='R')  # Current
        pdf.cell(comp_col_widths[4], 8, f"{row.get('Amount M1', 0):.2f}", border=1, align='R')  # Prior
        pdf.cell(comp_col_widths[5], 8, f"{row.get('Difference', 0):.2f}", border=1, align='R')
        changed_value = row.get('Change', False)
        if changed_value:
            pdf.set_text_color(255, 0, 0)
            pdf.cell(comp_col_widths[6], 8, "Changed", border=1, align='C')
            pdf.set_text_color(0, 0, 0)
        else:
            pdf.cell(comp_col_widths[6], 8, "", border=1, align='C')
        pdf.ln()

    # Journal table for related parties
    if is_related and may_df is not None:
        unmatched_cards = []
        journal_entries = []
        
        # Generate journal entries from may_df
        for _, row in may_df.iterrows():
            card_number = row['CARD NUMBER']
            amount = row['TOTAL AMOUNT']
            initials = row['MEMBER INITIAL']
            surname = row['MEMBER SURNAME']
            description = f"{initials} {surname} - Discovery"

            if card_number in account_mapping:
                # Credit entry
                journal_entries.append({
                    "CARD NUMBER": card_number,
                    "Account": account_mapping[card_number]["credit"],
                    "Description": description,
                    "Debit": "",
                    "Credit": f"{amount:.2f}"
                })
                # Debit entry
                journal_entries.append({
                    "CARD NUMBER": card_number,
                    "Account": account_mapping[card_number]["debit"],
                    "Description": description,
                    "Debit": f"{amount:.2f}",
                    "Credit": ""
                })
            else:
                unmatched_cards.append(card_number)

        if unmatched_cards:
            st.warning(f"No account mapping found for CARD NUMBERs: {', '.join(unmatched_cards)}")

    
        if journal_entries:
            pdf.ln(10)
            pdf.set_font("Arial", style='B', size=20)
            pdf.cell(0, 10, txt="Related Parties Medical Journal", ln=True, align='C')
            pdf.set_font("Arial", style='I', size=11)
            pdf.set_text_color(0, 40, 80)
            pdf.cell(0, 10, txt="(Automatically generated by the system without human intervention)", ln=True, align='C')
            pdf.set_text_color(0, 0, 0)
            pdf.set_font("Arial", style='', size=12)
            pdf.ln(2)
            

            journal_col_names = ['CARD NUMBER', 'Account', 'Description', 'Debit', 'Credit']
            journal_col_widths = [25, 70, 50, 22, 22]

            def print_journal_header():
                pdf.set_font('Arial', style='B', size=8)
                pdf.set_fill_color(200, 200, 200)
                pdf.set_text_color(0, 0, 0)
                for i, col_name in enumerate(journal_col_names):
                    pdf.cell(journal_col_widths[i], 10, col_name.upper(), border=1, align='C', fill=True)
                pdf.ln(10)
                pdf.set_font('Arial', style='', size=10)
                pdf.set_text_color(0, 0, 0)
                pdf.set_fill_color(255, 255, 255)

            print_journal_header()

            for entry in journal_entries:
                if pdf.get_y() + 10 > page_height_limit:
                    pdf.add_page()
                    pdf.set_font("Arial", style='B', size=20)
                    pdf.set_xy(0, 10)
                    pdf.cell(0, 10, txt=title, ln=True, align='C')
                    pdf.ln(2)
                    pdf.set_font("Arial", style='B', size=12)
                    pdf.cell(0, 10, txt="Journal: Related Parties Discovery Medical Aid", ln=True, align='C')
                    pdf.cell(0, 10, txt="(Automatically generated by the system without human intervention)", ln=True, align='C')
                    pdf.ln(2)
                    print_journal_header()

                pdf.cell(journal_col_widths[0], 8, entry['CARD NUMBER'], border=1)
                pdf.cell(journal_col_widths[1], 8, entry['Account'], border=1)
                pdf.cell(journal_col_widths[2], 8, entry['Description'], border=1)
                pdf.cell(journal_col_widths[3], 8, entry['Debit'], border=1, align='R')
                pdf.cell(journal_col_widths[4], 8, entry['Credit'], border=1, align='R')
                pdf.ln()

    pdf_bytes = pdf.output(dest='S').encode('latin1')
    return io.BytesIO(pdf_bytes)

# Custom CSS
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

# Password form
if may_file and april_file and not st.session_state['password_valid']:
    with st.form(key="password_form"):
        password_input = st.text_input("Enter password to decrypt Excel files:", type="password", key="password_input")
        submit_password = st.form_submit_button("Submit Password")
        if submit_password and password_input:
            st.session_state['password'] = password_input
            # Attempt to validate password by loading both files
            with st.spinner("Validating password..."):
                may_df_all, password_failed_may = load_data_all(may_file, password_input)
                april_df_all, password_failed_april = load_data_all(april_file, password_input)
                if password_failed_may or password_failed_april:
                    st.error("Incorrect password, please try again.")
                    st.session_state['password_valid'] = False
                elif may_df_all is None or april_df_all is None:
                    st.error("Failed to load one or both files. Check file format or password.")
                    st.session_state['password_valid'] = False
                else:
                    st.session_state['password_valid'] = True
                    st.success("Password validated successfully!")

# Report type selection and processing
if may_file and april_file and st.session_state['password_valid']:
    report_type = st.selectbox("Select Report Type:", ["All Data", "Related Parties"], key="report_type")
    if st.button("Generate Report"):
        with st.spinner(f"Generating {report_type} report..."):
            if report_type == "All Data":
                may_df, password_failed_may = load_data_all(may_file, st.session_state['password'])
                april_df, password_failed_april = load_data_all(april_file, st.session_state['password'])
                if password_failed_may or password_failed_april:
                    st.error("Password validation failed for all data, please re-enter.")
                    st.session_state['password_valid'] = False
                elif may_df is not None and april_df is not None:
                    merged_df = compare_data(may_df, april_df)
                    if merged_df is not None:
                        st.session_state['compared_all'] = True
                        st.session_state['compared_related'] = False
                        st.session_state['merged_df_all'] = merged_df
                        st.session_state['merged_df_related'] = None
                        st.session_state['may_df_related'] = None
                        st.success("All data comparison completed successfully!")
                    else:
                        st.error("All data comparison failed due to data issues.")
                        st.session_state['compared_all'] = False
                else:
                    st.error("Failed to load one or both files for all data.")
                    st.session_state['password_valid'] = False
                    st.session_state['compared_all'] = False
            else:  # Related Parties
                may_df, password_failed_may = load_data_related(may_file, st.session_state['password'])
                april_df, password_failed_april = load_data_related(april_file, st.session_state['password'])
                if password_failed_may or password_failed_april:
                    st.error("Password validation failed for related parties, please re-enter.")
                    st.session_state['password_valid'] = False
                elif may_df is not None and april_df is not None:
                    merged_df = compare_data(may_df, april_df)
                    if merged_df is not None:
                        st.session_state['compared_related'] = True
                        st.session_state['compared_all'] = False
                        st.session_state['merged_df_related'] = merged_df
                        st.session_state['merged_df_all'] = None
                        st.session_state['may_df_related'] = may_df  # Store may_df for journal
                        st.success("Related party comparison completed successfully!")
                    else:
                        st.error("Related party comparison failed due to data issues.")
                        st.session_state['compared_related'] = False
                else:
                    st.error("Failed to load one or both files for related parties.")
                    st.session_state['password_valid'] = False
                    st.session_state['compared_related'] = False
elif not (may_file and april_file):
    st.warning("Please upload both Excel files to proceed.")

# Display all data comparison results
if st.session_state['compared_all']:
    st.subheader("All Data Comparison Report")
    st.dataframe(st.session_state['merged_df_all'][['Name M1', 'Name M2', 'ID No', 'Amount M1', 'Amount M2', 'Difference', 'Change']])

    csv = st.session_state['merged_df_all'].to_csv(index=False)
    st.download_button("Download All Data Report as CSV", csv, "all_data_change_report.csv", "text/csv")

    pdf_buffer = create_pdf(st.session_state['merged_df_all'], "Rham Discovery Medical Contribution Change Report - All Data")
    st.download_button("Download All Data Report as PDF", pdf_buffer, "Rham_All_Data_Change_Report.pdf", "application/pdf")

# Display related party comparison results
if st.session_state['compared_related']:
    st.subheader("Related Party Comparison Report")
    st.dataframe(st.session_state['merged_df_related'][['Name M1', 'Name M2', 'ID No', 'Amount M1', 'Amount M2', 'Difference', 'Change']])

    csv = st.session_state['merged_df_related'].to_csv(index=False)
    st.download_button("Download Related Party Report as CSV", csv, "related_party_change_report.csv", "text/csv")

    pdf_buffer = create_pdf(
        st.session_state['merged_df_related'],
        "Related Party - Discovery Medical Aid Change Report",
        is_related=True,
        may_df=st.session_state.get('may_df_related')
    )
    st.download_button("Download Related Party Report as PDF", pdf_buffer, "Rham_Related_Party_Change_Report.pdf", "application/pdf")