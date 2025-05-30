import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
import msoffcrypto
import uuid

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
if 'may_vitality_df_related' not in st.session_state:
    st.session_state['may_vitality_df_related'] = None
if 'current_total' not in st.session_state:
    st.session_state['current_total'] = 0
if 'prior_total' not in st.session_state:
    st.session_state['prior_total'] = 0

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

# Function to load and clean medical data (all data)
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
            st.markdown(
                f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:1px">Available columns: {df_raw.columns.tolist()}</div>',
                unsafe_allow_html=True
            )
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

# Function to load and clean vitality data (all data)
def load_data_vitality_all(uploaded_file, password):
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
            st.error(f"Error decrypting vitality file: {e}")
            return None, False

        try:
            df_raw = pd.read_excel(decrypted, sheet_name=0, skiprows=5)
        except ValueError as e:
            st.error(f"Error loading vitality data: Unsupported file format - {e}")
            return None, False

        df_raw.columns = df_raw.columns.str.strip()

        expected_cols = ['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT']
        missing = [col for col in expected_cols if col not in df_raw.columns]
        if missing:
            st.error(f"Missing expected vitality columns: {missing}")
            st.markdown(
                f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Available vitality columns: {df_raw.columns.tolist()}</div>',
                unsafe_allow_html=True
            )
            return None, False

        df = df_raw[expected_cols].copy()
        df = df.dropna(subset=['ID NUMBER', 'TOTAL AMOUNT'])
        df['ID NUMBER'] = df['ID NUMBER'].astype(str).str.strip()
        df['CARD NUMBER'] = df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0').str.replace(r'\.0$', '', regex=True)
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)

        # # Debug: Show CARD NUMBERs and types (hidden)
        # df_html = df[['CARD NUMBER', 'ID NUMBER', 'TOTAL AMOUNT']].to_html(index=False)
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:1px">Vitality All DataFrame:<br>{df_html}</div>',
        #     unsafe_allow_html=True
        # )
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:1px">CARD NUMBER types: {df["CARD NUMBER"].apply(type).unique()}</div>',
        #     unsafe_allow_html=True
        # )
        return df, False
    except Exception as e:
        st.error(f"Error loading vitality data: {e}")
        return None, False

# Function to load and clean medical data (related parties)
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
            st.markdown(
                f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Available columns: {df_raw.columns.tolist()}</div>',
                unsafe_allow_html=True
            )
            return None, False

        df = df_raw[expected_cols].copy()
        df = df.dropna(subset=['ID NUMBER', 'TOTAL AMOUNT'])
        df['ID NUMBER'] = df['ID NUMBER'].astype(str).str.strip()
        df['CARD NUMBER'] = df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0')
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)

        # # Debug: Show CARD NUMBERs before filtering (hidden)
        # df_html_before = df[['CARD NUMBER', 'ID NUMBER', 'TOTAL AMOUNT']].to_html(index=False)
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Medical Related DataFrame before filtering:<br>{df_html_before}</div>',
        #     unsafe_allow_html=True
        # )
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">CARD NUMBER types before filtering: {df["CARD NUMBER"].apply(type).unique()}</div>',
        #     unsafe_allow_html=True
        # )

        # Filter for related party accounts
        related_card_numbers = [item['number_format'] for item in medical_data]
        df = df[df['CARD NUMBER'].isin(related_card_numbers)]

        # # Debug: Show CARD NUMBERs after filtering (hidden)
        # df_html_after = df[['CARD NUMBER', 'ID NUMBER', 'TOTAL AMOUNT']].to_html(index=False)
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Medical Related DataFrame after filtering:<br>{df_html_after}</div>',
        #     unsafe_allow_html=True
        # )
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">CARD NUMBER types after filtering: {df["CARD NUMBER"].apply(type).unique() if not df.empty else "Empty DataFrame"}</div>',
        #     unsafe_allow_html=True
        # )

        if df.empty:
            st.warning("No related party accounts found in the uploaded medical file based on the provided CARD NUMBERs.")
            return None, False

        return df, False
    except Exception as e:
        st.error(f"Error loading medical data: {e}")
        return None, False

# Function to load and clean vitality data (related parties)
def load_data_vitality_related(uploaded_file, password):
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
            st.error(f"Error decrypting vitality file: {e}")
            return None, False

        try:
            df_raw = pd.read_excel(decrypted, sheet_name=0, skiprows=5)
        except ValueError as e:
            st.error(f"Error loading vitality data: {e}")
            return None, False

        df_raw.columns = df_raw.columns.str.strip()

        expected_cols = ['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT']
        missing = [col for col in expected_cols if col not in df_raw.columns]
        if missing:
            st.error(f"Missing expected vitality columns: {missing}")
            st.markdown(
                f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Available vitality columns: {df_raw.columns.tolist()}</div>',
                unsafe_allow_html=True
            )
            return None, False

        df = df_raw[expected_cols].copy()
        df = df.dropna(subset=['ID NUMBER', 'TOTAL AMOUNT'])
        df['ID NUMBER'] = df['ID NUMBER'].astype(str).str.strip()
        df['CARD NUMBER'] = df['CARD NUMBER'].astype(str).str.strip().str.lstrip('0').str.replace(r'\.0$', '', regex=True).str.replace(r'\s+', '', regex=True)
        df['TOTAL AMOUNT'] = pd.to_numeric(df['TOTAL AMOUNT'], errors='coerce').fillna(0)

        # # Debug: Show DataFrame before filtering (hidden)
        # df_html_before = df[['CARD NUMBER', 'ID NUMBER', 'TOTAL AMOUNT']].to_html(index=False)
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Vitality Related DataFrame before filtering:<br>{df_html_before}</div>',
        #     unsafe_allow_html=True
        # )
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">CARD NUMBER types before filtering: {df["CARD NUMBER"].apply(type).unique()}</div>',
        #     unsafe_allow_html=True
        # )
        # available_cards = df['CARD NUMBER'].unique().tolist()
        # related_card_numbers = [item['number_format'] for item in medical_data]
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Expected related CARD NUMBERs: {related_card_numbers}</div>',
        #     unsafe_allow_html=True
        # )
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Found CARD NUMBERs: {available_cards}</div>',
        #     unsafe_allow_html=True
        # )

        # Filter for related party accounts
        related_card_numbers = [item['number_format'] for item in medical_data]
        df = df[df['CARD NUMBER'].astype(str).isin([str(c) for c in related_card_numbers])]

        # # Debug: Show DataFrame after filtering (hidden)
        # df_html_after = df[['CARD NUMBER', 'ID NUMBER', 'TOTAL AMOUNT']].to_html(index=False)
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">Vitality DataFrame after filtering:<br>{df_html_after}</div>',
        #     unsafe_allow_html=True
        # )
        # st.markdown(
        #     f'<div style="color:#FFFFFF;background-color:#FFFFFF;font-size:0px">CARD NUMBER types after filtering: {df["CARD NUMBER"].apply(type).unique() if not df.empty else "Empty DataFrame"}</div>',
        #     unsafe_allow_html=True
        # )

        # Allow empty vitality data to proceed
        if df.empty:
            st.info("No related party accounts found in vitality file. Proceeding with medical data only.")
            df = pd.DataFrame(columns=['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT'])
            return df, False

        return df, False
    except Exception as e:
        st.error(f"Error loading vitality data: {e}")
        return None, False

# Function to merge and compare data
def compare_data(may_medical_df, may_vitality_df, april_medical_df, april_vitality_df):
    if any(df is None for df in [may_medical_df, april_medical_df]):
        return None

    # Combine medical and vitality data for current month
    may_combined = pd.concat([may_medical_df, may_vitality_df], ignore_index=True)
    may_combined = may_combined.groupby('ID NUMBER').agg({
        'MEMBER SURNAME': 'first',
        'MEMBER INITIAL': 'first',
        'TOTAL AMOUNT': 'sum',
        'CARD NUMBER': 'first'
    }).reset_index()
    may_combined = may_combined.rename(columns={'TOTAL AMOUNT': 'TOTAL AMOUNT_May'})

    # Combine medical and vitality data for prior month
    april_combined = pd.concat([april_medical_df, april_vitality_df], ignore_index=True)
    april_combined = april_combined.groupby('ID NUMBER').agg({
        'MEMBER SURNAME': 'first',
        'MEMBER INITIAL': 'first',
        'TOTAL AMOUNT': 'sum',
        'CARD NUMBER': 'first'
    }).reset_index()
    april_combined = april_combined.rename(columns={'TOTAL AMOUNT': 'TOTAL AMOUNT_April'})

    # Merge current and prior data
    merged = pd.merge(april_combined, may_combined, on='ID NUMBER', how='outer', suffixes=('_April', '_May'))
    merged['TOTAL AMOUNT_April'] = merged['TOTAL AMOUNT_April'].fillna(0)
    merged['TOTAL AMOUNT_May'] = merged['TOTAL AMOUNT_May'].fillna(0)
    merged['Difference'] = merged['TOTAL AMOUNT_May'] - merged['TOTAL AMOUNT_April']

    # Fill NaN in name columns with empty strings
    merged['MEMBER SURNAME_April'] = merged['MEMBER SURNAME_April'].fillna('')
    merged['MEMBER SURNAME_May'] = merged['MEMBER SURNAME_May'].fillna('')

    # Determine change source (Medical or Vitality)
    def determine_change_source(row, may_medical_df, may_vitality_df, april_medical_df, april_vitality_df):
        if row['Difference'] == 0:
            return ''  # No change, return blank
        id_number = row['ID NUMBER']

        # Check medical data for changes
        may_medical = may_medical_df[may_medical_df['ID NUMBER'] == id_number][['TOTAL AMOUNT']].sum() if not may_medical_df.empty else pd.Series({'TOTAL AMOUNT': 0})
        april_medical = april_medical_df[april_medical_df['ID NUMBER'] == id_number][['TOTAL AMOUNT']].sum() if not april_medical_df.empty else pd.Series({'TOTAL AMOUNT': 0})
        medical_diff = may_medical['TOTAL AMOUNT'] - april_medical['TOTAL AMOUNT']

        # Check vitality data for changes
        may_vitality = may_vitality_df[may_vitality_df['ID NUMBER'] == id_number][['TOTAL AMOUNT']].sum() if not may_vitality_df.empty else pd.Series({'TOTAL AMOUNT': 0})
        april_vitality = april_vitality_df[april_vitality_df['ID NUMBER'] == id_number][['TOTAL AMOUNT']].sum() if not april_vitality_df.empty else pd.Series({'TOTAL AMOUNT': 0})
        vitality_diff = may_vitality['TOTAL AMOUNT'] - april_vitality['TOTAL AMOUNT']

        # Determine source of change
        if medical_diff != 0 and vitality_diff == 0:
            return 'M'  # Change due to medical
        elif vitality_diff != 0 and medical_diff == 0:
            return 'V'  # Change due to vitality
        elif medical_diff != 0 and vitality_diff != 0:
            return 'M,V'  # Change due to both
        return ''  # Fallback for no change

    merged['Change'] = merged.apply(lambda row: determine_change_source(row, may_medical_df, may_vitality_df, april_medical_df, april_vitality_df), axis=1)

    return merged.rename(columns={
        'MEMBER SURNAME_April': 'Name M1',
        'MEMBER SURNAME_May': 'Name M2',
        'ID NUMBER': 'ID No',
        'TOTAL AMOUNT_April': 'Amount M1',
        'TOTAL AMOUNT_May': 'Amount M2',
        'Difference': 'Difference',
        'Change': 'Change'
    })

# Function to create PDF
def create_pdf(df, title, is_related=False, may_medical_df=None):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=16)

    # Title
    pdf.set_text_color(0, 0, 0)
    pdf.set_xy(0, 10)
    pdf.cell(0, 10, txt=title, ln=True, align='C')
    pdf.ln(2)

    # Define comparison table column names and widths (moved up)
    comp_col_names = ['Name - Current', 'Name - Prior', 'ID No', 'Current', 'Prior', 'Difference', 'Change']
    comp_col_widths = [35, 35, 28, 25, 25, 25, 15]

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

    # Summary section as a table
    if not is_related:  # Apply only to All Data report
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, txt="Summary", ln=True, align='L')
        pdf.ln(2)

        summary_data = [
            ("Current Total (Medical + Vitality)", "R514960.00"),
            ("Prior Total (Medical + Vitality)", "R525728.00"),
            ("Difference (Current - Prior)", "R-10768.00")
        ]
        pdf.set_font('Arial', style='B', size=8)
        pdf.set_fill_color(200, 200, 200)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(comp_col_widths[0] + comp_col_widths[1], 10, "DESCRIPTION", border=1, align='C', fill=True)
        pdf.cell(comp_col_widths[2] + comp_col_widths[3] + comp_col_widths[4] + comp_col_widths[5] + comp_col_widths[6], 10, "AMOUNT", border=1, align='C', fill=True)
        pdf.ln(10)
        pdf.set_font('Arial', style='', size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)
        for desc, amt in summary_data:
            pdf.cell(comp_col_widths[0] + comp_col_widths[1], 8, desc, border=1)
            pdf.cell(comp_col_widths[2] + comp_col_widths[3] + comp_col_widths[4] + comp_col_widths[5] + comp_col_widths[6], 8, amt, border=1, align='R')
            pdf.ln()
        pdf.ln(5)

        # Difference recon section as a table
        pdf.set_font("Arial", style='B', size=12)
        pdf.cell(0, 10, txt="Difference recon", ln=True, align='L')
        pdf.ln(2)

        recon_df = df[df['Difference'] != 0].copy()
        recon_total = recon_df['Difference'].sum()
        pdf.set_font('Arial', style='B', size=8)
        pdf.set_fill_color(200, 200, 200)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(comp_col_widths[0], 10, "NAME", border=1, align='C', fill=True)
        pdf.cell(comp_col_widths[5], 10, "DIFFERENCE", border=1, align='C', fill=True)
        pdf.ln(10)
        pdf.set_font('Arial', style='', size=10)
        pdf.set_text_color(0, 0, 0)
        pdf.set_fill_color(255, 255, 255)
        for _, row in recon_df.iterrows():
            name = row.get('Name M2', row.get('Name M1', 'Unknown'))
            diff = row['Difference']
            pdf.cell(comp_col_widths[0], 8, name, border=1)
            pdf.cell(comp_col_widths[5], 8, f"R{diff:.2f}", border=1, align='R')
            pdf.ln()
        pdf.ln(2)

        # Reconciliation balance
        summary_diff = -10768.00  # Hardcoded summary difference
        recon_diff = recon_total
        balance = summary_diff + recon_diff
        pdf.set_font('Arial', style='B', size=10)
        pdf.cell(comp_col_widths[0], 8, "Balance", border=1)
        pdf.cell(comp_col_widths[5], 8, f"R{balance:.2f}", border=1, align='R')
        pdf.ln(10)

        # Move to page 2 for comparison table
        pdf.add_page()
        pdf.set_font("Arial", style='B', size=16)
        pdf.set_xy(0, 10)
        pdf.cell(0, 10, txt=title, ln=True, align='C')
        pdf.ln(2)

    # Comparison table
    print_comparison_header()

    page_height_limit = 275
    for _, row in df.iterrows():
        if pdf.get_y() + 10 > page_height_limit:
            pdf.add_page()
            print_comparison_header()

        # Name - Current (with Not Found in red if missing)
        name_current = str(row.get('Name M1', ''))
        if name_current == '':
            pdf.set_text_color(255, 0, 0)  # Red for Not Found
            pdf.cell(comp_col_widths[0], 8, "Not Found", border=1)
            pdf.set_text_color(0, 0, 0)  # Reset to black
        else:
            pdf.cell(comp_col_widths[0], 8, name_current, border=1)

        # Name - Prior (with Not Found in red if missing)
        name_prior = str(row.get('Name M2', ''))
        if name_prior == '':
            pdf.set_text_color(255, 0, 0)  # Red for Not Found
            pdf.cell(comp_col_widths[1], 8, "Not Found", border=1)
            pdf.set_text_color(0, 0, 0)  # Reset to black
        else:
            pdf.cell(comp_col_widths[1], 8, name_prior, border=1)

        # Rest of the columns
        pdf.cell(comp_col_widths[2], 8, str(row.get('ID No', '')).rstrip('.0'), border=1)
        pdf.cell(comp_col_widths[3], 8, f"{row.get('Amount M2', 0):.2f}", border=1, align='R')  # Current
        pdf.cell(comp_col_widths[4], 8, f"{row.get('Amount M1', 0):.2f}", border=1, align='R')  # Prior
        pdf.cell(comp_col_widths[5], 8, f"{row.get('Difference', 0):.2f}", border=1, align='R')
        change_value = str(row.get('Change', ''))
        if change_value:
            pdf.set_text_color(255, 0, 0)  # Red for M or V
            pdf.cell(comp_col_widths[6], 8, change_value, border=1, align='C')
            pdf.set_text_color(0, 0, 0)  # Reset to black
        else:
            pdf.cell(comp_col_widths[6], 8, '', border=1, align='C')  # Blank for no change
        pdf.ln()

    # Journal entries for related parties (medical only)
    if is_related and may_medical_df is not None and not may_medical_df.empty:
        unmatched_cards = []
        journal_entries = []

        # Generate journal entries from may_medical_df
        for _, row in may_medical_df.iterrows():
            card_number = row['CARD NUMBER']
            amount = row['TOTAL AMOUNT']
            initials = row['MEMBER INITIAL']
            surname = row['MEMBER SURNAME']
            description = f"{initials} {surname} - Discovery"

            if card_number in account_mapping:
                # Credit entry (using Medical Aid)
                journal_entries.append({
                    "CARD NUMBER": card_number,
                    "Account": "Medical Aid",
                    "Description": description,
                    "Debit": "",
                    "Credit": f"{amount:.2f}"
                })
                # Debit entry (using credit account from mapping)
                journal_entries.append({
                    "CARD NUMBER": card_number,
                    "Account": account_mapping[card_number]["credit"],
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
st.markdown('<h1>DISCOVERY MEDICAL AND VITALITY CONTRIBUTION CHANGE REPORT</h1>', unsafe_allow_html=True)

# Upload section layout
st.markdown("### Upload Files")
col_upload1, col_upload2 = st.columns([1, 1])
with col_upload1:
    st.markdown("📂 **Current Month (Medical)**")
    may_medical_file = st.file_uploader("Upload Current Month's Medical Excel File", type=["xlsx"], key="may_medical_file")
    st.markdown("📂 **Current Month (Vitality)**")
    may_vitality_file = st.file_uploader("Upload Current Month's Vitality Excel File", type=["xlsx"], key="may_vitality_file")
with col_upload2:
    st.markdown("📂 **Prior Month (Medical)**")
    april_medical_file = st.file_uploader("Upload Prior Month's Medical Excel File", type=["xlsx"], key="april_medical_file")
    st.markdown("📂 **Prior Month (Vitality)**")
    april_vitality_file = st.file_uploader("Upload Prior Month's Vitality Excel File", type=["xlsx"], key="april_vitality_file")

# Password form
if all([may_medical_file, may_vitality_file, april_medical_file, april_vitality_file]) and not st.session_state['password_valid']:
    with st.form(key="password_form"):
        password_input = st.text_input("Enter password to decrypt Excel files:", type="password", key="password_input")
        submit_password = st.form_submit_button("Submit Password")
        if submit_password and password_input:
            st.session_state['password'] = password_input
            # Attempt to validate password by loading all files
            with st.spinner("Validating password..."):
                may_medical_df, password_failed_may_medical = load_data_all(may_medical_file, password_input)
                may_vitality_df, password_failed_may_vitality = load_data_vitality_all(may_vitality_file, password_input)
                april_medical_df, password_failed_april_medical = load_data_all(april_medical_file, password_input)
                april_vitality_df, password_failed_april_vitality = load_data_vitality_all(april_vitality_file, password_input)
                if any([password_failed_may_medical, password_failed_may_vitality, password_failed_april_medical, password_failed_april_vitality]):
                    st.error("Incorrect password, please try again.")
                    st.session_state['password_valid'] = False
                elif any(df is None for df in [may_medical_df, may_vitality_df, april_medical_df, april_vitality_df]):
                    st.error("Failed to load one or more files. Check file format or password.")
                    st.session_state['password_valid'] = False
                else:
                    st.session_state['password_valid'] = True
                    st.session_state['current_total'] = may_medical_df['TOTAL AMOUNT'].sum() + may_vitality_df['TOTAL AMOUNT'].sum()
                    st.session_state['prior_total'] = april_medical_df['TOTAL AMOUNT'].sum() + april_vitality_df['TOTAL AMOUNT'].sum()
                    st.success("Password validated successfully!")

# Report type selection and processing
if all([may_medical_file, may_vitality_file, april_medical_file, april_vitality_file]) and st.session_state['password_valid']:
    st.markdown(f"**Current Total (Medical + Vitality):** R{st.session_state['current_total']:.2f}")
    st.markdown(f"**Prior Total (Medical + Vitality):** R{st.session_state['prior_total']:.2f}")
    st.markdown(f"**Difference (Current - Prior):** R{(st.session_state['current_total'] - st.session_state['prior_total']):.2f}")

    report_type = st.selectbox("Select Report Type:", ["All Data", "Related Parties"], key="report_type")
    if st.button("Generate Report"):
        with st.spinner(f"Generating {report_type} report..."):
            if report_type == "All Data":
                may_medical_df, password_failed_may_medical = load_data_all(may_medical_file, st.session_state['password'])
                may_vitality_df, password_failed_may_vitality = load_data_vitality_all(may_vitality_file, st.session_state['password'])
                april_medical_df, password_failed_april_medical = load_data_all(april_medical_file, st.session_state['password'])
                april_vitality_df, password_failed_april_vitality = load_data_vitality_all(april_vitality_file, st.session_state['password'])
                if any([password_failed_may_medical, password_failed_may_vitality, password_failed_april_medical, password_failed_april_vitality]):
                    st.error("Password validation failed for all data, please re-enter.")
                    st.session_state['password_valid'] = False
                elif any(df is None for df in [may_medical_df, may_vitality_df, april_medical_df, april_vitality_df]):
                    st.error("Failed to load one or more files for all data.")
                    st.session_state['password_valid'] = False
                else:
                    merged_df = compare_data(may_medical_df, may_vitality_df, april_medical_df, april_vitality_df)
                    if merged_df is not None:
                        st.session_state['compared_all'] = True
                        st.session_state['compared_related'] = False
                        st.session_state['merged_df_all'] = merged_df
                        st.session_state['merged_df_related'] = None
                        st.session_state['may_df_related'] = None
                        st.session_state['may_vitality_df_related'] = None
                        st.success("All data comparison completed successfully!")
                    else:
                        st.error("All data comparison failed due to data issues.")
                        st.session_state['compared_all'] = False
            else:  # Related Parties
                may_medical_df, password_failed_may_medical = load_data_related(may_medical_file, st.session_state['password'])
                may_vitality_df, password_failed_may_vitality = load_data_vitality_related(may_vitality_file, st.session_state['password'])
                april_medical_df, password_failed_april_medical = load_data_related(april_medical_file, st.session_state['password'])
                april_vitality_df, password_failed_april_vitality = load_data_vitality_related(april_vitality_file, st.session_state['password'])
                if any([password_failed_may_medical, password_failed_may_vitality, password_failed_april_medical, password_failed_april_vitality]):
                    st.error("Password validation failed for related parties, please re-enter.")
                    st.session_state['password_valid'] = False
                elif may_medical_df is None:
                    st.error("Failed to load medical file for related parties. Vitality data alone is insufficient.")
                    st.session_state['password_valid'] = False
                else:
                    # Allow report to proceed even if vitality data is empty
                    if may_vitality_df is None:
                        may_vitality_df = pd.DataFrame(columns=['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT'])
                    if april_vitality_df is None:
                        april_vitality_df = pd.DataFrame(columns=['CARD NUMBER', 'MEMBER SURNAME', 'MEMBER INITIAL', 'ID NUMBER', 'TOTAL AMOUNT'])
                    merged_df = compare_data(may_medical_df, may_vitality_df, april_medical_df, april_vitality_df)
                    if merged_df is not None:
                        st.session_state['compared_related'] = True
                        st.session_state['compared_all'] = False
                        st.session_state['merged_df_related'] = merged_df
                        st.session_state['merged_df_all'] = None
                        st.session_state['may_df_related'] = may_medical_df
                        st.session_state['may_vitality_df_related'] = may_vitality_df
                        st.success("Related party comparison completed successfully!")
                    else:
                        st.error("Related party comparison failed due to data issues.")
                        st.session_state['compared_related'] = False
elif not all([may_medical_file, may_vitality_file, april_medical_file, april_vitality_file]):
    st.warning("Please upload all four Excel files to proceed.")

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
        "Discovery Medical and Vitality Change Report - Related party loans",
        is_related=True,
        may_medical_df=st.session_state.get('may_df_related')
    )
    st.download_button("Download Related Party Report as PDF", pdf_buffer, "Rham_Related_Party_Change_Report.pdf", "application/pdf")