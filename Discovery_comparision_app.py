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

        # Replace nan with "None" for Name M1 and Name M2
        name_m1 = "None" if pd.isna(row.get('Name M1')) else str(row.get('Name M1', ''))
        name_m2 = "None" if pd.isna(row.get('Name M2')) else str(row.get('Name M2', ''))

        pdf.cell(col_widths[0], 8, name_m1, border=1)
        pdf.cell(col_widths[1], 8, name_m2, border=1)
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