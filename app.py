import streamlit as st
import pdfplumber
import re
import pandas as pd
import os
import tempfile
from io import BytesIO
import base64

def excel_field(excel_file):
    df = pd.read_excel(excel_file, sheet_name=0, header=None)
    
    # to avoid "FOCR data extraction exercise template" present at the top of the sheet
    header_row_index = 0
    for i, row in df.iterrows():
        non_null_count = row.notna().sum()
        if non_null_count >= 3:
            row_values = [str(val).strip().upper() for val in row if pd.notna(val)]
            row_text = ' '.join(row_values)
            
            ship_indicators = ['IMO', 'YEAR', 'GROSS', 'TONNAGE', 'EEDI', 'EEXI', 'DEADWEIGHT', 'SHIP', 'VESSEL']
            if any(indicator in row_text for indicator in ship_indicators):
                header_row_index = i
                break
    
    df_with_header = pd.read_excel(excel_file, sheet_name=0, header=header_row_index)
    
    # to extract fields from the header row found previously
    fields = []
    for col in df_with_header.columns:
        col_str = str(col).strip()
        if (not col_str.startswith('Unnamed') and 
            col_str != 'nan' and 
            len(col_str) > 0 and
            col_str.lower() != 'none'):
            fields.append(col_str)
    
    return fields, header_row_index

def extract_period_dates(text):
    period_data = {
        "Period start date": "Not found",
        "Period end date": "Not found"
    }
    
    start_patterns = [
        r"Period start date\s*:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})",
        r"Period start date\s*:\s*([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})",
        r"Period start date.*?([0-9]{4}-[0-9]{2}-[0-9]{2})",
        r"Start date\s*:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})"  
    ]
    
    end_patterns = [
        r"Period end date\s*:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})",
        r"Period end date\s*:\s*([0-9]{1,2}[-/][0-9]{1,2}[-/][0-9]{4})",
        r"Period end date.*?([0-9]{4}-[0-9]{2}-[0-9]{2})",
        r"End date\s*:\s*([0-9]{4}-[0-9]{2}-[0-9]{2})"
    ]
    
    for pattern in start_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            period_data["Period start date"] = match.group(1)
            break
    
    for pattern in end_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            period_data["Period end date"] = match.group(1)
            break
    
    return period_data

def extract_pdf_values(pdf_file, excel_file):
    extracted_data = {}
    raw_text = ""

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                raw_text += text + "\n"
            
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row:
                        table_text = " ".join([str(cell) if cell else "" for cell in row])
                        raw_text += "\n" + table_text
    
    fields_to_extract, _ = excel_field(excel_file)
    
    imo_number = "Not found"
    imo_patterns = [
        r"Particulars of ship\s*\n\s*(\d{7,8})",
        r"IMO\s*number\s*:?\s*(\d{7,8})",
        r"IMO\s*No\.?\s*:?\s*(\d{7,8})",
        r"IMO\s*(\d{7,8})"
    ]
    
    for pattern in imo_patterns:
        imo_match = re.search(pattern, raw_text, re.IGNORECASE)
        if imo_match:
            imo_number = imo_match.group(1)
            break
    
    if imo_number == "Not found":
        filename = pdf_file if isinstance(pdf_file, str) else getattr(pdf_file, 'name', str(pdf_file))
        filename_imo_match = re.search(r"IMO-?(\d{7,8})", filename, re.IGNORECASE)
        if filename_imo_match:
            imo_number = filename_imo_match.group(1)
    
    if "IMO number" in fields_to_extract:
        extracted_data["IMO number"] = imo_number
    
    period_dates = extract_period_dates(raw_text)
    for period_field, period_value in period_dates.items():
        if period_field in fields_to_extract:
            extracted_data[period_field] = period_value
    
    for field in fields_to_extract:
        if field not in extracted_data:
            value = find_value_in_text(raw_text, field)
            extracted_data[field] = value
            
    return extracted_data, raw_text, period_dates

def find_value_in_text(text, field):
    
    if field == "clDIST (g CO2/m∙nm)":
        lines = text.split('\n')
        for i, line in enumerate(lines):
            line_upper = line.upper()
            if any(term in line_upper for term in ['CLDIST', 'CIDIST', 'CL DIST', 'CI DIST']):
                number = re.search(r'([0-9]+\.?[0-9]*)', line)
                if number:
                    return number.group(1)
        return "Not found"
    
    if field == "EEPI (g CO2/t∙nm)":
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if 'EEPI' in line.upper() and 'EEOI' not in line.upper():
                number = re.search(r'([0-9]+\.?[0-9]*)', line)
                if number:
                    return number.group(1)
        return "Not found"
    
    if field == "EEOI (g CO2/t∙nm or others)":
        lines = text.split('\n')
        for line in lines:
            if 'EEOI' in line.upper() and ':' in line:
                after_colon = line.split(':')[-1].strip()
                number = re.search(r'([0-9]+\.?[0-9]*)', after_colon)
                if number:
                    return number.group(1)
        return "Not found"
    
    if field == "cbDIST (g CO2/berth∙nm)":
        lines = text.split('\n')
        for i, line in enumerate(lines):
            line_upper = line.upper()
            if any(term in line_upper for term in ['CBDIST', 'CB DIST', 'CBBERTH', 'BERTH']):
                number = re.search(r'([0-9]+\.?[0-9]*)', line)
                if number:
                    return number.group(1)
        return "Not found"
    
    if "Main propulsion power" in field:
        power_match = re.search(r"Main propulsion power\s*:\s*(\d+)", text, re.IGNORECASE)
        if power_match:
            return power_match.group(1)
        return "Not found"
    
    if "Auxiliary engine(s)" in field:
        aux_match = re.search(r"Auxiliary engine\(s\)\s*:\s*(\d+)", text, re.IGNORECASE)
        if aux_match:
            return aux_match.group(1)
        return "Not found"
    
    if "Distance travelled (nm)" in field:
        distance_match = re.search(r"Distance travelled \(nm\)\s*:\s*(\d+)", text, re.IGNORECASE)
        if distance_match:
            return distance_match.group(1)
        return "Not found"
    
    if "Hours underway (h)" in field:
        hours_match = re.search(r"Hours underway \(h\)\s*:\s*(\d+)", text, re.IGNORECASE)
        if hours_match:
            return hours_match.group(1)
        return "Not found"

    if "Attained annual operational CII before any correction" in field:
        patterns = [
            r"Attained annual operational CII before any\s*correction\s*:\s*([0-9\.]+)",
            r"Attained annual operational CII before any\s*correction.*?([0-9\.]+)",
            r"CII before any\s*correction\s*:\s*([0-9\.]+)"
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                return match.group(1)
        return "Not found"
    
    if "Attained EEDI" in field:
        eedi_pattern = r"Attained EEDI \(if applicable\) \(g[^:]*:\s*([^\n]*)"
        eedi_match = re.search(eedi_pattern, text, re.IGNORECASE | re.DOTALL)
        if eedi_match:
            value = eedi_match.group(1).strip()
            value = re.sub(r'[\.]{2,}', '', value).strip()
            return "Not applicable" if not value or len(value) == 0 else value
        
        eedi_dots = re.search(r"Attained EEDI.*?[\.]{3,}", text, re.IGNORECASE | re.DOTALL)
        if eedi_dots:
            return "Not applicable"
        
        return "Not found"

    if "Attained EEXI" in field:
        eexi_match = re.search(r"Attained EEXI \(if applicable\) \(g\s*([0-9\.]+)", text, re.IGNORECASE)
        if eexi_match:
            return eexi_match.group(1)
        return "Not found"
    
    if "Ice class" in field:
        ice_patterns = [
            r"Ice class \(if applicable\)\s*:\s*([^\n\r]*)",
            r"Ice class\s*:\s*([^\n\r]*)",
            r"Ice\s*class\s*\(if\s*applicable\)\s*:\s*([^\n\r]*)"
        ]
        for pattern in ice_patterns:
            ice_match = re.search(pattern, text, re.IGNORECASE)
            if ice_match:
                value = ice_match.group(1).strip()
                value = re.sub(r'\s+', ' ', value)
                value = re.sub(r'[\.]{2,}', '', value)
                value = value.strip(" ,.;:()")
                
                if not value or value.isspace():
                    return "Not applicable"
                
                if value and len(value) > 0:
                    return value
        return "Not found"
        
    if "DieselGasOil" in field:
        table_match = re.search(r"DieselGasOil\s+[\d\.]+\s+(\d+)", text, re.IGNORECASE)
        if table_match:
            return table_match.group(1)

    if "HeavyFuel" in field:
        table_match = re.search(r"HeavyFuel\s+[\d\.]+\s+(\d+)", text, re.IGNORECASE)
        if table_match:
            return table_match.group(1)

    if "LightFuel" in field:
        table_match = re.search(r"LightFuel\s+[\d\.]+\s+(\d+)", text, re.IGNORECASE)
        if table_match:
            return table_match.group(1)
    
    pattern = rf"{re.escape(field)}\s*:\s*([^\n\r]*)"
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        value = match.group(1).strip()
        value = re.sub(r'\s+', ' ', value)
        value = re.sub(r'[\.]{2,}', '', value)
        value = value.strip(" ,.;:()")
        
        if not value or value.isspace():
            return "Not applicable"
        
        if value and len(value) > 0:
            return value
    
    return "Not found"

def get_download_link(df, filename):
    towrite = BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download {filename}</a>'

def main():
    st.set_page_config(
        page_title="PDF Ship Data Extraction Tool",
        page_icon="🚢",
        layout="wide"
    )
    
    st.title("PDF Data Extraction Tool")
    st.markdown("Extract data from PDF files, fields are based on an Excel template")
    
    st.sidebar.header("File Upload")
    
    
    excel_file = st.sidebar.file_uploader(
        "Upload Excel Template", 
        type=['xlsx', 'xls'],
        help="Upload the Excel template with column headers"
    )
    
    pdf_files = st.sidebar.file_uploader(
        "Upload PDF Files", 
        type=['pdf'],
        accept_multiple_files=True,
        help="Upload one or more ship PDF files (preferably named IMO-*.pdf)"
    )
    
    if excel_file and pdf_files:
        
        st.header("Excel Template Analysis")
        fields, header_row = excel_field(excel_file)
        
        if len(fields) == 0:
            st.error("No valid column headers found in Excel template!")
            st.info("Please check that your Excel file has proper column headers.")
            return
        
        st.success(f"Excel template loaded with {len(fields)} columns")
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Template Fields:")
            for i, field in enumerate(fields[:10], 1):
                st.write(f"{i}. {field}")
            if len(fields) > 10:
                st.write(f"... and {len(fields) - 10} more fields")
        
        with col2:
            st.subheader("PDF Files:")
            st.write(f"Total files: {len(pdf_files)}")
            for pdf in pdf_files[:5]:
                st.write(f"• {pdf.name}")
            if len(pdf_files) > 5:
                st.write(f"... and {len(pdf_files) - 5} more files")
        
        if st.button("Start Extraction", type="primary"):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            all_data = []
            
            for i, pdf_file in enumerate(pdf_files):
                
                progress = (i + 1) / len(pdf_files)
                progress_bar.progress(progress)
                status_text.text(f"Processing {i+1}/{len(pdf_files)}: {pdf_file.name}")
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(pdf_file.read())
                    tmp_file_path = tmp_file.name
                
                extracted_data, raw_text, period_dates = extract_pdf_values(tmp_file_path, excel_file)
                
                os.unlink(tmp_file_path)
                
                row_data = {"Filename": pdf_file.name}
                for field in fields:
                    if field in extracted_data:
                        row_data[field] = extracted_data[field]
                    else:
                        value = find_value_in_text(raw_text, field)
                        row_data[field] = value
                
                all_data.append(row_data)
            
            progress_bar.empty()
            status_text.empty()
            
            if all_data:
                df_results = pd.DataFrame(all_data)
                column_order = ["Filename"] + fields
                df_results = df_results.reindex(columns=column_order)
                
                st.header("Extraction Results")
                
                data_columns = [col for col in df_results.columns if col != "Filename"]
                total_cells = len(all_data) * len(data_columns)
                success_count = 0
                not_found_count = 0
                not_applicable_count = 0
                
                for _, row in df_results.iterrows():
                    for col in data_columns:
                        value = row[col]
                        if value == "Not found":
                            not_found_count += 1
                        elif value == "Not applicable":
                            not_applicable_count += 1
                        else:
                            success_count += 1
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Cells", total_cells)
                with col2:
                    st.metric("Successfully Extracted", f"{success_count} ({success_count/total_cells*100:.1f}%)")
                with col3:
                    st.metric("Not Found", f"{not_found_count} ({not_found_count/total_cells*100:.1f}%)")
                with col4:
                    st.metric("Not Applicable", f"{not_applicable_count} ({not_applicable_count/total_cells*100:.1f}%)")
                
                st.subheader("Extracted PDF Data")
                st.dataframe(df_results, use_container_width=True)
                
                st.header("Download Results")
                st.markdown(
                    get_download_link(df_results, "extracted_ship_data.xlsx"),
                    unsafe_allow_html=True
                )
                
                st.success(f"Successfully processed {len(all_data)} PDF files!")
                
            else:
                st.error("No data extracted from any PDF files")
    
    else:
        st.info("Please upload an Excel template and PDF files to get started")
        
        st.markdown("""
        ### 📝 Instructions:
        1. **Upload Excel Template**: Upload your Excel file with the column headers you want to extract
        2. **Upload PDF Files**: Upload one or more PDF files (preferably named IMO-*.pdf)
        3. **Click Start Extraction**: The tool will process all PDFs and extract the required data
        4. **Download Results**: Get the extracted data as an Excel file
        5. **View Results**: See the extracted data in a table format
        6. **Check Metrics**: View extraction success metrics for better insights
        """)

if __name__ == "__main__":
    main()