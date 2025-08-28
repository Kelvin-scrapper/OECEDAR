import os
import re
import pandas as pd
import xlwings as xw

# =====================================================================================
# HARDCODED HEADER STRUCTURE FOR THE PENSION DATA OUTPUT
# This defines the exact, non-negotiable structure of the output file.
# =====================================================================================
HARDCODED_HEADER_ROW_1 = [
    'OECDAR.EFFECTIVELABOUREXITAGE.AUS.A', 'OECDAR.EFFECTIVELABOUREXITAGE.BEL.A', 'OECDAR.EFFECTIVELABOUREXITAGE.BRA.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.CAN.A', 'OECDAR.EFFECTIVELABOUREXITAGE.CHE.A', 'OECDAR.EFFECTIVELABOUREXITAGE.CHL.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.CHN.A', 'OECDAR.EFFECTIVELABOUREXITAGE.CZE.A', 'OECDAR.EFFECTIVELABOUREXITAGE.DEU.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.ESP.A', 'OECDAR.EFFECTIVELABOUREXITAGE.FIN.A', 'OECDAR.EFFECTIVELABOUREXITAGE.FRA.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.GBR.A', 'OECDAR.EFFECTIVELABOUREXITAGE.GRC.A', 'OECDAR.EFFECTIVELABOUREXITAGE.HUN.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.IRL.A', 'OECDAR.EFFECTIVELABOUREXITAGE.ITA.A', 'OECDAR.EFFECTIVELABOUREXITAGE.JPN.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.KOR.A', 'OECDAR.EFFECTIVELABOUREXITAGE.MEX.A', 'OECDAR.EFFECTIVELABOUREXITAGE.NLD.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.NOR.A', 'OECDAR.EFFECTIVELABOUREXITAGE.NZL.A', 'OECDAR.EFFECTIVELABOUREXITAGE.POL.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.PRT.A', 'OECDAR.EFFECTIVELABOUREXITAGE.RUS.A', 'OECDAR.EFFECTIVELABOUREXITAGE.SWE.A',
    'OECDAR.EFFECTIVELABOUREXITAGE.TUR.A', 'OECDAR.EFFECTIVELABOUREXITAGE.USA.A'
]

HARDCODED_HEADER_ROW_2 = [
    "Pensions at a Glance, Effective labour market exit age, men, Australia", "Pensions at a Glance, Effective labour market exit age, men, Belgium",
    "Pensions at a Glance, Effective labour market exit age, men, Brazil", "Pensions at a Glance, Effective labour market exit age, men, Canada",
    "Pensions at a Glance, Effective labour market exit age, men, Switzerland", "Pensions at a Glance, Effective labour market exit age, men, Chile",
    "Pensions at a Glance, Effective labour market exit age, men, China", "Pensions at a Glance, Effective labour market exit age, men, Czech Republic",
    "Pensions at a Glance, Effective labour market exit age, men, Germany", "Pensions at a Glance, Effective labour market exit age, men, Spain",
    "Pensions at a Glance, Effective labour market exit age, men, Finland", "Pensions at a Glance, Effective labour market exit age, men, France",
    "Pensions at a Glance, Effective labour market exit age, men, United Kingdom", "Pensions at a Glance, Effective labour market exit age, men, Greece",
    "Pensions at a Glance, Effective labour market exit age, men, Hungary", "Pensions at a Glance, Effective labour market exit age, men, Ireland",
    "Pensions at a Glance, Effective labour market exit age, men, Italy", "Pensions at a Glance, Effective labour market exit age, men, Japan",
    "Pensions at a Glance, Effective labour market exit age, men, Korea", "Pensions at a Glance, Effective labour market exit age, men, Mexico",
    "Pensions at a Glance, Effective labour market exit age, men, Netherlands", "Pensions at a Glance, Effective labour market exit age, men, Norway",
    "Pensions at a Glance, Effective labour market exit age, men, New Zealand", "Pensions at a Glance, Effective labour market exit age, men, Poland",
    "Pensions at a Glance, Effective labour market exit age, men, Portugal", "Pensions at a Glance, Effective labour market exit age, men, Russia",
    "Pensions at a Glance, Effective labour market exit age, men, Sweden", "Pensions at a Glance, Effective labour market exit age, men, Turkey",
    "Pensions at a Glance, Effective labour market exit age, men, United States"
]
# =====================================================================================

def find_any_excel_file(root_directory='.'):
    """
    Scans to find the first valid Excel file, ignoring temporary/lock files.
    """
    print(f"Searching for any valid Excel file (.xlsx, .xls) in '{os.path.abspath(root_directory)}'...")
    for root, dirs, files in os.walk(root_directory):
        for file in files:
            if file.startswith('~$') or "OECDAR_DATA" in file: # Ignore temp files AND the output file itself
                continue
            if file.endswith(('.xlsx', '.xls')):
                found_path = os.path.join(root, file)
                print(f"\nFound valid Excel file to process: {found_path}")
                return found_path
    return None

def parse_pension_data_with_xlwings(file_path, sheet_name):
    """
    Uses xlwings to robustly read and parse the 'Pensions at a glance' file format.
    The logic is universal and does not depend on a fixed order in the source.
    """
    print("Using xlwings to open Excel in the background and read the data...")
    with xw.App(visible=False) as app:
        try:
            workbook = app.books.open(file_path)
            sheet = workbook.sheets[sheet_name]
            df = sheet.used_range.options(pd.DataFrame, header=False, index=False).value
            workbook.close()
        except Exception as e:
            print(f"CRITICAL ERROR: xlwings could not open or read the file. Error: {e}")
            return []

    print("Successfully read data with xlwings. Now parsing...")
    try:
        # Dynamically finds the row containing years
        year_row_index = df[df[0] == 'Time period'].index[0]
        years = df.iloc[year_row_index].tolist()
    except (IndexError, KeyError):
        print("Error: Could not find 'Time period' row in the data read from Excel.")
        return []

    structured_data = []
    # Iterates through rows, handling any order of countries
    for _, row in df.iloc[year_row_index + 1:].iterrows():
        # Clean the country name (e.g., "China (People’s Republic of)" -> "China")
        country_raw = str(row.get(0) or '').strip()
        if country_raw and not country_raw.startswith("©"):
            country = country_raw.split('(')[0].strip()
            # Iterates through columns, handling any order of years
            for i, value in enumerate(row):
                if i > 0 and value is not None and value != '':
                    try:
                        # Years can be floats in pandas, so we handle that
                        year = int(float(years[i]))
                        # The value is already a number
                        data_value = float(value)
                        structured_data.append({
                            "Country": country,
                            "Year": year,
                            "Value": data_value
                        })
                    except (ValueError, TypeError, IndexError):
                        # Skip if year or value is not a valid number
                        continue
    return structured_data

def create_output_with_hardcoded_structure(structured_data, output_filename):
    """
    Builds the final Excel file using the hardcoded Pensions header structure.
    """
    if not structured_data:
        print("No data was parsed. Cannot create output file.")
        return

    df = pd.DataFrame(structured_data)
    
    # Map from the human-readable description to the machine-code header
    desc_to_code_map = dict(zip(HARDCODED_HEADER_ROW_2, HARDCODED_HEADER_ROW_1))

    # For the special case of Czechia vs Czech Republic
    country_name_map = {"Czechia": "Czech Republic"}
    df['Country_normalized'] = df['Country'].replace(country_name_map)

    # Maps each data point to its correct column in the final, rigid template
    df['HeaderCode'] = df.apply(
        lambda r: desc_to_code_map.get(
            f"Pensions at a Glance, Effective labour market exit age, men, {r['Country_normalized']}"
        ),
        axis=1
    )
    
    df.dropna(subset=['HeaderCode'], inplace=True)
    if df.empty:
        print("Warning: Parsed data did not match any columns in the hardcoded template.")
        return

    # Pivot the data with years as rows and countries (via HeaderCode) as columns
    data_pivot = df.pivot_table(index='Year', columns='HeaderCode', values='Value')
    
    # Forces the output to conform to the exact header structure and order
    data_body = data_pivot.reindex(columns=HARDCODED_HEADER_ROW_1).sort_index()
    data_body.reset_index(inplace=True)

    # Assemble the final file
    header_df_1 = pd.DataFrame([HARDCODED_HEADER_ROW_1])
    header_df_2 = pd.DataFrame([HARDCODED_HEADER_ROW_2])
    header_df_1.insert(0, '', '')
    header_df_2.insert(0, '', '')
    
    final_data_body = data_body.astype(object).where(pd.notna(data_body), '')
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        header_df_1.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=0)
        header_df_2.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=1)
        final_data_body.to_excel(writer, sheet_name="Sheet1", index=False, header=False, startrow=2)
        
    print(f"\nSuccessfully created output file with hardcoded structure: '{output_filename}'")


# --- Main Execution Block ---
search_directory = '.'
source_sheet_name = "Table" # The name of the sheet to read in the source file
output_filename = "OECDAR_DATA.xlsx"

# 1. Find any valid Excel file to process
source_file_path = find_any_excel_file(search_directory)

if source_file_path:
    # 2. Parse the found file using the universal pensions logic
    structured_data = parse_pension_data_with_xlwings(source_file_path, source_sheet_name)
    
    # 3. Create the final output file with the hardcoded pensions structure
    create_output_with_hardcoded_structure(structured_data, output_filename)
else:
    print(f"\nOperation stopped: Could not find any valid Excel file in '{os.path.abspath(search_directory)}' or its subfolders.")