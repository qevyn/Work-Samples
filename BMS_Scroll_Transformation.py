import pandas as pd
import sys
import re
import os
from datetime import datetime, timedelta
import calendar

def is_leap_year(year):
    """Check if the given year is a leap year."""
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)

def days_in_month(year, month):
    """Return the number of days in the given month/year."""
    return calendar.monthrange(year, month)[1]

def check_full_period(start_date, end_date):
    """Check if the date range spans full months or a full year."""
    start = datetime.strptime(start_date, '%m%d%Y')
    end = datetime.strptime(end_date, '%m%d%Y')
    
    # Check for full year
    if start.year != end.year:
        if start.month == 1 and start.day == 1 and end.month == 12 and end.day == 31:
            return True
        elif (end - start).days >= (365 if not is_leap_year(start.year) else 366):
            return True
    
    # Check for full months
    current = start
    full_months_count = 0
    while current <= end:
        month_days = days_in_month(current.year, current.month)
        month_end = current.replace(day=month_days)
        if month_end <= end:
            full_months_count += 1
            current = month_end + timedelta(days=1)
        else:
            # If we haven't reached the end, the last month isn't fully covered
            break
    
    # If we've iterated through all days up to 'end', then we have full months for the range in question
    return current > end and full_months_count > 0

def convert_date_format(date_str):
    """Split the date string at the space, take the left part, and convert to 'M/D/YYYY' format."""
    try:
        # Split the string at the first space and take the left part
        date_part = date_str.split(' ')[0]
        #date_obj = datetime.strptime(date_part, '%Y-%m-%d')
        #print(date_part, date_obj)
        return date_part #.strftime('%-m/%-d/%Y')
    except ValueError:
        return date_str  # Return original if conversion fails
    
def find_brand_and_indication(file_path):
    """Find the brand and indication from the Adobe Analytics suite in the Excel file."""
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name=0, header=None)

    # Mapping table for brand and indication
    mapping = {
        'www.sotyktu.com_US_US': ('Sotyktu', 'psoriasis'),
        'www.opdivo.com_US_US': ('Opdivo', 'cross indication'),
        'www.augtyro.com_US_US': ('Augtyro', 'cross indication'),
        'www.opdualag.com_US_US': ('Opdualag', 'cross indication'),
        'www.breyanzi.com_US_US': ('Breyanzi', 'cross indication'),
        'www.reblozyl.com_US_US': ('Reblozyl', 'cross indication'),
        'www.onureg.com_US_US': ('Onureg', 'cross indication'),
        'www.abecma.com_US_US': ('Abecma', 'cross indication'),
        'www.krazati.com_Global_AA': ('Krazati', 'cross indication'),
        'www.sprycel.com_US_US': ('Sprycel', 'cross indication'),
        'www.orencia.com_US_US': ('Orencia', 'cross indication'),
        'www.camzyos.com_US_US': ('Camzyos Branded', 'cross indication'),
        'www.hcmrealtalk.com_US_US': ('Camzyos Non-branded', 'cross indication'),
        'www.coulditbehcm.com_US_US': ('Camzyos ME', 'cross indication'),
        'www.zeposia.com_US_US': ('Zeposia', 'cross indication'),
        'cartautoimmune.com_US_AA': ('CarT', 'cross indication'),
        'www.cobenfy.com_US_US': ('Cobenfy', 'cross indication')
    }

    # Find Adobe Analytics suite
    adobe_suite = None
    for i, row in df.iterrows():
        if isinstance(row[0], str) and row[0].startswith("# Report suite: "):
            adobe_suite = row[0].split("# Report suite: ")[1].strip()
            break
    
    if adobe_suite is None:
        raise ValueError("Could not find the Adobe Analytics suite marker in the Excel file.")

    print(f"Found Adobe Analytics suite: {adobe_suite}")

    # Check if the found suite is in the mapping
    if adobe_suite not in mapping:
        raise ValueError(f"No mapping found for Adobe Analytics suite: {adobe_suite}")

    brand, indication = mapping[adobe_suite]
    return brand, indication

def find_report_date_range(file_path):
    """Find the report date range from the Excel file and format it as MMDDYYYY-MMDDYYYY."""
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name=0, header=None)

    # Find Adobe Analytics suite
    date_range = None
    for i, row in df.iterrows():
        if isinstance(row[0], str) and row[0].startswith("# Date: "):
            date_range = row[0].split("# Date: ")[1].strip()
            break
    
    if date_range is None:
        raise ValueError("Could not find the report date range marker in the Excel file.")

    # Parse and format the date range
    try:
        # Assuming the format is always "Month DD, YYYY - Month DD, YYYY"
        start_date, end_date = date_range.split(' - ')

        # Convert string dates to datetime objects
        start_dt = datetime.strptime(start_date, '%b %d, %Y')
        end_dt = datetime.strptime(end_date, '%b %d, %Y')

        # Format back to string with desired format
        formatted_date_range = f"{start_dt.strftime('%m%d%Y')}-{end_dt.strftime('%m%d%Y')}"
        
        print(f"Found Adobe Analytics suite date range: {formatted_date_range}")
        return formatted_date_range
    except ValueError:
        raise ValueError(f"Unexpected date format in the Excel file: {date_range}")

def extract_scroll_from_excel(file_path):
    """Extract the last freeform table from the Excel file."""
    # Read the entire Excel file
    xls = pd.ExcelFile(file_path)
    
    # Assume data is in the first sheet
    df = pd.read_excel(xls, sheet_name=0, header=None)

    # Identify the last occurrence of "##############################################"
    freeform_start_idx = None
    for i in range(len(df) - 1, -1, -1):
        if isinstance(df.iloc[i, 0], str) and "##############################################" in df.iloc[i, 0]:
            freeform_start_idx = i
            break

    if freeform_start_idx is None:
        raise ValueError("Could not locate the Freeform table (2) in the Excel file.")

    # Extract the table starting from the identified position
    df_table = df.iloc[freeform_start_idx + 1:].reset_index(drop=True)


    # Save the extracted table to a temporary Excel file
    temp_excel_path = "temp_extracted_table.xlsx"
    df_table.to_excel(temp_excel_path, index=False, header=False)

    return temp_excel_path

def transform_scroll_excel(file_path, brand, indication, date_range):
    """Transform the extracted table into the desired format."""
    df = pd.read_excel(file_path, header=[0, 1])  # Read the first two rows as multi-index header

    # Rename only the first level of columns to remove spaces
    new_columns = []
    for col in df.columns:
        new_col_name = (col[0].replace(' ', ''), col[1])
        new_columns.append(new_col_name)
    df.columns = pd.MultiIndex.from_tuples(new_columns, names=df.columns.names)

    # Drop the third row which seems to be the data we don't need
    df = df.iloc[1:].reset_index(drop=True)

    # Combine the multi-index header into a single row
    df.columns = [f"{col[0]} {col[1]}" for col in df.columns]
    
    # Rename the first column for clarity
    df.rename(columns={df.columns[0]: "PAGE"}, inplace=True)

    # Identify date columns dynamically based on 'YYYY-MM-DD HH:MM:SS' pattern within the combined header
    date_pattern = r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}"  # Matches formats like "2024-01-01 00:00:00"
    date_columns = [col for col in df.columns if re.search(date_pattern, col)]

    # Check if there are any date columns
    if not date_columns:
        print("No date columns found. Check the format of your data.")
        return

    # Melt the dataframe into long format
    df_melted = pd.melt(df, id_vars=["PAGE"], value_vars=date_columns, var_name="ACTIVITY_MONTH", value_name="VALUE")


    # Extract metric names from the combined header
    df_melted["METRIC"] = df_melted["ACTIVITY_MONTH"].apply(lambda x: x.split()[0])

    # Convert date format - this is where we need to ensure the conversion happens
    df_melted['ACTIVITY_MONTH'] = df_melted['ACTIVITY_MONTH'].apply(lambda x: convert_date_format(re.search(date_pattern, x).group() if re.search(date_pattern, x) else x))

    # Pivot back into the required format
    df_final = df_melted.pivot(index=["ACTIVITY_MONTH", "PAGE"], columns="METRIC", values="VALUE").reset_index()

    # Rename columns to match the desired output
    df_final.columns.name = None
    df_final = df_final.rename(columns={
        'PageScroll100%': 'PAGE SCROLL 100', 
        'PageScroll25%': 'PAGE SCROLL 25',
        'PageScroll50%': 'PAGE SCROLL 50', 
        'PageScroll75%': 'PAGE SCROLL 75', 
        'PageViews': 'PAGE VIEWS',
        'ACTIVITY_MONTH': 'ACTIVITY MONTH'
    })

    # Reorder columns
    desired_order = ['PAGE', 'ACTIVITY MONTH', 'PAGE VIEWS', 'PAGE SCROLL 25', 'PAGE SCROLL 50', 'PAGE SCROLL 75', 'PAGE SCROLL 100']
    df_final = df_final[desired_order]

    # Sort by LAST_TOUCH_CHANNEL then ACTIVITY_MONTH
    df_final = df_final.sort_values(by=['PAGE', 'ACTIVITY MONTH'])

    # Fill blanks as zeroes
    numeric_columns = ['PAGE VIEWS', 'PAGE SCROLL 25', 'PAGE SCROLL 50', 'PAGE SCROLL 75', 'PAGE SCROLL 100']
    df_final[numeric_columns] = df_final[numeric_columns].fillna(0)
    df_final['ACTIVITY MONTH'] = pd.to_datetime(df_final['ACTIVITY MONTH'])

    # Add BRAND and INDICATION columns
    df_final['BRAND'] = brand
    df_final['INDICATION'] = indication

    # Save final output to an Excel file
    output_path = f"{brand}_{indication}_{date_range}_SCROLL.xlsx"
    with pd.ExcelWriter(output_path, engine='openpyxl', mode='w') as writer:
        df_final.to_excel(writer, sheet_name='BMS Adobe Scroll Website', index=False)
    print(f"Transformation complete. Output saved to: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py <input_excel_file>")
        sys.exit(1)

    input_excel = sys.argv[1]
    
    temp_files = []

    # Step 1: Find the date range, brand and indication
    try:
        brand, indication = find_brand_and_indication(input_excel)
        date_range = find_report_date_range(input_excel)
        
        # Check if date_range is in the correct format
        if not re.match(r'^\d{8}-\d{8}$', date_range):
            print("Script unable to run due to only 1 day present in the report or incorrect date format.")
            sys.exit(1)
        
        start_date, end_date = date_range.split('-')
        
        if not check_full_period(start_date, end_date):
            print("The date range does not include full months or a full year.")
            sys.exit(1)
        
        print("Report Date Range is acceptable")
        print(f"Report Date Range: {date_range}, Brand: {brand}, Indication: {indication}")
    except ValueError as e:
        print(e)
        print("Report Date Range needs to be full months.")
        sys.exit(1)

    # Step 2: Extract and transform the SCROLL table
    try:
        extracted_scroll_table_path = extract_scroll_from_excel(input_excel)
        temp_files.append(extracted_scroll_table_path)  # Keep track of temporary file
        transform_scroll_excel(extracted_scroll_table_path, brand, indication, date_range)
    except ValueError as e:
        print(f"Error processing SCROLL table: {e}")

    # Step 3: Clean up temporary files
    for temp_file in temp_files:
        try:
            os.remove(temp_file)
            print(f"Deleted temporary file: {temp_file}")
        except OSError as e:
            print(f"Error deleting {temp_file}: {e}")