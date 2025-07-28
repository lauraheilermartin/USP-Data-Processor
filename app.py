import streamlit as st
import pandas as pd
import re
import io
from openpyxl.worksheet.table import Table, TableStyleInfo



st.title("USP Data Processor")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsm"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, "Data Summary Table")

    # Create dictionary
    stage_rules = {
        'Thawing': {
            'cultures': ['ORQ933', 'FVF959'],
            'must_contain': ['25302A', '000']
        },
        'Expansion': {
            'ORQ933': ['25302A', '25303A'],
            'FVF959': ['25302A', '25303A'],
            'DDF919': ['25310A', '25310B', '25310C']
        },
        'N-2': {
            'ORQ933': ['25307A', '25308A'],
            'FVF959': ['25306A', '25307'],
            'DDF919': ['25307A', '25308A', '25309A', '25311A', '25311B', '25311C']
        },
        'N-1': {
            'ORQ933': ['25309A', '25310A'],
            'FVF959': ['25308A', '25309A'],
            'DDF919': ['25313A', '25314A', '25315A', '25316A', '25316B', '25316C']
        }
    }

    culture_ratios = {
        'ORQ933': '0:1',
        'FVF959': '1:0'
    }

    ddf919_rules = {
        'SF': {
            '1:1': ['-25310A', '-25311A', '-25316A'],
            '1:2': ['-25310B', '-25311B', '-25316B'],
            '2:1': ['-25310C', '-25311C', '-25316C']
        },
        'BR': {
            '1:1': ['25307A', '25313A'],
            '1:2': ['25308A', '25314A'],
            '2:1': ['25309A', '25315A']
        }
    }


    def classify_stage_dict(row):
        sample_id = row['Sample ID']
        culture = row['Culture']

        # Thawing
        if culture in stage_rules['Thawing']['cultures'] and \
        all(code in sample_id for code in stage_rules['Thawing']['must_contain']):
            return 'Thawing'

        # N-2 and N-1
        for stage in ['N-2', 'N-1']:
            if culture in stage_rules[stage]:
                if any(code in sample_id for code in stage_rules[stage][culture]):
                    return stage

        # Expansion
        if culture in stage_rules['Expansion']:
            for code in stage_rules['Expansion'][culture]:
                if code in sample_id:
                    # Special case: ORQ933 with 25302A but not 000
                    if culture == 'ORQ933' and code == '25302A' and '000' in sample_id:
                        continue  # skip this one, it's Thawing
                    return 'Expansion'

        return 'Expansion' if culture in ['ORQ933', 'FVF959'] else 'Unknown'


    def classify_vessel(sample_id):
        sample_str = str(sample_id)
        return 'BR' if re.search(r'-Q\d+', sample_str) else 'SF'

    def assign_ratio(row):
        culture = row['Culture']
        vessel = row['Vessel']
        sample_id = row['Sample ID']

        if culture in culture_ratios:
            return culture_ratios[culture]
        elif culture == 'DDF919':
            for ratio, codes in ddf919_rules.get(vessel, {}).items():
                if any(code in sample_id for code in codes):
                    return ratio
        return 'Unknown'


    # Assign columns to new data frame
    new_df = pd.DataFrame()
    new_df['Sample ID'] = df['Sample ID']
    new_df['Culture'] = new_df['Sample ID'].apply(lambda x: str(x).split('-')[0])

    new_df['Stage'] = new_df.apply(classify_stage_dict, axis=1)
    new_df['Vessel'] = new_df['Sample ID'].apply(classify_vessel)
    new_df["VCD [cells/ml]"] = df["Viable Cell Density [cells/ml]"] 
    new_df["Viability [%]"] = df["Viability [%]"]
    new_df["cGlc FLEX [g/L]"] = df["cGlc FLEX [g/L]"]
    new_df["cGln FLEX [g/L]"] = df["cGln FLEX [g/L]"]
    new_df["cLac FLEX [g/L]"] = df["cLac FLEX [g/L]"]
    new_df["cNH4+ FLEX [mmol/L]"] = df["cNH4+ FLEX [mmol/L]"]
    new_df["Volume"] = 1
    new_df["Sample Date"] = df["Sample Date"]

    # Remove rows where Stage is "Unknown"
    new_df = new_df[new_df["Stage"] != "Unknown"]

    # Create "Sample" column to differentiate between -01 and -02
    new_df["Sample"] = new_df["Sample ID"].apply(
        lambda x: 1 if x.endswith("-01") else 2 if x.endswith("-02") else 0
    )    

    # Create "Data set" column
    # Convert 'Sample Date' to datetime
    new_df['Sample Date'] = pd.to_datetime(new_df['Sample Date'])
    new_df['Sample Day'] = new_df['Sample Date'].dt.date
    unique_days = sorted(new_df['Sample Day'].unique())

    # Calculate day offsets from the first date
    day_offsets = [0]
    for i in range(1, len(unique_days)):
        offset = (unique_days[i] - unique_days[i-1]).days
        day_offsets.append(day_offsets[-1] + offset)

    # Map each date to its corresponding dataset value
    date_to_dataset = dict(zip(unique_days, day_offsets))

    # Assign the 'Data set' values
    new_df['Data set'] = new_df['Sample Day'].map(date_to_dataset)
    new_df.drop(columns=["Sample Day"], inplace=True)
    new_df['Ratio FVF959/ORQ933'] = new_df.apply(assign_ratio, axis=1)


    # Reorder columns
    cols = list(new_df.columns)
    cols.insert(2,cols.pop(cols.index("Sample")))
    cols.insert(5, cols.pop(cols.index("Ratio FVF959/ORQ933")))
    new_df = new_df[cols]


    st.success("File processed successfully!")
    st.dataframe(new_df) 

    # Option to download the result

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        new_df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        
        # Define the table range
        table_range = f"A1:{chr(65 + len(new_df.columns) - 1)}{len(new_df) + 1}"
        table = Table(displayName="DataTable", ref=table_range)

        # Style the table
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                            showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        table.tableStyleInfo = style
        worksheet.add_table(table)


    buffer.seek(0)

    st.download_button(
        label="Download Processed File",
        data=buffer,
        file_name="Processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )







