import streamlit as st
import pandas as pd
import re
import io
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image


st.title("USP Co-cultivation processor")

project = st.selectbox("Choose a project:", ["DDF919", "HKR357"])
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsm"])



# Select project
if project == "DDF919":

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
        new_df["ddPCR 'cell ratio'"] = df["ddPCR 'cell ratio'"]
        new_df["ddPCR 'cell ratio project relationship'"] = df["ddPCR 'cell ratio project relationship'"]
        if "FACS 'cell ratio'" in df.columns and df["FACS 'cell ratio'"].notna().any():
            new_df["FACS 'cell ratio'"] = df["FACS 'cell ratio'"]
        if "Cell Bank Type" in df.columns and df["Cell Bank Type"].notna().any():
            new_df["Cell Bank Type"] = df["Cell Bank Type"]

        # Remove rows where Stage is "Unknown"
        new_df = new_df[new_df["Stage"] != "Unknown"]

        # Create "Sample" column to differentiate between -01 and -02
        new_df["Sample"] = new_df["Sample ID"].apply(
            lambda x: 1 if x.endswith("-01") else 2 if x.endswith("-02") else 0
        )  

        # Create "Day" column
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

        # Assign the 'Day' values
        new_df['Day'] = new_df['Sample Day'].map(date_to_dataset)
        new_df.drop(columns=["Sample Day"], inplace=True)
        new_df['Ratio FVF959/ORQ933'] = new_df.apply(assign_ratio, axis=1)


        # Reorder columns
        cols = list(new_df.columns)
        cols.insert(2,cols.pop(cols.index("Sample")))
        cols.insert(5, cols.pop(cols.index("Ratio FVF959/ORQ933")))
        new_df = new_df[cols]

        # Define the condition for rows to be removed: 25302A was used for MS (002 was used for 25303A)
        new_df = new_df[~(
            new_df['Sample ID'].str.startswith(('ORQ933-25302A', 'FVF959-25302A')) &
            ~new_df['Sample ID'].str.contains(r'-000-|-002-'))]

        # Assign culture flows 
        def matches(sample_id, pattern):
            return bool(re.search(pattern, sample_id))

        # Define flow 
        flow_rules = {
            "ORQ933 SF 01": lambda r: r['Culture'] == 'ORQ933' and r['Vessel'] == 'SF' and (
                matches(r['Sample ID'], r'N\d+S-\d+-01') or matches(r['Sample ID'], r'S\d+S-\d+-01')),

            "ORQ933 SF 02": lambda r: r['Culture'] == 'ORQ933' and r['Vessel'] == 'SF' and (
                matches(r['Sample ID'], r'N\d+S-\d+-01') or matches(r['Sample ID'], r'S\d+S-\d+-02')),

            "FVF959 SF 01": lambda r: r['Culture'] == 'FVF959' and r['Vessel'] == 'SF' and (
                matches(r['Sample ID'], r'N\d+S-\d+-01') or matches(r['Sample ID'], r'S\d+S-\d+-01')),

            "FVF959 SF 02": lambda r: r['Culture'] == 'FVF959' and r['Vessel'] == 'SF' and (
                matches(r['Sample ID'], r'N\d+S-\d+-01') or matches(r['Sample ID'], r'S\d+S-\d+-02')),

            "ORQ933 BR": lambda r: r['Culture'] == 'ORQ933' and (
                r['Vessel'] == 'BR' or matches(r['Sample ID'], r'N\d+S-\d+-01')),

            "FVF959 BR": lambda r: r['Culture'] == 'FVF959' and (
                r['Vessel'] == 'BR' or matches(r['Sample ID'], r'N\d+S-\d+-01')),

            "DDF919 BR ratio 1:2": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:2' and not matches(r['Sample ID'], r'S\d+S-\d+-0\d')) or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 BR ratio 1:1": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:1' and not matches(r['Sample ID'], r'S\d+S-\d+-0\d')) or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 BR ratio 2:1": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '2:1' and not matches(r['Sample ID'], r'S\d+S-\d+-0\d')) or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 SF 01 ratio 1:2": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:2' and matches(r['Sample ID'], r'S\d+S-\d+-01')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:2' and r["Stage"] == "Expansion") or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 SF 02 ratio 1:2": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:2' and matches(r['Sample ID'], r'S\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:2' and r["Stage"] == "Expansion") or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 SF 01 ratio 1:1": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:1' and matches(r['Sample ID'], r'S\d+S-\d+-01')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:1' and r["Stage"] == "Expansion") or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 SF 02 ratio 1:1": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:1' and matches(r['Sample ID'], r'S\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '1:1' and r["Stage"] == "Expansion") or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),

            "DDF919 SF 01 ratio 2:1": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '2:1' and matches(r['Sample ID'], r'S\d+S-\d+-01')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '2:1' and r["Stage"] == "Expansion") or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),
                                                                                            
            "DDF919 SF 02 ratio 2:1": lambda r: (
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Stage'] == 'Thawing') or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'N\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '2:1' and matches(r['Sample ID'], r'S\d+S-\d+-02')) or
                (r['Culture'] == 'DDF919' and r['Ratio FVF959/ORQ933'] == '2:1' and r["Stage"] == "Expansion") or
                (r['Culture'] in ['ORQ933', 'FVF959'] and r['Vessel'] == 'SF' and matches(r['Sample ID'], r'25302A-N\d+S+-00\d-'))),
        }

        # Apply tagging logic
        def assign_cultureflow(row):
            return [name for name, rule in flow_rules.items() if rule(row)]

        new_df['Culture flow'] = new_df.apply(assign_cultureflow, axis=1)

        # Add image with title and description
        
        with st.sidebar.container():
            image = Image.open("Nomenclature_DDF919.png")
        
            # Styled caption
            st.markdown("""
            <div style='font-size:18px; color:black; font-weight:bold;'>
                Nomenclature DDF
            </div>
            """, unsafe_allow_html=True)
        
            # Justified description with line break
            st.markdown("""
            <div style='text-align: justify; font-size:16px; color:black; font-weight:500;'>
                This diagram presents the upstream workflow of the DDF co-cultivation project, detailing sequential stages from cell thawing to the N-1 phase.<br>
                It includes standardised nomenclature and hierarchical relationships.
            </div>
            """, unsafe_allow_html=True)
        
            # Image display
            st.image(image, use_container_width=True)

        # Display Canva link
        st.sidebar.markdown(
            "https://www.canva.com/design/DAGvG6ozxpY/ErpDaPKfF5OiHHQO7CQ1bA/edit",
            unsafe_allow_html=True)

        # Add Culture flow filter
        culture_flows = ['All'] + sorted(set(flow for flows in new_df['Culture flow'] for flow in flows))
        selected_flow = st.sidebar.selectbox("Select Culture flow:", culture_flows)

        # Apply filter if not 'All'
        if selected_flow != 'All':
            new_df = new_df[new_df['Culture flow'].apply(lambda flows: selected_flow in flows)]

        # Optional: Remove Sample IDs entered by the user
        sample_ids_input = st.sidebar.text_input("Enter Sample IDs to remove (comma-separated):")
        if sample_ids_input:
            sample_ids_to_remove = [s.strip() for s in sample_ids_input.split(",")]
            new_df = new_df[~new_df['Sample ID'].isin(sample_ids_to_remove)]


        # Functions to add Time [d] and Data set columns
        def parse_sample_id(sample_id):
            parts = sample_id.split("-")
            project, batch = parts[0], parts[1]
            sf = parts[2] if len(parts) > 2 else ""
            expansion = parts[3] if len(parts) > 3 else ""
            sample = int(parts[4]) if len(parts) > 4 else None
            j, k = (expansion[1], expansion[2]) if len(expansion) == 3 and expansion.isdigit() else (None, None)
            return project, batch, sf, expansion, j, k, sample

        def extract_batch_j(sample_id):
            parts = sample_id.split('-')
            return (parts[1], parts[3][1]) if len(parts) >= 4 and len(parts[3]) >= 2 else (None, None)

        def assign_data_set(df, start_value):
            df = df.sort_values(by='Sample Date').copy()
            df[['Batch', 'j']] = df['Sample ID'].apply(lambda x: pd.Series(extract_batch_j(x)))
            data_set, current_set = [], start_value
            prev_batch, prev_j = None, None
            for batch, j in zip(df['Batch'], df['j']):
                if batch != prev_batch or j != prev_j:
                    current_set += 1 if prev_batch is not None else 0
                data_set.append(current_set)
                prev_batch, prev_j = batch, j
            df['Data set'] = data_set
            return df[['Sample ID', 'Data set']]

        def calculate_time_differences(df):
            df["Time [d]"] = 0.0
            df = df.sort_values("Sample Date").reset_index(drop=True)
            ref_n, ref_qs = {}, {}
            for i, row in df.iterrows():
                sample_id, sample_date = row["Sample ID"], row["Sample Date"]
                project, batch, sf, expansion, j, k, sample = parse_sample_id(sample_id)
                group_key = (project, batch, sample)
                if "-000" in sample_id:
                    if sf.startswith("N"):
                        ref_n[(project, batch, sample, j)] = {0: sample_date}
                    elif sf.startswith(("Q", "S")):
                        ref_qs[group_key] = sample_date
                else:
                    try: k_int = int(k)
                    except: k_int = None
                    if sf.startswith("N") and k_int is not None:
                        j_key = (project, batch, sample, j)
                        ref_n.setdefault(j_key, {})
                        for prev_k in range(k_int - 1, -1, -1):
                            if prev_k in ref_n[j_key]:
                                df.at[i, "Time [d]"] = (sample_date - ref_n[j_key][prev_k]).total_seconds() / 86400
                                break
                        ref_n[j_key][k_int] = sample_date
                    elif sf.startswith(("Q", "S")) and group_key in ref_qs:
                        df.at[i, "Time [d]"] = (sample_date - ref_qs[group_key]).total_seconds() / 86400
            return df

        def filter_and_assign(df):
            filtered = []
            for culture, sample, stages, start_val in [('ORQ933', 1, ['Thawing', 'Expansion'], 1),
                                                    ('FVF959', 1, ['Thawing', 'Expansion'], 1),
                                                    ('ORQ933', 2, ['Expansion'], 2),
                                                    ('FVF959', 2, ['Expansion'], 2)]:
                subset = df[(df['Stage'].isin(stages)) & (df['Culture'] == culture) & (df['Sample'] == sample)].copy()
                if not subset.empty:
                    filtered.append(assign_data_set(subset, start_val))
            return pd.concat(filtered)

        # Add Time [d] and Data set
        new_df = calculate_time_differences(new_df)
        combined_filtered = filter_and_assign(new_df)
        new_df = new_df.merge(combined_filtered, on='Sample ID', how='left')

        valid = (((new_df['Stage'].isin(['Thawing', 'Expansion'])) & (new_df['Culture'].isin(['ORQ933', 'FVF959'])) & (new_df['Sample'] == 1)) |
                ((new_df['Stage'] == 'Expansion') & (new_df['Culture'].isin(['ORQ933', 'FVF959'])) & (new_df['Sample'] == 2)))
        new_df['Data set'] = new_df['Data set'].where(valid, None)

        max_orq933 = new_df.loc[(new_df['Culture'] == 'ORQ933') & (new_df['Data set'].notna()), 'Data set'].max()
        max_fvf959 = new_df.loc[(new_df['Culture'] == 'FVF959') & (new_df['Data set'].notna()), 'Data set'].max()
        if max_orq933 == max_fvf959:
            new_value = max_orq933 + 1
            new_df.loc[(new_df['Culture'] == 'DDF919') & (new_df['Stage'] == 'Expansion'), 'Data set'] = new_value
        else:
            print("Different expansion numbers between cell lines")

        for stage, offset in [('N-2', 1), ('N-1', 1)]:
            for culture in ['ORQ933', 'FVF959', 'DDF919']:
                # Determine the reference stage for max lookup
                reference_stage = 'Expansion' if stage == 'N-2' else 'N-2'
                max_val = new_df.loc[
                    (new_df['Culture'] == culture) &
                    (new_df['Stage'] == reference_stage) &
                    (new_df['Data set'].notna()),
                    'Data set'
                ].max()
                if pd.notna(max_val):
                    new_df.loc[
                        (new_df['Culture'] == culture) &
                        (new_df['Stage'] == stage),
                        'Data set'
                    ] = max_val + offset

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


elif project == "HKR357":
    st.write("Running code for project HKR357...")
    # Add code for HKR357 here
    st.success("Code for HKR357 executed successfully.")




