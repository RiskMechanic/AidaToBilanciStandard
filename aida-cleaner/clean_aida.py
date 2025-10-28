import os

## function to clean datas
def cleandocs(raw_path, output_folder):

    import pandas as pd
    import re
    import numpy as np
    import traceback

    try:
        #load raw data
        df = pd.read_excel(raw_path, engine='xlrd', header=None)

        company_name = str(df.iloc[0, 1]).strip()
        ## find p.iva location dynamically
        piva = None
        for row_idx in range(df.shape[0]):
            for col_idx in range(df.shape[1]):
                cell_value = str(df.iat[row_idx, col_idx]).strip()
                if cell_value == "Codice fiscale":
                    # Scan to the right for the next non-empty cell
                    for next_col in range(col_idx + 1, df.shape[1]):
                        next_value = str(df.iat[row_idx, next_col]).strip()
                        if next_value and next_value.lower() != "nan":
                            piva = next_value
                            break
                    break
            if piva:
                break
        
        print("Cleaning data for company:", company_name)
        print("P.IVA:", piva)
        safecn = re.sub(r'[\\/*?:"<>|]', "_", company_name)
        
        # find header dynamically
        # Find the row containing "Bilancio non consolidato" in the first column
        target_row = None
        for i in range(df.shape[0]):
            cell = str(df.iat[i, 0]).strip()
            if "Bilancio non consolidato" in cell:
                target_row = i
                break

        if target_row is not None:
            # Extract non-empty values to the right of that cell
            header_values = []
            for col in range(1, df.shape[1]):
                val = str(df.iat[target_row, col]).strip()
                if val and val.lower() != "nan":
                    header_values.append(val)            
        else:
            print("'Bilancio non consolidato' not found in first column.")



        # Filter data
        target_index = df[df[0].astype(str).str.contains(" A. CREDITI VERSO SOCI", case=False)].index[0]
        end_index = df[df[0].astype(str).str.contains("  UTILE/PERDITA DI ESERCIZIO di pert. del GRUPPO", case=False)].index[0]
        df = df.iloc[target_index:end_index] # Keep everything from that row onward

        # Remove columns that contain any cell with exactly "EUR" 
        df = df.loc[:, ~df.apply(lambda col: col.astype(str).str.strip().str.upper().eq("EUR").any())]
        df.reset_index(drop=True, inplace=True)

        #removes empty rows and columns
        df.dropna(how='all', inplace=True)  # Drop rows where all elements are NaN
        df.dropna(axis=1, how='all', inplace=True)  # Drop columns where all elements are NaN

        #replace 'n.d.' with 0
        df.replace(to_replace=r'^\s*n\.d\.\s*$', value=0, regex=True, inplace=True)

        #remove empty spaces before text in fisrt column
        df[0] = df[0].astype(str).str.lstrip()

        #Remove specific empty dataframe
        df.reset_index(drop=True, inplace=True)
        start_idx = df[df[0].astype(str).str.contains("Garanzie prestate", case=False)].index[0]
        end_idx = df[df[0].astype(str).str.contains("A. TOT. VAL. DELLA PRODUZIONE", case=False)].index[0]
        df.drop(index=range(start_idx, end_idx), inplace=True)
        df.reset_index(drop=True, inplace=True)

        #add text "Conto economico" before Risultato di esecizio
        target_idx = df[df[0].astype(str).str.strip() == "A. TOT. VAL. DELLA PRODUZIONE"].index[0]
        new_row = ["Conto economico"] + [""] * (df.shape[1] - 1)
        df.loc[target_idx - 0.5] = new_row  # Insert with fractional index to avoid collision
        df = df.sort_index().reset_index(drop=True)

        df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        df.dropna(axis=1, how='all', inplace=True)

        # Insert header row using previously extracted values
        new_head = ["Anno"] + header_values
        header_row = [""] * df.shape[1]
        for i, val in enumerate(new_head):
            if i < df.shape[1]:
                header_row[i] = val

        df.loc[-1] = header_row
        df.index = df.index + 1
        df.sort_index(inplace=True)
        df.iloc[0] = df.iloc[0].apply(lambda x: int(x) if str(x).isdigit() else x)


        #append company name and piva at the end of the dataframe
        new_rows = pd.DataFrame(np.nan, index=[0, 1], columns=df.columns)
        new_rows[0] = new_rows[0].astype(object)
        new_rows.iloc[0, 0] = f'{safecn}'
        new_rows.iloc[1, 0] = f'{piva}'
        df = pd.concat([df, new_rows], ignore_index=True)


        #save cleaned data
        clean_folder='aida_clean'
        os.makedirs(clean_folder, exist_ok=True)
        clean_path = os.path.join(clean_folder,f'{safecn}.xlsx')
        df.to_excel(clean_path, index=False,header=False)
        print("Data cleaned and saved to", clean_path)
    except Exception as e:
        print(f"Error processing {raw_path}: {e}")
        traceback.print_exc()

# run cleandocs function for every files .xls in input folder and save in outputfolder
input_folder = 'aida_raw'
output_folder = 'aida_clean'
os.makedirs(output_folder, exist_ok=True)

for filename in os.listdir(input_folder):
    if filename.lower().endswith('.xls'):
        raw_path = os.path.join(input_folder, filename)
        print(f"Processing file: {filename}")
        cleandocs(raw_path, output_folder)


 ## source aida-env/bin/activate
