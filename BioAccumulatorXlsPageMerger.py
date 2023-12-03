import pandas as pd

xls_file = r'C:\Users\caleb\Downloads\Copy of Nutrient Bio Accumulators.xlsx'
xls = pd.read_excel(xls_file, sheet_name=None, header=None)

dfs = []  # list to store modified DataFrames

def rename_duplicates(old):
    seen = {}
    for x in old:
        if x in seen:
            seen[x] += 1
            yield "%s.%d" % (x, seen[x])
        else:
            seen[x] = 0
            yield x

# Iterate over each sheet
for sheet_name, df in xls.items():
    # Remove the first two rows
    df = df.iloc[2:]
    df.columns = df.iloc[0]
    df = df[1:]

    print(f'Sheet name: {sheet_name}')
    print(f'Columns: {df.columns.tolist()}')

    # Rename 'Min' and 'Max' columns
    if 'Min' in df.columns and 'Max' in df.columns:
        df = df.rename(columns={'Min': f'Min. {sheet_name}', 'Max': f'Max. {sheet_name}'})
    
    df.columns = list(rename_duplicates(df.columns))  # Rename duplicate columns

    dfs.append(df)

# Concatenate all DataFrames vertically
combined_df = pd.concat(dfs, ignore_index=True)

# If combined_df is still empty after concatenation, print an error message
if combined_df.empty:
    print("Error: The combined DataFrame is empty. Please check the data in the Excel file.")
else:
    # Save the combined DataFrame to a new Excel file
    output_file = 'C:/Users/caleb/Downloads/Nutrient_Bioaccumulator_combined_data.xlsx'
    combined_df.to_excel(output_file, index=False)