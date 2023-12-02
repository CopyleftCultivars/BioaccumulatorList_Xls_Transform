import pandas as pd

xls_file = r'C:\Users\caleb\Downloads\Nutrient Bio Accumulators.xlsx'
xls = pd.read_excel(xls_file, sheet_name=None)

dfs = []  # list to store modified DataFrames

# Iterate over each sheet
for sheet_name, df in xls.items():
    print(f'Sheet name: {sheet_name}')
    print(f'Columns: {df.columns.tolist()}')

    # Rename 'Min' and 'Max' columns
    if 'Min' in df.columns and 'Max' in df.columns:
        df = df.rename(columns={'Min': f'Min. {sheet_name}', 'Max': f'Max. {sheet_name}'})
        dfs.append(df)

# Check if dfs is empty
if not dfs:
    print('No DataFrames to concat.')
else:
    # Concat. all DataFrames
    combined_df = pd.concat(dfs)

    # Remove rows containing the specific string
    string_to_remove = "Questions, Comments, Additions, New Sources of Data? Email us:  OpenNutrientProject@gmail.com          Compiled by Originally by Stephen Raisner : PotentPonics@gmail.com. "
    mask = combined_df.applymap(lambda cell: string_to_remove in str(cell)).any(axis=1)
    combined_df = combined_df[~mask]

# Merge rows with duplicated text in the 'Scientific Name' and 'Latin Name' columns
combined_df = combined_df.groupby(['Scientific Name', 'Latin Name']).first().reset_index()

# Save the combined DataFrame to a new Excel file
output_file = 'C:/Users/caleb/Downloads/Nutrient_Bioaccumulator_combined_data.xlsx'
combined_df.to_excel(output_file, index=False)
