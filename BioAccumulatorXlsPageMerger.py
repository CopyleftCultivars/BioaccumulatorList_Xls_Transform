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

    print(f'Sheet name: {sheet_name}')
    print(f'Columns: {df.columns.tolist()}')

    # Rename 'Min' and 'Max' columns
    if 'Min' in df.columns and 'Max' in df.columns:
        df = df.rename(columns={'Min': f'Min. {sheet_name}', 'Max': f'Max. {sheet_name}'})
    
    df.columns = list(rename_duplicates(df.columns))  # Rename duplicate columns
    dfs.append(df)

# Concatenate all DataFrames
combined_df = pd.concat(dfs)

# Define a function to apply for each column
def agg_func(x):
    if x.dtypes == 'float64' and x.notnull().all():
        return x.mean()
    else:
        return x.first()

# Combine rows with the same scientific or latin name
combined_df = combined_df.groupby(['Scientific Name', 'Latin Name']).apply(agg_func)

# Remove rows with no data
combined_df = combined_df.dropna(how='all')

# Remove rows containing the specific string
string_to_remove = "Questions, Comments, Additions, New Sources of Data? Email us:  OpenNutrientProject@gmail.com          Compiled by Originally by Stephen Raisner : PotentPonics@gmail.com. "
mask = combined_df.applymap(lambda cell: string_to_remove in str(cell)).any(axis=1)
combined_df = combined_df[~mask]

# Reset the index without adding it as a column
combined_df.reset_index(drop=True, inplace=True)

# Merge rows with duplicated text in the 'Scientific Name' and 'Latin Name' columns
combined_df = combined_df.groupby(['Scientific Name', 'Latin Name'], as_index=False).first()

# Save the combined DataFrame to a new Excel file
output_file = 'C:/Users/caleb/Downloads/Nutrient_Bioaccumulator_combined_data.xlsx'
combined_df.to_excel(output_file, index=False)
