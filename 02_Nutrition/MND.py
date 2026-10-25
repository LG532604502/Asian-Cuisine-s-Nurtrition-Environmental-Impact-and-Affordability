import pandas as pd
import numpy as np

# 1. Load data
rda_df = pd.read_excel('RDA_input.xlsx')
recipe_df = pd.read_excel('Asia_recipe_nutrition.xlsx', sheet_name='Nutrition')

# ---------------------------------------------------------
# 2. Specify the target population group (the range is “0 to 14”; here, “Children 1-3 y” is used as an example, corresponding to index 0)
# ---------------------------------------------------------
target_group_idx = 14  # This number can be changed to switch to another population group,

# Extract the RDA data for the target population group and remove the population group name column
rda_values = rda_df.iloc[target_group_idx].drop('Life Stage Group')

# Data cleaning
rda_values = pd.to_numeric(rda_values, errors='coerce')

# ---------------------------------------------------------
# 3. Identify the nutrient columns shared by the recipe table and the RDA table
# ---------------------------------------------------------
recipe_nutrients = recipe_df.columns
rda_nutrients = rda_values.index

# Select nutrients that are present in both tables and have non-missing RDA values for the target population group
common_nutrients = [n for n in rda_nutrients if n in recipe_nutrients and pd.notna(rda_values[n])]

# ---------------------------------------------------------
# 4. Define the function: calculate ND and MND
# ---------------------------------------------------------
def calculate_nd_mnd(row):
    nd_dict = {}
    
    # Calculate the ND for each nutrient
    for nutrient in common_nutrients:
        content = row[nutrient]      # Content of the nutrient in the dish
        rda = rda_values[nutrient]   # RDA value for the corresponding population group
        
        # Ensure that the data are valid and that the RDA is greater than 0
        if pd.notna(content) and pd.notna(rda) and rda > 0:
            nd = 1 - content / rda
            nd = min(nd, 1.0)      # Truncation: limit the maximum value to 1
        else:
            nd = np.nan            # Set to NaN when data are missing
            
        nd_dict[f'ND_{nutrient}'] = nd
    
    # 5. Calculate MND: take the mean of all valid ND values (non-NaN)
    valid_nds = [v for v in nd_dict.values() if pd.notna(v)]
    mnd = np.mean(valid_nds) if len(valid_nds) > 0 else np.nan
    
    nd_dict['MND'] = mnd
    return pd.Series(nd_dict)

# ---------------------------------------------------------
# 6. Apply the calculation function to all dishes and output the results
# ---------------------------------------------------------
# Apply the function row by row
results_df = recipe_df.apply(calculate_nd_mnd, axis=1)


final_df = pd.concat([recipe_df.iloc[:, :6], results_df], axis=1)

# Print the first 5 rows for inspection
print(final_df.head())

# Save the results to a new CSV file
final_df.to_csv('calculated_ND_MND_Results#14.csv', index=False, encoding='utf-8-sig')
