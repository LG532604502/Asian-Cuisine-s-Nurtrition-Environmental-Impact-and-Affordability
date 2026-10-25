import pandas as pd

#Read the data
df = pd.read_excel('Asia_recipe_nutrition_+Sugar.xlsx')

#Define the nutrient columns
nutrients = {
    'WATER': 'WATER_density',
    'PROCNT': 'PROCNT_density',
    'FAT': 'FAT_density',
    'CHOAVLDF': 'CHOAVLDF_density',
    'FIBTG': 'FIBTG_density',
    'ASH': 'ASH_density',
    'CA': 'CA_density',
    'P': 'P_density',
    'NA': 'NA_density',
    'K': 'K_density',
    'FE': 'FE_density',
    'CU': 'CU_density',
    'ZN': 'ZN_density',
    'RETOL': 'RETOL_density',
    'CARTB': 'CARTB_density',
    'VITA_RAE': 'VITA_RAE_density',
    'THIA': 'THIA_density',
    'RIBF': 'RIBF_density',
    'NIA': 'NIA_density',
    'VITC': 'VITC_density',
    'SUGAR': 'SUGAR_density',
    }

#Avoid division by zero
df = df[df['ENERC'] > 0].copy()

#Calculate nutrient density (per 100 kcal)
for nutrient, new_col in nutrients.items():
    df[new_col] = df[nutrient] / df['ENERC'] * 100


#Save the results
df.to_excel('Nutrition_density_results.xlsx', index=False)