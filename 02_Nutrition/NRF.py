import pandas as pd

###### Nutrient density######

# 1. Read the data
df = pd.read_excel('Asia_recipe_nutrition_+Sugar.xlsx')
rda_df = pd.read_excel('RDA_input.xlsx')

# 2. Define the nutrient columns
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

# 3. Avoid division by zero
df = df[df['ENERC'] > 0].copy()

# 4. Calculate nutrient density (per 100 kcal)
for nutrient, new_col in nutrients.items():
    df[new_col] = df[nutrient] / df['ENERC'] * 100

# 5. View the results
print(df.head())

# 6. Save the results
df.to_excel('Nutrition_density_results.xlsx', index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### I. Nutrient index Children 1–3 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 17.10,   # protein (g)
    'FIBTG': 5,    # fiber (g)
    'FE': 5,       # iron (mg)
    'CA': 390,      # calcium (mg)
    'K': 2000,      # potassium (mg)
    'RETOL': 205,    # vitamin A (ug)
    'VITC': 15,    # vitamin C (mg)
    'ZN': 3,       # zinc (mg)
    'RIBF': 0.5    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 800,     # sodium (mg)
    'FAT': 25.41,       # fat (g)
    'SUGAR':13.70     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#1.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### II. Nutrient index Children 4–6 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 21.51,   # protein (g)
    'FIBTG': 9,    # fiber (g)
    'FE': 5,       # iron (mg)
    'CA': 680,      # calcium (mg)
    'K': 2300,      # potassium (mg)
    'RETOL': 245,    # vitamin A (ug)
    'VITC': 25,    # vitamin C (mg)
    'ZN': 5,       # zinc (mg)
    'RIBF': 0.6    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1000,     # sodium (mg)
    'FAT': 32.23,       # fat (g)
    'SUGAR':17.52     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#2.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### III. Nutrient index Children 7–10 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 32.96,   # protein (g)
    'FIBTG': 13,    # fiber (g)
    'FE': 8,       # iron (mg)
    'CA': 680,      # calcium (mg)
    'K': 2400,      # potassium (mg)
    'RETOL': 320,    # vitamin A (ug)
    'VITC': 40,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 0.8    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1100,     # sodium (mg)
    'FAT': 39.3,       # fat (g)
    'SUGAR':21.38     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#3.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### IV. Nutrient index Males 11–14 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 46.75,   # protein (g)
    'FIBTG': 17.5,    # fiber (g)
    'FE': 12.8,       # iron (mg)
    'CA': 960,      # calcium (mg)
    'K': 2500,      # potassium (mg)
    'RETOL': 480,    # vitamin A (ug)
    'VITC': 60,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 1.1    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1200,     # sodium (mg)
    'FAT': 26.11,       # fat (g)
    'SUGAR':30.51     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#4.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### V. Nutrient index Males 15–17 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 62.25,   # protein (g)
    'FIBTG': 22,    # fiber (g)
    'FE': 12.8,       # iron (mg)
    'CA': 960,      # calcium (mg)
    'K': 3000,      # potassium (mg)
    'RETOL': 580,    # vitamin A (ug)
    'VITC': 85,    # vitamin C (mg)
    'ZN': 11,       # zinc (mg)
    'RIBF': 1.4    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 68.39,       # fat (g)
    'SUGAR':37.31     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#5.xlsx', columns=cols, index=False)


#——————————————————————————————————————————————————————————————————————————————————
###### VI. Nutrient index Males 18–24 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 58.345,   # protein (g)
    'FIBTG': 22.5,    # fiber (g)
    'FE': 9.6,       # iron (mg)
    'CA': 860,      # calcium (mg)
    'K': 3400,      # potassium (mg)
    'RETOL': 570,    # vitamin A (ug)
    'VITC': 90,    # vitamin C (mg)
    'ZN': 11,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 60.7266666666667,       # fat (g)
    'SUGAR':32.999375     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#6.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### VII. Nutrient index Males 25–50 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 56.95,   # protein (g)
    'FIBTG': 23.5,    # fiber (g)
    'FE': 9.6,       # iron (mg)
    'CA': 750,      # calcium (mg)
    'K': 3400,      # potassium (mg)
    'RETOL': 570,    # vitamin A (ug)
    'VITC': 90,    # vitamin C (mg)
    'ZN': 11,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 60.63,       # fat (g)
    'SUGAR':32.99     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#7.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### VIII. Nutrient index Males 51–70 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 57.56,   # protein (g)
    'FIBTG': 23,    # fiber (g)
    'FE': 9.6,       # iron (mg)
    'CA': 750,      # calcium (mg)
    'K': 3400,      # potassium (mg)
    'RETOL': 570,    # vitamin A (ug)
    'VITC': 90,    # vitamin C (mg)
    'ZN': 11,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 56.73,       # fat (g)
    'SUGAR':31.12     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#8.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### IX. Nutrient index Males >70 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 58.03,   # protein (g)
    'FIBTG': 22.5,    # fiber (g)
    'FE': 9.6,       # iron (mg)
    'CA': 750,      # calcium (mg)
    'K': 3400,      # potassium (mg)
    'RETOL': 570,    # vitamin A (ug)
    'VITC': 90,    # vitamin C (mg)
    'ZN': 11,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 52.54,       # fat (g)
    'SUGAR':28.86     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#9.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### X. Nutrient index Females 11-14 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 45.53,   # protein (g)
    'FIBTG': 17,    # fiber (g)
    'FE': 11.2,       # iron (mg)
    'CA': 960,      # calcium (mg)
    'K': 2300,      # potassium (mg)
    'RETOL': 480,    # vitamin A (ug)
    'VITC': 60,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 1.1    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1200,     # sodium (mg)
    'FAT': 50.49,       # fat (g)
    'SUGAR':27.57     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#10.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### XI. Nutrient index Females 15-17 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 52.77,   # protein (g)
    'FIBTG': 21.5,    # fiber (g)
    'FE': 11.2,       # iron (mg)
    'CA': 960,      # calcium (mg)
    'K': 2300,      # potassium (mg)
    'RETOL': 490,    # vitamin A (ug)
    'VITC': 75,    # vitamin C (mg)
    'ZN': 9,       # zinc (mg)
    'RIBF': 1.4    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 58.01,       # fat (g)
    'SUGAR':31.73     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#11.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### XII. Nutrient index Females 18-24 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 49.06,   # protein (g)
    'FIBTG': 21.5,    # fiber (g)
    'FE': 11.2,       # iron (mg)
    'CA': 860,      # calcium (mg)
    'K': 2600,      # potassium (mg)
    'RETOL': 490,    # vitamin A (ug)
    'VITC': 80,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 50.65,       # fat (g)
    'SUGAR':27.57     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#12.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### XIII. Nutrient index Females 25–50 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 48.36,   # protein (g)
    'FIBTG': 21.5,    # fiber (g)
    'FE': 11.2,       # iron (mg)
    'CA': 750,      # calcium (mg)
    'K': 2600,      # potassium (mg)
    'RETOL': 490,    # vitamin A (ug)
    'VITC': 80,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 50.94,       # fat (g)
    'SUGAR':27.71     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#13.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### XIV. Nutrient index Females 25–50 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 48.36,   # protein (g)
    'FIBTG': 21.5,    # fiber (g)
    'FE': 9.6,       # iron (mg)
    'CA': 750,      # calcium (mg)
    'K': 2600,      # potassium (mg)
    'RETOL': 490,    # vitamin A (ug)
    'VITC': 80,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 47.64,       # fat (g)
    'SUGAR':26.12     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#14.xlsx', columns=cols, index=False)

#——————————————————————————————————————————————————————————————————————————————————
###### XV. Nutrient index Females 25–50 y ######

# 1. Define the RDA (adjust according to the population group)
RDA = {
    'PROCNT': 49.30,   # protein (g)
    'FIBTG': 21,    # fiber (g)
    'FE': 9.6,       # iron (mg)
    'CA': 750,      # calcium (mg)
    'K': 2600,      # potassium (mg)
    'RETOL': 490,    # vitamin A (ug)
    'VITC': 80,    # vitamin C (mg)
    'ZN': 8,       # zinc (mg)
    'RIBF': 1.3    # B2 (mg)
}

# Limiting items (recommended values)
LIMIT = {
    'NA': 1500,     # sodium (mg)
    'FAT': 45.57,       # fat (g)
    'SUGAR':25.02     # sugar (g)
}

#2. Calculate scores for positive nutrients
for nut, rda in RDA.items():
    df[nut + '_score'] = (df[nut] / rda).clip(upper=1)

#3. Calculate limiting items
for nut, limit in LIMIT.items():
    df[nut + '_score'] = (df[nut] / limit).clip(upper=1)

#4. Summarize
positive_cols = [k + '_score' for k in RDA.keys()]
negative_cols = [k + '_score' for k in LIMIT.keys()]

df['NRF_score'] = df[positive_cols].sum(axis=1) - df[negative_cols].sum(axis=1)

#5. Output the results
cols = list(df.columns[0:6]) + list(df.columns[54:67])
df.to_excel('Nutrition_index_results#15.xlsx', columns=cols, index=False)



