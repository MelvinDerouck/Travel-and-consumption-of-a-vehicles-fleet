# %%
import pandas as pd
import numpy as np

# %%
affectation_trackers = pd.read_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\Affectation trackers.csv",
                                      sep=";")


final_file = pd.read_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\final_file.csv",
                            sep=",")


suivi_conso_v2 = pd.read_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\Suivi Consommation v2.csv",
                                sep=",")

# %%
suivi_conso_v2

# %%
df_total = pd.merge(affectation_trackers, suivi_conso_v2, on='nom', how='outer')

df_total

# %%
df_total.drop('Carte GR_x', axis=1, inplace=True)
df_total.drop('Immat_y', axis=1, inplace=True)
df_total.drop('Division', axis=1, inplace=True)

df_total.rename(columns={'Carte GR_y': 'Carte GR'}, inplace=True)
df_total.rename(columns={'Immat_x': 'Immat'}, inplace=True)

# %%
df_total.loc[2, 'Litrage-Juil.'] = 142.16
df_total.loc[3, 'Litrage-Juil.'] = 278.23
df_total.loc[4, 'Litrage-Juil.'] = 218.67

rows_to_update = [2, 3, 4]
columns_to_sum = ['Litrage-Juil.', 'Litrage-Aout', 'Litrage-Sept.']

for row in rows_to_update:
    df_total.at[row, "Litrage-total"] = df_total.loc[row, columns_to_sum].sum()


# %%
pd.set_option('display.max_columns', None) 

# %%
df_invoxia = pd.merge(affectation_trackers, final_file, on='nom', how='outer')


# %%
df_main = pd.merge(df_total, final_file, on='nom', how='outer')

# %%
columns_del = ["Km-Juil.", "Km-Aout", "Km-Sept.", "Dernier Km-Juil.", "Dernier Km-Aout", "Dernier Km-Sept.", "Km-total", "Moyenne Km"]

df_main.drop(columns_del, axis=1, inplace=True)

# %%
df_main

# %%
df_main.to_excel(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\XLSX\df_main.xlsx", index=False)

df_main.to_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\df_main.csv", index=False)

# %%
print('-------- SUCCESSFUL --------')


