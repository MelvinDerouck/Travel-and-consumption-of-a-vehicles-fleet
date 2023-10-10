# %%
import pandas as pd
import numpy as np

# %%
xls_file_path = r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\TOTAL CARTES GR\exports 2023\Suivi Consommation.xls"

# Chemin de destination pour le fichier XLSX
xlsx_file_path = r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\TOTAL CARTES GR\exports 2023\Suivi Consommation.xlsx"

try:
   # Lire le fichier XLS en utilisant pandas
   df = pd.read_excel(xls_file_path)

   # Enregistrez le DataFrame dans un fichier XLSX
   df.to_excel(xlsx_file_path, index=False)

   print(f'Conversion terminée. Le fichier XLS a été converti en {xlsx_file_path}')
except FileNotFoundError:
   print(f'Le fichier {xls_file_path} n\'existe pas.')

df = df.iloc[6:]

# %%
record_count = len(df) // 4

new_columns = [
    "N° de carte", "Immat./Nom", "Mention compl.", "Division",
    "Km-Juil.", "L/100km-Juil.", "Litrage-Juil.", "Dernier Km-Juil.",
    "Km-Aout", "L/100km-Aout", "Litrage-Aout", "Dernier Km-Aout",
    "Km-Sept.", "L/100km-Sept.", "Litrage-Sept.", "Dernier Km-Sept.",
    "Km-total", "L/100km-total", "Litrage-total", "Dernier Km-total"
]

new_df = pd.DataFrame(columns=new_columns)

for i in range(record_count):
    start_row = i * 4
    end_row = start_row + 4
    record = df.iloc[start_row:end_row, 0:5].T.values.flatten()
    
    if len(record) < 20:
        record = list(record) + [None] * (20 - len(record))
    
    new_df.loc[i] = record

new_df = new_df.drop(0, axis=0)

new_df.reset_index(drop=True, inplace=True)

new_df = new_df.drop('Dernier Km-total', axis=1)

# %%
liste_cartes = ["0056", "0126", "0129", "0163", "0171", "0205", "0159", "0186"]

df_filt = new_df[new_df["N° de carte"].isin(liste_cartes)]

# %%
reorg_columns = ["N° de carte", "Immat./Nom", "Mention compl.",	"Division", 
                "Litrage-Juil.", "Km-Juil.", "Dernier Km-Juil.", "L/100km-Juil.", 
                "Litrage-Aout", "Km-Aout", "Dernier Km-Aout", "L/100km-Aout", 
                "Litrage-Sept.", "Km-Sept.", "Dernier Km-Sept.", "L/100km-Sept.", 
                "Litrage-total", "Km-total", "L/100km-total"]

df_filt = df_filt[reorg_columns]

# %%
df_filt["Moyenne Km"] = df_filt[["Km-Juil.", "Km-Aout", "Km-Sept."]].mean(axis=1)

df_filt["Moyenne Litrage"] = df_filt[["Litrage-Juil.", "Litrage-Aout", "Litrage-Sept."]].mean(axis=1)

df_filt["Moyenne L/100km"] = df_filt[["L/100km-Juil.", "L/100km-Aout", "L/100km-Sept."]].mean(axis=1)

# %%
df_filt.drop('L/100km-total', axis=1, inplace=True)

# %%
colonnes_int = ["Km-Juil.", "Km-Aout", "Km-Sept.", "Km-total", 
                "Moyenne Km", "L/100km-Juil.", "L/100km-Aout", "L/100km-Sept.", 
                "Litrage-Juil.", "Litrage-Aout", "Litrage-Sept.", "Litrage-total",
                "Dernier Km-Juil.", "Dernier Km-Aout", "Dernier Km-Sept."
]

df_filt[colonnes_int] = df_filt[colonnes_int].astype(float)


# %%
df_filt.rename(columns={'N° de carte': 'Carte GR'}, inplace=True)
df_filt.rename(columns={'Immat./Nom': 'Immat'}, inplace=True)
df_filt.rename(columns={'Mention compl.': 'nom'}, inplace=True)

# %%
df_filt

# %%
correspondance = {
    "0056": 'BOURBOTTE',
    "0126": 'ALIMI',
    "0129": 'MARINELLO',
    "0159": 'MATHELIN',
    "0163": 'OCULI',
    "0171": 'COLLI',
    "0186": 'LAROCHE',
    "0205": 'MENESES G.'
}

def associer_nom(carte_gr):
    return correspondance.get(carte_gr, 'Inconnu')

df_filt['nom'] = df_filt['Carte GR'].apply(associer_nom)

# %%
df_filt

# %%
df_filt.to_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\Suivi Consommation v2.csv", index=False)

# %%
with pd.ExcelWriter(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\XLSX\Suivi Consommation v2.xlsx") as writer:
    df_filt.to_excel(writer, sheet_name='Conso', index=False)

print('-------- SUCCESSFUL --------')
