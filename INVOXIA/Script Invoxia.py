
"""
Created on Mon Oct 18 14:31:07 2021

Author : melvin.derouck (03/08/2023)

"""

from geopy import distance
import numpy as np
import pandas as pd
import xlsxwriter
import seaborn as sns
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image
import os

csv_file = r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\export Invoxia\export_2023-07-03_2023-09-30.csv"

df_nom= pd.read_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\Affectation trackers.csv", sep=";")

df = pd.read_csv(csv_file)

df_sliced_dict = {}

df=df.sort_values(by=['name','datetime'])

df_merged = pd.merge(df, df_nom[['name', 'nom']], left_on='name', right_on='name', how='left')

df_merged['name']

valeurs_a_conserver = ["b'T 1'", "b'T 2'", "b'T 3'", "b'T 4'", "b'T 5'", "b'T 6'", "b'T11'", "b'T12'"]

masque = df_merged['name'].isin(valeurs_a_conserver)

df_filtre = df_merged[masque]

for serial in df['name'].unique():
    
    df_sliced_dict[serial] = df[df['name'] == serial ]
  
    df = df.fillna(method='ffill')
    df = df.fillna(method='bfill')
    
    df['lon-start'] = df['lng']
    df['lon-start'].loc[-1] = np.nan
    df['lon-start'] = np.roll(df['lon-start'], 1)
    df['lat-start'] = df['lat']
    df['lat-start'].loc[-1] = np.nan
    df['lat-start'] = np.roll(df['lat-start'], 1)
    df['time-start'] = df['datetime']
    df['time-start'].loc[-1] = np.nan
    df['time-start'] = np.roll(df['time-start'], 1)
    df = df.fillna(method='bfill')
    
    df['datetime'] = pd.to_datetime(df['datetime'], format='mixed')
    df['datetime'] = df['datetime'].dt.tz_localize(tz=None)
    df['time-start'] = pd.to_datetime(df['time-start'], format='mixed')
    df['time-start'] = df['time-start'].dt.tz_localize(tz=None)
    
    df['distance_dis_2d'] = df.apply(lambda x: distance.distance((x['lat-start'], x['lon-start']), (x['lat'], x['lng'])).km, axis = 1)
    df['time_delta_minutes'] = df.apply(lambda x: (x['datetime'] - x['time-start']).total_seconds()/60, axis=1) 

df['datetime']=(df['datetime'].dt.to_period('D'))

df= df[(df['time_delta_minutes']>0)]  
df=df[(df['stationary']=='move')]
df.insert(2, 'nom', df['name'].map(df_nom.set_index('name')['nom']))

#######################

jour_astreinte = {
    "2023-07-08", "2023-07-09", "2023-07-14",
    "2023-07-15", "2023-07-16", "2023-07-22",
    "2023-07-23", "2023-07-29", "2023-07-30",
    "2023-08-05", "2023-08-06", "2023-08-12",
    "2023-08-13", "2023-08-15", "2023-08-19", 
    "2023-08-20", "2023-08-26", "2023-08-27", 
    "2023-09-02", "2023-09-03", "2023-09-09",
    "2023-09-10", "2023-09-16", "2023-09-17", 
    "2023-09-23", "2023-09-24", "2023-09-30"
    }

df['jours_astreinte'] = df['datetime'].dt.strftime("%Y-%m-%d").isin(jour_astreinte)

astreinte_oculi = df[(df['nom']=="OCULI") & (df['jours_astreinte'])]
astreinte_alimi = df[(df['nom']=="ALIMI") & (df['jours_astreinte'])]
astreinte_bourbotte = df[(df['nom']=="BOURBOTTE") & (df['jours_astreinte'])]
astreinte_marinello = df[(df['nom']=="MARINELLO") & (df['jours_astreinte'])]
astreinte_mathelin = df[(df['nom']=="MATHELIN") & (df['jours_astreinte'])]
astreinte_meneses = df[(df['nom']=="MENESES G.") & (df['jours_astreinte'])]

ind_excl = astreinte_oculi.index.tolist() + astreinte_alimi.index.tolist() + astreinte_bourbotte.index.tolist() + astreinte_marinello.index.tolist() + astreinte_mathelin.index.tolist() + astreinte_meneses.index.tolist()

df = df.drop(ind_excl)

#######################

coeff_list = {"ALIMI": 1.3, "BOURBOTTE": 1.24, "MENESES G.": 1.3, "LAROCHE": 1.3, "OCULI": 1.3, "MATHELIN": 1.3, "MARINELLO": 1.35, "COLLI": 1.3}
                 
dfinale= df.groupby(['datetime','nom'])['distance_dis_2d'].agg(['count'])

dfinale['kilometre parcourus']= df.groupby(['datetime','nom'])['distance_dis_2d'].agg(['sum']).round(0)
dfinale['temps en minutes']= df.groupby(['datetime','nom'])['time_delta_minutes'].agg(['sum']).round(0)

dfinale = dfinale.reset_index()

dfinale['coeff'] = dfinale['nom'].map(coeff_list)
dfinale['kilometres corrigés'] = dfinale['kilometre parcourus']*dfinale['coeff']

dfinale['-100km']=np.where(dfinale['kilometres corrigés'] < 100, 1, 0)
dfinale['100-150km'] = np.where((dfinale['kilometres corrigés'] >= 100) & (dfinale['kilometres corrigés'] < 150), 1, 0)
dfinale['150-200km'] = np.where((dfinale['kilometres corrigés'] >= 150) & (dfinale['kilometres corrigés'] < 200), 1, 0)
dfinale['200-250km']= np.where((dfinale['kilometres corrigés'] >= 200) & (dfinale['kilometres corrigés'] < 250), 1, 0)
dfinale['250-300km']= np.where((dfinale['kilometres corrigés'] >= 250) & (dfinale['kilometres corrigés'] < 300), 1, 0)
dfinale['+ de 300km']=np.where(dfinale['kilometres corrigés']>=300,1,0)

dfinale=dfinale[(dfinale['kilometres corrigés']>3)]

########################

dfstat= dfinale.groupby(['nom'])['kilometres corrigés'].agg(['sum'])

dfstat= dfstat.rename(columns={"sum":"km total"})

dfstat['min km']= dfinale.groupby(['nom'])['kilometres corrigés'].agg(['min']).round(0)
dfstat['max km']= dfinale.groupby(['nom'])['kilometres corrigés'].agg(['max']).round(0)
dfstat['mean km']= dfinale.groupby(['nom'])['kilometres corrigés'].agg(['mean']).round(0)


dfstat['nb trajets -100 km']= dfinale.groupby(['nom'])['-100km'].agg(['sum'])
dfstat['nb trajets 100-150 km']= dfinale.groupby(['nom'])['100-150km'].agg(['sum'])
dfstat['nb trajets 150-200 km']= dfinale.groupby(['nom'])['150-200km'].agg(['sum'])
dfstat['nb trajets 200-250 km']= dfinale.groupby(['nom'])['200-250km'].agg(['sum'])
dfstat['nb trajets 250-300 km']= dfinale.groupby(['nom'])['250-300km'].agg(['sum'])
dfstat['nb trajets +300 km']= dfinale.groupby(['nom'])['+ de 300km'].agg(['sum'])

excel_writer = pd.ExcelWriter(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\final_file.xlsx", engine='xlsxwriter')

df_filtre.to_excel(excel_writer, sheet_name='data')
dfinale.to_excel(excel_writer, sheet_name='kilometrage_par_jour')
dfstat.to_excel(excel_writer, sheet_name='stat_sur_la_periode')

excel_writer._save()

xlsx_file = r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\final_file.xlsx"
sheet_2 = 'stat_sur_la_periode'

df_main = pd.read_excel(xlsx_file)
dfstat = pd.read_excel(xlsx_file, sheet_name=sheet_2)

dfstat.to_csv(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\CSV\final_file.csv", index=False)
dfstat.to_excel(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\XLSX\final_file.xlsx", index=False)

########################

df_main['datetime'] = pd.to_datetime(df_main['datetime'])

current_date = None
for idx, row in df_main.iterrows():
    if not pd.isnull(row['datetime']):
        current_date = row['datetime']
    else:
        df_main.at[idx, 'datetime'] = current_date

wb = openpyxl.load_workbook(xlsx_file)

plots_sheet = wb.create_sheet(title='Plots')

wb.save(xlsx_file)


plt.figure(figsize=(8, 6))

dfstat = dfstat.sort_values(by='mean km', ascending=False)

top_moyenne_km = sns.barplot(x="nom", y="mean km", data=dfstat, palette='Set2')

for p in top_moyenne_km.patches:
    height = p.get_height()
    top_moyenne_km.annotate(f'{height:.0f}', (p.get_x() + p.get_width() / 2., height),
                ha='center', va='bottom', fontsize=10)

plt.xlabel('')
plt.ylabel('km')
plt.title('Moyenne de km par tech')

plt.xticks(rotation=45)  
plt.tight_layout()

plt.savefig(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\top_moyenne_km.png")

image = Image(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\top_moyenne_km.png")

plots_sheet.add_image(image, "B2")

wb.save(xlsx_file)


sub_df = df_main['kilometres corrigés']

plt.figure(figsize=(8, 6))

distrib_km = sns.histplot(data=sub_df, bins=10, kde=True, color='skyblue')

plt.title('Distribution de la variable km parcourus')
plt.ylabel("km parcourus")
plt.xlabel('')

plt.savefig(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\distrib_km")

image2 = Image(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\distrib_moyenne_km.png")

plots_sheet.add_image(image2, "O2")

wb.save(xlsx_file)


list_of_dfs = [] 

for nom, data in df_main.groupby('nom'):
    technician_df = data.copy() 
    list_of_dfs.append(technician_df)

mean_km_by_name = df_main.groupby('nom')['kilometres corrigés'].mean()

for technician_df in list_of_dfs:
    nom = technician_df['nom'].iloc[0] 
    plt.figure(figsize=(15, 7))
    sns.lineplot(data=technician_df, x='datetime', y='kilometres corrigés', color='blue')
    mean_km_for_technician = round(mean_km_by_name[nom], 2) 
    plt.axhline(y=mean_km_for_technician, color='r', linestyle='--', alpha=0.5, label=f"Moyenne: {mean_km_for_technician:.2f}")
    
    plt.legend()
    
    plt.xlabel('')
    plt.ylabel('km')
    plt.title(f'Evolution du nombre total de km parcourus par jour - {nom}')
    plt.grid(alpha=0.3)
    
    output_folder = r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots"
    plot_image_path = os.path.join(output_folder, f"km_total_jour_{nom}.png")
    plt.savefig(plot_image_path)
    plt.close()

added_technicians = []

for technician_df in list_of_dfs:
    nom = technician_df['nom'].iloc[0] 
    
    if nom not in added_technicians:
        plot_image_path = os.path.join(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots", f"km_total_jour_{nom}.png")
        imgtech = Image(plot_image_path)
        
        ws = plots_sheet
        
        if added_technicians:  
             next_row += 40
        else:
            next_row = 130
        
        added_technicians.append(nom)
        
        ws.add_image(imgtech, f"C{next_row}")
        
        wb.save(xlsx_file)


melted_df = dfstat.melt(id_vars=['nom'], value_vars=['min km', 'max km', 'mean km'])

melted_df = melted_df.sort_values(by='value', ascending=False)

plt.figure(figsize=(12, 8))

melted_barplot = sns.barplot(data=melted_df, x='nom', y='value', hue='variable', palette='Paired')

mean_value = melted_df[melted_df['variable'] == 'mean km']['value'].mean()

plt.axhline(y=mean_value, color='red', linestyle='--', label=f'Agg. avg = {mean_value:.2f}')

plt.xticks(rotation=45, ha='right')

for p in melted_barplot.patches:
    height = p.get_height()
    melted_barplot.annotate(f'{height:.0f}', (p.get_x() + p.get_width() / 2., height),
                ha='center', va='bottom', fontsize=10)

plt.xlabel('')
plt.ylabel('Km')
plt.title('Kilométrage par technicien')

plt.grid(alpha=0.2)

plt.legend(title='')

plt.savefig(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\melted_barplot")

image6 = Image(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\melted_barplot.png")

plots_sheet.add_image(image6, "E37")

wb.save(xlsx_file)


tranche_cols = ['nb trajets -100 km', 'nb trajets 100-150 km', 'nb trajets 150-200 km', 'nb trajets 200-250 km', 'nb trajets 250-300 km', 'nb trajets +300 km']
df_tranches = dfstat[['nom'] + tranche_cols]

df_melted = df_tranches.melt(id_vars=['nom'], var_name='tranche', value_name='jours')

df_melted = df_melted.sort_values(by="jours", ascending=False)

plt.figure(figsize=(12, 8))

sns.set_palette("Set2")

tranche_cols = ['-100km', '100-150km', '150-200km', '200-250km', '250-300km', '+ de 300km']

custom_palette = sns.color_palette("Spectral", len(tranche_cols))[::-1]
custom_cmap = sns.color_palette(custom_palette, as_cmap=True)

stacked_bar = sns.barplot(data=df_melted, x='nom', y='jours', hue='tranche', palette=custom_cmap)

plt.xlabel('')
plt.ylabel('Nombre de jours')
plt.title('Nombre de trajets par tranche de kilomètres, par technicien')
plt.xticks(rotation=45)
plt.tight_layout()

plt.grid(alpha=0.2)

plt.legend(title='Tranche de kilomètres')

plt.savefig(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\grouped_barplot")

image7 = Image(r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\output results\plots\grouped_barplot.png")

plots_sheet.add_image(image7, "E82")

wb.save(xlsx_file)


print('-------- SUCCESSFUL --------')

 