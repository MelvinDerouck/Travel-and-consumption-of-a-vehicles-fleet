
"""
Created on Mon Sept 18 14:31:07 2023

Author : melvin.derouck

"""

from geopy import distance
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
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


print('-------- SUCCESSFUL --------')

 