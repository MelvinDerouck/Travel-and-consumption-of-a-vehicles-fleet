# -*- coding: utf-8 -*-
"""
03/08/2023 14:25:41

@author: melvin.derouck (03/08/2023)

"""
import os
import datetime
import pandas as pd
import folium

csv_file = r"C:\Users\melvin.derouk\VINCI Energies\VF DAUPHINE SAVOIE - METHODES\16-Melvin DEROUCK\P8 - Accompagner le passage à l'électrique du parc auto\INVOXIA\export Invoxia\export_2023-07-03_2023-09-30.csv"
data = pd.read_csv(csv_file, sep=",")

filtered_data = data[data['name'] == "b'T11'"]

min_lat, max_lat, min_lon, max_lon = min(filtered_data['lat']), max(filtered_data['lat']), min(filtered_data['lng']), max(filtered_data['lng'])

mymap = folium.Map(location=[(min_lat + max_lat) / 2, (min_lon + max_lon) / 2], zoom_start=10)
                              
points = list(zip(filtered_data['lat'], filtered_data['lng']))
folium.PolyLine(points, color='blue', weight=1.5, opacity=1).add_to(mymap)

mymap.save('my_folium_map.html')
os.system('my_folium_map.html')