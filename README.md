# Travel-and-consumption-of-a-vehicles-fleet
Analysis of travels and consumption of a vehicles fleet for a private company

![Texte alternatif](pic/sheet_5_censored.jpg)

## Context
The project was conducted for Vinci Facilities, a subsidiary of Vinci Energies responsible for maintaining urban infrastructures. Planned interventions required daily travel by technicians. The project anticipated new regulations on low-emission zones scheduled for 2025, imposing restrictions on polluting vehicles.

## Objectives
Identify the travel and consumption habits of technicians.
Anticipate the electrification of the vehicle fleet before the 2025 restrictions.
Gradually transition to 100% electric by 2025 by assigning electric cars based on driving habits.

## Methodology
GPS tracking campaign for 8 technicians over 3 months (July to September 2023).
Fuel consumption monitoring via the Total platform.
Merge of GPS tracker, consumption, and technician assignment data.
Creation of a Power BI dashboard to analyze habits.
Data Preparation
Use of Python scripts to restructure GPS tracker data, clean consumption data, and merge data.

## Data Processing and Cleaning
GPS Tracker (Invoxia):
Aggregation of data per day and per tracker.
Adjustment of kilometers with a coefficient based on user data.

Consumption (Total):
Restructuring and cleaning of data.
Filtering to obtain technician data.

Technician Assignment:
Data fusion to create a main dataframe.

## Merge of Files
External join between the assignment and consumption tracking files.

## Dashboard Creation
Development of a blueprint to define the structure.
Creation of a mock-up on Canva.
Exploration of the dashboard on Power BI.

## Dashboard Sections
Home:
User manual for the dashboard.

Global Indicators:
Gauges for kilometers, liters consumed, and L/100km.
Visualization of individual averages.

Tracker Analysis (km):
Stacked barchart for categories of trips.
Details on the number of trips and kilometer intervals.

Consumption Analysis (Liters + L/100km):
Grouped barchart for the distribution of L/100km values per month.
Average curve in red.

Indicators per Technician:
Individual selection with a geographical map of routes.

Indicators per Vehicles:
Comparison of habits for each type of vehicle.

## Conclusion
The dashboard met expectations by clearly identifying differences in travel habits. It allowed the definition of a list of technicians prioritized for the allocation of electric vehicles. Positive user feedback opens up prospects for other projects aimed at optimizing intervention management and technician schedules.
