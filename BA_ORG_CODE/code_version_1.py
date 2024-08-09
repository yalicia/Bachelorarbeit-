import pandas as pd

# Excel-Datei einlesen
file_path = "pfad_zur_excel_datei.xlsx"
bestand_df = pd.read_excel(file_path, sheet_name="Bestand")
sterbetafel_df = pd.read_excel(file_path, sheet_name="Tabelle 2")

# Überblick über die Spalten
print(bestand_df.columns)
print(sterbetafel_df.columns)
