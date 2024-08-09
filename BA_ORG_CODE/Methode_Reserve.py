import pandas as pd

class ReserveBerechnung:
    def __init__(self, death_table):
        self.death_table = death_table
            
    
    def berechne_reserve_jährlich(self, alter, rente, uebersterblichkeit, zins, n):
       # v = Abzinsfaktor
        v = 1 / (1 + zins)
        adjusted_qx = self.death_table["q x+0"] * uebersterblichkeit # Wegen übersterblichkeit
        px = [1 - qx for qx in adjusted_qx]
        lx = [1000000]  # Startwert für lx
        for i in range(1, len(px)):
            lx.append(lx[i - 1] * px[i - 1])
        barwertfaktor = 0
        print(f"Länge von adjusted_qx: {len(adjusted_qx)}")
        print(f"Länge von px: {len(px)}")
        print(f"Länge von lx: {len(lx)}")

        n= int(n)
        alter = int(alter)
        for k in range(1, n+1): 
            barwertfaktor += (v ** k )* (lx[alter + k ] / lx[alter])
        barwert = rente * barwertfaktor
        return barwert
    
    """
    def berechne_reserve_monatlich(self, alter, rente, uebersterblichkeit):
        adjusted_qx = self.sterbetafel["qx"] * uebersterblichkeit
        barwert = rente * sum(self.sterbetafel["lx"] * adjusted_qx / 12)
        return barwert

    # Weitere Methoden für halbjährlich, vierteljährlich, etc.

    #def berechne_reserve(self, zahlungsweise, alter, rente, uebersterblichkeit=1.0):
        if zahlungsweise == "jährlich":
            return self.berechne_reserve_jährlich(alter, rente, uebersterblichkeit)
        elif zahlungsweise == "monatlich":
            return self.berechne_reserve_monatlich(alter, rente, uebersterblichkeit)
        else:
            raise ValueError(f"Unbekannte Zahlungsweise: {zahlungsweise}")

# Erstelle ein Objekt der Klasse
reserve_berechnung = ReserveBerechnung(sterbetafel=sterbetafel_df)

gesamt_reserve = 0
for index, row in bestand_df.iterrows():
    barwert = reserve_berechnung.berechne_reserve(
        zahlungsweise=row["Zahlungsweise"], 
        alter=row["Alter"], 
        rente=row["Rente"],
        uebersterblichkeit=row.get("Übersterblichkeitsfaktor", 1.0)
    )
    gesamt_reserve += barwert

print(f"Die Gesamtreserve beträgt: {gesamt_reserve}")
"""
# Excel-Datei einlesen
file_path = '/Users/alicetangyie/Downloads/Uni/BachelorArbeit/GI_annuities_data_template.v10_GE_2024_Q1.xlsx' # Muss von Anwender angepasst werden
bestand_df = pd.read_excel(file_path, sheet_name="Tabelle2", header=0)
death_table_df = pd.read_excel(file_path, sheet_name="Death Table", header=3)

# Überblick über die Spalten
#print(bestand_df.columns)
#print(death_table_df.columns)

reserve_berechnung = ReserveBerechnung(death_table=death_table_df)

for index, row in bestand_df.iterrows():
    alter = row["Age"]
    rente = row["Rente"]
    zins = row.get("Zins", 0.0025)  # Verwende den Zinssatz aus der Tabelle oder den Standardwert
    uebersterblichkeit = row.get("Übersterblichkeit", 1.0)  # Verwende den Übersterblichkeitsfaktor oder den Standardwert
   # zahlungsweise = row.get("Zahlungsweise", "jährlich")  # Verwende die Zahlungsweise oder den Standardwert
    n = row.get("Laufzeit", 80)  # Verwende die Laufzeit oder den Standardwert

# Berechnung der Reserve
barwert = reserve_berechnung.berechne_reserve_jährlich(alter, rente, uebersterblichkeit, zins, n)
bestand_df.at[index, 'Reserve'] = barwert
bestand_df.to_excel(file_path, sheet_name="Tabelle2", index=False)
print(bestand_df.head())

print(f"Der Barwert der Rente beträgt: {barwert:.2f}")