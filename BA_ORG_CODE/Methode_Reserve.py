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
# Beispielwerte
rente = 1200
zins = 0.0025
n = 80  # Laufzeit der Rente (z.B. von 40 bis 80 Jahren)
uebersterblichkeit = 1.0
alter = 40

# Berechnung der Reserve
barwert = reserve_berechnung.berechne_reserve_jährlich(alter, rente, uebersterblichkeit, zins, n)
print(f"Der Barwert der Rente beträgt: {barwert:.2f}")