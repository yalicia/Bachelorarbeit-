import pandas as pd

class ReserveBerechnung:
    def __init__(self, death_table):
        self.death_table = death_table
            
    def berechne_reserve(self, alter, rente, uebersterblichkeit, zins, n, sex, escRate,freq):
        v = 1 / (1 + zins)  # Abzinsfaktor

        if sex == 0:  #  "0" for male gender
            adjusted_qx = self.death_table["q x+0"] * uebersterblichkeit  # Übersterblichkeit berücksichtigen
            px = [1 - qx for qx in adjusted_qx]
            lx = [1000000]  # Startwert für lx
            for i in range(1, len(px)):
                lx.append(lx[i - 1] * px[i - 1])
        elif sex == 1:  #  "1" for female gender
            adjusted_qy = self.death_table["q y+0"] * uebersterblichkeit  # Übersterblichkeit berücksichtigen
            py = [1 - qy for qy in adjusted_qy]
            lx = [1000000]  # Startwert für lx (using lx to keep consistency)
            for i in range(1, len(py)):
                lx.append(lx[i - 1] * py[i - 1])
        else:
            return "Ungültige Geschlechtseingabe"
        geburtsjahr = [entry_year - alter]

        if alter + n >= len(lx):
            raise IndexError("Der Wert von alter + n überschreitet die Länge von lx.")
        if escRate == "YES":
            rente = rente * (1 +  Rate)
        
        barwertfaktor = 0
        barwertfaktornormal = 0
        n = int(n)
        alter = int(alter)
       
        
        for k in range(1, n + 1):
            if freq == 1:
                barwertfaktor += (v ** k) * (lx[alter + k] / lx[alter])
            else :
                barwertfaktornormal += (v ** k) * (lx[alter + k] / lx[alter])
                barwertfaktor = barwertfaktornormal - ((((lx[alter+n] * (v ** (alter+n))))/(lx[alter] * (v ** alter))-1)* ((freq-1)/2*freq))
        barwert = rente * barwertfaktor
        return barwert

# Excel-Datei einlesen
file_path = '/Users/alicetangyie/Downloads/Uni/BachelorArbeit/GI_annuities_data_template.v10_GE_2024_Q1.xlsx'  # Muss vom Anwender angepasst werden
bestand_df = pd.read_excel(file_path, sheet_name="Inputs - MPs", header=5)
death_table_df = pd.read_excel(file_path, sheet_name="Death Table", header=3)

# Benutzer nach dem Zinssatz fragen
zins = float(input("Bitte geben Sie den Zinssatz ein (z.B. 0.0025 für 0.25%): "))
Rate = float(input("Bitte geben Sie den esc Rate ein (z.B. 0.009 für 0.9%): "))
# Überblick über die Spalten
#print(bestand_df.columns)
#print(death_table_df.columns)

reserve_berechnung = ReserveBerechnung(death_table=death_table_df)

for index, row in bestand_df.iterrows():
    alter = row["AGE_AT_ENTRY"]
    rente = row["ANN_ANNUITY"]
    uebersterblichkeit = row.get("Q_CORR_PN", 1.0)  # Verwende den Übersterblichkeitsfaktor oder den Standardwert
    n = row["POL_TERM_Y"]  # Verwende die Laufzeit aus der Tabelle
    sex = row["SEX"]
    escRate = row["ESC_RATE"] 
    freq = row["ANNUITY_FREQ"]
    entry_year = row["ENTRY_YEAR"]
    
    # Debug-Ausgabe der Eingabewerte
    print(f"Zeile {index}: Alter={alter}, Rente={rente}, Zins={zins}, Übersterblichkeit={uebersterblichkeit}, Laufzeit={n}, Geschlecht={sex}, escRate={escRate}")

    # Berechnung der Reserve
    barwert = reserve_berechnung.berechne_reserve(alter, rente, uebersterblichkeit, zins, n, sex, escRate,freq)
    bestand_df.at[index, 'Reserve'] = barwert

# Speichern der Ergebnisse in einer neuen Excel-Datei
new_file_path = '/Users/alicetangyie/Downloads/Uni/BachelorArbeit/GI_annuities_data_template_with_reserves.xlsx'
bestand_df.to_excel(new_file_path, sheet_name="Tabelle2", index=False)

print(f"Die Ergebnisse wurden in die neue Datei gespeichert: {new_file_path}")
print(bestand_df.head())

