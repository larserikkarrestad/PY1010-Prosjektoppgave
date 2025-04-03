"""
PY1010 - Prosjektoppgave

@author: Lars-Erik Karrestad
"""
# Import av nødvendige biblioteker for å løse oppgavene
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

# ##### Del A #####
# Lese inn Excel-filen
fil = pd.read_excel('support_uke_24.xlsx')

# Legge kolonner fra Excel i et array
u_dag = fil['Ukedag'].values
kl_slett = fil['Klokkeslett'].values
varighet = fil['Varighet'].values
score = fil['Tilfredshet'].values

# ##### Del B #####
# Henter ukedagene
data = u_dag

# Oppretter en tom dictionary
forekomster = {}

# Iterer gjennom arrayet og teller forekomster av ukedager
# Legger antall forekomster til i dictionary'en
for dag in data:
    if dag in forekomster:
        forekomster[dag] += 1
    else:
        forekomster[dag] = 1

# Sorterer etter ukedagene
ukedager = ["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag"]
antall_forekomster = [forekomster.get(dag, 0) for dag in ukedager]

# Lager stolpediagram
plt.bar(ukedager, antall_forekomster)
plt.xlabel('Ukedag')
plt.ylabel('Antall henvendelser')
plt.title('Antall henvendelser for hver ukedag')
plt.show()

# ##### Del C #####
# Henter samtaletider
samtaletider = varighet

# Konverter tidene til datetime-objekter
tid_objekter = [datetime.strptime(tid, "%H:%M:%S") for tid in samtaletider]

# Finner korteste og lengste samtaletid ved bruk av min og max
korteste_tid = min(tid_objekter).strftime("%H:%M:%S")
lengste_tid = max(tid_objekter).strftime("%H:%M:%S")

# Skriver resultatet til skjerm
print(f"Korteste samtaletid: {korteste_tid}")
print(f"Lengste samtaletid: {lengste_tid}")

# ##### Del D #####
# Henter data
# Her gjenbrukes variabelen "samtaletider" fra oppgavens del C

# Konverter tidene til sekunder
totale_sekunder = 0
for tid in samtaletider:
    tid_objekt = datetime.strptime(tid, "%H:%M:%S")
    sekunder = tid_objekt.hour * 3600 + tid_objekt.minute * 60 + tid_objekt.second
    totale_sekunder += sekunder

# Beregner gjennomsnittet i sekunder
gjennomsnitt_sekunder = totale_sekunder / len(samtaletider)

# Konverter gjennomsnittet tilbake til HH:MM:SS
gjennomsnitt_tid = str(timedelta(seconds=int(gjennomsnitt_sekunder)))

# Skriver resultatet til skjerm
print(f"Gjennomsnittlig samtaletid: {gjennomsnitt_tid}\n")

# ##### Del E #####
# Gjenbruker variabelen "fil" fra oppgavens del a
# Konverter klokkeslett til datetime-objekt for å kunne filtrere enklere
fil["Klokkeslett"] = pd.to_datetime(fil["Klokkeslett"], format="%H:%M:%S").dt.time

# Oppretter kategorier for tidsrommene
tidsrom_kategorier = {
    "08-10": (8, 10),
    "10-12": (10, 12),
    "12-14": (12, 14),
    "14-16": (14, 16)
}

# Teller antall henvendelser i hvert tidsrom
tidsrom_telling = {}
for tidsrom, (start, slutt) in tidsrom_kategorier.items():
    tidsrom_telling[tidsrom] = fil[(fil["Klokkeslett"] >= pd.to_datetime(f"{start}:00:00").time()) &
                                    (fil["Klokkeslett"] < pd.to_datetime(f"{slutt}:00:00").time())].shape[0]

# Plotter resultatet i et kakediagram
plt.figure(figsize=(8, 6))
plt.pie(tidsrom_telling.values(), labels=tidsrom_telling.keys(), autopct="%1.1f%%", startangle=90)
plt.title("Henvendelser per tidsrom")
plt.show()

# ##### Del F #####
# Fjerner rader der tilfredshet ikke er angitt. 
gyldige_data = fil.dropna(subset=["Tilfredshet"])

# Teller antall negative, nøytrale og positive svar
negative = gyldige_data[gyldige_data["Tilfredshet"] <= 6].shape[0]
noytrale = gyldige_data[(gyldige_data["Tilfredshet"] >= 7) & (gyldige_data["Tilfredshet"] <= 8)].shape[0]
positive = gyldige_data[gyldige_data["Tilfredshet"] >= 9].shape[0]

# Beregner prosentandel av innkomne svar og regner ut NPS
total_respons = negative + noytrale + positive
prosent_negative = (negative / total_respons) * 100
prosent_positive = (positive / total_respons) * 100
nps = prosent_positive - prosent_negative

# Skriver resultatet til skjerm
print(f"Supportavdelingens NPS er: {nps:.2f}", "prosent\n")
print("NPS'en er basert på disse tallene:\n")
print(f"Totalt antall gyldige tilbakemeldinger: {total_respons}")
print(f"Antall negative tilbakemeldinger: {negative}")
print(f"Antall nøytrale tilbakemeldinger: {noytrale}")
print(f"Antall positive tilbakemeldinger: {positive}")
