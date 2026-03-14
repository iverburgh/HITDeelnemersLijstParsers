# HITDeelnemersLijstParsers

PowerShell-scriptsuite voor het verwerken van HIT-deelnemerslijsten (Scouting Nederland), op basis van een Excel-export (.xlsx) uit het aanmeldingssysteem.

## Scripts

| Script | Doel |
|---|---|
| `Get-HitStatistic.ps1` | Toont statistieken over de aanmeldingen als tekstrapport |
| `Export-HitBijzonderheden.ps1` | Exporteert een Excel-lijst met alleen deelnemers met dieet of aandachtspunten |
| `Export-HitContactgegevens.ps1` | Exporteert een Excel-lijst met contactgegevens van alle deelnemers |
| `Compare-HitDeelnemerslijst.ps1` | Vergelijkt twee deelnemerslijsten op naam en geboortedatum (terugkerende deelnemers) |
| `Mail01-3_Weken_voor_Goede_Vrijdag.ps1` | Genereert een kopieerklare e-mail (BCC, onderwerp, body) voor de 'Het is bijna zover!'-mailing |
| `Mail02-1_Dag_voor_Merchandise_Deadline.ps1` | Genereert een kopieerklare herinneringsmail over de aankomende merchandise-besteldatum |
| `Mail03-1_Week_voor_Goede_Vrijdag.ps1` | Genereert een kopieerklare e-mail voor 1 week vóór het kamp, met automatisch opgehaalde weersvoorspelling |
| `HitHelpers.psm1` | Gedeelde helperfuncties, automatisch ingeladen door de scripts |

---

## Get-HitStatistic.ps1

Leest het deelnemers-Excel-bestand en toont een overzichtelijk rapport:

```
Goede Vrijdag:  3 april 2026
2e Paasdag:     6 april 2026
Statistieken over aanmeldingen [Zeilzwerf Fryslân] [2026]:
- 12 dames, 13 heren

- 13 subgroepjes
- 6 x alleen opgegeven
- 5 x met z'n tweeën
- 1 x met z'n vieren
- 1 x met z'n vijven

- 1 x 14 jaar oud
- 7 x 15 jaar oud
- 6 x 16 jaar oud
- 5 x 17 jaar oud
- 6 x 18 jaar oud

- Jan de Vries wordt 17 op 5 april
```

### Onderdelen rapport

| Sectie | Beschrijving |
|---|---|
| **Geslacht** | Verdeling dames/heren |
| **Subgroepen** | Aantal groepjes en verdeling op grootte |
| **Leeftijden** | Leeftijd per deelnemer berekend op de startdatum van het kamp |
| **Jarigen** | Wie is er jarig tijdens het kamp (naam, nieuwe leeftijd, datum) |

### Gebruik

```powershell
.\Get-HitStatistic.ps1
.\Get-HitStatistic.ps1 -Year 2025
.\Get-HitStatistic.ps1 -Verbose
```

Het script vraagt interactief om het jaar te bevestigen (standaard: huidig jaar).

### Parameters

| Parameter | Verplicht | Beschrijving |
|---|---|---|
| `-Year` | Nee | Jaar van het HIT-kamp (standaard: huidig jaar) |
| `-Verbose` | Nee | Toont gedetailleerde voortgangsberichten |

---

## Export-HitBijzonderheden.ps1

Exporteert een Excel-bestand met alleen deelnemers die een dieet of aandachtspunten hebben.

Kolommen in het outputbestand: `Voornaam`, `Achternaam`, `Geslacht`, `Leeftijd`, `Bijzonderheden`

- **Leeftijd** wordt berekend op de startdatum van het kamp. Als een deelnemer jarig is tijdens het kamp, worden beide leeftijden getoond (bijv. `16/17`).
- **Bijzonderheden** combineert de velden `Dieet` en `Aandachtspunten`, gescheiden door ` | `.

Outputbestandsnaam: `Deelnemerslijst_[naam]_[jaar]_Bijzonderheden.xlsx`

### Gebruik

```powershell
.\Export-HitBijzonderheden.ps1
.\Export-HitBijzonderheden.ps1 -Year 2025
.\Export-HitBijzonderheden.ps1 -Verbose
```

### Parameters

| Parameter | Verplicht | Beschrijving |
|---|---|---|
| `-Year` | Nee | Jaar van het HIT-kamp (standaard: huidig jaar) |
| `-Verbose` | Nee | Toont gedetailleerde voortgangsberichten |

---

## Export-HitContactgegevens.ps1

Exporteert een Excel-bestand met contactgegevens van alle deelnemers, gesorteerd op voornaam en achternaam.

Kolommen in het outputbestand: `Voornaam`, `Achternaam`, `Geslacht`, `Geboortedatum`, `Contactpersoon`, `Noodnummer`, `Lid mobiel`

Telefoonnummers en geboortedatum worden opgeslagen als tekst (voorloopnullen blijven behouden).

Outputbestandsnaam: `Deelnemerslijst_[naam]_[jaar]_Contactgegevens.xlsx`

### Gebruik

```powershell
.\Export-HitContactgegevens.ps1
.\Export-HitContactgegevens.ps1 -Year 2026
.\Export-HitContactgegevens.ps1 -Verbose
```

### Parameters

| Parameter | Verplicht | Beschrijving |
|---|---|---|
| `-Year` | Nee | Jaar voor de outputbestandsnaam (standaard: huidig jaar) |
| `-Verbose` | Nee | Toont gedetailleerde voortgangsberichten |

---

## Mail01-3_Weken_voor_Goede_Vrijdag.ps1

Genereert een kopieerklare e-mail voor de 'Het is bijna zover!'-mailing, klaar om te plakken in Gmail.

De output bestaat uit drie afzonderlijk te kopiëren secties:
- **BCC** — alle e-mailadressen van deelnemers uit het Excel-bestand (kolom `Mailadres`)
- **Onderwerp** — `[KampNaam] - Het is bijna zover!`
- **Body** — volledig opgemaakt bericht met deelnemersinformatie, merchandise en formulierverzoek

Bovenaan de output verschijnt een opvallende waarschuwing met de uiterste verzenddatum (3 weken vóór Goede Vrijdag).

Het script berekent automatisch:
- De kampnaam uit de kolom `Kamp` in het Excel-bestand
- Het aantal weken tot het kamp (op basis van de huidige datum)
- De uiterste besteldatum voor merchandise (donderdag 22:00, twee weken vóór het kamp)
- De uiterste verzenddatum van de e-mail (drie weken vóór Goede Vrijdag)

### Gebruik

```powershell
.\Mail01-3_Weken_voor_Goede_Vrijdag.ps1
.\Mail01-3_Weken_voor_Goede_Vrijdag.ps1 -TikkieLink "https://tikkie.me/pay/abc123" -GoogleFormLink "https://forms.gle/xyz789"
.\Mail01-3_Weken_voor_Goede_Vrijdag.ps1 -Verbose
```

### Parameters

| Parameter | Verplicht | Standaard | Beschrijving |
|---|---|---|---|
| `-Year` | Nee | Huidig jaar | Jaar van het HIT-kamp, voor paasdatumberekening |
| `-TikkieLink` | Nee | `https://tikkielink.nl/` | Tikkie-betaallink voor merchandise |
| `-DeelnemersinformatieLink` | Nee | `https://deelnemers.informatie.nl/` | Link naar de deelnemersinformatiepagina |
| `-GoogleFormLink` | Nee | `https://google.form.nl/` | Link naar het Google-formulier voor merchandise en aanvullende vragen |
| `-MerchandisePad` | Nee | `Merchandise.png` | Pad naar de merchandise-afbeelding (alleen bestandsnaam wordt in de mail getoond) |
| `-EmailKolom` | Nee | `Mailadres` | Kolomnaam in het Excel-bestand met de e-mailadressen |
| `-Verbose` | Nee | — | Toont gedetailleerde voortgangsberichten |

---

## Mail02-1_Dag_voor_Merchandise_Deadline.ps1

Genereert een kopieerklare herinneringsmail over de aankomende merchandise-besteldatum, klaar om te plakken in Gmail. Verstuur deze mail de dag vóór de merchandise-deadline (dus de donderdag ervoor).

De merchandise-deadline is de **donderdagavond om 22:00, twee weken vóór de start van het kamp**.
Ga terug vanuit (campStart − 14 dagen) naar de laatste donderdag op of vóór die datum.
Deze mail verstuur je uiterlijk **twee dagen vóór de deadline** (de dinsdag ervoor).

De output bestaat uit drie afzonderlijk te kopiëren secties:
- **BCC** — alle e-mailadressen van deelnemers uit het Excel-bestand (kolom `Mailadres`)
- **Onderwerp** — `[KampNaam] - Reminder merchandise bestelling`
- **Body** — korte herinnering met de uiterste besteldatum

Bovenaan de output verschijnt een waarschuwing met de uiterste verzenddatum (de donderdag vóór de merchandise-deadline).

### Gebruik

```powershell
.\Mail02-1_Dag_voor_Merchandise_Deadline.ps1
.\Mail02-1_Dag_voor_Merchandise_Deadline.ps1 -TikkieLink "https://tikkie.me/pay/abc123" -GoogleFormLink "https://forms.gle/xyz789"
.\Mail02-1_Dag_voor_Merchandise_Deadline.ps1 -Year 2027
.\Mail02-1_Dag_voor_Merchandise_Deadline.ps1 -Verbose
```

### Parameters

| Parameter | Verplicht | Standaard | Beschrijving |
|---|---|---|---|
| `-Year` | Nee | Huidig jaar | Jaar van het HIT-kamp, voor paasdatumberekening |
| `-TikkieLink` | Nee | `https://tikkielink.nl/` | Tikkie-betaallink voor merchandise |
| `-GoogleFormLink` | Nee | `https://google.form.nl/` | Link naar het Google-formulier voor de merchandise bestelling |
| `-EmailKolom` | Nee | `Mailadres` | Kolomnaam in het Excel-bestand met de e-mailadressen |
| `-Verbose` | Nee | — | Toont gedetailleerde voortgangsberichten |

---

## Mail03-1_Week_voor_Goede_Vrijdag.ps1

Genereert een kopieerklare e-mail voor de '1 week voor het kamp'-mailing, klaar om te plakken in Gmail.

De output bestaat uit drie afzonderlijk te kopiëren secties:
- **BCC** — alle e-mailadressen van deelnemers uit het Excel-bestand (kolom `Mailadres`)
- **Onderwerp** — `[KampNaam] - Nog maar 1 week!`
- **Body** — volledig opgemaakt bericht met start- en afsluitingsinfo, formulierstatus, en weersvoorspelling

Bovenaan de output verschijnt een waarschuwing met de uiterste verzenddatum (exact 1 week vóór Goede Vrijdag = campStart − 7 dagen, ook een vrijdag).

Het script berekent en haalt automatisch op:
- De kampnaam uit de kolom `Kamp` in het Excel-bestand
- De uiterste verzenddatum van de e-mail (vrijdag, 1 week vóór Goede Vrijdag)
- De weersvoorspelling voor het kampweekend via de **Open-Meteo API** (gratis, geen registratie, locatie Sneek)  
  — bij een mislukte API-aanroep verschijnt een waarschuwing en wordt een plaatshouder in de body geplaatst

### Gebruik

```powershell
.\Mail03-1_Week_voor_Goede_Vrijdag.ps1 -AantalIngevuldeFormulieren 16
.\Mail03-1_Week_voor_Goede_Vrijdag.ps1 -AantalIngevuldeFormulieren 20 -GoogleFormLink "https://forms.gle/xyz789" -DeelnemersinformatieLink "https://hit.scouting.nl/.../file"
.\Mail03-1_Week_voor_Goede_Vrijdag.ps1 -AantalIngevuldeFormulieren 16 -Year 2027
.\Mail03-1_Week_voor_Goede_Vrijdag.ps1 -AantalIngevuldeFormulieren 16 -Verbose
```

### Parameters

| Parameter | Verplicht | Standaard | Beschrijving |
|---|---|---|---|
| `-Year` | Nee | Huidig jaar | Jaar van het HIT-kamp, voor paasdatumberekening |
| `-AantalIngevuldeFormulieren` | **Ja** | — | Aantal deelnemers dat het Google-formulier al heeft ingevuld (te controleren in Google Forms) |
| `-DeelnemersinformatieLink` | Nee | `https://deelnemers.informatie.nl/` | Link naar de deelnemersinformatiepagina |
| `-GoogleFormLink` | Nee | `https://google.form.nl/` | Link naar het Google-formulier voor aanvullende kampinformatie |
| `-StartLocatie` | Nee | `Scoutingcentrum Sneek in Sneek` | Naam en plaats van de startlocatie van het kamp |
| `-StartTijd` | Nee | `19:00` | Tijdstip waarop deelnemers aanwezig moeten zijn voor de start |
| `-OpeningsTijd` | Nee | `19:30` | Tijdstip waarop het kamp officieel opent |
| `-MedeKampNaam` | Nee | `HIT Sail Fryslân` | Naam van het mede-kamp waarmee de afsluiting gezamenlijk plaatsvindt |
| `-AfsluitingsTijd` | Nee | `13:00` | Tijdstip van de gezamenlijke afsluiting |
| `-WelkomOudersTijd` | Nee | `12:30` | Tijdstip waarop ouders welkom zijn bij de afsluiting |
| `-EmailKolom` | Nee | `Mailadres` | Kolomnaam in het Excel-bestand met de e-mailadressen |
| `-Verbose` | Nee | — | Toont gedetailleerde voortgangsberichten |

---

## Compare-HitDeelnemerslijst.ps1

Vergelijkt twee deelnemerslijsten (xlsx of csv) en zoekt naar deelnemers die in beide jaren voorkomen, op basis van **volledige naam** en **geboortedatum**.

De gebruiker kiest interactief via een console-keuzemenu:
1. **Deelnemerslijst 1** — de lijst van het voorgaande jaar
2. **Deelnemerslijst 2** — de lijst van het huidige jaar

De output toont het aantal deelnemers per lijst en een tabel met de overeenkomende deelnemers (naam + geboortedatum).

```
══════════════════════════════════════════════════
  Resultaat
══════════════════════════════════════════════════

Deelnemerslijst 1: 25 deelnemer(s)
  D:\...\Zeilzwerf_2025-alles.xlsx
Deelnemerslijst 2: 28 deelnemer(s)
  D:\...\Zeilzwerf_2026-alles.xlsx

In deelnemerslijst 2 staan 7 deelnemer(s) die ook in deelnemerslijst 1 staan:

VolledigeNaam      Geboortedatum
-------------      -------------
Jan de Vries       13-06-2008
...
```

### Ondersteunde bestandsformaten

Beide bestanden mogen elk afzonderlijk `.xlsx` of `.csv` zijn. Kolommen worden automatisch gedetecteerd:

| Formaat | Voornaam | Tussenvoegsel | Achternaam | Geboortedatum |
|---|---|---|---|---|
| Raw CSV (puntkomma) | `Lid voornaam` | `Lid tussenvoegsel` | `Lid achternaam` | `Lid geboortedatum` |
| Excel (hernoemde headers) | `Voornaam` | `Tussenvoegsel` (optioneel) | `Achternaam` | `Geboortedatum` |

### Gebruik

```powershell
.\Compare-HitDeelnemerslijst.ps1
```

Geen parameters — al het interactieve verloopt via console-prompts.

---

## Bestandsselectie (alle scripts)

Alle scripts zoeken automatisch naar het invoerbestand in dezelfde map als het script:

1. Zoek op `*-alles.xlsx`
   - 1 treffer → automatisch geselecteerd
   - Meerdere treffers → interactief keuzemenu
2. Als geen `*-alles.xlsx` gevonden: zoek op `*.xlsx` (exclusief eerder gegenereerde `Deelnemerslijst_*`-bestanden), met dezelfde selectielogica
3. Als ook dat niets oplevert → foutmelding

`Compare-HitDeelnemerslijst.ps1` gebruikt een ruimere zoekopdracht: alle `*.xlsx`- en `*.csv`-bestanden in de scriptmap (exclusief `Deelnemerslijst_*`). De gebruiker kiest telkens interactief voor zowel lijst 1 als lijst 2.

Gegenereerde outputbestanden (`Deelnemerslijst_*`) worden bij de automatische selectie altijd overgeslagen.

---

## Excel-bestandsformaat

Alle scripts verwachten een `.xlsx`-bestand zoals geëxporteerd uit het Scouting Nederland aanmeldingssysteem. De volgende kolommen worden gebruikt:

| Kolom | Gebruikt door | Beschrijving |
|---|---|---|
| `Kamp` | Statistieken | Kampnaam in de rapporttitel |
| `Voornaam` | Alle scripts | Voornaam deelnemer |
| `Achternaam` | Alle scripts | Achternaam deelnemer |
| `Gender` | Alle scripts | `man` of `vrouw` |
| `Geboortedatum` | Alle scripts | Geboortedatum (datetime of tekst) |
| `Subgroep` | Statistieken | Naam van de subgroep |
| `Dieet` | Bijzonderheden | Dieetwens / -beperking |
| `Aandachtspunten` | Bijzonderheden | Medische/allergie-aandachtspunten |
| `Naam noodcontact` | Contactgegevens | Naam van de contactpersoon |
| `Telefoonnummer noodcontact` | Contactgegevens | Telefoonnummer noodcontact |
| `Mobiel` | Contactgegevens | Mobiel nummer van de deelnemer |
| `Mailadres` | E-mailgenerator | E-mailadres van de deelnemer, gebruikt voor de BCC-lijst |

Overige kolommen worden genegeerd.

---

## Vereisten

- **PowerShell 5.1+** of **PowerShell 7+**
- **ImportExcel-module** — wordt automatisch geïnstalleerd als die nog niet aanwezig is (`Install-Module -Scope CurrentUser`). Hiervoor is eenmalig een internetverbinding nodig.

---

## Foutafhandeling

| Situatie | Gedrag |
|---|---|
| Geen xlsx- of csv-bestand gevonden | Terminating error met duidelijke melding |
| Leeg Excel-bestand | Terminating error |
| Vereiste kolommen ontbreken | Terminating error met naam van ontbrekende kolom(men) |
| Onbekend kolomformaat in vergelijking | Warning per rij, rij overgeslagen |
| Ongeldig jaar opgegeven | Terminating error |
| Ongeldige geboortedatum | Warning per deelnemer, deelnemer overgeslagen |
| ImportExcel kan niet worden geïnstalleerd | Terminating error |

---

## Licentie

Dit project is bedoeld voor intern gebruik binnen Scouting HIT-organisaties.
