# HITDeelnemersLijstParsers

PowerShell-scriptsuite voor het verwerken van HIT-deelnemerslijsten (Scouting Nederland), op basis van een Excel-export (.xlsx) uit het aanmeldingssysteem.

## Scripts

| Script | Doel |
|---|---|
| `Get-HitStatistic.ps1` | Toont statistieken over de aanmeldingen als tekstrapport |
| `Export-HitBijzonderheden.ps1` | Exporteert een Excel-lijst met alleen deelnemers met dieet of aandachtspunten |
| `Export-HitContactgegevens.ps1` | Exporteert een Excel-lijst met contactgegevens van alle deelnemers |
| `GenerateBootIndelingBaseData.ps1` | Genereert een Excel-basisbestand voor de bootindeling, gecombineerd uit het deelnemersbestand en de Google Forms-export |
| `GenerateMerchandiseBestelling.ps1` | Toont een merchandise-bestellingsoverzicht (hoodies en t-shirts) voor HIT Sail Fryslân en Zeilzwerf Fryslân op de console |
| `Mail01-3_Weken_voor_Goede_Vrijdag.ps1` | Genereert een kopieerklare e-mail (BCC, onderwerp, body) voor de 'Het is bijna zover!'-mailing |
| `Mail02-1_Dag_voor_Merchandise_Deadline.ps1` | Genereert een kopieerklare herinneringsmail over de aankomende merchandise-besteldatum |
| `Mail03-1_Week_voor_Goede_Vrijdag.ps1` | Genereert een kopieerklare e-mail voor 1 week vóór het kamp, met automatisch opgehaalde weersvoorspelling |
| `HitHelpers.psm1` | Gedeelde helperfuncties, automatisch ingeladen door de scripts |

---

## Get-HitStatistic.ps1

Leest twee deelnemers-Excel-bestanden (huidig jaar en vorig jaar) en toont een overzichtelijk rapport op de console. De gebruiker selecteert beide bestanden via een interactief keuzemenu.

```
Goede Vrijdag:  3 april 2026
2e Paasdag:     6 april 2026
Statistieken over aanmeldingen Zeilzwerf Fryslân 2026 :
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

- 3 deelnemer(s) waren er vorig jaar ook bij:
  - Anna de Boer (12-03-2010)
  - Bram van Dijk (05-07-2009)
  - Laura Jansen (22-11-2008)
```

### Onderdelen rapport

| Sectie | Beschrijving |
|---|---|
| **Geslacht** | Verdeling dames/heren |
| **Subgroepen** | Aantal groepjes en verdeling op grootte |
| **Leeftijden** | Leeftijd per deelnemer berekend op de startdatum van het kamp |
| **Jarigen** | Wie is er jarig tijdens het kamp (naam, nieuwe leeftijd, datum) |
| **Terugkerende deelnemers** | Deelnemers die ook in de vorigjaarlijst voorkomen (naam + geboortedatum) |

### Gebruik

```powershell
.\Get-HitStatistic.ps1
.\Get-HitStatistic.ps1 -Year 2025
.\Get-HitStatistic.ps1 -Verbose
```

Het script vraagt interactief om het jaar te bevestigen. Vervolgens kiest de gebruiker via een keuzemenu het huidigjaar-bestand, gevolgd door het vorigjaar-bestand (ter vergelijking).

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

Kolommen in het outputbestand: `Groep`, `Voornaam`, `Achternaam`, `Geslacht`, `Geboortedatum`, `Contactpersoon`, `Noodnummer`, `Lid mobiel`, `Bijzonderheden`

- **Bijzonderheden** bevat automatisch gegenereerde aantekeningen, gescheiden door ` - `:
  - De verjaardagsdatum als een deelnemer jarig is tijdens het kamp (bijv. `vrijdag 3 april jarig (wordt 15)`).
  - `Was er vorig jaar ook` als de deelnemer ook in het optioneel geselecteerde vorigjaar-bestand staat.
- Telefoonnummers en geboortedatum worden opgeslagen als tekst (voorloopnullen blijven behouden).

Na het kiezen van het huidigjaar-bestand vraagt het script optioneel om een **vorigjaar-bestand** (keuze `[0]` of lege invoer = overslaan). Als dit bestand is geselecteerd, wordt de `Bijzonderheden`-kolom aangevuld met `Was er vorig jaar ook` voor terugkerende deelnemers.

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

Genereert een kopieerklare herinneringsmail over de aankomende merchandise-besteldatum, klaar om te plakken in Gmail. Verstuur deze mail uiterlijk twee dagen vóór de merchandise-deadline (de dinsdag ervoor).

De merchandise-deadline is de **donderdagavond om 22:00, twee weken vóór de start van het kamp**.
Ga terug vanuit (campStart − 14 dagen) naar de laatste donderdag op of vóór die datum.
Deze mail verstuur je uiterlijk **twee dagen vóór de deadline** (de dinsdag ervoor).

De output bestaat uit drie afzonderlijk te kopiëren secties:
- **BCC** — alle e-mailadressen van deelnemers uit het Excel-bestand (kolom `Mailadres`)
- **Onderwerp** — `[KampNaam] - Reminder merchandise bestelling`
- **Body** — korte herinnering met de uiterste besteldatum

Bovenaan de output verschijnt een waarschuwing met de uiterste verzenddatum (twee dagen vóór de merchandise-deadline, de dinsdag ervoor).

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
| `-TestWeerDatum` | Nee | Huidige datum | Startdatum voor de weerscheck, uitsluitend voor testdoeleinden. Wordt genegeerd als Goede Vrijdag binnen het 16-daagse voorspellingsvenster van Open-Meteo valt; anders worden datums vanaf deze datum gebruikt |
| `-EmailKolom` | Nee | `Mailadres` | Kolomnaam in het Excel-bestand met de e-mailadressen |
| `-Verbose` | Nee | — | Toont gedetailleerde voortgangsberichten |

---

## GenerateBootIndelingBaseData.ps1

Combineert het deelnemersbestand (xlsx of puntkomma-gescheiden csv) met een Google Forms-export (komma-gescheiden csv) en genereert een Excel-bestand als basismateriaal voor de bootindeling.

De gebruiker kiest interactief:
1. **Het deelnemersbestand** — xlsx of csv met alle ingeschreven deelnemers
2. **De Google Forms-export** — csv met de antwoorden op de pré-kampformulier

Kolommen in het outputbestand:

| Kolom | Bron | Beschrijving |
|---|---|---|
| `Naam` | Deelnemersbestand | Voornaam; als meerdere deelnemers dezelfde voornaam hebben ook de achternaam |
| `Leeftijd` | Deelnemersbestand | Leeftijd op de startdag van het kamp; `14/15` als de deelnemer jarig is tijdens het kamp |
| `Geslacht` | Deelnemersbestand | Geslacht van de deelnemer |
| `Groep` | Deelnemersbestand | Subgroep van de deelnemer |
| `Zeilen met je groep?` | Google Forms | Afgeleid van de bezwaar-vraag: bezwaar=Ja → `Ja` (wil met groep), bezwaar=Nee → `Nee` |
| `Zeilervaring` | Google Forms | Antwoord op de zeilervaring-vraag |
| `Eigen reddingsvest?` | Google Forms | `Ja` of `Nee`, genormaliseerd vanuit het formulier |
| `Gewicht (kg)` | Google Forms | Opgegeven gewicht (nodig bij huur reddingsvest) |
| `Emailadres` | Deelnemersbestand | E-mailadres van de deelnemer |

- Rijen worden gesorteerd op **Groep**, daarbinnen op **Naam**.
- Deelnemers die het formulier **niet** hebben ingevuld worden **rood gemarkeerd** (RGB 255, 199, 206).
- Google Forms-kolommen worden herkend via **substring-match** op de kolomnaam, zodat kleine tekstverschillen in formuliervragen geen breuk veroorzaken.
- Namen worden vergeleken via **fuzzy matching** (standaard: 85% gelijkenis). Zo wordt bijv. `"Yvonne's van Buuren"` (formulier) automatisch gematcht op `"Yvonne van Buuren"` (deelnemersbestand). Fuzzy matches worden als informatieve melding op de console getoond.

Outputbestandsnaam: `Deelnemerslijst_[naam]_[jaar]_BootIndeling.xlsx`

### Gebruik

```powershell
.\GenerateBootIndelingBaseData.ps1
.\GenerateBootIndelingBaseData.ps1 -Year 2026
.\GenerateBootIndelingBaseData.ps1 -MatchThreshold 0.90
.\GenerateBootIndelingBaseData.ps1 -Verbose
```

### Parameters

| Parameter | Verplicht | Standaard | Beschrijving |
|---|---|---|---|
| `-Year` | Nee | Huidig jaar | Jaar van het HIT-kamp (paasdatum + leeftijdsberekening) |
| `-MatchThreshold` | Nee | `0.85` | Minimale naamsovereenkomst (0.0–1.0) voor fuzzy matching. Verlaag bij veel schrijfvarianten, verhoog bij risico op verkeerde koppelingen |
| `-Verbose` | Nee | — | Toont gedetailleerde voortgangsberichten, inclusief alle fuzzy matches |

---

## GenerateMerchandiseBestelling.ps1

Genereert een merchandise-bestellingsoverzicht voor **HIT Sail Fryslân** en **Zeilzwerf Fryslân**, gebaseerd op twee Google Forms-exportbestanden (csv).

De gebruiker kiest interactief:
1. **Het CSV-bestand voor HIT Sail Fryslân** — export uit het merchandise-formulier
2. **Het CSV-bestand voor Zeilzwerf Fryslân** — export uit het merchandise-formulier

Per evenement worden hoodies en t-shirts gegroepeerd op itemtype (Hoodie → T-shirt) en maat (XS → XXXL) en afgedrukt op de console, met voor elk item de voornamen van de bestellers.

Er wordt geen outputbestand aangemaakt — de output verschijnt alleen op de console.

### Verwacht CSV-formaat

| Kolom | Beschrijving |
|---|---|
| `Wat is je voor- en achternaam?` | Volledige naam van de besteller |
| `Hoodie € 39,95` | Gekozen hoodie-maat (bijv. `M`, `XL`), leeg als niet besteld |
| `T-shirt € 39,95` | Gekozen t-shirt-maat (bijv. `S`, `L`), leeg als niet besteld |

### Gebruik

```powershell
.\GenerateMerchandiseBestelling.ps1
.\GenerateMerchandiseBestelling.ps1 -Verbose
```

Geen bestandsparameters — de CSV-selectie verloopt volledig via interactieve console-prompts.

---

## Bestandsselectie

Scripts gebruiken twee verschillende methodes om het invoerbestand te selecteren:

### Resolve-HitExcelPath — automatisch, met prioriteit op `*-alles.xlsx`

Gebruikt door: `Export-HitBijzonderheden.ps1`, `Mail01-3_Weken_voor_Goede_Vrijdag.ps1`, `Mail02-1_Dag_voor_Merchandise_Deadline.ps1`, `Mail03-1_Week_voor_Goede_Vrijdag.ps1`

1. Zoek op `*-alles.xlsx` in de scriptmap
   - 1 treffer → automatisch geselecteerd
   - Meerdere treffers → interactief keuzemenu
2. Als geen `*-alles.xlsx` gevonden: zoek op `*.xlsx` (exclusief `Deelnemerslijst_*`-bestanden), met dezelfde selectielogica
3. Als ook dat niets oplevert → foutmelding

### Select-HitFilePath — interactief keuzemenu voor xlsx én csv

Gebruikt door: `Get-HitStatistic.ps1`, `Export-HitContactgegevens.ps1`, `GenerateBootIndelingBaseData.ps1`, `GenerateMerchandiseBestelling.ps1`

Toont een keuzemenu met alle `*.xlsx`- en `*.csv`-bestanden in de scriptmap (exclusief `Deelnemerslijst_*`). Bij slechts één beschikbaar bestand (zonder `-AllowSkip`) wordt automatisch geselecteerd. Scripts die meerdere bestanden nodig hebben (bijv. huidigjaar + vorigjaar) roepen het menu twee keer aan; het **als eerste geselecteerde bestand wordt bij de tweede keuze automatisch uitgesloten** zodat hetzelfde bestand niet dubbel gekozen kan worden.

`Export-HitContactgegevens.ps1` en `Get-HitStatistic.ps1` bieden ook de `-AllowSkip`-optie voor het tweede bestand: keuze `[0]` of lege invoer slaat de selectie over en retourneert `$null`.

Gegenereerde outputbestanden (`Deelnemerslijst_*`) worden altijd overgeslagen.

---

## Excel-bestandsformaat

Alle scripts verwachten een `.xlsx`-bestand zoals geëxporteerd uit het Scouting Nederland aanmeldingssysteem. De volgende kolommen worden gebruikt:

| Kolom | Gebruikt door | Beschrijving |
|---|---|---|
| `Kamp` | Statistieken, E-mailgenerator | Kampnaam in de rapporttitel en het e-mailonderwerp |
| `Voornaam` | Alle scripts | Voornaam deelnemer |
| `Achternaam` | Alle scripts | Achternaam deelnemer |
| `Gender` | Alle scripts | `man` of `vrouw` |
| `Geboortedatum` | Alle scripts | Geboortedatum (datetime of tekst) |
| `Subgroep` | Statistieken, Contactgegevens, Bootindeling | Subgroep van de deelnemer (kolom `Groep` in de output); fallbacks: `Groep`, `Subgroepnaam` |
| `Dieet` | Bijzonderheden | Dieetwens / -beperking |
| `Aandachtspunten` | Bijzonderheden | Medische/allergie-aandachtspunten |
| `Naam noodcontact` | Contactgegevens | Naam van de contactpersoon |
| `Telefoonnummer noodcontact` | Contactgegevens | Telefoonnummer noodcontact |
| `Mobiel` | Contactgegevens | Mobiel nummer van de deelnemer |
| `Mailadres` | E-mailgenerator, Bootindeling | E-mailadres van de deelnemer, gebruikt voor de BCC-lijst en de bootindelingsexport |

Overige kolommen worden genegeerd.

---

## HitHelpers.psm1

Gedeelde module die automatisch wordt ingeladen door alle scripts. Bevat:

| Functie | Beschrijving |
|---|---|
| `Get-EasterSunday` | Berekent Eerste Paasdag (Meeus/Jones/Butcher-algoritme) |
| `Get-AgeAtDate` | Berekent leeftijd in jaren op een gegeven datum |
| `Get-BirthdayDuringCamp` | Controleert of een geboortedatum valt binnen de kampperiode |
| `Get-BirthdayDateDuringCamp` | Geeft de concrete datum terug waarop een persoon jarig is tijdens het kamp (`$null` als niet van toepassing) |
| `Get-DutchMonthName` / `Get-DutchDayName` | Nederlandse dag- en maandnamen |
| `Get-DutchGroupSizeLabel` | Nederlandse omschrijving voor groepsgrootte |
| `Assert-HitImportExcel` | Installeert de ImportExcel-module automatisch als die ontbreekt |
| `Resolve-HitExcelPath` | Zoekt automatisch het `*-alles.xlsx`-bestand op in de scriptmap |
| `Select-HitFilePath` | Toont een interactief keuzemenu voor alle xlsx/csv-bestanden in de scriptmap. De optionele parameter `-ExcludePaths` sluit opgegeven paden uit van de lijst (gebruikt door `GenerateBootIndelingBaseData.ps1` en `GenerateMerchandiseBestelling.ps1` om het al gekozen bestand te verbergen bij de tweede keuze). Met `-AllowSkip` verschijnt optie `[0] Overslaan` en retourneert het menu `$null` |
| `ConvertFrom-HitBirthDate` | Parseert geboortedatums in meerdere formaten naar `DateTime` |
| `Get-HitOutputBaseName` | Genereert een schone basisnaam voor het outputbestand op basis van het inputpad |
| `Import-ParticipantFile` | Importeert een deelnemersbestand (xlsx of csv) en retourneert de rijen |
| `ConvertTo-NormalizedParticipant` | Normaliseert een rij uit een deelnemersbestand naar een gestandaardiseerd object met `Sleutel`, `VolledigeNaam` en `Geboortedatum` |
| `Get-HitCampDates` | Berekent de kampdatums voor een jaar op basis van Pasen: `CampStart` (Goede Vrijdag), `CampEnd` (Tweede Paasdag) en `EasterSunday` |
| `Get-HitMerchandiseDeadline` | Berekent de uiterste merchandise-besteldatum: donderdag 22:00, twee weken vóór het kamp. Geeft `DateTime`, `Formatted` en `Time` terug |
| `Import-HitMailData` | Laadt het Excel-bestand, leest de kampnaam (`Kamp`-kolom) en bouwt de BCC-string op. Geeft `KampNaam`, `BccString`, `AllRows` en `ActualColumns` terug |
| `Write-HitMailOutput` | Schrijft de gestandaardiseerde console-uitvoer: deadline-banner (geel/rood) gevolgd door afgescheiden `BCC`-, `ONDERWERP`- en `EMAIL BODY`-secties |

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
| Google Forms-kolom niet gevonden (`voor- en achternaam`) | Terminating error met lijst van gevonden kolommen |
| Formuliernaam niet exact of fuzzy te matchen | Deelnemer rood gemarkeerd in Excel; melding op console |
| Formulierregel niet te koppelen aan deelnemer | Gele warning op console met regelnummer en naam |
| ImportExcel kan niet worden geïnstalleerd | Terminating error |

---

## Licentie

Dit project is bedoeld voor intern gebruik binnen Scouting HIT-organisaties.
