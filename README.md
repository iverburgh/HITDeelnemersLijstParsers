# HITSail — HIT Aanmeldingen Statistieken

PowerShell-script dat statistieken genereert over deelnemersaanmeldingen voor HIT-kampen (Scouting Nederland), op basis van een CSV-export uit het aanmeldingssysteem.

## Wat doet het?

`Get-HitStatistic.ps1` leest een CSV-bestand en toont een overzichtelijk rapport:

```
Statistieken over aanmeldingen:
- 14 dames, 5 heren

- 10 subgroepjes
- 5 x alleen opgegeven
- 1 x met z'n tweeën
- 4 x met z'n drieën

- 6 x 15 jaar oud
- 9 x 16 jaar oud
- 4 x 17 jaar oud

- geen jarige deelnemers tijdens hit

- 6 x uit Gelderland
- 6 x uit Zuid-Holland
- 2 x uit Drenthe
- 2 x uit Overijssel
- 1 x uit Fryslân
- 1 x uit Groningen
- 1 x uit Utrecht
```

### Onderdelen

| Sectie | Beschrijving |
|---|---|
| **Geslacht** | Verdeling dames/heren |
| **Subgroepen** | Aantal groepjes + verdeling op grootte (alleen, tweeën, drieën, …) |
| **Leeftijden** | Leeftijd berekend op de startdatum van het kamp |
| **Jarigen** | Wie is er jarig tijdens het kamp (met naam, nieuwe leeftijd en datum) |
| **Provincies** | Woonprovincie per deelnemer, opgehaald via de PDOK Locatieserver API op basis van postcode |

## Gebruik

```powershell
.\Get-HitStatistic.ps1
```

Het script vraagt interactief om het jaar te bevestigen (standaard: huidig jaar) en berekent
automatisch Goede Vrijdag en 2e Paasdag als kampdatums.

### Parameters

| Parameter | Verplicht | Beschrijving |
|---|---|---|
| `-Year` | Nee | Jaar van het HIT-kamp (standaard: huidig jaar). De gebruiker krijgt een interactieve prompt om te bevestigen of te wijzigen. |
| `-Verbose` | Nee | Toont gedetailleerde voortgangsberichten (API-calls, CSV-import, Paasberekening, etc.) |

### CSV-selectie

Bij het starten zoekt het script automatisch naar `*.csv`-bestanden in dezelfde map als het script:

- **1 CSV gevonden** → wordt automatisch gebruikt
- **Meerdere CSVs** → er verschijnt een genummerde lijst waaruit je kiest

### Voorbeeld met uitgebreide output

```powershell
.\Get-HitStatistic.ps1 -Year 2025 -Verbose
```

### Paasberekening

Het script berekent automatisch de Paasdatums via het **Computus-algoritme** (Anonymous Gregorian / Meeus-Jones-Butcher):

- **Goede Vrijdag** = Paaszondag − 2 dagen (startdatum kamp)
- **2e Paasdag** = Paaszondag + 1 dag (einddatum kamp)

Voorbeeld voor 2025:
```
Goede Vrijdag:  18 april 2025
2e Paasdag:     21 april 2025
```

## Vereisten

- **PowerShell 5.1+** of **PowerShell 7+**
- **Internetverbinding** — voor postcode → provincie lookup via de [PDOK Locatieserver API](https://api.pdok.nl/bzk/locatieserver/search/v3_1/ui/)
  - Bij geen verbinding wordt een grove fallback gebruikt op basis van het eerste cijfer van de postcode (minder nauwkeurig)

## CSV-formaat

Het script verwacht een puntkomma-gescheiden (`;`) CSV zoals geëxporteerd uit het Scouting Nederland aanmeldingssysteem. De volgende kolommen zijn vereist:

| Kolom | Voorbeeld | Gebruik |
|---|---|---|
| `Lid geslacht` | `Vrouw` / `Man` | Geslachtsverdeling |
| `Subgroepnaam` | `De Dolfijnen` | Subgroep-analyse |
| `Lid geboortedatum` | `13-06-2008` | Leeftijd + jarigen (format: `dd-MM-yyyy`) |
| `Lid postcode` | `2311EJ` | Provincie via PDOK |
| `Lid voornaam` | `Sophie` | Naam bij jarigen |
| `Lid tussenvoegsel` | `van` | Naam bij jarigen |
| `Lid achternaam` | `Bakker` | Naam bij jarigen |

Overige kolommen worden genegeerd en hoeven niet aanwezig te zijn.

## Foutafhandeling

| Situatie | Gedrag |
|---|---|
| `EndDate` vóór `StartDate` | Kan niet meer voorkomen (automatisch berekend) |
| Ongeldig jaar opgegeven | Terminating error met duidelijke melding |
| Geen CSV-bestanden in de map | Terminating error |
| Leeg CSV-bestand | Terminating error |
| Vereiste kolommen ontbreken | Terminating error met lijst van ontbrekende kolommen |
| Ongeldige geboortedatum | Warning per deelnemer, overgeslagen |
| PDOK API niet bereikbaar | Warning + fallback op eerste-cijfer postcodetabel |

## Licentie

Dit project is bedoeld voor intern gebruik binnen Scouting HIT-organisaties.
