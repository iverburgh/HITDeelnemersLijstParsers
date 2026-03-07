# HITDeelnemersLijstParsers

PowerShell-scriptsuite voor het verwerken van HIT-deelnemerslijsten (Scouting Nederland), op basis van een Excel-export (.xlsx) uit het aanmeldingssysteem.

## Scripts

| Script | Doel |
|---|---|
| `Get-HitStatistic.ps1` | Toont statistieken over de aanmeldingen als tekstrapport |
| `Export-HitBijzonderheden.ps1` | Exporteert een Excel-lijst met alleen deelnemers met dieet of aandachtspunten |
| `Export-HitContactgegevens.ps1` | Exporteert een Excel-lijst met contactgegevens van alle deelnemers |
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

## Bestandsselectie (alle scripts)

Alle scripts zoeken automatisch naar het invoerbestand in dezelfde map als het script:

1. Zoek op `*-alles.xlsx`
   - 1 treffer → automatisch geselecteerd
   - Meerdere treffers → interactief keuzemenu
2. Als geen `*-alles.xlsx` gevonden: zoek op `*.xlsx` (exclusief eerder gegenereerde `Deelnemerslijst_*`-bestanden), met dezelfde selectielogica
3. Als ook dat niets oplevert → foutmelding

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

Overige kolommen worden genegeerd.

---

## Vereisten

- **PowerShell 5.1+** of **PowerShell 7+**
- **ImportExcel-module** — wordt automatisch geïnstalleerd als die nog niet aanwezig is (`Install-Module -Scope CurrentUser`). Hiervoor is eenmalig een internetverbinding nodig.

---

## Foutafhandeling

| Situatie | Gedrag |
|---|---|
| Geen xlsx-bestand gevonden | Terminating error met duidelijke melding |
| Leeg Excel-bestand | Terminating error |
| Vereiste kolommen ontbreken | Terminating error met naam van ontbrekende kolom(men) |
| Ongeldig jaar opgegeven | Terminating error |
| Ongeldige geboortedatum | Warning per deelnemer, deelnemer overgeslagen |
| ImportExcel kan niet worden geïnstalleerd | Terminating error |

---

## Licentie

Dit project is bedoeld voor intern gebruik binnen Scouting HIT-organisaties.
