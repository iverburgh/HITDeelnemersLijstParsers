<#
    .SYNOPSIS
    Genereert een kopieerklare e-mail voor de '1 week voor het kamp'-mailing.

    .DESCRIPTION
    Leest de e-mailadressen van deelnemers uit het HIT-aanmeldingen Excel-bestand en
    genereert een informatieve e-mail die verstuurd wordt 1 week vóór Goede Vrijdag,
    klaar om te kopiëren naar Gmail.

    De output bestaat uit drie afzonderlijke secties:
    - BCC:       alle e-mailadressen van deelnemers uit het Excel-bestand
    - Onderwerp: "[KampNaam] - Nog maar 1 week!"
    - Body:      een bericht met start- en afsluitingsinfo, formulierstatus, en weersvoorspelling

    Het script haalt automatisch de weersvoorspelling op via de Open-Meteo API (gratis, geen
    API-sleutel vereist) voor de locatie Sneek, voor het camp-startweekend. Als het ophalen
    mislukt, verschijnt er een waarschuwing en wordt een plaatshouder in de body geplaatst.

    Bovenaan de output verschijnt een waarschuwing met de uiterste verzenddatum
    (exact 1 week vóór Goede Vrijdag = campStart - 7 dagen).

    Het script zoekt automatisch naar een bestand met het patroon "*-alles.xlsx" in de
    scriptmap. Als er meerdere zijn, verschijnt een keuzemenu. Als er geen zijn, wordt
    gezocht op "*.xlsx". Als er dan ook geen is, volgt een foutmelding.

    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is.

    .PARAMETER Year
    Het jaar van het HIT-kamp. Wordt gebruikt voor de paasdatumberekening.
    Standaard: het huidige jaar.

    .PARAMETER AantalIngevuldeFormulieren
    Het actuele aantal deelnemers dat het Google-formulier al heeft ingevuld.
    Dit getal moet handmatig worden opgegeven (te controleren in de Google Forms-respons).

    .PARAMETER DeelnemersinformatieLink
    De link naar de deelnemersinformatiepagina op het HIT-portaal.
    Standaard: 'https://deelnemers.informatie.nl/'

    .PARAMETER GoogleFormLink
    De link naar het Google-formulier voor aanvullende kampinformatie.
    Standaard: 'https://google.form.nl/'

    .PARAMETER StartLocatie
    De naam en plaats van de startlocatie van het kamp.
    Standaard: 'Scoutingcentrum Sneek in Sneek'

    .PARAMETER StartTijd
    Het tijdstip waarop deelnemers aanwezig moeten zijn voor de start.
    Standaard: '19:00'

    .PARAMETER OpeningsTijd
    Het tijdstip waarop het kamp officieel opent.
    Standaard: '19:30'

    .PARAMETER MedeKampNaam
    De naam van het mede-kamp waarmee de gezamenlijke afsluiting plaatsvindt.
    Standaard: 'HIT Sail Fryslân'

    .PARAMETER AfsluitingsTijd
    Het tijdstip van de gezamenlijke afsluiting.
    Standaard: '13:00'

    .PARAMETER WelkomOudersTijd
    Het tijdstip waarop ouders welkom zijn bij de afsluiting.
    Standaard: '12:30'

    .PARAMETER TestWeerDatum
    Optionele startdatum voor de weerscheck, uitsluitend bedoeld voor testdoeleinden.
    Gebruik dit als Goede Vrijdag nog buiten het 16-daagse voorspellingsvenster van Open-Meteo valt.
    Als Goede Vrijdag wél binnen het venster valt, wordt deze parameter genegeerd en worden de
    echte kampdatums gebruikt. Formaat: 'yyyy-MM-dd' of een DateTime-waarde.

    .PARAMETER EmailKolom
    De naam van de kolom in het Excel-bestand die de e-mailadressen van de deelnemers bevat.
    Standaard: 'Mailadres'

    .EXAMPLE
    .\Mail03-1_Week_voor_Goede_Vrijdag.ps1 -AantalIngevuldeFormulieren 16

    Genereert de e-mail met standaardwaarden en 16 ingevulde formulieren.
    De kampnaam en weersvoorspelling worden automatisch opgehaald.

    .EXAMPLE
    .\Mail03-1_Week_voor_Goede_Vrijdag.ps1 `
        -AantalIngevuldeFormulieren 20 `
        -GoogleFormLink "https://forms.gle/xyz789" `
        -DeelnemersinformatieLink "https://hit.scouting.nl/.../file" `
        -Year 2027

    Genereert de e-mail met alle parameters expliciet ingevuld.

    .EXAMPLE
    .\Mail03-1_Week_voor_Goede_Vrijdag.ps1 -AantalIngevuldeFormulieren 16 -Verbose

    Toont uitgebreide voortgangsberichten, inclusief de opgehaalde weersdata.

    .OUTPUTS
    System.String
    Drie secties (BCC, Onderwerp, Body) als tekstuitvoer op de console.

    .NOTES
    Vereist: PowerShell 5.1+, ImportExcel-module.
    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is
    (via Install-Module -Scope CurrentUser). Hiervoor is een internetverbinding nodig.
    Weersdata: Open-Meteo API (https://api.open-meteo.com), vrij beschikbaar, geen registratie vereist.
    Paasdatum-berekening: Anonymous Gregorian-algoritme (Meeus/Jones/Butcher).
#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = 'Jaar van het HIT-kamp (standaard: huidig jaar).')]
    [ValidateRange(2000, 2099)]
    [int]$Year = (Get-Date).Year,

    [Parameter(Mandatory, HelpMessage = 'Aantal deelnemers dat het Google-formulier al heeft ingevuld (te controleren in Google Forms).')]
    [ValidateRange(0, 9999)]
    [int]$AantalIngevuldeFormulieren,

    [Parameter(HelpMessage = 'Link naar de deelnemersinformatiepagina.')]
    [string]$DeelnemersinformatieLink = 'https://deelnemers.informatie.nl/',

    [Parameter(HelpMessage = 'Link naar het Google-formulier voor aanvullende kampinformatie.')]
    [string]$GoogleFormLink = 'https://google.form.nl/',

    [Parameter(HelpMessage = 'Naam en plaats van de startlocatie van het kamp.')]
    [string]$StartLocatie = 'Scoutingcentrum Sneek in Sneek',

    [Parameter(HelpMessage = 'Tijdstip waarop deelnemers aanwezig moeten zijn voor de start.')]
    [string]$StartTijd = '19:00',

    [Parameter(HelpMessage = 'Tijdstip waarop het kamp officieel opent.')]
    [string]$OpeningsTijd = '19:30',

    [Parameter(HelpMessage = 'Naam van het mede-kamp waarmee de gezamenlijke afsluiting plaatsvindt.')]
    [string]$MedeKampNaam = 'HIT Sail Fryslân',

    [Parameter(HelpMessage = 'Tijdstip van de gezamenlijke afsluiting.')]
    [string]$AfsluitingsTijd = '13:00',

    [Parameter(HelpMessage = 'Tijdstip waarop ouders welkom zijn bij de afsluiting.')]
    [string]$WelkomOudersTijd = '12:30',

    [Parameter(HelpMessage = 'Startdatum voor de weerscheck. Wordt automatisch genegeerd als Goede Vrijdag binnen het 16-daagse voorspellingsvenster valt; anders wordt deze datum gebruikt. Standaard: de huidige datum.')]
    [datetime]$TestWeerDatum = (Get-Date),

    [Parameter(HelpMessage = 'Kolomnaam in het Excel-bestand met de e-mailadressen van de deelnemers.')]
    [string]$EmailKolom = 'Mailadres'
)

#region Import Shared Module

$_moduleRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
Import-Module -Name (Join-Path -Path $_moduleRoot -ChildPath 'HitHelpers.psm1') -Force
Remove-Variable -Name _moduleRoot

#endregion Import Shared Module

#region Initialisation

Assert-HitImportExcel

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
$resolvedInputPath = Resolve-HitExcelPath -ScriptDir $scriptDir

Write-Verbose "Input-bestand: $resolvedInputPath"

#endregion Initialisation

#region Camp Date Calculation

$campDates = Get-HitCampDates -Year $Year
$campStart  = $campDates.CampStart
$campEnd    = $campDates.CampEnd

#endregion Camp Date Calculation

#region Mail Deadline Calculation

# Uiterste verzenddatum: exact 1 week vóór Goede Vrijdag (= campStart - 7 dagen = ook een vrijdag)
$mailDeadlineDate      = $campStart.AddDays(-7)
$mailDeadlineDayName   = Get-DutchDayName -DayOfWeek ([int]$mailDeadlineDate.DayOfWeek)
$mailDeadlineMonthName = Get-DutchMonthName -Month $mailDeadlineDate.Month
$mailDeadlineFormatted = "$mailDeadlineDayName $($mailDeadlineDate.Day) $mailDeadlineMonthName $($mailDeadlineDate.Year)"

Write-Verbose "Uiterste verzenddatum e-mail: $mailDeadlineFormatted"

#endregion Mail Deadline Calculation

#region Weather Forecast

# Open-Meteo API — vrij beschikbaar, geen API-sleutel vereist.
# Coördinaten Sneek: 53.0325°N, 5.6575°E
$weerOmschrijving = '[weersomschrijving kon niet automatisch worden opgehaald — vul dit handmatig in]'
$tempOverdag      = '?'
$tempNacht        = '?'
$weatherFetched   = $false

$campStartStr = $campStart.ToString('yyyy-MM-dd')
$campEndStr   = $campEnd.ToString('yyyy-MM-dd')

# Open-Meteo forecast reikt maximaal 16 dagen vooruit.
# Goede Vrijdag binnen het venster → echte kampdata; anders → TestWeerDatum (standaard: vandaag).
$forecastWindowEnd = (Get-Date).AddDays(16)
if ($campStart -le $forecastWindowEnd) {
    $weerStartStr = $campStartStr
    $weerEndStr   = $campEndStr
    Write-Verbose "Goede Vrijdag valt binnen het voorspellingsvenster: kampdata gebruikt voor weerscheck."
}
else {
    $weerStartStr = $TestWeerDatum.ToString('yyyy-MM-dd')
    $weerEndStr   = $TestWeerDatum.AddDays(3).ToString('yyyy-MM-dd')
    Write-Verbose "Kamp valt nog buiten de 16-daagse voorspelling: weerscheck op $weerStartStr t/m $weerEndStr."
}

try {
    $weatherUri = (
        'https://api.open-meteo.com/v1/forecast' +
        '?latitude=53.0325&longitude=5.6575' +
        '&daily=weather_code,temperature_2m_max,temperature_2m_min' +
        '&timezone=Europe%2FAmsterdam' +
        "&start_date=$weerStartStr&end_date=$weerEndStr"
    )

    Write-Verbose "Weersvoorspelling ophalen: $weatherUri"
    $weatherData = Invoke-RestMethod -Uri $weatherUri -Method Get -ErrorAction Stop

    # Gebruik de eerste dag (Goede Vrijdag) voor de weersomschrijving
    $wmoCode = [int]$weatherData.daily.weather_code[0]

    # Gemiddelde max- en min-temperatuur over het hele kampweekend
    $tempOverdag = [Math]::Round(($weatherData.daily.temperature_2m_max | Measure-Object -Average).Average)
    $tempNacht   = [Math]::Round(($weatherData.daily.temperature_2m_min | Measure-Object -Average).Average)

    # WMO Weather Interpretation Codes → Nederlandse omschrijving
    $weerOmschrijving = switch ($wmoCode) {
        { $_ -eq 0 }                          { 'stralend zonnig weer' }
        { $_ -in 1..2 }                       { 'overwegend zonnig met wat bewolking' }
        { $_ -eq 3 }                          { 'bewolkt' }
        { $_ -in @(45, 48) }                  { 'mist' }
        { $_ -in 51..57 }                     { 'motregen' }
        { $_ -in @(61, 63) }                  { 'lichte regen' }
        { $_ -in @(65, 67) }                  { 'zware regen' }
        { $_ -in @(71, 73, 75, 77, 85, 86) }  { 'sneeuw' }
        { $_ -in 80..82 }                     { 'regenbuien' }
        { $_ -ge 95 }                         { 'onweer' }
        default                                { 'wisselvallig weer' }
    }

    $weatherFetched = $true
    Write-Verbose "Weersvoorspelling: $weerOmschrijving, overdag ${tempOverdag}°C, nacht ${tempNacht}°C"
}
catch {
    $apiReason = $null
    if ($_.ErrorDetails.Message) {
        try { $apiReason = ($_.ErrorDetails.Message | ConvertFrom-Json).reason } catch {}
    }
    $foutmelding = if ($apiReason) { $apiReason } else { $_.Exception.Message }
    Write-Warning "Weersvoorspelling kon niet worden opgehaald via Open-Meteo: $foutmelding"
}

if ($weatherFetched) {
    $weatherSection = @"

Misschien hebben jullie het weerbericht al even bekeken: Overdag verwachten we $weerOmschrijving.
De temperatuur zal $tempOverdag graden overdag zijn en $tempNacht graden 's nachts.
"@
}
else {
    $weatherSection = @"

[Weersomschrijving kon niet automatisch worden opgehaald. Vul handmatig de verwachte weersomstandigheden in.]
"@
}

#endregion Weather Forecast

#region Import Data

$mailData  = Import-HitMailData -Path $resolvedInputPath -EmailKolom $EmailKolom
$KampNaam  = $mailData.KampNaam
$bccString = $mailData.BccString

#endregion Import Data

#region Build Email

$subject = "$KampNaam - Nog maar 1 week!"

$body = @"
Hallo,

Nog maar 1 week en dan begint 's avonds de $KampNaam!

*Start*
Zorg ervoor dat je rond $StartTijd op $StartLocatie bent, zodat we om $OpeningsTijd kunnen openen!
De routebeschrijving staat in de deelnemersinformatie die hier te vinden is: $DeelnemersinformatieLink

*Afsluiting*
De gezamenlijke afsluiting met $MedeKampNaam zal op maandagmiddag om $AfsluitingsTijd zijn bij $StartLocatie. Ouders zijn welkom vanaf $WelkomOudersTijd.

*Ontvangen extra informatie - Google Forms*
We hebben van $AantalIngevuldeFormulieren deelnemers alle extra informatie ontvangen. Dit gaat ons helpen met een goede bootverdeling!
Heb je dit formulier nog niet ingevuld? Doe het graag snel via: $GoogleFormLink

*Bootindeling*
Deze wordt op de eerste avond bekendgemaakt. Wij houden zoveel mogelijk rekening met de ontvangen extra informatie.

*Weer*$weatherSection
*Wat neem ik mee?*
- Zonnebrand
Op het water verbrand je sneller dan op de kant. Neem sterke zonnebrand mee zodat je niet verbrandt. Een pet, zonnebril en een hervulbare fles voor drinken helpt ook om tegen de zon te kunnen.
- Zeil/regenkleding
Als het regent tijdens het zeilen krijg je natte billen. Hier krijg je het heel koud van. Een zeilbroek of regenbroek voorkomt dit. Neem dit mee. Een water/winddichte jas erboven maakt het helemaal comfortabel.
- Isolatiematje
De meeste deelnemers slapen in de boten (lelievletten). Zorg voor een goed isolatiematje om op te liggen die niet te breed is, want het zal wat krap zijn. Een aluminiumfolie matje of een stuk karton onder je isolatiematje maakt het helemaal comfortabel.
- Kampvuurdeken
Lekker voor 's avonds bij het kampvuur om op te zitten (we hebben geen banken of stoelen), of om over je heen te slaan. Ook lekker voor over je slaapzak 's nachts.
- Powerbank
Mobiele telefoons zijn toegestaan tijdens het kamp en met sommige activiteiten zelfs handig. Er zijn alleen geen stopcontacten op de eilanden waar we overnachten en de accu van het volgschip is niet sterk genoeg voor alle deelnemers. Neem dus vooral een opgeladen powerbank mee. Let er vooral op dat je telefoon niet in het water kan vallen rondom het varen. Dit gebeurt jammer genoeg nogal vaak. Wij zijn hiervoor niet verantwoordelijk.

*Wat neem ik niet mee?*
- Laarzen
We raden af om plastic regenlaarzen aan te hebben tijdens het zeilen. Dit is niet veilig als je in het water valt.

We zien jullie volgende week vrijdag!

Groetjes van de $KampNaam staf
"@

#endregion Build Email

#region Output

$extraWarnings = @()
if (-not $weatherFetched) {
    $extraWarnings += 'weersvoorspelling niet opgehaald — pas de body handmatig aan!'
}

Write-HitMailOutput `
    -BccString             $bccString `
    -Subject               $subject `
    -Body                  $body `
    -MailDeadlineFormatted $mailDeadlineFormatted `
    -ExtraWarnings         $extraWarnings

#endregion Output
