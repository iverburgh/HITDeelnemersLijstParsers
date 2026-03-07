<#
    .SYNOPSIS
    Genereert een kopieerklare herinneringsmail voor de merchandise-deadline.

    .DESCRIPTION
    Leest de e-mailadressen van deelnemers uit het HIT-aanmeldingen Excel-bestand en
    genereert een korte herinneringsmail over de aankomende merchandise-besteldatum,
    klaar om te kopiëren naar Gmail.

    De output bestaat uit drie afzonderlijke secties:
    - BCC:       alle e-mailadressen van deelnemers uit het Excel-bestand
    - Onderwerp: "[KampNaam] - Reminder merchandise bestelling"
    - Body:      een korte herinnering met de uiterste besteldatum

    De merchandise-deadline is de donderdagavond om 22:00, twee weken vóór de start van het kamp.
    Ga terug vanuit (campStart - 14 dagen) naar de laatste donderdag op of vóór die datum.
    Deze mail verstuur je uiterlijk twee dagen vóór de deadline (de dinsdag ervoor).

    Bovenaan de output verschijnt een waarschuwing met de uiterste verzenddatum
    (twee dagen vóór de merchandise-deadline).

    Het script zoekt automatisch naar een bestand met het patroon "*-alles.xlsx" in de
    scriptmap. Als er meerdere zijn, verschijnt een keuzemenu. Als er geen zijn, wordt
    gezocht op "*.xlsx". Als er dan ook geen is, volgt een foutmelding.

    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is.

    .PARAMETER Year
    Het jaar van het HIT-kamp. Wordt gebruikt voor de paasdatumberekening.
    Standaard: het huidige jaar.

    .PARAMETER TikkieLink
    De Tikkie-betaallink voor merchandise.
    Standaard: 'https://tikkielink.nl/'

    .PARAMETER GoogleFormLink
    De link naar het Google-formulier voor de merchandise bestelling.
    Standaard: 'https://google.form.nl/'

    .PARAMETER EmailKolom
    De naam van de kolom in het Excel-bestand die de e-mailadressen van de deelnemers bevat.
    Standaard: 'Mailadres'

    .EXAMPLE
    .\Mail02-1_Dag_voor_Merchandise_Deadline.ps1

    Genereert de herinneringsmail voor het huidige jaar.
    De kampnaam wordt automatisch uitgelezen uit het Excel-bestand.

    .EXAMPLE
    .\Mail02-1_Dag_voor_Merchandise_Deadline.ps1 `
        -TikkieLink "https://tikkie.me/pay/abc123" `
        -GoogleFormLink "https://forms.gle/xyz789" `
        -Year 2027

    Genereert de herinneringsmail met alle parameters expliciet ingevuld.

    .EXAMPLE
    .\Mail02-1_Dag_voor_Merchandise_Deadline.ps1 -Verbose

    Toont uitgebreide voortgangsberichten tijdens de verwerking.

    .OUTPUTS
    System.String
    Drie secties (BCC, Onderwerp, Body) als tekstuitvoer op de console.

    .NOTES
    Vereist: PowerShell 5.1+, ImportExcel-module.
    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is
    (via Install-Module -Scope CurrentUser). Hiervoor is een internetverbinding nodig.
    Paasdatum-berekening: Anonymous Gregorian-algoritme (Meeus/Jones/Butcher).
#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = 'Jaar van het HIT-kamp (standaard: huidig jaar).')]
    [ValidateRange(2000, 2099)]
    [int]$Year = (Get-Date).Year,

    [Parameter(HelpMessage = 'Tikkie-betaallink voor merchandise.')]
    [string]$TikkieLink = 'https://tikkielink.nl/',

    [Parameter(HelpMessage = 'Link naar het Google-formulier voor de merchandise bestelling.')]
    [string]$GoogleFormLink = 'https://google.form.nl/',

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

$easterSunday = Get-EasterSunday -Year $Year
$campStart    = $easterSunday.AddDays(-2)   # Goede Vrijdag

Write-Verbose ("Kamp {0}: Goede Vrijdag {1:dd-MM-yyyy}" -f $Year, $campStart)

#endregion Camp Date Calculation

#region Deadline Calculation

# Merchandise-deadline: donderdagavond 22:00, twee weken vóór de start van het kamp.
# Ga terug vanuit (campStart - 14 dagen) naar de laatste donderdag op of vóór die datum.
$rawDeadline    = $campStart.AddDays(-14)
$dayOfWeekValue = [int]$rawDeadline.DayOfWeek      # Sunday=0 ... Thursday=4 ... Saturday=6
$daysToSubtract = ($dayOfWeekValue - 4 + 7) % 7    # stappen terug naar de dichtstbijzijnde of huidige donderdag
$deadlineDate   = $rawDeadline.AddDays(-$daysToSubtract)

$deadlineDateTime = [datetime]::new(
    $deadlineDate.Year, $deadlineDate.Month, $deadlineDate.Day, 22, 0, 0
)

$deadlineDayName   = Get-DutchDayName -DayOfWeek ([int]$deadlineDateTime.DayOfWeek)
$deadlineMonthName = Get-DutchMonthName -Month $deadlineDateTime.Month
$deadlineFormatted = "$deadlineDayName $($deadlineDateTime.Day) $deadlineMonthName"
$deadlineTime      = $deadlineDateTime.ToString('HH:mm')

Write-Verbose "Merchandise-deadline: $deadlineFormatted om $deadlineTime"

# Uiterste verzenddatum e-mail: 2 dagen vóór de merchandise-deadline (de dinsdag ervoor)
$mailDeadlineDate      = $deadlineDateTime.AddDays(-2)
$mailDeadlineDayName   = Get-DutchDayName -DayOfWeek ([int]$mailDeadlineDate.DayOfWeek)
$mailDeadlineMonthName = Get-DutchMonthName -Month $mailDeadlineDate.Month
$mailDeadlineFormatted = "$mailDeadlineDayName $($mailDeadlineDate.Day) $mailDeadlineMonthName $($mailDeadlineDate.Year)"

Write-Verbose "Uiterste verzenddatum e-mail: $mailDeadlineFormatted"

#endregion Deadline Calculation

#region Import Data

Write-Verbose 'Inlezen van Excel-bestand...'
$allRows   = Import-Excel -Path $resolvedInputPath -ErrorAction Stop
$totalRows = ($allRows | Measure-Object).Count
Write-Verbose "Ingelezen: $totalRows rij(en)."

if ($totalRows -eq 0) {
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.InvalidOperationException]::new('Het Excel-bestand bevat geen gegevensrijen.'),
        'ExcelEmpty',
        [System.Management.Automation.ErrorCategory]::InvalidData,
        $resolvedInputPath
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

$actualColumns = $allRows[0].PSObject.Properties.Name

# Kampnaam afleiden uit de 'Kamp'-kolom
if ('Kamp' -in $actualColumns) {
    $KampNaam = ($allRows | Select-Object -First 1).Kamp
    Write-Verbose "Kampnaam: $KampNaam"
}
else {
    Write-Warning "Kolom 'Kamp' niet gevonden in het Excel-bestand. Kampnaam ingesteld op 'HIT-kamp'."
    $KampNaam = 'HIT-kamp'
}

#endregion Import Data

#region Build BCC

if ($EmailKolom -notin $actualColumns) {
    Write-Warning (
        "Kolom '$EmailKolom' niet gevonden in het Excel-bestand. BCC-lijst is leeg. " +
        "Beschikbare kolommen: $($actualColumns -join ', ')"
    )
    $bccAddresses = @()
}
else {
    $bccAddresses = @(
        $allRows |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_."$EmailKolom") } |
            Select-Object -ExpandProperty $EmailKolom |
            Sort-Object -Unique
    )
}

$bccString = $bccAddresses -join ', '
Write-Verbose "$($bccAddresses.Count) e-mailadres(sen) gevonden voor BCC."

#endregion Build BCC

#region Build Email

$subject = "$KampNaam - Reminder merchandise bestelling"

$body = @"
Hallo,

We hebben de afgelopen dagen al heel wat gegevens van jullie ontvangen met daarin bestellingen voor de $KampNaam merchandise. Super moai!

Bij deze een reminder voor de deadline van het bestellen van de merchandise: dit kan tot uiterlijk $deadlineFormatted om $deadlineTime.

Bestellen kan via het formulier op de volgende link: $GoogleFormLink

Betalen kan via deze link: $TikkieLink

Tot binnenkort!

Ivo Verburgh
$KampNaam
HIT Heerenveen
"@

#endregion Build Email

#region Output

$separator = '=' * 60

Write-Host ''
Write-Host ('!' * 60) -ForegroundColor Yellow
Write-Host "  Verstuur deze mail uiterlijk op $mailDeadlineFormatted" -ForegroundColor Yellow -BackgroundColor DarkRed
Write-Host ('!' * 60) -ForegroundColor Yellow

Write-Host ''
Write-Host $separator -ForegroundColor Cyan
Write-Host '  BCC' -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host $bccString

Write-Host ''
Write-Host $separator -ForegroundColor Cyan
Write-Host '  ONDERWERP' -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host $subject

Write-Host ''
Write-Host $separator -ForegroundColor Cyan
Write-Host '  EMAIL BODY' -ForegroundColor Cyan
Write-Host $separator -ForegroundColor Cyan
Write-Host $body

#endregion Output
