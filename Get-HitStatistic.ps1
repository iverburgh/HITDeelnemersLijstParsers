<#
    .SYNOPSIS
    Toont statistieken over HIT-aanmeldingen en vergelijkt met het vorig jaar.

    .DESCRIPTION
    Leest twee HIT-aanmeldingenbestanden (huidig jaar en vorig jaar) en toont een samenvatting met:
    - Geslachtsverdeling (dames/heren)
    - Subgroep-analyse (groepsgrootte-verdeling)
    - Leeftijdsverdeling op de startdatum van het kamp
    - Jarigen-check tijdens het kamp
    - Terugkerende deelnemers (ook aanwezig in de vorigjaarlijst)

    De kampdatums worden automatisch berekend aan de hand van Pasen:
    - Startdatum = Goede Vrijdag (Paaszondag − 2 dagen)
    - Einddatum = Tweede Paasdag (Paaszondag + 1 dag)

    De gebruiker geeft alleen het jaar op (standaard: huidig jaar) en bevestigt
    interactief. Het script berekent de juiste Paasdatums via het Computus-algoritme.
    Beide bestanden worden geselecteerd via een interactief keuzemenu.

    .PARAMETER Year
    Het jaar van het HIT-kamp. Standaard het huidige jaar.
    De gebruiker krijgt een interactieve prompt om het jaar te bevestigen of te wijzigen.

    .EXAMPLE
    .\Get-HitStatistic.ps1

    Start het script met het huidige jaar. De gebruiker bevestigt het jaar met ENTER.
    Beide bestanden (huidig en vorig jaar) worden geselecteerd via een keuzemenu.

    .EXAMPLE
    .\Get-HitStatistic.ps1 -Year 2025

    Toont statistieken voor het HIT-kamp op Goede Vrijdag t/m 2e Paasdag 2025.

    .EXAMPLE
    .\Get-HitStatistic.ps1 -Verbose

    Zoals hierboven, maar met uitgebreide voortgangsberichten.

    .OUTPUTS
    System.String
    Een tekstrapport met statistieken over de aanmeldingen, inclusief terugkerende deelnemers.

    .NOTES
    Vereist: PowerShell 5.1+, ImportExcel-module.
    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is
    (via Install-Module -Scope CurrentUser). Hiervoor is een internetverbinding nodig.
    Huidigjaar-bestand moet de kolommen bevatten: Gender, Subgroep, Geboortedatum, Voornaam, Achternaam.
    Vorigjaar-bestand ondersteunt Excel (.xlsx) én raw CSV (puntkommagescheiden, kolommen 'Lid voornaam' e.d.).

    Paasberekening: Anonymous Gregorian-algoritme (Meeus/Jones/Butcher).
#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = 'Jaar van het HIT-kamp (standaard: huidig jaar)')]
    [ValidateRange(2000, 2099)]
    [int]$Year = (Get-Date).Year
)

#region Import Shared Module

$_moduleRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
Import-Module -Name (Join-Path -Path $_moduleRoot -ChildPath 'HitHelpers.psm1') -Force
Remove-Variable -Name _moduleRoot

#endregion Import Shared Module

#region Year Confirmation & Date Computation

Write-Host ''
$yearInput = Read-Host "Jaar van het HIT-kamp [$Year]"

if (-not [string]::IsNullOrWhiteSpace($yearInput)) {
    $parsedYear = 0
    if ([int]::TryParse($yearInput, [ref]$parsedYear) -and $parsedYear -ge 2000 -and $parsedYear -le 2099) {
        $Year = $parsedYear
    }
    else {
        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
            [System.ArgumentException]::new("Ongeldig jaar: '$yearInput'. Voer een jaar in tussen 2000 en 2099."),
            'InvalidYear',
            [System.Management.Automation.ErrorCategory]::InvalidArgument,
            $null
        )
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }
}

$easterSunday = Get-EasterSunday -Year $Year
$StartDate = $easterSunday.AddDays(-2)   # Goede Vrijdag
$EndDate = $easterSunday.AddDays(1)       # Tweede Paasdag

$startMonthName = Get-DutchMonthName -Month $StartDate.Month
$endMonthName = Get-DutchMonthName -Month $EndDate.Month

Write-Host ''
Write-Host "Goede Vrijdag:  $($StartDate.Day) $startMonthName $($StartDate.Year)" -ForegroundColor Green
Write-Host "2e Paasdag:     $($EndDate.Day) $endMonthName $($EndDate.Year)" -ForegroundColor Green

Write-Verbose "Paaszondag: $($easterSunday.ToString('yyyy-MM-dd'))"
Write-Verbose "Kampdatums: $($StartDate.ToString('yyyy-MM-dd')) t/m $($EndDate.ToString('yyyy-MM-dd'))"

#endregion Year Confirmation & Date Computation

#region Bestandsselectie

Assert-HitImportExcel

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }

Write-Host ''
$selectedExcelPath = Select-HitFilePath -Prompt 'Kies huidigjaar deelnemerslijst' -ScriptDir $scriptDirectory
$vorigJaarPath     = Select-HitFilePath -Prompt 'Kies vorigjaar deelnemerslijst (ter vergelijking)' -ScriptDir $scriptDirectory -ExcludePaths @($selectedExcelPath)

#endregion Bestandsselectie

#region Excel Import & Validation

Write-Verbose "Excel importeren: $selectedExcelPath"

$deelnemers = @(Import-Excel -Path $selectedExcelPath -ErrorAction Stop)

if ($deelnemers.Count -eq 0) {
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.InvalidOperationException]::new('Het Excel-bestand bevat geen gegevensrijen.'),
        'ExcelEmpty',
        [System.Management.Automation.ErrorCategory]::InvalidData,
        $selectedExcelPath
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

# Controleer vereiste kolommen
$requiredColumns = @('Gender', 'Subgroep', 'Geboortedatum', 'Voornaam', 'Achternaam')
$actualColumns = $deelnemers[0].PSObject.Properties.Name
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $actualColumns })

if ($missingColumns.Count -gt 0) {
    $missingList = $missingColumns -join ', '
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.InvalidOperationException]::new("Vereiste kolommen ontbreken in het Excel-bestand: $missingList"),
        'ExcelMissingColumns',
        [System.Management.Automation.ErrorCategory]::InvalidData,
        $selectedExcelPath
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

Write-Verbose "$($deelnemers.Count) deelnemers geladen uit Excel."

#endregion Excel Import & Validation

#region Statistics Computation

$kampNaam = ($deelnemers | Select-Object -First 1).Kamp
$outputLines = [System.Collections.Generic.List[string]]::new()
$outputLines.Add("Statistieken over aanmeldingen $kampNaam $Year :")

# --- Geslacht ---
$genderGroups = $deelnemers | Group-Object -Property 'Gender'
$countDames = ($genderGroups | Where-Object { $_.Name -eq 'vrouw' }).Count
$countHeren = ($genderGroups | Where-Object { $_.Name -eq 'man' }).Count

$outputLines.Add("- $countDames dames, $countHeren heren")
$outputLines.Add('')

# --- Subgroepen ---
$subgroepGroups = $deelnemers | Group-Object -Property 'Subgroep'
$totalSubgroepen = $subgroepGroups.Count

$outputLines.Add("- $totalSubgroepen subgroepjes")

# Groepeer subgroepen op grootte
$sizeDistribution = $subgroepGroups |
    Group-Object -Property Count |
    Sort-Object -Property { [int]$_.Name } |
    ForEach-Object {
        [PSCustomObject]@{
            GroupSize  = [int]$_.Name
            Frequency  = $_.Count
            SizeLabel  = Get-DutchGroupSizeLabel -Size ([int]$_.Name)
        }
    }

foreach ($sizeGroup in $sizeDistribution) {
    $outputLines.Add("- $($sizeGroup.Frequency) x $($sizeGroup.SizeLabel)")
}

$outputLines.Add('')

# --- Leeftijdsverdeling ---
$birthDates = @{}
$ages = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($deelnemer in $deelnemers) {
    $birthDateRaw = $deelnemer.Geboortedatum
    $birthDate    = ConvertFrom-HitBirthDate -RawValue $birthDateRaw

    if ($null -eq $birthDate) {
        Write-Warning "Kan geboortedatum '$birthDateRaw' niet parsen voor $($deelnemer.Voornaam) $($deelnemer.Achternaam). Overgeslagen."
        continue
    }

    $age = Get-AgeAtDate -BirthDate $birthDate -ReferenceDate $StartDate
    $volledigeNaam = "$($deelnemer.Voornaam) $($deelnemer.Achternaam)"

    $ages.Add([PSCustomObject]@{
        Naam          = $volledigeNaam
        Geboortedatum = $birthDate
        Leeftijd      = $age
    })
}

$ageDistribution = $ages |
    Group-Object -Property Leeftijd |
    Sort-Object -Property { [int]$_.Name } -Descending |
    ForEach-Object {
        [PSCustomObject]@{
            Leeftijd  = [int]$_.Name
            Aantal    = $_.Count
        }
    }

# Sorteer op leeftijd (oplopend)
$ageDistribution = $ageDistribution | Sort-Object -Property Leeftijd

foreach ($ageGroup in $ageDistribution) {
    $outputLines.Add("- $($ageGroup.Aantal) x $($ageGroup.Leeftijd) jaar oud")
}

$outputLines.Add('')

# --- Jarigen tijdens het kamp ---
$jarigen = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($deelnemerAge in $ages) {
    $birthDate = $deelnemerAge.Geboortedatum

    # Controleer verjaardag in elk jaar dat het kamp beslaat
    $startYear = $StartDate.Year
    $endYear = $EndDate.Year

    for ($checkYear = $startYear; $checkYear -le $endYear; $checkYear++) {
        try {
            # Behandel schrikkeljaar: 29 feb op een niet-schrikkelj. → niet jarig
            $birthdayThisYear = [datetime]::new($checkYear, $birthDate.Month, $birthDate.Day)
        }
        catch {
            continue
        }

        if ($birthdayThisYear.Date -ge $StartDate.Date -and $birthdayThisYear.Date -le $EndDate.Date) {
            $newAge = Get-AgeAtDate -BirthDate $birthDate -ReferenceDate $birthdayThisYear
            $monthName = Get-DutchMonthName -Month $birthdayThisYear.Month

            $jarigen.Add([PSCustomObject]@{
                Naam     = $deelnemerAge.Naam
                Leeftijd = $newAge
                Dag      = $birthdayThisYear.Day
                Maand    = $monthName
            })
        }
    }
}

if ($jarigen.Count -eq 0) {
    $outputLines.Add('- geen jarige deelnemers tijdens hit')
}
else {
    foreach ($jarige in $jarigen) {
        $outputLines.Add("- $($jarige.Naam) wordt $($jarige.Leeftijd) op $($jarige.Dag) $($jarige.Maand)")
    }
}

$outputLines.Add('')

# --- Terugkerende deelnemers ---
Write-Verbose "Vorigjaar-bestand laden: $vorigJaarPath"

$vorigJaarRijen = Import-ParticipantFile -Path $vorigJaarPath
$vorigJaarDeelnemers = @(
    $vorigJaarRijen | ForEach-Object { ConvertTo-NormalizedParticipant -Row $_ } | Where-Object { $null -ne $_ }
)

Write-Verbose "$($vorigJaarDeelnemers.Count) deelnemers geladen uit vorigjaar-bestand."

$vorigJaarLookup = @{}
foreach ($vorigDeelnemer in $vorigJaarDeelnemers) {
    $vorigJaarLookup[$vorigDeelnemer.Sleutel] = $true
}

$huidigJaarNormalized = @(
    $deelnemers | ForEach-Object { ConvertTo-NormalizedParticipant -Row $_ } | Where-Object { $null -ne $_ }
)

$terugkerendeDeelnemers = @($huidigJaarNormalized | Where-Object { $vorigJaarLookup.ContainsKey($_.Sleutel) })

if ($terugkerendeDeelnemers.Count -eq 0) {
    $outputLines.Add('- geen deelnemers die vorig jaar ook deelnamen')
}
else {
    $outputLines.Add("- $($terugkerendeDeelnemers.Count) deelnemer(s) waren er vorig jaar ook bij:")
    foreach ($terugkerende in $terugkerendeDeelnemers) {
        $geboortedatumTekst = if ($null -ne $terugkerende.Geboortedatum) {
            $terugkerende.Geboortedatum.ToString('dd-MM-yyyy')
        }
        else {
            '(onbekend)'
        }
        $outputLines.Add("  - $($terugkerende.VolledigeNaam) ($geboortedatumTekst)")
    }
}


#endregion Statistics Computation

#region Output

Write-Output ($outputLines -join [System.Environment]::NewLine)

#endregion Output
