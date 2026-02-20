<#
    .SYNOPSIS
    Toont statistieken over HIT-aanmeldingen uit een CSV-bestand.

    .DESCRIPTION
    Leest een HIT-aanmeldingen CSV-bestand (puntkomma-gescheiden, zoals geëxporteerd uit
    het Scouting-aanmeldingssysteem) en toont een samenvatting met:
    - Geslachtsverdeling (dames/heren)
    - Subgroep-analyse (groepsgrootte-verdeling)
    - Leeftijdsverdeling op de startdatum van het kamp
    - Jarigen-check tijdens het kamp
    - Provincieverdeling op basis van woonadres (postcode via PDOK Locatieserver API)

    De kampdatums worden automatisch berekend aan de hand van Pasen:
    - Startdatum = Goede Vrijdag (Paaszondag − 2 dagen)
    - Einddatum = Tweede Paasdag (Paaszondag + 1 dag)

    De gebruiker geeft alleen het jaar op (standaard: huidig jaar) en bevestigt
    interactief. Het script berekent de juiste Paasdatums via het Computus-algoritme.

    .PARAMETER Year
    Het jaar van het HIT-kamp. Standaard het huidige jaar.
    De gebruiker krijgt een interactieve prompt om het jaar te bevestigen of te wijzigen.

    .EXAMPLE
    .\Get-HitStatistic.ps1

    Start het script met het huidige jaar. De gebruiker bevestigt het jaar met ENTER
    en kiest interactief een CSV-bestand.

    .EXAMPLE
    .\Get-HitStatistic.ps1 -Year 2025

    Toont statistieken voor het HIT-kamp op Goede Vrijdag t/m 2e Paasdag 2025.

    .EXAMPLE
    .\Get-HitStatistic.ps1 -Verbose

    Zoals hierboven, maar met uitgebreide voortgangsberichten.

    .OUTPUTS
    System.String
    Een tekstrapport met statistieken over de aanmeldingen.

    .NOTES
    Vereist een internetverbinding voor postcode-naar-provincie lookup via PDOK.
    Bij geen verbinding wordt een fallback-tabel op basis van de eerste twee postcodecijfers gebruikt.
    CSV moet puntkomma-gescheiden zijn met kolommen: Lid geslacht, Subgroepnaam,
    Lid geboortedatum (dd-MM-yyyy), Lid postcode.

    Paasberekening: Anonymous Gregorian-algoritme (Meeus/Jones/Butcher).
#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = 'Jaar van het HIT-kamp (standaard: huidig jaar)')]
    [ValidateRange(2000, 2099)]
    [int]$Year = (Get-Date).Year
)

#region Helper Functions

function Get-EasterSunday {
    <#
        .SYNOPSIS
        Berekent de datum van Eerste Paasdag (Easter Sunday) voor een gegeven jaar.

        .DESCRIPTION
        Gebruikt het Anonymous Gregorian-algoritme (Meeus/Jones/Butcher) om de datum
        van Eerste Paasdag te berekenen. Geldig voor jaren 1583–4099.

        .PARAMETER Year
        Het jaar waarvoor de Paasdatum berekend moet worden.

        .OUTPUTS
        System.DateTime — De datum van Eerste Paasdag.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1583, 4099)]
        [int]$Year
    )

    $a = $Year % 19
    $b = [Math]::Floor($Year / 100)
    $c = $Year % 100
    $d = [Math]::Floor($b / 4)
    $e = $b % 4
    $f = [Math]::Floor(($b + 8) / 25)
    $g = [Math]::Floor(($b - $f + 1) / 3)
    $h = (19 * $a + $b - $d - $g + 15) % 30
    $i = [Math]::Floor($c / 4)
    $k = $c % 4
    $l = (32 + 2 * $e + 2 * $i - $h - $k) % 7
    $m = [Math]::Floor(($a + 11 * $h + 22 * $l) / 451)

    $month = [Math]::Floor(($h + $l - 7 * $m + 114) / 31)
    $day = (($h + $l - 7 * $m + 114) % 31) + 1

    return [datetime]::new($Year, $month, $day)
}

function Get-ProvinceFromPostcode {
    <#
        .SYNOPSIS
        Vertaalt een Nederlandse postcode naar een provincienaam via de PDOK Locatieserver API.

        .DESCRIPTION
        Roept de PDOK Locatieserver v3_1 API aan om de provincie op te halen die hoort bij
        de opgegeven postcode. Bij falen (geen internet, timeout, onbekende postcode) wordt
        een fallback-tabel op basis van het eerste postcodecijfer gebruikt.

        .PARAMETER Postcode
        Een Nederlandse postcode (4 cijfers + 2 letters, bijv. '2353VL').

        .OUTPUTS
        System.String — De provincienaam, of 'Onbekend' als de lookup niet lukt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Postcode
    )

    # Normaliseer postcode: verwijder spaties
    $normalizedPostcode = $Postcode -replace '\s', ''

    # Controleer of het een Nederlands postcodeformaat is (4 cijfers + 2 letters)
    if ($normalizedPostcode -notmatch '^\d{4}[A-Za-z]{2}$') {
        Write-Warning "Postcode '$normalizedPostcode' is geen geldig Nederlands formaat. Overgeslagen voor provincie-bepaling."
        return $null
    }

    # Probeer PDOK API met volledige postcode (PC6)
    try {
        $encodedPostcode = [System.Uri]::EscapeDataString($normalizedPostcode)
        $uri = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q=$encodedPostcode&fq=type:postcode&rows=1"
        Write-Verbose "PDOK API aanroep voor postcode: $normalizedPostcode"

        $response = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 10 -ErrorAction Stop

        if ($response.response.numFound -gt 0) {
            $provinceRaw = $response.response.docs[0].provincienaam
            if (-not [string]::IsNullOrWhiteSpace($provinceRaw)) {
                Write-Verbose "PDOK resultaat voor ${normalizedPostcode}: $provinceRaw"
                return $provinceRaw
            }
        }

        # Probeer met alleen de 4-cijferige postcode (PC4) als fallback
        $pc4 = $normalizedPostcode.Substring(0, 4)
        $encodedPc4 = [System.Uri]::EscapeDataString($pc4)
        $uri4 = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q=$encodedPc4&fq=type:postcode&rows=1"
        Write-Verbose "PDOK API PC4-fallback voor postcode: $pc4"

        $response4 = Invoke-RestMethod -Uri $uri4 -Method Get -TimeoutSec 10 -ErrorAction Stop

        if ($response4.response.numFound -gt 0) {
            $provinceRaw4 = $response4.response.docs[0].provincienaam
            if (-not [string]::IsNullOrWhiteSpace($provinceRaw4)) {
                Write-Verbose "PDOK PC4-resultaat voor ${pc4}: $provinceRaw4"
                return $provinceRaw4
            }
        }

        Write-Warning "PDOK API gaf geen resultaat voor postcode '$normalizedPostcode'. Fallback wordt gebruikt."
    }
    catch {
        Write-Warning "PDOK API-aanroep mislukt voor postcode '$normalizedPostcode': $($_.Exception.Message). Fallback wordt gebruikt."
    }

    # Fallback: eerste twee cijfers van postcode -> provincie
    $prefix2 = $normalizedPostcode.Substring(0, 2)
    $fallbackMap = @{
        '10' = 'Noord-Holland'; '11' = 'Noord-Holland'; '12' = 'Noord-Holland'; '13' = 'Noord-Holland'
        '14' = 'Noord-Holland'; '15' = 'Noord-Holland'; '16' = 'Noord-Holland'; '17' = 'Noord-Holland'
        '18' = 'Noord-Holland'; '19' = 'Noord-Holland'; '20' = 'Noord-Holland'; '21' = 'Noord-Holland'
        '22' = 'Zuid-Holland'; '23' = 'Zuid-Holland'; '24' = 'Zuid-Holland'; '25' = 'Zuid-Holland'
        '26' = 'Zuid-Holland'; '27' = 'Zuid-Holland'; '28' = 'Zuid-Holland'; '29' = 'Zuid-Holland'
        '30' = 'Zuid-Holland'; '31' = 'Zuid-Holland'; '32' = 'Zuid-Holland'; '33' = 'Zuid-Holland'
        '34' = 'Utrecht'; '35' = 'Utrecht'; '36' = 'Utrecht'; '37' = 'Utrecht'
        '38' = 'Utrecht'; '39' = 'Utrecht'
        '40' = 'Zeeland'; '41' = 'Zeeland'; '42' = 'Zeeland'; '43' = 'Zeeland'; '44' = 'Zeeland'
        '45' = 'Noord-Brabant'; '46' = 'Noord-Brabant'; '47' = 'Noord-Brabant'
        '48' = 'Noord-Brabant'; '49' = 'Noord-Brabant'
        '50' = 'Noord-Brabant'; '51' = 'Noord-Brabant'; '52' = 'Noord-Brabant'
        '53' = 'Noord-Brabant'; '54' = 'Noord-Brabant'; '55' = 'Noord-Brabant'; '56' = 'Noord-Brabant'
        '57' = 'Limburg'; '58' = 'Limburg'; '59' = 'Limburg'
        '60' = 'Limburg'; '61' = 'Limburg'; '62' = 'Limburg'; '63' = 'Limburg'; '64' = 'Limburg'
        '65' = 'Gelderland'; '66' = 'Gelderland'; '67' = 'Gelderland'
        '68' = 'Gelderland'; '69' = 'Gelderland'
        '70' = 'Overijssel'; '71' = 'Overijssel'; '72' = 'Overijssel'
        '73' = 'Overijssel'; '74' = 'Overijssel'; '75' = 'Overijssel'; '76' = 'Overijssel'
        '77' = 'Drenthe'; '78' = 'Drenthe'; '79' = 'Drenthe'
        '80' = 'Overijssel'; '81' = 'Overijssel'; '82' = 'Flevoland'; '83' = 'Drenthe'
        '84' = 'Frysl\u00e2n'; '85' = 'Frysl\u00e2n'; '86' = 'Frysl\u00e2n'
        '87' = 'Frysl\u00e2n'; '88' = 'Frysl\u00e2n'; '89' = 'Frysl\u00e2n'
        '90' = 'Groningen'; '91' = 'Frysl\u00e2n'; '92' = 'Frysl\u00e2n'
        '93' = 'Drenthe'; '94' = 'Drenthe'
        '95' = 'Groningen'; '96' = 'Groningen'; '97' = 'Groningen'
        '98' = 'Groningen'; '99' = 'Groningen'
    }

    $fallbackProvince = $fallbackMap[$prefix2]
    if ($fallbackProvince) {
        Write-Warning "Fallback-provincie voor postcode '$normalizedPostcode': $fallbackProvince (gebaseerd op eerste twee cijfers, kan onnauwkeurig zijn)"
        return $fallbackProvince
    }

    return 'Onbekend'
}

function Get-AgeAtDate {
    <#
        .SYNOPSIS
        Berekent de leeftijd in jaren op een bepaalde datum.

        .PARAMETER BirthDate
        De geboortedatum.

        .PARAMETER ReferenceDate
        De datum waarop de leeftijd berekend wordt.

        .OUTPUTS
        System.Int32 — De leeftijd in hele jaren.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$BirthDate,

        [Parameter(Mandatory = $true)]
        [datetime]$ReferenceDate
    )

    $age = $ReferenceDate.Year - $BirthDate.Year
    if ($ReferenceDate.Date -lt $BirthDate.Date.AddYears($age)) {
        $age--
    }
    return $age
}

function Get-DutchGroupSizeLabel {
    <#
        .SYNOPSIS
        Geeft de Nederlandse beschrijving van een (sub)groepsgrootte.

        .PARAMETER Size
        Het aantal personen in de groep.

        .OUTPUTS
        System.String — Nederlandse beschrijving (bijv. 'alleen opgegeven', 'met z'n drieën').
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 99)]
        [int]$Size
    )

    $labels = @{
        1  = 'alleen opgegeven'
        2  = "met z'n twee" + [char]0x00EB + "n"   # tweeën
        3  = "met z'n drie" + [char]0x00EB + "n"    # drieën
        4  = "met z'n vieren"
        5  = "met z'n vijven"
        6  = "met z'n zessen"
        7  = "met z'n zevenen"
        8  = "met z'n achten"
        9  = "met z'n negenen"
        10 = "met z'n tienen"
    }

    if ($labels.ContainsKey($Size)) {
        return $labels[$Size]
    }
    return "met z'n ${Size}en"
}

function Get-DutchMonthName {
    <#
        .SYNOPSIS
        Geeft de Nederlandse maandnaam voor een maandnummer.

        .PARAMETER Month
        Het maandnummer (1-12).

        .OUTPUTS
        System.String — De Nederlandse maandnaam.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 12)]
        [int]$Month
    )

    $months = @(
        'januari', 'februari', 'maart', 'april', 'mei', 'juni',
        'juli', 'augustus', 'september', 'oktober', 'november', 'december'
    )
    return $months[$Month - 1]
}

#endregion Helper Functions

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

#region CSV Selection

$scriptDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$csvFiles = @(Get-ChildItem -Path $scriptDirectory -Filter '*.csv' -File)

if ($csvFiles.Count -eq 0) {
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.IO.FileNotFoundException]::new("Geen CSV-bestanden gevonden in '$scriptDirectory'."),
        'NoCsvFilesFound',
        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
        $scriptDirectory
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

$selectedCsvPath = $null

if ($csvFiles.Count -eq 1) {
    $selectedCsvPath = $csvFiles[0].FullName
    Write-Verbose "Eén CSV-bestand gevonden, automatisch geselecteerd: $($csvFiles[0].Name)"
}
else {
    Write-Host ''
    Write-Host 'Beschikbare CSV-bestanden:' -ForegroundColor Cyan
    Write-Host ''

    for ($i = 0; $i -lt $csvFiles.Count; $i++) {
        Write-Host "  [$($i + 1)] $($csvFiles[$i].Name)" -ForegroundColor White
    }

    Write-Host ''

    $validChoice = $false
    while (-not $validChoice) {
        $choiceInput = Read-Host "Kies een bestand (1-$($csvFiles.Count))"

        $choiceNumber = 0
        if ([int]::TryParse($choiceInput, [ref]$choiceNumber)) {
            if ($choiceNumber -ge 1 -and $choiceNumber -le $csvFiles.Count) {
                $selectedCsvPath = $csvFiles[$choiceNumber - 1].FullName
                $validChoice = $true
            }
        }

        if (-not $validChoice) {
            Write-Host "Ongeldige keuze. Voer een nummer in van 1 t/m $($csvFiles.Count)." -ForegroundColor Yellow
        }
    }

    Write-Verbose "Geselecteerd CSV-bestand: $selectedCsvPath"
}

#endregion CSV Selection

#region CSV Import & Validation

Write-Verbose "CSV importeren: $selectedCsvPath"

try {
    $deelnemers = @(Import-Csv -Path $selectedCsvPath -Delimiter ';' -Encoding UTF8 -ErrorAction Stop)
}
catch {
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.IO.IOException]::new("Kan CSV-bestand niet lezen: $($_.Exception.Message)"),
        'CsvImportFailed',
        [System.Management.Automation.ErrorCategory]::ReadError,
        $selectedCsvPath
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

if ($deelnemers.Count -eq 0) {
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.InvalidOperationException]::new('Het CSV-bestand bevat geen gegevensrijen.'),
        'CsvEmpty',
        [System.Management.Automation.ErrorCategory]::InvalidData,
        $selectedCsvPath
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

# Controleer vereiste kolommen
$requiredColumns = @('Lid geslacht', 'Subgroepnaam', 'Lid geboortedatum', 'Lid postcode')
$actualColumns = $deelnemers[0].PSObject.Properties.Name
$missingColumns = @($requiredColumns | Where-Object { $_ -notin $actualColumns })

if ($missingColumns.Count -gt 0) {
    $missingList = $missingColumns -join ', '
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.InvalidOperationException]::new("Vereiste kolommen ontbreken in het CSV-bestand: $missingList"),
        'CsvMissingColumns',
        [System.Management.Automation.ErrorCategory]::InvalidData,
        $selectedCsvPath
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

Write-Verbose "$($deelnemers.Count) deelnemers geladen uit CSV."

#endregion CSV Import & Validation

#region Statistics Computation

$outputLines = [System.Collections.Generic.List[string]]::new()
$outputLines.Add('Statistieken over aanmeldingen:')

# --- Geslacht ---
$genderGroups = $deelnemers | Group-Object -Property 'Lid geslacht'
$countDames = ($genderGroups | Where-Object { $_.Name -eq 'Vrouw' }).Count
$countHeren = ($genderGroups | Where-Object { $_.Name -eq 'Man' }).Count

$outputLines.Add("- $countDames dames, $countHeren heren")
$outputLines.Add('')

# --- Subgroepen ---
$subgroepGroups = $deelnemers | Group-Object -Property 'Subgroepnaam'
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
    $dateString = $deelnemer.'Lid geboortedatum'
    try {
        $birthDate = [datetime]::ParseExact($dateString, 'dd-MM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        $voornaam = $deelnemer.'Lid voornaam'
        $achternaam = $deelnemer.'Lid achternaam'
        Write-Warning "Kan geboortedatum '$dateString' niet parsen voor $voornaam $achternaam. Overgeslagen."
        continue
    }

    $age = Get-AgeAtDate -BirthDate $birthDate -ReferenceDate $StartDate
    $voornaam = $deelnemer.'Lid voornaam'
    $tussenvoegsel = $deelnemer.'Lid tussenvoegsel'
    $achternaam = $deelnemer.'Lid achternaam'
    $volledigeNaam = if ([string]::IsNullOrWhiteSpace($tussenvoegsel)) { "$voornaam $achternaam" } else { "$voornaam $tussenvoegsel $achternaam" }

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

# --- Provincie via postcode (PDOK) ---
Write-Verbose 'Provincies ophalen via PDOK Locatieserver...'

# Cache: postcode -> provincie
$postcodeCache = @{}
$deelnemerProvinces = [System.Collections.Generic.List[string]]::new()

foreach ($deelnemer in $deelnemers) {
    $postcode = $deelnemer.'Lid postcode' -replace '\s', ''

    if (-not $postcodeCache.ContainsKey($postcode)) {
        $province = Get-ProvinceFromPostcode -Postcode $postcode
        $postcodeCache[$postcode] = $province
    }

    # Sla deelnemers met niet-Nederlandse postcodes (null) over in de provincietelling
    if ($null -ne $postcodeCache[$postcode]) {
        $deelnemerProvinces.Add($postcodeCache[$postcode])
    }
}

$provincieDistribution = $deelnemerProvinces |
    Group-Object |
    Sort-Object -Property @{Expression = 'Count'; Descending = $true}, @{Expression = 'Name'; Ascending = $true}

foreach ($provGroup in $provincieDistribution) {
    $outputLines.Add("- $($provGroup.Count) x uit $($provGroup.Name)")
}

#endregion Statistics Computation

#region Output

Write-Output ($outputLines -join [System.Environment]::NewLine)

#endregion Output
