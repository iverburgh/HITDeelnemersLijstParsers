<#
    .SYNOPSIS
    Gedeelde helperfuncties voor de HIT-aanmeldingen scriptsuite.

    .DESCRIPTION
    Dit module bevat alle functies die door meerdere HIT-scripts worden gedeeld:
    - Paasdatumberekening (Get-EasterSunday)
    - Leeftijdberekening (Get-AgeAtDate)
    - Verjaardagscheck (Get-BirthdayDuringCamp)
    - Nederlandse labels (Get-DutchMonthName, Get-DutchDayName, Get-DutchGroupSizeLabel)
    - ImportExcel-installatie (Assert-HitImportExcel)
    - Excel-bestandsselectie met menu (Resolve-HitExcelPath)
    - Geboortedatum-parsing (ConvertFrom-HitBirthDate)
    - Schone outputnaam genereren (Get-HitOutputBaseName)
    - Concrete verjaardagsdatum tijdens kamp (Get-BirthdayDateDuringCamp)
    - Bestandsimport deelnemers (Import-ParticipantFile)
    - Deelnemerrij normaliseren (ConvertTo-NormalizedParticipant)
#>

function Get-EasterSunday {
    <#
        .SYNOPSIS
        Berekent de datum van Eerste Paasdag (Easter Sunday) voor een gegeven jaar.

        .DESCRIPTION
        Gebruikt het Anonymous Gregorian-algoritme (Meeus/Jones/Butcher) om de datum
        van Eerste Paasdag te berekenen. Geldig voor jaren 1583-4099.

        .PARAMETER Year
        Het jaar waarvoor de Paasdatum berekend moet worden.

        .OUTPUTS
        System.DateTime -- De datum van Eerste Paasdag.
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

function Get-AgeAtDate {
    <#
        .SYNOPSIS
        Berekent de leeftijd in jaren op een bepaalde datum.

        .PARAMETER BirthDate
        De geboortedatum.

        .PARAMETER ReferenceDate
        De datum waarop de leeftijd berekend wordt.

        .OUTPUTS
        System.Int32 -- De leeftijd in hele jaren.
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

function Get-BirthdayDuringCamp {
    <#
        .SYNOPSIS
        Controleert of een geboortedatum valt binnen een opgegeven datumrange.

        .DESCRIPTION
        Bepaalt of de verjaardag (dag en maand) van een persoon valt op een datum
        binnen de opgegeven campperiode, ongeacht het jaar.

        .PARAMETER BirthDate
        De geboortedatum van de deelnemer.

        .PARAMETER CampStart
        De eerste dag van het kamp (inclusief).

        .PARAMETER CampEnd
        De laatste dag van het kamp (inclusief).

        .OUTPUTS
        System.Boolean -- $true als de verjaardag tijdens het kamp valt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$BirthDate,

        [Parameter(Mandatory = $true)]
        [datetime]$CampStart,

        [Parameter(Mandatory = $true)]
        [datetime]$CampEnd
    )

    $current = $CampStart.Date
    while ($current -le $CampEnd.Date) {
        if ($current.Day -eq $BirthDate.Day -and $current.Month -eq $BirthDate.Month) {
            return $true
        }
        $current = $current.AddDays(1)
    }
    return $false
}

function Get-BirthdayDateDuringCamp {
    <#
        .SYNOPSIS
        Geeft de concrete datum terug waarop een persoon jarig is tijdens het kamp.

        .DESCRIPTION
        Bepaalt of de verjaardag (dag en maand) van een persoon valt binnen de opgegeven
        campperiode en retourneert de concrete datum. Houdt rekening met kampspans over
        meerdere jaren en met 29-februariverjaardagen. Retourneert $null als de verjaardag
        niet tijdens het kamp valt.

        .PARAMETER BirthDate
        De geboortedatum van de deelnemer.

        .PARAMETER CampStart
        De eerste dag van het kamp (inclusief).

        .PARAMETER CampEnd
        De laatste dag van het kamp (inclusief).

        .OUTPUTS
        System.DateTime of $null -- De concrete verjaardagsdatum tijdens het kamp, of $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$BirthDate,

        [Parameter(Mandatory = $true)]
        [datetime]$CampStart,

        [Parameter(Mandatory = $true)]
        [datetime]$CampEnd
    )

    for ($checkYear = $CampStart.Year; $checkYear -le $CampEnd.Year; $checkYear++) {
        try {
            $birthdayThisYear = [datetime]::new($checkYear, $BirthDate.Month, $BirthDate.Day)
        }
        catch {
            continue   # 29 feb op niet-schrikkeljaar
        }

        if ($birthdayThisYear.Date -ge $CampStart.Date -and $birthdayThisYear.Date -le $CampEnd.Date) {
            return $birthdayThisYear
        }
    }
    return $null
}

function Get-DutchMonthName {
    <#
        .SYNOPSIS
        Geeft de Nederlandse maandnaam voor een maandnummer.

        .PARAMETER Month
        Het maandnummer (1-12).

        .OUTPUTS
        System.String -- De Nederlandse maandnaam.
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

function Get-DutchDayName {
    <#
        .SYNOPSIS
        Geeft de Nederlandse dagnaam voor een dag van de week.

        .PARAMETER DayOfWeek
        De dag van de week als [System.DayOfWeek]-enum of als integer (0=zondag, 1=maandag, ..., 6=zaterdag).

        .OUTPUTS
        System.String -- De Nederlandse dagnaam (bijv. 'donderdag').
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 6)]
        [int]$DayOfWeek
    )

    $days = @('zondag', 'maandag', 'dinsdag', 'woensdag', 'donderdag', 'vrijdag', 'zaterdag')
    return $days[$DayOfWeek]
}

function Get-DutchGroupSizeLabel {
    <#
        .SYNOPSIS
        Geeft de Nederlandse beschrijving van een (sub)groepsgrootte.

        .PARAMETER Size
        Het aantal personen in de groep.

        .OUTPUTS
        System.String -- Nederlandse beschrijving (bijv. 'alleen opgegeven', "met z'n drieen").
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 99)]
        [int]$Size
    )

    $labels = @{
        1  = 'alleen opgegeven'
        2  = "met z'n twee" + [char]0x00EB + 'n'   # tweeën
        3  = "met z'n drie" + [char]0x00EB + 'n'   # drieën
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

function Assert-HitImportExcel {
    <#
        .SYNOPSIS
        Controleert of de module ImportExcel beschikbaar is en installeert deze indien nodig.

        .DESCRIPTION
        Controleert of ImportExcel geinstalleerd is. Zo niet, dan wordt de module automatisch
        geinstalleerd voor de huidige gebruiker. Vervolgens wordt de module geladen.
        Genereert een afbrekende fout als de installatie mislukt.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Verbose "Module 'ImportExcel' niet gevonden. Installeren voor de huidige gebruiker..."
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Verbose "Module 'ImportExcel' succesvol geinstalleerd."
        }
        catch {
            $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                [System.InvalidOperationException]::new(
                    "De module 'ImportExcel' kon niet automatisch worden geinstalleerd: $($_.Exception.Message)"
                ),
                'ImportExcelInstallFailed',
                [System.Management.Automation.ErrorCategory]::NotInstalled,
                $null
            )
            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }
    }

    Import-Module ImportExcel -ErrorAction Stop
}

function Resolve-HitExcelPath {
    <#
        .SYNOPSIS
        Bepaalt het pad naar het te gebruiken Excel-invoerbestand.

        .DESCRIPTION
        Zoekt eerst naar bestanden met het patroon "*-alles.xlsx" in ScriptDir.
        Als er precies één treffer is, wordt dat bestand automatisch geselecteerd.
        Als er meerdere zijn, krijgt de gebruiker een keuzemenu.
        Als er geen zijn, wordt opnieuw gezocht op "*.xlsx" (exclusief Deelnemerslijst_*-bestanden),
        met dezelfde automatische selectie of keuzemenu-logica.
        Als ook dat niets oplevert, wordt een afbrekende fout gegenereerd.

        .PARAMETER ScriptDir
        De map waarin naar xlsx-bestanden gezocht wordt.

        .OUTPUTS
        System.String -- Het absolute pad naar het geselecteerde Excel-bestand.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptDir
    )

    # Intern: toon keuzemenu en retourneer het geselecteerde pad
    function Select-FromMenu {
        param([System.IO.FileInfo[]]$Files)

        Write-Host ''
        Write-Host 'Beschikbare Excel-bestanden:' -ForegroundColor Cyan
        Write-Host ''
        for ($i = 0; $i -lt $Files.Count; $i++) {
            Write-Host "  [$($i + 1)] $($Files[$i].Name)" -ForegroundColor White
        }
        Write-Host ''

        $selectedPath = $null
        $validChoice  = $false
        while (-not $validChoice) {
            $choiceInput  = Read-Host "Kies een bestand (1-$($Files.Count))"
            $choiceNumber = 0
            if ([int]::TryParse($choiceInput, [ref]$choiceNumber) -and
                $choiceNumber -ge 1 -and $choiceNumber -le $Files.Count) {
                $selectedPath = $Files[$choiceNumber - 1].FullName
                $validChoice  = $true
            }
            else {
                Write-Host "Ongeldige keuze. Voer een nummer in van 1 t/m $($Files.Count)." -ForegroundColor Yellow
            }
        }
        return $selectedPath
    }

    # Stap 1: zoek op "*-alles.xlsx"
    $allesFiles = @(
        Get-ChildItem -Path $ScriptDir -Filter '*-alles.xlsx' -File | Sort-Object Name
    )

    if ($allesFiles.Count -eq 1) {
        Write-Verbose "Automatisch geselecteerd (*-alles.xlsx): $($allesFiles[0].Name)"
        return $allesFiles[0].FullName
    }

    if ($allesFiles.Count -gt 1) {
        Write-Verbose "Meerdere *-alles.xlsx gevonden in: $ScriptDir"
        $selected = Select-FromMenu -Files $allesFiles
        Write-Verbose "Geselecteerd: $selected"
        return $selected
    }

    # Stap 2: geen *-alles.xlsx gevonden — zoek op alle *.xlsx (excl. Deelnemerslijst_*)
    Write-Verbose "Geen *-alles.xlsx gevonden. Zoeken op *.xlsx in: $ScriptDir"
    $xlsxFiles = @(
        Get-ChildItem -Path $ScriptDir -Filter '*.xlsx' -File |
            Where-Object { $_.Name -notlike 'Deelnemerslijst_*' } |
            Sort-Object Name
    )

    if ($xlsxFiles.Count -eq 1) {
        Write-Verbose "Automatisch geselecteerd (*.xlsx): $($xlsxFiles[0].Name)"
        return $xlsxFiles[0].FullName
    }

    if ($xlsxFiles.Count -gt 1) {
        Write-Verbose "Meerdere *.xlsx gevonden in: $ScriptDir"
        $selected = Select-FromMenu -Files $xlsxFiles
        Write-Verbose "Geselecteerd: $selected"
        return $selected
    }

    # Stap 3: niets gevonden
    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
        [System.IO.FileNotFoundException]::new("Geen xlsx-bestanden gevonden in: $ScriptDir"),
        'NoXlsxFilesFound',
        [System.Management.Automation.ErrorCategory]::ObjectNotFound,
        $ScriptDir
    )
    $PSCmdlet.ThrowTerminatingError($errorRecord)
}

function ConvertFrom-HitBirthDate {
    <#
        .SYNOPSIS
        Parseert een ruwe geboortedatumwaarde naar een DateTime-object.

        .DESCRIPTION
        Verwerkt zowel al-geparseerde DateTime-objecten als tekstrepresentaties
        in meerdere gangbare formaten (dd-MM-yyyy, yyyy-MM-dd, M/d/yyyy, d-M-yyyy).
        Retourneert $null als de waarde leeg is of niet geparseerd kan worden.

        .PARAMETER RawValue
        De ruwe waarde zoals ingelezen uit Excel (DateTime of string).

        .OUTPUTS
        System.DateTime of $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        $RawValue
    )

    if ($null -eq $RawValue -or [string]::IsNullOrWhiteSpace($RawValue.ToString())) {
        return $null
    }

    if ($RawValue -is [datetime]) {
        return $RawValue
    }

    $formats = @('dd-MM-yyyy', 'yyyy-MM-dd', 'M/d/yyyy', 'd-M-yyyy')
    foreach ($fmt in $formats) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact(
                $RawValue.ToString().Trim(), $fmt,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [System.Globalization.DateTimeStyles]::None,
                [ref]$parsed)) {
            return $parsed
        }
    }

    return $null
}

function Select-HitFilePath {
    <#
        .SYNOPSIS
        Toont een keuzemenu voor het selecteren van een xlsx- of csv-bestand.

        .DESCRIPTION
        Zoekt naar alle *.xlsx- en *.csv-bestanden in ScriptDir, exclusief bestanden
        die beginnen met 'Deelnemerslijst_'. Bij één treffer wordt automatisch geselecteerd;
        bij meerdere treffers verschijnt een genummerd console-keuzemenu.
        Genereert een afbrekende fout als er geen bestanden worden gevonden.

        .PARAMETER Prompt
        De omschrijving die boven het keuzemenu wordt weergegeven, bijvoorbeeld
        'Kies deelnemerslijst 1 (vorig jaar)'.

        .PARAMETER ScriptDir
        De map waarin naar xlsx- en csv-bestanden gezocht wordt.

        .PARAMETER ExcludePaths
        Optionele lijst van volledige bestandspaden die uitgesloten worden van de selectie.
        Handig om een eerder geselecteerd bestand te verbergen bij een tweede keuze.

        .PARAMETER AllowSkip
        Als opgegeven, verschijnt optie [0] Overslaan bovenaan het keuzemenu.
        Bij keuze 0 of lege invoer wordt $null geretourneerd (geen fout).
        Als er geen bestanden worden gevonden én AllowSkip opgegeven is, wordt eveneens
        $null geretourneerd in plaats van een afbrekende fout.

        .OUTPUTS
        System.String of $null -- Het absolute pad naar het geselecteerde bestand, of $null bij overslaan.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [string]$ScriptDir,

        [Parameter()]
        [string[]]$ExcludePaths = @(),

        [Parameter()]
        [switch]$AllowSkip
    )

    $candidateFiles = @(
        Get-ChildItem -Path $ScriptDir -File |
            Where-Object {
                $_.Extension -in @('.xlsx', '.csv') -and
                $_.Name -notlike 'Deelnemerslijst_*' -and
                $_.FullName -notin $ExcludePaths
            } |
            Sort-Object Name
    )

    if ($candidateFiles.Count -eq 0) {
        if ($AllowSkip) {
            Write-Verbose "Geen bestanden gevonden voor '$Prompt'. Overgeslagen."
            return $null
        }
        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
            [System.IO.FileNotFoundException]::new("Geen xlsx- of csv-bestanden gevonden in: $ScriptDir"),
            'NoDataFilesFound',
            [System.Management.Automation.ErrorCategory]::ObjectNotFound,
            $ScriptDir
        )
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }

    if ($candidateFiles.Count -eq 1 -and -not $AllowSkip) {
        Write-Verbose "Automatisch geselecteerd: $($candidateFiles[0].Name)"
        return $candidateFiles[0].FullName
    }

    Write-Host ''
    Write-Host "${Prompt}:" -ForegroundColor Cyan
    Write-Host ''
    if ($AllowSkip) {
        Write-Host '  [0] Overslaan' -ForegroundColor DarkGray
    }
    for ($i = 0; $i -lt $candidateFiles.Count; $i++) {
        Write-Host "  [$($i + 1)] $($candidateFiles[$i].Name)" -ForegroundColor White
    }
    Write-Host ''

    $selectedPath = $null
    $validChoice  = $false
    while (-not $validChoice) {
        $prompt = if ($AllowSkip) { "Kies een bestand (0 = overslaan, 1-$($candidateFiles.Count))" } else { "Kies een bestand (1-$($candidateFiles.Count))" }
        $choiceInput  = Read-Host $prompt

        if ($AllowSkip -and [string]::IsNullOrWhiteSpace($choiceInput)) {
            Write-Verbose 'Overgeslagen (lege invoer).'
            return $null
        }

        $choiceNumber = 0
        if ([int]::TryParse($choiceInput, [ref]$choiceNumber)) {
            if ($AllowSkip -and $choiceNumber -eq 0) {
                Write-Verbose 'Overgeslagen (keuze 0).'
                return $null
            }
            if ($choiceNumber -ge 1 -and $choiceNumber -le $candidateFiles.Count) {
                $selectedPath = $candidateFiles[$choiceNumber - 1].FullName
                $validChoice  = $true
            }
        }

        if (-not $validChoice) {
            $rangeHint = if ($AllowSkip) { "0 t/m $($candidateFiles.Count)" } else { "1 t/m $($candidateFiles.Count)" }
            Write-Host "Ongeldige keuze. Voer een nummer in van $rangeHint." -ForegroundColor Yellow
        }
    }
    Write-Verbose "Geselecteerd: $selectedPath"
    return $selectedPath
}

function Get-HitOutputBaseName {
    <#
        .SYNOPSIS
        Genereert een schone basisnaam voor een outputbestand op basis van het inputpad.

        .DESCRIPTION
        Verwijdert de bestandsextensie en bekende suffixen zoals " (NNNN)-alles" en "-alles",
        trimmt spaties en vervangt resterende spaties door underscores.

        .PARAMETER InputPath
        Het volledige pad naar het input-bestand.

        .OUTPUTS
        System.String -- De schone basisnaam, geschikt voor gebruik in een outputbestandsnaam.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $cleanName = $baseName -replace '\s*\(\d+\)-alles$', '' -replace '-alles$', ''
    $cleanName = $cleanName.Trim() -replace '\s+', '_'
    return $cleanName
}

function Import-ParticipantFile {
    <#
        .SYNOPSIS
        Importeert een deelnemersbestand (xlsx of csv) en retourneert de rijen.

        .DESCRIPTION
        Ondersteunt Excel-bestanden (.xlsx, via Import-Excel) en
        puntkommagescheiden CSV-bestanden (.csv, via Import-Csv -Delimiter ';').
        Genereert een afbrekende fout bij een niet-ondersteund bestandsformaat.

        .PARAMETER Path
        Het volledige pad naar het te importeren bestand.

        .OUTPUTS
        System.Object[] -- Een array van geparseerde rijen.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        '.csv'  { return @(Import-Csv -Path $Path -Delimiter ';' -ErrorAction Stop) }
        '.xlsx' { return @(Import-Excel -Path $Path -ErrorAction Stop) }
        default {
            $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                [System.ArgumentException]::new("Niet-ondersteund bestandsformaat: $extension"),
                'UnsupportedFileFormat',
                [System.Management.Automation.ErrorCategory]::InvalidArgument,
                $Path
            )
            $PSCmdlet.ThrowTerminatingError($errorRecord)
        }
    }
}

function ConvertTo-NormalizedParticipant {
    <#
        .SYNOPSIS
        Normaliseert een rij uit een deelnemersbestand naar een gestandaardiseerd object.

        .DESCRIPTION
        Ondersteunt twee kolomformaten:
        - HIT-registratiesysteem: 'Voornaam', 'Tussenvoegsel', 'Achternaam', 'Geboortedatum'
        - Scouting-CSV: 'Lid voornaam', 'Lid tussenvoegsel', 'Lid achternaam', 'Lid geboortedatum'
        Retourneert $null voor lege of niet-herkenbare rijen.

        .PARAMETER Row
        De te normaliseren rij als PSObject (afkomstig uit Import-Excel of Import-Csv).

        .OUTPUTS
        PSCustomObject met Sleutel, VolledigeNaam, Geboortedatum — of $null.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Row
    )

    $columnNames = $Row.PSObject.Properties.Name

    if ($columnNames -contains 'Lid voornaam') {
        $firstName    = [string]$Row.'Lid voornaam'
        $insertion    = [string]$Row.'Lid tussenvoegsel'
        $lastName     = [string]$Row.'Lid achternaam'
        $rawBirthDate = $Row.'Lid geboortedatum'
    }
    elseif ($columnNames -contains 'Voornaam') {
        $firstName    = [string]$Row.'Voornaam'
        $insertion    = if ($columnNames -contains 'Tussenvoegsel') { [string]$Row.'Tussenvoegsel' } else { '' }
        $lastName     = [string]$Row.'Achternaam'
        $rawBirthDate = $Row.'Geboortedatum'
    }
    else {
        Write-Warning 'Onbekend kolomformaat — rij overgeslagen.'
        return $null
    }

    $nameParts = @($firstName, $insertion, $lastName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $fullName  = ($nameParts -join ' ').Trim()

    if ([string]::IsNullOrWhiteSpace($fullName)) {
        return $null
    }

    $birthDate    = ConvertFrom-HitBirthDate -RawValue $rawBirthDate
    $birthDateKey = if ($null -ne $birthDate) { $birthDate.ToString('yyyy-MM-dd') } else { '' }
    $key          = "$($fullName.ToLowerInvariant())|$birthDateKey"

    return [PSCustomObject]@{
        Sleutel       = $key
        VolledigeNaam = $fullName
        Geboortedatum = $birthDate
    }
}

Export-ModuleMember -Function @(
    'Get-EasterSunday',
    'Get-AgeAtDate',
    'Get-BirthdayDuringCamp',
    'Get-BirthdayDateDuringCamp',
    'Get-DutchMonthName',
    'Get-DutchDayName',
    'Get-DutchGroupSizeLabel',
    'Assert-HitImportExcel',
    'Resolve-HitExcelPath',
    'Select-HitFilePath',
    'ConvertFrom-HitBirthDate',
    'Get-HitOutputBaseName',
    'Import-ParticipantFile',
    'ConvertTo-NormalizedParticipant'
)
