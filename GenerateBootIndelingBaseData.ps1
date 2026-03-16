<#
    .SYNOPSIS
    Genereert basisdata voor de bootindeling vanuit het deelnemersbestand en de Google Forms-export.

    .DESCRIPTION
    Combineert twee bestanden:
    1. Het deelnemersbestand (xlsx of puntkomma-gescheiden csv) met naam, geboortedatum,
       geslacht en subgroep van alle ingeschreven deelnemers.
    2. De Google Forms-export (komma-gescheiden csv) met antwoorden over bezwaar tegen
       zeilen in een andere boot dan de groep, en de zeilervaring.

    Per deelnemer worden de volgende kolommen gegenereerd:
    - Naam          : Voornaam (alleen achternaam erbij als meerdere deelnemers dezelfde voornaam hebben)
    - Leeftijd      : Leeftijd op de eerste dag van het kamp. Deelnemers die tijdens het kamp
                      jarig zijn krijgen beide leeftijden te zien (bijv. "14/15").
    - Geslacht      : Geslacht uit het deelnemersbestand.
    - Groep         : Subgroep uit het deelnemersbestand.
    - Zeilen met je groep? : Op basis van de bezwaar-vraag uit het formulier (bezwaar=Ja → "Ja").
    - Zeilervaring  : Antwoord op de zeilervaring-vraag uit het formulier.

    Deelnemers die het formulier NIET hebben ingevuld worden in het Excel-bestand rood gemarkeerd
    (achtergrond RGB 255, 199, 206). De rijen worden gesorteerd op Groep, daarbinnen op Naam.

    Het output-bestand wordt opgeslagen in dezelfde map als het script:
    Deelnemerslijst_[basisnaam]_[jaar]_BootIndeling.xlsx

    .PARAMETER Year
    Het jaar van het HIT-kamp. Wordt gebruikt om de paasdatum en leeftijden te berekenen.
    Standaard: het huidige jaar.

    .EXAMPLE
    .\GenerateBootIndelingBaseData.ps1

    Selecteer interactief het deelnemersbestand en de Google Forms-export voor het huidige jaar.

    .EXAMPLE
    .\GenerateBootIndelingBaseData.ps1 -Year 2026

    Zoals hierboven, maar voor het opgegeven jaar.

    .OUTPUTS
    System.IO.FileInfo
    Het pad naar het gemaakte Excel-bestand.

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

    [Parameter(HelpMessage = 'Minimale naamsovereenkomst (0.0-1.0) voor fuzzy name matching. Standaard: 0.85 (85%).')]
    [ValidateRange(0.0, 1.0)]
    [double]$MatchThreshold = 0.85
)

#region Import Shared Module

$_moduleRoot = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }
Import-Module -Name (Join-Path -Path $_moduleRoot -ChildPath 'HitHelpers.psm1') -Force
Remove-Variable -Name _moduleRoot

#endregion Import Shared Module

#region Initialisation

Assert-HitImportExcel

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { $PWD.Path }

#endregion Initialisation

#region Helper Functions

function Import-ParticipantFile {
    <#
        .SYNOPSIS
        Laadt een xlsx- of puntkomma-gescheiden csv-bestand als array van rijen.
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
            throw "Niet-ondersteund bestandsformaat: $extension"
        }
    }
}

function Get-NormalizedName {
    <#
        .SYNOPSIS
        Normaliseert een volledige naam voor vergelijking: lowercase, getrimde spaties.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullName
    )

    return ($FullName.Trim() -replace '\s+', ' ').ToLowerInvariant()
}

function Find-FormColumn {
    <#
        .SYNOPSIS
        Zoekt een kolomnaam in een array van kolomnamen op basis van een wildcard-patroon.
        Geeft de eerste match terug, of $null als er geen match is.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ColumnNames,

        [Parameter(Mandatory = $true)]
        [string]$Pattern
    )

    return $ColumnNames | Where-Object { $_ -ilike $Pattern } | Select-Object -First 1
}

function Get-LevenshteinDistance {
    <#
        .SYNOPSIS
        Berekent de Levenshtein-afstand (edit distance) tussen twee strings.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$Source,
        [Parameter(Mandatory = $true)] [string]$Target
    )

    $n = $Source.Length
    $m = $Target.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }

    $prev = [int[]](0..$m)
    $curr = [int[]]::new($m + 1)

    for ($i = 1; $i -le $n; $i++) {
        $curr[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($Source[$i - 1] -eq $Target[$j - 1]) { 0 } else { 1 }
            $curr[$j] = [Math]::Min(
                [Math]::Min($prev[$j] + 1, $curr[$j - 1] + 1),
                $prev[$j - 1] + $cost
            )
        }
        $temp = $prev; $prev = $curr; $curr = $temp
    }
    return $prev[$m]
}

function Get-NameSimilarity {
    <#
        .SYNOPSIS
        Geeft de mate van gelijkenis (0.0-1.0) tussen twee genormaliseerde namen.
        Gebaseerd op Levenshtein-afstand gedeeld door de maximale stringlengte.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$NameA,
        [Parameter(Mandatory = $true)] [string]$NameB
    )

    $maxLen = [Math]::Max($NameA.Length, $NameB.Length)
    if ($maxLen -eq 0) { return 1.0 }
    return 1.0 - ((Get-LevenshteinDistance -Source $NameA -Target $NameB) / $maxLen)
}

function Find-BestFuzzyMatch {
    <#
        .SYNOPSIS
        Zoekt de best overeenkomende naam in een lijst op basis van Get-NameSimilarity.
        Geeft $null terug als de beste score onder de drempelwaarde valt.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string]$NormalizedName,
        [Parameter(Mandatory = $true)] [string[]]$Candidates,
        [Parameter(Mandatory = $true)] [double]$Threshold
    )

    $bestCandidate = $null
    $bestScore     = 0.0
    foreach ($candidate in $Candidates) {
        $score = Get-NameSimilarity -NameA $NormalizedName -NameB $candidate
        if ($score -gt $bestScore) {
            $bestScore     = $score
            $bestCandidate = $candidate
        }
    }
    if ($bestScore -ge $Threshold) {
        return [PSCustomObject]@{ Name = $bestCandidate; Score = $bestScore }
    }
    return $null
}

#endregion Helper Functions

#region Banner

Write-Host ''
Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host '  HIT Bootindeling Basisdata Generator' -ForegroundColor Cyan
Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host ''

#endregion Banner

#region File Selection

$deelnemersPath = Select-HitFilePath -Prompt 'Kies het deelnemersbestand (bijv. Zeilzwerf Fryslan-alles.xlsx)' -ScriptDir $scriptDir

$formsPath = Select-HitFilePath `
    -Prompt 'Kies de Google Forms export (bijv. Zeilzwerf Fryslân 2026.csv)' `
    -ScriptDir $scriptDir `
    -ExcludePaths @($deelnemersPath)

Write-Verbose "Deelnemersbestand : $deelnemersPath"
Write-Verbose "Google Forms bestand: $formsPath"

#endregion File Selection

#region Camp Date Calculation

$easterSunday = Get-EasterSunday -Year $Year
$campStart    = $easterSunday.AddDays(-2)   # Goede Vrijdag
$campEnd      = $easterSunday.AddDays(1)    # Tweede Paasdag

Write-Verbose ("Kamp {0}: Goede Vrijdag {1:dd-MM-yyyy} t/m Tweede Paasdag {2:dd-MM-yyyy}" -f $Year, $campStart, $campEnd)

#endregion Camp Date Calculation

#region Load Participant File

Write-Host 'Deelnemersbestand inladen...' -ForegroundColor DarkGray
$deelnemersRows = Import-ParticipantFile -Path $deelnemersPath
$totalDeelnemers = $deelnemersRows.Count
Write-Verbose "Ingelezen: $totalDeelnemers rij(en) uit deelnemersbestand."

if ($totalDeelnemers -eq 0) {
    throw "Het deelnemersbestand bevat geen rijen: $deelnemersPath"
}

# Detecteer kolomformaat: raw CSV ("Lid voornaam") vs. xlsx ("Voornaam")
$sampleColumns = $deelnemersRows[0].PSObject.Properties.Name
$isRawCsv      = $sampleColumns -contains 'Lid voornaam'

if (-not $isRawCsv -and $sampleColumns -notcontains 'Voornaam') {
    throw "Onbekend kolomformaat in deelnemersbestand. Verwacht 'Lid voornaam' (csv) of 'Voornaam' (xlsx)."
}

$kolomformaatLabel = if ($isRawCsv) { 'raw CSV (puntkomma)' } else { 'xlsx (hernoemd)' }
Write-Verbose "Kolomformaat: $kolomformaatLabel"

#endregion Load Participant File

#region Load Google Forms File

Write-Host 'Google Forms export inladen...' -ForegroundColor DarkGray
$formsRows = @(Import-Csv -Path $formsPath -ErrorAction Stop)
Write-Verbose "Ingelezen: $($formsRows.Count) rij(en) uit Google Forms export."

# Detecteer benodigde kolommen op basis van substring-match (case-insensitief)
$formsColumns = if ($formsRows.Count -gt 0) { $formsRows[0].PSObject.Properties.Name } else { @() }

$formsNaamKolom           = Find-FormColumn -ColumnNames $formsColumns -Pattern '*voor- en achternaam*'
$formsBezwaarKolom        = Find-FormColumn -ColumnNames $formsColumns -Pattern '*andere boot*'
$formsZeilervaringKolom   = Find-FormColumn -ColumnNames $formsColumns -Pattern '*zeilervaring*'
$formsReddingsvestKolom   = Find-FormColumn -ColumnNames $formsColumns -Pattern '*reddingsvest mee*'
$formsGewichtKolom        = Find-FormColumn -ColumnNames $formsColumns -Pattern '*gewicht*'

if ($null -eq $formsNaamKolom) {
    throw "Geen naamkolom gevonden in Google Forms export. Verwacht een kolom met 'voor- en achternaam' in de naam. Gevonden kolommen: $($formsColumns -join ', ')"
}

Write-Verbose "Forms naamkolom          : '$formsNaamKolom'"
Write-Verbose "Forms bezwaarkolom       : $(if ($null -ne $formsBezwaarKolom) { "'$formsBezwaarKolom'" } else { '(niet gevonden)' })"
Write-Verbose "Forms zeilervaringkolom  : $(if ($null -ne $formsZeilervaringKolom) { "'$formsZeilervaringKolom'" } else { '(niet gevonden)' })"
Write-Verbose "Forms reddingsvestkolom  : $(if ($null -ne $formsReddingsvestKolom) { "'$formsReddingsvestKolom'" } else { '(niet gevonden)' })"
Write-Verbose "Forms gewichtkolom       : $(if ($null -ne $formsGewichtKolom) { "'$formsGewichtKolom'" } else { '(niet gevonden)' })"

# Bouw lookup-hashtable: genormaliseerde naam → form-rij
# Bij dubbele inzendingen wint de laatste (meest recente indien gesorteerd op Tijdstempel)
$formsLookup = @{}
foreach ($formRow in $formsRows) {
    $rawName = [string]$formRow.$formsNaamKolom
    if ([string]::IsNullOrWhiteSpace($rawName)) { continue }
    $normalizedName = Get-NormalizedName -FullName $rawName
    $formsLookup[$normalizedName] = $formRow
}

Write-Verbose "$($formsLookup.Count) unieke deelnemers gevonden in Google Forms export."
$allFormKeys = @($formsLookup.Keys)

#endregion Load Google Forms File

#region Normalize Participants and Detect Duplicate First Names

$participants = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $deelnemersRows) {
    if ($isRawCsv) {
        $firstName   = [string]$row.'Lid voornaam'
        $insertion   = [string]$row.'Lid tussenvoegsel'
        $lastName    = [string]$row.'Lid achternaam'
        $geslacht    = [string]$row.'Lid geslacht'
        $rawBirth    = $row.'Lid geboortedatum'
        $groep       = [string]$row.'Subgroepnaam'
        $emailadres  = [string]$row.'Lid e-mailadres'
    }
    else {
        $firstName   = [string]$row.'Voornaam'
        $insertion   = if ($sampleColumns -contains 'Tussenvoegsel') { [string]$row.'Tussenvoegsel' } else { '' }
        $lastName    = [string]$row.'Achternaam'
        $geslacht    = if ($sampleColumns -contains 'Gender') { [string]$row.'Gender' } else { [string]$row.'Geslacht' }
        $rawBirth    = $row.'Geboortedatum'
        # Probeer 'Subgroep' eerst (xlsx), dan 'Groep', dan 'Subgroepnaam' als fallback
        $groep       = if ($sampleColumns -contains 'Subgroep') { [string]$row.'Subgroep' }
                       elseif ($sampleColumns -contains 'Groep') { [string]$row.'Groep' }
                       elseif ($sampleColumns -contains 'Subgroepnaam') { [string]$row.'Subgroepnaam' }
                       else { '' }
        $emailadres  = if ($sampleColumns -contains 'Mailadres') { [string]$row.'Mailadres' }
                       elseif ($sampleColumns -contains 'E-mailadres') { [string]$row.'E-mailadres' }
                       else { '' }
    }

    if ([string]::IsNullOrWhiteSpace($firstName) -and [string]::IsNullOrWhiteSpace($lastName)) {
        continue
    }

    # Bouw volledige naam (voornaam + tussenvoegsel + achternaam)
    $nameParts = @($firstName, $insertion, $lastName) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    $fullName  = ($nameParts -join ' ').Trim()

    $birthDate = ConvertFrom-HitBirthDate -RawValue $rawBirth

    $participants.Add([PSCustomObject]@{
        Voornaam      = $firstName.Trim()
        Achternaam    = $lastName.Trim()
        VolledNaam    = $fullName
        Geslacht      = $geslacht.Trim()
        Geboortedatum = $birthDate
        Groep         = $groep.Trim()
        Emailadres    = $emailadres.Trim()
    })
}

Write-Verbose "$($participants.Count) deelnemers genormaliseerd."

# Tel how many deelnemers per voornaam (lowercase) → voor Naam-kolom bepalen
$firstNameCounts = @{}
foreach ($p in $participants) {
    $key = $p.Voornaam.ToLowerInvariant()
    $firstNameCounts[$key] = ($firstNameCounts[$key] -as [int]) + 1
}

#endregion Normalize Participants and Detect Duplicate First Names

#region Report Unmatched Form Entries

# Bouw een set + omgekeerde lookup van genormaliseerde namen uit het deelnemersbestand
$participantNameSet       = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$participantKeyToOriginal = @{}
foreach ($p in $participants) {
    $key = Get-NormalizedName -FullName $p.VolledNaam
    $participantNameSet.Add($key) | Out-Null
    $participantKeyToOriginal[$key] = $p.VolledNaam
}
$allParticipantKeys = [string[]]$participantNameSet

# Vergelijk elke formulierregel met de deelnemersset: exact, dan fuzzy, anders onbekend
# Regelnummer in de CSV: header = regel 1, eerste dataregel = regel 2
$unmatchedFormEntries = [System.Collections.Generic.List[PSCustomObject]]::new()
$fuzzyFormEntries     = [System.Collections.Generic.List[PSCustomObject]]::new()

for ($fi = 0; $fi -lt $formsRows.Count; $fi++) {
    $rawName = [string]$formsRows[$fi].$formsNaamKolom
    if ([string]::IsNullOrWhiteSpace($rawName)) { continue }
    $normalizedName = Get-NormalizedName -FullName $rawName
    if ($participantNameSet.Contains($normalizedName)) { continue }   # exacte match

    $fuzzyResult = Find-BestFuzzyMatch -NormalizedName $normalizedName -Candidates $allParticipantKeys -Threshold $MatchThreshold
    if ($null -ne $fuzzyResult) {
        $fuzzyFormEntries.Add([PSCustomObject]@{
            Regelnummer   = $fi + 2
            FormNaam      = $rawName.Trim()
            DeelnemerNaam = $participantKeyToOriginal[$fuzzyResult.Name]
            Score         = $fuzzyResult.Score
        })
    }
    else {
        $unmatchedFormEntries.Add([PSCustomObject]@{
            Regelnummer = $fi + 2
            Naam        = $rawName.Trim()
        })
    }
}

if ($fuzzyFormEntries.Count -gt 0) {
    Write-Host ''
    Write-Host ("  Info: {0} formulierregel(s) automatisch gematcht via naamsgelijkenis (drempel: {1:P0}):" -f $fuzzyFormEntries.Count, $MatchThreshold) -ForegroundColor Cyan
    foreach ($entry in $fuzzyFormEntries) {
        Write-Host ("    Regel {0,3}: '{1}'  ->  '{2}'  ({3:P0})" -f $entry.Regelnummer, $entry.FormNaam, $entry.DeelnemerNaam, $entry.Score) -ForegroundColor Cyan
    }
}

if ($unmatchedFormEntries.Count -gt 0) {
    Write-Host ''
    Write-Host "  Let op: $($unmatchedFormEntries.Count) formulierregels konden niet gematcht worden" `
        -ForegroundColor Yellow
    Write-Host '  (naam in het formulier komt niet voor in het deelnemersbestand):' `
        -ForegroundColor Yellow
    foreach ($entry in $unmatchedFormEntries) {
        Write-Host ("    Regel {0,3}: {1}" -f $entry.Regelnummer, $entry.Naam) -ForegroundColor Yellow
    }
}

#endregion Report Unmatched Form Entries

#region Build Output Rows

$outputRows   = [System.Collections.Generic.List[PSCustomObject]]::new()
$matchedFlags = [System.Collections.Generic.List[bool]]::new()

foreach ($p in $participants) {
    # ── Naam ──────────────────────────────────────────────────────────────────
    $firstNameKey = $p.Voornaam.ToLowerInvariant()
    $displayName  = if ($firstNameCounts[$firstNameKey] -gt 1) {
                        # Bouw naam met tussenvoegsel + achternaam (als aanwezig)
                        $nameParts = @($p.Voornaam, $p.Achternaam) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                        ($nameParts -join ' ').Trim()
                    }
                    else {
                        $p.Voornaam
                    }

    # ── Leeftijd ──────────────────────────────────────────────────────────────
    $leeftijdValue = ''
    if ($null -ne $p.Geboortedatum) {
        $ageAtStart  = Get-AgeAtDate -BirthDate $p.Geboortedatum -ReferenceDate $campStart
        $hasBirthday = Get-BirthdayDuringCamp -BirthDate $p.Geboortedatum -CampStart $campStart -CampEnd $campEnd
        $leeftijdValue = if ($hasBirthday) { "$ageAtStart/$($ageAtStart + 1)" } else { "$ageAtStart" }
    }
    else {
        Write-Warning "Kon geboortedatum niet verwerken voor '$($p.VolledNaam)'"
    }

    # ── Forms opzoeken: exact eerst, daarna fuzzy als exacte match ontbreekt ──
    $normalizedKey = Get-NormalizedName -FullName $p.VolledNaam
    $formRow       = $formsLookup[$normalizedKey]
    if ($null -eq $formRow) {
        $fuzzyResult = Find-BestFuzzyMatch -NormalizedName $normalizedKey -Candidates $allFormKeys -Threshold $MatchThreshold
        if ($null -ne $fuzzyResult) {
            $formRow = $formsLookup[$fuzzyResult.Name]
            Write-Verbose ("Fuzzy match: '{0}' -> '{1}' ({2:P0})" -f $normalizedKey, $fuzzyResult.Name, $fuzzyResult.Score)
        }
    }
    $isMatched     = $null -ne $formRow

    $zeilenMetGroep  = ''
    $zeilervaring    = ''
    $reddingsvest    = ''
    $gewicht         = ''

    if ($isMatched) {
        if ($null -ne $formsBezwaarKolom) {
            $bezwaarRaw = [string]$formRow.$formsBezwaarKolom
            # Bezwaar=Ja betekent: wil NIET in andere boot → wil WEL met groep
            # Bezwaar=Nee betekent: wil WEL in andere boot → wil NIET per se met groep
            $zeilenMetGroep = switch ($bezwaarRaw.Trim().ToLowerInvariant()) {
                'ja'  { 'Ja' }
                'nee' { 'Nee' }
                default { $bezwaarRaw.Trim() }
            }
        }
        if ($null -ne $formsZeilervaringKolom) {
            $zeilervaring = ([string]$formRow.$formsZeilervaringKolom).Trim()
        }
        if ($null -ne $formsReddingsvestKolom) {
            $reddingsvestRaw = ([string]$formRow.$formsReddingsvestKolom).Trim()
            $reddingsvest = switch ($reddingsvestRaw.ToLowerInvariant()) {
                'ja'  { 'Ja' }
                'nee' { 'Nee' }
                default { $reddingsvestRaw }
            }
        }
        if ($null -ne $formsGewichtKolom) {
            $gewicht = ([string]$formRow.$formsGewichtKolom).Trim()
        }
    }

    $outputRows.Add([PSCustomObject]@{
        Naam                       = $displayName
        Leeftijd                   = $leeftijdValue
        Geslacht                   = $p.Geslacht
        Groep                      = $p.Groep
        'Zeilen met je groep?'     = $zeilenMetGroep
        Zeilervaring               = $zeilervaring
        'Eigen reddingsvest?'      = $reddingsvest
        'Gewicht (kg)'             = $gewicht
        Emailadres                 = $p.Emailadres
    })
    $matchedFlags.Add($isMatched)
}

Write-Verbose "$($outputRows.Count) uitvoerrijen opgebouwd."
$unmatchedCount = ($matchedFlags | Where-Object { -not $_ } | Measure-Object).Count

#endregion Build Output Rows

#region Sort Output

$sortedPairs = [System.Collections.Generic.List[PSCustomObject]]::new()
for ($i = 0; $i -lt $outputRows.Count; $i++) {
    $sortedPairs.Add([PSCustomObject]@{
        Row       = $outputRows[$i]
        IsMatched = $matchedFlags[$i]
    })
}

$sortedPairs = @($sortedPairs | Sort-Object -Property { $_.Row.Groep }, { $_.Row.Naam })

$sortedRows    = $sortedPairs | ForEach-Object { $_.Row }
$sortedMatched = $sortedPairs | ForEach-Object { $_.IsMatched }

#endregion Sort Output

#region Export to Excel

$cleanBaseName  = Get-HitOutputBaseName -InputPath $deelnemersPath
$outputFileName = "Deelnemerslijst_${cleanBaseName}_${Year}_BootIndeling.xlsx"
$outputDir      = [System.IO.Path]::GetDirectoryName($deelnemersPath)
$outputPath     = Join-Path -Path $outputDir -ChildPath $outputFileName

Write-Verbose "Exporteren naar: $outputPath"
Write-Host "Exporteren naar: $outputFileName" -ForegroundColor DarkGray

$excelParams = @{
    Path          = $outputPath
    WorksheetName = 'BootIndeling'
    TableName     = 'Deelnemers'
    TableStyle    = 'Medium2'
    AutoSize      = $true
    ClearSheet    = $true
    PassThru      = $true
}

$excelPackage = $sortedRows | Export-Excel @excelParams
$worksheet    = $excelPackage.Workbook.Worksheets['BootIndeling']

# Bepaal kolomindices (1-gebaseerd) op basis van koptekstrij
$headerRow     = $worksheet.Dimension.Start.Row
$lastCol       = $worksheet.Dimension.End.Column
$leeftijdColIndex = $null

for ($col = 1; $col -le $lastCol; $col++) {
    $headerValue = $worksheet.Cells[$headerRow, $col].Text
    if ($headerValue -eq 'Leeftijd') {
        $leeftijdColIndex = $col
        break
    }
}

$lastDataRow = $worksheet.Dimension.End.Row

# ── Forceer Leeftijd-kolom als tekst ─────────────────────────────────────────
# EPPlus slaat "14/15" anders op als datum; herschrijf expliciet als string.
if ($null -ne $leeftijdColIndex) {
    for ($rowIdx = ($headerRow + 1); $rowIdx -le $lastDataRow; $rowIdx++) {
        $cell = $worksheet.Cells[$rowIdx, $leeftijdColIndex]
        $cellText = $cell.Text
        $cell.Style.Numberformat.Format = '@'
        $cell.Value = $cellText
    }
}

# ── Markeer ongematchte deelnemers rood ──────────────────────────────────────
# RGB 255, 199, 206 = lichtroze/rood (standaard Excel "slechte cel"-kleur)
for ($i = 0; $i -lt $sortedMatched.Count; $i++) {
    if (-not $sortedMatched[$i]) {
        $dataRowIndex = $headerRow + 1 + $i
        if ($dataRowIndex -le $lastDataRow) {
            $rowRange = $worksheet.Cells[$dataRowIndex, 1, $dataRowIndex, $lastCol]
            $rowRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $rowRange.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255, 255, 199, 206))
        }
    }
}

Close-ExcelPackage -ExcelPackage $excelPackage -Show:$false

Write-Host ''
Write-Host "Export voltooid: $outputFileName" -ForegroundColor Green
Write-Host "  Totaal deelnemers : $($sortedRows.Count)" -ForegroundColor White
Write-Host "  Formulier ingevuld: $($sortedRows.Count - $unmatchedCount)" -ForegroundColor White
if ($unmatchedCount -gt 0) {
    Write-Host "  Niet gevonden     : $unmatchedCount (rood gemarkeerd)" -ForegroundColor Yellow
}
Write-Host ''

Get-Item -Path $outputPath

#endregion Export to Excel
