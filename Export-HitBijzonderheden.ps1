<#
    .SYNOPSIS
    Exporteert een gefilterde deelnemerslijst met bijzonderheden naar een Excel-bestand.

    .DESCRIPTION
    Leest een HIT-registratie Excel-bestand (.xlsx, zoals geëxporteerd uit het
    Scouting-aanmeldingssysteem) en exporteert een nieuw Excel-bestand met de kolommen:
    Voornaam, Achternaam, Geslacht, Leeftijd, Bijzonderheden.

    Alleen deelnemers met minstens één ingevuld dieet of aandachtspunt worden opgenomen.

    De leeftijd wordt berekend voor tijdens het kamp (Goede Vrijdag t/m Tweede Paasdag).
    Als een deelnemer jarig is tijdens het kamp, worden beide leeftijden getoond
    (bijv. "14/15").

    De Bijzonderheden-kolom combineert:
    - Alle ingevulde dieetbeperkingen (exclusief "Geen dieet"), inclusief vrije-tekst
      redenen uit de bijbehorende "Dieet Reden:"-kolommen.
    - De aandachtspunten uit de medische/allergie-kolom.

    Het output-bestand wordt opgeslagen in dezelfde map als het script, met de naam:
    Deelnemerslijst_[basisnaam]_[jaar]_Bijzonderheden.xlsx

    Het script zoekt automatisch naar een bestand met het patroon "*-alles.xlsx" in de
    scriptmap. Als er meerdere zijn, verschijnt een keuzemenu. Als er geen zijn, wordt
    gezocht op "*.xlsx". Als er dan ook geen is, volgt een foutmelding.

    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is.

    .PARAMETER Year
    Het jaar van het HIT-kamp. Wordt gebruikt om de paasdatum te berekenen.
    Standaard: het huidige jaar.

    .EXAMPLE
    .\Export-HitBijzonderheden.ps1

    Verwerkt het gevonden *-alles.xlsx-bestand in de scriptmap voor het huidige jaar.

    .EXAMPLE
    .\Export-HitBijzonderheden.ps1 -Verbose

    Zoals hierboven, maar met uitgebreide voortgangsberichten.

    .OUTPUTS
    System.IO.FileInfo
    Het pad naar het gemaakte Excel-bestand met bijzonderheden.

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
    [int]$Year = (Get-Date).Year
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

#region Import Data

Write-Verbose "Inlezen van Excel-bestand..."
$allRows = Import-Excel -Path $resolvedInputPath -ErrorAction Stop
$totalRows = ($allRows | Measure-Object).Count
Write-Verbose "Ingelezen: $totalRows rij(en)."

# Verplichte kolommen
$requiredColumns = @(
    'Voornaam',
    'Achternaam',
    'Gender',
    'Geboortedatum'
)

$actualColumns = $allRows[0].PSObject.Properties.Name
foreach ($col in $requiredColumns) {
    if ($col -notin $actualColumns) {
        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
            [System.Exception]::new("Verplichte kolom ontbreekt in het input-bestand: '$col'"),
            'MissingRequiredColumn',
            [System.Management.Automation.ErrorCategory]::InvalidData,
            $col
        )
        $PSCmdlet.ThrowTerminatingError($errorRecord)
    }
}

#endregion Import Data

#region Build Output Rows

$outputRows = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($row in $allRows) {
    # --- Leeftijd ---
    $birthDateRaw  = $row.Geboortedatum
    $birthDate     = ConvertFrom-HitBirthDate -RawValue $birthDateRaw
    $leeftijdValue = ''

    if ($null -ne $birthDate) {
        $ageAtStart  = Get-AgeAtDate -BirthDate $birthDate -ReferenceDate $campStart
        $hasBirthday = Get-BirthdayDuringCamp -BirthDate $birthDate -CampStart $campStart -CampEnd $campEnd

        if ($hasBirthday) {
            $leeftijdValue = "$ageAtStart/$($ageAtStart + 1)"
        }
        else {
            $leeftijdValue = "$ageAtStart"
        }
    }
    else {
        Write-Warning "Kon geboortedatum niet verwerken voor '$($row.Voornaam) $($row.Achternaam)': '$birthDateRaw'"
    }

    # --- Bijzonderheden: dieet + aandachtspunten ---
    $bijzonderhedenParts = [System.Collections.Generic.List[string]]::new()

    $dieetValue = if ('Dieet' -in $actualColumns) { $row.Dieet } else { $null }
    if (-not [string]::IsNullOrWhiteSpace($dieetValue)) {
        $bijzonderhedenParts.Add("Dieet: $($dieetValue.ToString().Trim())")
    }

    $aandachtValue = if ('Aandachtspunten' -in $actualColumns) { $row.Aandachtspunten } else { $null }
    if (-not [string]::IsNullOrWhiteSpace($aandachtValue)) {
        $bijzonderhedenParts.Add($aandachtValue.ToString().Trim())
    }

    # --- Filter: alleen rijen met minstens één bijzonderheid ---
    if ($bijzonderhedenParts.Count -eq 0) {
        continue
    }

    $bijzonderheden = $bijzonderhedenParts -join ' | '

    $outputRows.Add([PSCustomObject]@{
        Voornaam       = $row.Voornaam
        Achternaam     = $row.Achternaam
        Geslacht       = $row.Gender
        Leeftijd       = $leeftijdValue
        Bijzonderheden = $bijzonderheden
    })
}

Write-Verbose "$($outputRows.Count) deelnemer(s) met bijzonderheden (van $totalRows totaal)."

#endregion Build Output Rows

#region Export to Excel

# Bepaal outputbestandsnaam op basis van inputbestandsnaam
$cleanBaseName = Get-HitOutputBaseName -InputPath $resolvedInputPath

$outputFileName = "Deelnemerslijst_${cleanBaseName}_${Year}_Bijzonderheden.xlsx"
$outputDir      = [System.IO.Path]::GetDirectoryName($resolvedInputPath)
$outputPath     = Join-Path -Path $outputDir -ChildPath $outputFileName

Write-Verbose "Exporteren naar: $outputPath"

$excelParams = @{
    Path          = $outputPath
    WorksheetName = 'Bijzonderheden'
    TableName     = 'Deelnemers'
    TableStyle    = 'Medium2'
    AutoSize      = $true
    ClearSheet    = $true
    PassThru      = $true
}

$excelPackage = $outputRows | Export-Excel @excelParams
$worksheet    = $excelPackage.Workbook.Worksheets['Bijzonderheden']

# Forceer de kolom 'Leeftijd' (4e kolom, rijen 2 t/m einde) als tekst.
# EPPlus slaat numerieke strings anders op als getal; herschrijf ze expliciet als string.
$leeftijdColIndex = 4
$lastRow = $worksheet.Dimension.End.Row
for ($rowIdx = 2; $rowIdx -le $lastRow; $rowIdx++) {
    $cell = $worksheet.Cells[$rowIdx, $leeftijdColIndex]
    $cellText = $cell.Text
    $cell.Style.Numberformat.Format = '@'   # Eerst format instellen op tekst
    $cell.Value = $cellText                  # Dan waarde als string toewijzen
}

Close-ExcelPackage -ExcelPackage $excelPackage -Show:$false

Write-Verbose "Export voltooid."

Get-Item -Path $outputPath

#endregion Export to Excel
