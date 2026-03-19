<#
    .SYNOPSIS
    Exporteert een contactgegevenslijst van HIT-deelnemers naar een Excel-bestand.

    .DESCRIPTION
    Leest een HIT-registratie Excel-bestand (.xlsx, zoals geëxporteerd uit het
    Scouting-aanmeldingssysteem) en exporteert een nieuw Excel-bestand met de kolommen:
    Voornaam, Achternaam, Geslacht, Geboortedatum, Contactpersoon, Noodnummer, Lid mobiel,
    Bijzonderheden.

    De Bijzonderheden-kolom bevat:
    - De verjaardagsdatum (bijv. "zaterdag 3 april jarig (wordt 15)") als een deelnemer
      jarig is tijdens het kamp (Goede Vrijdag t/m Tweede Paasdag).
    - "Was er vorig jaar ook" als de deelnemer ook in een optioneel vorigjaar-bestand staat.
    - Beide gecombineerd met " - " als allebei van toepassing zijn.

    Alle deelnemers worden opgenomen (geen filter).

    Het output-bestand wordt opgeslagen in dezelfde map als het script, met de naam:
    Deelnemerslijst_[basisnaam]_[jaar]_Contactgegevens.xlsx

    Het script zoekt automatisch naar een bestand met het patroon "*-alles.xlsx" in de
    scriptmap. Als er meerdere zijn, verschijnt een keuzemenu. Als er geen zijn, wordt
    gezocht op "*.xlsx". Als er dan ook geen is, volgt een foutmelding.

    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is.

    .PARAMETER Year
    Het jaar dat wordt opgenomen in de bestandsnaam van het output-bestand en wordt gebruikt
    voor de kampdatumberekening (Bijzonderheden-kolom). Standaard: het huidige jaar.

    .EXAMPLE
    .\Export-HitContactgegevens.ps1

    Verwerkt het gevonden *-alles.xlsx-bestand in de scriptmap voor het huidige jaar.

    .EXAMPLE
    .\Export-HitContactgegevens.ps1 -Verbose

    Zoals hierboven, maar met uitgebreide voortgangsberichten.

    .OUTPUTS
    System.IO.FileInfo
    Het pad naar het gemaakte Excel-bestand met contactgegevens.

    .NOTES
    Vereist: PowerShell 5.1+, ImportExcel-module.
    De module ImportExcel wordt automatisch geïnstalleerd als die nog niet aanwezig is
    (via Install-Module -Scope CurrentUser). Hiervoor is een internetverbinding nodig.
#>
[CmdletBinding()]
param(
    [Parameter(HelpMessage = 'Jaar dat in de outputbestandsnaam wordt gebruikt (standaard: huidig jaar).')]
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

$resolvedInputPath = Select-HitFilePath `
    -Prompt 'Kies huidigjaar deelnemerslijst' `
    -ScriptDir $scriptDir

Write-Verbose "Input-bestand: $resolvedInputPath"

#endregion Initialisation

#region Camp Date Calculation

$campDates = Get-HitCampDates -Year $Year
$campStart  = $campDates.CampStart
$campEnd    = $campDates.CampEnd

#endregion Camp Date Calculation

#region Vorigjaar Selectie

$vorigJaarPath = Select-HitFilePath `
    -Prompt 'Kies vorigjaar-deelnemerslijst (0 of Enter = overslaan)' `
    -ScriptDir $scriptDir `
    -ExcludePaths @($resolvedInputPath) `
    -AllowSkip

if ($null -ne $vorigJaarPath) {
    Write-Verbose "Vorigjaar-bestand: $vorigJaarPath"
    $vorigJaarRijen      = Import-ParticipantFile -Path $vorigJaarPath
    $vorigJaarDeelnemers = @(
        $vorigJaarRijen | ForEach-Object { ConvertTo-NormalizedParticipant -Row $_ } | Where-Object { $null -ne $_ }
    )
    $vorigJaarLookup = @{}
    foreach ($vd in $vorigJaarDeelnemers) {
        $vorigJaarLookup[$vd.Sleutel] = $true
    }
    Write-Verbose "$($vorigJaarDeelnemers.Count) deelnemer(s) geladen uit vorigjaar-bestand."
}
else {
    $vorigJaarLookup = $null
    Write-Verbose 'Geen vorigjaar-bestand geselecteerd; vorig-jaar-check overgeslagen.'
}

#endregion Vorigjaar Selectie

#region Import Data

Write-Verbose "Inlezen van Excel-bestand..."
$allRows   = Import-Excel -Path $resolvedInputPath -ErrorAction Stop
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
    # Geboortedatum: zorg voor een datetime-object
    $birthDateRaw = $row.Geboortedatum
    $birthDate    = ConvertFrom-HitBirthDate -RawValue $birthDateRaw

    if ($null -eq $birthDate) {
        Write-Warning "Kon geboortedatum niet verwerken voor '$($row.Voornaam) $($row.Achternaam)': '$birthDateRaw'"
    }

    # --- Bijzonderheden ---
    $bijzonderhedenParts = [System.Collections.Generic.List[string]]::new()

    # Verjaardag tijdens het kamp?
    if ($null -ne $birthDate) {
        $bdDate = Get-BirthdayDateDuringCamp -BirthDate $birthDate -CampStart $campStart -CampEnd $campEnd
        if ($null -ne $bdDate) {
            $dayName   = Get-DutchDayName   -DayOfWeek ([int]$bdDate.DayOfWeek)
            $monthName = Get-DutchMonthName -Month $bdDate.Month
            $newAge    = Get-AgeAtDate -BirthDate $birthDate -ReferenceDate $bdDate
            $bijzonderhedenParts.Add("$dayName $($bdDate.Day) $monthName jarig (wordt $newAge)")
        }
    }

    # Was er vorig jaar ook?
    if ($null -ne $vorigJaarLookup) {
        $normalized = ConvertTo-NormalizedParticipant -Row $row
        if ($null -ne $normalized -and $vorigJaarLookup.ContainsKey($normalized.Sleutel)) {
            $bijzonderhedenParts.Add('Was er vorig jaar ook')
        }
    }

    $bijzonderheden = $bijzonderhedenParts -join ' - '

    $groep = if ($actualColumns -contains 'Subgroep')     { [string]$row.'Subgroep' }
             elseif ($actualColumns -contains 'Groep')    { [string]$row.'Groep' }
             elseif ($actualColumns -contains 'Subgroepnaam') { [string]$row.'Subgroepnaam' }
             else { '' }

    $outputRows.Add([PSCustomObject]@{
        Groep           = $groep
        Voornaam        = $row.Voornaam
        Achternaam      = $row.Achternaam
        Geslacht        = $row.Gender
        Geboortedatum   = if ($null -ne $birthDate) { $birthDate.ToString('dd-MM-yyyy') } else { '' }
        Contactpersoon  = if ('Naam noodcontact' -in $actualColumns) { $row.'Naam noodcontact' } else { '' }
        Noodnummer      = if ('Telefoonnummer noodcontact' -in $actualColumns) { $row.'Telefoonnummer noodcontact' } else { '' }
        'Lid mobiel'    = if ('Mobiel' -in $actualColumns) { $row.Mobiel } else { '' }
        Bijzonderheden  = $bijzonderheden
    })
}

Write-Verbose "$($outputRows.Count) deelnemer(s) verwerkt."

# Sorteren op Voornaam, dan Achternaam
$outputRows = [System.Collections.Generic.List[PSCustomObject]](
    $outputRows | Sort-Object Voornaam, Achternaam
)

#endregion Build Output Rows

#region Export to Excel

# Bepaal outputbestandsnaam op basis van inputbestandsnaam
$cleanBaseName = Get-HitOutputBaseName -InputPath $resolvedInputPath

$outputFileName = "Deelnemerslijst_${cleanBaseName}_${Year}_Contactgegevens.xlsx"
$outputDir      = [System.IO.Path]::GetDirectoryName($resolvedInputPath)
$outputPath     = Join-Path -Path $outputDir -ChildPath $outputFileName

Write-Verbose "Exporteren naar: $outputPath"

$excelParams = @{
    Path          = $outputPath
    WorksheetName = 'Contactgegevens'
    TableName     = 'Deelnemers'
    TableStyle    = 'Medium2'
    AutoSize      = $true
    ClearSheet    = $true
    PassThru      = $true
}

$excelPackage = $outputRows | Export-Excel @excelParams
$worksheet    = $excelPackage.Workbook.Worksheets['Contactgegevens']

# Schrijf tekstvelden opnieuw met expliciete '@' celstijl vóór de waarde.
# EPPlus converteert numeriek-uitziende strings (zoals 0612345678) automatisch naar getallen
# waardoor voorloopnullen verdwijnen. Door '@' in te stellen vóór het schrijven van de waarde
# en de waarde direct uit $outputRows te lezen (niet via $cell.Text), blijft de string intact.
$lastRow = $worksheet.Dimension.End.Row
for ($rowIdx = 2; $rowIdx -le $lastRow; $rowIdx++) {
    $sourceRow = $outputRows[$rowIdx - 2]

    # Geboortedatum (kolom 4)
    $cell = $worksheet.Cells[$rowIdx, 4]
    $cell.Style.Numberformat.Format = '@'
    $cell.Value = $sourceRow.Geboortedatum

    # Noodnummer (kolom 6)
    $cell = $worksheet.Cells[$rowIdx, 6]
    $cell.Style.Numberformat.Format = '@'
    $cell.Value = $sourceRow.Noodnummer

    # Lid mobiel (kolom 7)
    $cell = $worksheet.Cells[$rowIdx, 7]
    $cell.Style.Numberformat.Format = '@'
    $cell.Value = $sourceRow.'Lid mobiel'
}

Close-ExcelPackage -ExcelPackage $excelPackage -Show:$false

Write-Verbose "Export voltooid."

Get-Item -Path $outputPath

#endregion Export to Excel
