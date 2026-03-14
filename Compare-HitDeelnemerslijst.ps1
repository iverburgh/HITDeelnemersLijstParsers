<#
    .SYNOPSIS
    Vergelijkt twee aanmeldingslijsten op naam en geboortedatum.

    .DESCRIPTION
    Laadt twee deelnemerslijsten (xlsx of csv) en zoekt naar deelnemers die in beide
    lijsten voorkomen op basis van volledige naam (voornaam + tussenvoegsel + achternaam)
    en geboortedatum. De gebruiker selecteert beide bestanden via een console-keuzemenu.

    .NOTES
    Ondersteunde kolomformaten:
    - Raw CSV (puntkomma-gescheiden) met kolommen:
        Lid voornaam, Lid tussenvoegsel, Lid achternaam, Lid geboortedatum
    - Excel (.xlsx) met hernoemde kolommen:
        Voornaam, Achternaam, Geboortedatum (tussenvoegsel optioneel)
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module -Name (Join-Path $PSScriptRoot 'HitHelpers.psm1') -Force
Assert-HitImportExcel

# ─────────────────────────────────────────────────────────────────────────────
# Hulpfunctie: laad een bestand (xlsx of csv) als array van rijen
# ─────────────────────────────────────────────────────────────────────────────
function Import-ParticipantFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        '.csv'  { return @(Import-Csv -Path $Path -Delimiter ';' -ErrorAction Stop) }
        '.xlsx' { return @(Import-Excel -Path $Path -ErrorAction Stop) }
        default { throw "Niet-ondersteund bestandsformaat: $extension" }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# Hulpfunctie: normaliseer een rij naar een PSCustomObject met vergelijksleutel
# ─────────────────────────────────────────────────────────────────────────────
function ConvertTo-NormalizedParticipant {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Row
    )

    $columnNames = $Row.PSObject.Properties.Name

    # Detecteer het kolomformaat: raw CSV vs. hernoemde xlsx
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
        Write-Warning "Onbekend kolomformaat — rij overgeslagen."
        return $null
    }

    # Bouw volledige naam met genormaliseerde spaties
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

# ─────────────────────────────────────────────────────────────────────────────
# Bestandsselectie
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host '  HIT Deelnemerslijst Vergelijker' -ForegroundColor Cyan
Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host ''

$list1Path = Select-HitFilePath -Prompt 'Kies deelnemerslijst 1 (vorig jaar)' -ScriptDir $PSScriptRoot
$list2Path = Select-HitFilePath -Prompt 'Kies deelnemerslijst 2 (huidig jaar)' -ScriptDir $PSScriptRoot

# ─────────────────────────────────────────────────────────────────────────────
# Inladen en normaliseren
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host 'Bestanden worden ingeladen...' -ForegroundColor DarkGray

$rawRows1 = Import-ParticipantFile -Path $list1Path
$rawRows2 = Import-ParticipantFile -Path $list2Path

$participants1 = @(
    $rawRows1 | ForEach-Object { ConvertTo-NormalizedParticipant -Row $_ } | Where-Object { $null -ne $_ }
)
$participants2 = @(
    $rawRows2 | ForEach-Object { ConvertTo-NormalizedParticipant -Row $_ } | Where-Object { $null -ne $_ }
)

# ─────────────────────────────────────────────────────────────────────────────
# Opzoektabel van lijst 1
# ─────────────────────────────────────────────────────────────────────────────
$lookup = @{}
foreach ($participant in $participants1) {
    $lookup[$participant.Sleutel] = $true
}

# ─────────────────────────────────────────────────────────────────────────────
# Vergelijking
# ─────────────────────────────────────────────────────────────────────────────
$matchedParticipants = @($participants2 | Where-Object { $lookup.ContainsKey($_.Sleutel) })

# ─────────────────────────────────────────────────────────────────────────────
# Output
# ─────────────────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host '  Resultaat' -ForegroundColor Cyan
Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host ''
Write-Host "Deelnemerslijst 1: $($participants1.Count) deelnemer(s)" -ForegroundColor White
Write-Host "  $list1Path" -ForegroundColor DarkGray
Write-Host "Deelnemerslijst 2: $($participants2.Count) deelnemer(s)" -ForegroundColor White
Write-Host "  $list2Path" -ForegroundColor DarkGray
Write-Host ''

if ($matchedParticipants.Count -eq 0) {
    Write-Host 'Er zijn geen deelnemers gevonden die in beide lijsten voorkomen.' -ForegroundColor Yellow
}
else {
    Write-Host "In deelnemerslijst 2 staan $($matchedParticipants.Count) deelnemer(s) die ook in deelnemerslijst 1 staan:" -ForegroundColor Green
    Write-Host ''

    $matchedParticipants |
        Select-Object -Property VolledigeNaam,
            @{
                Name       = 'Geboortedatum'
                Expression = {
                    if ($null -ne $_.Geboortedatum) {
                        $_.Geboortedatum.ToString('dd-MM-yyyy')
                    }
                    else {
                        '(onbekend)'
                    }
                }
            } |
        Format-Table -AutoSize |
        Out-Host
}

Write-Host '══════════════════════════════════════════════════' -ForegroundColor DarkCyan
Write-Host ''
