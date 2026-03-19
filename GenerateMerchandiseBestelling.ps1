#Requires -Version 5.1
<#
.SYNOPSIS
    Genereert een merchandise-bestellingsoverzicht voor HIT Sail Fryslân en Zeilzwerf Fryslân.

.DESCRIPTION
    Vraagt de gebruiker twee CSV-bestanden te selecteren: één voor HIT Sail Fryslân en één voor
    Zeilzwerf Fryslân. Vervolgens worden de hoodies en t-shirts per evenement gegroepeerd,
    gesorteerd op itemtype (Hoodie → T-shirt) en maat (XS → XXXL), en op de console afgedrukt.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'HitHelpers.psm1') -Force

#region Private helpers

function Get-MerchandiseItems {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$CsvPath
    )

    $rows = Import-Csv -Path $CsvPath -Encoding UTF8

    foreach ($row in $rows) {
        $fullName = $row.'Wat is je voor- en achternaam?'
        if ([string]::IsNullOrWhiteSpace($fullName)) {
            continue
        }

        $firstName = ($fullName.Trim() -split '\s+')[0]

        $hoodieSize = ($row.'Hoodie € 39,95').Trim()
        if (-not [string]::IsNullOrWhiteSpace($hoodieSize)) {
            [PSCustomObject]@{
                Type     = 'Hoodie'
                Size     = $hoodieSize.ToUpper()
                Voornaam = $firstName
            }
        }

        $tshirtSize = ($row.'T-shirt € 39,95').Trim()
        if (-not [string]::IsNullOrWhiteSpace($tshirtSize)) {
            [PSCustomObject]@{
                Type     = 'T-shirt'
                Size     = $tshirtSize.ToUpper()
                Voornaam = $firstName
            }
        }
    }
}

function Write-MerchandiseOverzicht {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$EventName,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [PSCustomObject[]]$Items
    )

    $typeOrder = @('Hoodie', 'T-shirt')
    $sizeOrder = @('XS', 'S', 'M', 'L', 'XL', 'XXL', 'XXXL')

    # Add sort-key properties explicitly to avoid scriptblock scoping issues in PS 5.1
    $withKeys = foreach ($item in $Items) {
        $typeIdx = [array]::IndexOf($typeOrder, $item.Type)
        $sizeIdx = [array]::IndexOf($sizeOrder, $item.Size)
        [PSCustomObject]@{
            TypeIdx  = if ($typeIdx -lt 0) { 999 } else { $typeIdx }
            SizeIdx  = if ($sizeIdx -lt 0) { 999 } else { $sizeIdx }
            Type     = $item.Type
            Size     = $item.Size
            Voornaam = $item.Voornaam
        }
    }
    $sortedItems = @($withKeys | Sort-Object -Property TypeIdx, SizeIdx)

    # Group-Object sorts groups alphabetically and loses the sort order, so group manually
    # using an ordered dictionary to preserve the sorted sequence.
    $groups = [ordered]@{}
    foreach ($item in $sortedItems) {
        $key = "$($item.Type)|$($item.Size)"
        if (-not $groups.Contains($key)) {
            $groups[$key] = [System.Collections.Generic.List[PSCustomObject]]::new()
        }
        $groups[$key].Add($item)
    }

    Write-Host $EventName

    $isFirstGroup = $true
    foreach ($key in $groups.Keys) {
        if (-not $isFirstGroup) {
            Write-Host ''
        }
        $isFirstGroup = $false

        $members   = $groups[$key]
        $firstItem = $members[0]
        Write-Host "$($members.Count)x $($firstItem.Type) $($firstItem.Size)"

        foreach ($item in $members) {
            Write-Host "- $($item.Voornaam)"
        }
    }
}

#endregion

#region Main

$hitCsvPath = Select-HitFilePath `
    -Prompt 'Selecteer het CSV-bestand voor HIT Sail Fryslân' `
    -ScriptDir $PSScriptRoot

$zeilzwerfCsvPath = Select-HitFilePath `
    -Prompt 'Selecteer het CSV-bestand voor Zeilzwerf Fryslân' `
    -ScriptDir $PSScriptRoot `
    -ExcludePaths @($hitCsvPath)

$hitItems       = @(Get-MerchandiseItems -CsvPath $hitCsvPath)
$zeilzwerfItems = @(Get-MerchandiseItems -CsvPath $zeilzwerfCsvPath)

$hitEventName       = [System.IO.Path]::GetFileNameWithoutExtension($hitCsvPath)
$zeilzwerfEventName = [System.IO.Path]::GetFileNameWithoutExtension($zeilzwerfCsvPath)

Write-Host ''
Write-MerchandiseOverzicht -EventName $hitEventName -Items $hitItems

Write-Host ''
Write-MerchandiseOverzicht -EventName $zeilzwerfEventName -Items $zeilzwerfItems
Write-Host ''

#endregion
