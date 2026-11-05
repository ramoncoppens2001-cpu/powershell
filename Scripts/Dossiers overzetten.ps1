<#
.SYNOPSIS
Zet de dossierstukken van vorig jaar (foto's, kadaster, huur, meetstaten) over naar de mappen van dit jaar, in een mapje per jaar.

.DESCRIPTION
Per dossiernummer (bijv. 07917) wordt gekeken welke mappen er vorig jaar
waren, en de inhoud gaat naar de bijbehorende map van dit jaar, in een
submapje met de naam van het jaar:

  2025\07917\09.0 Foto's\...          ->  2026\07917\1. Inspectie Formulier & Foto's\2025\...
  2025\07917\04.0 Kadaster - Aktes    ->  2026\07917\2. Kadaster\2025\
  enzovoort.

Wat er meegaat staat in de tabel KOPPELINGEN hieronder. Rapporten,
berekeningen, checklists, opdracht en correspondentie gaan niet mee.

Een dossier dat alleen in vorig jaar bestaat, wordt overgeslagen met een
melding. Lege mappen worden overgeslagen, zodat er geen lege 2025-mapjes
ontstaan. Bestanden die al in het doelmapje staan worden niet overschreven;
je kunt het script dus veilig nog een keer draaien.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, HelpMessage = "Projectmap van vorig jaar, bijv. T:\Dossiers\2025\Heitman\PERL\Eramus I")]
    [string]$Bronmap,

    [Parameter(Mandatory, HelpMessage = "Projectmap van dit jaar, bijv. T:\Dossiers\2026\Heitman\PERL\Erasmus I")]
    [string]$Doelmap,

    [Parameter(HelpMessage = "Naam van het mapje waarin de bestanden van vorig jaar komen")]
    [string]$Jaar = "2025",

    [Parameter(HelpMessage = "Alleen deze dossiernummers, een per regel. Leeg laten voor alle dossiers.")]
    [string[]]$Nummers = @()
)

# ==========================================================================
# KOPPELINGEN: welke map van vorig jaar naar welke map van dit jaar.
#
# Er wordt gezocht op een stukje van de mapnaam (* = wat dan ook), zodat
# "09.0 Foto's" en "1. Inspectie Formulier & Foto's" elkaar vinden zonder
# dat de nummering hoeft te kloppen.
#
# Extra = optioneel mapje binnen het jaarmapje. Leeg = direct in het jaarmapje.
#
# Wil je iets extra meenemen, voeg dan een regel toe. Bijvoorbeeld:
#   @{ Bron = "*Correspondentie*"; Doel = "*Diversen*"; Extra = "Correspondentie" }
# ==========================================================================
$Koppelingen = @(
    @{ Bron = "*Foto*";       Doel = "*Foto*";       Extra = "" }
    @{ Bron = "*Inspectie*";  Doel = "*Foto*";       Extra = "Inspectieformulier" }
    @{ Bron = "*Kadaster*";   Doel = "*Kadaster*";   Extra = "" }
    @{ Bron = "*Bestemming*"; Doel = "*Bestemming*"; Extra = "" }
    @{ Bron = "*Meetstat*";   Doel = "*Meetstat*";   Extra = "" }
    @{ Bron = "*Huur*";       Doel = "*Huur*";       Extra = "" }
    @{ Bron = "*Diversen*";   Doel = "*Diversen*";   Extra = "" }
)

# Bestanden die nooit mee hoeven: Windows-miniaturen en Office-slotbestanden.
$NietMee = @('Thumbs.db', 'desktop.ini', '~$*')

# ==========================================================================

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = "Stop"

foreach ($pad in @($Bronmap, $Doelmap)) {
    if (-not (Test-Path -LiteralPath $pad -PathType Container)) {
        Write-Host "Map niet gevonden: $pad" -ForegroundColor Red
        exit 1
    }
}
if ((Resolve-Path -LiteralPath $Bronmap).Path -eq (Resolve-Path -LiteralPath $Doelmap).Path) {
    Write-Host "Bronmap en doelmap zijn dezelfde map." -ForegroundColor Red
    exit 1
}

$Jaar = $Jaar.Trim()
if ($Jaar -eq "") {
    Write-Host "Vul bij Jaar de naam van het mapje in, bijvoorbeeld 2025." -ForegroundColor Red
    exit 1
}

function Zoek-Map([string]$ouder, [string]$patroon) {
    @(Get-ChildItem -LiteralPath $ouder -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like $patroon })
}

function Als-MB([long]$bytes) {
    "{0:N0} MB" -f ($bytes / 1MB)
}

# Dossiermappen herkennen aan een naam van alleen cijfers (07917 enz.).
# De algemene mappen (01.1 Algemeen Secretariaat enz.) vallen daar vanzelf buiten.
$Nummers = @($Nummers | ForEach-Object { $_ -split '[,;\s]+' } | Where-Object { $_ -ne "" })
if ($Nummers.Count -eq 0) {
    $Nummers = @(Get-ChildItem -LiteralPath $Bronmap -Directory |
        Where-Object { $_.Name -match '^\d+$' } |
        Sort-Object Name | Select-Object -ExpandProperty Name)
}

Write-Host "Van      : $Bronmap"
Write-Host "Naar     : $Doelmap"
Write-Host "Mapje    : $Jaar"
Write-Host "Dossiers : $($Nummers.Count)"
Write-Host ""

$verwerkt = 0
$overgeslagen = @()
$aantalBestanden = 0
$totaalBytes = 0
$mislukt = 0
$meldingen = 0

foreach ($nummer in $Nummers) {
    $bronDossier = Join-Path $Bronmap $nummer
    $doelDossier = Join-Path $Doelmap $nummer

    if (-not (Test-Path -LiteralPath $bronDossier -PathType Container)) {
        Write-Host "${nummer}: niet aanwezig in vorig jaar, overgeslagen" -ForegroundColor Yellow
        $overgeslagen += $nummer
        continue
    }
    if (-not (Test-Path -LiteralPath $doelDossier -PathType Container)) {
        Write-Host "${nummer}: geen dossiermap in de doelmap, overgeslagen" -ForegroundColor Yellow
        $overgeslagen += $nummer
        continue
    }

    Write-Host "[$nummer]" -ForegroundColor Cyan
    $gebruikt = @{}

    foreach ($k in $Koppelingen) {
        $bronnen = Zoek-Map $bronDossier $k.Bron
        if ($bronnen.Count -eq 0) { continue }
        if ($bronnen.Count -gt 1) {
            Write-Host "  let op: meerdere mappen passen bij $($k.Bron): $($bronnen.Name -join ', ') - niet gekopieerd" -ForegroundColor Yellow
            $meldingen++
            continue
        }
        $bron = $bronnen[0]
        if ($gebruikt.ContainsKey($bron.FullName)) { continue }
        $gebruikt[$bron.FullName] = $true

        $doelen = Zoek-Map $doelDossier $k.Doel
        if ($doelen.Count -ne 1) {
            $wat = if ($doelen.Count -eq 0) { "geen map" } else { "meerdere mappen" }
            Write-Host "  let op: $wat voor $($k.Doel) in dit jaar, '$($bron.Name)' niet gekopieerd" -ForegroundColor Yellow
            $meldingen++
            continue
        }

        $bestanden = @(Get-ChildItem -LiteralPath $bron.FullName -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $n = $_.Name; -not ($NietMee | Where-Object { $n -like $_ }) })
        if ($bestanden.Count -eq 0) {
            Write-Host "  $($bron.Name): leeg, overgeslagen" -ForegroundColor DarkGray
            continue
        }
        $bytes = ($bestanden | Measure-Object -Property Length -Sum).Sum

        $doel = Join-Path $doelen[0].FullName $Jaar
        $toon = "$($doelen[0].Name)\$Jaar"
        if ($k.Extra) {
            $doel = Join-Path $doel $k.Extra
            $toon += "\$($k.Extra)"
        }

        Write-Host ("  {0}  ->  {1}   ({2} bestanden, {3})" -f $bron.Name, $toon, $bestanden.Count, (Als-MB $bytes))

        if (-not $PSCmdlet.ShouldProcess($doel, "$($bestanden.Count) bestanden kopieren uit $($bron.FullName)")) {
            continue
        }

        # /E submappen meenemen (Summary 1, Summary 2 enz.)
        # /XC /XN /XO bestaande bestanden nooit overschrijven
        # /R:2 /W:2 niet minutenlang blijven hangen op een vergrendeld bestand
        $opties = @("/E", "/XC", "/XN", "/XO", "/R:2", "/W:2",
                    "/NFL", "/NDL", "/NJH", "/NJS", "/NP", "/XF") + $NietMee
        & robocopy $bron.FullName $doel *.* @opties | Out-Null
        $code = $LASTEXITCODE

        if ($code -ge 8) {
            Write-Host "    MISLUKT, robocopy code $code" -ForegroundColor Red
            $mislukt++
        }
        elseif ($code -band 1) {
            $aantalBestanden += $bestanden.Count
            $totaalBytes += $bytes
        }
        else {
            Write-Host "    stond er al, niets nieuws gekopieerd" -ForegroundColor DarkGray
        }
    }
    $verwerkt++
}

Write-Host ""
if ($overgeslagen.Count -gt 0) {
    Write-Host "Overgeslagen dossiers: $($overgeslagen -join ', ')" -ForegroundColor Yellow
}
if ($meldingen -gt 0) {
    Write-Host "$meldingen keer 'let op' hierboven: die mappen zijn niet gekopieerd." -ForegroundColor Yellow
}

if ($WhatIfPreference) {
    Write-Host "Klaar met proefrun. Er is niets gekopieerd; $verwerkt dossiers bekeken." -ForegroundColor Green
}
elseif ($mislukt -eq 0) {
    Write-Host ("Klaar. {0} dossiers verwerkt, {1} bestanden gekopieerd ({2})." -f $verwerkt, $aantalBestanden, (Als-MB $totaalBytes)) -ForegroundColor Green
}
else {
    Write-Host "Klaar, maar $mislukt kopieeracties zijn MISLUKT. Kijk hierboven." -ForegroundColor Red
    exit 1
}
