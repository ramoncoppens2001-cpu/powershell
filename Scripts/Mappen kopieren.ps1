<#
.SYNOPSIS
Kopieert dossiermappen van de ene locatie naar de andere, submappen behouden.

.DESCRIPTION
Geef een bronmap en een doelmap. Vul je bij Nummers een lijst met mapnamen in,
dan worden alleen die submappen gekopieerd; laat je dat leeg, dan gaan alle
submappen mee.

Voor het kopiëren zelf zijn er twee wegen:

  Robocopy      Windows' eigen kopieergereedschap. Veel sneller over
                netwerkschijven als T:\ en Z:\, hervat na een storing, en
                slaat bestanden over die al gelijk zijn. Kan geen duplicaten
                omnoemen.

  PerBestand    Bestand voor bestand, zoals de oude versie. Trager, maar kan
                een bestaand bestand laten staan en de nieuwe ernaast zetten
                als naam_copy1.

Bestaat een bronmap niet, dan wordt die overgeslagen met een melding; de rest
gaat gewoon door.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, HelpMessage = "Map waarin de te kopiëren submappen staan")]
    [string]$Bronmap,

    [Parameter(Mandatory, HelpMessage = "Map waar de kopieën terechtkomen")]
    [string]$Doelmap,

    [Parameter(HelpMessage = "Mapnamen, één per regel. Leeg laten voor alle submappen.")]
    [string[]]$Nummers = @(),

    [Parameter(HelpMessage = "Robocopy is sneller; PerBestand kan duplicaten omnoemen")]
    [ValidateSet("Robocopy", "PerBestand")]
    [string]$Methode = "Robocopy",

    [Parameter(HelpMessage = "Wat te doen met een bestand dat al bestaat")]
    [ValidateSet("Overslaan", "Overschrijven", "NieuweNaam")]
    [string]$Bestaande = "Overslaan"
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Bronmap -PathType Container)) {
    Write-Host "Bronmap bestaat niet: $Bronmap" -ForegroundColor Red
    exit 1
}

# NieuweNaam kan robocopy niet; dan is er maar één werkbare weg.
if ($Bestaande -eq "NieuweNaam" -and $Methode -eq "Robocopy") {
    Write-Host "Robocopy kan geen duplicaten omnoemen. Overgeschakeld op PerBestand." `
        -ForegroundColor Yellow
    $Methode = "PerBestand"
}

if (-not $Nummers -or $Nummers.Count -eq 0) {
    $Nummers = Get-ChildItem -LiteralPath $Bronmap -Directory | Select-Object -ExpandProperty Name
    Write-Host "Geen lijst opgegeven, alle $($Nummers.Count) submappen worden gekopieerd."
}

# Lege regels en spaties uit de lijst halen; die komen er bij plakken makkelijk in
$Nummers = $Nummers | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }

Write-Host "Van    : $Bronmap"
Write-Host "Naar   : $Doelmap"
Write-Host "Mappen : $($Nummers.Count)"
Write-Host "Methode: $Methode, bestaande bestanden: $Bestaande"
Write-Host ""

$gekopieerd = 0
$overgeslagenMappen = 0
$fouten = 0

foreach ($naam in $Nummers) {
    $bron = Join-Path $Bronmap $naam
    $doel = Join-Path $Doelmap $naam

    if (-not (Test-Path -LiteralPath $bron -PathType Container)) {
        Write-Host "BRON NIET GEVONDEN, overgeslagen: $bron" -ForegroundColor Yellow
        $overgeslagenMappen++
        continue
    }

    if (-not $PSCmdlet.ShouldProcess($doel, "kopiëren van $bron")) { continue }

    Write-Host "[$naam]" -ForegroundColor Cyan

    if ($Methode -eq "Robocopy") {
        # /E submappen incl. lege, /NFL /NDL /NJH /NJS beperken de ruis,
        # /R:2 /W:2 blijven niet minutenlang hangen op een vergrendeld bestand.
        $opties = @("/E", "/R:2", "/W:2", "/NFL", "/NDL", "/NJH", "/NJS", "/NP")
        if ($Bestaande -eq "Overslaan") { $opties += "/XC"; $opties += "/XN"; $opties += "/XO" }

        & robocopy $bron $doel *.* @opties | Out-Null
        $code = $LASTEXITCODE

        # Robocopy: 0 = niets te doen, 1 = gekopieerd, 2/3 = extra bestanden,
        # 8 en hoger = echte fout.
        if ($code -ge 8) {
            Write-Host "  MISLUKT, robocopy code $code" -ForegroundColor Red
            $fouten++
        }
        else {
            $aantal = (Get-ChildItem -LiteralPath $doel -Recurse -File -ErrorAction SilentlyContinue).Count
            Write-Host "  klaar, $aantal bestanden in de doelmap"
            $gekopieerd++
        }
        continue
    }

    # ---- PerBestand ----
    $lengte = $bron.TrimEnd('\').Length + 1
    $bestanden = Get-ChildItem -LiteralPath $bron -Recurse -File -ErrorAction SilentlyContinue
    $inDezeMap = 0

    foreach ($bestand in $bestanden) {
        $relatief = $bestand.FullName.Substring($lengte)
        $doelpad = Join-Path $doel $relatief
        $doelmapVanBestand = Split-Path -Parent $doelpad

        if (-not (Test-Path -LiteralPath $doelmapVanBestand)) {
            New-Item -ItemType Directory -Path $doelmapVanBestand -Force | Out-Null
        }

        if (Test-Path -LiteralPath $doelpad) {
            switch ($Bestaande) {
                "Overslaan" {
                    continue
                }
                "NieuweNaam" {
                    $kern = [System.IO.Path]::GetFileNameWithoutExtension($bestand.Name)
                    $ext = $bestand.Extension
                    $teller = 1
                    do {
                        $nieuw = "${kern}_copy${teller}${ext}"
                        $doelpad = Join-Path $doelmapVanBestand $nieuw
                        $teller++
                    } while (Test-Path -LiteralPath $doelpad)
                    Write-Host "  bestaat al, kopie als: $nieuw" -ForegroundColor Cyan
                }
            }
        }

        try {
            Copy-Item -LiteralPath $bestand.FullName -Destination $doelpad -Force:($Bestaande -eq "Overschrijven")
            $inDezeMap++
        }
        catch {
            Write-Host "  MISLUKT: $($bestand.FullName)" -ForegroundColor Red
            Write-Host "           $($_.Exception.Message)" -ForegroundColor Red
            $fouten++
        }
    }

    Write-Host "  $inDezeMap bestanden gekopieerd"
    $gekopieerd++
}

Write-Host ""
Write-Host "Klaar. $gekopieerd mappen verwerkt, $overgeslagenMappen niet gevonden, $fouten fouten." `
    -ForegroundColor Green
if ($fouten -gt 0) { exit 1 }
