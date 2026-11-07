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

# --------------------------------------------------------------------------
# Voortgang. De PowerShell-runner toont dit als blauwe balk met percentage;
# los gedraaid zie je de gewone voortgangsbalk van PowerShell. Hooguit vijf
# keer per seconde bijwerken: bij tienduizenden bestanden kost elke melding
# anders merkbaar tijd.
# --------------------------------------------------------------------------
$script:voortgangKlok = [System.Diagnostics.Stopwatch]::StartNew()
$script:voortgangEerste = $true

function Meld-Voortgang {
    <# Totaal 0 = nog onbekend: dan alleen een teller, zonder percentage. #>
    param([string]$Wat, [long]$Nu, [long]$Totaal = 0, [string]$Status = "", [string]$Detail = "")
    $laatste = $Totaal -gt 0 -and $Nu -ge $Totaal
    if (-not ($script:voortgangEerste -or $laatste) -and
        $script:voortgangKlok.ElapsedMilliseconds -lt 200) { return }
    $script:voortgangEerste = $false
    $script:voortgangKlok.Restart()
    if (-not $Status) {
        $Status = if ($Totaal -gt 0) { "{0:N0} van {1:N0}" -f $Nu, $Totaal } else { "{0:N0} gevonden" -f $Nu }
    }
    $melding = @{ Activity = $Wat; Status = $Status }
    if ($Detail) { $melding.CurrentOperation = $Detail }
    if ($Totaal -gt 0) { $melding.PercentComplete = [Math]::Min(100, [int][Math]::Floor($Nu * 100.0 / $Totaal)) }
    Write-Progress @melding
}

function Stop-Voortgang {
    param([string]$Wat)
    $script:voortgangEerste = $true
    Write-Progress -Activity $Wat -Completed
}

function Tel-Onderweg {
    <# Telt mee terwijl Get-ChildItem nog leest; geeft alles ongewijzigd door.
       Verwacht > 0 geeft een percentage, als het aantal ongeveer bekend is. #>
    param([Parameter(ValueFromPipeline)]$Item, [string]$Wat, [long]$Verwacht = 0)
    begin { $n = 0 }
    process {
        $n++
        if ($Verwacht -gt 0) {
            Meld-Voortgang -Wat $Wat -Nu ([Math]::Min($n, $Verwacht - 1)) -Totaal $Verwacht `
                -Status ("{0:N0} gevonden" -f $n)
        }
        else { Meld-Voortgang -Wat $Wat -Nu $n }
        $Item
    }
}

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
$Nummers = @($Nummers)
$mapNr = 0
# Voortgang in duizendsten per map, zodat PerBestand binnen een map kan
# doorlopen. Robocopy meldt niets tussendoor; daar springt de balk per map.
$stappen = [long]$Nummers.Count * 1000

foreach ($naam in $Nummers) {
    $mapNr++
    $bron = Join-Path $Bronmap $naam
    $doel = Join-Path $Doelmap $naam
    Meld-Voortgang -Wat "Mappen kopiëren" -Nu (($mapNr - 1) * 1000) -Totaal $stappen `
        -Status "map $mapNr van $($Nummers.Count): $naam"

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
    $bestanden = @(Get-ChildItem -LiteralPath $bron -Recurse -File -ErrorAction SilentlyContinue)
    $inDezeMap = 0
    $bestandNr = 0

    foreach ($bestand in $bestanden) {
        $bestandNr++
        Meld-Voortgang -Wat "Mappen kopiëren" `
            -Nu (($mapNr - 1) * 1000 + [long][Math]::Floor($bestandNr * 1000.0 / $bestanden.Count)) `
            -Totaal $stappen -Detail $bestand.Name `
            -Status "map $mapNr van $($Nummers.Count): $naam · bestand $bestandNr van $($bestanden.Count)"
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
Stop-Voortgang -Wat "Mappen kopiëren"

Write-Host ""
Write-Host "Klaar. $gekopieerd mappen verwerkt, $overgeslagenMappen niet gevonden, $fouten fouten." `
    -ForegroundColor Green
if ($fouten -gt 0) { exit 1 }
