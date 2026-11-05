<#
.SYNOPSIS
Maakt een overzicht van de mappenstructuur onder een map, met een leesbaar
HTML-rapport. Bestandsnamen komen er niet in, alleen mapnamen.

.DESCRIPTION
De tegenhanger van Map_inventariseren.ps1: die kijkt naar bestanden, deze kijkt
naar de mappen zelf. Je krijgt de boom met mapnamen, hoe diep die gaat, waar de
vertakkingen zitten, welke mappen leeg zijn en welke niet leesbaar waren.

Handig om de indeling van een dossier te zien zonder dat de inhoud in de weg
zit, en om een mappenstructuur ergens anders na te bouwen.

Er wordt niets gewijzigd, alleen gelezen.

.EXAMPLE
.\Mapstructuur_inventariseren.ps1 -Map "D:\Dossiers" -Openen

.EXAMPLE
.\Mapstructuur_inventariseren.ps1 -Map "D:\Projecten" -MaxDiepte 3 -Negeer node_modules,.git -Openen
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "De map waarvan je de structuur wilt zien")]
    [string]$Map,

    [Parameter(HelpMessage = "Pad voor het HTML-rapport. Alleen een bestandsnaam mag ook: dan komt het in Documenten. Leeg laten in combinatie met Openen geeft een automatische naam.")]
    [string]$RapportHtml = "",

    [Parameter(HelpMessage = "Hoeveel niveaus diep kijken. 0 is zover als het gaat.")]
    [int]$MaxDiepte = 0,

    [Parameter(HelpMessage = "Mapnamen om over te slaan, bijvoorbeeld node_modules en .git. Jokertekens mogen.")]
    [string[]]$Negeer = @(),

    [Parameter(HelpMessage = "Ook tellen hoeveel bestanden en hoeveel ruimte er per map in zit")]
    [switch]$MetBestandstelling,

    [Parameter(HelpMessage = "Verborgen en systeemmappen ook meenemen")]
    [switch]$MetVerborgen,

    [Parameter(HelpMessage = "Ook de volledige mappenlijst als tabel in het rapport")]
    [switch]$AlleMappenTonen,

    [Parameter(HelpMessage = "Extra CSV-bestand om in Excel mee door te werken")]
    [string]$RapportCsv = "",

    [Parameter(HelpMessage = "Het rapport na afloop meteen openen")]
    [switch]$Openen
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Map -PathType Container)) {
    Write-Host "Map bestaat niet: $Map" -ForegroundColor Red
    exit 1
}

# --------------------------------------------------------------------------
# Hulpjes
# --------------------------------------------------------------------------

function Leesbaar {
    param([Nullable[long]]$Bytes)
    if ($null -eq $Bytes) { return "" }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N1} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N1} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Meervoud {
    <# 1 map, 2 mappen. Scheelt lelijke haakjes in de teksten. #>
    param([int]$Aantal, [string]$Enkelvoud, [string]$Meervoud)
    if ($Aantal -eq 1) { return "$Aantal $Enkelvoud" }
    return "$Aantal $Meervoud"
}

function Veilige-Naam {
    param([string]$Tekst)
    $ongeldig = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $patroon = "[{0}]" -f [regex]::Escape($ongeldig)
    return ($Tekst -replace $patroon, '_')
}

function Bepaal-Rapportpad {
    <#
    Vul je alleen een bestandsnaam in, dan komt het rapport in Documenten en
    niet in een willekeurige werkmap. Vul je een bestaande map in, dan komt het
    daarin met een automatische naam.
    #>
    param([string]$Opgegeven, [string]$Standaardnaam, [string]$Extensie)

    $pad = "$Opgegeven".Trim().Trim('"')
    if (-not $pad) { $pad = $Standaardnaam }

    if (Test-Path -LiteralPath $pad -PathType Container) {
        $pad = Join-Path $pad $Standaardnaam
    }
    if (-not [System.IO.Path]::GetExtension($pad)) { $pad += $Extensie }

    if (-not [System.IO.Path]::IsPathRooted($pad)) {
        $basis = if (Split-Path -Parent $pad) {
            (Get-Location).Path
        }
        else {
            [Environment]::GetFolderPath('MyDocuments')
        }
        $pad = Join-Path $basis $pad
    }
    return [System.IO.Path]::GetFullPath($pad)
}

function HtmlVeilig {
    param([string]$Tekst)
    if (-not $Tekst) { return "" }
    return $Tekst.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function Openen-Bestand {
    param([string]$Pad)
    try {
        Start-Process -FilePath $Pad -ErrorAction Stop
        Write-Host "Geopend."
    }
    catch {
        Write-Host "Openen lukte niet: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Dubbelklik het bestand hierboven."
    }
}

# --------------------------------------------------------------------------
# Inlezen
# --------------------------------------------------------------------------

$script:knopen = New-Object System.Collections.ArrayList
$script:zonderToegang = New-Object System.Collections.ArrayList

function Overslaan {
    param([string]$Naam)
    foreach ($patroon in $Negeer) {
        if ($Naam -like $patroon) { return $true }
    }
    return $false
}

function Lees-Map {
    <#
    Leest één map en daaronder alles wat mag, en geeft een knoop terug met de
    eigen gegevens plus de totalen van alles eronder.
    #>
    param([string]$Pad, [string]$Naam, [int]$Diepte)

    # Een PSCustomObject en geen hashtable: alleen zo zien Sort-Object,
    # Group-Object en Measure-Object de eigenschappen verderop.
    $knoop = [pscustomobject]@{
        naam            = $Naam
        volledig        = $Pad
        diepte          = $Diepte
        gewijzigd       = $null
        submappen       = @()
        directBestanden = 0
        directBytes     = [long]0
        totaalMappen    = 0
        totaalBestanden = 0
        totaalBytes     = [long]0
        afgekapt        = $false
        fout            = $false
    }

    try {
        $knoop.gewijzigd = (Get-Item -LiteralPath $Pad -Force -ErrorAction Stop).LastWriteTime
    }
    catch { }

    if ($MetBestandstelling) {
        try {
            $bestanden = @(Get-ChildItem -LiteralPath $Pad -File -Force:$MetVerborgen -ErrorAction Stop)
            $knoop.directBestanden = $bestanden.Count
            if ($bestanden.Count -gt 0) {
                $knoop.directBytes = [long](($bestanden | Measure-Object -Property Length -Sum).Sum)
            }
        }
        catch {
            $knoop.fout = $true
        }
    }

    $kinderen = @()
    try {
        $kinderen = @(Get-ChildItem -LiteralPath $Pad -Directory -Force:$MetVerborgen -ErrorAction Stop |
            Sort-Object Name)
    }
    catch {
        $knoop.fout = $true
        $null = $script:zonderToegang.Add($Pad)
    }

    $kinderen = @($kinderen | Where-Object { -not (Overslaan $_.Name) })

    $mag = ($MaxDiepte -le 0) -or ($Diepte -lt $MaxDiepte)
    if ($kinderen.Count -gt 0 -and -not $mag) {
        $knoop.afgekapt = $true
    }

    $knoop.totaalBestanden = $knoop.directBestanden
    $knoop.totaalBytes = $knoop.directBytes

    if ($mag) {
        $lijst = New-Object System.Collections.ArrayList
        foreach ($kind in $kinderen) {
            $kindKnoop = Lees-Map -Pad $kind.FullName -Naam $kind.Name -Diepte ($Diepte + 1)
            $null = $lijst.Add($kindKnoop)
            $knoop.totaalMappen += 1 + $kindKnoop.totaalMappen
            $knoop.totaalBestanden += $kindKnoop.totaalBestanden
            $knoop.totaalBytes += $kindKnoop.totaalBytes
        }
        $knoop.submappen = $lijst.ToArray()
    }

    $null = $script:knopen.Add($knoop)
    if ($script:knopen.Count % 250 -eq 0) {
        Write-Progress -Activity "Mappen lezen" -Status "$($script:knopen.Count) mappen" -CurrentOperation $Pad
    }
    return $knoop
}

$volledigMap = (Resolve-Path -LiteralPath $Map).Path.TrimEnd('\')
$wortelNaam = Split-Path -Leaf $volledigMap
if (-not $wortelNaam) { $wortelNaam = $volledigMap }
$basislengte = $volledigMap.Length + 1

Write-Host "Map      : $volledigMap"
Write-Host "Diepte   : $(if ($MaxDiepte -le 0) { 'alles' } else { "$MaxDiepte niveaus" })"
Write-Host "Overslaan: $(if ($Negeer.Count -gt 0) { $Negeer -join ', ' } else { 'niets' })"
Write-Host "Tellen   : $(if ($MetBestandstelling) { 'mappen en bestanden' } else { 'alleen mappen' })"
Write-Host ""
Write-Host "Bezig met lezen..." -ForegroundColor DarkGray

$wortel = Lees-Map -Pad $volledigMap -Naam $wortelNaam -Diepte 0
Write-Progress -Activity "Mappen lezen" -Completed

# Alle knopen behalve de wortel, op pad gesorteerd.
$alleMappen = @($script:knopen | Where-Object { $_.diepte -gt 0 } | Sort-Object volledig)

if ($alleMappen.Count -eq 0) {
    Write-Host "Geen submappen gevonden onder deze map." -ForegroundColor Yellow
    Write-Host "Er valt dan weinig structuur te tonen." -ForegroundColor Yellow
    exit 0
}

# --------------------------------------------------------------------------
# Rekenen
# --------------------------------------------------------------------------

$diepste = ($alleMappen | Measure-Object -Property diepte -Maximum).Maximum

$perNiveau = $alleMappen | Group-Object diepte | ForEach-Object {
    [pscustomobject]@{
        Niveau  = [int]$_.Name
        Aantal  = $_.Count
        Aandeel = [math]::Round($_.Count / $alleMappen.Count * 100, 1)
    }
} | Sort-Object Niveau

$eindmappen = @($alleMappen | Where-Object { @($_.submappen).Count -eq 0 -and -not $_.afgekapt })

$legeMappen = @()
if ($MetBestandstelling) {
    $legeMappen = @($alleMappen | Where-Object {
            @($_.submappen).Count -eq 0 -and $_.totaalBestanden -eq 0 -and -not $_.afgekapt
        })
}

$vertakkingen = @($alleMappen |
    Where-Object { @($_.submappen).Count -gt 0 } |
    Sort-Object @{Expression = { @($_.submappen).Count }; Descending = $true }, volledig |
    Select-Object -First 20)

$diepsteMappen = @($alleMappen |
    Sort-Object @{Expression = "diepte"; Descending = $true }, volledig |
    Select-Object -First 20)

$grootsteMappen = @()
if ($MetBestandstelling) {
    $grootsteMappen = @($alleMappen |
        Sort-Object @{Expression = "totaalBytes"; Descending = $true } |
        Select-Object -First 20)
}

# --------------------------------------------------------------------------
# In beeld
# --------------------------------------------------------------------------

Write-Host ""
Write-Host "$(Meervoud $alleMappen.Count 'submap' 'submappen'), $(Meervoud ($diepste) 'niveau' 'niveaus') diep" -ForegroundColor Green
if ($MetBestandstelling) {
    Write-Host "$($wortel.totaalBestanden) bestanden, $(Leesbaar $wortel.totaalBytes) in totaal"
    if ($legeMappen.Count -gt 0) { Write-Host "$($legeMappen.Count) mappen zijn helemaal leeg" }
}
if ($script:zonderToegang.Count -gt 0) {
    Write-Host "$($script:zonderToegang.Count) mappen konden niet gelezen worden" -ForegroundColor Yellow
}

function Toon-Boom {
    <# De boom op het scherm, ingesprongen, tot een beperkte diepte. #>
    param($Knoop, [string]$Voorvoegsel, [bool]$Laatste, [bool]$IsWortel, [int]$Tot)

    if ($IsWortel) {
        Write-Host "$($Knoop.naam)\" -ForegroundColor Cyan
        $kindVoorvoegsel = ""
    }
    else {
        $tak = if ($Laatste) { "\-- " } else { "|-- " }
        $staart = ""
        if (@($Knoop.submappen).Count -gt 0) { $staart = "  ($(@($Knoop.submappen).Count))" }
        elseif ($Knoop.afgekapt) { $staart = "  (...)" }
        Write-Host "$Voorvoegsel$tak$($Knoop.naam)\$staart"
        $kindVoorvoegsel = $Voorvoegsel + $(if ($Laatste) { "    " } else { "|   " })
    }

    if ($Knoop.diepte -ge $Tot) {
        if (@($Knoop.submappen).Count -gt 0) {
            Write-Host "$kindVoorvoegsel\-- ... $(Meervoud $Knoop.totaalMappen 'map' 'mappen') hieronder" -ForegroundColor DarkGray
        }
        return
    }

    $kinderen = @($Knoop.submappen)
    for ($i = 0; $i -lt $kinderen.Count; $i++) {
        Toon-Boom -Knoop $kinderen[$i] -Voorvoegsel $kindVoorvoegsel `
            -Laatste ($i -eq $kinderen.Count - 1) -IsWortel $false -Tot $Tot
    }
}

Write-Host ""
Write-Host "MAPPENSTRUCTUUR (op het scherm tot drie niveaus)" -ForegroundColor White
Toon-Boom -Knoop $wortel -Voorvoegsel "" -Laatste $true -IsWortel $true -Tot 3

Write-Host ""
Write-Host "MAPPEN PER NIVEAU" -ForegroundColor White
foreach ($n in $perNiveau) {
    Write-Host ("  niveau {0,-3} {1,6} mappen {2,6}%" -f $n.Niveau, $n.Aantal, $n.Aandeel)
}

# --------------------------------------------------------------------------
# HTML-rapport
# --------------------------------------------------------------------------

$maakRapport = [bool]$RapportHtml -or $Openen

if ($maakRapport) {

    $RapportHtml = Bepaal-Rapportpad -Opgegeven $RapportHtml -Extensie ".html" `
        -Standaardnaam (Veilige-Naam ("Mapstructuur " + $wortelNaam + " " +
            (Get-Date -Format 'yyyyMMdd-HHmm') + ".html"))

    $css = @'
:root { color-scheme: light; }
* { box-sizing: border-box; }
body { font-family: "Segoe UI", system-ui, sans-serif; margin: 0; padding: 32px;
       background: #F5F5F7; color: #1D1D1F; font-size: 14px; }
h1 { font-size: 24px; margin: 0 0 4px; }
h2 { font-size: 17px; margin: 36px 0 10px; padding-bottom: 6px;
     border-bottom: 2px solid #E3E3E6; }
.sub { color: #6E6E73; margin: 0 0 24px; }
.pad-groot { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
             padding: 12px 16px; margin: 0 0 24px; font-family: Consolas, monospace;
             font-size: 13px; word-break: break-all; }
.kaarten { display: flex; gap: 14px; flex-wrap: wrap; margin: 0 0 8px; }
.kaart { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
         padding: 14px 18px; min-width: 150px; border-left: 4px solid #2C3E50; }
.kaart .getal { font-size: 28px; font-weight: 600; line-height: 1.1; }
.kaart .label { color: #6E6E73; font-size: 12px; margin-top: 2px; }
.kaart.b { border-left-color: #3D8FE8; }
.kaart.c { border-left-color: #34A853; }
.kaart.d { border-left-color: #E8A33D; }
table { border-collapse: collapse; width: 100%; background: #fff;
        border: 1px solid #D1D1D6; border-radius: 10px; overflow: hidden; }
th { background: #2C3E50; color: #fff; text-align: left; padding: 9px 12px;
     font-weight: 600; font-size: 13px; position: sticky; top: 0; }
td { padding: 8px 12px; border-top: 1px solid #EDEDF0; vertical-align: middle; }
tr:nth-child(even) td { background: #FAFAFB; }
td.num, th.num { text-align: right; white-space: nowrap;
                 font-variant-numeric: tabular-nums; }
td.pad { font-family: Consolas, monospace; font-size: 12.5px; word-break: break-all; }
.spoor { background: #EDEDF0; border-radius: 3px; height: 8px; width: 120px;
         display: inline-block; vertical-align: middle; overflow: hidden; }
.spoor span { display: block; height: 100%; background: #3D8FE8; }
.mono { font-family: Consolas, monospace; font-size: 12.5px; }
.badge { display: inline-block; padding: 1px 8px; border-radius: 10px;
         font-size: 12px; font-weight: 600; background: #EDEDF0; color: #44443F; }
footer { margin-top: 40px; color: #8E8E93; font-size: 12px; }
.boomvak { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
           padding: 10px 14px; }
details.boom > summary { cursor: pointer; padding: 4px 6px; border-radius: 6px;
                         list-style: none; display: flex; align-items: center; gap: 8px; }
details.boom > summary::-webkit-details-marker { display: none; }
details.boom > summary::before { content: "\25B8"; color: #8E8E93; font-size: 11px;
                                 width: 10px; flex: 0 0 10px; }
details.boom[open] > summary::before { content: "\25BE"; }
details.boom > summary:hover { background: #F0F0F3; }
.blad { padding: 4px 6px; border-radius: 6px; display: flex; align-items: center; gap: 8px; }
.blad::before { content: ""; width: 10px; flex: 0 0 10px; }
.blad:hover { background: #F7F7F9; }
.mapnaam { font-family: Consolas, monospace; font-weight: 600; }
.mapnaam.leeg { font-weight: 400; color: #8E8E93; }
.tel { color: #6E6E73; font-size: 12px; }
.merk { font-size: 11px; padding: 0 6px; border-radius: 8px; background: #FBEEDA;
        color: #8A5B10; }
.merk.rood { background: #FBE0DC; color: #93291C; }
.kinderen { margin-left: 16px; border-left: 1px solid #E9E9EC; padding-left: 10px; }
.grootte { color: #6E6E73; font-size: 12px; margin-left: auto; white-space: nowrap; }
.knoppen { margin: 0 0 10px; display: flex; gap: 8px; flex-wrap: wrap; }
.knoppen button { font: inherit; font-size: 12px; padding: 4px 10px; cursor: pointer;
                  border: 1px solid #D1D1D6; border-radius: 6px; background: #fff;
                  color: #1D1D1F; }
.knoppen button:hover { background: #F0F0F3; }
pre.tekstboom { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
                padding: 14px 16px; font-family: Consolas, monospace; font-size: 12.5px;
                line-height: 1.45; overflow-x: auto; margin: 0; }
@media print {
  .knoppen { display: none; }
  details.boom { break-inside: avoid; }
  body { background: #fff; padding: 0; }
  th { position: static; }
  h2 { page-break-after: avoid; }
  tr { page-break-inside: avoid; }
}
'@

    function Teken-Knoop {
        <# Eén map in de HTML-boom, met zijn submappen eronder. #>
        param($Knoop, [int]$OpenTot)

        $sb = New-Object System.Text.StringBuilder
        $kinderen = @($Knoop.submappen)

        $klasse = if ($kinderen.Count -eq 0 -and $Knoop.totaalBestanden -eq 0 -and $MetBestandstelling) {
            "mapnaam leeg"
        }
        else { "mapnaam" }

        $merken = ""
        if ($Knoop.afgekapt) { $merken += "<span class='merk'>niet verder gekeken</span>" }
        if ($Knoop.fout) { $merken += "<span class='merk rood'>geen toegang</span>" }

        $tel = ""
        if ($kinderen.Count -gt 0) {
            $tel = "<span class='tel'>$(Meervoud $kinderen.Count 'submap' 'submappen')</span>"
        }

        $rechts = ""
        if ($MetBestandstelling) {
            $rechts = "<span class='grootte'>$($Knoop.totaalBestanden) best. &middot; " +
            "$(Leesbaar $Knoop.totaalBytes)</span>"
        }

        $kop = "<span class='$klasse'>$(HtmlVeilig $Knoop.naam)</span>$tel$merken$rechts"

        if ($kinderen.Count -eq 0) {
            $null = $sb.AppendLine("<div class='blad'>$kop</div>")
            return $sb.ToString()
        }

        $open = if ($Knoop.diepte -lt $OpenTot) { " open" } else { "" }
        $null = $sb.AppendLine("<details class='boom'$open><summary>$kop</summary>")
        $null = $sb.AppendLine("<div class='kinderen'>")
        foreach ($kind in $kinderen) {
            $null = $sb.Append((Teken-Knoop -Knoop $kind -OpenTot $OpenTot))
        }
        $null = $sb.AppendLine("</div></details>")
        return $sb.ToString()
    }

    function Teken-Tekstboom {
        <# Dezelfde boom als platte tekst, om te kopieren naar een mail of notitie. #>
        param($Knoop, [string]$Voorvoegsel, [bool]$Laatste, [bool]$IsWortel, $Sb)

        if ($IsWortel) {
            $null = $Sb.AppendLine("$(HtmlVeilig $Knoop.naam)\")
            $kindVoorvoegsel = ""
        }
        else {
            $streep = "$([char]0x2500)$([char]0x2500)"
            $tak = if ($Laatste) { "$([char]0x2514)$streep " } else { "$([char]0x251C)$streep " }
            $null = $Sb.AppendLine("$Voorvoegsel$tak$(HtmlVeilig $Knoop.naam)\")
            $kindVoorvoegsel = $Voorvoegsel + $(if ($Laatste) { "    " } else { "$([char]0x2502)   " })
        }

        $kinderen = @($Knoop.submappen)
        for ($i = 0; $i -lt $kinderen.Count; $i++) {
            Teken-Tekstboom -Knoop $kinderen[$i] -Voorvoegsel $kindVoorvoegsel `
                -Laatste ($i -eq $kinderen.Count - 1) -IsWortel $false -Sb $Sb
        }
    }

    $html = New-Object System.Text.StringBuilder
    $null = $html.AppendLine('<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8">')
    $null = $html.AppendLine('<title>Mapstructuur</title>')
    $null = $html.AppendLine("<style>$css</style></head><body>")

    $null = $html.AppendLine("<h1>Mapstructuur: $(HtmlVeilig $wortelNaam)</h1>")
    $null = $html.AppendLine("<p class='sub'>Gemaakt op $(Get-Date -Format 'dddd d MMMM yyyy, HH:mm')" +
        "$(if ($MaxDiepte -gt 0) { " &middot; tot $MaxDiepte niveaus" })" +
        "$(if ($Negeer.Count -gt 0) { " &middot; overgeslagen: $(HtmlVeilig ($Negeer -join ', '))" })</p>")
    $null = $html.AppendLine("<div class='pad-groot'>$(HtmlVeilig $volledigMap)</div>")

    $null = $html.AppendLine("<div class='kaarten'>")
    $null = $html.AppendLine("<div class='kaart'><div class='getal'>$($alleMappen.Count)</div><div class='label'>submappen</div></div>")
    $null = $html.AppendLine("<div class='kaart b'><div class='getal'>$diepste</div><div class='label'>niveaus diep</div></div>")
    $null = $html.AppendLine("<div class='kaart c'><div class='getal'>$($eindmappen.Count)</div><div class='label'>eindmappen</div></div>")
    if ($MetBestandstelling) {
        $null = $html.AppendLine("<div class='kaart d'><div class='getal'>$($wortel.totaalBestanden)</div><div class='label'>bestanden, $(Leesbaar $wortel.totaalBytes)</div></div>")
    }
    else {
        $null = $html.AppendLine("<div class='kaart d'><div class='getal'>$(@($wortel.submappen).Count)</div><div class='label'>mappen op niveau 1</div></div>")
    }
    $null = $html.AppendLine("</div>")

    $opmerkingen = @()
    if ($legeMappen.Count -gt 0) { $opmerkingen += "$($legeMappen.Count) mappen zijn helemaal leeg" }
    if ($script:zonderToegang.Count -gt 0) { $opmerkingen += "$($script:zonderToegang.Count) mappen waren niet leesbaar" }
    if (-not $MetBestandstelling) { $opmerkingen += "bestanden zijn niet meegeteld" }
    $null = $html.AppendLine("<p class='sub'>$(HtmlVeilig ($opmerkingen -join ' &middot; '))</p>")

    # ---- de boom ----
    $null = $html.AppendLine("<h2>Mappenboom</h2>")
    $null = $html.AppendLine("<div class='knoppen'>" +
        "<button onclick=""alles(true)"">Alles uitklappen</button>" +
        "<button onclick=""alles(false)"">Alles inklappen</button></div>")
    $null = $html.AppendLine("<div class='boomvak'>")
    $null = $html.AppendLine((Teken-Knoop -Knoop $wortel -OpenTot 2))
    $null = $html.AppendLine("</div>")

    # ---- per niveau ----
    $null = $html.AppendLine("<h2>Mappen per niveau</h2>")
    $null = $html.AppendLine("<table><thead><tr><th>Niveau</th><th class='num'>Aantal</th>" +
        "<th class='num'>Aandeel</th><th></th></tr></thead><tbody>")
    foreach ($n in $perNiveau) {
        $null = $html.AppendLine("<tr><td>niveau $($n.Niveau)</td>" +
            "<td class='num'>$($n.Aantal)</td><td class='num'>$($n.Aandeel)%</td>" +
            "<td><span class='spoor'><span style='width:$($n.Aandeel)%'></span></span></td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")

    # ---- breedste vertakkingen ----
    if ($vertakkingen.Count -gt 0) {
        $null = $html.AppendLine("<h2>Mappen met de meeste directe submappen</h2>")
        $null = $html.AppendLine("<table><thead><tr><th class='num'>Submappen</th><th>Map</th>" +
            "<th class='num'>Totaal eronder</th></tr></thead><tbody>")
        foreach ($v in $vertakkingen) {
            $null = $html.AppendLine("<tr><td class='num'>$(@($v.submappen).Count)</td>" +
                "<td class='pad'>$(HtmlVeilig $v.volledig.Substring($basislengte))</td>" +
                "<td class='num'>$($v.totaalMappen)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    # ---- diepste paden ----
    $null = $html.AppendLine("<h2>Diepste paden</h2>")
    $null = $html.AppendLine("<table><thead><tr><th class='num'>Niveau</th><th>Map</th>" +
        "<th class='num'>Padlengte</th></tr></thead><tbody>")
    foreach ($d in $diepsteMappen) {
        $null = $html.AppendLine("<tr><td class='num'>$($d.diepte)</td>" +
            "<td class='pad'>$(HtmlVeilig $d.volledig.Substring($basislengte))</td>" +
            "<td class='num'>$($d.volledig.Length)</td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")
    $langePaden = @($alleMappen | Where-Object { $_.volledig.Length -gt 240 })
    if ($langePaden.Count -gt 0) {
        $null = $html.AppendLine("<p class='sub'>$($langePaden.Count) paden zijn langer dan 240 tekens. " +
            "Dat geeft in oudere programma's problemen bij kopieren.</p>")
    }

    # ---- grootste mappen ----
    if ($MetBestandstelling -and $grootsteMappen.Count -gt 0) {
        $null = $html.AppendLine("<h2>Twintig grootste mappen, alles eronder meegeteld</h2>")
        $null = $html.AppendLine("<table><thead><tr><th class='num'>Grootte</th><th>Map</th>" +
            "<th class='num'>Bestanden</th></tr></thead><tbody>")
        foreach ($g in $grootsteMappen) {
            $null = $html.AppendLine("<tr><td class='num'>$(Leesbaar $g.totaalBytes)</td>" +
                "<td class='pad'>$(HtmlVeilig $g.volledig.Substring($basislengte))</td>" +
                "<td class='num'>$($g.totaalBestanden)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    # ---- lege mappen ----
    if ($legeMappen.Count -gt 0) {
        $null = $html.AppendLine("<h2>Lege mappen <span class='badge'>$($legeMappen.Count)</span></h2>")
        $null = $html.AppendLine("<table><thead><tr><th>Map</th>" +
            "<th class='num'>Laatst gewijzigd</th></tr></thead><tbody>")
        foreach ($l in ($legeMappen | Select-Object -First 200)) {
            $gew = if ($l.gewijzigd) { $l.gewijzigd.ToString('yyyy-MM-dd HH:mm') } else { "" }
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $l.volledig.Substring($basislengte))</td>" +
                "<td class='num'>$gew</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
        if ($legeMappen.Count -gt 200) {
            $null = $html.AppendLine("<p class='sub'>Alleen de eerste 200 staan hierboven. De rest staat in de CSV.</p>")
        }
    }

    # ---- niet leesbaar ----
    if ($script:zonderToegang.Count -gt 0) {
        $null = $html.AppendLine("<h2>Niet leesbaar <span class='badge'>$($script:zonderToegang.Count)</span></h2>")
        $null = $html.AppendLine("<table><thead><tr><th>Map</th></tr></thead><tbody>")
        foreach ($z in ($script:zonderToegang | Select-Object -First 100)) {
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $z)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    # ---- boom als platte tekst ----
    $tekst = New-Object System.Text.StringBuilder
    Teken-Tekstboom -Knoop $wortel -Voorvoegsel "" -Laatste $true -IsWortel $true -Sb $tekst
    $null = $html.AppendLine("<h2>Boom als platte tekst</h2>")
    $null = $html.AppendLine("<div class='knoppen'><button onclick=""kopieer()"">Kopieren</button></div>")
    $null = $html.AppendLine("<pre class='tekstboom' id='tekstboom'>$($tekst.ToString())</pre>")

    # ---- volledige lijst, alleen op verzoek ----
    if ($AlleMappenTonen) {
        $null = $html.AppendLine("<h2>Alle mappen <span class='badge'>$($alleMappen.Count)</span></h2>")
        $null = $html.AppendLine("<table><thead><tr><th>Pad</th><th>Naam</th>" +
            "<th class='num'>Niveau</th><th class='num'>Submappen</th>" +
            "<th class='num'>Gewijzigd</th></tr></thead><tbody>")
        foreach ($m in $alleMappen) {
            $gew = if ($m.gewijzigd) { $m.gewijzigd.ToString('yyyy-MM-dd HH:mm') } else { "" }
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $m.volledig.Substring($basislengte))</td>" +
                "<td class='pad'>$(HtmlVeilig $m.naam)</td>" +
                "<td class='num'>$($m.diepte)</td>" +
                "<td class='num'>$(@($m.submappen).Count)</td>" +
                "<td class='num'>$gew</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }
    else {
        $null = $html.AppendLine("<footer>De volledige mappenlijst staat niet in dit rapport. " +
            "Vink 'AlleMappenTonen' aan als je die erbij wilt.</footer>")
    }

    $null = $html.AppendLine("<footer>$(HtmlVeilig $volledigMap)<br>" +
        "$($alleMappen.Count) submappen, $diepste niveaus diep</footer>")
    $null = $html.AppendLine(@'
<script>
function alles(open) {
  document.querySelectorAll("details.boom").forEach(function (d) { d.open = open; });
}
function kopieer() {
  var t = document.getElementById("tekstboom").innerText;
  navigator.clipboard.writeText(t).then(function () { alert("Gekopieerd."); },
    function () { alert("Kopieren lukte niet. Selecteer de tekst met de muis."); });
}
</script>
'@)
    $null = $html.AppendLine("</body></html>")

    $doelmap = Split-Path -Parent $RapportHtml
    if ($doelmap -and -not (Test-Path -LiteralPath $doelmap)) {
        New-Item -ItemType Directory -Path $doelmap -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($RapportHtml, $html.ToString(), [System.Text.UTF8Encoding]::new($true))

    Write-Host ""
    Write-Host "Rapport opgeslagen: $RapportHtml" -ForegroundColor Green
    if ($Openen) { Openen-Bestand $RapportHtml }
    else { Write-Host "Dubbelklikken opent het in je browser." }
}

# --------------------------------------------------------------------------
# CSV, voor wie ermee wil doorwerken
# --------------------------------------------------------------------------

if ($RapportCsv) {
    $RapportCsv = Bepaal-Rapportpad -Opgegeven $RapportCsv -Extensie ".csv" `
        -Standaardnaam (Veilige-Naam ("Mapstructuur " + $wortelNaam + " " +
            (Get-Date -Format 'yyyyMMdd-HHmm') + ".csv"))

    $rijen = $alleMappen | ForEach-Object {
        [pscustomobject]@{
            Niveau           = $_.diepte
            Mapnaam          = $_.naam
            RelatiefPad      = $_.volledig.Substring($basislengte)
            BovenliggendeMap = Split-Path -Parent $_.volledig.Substring($basislengte)
            DirecteSubmappen = @($_.submappen).Count
            MappenEronder    = $_.totaalMappen
            Bestanden        = if ($MetBestandstelling) { $_.totaalBestanden } else { "" }
            Bytes            = if ($MetBestandstelling) { $_.totaalBytes } else { "" }
            Grootte          = if ($MetBestandstelling) { Leesbaar $_.totaalBytes } else { "" }
            Gewijzigd        = if ($_.gewijzigd) { $_.gewijzigd.ToString("yyyy-MM-dd HH:mm") } else { "" }
            Padlengte        = $_.volledig.Length
            VolledigPad      = $_.volledig
        }
    }

    $csvmap = Split-Path -Parent $RapportCsv
    if ($csvmap -and -not (Test-Path -LiteralPath $csvmap)) {
        New-Item -ItemType Directory -Path $csvmap -Force | Out-Null
    }
    # Puntkomma en UTF8-BOM: dan opent Excel het meteen goed.
    $rijen | Export-Csv -LiteralPath $RapportCsv -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Write-Host "CSV opgeslagen: $RapportCsv" -ForegroundColor Green
}

Write-Host ""
Write-Host "Klaar." -ForegroundColor Green
