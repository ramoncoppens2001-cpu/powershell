<#
.SYNOPSIS
Vergelijkt twee mappenbomen en maakt een leesbaar HTML-rapport van de verschillen.

.DESCRIPTION
Bedoeld om een archiefmap te controleren tegen de werkmap: welke bestanden zijn
nooit gearchiveerd, wat staat er in het archief dat uit de werkmap verdwenen is,
en welke bestanden staan in beide maar zijn niet gelijk.

Het rapport begint met een overzicht per dossiermap, zodat je in één oogopslag
ziet welk dossier aandacht nodig heeft. Daaronder staan de bestanden per
categorie. De namen van je eigen mappen worden als label gebruikt, dus je hoeft
niet te onthouden welke kant A was.

Bestanden worden vergeleken op relatief pad en op grootte. De inhoud wordt niet
gelezen, dus twee bestanden van gelijke grootte met andere inhoud gelden als
gelijk.

Er wordt niets gewijzigd.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "Eerste map, bijvoorbeeld het archief")]
    [string]$MapA,

    [Parameter(Mandatory, HelpMessage = "Tweede map, bijvoorbeeld het dossier")]
    [string]$MapB,

    [Parameter(HelpMessage = "Pad voor het HTML-rapport. Alleen een bestandsnaam mag ook: dan komt het in Documenten. Leeg laten in combinatie met Openen geeft een automatische naam.")]
    [string]$RapportHtml = "",

    [Parameter(HelpMessage = "Naam voor de eerste map in het rapport. Leeg = de mapnaam zelf.")]
    [string]$LabelA = "",

    [Parameter(HelpMessage = "Naam voor de tweede map in het rapport. Leeg = de mapnaam zelf.")]
    [string]$LabelB = "",

    [Parameter(HelpMessage = "Ook de bestanden opnemen die in beide mappen gelijk zijn")]
    [switch]$GelijkeOokTonen,

    [Parameter(HelpMessage = "Het rapport na afloop meteen openen (standaard aan)")]
    [switch]$Openen
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = "Stop"

foreach ($paar in @(@{n = "MapA"; p = $MapA }, @{n = "MapB"; p = $MapB })) {
    if (-not (Test-Path -LiteralPath $paar.p -PathType Container)) {
        Write-Host "$($paar.n) bestaat niet: $($paar.p)" -ForegroundColor Red
        exit 1
    }
}

if (-not $LabelA) { $LabelA = Split-Path -Leaf $MapA.TrimEnd('\') }
if (-not $LabelB) { $LabelB = Split-Path -Leaf $MapB.TrimEnd('\') }
if ($LabelA -eq $LabelB) { $LabelA = "$LabelA (1)"; $LabelB = "$LabelB (2)" }

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

function VerschilTekst {
    param([long]$Van, [long]$Naar)
    $delta = $Naar - $Van
    $teken = if ($delta -gt 0) { "+" } else { "-" }
    $tekst = "$teken$(Leesbaar ([math]::Abs($delta)))"
    if ($Van -gt 0) {
        $procent = [math]::Round(($delta / $Van) * 100, 1)
        $tekst += "  ($teken$([math]::Abs($procent))%)"
    }
    return $tekst
}

function Dossier {
    # Eerste mapniveau onder de basis; daarop groeperen we het overzicht.
    param([string]$RelatiefPad)
    $delen = $RelatiefPad -split '\\'
    if ($delen.Count -le 1) { return "(hoofdmap)" }
    return $delen[0]
}

function HtmlVeilig {
    param([string]$Tekst)
    if (-not $Tekst) { return "" }
    return [System.Web.HttpUtility]::HtmlEncode($Tekst)
}

# System.Web is niet altijd geladen; zonder deze regel valt HtmlVeilig om.
try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch {
    function HtmlVeilig {
        param([string]$Tekst)
        if (-not $Tekst) { return "" }
        return $Tekst.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
    }
}

function Veilige-Naam {
    param([string]$Tekst)
    $ongeldig = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    $patroon = "[{0}]" -f [regex]::Escape($ongeldig)
    return ($Tekst -replace $patroon, '_')
}

function Bepaal-Rapportpad {
    <#
    Maakt van wat de gebruiker invulde een volledig pad.

    Vul je alleen een bestandsnaam in, dan komt het rapport in Documenten en
    niet in een willekeurige werkmap waar je het nooit meer terugvindt. Vul je
    een bestaande map in, dan komt het daarin met een automatische naam. Vul
    je niets in, dan wordt de hele naam automatisch bepaald.
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

function Lees-Boom {
    param([string]$Basis)
    $lengte = $Basis.TrimEnd('\').Length + 1
    $tabel = @{}
    Get-ChildItem -LiteralPath $Basis -Recurse -File -ErrorAction SilentlyContinue |
        ForEach-Object { $tabel[$_.FullName.Substring($lengte)] = $_ }
    return $tabel
}

# --------------------------------------------------------------------------
# Mappenboom opbouwen en tekenen
# --------------------------------------------------------------------------

function Nieuwe-Knoop {
    return [ordered]@{ mappen = [ordered]@{}; bestanden = @() }
}

function Voeg-InBoom {
    <# Zet een bestand op zijn plek in de boom, submappen aanmakend waar nodig. #>
    param($Knoop, [string[]]$Delen, $Bestand)

    if ($Delen.Count -le 1) {
        $Knoop.bestanden = @($Knoop.bestanden) + @($Bestand)
        return
    }
    $naam = $Delen[0]
    if (-not $Knoop.mappen.Contains($naam)) {
        $Knoop.mappen[$naam] = Nieuwe-Knoop
    }
    Voeg-InBoom -Knoop $Knoop.mappen[$naam] -Delen $Delen[1..($Delen.Count - 1)] -Bestand $Bestand
}

function Teken-Knoop {
    <#
    Tekent één map met zijn inhoud, en geeft de HTML plus de aantallen terug.

    De aantallen zijn nodig in de kop van de map, maar die kop staat vóór de
    kinderen in de HTML. Daarom eerst de kinderen tekenen, dan de kop.
    #>
    param($Knoop, [string]$Naam, [int]$Diepte, [int]$OpenTot = 1)

    $binnen = New-Object System.Text.StringBuilder
    $telling = @{ ok = 0; mist = 0; mist2 = 0; afw = 0 }

    foreach ($submap in $Knoop.mappen.Keys) {
        $kind = Teken-Knoop -Knoop $Knoop.mappen[$submap] -Naam $submap `
            -Diepte ($Diepte + 1) -OpenTot $OpenTot
        $null = $binnen.Append($kind.html)
        foreach ($sleutel in @('ok', 'mist', 'mist2', 'afw')) {
            $telling[$sleutel] += $kind.telling[$sleutel]
        }
    }

    if (@($Knoop.bestanden).Count -gt 0) {
        $null = $binnen.AppendLine("<ul class='bestanden'>")
        foreach ($bestand in ($Knoop.bestanden | Sort-Object naam)) {
            $telling[$bestand.klasse]++
            $null = $binnen.AppendLine("<li><span class='stip s-$($bestand.klasse)'></span>" +
                "<span>$(HtmlVeilig $bestand.naam)</span>" +
                "<span class='grootte'>$(HtmlVeilig $bestand.uitleg)</span></li>")
        }
        $null = $binnen.AppendLine("</ul>")
    }

    $problemen = $telling.mist + $telling.mist2 + $telling.afw
    $totaal = $problemen + $telling.ok

    $merken = New-Object System.Text.StringBuilder
    $null = $merken.Append("<span class='tel'>$totaal bestand$(if ($totaal -ne 1) { 'en' })</span>")
    if ($telling.mist -gt 0) {
        $null = $merken.Append("<span class='badge b-mist'>$($telling.mist) ontbreekt in $(HtmlVeilig $LabelB)</span>")
    }
    if ($telling.mist2 -gt 0) {
        $null = $merken.Append("<span class='badge b-mist2'>$($telling.mist2) ontbreekt in $(HtmlVeilig $LabelA)</span>")
    }
    if ($telling.afw -gt 0) {
        $null = $merken.Append("<span class='badge b-afw'>$($telling.afw) afwijkend</span>")
    }
    if ($problemen -eq 0 -and $totaal -gt 0) {
        $null = $merken.Append("<span class='badge b-ok'>in orde</span>")
    }

    # Mappen met problemen staan open, de rest dicht: zo zie je bij openen
    # meteen waar het zit zonder alles te hoeven doorklikken.
    $open = if ($Diepte -lt $OpenTot -or $problemen -gt 0) { " open" } else { "" }

    $html = New-Object System.Text.StringBuilder
    $null = $html.AppendLine("<details class='boom'$open><summary>" +
        "<span class='mapnaam'>$(HtmlVeilig $Naam)</span>$($merken.ToString())</summary>")
    $null = $html.AppendLine("<div class='kinderen'>$($binnen.ToString())</div></details>")

    return @{ html = $html.ToString(); telling = $telling }
}

# --------------------------------------------------------------------------
# Vergelijken
# --------------------------------------------------------------------------

Write-Host "$LabelA : $MapA"
Write-Host "$LabelB : $MapB"
Write-Host "Inlezen..."

$a = Lees-Boom -Basis $MapA
$b = Lees-Boom -Basis $MapB

Write-Host "$($a.Count) bestanden in $LabelA, $($b.Count) in $LabelB."
Write-Host ""

$ontbreektInB = @($a.Keys | Where-Object { -not $b.ContainsKey($_) } | Sort-Object)
$ontbreektInA = @($b.Keys | Where-Object { -not $a.ContainsKey($_) } | Sort-Object)
$beide = @($a.Keys | Where-Object { $b.ContainsKey($_) } | Sort-Object)
$afwijkend = @($beide | Where-Object { $a[$_].Length -ne $b[$_].Length })
$gelijk = @($beide | Where-Object { $a[$_].Length -eq $b[$_].Length })

# --------------------------------------------------------------------------
# Overzicht per dossiermap
# --------------------------------------------------------------------------

$mappen = @{}
function Tel {
    param([string]$Pad, [string]$Soort)
    $naam = Dossier $Pad
    if (-not $mappen.ContainsKey($naam)) {
        $mappen[$naam] = [ordered]@{ Map = $naam; MistB = 0; MistA = 0; Afwijkend = 0; Gelijk = 0 }
    }
    $mappen[$naam][$Soort]++
}
foreach ($p in $ontbreektInB) { Tel $p "MistB" }
foreach ($p in $ontbreektInA) { Tel $p "MistA" }
foreach ($p in $afwijkend) { Tel $p "Afwijkend" }
foreach ($p in $gelijk) { Tel $p "Gelijk" }

$overzicht = $mappen.Values | ForEach-Object {
    $problemen = $_.MistB + $_.MistA + $_.Afwijkend
    [pscustomobject]@{
        Map       = $_.Map
        MistB     = $_.MistB
        MistA     = $_.MistA
        Afwijkend = $_.Afwijkend
        Gelijk    = $_.Gelijk
        Problemen = $problemen
    }
} | Sort-Object -Property @{Expression = "Problemen"; Descending = $true }, "Map"

# --------------------------------------------------------------------------
# In beeld
# --------------------------------------------------------------------------

Write-Host "OVERZICHT PER MAP" -ForegroundColor White
Write-Host ("  {0,-24} {1,8} {2,8} {3,10} {4,8}" -f "map", "mist $LabelB", "mist $LabelA", "afwijkend", "gelijk")
foreach ($rij in $overzicht) {
    $kleur = if ($rij.Problemen -gt 0) { "Yellow" } else { "Gray" }
    Write-Host ("  {0,-24} {1,8} {2,8} {3,10} {4,8}" -f `
            $rij.Map, $rij.MistB, $rij.MistA, $rij.Afwijkend, $rij.Gelijk) -ForegroundColor $kleur
}

Write-Host ""
Write-Host "ONTBREEKT IN $LabelB  ($($ontbreektInB.Count))" -ForegroundColor Yellow
if ($ontbreektInB.Count -eq 0) { Write-Host "  niets" } else {
    foreach ($p in $ontbreektInB) { Write-Host ("  {0,10}  {1}" -f (Leesbaar $a[$p].Length), $p) }
}

Write-Host ""
Write-Host "ONTBREEKT IN $LabelA  ($($ontbreektInA.Count))" -ForegroundColor Cyan
if ($ontbreektInA.Count -eq 0) { Write-Host "  niets" } else {
    foreach ($p in $ontbreektInA) { Write-Host ("  {0,10}  {1}" -f (Leesbaar $b[$p].Length), $p) }
}

Write-Host ""
Write-Host "IN BEIDE, ANDERE GROOTTE  ($($afwijkend.Count))" -ForegroundColor Magenta
if ($afwijkend.Count -eq 0) { Write-Host "  niets" } else {
    foreach ($p in $afwijkend) {
        Write-Host ("  {0}" -f $p)
        Write-Host ("      {0}: {1,10}    {2}: {3,10}    {4}" -f `
                $LabelA, (Leesbaar $a[$p].Length), $LabelB, (Leesbaar $b[$p].Length),
            (VerschilTekst $a[$p].Length $b[$p].Length))
    }
}

Write-Host ""
Write-Host "Gelijk in beide: $($gelijk.Count) bestanden." -ForegroundColor Green

# --------------------------------------------------------------------------
# HTML-rapport
# --------------------------------------------------------------------------

# Vraag je om openen zonder pad, dan is een rapport toch nodig: anders valt er
# niets te openen en gebeurt er ogenschijnlijk niets.
$maakRapport = [bool]$RapportHtml -or $Openen

if ($maakRapport) {

    $RapportHtml = Bepaal-Rapportpad -Opgegeven $RapportHtml -Extensie ".html" `
        -Standaardnaam (Veilige-Naam ("Vergelijking $LabelA - $LabelB " +
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
.mappen { display: flex; gap: 14px; flex-wrap: wrap; margin: 0 0 24px; }
.map { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
       padding: 12px 16px; min-width: 260px; }
.map .naam { font-weight: 600; }
.map .pad { color: #6E6E73; font-size: 12px; word-break: break-all;
            font-family: Consolas, monospace; }
.kaarten { display: flex; gap: 14px; flex-wrap: wrap; margin: 0 0 8px; }
.kaart { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
         padding: 14px 18px; min-width: 150px; }
.kaart .getal { font-size: 28px; font-weight: 600; line-height: 1.1; }
.kaart .label { color: #6E6E73; font-size: 12px; margin-top: 2px; }
.kaart.mist   { border-left: 4px solid #E8A33D; }
.kaart.mist2  { border-left: 4px solid #3D8FE8; }
.kaart.afw    { border-left: 4px solid #B85CC7; }
.kaart.ok     { border-left: 4px solid #34A853; }
table { border-collapse: collapse; width: 100%; background: #fff;
        border: 1px solid #D1D1D6; border-radius: 10px; overflow: hidden; }
th { background: #2C3E50; color: #fff; text-align: left; padding: 9px 12px;
     font-weight: 600; font-size: 13px; position: sticky; top: 0; }
td { padding: 8px 12px; border-top: 1px solid #EDEDF0; vertical-align: top; }
tr:nth-child(even) td { background: #FAFAFB; }
td.num, th.num { text-align: right; white-space: nowrap;
                 font-variant-numeric: tabular-nums; }
td.pad { font-family: Consolas, monospace; font-size: 12.5px; word-break: break-all; }
tr.let td { background: #FFF9EC; }
.badge { display: inline-block; padding: 1px 8px; border-radius: 10px;
         font-size: 12px; font-weight: 600; white-space: nowrap; }
.b-mist  { background: #FDF0DC; color: #8A5A12; }
.b-mist2 { background: #E3EFFB; color: #1B4E80; }
.b-afw   { background: #F6E6F8; color: #6E2A7A; }
.b-ok    { background: #E4F4E7; color: #1E6B2E; }
.groter { color: #B02A2A; font-weight: 600; }
.kleiner { color: #1E6B2E; font-weight: 600; }
.leeg { color: #8E8E93; font-style: italic; padding: 14px 12px;
        background: #fff; border: 1px solid #D1D1D6; border-radius: 10px; }
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
.mapnaam { font-family: Consolas, monospace; font-weight: 600; }
.tel { color: #6E6E73; font-size: 12px; }
.kinderen { margin-left: 16px; border-left: 1px solid #E9E9EC; padding-left: 10px; }
ul.bestanden { list-style: none; margin: 2px 0 6px; padding: 0 0 0 18px; }
ul.bestanden li { display: flex; align-items: center; gap: 8px; padding: 2px 0;
                  font-family: Consolas, monospace; font-size: 12.5px; }
.stip { width: 8px; height: 8px; border-radius: 50%; flex: 0 0 8px; }
.s-ok { background: #CFD2D7; }
.s-mist { background: #E8A33D; }
.s-mist2 { background: #3D8FE8; }
.s-afw { background: #B85CC7; }
.grootte { color: #6E6E73; font-size: 12px; margin-left: auto; white-space: nowrap; }
.legenda { display: flex; gap: 16px; flex-wrap: wrap; margin: 0 0 10px;
           color: #6E6E73; font-size: 12px; align-items: center; }
.legenda .item { display: flex; align-items: center; gap: 5px; }
.knoppen { margin: 0 0 10px; display: flex; gap: 8px; }
.knoppen button { font: inherit; font-size: 12px; padding: 4px 10px; cursor: pointer;
                  border: 1px solid #D1D1D6; border-radius: 6px; background: #fff;
                  color: #1D1D1F; }
.knoppen button:hover { background: #F0F0F3; }
@media print {
  .knoppen { display: none; }
  details.boom { break-inside: avoid; }
  body { background: #fff; padding: 0; }
  th { position: static; }
  h2 { page-break-after: avoid; }
  tr { page-break-inside: avoid; }
}
'@

    $html = New-Object System.Text.StringBuilder
    $null = $html.AppendLine('<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8">')
    $null = $html.AppendLine('<title>Mappen vergelijken</title>')
    $null = $html.AppendLine("<style>$css</style></head><body>")

    $null = $html.AppendLine("<h1>Mappenvergelijking</h1>")
    $null = $html.AppendLine("<p class='sub'>Gemaakt op $(Get-Date -Format 'dddd d MMMM yyyy, HH:mm')</p>")

    $null = $html.AppendLine("<div class='mappen'>")
    foreach ($m in @(@{l = $LabelA; p = $MapA; n = $a.Count }, @{l = $LabelB; p = $MapB; n = $b.Count })) {
        $null = $html.AppendLine("<div class='map'><div class='naam'>$(HtmlVeilig $m.l) &mdash; $($m.n) bestanden</div>")
        $null = $html.AppendLine("<div class='pad'>$(HtmlVeilig $m.p)</div></div>")
    }
    $null = $html.AppendLine("</div>")

    $null = $html.AppendLine("<div class='kaarten'>")
    $null = $html.AppendLine("<div class='kaart mist'><div class='getal'>$($ontbreektInB.Count)</div><div class='label'>ontbreekt in $(HtmlVeilig $LabelB)</div></div>")
    $null = $html.AppendLine("<div class='kaart mist2'><div class='getal'>$($ontbreektInA.Count)</div><div class='label'>ontbreekt in $(HtmlVeilig $LabelA)</div></div>")
    $null = $html.AppendLine("<div class='kaart afw'><div class='getal'>$($afwijkend.Count)</div><div class='label'>in beide, andere grootte</div></div>")
    $null = $html.AppendLine("<div class='kaart ok'><div class='getal'>$($gelijk.Count)</div><div class='label'>gelijk</div></div>")
    $null = $html.AppendLine("</div>")

    # ---- overzicht per map ----
    $null = $html.AppendLine("<h2>Overzicht per map</h2>")
    $null = $html.AppendLine("<table><thead><tr><th>Map</th>")
    $null = $html.AppendLine("<th class='num'>Ontbreekt in $(HtmlVeilig $LabelB)</th>")
    $null = $html.AppendLine("<th class='num'>Ontbreekt in $(HtmlVeilig $LabelA)</th>")
    $null = $html.AppendLine("<th class='num'>Andere grootte</th><th class='num'>Gelijk</th>")
    $null = $html.AppendLine("<th>Status</th></tr></thead><tbody>")
    foreach ($rij in $overzicht) {
        $klasse = if ($rij.Problemen -gt 0) { " class='let'" } else { "" }
        $status = if ($rij.Problemen -gt 0) {
            "<span class='badge b-mist'>$($rij.Problemen) aandachtspunt$(if ($rij.Problemen -ne 1) { 'en' })</span>"
        }
        else { "<span class='badge b-ok'>in orde</span>" }
        $null = $html.AppendLine("<tr$klasse><td>$(HtmlVeilig $rij.Map)</td>" +
            "<td class='num'>$($rij.MistB)</td><td class='num'>$($rij.MistA)</td>" +
            "<td class='num'>$($rij.Afwijkend)</td><td class='num'>$($rij.Gelijk)</td>" +
            "<td>$status</td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")

    # ---- mappenboom ----
    $wortel = Nieuwe-Knoop
    foreach ($p in $ontbreektInB) {
        Voeg-InBoom -Knoop $wortel -Delen ($p -split '\\') -Bestand @{
            naam = Split-Path -Leaf $p; klasse = "mist"
            uitleg = "$(Leesbaar $a[$p].Length) - alleen in $LabelA"
        }
    }
    foreach ($p in $ontbreektInA) {
        Voeg-InBoom -Knoop $wortel -Delen ($p -split '\\') -Bestand @{
            naam = Split-Path -Leaf $p; klasse = "mist2"
            uitleg = "$(Leesbaar $b[$p].Length) - alleen in $LabelB"
        }
    }
    foreach ($p in $afwijkend) {
        Voeg-InBoom -Knoop $wortel -Delen ($p -split '\\') -Bestand @{
            naam = Split-Path -Leaf $p; klasse = "afw"
            uitleg = "$(Leesbaar $a[$p].Length) tegen $(Leesbaar $b[$p].Length)  ($(VerschilTekst $a[$p].Length $b[$p].Length))"
        }
    }
    foreach ($p in $gelijk) {
        Voeg-InBoom -Knoop $wortel -Delen ($p -split '\\') -Bestand @{
            naam = Split-Path -Leaf $p; klasse = "ok"
            uitleg = "$(Leesbaar $a[$p].Length)"
        }
    }

    $null = $html.AppendLine("<h2>Structuur</h2>")
    $null = $html.AppendLine("<div class='legenda'>" +
        "<span class='item'><span class='stip s-mist'></span>alleen in $(HtmlVeilig $LabelA)</span>" +
        "<span class='item'><span class='stip s-mist2'></span>alleen in $(HtmlVeilig $LabelB)</span>" +
        "<span class='item'><span class='stip s-afw'></span>andere grootte</span>" +
        "<span class='item'><span class='stip s-ok'></span>gelijk</span></div>")
    $null = $html.AppendLine("<div class='knoppen'>" +
        "<button onclick=""alles(true)"">Alles uitklappen</button>" +
        "<button onclick=""alles(false)"">Alles inklappen</button></div>")
    $null = $html.AppendLine("<div class='boomvak'>")

    $wortelnaam = "$(Split-Path -Leaf $MapA.TrimEnd('\'))  /  $(Split-Path -Leaf $MapB.TrimEnd('\'))"
    $getekend = Teken-Knoop -Knoop $wortel -Naam $wortelnaam -Diepte 0 -OpenTot 2
    $null = $html.AppendLine($getekend.html)
    $null = $html.AppendLine("</div>")

    # ---- ontbrekende bestanden ----
    foreach ($blok in @(
            @{ titel = "Ontbreekt in $LabelB"; lijst = $ontbreektInB; bron = $a; badge = "b-mist" },
            @{ titel = "Ontbreekt in $LabelA"; lijst = $ontbreektInA; bron = $b; badge = "b-mist2" }
        )) {
        $null = $html.AppendLine("<h2>$(HtmlVeilig $blok.titel) <span class='badge $($blok.badge)'>$($blok.lijst.Count)</span></h2>")
        if ($blok.lijst.Count -eq 0) {
            $null = $html.AppendLine("<div class='leeg'>Niets. Alles wat hier hoort te staan is aanwezig.</div>")
            continue
        }
        $null = $html.AppendLine("<table><thead><tr><th>Map</th><th>Bestand</th><th class='num'>Grootte</th><th class='num'>Gewijzigd</th></tr></thead><tbody>")
        foreach ($p in $blok.lijst) {
            $info = $blok.bron[$p]
            $submap = Split-Path -Parent $p
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $(if ($submap) { $submap } else { '.' }))</td>" +
                "<td class='pad'>$(HtmlVeilig $info.Name)</td>" +
                "<td class='num'>$(Leesbaar $info.Length)</td>" +
                "<td class='num'>$($info.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    # ---- afwijkende grootte ----
    $null = $html.AppendLine("<h2>In beide, maar andere grootte <span class='badge b-afw'>$($afwijkend.Count)</span></h2>")
    if ($afwijkend.Count -eq 0) {
        $null = $html.AppendLine("<div class='leeg'>Niets. Alle bestanden die in beide mappen staan zijn even groot.</div>")
    }
    else {
        $null = $html.AppendLine("<table><thead><tr><th>Map</th><th>Bestand</th>" +
            "<th class='num'>$(HtmlVeilig $LabelA)</th><th class='num'>$(HtmlVeilig $LabelB)</th>" +
            "<th class='num'>Verschil</th><th class='num'>Gewijzigd $(HtmlVeilig $LabelB)</th></tr></thead><tbody>")
        foreach ($p in ($afwijkend | Sort-Object { [math]::Abs($b[$_].Length - $a[$_].Length) } -Descending)) {
            $submap = Split-Path -Parent $p
            $groei = if ($b[$p].Length -gt $a[$p].Length) { "groter" } else { "kleiner" }
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $(if ($submap) { $submap } else { '.' }))</td>" +
                "<td class='pad'>$(HtmlVeilig $a[$p].Name)</td>" +
                "<td class='num'>$(Leesbaar $a[$p].Length)</td>" +
                "<td class='num'>$(Leesbaar $b[$p].Length)</td>" +
                "<td class='num $groei'>$(VerschilTekst $a[$p].Length $b[$p].Length)</td>" +
                "<td class='num'>$($b[$p].LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    # ---- gelijke bestanden, alleen op verzoek ----
    if ($GelijkeOokTonen) {
        $null = $html.AppendLine("<h2>Gelijk in beide <span class='badge b-ok'>$($gelijk.Count)</span></h2>")
        if ($gelijk.Count -eq 0) {
            $null = $html.AppendLine("<div class='leeg'>Niets.</div>")
        }
        else {
            $null = $html.AppendLine("<table><thead><tr><th>Map</th><th>Bestand</th><th class='num'>Grootte</th></tr></thead><tbody>")
            foreach ($p in $gelijk) {
                $submap = Split-Path -Parent $p
                $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $(if ($submap) { $submap } else { '.' }))</td>" +
                    "<td class='pad'>$(HtmlVeilig $a[$p].Name)</td>" +
                    "<td class='num'>$(Leesbaar $a[$p].Length)</td></tr>")
            }
            $null = $html.AppendLine("</tbody></table>")
        }
    }

    $null = $html.AppendLine("<footer>Vergelijking op relatief pad en bestandsgrootte. " +
        "De inhoud van bestanden is niet gelezen, dus twee bestanden van gelijke " +
        "grootte met andere inhoud staan hier als gelijk.<br>" +
        "$(HtmlVeilig $LabelA): $(HtmlVeilig $MapA)<br>$(HtmlVeilig $LabelB): $(HtmlVeilig $MapB)</footer>")
    $null = $html.AppendLine(@'
<script>
function alles(open) {
  document.querySelectorAll("details.boom").forEach(function (d) { d.open = open; });
}
</script>
'@)
    $null = $html.AppendLine("</body></html>")

    $map = Split-Path -Parent $RapportHtml
    if ($map -and -not (Test-Path -LiteralPath $map)) {
        New-Item -ItemType Directory -Path $map -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($RapportHtml, $html.ToString(), [System.Text.UTF8Encoding]::new($true))

    Write-Host ""
    Write-Host "Rapport opgeslagen: $RapportHtml" -ForegroundColor Green

    if ($Openen) {
        try {
            # Start-Process pakt de standaardbrowser; Invoke-Item werkt niet
            # altijd als de tool zonder console draait.
            Start-Process -FilePath $RapportHtml -ErrorAction Stop
            Write-Host "Geopend in je browser."
        }
        catch {
            Write-Host "Openen lukte niet: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Dubbelklik het bestand hierboven."
        }
    }
    else {
        Write-Host "Dubbelklikken opent het in je browser."
    }
}

Write-Host ""
Write-Host "Klaar." -ForegroundColor Green
