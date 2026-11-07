<#
.SYNOPSIS
Maakt een overzicht van alle bestanden onder een map, met een leesbaar HTML-rapport.

.DESCRIPTION
Vervangt de oude eenregelige variant die alleen naar het scherm schreef. Je
krijgt een rapport met de verdeling per submap en per bestandstype, de grootste
bestanden, en desgewenst de volledige bestandslijst.

Bedoeld om snel te zien wat er in een dossier zit: hoeveel, hoe groot, van
wanneer, en waar de ruimte zit.

Er wordt niets gewijzigd, alleen gelezen.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "De map die je wilt inventariseren")]
    [string]$Map,

    [Parameter(HelpMessage = "Pad voor het HTML-rapport. Alleen een bestandsnaam mag ook: dan komt het in Documenten. Leeg laten in combinatie met Openen geeft een automatische naam.")]
    [string]$RapportHtml = "",

    [Parameter(HelpMessage = "Alleen deze bestandstypen, bijvoorbeeld *.xlsm en *.pdf")]
    [string[]]$Filter = @("*.*"),

    [Parameter(HelpMessage = "Aanvinken om submappen NIET mee te nemen")]
    [switch]$AlleenBovensteMap,

    [Parameter(HelpMessage = "Ook de volledige bestandslijst in het rapport. Bij duizenden bestanden wordt het rapport dan groot.")]
    [switch]$AlleBestandenTonen,

    [Parameter(HelpMessage = "Extra CSV-bestand om in Excel mee door te rekenen")]
    [string]$RapportCsv = "",

    [Parameter(HelpMessage = "Het rapport na afloop meteen openen (standaard aan)")]
    [switch]$Openen
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

try { Add-Type -AssemblyName System.Web -ErrorAction Stop } catch { }
function HtmlVeilig {
    param([string]$Tekst)
    if (-not $Tekst) { return "" }
    return $Tekst.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

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
    Tekent één map met zijn inhoud, en geeft de HTML plus aantal en omvang
    terug. Die staan in de kop van de map, die vóór de kinderen komt, dus de
    kinderen worden eerst getekend.
    #>
    param($Knoop, [string]$Naam, [int]$Diepte, [int]$OpenTot = 1)

    $binnen = New-Object System.Text.StringBuilder
    $aantal = 0
    [long]$bytes = 0

    foreach ($submap in $Knoop.mappen.Keys) {
        $kind = Teken-Knoop -Knoop $Knoop.mappen[$submap] -Naam $submap `
            -Diepte ($Diepte + 1) -OpenTot $OpenTot
        $null = $binnen.Append($kind.html)
        $aantal += $kind.aantal
        $bytes += $kind.bytes
    }

    if (@($Knoop.bestanden).Count -gt 0) {
        $null = $binnen.AppendLine("<ul class='bestanden'>")
        foreach ($bestand in ($Knoop.bestanden | Sort-Object -Property @{Expression = "bytes"; Descending = $true })) {
            $aantal++
            $bytes += [long]$bestand.bytes
            $klasse = if ($bestand.alleenlezen) { "s-slot" } else { "s-ok" }
            $extra = if ($bestand.alleenlezen) { " alleen-lezen" } else { "" }
            $null = $binnen.AppendLine("<li><span class='stip $klasse'></span>" +
                "<span>$(HtmlVeilig $bestand.naam)</span>" +
                "<span class='grootte'>$(Leesbaar $bestand.bytes)$extra</span></li>")
        }
        $null = $binnen.AppendLine("</ul>")
    }

    $html = New-Object System.Text.StringBuilder
    $open = if ($Diepte -lt $OpenTot) { " open" } else { "" }
    $null = $html.AppendLine("<details class='boom'$open><summary>" +
        "<span class='mapnaam'>$(HtmlVeilig $Naam)</span>" +
        "<span class='tel'>$aantal bestand$(if ($aantal -ne 1) { 'en' })</span>" +
        "<span class='grootte'>$(Leesbaar $bytes)</span></summary>")
    $null = $html.AppendLine("<div class='kinderen'>$($binnen.ToString())</div></details>")

    return @{ html = $html.ToString(); aantal = $aantal; bytes = $bytes }
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

Write-Host "Map    : $Map"
Write-Host "Filter : $($Filter -join ', ')"
Write-Host "Submappen: $(if ($AlleenBovensteMap) { 'nee' } else { 'ja' })"
Write-Host ""

if ($AlleenBovensteMap) {
    $bestanden = Get-ChildItem -LiteralPath $Map -File | Where-Object {
        $naam = $_.Name
        ($Filter | Where-Object { $naam -like $_ }).Count -gt 0
    }
}
else {
    $bestanden = @(Get-ChildItem -LiteralPath $Map -Include $Filter -File -Recurse `
            -ErrorAction SilentlyContinue | Tel-Onderweg -Wat "Bestanden zoeken")
}
Stop-Voortgang -Wat "Bestanden zoeken"

if (-not $bestanden -or $bestanden.Count -eq 0) {
    Write-Host "Geen bestanden gevonden." -ForegroundColor Yellow
    exit 0
}

$basislengte = $Map.TrimEnd('\').Length + 1
$totaalBytes = ($bestanden | Measure-Object -Property Length -Sum).Sum
$grootste = $bestanden | Sort-Object Length -Descending | Select-Object -First 20
$oudste = ($bestanden | Sort-Object LastWriteTime | Select-Object -First 1).LastWriteTime
$nieuwste = ($bestanden | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
$alleenLezen = @($bestanden | Where-Object { $_.IsReadOnly }).Count

# per submap
$permap = $bestanden | Group-Object DirectoryName | ForEach-Object {
    $relatief = if ($_.Name.Length -gt $basislengte) { $_.Name.Substring($basislengte) } else { "(hoofdmap)" }
    $som = ($_.Group | Measure-Object -Property Length -Sum).Sum
    [pscustomobject]@{
        Map      = $relatief
        Aantal   = $_.Count
        Bytes    = $som
        Aandeel  = if ($totaalBytes -gt 0) { [math]::Round($som / $totaalBytes * 100, 1) } else { 0 }
        Nieuwste = ($_.Group | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
    }
} | Sort-Object Bytes -Descending

# per bestandstype
$pertype = $bestanden | Group-Object { if ($_.Extension) { $_.Extension.ToLower() } else { "(geen)" } } |
    ForEach-Object {
        $som = ($_.Group | Measure-Object -Property Length -Sum).Sum
        [pscustomobject]@{
            Type    = $_.Name
            Aantal  = $_.Count
            Bytes   = $som
            Aandeel = if ($totaalBytes -gt 0) { [math]::Round($som / $totaalBytes * 100, 1) } else { 0 }
        }
    } | Sort-Object Bytes -Descending

# --------------------------------------------------------------------------
# In beeld
# --------------------------------------------------------------------------

Write-Host "$($bestanden.Count) bestanden in $($permap.Count) mappen, $(Leesbaar $totaalBytes)" `
    -ForegroundColor Green
Write-Host "Gewijzigd tussen $($oudste.ToString('yyyy-MM-dd')) en $($nieuwste.ToString('yyyy-MM-dd'))"
if ($alleenLezen -gt 0) { Write-Host "$alleenLezen bestanden staan op alleen-lezen" }

Write-Host ""
Write-Host "PER BESTANDSTYPE" -ForegroundColor White
foreach ($t in $pertype) {
    Write-Host ("  {0,-10} {1,6} stuks {2,12} {3,6}%" -f $t.Type, $t.Aantal, (Leesbaar $t.Bytes), $t.Aandeel)
}

Write-Host ""
Write-Host "PER MAP, grootste eerst" -ForegroundColor White
foreach ($m in $permap) {
    Write-Host ("  {0,6} stuks {1,12} {2,6}%   {3}" -f $m.Aantal, (Leesbaar $m.Bytes), $m.Aandeel, $m.Map)
}

Write-Host ""
Write-Host "TIEN GROOTSTE BESTANDEN" -ForegroundColor Cyan
foreach ($b in ($grootste | Select-Object -First 10)) {
    Write-Host ("  {0,12}  {1}" -f (Leesbaar $b.Length), $b.FullName.Substring($basislengte))
}

# --------------------------------------------------------------------------
# HTML-rapport
# --------------------------------------------------------------------------

$maakRapport = [bool]$RapportHtml -or $Openen

if ($maakRapport) {

    $RapportHtml = Bepaal-Rapportpad -Opgegeven $RapportHtml -Extensie ".html" `
        -Standaardnaam (Veilige-Naam ("Inventarisatie " +
            (Split-Path -Leaf $Map.TrimEnd('\')) + " " +
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
.mapnaam { font-family: Consolas, monospace; font-weight: 600; }
.tel { color: #6E6E73; font-size: 12px; }
.kinderen { margin-left: 16px; border-left: 1px solid #E9E9EC; padding-left: 10px; }
ul.bestanden { list-style: none; margin: 2px 0 6px; padding: 0 0 0 18px; }
ul.bestanden li { display: flex; align-items: center; gap: 8px; padding: 2px 0;
                  font-family: Consolas, monospace; font-size: 12.5px; }
.stip { width: 8px; height: 8px; border-radius: 50%; flex: 0 0 8px; }
.s-ok { background: #CFD2D7; }
.s-slot { background: #E8A33D; }
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
    $null = $html.AppendLine('<title>Inventarisatie</title>')
    $null = $html.AppendLine("<style>$css</style></head><body>")

    $null = $html.AppendLine("<h1>Inventarisatie: $(HtmlVeilig (Split-Path -Leaf $Map.TrimEnd('\')))</h1>")
    $null = $html.AppendLine("<p class='sub'>Gemaakt op $(Get-Date -Format 'dddd d MMMM yyyy, HH:mm')" +
        "$(if ($Filter -join ',' -ne '*.*') { " &middot; filter: $(HtmlVeilig ($Filter -join ', '))" })</p>")
    $null = $html.AppendLine("<div class='pad-groot'>$(HtmlVeilig $Map)</div>")

    $null = $html.AppendLine("<div class='kaarten'>")
    $null = $html.AppendLine("<div class='kaart'><div class='getal'>$($bestanden.Count)</div><div class='label'>bestanden</div></div>")
    $null = $html.AppendLine("<div class='kaart b'><div class='getal'>$($permap.Count)</div><div class='label'>mappen</div></div>")
    $null = $html.AppendLine("<div class='kaart c'><div class='getal'>$(Leesbaar $totaalBytes)</div><div class='label'>totaal</div></div>")
    $null = $html.AppendLine("<div class='kaart d'><div class='getal'>$($pertype.Count)</div><div class='label'>bestandstypen</div></div>")
    $null = $html.AppendLine("</div>")

    $null = $html.AppendLine("<p class='sub'>Gewijzigd tussen $($oudste.ToString('d MMMM yyyy')) en " +
        "$($nieuwste.ToString('d MMMM yyyy'))" +
        "$(if ($alleenLezen -gt 0) { " &middot; $alleenLezen bestanden op alleen-lezen" })</p>")

    # ---- per bestandstype ----
    $null = $html.AppendLine("<h2>Per bestandstype</h2>")
    $null = $html.AppendLine("<table><thead><tr><th>Type</th><th class='num'>Aantal</th>" +
        "<th class='num'>Grootte</th><th class='num'>Aandeel</th><th></th></tr></thead><tbody>")
    foreach ($t in $pertype) {
        $null = $html.AppendLine("<tr><td class='mono'>$(HtmlVeilig $t.Type)</td>" +
            "<td class='num'>$($t.Aantal)</td><td class='num'>$(Leesbaar $t.Bytes)</td>" +
            "<td class='num'>$($t.Aandeel)%</td>" +
            "<td><span class='spoor'><span style='width:$($t.Aandeel)%'></span></span></td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")

    # ---- per map ----
    $null = $html.AppendLine("<h2>Per map, grootste eerst</h2>")
    $null = $html.AppendLine("<table><thead><tr><th>Map</th><th class='num'>Aantal</th>" +
        "<th class='num'>Grootte</th><th class='num'>Aandeel</th><th></th>" +
        "<th class='num'>Laatst gewijzigd</th></tr></thead><tbody>")
    foreach ($m in $permap) {
        $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $m.Map)</td>" +
            "<td class='num'>$($m.Aantal)</td><td class='num'>$(Leesbaar $m.Bytes)</td>" +
            "<td class='num'>$($m.Aandeel)%</td>" +
            "<td><span class='spoor'><span style='width:$($m.Aandeel)%'></span></span></td>" +
            "<td class='num'>$($m.Nieuwste.ToString('yyyy-MM-dd HH:mm'))</td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")

    # ---- mappenboom ----
    $wortel = Nieuwe-Knoop
    $boomNr = 0
    foreach ($b in $bestanden) {
        $boomNr++
        Meld-Voortgang -Wat "Rapport opbouwen" -Nu $boomNr -Totaal $bestanden.Count
        $relatief = $b.FullName.Substring($basislengte)
        Voeg-InBoom -Knoop $wortel -Delen ($relatief -split '\\') -Bestand @{
            naam = $b.Name; bytes = $b.Length; alleenlezen = $b.IsReadOnly
        }
    }
    Stop-Voortgang -Wat "Rapport opbouwen"

    $null = $html.AppendLine("<h2>Mapstructuur</h2>")
    if ($alleenLezen -gt 0) {
        $null = $html.AppendLine("<div class='legenda'>" +
            "<span class='item'><span class='stip s-ok'></span>gewoon bestand</span>" +
            "<span class='item'><span class='stip s-slot'></span>alleen-lezen</span></div>")
    }
    $null = $html.AppendLine("<div class='knoppen'>" +
        "<button onclick=""alles(true)"">Alles uitklappen</button>" +
        "<button onclick=""alles(false)"">Alles inklappen</button></div>")
    $null = $html.AppendLine("<div class='boomvak'>")
    $getekend = Teken-Knoop -Knoop $wortel -Naam (Split-Path -Leaf $Map.TrimEnd('\')) `
        -Diepte 0 -OpenTot 2
    $null = $html.AppendLine($getekend.html)
    $null = $html.AppendLine("</div>")

    # ---- grootste bestanden ----
    $null = $html.AppendLine("<h2>Twintig grootste bestanden</h2>")
    $null = $html.AppendLine("<table><thead><tr><th class='num'>Grootte</th><th>Bestand</th>" +
        "<th class='num'>Gewijzigd</th></tr></thead><tbody>")
    foreach ($b in $grootste) {
        $null = $html.AppendLine("<tr><td class='num'>$(Leesbaar $b.Length)</td>" +
            "<td class='pad'>$(HtmlVeilig $b.FullName.Substring($basislengte))</td>" +
            "<td class='num'>$($b.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")

    # ---- volledige lijst, alleen op verzoek ----
    if ($AlleBestandenTonen) {
        $null = $html.AppendLine("<h2>Alle bestanden <span class='badge'>$($bestanden.Count)</span></h2>")
        $null = $html.AppendLine("<table><thead><tr><th>Map</th><th>Bestand</th>" +
            "<th class='num'>Grootte</th><th class='num'>Gewijzigd</th>" +
            "<th class='num'>Alleen-lezen</th></tr></thead><tbody>")
        foreach ($b in ($bestanden | Sort-Object FullName)) {
            $relatief = $b.FullName.Substring($basislengte)
            $submap = Split-Path -Parent $relatief
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $(if ($submap) { $submap } else { '.' }))</td>" +
                "<td class='pad'>$(HtmlVeilig $b.Name)</td>" +
                "<td class='num'>$(Leesbaar $b.Length)</td>" +
                "<td class='num'>$($b.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))</td>" +
                "<td class='num'>$(if ($b.IsReadOnly) { 'ja' } else { '' })</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }
    else {
        $null = $html.AppendLine("<footer>De volledige bestandslijst staat niet in dit " +
            "rapport. Vink 'AlleBestandenTonen' aan als je die erbij wilt.</footer>")
    }

    $null = $html.AppendLine("<footer>$(HtmlVeilig $Map)<br>" +
        "$($bestanden.Count) bestanden, $(Leesbaar $totaalBytes)</footer>")
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
    if ($Openen) { Openen-Bestand $RapportHtml }
    else { Write-Host "Dubbelklikken opent het in je browser." }
}

# --------------------------------------------------------------------------
# CSV, voor wie ermee wil doorrekenen
# --------------------------------------------------------------------------

if ($RapportCsv) {
    $RapportCsv = Bepaal-Rapportpad -Opgegeven $RapportCsv -Extensie ".csv" `
        -Standaardnaam (Veilige-Naam ("Inventarisatie " +
            (Split-Path -Leaf $Map.TrimEnd('\')) + " " +
            (Get-Date -Format 'yyyyMMdd-HHmm') + ".csv"))

    $rijen = $bestanden | ForEach-Object {
        $relatiefPad = $_.FullName.Substring($basislengte)
        [pscustomobject]@{
            Submap       = Split-Path -Parent $relatiefPad
            Bestandsnaam = $_.Name
            Extensie     = $_.Extension
            Bytes        = $_.Length
            Grootte      = Leesbaar $_.Length
            Gewijzigd    = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
            AlleenLezen  = $_.IsReadOnly
            RelatiefPad  = $relatiefPad
            VolledigPad  = $_.FullName
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
