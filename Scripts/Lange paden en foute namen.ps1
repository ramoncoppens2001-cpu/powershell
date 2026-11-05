<#
.SYNOPSIS
Zoekt te lange paden en bestandsnamen die problemen geven bij kopiëren, zippen of SharePoint.

.DESCRIPTION
Windows kapt van huis uit af op 260 tekens voor een volledig pad. Bij mappen als
T:\Dossiers\2026\ING Bank N.V\Albatross Portfolio\... met lange klantnamen is dat
sneller bereikt dan je denkt, en dan mislukt zippen, kopiëren of een back-up
zonder duidelijke reden.

Daarnaast weigeren SharePoint en OneDrive bepaalde tekens en namen. Wie ooit een
map naar Teams of SharePoint tilt, loopt daar tegenaan. Ook nu al kan het
synchronisatiefouten geven.

Het rapport laat per bestand zien wat er aan de hand is en hoeveel tekens eraf
moeten, met een overzicht per dossiermap zodat je ziet waar het zit.

Er wordt niets gewijzigd of hernoemd, alleen gelezen en gerapporteerd.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "De map die je wilt controleren")]
    [string]$Map,

    [Parameter(HelpMessage = "Pad voor het HTML-rapport. Alleen een bestandsnaam mag ook: dan komt het in Documenten.")]
    [string]$RapportHtml = "",

    [Parameter(HelpMessage = "Vanaf hoeveel tekens een pad te lang heet. 240 laat marge tot de Windows-grens van 260.")]
    [int]$MaxPadlengte = 240,

    [Parameter(HelpMessage = "Vanaf hoeveel tekens één bestands- of mapnaam te lang heet")]
    [int]$MaxNaamlengte = 128,

    [Parameter(HelpMessage = "Ook controleren op tekens en namen die SharePoint en OneDrive weigeren")]
    [switch]$SharePointControleOverslaan,

    [Parameter(HelpMessage = "Het rapport na afloop meteen openen (standaard aan)")]
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

function Veilige-Naam {
    param([string]$Tekst)
    $ongeldig = [System.IO.Path]::GetInvalidFileNameChars() -join ''
    return ($Tekst -replace ("[{0}]" -f [regex]::Escape($ongeldig)), '_')
}

function Bepaal-Rapportpad {
    param([string]$Opgegeven, [string]$Standaardnaam, [string]$Extensie)
    $pad = "$Opgegeven".Trim().Trim('"')
    if (-not $pad) { $pad = $Standaardnaam }
    if (Test-Path -LiteralPath $pad -PathType Container) { $pad = Join-Path $pad $Standaardnaam }
    if (-not [System.IO.Path]::GetExtension($pad)) { $pad += $Extensie }
    if (-not [System.IO.Path]::IsPathRooted($pad)) {
        $basis = if (Split-Path -Parent $pad) { (Get-Location).Path }
        else { [Environment]::GetFolderPath('MyDocuments') }
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

function Dossier {
    param([string]$RelatiefPad)
    $delen = $RelatiefPad -split '\\'
    if ($delen.Count -le 1) { return "(hoofdmap)" }
    return $delen[0]
}

# --------------------------------------------------------------------------
# Regels voor namen
# --------------------------------------------------------------------------

# SharePoint en OneDrive weigeren deze tekens in namen. \ / : * ? " < > | kunnen
# in Windows al niet, dus die kunnen we niet aantreffen; # en % geven wel
# problemen en komen in de praktijk voor.
$SP_TEKENS = @('#', '%', '&', '{', '}', '~')

# Namen die MS-DOS ooit voor apparaten gebruikte. Windows weigert ze nog steeds.
$GERESERVEERD = @('CON', 'PRN', 'AUX', 'NUL',
    'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
    'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')

function Beoordeel-Naam {
    <# Geeft een lijst bezwaren tegen deze naam. Leeg = in orde. #>
    param([string]$Naam)

    $bezwaren = @()
    $tijdelijk = $Naam.StartsWith('~$')

    if ($Naam.Length -gt $MaxNaamlengte) {
        $bezwaren += "naam is $($Naam.Length) tekens"
    }
    if ($Naam -ne $Naam.Trim()) {
        $bezwaren += "begint of eindigt met een spatie"
    }
    if ($Naam.EndsWith('.')) {
        $bezwaren += "eindigt op een punt"
    }
    # 'Aantekening .txt' mag van Windows, maar hulpprogramma's die de naam
    # opschonen kunnen het bestand daarna niet meer vinden.
    $kern = [System.IO.Path]::GetFileNameWithoutExtension($Naam)
    if ($kern -ne $kern.TrimEnd() -and -not $tijdelijk) {
        $bezwaren += "spatie vlak voor de extensie"
    }

    if (-not $SharePointControleOverslaan) {
        if ($tijdelijk) {
            # De tilde hier melden is dubbelop; dit is een achtergebleven
            # vergrendelingsbestand van Excel en dat is het echte bezwaar.
            $bezwaren += "tijdelijk Excel-bestand, hoort hier niet te blijven staan"
        }
        else {
            $gevonden = @($SP_TEKENS | Where-Object { $Naam.Contains($_) })
            if ($gevonden.Count -gt 0) {
                $bezwaren += "bevat $($gevonden -join ' ') - SharePoint en OneDrive weigeren dat"
            }
        }
        if ($GERESERVEERD -contains $kern.Trim().ToUpper()) {
            $bezwaren += "'$($kern.Trim())' is een door Windows gereserveerde naam"
        }
    }
    return $bezwaren
}

# --------------------------------------------------------------------------
# Doorlopen
# --------------------------------------------------------------------------

Write-Host "Map           : $Map"
Write-Host "Max padlengte : $MaxPadlengte tekens"
Write-Host "Max naamlengte: $MaxNaamlengte tekens"
Write-Host "SharePoint    : $(if ($SharePointControleOverslaan) { 'niet gecontroleerd' } else { 'gecontroleerd' })"
Write-Host ""
Write-Host "Doorlopen..."

$basislengte = $Map.TrimEnd('\').Length + 1
$alles = @(Get-ChildItem -LiteralPath $Map -Recurse -Force -ErrorAction SilentlyContinue)

Write-Host "$($alles.Count) items gevonden. Beoordelen..."

$telAlles = 0
$problemen = New-Object System.Collections.Generic.List[object]
$langstePad = 0

foreach ($item in $alles) {
    $telAlles++
    if ($telAlles % 2000 -eq 0) { Write-Host "  $telAlles van $($alles.Count)..." }

    $vol = $item.FullName
    if ($vol.Length -gt $langstePad) { $langstePad = $vol.Length }

    $bezwaren = @()
    if ($vol.Length -gt $MaxPadlengte) {
        $bezwaren += "pad is $($vol.Length) tekens, $($vol.Length - $MaxPadlengte) te veel"
    }
    $bezwaren += Beoordeel-Naam -Naam $item.Name

    if ($bezwaren.Count -eq 0) { continue }

    $relatief = if ($vol.Length -gt $basislengte) { $vol.Substring($basislengte) } else { $item.Name }
    $problemen.Add([pscustomobject]@{
            Relatief  = $relatief
            Naam      = $item.Name
            Submap    = Split-Path -Parent $relatief
            Dossier   = Dossier $relatief
            IsMap     = $item.PSIsContainer
            Padlengte = $vol.Length
            TeVeel    = [math]::Max(0, $vol.Length - $MaxPadlengte)
            Bytes     = if ($item.PSIsContainer) { $null } else { $item.Length }
            Bezwaren  = $bezwaren
            TeLang    = $vol.Length -gt $MaxPadlengte
        })
}

$teLang = @($problemen | Where-Object { $_.TeLang })
$fouteNamen = @($problemen | Where-Object { -not $_.TeLang })

Write-Host ""
Write-Host "$($alles.Count) items nagelopen, langste pad is $langstePad tekens." -ForegroundColor Green
Write-Host "$($teLang.Count) te lange paden, $($fouteNamen.Count) namen met een bezwaar." `
    -ForegroundColor $(if ($problemen.Count -gt 0) { "Yellow" } else { "Green" })

if ($teLang.Count -gt 0) {
    Write-Host ""
    Write-Host "TE LANGE PADEN, langste eerst" -ForegroundColor Yellow
    foreach ($p in ($teLang | Sort-Object Padlengte -Descending | Select-Object -First 20)) {
        Write-Host ("  {0,4} tekens ({1,3} te veel)  {2}" -f $p.Padlengte, $p.TeVeel, $p.Relatief)
    }
    if ($teLang.Count -gt 20) { Write-Host "  ... en $($teLang.Count - 20) meer, zie het rapport" }
}

if ($fouteNamen.Count -gt 0) {
    Write-Host ""
    Write-Host "NAMEN MET EEN BEZWAAR" -ForegroundColor Cyan
    foreach ($p in ($fouteNamen | Select-Object -First 20)) {
        Write-Host "  $($p.Relatief)"
        foreach ($b in $p.Bezwaren) { Write-Host "      $b" }
    }
    if ($fouteNamen.Count -gt 20) { Write-Host "  ... en $($fouteNamen.Count - 20) meer, zie het rapport" }
}

# --------------------------------------------------------------------------
# HTML-rapport
# --------------------------------------------------------------------------

$maakRapport = [bool]$RapportHtml -or $Openen

if ($maakRapport) {

    $RapportHtml = Bepaal-Rapportpad -Opgegeven $RapportHtml -Extensie ".html" `
        -Standaardnaam (Veilige-Naam ("Padcontrole " + (Split-Path -Leaf $Map.TrimEnd('\')) +
            " " + (Get-Date -Format 'yyyyMMdd-HHmm') + ".html"))

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
         padding: 14px 18px; min-width: 160px; border-left: 4px solid #2C3E50; }
.kaart .getal { font-size: 28px; font-weight: 600; line-height: 1.1; }
.kaart .label { color: #6E6E73; font-size: 12px; margin-top: 2px; }
.kaart.lang { border-left-color: #E8A33D; }
.kaart.naam { border-left-color: #B85CC7; }
.kaart.ok { border-left-color: #34A853; }
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
.b-lang { background: #FDF0DC; color: #8A5A12; }
.b-naam { background: #F6E6F8; color: #6E2A7A; }
.b-ok { background: #E4F4E7; color: #1E6B2E; }
.balk { background: #EDEDF0; border-radius: 3px; height: 8px; width: 90px;
        display: inline-block; vertical-align: middle; overflow: hidden; }
.balk span { display: block; height: 100%; background: #E8A33D; }
ul.bezwaren { margin: 0; padding-left: 18px; }
ul.bezwaren li { padding: 1px 0; }
.leeg { color: #8E8E93; font-style: italic; padding: 14px 12px; background: #fff;
        border: 1px solid #D1D1D6; border-radius: 10px; }
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
.s-lang { background: #E8A33D; }
.s-naam { background: #B85CC7; }
.grootte { color: #6E6E73; font-size: 12px; margin-left: auto; white-space: nowrap; }
.legenda { display: flex; gap: 16px; flex-wrap: wrap; margin: 0 0 10px;
           color: #6E6E73; font-size: 12px; align-items: center; }
.legenda .item { display: flex; align-items: center; gap: 5px; }
.knoppen { margin: 0 0 10px; display: flex; gap: 8px; }
.knoppen button { font: inherit; font-size: 12px; padding: 4px 10px; cursor: pointer;
                  border: 1px solid #D1D1D6; border-radius: 6px; background: #fff;
                  color: #1D1D1F; }
.knoppen button:hover { background: #F0F0F3; }
footer { margin-top: 40px; color: #8E8E93; font-size: 12px; }
@media print {
  body { background: #fff; padding: 0; }
  th { position: static; }
  h2 { page-break-after: avoid; }
  tr { page-break-inside: avoid; }
  .knoppen { display: none; }
}
'@

    # ---- boom met alleen de probleemgevallen ----
    function Nieuwe-Knoop { return [ordered]@{ mappen = [ordered]@{}; bestanden = @() } }

    function Voeg-InBoom {
        param($Knoop, [string[]]$Delen, $Bestand)
        if ($Delen.Count -le 1) {
            $Knoop.bestanden = @($Knoop.bestanden) + @($Bestand); return
        }
        if (-not $Knoop.mappen.Contains($Delen[0])) { $Knoop.mappen[$Delen[0]] = Nieuwe-Knoop }
        Voeg-InBoom -Knoop $Knoop.mappen[$Delen[0]] -Delen $Delen[1..($Delen.Count - 1)] -Bestand $Bestand
    }

    function Teken-Knoop {
        param($Knoop, [string]$Naam, [int]$Diepte, [int]$OpenTot = 2)
        $binnen = New-Object System.Text.StringBuilder
        $telling = @{ lang = 0; naam = 0 }

        foreach ($submap in $Knoop.mappen.Keys) {
            $kind = Teken-Knoop -Knoop $Knoop.mappen[$submap] -Naam $submap `
                -Diepte ($Diepte + 1) -OpenTot $OpenTot
            $null = $binnen.Append($kind.html)
            $telling.lang += $kind.telling.lang
            $telling.naam += $kind.telling.naam
        }
        if (@($Knoop.bestanden).Count -gt 0) {
            $null = $binnen.AppendLine("<ul class='bestanden'>")
            foreach ($b in ($Knoop.bestanden | Sort-Object naam)) {
                $telling[$b.klasse]++
                $null = $binnen.AppendLine("<li><span class='stip s-$($b.klasse)'></span>" +
                    "<span>$(HtmlVeilig $b.naam)</span>" +
                    "<span class='grootte'>$(HtmlVeilig $b.uitleg)</span></li>")
            }
            $null = $binnen.AppendLine("</ul>")
        }

        $merken = New-Object System.Text.StringBuilder
        if ($telling.lang -gt 0) {
            $null = $merken.Append("<span class='badge b-lang'>$($telling.lang) te lang</span>")
        }
        if ($telling.naam -gt 0) {
            $null = $merken.Append("<span class='badge b-naam'>$($telling.naam) naam</span>")
        }

        $html = New-Object System.Text.StringBuilder
        $null = $html.AppendLine("<details class='boom' open><summary>" +
            "<span class='mapnaam'>$(HtmlVeilig $Naam)</span>$($merken.ToString())</summary>")
        $null = $html.AppendLine("<div class='kinderen'>$($binnen.ToString())</div></details>")
        return @{ html = $html.ToString(); telling = $telling }
    }

    $html = New-Object System.Text.StringBuilder
    $null = $html.AppendLine('<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8">')
    $null = $html.AppendLine('<title>Padcontrole</title>')
    $null = $html.AppendLine("<style>$css</style></head><body>")

    $null = $html.AppendLine("<h1>Padcontrole: $(HtmlVeilig (Split-Path -Leaf $Map.TrimEnd('\')))</h1>")
    $null = $html.AppendLine("<p class='sub'>Gemaakt op $(Get-Date -Format 'dddd d MMMM yyyy, HH:mm')" +
        " &middot; grens $MaxPadlengte tekens voor het pad, $MaxNaamlengte voor een naam</p>")
    $null = $html.AppendLine("<div class='pad-groot'>$(HtmlVeilig $Map)</div>")

    $null = $html.AppendLine("<div class='kaarten'>")
    $null = $html.AppendLine("<div class='kaart'><div class='getal'>$($alles.Count)</div><div class='label'>items nagelopen</div></div>")
    $null = $html.AppendLine("<div class='kaart lang'><div class='getal'>$($teLang.Count)</div><div class='label'>te lange paden</div></div>")
    $null = $html.AppendLine("<div class='kaart naam'><div class='getal'>$($fouteNamen.Count)</div><div class='label'>namen met bezwaar</div></div>")
    $null = $html.AppendLine("<div class='kaart $(if ($langstePad -gt 260) { 'lang' } else { 'ok' })'>" +
        "<div class='getal'>$langstePad</div><div class='label'>langste pad in tekens</div></div>")
    $null = $html.AppendLine("</div>")

    if ($problemen.Count -eq 0) {
        $null = $html.AppendLine("<p class='sub'><span class='badge b-ok'>alles in orde</span> " +
            "Geen te lange paden en geen namen met bezwaar.</p>")
    }
    else {
        # ---- overzicht per dossiermap ----
        $perDossier = $problemen | Group-Object Dossier | ForEach-Object {
            [pscustomobject]@{
                Map    = $_.Name
                Lang   = @($_.Group | Where-Object { $_.TeLang }).Count
                Naam   = @($_.Group | Where-Object { -not $_.TeLang }).Count
                Totaal = $_.Count
            }
        } | Sort-Object Totaal -Descending

        $null = $html.AppendLine("<h2>Overzicht per map</h2>")
        $null = $html.AppendLine("<table><thead><tr><th>Map</th><th class='num'>Te lang</th>" +
            "<th class='num'>Naam</th><th class='num'>Totaal</th></tr></thead><tbody>")
        foreach ($d in $perDossier) {
            $null = $html.AppendLine("<tr class='let'><td class='pad'>$(HtmlVeilig $d.Map)</td>" +
                "<td class='num'>$($d.Lang)</td><td class='num'>$($d.Naam)</td>" +
                "<td class='num'>$($d.Totaal)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")

        # ---- structuur ----
        $wortel = Nieuwe-Knoop
        foreach ($p in $problemen) {
            $klasse = if ($p.TeLang) { "lang" } else { "naam" }
            $uitleg = if ($p.TeLang) { "$($p.Padlengte) tekens, $($p.TeVeel) te veel" }
            else { $p.Bezwaren -join "; " }
            Voeg-InBoom -Knoop $wortel -Delen ($p.Relatief -split '\\') -Bestand @{
                naam = $p.Naam; klasse = $klasse; uitleg = $uitleg
            }
        }
        $null = $html.AppendLine("<h2>Structuur, alleen de probleemgevallen</h2>")
        $null = $html.AppendLine("<div class='legenda'>" +
            "<span class='item'><span class='stip s-lang'></span>pad te lang</span>" +
            "<span class='item'><span class='stip s-naam'></span>naam met bezwaar</span></div>")
        $null = $html.AppendLine("<div class='knoppen'>" +
            "<button onclick=""alles(true)"">Alles uitklappen</button>" +
            "<button onclick=""alles(false)"">Alles inklappen</button></div>")
        $null = $html.AppendLine("<div class='boomvak'>")
        $getekend = Teken-Knoop -Knoop $wortel -Naam (Split-Path -Leaf $Map.TrimEnd('\')) -Diepte 0
        $null = $html.AppendLine($getekend.html)
        $null = $html.AppendLine("</div>")

        # ---- te lange paden ----
        $null = $html.AppendLine("<h2>Te lange paden <span class='badge b-lang'>$($teLang.Count)</span></h2>")
        if ($teLang.Count -eq 0) {
            $null = $html.AppendLine("<div class='leeg'>Geen enkel pad boven de $MaxPadlengte tekens.</div>")
        }
        else {
            $null = $html.AppendLine("<table><thead><tr><th class='num'>Tekens</th>" +
                "<th class='num'>Te veel</th><th></th><th>Pad onder de bronmap</th>" +
                "<th class='num'>Grootte</th></tr></thead><tbody>")
            foreach ($p in ($teLang | Sort-Object Padlengte -Descending)) {
                $breedte = [math]::Min(100, [math]::Round($p.Padlengte / 260 * 100))
                $null = $html.AppendLine("<tr><td class='num'>$($p.Padlengte)</td>" +
                    "<td class='num'>$($p.TeVeel)</td>" +
                    "<td><span class='balk'><span style='width:$breedte%'></span></span></td>" +
                    "<td class='pad'>$(HtmlVeilig $p.Relatief)</td>" +
                    "<td class='num'>$(Leesbaar $p.Bytes)</td></tr>")
            }
            $null = $html.AppendLine("</tbody></table>")
        }

        # ---- namen met bezwaar ----
        $null = $html.AppendLine("<h2>Namen met een bezwaar <span class='badge b-naam'>$($fouteNamen.Count)</span></h2>")
        if ($fouteNamen.Count -eq 0) {
            $null = $html.AppendLine("<div class='leeg'>Alle namen zijn bruikbaar.</div>")
        }
        else {
            $null = $html.AppendLine("<table><thead><tr><th>Pad onder de bronmap</th>" +
                "<th>Wat er aan de hand is</th></tr></thead><tbody>")
            foreach ($p in ($fouteNamen | Sort-Object Relatief)) {
                $lijst = ($p.Bezwaren | ForEach-Object { "<li>$(HtmlVeilig $_)</li>" }) -join ""
                $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $p.Relatief)</td>" +
                    "<td><ul class='bezwaren'>$lijst</ul></td></tr>")
            }
            $null = $html.AppendLine("</tbody></table>")
        }
    }

    $null = $html.AppendLine("<footer>Windows kapt van huis uit af op 260 tekens voor een " +
        "volledig pad. De grens van $MaxPadlengte laat marge, want bij kopiëren naar een " +
        "diepere map komen er tekens bij.<br>" +
        "SharePoint en OneDrive weigeren onder meer $(HtmlVeilig ($SP_TEKENS -join ' ')) in " +
        "namen, plus namen die op een punt of spatie eindigen.<br>" +
        "$(HtmlVeilig $Map)</footer>")
    $null = $html.AppendLine(@'
<script>
function alles(open) {
  document.querySelectorAll("details.boom").forEach(function (d) { d.open = open; });
}
</script>
'@)
    $null = $html.AppendLine("</body></html>")

    $rapportmap = Split-Path -Parent $RapportHtml
    if ($rapportmap -and -not (Test-Path -LiteralPath $rapportmap)) {
        New-Item -ItemType Directory -Path $rapportmap -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($RapportHtml, $html.ToString(), [System.Text.UTF8Encoding]::new($true))

    Write-Host ""
    Write-Host "Rapport opgeslagen: $RapportHtml" -ForegroundColor Green
    if ($Openen) { Openen-Bestand $RapportHtml }
    else { Write-Host "Dubbelklikken opent het in je browser." }
}

Write-Host ""
Write-Host "Klaar." -ForegroundColor Green
