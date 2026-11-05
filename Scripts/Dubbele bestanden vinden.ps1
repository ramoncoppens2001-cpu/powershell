<#
.SYNOPSIS
Vindt bestanden met identieke inhoud, ook als ze anders heten of ergens anders staan.

.DESCRIPTION
Vergelijkt op inhoud, niet op naam. Zo vind je dezelfde taxatie die vier keer op
verschillende plekken staat, ook als iemand er "_definitief" of "_v2" achter heeft
gezet. Handig om ruimte terug te winnen, maar vooral om te weten welke versie de
echte is.

Werkwijze, en waarom die snel is: eerst worden bestanden op grootte gegroepeerd.
Alleen groepen waarin meerdere bestanden dezelfde grootte hebben worden echt
gelezen en gehasht. Dat scheelt bij een dossier met tienduizend bestanden het
grootste deel van het werk, want unieke groottes kunnen per definitie geen
duplicaat zijn.

Vul je een tweede map in, dan worden beide bomen samen doorzocht. Met
AlleenTussenDeMappen zie je uitsluitend bestanden die in beide mappen voorkomen;
dat is de vraag "wat staat er dubbel in archief en dossier".

Er wordt niets verwijderd of gewijzigd. Het script maakt alleen een rapport.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "De map die je wilt doorzoeken")]
    [string]$Map,

    [Parameter(HelpMessage = "Optionele tweede map, bijvoorbeeld het archief. Leeg laten om binnen één map te zoeken.")]
    [string]$TweedeMap = "",

    [Parameter(HelpMessage = "Pad voor het HTML-rapport. Alleen een bestandsnaam mag ook: dan komt het in Documenten.")]
    [string]$RapportHtml = "",

    [Parameter(HelpMessage = "Bestanden kleiner dan dit aantal kB overslaan. Kleine bestanden zijn vaak toevallig gelijk.")]
    [int]$MinimaleGrootteKB = 4,

    [Parameter(HelpMessage = "Alleen deze bestandstypen, bijvoorbeeld *.xlsm en *.pdf")]
    [string[]]$Filter = @("*.*"),

    [Parameter(HelpMessage = "Alleen duplicaten tonen die in beide mappen voorkomen. Vereist een tweede map.")]
    [switch]$AlleenTussenDeMappen,

    [Parameter(HelpMessage = "Het rapport na afloop meteen openen (standaard aan)")]
    [switch]$Openen
)

[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $Map -PathType Container)) {
    Write-Host "Map bestaat niet: $Map" -ForegroundColor Red
    exit 1
}
if ($TweedeMap -and -not (Test-Path -LiteralPath $TweedeMap -PathType Container)) {
    Write-Host "Tweede map bestaat niet: $TweedeMap" -ForegroundColor Red
    exit 1
}
if ($AlleenTussenDeMappen -and -not $TweedeMap) {
    Write-Host "AlleenTussenDeMappen kan alleen met een tweede map erbij." -ForegroundColor Yellow
    $AlleenTussenDeMappen = $false
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

# --------------------------------------------------------------------------
# Inlezen
# --------------------------------------------------------------------------

$drempel = [long]$MinimaleGrootteKB * 1KB

Write-Host "Map        : $Map"
if ($TweedeMap) { Write-Host "Tweede map : $TweedeMap" }
Write-Host "Filter     : $($Filter -join ', ')"
Write-Host "Vanaf      : $(Leesbaar $drempel)"
Write-Host ""

function Lees-Map {
    param([string]$Basis, [string]$Label)
    $lengte = $Basis.TrimEnd('\').Length + 1
    $uit = New-Object System.Collections.Generic.List[object]
    foreach ($b in (Get-ChildItem -LiteralPath $Basis -Include $Filter -File -Recurse `
                -ErrorAction SilentlyContinue)) {
        if ($b.Length -lt $drempel) { continue }
        if ($b.Name.StartsWith('~$')) { continue }
        $uit.Add([pscustomobject]@{
                Bron      = $Label
                Naam      = $b.Name
                Relatief  = $b.FullName.Substring($lengte)
                Vol       = $b.FullName
                Bytes     = $b.Length
                Gewijzigd = $b.LastWriteTime
                Hash      = $null
            })
    }
    return $uit
}

$labelA = Split-Path -Leaf $Map.TrimEnd('\')
$labelB = if ($TweedeMap) { Split-Path -Leaf $TweedeMap.TrimEnd('\') } else { "" }
if ($labelB -and $labelA -eq $labelB) { $labelA = "$labelA (1)"; $labelB = "$labelB (2)" }

Write-Host "Bestanden zoeken..."
$bestanden = New-Object System.Collections.Generic.List[object]
foreach ($b in (Lees-Map -Basis $Map -Label $labelA)) { $bestanden.Add($b) }
if ($TweedeMap) {
    foreach ($b in (Lees-Map -Basis $TweedeMap -Label $labelB)) { $bestanden.Add($b) }
}

Write-Host "$($bestanden.Count) bestanden boven de drempel."

if ($bestanden.Count -lt 2) {
    Write-Host "Te weinig bestanden om te vergelijken." -ForegroundColor Yellow
    exit 0
}

# --------------------------------------------------------------------------
# Eerst op grootte, dan hashen
# --------------------------------------------------------------------------

$opGrootte = $bestanden | Group-Object Bytes | Where-Object { $_.Count -gt 1 }
$teHashen = @($opGrootte | ForEach-Object { $_.Group })

Write-Host "$($teHashen.Count) bestanden hebben een grootte die ook bij een ander voorkomt."
Write-Host "$($bestanden.Count - $teHashen.Count) vallen daarmee al af zonder ze te lezen."

if ($teHashen.Count -eq 0) {
    Write-Host ""
    Write-Host "Geen enkel bestand heeft dezelfde grootte als een ander. Geen duplicaten." `
        -ForegroundColor Green
    exit 0
}

Write-Host ""
Write-Host "Inhoud lezen en vergelijken..."
$geteld = 0
$mislukt = New-Object System.Collections.Generic.List[object]

foreach ($b in $teHashen) {
    $geteld++
    if ($geteld % 100 -eq 0) {
        Write-Host "  $geteld van $($teHashen.Count)..."
    }
    try {
        $b.Hash = (Get-FileHash -LiteralPath $b.Vol -Algorithm SHA256).Hash
    }
    catch {
        $mislukt.Add([pscustomobject]@{ Pad = $b.Vol; Fout = $_.Exception.Message })
    }
}

if ($mislukt.Count -gt 0) {
    Write-Host "$($mislukt.Count) bestanden konden niet gelezen worden (vergrendeld of geen rechten)." `
        -ForegroundColor Yellow
}

# --------------------------------------------------------------------------
# Groepen samenstellen
# --------------------------------------------------------------------------

$groepen = $teHashen | Where-Object { $_.Hash } | Group-Object Hash |
    Where-Object { $_.Count -gt 1 } | ForEach-Object {
        $kopieen = @($_.Group | Sort-Object Gewijzigd)
        $bronnen = @($kopieen | Select-Object -ExpandProperty Bron -Unique)
        [pscustomobject]@{
            Hash       = $_.Name
            Aantal     = $_.Count
            Bytes      = $kopieen[0].Bytes
            Verspild   = $kopieen[0].Bytes * ($_.Count - 1)
            Kopieen    = $kopieen
            Bronnen    = $bronnen
            OverMappen = $bronnen.Count -gt 1
            Namen      = @($kopieen | Select-Object -ExpandProperty Naam -Unique)
        }
    }

if ($AlleenTussenDeMappen) {
    $groepen = @($groepen | Where-Object { $_.OverMappen })
}
$groepen = @($groepen | Sort-Object Verspild -Descending)

$totaalVerspild = ($groepen | Measure-Object -Property Verspild -Sum).Sum
if (-not $totaalVerspild) { $totaalVerspild = 0 }
$totaalKopieen = ($groepen | Measure-Object -Property Aantal -Sum).Sum
if (-not $totaalKopieen) { $totaalKopieen = 0 }

# --------------------------------------------------------------------------
# In beeld
# --------------------------------------------------------------------------

Write-Host ""
if ($groepen.Count -eq 0) {
    Write-Host "Geen duplicaten gevonden." -ForegroundColor Green
}
else {
    Write-Host "$($groepen.Count) groepen duplicaten, $totaalKopieen bestanden, $(Leesbaar $totaalVerspild) onnodig in gebruik." `
        -ForegroundColor Yellow
    Write-Host ""
    foreach ($g in ($groepen | Select-Object -First 15)) {
        Write-Host "$(Leesbaar $g.Bytes) x $($g.Aantal)  -  $(Leesbaar $g.Verspild) te winnen" `
            -ForegroundColor Cyan
        foreach ($k in $g.Kopieen) {
            $merk = if ($k -eq $g.Kopieen[0]) { "oudste" } else { "      " }
            $bron = if ($TweedeMap) { "[$($k.Bron)] " } else { "" }
            Write-Host ("  $merk  {0}{1}   {2}" -f $bron, $k.Relatief,
                $k.Gewijzigd.ToString('yyyy-MM-dd HH:mm'))
        }
        Write-Host ""
    }
    if ($groepen.Count -gt 15) {
        Write-Host "... en $($groepen.Count - 15) groepen meer, zie het rapport"
    }
}

# --------------------------------------------------------------------------
# HTML-rapport
# --------------------------------------------------------------------------

$maakRapport = [bool]$RapportHtml -or $Openen

if ($maakRapport) {

    $RapportHtml = Bepaal-Rapportpad -Opgegeven $RapportHtml -Extensie ".html" `
        -Standaardnaam (Veilige-Naam ("Duplicaten " + (Split-Path -Leaf $Map.TrimEnd('\')) +
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
.mappen { display: flex; gap: 14px; flex-wrap: wrap; margin: 0 0 24px; }
.map { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
       padding: 12px 16px; min-width: 260px; }
.map .naam { font-weight: 600; }
.map .pad { color: #6E6E73; font-size: 12px; word-break: break-all;
            font-family: Consolas, monospace; }
.kaarten { display: flex; gap: 14px; flex-wrap: wrap; margin: 0 0 8px; }
.kaart { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
         padding: 14px 18px; min-width: 160px; border-left: 4px solid #2C3E50; }
.kaart .getal { font-size: 28px; font-weight: 600; line-height: 1.1; }
.kaart .label { color: #6E6E73; font-size: 12px; margin-top: 2px; }
.kaart.dup { border-left-color: #B85CC7; }
.kaart.ruimte { border-left-color: #E8A33D; }
.kaart.ok { border-left-color: #34A853; }
.groep { background: #fff; border: 1px solid #D1D1D6; border-radius: 10px;
         padding: 12px 16px; margin: 0 0 12px; }
.groepkop { display: flex; align-items: baseline; gap: 12px; flex-wrap: wrap;
            margin: 0 0 8px; }
.groepkop .titel { font-weight: 600; font-family: Consolas, monospace; }
.groepkop .winst { color: #8A5A12; font-weight: 600; }
table.kopieen { border-collapse: collapse; width: 100%; }
table.kopieen td { padding: 5px 8px; border-top: 1px solid #F0F0F3;
                   font-family: Consolas, monospace; font-size: 12.5px;
                   vertical-align: top; }
table.kopieen tr:first-child td { border-top: none; }
td.datum, td.bron { white-space: nowrap; color: #6E6E73; font-family: "Segoe UI", sans-serif;
                    font-size: 12px; }
.badge { display: inline-block; padding: 1px 8px; border-radius: 10px;
         font-size: 12px; font-weight: 600; white-space: nowrap; }
.b-oud { background: #E4F4E7; color: #1E6B2E; }
.b-dup { background: #F6E6F8; color: #6E2A7A; }
.b-over { background: #E3EFFB; color: #1B4E80; }
.b-naam { background: #FDF0DC; color: #8A5A12; }
.leeg { color: #8E8E93; font-style: italic; padding: 14px 12px; background: #fff;
        border: 1px solid #D1D1D6; border-radius: 10px; }
table.plat { border-collapse: collapse; width: 100%; background: #fff;
             border: 1px solid #D1D1D6; border-radius: 10px; overflow: hidden; }
table.plat th { background: #2C3E50; color: #fff; text-align: left; padding: 9px 12px;
                font-weight: 600; font-size: 13px; }
table.plat td { padding: 8px 12px; border-top: 1px solid #EDEDF0; }
table.plat tr:nth-child(even) td { background: #FAFAFB; }
.num { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }
footer { margin-top: 40px; color: #8E8E93; font-size: 12px; }
@media print {
  body { background: #fff; padding: 0; }
  h2 { page-break-after: avoid; }
  .groep { page-break-inside: avoid; }
}
'@

    $html = New-Object System.Text.StringBuilder
    $null = $html.AppendLine('<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8">')
    $null = $html.AppendLine('<title>Dubbele bestanden</title>')
    $null = $html.AppendLine("<style>$css</style></head><body>")

    $null = $html.AppendLine("<h1>Dubbele bestanden</h1>")
    $null = $html.AppendLine("<p class='sub'>Gemaakt op $(Get-Date -Format 'dddd d MMMM yyyy, HH:mm')" +
        " &middot; vergeleken op inhoud, vanaf $(Leesbaar $drempel)</p>")

    $null = $html.AppendLine("<div class='mappen'>")
    $null = $html.AppendLine("<div class='map'><div class='naam'>$(HtmlVeilig $labelA)</div>" +
        "<div class='pad'>$(HtmlVeilig $Map)</div></div>")
    if ($TweedeMap) {
        $null = $html.AppendLine("<div class='map'><div class='naam'>$(HtmlVeilig $labelB)</div>" +
            "<div class='pad'>$(HtmlVeilig $TweedeMap)</div></div>")
    }
    $null = $html.AppendLine("</div>")

    $null = $html.AppendLine("<div class='kaarten'>")
    $null = $html.AppendLine("<div class='kaart dup'><div class='getal'>$($groepen.Count)</div><div class='label'>groepen duplicaten</div></div>")
    $null = $html.AppendLine("<div class='kaart dup'><div class='getal'>$totaalKopieen</div><div class='label'>bestanden in die groepen</div></div>")
    $null = $html.AppendLine("<div class='kaart ruimte'><div class='getal'>$(Leesbaar $totaalVerspild)</div><div class='label'>onnodig in gebruik</div></div>")
    $null = $html.AppendLine("<div class='kaart ok'><div class='getal'>$($bestanden.Count)</div><div class='label'>bestanden nagelopen</div></div>")
    $null = $html.AppendLine("</div>")

    $null = $html.AppendLine("<p class='sub'>Van de $($bestanden.Count) bestanden vielen " +
        "$($bestanden.Count - $teHashen.Count) af op grootte alleen; " +
        "$($teHashen.Count) zijn echt gelezen en vergeleken.</p>")

    if ($groepen.Count -eq 0) {
        $null = $html.AppendLine("<div class='leeg'>Geen duplicaten gevonden.</div>")
    }
    else {
        # ---- overzicht ----
        $null = $html.AppendLine("<h2>Overzicht, meeste winst eerst</h2>")
        $null = $html.AppendLine("<table class='plat'><thead><tr><th>Bestand</th>" +
            "<th class='num'>Kopieën</th><th class='num'>Per stuk</th>" +
            "<th class='num'>Te winnen</th></tr></thead><tbody>")
        foreach ($g in $groepen) {
            $naam = $g.Namen[0]
            $extra = if ($g.Namen.Count -gt 1) { " <span class='badge b-naam'>$($g.Namen.Count) verschillende namen</span>" } else { "" }
            $over = if ($g.OverMappen -and $TweedeMap) { " <span class='badge b-over'>in beide mappen</span>" } else { "" }
            $null = $html.AppendLine("<tr><td>$(HtmlVeilig $naam)$extra$over</td>" +
                "<td class='num'>$($g.Aantal)</td><td class='num'>$(Leesbaar $g.Bytes)</td>" +
                "<td class='num'>$(Leesbaar $g.Verspild)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")

        # ---- per groep ----
        $null = $html.AppendLine("<h2>Alle groepen</h2>")
        $null = $html.AppendLine("<p class='sub'>Het oudste bestand is gemarkeerd; dat is " +
            "meestal het origineel waar de kopieën van afstammen. Er wordt niets verwijderd, " +
            "dat blijft aan jou.</p>")

        foreach ($g in $groepen) {
            $null = $html.AppendLine("<div class='groep'><div class='groepkop'>")
            $null = $html.AppendLine("<span class='titel'>$(HtmlVeilig $g.Namen[0])</span>")
            $null = $html.AppendLine("<span class='badge b-dup'>$($g.Aantal) kopieën</span>")
            $null = $html.AppendLine("<span>$(Leesbaar $g.Bytes) per stuk</span>")
            $null = $html.AppendLine("<span class='winst'>$(Leesbaar $g.Verspild) te winnen</span>")
            if ($g.OverMappen -and $TweedeMap) {
                $null = $html.AppendLine("<span class='badge b-over'>in beide mappen</span>")
            }
            $null = $html.AppendLine("</div><table class='kopieen'>")
            foreach ($k in $g.Kopieen) {
                $merk = if ($k -eq $g.Kopieen[0]) { "<span class='badge b-oud'>oudste</span>" } else { "" }
                $bron = if ($TweedeMap) { "<td class='bron'>$(HtmlVeilig $k.Bron)</td>" } else { "" }
                $null = $html.AppendLine("<tr>$bron<td>$(HtmlVeilig $k.Relatief)</td>" +
                    "<td class='datum'>$($k.Gewijzigd.ToString('yyyy-MM-dd HH:mm'))</td>" +
                    "<td class='datum'>$merk</td></tr>")
            }
            $null = $html.AppendLine("</table></div>")
        }
    }

    if ($mislukt.Count -gt 0) {
        $null = $html.AppendLine("<h2>Niet kunnen lezen <span class='badge b-naam'>$($mislukt.Count)</span></h2>")
        $null = $html.AppendLine("<table class='plat'><thead><tr><th>Bestand</th><th>Reden</th></tr></thead><tbody>")
        foreach ($m in $mislukt) {
            $null = $html.AppendLine("<tr><td>$(HtmlVeilig $m.Pad)</td><td>$(HtmlVeilig $m.Fout)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    $null = $html.AppendLine("<footer>Vergeleken op de SHA256-som van de inhoud. Bestanden met " +
        "dezelfde som zijn inhoudelijk identiek, ongeacht hun naam of datum.<br>" +
        "Bestanden kleiner dan $(Leesbaar $drempel) zijn overgeslagen, en tijdelijke " +
        "Excel-bestanden die met ~`$ beginnen ook.<br>Er is niets verwijderd of gewijzigd." +
        "</footer>")
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
