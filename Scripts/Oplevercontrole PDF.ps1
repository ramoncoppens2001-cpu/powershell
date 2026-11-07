<#
.SYNOPSIS
Controleert of er bij elke Excel-werkmap een actuele PDF staat.

.DESCRIPTION
Het sluitstuk na een run met Excel-naar-PDF. Per werkmap wordt gekeken of er een
PDF met dezelfde naam bestaat, of die nieuwer is dan de werkmap, en of hij niet
verdacht klein of verdacht kort is.

Vier dingen die misgaan en die je hier ziet:

  PDF ontbreekt        er is nooit een PDF gemaakt, of de run is halverwege
                       gestopt
  PDF is ouder         de werkmap is na het printen nog gewijzigd, dus je
                       levert een PDF van vóór je laatste aanpassing. Dit is
                       de gevaarlijkste van de vier, want er lijkt niets
                       aan de hand
  verdacht klein       een PDF van een paar kB is meestal een lege pagina
  te weinig pagina's   het aantal pagina's wordt geteld; één pagina bij een
                       model met tien tabbladen klopt zelden

Ook PDF's zonder bijbehorende werkmap worden gemeld: dan is de werkmap verplaatst
of hernoemd, en verwijst je oplevering naar iets wat niet meer bestaat.

Er wordt niets gewijzigd, alleen gelezen.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory, HelpMessage = "Map met de Excel-werkmappen")]
    [string]$Modelmap,

    [Parameter(HelpMessage = "Map met de PDF's. Leeg laten als de PDF's naast de werkmappen staan.")]
    [string]$Pdfmap = "",

    [Parameter(HelpMessage = "Pad voor het HTML-rapport. Alleen een bestandsnaam mag ook: dan komt het in Documenten.")]
    [string]$RapportHtml = "",

    [Parameter(HelpMessage = "Welke werkmappen meedoen")]
    [string[]]$Filter = @("*.xlsm", "*.xlsx"),

    [Parameter(HelpMessage = "Een PDF kleiner dan dit aantal kB is verdacht")]
    [int]$MinimalePdfKB = 20,

    [Parameter(HelpMessage = "Een PDF met minder pagina's dan dit is verdacht. 0 om niet te controleren.")]
    [int]$MinimaalAantalPaginas = 1,

    [Parameter(HelpMessage = "Hoeveel minuten een PDF ouder mag zijn dan de werkmap voordat het opvalt")]
    [int]$MargeMinuten = 5,

    [Parameter(HelpMessage = "Aanvinken om submappen NIET mee te nemen")]
    [switch]$AlleenBovensteMap,

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

if (-not (Test-Path -LiteralPath $Modelmap -PathType Container)) {
    Write-Host "Modelmap bestaat niet: $Modelmap" -ForegroundColor Red
    exit 1
}
if ($Pdfmap -and -not (Test-Path -LiteralPath $Pdfmap -PathType Container)) {
    Write-Host "PDF-map bestaat niet: $Pdfmap" -ForegroundColor Red
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
    }
}

function Pdf-Paginas {
    <#
    Ruwe telling van de pagina-objecten in een PDF, zonder bibliotheek.

    Goed genoeg om te zien of Excel de helft heeft laten liggen. Grote
    bestanden worden overgeslagen om niet onnodig geheugen te vullen.
    #>
    param([string]$Pad, [long]$Bytes)

    if ($Bytes -gt 80MB) { return $null }
    try {
        $rauw = [System.IO.File]::ReadAllBytes($Pad)
        # 28591 is Latin-1: elke byte wordt één teken, dus zoeken blijft betrouwbaar
        $tekst = [System.Text.Encoding]::GetEncoding(28591).GetString($rauw)
        $paginas = ([regex]::Matches($tekst, '/Type\s*/Page[^s]')).Count
        if ($paginas -eq 0) {
            $paginas = ([regex]::Matches($tekst, '/Count\s+(\d+)')).Count
            if ($paginas -gt 0) {
                $eerste = [regex]::Match($tekst, '/Count\s+(\d+)')
                return [int]$eerste.Groups[1].Value
            }
            return $null
        }
        return $paginas
    }
    catch {
        return $null
    }
}

function Dossier {
    param([string]$RelatiefPad)
    $delen = $RelatiefPad -split '\\'
    if ($delen.Count -le 1) { return "(hoofdmap)" }
    return $delen[0]
}

# --------------------------------------------------------------------------
# Inlezen
# --------------------------------------------------------------------------

$OK = "in orde"
$ONTBREEKT = "PDF ontbreekt"
$OUDER = "PDF is ouder dan het model"
$KLEIN = "PDF verdacht klein"
$KORT = "PDF heeft te weinig pagina's"

$naastElkaar = -not $Pdfmap
$pdfBasis = if ($naastElkaar) { $Modelmap } else { $Pdfmap }

Write-Host "Modellen : $Modelmap"
Write-Host "PDF's    : $(if ($naastElkaar) { 'naast de werkmappen' } else { $Pdfmap })"
Write-Host "Filter   : $($Filter -join ', ')"
Write-Host ""

$modelBasislengte = $Modelmap.TrimEnd('\').Length + 1
$pdfBasislengte = $pdfBasis.TrimEnd('\').Length + 1

if ($AlleenBovensteMap) {
    $modellen = @(Get-ChildItem -LiteralPath $Modelmap -File | Where-Object {
            $naam = $_.Name
            -not $naam.StartsWith('~$') -and ($Filter | Where-Object { $naam -like $_ }).Count -gt 0
        })
}
else {
    $modellen = @(Get-ChildItem -LiteralPath $Modelmap -Include $Filter -File -Recurse `
            -ErrorAction SilentlyContinue | Tel-Onderweg -Wat "Werkmappen zoeken" |
            Where-Object { -not $_.Name.StartsWith('~$') })
}
Stop-Voortgang -Wat "Werkmappen zoeken"

if ($modellen.Count -eq 0) {
    Write-Host "Geen werkmappen gevonden die aan het filter voldoen." -ForegroundColor Yellow
    exit 0
}

Write-Host "$($modellen.Count) werkmappen gevonden. Controleren..."

$drempel = [long]$MinimalePdfKB * 1KB
$resultaten = New-Object System.Collections.Generic.List[object]
$gebruiktePdfs = New-Object System.Collections.Generic.HashSet[string]
$geteld = 0

foreach ($model in $modellen) {
    $geteld++
    Meld-Voortgang -Wat "Werkmappen controleren" -Nu $geteld -Totaal $modellen.Count -Detail $model.Name
    if ($geteld % 50 -eq 0) { Write-Host "  $geteld van $($modellen.Count)..." }

    $relatief = $model.FullName.Substring($modelBasislengte)
    $verwachtRelatief = [System.IO.Path]::ChangeExtension($relatief, ".pdf")
    $pdfPad = Join-Path $pdfBasis $verwachtRelatief

    $rij = [ordered]@{
        Model        = $relatief
        Modelnaam    = $model.Name
        Dossier      = Dossier $relatief
        ModelBytes   = $model.Length
        ModelDatum   = $model.LastWriteTime
        PdfRelatief  = $verwachtRelatief
        PdfBytes     = $null
        PdfDatum     = $null
        Paginas      = $null
        Status       = $OK
        Toelichting  = ""
    }

    if (-not (Test-Path -LiteralPath $pdfPad -PathType Leaf)) {
        $rij.Status = $ONTBREEKT
        $rij.Toelichting = "verwacht op $verwachtRelatief"
        $resultaten.Add([pscustomobject]$rij)
        continue
    }

    $pdf = Get-Item -LiteralPath $pdfPad
    $null = $gebruiktePdfs.Add($pdf.FullName.ToLower())
    $rij.PdfBytes = $pdf.Length
    $rij.PdfDatum = $pdf.LastWriteTime
    $rij.Paginas = Pdf-Paginas -Pad $pdf.FullName -Bytes $pdf.Length

    $bezwaren = @()

    $verschil = ($pdf.LastWriteTime - $model.LastWriteTime).TotalMinutes
    if ($verschil -lt (-1 * $MargeMinuten)) {
        $ouder = [math]::Round([math]::Abs($verschil))
        $tekst = if ($ouder -ge 1440) { "$([math]::Round($ouder / 1440, 1)) dagen" }
        elseif ($ouder -ge 60) { "$([math]::Round($ouder / 60, 1)) uur" }
        else { "$ouder minuten" }
        $bezwaren += "PDF is $tekst ouder dan de werkmap"
        $rij.Status = $OUDER
    }

    if ($pdf.Length -lt $drempel) {
        $bezwaren += "PDF is maar $(Leesbaar $pdf.Length)"
        if ($rij.Status -eq $OK) { $rij.Status = $KLEIN }
    }

    if ($MinimaalAantalPaginas -gt 0 -and $null -ne $rij.Paginas -and
        $rij.Paginas -lt $MinimaalAantalPaginas) {
        $bezwaren += "PDF heeft $($rij.Paginas) pagina's"
        if ($rij.Status -eq $OK) { $rij.Status = $KORT }
    }

    $rij.Toelichting = $bezwaren -join "; "
    $resultaten.Add([pscustomobject]$rij)
}
Stop-Voortgang -Wat "Werkmappen controleren"

# PDF's zonder werkmap
$alleWeesPdfs = New-Object System.Collections.Generic.List[object]
if ($AlleenBovensteMap) {
    $pdfLijst = @(Get-ChildItem -LiteralPath $pdfBasis -Filter "*.pdf" -File)
}
else {
    $pdfLijst = @(Get-ChildItem -LiteralPath $pdfBasis -Filter "*.pdf" -File -Recurse `
            -ErrorAction SilentlyContinue)
}
foreach ($pdf in $pdfLijst) {
    if ($gebruiktePdfs.Contains($pdf.FullName.ToLower())) { continue }
    $alleWeesPdfs.Add([pscustomobject]@{
            Relatief = $pdf.FullName.Substring($pdfBasislengte)
            Bytes    = $pdf.Length
            Datum    = $pdf.LastWriteTime
        })
}

# --------------------------------------------------------------------------
# In beeld
# --------------------------------------------------------------------------

$telling = [ordered]@{}
foreach ($s in @($OK, $ONTBREEKT, $OUDER, $KLEIN, $KORT)) {
    $telling[$s] = @($resultaten | Where-Object { $_.Status -eq $s }).Count
}

Write-Host ""
Write-Host "$($telling[$OK]) van $($resultaten.Count) in orde." -ForegroundColor Green
foreach ($s in @($ONTBREEKT, $OUDER, $KLEIN, $KORT)) {
    if ($telling[$s] -gt 0) { Write-Host "$($telling[$s])x $s" -ForegroundColor Yellow }
}
if ($alleWeesPdfs.Count -gt 0) {
    Write-Host "$($alleWeesPdfs.Count) PDF's zonder bijbehorende werkmap" -ForegroundColor Cyan
}

foreach ($s in @($OUDER, $ONTBREEKT, $KLEIN, $KORT)) {
    $groep = @($resultaten | Where-Object { $_.Status -eq $s })
    if ($groep.Count -eq 0) { continue }
    Write-Host ""
    Write-Host $s.ToUpper() -ForegroundColor Yellow
    foreach ($r in ($groep | Select-Object -First 25)) {
        Write-Host "  $($r.Model)"
        if ($r.Toelichting) { Write-Host "      $($r.Toelichting)" }
    }
    if ($groep.Count -gt 25) { Write-Host "  ... en $($groep.Count - 25) meer, zie het rapport" }
}

# --------------------------------------------------------------------------
# HTML-rapport
# --------------------------------------------------------------------------

$maakRapport = [bool]$RapportHtml -or $Openen

if ($maakRapport) {

    $RapportHtml = Bepaal-Rapportpad -Opgegeven $RapportHtml -Extensie ".html" `
        -Standaardnaam (Veilige-Naam ("Oplevercontrole " +
            (Split-Path -Leaf $Modelmap.TrimEnd('\')) + " " +
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
         padding: 14px 18px; min-width: 150px; border-left: 4px solid #2C3E50; }
.kaart .getal { font-size: 28px; font-weight: 600; line-height: 1.1; }
.kaart .label { color: #6E6E73; font-size: 12px; margin-top: 2px; }
.kaart.ok { border-left-color: #34A853; }
.kaart.ontbreekt { border-left-color: #C0392B; }
.kaart.ouder { border-left-color: #E8A33D; }
.kaart.klein { border-left-color: #B85CC7; }
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
tr.mis td { background: #FDECEA; }
.badge { display: inline-block; padding: 1px 8px; border-radius: 10px;
         font-size: 12px; font-weight: 600; white-space: nowrap; }
.b-ok { background: #E4F4E7; color: #1E6B2E; }
.b-ontbreekt { background: #FDECEA; color: #8E2420; }
.b-ouder { background: #FDF0DC; color: #8A5A12; }
.b-klein { background: #F6E6F8; color: #6E2A7A; }
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
.s-ok { background: #CFD2D7; }
.s-ontbreekt { background: #C0392B; }
.s-ouder { background: #E8A33D; }
.s-klein { background: #B85CC7; }
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
        $telling = @{ ok = 0; ontbreekt = 0; ouder = 0; klein = 0 }

        foreach ($submap in $Knoop.mappen.Keys) {
            $kind = Teken-Knoop -Knoop $Knoop.mappen[$submap] -Naam $submap `
                -Diepte ($Diepte + 1) -OpenTot $OpenTot
            $null = $binnen.Append($kind.html)
            foreach ($s in @('ok', 'ontbreekt', 'ouder', 'klein')) {
                $telling[$s] += $kind.telling[$s]
            }
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

        $problemen = $telling.ontbreekt + $telling.ouder + $telling.klein
        $totaal = $problemen + $telling.ok
        $merken = New-Object System.Text.StringBuilder
        $null = $merken.Append("<span class='tel'>$totaal model$(if ($totaal -ne 1) { 'len' })</span>")
        if ($telling.ontbreekt -gt 0) {
            $null = $merken.Append("<span class='badge b-ontbreekt'>$($telling.ontbreekt) zonder PDF</span>")
        }
        if ($telling.ouder -gt 0) {
            $null = $merken.Append("<span class='badge b-ouder'>$($telling.ouder) verouderd</span>")
        }
        if ($telling.klein -gt 0) {
            $null = $merken.Append("<span class='badge b-klein'>$($telling.klein) verdacht</span>")
        }
        if ($problemen -eq 0 -and $totaal -gt 0) {
            $null = $merken.Append("<span class='badge b-ok'>in orde</span>")
        }

        $open = if ($Diepte -lt $OpenTot -or $problemen -gt 0) { " open" } else { "" }
        $html = New-Object System.Text.StringBuilder
        $null = $html.AppendLine("<details class='boom'$open><summary>" +
            "<span class='mapnaam'>$(HtmlVeilig $Naam)</span>$($merken.ToString())</summary>")
        $null = $html.AppendLine("<div class='kinderen'>$($binnen.ToString())</div></details>")
        return @{ html = $html.ToString(); telling = $telling }
    }

    $klasseVan = @{
        $OK = "ok"; $ONTBREEKT = "ontbreekt"; $OUDER = "ouder"
        $KLEIN = "klein"; $KORT = "klein"
    }

    $html = New-Object System.Text.StringBuilder
    $null = $html.AppendLine('<!DOCTYPE html><html lang="nl"><head><meta charset="utf-8">')
    $null = $html.AppendLine('<title>Oplevercontrole</title>')
    $null = $html.AppendLine("<style>$css</style></head><body>")

    $null = $html.AppendLine("<h1>Oplevercontrole: $(HtmlVeilig (Split-Path -Leaf $Modelmap.TrimEnd('\')))</h1>")
    $null = $html.AppendLine("<p class='sub'>Gemaakt op $(Get-Date -Format 'dddd d MMMM yyyy, HH:mm')</p>")

    $null = $html.AppendLine("<div class='mappen'>")
    $null = $html.AppendLine("<div class='map'><div class='naam'>Werkmappen &mdash; $($resultaten.Count)</div>" +
        "<div class='pad'>$(HtmlVeilig $Modelmap)</div></div>")
    $null = $html.AppendLine("<div class='map'><div class='naam'>PDF's</div>" +
        "<div class='pad'>$(HtmlVeilig $(if ($naastElkaar) { 'naast de werkmappen' } else { $Pdfmap }))</div></div>")
    $null = $html.AppendLine("</div>")

    $null = $html.AppendLine("<div class='kaarten'>")
    $null = $html.AppendLine("<div class='kaart ok'><div class='getal'>$($telling[$OK])</div><div class='label'>in orde</div></div>")
    $null = $html.AppendLine("<div class='kaart ontbreekt'><div class='getal'>$($telling[$ONTBREEKT])</div><div class='label'>PDF ontbreekt</div></div>")
    $null = $html.AppendLine("<div class='kaart ouder'><div class='getal'>$($telling[$OUDER])</div><div class='label'>PDF is ouder</div></div>")
    $null = $html.AppendLine("<div class='kaart klein'><div class='getal'>$($telling[$KLEIN] + $telling[$KORT])</div><div class='label'>verdacht klein of kort</div></div>")
    $null = $html.AppendLine("</div>")

    if ($telling[$OUDER] -gt 0) {
        $null = $html.AppendLine("<p class='sub'><span class='badge b-ouder'>let op</span> " +
            "Een PDF die ouder is dan de werkmap betekent dat het model daarna nog " +
            "gewijzigd is. Die PDF ziet er goed uit maar mist je laatste aanpassing.</p>")
    }

    # ---- per dossiermap ----
    $perDossier = $resultaten | Group-Object Dossier | ForEach-Object {
        [pscustomobject]@{
            Map       = $_.Name
            Totaal    = $_.Count
            Ok        = @($_.Group | Where-Object { $_.Status -eq $OK }).Count
            Ontbreekt = @($_.Group | Where-Object { $_.Status -eq $ONTBREEKT }).Count
            Ouder     = @($_.Group | Where-Object { $_.Status -eq $OUDER }).Count
            Verdacht  = @($_.Group | Where-Object { $_.Status -in @($KLEIN, $KORT) }).Count
        }
    } | Sort-Object -Property @{Expression = { $_.Totaal - $_.Ok }; Descending = $true }, "Map"

    $null = $html.AppendLine("<h2>Overzicht per map</h2>")
    $null = $html.AppendLine("<table><thead><tr><th>Map</th><th class='num'>Modellen</th>" +
        "<th class='num'>In orde</th><th class='num'>Geen PDF</th><th class='num'>Ouder</th>" +
        "<th class='num'>Verdacht</th><th>Status</th></tr></thead><tbody>")
    foreach ($d in $perDossier) {
        $problemen = $d.Totaal - $d.Ok
        $klasse = if ($d.Ontbreekt -gt 0) { " class='mis'" } elseif ($problemen -gt 0) { " class='let'" } else { "" }
        $status = if ($problemen -eq 0) { "<span class='badge b-ok'>in orde</span>" }
        else { "<span class='badge b-ouder'>$problemen aandachtspunt$(if ($problemen -ne 1) { 'en' })</span>" }
        $null = $html.AppendLine("<tr$klasse><td class='pad'>$(HtmlVeilig $d.Map)</td>" +
            "<td class='num'>$($d.Totaal)</td><td class='num'>$($d.Ok)</td>" +
            "<td class='num'>$($d.Ontbreekt)</td><td class='num'>$($d.Ouder)</td>" +
            "<td class='num'>$($d.Verdacht)</td><td>$status</td></tr>")
    }
    $null = $html.AppendLine("</tbody></table>")

    # ---- structuur ----
    $wortel = Nieuwe-Knoop
    foreach ($r in $resultaten) {
        $uitleg = switch ($r.Status) {
            $OK { "$(Leesbaar $r.PdfBytes)$(if ($r.Paginas) { ", $($r.Paginas) pagina's" })" }
            $ONTBREEKT { "geen PDF" }
            default { $r.Toelichting }
        }
        Voeg-InBoom -Knoop $wortel -Delen ($r.Model -split '\\') -Bestand @{
            naam = $r.Modelnaam; klasse = $klasseVan[$r.Status]; uitleg = $uitleg
        }
    }
    $null = $html.AppendLine("<h2>Structuur</h2>")
    $null = $html.AppendLine("<div class='legenda'>" +
        "<span class='item'><span class='stip s-ok'></span>in orde</span>" +
        "<span class='item'><span class='stip s-ontbreekt'></span>PDF ontbreekt</span>" +
        "<span class='item'><span class='stip s-ouder'></span>PDF is ouder</span>" +
        "<span class='item'><span class='stip s-klein'></span>verdacht klein of kort</span></div>")
    $null = $html.AppendLine("<div class='knoppen'>" +
        "<button onclick=""alles(true)"">Alles uitklappen</button>" +
        "<button onclick=""alles(false)"">Alles inklappen</button></div>")
    $null = $html.AppendLine("<div class='boomvak'>")
    $getekend = Teken-Knoop -Knoop $wortel -Naam (Split-Path -Leaf $Modelmap.TrimEnd('\')) -Diepte 0
    $null = $html.AppendLine($getekend.html)
    $null = $html.AppendLine("</div>")

    # ---- per categorie ----
    foreach ($blok in @(
            @{ titel = $OUDER; badge = "b-ouder" },
            @{ titel = $ONTBREEKT; badge = "b-ontbreekt" },
            @{ titel = $KLEIN; badge = "b-klein" },
            @{ titel = $KORT; badge = "b-klein" }
        )) {
        $groep = @($resultaten | Where-Object { $_.Status -eq $blok.titel })
        if ($groep.Count -eq 0) { continue }
        $null = $html.AppendLine("<h2>$(HtmlVeilig $blok.titel) " +
            "<span class='badge $($blok.badge)'>$($groep.Count)</span></h2>")
        $null = $html.AppendLine("<table><thead><tr><th>Werkmap</th>" +
            "<th class='num'>Model gewijzigd</th><th class='num'>PDF gewijzigd</th>" +
            "<th class='num'>PDF</th><th class='num'>Pagina's</th>" +
            "<th>Wat er aan de hand is</th></tr></thead><tbody>")
        foreach ($r in ($groep | Sort-Object Model)) {
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $r.Model)</td>" +
                "<td class='num'>$($r.ModelDatum.ToString('yyyy-MM-dd HH:mm'))</td>" +
                "<td class='num'>$(if ($r.PdfDatum) { $r.PdfDatum.ToString('yyyy-MM-dd HH:mm') } else { '' })</td>" +
                "<td class='num'>$(Leesbaar $r.PdfBytes)</td>" +
                "<td class='num'>$(if ($null -ne $r.Paginas) { $r.Paginas } else { '?' })</td>" +
                "<td>$(HtmlVeilig $r.Toelichting)</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    # ---- PDF's zonder werkmap ----
    $null = $html.AppendLine("<h2>PDF's zonder bijbehorende werkmap " +
        "<span class='badge b-klein'>$($alleWeesPdfs.Count)</span></h2>")
    if ($alleWeesPdfs.Count -eq 0) {
        $null = $html.AppendLine("<div class='leeg'>Bij elke PDF hoort een werkmap.</div>")
    }
    else {
        $null = $html.AppendLine("<p class='sub'>Deze PDF's horen bij geen enkele werkmap in " +
            "de modelmap. Meestal is de werkmap hernoemd of verplaatst.</p>")
        $null = $html.AppendLine("<table><thead><tr><th>PDF</th><th class='num'>Grootte</th>" +
            "<th class='num'>Gewijzigd</th></tr></thead><tbody>")
        foreach ($p in ($alleWeesPdfs | Sort-Object Relatief)) {
            $null = $html.AppendLine("<tr><td class='pad'>$(HtmlVeilig $p.Relatief)</td>" +
                "<td class='num'>$(Leesbaar $p.Bytes)</td>" +
                "<td class='num'>$($p.Datum.ToString('yyyy-MM-dd HH:mm'))</td></tr>")
        }
        $null = $html.AppendLine("</tbody></table>")
    }

    $null = $html.AppendLine("<footer>Een PDF hoort nieuwer te zijn dan de werkmap waaruit " +
        "hij komt. Er is $MargeMinuten minuten marge aangehouden.<br>" +
        "PDF's onder $(Leesbaar $drempel) of met minder dan $MinimaalAantalPaginas pagina's " +
        "worden als verdacht gemeld. Het aantal pagina's is geteld in het PDF-bestand zelf.<br>" +
        "$(HtmlVeilig $Modelmap)</footer>")
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
