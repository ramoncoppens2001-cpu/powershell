<#
.SYNOPSIS
Zet bestanden onder een map op alleen-lezen, of heft dat weer op.

.DESCRIPTION
Handig om een archiefmap vast te zetten zodat niemand er per ongeluk in werkt.
Anders dan de oude versie kun je dit ook terugdraaien, en kies je zelf welke
bestandstypen meedoen.

Er wordt niets aan de inhoud van bestanden gewijzigd, alleen het kenmerk
alleen-lezen.

.EXAMPLE
Alleen alle Excel-werkmappen met macro's vastzetten:
  Map = T:\Archief\2025\Klant, Filter = *.xlsm

.EXAMPLE
Alles weer vrijgeven:
  Map = T:\Archief\2025\Klant, Filter = *.*, Terugdraaien aan
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, HelpMessage = "De map waarin gezocht wordt")]
    [string]$Map,

    [Parameter(HelpMessage = "Welke bestanden, bijvoorbeeld *.xlsm of *.xlsx en *.xlsm")]
    [string[]]$Filter = @("*.xlsm"),

    [Parameter(HelpMessage = "Aanvinken om alleen-lezen juist op te heffen")]
    [switch]$Terugdraaien,

    [Parameter(HelpMessage = "Aanvinken om submappen NIET mee te nemen")]
    [switch]$AlleenBovensteMap
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

$doel = -not $Terugdraaien
$wat  = if ($doel) { "alleen-lezen" } else { "schrijfbaar" }

Write-Host "Map      : $Map"
Write-Host "Filter   : $($Filter -join ', ')"
Write-Host "Actie    : bestanden $wat maken"
Write-Host "Submappen: $(if ($AlleenBovensteMap) { 'nee' } else { 'ja' })"
Write-Host ""

$bestanden = Get-ChildItem -LiteralPath $Map -Include $Filter -File `
    -Recurse:(-not $AlleenBovensteMap) -ErrorAction SilentlyContinue |
    Tel-Onderweg -Wat "Bestanden zoeken"
Stop-Voortgang -Wat "Bestanden zoeken"

# Get-ChildItem -Include werkt alleen samen met -Recurse of een pad met wildcard.
# Zonder submappen vallen we daarom terug op filteren per bestand.
if ($AlleenBovensteMap) {
    $bestanden = Get-ChildItem -LiteralPath $Map -File | Where-Object {
        $naam = $_.Name
        ($Filter | Where-Object { $naam -like $_ }).Count -gt 0
    }
}

if (-not $bestanden) {
    Write-Host "Geen bestanden gevonden die aan het filter voldoen." -ForegroundColor Yellow
    exit 0
}

$bestanden = @($bestanden)
$aangepast = 0
$overgeslagen = 0
$mislukt = 0
$geteld = 0

foreach ($bestand in $bestanden) {
    $geteld++
    Meld-Voortgang -Wat "Bestanden $wat maken" -Nu $geteld -Totaal $bestanden.Count -Detail $bestand.Name
    if ($bestand.IsReadOnly -eq $doel) {
        $overgeslagen++
        continue
    }
    if ($PSCmdlet.ShouldProcess($bestand.FullName, "$wat maken")) {
        try {
            Set-ItemProperty -LiteralPath $bestand.FullName -Name IsReadOnly -Value $doel
            $aangepast++
            Write-Host "  $wat : $($bestand.FullName)"
        }
        catch {
            $mislukt++
            Write-Host "  MISLUKT: $($bestand.FullName)" -ForegroundColor Red
            Write-Host "           $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}
Stop-Voortgang -Wat "Bestanden $wat maken"

Write-Host ""
Write-Host "Klaar. $aangepast aangepast, $overgeslagen stonden al goed, $mislukt mislukt." `
    -ForegroundColor Green
if ($mislukt -gt 0) { exit 1 }
