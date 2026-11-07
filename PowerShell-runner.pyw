"""
PowerShell-runner
=================

Draai je PowerShell-scripts vanuit een venster, zonder in code te typen.

De runner leest het `param()`-blok van elk script en maakt daar automatisch
invulvelden van: mapkiezers voor mappen, aan/uit-knopjes voor schakelaars,
keuzeknoppen waar het script een vaste keuze verwacht. De uitvoer loopt live
mee en wordt opgeslagen in een logbestand.

Dubbelklik dit bestand. De .pyw-extensie zorgt dat Windows het zonder zwart
terminalvenster start. Er is niets te installeren; dit werkt op standaard
Windows met PowerShell.

Het venster (versie 2)
----------------------
- Zelfde stijl als Excel naar PDF en de Adressen-tool: lichtgrijze
  achtergrond, witte kaarten, knoppen als pil, aan/uit-knopjes in plaats van
  vierkante vinkjes, een dunne blauwe voortgangslijn en scherpe tekst bij
  schermvergroting.
- Favorieten: klik op het sterretje voor een script en het staat bovenaan.
- Wat je de vorige keer per script invulde, staat er de volgende keer weer.
  Met "Standaardwaarden" zet je alles terug zoals het script het aangeeft.
- Mapvelden onthouden de mappen die je eerder gebruikte (uitklappen), en je
  kunt een map uit de Verkenner op het venster slepen.
- Filtervelden hebben knopjes voor de gangbare bestandstypen (*.xlsm, *.pdf…).
- Een vaste keuze met weinig opties (Robocopy / PerBestand) is een rij
  knopjes in plaats van een uitklaplijst.
- De voortgangslijn laat zien hoe ver het is. Gebruikt het script
  Write-Progress met een percentage, dan is dat exact. Anders schat de runner
  het op basis van hoe lang het script de vorige keer duurde, en anders loopt
  er een bewegend lijntje met de verstreken tijd erbij.
- Opgelost: een lijstveld met meer dan één waarde (bijvoorbeeld Filter met
  *.xlsm en *.xlsx) kwam als één tekst "*.xlsm,*.xlsx" bij het script aan,
  zodat er niets gevonden werd. En de foutmelding bij een opstartfout
  verwees naar een functie die niet bestond.

Waar de scripts staan
---------------------
In de submap `Scripts` naast dit bestand. Elk .ps1-bestand daar verschijnt
automatisch in de lijst. Voeg er zelf een script bij en het staat er de
volgende keer bij.

Hoe je een script geschikt maakt
--------------------------------
Zet bovenaan een `param()`-blok met per parameter een HelpMessage:

    <#
    .SYNOPSIS
    Korte omschrijving die in de lijst komt te staan.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, HelpMessage = "De map waarin gezocht wordt")]
        [string]$Map,

        [Parameter(HelpMessage = "Aanvinken om het juist terug te draaien")]
        [switch]$Terugdraaien
    )

Namen die op map, pad, folder of directory lijken krijgen een mapkiezer.
Namen met bestand of csv erin krijgen een bestandskiezer. `[switch]` wordt een
aan/uit-knopje, `[ValidateSet(...)]` een keuze, `[string[]]` een invoervak
waarin je één waarde per regel zet.

Gebruik je `SupportsShouldProcess` in je script, dan verschijnt er een
Proefrun-knopje: het script vertelt dan wat het zou doen, zonder het te doen.
"""

from __future__ import annotations

import base64
import json
import math
import os
import queue
import re
import shutil
import struct
import subprocess
import sys
import threading
import time
import traceback
import zlib
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

HIER = Path(__file__).resolve().parent
SCRIPTMAP = HIER / "Scripts"
LOGMAP = HIER / "Logs"


# ==========================================================================
# INSTELLINGEN
# ==========================================================================

# Logbestanden ouder dan dit aantal dagen worden bij het starten opgeruimd.
# 0 = nooit opruimen.
LOG_BEWAARDAGEN = 90

# Hoeveel eerder gebruikte mappen er in het uitklaplijstje van een mapveld
# blijven staan.
MAX_RECENTE_MAPPEN = 12

# De knopjes onder een filterveld. Aanklikken zet het bestandstype erbij,
# nog eens klikken haalt het weg. Zelf typen in het vak blijft gewoon kan.
SNELFILTERS = ["*.*", "*.xlsm", "*.xlsx", "*.pdf", "*.docx", "*.csv"]

# Een vaste keuze met zoveel opties of minder wordt een rij knopjes; met
# meer opties een uitklaplijst.
MAX_KEUZEKNOPPEN = 4

# ==========================================================================


def pas_venstergrootte(venster, breedte: int, hoogte: int) -> None:
    """Open zo groot als gevraagd, maar nooit groter dan het scherm.

    Een vaste hoogte valt bij een lager scherm of een grote taakbalk deels
    buiten beeld, en dan staan de knoppen onderaan er niet op.
    """
    scherm_b = venster.winfo_screenwidth()
    scherm_h = venster.winfo_screenheight()
    b = max(700, min(breedte, scherm_b - 80))
    h = max(560, min(hoogte, scherm_h - 80))
    venster.geometry(f"{b}x{h}+{max(0, (scherm_b - b) // 2)}+{max(0, (scherm_h - h) // 3)}")


def toon_startfout(titel: str, tekst: str, kopieer: str = "") -> None:
    """Foutmelding in een venster; er is geen terminal om in te kijken."""
    root = tk.Tk()
    root.withdraw()
    if kopieer:
        try:
            root.clipboard_clear()
            root.clipboard_append(kopieer)
            root.update()
        except tk.TclError:
            pass
    messagebox.showerror(titel, tekst)
    root.destroy()


def vind_powershell() -> str | None:
    """PowerShell 7 als die er is, anders de Windows-versie die er altijd is."""
    for naam in ("pwsh.exe", "pwsh", "powershell.exe", "powershell"):
        pad = shutil.which(naam)
        if pad:
            return pad
    standaard = Path(os.environ.get("SystemRoot", r"C:\Windows")) / \
        "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    return str(standaard) if standaard.exists() else None


POWERSHELL = vind_powershell()

GEEN_VENSTER = getattr(subprocess, "CREATE_NO_WINDOW", 0)

INSTELLINGEN_PAD = (
    Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "PowerShell-runner" / "instellingen.json"
)


# ==========================================================================
# 1. Parameters van een script uitlezen
# ==========================================================================

@dataclass
class Parameter:
    naam: str
    soort: str = "tekst"          # tekst, map, bestand, vinkje, keuze, lijst, getal
    verplicht: bool = False
    uitleg: str = ""
    keuzes: list[str] = field(default_factory=list)
    standaard: str = ""

    @property
    def label(self) -> str:
        return leesbare_naam(self.naam)


@dataclass
class Script:
    pad: Path
    omschrijving: str = ""
    parameters: list[Parameter] = field(default_factory=list)
    proefrun_mogelijk: bool = False
    fout: str = ""
    waarschuwing: str = ""

    @property
    def naam(self) -> str:
        return self.pad.stem


# Woorden die in hoofdletters horen, en namen die aan elkaar blijven.
_AFKORTINGEN = {"html": "HTML", "csv": "CSV", "pdf": "PDF", "kb": "KB", "mb": "MB",
                "gb": "GB", "id": "ID", "url": "URL"}


def leesbare_naam(naam: str) -> str:
    """Van 'MinimaleGrootteKB' naar 'Minimale grootte KB'.

    Parameternamen zijn aan elkaar geschreven omdat PowerShell dat wil. In
    het venster lees je liever gewone woorden.
    """
    delen = re.findall(r"[A-Z]+(?![a-z])|[A-Z]?[a-z]+|\d+", naam)
    if not delen:
        return naam
    woorden = []
    for i, deel in enumerate(delen):
        laag = deel.lower()
        if laag in _AFKORTINGEN:
            woorden.append(_AFKORTINGEN[laag])
        elif deel.isupper():                    # KB, of de A in MapA
            woorden.append(deel)
        elif i == 0:
            woorden.append(deel[:1].upper() + deel[1:])
        else:
            woorden.append(laag)
    tekst = " ".join(woorden)
    return tekst.replace("Share point", "SharePoint").replace("share point", "SharePoint")


def controleer_codering(pad: Path) -> str:
    """Waarschuw als een script accenten heeft maar geen BOM.

    PowerShell 5.1, de versie die standaard op Windows staat, leest een
    .ps1-bestand als Windows-1252 tenzij er een byte order mark voor staat.
    Zonder BOM wordt 'kopieën' dan 'kopieÃ«n' — en dat zie je pas terug in de
    uitvoer of in een rapport, niet in de code.
    """
    try:
        rauw = pad.read_bytes()
    except OSError:
        return ""
    if rauw.startswith(b"\xef\xbb\xbf"):
        return ""
    try:
        tekst = rauw.decode("utf-8")
    except UnicodeDecodeError:
        return ""                      # geen UTF-8, dus geen UTF-8-probleem
    accenten = sorted({c for c in tekst if ord(c) > 127})
    if not accenten:
        return ""
    return ("bevat " + "".join(accenten[:6]) + " maar is opgeslagen zonder BOM. "
            "PowerShell 5.1 leest die tekens dan verkeerd. Sla het bestand op als "
            "'UTF-8 met BOM' (in VS Code rechtsonder bij de codering).")


PAD_WOORDEN = ("map", "pad", "folder", "directory", "dir")
BESTAND_WOORDEN = ("bestand", "file", "csv", "rapport", "log")


def soort_van(naam: str, type_naam: str, keuzes: list[str]) -> str:
    """Welk invoerveld hoort bij deze parameter?"""
    laag = naam.lower()
    type_laag = (type_naam or "").lower()

    if "switch" in type_laag or "boolean" in type_laag:
        return "vinkje"
    if keuzes:
        return "keuze"
    if "[]" in type_laag:
        return "lijst"
    if any(x in type_laag for x in ("int", "double", "decimal", "single")):
        return "getal"
    # Naam is hier betrouwbaarder dan het type: alles is [string].
    # Bestand vóór map controleren, want 'RapportCsv' bevat geen 'map'
    # maar 'Doelmap' wel, en 'Logbestand' moet een bestandskiezer krijgen.
    if any(w in laag for w in BESTAND_WOORDEN):
        return "bestand"
    if any(w in laag for w in PAD_WOORDEN):
        return "map"
    return "tekst"


# PowerShell zelf laten vertellen wat de parameters zijn: betrouwbaarder dan
# de tekst ontleden, want het snapt attributen, types en overervingen.
INTROSPECTIE = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$algemeen = @(
    'Verbose','Debug','ErrorAction','WarningAction','InformationAction',
    'ErrorVariable','WarningVariable','InformationVariable','OutVariable',
    'OutBuffer','PipelineVariable','WhatIf','Confirm','ProgressAction'
)
$uit = @()
foreach ($pad in $args) {
    $rij = [ordered]@{ Pad = $pad; Omschrijving = ''; Proefrun = $false; Parameters = @(); Fout = '' }
    try {
        $cmd = Get-Command -Name $pad -ErrorAction Stop
        try {
            $hulp = Get-Help $pad -ErrorAction SilentlyContinue
            if ($hulp -and $hulp.Synopsis) { $rij.Omschrijving = ($hulp.Synopsis -join ' ').Trim() }
        } catch { }
        $rij.Proefrun = $cmd.Parameters.ContainsKey('WhatIf')
        $lijst = @()
        foreach ($sleutel in $cmd.Parameters.Keys) {
            if ($algemeen -contains $sleutel) { continue }
            $p = $cmd.Parameters[$sleutel]
            $attr = $p.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] } | Select-Object -First 1
            $set = $p.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | Select-Object -First 1
            $lijst += [ordered]@{
                Naam      = $p.Name
                Type      = $p.ParameterType.ToString()
                Verplicht = [bool]($attr -and $attr.Mandatory)
                Uitleg    = if ($attr) { [string]$attr.HelpMessage } else { '' }
                Keuzes    = if ($set) { @($set.ValidValues) } else { @() }
            }
        }
        $rij.Parameters = $lijst
    }
    catch {
        $rij.Fout = $_.Exception.Message
    }
    $uit += $rij
}
ConvertTo-Json -InputObject $uit -Depth 6 -Compress
"""


def regex_parameters(tekst: str) -> tuple[list[Parameter], bool, str]:
    """Terugval als PowerShell niet bevraagd kan worden.

    Minder slim dan PowerShell zelf, maar genoeg om een formulier te bouwen.
    """
    omschrijving = ""
    synopsis = re.search(r"\.SYNOPSIS\s*\n\s*(.+)", tekst)
    if synopsis:
        omschrijving = synopsis.group(1).strip()

    proefrun = bool(re.search(r"SupportsShouldProcess", tekst, re.IGNORECASE))

    blok = re.search(r"\bparam\s*\((.*?)\n\s*\)\s*$", tekst,
                     re.IGNORECASE | re.DOTALL | re.MULTILINE)
    if not blok:
        blok = re.search(r"\bparam\s*\((.*)\)", tekst, re.IGNORECASE | re.DOTALL)
    if not blok:
        return [], proefrun, omschrijving

    inhoud = blok.group(1)
    parameters: list[Parameter] = []

    # Splitsen op de komma's die een parametergroep afsluiten: vlak voor de
    # volgende [Parameter( of het volgende [type]$naam.
    stukken = re.split(r",\s*(?=(?:\[|\$))", inhoud)
    for stuk in stukken:
        naam_match = re.search(r"\$([A-Za-z_]\w*)", stuk)
        if not naam_match:
            continue
        naam = naam_match.group(1)

        type_match = re.findall(r"\[([A-Za-z_][\w\.]*(?:\[\])?)\]", stuk)
        types = [t for t in type_match
                 if not t.lower().startswith(("parameter", "validate", "cmdletbinding",
                                              "alias", "allownull", "allowempty",
                                              "supportsshouldprocess"))]
        type_naam = types[-1] if types else "string"

        keuzes = []
        set_match = re.search(r"ValidateSet\((.*?)\)", stuk, re.DOTALL)
        if set_match:
            keuzes = [k.strip().strip('"\'') for k in set_match.group(1).split(",")
                      if k.strip()]

        uitleg = ""
        help_match = re.search(r'HelpMessage\s*=\s*["\'](.*?)["\']', stuk, re.DOTALL)
        if help_match:
            uitleg = help_match.group(1).strip()

        verplicht = bool(re.search(r"Mandatory(?!\s*=\s*\$false)", stuk, re.IGNORECASE))

        standaard = ""
        std_match = re.search(rf"\${naam}\s*=\s*(.+)$", stuk, re.DOTALL)
        if std_match:
            rauw = std_match.group(1).strip().rstrip(",").strip()
            rauw = re.sub(r"^@\(|\)$", "", rauw).strip()
            standaard = ", ".join(d.strip().strip('"\'') for d in rauw.split(",")
                                  if d.strip() and d.strip() not in ('""', "''"))

        parameters.append(Parameter(
            naam=naam,
            soort=soort_van(naam, type_naam, keuzes),
            verplicht=verplicht,
            uitleg=uitleg,
            keuzes=keuzes,
            standaard=standaard,
        ))

    return parameters, proefrun, omschrijving


def lees_scripts(map_: Path) -> list[Script]:
    """Alle .ps1-scripts in de map, met hun parameters."""
    paden = sorted(map_.glob("*.ps1"), key=lambda p: p.name.lower()) if map_.is_dir() else []
    if not paden:
        return []

    scripts = {p: Script(pad=p) for p in paden}

    # Eerst de tekstuele terugval, zodat er altijd iets staat.
    for pad, script in scripts.items():
        try:
            tekst = pad.read_text(encoding="utf-8-sig", errors="replace")
        except OSError as e:
            script.fout = str(e)
            continue
        script.parameters, script.proefrun_mogelijk, script.omschrijving = \
            regex_parameters(tekst)
        script.waarschuwing = controleer_codering(pad)

    # Daarna PowerShell zelf laten kijken; dat overschrijft de terugval.
    if POWERSHELL:
        try:
            resultaat = subprocess.run(
                [POWERSHELL, "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-Command", INTROSPECTIE,
                 *[str(p) for p in paden]],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
                timeout=90, creationflags=GEEN_VENSTER)
            gegevens = json.loads(resultaat.stdout.strip() or "[]")
            if isinstance(gegevens, dict):
                gegevens = [gegevens]
            for rij in gegevens:
                pad = Path(rij.get("Pad", ""))
                script = scripts.get(pad)
                if script is None:
                    continue
                if rij.get("Fout"):
                    script.fout = rij["Fout"]
                    continue
                if rij.get("Omschrijving"):
                    script.omschrijving = rij["Omschrijving"]
                script.proefrun_mogelijk = bool(rij.get("Proefrun"))
                bestaand = {p.naam.lower(): p for p in script.parameters}
                nieuw = []
                for p in rij.get("Parameters") or []:
                    keuzes = list(p.get("Keuzes") or [])
                    param = Parameter(
                        naam=p.get("Naam", ""),
                        soort=soort_van(p.get("Naam", ""), p.get("Type", ""), keuzes),
                        verplicht=bool(p.get("Verplicht")),
                        uitleg=(p.get("Uitleg") or "").strip(),
                        keuzes=keuzes,
                        # standaardwaarden staan niet in de metadata; die komen
                        # uit de tekstuele terugval van hierboven
                        standaard=getattr(bestaand.get(p.get("Naam", "").lower()),
                                          "standaard", ""),
                    )
                    if not param.uitleg:
                        param.uitleg = getattr(
                            bestaand.get(param.naam.lower()), "uitleg", "")
                    nieuw.append(param)
                if nieuw:
                    script.parameters = nieuw
        except (subprocess.SubprocessError, json.JSONDecodeError, OSError):
            pass                      # terugval blijft staan

    return [scripts[p] for p in paden]


# ==========================================================================
# 2. Opdracht opbouwen
# ==========================================================================
#
# Het script wordt niet meer met -File gestart, maar via een klein
# PowerShell-voorprogramma (STARTER). Twee redenen:
#
# 1. Lijsten. Met -File kun je geen lijst meegeven: "-Filter *.xlsm,*.xlsx"
#    komt bij het script aan als één tekst met een komma erin, en dan vindt
#    het niets. De starter krijgt alle invoer als JSON en geeft ze door als
#    echte waarden, dus een lijst blijft een lijst en een getal een getal.
# 2. Voortgang. De starter vervangt Write-Progress door een versie die een
#    herkenbare regel op de uitvoer zet. Zo kan de runner een echte
#    voortgangslijn tonen, zonder dat je in je script iets anders hoeft te
#    doen dan gewoon Write-Progress te gebruiken.
#
# De exitcode: een script dat zelf 'exit 1' doet, eindigt met 1. Eindigt het
# gewoon, dan is het 0, ook als de laatste opdracht robocopy was (dat geeft
# 1 terug als er iets gekopieerd is, en dat is geen fout).

VOORTGANG_MERK = "##VOORTGANG##"

STARTER = r"""
try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch { }
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
if ($PSStyle) { $PSStyle.OutputRendering = 'PlainText' }
$ProgressPreference = 'SilentlyContinue'
function global:Write-Progress {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)][string]$Activity = '',
        [Parameter(Position = 1)][string]$Status = '',
        [Parameter(Position = 2)][int]$Id = 0,
        [int]$PercentComplete = -1,
        [int]$SecondsRemaining = -1,
        [string]$CurrentOperation = '',
        [int]$ParentId = -1,
        [switch]$Completed,
        [int]$SourceId = 0
    )
    if ($ParentId -ge 0) { return }
    $delen = @($Activity, $Status, $CurrentOperation) | ForEach-Object { ([string]$_) -replace '[\r\n|]', ' ' }
    [Console]::Out.WriteLine("##VOORTGANG##|$PercentComplete|$([int]$Completed.IsPresent)|$($delen[0])|$($delen[1])|$($delen[2])")
    [Console]::Out.Flush()
}
$invoer = $env:PSRUNNER_INVOER | ConvertFrom-Json
$p = @{}
if ($invoer) { foreach ($e in $invoer.PSObject.Properties) { $p[$e.Name] = $e.Value } }
$global:LASTEXITCODE = 0
try {
    & $env:PSRUNNER_SCRIPT @p
    $gelukt = $?
}
catch {
    Write-Host ("FOUT: " + $_.Exception.Message)
    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
        Write-Host $_.InvocationInfo.PositionMessage
    }
    exit 1
}
if ($gelukt) { exit 0 } else { exit [Math]::Max(1, [int]$LASTEXITCODE) }
"""


def bouw_invoer(script: Script, waarden: dict[str, object], proefrun: bool) -> dict:
    """De invoer voor het script, als waarden van het juiste soort.

    Een leeg veld laat de parameter weg, zodat de standaard uit het script
    zelf geldt. Een uitgezet knopje ook.
    """
    invoer: dict[str, object] = {}
    for parameter in script.parameters:
        waarde = waarden.get(parameter.naam)

        if parameter.soort == "vinkje":
            if waarde:
                invoer[parameter.naam] = True
            continue

        if parameter.soort == "lijst":
            regels = [r.strip() for r in re.split(r"[\r\n]+", str(waarde or ""))
                      if r.strip()]
            if regels:
                invoer[parameter.naam] = regels
            continue

        tekst = str(waarde or "").strip()
        if tekst == "":
            continue
        if parameter.soort == "getal":
            try:
                invoer[parameter.naam] = int(tekst)
            except ValueError:
                try:
                    invoer[parameter.naam] = float(tekst.replace(",", "."))
                except ValueError:
                    invoer[parameter.naam] = tekst      # PowerShell meldt het zelf
            continue
        invoer[parameter.naam] = tekst

    if proefrun and script.proefrun_mogelijk:
        invoer["WhatIf"] = True
    return invoer


def bouw_uitvoering(script: Script, invoer: dict,
                    powershell: str | None = None) -> tuple[list[str], dict]:
    """De opdracht en de omgeving waarmee het script echt gestart wordt."""
    ps = powershell or POWERSHELL
    if not ps:
        raise RuntimeError("PowerShell niet gevonden op deze computer.")
    gecodeerd = base64.b64encode(STARTER.encode("utf-16-le")).decode("ascii")
    # -OutputFormat Text: zonder dit stuurt PowerShell, als het door een
    # ander programma gestart wordt, alle uitvoer óók nog als XML mee
    # ("#< CLIXML <Objs ..."), dwars door de gewone regels heen.
    opdracht = [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                "-OutputFormat", "Text", "-EncodedCommand", gecodeerd]
    omgeving = dict(os.environ)
    omgeving["PSRUNNER_SCRIPT"] = str(script.pad)
    omgeving["PSRUNNER_INVOER"] = json.dumps(invoer, ensure_ascii=False)
    return opdracht, omgeving


def _ps_tekst(waarde: str) -> str:
    return "'" + str(waarde).replace("'", "''") + "'"


def opdracht_als_tekst(script: Script, invoer: dict) -> str:
    """Hoe je hetzelfde zelf in een PowerShell-venster zou typen.

    Komt bovenaan het logbestand, zodat je later kunt zien én nadoen wat er
    precies gedraaid heeft.
    """
    delen = ["&", _ps_tekst(script.pad)]
    for naam, waarde in invoer.items():
        if waarde is True:
            delen.append(f"-{naam}")
        elif isinstance(waarde, list):
            delen.append(f"-{naam} " + ",".join(_ps_tekst(w) for w in waarde))
        elif isinstance(waarde, (int, float)):
            delen.append(f"-{naam} {waarde}")
        else:
            delen.append(f"-{naam} {_ps_tekst(waarde)}")
    return " ".join(delen)


def ontbrekende_verplichte(script: Script, waarden: dict[str, object]) -> list[Parameter]:
    ontbreekt = []
    for parameter in script.parameters:
        if not parameter.verplicht or parameter.soort == "vinkje":
            continue
        if not str(waarden.get(parameter.naam) or "").strip():
            ontbreekt.append(parameter)
    return ontbreekt


def foute_getallen(script: Script, waarden: dict[str, object]) -> list[Parameter]:
    fout = []
    for parameter in script.parameters:
        if parameter.soort != "getal":
            continue
        tekst = str(waarden.get(parameter.naam) or "").strip()
        if tekst and not re.fullmatch(r"-?\d+([.,]\d+)?", tekst):
            fout.append(parameter)
    return fout


def lees_voortgang(regel: str) -> dict | None:
    """Een voortgangsregel van de starter uit elkaar halen."""
    if not regel.startswith(VOORTGANG_MERK + "|"):
        return None
    delen = regel.split("|", 5)
    while len(delen) < 6:
        delen.append("")
    try:
        procent = int(delen[1])
    except ValueError:
        procent = -1
    return {"procent": procent, "klaar": delen[2].strip() == "1",
            "activiteit": delen[3].strip(), "status": delen[4].strip(),
            "bezig_met": delen[5].strip()}


_ANSI = re.compile(r"\x1b\[[0-9;?]*[A-Za-z]")


def is_clixml(regel: str) -> bool:
    """PowerShell-XML die niet voor mensen bedoeld is (zie bouw_uitvoering)."""
    begin = regel.lstrip()
    return begin.startswith(("#< CLIXML", "<Objs Version=")) or (
        begin.startswith("<") and "schemas.microsoft.com/powershell" in begin)
_RAPPORT = re.compile(r"([A-Za-z]:\\[^\"<>|\r\n*?]+?\.(?:html?|csv|xlsx))", re.IGNORECASE)


def vind_rapporten(regel: str) -> list[Path]:
    """Paden naar rapporten die in een uitvoerregel genoemd worden."""
    gevonden = []
    for tekst in _RAPPORT.findall(regel):
        pad = Path(tekst.strip())
        try:
            if pad.is_file():
                gevonden.append(pad)
        except OSError:
            pass
    return gevonden


def klok(seconden: float) -> str:
    seconden = max(0, int(round(seconden)))
    minuten, sec = divmod(seconden, 60)
    uren, minuten = divmod(minuten, 60)
    return f"{uren}:{minuten:02d}:{sec:02d}" if uren else f"{minuten}:{sec:02d}"


# ==========================================================================
# 3. Logbestanden en onthouden instellingen
# ==========================================================================

def logpad(scriptnaam: str) -> Path:
    veilig = re.sub(r"[^\w\-. ]", "_", scriptnaam).strip()
    stempel = datetime.now().strftime("%Y%m%d-%H%M%S")
    return LOGMAP / f"{stempel} {veilig}.log"


def ruim_logs_op(dagen: int = LOG_BEWAARDAGEN) -> int:
    """Oude logbestanden weggooien zodat de map niet dichtslibt."""
    if not dagen or not LOGMAP.is_dir():
        return 0
    grens = datetime.now().timestamp() - dagen * 86400
    verwijderd = 0
    for bestand in LOGMAP.glob("*.log"):
        try:
            if bestand.stat().st_mtime < grens:
                bestand.unlink()
                verwijderd += 1
        except OSError:
            pass
    return verwijderd


def laad_instellingen(pad: Path | None = None) -> dict:
    """Favorieten, laatst ingevulde waarden, recente mappen en looptijden.

    Staat in je eigen profiel (LOCALAPPDATA), niet naast de scripts: zo heeft
    iedere collega die dezelfde map gebruikt zijn eigen favorieten.
    """
    pad = pad or INSTELLINGEN_PAD
    try:
        data = json.loads(pad.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def bewaar_instellingen(data: dict, pad: Path | None = None) -> None:
    pad = pad or INSTELLINGEN_PAD
    try:
        pad.parent.mkdir(parents=True, exist_ok=True)
        tijdelijk = pad.with_suffix(".tmp")
        tijdelijk.write_text(json.dumps(data, ensure_ascii=False, indent=1),
                             encoding="utf-8")
        tijdelijk.replace(pad)
    except OSError:
        pass                     # niet kunnen bewaren is geen reden om te stoppen


def hoe_lang_geleden(iso: str) -> str:
    try:
        toen = datetime.fromisoformat(iso)
    except (TypeError, ValueError):
        return ""
    dagen = (datetime.now().date() - toen.date()).days
    if dagen <= 0:
        return f"vandaag om {toen:%H:%M}"
    if dagen == 1:
        return "gisteren"
    if dagen < 14:
        return f"{dagen} dagen geleden"
    return f"op {toen:%d-%m-%Y}"


# ==========================================================================
# 4. De stijl van het venster
# ==========================================================================
#
# Overgenomen uit Excel naar PDF en de Adressen-tool, zodat de tools er
# hetzelfde uitzien: lichtgrijze achtergrond, witte kaarten met ronde hoeken,
# knoppen als pil.

KLEUR = {
    "achtergrond": "#F5F5F7",
    "kaart": "#FFFFFF",
    "tekst": "#1D1D1F",
    "zacht": "#636366",
    "lichter": "#A1A1A6",
    "lijn": "#E5E5EA",
    "rand": "#D2D2D7",
    "blauw": "#0071E3",
    "blauw_donker": "#005BBF",
    "blauw_diep": "#004A9E",
    "blauw_licht": "#EAF3FE",
    "blauw_uit": "#9CC3F0",
    "groen": "#248A3D",
    "oranje": "#B25000",
    "rood": "#D70015",
    "ster": "#FF9F0A",
    "geel_vlak": "#FFF6DB",
    "rood_vlak": "#FDE8E6",
    "zebra": "#FAFAFC",
    "selectie": "#D8E8FC",
    "vlak": "#FBFBFD",
    "knop": "#E8E8ED",
    "knop_zweef": "#DCDCE1",
    "knop_druk": "#D1D1D6",
    "knop_uit": "#F0F0F3",
    "duim": "#C7C7CC",
    "duim_zweef": "#AEAEB2",
    "duim_druk": "#8E8E93",
}

# Hoe rond de hoeken zijn, in beeldpunten bij 100% schermvergroting.
ROND = {
    "kaart": 18,
    "veld": 10,
    "schuif": 4,
}

# Binnenwit van knoppen, in beeldpunten bij 100%: (links/rechts, boven/onder).
KNOP_BINNENWIT = (16, 6)

LETTER_VENSTER = "Segoe UI"
LETTER_CODE = "Consolas"
SCHAAL = 1.0          # 1.0 bij 100% schermvergroting, 1.5 bij 150%


def letter(grootte: int = 10, gewicht: str = "normal") -> tuple:
    return (LETTER_VENSTER, grootte, gewicht)


def px(waarde: float) -> int:
    """Een maat in beeldpunten, meegeschaald met de Windows-vergroting."""
    return int(round(waarde * SCHAAL))


def maak_scherp() -> None:
    """Tekst scherp in plaats van wazig op een scherm met vergroting.

    Zonder dit rekt Windows het hele venster als plaatje op, en dan wordt
    alle tekst vaag. Moet gebeuren voordat het eerste venster er is.
    """
    if sys.platform != "win32":
        return
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            import ctypes
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


# ------------------------------------------------------- ronde vormen
#
# Tk tekent zelf geen ronde hoeken. Daarom maken we de vormen zelf als
# PNG-plaatje met een zachte rand, pixel voor pixel. Een knop of vak is
# zo'n plaatje dat ttk uitrekt: de hoeken blijven staan, het midden groeit
# mee. Alles met de standaardbibliotheek, zonder extra pakket.

_BEELDEN: list = []      # Tk gooit een plaatje weg zodra Python het vergeet


def _rgb(kleur: str) -> tuple[int, int, int]:
    kleur = kleur.lstrip("#")
    return int(kleur[0:2], 16), int(kleur[2:4], 16), int(kleur[4:6], 16)


def _png(breedte: int, hoogte: int, rijen: list[bytes]) -> bytes:
    """Een RGBA-plaatje als PNG-bestand in het geheugen."""
    def blok(soort: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + soort + data
                + struct.pack(">I", zlib.crc32(soort + data) & 0xFFFFFFFF))
    ruw = b"".join(b"\x00" + rij for rij in rijen)
    return (b"\x89PNG\r\n\x1a\n"
            + blok(b"IHDR", struct.pack(">IIBBBBB", breedte, hoogte, 8, 6, 0, 0, 0))
            + blok(b"IDAT", zlib.compress(ruw, 6)) + blok(b"IEND", b""))


def afgerond_png(breedte: int, hoogte: int, straal: float, vulling: str,
                 rand: str | None = None, dikte: float = 1.0) -> bytes:
    """Een rechthoek met ronde hoeken en zachte randen, als PNG.

    Per pixel wordt uitgerekend hoe ver hij van de rand ligt; een pixel die
    half op de rand valt wordt half doorzichtig. Dat maakt de hoeken glad.
    """
    vul = _rgb(vulling)
    lijn = _rgb(rand) if rand else vul
    hb, hh = breedte / 2, hoogte / 2
    r = max(0.5, min(straal, hb, hh))

    def pixel(x: int, y: int) -> bytes:
        qx = abs(x + 0.5 - hb) - (hb - r)
        qy = abs(y + 0.5 - hh) - (hh - r)
        afstand = math.hypot(max(qx, 0.0), max(qy, 0.0)) + min(max(qx, qy), 0.0) - r
        dekking = min(1.0, max(0.0, 0.5 - afstand))
        if dekking <= 0:
            return b"\x00\x00\x00\x00"
        if rand:
            binnen = min(1.0, max(0.0, 0.5 - (afstand + dikte)))
            deel = (dekking - binnen) / dekking
            kleur = tuple(round(v * (1 - deel) + l * deel) for v, l in zip(vul, lijn))
        else:
            kleur = vul
        return bytes((*kleur, round(dekking * 255)))

    rand_zone = int(math.ceil(dikte)) + 2
    vol = bytes((*vul, 255))
    rijen = []
    for y in range(hoogte):
        if y < r or y >= hoogte - r:
            # Alleen de hoeken zelf hoeven pixel voor pixel; daartussen is de
            # boven- of onderrand overal gelijk.
            hoek = min(breedte // 2, int(math.ceil(r)) + 2)
            links = b"".join(pixel(x, y) for x in range(hoek))
            rechts = b"".join(pixel(x, y) for x in range(breedte - hoek, breedte))
            rij = links + pixel(hoek, y) * (breedte - 2 * hoek) + rechts
        else:
            links = b"".join(pixel(x, y) for x in range(min(rand_zone, breedte)))
            midden = max(0, breedte - 2 * rand_zone)
            rechts = b"".join(pixel(x, y)
                              for x in range(max(rand_zone, breedte - rand_zone), breedte))
            rij = links + vol * midden + rechts
        rijen.append(rij)
    return _png(breedte, hoogte, rijen)


def afgerond_beeld(breedte: int, hoogte: int, straal: float, vulling: str,
                   rand: str | None = None, dikte: float = 1.0) -> tk.PhotoImage:
    data = base64.b64encode(afgerond_png(breedte, hoogte, straal, vulling, rand, dikte))
    beeld = tk.PhotoImage(data=data, format="png")
    _BEELDEN.append(beeld)
    return beeld


def los_vinkje_png(maat: int, kleur: str) -> bytes:
    """Alleen een vinkje, op een doorzichtige achtergrond, voor de knopjes.

    Een plaatje in plaats van het teken ✓: dat teken zit niet in Segoe UI,
    en dan zoekt Windows bij het eerste gebruik alle lettertypen af.
    """
    rgb = _rgb(kleur)
    streep = maat * 0.085
    punten = [(x * maat, y * maat) for x, y in ((0.16, 0.53), (0.39, 0.76), (0.84, 0.27))]

    def tot_lijnstuk(px_, py_, a, b):
        dx, dy = b[0] - a[0], b[1] - a[1]
        t = max(0.0, min(1.0, ((px_ - a[0]) * dx + (py_ - a[1]) * dy) / (dx * dx + dy * dy)))
        return math.hypot(px_ - (a[0] + t * dx), py_ - (a[1] + t * dy))

    rijen = []
    for y in range(maat):
        rij = bytearray()
        for x in range(maat):
            d = min(tot_lijnstuk(x + 0.5, y + 0.5, punten[0], punten[1]),
                    tot_lijnstuk(x + 0.5, y + 0.5, punten[1], punten[2]))
            dekking = min(1.0, max(0.0, streep - d + 0.5))
            rij += bytes((*rgb, round(dekking * 255)))
        rijen.append(bytes(rij))
    return _png(maat, maat, rijen)


def ster_png(maat: int, kleur: str, vol: bool) -> bytes:
    """Een sterretje, vol of als omtrek, met gladde randen.

    Zelfde aanpak als de ronde vormen: per pixel de afstand tot de rand van
    de ster, en een pixel die half op de rand valt wordt half doorzichtig.
    """
    rgb = _rgb(kleur)
    midden = maat / 2
    buiten = maat * 0.47
    binnen = buiten * 0.45
    hoeken = []
    for i in range(10):
        hoek = -math.pi / 2 + i * math.pi / 5
        r = buiten if i % 2 == 0 else binnen
        hoeken.append((midden + r * math.cos(hoek), midden + 0.04 * maat + r * math.sin(hoek)))
    randen = list(zip(hoeken, hoeken[1:] + hoeken[:1]))
    dikte = max(1.0, maat / 12)

    def afstand(x: float, y: float) -> float:
        kortst = min(_tot_lijnstuk(x, y, a, b) for a, b in randen)
        binnen_ster = False
        for (ax, ay), (bx, by) in randen:
            if (ay > y) != (by > y) and x < (bx - ax) * (y - ay) / (by - ay) + ax:
                binnen_ster = not binnen_ster
        return -kortst if binnen_ster else kortst

    rijen = []
    for y in range(maat):
        rij = bytearray()
        for x in range(maat):
            d = afstand(x + 0.5, y + 0.5)
            if vol:
                dekking = min(1.0, max(0.0, 0.5 - d))
            else:
                dekking = min(1.0, max(0.0, 0.5 - (abs(d) - dikte / 2)))
            rij += bytes((*rgb, round(dekking * 255)))
        rijen.append(bytes(rij))
    return _png(maat, maat, rijen)


def _tot_lijnstuk(px_: float, py_: float, a, b) -> float:
    dx, dy = b[0] - a[0], b[1] - a[1]
    lengte = dx * dx + dy * dy
    t = 0.0 if lengte == 0 else max(0.0, min(1.0, ((px_ - a[0]) * dx + (py_ - a[1]) * dy) / lengte))
    return math.hypot(px_ - (a[0] + t * dx), py_ - (a[1] + t * dy))


def beeld_van(png: bytes) -> tk.PhotoImage:
    beeld = tk.PhotoImage(data=base64.b64encode(png), format="png")
    _BEELDEN.append(beeld)
    return beeld


# Hoe groot het rekbare middenstuk van elk plaatje is. Tk vult een groot vak
# door dat middenstuk telkens opnieuw te tekenen; met een groot middenstuk
# zijn dat maar een handvol stukjes, en dat houdt het vergroten vlot.
MIDDEN_PX = 256


def rond_element(stijl: ttk.Style, naam: str, straal: int,
                 standen: list[tuple[tuple, str, str | None]],
                 dikte: float = 1.0, binnen=0) -> None:
    """Een rekbaar rond vlak als ttk-element, met een plaatje per toestand.

    standen: (toestand, vulling, rand). De eerste is de gewone stand;
    de rest bijvoorbeeld (("pressed",), ...) of (("focus",), ...).
    """
    zijde = 2 * straal + 2 + MIDDEN_PX
    beelden = [(toestand, afgerond_beeld(zijde, zijde, straal, vulling, rand, dikte))
               for toestand, vulling, rand in standen]
    extra = [(*toestand, beeld) for toestand, beeld in beelden[1:]]
    try:
        stijl.element_create(naam, "image", beelden[0][1], *extra,
                             border=straal + 1, padding=binnen, sticky="nsew",
                             width=2 * straal + 3, height=2 * straal + 3)
    except tk.TclError:
        pass                                   # bestaat al (tweede venster)


def _laag(vlak: str, onderdeel: str, binnen: list) -> list:
    """Een layout: achtergrond, daarop het ronde vlak, daarin de inhoud."""
    inhoud = [(f"{onderdeel}.padding", {"sticky": "nswe", "children": binnen})]
    if vlak:
        inhoud = [(vlak, {"sticky": "nswe", "children": inhoud})]
    return [(f"{onderdeel}.background", {"sticky": "nswe"})] + inhoud


def maak_stijl(root: tk.Tk) -> None:
    """Alle ttk-onderdelen in dezelfde lichte stijl, met ronde hoeken.

    Basis is het thema 'clam': het enige ingebouwde thema waarvan je de
    kleuren echt kunt instellen.
    """
    global LETTER_VENSTER, LETTER_CODE, SCHAAL
    from tkinter import font as tkfont

    try:
        SCHAAL = max(1.0, root.winfo_fpixels("1i") / 96)
    except tk.TclError:
        SCHAAL = 1.0
    beschikbaar = set(tkfont.families(root))
    for naam in ("Segoe UI", "Segoe UI Variable Text", "SF Pro Text",
                 "Helvetica Neue", "Arial"):
        if naam in beschikbaar:
            LETTER_VENSTER = naam
            break
    for naam in ("Cascadia Mono", "Consolas", "Courier New"):
        if naam in beschikbaar:
            LETTER_CODE = naam
            break

    k = KLEUR
    stijl = ttk.Style(root)
    stijl.theme_use("clam")
    root.configure(background=k["achtergrond"])
    for naam in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(naam).configure(family=LETTER_VENSTER, size=10)
        except tk.TclError:
            pass

    stijl.configure(".", background=k["achtergrond"], foreground=k["tekst"],
                    font=letter(10), bordercolor=k["rand"], lightcolor=k["kaart"],
                    darkcolor=k["rand"], troughcolor=k["lijn"],
                    focuscolor=k["blauw"], selectbackground=k["selectie"],
                    selectforeground=k["tekst"])

    # ------------------------------------------------ de ronde vormen
    regel = tkfont.Font(root, font=letter(10)).metrics("linespace")
    r_knop = max(px(8), (regel + 2 * px(KNOP_BINNENWIT[1])) // 2 - 1)
    r_keuze = max(px(8), (regel + 2 * px(5)) // 2 - 1)
    r_veld, r_kaart, r_schuif = px(ROND["veld"]), px(ROND["kaart"]), px(ROND["schuif"])
    rond_element(stijl, "Grijs.vlak", r_knop, [
        ((), k["knop"], None), (("disabled",), k["knop_uit"], None),
        (("pressed",), k["knop_druk"], None), (("active",), k["knop_zweef"], None)])
    rond_element(stijl, "Blauw.vlak", r_knop, [
        ((), k["blauw"], None), (("disabled",), k["blauw_uit"], None),
        (("pressed",), k["blauw_diep"], None), (("active",), k["blauw_donker"], None)])
    rond_element(stijl, "Veld.vlak", r_veld, [
        ((), k["kaart"], k["rand"]), (("disabled",), k["achtergrond"], k["lijn"]),
        (("focus",), k["kaart"], k["blauw"])], dikte=1.0)
    rond_element(stijl, "Keuzeveld.vlak", r_veld, [
        ((), k["kaart"], k["rand"]), (("disabled",), k["achtergrond"], k["lijn"]),
        (("focus",), k["kaart"], k["blauw"])], dikte=1.0,
        binnen=(px(1), px(2), max(px(6), r_veld // 2), px(2)))
    rond_element(stijl, "Kaart.vlak", r_kaart, [((), k["kaart"], k["lijn"])])
    rond_element(stijl, "Aan.vlak", r_keuze, [((), k["blauw_licht"], k["blauw"])])
    rond_element(stijl, "Uit.vlak", r_keuze, [((), k["kaart"], k["rand"])])
    rond_element(stijl, "Rond.Vertical.Scrollbar.thumb", r_schuif, [
        ((), k["duim"], None), (("pressed",), k["duim_druk"], None),
        (("active",), k["duim_zweef"], None)])
    rond_element(stijl, "Rond.Horizontal.Scrollbar.thumb", r_schuif, [
        ((), k["duim"], None), (("pressed",), k["duim_druk"], None),
        (("active",), k["duim_zweef"], None)])

    # tekst
    stijl.configure("TLabel", background=k["achtergrond"])
    stijl.configure("Titel.TLabel", font=letter(20, "bold"))
    stijl.configure("Ondertitel.TLabel", foreground=k["zacht"], font=letter(10))
    stijl.configure("Zacht.TLabel", foreground=k["zacht"], font=letter(9))
    stijl.configure("Kaart.TLabel", background=k["kaart"])
    stijl.configure("Kop.Kaart.TLabel", background=k["kaart"], font=letter(12, "bold"))
    stijl.configure("Groot.Kaart.TLabel", background=k["kaart"], font=letter(14, "bold"))
    stijl.configure("Naam.Kaart.TLabel", background=k["kaart"], font=letter(10, "bold"))
    stijl.configure("Veld.Kaart.TLabel", background=k["kaart"], foreground=k["zacht"],
                    font=letter(9, "bold"))
    stijl.configure("Zacht.Kaart.TLabel", background=k["kaart"],
                    foreground=k["zacht"], font=letter(9))
    stijl.configure("Licht.Kaart.TLabel", background=k["kaart"],
                    foreground=k["lichter"], font=letter(9))
    stijl.configure("Verplicht.Kaart.TLabel", background=k["kaart"],
                    foreground=k["blauw"], font=letter(8))
    stijl.configure("Fout.Kaart.TLabel", background=k["kaart"],
                    foreground=k["rood"], font=letter(9))

    # vakken: een ttk.Frame met het ronde kaartplaatje als achtergrond
    stijl.layout("Kaart.TFrame", [("Frame.background", {"sticky": "nswe"}),
                                  ("Kaart.vlak", {"sticky": "nswe"})])
    stijl.configure("Kaart.TFrame", background=k["achtergrond"])
    stijl.configure("Wit.TFrame", background=k["kaart"])

    # knoppen: grijs en rond, de hoofdknop blauw, de rest als link
    label = [("Button.label", {"sticky": "nswe"})]
    for naam, vlak, achter in (("TButton", "Grijs.vlak", k["achtergrond"]),
                               ("Kaart.TButton", "Grijs.vlak", k["kaart"]),
                               ("Primair.TButton", "Blauw.vlak", k["achtergrond"]),
                               ("Kaart.Primair.TButton", "Blauw.vlak", k["kaart"]),
                               ("Link.TButton", "", k["achtergrond"]),
                               ("Kaart.Link.TButton", "", k["kaart"])):
        stijl.layout(naam, _laag(vlak, "Button", label))
        stijl.configure(naam, background=achter)
        stijl.map(naam, background=[], lightcolor=[], darkcolor=[], bordercolor=[])
    stijl.configure("TButton", foreground=k["tekst"], font=letter(10), width=0,
                    padding=(px(KNOP_BINNENWIT[0]), px(KNOP_BINNENWIT[1])))
    stijl.map("TButton", foreground=[("disabled", k["lichter"])])
    stijl.configure("Kaart.TButton", padding=(px(12), px(KNOP_BINNENWIT[1])))
    for naam in ("Primair.TButton", "Kaart.Primair.TButton"):
        stijl.configure(naam, foreground="#FFFFFF", font=letter(10, "bold"),
                        padding=(px(KNOP_BINNENWIT[0] + 8), px(KNOP_BINNENWIT[1] + 1)))
        stijl.map(naam, foreground=[("disabled", "#FFFFFF")])
    for naam in ("Link.TButton", "Kaart.Link.TButton"):
        stijl.configure(naam, foreground=k["blauw"], padding=(px(6), px(3)),
                        font=letter(10), width=0)
        stijl.map(naam, foreground=[("disabled", k["lichter"]),
                                    ("active", k["blauw_donker"])])

    # aan/uit-knopjes (Keuzeknop): een pil met tekst erin
    for naam, vlak, kleur in (("Aan.Keuze.TLabel", "Aan.vlak", k["blauw"]),
                              ("Uit.Keuze.TLabel", "Uit.vlak", k["zacht"])):
        stijl.layout(naam, _laag(vlak, "Label", [("Label.label", {"sticky": "nswe"})]))
        stijl.configure(naam, background=k["kaart"], foreground=kleur,
                        font=letter(10), padding=(px(13), px(5)), anchor="center")

    # invulvelden: rond, blauwe rand zodra je erin typt
    stijl.layout("TEntry", _laag("Veld.vlak", "Entry",
                                 [("Entry.textarea", {"sticky": "nswe"})]))
    stijl.configure("TEntry", background=k["kaart"], foreground=k["tekst"],
                    fieldbackground=k["kaart"], padding=(px(10), px(6)),
                    insertcolor=k["tekst"])
    stijl.map("TEntry", background=[], fieldbackground=[], bordercolor=[],
              lightcolor=[], darkcolor=[])

    # keuzelijsten: hetzelfde ronde veld, met een pijltje rechts
    stijl.layout("TCombobox", [
        ("Combobox.background", {"sticky": "nswe"}),
        ("Keuzeveld.vlak", {"sticky": "nswe", "children": [
            ("Combobox.downarrow", {"side": "right", "sticky": "ns"}),
            ("Combobox.padding", {"sticky": "nswe", "children": [
                ("Combobox.textarea", {"sticky": "nswe"})]})]})])
    stijl.configure("TCombobox", background=k["kaart"], foreground=k["tekst"],
                    fieldbackground=k["kaart"], padding=(px(9), px(4), px(2), px(4)),
                    arrowcolor=k["zacht"], arrowsize=px(11), bordercolor=k["kaart"],
                    lightcolor=k["kaart"], darkcolor=k["kaart"], insertcolor=k["tekst"])
    stijl.map("TCombobox", background=[], fieldbackground=[], bordercolor=[],
              lightcolor=[], darkcolor=[],
              arrowcolor=[("disabled", k["lichter"]), ("active", k["blauw"])],
              selectbackground=[("readonly", k["kaart"])],
              selectforeground=[("readonly", k["tekst"])])
    root.option_add("*TCombobox*Listbox.font", letter(10))
    root.option_add("*TCombobox*Listbox.background", k["kaart"])
    root.option_add("*TCombobox*Listbox.foreground", k["tekst"])
    root.option_add("*TCombobox*Listbox.selectBackground", k["selectie"])
    root.option_add("*TCombobox*Listbox.selectForeground", k["tekst"])
    root.option_add("*TCombobox*Listbox.borderWidth", 0)

    # de scriptlijst
    stijl.configure("Treeview", background=k["kaart"], fieldbackground=k["kaart"],
                    foreground=k["tekst"], rowheight=regel + px(14),
                    borderwidth=0, font=letter(10), indent=0)
    stijl.map("Treeview", background=[("selected", k["selectie"])],
              foreground=[("selected", k["tekst"])])
    stijl.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
    stijl.layout("Item", [
        ("Treeitem.padding", {"sticky": "nswe", "children": [
            ("Treeitem.image", {"side": "left", "sticky": ""}),
            ("Treeitem.text", {"sticky": "nswe"})]})])
    stijl.configure("Item", padding=(px(6), 0))

    # schuifbalken: een smal rond duimpje, zonder pijltjes
    stijl.layout("Vertical.TScrollbar", [
        ("Vertical.Scrollbar.trough", {"sticky": "ns", "children": [
            ("Rond.Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    stijl.layout("Horizontal.TScrollbar", [
        ("Horizontal.Scrollbar.trough", {"sticky": "ew", "children": [
            ("Rond.Horizontal.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    for naam in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
        stijl.configure(naam, troughcolor=k["kaart"], bordercolor=k["kaart"],
                        lightcolor=k["kaart"], darkcolor=k["kaart"],
                        background=k["kaart"], relief="flat", borderwidth=0,
                        gripcount=0, arrowsize=px(8))

    # de schuiven tussen de kaarten: onzichtbaar, wel te pakken
    stijl.configure("TPanedwindow", background=k["achtergrond"])
    stijl.configure("Sash", sashthickness=px(14), gripcount=0,
                    background=k["achtergrond"], lightcolor=k["achtergrond"],
                    darkcolor=k["achtergrond"], bordercolor=k["achtergrond"])


def meebuigen(label) -> None:
    """Laat de tekst van een label teruglopen op de breedte die het krijgt."""
    label.bind("<Configure>", lambda e: label.configure(
        wraplength=max(px(120), e.width - px(4))), add="+")


def kaart(ouder, padx: int | None = None, pady: int | None = None) -> ttk.Frame:
    """Een wit vlak met ronde hoeken en een dun randje."""
    return ttk.Frame(ouder, style="Kaart.TFrame",
                     padding=(px(18) if padx is None else padx,
                              px(14) if pady is None else pady))


class Keuzeknop(ttk.Label):
    """Een aan/uit-knopje in plaats van een vinkje.

    Een rond pilletje dat blauw kleurt als het aan staat, met een vinkje
    ervoor. Het hele knopje is klikbaar, en met de spatiebalk werkt het ook.
    """

    _vink = None                      # het vinkje-plaatje, één keer gemaakt

    def __init__(self, ouder, tekst: str, variabele: tk.BooleanVar, bij_wissel=None):
        super().__init__(ouder, cursor="hand2", compound="left", takefocus=True)
        if Keuzeknop._vink is None:
            Keuzeknop._vink = beeld_van(los_vinkje_png(px(12), KLEUR["blauw"]))
        self.tekst = tekst
        self.variabele = variabele
        self.bij_wissel = bij_wissel
        self.bind("<Button-1>", self._klik)
        self.bind("<space>", self._klik)
        self._spoor = variabele.trace_add("write", lambda *_a: self._teken())
        self.bind("<Destroy>", self._weg, add="+")
        self._teken()

    def _weg(self, _e=None) -> None:
        try:
            self.variabele.trace_remove("write", self._spoor)
        except (tk.TclError, ValueError):
            pass

    def _klik(self, _gebeurtenis=None) -> None:
        if str(self.cget("state")) == "disabled":
            return
        self.variabele.set(not self.variabele.get())
        if self.bij_wissel:
            self.bij_wissel()

    def _teken(self) -> None:
        try:
            if self.variabele.get():
                self.configure(text=f" {self.tekst}", image=Keuzeknop._vink,
                               style="Aan.Keuze.TLabel")
            else:
                self.configure(text=self.tekst, image="", style="Uit.Keuze.TLabel")
        except tk.TclError:
            pass


class Keuzerij(tk.Frame):
    """Een rij knopjes waarvan er precies één aan staat.

    Voor een vaste keuze met weinig opties: je ziet meteen alle keuzes, en
    het is één klik in plaats van uitklappen en kiezen.
    """

    def __init__(self, ouder, opties: list[str], variabele: tk.StringVar, bij_wissel=None):
        super().__init__(ouder, background=KLEUR["kaart"])
        self.variabele = variabele
        self.bij_wissel = bij_wissel
        self.knoppen = {}
        for optie in opties:
            knop = ttk.Label(self, text=optie, cursor="hand2", style="Uit.Keuze.TLabel")
            knop.pack(side="left", padx=(0, px(6)))
            knop.bind("<Button-1>", lambda _e, o=optie: self._kies(o))
            self.knoppen[optie] = knop
        self._spoor = variabele.trace_add("write", lambda *_a: self._teken())
        self.bind("<Destroy>", self._weg, add="+")
        self._teken()

    def _weg(self, _e=None) -> None:
        try:
            self.variabele.trace_remove("write", self._spoor)
        except (tk.TclError, ValueError):
            pass

    def _kies(self, optie: str) -> None:
        self.variabele.set(optie)
        if self.bij_wissel:
            self.bij_wissel()

    def _teken(self) -> None:
        huidig = self.variabele.get()
        for optie, knop in self.knoppen.items():
            try:
                knop.configure(style="Aan.Keuze.TLabel" if optie == huidig
                               else "Uit.Keuze.TLabel")
            except tk.TclError:
                pass


class Voortgangslijn(tk.Frame):
    """Een dun blauw lijntje dat volloopt, of heen en weer loopt.

    zet(deel)  vult het lijntje tot dat deel (0..1): echte voortgang.
    lopend()   een stukje dat van links naar rechts schuift: hij is bezig,
               maar hoe ver is niet te zeggen.
    leeg()     niets.
    """

    def __init__(self, ouder):
        super().__init__(ouder, height=px(5), background=KLEUR["lijn"])
        self.vulling = tk.Frame(self, background=KLEUR["blauw"])
        self._lopend = False
        self._positie = 0.0
        self._deel = 0.0

    def zet(self, deel: float) -> None:
        self._lopend = False
        deel = max(0.0, min(1.0, deel))
        # Zacht bijschuiven in plaats van verspringen.
        self._deel = deel if deel < self._deel else self._deel + (deel - self._deel) * 0.5
        if self._deel <= 0.001:
            self.vulling.place_forget()
        else:
            self.vulling.place(x=0, y=0, relheight=1, relx=0, relwidth=self._deel)

    def vol(self) -> None:
        self._lopend = False
        self._deel = 1.0
        self.vulling.place(x=0, y=0, relheight=1, relx=0, relwidth=1)

    def lopend(self) -> None:
        if self._lopend:
            return
        self._lopend = True
        self._deel = 0.0
        self._positie = 0.0
        self._beweeg()

    def leeg(self) -> None:
        self._lopend = False
        self._deel = 0.0
        self.vulling.place_forget()

    def _beweeg(self) -> None:
        if not self._lopend:
            return
        self._positie = (self._positie + 0.012) % 1.35
        try:
            self.vulling.place(x=0, y=0, relheight=1, relx=self._positie - 0.35,
                               relwidth=0.35)
            self.after(16, self._beweeg)
        except tk.TclError:
            self._lopend = False


def maak_sleepbaar(venster: tk.Tk, bij_neerzetten) -> bool:
    """Mappen op het venster kunnen slepen, zonder extra pakket.

    Windows stuurt met DragAcceptFiles een WM_DROPFILES-bericht als je iets
    op het venster loslaat; dat vangen we hier af. bij_neerzetten wordt
    aangeroepen midden in de berichtenlus van Tk: daar mag niets met Tk
    gebeuren, dus geef hem iets dat de paden alleen in een wachtrij legt.
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        from ctypes import wintypes

        user32, shell32 = ctypes.windll.user32, ctypes.windll.shell32
        WM_DROPFILES, GWLP_WNDPROC = 0x0233, -4
        LRESULT = ctypes.c_ssize_t
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT,
                                     wintypes.WPARAM, wintypes.LPARAM)
        zet_proc = (user32.SetWindowLongPtrW if ctypes.sizeof(ctypes.c_void_p) == 8
                    else user32.SetWindowLongW)
        zet_proc.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        zet_proc.restype = ctypes.c_void_p
        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                           wintypes.WPARAM, wintypes.LPARAM]
        user32.CallWindowProcW.restype = LRESULT
        shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT,
                                           ctypes.c_wchar_p, wintypes.UINT]
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragFinish.argtypes = [ctypes.c_void_p]
        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]

        hwnd = venster.winfo_id()
        oud: dict = {}

        def afhandelen(h, bericht, wparam, lparam):
            if bericht == WM_DROPFILES:
                paden = []
                try:
                    aantal = shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)
                    for i in range(aantal):
                        lengte = shell32.DragQueryFileW(wparam, i, None, 0)
                        buffer = ctypes.create_unicode_buffer(lengte + 1)
                        shell32.DragQueryFileW(wparam, i, buffer, lengte + 1)
                        paden.append(buffer.value)
                    bij_neerzetten(paden)
                except Exception:
                    pass
                finally:
                    shell32.DragFinish(wparam)
                return 0
            return user32.CallWindowProcW(oud["proc"], h, bericht, wparam, lparam)

        terugroep = WNDPROC(afhandelen)
        shell32.DragAcceptFiles(hwnd, True)
        oud["proc"] = zet_proc(hwnd, GWLP_WNDPROC,
                               ctypes.cast(terugroep, ctypes.c_void_p))
        venster._sleep_terugroep = terugroep     # anders ruimt Python hem op
        return True
    except Exception:
        return False


# ==========================================================================
# 5. Het venster
# ==========================================================================

UITLEG = """Wat dit doet

Je kiest links een script, vult rechts de velden in en klikt op Uitvoeren. De
uitvoer loopt live mee in het onderste vak en wordt opgeslagen als logbestand.

De invulvelden worden gemaakt op basis van het script zelf. Verandert een
script, dan verandert het formulier mee; er is niets dubbel te onderhouden.


Favorieten en onthouden waarden

Klik op het sterretje voor een script (in de lijst, of rechtsboven bij het
script) en het komt bovenaan onder Favorieten te staan.

Wat je bij een script invult, wordt onthouden zodra je op Uitvoeren klikt.
Kies je het script de volgende keer, dan staat alles er weer. Met
Standaardwaarden zet je de velden terug op wat het script zelf voorstelt.

Mapvelden onthouden de laatste mappen die je gebruikte: klik op het pijltje
in het veld. Je kunt ook een map uit de Verkenner op het venster slepen; die
komt in het mapveld waar je in stond, of anders in het eerste lege mapveld.

Dit alles staat in je eigen profiel, in
%LOCALAPPDATA%\\PowerShell-runner\\instellingen.json. Collega's die dezelfde
scriptmap gebruiken hebben dus hun eigen favorieten.


Voortgang

Tijdens het draaien loopt er een blauwe lijn onderin, met rechts ernaast wat
er gebeurt en hoe lang het al duurt.

    Het script meldt zelf een percentage (Write-Progress)  ->  exacte voortgang
    Het script draaide eerder al eens                      ->  schatting op basis
                                                               van de vorige keer
    Geen van beide                                         ->  bewegend lijntje

Wil je exacte voortgang in een eigen script, gebruik dan gewoon:

    Write-Progress -Activity "Bestanden lezen" -Status "$i van $totaal" `
                   -PercentComplete ($i / $totaal * 100)


Proefrun

Staat er een Proefrun-knopje, dan kan het script vertellen wat het zou doen
zonder het te doen. Gebruik dat bij scripts die iets veranderen (alleen-lezen
zetten, kopiëren, overzetten) voordat je het echt laat lopen, vooral als je
een pad met de hand hebt ingetypt.


Logbestanden

Elke run schrijft een logbestand in de map Logs naast dit bestand, met
tijdstempel en scriptnaam. Bovenaan staat de opdracht zoals je hem zelf in
PowerShell zou typen. Logs ouder dan 90 dagen worden bij het opstarten
opgeruimd; dat getal staat bovenin dit bestand bij LOG_BEWAARDAGEN.


Sneltoetsen

    Ctrl+Enter   uitvoeren
    Esc          afbreken
    F5           scripts opnieuw inlezen


Zelf een script toevoegen

Zet een .ps1-bestand in de map Scripts. Bij de volgende start staat het in de
lijst; met Herlaad ook meteen. Zorg voor een param()-blok met HelpMessage per
parameter, dan komen de labels en uitleg goed in beeld:

    <#
    .SYNOPSIS
    Korte omschrijving die in de lijst komt.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, HelpMessage = "De map waarin gezocht wordt")]
        [string]$Map,

        [Parameter(HelpMessage = "Aanvinken om terug te draaien")]
        [switch]$Terugdraaien
    )

Hoe het veldtype gekozen wordt:

    naam bevat bestand, file, csv, rapport of log   ->  bestandskiezer
    naam bevat map, pad, folder of directory        ->  mapkiezer
    [switch]                                        ->  aan/uit-knopje
    [ValidateSet("a","b")]                          ->  rij keuzeknoppen
    [string[]]                                      ->  invoervak, per regel
    getaltype                                       ->  getalveld

Heet een lijstveld Filter, dan komen er knopjes onder voor de gangbare
bestandstypen. De parameternaam wordt leesbaar gemaakt: MinimaleGrootteKB
staat in beeld als "Minimale grootte KB".

Wil je dat een knopje aan in beeld komt, zet dan "(standaard aan)" in de
HelpMessage van die parameter. Dat scheelt dubbele ontkenningen als NietOpenen:

    [Parameter(HelpMessage = "Rapport meteen openen (standaard aan)")]
    [switch]$Openen

De scripts blijven gewone PowerShell-bestanden. Je kunt ze los draaien in een
PowerShell-venster, en de runner wijzigt ze nooit.


Let op bij accenten: sla op als UTF-8 met BOM

PowerShell 5.1, de versie die standaard op Windows staat, leest een .ps1-bestand
als Windows-1252 tenzij er een byte order mark voor staat. Zonder BOM wordt
'kopieën' dan 'kopieÃ«n', en dat zie je pas terug in de uitvoer of in een
rapport, niet in je code.

In VS Code klik je rechtsonder op de codering en kies je "Save with encoding" ->
"UTF-8 with BOM". De runner controleert dit bij elke Herlaad en zet een
uitroepteken achter de scriptnaam als het misgaat.
"""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        maak_stijl(self)
        self.title("PowerShell-runner")
        pas_venstergrootte(self, px(1180), px(940))
        self.minsize(min(px(860), self.winfo_screenwidth() - 80),
                     min(px(620), self.winfo_screenheight() - 80))

        self.scripts: list[Script] = []
        self.huidig: Script | None = None
        self.velden: dict[str, object] = {}         # naam -> var of Text
        self.mapvelden: dict[str, tuple] = {}       # naam -> (widget, var)
        self.proefrun_var = tk.BooleanVar(value=False)
        self.berichten: queue.Queue = queue.Queue()
        self.proces: subprocess.Popen | None = None
        self.bezig = False
        self.afgebroken = False
        self.lopend_script = ""
        self.lopende_proefrun = False
        self.logbestand: Path | None = None
        self.rapporten: list[Path] = []
        self.status_var = tk.StringVar(value="")
        self._laatste_status = 0.0
        self._wacht_herlaad = False

        # voortgang van de lopende run
        self.start_tijd = 0.0
        self.verwacht: float | None = None
        self.vg_procent: int | None = None
        self.vg_tekst = ""
        self.vg_actief = False

        self.instellingen = laad_instellingen()
        for sleutel, leeg in (("favorieten", []), ("gebruik", {}), ("waarden", {}),
                              ("recente_mappen", [])):
            if not isinstance(self.instellingen.get(sleutel), type(leeg)):
                self.instellingen[sleutel] = leeg

        self.ster_vol = beeld_van(ster_png(px(16), KLEUR["ster"], True))
        self.ster_leeg = beeld_van(ster_png(px(16), KLEUR["duim"], False))
        self.ster_knop_vol = beeld_van(ster_png(px(15), KLEUR["ster"], True))
        self.ster_knop_leeg = beeld_van(ster_png(px(15), KLEUR["blauw"], False))

        self._bouw()
        self.protocol("WM_DELETE_WINDOW", self.sluiten)
        self.bind_all("<Control-Return>", lambda _e: self.start())
        self.bind_all("<Escape>", lambda _e: self.afbreken() if self.bezig else None)
        self.bind_all("<F5>", lambda _e: self.herlaad())

        self.sleepbaar = maak_sleepbaar(
            self, lambda paden: self.berichten.put(("neergezet", paden)))

        self.after(100, self.verwerk_berichten)

        if POWERSHELL is None:
            self.schrijf("PowerShell niet gevonden op deze computer. Dat is "
                         "ongebruikelijk op Windows.\n", "fout")
        opgeruimd = ruim_logs_op()
        if opgeruimd:
            self.schrijf(f"{opgeruimd} oude logbestanden opgeruimd.\n", "grijs")
        # Eerst het venster tonen, dan pas de scripts inlezen: dat bevragen
        # van PowerShell duurt een paar tellen.
        self.after(50, self.herlaad)

    def report_callback_exception(self, soort, waarde, spoor) -> None:
        """Een fout in een knop zichtbaar maken; een .pyw heeft geen console."""
        tekst = "".join(traceback.format_exception(soort, waarde, spoor))
        try:
            self.schrijf("\nEr ging iets mis in het venster:\n", "fout")
            self.schrijf(tekst, "grijs")
        except Exception:
            pass
        try:
            self.clipboard_clear()
            self.clipboard_append(tekst)
        except Exception:
            pass
        messagebox.showerror(
            "Er ging iets mis",
            f"{type(waarde).__name__}: {waarde}\n\n"
            "De volledige melding staat in het uitvoervak en op je klembord.")

    # ------------------------------------------------------------- opbouw
    #
    # Van boven naar beneden:
    #   kop          titel, Uitleg
    #   midden       Scripts | Instellingen van het gekozen script
    #   Uitvoer      het logboek (versleepbaar tussen midden en uitvoer)
    #   voortgang    dunne blauwe lijn
    #   actiebalk    Uitvoeren, of tijdens het werk Afbreken; rechts de status
    #
    # De actiebalk wordt als eerste neergezet (side="bottom"), zodat de knoppen
    # nooit uit beeld geduwd worden.

    def _bouw(self) -> None:
        rand = px(22)
        self._bouw_actiebalk(rand)

        self.balk = Voortgangslijn(self)
        self.balk.pack(side="bottom", fill="x", padx=rand, pady=(px(12), 0))

        self._bouw_kop(rand)

        verticaal = ttk.PanedWindow(self, orient="vertical")
        verticaal.pack(fill="both", expand=True, padx=rand, pady=(px(4), 0))

        midden = ttk.PanedWindow(verticaal, orient="horizontal")
        midden.add(self._bouw_scriptlijst(midden), weight=2)
        midden.add(self._bouw_formulier(midden), weight=5)
        verticaal.add(midden, weight=3)
        verticaal.add(self._bouw_uitvoer(verticaal), weight=1)
        self.verticaal = verticaal
        # Het formulier is waar je werkt: dat krijgt bij het openen ongeveer
        # twee derde van de hoogte. De schuif ertussen blijft te verslepen.
        self.after(120, self._verdeel_hoogte)

        self.werk_startknop_bij()

    def _verdeel_hoogte(self) -> None:
        try:
            hoogte = self.verticaal.winfo_height()
            if hoogte > 100:
                self.verticaal.sashpos(0, int(hoogte * 0.66))
        except tk.TclError:
            pass

    def _bouw_kop(self, rand: int) -> None:
        kop = ttk.Frame(self, padding=(rand, px(10), rand, px(8)))
        kop.pack(fill="x")
        ttk.Label(kop, text="PowerShell-runner", style="Titel.TLabel").pack(
            side="left", anchor="s")
        ttk.Label(kop, style="Ondertitel.TLabel",
                  text="Scripts draaien zonder in code te typen").pack(
            side="left", anchor="s", padx=(px(14), 0), pady=(0, px(5)))
        ttk.Button(kop, text="Uitleg", command=self.toon_uitleg,
                   style="Link.TButton").pack(side="right", anchor="s")
        ttk.Button(kop, text="Logmap", command=lambda: self.open_pad(LOGMAP),
                   style="Link.TButton").pack(side="right", anchor="s")

    def _bouw_actiebalk(self, rand: int) -> None:
        """De hoofdknop, en wat er naast hoort te staan.

        Eén blauwe knop die meeloopt met wat er kan: uit met een reden ernaast,
        aan met de naam van het script, tijdens het werk Afbreken ernaast, en
        na afloop Open rapport als het script een rapport noemde.
        """
        buiten = ttk.Frame(self, padding=(rand, px(12), rand, px(16)))
        buiten.pack(side="bottom", fill="x")
        buiten.columnconfigure(1, weight=1)
        balk = ttk.Frame(buiten)
        balk.grid(row=0, column=0, sticky="w")
        self.label_voortgang = ttk.Label(buiten, textvariable=self.status_var,
                                         style="Zacht.TLabel", anchor="e")
        self.label_voortgang.grid(row=0, column=1, sticky="e", padx=(px(14), 0))
        self.knop_start = ttk.Button(balk, text="Uitvoeren", command=self.start,
                                     style="Primair.TButton")
        self.knop_start.pack(side="left")
        self.knop_stop = ttk.Button(balk, text="Afbreken", command=self.afbreken)
        self.knop_rapport = ttk.Button(balk, text="Open rapport",
                                       command=self.open_rapport)
        self.knop_hint = ttk.Label(balk, text="", style="Zacht.TLabel")

    def _bouw_scriptlijst(self, ouder) -> ttk.Frame:
        k = KLEUR
        vak = kaart(ouder, padx=px(5), pady=px(5))

        kop = tk.Frame(vak, background=k["kaart"], padx=px(13), pady=px(9))
        kop.pack(fill="x")
        ttk.Label(kop, text="Scripts", style="Kop.Kaart.TLabel").pack(side="left")
        self.knop_herlaad = ttk.Button(kop, text="Herlaad", command=self.herlaad,
                                       style="Kaart.Link.TButton")
        self.knop_herlaad.pack(side="right")
        ttk.Button(kop, text="Map", command=lambda: self.open_pad(SCRIPTMAP),
                   style="Kaart.Link.TButton").pack(side="right")
        tk.Frame(vak, background=k["lijn"], height=1).pack(fill="x")

        onder = tk.Frame(vak, background=k["kaart"], padx=px(13), pady=px(8))
        onder.pack(side="bottom", fill="x")
        self.label_lijst = ttk.Label(onder, text="Scripts inlezen…",
                                     style="Zacht.Kaart.TLabel")
        self.label_lijst.pack(side="left")
        tk.Frame(vak, background=k["lijn"], height=1).pack(side="bottom", fill="x")

        binnen = tk.Frame(vak, background=k["kaart"])
        binnen.pack(fill="both", expand=True)
        self.lijst = ttk.Treeview(binnen, columns=("gebruik",), show="tree",
                                  selectmode="browse", height=6)
        self.lijst.column("#0", width=px(240), stretch=True)
        self.lijst.column("gebruik", width=px(46), anchor="e", stretch=False)
        self.lijst.tag_configure("kop", foreground=k["zacht"], font=letter(8, "bold"))
        self.lijst.tag_configure("probleem", foreground=k["oranje"])
        schuif = ttk.Scrollbar(binnen, command=self.lijst.yview)
        self.lijst.configure(yscrollcommand=schuif.set)
        schuif.pack(side="right", fill="y", padx=(px(2), px(2)))
        self.lijst.pack(side="left", fill="both", expand=True, padx=(px(4), 0))
        self.lijst.bind("<<TreeviewSelect>>", self.op_script_gekozen)
        self.lijst.bind("<Button-1>", self.op_klik_lijst)
        self.lijst.bind("<Double-Button-1>", lambda _e: self.start())
        return vak

    def _bouw_formulier(self, ouder) -> ttk.Frame:
        k = KLEUR
        vak = kaart(ouder, padx=px(5), pady=px(5))

        kop = tk.Frame(vak, background=k["kaart"], padx=px(13), pady=px(9))
        kop.pack(fill="x")
        kop.columnconfigure(0, weight=1)
        self.label_script = ttk.Label(kop, text="Kies links een script",
                                      style="Groot.Kaart.TLabel")
        self.label_script.grid(row=0, column=0, sticky="w")
        self.knop_favoriet = ttk.Button(kop, text=" Favoriet", compound="left",
                                        image=self.ster_knop_leeg,
                                        command=self.wissel_favoriet_huidig,
                                        style="Kaart.Link.TButton")
        self.label_omschrijving = ttk.Label(kop, text="", style="Zacht.Kaart.TLabel",
                                            justify="left")
        self.label_omschrijving.grid(row=1, column=0, columnspan=2, sticky="ew",
                                     pady=(px(2), 0))
        meebuigen(self.label_omschrijving)
        regel = tk.Frame(kop, background=k["kaart"])
        regel.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(px(4), 0))
        self.label_onthouden = ttk.Label(regel, text="", style="Licht.Kaart.TLabel")
        self.label_onthouden.pack(side="left")
        self.knop_standaard = ttk.Button(regel, text="Standaardwaarden",
                                         command=self.zet_standaardwaarden,
                                         style="Kaart.Link.TButton")
        tk.Frame(vak, background=k["lijn"], height=1).pack(fill="x")

        # onderaan: proefrun
        self.proefrun_vak = tk.Frame(vak, background=k["kaart"], padx=px(13), pady=px(9))
        self.proefrun_vak.pack(side="bottom", fill="x")
        self.knop_proefrun = Keuzeknop(self.proefrun_vak, "Proefrun", self.proefrun_var,
                                       bij_wissel=self.werk_startknop_bij)
        self.knop_proefrun.pack(side="left")
        ttk.Label(self.proefrun_vak, style="Zacht.Kaart.TLabel",
                  text="laat zien wat het script zou doen, zonder het te doen").pack(
            side="left", padx=(px(10), 0))
        self.proefrun_lijn = tk.Frame(vak, background=k["lijn"], height=1)
        self.proefrun_lijn.pack(side="bottom", fill="x")

        # de velden, met een schuifbalk
        binnen = tk.Frame(vak, background=k["kaart"])
        binnen.pack(fill="both", expand=True)
        self.doek = tk.Canvas(binnen, highlightthickness=0, borderwidth=0,
                              background=k["kaart"], height=px(120))
        schuif = ttk.Scrollbar(binnen, command=self.doek.yview)
        self.doek.configure(yscrollcommand=schuif.set)
        schuif.pack(side="right", fill="y", padx=(px(2), px(2)))
        self.doek.pack(side="left", fill="both", expand=True)
        self.formulier = tk.Frame(self.doek, background=k["kaart"], padx=px(13),
                                  pady=px(10))
        self._venster = self.doek.create_window((0, 0), window=self.formulier, anchor="nw")
        self.formulier.bind("<Configure>", lambda e: self.doek.configure(
            scrollregion=self.doek.bbox("all")))
        self.doek.bind("<Configure>", lambda e: self.doek.itemconfig(
            self._venster, width=e.width))
        self.bind_all("<MouseWheel>", self._wiel, add="+")

        self._zet_proefrun_zichtbaar(False)
        return vak

    def _bouw_uitvoer(self, ouder) -> ttk.Frame:
        k = KLEUR
        vak = kaart(ouder, padx=px(5), pady=px(5))
        kop = tk.Frame(vak, background=k["kaart"], padx=px(13), pady=px(7))
        kop.pack(fill="x")
        ttk.Label(kop, text="Uitvoer", style="Kop.Kaart.TLabel").pack(side="left")
        ttk.Button(kop, text="Wissen", command=self.leeg_log,
                   style="Kaart.Link.TButton").pack(side="right")
        ttk.Button(kop, text="Kopiëren", command=self.kopieer_log,
                   style="Kaart.Link.TButton").pack(side="right")
        self.knop_log = ttk.Button(kop, text="Open log", command=self.open_log,
                                   style="Kaart.Link.TButton", state="disabled")
        self.knop_log.pack(side="right")
        tk.Frame(vak, background=k["lijn"], height=1).pack(fill="x")

        binnen = tk.Frame(vak, background=k["kaart"])
        binnen.pack(fill="both", expand=True)
        self.log = tk.Text(binnen, height=6, wrap="none", relief="flat",
                           borderwidth=0, highlightthickness=0,
                           font=(LETTER_CODE, 9), background=k["kaart"],
                           foreground=k["tekst"], padx=px(13), pady=px(8),
                           spacing1=px(1), insertbackground=k["tekst"])
        vert = ttk.Scrollbar(binnen, command=self.log.yview)
        hor = ttk.Scrollbar(vak, orient="horizontal", command=self.log.xview)
        self.log.configure(yscrollcommand=vert.set, xscrollcommand=hor.set,
                           state="disabled")
        hor.pack(side="bottom", fill="x", padx=px(8))
        vert.pack(side="right", fill="y", padx=(px(2), px(2)))
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_configure("kop", font=(LETTER_CODE, 9, "bold"))
        self.log.tag_configure("ok", foreground=k["groen"])
        self.log.tag_configure("fout", foreground=k["rood"])
        self.log.tag_configure("let_op", foreground=k["oranje"])
        self.log.tag_configure("grijs", foreground=k["zacht"])
        return vak

    def _zet_proefrun_zichtbaar(self, zichtbaar: bool) -> None:
        if zichtbaar and not self.proefrun_vak.winfo_ismapped():
            self.proefrun_lijn.pack(side="bottom", fill="x", before=self.doek.master)
            self.proefrun_vak.pack(side="bottom", fill="x", before=self.proefrun_lijn)
        elif not zichtbaar:
            self.proefrun_vak.pack_forget()
            self.proefrun_lijn.pack_forget()

    def _wiel(self, e) -> None:
        """Het muiswiel schuift het formulier, maar alleen als de muis erboven
        staat. Boven de lijst of de uitvoer doen die zelf het werk."""
        try:
            w = self.winfo_containing(e.x_root, e.y_root)
        except (KeyError, tk.TclError):
            return
        while w is not None:
            if w is self.doek:
                if isinstance(self.winfo_containing(e.x_root, e.y_root), tk.Text):
                    return                     # een lijstveld schuift zelf
                try:
                    self.doek.yview_scroll(int(-1 * (e.delta / 120)), "units")
                except tk.TclError:
                    pass
                return
            w = getattr(w, "master", None)

    # --------------------------------------------------------------- log
    def schrijf(self, tekst: str, tag: str = "") -> None:
        self.log.configure(state="normal")
        self.log.insert("end", tekst, tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def leeg_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def kopieer_log(self) -> None:
        self.clipboard_clear()
        self.clipboard_append(self.log.get("1.0", "end").rstrip())
        self.status_var.set("Uitvoer gekopieerd naar het klembord.")

    def toon_uitleg(self) -> None:
        k = KLEUR
        venster = tk.Toplevel(self, background=k["achtergrond"])
        venster.title("Uitleg")
        pas_venstergrootte(venster, px(820), px(760))
        vak = kaart(venster, padx=px(6), pady=px(6))
        vak.pack(fill="both", expand=True, padx=px(18), pady=(px(18), px(8)))
        binnen = tk.Frame(vak, background=k["kaart"])
        binnen.pack(fill="both", expand=True)
        tekst = tk.Text(binnen, wrap="word", relief="flat", borderwidth=0,
                        highlightthickness=0, font=letter(10),
                        background=k["kaart"], foreground=k["tekst"],
                        padx=px(18), pady=px(14), spacing1=px(1), spacing3=px(2))
        schuif = ttk.Scrollbar(binnen, command=tekst.yview)
        tekst.configure(yscrollcommand=schuif.set)
        schuif.pack(side="right", fill="y", padx=(0, px(2)))
        tekst.pack(side="left", fill="both", expand=True)
        tekst.tag_configure("kop", font=letter(12, "bold"), spacing1=px(10),
                            spacing3=px(4))
        tekst.tag_configure("code", font=(LETTER_CODE, 9), foreground=k["zacht"])

        # Een kopje is een korte regel met een witregel ervoor en erna.
        regels = UITLEG.splitlines()
        for i, regel in enumerate(regels):
            voor = i == 0 or not regels[i - 1].strip()
            na = i + 1 < len(regels) and not regels[i + 1].strip()
            if (regel.strip() and not regel.startswith(" ") and voor and na
                    and len(regel) < 60 and not regel.rstrip().endswith((".", ":"))):
                tag = "kop"
            elif regel.startswith("    "):
                tag = "code"
            else:
                tag = ""
            tekst.insert("end", regel + "\n", tag)
        tekst.configure(state="disabled")
        onder = ttk.Frame(venster, padding=(px(18), 0, px(18), px(12)))
        onder.pack(fill="x")
        ttk.Button(onder, text="Sluiten", command=venster.destroy,
                   style="Primair.TButton").pack(side="right")

    # ---------------------------------------------------- onthouden
    def bewaar(self) -> None:
        bewaar_instellingen(self.instellingen)

    def is_favoriet(self, script: Script) -> bool:
        return script.naam in self.instellingen["favorieten"]

    def wissel_favoriet(self, script: Script) -> None:
        favorieten = self.instellingen["favorieten"]
        if script.naam in favorieten:
            favorieten.remove(script.naam)
        else:
            favorieten.append(script.naam)
        self.bewaar()
        self.vul_lijst()
        self.werk_favorietknop_bij()

    def wissel_favoriet_huidig(self) -> None:
        if self.huidig is not None:
            self.wissel_favoriet(self.huidig)

    def werk_favorietknop_bij(self) -> None:
        if self.huidig is None or self.huidig.fout:
            self.knop_favoriet.grid_remove()
            return
        if self.is_favoriet(self.huidig):
            self.knop_favoriet.configure(image=self.ster_knop_vol, text=" Favoriet")
        else:
            self.knop_favoriet.configure(image=self.ster_knop_leeg,
                                         text=" Maak favoriet")
        self.knop_favoriet.grid(row=0, column=1, sticky="e")

    def onthoud_mappen(self, mappen: list[str]) -> None:
        recent = self.instellingen["recente_mappen"]
        for map_ in reversed(mappen):
            for bestaand in [r for r in recent if r.lower() == map_.lower()]:
                recent.remove(bestaand)
            recent.insert(0, map_)
        del recent[MAX_RECENTE_MAPPEN:]

    # ------------------------------------------------------------ scripts
    def herlaad(self) -> None:
        """Scripts opnieuw inlezen, in een thread zodat het venster vrij blijft."""
        if self._wacht_herlaad:
            return
        self._wacht_herlaad = True
        self.knop_herlaad.configure(state="disabled")
        self.label_lijst.configure(text="Scripts inlezen…")
        SCRIPTMAP.mkdir(parents=True, exist_ok=True)

        def lees():
            try:
                scripts = lees_scripts(SCRIPTMAP)
            except Exception as e:                       # pragma: no cover
                scripts = []
                self.berichten.put(("regel_los", f"Scripts inlezen mislukt: {e}", "fout"))
            self.berichten.put(("ingelezen", scripts))

        threading.Thread(target=lees, daemon=True).start()

    def scripts_ingelezen(self, scripts: list[Script]) -> None:
        self._wacht_herlaad = False
        self.knop_herlaad.configure(state="normal")
        vorige = self.huidig.naam if self.huidig else self.instellingen.get("laatste_script")
        self.scripts = scripts
        self.vul_lijst()

        if not scripts:
            self.schrijf(f"Geen .ps1-scripts gevonden in {SCRIPTMAP}\n", "fout")
        for script in scripts:
            if script.waarschuwing:
                self.schrijf(f"{script.naam}: {script.waarschuwing}\n", "let_op")
            if script.fout:
                self.schrijf(f"{script.naam}: kon niet gelezen worden: {script.fout}\n",
                             "fout")

        doel = next((s for s in scripts if s.naam == vorige), None)
        if doel is not None:
            self.lijst.selection_set(str(doel.pad))
            self.lijst.see(str(doel.pad))
        else:
            self.huidig = None
            self.bouw_velden()
        self.werk_startknop_bij()

    def vul_lijst(self) -> None:
        """De lijst opbouwen: eerst de favorieten, dan de rest.

        Voor elk script een sterretje: oranje en vol bij een favoriet, een
        grijze omtrek bij de rest. Klikken op het sterretje zet hem aan of
        uit. Rechts hoe vaak je het script gebruikt hebt.
        """
        gekozen = self.lijst.selection()
        self.lijst.delete(*self.lijst.get_children())
        favorieten = [s for s in self.scripts if self.is_favoriet(s)]
        overig = [s for s in self.scripts if not self.is_favoriet(s)]

        def voeg_toe(script: Script) -> None:
            gebruik = self.instellingen["gebruik"].get(script.naam, {})
            aantal = gebruik.get("aantal", 0)
            probleem = bool(script.fout or script.waarschuwing)
            self.lijst.insert(
                "", "end", iid=str(script.pad),
                text=" " + script.naam + ("  !" if probleem else ""),
                image=self.ster_vol if self.is_favoriet(script) else self.ster_leeg,
                values=(f"{aantal}×" if aantal else "",),
                tags=("probleem",) if probleem else ())

        if favorieten:
            self.lijst.insert("", "end", iid="#favorieten", text="FAVORIETEN", tags=("kop",))
            for script in favorieten:
                voeg_toe(script)
            if overig:
                self.lijst.insert("", "end", iid="#overig", text="ALLE SCRIPTS",
                                  tags=("kop",))
        for script in overig:
            voeg_toe(script)

        for iid in gekozen:
            if self.lijst.exists(iid):
                self.lijst.selection_set(iid)

        if favorieten:
            tekst = f"{len(self.scripts)} scripts  ·  {len(favorieten)} favoriet"
        else:
            tekst = f"{len(self.scripts)} scripts  ·  ster = favoriet"
        self.label_lijst.configure(text=tekst)

    def op_klik_lijst(self, e) -> str | None:
        """Klik op het sterretje: favoriet aan of uit, zonder te selecteren."""
        rij = self.lijst.identify_row(e.y)
        if not rij or rij.startswith("#"):
            return "break" if rij else None
        if self.lijst.identify_element(e.x, e.y) == "Treeitem.image":
            script = next((s for s in self.scripts if str(s.pad) == rij), None)
            if script is not None:
                self.wissel_favoriet(script)
            return "break"
        return None

    def op_script_gekozen(self, *_) -> None:
        keuze = self.lijst.selection()
        if not keuze:
            return
        if keuze[0].startswith("#"):                     # een kopje
            self.lijst.selection_remove(keuze[0])
            return
        nieuw = next((s for s in self.scripts if str(s.pad) == keuze[0]), None)
        if nieuw is self.huidig and self.velden:
            return
        self.huidig = nieuw
        if nieuw is not None:
            self.instellingen["laatste_script"] = nieuw.naam
        self.bouw_velden()

    # ------------------------------------------------------------ formulier
    def bouw_velden(self, standaard: bool = False) -> None:
        """Het formulier voor het gekozen script.

        Velden krijgen de waarden van de vorige keer, tenzij standaard=True
        of het script nog nooit gedraaid heeft: dan wat het script voorstelt.
        """
        for widget in self.formulier.winfo_children():
            widget.destroy()
        self.velden.clear()
        self.mapvelden.clear()
        self.doek.yview_moveto(0)

        script = self.huidig
        self.werk_favorietknop_bij()
        self.knop_standaard.pack_forget()
        self.label_onthouden.configure(text="")
        if script is None:
            self.label_script.configure(text="Kies links een script")
            self.label_omschrijving.configure(text="")
            self._zet_proefrun_zichtbaar(False)
            self.werk_startknop_bij()
            return

        self.label_script.configure(text=script.naam)
        self.label_omschrijving.configure(text=script.omschrijving or "(geen omschrijving)")

        if script.fout:
            ttk.Label(self.formulier, style="Fout.Kaart.TLabel", wraplength=px(460),
                      text=f"Dit script kon niet gelezen worden:\n{script.fout}").pack(
                anchor="w")
            self._zet_proefrun_zichtbaar(False)
            self.werk_startknop_bij()
            return

        if script.waarschuwing:
            noot = ttk.Label(self.formulier, style="Fout.Kaart.TLabel", justify="left",
                             text=f"Let op: dit script {script.waarschuwing}")
            noot.pack(anchor="w", fill="x", pady=(0, px(12)))
            meebuigen(noot)

        if not script.parameters:
            ttk.Label(self.formulier, style="Zacht.Kaart.TLabel",
                      text="Dit script heeft geen instellingen; je kunt het direct "
                           "uitvoeren.").pack(anchor="w")

        bewaard = {} if standaard else self.instellingen["waarden"].get(script.naam, {})
        gebruik = self.instellingen["gebruik"].get(script.naam, {})
        if bewaard and gebruik.get("laatst"):
            self.label_onthouden.configure(
                text=f"Ingevuld zoals de vorige keer ({hoe_lang_geleden(gebruik['laatst'])})")
        elif standaard:
            self.label_onthouden.configure(text="Standaardwaarden van het script")
        if self.instellingen["waarden"].get(script.naam) and script.parameters:
            self.knop_standaard.pack(side="left", padx=(px(4), 0))

        # Eerst de verplichte velden, dan de rest; binnen die groepen de
        # volgorde van het script.
        volgorde = ([p for p in script.parameters if p.verplicht and p.soort != "vinkje"]
                    + [p for p in script.parameters
                       if not (p.verplicht and p.soort != "vinkje") and p.soort != "vinkje"])
        knopjes = [p for p in script.parameters if p.soort == "vinkje"]
        for parameter in volgorde:
            self._bouw_veld(parameter, bewaard)
        if knopjes:
            self._bouw_knopjes(knopjes, bewaard)

        self._zet_proefrun_zichtbaar(script.proefrun_mogelijk)
        if not script.proefrun_mogelijk:
            self.proefrun_var.set(False)
        self.werk_startknop_bij()

    def zet_standaardwaarden(self) -> None:
        self.bouw_velden(standaard=True)

    def _veldkop(self, ouder, parameter: Parameter) -> None:
        regel = tk.Frame(ouder, background=KLEUR["kaart"])
        regel.pack(fill="x")
        ttk.Label(regel, text=parameter.label, style="Naam.Kaart.TLabel").pack(side="left")
        if parameter.verplicht:
            ttk.Label(regel, text="verplicht", style="Verplicht.Kaart.TLabel").pack(
                side="left", padx=(px(8), 0), pady=(px(2), 0))
        if parameter.uitleg:
            uitleg = ttk.Label(ouder, text=parameter.uitleg, style="Zacht.Kaart.TLabel",
                               justify="left")
            uitleg.pack(anchor="w", fill="x")
            meebuigen(uitleg)

    def _bouw_veld(self, parameter: Parameter, bewaard: dict) -> None:
        k = KLEUR
        blok = tk.Frame(self.formulier, background=k["kaart"])
        blok.pack(fill="x", pady=(0, px(14)))
        self._veldkop(blok, parameter)
        waarde = bewaard.get(parameter.naam)

        if parameter.soort == "keuze":
            begin = waarde if waarde in parameter.keuzes else (
                parameter.standaard if parameter.standaard in parameter.keuzes
                else (parameter.keuzes[0] if parameter.keuzes else ""))
            var = tk.StringVar(value=begin)
            if len(parameter.keuzes) <= MAX_KEUZEKNOPPEN:
                Keuzerij(blok, parameter.keuzes, var).pack(anchor="w", pady=(px(4), 0))
            else:
                ttk.Combobox(blok, textvariable=var, values=parameter.keuzes,
                             state="readonly", width=28, font=letter(10)).pack(
                    anchor="w", pady=(px(4), 0))
            self.velden[parameter.naam] = var
            return

        if parameter.soort == "lijst":
            if isinstance(waarde, str):
                begin = waarde.strip()
            else:
                begin = "\n".join(d.strip() for d in parameter.standaard.split(",")
                                  if d.strip())
            vak = tk.Text(blok, height=4, wrap="none", font=(LETTER_CODE, 9),
                          relief="flat", borderwidth=0, background=k["kaart"],
                          foreground=k["tekst"], insertbackground=k["tekst"],
                          highlightthickness=1, highlightbackground=k["rand"],
                          highlightcolor=k["blauw"], padx=px(8), pady=px(6),
                          undo=True)
            vak.pack(fill="x", pady=(px(4), 0))
            if begin:
                vak.insert("1.0", begin)
            vak.bind("<KeyRelease>", lambda _e: self.werk_startknop_bij(), add="+")
            self.velden[parameter.naam] = vak
            if "filter" in parameter.naam.lower():
                self._bouw_snelfilters(blok, vak)
            else:
                ttk.Label(blok, text="één waarde per regel",
                          style="Licht.Kaart.TLabel").pack(anchor="w", pady=(px(2), 0))
            return

        rij = tk.Frame(blok, background=k["kaart"])
        rij.pack(fill="x", pady=(px(4), 0))
        begin = str(waarde) if isinstance(waarde, (str, int, float)) and not isinstance(
            waarde, bool) else parameter.standaard
        var = tk.StringVar(value=begin)
        if parameter.soort == "map":
            veld = ttk.Combobox(rij, textvariable=var, font=letter(10),
                                values=self.instellingen["recente_mappen"],
                                postcommand=lambda: veld.configure(
                                    values=self.instellingen["recente_mappen"]))
            veld.pack(side="left", fill="x", expand=True)
            ttk.Button(rij, text="Kies…", style="Kaart.TButton",
                       command=lambda v=var, p=parameter: self.kies_map(v, p)).pack(
                side="left", padx=(px(8), 0))
            self.mapvelden[parameter.naam] = (veld, var)
        elif parameter.soort == "bestand":
            veld = ttk.Entry(rij, textvariable=var, font=letter(10))
            veld.pack(side="left", fill="x", expand=True)
            ttk.Button(rij, text="Kies…", style="Kaart.TButton",
                       command=lambda v=var, p=parameter: self.kies_bestand(v, p)).pack(
                side="left", padx=(px(8), 0))
        elif parameter.soort == "getal":
            veld = ttk.Entry(rij, textvariable=var, font=letter(10), width=10,
                             justify="right")
            veld.pack(side="left")
            if parameter.standaard:
                ttk.Label(rij, text=f"standaard {parameter.standaard}",
                          style="Licht.Kaart.TLabel").pack(side="left", padx=(px(10), 0))
        else:
            veld = ttk.Entry(rij, textvariable=var, font=letter(10))
            veld.pack(side="left", fill="x", expand=True)
        veld.bind("<Return>", lambda _e: self.start())
        var.trace_add("write", lambda *_: self.werk_startknop_bij())
        self.velden[parameter.naam] = var

    def _bouw_snelfilters(self, blok, vak: tk.Text) -> None:
        """Knopjes voor gangbare bestandstypen onder een filterveld.

        Aanklikken zet het type in het vak, nog eens klikken haalt het weg.
        Wat je zelf typt blijft staan, en de knopjes volgen wat er staat.
        """
        rij = tk.Frame(blok, background=KLEUR["kaart"])
        rij.pack(fill="x", pady=(px(6), 0))
        knoppen: dict[str, tk.BooleanVar] = {}

        def regels() -> list[str]:
            return [r.strip() for r in vak.get("1.0", "end").splitlines() if r.strip()]

        def volg_tekst(_e=None) -> None:
            aanwezig = {r.lower() for r in regels()}
            for filter_, var in knoppen.items():
                if var.get() != (filter_.lower() in aanwezig):
                    var.set(filter_.lower() in aanwezig)

        def wissel(filter_: str) -> None:
            huidig = regels()
            if knoppen[filter_].get():
                if filter_ == "*.*":
                    huidig = []                 # alles: de rest heeft geen zin
                else:
                    huidig = [r for r in huidig if r != "*.*"]
                if filter_.lower() not in {r.lower() for r in huidig}:
                    huidig.append(filter_)
            else:
                huidig = [r for r in huidig if r.lower() != filter_.lower()]
            vak.delete("1.0", "end")
            vak.insert("1.0", "\n".join(huidig))
            volg_tekst()
            self.werk_startknop_bij()

        for filter_ in SNELFILTERS:
            var = tk.BooleanVar(value=False)
            knoppen[filter_] = var
            Keuzeknop(rij, filter_, var, bij_wissel=lambda f=filter_: wissel(f)).pack(
                side="left", padx=(0, px(6)), pady=(0, px(4)))
        vak.bind("<KeyRelease>", volg_tekst, add="+")
        volg_tekst()
        ttk.Label(blok, text="klik een type aan, of typ zelf: één per regel",
                  style="Licht.Kaart.TLabel").pack(anchor="w")

    def _bouw_knopjes(self, parameters: list[Parameter], bewaard: dict) -> None:
        """Alle aan/uit-schakelaars samen onder één kopje.

        Elk knopje met de uitleg er direct naast, zodat je niet hoeft te
        zoeken welk stukje tekst bij welk knopje hoort.
        """
        k = KLEUR
        blok = tk.Frame(self.formulier, background=k["kaart"])
        blok.pack(fill="x", pady=(px(2), px(6)))
        ttk.Label(blok, text="Opties", style="Naam.Kaart.TLabel").pack(anchor="w")
        raster = tk.Frame(blok, background=k["kaart"])
        raster.pack(fill="x", pady=(px(6), 0))
        raster.columnconfigure(1, weight=1)
        for rijnr, parameter in enumerate(parameters):
            if parameter.naam in bewaard:
                aan = bool(bewaard[parameter.naam])
            else:
                # Staat "(standaard aan)" in de HelpMessage van het script, dan
                # komt het knopje aan in beeld.
                aan = "standaard aan" in (parameter.uitleg or "").lower()
            var = tk.BooleanVar(value=aan)
            Keuzeknop(raster, parameter.label, var).grid(
                row=rijnr, column=0, sticky="w", pady=(0, px(6)))
            uitleg = re.sub(r"\s*\(standaard aan\)", "", parameter.uitleg or "").strip()
            if uitleg:
                label = ttk.Label(raster, text=uitleg, style="Zacht.Kaart.TLabel",
                                  justify="left")
                label.grid(row=rijnr, column=1, sticky="ew", padx=(px(12), 0),
                           pady=(0, px(6)))
                meebuigen(label)
            self.velden[parameter.naam] = var

    def kies_map(self, var: tk.StringVar, parameter: Parameter) -> None:
        huidig = var.get().strip()
        if not (huidig and Path(huidig).is_dir()) and self.instellingen["recente_mappen"]:
            huidig = self.instellingen["recente_mappen"][0]
        gekozen = filedialog.askdirectory(
            title=f"Kies: {parameter.label}",
            initialdir=huidig if huidig and Path(huidig).is_dir() else None)
        if gekozen:
            var.set(str(Path(gekozen)))

    def kies_bestand(self, var: tk.StringVar, parameter: Parameter) -> None:
        """Bestandskiezer met het juiste bestandstype erbij.

        De extensie volgt uit de parameternaam: een veld dat RapportHtml heet
        moet geen .csv voorstellen.
        """
        laag = parameter.naam.lower()
        if "html" in laag:
            extensie, soorten = ".html", [("HTML-pagina", "*.html;*.htm")]
        elif "csv" in laag:
            extensie, soorten = ".csv", [("CSV", "*.csv")]
        elif "xlsx" in laag or "excel" in laag:
            extensie, soorten = ".xlsx", [("Excel", "*.xlsx")]
        elif "log" in laag:
            extensie, soorten = ".log", [("Logbestand", "*.log;*.txt")]
        else:
            extensie, soorten = "", []
        soorten.append(("Alle bestanden", "*.*"))

        # Rapportvelden verwachten een bestand dat nog niet bestaat.
        opslaan = any(w in laag for w in ("rapport", "csv", "html", "xlsx", "log", "uit"))
        huidig = var.get().strip()
        beginmap = str(Path(huidig).parent) if huidig and Path(huidig).parent.is_dir() else None

        if opslaan:
            gekozen = filedialog.asksaveasfilename(
                title=f"Waar komt {parameter.label}?", defaultextension=extensie,
                filetypes=soorten, initialdir=beginmap,
                initialfile=Path(huidig).name if huidig else "")
        else:
            gekozen = filedialog.askopenfilename(title=f"Kies: {parameter.label}",
                                                 filetypes=soorten, initialdir=beginmap)
        if gekozen:
            var.set(str(Path(gekozen)))

    def zet_neergezet(self, paden: list[str]) -> None:
        """Een map uit de Verkenner op het venster gezet."""
        mappen = [p for p in paden if Path(p).is_dir()]
        if not mappen:
            self.status_var.set("Sleep een map op het venster, geen los bestand.")
            return
        if not self.mapvelden:
            self.status_var.set("Kies eerst een script met een mapveld.")
            return
        focus = self.focus_get()
        namen = list(self.mapvelden)
        doelen = [n for n in namen if self.mapvelden[n][0] is focus]
        doelen += [n for n in namen if not self.mapvelden[n][1].get().strip()]
        doelen += namen
        # Meer mappen tegelijk: vul de velden op volgorde.
        gebruikt = []
        for map_ in mappen:
            doel = next((n for n in doelen if n not in gebruikt), None)
            if doel is None:
                break
            self.mapvelden[doel][1].set(str(Path(map_)))
            gebruikt.append(doel)
        if self.huidig:
            labels = {p.naam: p.label for p in self.huidig.parameters}
            self.status_var.set("Ingevuld bij: " + ", ".join(labels.get(n, n) for n in gebruikt))

    def huidige_waarden(self) -> dict[str, object]:
        waarden: dict[str, object] = {}
        for naam, widget in self.velden.items():
            try:
                if isinstance(widget, tk.Text):
                    waarden[naam] = widget.get("1.0", "end").strip()
                else:
                    waarden[naam] = widget.get()
            except tk.TclError:
                pass
        return waarden

    # ----------------------------------------------------------- knoppen
    def werk_startknop_bij(self) -> None:
        """De actiebalk laten kloppen met wat er nu kan."""
        if not hasattr(self, "knop_start"):
            return
        if self.bezig:
            reden = "bezig…"
        elif POWERSHELL is None:
            reden = "PowerShell niet gevonden"
        elif self.huidig is None:
            reden = "kies links een script"
        elif self.huidig.fout:
            reden = "dit script kan niet gelezen worden"
        else:
            waarden = self.huidige_waarden()
            ontbreekt = ontbrekende_verplichte(self.huidig, waarden)
            fout = foute_getallen(self.huidig, waarden)
            if ontbreekt:
                reden = "vul nog in: " + ", ".join(p.label for p in ontbreekt)
            elif fout:
                reden = "geen getal: " + ", ".join(p.label for p in fout)
            else:
                reden = ""

        proef = self.proefrun_var.get() and self.huidig is not None \
            and self.huidig.proefrun_mogelijk
        if self.bezig:
            tekst = "Bezig…"
        elif proef:
            tekst = "Proefrun starten"
        else:
            tekst = "Uitvoeren"
        self.knop_start.configure(state="disabled" if reden else "normal", text=tekst)

        for knop in (self.knop_stop, self.knop_rapport, self.knop_hint):
            knop.pack_forget()
        if self.bezig:
            self.knop_stop.configure(state="disabled" if self.afgebroken else "normal")
            self.knop_stop.pack(side="left", padx=(px(8), 0))
        else:
            if self.rapporten:
                self.knop_rapport.pack(side="left", padx=(px(8), 0))
            if reden:
                self.knop_hint.configure(text=reden)
                self.knop_hint.pack(side="left", padx=(px(14), 0))

    # ------------------------------------------------------------ draaien
    def start(self) -> None:
        if self.bezig or self.huidig is None or self.huidig.fout:
            return
        script = self.huidig
        waarden = self.huidige_waarden()
        ontbreekt = ontbrekende_verplichte(script, waarden)
        if ontbreekt:
            messagebox.showwarning("Nog niet compleet",
                                   "Deze velden zijn verplicht:\n\n"
                                   + "\n".join(p.label for p in ontbreekt))
            return
        if foute_getallen(script, waarden):
            self.werk_startknop_bij()
            return

        proefrun = bool(self.proefrun_var.get() and script.proefrun_mogelijk)
        invoer = bouw_invoer(script, waarden, proefrun)
        try:
            opdracht, omgeving = bouw_uitvoering(script, invoer)
        except RuntimeError as e:
            messagebox.showerror("Kan niet starten", str(e))
            return

        # Onthouden: wat er is ingevuld, welke mappen, en hoe vaak.
        self.instellingen["waarden"][script.naam] = waarden
        self.onthoud_mappen([str(waarden.get(p.naam) or "").strip()
                             for p in script.parameters
                             if p.soort == "map" and str(waarden.get(p.naam) or "").strip()
                             and Path(str(waarden[p.naam]).strip()).is_dir()])
        gebruik = self.instellingen["gebruik"].setdefault(script.naam, {})
        gebruik["aantal"] = int(gebruik.get("aantal", 0)) + 1
        gebruik["laatst"] = datetime.now().isoformat(timespec="seconds")
        self.instellingen["laatste_script"] = script.naam
        self.bewaar()
        self.vul_lijst()
        self.label_onthouden.configure(text="Deze waarden worden onthouden")
        self.knop_standaard.pack(side="left", padx=(px(4), 0))

        self.leeg_log()
        self.bezig = True
        self.afgebroken = False
        self.lopend_script = script.naam
        self.lopende_proefrun = proefrun
        self.rapporten = []
        self.knop_log.configure(state="disabled")
        self.start_tijd = time.monotonic()
        duur = gebruik.get("duur")
        self.verwacht = float(duur) if isinstance(duur, (int, float)) and duur >= 3 \
            and not proefrun else None
        self.vg_procent = None
        self.vg_tekst = ""
        self.vg_actief = False
        self.balk.leeg()
        self.werk_voortgang_bij()
        self.werk_startknop_bij()

        LOGMAP.mkdir(parents=True, exist_ok=True)
        self.logbestand = logpad(script.naam)

        self.schrijf(f"{script.naam}\n", "kop")
        self.schrijf(f"gestart {datetime.now():%d-%m-%Y %H:%M:%S}"
                     + ("  ·  PROEFRUN, er wordt niets veranderd" if proefrun else "")
                     + "\n\n", "grijs")

        kop = (f"{script.naam}\ngestart {datetime.now():%Y-%m-%d %H:%M:%S}\n"
               f"proefrun: {'ja' if proefrun else 'nee'}\n"
               f"zelf draaien: {opdracht_als_tekst(script, invoer)}\n\n")
        threading.Thread(target=self.werk, args=(opdracht, omgeving, self.logbestand, kop),
                         daemon=True).start()

    def werk(self, opdracht: list[str], omgeving: dict, log: Path, kop: str) -> None:
        """Draait in een eigen thread; praat alleen via de wachtrij met het venster."""
        code = -1
        try:
            with open(log, "w", encoding="utf-8") as bestand:
                bestand.write(kop)
                bestand.flush()

                self.proces = subprocess.Popen(
                    opdracht, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    stdin=subprocess.DEVNULL, text=True, encoding="utf-8",
                    errors="replace", bufsize=1, env=omgeving,
                    creationflags=GEEN_VENSTER)

                for regel in self.proces.stdout:
                    schoon = _ANSI.sub("", regel.rstrip("\r\n"))
                    if is_clixml(schoon):
                        continue
                    voortgang = lees_voortgang(schoon)
                    if voortgang is not None:
                        self.berichten.put(("voortgang", voortgang))
                        continue
                    bestand.write(schoon + "\n")
                    self.berichten.put(("regel", schoon))
                self.proces.wait()
                code = self.proces.returncode
                bestand.write(f"\nafgesloten met code {code}\n")
            self.berichten.put(("klaar", code, str(log)))
        except Exception as e:
            self.berichten.put(("klaar", -1, str(log), f"{type(e).__name__}: {e}"))
        finally:
            self.proces = None

    def afbreken(self) -> None:
        proces = self.proces
        if proces is None:
            return
        self.afgebroken = True
        self.werk_startknop_bij()
        self.schrijf("\nAfbreken…\n", "grijs")
        try:
            # /T pakt ook de kindprocessen, bijvoorbeeld robocopy.
            subprocess.run(["taskkill", "/T", "/F", "/PID", str(proces.pid)],
                           capture_output=True, creationflags=GEEN_VENSTER)
        except Exception:
            try:
                proces.kill()
            except Exception:
                pass

    def verwerk_berichten(self) -> None:
        try:
            for _ in range(500):                 # niet eindeloos: venster moet bijtekenen
                bericht = self.berichten.get_nowait()
                soort = bericht[0]
                if soort == "regel":
                    self.toon_regel(bericht[1])
                elif soort == "regel_los":
                    self.schrijf(bericht[1] + "\n", bericht[2])
                elif soort == "voortgang":
                    self.verwerk_voortgang(bericht[1])
                elif soort == "klaar":
                    self.rond_af(*bericht[1:])
                elif soort == "ingelezen":
                    self.scripts_ingelezen(bericht[1])
                elif soort == "neergezet":
                    self.zet_neergezet(bericht[1])
        except queue.Empty:
            pass
        if self.bezig and time.monotonic() - self._laatste_status > 0.25:
            self.werk_voortgang_bij()
        self.after(60, self.verwerk_berichten)

    def toon_regel(self, regel: str) -> None:
        laag = regel.lower()
        tag = ""
        if any(w in laag for w in ("mislukt", "fout", "error", "bestaat niet",
                                   "not found", "exception", "kan niet")):
            tag = "fout"
        elif laag.startswith(("let op", "waarschuwing", "warning", "what if", "wat als")):
            tag = "let_op"
        elif laag.startswith(("klaar", "gereed")) or "opgeslagen" in laag:
            tag = "ok"
        self.schrijf(regel + "\n", tag)
        for pad in vind_rapporten(regel):
            if pad not in self.rapporten:
                self.rapporten.append(pad)

    # ----------------------------------------------------------- voortgang
    def verwerk_voortgang(self, v: dict) -> None:
        if v["klaar"]:
            self.vg_actief = False
            self.vg_procent = None
            self.vg_tekst = ""
        else:
            self.vg_actief = True
            self.vg_procent = v["procent"] if 0 <= v["procent"] <= 100 else None
            self.vg_tekst = "  ·  ".join(d for d in (v["activiteit"], v["status"]) if d)
        self.werk_voortgang_bij()

    def werk_voortgang_bij(self) -> None:
        """Balk en statusregel laten zien hoe ver het is.

        Exact als het script een percentage meldt; anders een schatting op
        basis van de vorige keer; anders een bewegend lijntje.
        """
        self._laatste_status = time.monotonic()
        verstreken = time.monotonic() - self.start_tijd
        procent = None
        if self.vg_actief and self.vg_procent is not None:
            procent = self.vg_procent
            self.balk.zet(procent / 100)
            tekst = f"{self.vg_tekst}  ·  {procent}%" if self.vg_tekst else f"{procent}%"
        elif self.vg_actief:
            self.balk.lopend()
            tekst = self.vg_tekst or "bezig"
        elif self.verwacht:
            deel = verstreken / self.verwacht
            if deel < 0.97:
                procent = int(deel * 100)
                self.balk.zet(deel)
                tekst = f"nog ongeveer {klok(self.verwacht - verstreken)}"
            else:
                self.balk.lopend()
                tekst = f"duurt langer dan vorige keer ({klok(self.verwacht)})"
        else:
            self.balk.lopend()
            tekst = "bezig"
        if len(tekst) > 90:
            tekst = tekst[:87] + "…"
        self.status_var.set(f"{tekst}  ·  {klok(verstreken)}")
        naam = self.lopend_script
        self.title(f"{procent}% · {naam} · PowerShell-runner" if procent is not None
                   else f"{naam} loopt · PowerShell-runner")

    def rond_af(self, code: int, log: str, fout: str = "") -> None:
        self.bezig = False
        verstreken = time.monotonic() - self.start_tijd
        self.knop_log.configure(state="normal")
        self.logbestand = Path(log)
        self.title("PowerShell-runner")

        naam = self.lopend_script
        if fout:
            self.schrijf(f"\nKon het script niet draaien: {fout}\n", "fout")
            self.status_var.set("Niet gelukt")
            self.balk.leeg()
        elif self.afgebroken:
            self.schrijf(f"\nAfgebroken na {klok(verstreken)}.\n", "let_op")
            self.status_var.set(f"Afgebroken na {klok(verstreken)}")
            self.balk.leeg()
        elif code == 0:
            self.schrijf(f"\nKlaar zonder fouten, in {klok(verstreken)}.\n", "ok")
            self.status_var.set(f"Klaar in {klok(verstreken)}")
            self.balk.vol()
            self.after(2500, lambda: None if self.bezig else self.balk.leeg())
            # Looptijd onthouden voor de schatting van de volgende keer; een
            # proefrun telt niet mee, die doet het echte werk niet.
            if not self.lopende_proefrun and naam:
                gebruik = self.instellingen["gebruik"].setdefault(naam, {})
                gebruik["duur"] = round(verstreken, 1)
                self.bewaar()
        else:
            self.schrijf(f"\nAfgesloten met code {code}, na {klok(verstreken)}. "
                         "Kijk hierboven welke regels rood zijn.\n", "fout")
            self.status_var.set(f"Afgesloten met fouten (code {code})")
            self.balk.leeg()
        self.schrijf(f"Log: {log}\n", "grijs")
        self.werk_startknop_bij()

    # ------------------------------------------------------------ overig
    def open_log(self) -> None:
        if self.logbestand:
            self.open_pad(self.logbestand)

    def open_rapport(self) -> None:
        if self.rapporten:
            self.open_pad(self.rapporten[-1])

    def open_pad(self, pad: Path) -> None:
        try:
            if isinstance(pad, Path) and pad == LOGMAP:
                LOGMAP.mkdir(parents=True, exist_ok=True)
            if sys.platform.startswith("win"):
                os.startfile(str(pad))                    # noqa: S606
            elif sys.platform == "darwin":
                subprocess.run(["open", str(pad)], check=False)
            else:
                subprocess.run(["xdg-open", str(pad)], check=False)
        except Exception as e:
            messagebox.showerror("Openen mislukt", f"{pad}\n\n{e}")

    def sluiten(self) -> None:
        if self.bezig and not messagebox.askokcancel(
                "Nog bezig", "Er draait nog een script. Toch sluiten?"):
            return
        if self.bezig:
            self.afbreken()
        self.bewaar()
        self.destroy()


# ==========================================================================
# 6. Opstarten
# ==========================================================================

def main() -> int:
    if not sys.platform.startswith("win"):
        toon_startfout("Alleen Windows",
                       "Deze tool draait PowerShell-scripts en werkt daarom "
                       "alleen op Windows.")
        return 1
    maak_scherp()
    try:
        App().mainloop()
    except Exception:
        # Een .pyw heeft geen console: zonder dit zie je bij een fout tijdens
        # het opbouwen van het venster helemaal niets gebeuren.
        toon_startfout("Kon niet starten",
                       "De tool kon niet opstarten.\n\n" + traceback.format_exc()[-1500:],
                       kopieer=traceback.format_exc())
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
