"""
PowerShell-runner
=================

Draai je PowerShell-scripts vanuit een venster, zonder in code te typen.

De runner leest het `param()`-blok van elk script en maakt daar automatisch
invulvelden van: mapkiezers voor mappen, vinkjes voor schakelaars, keuzelijsten
waar het script een vaste keuze verwacht. De uitvoer loopt live mee en wordt
opgeslagen in een logbestand.

Dubbelklik dit bestand. De .pyw-extensie zorgt dat Windows het zonder zwart
terminalvenster start. Er is niets te installeren; dit werkt op standaard
Windows met PowerShell.

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
vinkje, `[ValidateSet(...)]` een keuzelijst, `[string[]]` een invoervak waarin
je één waarde per regel zet.

Gebruik je `SupportsShouldProcess` in je script, dan verschijnt er een
Proefrun-vinkje: het script vertelt dan wat het zou doen, zonder het te doen.
"""

from __future__ import annotations

import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import traceback
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

# ==========================================================================


def pas_venstergrootte(venster, breedte: int, hoogte: int) -> None:
    """Open zo groot als gevraagd, maar nooit groter dan het scherm."""
    scherm_b = venster.winfo_screenwidth()
    scherm_h = venster.winfo_screenheight()
    b = max(700, min(breedte, scherm_b - 80))
    h = max(500, min(hoogte, scherm_h - 120))
    venster.geometry(f"{b}x{h}+{max(0, (scherm_b - b) // 2)}+{max(0, (scherm_h - h) // 3)}")


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

def bouw_opdracht(script: Script, waarden: dict[str, object], proefrun: bool,
                  powershell: str | None = None) -> list[str]:
    """De volledige opdrachtregel voor dit script met deze invoer."""
    ps = powershell or POWERSHELL
    if not ps:
        raise RuntimeError("PowerShell niet gevonden op deze computer.")

    opdracht = [ps, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                "-File", str(script.pad)]

    for parameter in script.parameters:
        waarde = waarden.get(parameter.naam)

        if parameter.soort == "vinkje":
            if waarde:
                opdracht.append(f"-{parameter.naam}")
            continue

        if parameter.soort == "lijst":
            regels = [r.strip() for r in str(waarde or "").splitlines() if r.strip()]
            if not regels:
                continue
            opdracht += [f"-{parameter.naam}", ",".join(regels)]
            continue

        tekst = str(waarde or "").strip()
        if tekst == "":
            continue
        opdracht += [f"-{parameter.naam}", tekst]

    if proefrun and script.proefrun_mogelijk:
        opdracht.append("-WhatIf")
    return opdracht


def ontbrekende_verplichte(script: Script, waarden: dict[str, object]) -> list[str]:
    ontbreekt = []
    for parameter in script.parameters:
        if not parameter.verplicht or parameter.soort == "vinkje":
            continue
        if not str(waarden.get(parameter.naam) or "").strip():
            ontbreekt.append(parameter.naam)
    return ontbreekt


# ==========================================================================
# 3. Logbestanden
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


# ==========================================================================
# 4. Het venster
# ==========================================================================

UITLEG = """Wat dit doet

Je kiest links een script, vult rechts de velden in en klikt op Uitvoeren. De
uitvoer loopt live mee in het onderste vak en wordt opgeslagen als logbestand.

De invulvelden worden gemaakt op basis van het script zelf. Verandert een
script, dan verandert het formulier mee; er is niets dubbel te onderhouden.


De vier meegeleverde scripts

    Alleen-lezen instellen   Zet bestanden onder een map op alleen-lezen, of
                             heft dat weer op. Kies zelf welke bestandstypen.

    Mappen vergelijken       Vergelijkt twee mappenbomen. Toont wat alleen in
                             A staat, wat alleen in B staat, en wat in beide
                             staat maar verschilt in grootte. Optioneel een
                             CSV-rapport voor Excel.

    Map inventariseren       Alle bestanden per submap, met grootte en datum,
                             plus de tien grootste bestanden. Optioneel een
                             CSV-rapport.

    Mappen kopieren          Kopieert dossiermappen naar een nieuwe locatie.
                             Via robocopy (snel over netwerkschijven) of
                             bestand voor bestand (kan duplicaten omnoemen
                             als naam_copy1).


Proefrun

Staat er een Proefrun-vinkje, dan kan het script vertellen wat het zou doen
zonder het te doen. Gebruik dat bij Alleen-lezen instellen en Mappen kopieren
voordat je het echt laat lopen, vooral als je een pad met de hand hebt
ingetypt.


Logbestanden

Elke run schrijft een logbestand in de map Logs naast dit bestand, met
tijdstempel en scriptnaam. Handig om later te kunnen aantonen wat er wanneer
gebeurd is. Logs ouder dan 90 dagen worden bij het opstarten opgeruimd; dat
getal staat bovenin dit bestand bij LOG_BEWAARDAGEN.


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
    [switch]                                        ->  vinkje
    [ValidateSet("a","b")]                          ->  keuzelijst
    [string[]]                                      ->  invoervak, per regel
    getaltype                                       ->  getalveld

Wil je dat een vinkje aangevinkt in beeld komt, zet dan "(standaard aan)" in de
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
uitroepteken voor de scriptnaam als het misgaat.
"""


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PowerShell-runner")
        pas_venstergrootte(self, 1120, 900)
        self.minsize(900, 560)

        self.scripts: list[Script] = []
        self.huidig: Script | None = None
        self.velden: dict[str, object] = {}
        self.proefrun_var = tk.BooleanVar(value=False)
        self.berichten: queue.Queue = queue.Queue()
        self.proces: subprocess.Popen | None = None
        self.bezig = False
        self.logbestand: Path | None = None
        self.log_handle = None

        self._bouw()
        self.protocol("WM_DELETE_WINDOW", self.sluiten)
        self.after(100, self.verwerk_berichten)

        if POWERSHELL is None:
            self.schrijf("PowerShell niet gevonden op deze computer. Dat is "
                         "ongebruikelijk op Windows.\n", "fout")
        opgeruimd = ruim_logs_op()
        if opgeruimd:
            self.schrijf(f"{opgeruimd} oude logbestanden opgeruimd.\n", "grijs")
        self.herlaad()

    # ------------------------------------------------------------- opbouw
    def _bouw(self) -> None:
        try:
            ttk.Style().theme_use("vista")
        except tk.TclError:
            pass

        buiten = ttk.Frame(self, padding=12)
        buiten.pack(fill="both", expand=True)

        kop = ttk.Frame(buiten)
        kop.pack(fill="x")
        ttk.Label(kop, text="PowerShell-runner",
                  font=("Segoe UI", 15, "bold")).pack(side="left")
        ttk.Label(kop, text="  scripts draaien zonder in code te typen",
                  foreground="#777").pack(side="left", pady=(4, 0))
        ttk.Button(kop, text="Uitleg", command=self.toon_uitleg).pack(side="right")

        # onderkant vastzetten zodat de knoppen nooit buiten beeld vallen
        knoppen = ttk.Frame(buiten)
        knoppen.pack(side="bottom", fill="x", pady=(10, 0))
        self.knop_start = ttk.Button(knoppen, text="Uitvoeren",
                                     command=self.start, state="disabled")
        self.knop_start.pack(side="left")
        self.knop_stop = ttk.Button(knoppen, text="Afbreken",
                                    command=self.afbreken, state="disabled")
        self.knop_stop.pack(side="left", padx=8)
        self.knop_log = ttk.Button(knoppen, text="Open log",
                                   command=self.open_log, state="disabled")
        self.knop_log.pack(side="left")
        ttk.Button(knoppen, text="Open scriptmap",
                   command=lambda: self.open_pad(SCRIPTMAP)).pack(side="left", padx=8)
        ttk.Button(knoppen, text="Sluiten", command=self.sluiten).pack(side="right")
        self.knop_hint = ttk.Label(knoppen, text="", foreground="#777")
        self.knop_hint.pack(side="left", padx=12)

        uitvoervak = ttk.Frame(buiten)
        uitvoervak.pack(side="bottom", fill="both", expand=False, pady=(6, 0))
        ttk.Label(uitvoervak, text="Uitvoer").pack(anchor="w")
        binnen = ttk.Frame(uitvoervak)
        binnen.pack(fill="both", expand=True)
        self.log = tk.Text(binnen, height=14, wrap="none", relief="flat",
                           background="#1E1E1E", foreground="#DCDCDC",
                           insertbackground="#DCDCDC", font=("Consolas", 9))
        vert = ttk.Scrollbar(binnen, command=self.log.yview)
        hor = ttk.Scrollbar(uitvoervak, orient="horizontal", command=self.log.xview)
        self.log.configure(yscrollcommand=vert.set, xscrollcommand=hor.set,
                           state="disabled")
        self.log.pack(side="left", fill="both", expand=True)
        vert.pack(side="right", fill="y")
        hor.pack(fill="x")
        self.log.tag_configure("kop", foreground="#FFFFFF", font=("Consolas", 9, "bold"))
        self.log.tag_configure("ok", foreground="#6FCF6F")
        self.log.tag_configure("fout", foreground="#F08080")
        self.log.tag_configure("grijs", foreground="#9A9A9A")

        self.balk = ttk.Progressbar(buiten, mode="determinate")
        self.balk.pack(side="bottom", fill="x", pady=(10, 0))

        midden = ttk.PanedWindow(buiten, orient="horizontal")
        midden.pack(fill="both", expand=True, pady=(10, 0))
        midden.add(self._bouw_scriptlijst(midden), weight=2)
        midden.add(self._bouw_formulier(midden), weight=3)

    def _bouw_scriptlijst(self, ouder) -> ttk.Frame:
        vak = ttk.Frame(ouder)
        rij = ttk.Frame(vak)
        rij.pack(fill="x")
        ttk.Label(rij, text="Scripts").pack(side="left")
        ttk.Button(rij, text="Herlaad", command=self.herlaad).pack(side="right")

        binnen = ttk.Frame(vak)
        binnen.pack(fill="both", expand=True, pady=(4, 0))
        self.lijst = ttk.Treeview(binnen, columns=(), show="tree", selectmode="browse")
        schuif = ttk.Scrollbar(binnen, command=self.lijst.yview)
        self.lijst.configure(yscrollcommand=schuif.set)
        self.lijst.pack(side="left", fill="both", expand=True)
        schuif.pack(side="right", fill="y")
        self.lijst.bind("<<TreeviewSelect>>", self.op_script_gekozen)

        self.label_omschrijving = ttk.Label(vak, text="", foreground="#555",
                                            wraplength=320, justify="left")
        self.label_omschrijving.pack(anchor="w", pady=(6, 0))
        return vak

    def _bouw_formulier(self, ouder) -> ttk.Frame:
        vak = ttk.Frame(ouder)
        self.label_script = ttk.Label(vak, text="Instellingen",
                                      font=("Segoe UI", 9, "bold"))
        self.label_script.pack(anchor="w")

        binnen = ttk.Frame(vak)
        binnen.pack(fill="both", expand=True, pady=(4, 0))
        self.doek = tk.Canvas(binnen, highlightthickness=0, background="#FFFFFF")
        schuif = ttk.Scrollbar(binnen, command=self.doek.yview)
        self.doek.configure(yscrollcommand=schuif.set)
        self.doek.pack(side="left", fill="both", expand=True)
        schuif.pack(side="right", fill="y")
        self.formulier = ttk.Frame(self.doek, padding=8)
        self._venster = self.doek.create_window((0, 0), window=self.formulier, anchor="nw")
        self.formulier.bind("<Configure>", lambda e: self.doek.configure(
            scrollregion=self.doek.bbox("all")))
        self.doek.bind("<Configure>", lambda e: self.doek.itemconfig(
            self._venster, width=e.width))
        self.doek.bind_all("<MouseWheel>", self._wiel)

        self.proefrun_vinkje = ttk.Checkbutton(
            vak, text="Proefrun — laat zien wat het zou doen, zonder het te doen",
            variable=self.proefrun_var)
        self.proefrun_vinkje.pack(anchor="w", pady=(8, 0))
        return vak

    def _wiel(self, e) -> None:
        try:
            self.doek.yview_scroll(int(-1 * (e.delta / 120)), "units")
        except tk.TclError:
            pass

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

    def toon_uitleg(self) -> None:
        venster = tk.Toplevel(self)
        venster.title("Uitleg")
        venster.geometry("820x700")
        vak = ttk.Frame(venster, padding=12)
        vak.pack(fill="both", expand=True)
        tekst = tk.Text(vak, wrap="word", relief="flat", font=("Consolas", 9))
        schuif = ttk.Scrollbar(vak, command=tekst.yview)
        tekst.configure(yscrollcommand=schuif.set)
        tekst.pack(side="left", fill="both", expand=True)
        schuif.pack(side="right", fill="y")
        tekst.insert("1.0", UITLEG)
        tekst.configure(state="disabled")
        ttk.Button(venster, text="Sluiten", command=venster.destroy).pack(pady=(0, 10))

    # ------------------------------------------------------------ scripts
    def herlaad(self) -> None:
        SCRIPTMAP.mkdir(parents=True, exist_ok=True)
        self.scripts = lees_scripts(SCRIPTMAP)
        self.lijst.delete(*self.lijst.get_children())
        for script in self.scripts:
            merk = "!  " if (script.fout or script.waarschuwing) else ""
            self.lijst.insert("", "end", iid=str(script.pad), text=merk + script.naam)
        if not self.scripts:
            self.schrijf(f"Geen .ps1-scripts gevonden in {SCRIPTMAP}\n", "fout")
        else:
            self.schrijf(f"{len(self.scripts)} scripts gevonden.\n", "grijs")
        for script in self.scripts:
            if script.waarschuwing:
                self.schrijf(f"{script.naam}: {script.waarschuwing}\n", "fout")
        self.werk_startknop_bij()

    def op_script_gekozen(self, *_) -> None:
        keuze = self.lijst.selection()
        if not keuze:
            return
        self.huidig = next((s for s in self.scripts if str(s.pad) == keuze[0]), None)
        self.bouw_velden()

    def bouw_velden(self) -> None:
        for widget in self.formulier.winfo_children():
            widget.destroy()
        self.velden.clear()

        script = self.huidig
        if script is None:
            return

        self.label_script.configure(text=f"Instellingen — {script.naam}")
        self.label_omschrijving.configure(text=script.omschrijving or "(geen omschrijving)")

        if script.fout:
            ttk.Label(self.formulier, foreground="#C0392B", wraplength=460,
                      text=f"Dit script kon niet gelezen worden:\n{script.fout}").pack(
                anchor="w")
            self.werk_startknop_bij()
            return

        if not script.parameters:
            ttk.Label(self.formulier, foreground="#777", wraplength=460,
                      text="Dit script heeft geen parameters; je kunt het direct "
                           "uitvoeren.").pack(anchor="w")

        for parameter in script.parameters:
            self._bouw_veld(parameter)

        stand = "normal" if script.proefrun_mogelijk else "disabled"
        self.proefrun_vinkje.configure(state=stand)
        if not script.proefrun_mogelijk:
            self.proefrun_var.set(False)
        self.werk_startknop_bij()

    def _bouw_veld(self, parameter: Parameter) -> None:
        blok = ttk.Frame(self.formulier)
        blok.pack(fill="x", pady=(0, 10))

        label = parameter.naam + (" *" if parameter.verplicht else "")
        ttk.Label(blok, text=label, font=("Segoe UI", 9, "bold")).pack(anchor="w")
        if parameter.uitleg:
            ttk.Label(blok, text=parameter.uitleg, foreground="#777",
                      wraplength=460, justify="left").pack(anchor="w")

        if parameter.soort == "vinkje":
            # Staat "(standaard aan)" in de HelpMessage van het script, dan komt
            # het vinkje aangevinkt in beeld. Zo kan een script een schakelaar
            # standaard aan hebben zonder dubbele ontkenningen als NietOpenen.
            aan = "standaard aan" in (parameter.uitleg or "").lower()
            var = tk.BooleanVar(value=aan)
            ttk.Checkbutton(blok, text="aan", variable=var).pack(anchor="w", pady=(2, 0))
            self.velden[parameter.naam] = var
            return

        if parameter.soort == "keuze":
            var = tk.StringVar(value=parameter.standaard or
                               (parameter.keuzes[0] if parameter.keuzes else ""))
            ttk.Combobox(blok, textvariable=var, values=parameter.keuzes,
                         state="readonly", width=28).pack(anchor="w", pady=(2, 0))
            self.velden[parameter.naam] = var
            return

        if parameter.soort == "lijst":
            vak = tk.Text(blok, height=5, wrap="none", font=("Consolas", 9))
            vak.pack(fill="x", pady=(2, 0))
            if parameter.standaard:
                vak.insert("1.0", "\n".join(
                    d.strip() for d in parameter.standaard.split(",") if d.strip()))
            ttk.Label(blok, text="één waarde per regel", foreground="#999").pack(anchor="w")
            self.velden[parameter.naam] = vak
            return

        rij = ttk.Frame(blok)
        rij.pack(fill="x", pady=(2, 0))
        var = tk.StringVar(value=parameter.standaard)
        ttk.Entry(rij, textvariable=var).pack(side="left", fill="x", expand=True)
        if parameter.soort == "map":
            ttk.Button(rij, text="Map…", width=7,
                       command=lambda v=var: self.kies_map(v)).pack(side="left", padx=(6, 0))
        elif parameter.soort == "bestand":
            ttk.Button(rij, text="Bestand…", width=10,
                       command=lambda v=var, p=parameter: self.kies_bestand(v, p)).pack(
                side="left", padx=(6, 0))
        var.trace_add("write", lambda *_: self.werk_startknop_bij())
        self.velden[parameter.naam] = var

    def kies_map(self, var: tk.StringVar) -> None:
        huidig = var.get().strip()
        gekozen = filedialog.askdirectory(
            title="Kies een map",
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
                title="Kies waar het bestand komt", defaultextension=extensie,
                filetypes=soorten, initialdir=beginmap,
                initialfile=Path(huidig).name if huidig else "")
        else:
            gekozen = filedialog.askopenfilename(title="Kies een bestand",
                                                 filetypes=soorten, initialdir=beginmap)
        if gekozen:
            var.set(str(Path(gekozen)))

    def huidige_waarden(self) -> dict[str, object]:
        waarden: dict[str, object] = {}
        for naam, widget in self.velden.items():
            if isinstance(widget, tk.Text):
                waarden[naam] = widget.get("1.0", "end")
            else:
                waarden[naam] = widget.get()
        return waarden

    # ----------------------------------------------------------- knoppen
    def werk_startknop_bij(self) -> None:
        if self.bezig:
            reden = "bezig…"
        elif POWERSHELL is None:
            reden = "PowerShell niet gevonden"
        elif self.huidig is None:
            reden = "kies links een script"
        elif self.huidig.fout:
            reden = "dit script kan niet gelezen worden"
        else:
            ontbreekt = ontbrekende_verplichte(self.huidig, self.huidige_waarden())
            reden = ("vul in: " + ", ".join(ontbreekt)) if ontbreekt else ""
        self.knop_start.configure(state="disabled" if reden else "normal")
        if hasattr(self, "knop_hint"):
            self.knop_hint.configure(text=reden)

    # ------------------------------------------------------------ draaien
    def start(self) -> None:
        if self.bezig or self.huidig is None:
            return
        waarden = self.huidige_waarden()
        ontbreekt = ontbrekende_verplichte(self.huidig, waarden)
        if ontbreekt:
            messagebox.showwarning("Nog niet compleet",
                                   "Deze velden zijn verplicht:\n\n" + "\n".join(ontbreekt))
            return

        try:
            opdracht = bouw_opdracht(self.huidig, waarden, self.proefrun_var.get())
        except RuntimeError as e:
            messagebox.showerror("Kan niet starten", str(e))
            return

        self.leeg_log()
        self.bezig = True
        self.knop_stop.configure(state="normal")
        self.knop_log.configure(state="disabled")
        self.balk.configure(mode="indeterminate")
        self.balk.start(12)
        self.werk_startknop_bij()

        LOGMAP.mkdir(parents=True, exist_ok=True)
        self.logbestand = logpad(self.huidig.naam)

        kopregels = [
            f"{self.huidig.naam}",
            f"gestart {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            "proefrun: ja" if self.proefrun_var.get() and self.huidig.proefrun_mogelijk
            else "proefrun: nee",
            "",
        ]
        for regel in kopregels:
            self.schrijf(regel + "\n", "kop" if regel else "")

        threading.Thread(target=self.werk, args=(opdracht, self.logbestand),
                         daemon=True).start()

    def werk(self, opdracht: list[str], log: Path) -> None:
        try:
            with open(log, "w", encoding="utf-8") as bestand:
                bestand.write(" ".join(opdracht) + "\n\n")
                bestand.flush()

                self.proces = subprocess.Popen(
                    opdracht, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace", bufsize=1,
                    creationflags=GEEN_VENSTER)

                for regel in self.proces.stdout:
                    schoon = regel.rstrip("\n")
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
        self.knop_stop.configure(state="disabled")
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
            while True:
                bericht = self.berichten.get_nowait()
                if bericht[0] == "regel":
                    regel = bericht[1]
                    tag = ""
                    laag = regel.lower()
                    if any(w in laag for w in ("mislukt", "fout", "error", "bestaat niet",
                                               "not found", "exception")):
                        tag = "fout"
                    elif laag.startswith("klaar") or "opgeslagen" in laag:
                        tag = "ok"
                    self.schrijf(regel + "\n", tag)
                elif bericht[0] == "klaar":
                    self.rond_af(*bericht[1:])
        except queue.Empty:
            pass
        self.after(100, self.verwerk_berichten)

    def rond_af(self, code: int, log: str, fout: str = "") -> None:
        self.bezig = False
        self.balk.stop()
        self.balk.configure(mode="determinate", value=0)
        self.knop_stop.configure(state="disabled")
        self.knop_log.configure(state="normal")
        self.logbestand = Path(log)

        if fout:
            self.schrijf(f"\nKon het script niet draaien: {fout}\n", "fout")
        elif code == 0:
            self.schrijf("\nKlaar zonder fouten.\n", "ok")
        else:
            self.schrijf(f"\nAfgesloten met code {code}. "
                         "Kijk hierboven welke regels roodgekleurd zijn.\n", "fout")
        self.schrijf(f"Log: {log}\n", "grijs")
        self.werk_startknop_bij()

    # ------------------------------------------------------------ overig
    def open_log(self) -> None:
        if self.logbestand:
            self.open_pad(self.logbestand)

    def open_pad(self, pad: Path) -> None:
        try:
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
        self.destroy()


def main() -> int:
    if not sys.platform.startswith("win"):
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Alleen Windows",
                             "Deze tool draait PowerShell-scripts en werkt daarom "
                             "alleen op Windows.")
        root.destroy()
        return 1
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
