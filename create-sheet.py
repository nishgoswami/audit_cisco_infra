#!/usr/bin/env python3

"""
Cisco Migration Sheet Generator
================================

Purpose
-------
Reads a multi-switch audit text file (collected by a remote SSH script) and
produces a single Excel spreadsheet (.xlsx) with one row per MAC address seen
on each switch port.  Columns include:

  Switch hostname, Switch Model, Building, Floor, SNMP Location,
  Interface Status, Interface Number, Interface Description,
  CDP Neighbour, CDP Platform, Uplink (Yes/No),
  VLAN, MAC Address, IP Address, MAC Count

Usage
-----
  python3 create-sheet.py <input_file.txt> <output_file.xlsx>

Example:
  python3 create-sheet.py output_20260302_225516.txt migration.xlsx

Input file format
-----------------
The script expects one or more switch sections, each starting with:

  =================================================================
   Switch : <hostname>
   Time   : ...
  =================================================================

Within each section the following Cisco show-command outputs are parsed:
  - show running-config | include (hostname|snmp-server location)
  - show mac address-table
  - show interface description
  - show arp
  - show cdp neighbors
  - sh version  (model extraction)

Requirements
------------
  pip install pandas openpyxl rich
"""



import sys

import os

import argparse

import sqlite3

import datetime

import re

import time

import math                  # imported at module level (was incorrectly inside a loop)

import socket

import urllib.request

import urllib.error

import json

import ssl

from concurrent.futures import ThreadPoolExecutor, as_completed

from typing import List, Dict, Tuple

from collections import Counter



# ── Dependency check ──────────────────────────────────────────────────────────

# Run this BEFORE any third-party import so the error message is always visible.



_REQUIRED = {

    "pandas":   "pandas",

    "openpyxl": "openpyxl",

    "rich":     "rich",

}



import importlib.util as _iutil

_missing = [pkg for mod, pkg in _REQUIRED.items() if _iutil.find_spec(mod) is None]

if _missing:

    print()

    print("  ❌  Missing required package(s): " + ", ".join(_missing))

    print()

    print("  Install with:")

    print(f"    pip install {' '.join(_missing)}")

    print()

    sys.exit(1)



# ── Third-party imports (safe after the guard) ────────────────────────────────



import pandas as pd



from rich.console import Console

from rich.panel import Panel

from rich.table import Table

from rich.text import Text

from rich.progress import (

    Progress, SpinnerColumn, TextColumn, BarColumn,

    TaskProgressColumn, TimeElapsedColumn,

)

from rich import box  # Live and Columns were imported but never used — removed



# ── Terminal background detection ─────────────────────────────────────────────



def _is_dark_background() -> bool:

    """

    Detect whether the terminal has a dark background.

    Strategy (in order):

      1. COLORFGBG env var  – set by rxvt / konsole  ('fg;bg')

      2. OSC 11 query       – xterm-compatible escape (iTerm2, macOS Terminal,

                              kitty, alacritty, WezTerm …).  Reads back the

                              actual background RGB and computes luminance.

      3. Fallback           – assume dark (most dev terminals are dark).

    """

    # 1. COLORFGBG  (last semicolon-separated component is the bg ANSI index)

    colorfgbg = os.environ.get("COLORFGBG", "")

    if colorfgbg:

        try:

            bg_idx = int(colorfgbg.rsplit(";", 1)[-1])

            return bg_idx < 8     # 0-7 are the dark ANSI palette colours

        except ValueError:

            pass



    # 2. OSC 11 query – only attempt on a real interactive tty

    if sys.stdout.isatty() and sys.stdin.isatty():

        try:

            import termios, tty, select

            fd  = sys.stdin.fileno()

            old = termios.tcgetattr(fd)

            try:

                tty.setraw(fd)

                # Ask terminal: what is the background colour?

                sys.stdout.write("\033]11;?\033\\")

                sys.stdout.flush()

                resp = ""

                deadline = time.monotonic() + 0.25    # 250 ms budget

                while time.monotonic() < deadline:

                    ready, _, _ = select.select([sys.stdin], [], [], 0.05)

                    if not ready:

                        continue

                    ch = sys.stdin.read(1)

                    resp += ch

                    if resp.endswith("\033\\") or resp.endswith("\007"):

                        break

            finally:

                termios.tcsetattr(fd, termios.TCSADRAIN, old)



            # Response format: ESC ] 11 ; rgb:RRRR/GGGG/BBBB ST

            m = re.search(

                r"rgb:([0-9a-fA-F]+)/([0-9a-fA-F]+)/([0-9a-fA-F]+)", resp

            )

            if m:

                # Components may be 2 or 4 hex digits – normalise to 0-255

                def _to_byte(s: str) -> int:

                    v = int(s, 16)

                    return v >> 8 if len(s) == 4 else v

                r = _to_byte(m.group(1))

                g = _to_byte(m.group(2))

                b = _to_byte(m.group(3))

                luminance = 0.299 * r + 0.587 * g + 0.114 * b

                return luminance < 128

        except Exception:

            pass



    # 3. Fallback – assume dark

    return True



# ── Colour themes ─────────────────────────────────────────────────────────────

#

# Every style string used in the UI lives here so that swapping theme is a

# single-line change.  Colours were chosen for readability on each background.



_DARK: dict = {

    # header

    "badge":           "bold white on #b34700",

    "banner_text":     "bold #e8e8e8",

    "border":          "#b34700",

    # progress

    "spinner":         "bold #ff9500",

    "bar_empty":       "grey23",

    "bar_fill":        "#ff9500",

    "bar_done":        "#50fa7b",

    "pct":             "bold #ffd080",

    "task_text":       "bold #e8e8e8",

    # info lines

    "label":           "grey62",

    "value":           "bold #e8e8e8",

    "found_num":       "bold #ffd080",

    "done_word":       "#50fa7b",

    # table

    "tbl_hdr":         "bold #e8e8e8",

    "tbl_border":      "grey35",

    "col_hostname":    "bold #e8e8e8",

    "col_macs":        "#ffd080",

    "col_records":     "#50fa7b",

    "col_building":    "#6fa8dc",

    "col_floor":       "#c49ad8",

    "col_model":       "grey74",

    "check":           "#50fa7b",

    "nil":             "grey42",

    # summary panel

    "sum_label":       "grey62",

    "sum_file":        "bold #ff9500",

    "sum_switches":    "bold #ffd080",

    "sum_rows":        "bold #50fa7b",

    "complete_title":  "bold #50fa7b",

    "complete_border": "#50fa7b",

}



_LIGHT: dict = {

    # header

    "badge":           "bold white on #8b2500",

    "banner_text":     "bold #1a1a1a",

    "border":          "#8b2500",

    # progress

    "spinner":         "bold #8b2500",

    "bar_empty":       "grey82",

    "bar_fill":        "#8b2500",

    "bar_done":        "#1a6b1a",

    "pct":             "bold #5c3300",

    "task_text":       "bold #1a1a1a",

    # info lines

    "label":           "grey42",

    "value":           "bold #1a1a1a",

    "found_num":       "bold #5c3300",

    "done_word":       "#1a6b1a",

    # table

    "tbl_hdr":         "bold #1a1a1a",

    "tbl_border":      "grey58",

    "col_hostname":    "bold #1a1a1a",

    "col_macs":        "#5c3300",

    "col_records":     "#1a6b1a",

    "col_building":    "#1a4080",

    "col_floor":       "#5c1a7a",

    "col_model":       "grey42",

    "check":           "#1a6b1a",

    "nil":             "grey58",

    # summary panel

    "sum_label":       "grey42",

    "sum_file":        "bold #8b2500",

    "sum_switches":    "bold #5c3300",

    "sum_rows":        "bold #1a6b1a",

    "complete_title":  "bold #1a6b1a",

    "complete_border": "#1a6b1a",

}



T: dict = _DARK if _is_dark_background() else _LIGHT



console = Console(highlight=False)



def normalize_port(port: str) -> str:

    """

    Convert a Cisco interface name to its 2-letter abbreviated form.



    Examples:

      GigabitEthernet2/1   ->  Gi2/1

      TenGigabitEthernet1/4  ->  Te1/4

      Port-channel101      ->  Po101



    Normalising to a consistent short form ensures that MAC table, interface

    description, CDP and ARP entries all join on the same key.

    Any unrecognised prefix is returned unchanged.

    Supported families (keeps all keys consistent across Catalyst, Nexus, ISR):

      GigabitEthernet / Gig       -> Gi

      FastEthernet                -> Fa

      TenGigabitEthernet / Ten    -> Te

      TwentyFiveGigE / Twe        -> Twe   (Catalyst 9k)

      TwoGigabitEthernet          -> Tw    (Catalyst 9k low-end)

      FiveGigabitEthernet         -> Fi    (Catalyst 9k low-end)

      HundredGigE / Hu            -> Hu    (Catalyst 9k uplinks)

      FortyGigabitEthernet        -> Fo

      Ethernet                    -> Eth   (Nexus)

      AppGigabitEthernet          -> Ap    (Catalyst 9k)

      Management / mgmt           -> Ma

      Port-channel                -> Po

      Vlan                        -> Vl

    """

    # Order matters: check longer prefixes before shorter ones to avoid

    # partial matches (e.g. "TenGigabitEthernet" must come before "Te").

    _MAP = [

        ("GigabitEthernet",    "Gi"),

        ("FastEthernet",       "Fa"),

        ("TenGigabitEthernet", "Te"),

        ("TwentyFiveGigE",     "Twe"),

        ("TwoGigabitEthernet", "Tw"),

        ("FiveGigabitEthernet","Fi"),

        ("HundredGigE",        "Hu"),

        ("FortyGigabitEthernet","Fo"),

        ("AppGigabitEthernet", "Ap"),

        ("Ethernet",           "Eth"),   # Nexus uses bare 'Ethernet'

        ("Management",         "Ma"),

        ("mgmt",               "Ma"),    # some Nexus platforms

        ("Port-channel",       "Po"),

        ("port-channel",       "Po"),

        ("Vlan",               "Vl"),

        # Short abbreviations used in CDP / MAC table output

        ("Gig",  "Gi"),

        ("Ten",  "Te"),

        ("Twe",  "Twe"),

        ("Fo",   "Fo"),

        ("Hu",   "Hu"),

        ("Tw",   "Tw"),

        ("Fi",   "Fi"),

        ("Eth",  "Eth"),

        ("Ma",   "Ma"),

    ]

    for full, short in _MAP:

        if port.startswith(full):

            return short + port[len(full):]

    return port



def normalize_mac(mac: str) -> str:

    """

    Normalise a MAC address to lowercase colon-separated pairs.



    Accepts Cisco dot notation (aabb.ccdd.eeff), colons, or dashes.

    Example:  8843.e1a3.2c80  ->  88:43:e1:a3:2c:80

    """

    clean = re.sub(r'[.:\-]', '', mac).lower()

    return ':'.join(clean[i:i+2] for i in range(0, len(clean), 2))



def short_hostname(device_id: str) -> str:

    """

    Return the short hostname portion of a CDP device ID.



    Cisco CDP often reports FQDNs (e.g. mondtd02.network.rogers.com).

    We keep only the label before the first dot so the spreadsheet shows

    clean hostnames that match the Switch hostname column.

    Short names (e.g. mcgwc01) are returned unchanged.

    """

    return device_id.split('.')[0]


# ── OUI → Device-Type lookup ─────────────────────────────────────────────────
# Keyed by the first 6 hex digits (lower-case, no separators) of a MAC address.
# Sources: IEEE MA-L registry + well-known vendor assignments.
# Only high-confidence mappings are included; unknown/ambiguous OUIs fall
# through to "Unknown" rather than risk misclassifying a device.
_OUI_TABLE: Dict[str, str] = {
    # ── Printers ──────────────────────────────────────────────────────────────
    "000048": "Printer",    # Seiko Epson
    "0026ab": "Printer",    # Seiko Epson
    "0004b0": "Printer",    # Canon Inc.
    "001e8f": "Printer",    # Canon Information Systems
    "000085": "Printer",    # Canon KK
    "c0eefb": "Printer",    # Canon Inc.
    "0000aa": "Printer",    # Xerox Corporation
    "000065": "Printer",    # Xerox
    "080037": "Printer",    # Fuji Xerox
    "001722": "Printer",    # Ricoh Company Ltd
    "002673": "Printer",    # Ricoh Company Ltd
    "0017a4": "Printer",    # HP Inc.
    "001b78": "Printer",    # HP Inc.
    "001ee0": "Printer",    # HP Inc.
    "283737": "Printer",    # HP Inc.
    "38bb3c": "Printer",    # HP Inc.
    "3c29ab": "Printer",    # HP Inc.
    "48dc2d": "Printer",    # HP Inc.
    "4c391a": "Printer",    # HP Inc.
    "9cb654": "Printer",    # HP Inc. (LaserJet)
    "c8d9d2": "Printer",    # HP Inc.
    "f0921c": "Printer",    # HP Inc.
    "00c0ee": "Printer",    # Konica Minolta
    "009098": "Printer",    # Konica Minolta
    "000877": "Printer",    # Brother Industries
    "001ba9": "Printer",    # Murata Machinery (fax/printer)
    "0050f9": "Printer",    # Sharp Corporation
    # ── IP Phones ─────────────────────────────────────────────────────────────
    "000ef7": "IP Phone",   # Cisco IP Phone (Linksys VOIP)
    "000ed7": "IP Phone",   # Cisco VOIP
    "001cc0": "IP Phone",   # Cisco UC Phone
    "805ec0": "IP Phone",   # Yealink Network Technology
    "001b5c": "IP Phone",   # Snom Technology (SIP phones)
    "002e58": "IP Phone",   # Fanvil Technology
    # ── Video Conferencing ────────────────────────────────────────────────────
    "0004f2": "Video Conf", # Polycom Inc.
    "00e0db": "Video Conf", # Tandberg (Cisco TelePresence)
    "506b8d": "Video Conf", # Cisco TelePresence
    # ── Wireless APs ──────────────────────────────────────────────────────────
    "000b85": "Wireless AP", # Cisco Aironet
    "001a1e": "Wireless AP", # Aruba Networks (HPE)
    "d4684d": "Wireless AP", # Ruckus Networks
    "90b11c": "Wireless AP", # Ruckus Networks (ZoneFlex)  ← seen in Vlan113
    "58970b": "Wireless AP", # Cisco Meraki
    "88dc96": "Wireless AP", # Cisco Meraki
    "e0cbbc": "Wireless AP", # Cisco Meraki
    "0024d4": "Wireless AP", # Cisco Aironet 1600/2600
    "88155e": "Wireless AP", # Cisco Meraki MR
    # ── Virtual Machines ──────────────────────────────────────────────────────
    "005056": "Virtual Machine", # VMware vSphere / ESXi  ← seen in Vlan116
    "000c29": "Virtual Machine", # VMware Workstation
    "000569": "Virtual Machine", # VMware ESX
    "001c14": "Virtual Machine", # VMware
    "080027": "Virtual Machine", # Oracle VirtualBox
    "525400": "Virtual Machine", # QEMU / KVM
    "001c42": "Virtual Machine", # Parallels Desktop
    # ── IP Cameras ────────────────────────────────────────────────────────────
    "000f9b": "IP Camera",  # Axis Communications  ← 000f.9b02.4add seen in Vlan113
    "accc8e": "IP Camera",  # Axis Communications
    "00408c": "IP Camera",  # Axis Communications (legacy)
    "001277": "IP Camera",  # Axis Communications
    "d845cc": "IP Camera",  # Hanwha Techwin (Samsung cameras)
    "00166c": "IP Camera",  # Axis Communications
    "b4a20e": "IP Camera",  # Hikvision
    "c0562e": "IP Camera",  # Hikvision
    "4c5563": "IP Camera",  # Dahua Technology
    "e069a6": "IP Camera",  # Bosch Security
    # ── Servers (BMC / iDRAC / iLO interfaces — specific OUIs only) ───────────
    # Using OUI-specific entries here rather than broad Dell/HP vendor-name
    # matching, because those vendors also make workstations and printers.
    "18dbf2": "Server",     # Dell Inc. (iDRAC)
    "f0edc9": "Server",     # Dell Inc. (iDRAC)
    "f8db88": "Server",     # Dell Inc. (iDRAC)
    "8cec4b": "Server",     # Dell Inc. (iDRAC)
    "0021f6": "Server",     # Dell Inc. (iDRAC)
    "40a8f0": "Server",     # HP ProLiant (iLO)
    "001a4b": "Server",     # HP ProLiant (iLO)
    "00237d": "Server",     # HP ProLiant
    "3c4a92": "Server",     # HP Enterprise (ProLiant Gen9+)
    "001a64": "Server",     # IBM System x
    "00e081": "Server",     # IBM (BladeCenter)
    # ── Workstations / PCs ────────────────────────────────────────────────────
    # Only OUIs registered to vendors that exclusively make end-user client
    # hardware (laptops, desktops) are listed here.  Ambiguous vendors that
    # also make servers (Dell, HP) are handled via _vendor_to_device_type()
    # using more specific keyword rules.
    "28d244": "Workstation", # Lenovo
    "3417eb": "Workstation", # Lenovo
    "50795a": "Workstation", # Lenovo
    "54ee75": "Workstation", # Lenovo
    "6c4008": "Workstation", # Lenovo
    "70723c": "Workstation", # Lenovo
    "8c8d28": "Workstation", # Lenovo
    "acf2c5": "Workstation", # Lenovo
    "0022fb": "Workstation", # Lenovo
    "0024be": "Workstation", # Lenovo
    "001a92": "Workstation", # ASUSTeK Computer
    "04d4c4": "Workstation", # ASUSTeK Computer
    "08606e": "Workstation", # ASUSTeK Computer
    "1002b5": "Workstation", # ASUSTeK Computer
    "2cfda1": "Workstation", # ASUSTeK Computer
    "40b034": "Workstation", # ASUSTeK Computer
    "90e6ba": "Workstation", # ASUSTeK Computer
    "000393": "Workstation", # Apple Inc. (Mac — wired corporate context)
    "000a27": "Workstation", # Apple Inc. (Mac)
    "000a95": "Workstation", # Apple Inc. (Mac)
    "001124": "Workstation", # Apple Inc. (Mac)
    "001451": "Workstation", # Apple Inc. (Mac)
    "0016cb": "Workstation", # Apple Inc. (Mac)
    "0017f2": "Workstation", # Apple Inc. (Mac)
    "001b63": "Workstation", # Apple Inc. (Mac)
    "001cb3": "Workstation", # Apple Inc. (Mac)
    "001e52": "Workstation", # Apple Inc. (Mac)
    "28cfda": "Workstation", # Apple Inc. (Mac)
    "3c0754": "Workstation", # Apple Inc. (Mac)
    "a4c361": "Workstation", # Apple Inc. (Mac)
}


# ── Runtime OUI lookup state ─────────────────────────────────────────────────
# These are populated once at the start of process_file() and reused by every
# device_type() / _lookup_oui() call during the run.
_OUI_CACHE: Dict[str, str] = {}   # oui-6-hex-lower → vendor name string
_INTERNET_OK: bool         = False # set to True if connectivity confirmed


# Unverified SSL context used only for OUI vendor lookups (public, non-sensitive
# data).  Required on macOS Python 3.x where the system CA bundle is not wired
# into the default urllib context until 'Install Certificates.command' is run.
_OUI_SSL = ssl._create_unverified_context()


def _check_internet() -> bool:
    """
    Return True if the host can reach the public internet.

    Uses a 2-second TCP connect to 8.8.8.8:53 (Google DNS).  No DNS lookup
    needed so this works even when DNS itself is misconfigured.
    """
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(2)
        sock.connect(("8.8.8.8", 53))
        sock.close()
        return True
    except OSError:
        return False


def _lookup_oui(oui6: str) -> str:
    """
    Return the cached vendor name for a 6-hex-char OUI string.

    This is called per-row during row assembly and is always a pure cache hit
    after _prefetch_ouis() has run.  If somehow called before prefetch (e.g.
    in tests), it returns the cached value or an empty string.
    """
    return _OUI_CACHE.get(oui6, "")


def _prefetch_ouis(macs: list, rich_console) -> None:
    """
    Pre-fetch vendor names for all unique OUI prefixes found in *macs* and
    populate _OUI_CACHE so subsequent calls to _lookup_oui() are instant.

    Strategy (fastest first, automatic fallback):

      1. 15 concurrent GET threads to api.maclookup.app/v2/macs/AA:BB:CC
         Returns structured JSON with company name.  No rate-limiting on the
         free tier; 47 OUIs typically resolve in under 3 seconds.

      2. If maclookup.app is entirely unreachable (ALL requests fail), fall
         back to 10 threads against api.macvendors.com with 429 retry/backoff.

    The progress bar updates in real time as threads complete.
    """
    from rich.progress import (Progress as _Prog, SpinnerColumn as _Spin,
                                TextColumn as _Txt, BarColumn as _Bar,
                                MofNCompleteColumn as _MofN,
                                TimeElapsedColumn as _Time)

    unique: set = set()
    for mac in macs:
        clean = re.sub(r'[.:\-]', '', mac).lower()
        if len(clean) >= 6:
            unique.add(clean[:6])

    to_fetch = [o for o in sorted(unique) if o not in _OUI_CACHE]
    if not to_fetch:
        return

    def _fetch_maclookup(oui6: str) -> tuple:
        """Fetch one OUI from maclookup.app; returns (oui6, vendor_or_None)."""
        fmt = f"{oui6[0:2]}:{oui6[2:4]}:{oui6[4:6]}"
        url = f"https://api.maclookup.app/v2/macs/{fmt}"
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": "cisco-audit-script/1.0"}
            )
            with urllib.request.urlopen(req, timeout=8, context=_OUI_SSL) as resp:
                data = json.loads(resp.read().decode("utf-8"))
                company = data.get("company", "") if data.get("found", False) else ""
                return oui6, company or ""
        except Exception:
            return oui6, None   # None signals complete failure (vs "" = not found)

    def _fetch_macvendors(oui6: str) -> tuple:
        """Fallback: fetch one OUI from macvendors.com with 429 retry."""
        fmt = f"{oui6[0:2]}:{oui6[2:4]}:{oui6[4:6]}"
        url = f"https://api.macvendors.com/{fmt}"
        for attempt in range(3):
            try:
                req = urllib.request.Request(
                    url, headers={"User-Agent": "cisco-audit-script/1.0"}
                )
                with urllib.request.urlopen(req, timeout=8, context=_OUI_SSL) as resp:
                    return oui6, resp.read().decode("utf-8").strip()
            except urllib.error.HTTPError as exc:
                if exc.code == 429:
                    time.sleep(2 ** attempt)   # 1s, 2s, 4s back-off
                else:
                    return oui6, ""
            except Exception:
                return oui6, ""
        return oui6, ""

    # ── Attempt 1: maclookup.app (15 concurrent threads) ───────────────────────────
    failed: list = []
    with _Prog(
        _Spin(spinner_name="dots2", style=T["spinner"]),
        _Txt(f"[{T['task_text']}]" + "{task.description}"),
        _Bar(bar_width=28, style=T["bar_empty"], complete_style=T["bar_fill"],
             finished_style=T["bar_done"]),
        _MofN(),
        _Time(),
        console=rich_console,
        transient=False,
    ) as prog:
        task = prog.add_task(
            "OUI lookup — api.maclookup.app (15 threads)", total=len(to_fetch)
        )
        with ThreadPoolExecutor(max_workers=15) as pool:
            futures = {pool.submit(_fetch_maclookup, o): o for o in to_fetch}
            for fut in as_completed(futures):
                oui6, vendor = fut.result()
                if vendor is None:
                    failed.append(oui6)
                else:
                    _OUI_CACHE[oui6] = vendor
                prog.advance(task)

    # ── Attempt 2: macvendors.com fallback (only if ALL maclookup failed) ───────
    if failed and len(failed) == len(to_fetch):
        rich_console.print(
            f"  [{T['label']}]OUI[/{T['label']}]  "
            "[yellow]maclookup.app unreachable — trying macvendors.com[/yellow]"
        )
        with _Prog(
            _Spin(spinner_name="dots2", style=T["spinner"]),
            _Txt(f"[{T['task_text']}]" + "{task.description}"),
            _Bar(bar_width=28, style=T["bar_empty"], complete_style=T["bar_fill"],
                 finished_style=T["bar_done"]),
            _MofN(),
            _Time(),
            console=rich_console,
            transient=False,
        ) as prog2:
            task2 = prog2.add_task(
                "OUI lookup — macvendors.com (10 threads)", total=len(failed)
            )
            with ThreadPoolExecutor(max_workers=10) as pool:
                futures2 = {pool.submit(_fetch_macvendors, o): o for o in failed}
                for fut in as_completed(futures2):
                    oui6, vendor = fut.result()
                    _OUI_CACHE[oui6] = vendor
                    prog2.advance(task2)

    for oui6 in to_fetch:
        _OUI_CACHE.setdefault(oui6, "")

    resolved = sum(1 for o in to_fetch if _OUI_CACHE.get(o))
    rich_console.print(
        f"  [{T['label']}]OUI[/{T['label']}]  "
        f"[{T['done_word']}]{resolved} of {len(to_fetch)} resolved[/{T['done_word']}]"
    )


def _vendor_to_device_type(vendor: str) -> str:
    """
    Map a raw OUI vendor name (returned by macvendors.com) to a device-type
    category.  Returns "" when the vendor name is ambiguous or unknown.

    This supplements the hardcoded _OUI_TABLE: the online API returns the
    actual registered company name, which lets us classify new/unrecognised
    OUIs without updating code.

    Notes on ambiguous vendors:
      - "Cisco" is skipped — makes switches, phones, APs, routers; CDP signals
        provide much better classification for Cisco devices.
      - "Dell" is skipped as a general Server match — Dell makes servers AND
        workstations; specific iDRAC OUIs in _OUI_TABLE handle Dell servers,
        while generic Dell NICs fall back to Workstation via the rule below.
      - "HP Inc." is skipped — makes printers AND consumer PCs; the
        printer-specific OUIs in _OUI_TABLE handle HP printers.
    """
    v = vendor.lower()
    # Vendor → device type keyword patterns (most-specific first)
    _VENDOR_MAP = [
        # ── Cameras ───────────────────────────────────────────────────────────
        (("axis comm",),                                          "IP Camera"),
        (("hanwha", "hikvision", "dahua", "avigilon",
          "bosch security", "pelco", "genetec"),                  "IP Camera"),
        # ── Video Conferencing ────────────────────────────────────────────────
        (("polycom", "tandberg", "cisco telepresence",
          "lifesize", "yealink"),                                 "Video Conf"),
        # ── Printers ──────────────────────────────────────────────────────────
        (("canon", "ricoh", "xerox", "konica", "minolta",
          "fuji xerox", "seiko epson", "brother ind",
          "sharp corp", "oki data", "lexmark",
          "kyocera", "toshiba tec", "murata"),                    "Printer"),
        # ── Virtual Machines ──────────────────────────────────────────────────
        (("vmware", "oracle virtualbox", "qemu", "parallels"),   "Virtual Machine"),
        # ── Wireless APs ──────────────────────────────────────────────────────
        (("aruba net", "ruckus", "ubiquiti",
          "meraki", "aerohive", "extreme network"),               "Wireless AP"),
        # ── Servers (specific enterprise brands only) ─────────────────────────
        (("hp enterprise", "hewlett packard enterprise",
          "super micro", "supermicro", "ibm"),                    "Server"),
        # ── Workstations ──────────────────────────────────────────────────────
        # Lenovo, ASUS, Acer make exclusively client devices on wired ports.
        # Apple on a wired corporate port is almost certainly a Mac.
        # Dell falls here too for generic NIC OUIs — iDRAC-specific entries
        # in _OUI_TABLE will already have caught Dell server BMC ports.
        (("lenovo",),                                             "Workstation"),
        (("asustek", "asus"),                                     "Workstation"),
        (("acer"),                                                "Workstation"),
        (("apple",),                                              "Workstation"),
        (("dell",),                                               "Workstation"),
        # ── Skip too-broad vendors ────────────────────────────────────────────
        (("cisco",),                                              ""),
        (("intel corp",),                                         ""),  # makes server + PC NICs
        (("realtek",),                                            ""),  # makes server + PC NICs
    ]
    for keywords, dtype in _VENDOR_MAP:
        if any(kw in v for kw in keywords):
            return dtype
    return ""


def device_type(mac: str, cdp_platform: str = "",
                cdp_capabilities: str = "", description: str = "",
                oui_vendor: str = "") -> Dict[str, str]:
    """
    Classify a network-connected device and explain how the decision was made.

    Returns a dict with three keys:

      "type"     — classification string, one of:
                     Switch/Router, Wireless AP, IP Phone, Video Conf,
                     Printer, IP Camera, Server, Virtual Machine,
                     Workstation, Unknown

      "source"   — how the type was determined:
                     CDP-Caps      (CDP capability codes — highest confidence)
                     CDP-Platform  (CDP platform string)
                     Description   (interface description keyword)
                     OUI-Live      (macvendors.com real-time lookup)
                     OUI-Fallback  (hardcoded _OUI_TABLE)
                     Unknown       (no signal matched)

      "conflict" — non-empty warning when the OUI independently suggests a
                     *different* device type than what CDP or the description
                     determined, e.g.:
                       "OUI suggests: Printer (HP Inc.)"
                     Empty string when there is no conflict or when the type
                     was already determined from the OUI itself.

    Classification priority (highest first):
      1. CDP capability codes  — 'P' = IP Phone; 'R'/'S' = Switch/Router.
      2. CDP platform string   — explicit platform name from CDP.
      3. Interface description — admin-entered keywords.
      4. Live OUI vendor name  — macvendors.com mapped via _vendor_to_device_type().
      5. Hardcoded _OUI_TABLE  — static fallback when offline.
    """
    caps       = set(cdp_capabilities.split()) if cdp_capabilities else set()
    plat_lower = cdp_platform.lower()  if cdp_platform  else ""
    desc_lower = description.lower()   if description   else ""

    # ── Pre-compute what the OUI alone would say (used for conflict detection)
    oui_inferred = ""
    if oui_vendor:
        oui_inferred = _vendor_to_device_type(oui_vendor)
    if not oui_inferred:                         # live lookup empty → try static table
        _clean = re.sub(r'[.:\-]', '', mac).lower()
        if len(_clean) >= 6:
            oui_inferred = _OUI_TABLE.get(_clean[:6], "")

    def _result(typ: str, source: str) -> Dict[str, str]:
        """Package a classification result and attach a conflict note if needed."""
        conflict = ""
        # Only flag a conflict when the type came from CDP or description AND
        # the OUI independently resolves to a meaningfully different type.
        if source not in ("OUI-Live", "OUI-Fallback", "Unknown"):
            if oui_inferred and oui_inferred != typ:
                vendor_note = f" ({oui_vendor})" if oui_vendor else ""
                conflict = f"OUI suggests: {oui_inferred}{vendor_note}"
        return {"type": typ, "source": source, "conflict": conflict}

    # ── 1. CDP capability codes ───────────────────────────────────────────────
    if "P" in caps:
        return _result("IP Phone", "CDP-Caps")
    if "R" in caps or "S" in caps:
        return _result("Switch/Router", "CDP-Caps")

    # ── 2. CDP platform string ────────────────────────────────────────────────
    if any(x in plat_lower for x in ("air-cap", "air-ap", "aironet", "meraki mr", " lap")):
        return _result("Wireless AP", "CDP-Platform")
    if any(x in plat_lower for x in ("polycom", "room kit", "tandberg",
                                      "telepresence", "webex room")):
        return _result("Video Conf", "CDP-Platform")
    if any(x in plat_lower for x in ("cisco spa", "ip phone", "sep")):
        return _result("IP Phone", "CDP-Platform")
    if any(x in plat_lower for x in ("ws-c", "catalyst", "nexus", "n5k", "n7k", "n9k",
                                      "c9", "isr", "asr", "wlc", "air-ct",
                                      "router", "switch")):
        return _result("Switch/Router", "CDP-Platform")

    # ── 3. Interface description keywords ────────────────────────────────────
    _desc_map = [
        (("printer", "copier", "mfp", " fax", "laserjet", "officejet"), "Printer"),
        (("camera", "cam ", "cctv", "nvr", "dvr"),                       "IP Camera"),
        (("phone", "voip", "ip-phone", "polycom", "cisco ip phone"),     "IP Phone"),
        (("server",),                                                      "Server"),
        ((" esx", "vcenter", "esxi", "hyperv", "hyper-v"),               "Virtual Machine"),
        (("workstation", "desktop", "laptop", "thin client", "pc ",
          " pc", "imac", "macbook"),                                       "Workstation"),
        ((" ap ", "wireless ap", "access point"),                          "Wireless AP"),
    ]
    for keywords, dtype in _desc_map:
        if any(kw in desc_lower for kw in keywords):
            return _result(dtype, "Description")

    # ── 4. Live OUI vendor name (from macvendors.com) ─────────────────────────
    if oui_vendor:
        inferred = _vendor_to_device_type(oui_vendor)
        if inferred:
            return _result(inferred, "OUI-Live")

    # ── 5. Hardcoded OUI table (offline fallback) ────────────────────────────
    clean = re.sub(r'[.:\-]', '', mac).lower()
    if len(clean) >= 6:
        oui = clean[:6]
        if oui in _OUI_TABLE:
            return _result(_OUI_TABLE[oui], "OUI-Fallback")

    return _result("Unknown", "Unknown")


def is_uplink(description: str, cdp_platform: str,

               port: str = "", mac_count: int = 0,

               cdp_neighbour: str = "", cdp_capabilities: str = "") -> str:

    """

    Return "Yes", "Possible", or "No" indicating uplink confidence.



    Strong signals — any one alone returns "Yes" immediately:

      * CDP platform is a known network-infrastructure device

        (WS-C, Catalyst, Nexus, C9k, ISR, ASR, WLC, AIR-CT, router, switch).

      * CDP capability codes include 'R' (Router) or 'S' (Switch) — the

        gold standard: the remote device itself declares what it is.

      * Description contains the word "switch" (case-insensitive whole word) —

        admins label downstream links as "MONDTSW08_G3/3_8th_Floor_Switch".

      * Description ends with what looks like a remote-port reference, e.g.

        "MONDTSW10_G1/0/1" or "<<Link To MONDT101_G0/0/1>>" — admins

        record the remote end only for infrastructure links.



    Moderate signals — each scores 1 point.  Score >= 2 → "Yes",

                       score == 1 → "Possible", score 0 → "No":

      * Interface is 10G or faster (Te/25G/40G/100G).

      * Description contains a routing/infra keyword

        (uplink, trunk, core, backbone, dist, distribution, wan, isp, p2p).

      * A CDP neighbour is present AND is NOT a known endpoint — we skip

        this point when capabilities show a phone ('P') or wireless AP

        ('T' + 'B' = Trans-Bridge mode, typical of Cisco lightweight APs).

      * MAC count > 5 — trunk-like behaviour.

    """

    desc_lower     = description.lower()  if description   else ""

    platform_lower = cdp_platform.lower() if cdp_platform  else ""

    caps           = set(cdp_capabilities.split()) if cdp_capabilities else set()



    # ── Strong signals (any one -> "Yes") ────────────────────────────────────

    INFRA_PLATFORMS = ["ws-c", "catalyst", "nexus", "c9", "isr", "asr",

                       "router", "switch", "n5k", "n7k", "n9k",

                       "air-ct", "wlc"]  # WLC / AIR-CT = Wireless LAN Controller

    if any(x in platform_lower for x in INFRA_PLATFORMS):

        return "Yes"



    # CDP capability R (Router) or S (Switch) declared by the remote device itself

    if "R" in caps or "S" in caps:

        return "Yes"



    # Description contains the word "switch" (whole word, case-insensitive)

    if re.search(r'\bswitch\b', desc_lower):

        return "Yes"



    # ── Moderate signals: scored ──────────────────────────────────────────

    score = 0



    # 10G+ port speed (access ports are almost never this fast)

    if port and any(port.startswith(p) for p in ("Te", "Twe", "Fo", "Hu")):

        score += 1



    # Description routing/infra keyword

    if any(x in desc_lower for x in ["uplink", "trunk", "core", "backbone",

                                      "dist", "distribution", "wan", "isp", "p2p"]):

        score += 1


    # CDP neighbour present, but only if NOT a confirmed endpoint.

    # Endpoints:

    #   P        = Phone

    #   T + B    = Cisco lightweight AP (Trans-Bridge + Source-Route-Bridge)

    #   H alone  = generic host (PC, server, video-conf unit, etc.)

    #              NOTE: a WLC also reports H, but it's already caught by

    #              INFRA_PLATFORMS above so we never reach this point for a WLC.

    is_endpoint = (

        ("P" in caps) or

        ("T" in caps and "B" in caps) or

        (caps == {"H"})   # pure host with no router/switch capability

    )

    if cdp_neighbour and not is_endpoint:

        score += 1



    # Multiple MACs suggest a trunk or downstream switch

    if mac_count > 5:

        score += 1



    if score >= 2:

        return "Yes"

    elif score == 1:

        return "Possible"

    return "No"



def parse_location_details(location: str) -> Tuple[str, str]:

    """

    Extract a (building, floor) tuple from a free-form SNMP location string.



    Building — everything before the first '.' or ',' in the string, with

               any leading non-alphanumeric characters stripped.

               e.g. "1200 McGill. 10th. Floor..." -> "1200 McGill"



    Floor    — first integer followed by an optional ordinal suffix and

               the word "floor" or abbreviation "fl".

               Also recognises "first floor", "ground", "basement".



    Returns ("", "") for an empty/missing location.  Never raises.

    """

    if not location:

        return "", ""



    building = ""

    floor = ""



    # Floor: look for digits followed by optional ordinal suffix then 'floor' or 'fl'

    floor_match = re.search(r'(\d+)\s*(?:st|nd|rd|th)?\.?\s*[Ff]l(?:oor)?', location, re.IGNORECASE)

    if floor_match:

        floor = floor_match.group(1)

    elif re.search(r'\bfirst\s+floor\b', location, re.IGNORECASE):

        floor = "1"

    elif re.search(r'\bground\b|\bgrnd\b', location, re.IGNORECASE):

        floor = "Ground"

    elif re.search(r'\bbasement\b', location, re.IGNORECASE):

        floor = "Basement"



    # Building: take everything before the first '.' or ',' as the building/address

    # e.g. "1200 McGill. 10th. Floor ..." -> "1200 McGill"

    first_delim = re.search(r'[.,]', location)

    if first_delim:

        building = location[:first_delim.start()].strip()

    else:

        building = location.strip()

    # Strip leading and trailing non-alphanumeric junk (e.g. "<<", ">>", spaces)

    building = re.sub(r'^[^A-Za-z0-9]+', '', building).strip()

    building = re.sub(r'[^A-Za-z0-9]+$', '', building).strip()



    return building, floor



def parse_switch_section(lines: List[str]) -> Dict:

    """

    Parse all show-command output for one switch section.



    Line-by-line state machine.  Transitions to a new state whenever a

    show-command prompt is recognised, then collects data in that state.



    States:

      hostname_snmp  — hostname + SNMP location string

      mac_table      — MAC address table (vlan / mac / port)

      int_desc       — interface descriptions and up/down status

      arp            — IP to MAC mapping

      cdp_neighbors  — CDP neighbour hostname + platform per port

      version        — switch model from 'sh version' output



    Returns a dict with keys:

      hostname, snmp_location, mac_table, int_desc, int_status,

      cdp_neighbors, cdp_platforms, arp_table, switch_model

    """

    hostname = ""

    snmp_location = ""

    mac_table: Dict[str, List[Tuple[str, str]]] = {}  # port -> [(vlan, mac)]

    int_desc: Dict[str, str] = {}

    int_status: Dict[str, str] = {}

    cdp_neighbors: Dict[str, str] = {}

    cdp_platforms: Dict[str, str] = {}

    cdp_capabilities: Dict[str, str] = {}  # port -> capability string, e.g. "R S", "H P", "T B I"

    arp_table: Dict[str, str] = {}  # mac -> ip

    switch_model = ""

    # --- New metadata fields extracted from 'sh version' ---

    ios_version   = ""  # e.g. "15.1(2)SY" or "03.06.10.E"

    serial_number = ""  # chassis serial (Processor board ID / System Serial Number)

    uptime        = ""  # e.g. "1 year, 6 weeks, 3 days"

    last_restart  = ""  # e.g. "21:25:18 EST Thu Jan 16 2025"

    # --- Interface traffic stats from 'show interfaces stat' ---

    # port -> {pkts_in: int, pkts_out: int, disabled: bool}

    int_stats: Dict[str, Dict] = {}

    _int_stats_cur = ""   # interface currently being accumulated

    # List of dicts: {subnet, interface, route_type}

    # route_type: 'C' = connected network, 'L' = local (switch's own IP on that interface)

    route_connected: List[Dict] = []

    state = "none"

    cdp_pending_device = ""  # holds device ID when CDP entry spans two lines

    

    for raw_line in lines:

        line = raw_line.strip()

        if not line:

            if state == "cdp_neighbors":

                cdp_pending_device = ""  # blank line resets pending device

            if state == "int_stats":

                _int_stats_cur = ""      # blank line ends current interface block

            continue

        

        # ── State transitions ────────────────────────────────────────────────

        # Each show-command prompt line switches the parser into the matching

        # state.  A "<hostname>#" prompt resets to the idle state.

        #

        # NOTE: Cisco allows abbreviated commands, e.g. "sh version" instead

        # of "show version", so we check for both forms where needed.



        if "show running-config | include (hostname|snmp" in line:

            state = "hostname_snmp"

            continue



        elif "show mac address-table" in line:

            state = "mac_table"

            continue



        elif "show interface description" in line:

            state = "int_desc"

            continue



        elif "show arp" in line:

            state = "arp"

            continue



        elif "show cdp neighbors" in line:

            state = "cdp_neighbors"

            continue



        elif "show version" in line or "sh version" in line:

            # BUG FIX: audit files use abbreviated 'sh version', not 'show version',

            # so the model was never extracted before this fix.

            state = "version"

            continue



        elif "show ip route connected" in line or "sh ip route connected" in line:

            state = "route_connected"

            continue



        elif "show interfaces stat" in line:

            state = "int_stats"

            _int_stats_cur = ""

            continue



        elif hostname and line.startswith(hostname + "#"):

            # BUG FIX: when hostname is still empty, startswith("") matches

            # every line and would reset state prematurely.  Only match the

            # prompt once hostname is known; fall back to bare '#' ending.

            state = "none"

            continue



        elif not hostname and line.endswith("#"):

            state = "none"

            continue

        

        # Parse based on state

        if state == "hostname_snmp":

            if line.startswith("hostname"):

                hostname = line.split()[1]

            elif "snmp-server location" in line or "snmp-location" in line:

                snmp_location = line.split("snmp-server location", 1)[1].strip() if "snmp-server location" in line else line.split("snmp-location", 1)[1].strip()

        

        elif state == "mac_table":

            # Format: [*] vlan   mac-address   type  learn  age   port

            # The '*' primary-entry marker is optional - strip it before parsing.

            parts = line.split()

            if parts and parts[0] == '*':

                parts = parts[1:]  # drop the '*' marker

            if len(parts) >= 4 and parts[0].isdigit():

                vlan = parts[0]

                mac = normalize_mac(parts[1])

                port = normalize_port(parts[-1])

                # Skip special/internal entries with pseudo-ports

                if port in ("Router", "Switch", "CPU", "Drop"):

                    continue

                if port not in mac_table:

                    mac_table[port] = []

                mac_table[port].append((vlan, mac))

        

        elif state == "int_desc":

            # Format: Interface  Status  Protocol  Description

            # Accept any Cisco interface prefix so Nexus (Eth), Catalyst 9k (Twe/Hu/Tw/Fi/Ap)

            # and older IOS (Gi/Fa/Te/Po) all get parsed.  normalize_port() handles

            # the full-name -> abbreviation conversion for all of them.

            if re.match(r'^(GigabitEthernet|FastEthernet|TenGigabitEthernet|TwentyFiveGigE'

                        r'|TwoGigabitEthernet|FiveGigabitEthernet|HundredGigE|FortyGigabitEthernet'

                        r'|AppGigabitEthernet|Ethernet|Management|mgmt'

                        r'|Gi|Fa|Te|Twe|Tw|Fi|Hu|Fo|Ap|Eth|Ma|Po)', line):

                parts = line.split(None, 3)

                if len(parts) >= 2:

                    port = normalize_port(parts[0])

                    status = parts[1]  # up/down

                    desc = parts[3] if len(parts) >= 4 else ""

                    int_status[port] = status

                    int_desc[port] = desc

        

        elif state == "arp":

            # Format: Protocol  Address  Age  Hardware_Addr  Type  Interface

            # Internet  10.14.113.11  34  000b.aa30.3855  ARPA  Vlan513

            parts = line.split()

            if len(parts) >= 5 and parts[0] == "Internet" and parts[3] != "Incomplete":

                ip = parts[1]

                mac = normalize_mac(parts[3])

                arp_table[mac] = ip

        

        elif state == "cdp_neighbors":

            # CDP entries appear in two formats:

            #   Multi-line (long FQDN): device_id alone on line 1,

            #                           then "Port_abbr port_num holdtme caps platform remote_port" on indented line 2

            #   Single-line (short name): "device_id  Port_abbr port_num holdtme caps platform remote_port"

            #

            # Port abbreviations used in local-interface field: Ten, Gig, Gi, Fa, Po



            # Skip header/legend lines

            if any(x in line for x in ["Device ID", "Capability Codes", "Local Intrfce", "Holdtme",

                                        "Trans Bridge", "Source Route", "Repeater", "Two-port"]):

                continue



            parts = line.split()

            if not parts:

                cdp_pending_device = ""

                continue



            # Full set of Cisco interface-type prefixes seen in CDP local-port fields.

            # Covers: IOS classic, IOS-XE Catalyst 9k, NX-OS Nexus, ISR routers.

            PORT_ABBREVS = {

                # Full names (show cdp neighbors detail sometimes uses them)

                "GigabitEthernet", "FastEthernet", "TenGigabitEthernet",

                "TwentyFiveGigE", "TwoGigabitEthernet", "FiveGigabitEthernet",

                "HundredGigE", "FortyGigabitEthernet", "AppGigabitEthernet",

                "Ethernet",   # Nexus

                "Management", "mgmt",

                # Short abbreviations used in 'show cdp neighbors' output

                "Gig", "Gi", "Fa", "Ten", "Te",

                "Twe", "Tw", "Fi", "Hu", "Fo", "Ap", "Eth", "Ma",

                "Po",   # Port-channel

            }



            if parts[0] in PORT_ABBREVS:

                # ---- Continuation line of a multi-line CDP entry ----

                # Format: port_abbr port_num holdtme [caps...] platform remote_abbr remote_num

                # e.g.:  Ten 1/4  168  R S  WS-C6504-  Ten 1/4

                if cdp_pending_device and len(parts) >= 4:

                    port = normalize_port(parts[0] + parts[1])  # e.g. Ten1/4 -> Te1/4

                    # Platform = parts[-3]  (last two tokens are remote port abbr+num)

                    platform = parts[-3] if len(parts) >= 5 else ""

                    # Capabilities = letters between holdtime and platform

                    # parts: [abbr, num, holdtime, cap, cap, ..., platform, rem_abbr, rem_num]

                    caps = " ".join(parts[3:-3]) if len(parts) >= 7 else ""

                    cdp_neighbors[port]     = short_hostname(cdp_pending_device)

                    cdp_platforms[port]     = platform

                    cdp_capabilities[port]  = caps

                cdp_pending_device = ""



            elif len(parts) == 1:

                # ---- Device ID on its own line (FQDN too long for one line) ----

                cdp_pending_device = parts[0]



            else:

                # ---- Single-line entry: device_id  port_abbr  port_num  ... ----

                # e.g.: mcgwc01  Gig 3/5  161  H  AIR-CT552  Ten 0/0/1

                device_id = parts[0]

                cdp_pending_device = ""

                if len(parts) >= 3 and parts[1] in PORT_ABBREVS:

                    port = normalize_port(parts[1] + parts[2])  # e.g. Gig3/5 -> Gi3/5

                    platform = parts[-3] if len(parts) >= 6 else ""

                    # Capabilities sit between holdtime (parts[3]) and platform (parts[-3])

                    caps = " ".join(parts[4:-3]) if len(parts) >= 8 else ""

                    cdp_neighbors[port]     = short_hostname(device_id)

                    cdp_platforms[port]     = platform

                    cdp_capabilities[port]  = caps

        

        elif state == "version":

            # IOS version string — first match of "Version X.Y.Z" in the banner lines

            if not ios_version:

                m = re.search(r'\bVersion\s+([\d().A-Za-z/-]+)', line)

                if m:

                    ios_version = m.group(1)



            # Uptime — "<hostname> uptime is <duration>"

            if not uptime and "uptime is" in line.lower():

                m = re.search(r'uptime is\s+(.+)', line, re.IGNORECASE)

                if m:

                    uptime = m.group(1).strip()



            # Last restart — "System restarted at <timestamp>"

            if not last_restart and "System restarted at" in line:

                m = re.search(r'System restarted at\s+(.+)', line)

                if m:

                    last_restart = m.group(1).strip()



            # Serial number — two formats:

            #   Old IOS : "Processor board ID FOX1415GTFT"

            #   IOS-XE  : "System Serial Number               : FDO2013Q07L"

            if not serial_number:

                m = re.search(r'Processor board ID\s+(\S+)', line)

                if m:

                    serial_number = m.group(1)

            if "System Serial Number" in line:

                m = re.search(r'System Serial Number\s*:\s*(\S+)', line)

                if m:

                    serial_number = m.group(1)  # IOS-XE value overrides if both present



            # Explicit model line (IOS-XE): "Model Number        : WS-C3650-48PQ"

            if "Model number" in line or "Model Number" in line:

                switch_model = line.split(":")[-1].strip()



            # Fallback model: "cisco WS-C6504-E (R7000) processor..." (old IOS)
            # Only match the hardware synopsis line, not the "Cisco IOS Software..." banner.
            # The hardware line always contains the word "processor".
            if not switch_model and "processor" in line.lower() and re.match(r'^cisco\s+(\S+)', line.strip(), re.IGNORECASE):

                m = re.match(r'^cisco\s+(\S+)', line.strip(), re.IGNORECASE)

                if m:

                    switch_model = m.group(1)



        elif state == "route_connected":

            # Lines of interest:

            #   C     10.64.111.0/24 is directly connected, Vlan111

            #   L     10.64.111.3/32 is directly connected, Vlan111

            # Skip header/legend lines; only grab C and L entries.

            if line and line[0] in ("C", "L"):

                m = re.match(r'^([CL])\s+(\S+)\s+is directly connected,\s+(\S+)', line)

                if m:

                    route_type = "Connected" if m.group(1) == "C" else "Local"

                    route_connected.append({

                        "subnet":     m.group(2),

                        "interface":  normalize_port(m.group(3)),

                        "route_type": route_type,

                    })



        elif state == "int_stats":

            # Two relevant line types:

            #

            # 1. "Interface GigabitEthernet2/4 is disabled"

            #    → mark port as disabled with zero traffic

            #

            # 2. Interface header (just the name alone on a line):

            #    "GigabitEthernet2/3"  or  "TenGigabitEthernet1/4"

            #    → set _int_stats_cur so the Total row knows which port it belongs to

            #

            # 3. Total line: "               Total  <pkts_in>  <chars_in>  <pkts_out>  <chars_out>"

            #    → store traffic counters for the current interface

            #

            # Lines like "% Incomplete command" or header rows are silently ignored.



            m_disabled = re.match(r'^Interface\s+(\S+)\s+is disabled', line)

            if m_disabled:

                port = normalize_port(m_disabled.group(1))

                int_stats[port] = {"pkts_in": 0, "pkts_out": 0, "disabled": True}

                _int_stats_cur = ""



            elif re.match(

                r'^(GigabitEthernet|FastEthernet|TenGigabitEthernet|TwentyFiveGigE'

                r'|TwoGigabitEthernet|FiveGigabitEthernet|HundredGigE|FortyGigabitEthernet'

                r'|AppGigabitEthernet|Ethernet|Management|Loopback|Vlan|Tunnel|Port-channel'

                r'|Gi|Fa|Te|Twe|Tw|Fi|Hu|Fo|Ap|Eth|Ma|Po|Lo|Vl)\S*$', line):

                _int_stats_cur = normalize_port(line.strip())



            elif line.strip().startswith("Total") and _int_stats_cur:

                parts = line.split()

                # Total row: Total <pkts_in> <chars_in> <pkts_out> <chars_out>

                if len(parts) >= 5:

                    try:

                        pkts_in  = int(parts[1])

                        pkts_out = int(parts[3])

                        int_stats[_int_stats_cur] = {

                            "pkts_in":  pkts_in,

                            "pkts_out": pkts_out,

                            "disabled": False,

                        }

                    except ValueError:

                        pass  # malformed line — skip gracefully



    return {

        "hostname": hostname,

        "snmp_location": snmp_location,

        "mac_table": mac_table,

        "int_desc": int_desc,

        "int_status": int_status,

        "cdp_neighbors":   cdp_neighbors,

        "cdp_platforms":   cdp_platforms,

        "cdp_capabilities": cdp_capabilities,

        "arp_table":       arp_table,

        "switch_model":    switch_model,

        "ios_version":     ios_version,

        "serial_number":   serial_number,

        "uptime":          uptime,

        "last_restart":    last_restart,

        "int_stats":       int_stats,

        "route_connected": route_connected,

    }



# ── Helpers ──────────────────────────────────────────────────────────────────



def _build_switch_rows(parsed: Dict) -> Tuple[list, str, str]:

    """

    Expand a parsed switch dict into a flat list of Excel row dicts.



    Each row = one (MAC address, port) combination.  Common switch-level fields

    (hostname, model, building, floor, SNMP location) are duplicated across all

    rows for that switch so every row is self-contained.



    Returns

    -------

    (rows, building, floor)

      rows     : list of dicts, one per MAC entry, ready for pd.DataFrame

      building : parsed building name (may be empty string)

      floor    : parsed floor string  (may be empty string)



    Returning building/floor avoids calling parse_location_details a second

    time in process_file for the done_infos summary table.

    """

    building, floor = parse_location_details(parsed["snmp_location"])

    mac_counts = Counter()

    for port in parsed["mac_table"]:

        mac_counts[port] = len(parsed["mac_table"][port])



    rows: list = []

    for port, entries in parsed["mac_table"].items():

        desc         = parsed["int_desc"].get(port, "")

        status       = parsed["int_status"].get(port, "")

        cdp          = parsed["cdp_neighbors"].get(port, "")

        cdp_platform = parsed["cdp_platforms"].get(port, "")

        cdp_cap      = parsed["cdp_capabilities"].get(port, "")

        mac_count    = mac_counts[port]

        uplink_flag  = is_uplink(desc, cdp_platform,

                                  port=port, mac_count=mac_count,

                                  cdp_neighbour=cdp, cdp_capabilities=cdp_cap)

        stats        = parsed["int_stats"].get(port, {})

        pkts_in      = stats.get("pkts_in", "")

        pkts_out     = stats.get("pkts_out", "")

        # "Traffic Active" = Yes if any packets seen; "-" if no stats available

        if stats:

            traffic_active = "Yes" if (stats["pkts_in"] + stats["pkts_out"]) > 0 else "No"

        else:

            traffic_active = "-"

        for vlan, mac in entries:

            ip_address  = parsed["arp_table"].get(mac, "")
            # "Has IP" = Yes when this MAC appears in the ARP table; devices
            # that always carry the same IP (even across reboots) are strong
            # candidates for statically-assigned addresses.
            has_ip      = "Yes" if ip_address else "No"
            # Look up OUI vendor name: uses in-memory cache populated before
            # the main loop — no extra network delay per row.
            clean_mac   = re.sub(r'[.:\-]', '', mac).lower()
            oui6        = clean_mac[:6] if len(clean_mac) >= 6 else ""
            oui_vendor  = _lookup_oui(oui6) if oui6 else ""
            dt          = device_type(mac,
                                      cdp_platform=cdp_platform,
                                      cdp_capabilities=cdp_cap,
                                      description=desc,
                                      oui_vendor=oui_vendor)

            rows.append({

                "Switch hostname":       parsed["hostname"],

                "Switch Model":          parsed["switch_model"],

                "Building":              building,

                "Floor":                 floor,

                "snmp location":         parsed["snmp_location"],

                "Interface Status":      status,

                "interface number":      port,

                "interface description": desc,

                "CDP Neighbour":         cdp,

                "CDP Platform":          cdp_platform,

                "Uplink":                uplink_flag,

                "Device Type":           dt["type"],

                "Classification Source": dt["source"],

                "Type Conflict":         dt["conflict"],

                "VLAN":                  vlan,

                "mac address":           mac,

                "IP Address":            ip_address,

                "Has IP":                has_ip,

                "Vendor (OUI)":          oui_vendor,

                "MAC Count":          mac_count,

                "Traffic Active":     traffic_active,

                "Pkts In":            pkts_in,

                "Pkts Out":           pkts_out,

            })

    # Return building and floor alongside rows so the caller doesn't need to
    # call parse_location_details a second time for the summary table.
    return rows, building, floor



# ── Rich live-table builder ───────────────────────────────────────────────────



def _make_switch_table(rows_done: list) -> Table:

    """
    Build the per-switch results Table printed after the progress bar.

    One row per switch, showing MACs, records, building, floor and model.
    Colours are taken from the global theme dict T so dark/light terminals
    both render cleanly.
    """

    tbl = Table(

        box=box.SIMPLE_HEAD,

        show_header=True,

        header_style=T["tbl_hdr"],

        expand=False,

        show_edge=False,

    )

    tbl.add_column("  ",       width=3,  justify="center")

    tbl.add_column("Switch",   style=T["col_hostname"], min_width=18)

    tbl.add_column("MACs",     style=T["col_macs"],     justify="right", width=6)

    tbl.add_column("Records",  style=T["col_records"],  justify="right", width=8)

    tbl.add_column("Building", style=T["col_building"], max_width=22, no_wrap=True)

    tbl.add_column("Floor",    style=T["col_floor"],    width=7)

    tbl.add_column("Model",    style=T["col_model"],    max_width=14, no_wrap=True)

    for info in rows_done:

        # Convert float floor (e.g. 10.0) to a clean int string; hide NaN/empty
        floor_val = info["floor"]

        if isinstance(floor_val, float) and not math.isnan(floor_val):

            floor_str = str(int(floor_val))

        else:

            floor_str = str(floor_val) if floor_val else f"[{T['nil']}]-[/{T['nil']}]"

        building  = (info["building"][:20] + "…") if info["building"] and len(info["building"]) > 21 else (info["building"] or f"[{T['nil']}]-[/{T['nil']}]")

        model_str = (info["model"][:12] + "…")    if info["model"]    and len(info["model"])    > 13 else (info["model"]    or f"[{T['nil']}]-[/{T['nil']}]")

        tbl.add_row(

            f"[{T['check']}]\u2714[/{T['check']}]",

            info["hostname"],

            str(info["mac_count"]),

            str(info["record_count"]),

            building,

            floor_str,

            model_str,

        )

    return tbl



# ── HTML5 report generator ───────────────────────────────────────────────────


def _write_html_report(
    df: "pd.DataFrame",
    df_subnets: "pd.DataFrame",
    df_devices: "pd.DataFrame",
    html_path: str,
    input_file: str = "",
) -> None:
    """Generate a self-contained interactive HTML5 migration dashboard."""
    import json as _json
    from datetime import datetime as _dt

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _safe(v):
        if v is None:
            return ""
        if isinstance(v, float) and v != v:
            return ""
        return v

    def _df_rows(frame: "pd.DataFrame") -> list:
        return [{k: _safe(v) for k, v in row.items()} for _, row in frame.iterrows()]

    def _filter_opts(frame: "pd.DataFrame", col: str) -> list:
        if frame.empty or col not in frame.columns:
            return []
        return sorted({str(v) for v in frame[col].dropna() if str(v).strip()}, key=str.lower)

    # ── Charts data ─────────────────────────────────────────────────────────

    dtype_counts = df["Device Type"].value_counts().to_dict() if not df.empty and "Device Type" in df.columns else {}
    upl_counts   = df["Uplink"].value_counts().to_dict()      if not df.empty and "Uplink"      in df.columns else {}

    if not df_devices.empty and "Hostname" in df_devices.columns:
        _sw = df_devices[["Hostname","Total MACs"]].sort_values("Total MACs", ascending=False)
        sw_labels = _sw["Hostname"].tolist()
        sw_values = [int(v) for v in _sw["Total MACs"].fillna(0)]
    else:
        sw_labels, sw_values = [], []

    if not df_devices.empty and "Building" in df_devices.columns:
        _b = (df_devices[df_devices["Building"].notna() & (df_devices["Building"] != "")]
              .groupby("Building")["Total MACs"].sum()
              .sort_values(ascending=False).head(20))
        bld_labels = _b.index.tolist()
        bld_values = [int(v) for v in _b.values]
    else:
        bld_labels, bld_values = [], []

    if not df_devices.empty and "Active Interfaces" in df_devices.columns:
        _id = df_devices[["Hostname","Active Interfaces","Inactive Interfaces","Disabled Interfaces"]].head(30)
        iface_labels   = _id["Hostname"].tolist()
        iface_active   = [int(v) for v in _id["Active Interfaces"].fillna(0)]
        iface_inactive = [int(v) for v in _id["Inactive Interfaces"].fillna(0)]
        iface_disabled = [int(v) for v in _id["Disabled Interfaces"].fillna(0)]
    else:
        iface_labels = iface_active = iface_inactive = iface_disabled = []  # type: ignore[assignment]

    # ── KPI numbers ─────────────────────────────────────────────────────────

    total_macs      = int(df_devices["Total MACs"].sum())          if not df_devices.empty else 0
    total_switches  = len(df_devices)                               if not df_devices.empty else 0
    total_subnets_c = (len(df_subnets[df_subnets["Route Type"] == "Connected"])
                       if not df_subnets.empty and "Route Type" in df_subnets.columns else 0)
    total_buildings = (int(df_devices["Building"].dropna().apply(str)
                           .replace("", "<blank>").nunique())
                       if not df_devices.empty else 0)
    total_conflicts = int((df["Type Conflict"] != "").sum()) if not df.empty and "Type Conflict" in df.columns else 0
    total_uplinks   = int((df["Uplink"] == "Yes").sum())     if not df.empty and "Uplink" in df.columns else 0
    no_ip_pct       = (round(100 * int((df["Has IP"] == "No").sum()) / max(len(df), 1))
                       if not df.empty and "Has IP" in df.columns else 0)

    # ── Full table data ──────────────────────────────────────────────────────

    mac_cols    = list(df.columns)         if not df.empty         else []
    subnet_cols = list(df_subnets.columns) if not df_subnets.empty else []
    device_cols = list(df_devices.columns) if not df_devices.empty else []

    # ── Topology data ────────────────────────────────────────────────────────
    _topo_pal  = ["#2563eb","#10b981","#f59e0b","#6366f1","#ec4899","#8b5cf6",
                  "#ef4444","#14b8a6","#f97316","#0891b2","#84cc16","#db2777"]
    _bld_list  = sorted({
        str(row.get("Building", "") or "")
        for _, row in df_devices.iterrows()
    } if not df_devices.empty else set())
    _bld_list  = [b for b in _bld_list if b]
    _bld_color = {b: _topo_pal[i % len(_topo_pal)] for i, b in enumerate(_bld_list)}

    _topo_nodes: list = []
    _known_sw:   set  = set()
    if not df_devices.empty:
        for _, _row in df_devices.iterrows():
            _hn  = str(_row.get("Hostname", "") or "")
            if not _hn:
                continue
            _bld = str(_row.get("Building", "") or "")
            _macs_cnt = int(_row.get("Total MACs", 0) or 0)
            _topo_nodes.append({
                "id":    _hn,
                "label": _hn,
                "title": (
                    f"<b>{_hn}</b><br>"
                    f"Model: {_row.get('Model', '\u2014')}<br>"
                    f"Building: {_bld}<br>"
                    f"Floor: {_row.get('Floor', '\u2014')}<br>"
                    f"Total MACs: {_macs_cnt}"
                ),
                "group":   _bld or "Unknown",
                "color":   {"background": _bld_color.get(_bld, "#94a3b8"),
                            "border":     "#0f172a",
                            "highlight":  {"background": "#fef08a", "border": "#0f172a"}},
                "font":    {"color": "#fff", "size": 11},
                "shape":   "box",
                "macs":    _macs_cnt,
                "model":   str(_row.get("Model", "") or ""),
                "serial":  str(_row.get("Serial Number", "") or ""),
                "ios":     str(_row.get("IOS Version", "") or ""),
                "uptime":  str(_row.get("Uptime", "") or ""),
                "building": _bld,
                "floor":   str(_row.get("Floor", "") or ""),
                "totalIf":    int(_row.get("Total Interfaces",    0) or 0),
                "activeIf":   int(_row.get("Active Interfaces",   0) or 0),
                "inactiveIf": int(_row.get("Inactive Interfaces", 0) or 0),
                "disabledIf": int(_row.get("Disabled Interfaces", 0) or 0),
                "dtypes": {
                    "Switch/Router": int(_row.get("Switches/Routers", 0) or 0),
                    "Wireless AP":   int(_row.get("Wireless APs",     0) or 0),
                    "IP Phone":      int(_row.get("IP Phones",        0) or 0),
                    "Printer":       int(_row.get("Printers",         0) or 0),
                    "Workstation":   int(_row.get("Workstations",     0) or 0),
                    "Unknown":       int(_row.get("Unknown Devices",  0) or 0),
                },
            })
            _known_sw.add(_hn)

    _topo_edges: list = []
    _seen_pairs: set  = set()
    if not df.empty and "CDP Neighbour" in df.columns and "Uplink" in df.columns:
        for _, _row in df.iterrows():
            _src = str(_row.get("Switch hostname", "") or "")
            _dst = str(_row.get("CDP Neighbour",   "") or "").strip()
            if not _src or not _dst:
                continue
            if (_row.get("Uplink", "") != "Yes"
                    and _row.get("Device Type", "") != "Switch/Router"):
                continue
            _pair = tuple(sorted([_src, _dst]))
            if _pair in _seen_pairs:
                continue
            _seen_pairs.add(_pair)
            if _dst not in _known_sw:
                _topo_nodes.append({
                    "id":    _dst,
                    "label": _dst,
                    "title": "External / not audited in this run",
                    "group": "External",
                    "color": {"background": "#94a3b8", "border": "#475569",
                              "highlight":  {"background": "#fef08a", "border": "#475569"}},
                    "font":    {"color": "#fff", "size": 11},
                    "shape":   "ellipse",
                    "macs":    0, "model": "\u2014", "building": "External", "floor": "",
                })
                _known_sw.add(_dst)
            _vlan_str = str(_row.get("VLAN", "") or "")
            _iface_str = str(_row.get("interface number", "") or "")
            _topo_edges.append({
                "from":  _src,
                "to":    _dst,
                "label": _iface_str,
                "vlan":  _vlan_str,
                "title": (
                    f"<b>{_src} \u2194 {_dst}</b><br>"
                    f"Interface: {_iface_str}<br>"
                    + (f"VLAN: {_vlan_str}" if _vlan_str else "")
                ),
                "color": {"color": "#64748b", "highlight": "#2563eb"},
                "width": 2,
            })

    # ── History data ─────────────────────────────────────────────────────────
    try:
        _history = _db_load_history()
    except Exception:
        _history = []

    # ── Sites data ────────────────────────────────────────────────────────────
    try:
        _sites_list = _db_load_sites()
    except Exception:
        _sites_list = []
    try:
        _device_site_map = _db_load_device_site_map()
    except Exception:
        _device_site_map = {}

    # Always append the current run as the last entry so it can be compared
    # against any previous run directly from the HTML without needing to re-run.
    _cur_ports = []
    if not df.empty:
        for _, _r in df.iterrows():
            _cur_ports.append({
                "switch":        str(_r.get("Switch hostname",    "") or ""),
                "interface":     str(_r.get("interface number",   "") or ""),
                "vlan":          str(_r.get("VLAN",               "") or ""),
                "mac_address":   str(_r.get("mac address",        "") or ""),
                "ip_address":    str(_r.get("IP Address",         "") or ""),
                "device_type":   str(_r.get("Device Type",        "") or ""),
                "status":        str(_r.get("Interface Status",   "") or ""),
                "cdp_neighbour": str(_r.get("CDP Neighbour",      "") or ""),
                "description":   str(_r.get("interface description", "") or ""),
                "building":      str(_r.get("Building",           "") or ""),
                "floor":         str(_r.get("Floor",              "") or ""),
            })
    _history.append({
        "run_id":      "current",
        "run_at":      _dt.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "source_file": os.path.basename(input_file) if input_file else "current",
        "ports":       _cur_ports,
    })

    payload = {
        "meta": {
            "generated":      _dt.now().strftime("%Y-%m-%d %H:%M"),
            "source":         input_file,
            "totalMacs":      total_macs,
            "totalSwitches":  total_switches,
            "totalSubnets":   total_subnets_c,
            "totalBuildings": total_buildings,
            "totalConflicts": total_conflicts,
            "totalUplinks":   total_uplinks,
            "noIpPct":        no_ip_pct,
        },
        "charts": {
            "deviceTypes": {"labels": list(dtype_counts.keys()), "values": list(dtype_counts.values())},
            "uplink":      {"labels": list(upl_counts.keys()),   "values": list(upl_counts.values())},
            "macPerSwitch":{"labels": sw_labels,   "values": sw_values},
            "buildings":   {"labels": bld_labels,  "values": bld_values},
            "interfaces":  {"labels": iface_labels, "active": iface_active,
                            "inactive": iface_inactive, "disabled": iface_disabled},
        },
        "macs": {
            "columns": mac_cols,
            "rows":    _df_rows(df) if not df.empty else [],
            "defaultHidden": ["snmp location", "Switch Model", "CDP Platform",
                              "Classification Source", "MAC Count", "Traffic Active",
                              "Pkts In", "Pkts Out"],
            "filters": {
                "Switch hostname": _filter_opts(df, "Switch hostname"),
                "Building":        _filter_opts(df, "Building"),
                "Floor":           _filter_opts(df, "Floor"),
                "Device Type":     _filter_opts(df, "Device Type"),
                "Uplink":          _filter_opts(df, "Uplink"),
                "Has IP":          _filter_opts(df, "Has IP"),
                "VLAN":            _filter_opts(df, "VLAN"),
            },
        },
        "subnets": {
            "columns": subnet_cols,
            "rows":    _df_rows(df_subnets) if not df_subnets.empty else [],
            "filters": {
                "Building":   _filter_opts(df_subnets, "Building"),
                "Route Type": _filter_opts(df_subnets, "Route Type"),
            },
        },
        "devices": {
            "columns": device_cols,
            "rows":    _df_rows(df_devices) if not df_devices.empty else [],
            "filters": {
                "Building": _filter_opts(df_devices, "Building"),
            },
        },
        "topology": {
            "nodes":          _topo_nodes,
            "edges":          _topo_edges,
            "buildingColors": _bld_color,
        },
        "history":       _history,
        "sites":         _sites_list,
        "deviceSiteMap": _device_site_map,
    }

    data_json = _json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")

    _T = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Cisco Migration Report</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<script src="https://unpkg.com/vis-network@9.1.9/standalone/umd/vis-network.min.js"></script>
<style>
:root{
  --bg:#f0f2f5;--card:#fff;--border:#e2e8f0;--text:#1e293b;--muted:#64748b;
  --hdr:#0f172a;--nav-act:#3b82f6;--nav-txt:#94a3b8;--nav-txt-act:#fff;
  --accent:#2563eb;--green:#16a34a;--red:#dc2626;--yellow:#ca8a04;--orange:#ea580c;
  --purple:#7c3aed;--teal:#0d9488;--cyan:#0891b2;--indigo:#6366f1;--pink:#db2777;
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
     background:var(--bg);color:var(--text);font-size:14px}

/* ══ Header / Nav ══════════════════════════════════════════════════ */
header{background:var(--hdr);color:#fff;padding:0 28px;
       display:flex;align-items:stretch;gap:0;position:sticky;top:0;z-index:100;
       box-shadow:0 2px 14px #0008}
.hdr-brand{display:flex;align-items:center;gap:11px;padding:13px 24px 13px 0;
           border-right:1px solid #ffffff14;margin-right:4px;flex-shrink:0}
.hdr-brand svg{width:26px;height:26px;flex-shrink:0}
.hdr-brand h1{font-size:1rem;font-weight:700;letter-spacing:.1px;white-space:nowrap}
.hdr-brand .sub{font-size:.7rem;color:#64748b;margin-top:1px}
nav{display:flex;align-items:stretch;flex:1;overflow-x:auto;gap:0}
nav button{background:none;border:none;cursor:pointer;padding:0 16px;
           color:var(--nav-txt);font-size:.8rem;font-weight:600;letter-spacing:.2px;
           border-bottom:3px solid transparent;white-space:nowrap;transition:.15s;
           display:flex;align-items:center;gap:6px}
nav button:hover{color:#d1d5db;border-bottom-color:#ffffff25}
nav button.active{color:#fff;border-bottom-color:var(--nav-act)}
nav button .badge{background:#ffffff14;border-radius:9px;padding:1px 7px;
                  font-size:.65rem;font-weight:700}
nav button.active .badge{background:var(--nav-act);color:#fff}
.hdr-right{margin-left:auto;display:flex;align-items:center;gap:14px;
           padding:0 0 0 16px;flex-shrink:0;font-size:.7rem;color:#475569}

/* ══ Sections ══════════════════════════════════════════════════════ */
section{display:none}
section.active{display:block}

/* ══ Dashboard: Hero ════════════════════════════════════════════════ */
.db-hero{background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 50%,#1e3a4a 100%);
         padding:28px 32px 24px;color:#fff;position:relative;overflow:hidden}
.db-hero::before{content:'';position:absolute;inset:0;
  background:url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.03'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E")}
.db-hero-top{display:flex;align-items:flex-start;justify-content:space-between;gap:24px}
.db-hero-title{font-size:1.5rem;font-weight:800;letter-spacing:-.3px;margin-bottom:4px}
.db-hero-sub{font-size:.82rem;color:#94a3b8;max-width:480px}
.db-readiness{text-align:right;flex-shrink:0}
.db-readiness .score{font-size:3rem;font-weight:900;line-height:1;
  background:linear-gradient(135deg,#60a5fa,#34d399);-webkit-background-clip:text;
  -webkit-text-fill-color:transparent;background-clip:text}
.db-readiness .score-lbl{font-size:.72rem;color:#94a3b8;text-transform:uppercase;
  letter-spacing:.6px;margin-top:2px}
.db-progress-bar{margin-top:16px;background:#ffffff10;border-radius:8px;
                 height:8px;overflow:hidden}
.db-progress-bar .fill{height:100%;border-radius:8px;
  background:linear-gradient(90deg,#3b82f6,#10b981);transition:width .8s ease}
.db-progress-labels{display:flex;justify-content:space-between;
  font-size:.68rem;color:#64748b;margin-top:5px}

/* ══ KPI strip ══════════════════════════════════════════════════════ */
.kpi-strip{display:grid;grid-template-columns:repeat(7,1fr);gap:0;
           border-bottom:1px solid var(--border)}
.kpi-s{padding:14px 16px;border-right:1px solid var(--border);background:var(--card);
       min-width:0;transition:.15s}
.kpi-s:last-child{border-right:none}
.kpi-s:hover{background:#f8fafc}
.kpi-s .kval{font-size:1.55rem;font-weight:800;color:var(--c,var(--accent));
             white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
             display:flex;align-items:center;gap:6px}
.kpi-s .kval .icon{font-size:1rem;opacity:.7}
.kpi-s .klbl{font-size:.66rem;color:var(--muted);margin-top:2px;
             text-transform:uppercase;letter-spacing:.4px}
.kpi-s .ktrend{font-size:.66rem;margin-top:4px;font-weight:600}
.kpi-s .ktrend.warn{color:var(--orange)}
.kpi-s .ktrend.ok{color:var(--green)}
.kpi-s .ktrend.info{color:var(--cyan)}

/* ══ Dashboard body ══════════════════════════════════════════════════ */
.db-body{display:grid;grid-template-columns:1fr 340px;gap:16px;
         padding:16px 28px 28px;align-items:start}
.db-left{display:flex;flex-direction:column;gap:14px;min-width:0}
.db-right{display:flex;flex-direction:column;gap:14px}
.ccard{background:var(--card);border-radius:10px;padding:18px 20px;
       box-shadow:0 1px 4px #00000009}
.ccard-hd{display:flex;align-items:center;justify-content:space-between;
          margin-bottom:14px}
.ccard-hd h3{font-size:.72rem;font-weight:700;color:var(--muted);
             text-transform:uppercase;letter-spacing:.6px}
.ccard-hd .tag{font-size:.65rem;padding:2px 8px;border-radius:12px;
               font-weight:700;background:#f1f5f9;color:var(--muted)}
.cwrap{position:relative}
.chart-row{display:grid;grid-template-columns:1fr 1fr;gap:14px}

/* ══ Attention panel ══════════════════════════════════════════════ */
.attn-list{display:flex;flex-direction:column;gap:8px}
.attn-item{display:flex;align-items:flex-start;gap:10px;
           padding:10px 12px;border-radius:8px;
           background:#f8fafc;border-left:3px solid var(--c)}
.attn-item .ai-icon{font-size:1rem;flex-shrink:0;margin-top:1px}
.attn-item .ai-text{min-width:0}
.attn-item .ai-title{font-size:.78rem;font-weight:700;color:var(--text)}
.attn-item .ai-desc{font-size:.71rem;color:var(--muted);margin-top:1px;
                    white-space:normal;line-height:1.4}
.attn-item .ai-cnt{font-size:1.2rem;font-weight:800;color:var(--c);
                   margin-left:auto;flex-shrink:0;padding-left:8px}

/* ══ Device type list (sidebar) ══════════════════════════════════ */
.dtype-row{display:flex;align-items:center;gap:8px;
           padding:7px 0;border-bottom:1px solid #f8fafc}
.dtype-row:last-child{border-bottom:none}
.dtype-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.dtype-name{flex:1;font-size:.79rem;color:var(--text)}
.dtype-cnt{font-size:.79rem;font-weight:700;color:var(--text)}
.dtype-pct{font-size:.68rem;color:var(--muted);width:32px;text-align:right}
.dtype-bar-bg{flex:0 0 70px;background:#f1f5f9;border-radius:4px;height:5px}
.dtype-bar-fill{height:100%;border-radius:4px}

/* ══ Modal overlay ═══════════════════════════════════════════════════ */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.45);display:flex;
  align-items:center;justify-content:center;z-index:1000;padding:24px}
.modal-overlay.hidden{display:none}
.modal-box{background:var(--card);border-radius:12px;width:100%;max-width:600px;
  max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal-hdr{display:flex;align-items:center;justify-content:space-between;
  padding:18px 20px;border-bottom:1px solid var(--border)}
.modal-hdr h3{margin:0;font-size:1rem;font-weight:700}
.modal-close{background:none;border:none;cursor:pointer;font-size:1.3rem;color:var(--muted);
  padding:0 4px;line-height:1}.modal-close:hover{color:var(--text)}
.modal-body{padding:20px}
.modal-footer{display:flex;gap:8px;justify-content:flex-end;
  padding:14px 20px;border-top:1px solid var(--border)}
.modal-form-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.modal-form-grid .span2{grid-column:1/-1}
label.flbl{display:block;font-size:.78rem;font-weight:600;margin-bottom:4px;color:var(--muted)}
input.sinp,textarea.sinp,select.sinp{width:100%;box-sizing:border-box;padding:7px 10px;
  border:1px solid var(--border);border-radius:6px;
  background:var(--bg);color:var(--text);font-size:.86rem}
textarea.sinp{resize:vertical;min-height:60px}
input.sinp:focus,textarea.sinp:focus,select.sinp:focus{outline:none;border-color:#2563eb}
.site-chip{display:inline-block;padding:2px 8px;border-radius:50px;font-size:.75rem;
  font-weight:600;background:#dbeafe;color:#1e40af}
.site-chip.her-RW{background:#d1fae5;color:#065f46}
.site-chip.her-Rogers{background:#ede9fe;color:#4c1d95}
.site-chip.her-TBD{background:#fef3c7;color:#92400e}
/* ══ Sites section ═══════════════════════════════════════════════════ */
.sites-stats{display:flex;gap:16px;padding:20px 24px 4px;flex-wrap:wrap}
.sites-stat{background:var(--card);border:1px solid var(--border);border-radius:10px;
  padding:12px 20px;text-align:center;min-width:90px}
.sites-stat .sv{font-size:1.5rem;font-weight:800;color:#2563eb}
.sites-stat .sl{font-size:.72rem;color:var(--muted);margin-top:2px}
.sites-tbl-wrap{overflow-x:auto;padding:0 16px 24px}
#sites-filter-bar{display:flex;gap:10px;flex-wrap:wrap;padding:8px 24px 4px;align-items:flex-end}
#sites-filter-bar select,#sites-filter-bar input{padding:6px 10px;
  border:1px solid var(--border);border-radius:6px;
  background:var(--bg);color:var(--text);font-size:.85rem}
#sites-filter-bar input{flex:1;min-width:160px}
#site-global-filter{display:flex;align-items:center;gap:8px;padding:0 8px}
#site-global-filter input{transition:border-color .15s}
#site-global-filter input:focus{outline:none;border-color:#2563eb}

/* ══ Topology toolbar, stats & detail panel ═══════════════════════ */
.topo-toolbar{background:var(--card);border-bottom:1px solid var(--border);
  padding:8px 20px;display:flex;align-items:flex-end;gap:12px;flex-wrap:wrap}
.topo-stats-strip{background:#f8fafc;border-bottom:1px solid var(--border);
  padding:6px 22px;display:flex;align-items:center;gap:24px;
  font-size:.74rem;color:var(--muted);font-weight:600}
.topo-stats-strip span{color:var(--text);font-weight:700}
.topo-search{padding:6px 12px 6px 30px;border:1px solid var(--border);
  border-radius:7px;font-size:.78rem;outline:none;width:175px;color:var(--text);
  background:#fff url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='13' height='13' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E") no-repeat 9px center}
.topo-search:focus{border-color:var(--accent);box-shadow:0 0 0 2px #2563eb18}
select.topo-sel{padding:6px 10px;border:1px solid var(--border);border-radius:7px;
  font-size:.78rem;background:#fff;color:var(--text);outline:none;cursor:pointer}
select.topo-sel:focus{border-color:var(--accent)}
.topo-detail-name{font-weight:700;font-size:.95rem;margin-bottom:8px;color:var(--text)}
.topo-detail-grid{display:grid;grid-template-columns:auto 1fr;gap:2px 10px;
  font-size:.76rem;line-height:1.9}
.topo-detail-grid .k{color:var(--muted);white-space:nowrap}
.topo-detail-grid .v{color:var(--text);font-weight:600;min-width:0;word-break:break-all}
.topo-iface-bar{display:grid;grid-template-columns:repeat(3,1fr);gap:4px}
.topo-iface-cell{background:#f8fafc;border-radius:6px;padding:6px 4px;text-align:center;
  border:1px solid var(--border)}
.topo-iface-cell .val{font-size:1rem;font-weight:800}
.topo-iface-cell .lbl{font-size:.62rem;color:var(--muted);text-transform:uppercase;
  letter-spacing:.3px;margin-top:1px}
.topo-nbr-item{padding:6px 10px;border-radius:6px;font-size:.76rem;cursor:pointer;
  border:1px solid var(--border);margin-bottom:4px;background:#fff;
  display:flex;align-items:center;gap:8px}
.topo-nbr-item:hover{background:#eff6ff;border-color:#93c5fd}

/* ══ Quick stats grid ══════════════════════════════════════════════ */
.qs-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.qs-item{background:#f8fafc;border-radius:8px;padding:10px 12px;}
.qs-item .qs-val{font-size:1.3rem;font-weight:800;color:var(--c,var(--accent))}
.qs-item .qs-lbl{font-size:.67rem;color:var(--muted);text-transform:uppercase;
                 letter-spacing:.4px;margin-top:2px}

/* ══ Table page layout ══════════════════════════════════════════════ */
.pg-header{background:var(--card);border-bottom:1px solid var(--border);
           padding:16px 28px 14px;display:flex;align-items:center;
           justify-content:space-between;gap:16px}
.pg-header-left .pg-title{font-size:1.05rem;font-weight:700;color:var(--text)}
.pg-header-left .pg-desc{font-size:.75rem;color:var(--muted);margin-top:2px}
.pg-header-right{display:flex;gap:8px;align-items:center;flex-shrink:0}

/* ══ Filter panel ══════════════════════════════════════════════════ */
.filter-panel{background:#fff;border-bottom:2px solid #e2e8f0;
              padding:12px 28px;display:flex;flex-wrap:wrap;
              align-items:flex-end;gap:10px}
.filter-panel-label{font-size:.65rem;font-weight:800;color:#94a3b8;
  text-transform:uppercase;letter-spacing:.6px;align-self:center;
  flex-shrink:0;border-right:1px solid var(--border);padding-right:12px;
  margin-right:2px}
.fp-field{display:flex;flex-direction:column;gap:3px}
.flabel{font-size:.65rem;font-weight:700;color:var(--muted);
        text-transform:uppercase;letter-spacing:.4px}
select.flt{padding:6px 10px;border:1px solid var(--border);border-radius:7px;
           font-size:.78rem;background:#fff;color:var(--text);
           outline:none;min-width:100px;max-width:180px}
select.flt:focus{border-color:var(--accent);box-shadow:0 0 0 2px #2563eb18}
input.srch{padding:7px 12px 7px 32px;border:1px solid var(--border);border-radius:7px;
           font-size:.82rem;background:#fff;color:var(--text);
           outline:none;width:220px;
           background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='14' height='14' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Ccircle cx='11' cy='11' r='8'/%3E%3Cpath d='m21 21-4.35-4.35'/%3E%3C/svg%3E");
           background-repeat:no-repeat;background-position:10px center}
input.srch:focus{border-color:var(--accent);box-shadow:0 0 0 2px #2563eb18}
.filter-actions{margin-left:auto;display:flex;gap:7px;align-items:flex-end;flex-shrink:0}
.btn{padding:7px 13px;border-radius:7px;font-size:.78rem;font-weight:600;
     cursor:pointer;border:1px solid var(--border);background:#fff;
     color:var(--text);display:flex;align-items:center;gap:5px;white-space:nowrap;
     transition:.15s}
.btn:hover{background:#f1f5f9}
.btn-primary{background:var(--accent);color:#fff;border-color:var(--accent)}
.btn-primary:hover{background:#1d4ed8}
.btn-warn{background:#fff7ed;color:#ea580c;border-color:#fed7aa}
.btn-warn:hover{background:#fff0dc}
.btn-warn.active{background:#ea580c;color:#fff;border-color:#ea580c}

/* ══ Column toggle ══════════════════════════════════════════════════ */
.col-toggle-wrap{position:relative}
.col-toggle-wrap .dropdown{display:none;position:absolute;right:0;top:calc(100% + 4px);
  background:#fff;border:1px solid var(--border);border-radius:9px;
  box-shadow:0 6px 24px #0003;z-index:50;padding:12px 14px;
  min-width:220px;max-height:320px;overflow-y:auto}
.col-toggle-wrap.open .dropdown{display:block}
.col-toggle-wrap .dropdown label{display:flex;align-items:center;gap:8px;
  padding:4px 0;cursor:pointer;font-size:.78rem;color:var(--text)}
.col-toggle-wrap .dropdown label:hover{color:var(--accent)}
.col-toggle-wrap .dropdown label input{accent-color:var(--accent)}
.col-toggle-wrap .dropdown .dd-hd{font-size:.65rem;font-weight:800;
  color:#94a3b8;margin-bottom:8px;text-transform:uppercase;letter-spacing:.5px}

/* ══ Table ══════════════════════════════════════════════════════════ */
.tsec{padding:16px 28px 32px}
.tbl-wrap{background:var(--card);border-radius:10px;box-shadow:0 1px 4px #00000009;overflow:hidden}
.tbl-scroll{overflow-x:auto;max-height:68vh}
table{width:100%;border-collapse:collapse;font-size:.76rem}
thead tr{position:sticky;top:0;z-index:10}
th{background:#f8fafc;padding:9px 12px;text-align:left;font-weight:700;color:var(--muted);
   border-bottom:2px solid var(--border);white-space:nowrap;cursor:pointer;
   user-select:none;font-size:.67rem;text-transform:uppercase;letter-spacing:.4px}
th:hover{background:#f1f5f9;color:var(--text)}
th.s-asc::after{content:" ▲";color:var(--accent)}
th.s-desc::after{content:" ▼";color:var(--accent)}
td{padding:7px 12px;border-bottom:1px solid #f8fafc;white-space:nowrap;
   max-width:260px;overflow:hidden;text-overflow:ellipsis}
td.wrap{white-space:normal;max-width:300px}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f8fafc}
.r{text-align:right}
.tfoot{display:flex;align-items:center;justify-content:space-between;
       padding:10px 16px;border-top:1px solid var(--border);
       font-size:.74rem;color:var(--muted)}
.pager{display:flex;gap:4px;align-items:center}
.pager button{padding:4px 9px;border:1px solid var(--border);border-radius:5px;
              background:#fff;cursor:pointer;font-size:.73rem;color:var(--text);transition:.1s}
.pager button:hover{background:#f1f5f9}
.pager button.cur{background:var(--accent);color:#fff;border-color:var(--accent)}
.pager button:disabled{opacity:.35;cursor:default}
select.pgsize{padding:3px 7px;border:1px solid var(--border);border-radius:5px;
              font-size:.73rem;background:#fff;color:var(--text)}

/* ══ Badges ══════════════════════════════════════════════════════════ */
.bdg{display:inline-block;padding:2px 8px;border-radius:20px;
     font-size:.67rem;font-weight:700;letter-spacing:.2px}
.no-data{text-align:center;padding:48px;color:var(--muted);font-size:.88rem}
.conflict-cell{color:var(--orange);font-weight:600}
.row-conflict td:first-child{border-left:3px solid var(--orange)}

/* ══ Responsive ══════════════════════════════════════════════════════ */
@media(max-width:1100px){
  .db-body{grid-template-columns:1fr}
  .db-right{display:grid;grid-template-columns:1fr 1fr;gap:14px}
}
@media(max-width:900px){
  .kpi-strip{grid-template-columns:repeat(4,1fr)}
  .chart-row{grid-template-columns:1fr}
  .db-right{grid-template-columns:1fr}
}
@media(max-width:640px){
  .kpi-strip{grid-template-columns:1fr 1fr}
  header{flex-wrap:wrap}.hdr-right{display:none}
  .db-hero-top{flex-direction:column}
  .db-readiness{text-align:left}
}
/* ══ Animations ══════════════════════════════════════════════════════ */
.tl-diff-add{color:var(--green);font-weight:700}
.tl-diff-del{color:var(--red);opacity:.7}
.tl-sw-hdr td{background:#f8fafc;padding:10px 14px;border-top:2px solid var(--border);
  border-bottom:1px solid var(--border);font-weight:700;font-size:.82rem}
.tl-row td{padding:8px 12px;vertical-align:top;border-bottom:1px solid var(--border)}
.tl-row.type-added{border-left:3px solid var(--green)}
.tl-row.type-removed{border-left:3px solid var(--red)}
.tl-row.type-changed{border-left:3px solid var(--orange)}
.tl-fbtn{padding:3px 10px;font-size:.72rem;border-radius:20px}
.tl-fbtn.active{background:var(--accent);color:#fff;border-color:var(--accent)}
.tl-scope-box{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;
  padding:10px 16px;display:flex;flex-direction:column;gap:6px}
.fp-field{display:flex;flex-direction:column;gap:3px}
@keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
.kpi-s{animation:fadeUp .3s ease both}
.kpi-s:nth-child(1){animation-delay:.05s}.kpi-s:nth-child(2){animation-delay:.1s}
.kpi-s:nth-child(3){animation-delay:.15s}.kpi-s:nth-child(4){animation-delay:.2s}
.kpi-s:nth-child(5){animation-delay:.25s}.kpi-s:nth-child(6){animation-delay:.3s}
.kpi-s:nth-child(7){animation-delay:.35s}
</style>
</head>
<body>

<!-- ══════════════════════ HEADER ══════════════════════════════════ -->
<header>
  <div class="hdr-brand">
    <svg viewBox="0 0 28 28" fill="none"><rect width="28" height="28" rx="6" fill="#2563eb"/>
    <path d="M5 14h3M20 14h3M9 9l2 2M17 17l2 2M9 19l2-2M17 11l2-2M14 5v3M14 20v3"
          stroke="#fff" stroke-width="1.8" stroke-linecap="round"/>
    <circle cx="14" cy="14" r="3" fill="#fff"/></svg>
    <div><h1>Cisco Migration Report</h1><div class="sub" id="hMeta"></div></div>
  </div>
  <nav>
    <button class="active" onclick="nav('dashboard')" id="nb-dashboard">📊 Dashboard</button>
    <button onclick="nav('macs')"    id="nb-macs">
      🔌 Port Inventory <span class="badge" id="badge-macs">0</span>
    </button>
    <button onclick="nav('subnets')" id="nb-subnets">
      🌐 Subnets &amp; Routes <span class="badge" id="badge-subnets">0</span>
    </button>
    <button onclick="nav('devices')" id="nb-devices">
      🖥 Switch Inventory <span class="badge" id="badge-devices">0</span>
    </button>
    <button onclick="nav('timeline')" id="nb-timeline">
      📅 Timeline <span class="badge" id="badge-timeline">0</span>
    </button>
    <button onclick="nav('topology')" id="nb-topology">🗺 Topology</button>
    <button onclick="nav('sites')" id="nb-sites">📍 Sites <span class="badge" id="badge-sites">0</span></button>
  </nav>
  <div id="site-global-filter">
    <label style="font-size:.8rem;color:var(--muted);font-weight:600">Site:</label>
    <input id="site-global-inp" type="text"
      placeholder="Search site / city / address…"
      oninput="setGlobalSite(this.value)"
      style="padding:5px 10px;border:1px solid var(--border);border-radius:6px;
             background:var(--bg);color:var(--text);font-size:.84rem;width:220px"
      title="Filter all tabs to devices at matching sites">
    <button onclick="document.getElementById('site-global-inp').value='';setGlobalSite('')"
      style="padding:4px 8px;border:1px solid var(--border);border-radius:6px;
             background:var(--bg);color:var(--muted);cursor:pointer;font-size:.8rem"
      title="Clear site filter">&times;</button>
  </div>
  <div class="hdr-right" id="hdrRight"></div>
</header>

<!-- ══════════════════════ DASHBOARD ═══════════════════════════════ -->
<section id="s-dashboard" class="active">

  <!-- Hero -->
  <div class="db-hero">
    <div class="db-hero-top">
      <div>
        <div class="db-hero-title">Network Migration Overview</div>
        <div class="db-hero-sub" id="db-hero-sub">Analysing switch estate — please wait…</div>
      </div>
      <div class="db-readiness">
        <div class="score" id="db-score">—</div>
        <div class="score-lbl">Classification Rate</div>
      </div>
    </div>
    <div class="db-progress-bar"><div class="fill" id="db-prog" style="width:0%"></div></div>
    <div class="db-progress-labels">
      <span id="db-prog-l">0 classified</span>
      <span id="db-prog-r">0 unclassified</span>
    </div>
  </div>

  <!-- KPI strip -->
  <div class="kpi-strip">
    <div class="kpi-s" style="--c:#2563eb">
      <div class="kval"><span class="icon">🔌</span><span id="k1">—</span></div>
      <div class="klbl">Port Entries</div>
    </div>
    <div class="kpi-s" style="--c:#0891b2">
      <div class="kval"><span class="icon">🖥</span><span id="k2">—</span></div>
      <div class="klbl">Switches</div>
    </div>
    <div class="kpi-s" style="--c:#059669">
      <div class="kval"><span class="icon">🌐</span><span id="k3">—</span></div>
      <div class="klbl">Subnets</div>
    </div>
    <div class="kpi-s" style="--c:#7c3aed">
      <div class="kval"><span class="icon">🏢</span><span id="k4">—</span></div>
      <div class="klbl">Buildings</div>
    </div>
    <div class="kpi-s" style="--c:#ea580c">
      <div class="kval"><span class="icon">⚠️</span><span id="k5">—</span></div>
      <div class="klbl">Conflicts</div>
    </div>
    <div class="kpi-s" style="--c:#0369a1">
      <div class="kval"><span class="icon">↑</span><span id="k6">—</span></div>
      <div class="klbl">Uplink Ports</div>
    </div>
    <div class="kpi-s" style="--c:#9f1239">
      <div class="kval"><span class="icon">❓</span><span id="k7">—</span></div>
      <div class="klbl">No-IP %</div>
    </div>
  </div>

  <!-- Main body -->
  <div class="db-body">

    <!-- LEFT column -->
    <div class="db-left">

      <!-- Charts row: donut pair -->
      <div class="chart-row">
        <div class="ccard">
          <div class="ccard-hd"><h3>Device Type Distribution</h3><span class="tag" id="tag-dtype">Total: 0</span></div>
          <div class="cwrap" style="height:230px"><canvas id="ch-dtype"></canvas></div>
        </div>
        <div class="ccard">
          <div class="ccard-hd"><h3>Port Role (Uplink vs Access)</h3><span class="tag" id="tag-upl">Total: 0</span></div>
          <div class="cwrap" style="height:230px"><canvas id="ch-upl"></canvas></div>
        </div>
      </div>

      <!-- MACs per switch bar -->
      <div class="ccard">
        <div class="ccard-hd"><h3>Devices per Switch</h3>
          <span class="tag" id="tag-sw">Sorted by count</span></div>
        <div class="cwrap" id="w-sw"><canvas id="ch-sw"></canvas></div>
      </div>

      <!-- Building bar -->
      <div class="ccard">
        <div class="ccard-hd"><h3>Devices by Building (Top 20)</h3></div>
        <div class="cwrap" id="w-bld"><canvas id="ch-bld"></canvas></div>
      </div>

      <!-- Interface status stacked bar -->
      <div class="ccard">
        <div class="ccard-hd"><h3>Interface Status per Switch</h3>
          <span class="tag">First 30 switches</span></div>
        <div class="cwrap" id="w-ifc"><canvas id="ch-ifc"></canvas></div>
      </div>
    </div>

    <!-- RIGHT column -->
    <div class="db-right">

      <!-- Attention required -->
      <div class="ccard">
        <div class="ccard-hd"><h3>⚠ Attention Required</h3></div>
        <div class="attn-list" id="attn-list">
          <div style="color:var(--muted);font-size:.78rem">Loading…</div>
        </div>
      </div>

      <!-- Device breakdown list -->
      <div class="ccard">
        <div class="ccard-hd"><h3>Device Breakdown</h3></div>
        <div id="dtype-list">
          <div style="color:var(--muted);font-size:.78rem">Loading…</div>
        </div>
      </div>

      <!-- Quick stats grid -->
      <div class="ccard">
        <div class="ccard-hd"><h3>Quick Stats</h3></div>
        <div class="qs-grid">
          <div class="qs-item" style="--c:#2563eb">
            <div class="qs-val" id="qs1">—</div><div class="qs-lbl">Avg MACs / Switch</div></div>
          <div class="qs-item" style="--c:#059669">
            <div class="qs-val" id="qs2">—</div><div class="qs-lbl">Active Switches</div></div>
          <div class="qs-item" style="--c:#0891b2">
            <div class="qs-val" id="qs3">—</div><div class="qs-lbl">Avg Subnets / Switch</div></div>
          <div class="qs-item" style="--c:#7c3aed">
            <div class="qs-val" id="qs4">—</div><div class="qs-lbl">With IP</div></div>
        </div>
      </div>

    </div>
  </div>
</section>

<!-- ══════════════════════ PORT INVENTORY ══════════════════════════ -->
<section id="s-macs">
  <div class="pg-header">
    <div class="pg-header-left">
      <div class="pg-title">🔌 Port Inventory</div>
      <div class="pg-desc">All switch ports and connected devices — one row per MAC address seen on a port. Use filters below to narrow down for migration planning.</div>
    </div>
    <div class="pg-header-right">
      <div class="col-toggle-wrap" id="mac-col-toggle">
        <button class="btn" onclick="toggleColMenu()" title="Show/hide columns">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/></svg>
          Columns
        </button>
        <div class="dropdown" id="mac-col-menu"><div class="dd-hd">Toggle Columns</div></div>
      </div>
      <button class="btn btn-warn" onclick="conflictOnly()" id="btn-conflict">⚠ Conflicts Only</button>
      <button class="btn btn-primary" onclick="exportCSV('mac')">⬇ Export CSV</button>
    </div>
  </div>
  <div class="filter-panel" id="mac-filters">
    <span class="filter-panel-label">🔍 Filter by</span>
  </div>
  <div class="tsec" style="padding-top:12px">
    <div class="tbl-wrap">
      <div class="tbl-scroll"><table id="mac-tbl"><thead></thead><tbody></tbody></table></div>
      <div class="tfoot">
        <div id="mac-info"></div>
        <div style="display:flex;align-items:center;gap:10px">
          <select class="pgsize" id="mac-pgsize" onchange="MacTV.setPageSize(this.value)">
            <option value="50">50 / page</option><option value="100">100 / page</option>
            <option value="200">200 / page</option><option value="500">500 / page</option>
            <option value="99999">All</option>
          </select>
          <div class="pager" id="mac-pager"></div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ══════════════════════ SUBNETS & ROUTES ═══════════════════════ -->
<section id="s-subnets">
  <div class="pg-header">
    <div class="pg-header-left">
      <div class="pg-title">🌐 Subnets &amp; Routes</div>
      <div class="pg-desc">Connected and static routes discovered on each switch. Use this to plan VLAN / IP scheme for the new environment.</div>
    </div>
    <div class="pg-header-right">
      <button class="btn btn-primary" onclick="exportCSV('sub')">⬇ Export CSV</button>
    </div>
  </div>
  <div class="filter-panel" id="sub-filters">
    <span class="filter-panel-label">🔍 Filter by</span>
  </div>
  <div class="tsec" style="padding-top:12px">
    <div class="tbl-wrap">
      <div class="tbl-scroll"><table id="sub-tbl"><thead></thead><tbody></tbody></table></div>
      <div class="tfoot">
        <div id="sub-info"></div>
        <div style="display:flex;align-items:center;gap:10px">
          <select class="pgsize" id="sub-pgsize" onchange="SubTV.setPageSize(this.value)">
            <option value="100">100 / page</option><option value="200">200 / page</option>
            <option value="99999">All</option>
          </select>
          <div class="pager" id="sub-pager"></div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- ══════════════════════ SWITCH INVENTORY ═══════════════════════ -->
<section id="s-devices">
  <div class="pg-header">
    <div class="pg-header-left">
      <div class="pg-title">🖥 Switch Inventory</div>
      <div class="pg-desc">One row per switch — model, serial, IOS version, uptime, and device-type breakdown. Essential for hardware EOL and replacement planning.
        Use the <strong>Assigned Site</strong> column to link each switch to a site — click <strong>+ Assign</strong> on any row.</div>
    </div>
    <div class="pg-header-right">
      <button class="btn btn-primary" onclick="exportCSV('dev')">⬇ Export CSV</button>
    </div>
  </div>
  <div class="filter-panel" id="dev-filters">
    <span class="filter-panel-label">🔍 Filter by</span>
  </div>
  <div class="tsec" style="padding-top:12px">
    <div class="tbl-wrap">
      <div class="tbl-scroll"><table id="dev-tbl"><thead></thead><tbody></tbody></table></div>
      <div class="tfoot">
        <div id="dev-info"></div>
        <div style="display:flex;align-items:center;gap:10px">
          <select class="pgsize" id="dev-pgsize" onchange="DevTV.setPageSize(this.value)">
            <option value="50">50 / page</option><option value="100">100 / page</option>
            <option value="99999">All</option>
          </select>
          <div class="pager" id="dev-pager"></div>
        </div>
      </div>
    </div>
  </div>
</section>

<!-- \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550 TIMELINE \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550 -->
<section id="s-timeline">
  <div class="pg-header">
    <div class="pg-header-left">
      <div class="pg-title">\U0001f4c5 Audit Timeline</div>
      <div class="pg-desc">Select any two audit runs to compare. Only switches present in <em>both</em> selected runs are diffed \u2014 switches missing from one run are not assumed removed.</div>
    </div>
  </div>

  <!-- Run selector bar -->
  <div class="filter-panel" id="tl-selector" style="gap:14px;align-items:flex-end;flex-wrap:wrap">
    <span class="filter-panel-label">Compare Runs</span>
    <div class="fp-field">
      <label class="flabel">From Run (A)</label>
      <select class="flt" id="tl-run-a" style="min-width:240px"></select>
    </div>
    <div class="fp-field">
      <label class="flabel">To Run (B)</label>
      <select class="flt" id="tl-run-b" style="min-width:240px"></select>
    </div>
    <div class="filter-actions">
      <button class="btn btn-primary" onclick="_doCompare()">\u21ba Compare</button>
    </div>
    <div style="margin-left:auto;display:flex;gap:6px;align-items:flex-end" id="tl-type-filters"></div>
  </div>

  <!-- Summary KPI strip (hidden until compare) -->
  <div id="tl-summary" style="display:none;background:var(--card);border-bottom:1px solid var(--border);padding:10px 28px;gap:20px;align-items:center;flex-wrap:wrap"></div>

  <!-- Scope notice (hidden until compare) -->
  <div id="tl-scope" style="display:none;padding:10px 28px 0"></div>

  <!-- Results body -->
  <div class="tsec" style="padding-top:12px">
    <div id="tl-results">
      <div class="no-data" style="padding:48px">Select two runs above and click Compare.</div>
    </div>
  </div>
</section>

<!-- \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550 TOPOLOGY \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550 -->
<section id="s-topology">
  <div class="pg-header">
    <div class="pg-header-left">
      <div class="pg-title">\U0001f5fa Network Topology</div>
      <div class="pg-desc">CDP switch links &mdash; colour = building. Click a node for details; double-click to isolate neighbours.</div>
    </div>
    <div class="pg-header-right" id="topo-hdr-btns" style="display:none">
      <button class="btn" id="btn-topo-ext" onclick="_topoToggleExt()" title="Toggle external/unaudited nodes">\U0001f441 External</button>
      <button class="btn btn-primary" onclick="_topoExport()">\u2b07 Export PNG</button>
      <button class="btn" onclick="if(_topoNet)_topoNet.fit()">\u22a1 Fit All</button>
    </div>
  </div>
  <!-- stats strip -->
  <div class="topo-stats-strip">
    Switches: <span id="ts-sw">&mdash;</span>
    &nbsp;&nbsp;Links: <span id="ts-lk">&mdash;</span>
    &nbsp;&nbsp;Buildings: <span id="ts-bld">&mdash;</span>
  </div>
  <!-- toolbar -->
  <div class="topo-toolbar" id="topo-toolbar" style="display:none">
    <div style="display:flex;flex-direction:column;gap:3px">
      <label class="flabel">Layout</label>
      <select class="topo-sel" id="topo-layout" onchange="_topoLayout(this.value)">
        <option value="physics">Physics (default)</option>
        <option value="hierarchical">Hierarchical \u2193</option>
        <option value="hierarchicalLR">Hierarchical \u2192</option>
      </select>
    </div>
    <div style="display:flex;flex-direction:column;gap:3px">
      <label class="flabel">Building</label>
      <select class="topo-sel" id="topo-bld-filter" onchange="_topoBldFilter(this.value)">
        <option value="">All Buildings</option>
      </select>
    </div>
    <div style="display:flex;flex-direction:column;gap:3px">
      <label class="flabel">Find switch</label>
      <input class="topo-search" id="topo-search-inp" placeholder="Hostname / model\u2026"
             oninput="_topoSearch(this.value)" />
    </div>
  </div>
  <!-- canvas + sidebar -->
  <div style="display:grid;grid-template-columns:1fr 275px;height:calc(100vh - 195px)">
    <div id="topo-canvas" style="width:100%;height:100%"></div>
    <div style="background:var(--card);border-left:1px solid var(--border);padding:16px;
                overflow-y:auto;display:flex;flex-direction:column;gap:16px">
      <div id="topo-detail-wrap">
        <div class="flabel" style="margin-bottom:8px">Selected Node</div>
        <div id="topo-detail" style="font-size:.78rem;color:var(--muted)">Click a switch to see details</div>
      </div>
      <div>
        <div class="flabel" style="margin-bottom:8px">Building Legend</div>
        <div id="topo-legend"></div>
      </div>
    </div>
  </div>
</section>

<!-- ══════════════════════ SITES ════════════════════════════════════ -->
<section id="s-sites">
  <div style="display:flex;align-items:center;justify-content:space-between;padding:16px 24px 0">
    <h2 style="margin:0;font-size:1.1rem">📍 Site Directory</h2>
    <div style="display:flex;gap:8px">
      <button class="btn-sm" onclick="openSiteModal(null)">+ Add Site</button>
      <button class="btn-sm" onclick="exportSiteChanges()">⬇ Export Changes</button>
    </div>
  </div>
  <div class="sites-stats">
    <div class="sites-stat"><div class="sv" id="ss-total">0</div><div class="sl">Total Sites</div></div>
    <div class="sites-stat"><div class="sv" id="ss-provs">0</div><div class="sl">Provinces</div></div>
    <div class="sites-stat"><div class="sv" id="ss-assigned">0</div><div class="sl">Assigned</div></div>
    <div class="sites-stat"><div class="sv" id="ss-2024">0</div><div class="sl">2024</div></div>
    <div class="sites-stat"><div class="sv" id="ss-2025">0</div><div class="sl">2025</div></div>
    <div class="sites-stat"><div class="sv" id="ss-2026">0</div><div class="sl">2026</div></div>
  </div>
  <div id="sites-filter-bar">
    <select id="sf-heritage" onchange="buildSites()"><option value="">All Heritage</option></select>
    <select id="sf-province" onchange="buildSites()"><option value="">All Provinces</option></select>
    <select id="sf-year" onchange="buildSites()">
      <option value="">All Years</option>
      <option>2024</option><option>2025</option><option>2026</option>
    </select>
    <input id="sf-search" placeholder="Search site code or address…" oninput="buildSites()">
    <button class="btn-sm" onclick="
      document.getElementById('sf-heritage').value='';
      document.getElementById('sf-province').value='';
      document.getElementById('sf-year').value='';
      document.getElementById('sf-search').value='';
      buildSites()">Clear</button>
  </div>
  <div class="sites-tbl-wrap">
    <table id="sites-tbl" class="data-tbl">
      <thead><tr>
        <th>Site Code</th><th>Heritage</th><th>Street Address</th>
        <th>City</th><th>Prov</th><th>Year</th>
        <th>Switches</th><th>Notes</th><th>Actions</th>
      </tr></thead>
      <tbody id="sites-tbody"></tbody>
    </table>
  </div>
</section>

<!-- ══ Add / Edit Site Modal ══════════════════════════════════════════ -->
<div class="modal-overlay hidden" id="site-modal-overlay" onclick="if(event.target===this)closeSiteModal()">
  <div class="modal-box">
    <div class="modal-hdr">
      <h3 id="site-modal-title">Add Site</h3>
      <button class="modal-close" onclick="closeSiteModal()">&#x2715;</button>
    </div>
    <div class="modal-body">
      <div class="modal-form-grid">
        <div><label class="flbl">Site Code *</label><input class="sinp" id="sm-code" placeholder="e.g. TORO333"></div>
        <div><label class="flbl">Heritage</label>
          <select class="sinp" id="sm-heritage">
            <option value="">&#x2014;</option><option>RE</option><option>RW</option><option>Rogers</option><option>TBD</option>
          </select></div>
        <div class="span2"><label class="flbl">Street Address</label>
          <div style="display:flex;gap:6px">
            <input class="sinp" id="sm-address" placeholder="123 Main St, City ON A1B 2C3" style="flex:1">
            <button class="btn-sm" onclick="lookupAddress()" title="Lookup via Nominatim">🔍 Lookup</button>
            <button class="btn-sm" onclick="openGoogleMaps()" title="Open in Google Maps">🗺</button>
          </div></div>
        <div><label class="flbl">City</label><input class="sinp" id="sm-city"></div>
        <div><label class="flbl">Province</label><input class="sinp" id="sm-province" placeholder="ON"></div>
        <div><label class="flbl">Postal Code</label><input class="sinp" id="sm-postal"></div>
        <div><label class="flbl">Planned Completion</label>
          <select class="sinp" id="sm-year">
            <option value="">&#x2014;</option><option>2024</option><option>2025</option><option>2026</option>
            <option>2024/2025</option><option>2025/2026</option>
          </select></div>
        <div id="sm-latlon" class="span2" style="font-size:.75rem;color:var(--muted);min-height:1.2em"></div>
        <div class="span2"><label class="flbl">Notes</label><textarea class="sinp" id="sm-notes" rows="2"></textarea></div>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn-sm" id="sm-delete-btn" onclick="deleteSite()"
        style="margin-right:auto;background:#fee2e2;color:#b91c1c;border-color:#fca5a5">🗑 Delete</button>
      <button class="btn-sm" onclick="closeSiteModal()">Cancel</button>
      <button class="btn-sm btn-primary" onclick="saveSiteModal()">Save</button>
    </div>
  </div>
</div>

<!-- ══ Assign Switch to Site Modal ═════════════════════════════════ -->
<div class="modal-overlay hidden" id="assign-modal-overlay" onclick="if(event.target===this)closeAssignModal()">
  <div class="modal-box" style="max-width:480px">
    <div class="modal-hdr">
      <h3>Assign Site to Switch</h3>
      <button class="modal-close" onclick="closeAssignModal()">&#x2715;</button>
    </div>
    <div class="modal-body">
      <p style="margin:0 0 12px;font-size:.9rem">Switch: <strong id="am-hostname"></strong></p>
      <label class="flbl">Search site</label>
      <input type="text" id="am-search" class="sinp" placeholder="Type site code, city, province…"
        oninput="_filterAssignSites(this.value)"
        style="margin-bottom:8px">
      <label class="flbl">Select Site <span id="am-count" style="font-weight:400;color:var(--muted)"></span></label>
      <select class="sinp" id="am-site-sel" size="8"
        style="height:200px;overflow-y:auto">
        <option value="">&#x2014; Unassign &#x2014;</option>
      </select>
    </div>
    <div class="modal-footer">
      <button class="btn-sm" onclick="closeAssignModal()">Cancel</button>
      <button class="btn-sm btn-primary" onclick="saveAssignModal()">Save</button>
    </div>
  </div>
</div>

<script>
const D = __DATA_JSON__;

// ── Sites eager state ───────────────────────────────────────────────────────────────
let _sitesBuilt=false;
let _sitesState=D.sites?JSON.parse(JSON.stringify(D.sites)):[];
let _dsmState=Object.assign({},D.deviceSiteMap||{});
let _sitesPending={upsert:[],delete:[]}, _dsmPending={};
// Merge any pending changes saved from a previous session
(function(){
  try{
    const ls=JSON.parse(localStorage.getItem('_sitesPending')||'null');
    if(ls){_sitesPending=ls;}
    const ld=JSON.parse(localStorage.getItem('_dsmPending')||'null');
    if(ld){
      _dsmPending=ld;
      Object.assign(_dsmState,ld);
      // Patch _sitesState assigned_switch_count to reflect pending assignments
      const cnts={};
      Object.values(_dsmState).forEach(sc=>{if(sc)cnts[sc]=(cnts[sc]||0)+1;});
      _sitesState.forEach(s=>{s.assigned_switch_count=cnts[s.site_code]||s.assigned_switch_count||0;});
    }
    // Apply pending upserts to _sitesState
    if(ls&&ls.upsert){
      ls.upsert.forEach(u=>{
        const idx=_sitesState.findIndex(s=>s.site_code===u.site_code);
        if(idx>=0)Object.assign(_sitesState[idx],u);
        else _sitesState.push(Object.assign({assigned_switch_count:0},u));
      });
    }
    if(ls&&ls.delete){
      ls.delete.forEach(code=>{const i=_sitesState.findIndex(s=>s.site_code===code);if(i>=0)_sitesState.splice(i,1);});
    }
  }catch(e){}
})();

// ── Palettes ───────────────────────────────────────────────────────────────
const PAL=['#2563eb','#10b981','#f59e0b','#6366f1','#ec4899','#8b5cf6',
           '#ef4444','#14b8a6','#f97316','#94a3b8','#0891b2','#84cc16',
           '#db2777','#7c3aed','#064e3b'];
const DTYPE_CLR={'Switch/Router':'#2563eb','Wireless AP':'#0891b2',
  'IP Phone':'#16a34a','Server':'#7c3aed','Printer':'#ea580c',
  'IP Camera':'#dc2626','Workstation':'#0d9488','Virtual Machine':'#6366f1',
  'Video Conf':'#db2777','Unknown':'#94a3b8'};
const UPL_CLR={'Yes':'#16a34a','No':'#2563eb','Possible':'#ca8a04'};
function pal(n){return Array.from({length:n},(_,i)=>PAL[i%PAL.length]);}

// ── Nav ───────────────────────────────────────────────────────────────────
const SECTIONS=['dashboard','macs','subnets','devices','timeline','topology','sites'];
function nav(id){
  SECTIONS.forEach(s=>{
    document.getElementById('s-'+s).classList.toggle('active',s===id);
    const nb=document.getElementById('nb-'+s);
    if(nb) nb.classList.toggle('active',s===id);
  });
  if(id==='macs'     && !MacTV._built) MacTV.build();
  if(id==='subnets'  && !SubTV._built) SubTV.build();
  if(id==='devices'  && !DevTV._built) DevTV.build();
  if(id==='timeline' && !_tlBuilt)     _buildTimeline();
  if(id==='topology' && !_topoNet)     _initTopology();
  if(id==='sites'    && !_sitesBuilt)  _initSites();
}

// ── Header meta ───────────────────────────────────────────────────────────
const M=D.meta;
document.getElementById('hMeta').textContent=(M.source?M.source+' · ':'')+M.generated;
document.getElementById('hdrRight').textContent='Generated '+M.generated;
document.getElementById('badge-macs').textContent    = D.macs.rows.length.toLocaleString();
document.getElementById('badge-subnets').textContent = D.subnets.rows.length.toLocaleString();
document.getElementById('badge-devices').textContent = D.devices.rows.length.toLocaleString();

// ── KPI strip ─────────────────────────────────────────────────────────────
document.getElementById('k1').textContent = M.totalMacs.toLocaleString();
document.getElementById('k2').textContent = M.totalSwitches.toLocaleString();
document.getElementById('k3').textContent = M.totalSubnets.toLocaleString();
document.getElementById('k4').textContent = M.totalBuildings.toLocaleString();
document.getElementById('k5').textContent = M.totalConflicts.toLocaleString();
document.getElementById('k6').textContent = M.totalUplinks.toLocaleString();
document.getElementById('k7').textContent = M.noIpPct+'%';

// ── Dashboard hero / readiness ────────────────────────────────────────────
(function(){
  const rows = D.macs.rows;
  const total = rows.length;
  const classified = rows.filter(r=>r['Device Type']&&r['Device Type']!=='Unknown').length;
  const pct = total ? Math.round(100*classified/total) : 0;
  document.getElementById('db-score').textContent = pct+'%';
  document.getElementById('db-prog').style.width = pct+'%';
  document.getElementById('db-prog-l').textContent = classified.toLocaleString()+' classified';
  document.getElementById('db-prog-r').textContent = (total-classified).toLocaleString()+' unknown / unclassified';
  document.getElementById('db-hero-sub').textContent =
    total.toLocaleString()+' port entries across '+M.totalSwitches+' switches'+
    (M.source ? ' · Source: '+M.source : '');

  // Quick stats
  const swWithMacs = D.devices.rows.filter(r=>Number(r['Total MACs'])>0).length;
  const avgMacs = M.totalSwitches ? Math.round(M.totalMacs/M.totalSwitches) : 0;
  const avgSubs = M.totalSwitches ? (M.totalSubnets/M.totalSwitches).toFixed(1) : 0;
  const withIp  = total ? Math.round(100*rows.filter(r=>r['Has IP']==='Yes').length/total)+'%' : '—';
  document.getElementById('qs1').textContent = avgMacs.toLocaleString();
  document.getElementById('qs2').textContent = swWithMacs.toLocaleString();
  document.getElementById('qs3').textContent = avgSubs;
  document.getElementById('qs4').textContent = withIp;

  // tag updates
  const dtTotal = D.charts.deviceTypes.values.reduce((a,b)=>a+b,0);
  document.getElementById('tag-dtype').textContent = 'Total: '+dtTotal.toLocaleString();
  const uplTotal = D.charts.uplink.values.reduce((a,b)=>a+b,0);
  document.getElementById('tag-upl').textContent = 'Total: '+uplTotal.toLocaleString();
  document.getElementById('tag-sw').textContent = M.totalSwitches+' switches';
})();

// ── Attention panel ───────────────────────────────────────────────────────
(function(){
  const items=[];
  if(M.totalConflicts>0)
    items.push({c:'var(--orange)',icon:'⚠️',
      title:M.totalConflicts.toLocaleString()+' Classification Conflicts',
      desc:'OUI vendor contradicts the port description. Review the Port Inventory → "Conflicts Only" filter.',
      cnt:null});
  const unknownCnt = D.macs.rows.filter(r=>r['Device Type']==='Unknown').length;
  if(unknownCnt>0)
    items.push({c:'var(--yellow)',icon:'❓',
      title:unknownCnt.toLocaleString()+' Unclassified Devices',
      desc:'Device type could not be inferred from CDP, description, or OUI. Manual review recommended.',
      cnt:null});
  const noIp = D.macs.rows.filter(r=>r['Has IP']==='No').length;
  if(noIp>0)
    items.push({c:'var(--cyan)',icon:'🔗',
      title:noIp.toLocaleString()+' Ports Without an IP',
      desc:'No ARP entry found. These devices may not be reachable at L3 or are purely L2.',
      cnt:null});
  const disabledSw = D.devices.rows.filter(r=>Number(r['Disabled Interfaces'])>0);
  if(disabledSw.length>0){
    const tot=disabledSw.reduce((a,r)=>a+Number(r['Disabled Interfaces']||0),0);
    items.push({c:'var(--muted)',icon:'🔒',
      title:tot.toLocaleString()+' Admin-Down Interfaces',
      desc:'Across '+disabledSw.length+' switches. Confirm these are intentionally disabled before migration.',
      cnt:null});
  }
  if(!items.length)
    items.push({c:'var(--green)',icon:'✅',
      title:'No issues detected',
      desc:'All devices classified, no conflicts found.',cnt:null});
  const list=document.getElementById('attn-list');
  list.innerHTML=items.map(it=>`
    <div class="attn-item" style="--c:${it.c}">
      <span class="ai-icon">${it.icon}</span>
      <div class="ai-text">
        <div class="ai-title">${it.title}</div>
        <div class="ai-desc">${it.desc}</div>
      </div>
    </div>`).join('');
})();

// ── Device breakdown list ─────────────────────────────────────────────────
(function(){
  const labels=D.charts.deviceTypes.labels;
  const vals  =D.charts.deviceTypes.values;
  const total =vals.reduce((a,b)=>a+b,0)||1;
  const max   =Math.max(...vals)||1;
  const listEl=document.getElementById('dtype-list');
  if(!labels.length){listEl.innerHTML='<div style="color:var(--muted);font-size:.78rem">No data</div>';return;}
  listEl.innerHTML=labels.map((l,i)=>{
    const c=DTYPE_CLR[l]||PAL[i%PAL.length];
    const pct=Math.round(100*vals[i]/total);
    const barW=Math.round(100*vals[i]/max);
    return `<div class="dtype-row">
      <div class="dtype-dot" style="background:${c}"></div>
      <div class="dtype-name">${l}</div>
      <div class="dtype-bar-bg"><div class="dtype-bar-fill" style="width:${barW}%;background:${c}"></div></div>
      <div class="dtype-cnt">${vals[i].toLocaleString()}</div>
      <div class="dtype-pct">${pct}%</div>
    </div>`;
  }).join('');
})();

// ── Charts ────────────────────────────────────────────────────────────────
const DONUT={responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'right',labels:{boxWidth:10,font:{size:11},padding:7}}}};

function hBar(id,wId,labels,datasets,stacked=false){
  const n=labels.length;
  const h=Math.max(200,n*22+50);
  if(wId)document.getElementById(wId).style.height=h+'px';
  new Chart(document.getElementById(id),{type:'bar',data:{labels,datasets},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:datasets.length>1,position:'top',
                       labels:{boxWidth:10,font:{size:11}}}},
      scales:{x:{stacked,beginAtZero:true,grid:{color:'#f1f5f9'},
                 ticks:{font:{size:11}}},
              y:{stacked,ticks:{font:{size:11}},grid:{display:false}}}}});
}

if(D.charts.deviceTypes.labels.length)
  new Chart(document.getElementById('ch-dtype'),{type:'doughnut',
    data:{labels:D.charts.deviceTypes.labels,
          datasets:[{data:D.charts.deviceTypes.values,
            backgroundColor:D.charts.deviceTypes.labels.map(l=>DTYPE_CLR[l]||'#94a3b8'),
            borderWidth:2,hoverOffset:4}]},options:DONUT});

if(D.charts.uplink.labels.length)
  new Chart(document.getElementById('ch-upl'),{type:'doughnut',
    data:{labels:D.charts.uplink.labels,
          datasets:[{data:D.charts.uplink.values,
            backgroundColor:D.charts.uplink.labels.map(l=>UPL_CLR[l]||'#94a3b8'),
            borderWidth:2,hoverOffset:4}]},options:DONUT});

if(D.charts.macPerSwitch.labels.length)
  hBar('ch-sw','w-sw',D.charts.macPerSwitch.labels,
    [{label:'Devices',data:D.charts.macPerSwitch.values,
      backgroundColor:'#2563eb44',borderColor:'#2563eb',borderWidth:1.5,
      borderRadius:3}]);

if(D.charts.buildings.labels.length)
  hBar('ch-bld','w-bld',D.charts.buildings.labels,
    [{label:'Devices',data:D.charts.buildings.values,
      backgroundColor:'#7c3aed44',borderColor:'#7c3aed',borderWidth:1.5,
      borderRadius:3}]);

if(D.charts.interfaces.labels.length)
  hBar('ch-ifc','w-ifc',D.charts.interfaces.labels,[
    {label:'Active',   data:D.charts.interfaces.active,   backgroundColor:'#10b98166',borderRadius:2},
    {label:'Inactive', data:D.charts.interfaces.inactive, backgroundColor:'#f59e0b66',borderRadius:2},
    {label:'Disabled', data:D.charts.interfaces.disabled, backgroundColor:'#ef444466',borderRadius:2},
  ],true);

// ── Misc ──────────────────────────────────────────────────────────────────
const NUM_COLS=new Set(['Total MACs','Total Subnets (C)','Active Interfaces',
  'Inactive Interfaces','Disabled Interfaces','MAC Count','Pkts In','Pkts Out',
  'Switches/Routers','Wireless APs','IP Phones','Video Conf','Printers',
  'IP Cameras','Servers','Virtual Machines','Workstations','Unknown Devices','VLAN']);

function _esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

// ── TableView ─────────────────────────────────────────────────────────────
function TableView(cfg){
  const self=this; self._built=false;
  let _cols,_rows,_vis,_sortCol=null,_sortDir=1,
      _filters={},_srch='',_pg=1,_pgSz=50,_filtered=[];

  self.setPageSize=function(v){_pgSz=parseInt(v);_pg=1;_render();};

  self.build=function(){
    self._built=true;
    _cols=cfg.data.columns; _rows=cfg.data.rows;
    _vis=new Set(_cols.filter(c=>!(cfg.data.defaultHidden||[]).includes(c)));

    // Build filter panel
    if(cfg.filterEl&&cfg.data.filters){
      const wrap=document.getElementById(cfg.filterEl);
      // keep the label span, append after it
      Object.entries(cfg.data.filters).forEach(([col,vals])=>{
        if(!vals||!vals.length)return;
        const div=document.createElement('div');
        div.className='fp-field';
        div.innerHTML=`<label class="flabel">${col}</label>`;
        const sel=document.createElement('select'); sel.className='flt';
        sel.innerHTML=`<option value="">All</option>`+
          vals.map(v=>`<option>${_esc(v)}</option>`).join('');
        sel.addEventListener('change',()=>{_filters[col]=sel.value;_pg=1;_render();});
        div.appendChild(sel); wrap.appendChild(div);
      });
      // Search at end of filter panel
      const sdiv=document.createElement('div');
      sdiv.className='fp-field';
      sdiv.innerHTML=`<label class="flabel">Search</label>`;
      const sinp=document.createElement('input');
      sinp.type='search'; sinp.className='srch';
      sinp.placeholder=cfg.searchPlaceholder||'Search…';
      sinp.addEventListener('input',e=>{_srch=e.target.value.toLowerCase();_pg=1;_render();});
      sdiv.appendChild(sinp); wrap.appendChild(sdiv);
    }

    // Column toggle
    if(cfg.colMenuEl){
      const menu=document.getElementById(cfg.colMenuEl);
      // keep dd-hd div if present
      const hd=menu.querySelector('.dd-hd');
      _cols.forEach(c=>{
        const lbl=document.createElement('label');
        const cb=document.createElement('input'); cb.type='checkbox'; cb.checked=_vis.has(c);
        cb.addEventListener('change',()=>{
          if(cb.checked)_vis.add(c);else _vis.delete(c);
          _rebuildHead();_render();
        });
        lbl.appendChild(cb); lbl.appendChild(document.createTextNode(' '+c));
        menu.appendChild(lbl);
      });
    }

    _rebuildHead(); _render();
  };

  function _rebuildHead(){
    const tbl=document.getElementById(cfg.id);
    const thead=tbl.querySelector('thead');
    const visCol=_cols.filter(c=>_vis.has(c));
    thead.innerHTML='';
    const tr=document.createElement('tr');
    visCol.forEach(c=>{
      const th=document.createElement('th'); th.textContent=c;
      if(c===_sortCol)th.className=_sortDir===1?'s-asc':'s-desc';
      th.addEventListener('click',()=>{
        _sortDir=c===_sortCol?-_sortDir:1; _sortCol=c; _pg=1;
        thead.querySelectorAll('th').forEach(t=>t.className='');
        th.className=_sortDir===1?'s-asc':'s-desc'; _render();
      });
      tr.appendChild(th);
    });
    thead.appendChild(tr);
  }

  function _filter_rows(){
    return _rows.filter(row=>{
      if(_srch&&!Object.values(row).some(v=>String(v).toLowerCase().includes(_srch)))return false;
      for(const [col,val]of Object.entries(_filters)){
        if(!val)continue;
        if(String(row[col]||'')!==val)return false;
      }
      if(cfg.extraFilter&&!cfg.extraFilter(row))return false;
      return true;
    });
  }

  function _sort(arr){
    if(!_sortCol)return arr;
    return [...arr].sort((a,b)=>{
      let va=a[_sortCol]??'', vb=b[_sortCol]??'';
      const na=Number(va),nb=Number(vb);
      if(!isNaN(na)&&!isNaN(nb)&&va!==''&&vb!==''){va=na;vb=nb;}
      else{va=String(va).toLowerCase();vb=String(vb).toLowerCase();}
      return(va<vb?-1:va>vb?1:0)*_sortDir;
    });
  }

  function _render(){
    _filtered=_sort(_filter_rows());
    const total=_filtered.length;
    const pages=Math.max(1,Math.ceil(total/_pgSz));
    if(_pg>pages)_pg=pages;
    const start=(_pg-1)*_pgSz,end=Math.min(_pg*_pgSz,total);
    const slice=_filtered.slice(start,end);
    const visCol=_cols.filter(c=>_vis.has(c));
    const tbody=document.getElementById(cfg.id).querySelector('tbody');
    tbody.innerHTML='';
    if(!slice.length){
      const tr=document.createElement('tr');
      const td=document.createElement('td'); td.colSpan=visCol.length;
      td.className='no-data'; td.textContent='No matching records';
      tbody.appendChild(tr); tr.appendChild(td);
    } else {
      slice.forEach(row=>{
        const tr=document.createElement('tr');
        if(cfg.rowCls)tr.className=cfg.rowCls(row);
        visCol.forEach(c=>{
          const td=document.createElement('td');
          const v=row[c];
          if(cfg.cellFn){
            const r=cfg.cellFn(c,v,row,td);
            if(r!==undefined)td.innerHTML=r;
            else td.textContent=(v===''||v==null)?'':v;
          } else td.textContent=(v===''||v==null)?'':v;
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
    }
    const infoEl=document.getElementById(cfg.infoEl);
    if(infoEl)infoEl.textContent=
      total.toLocaleString()+' row'+(total!==1?'s':'')+
      (_rows.length!==total?' (of '+_rows.length.toLocaleString()+' total)':'')+
      (total?'  ·  Showing '+(start+1)+'–'+end:'');
    const pagerEl=document.getElementById(cfg.pagerEl);
    if(pagerEl){
      pagerEl.innerHTML='';
      const pb=(lbl,pg,dis,cur)=>{
        const b=document.createElement('button'); b.textContent=lbl; b.disabled=dis;
        if(cur)b.className='cur';
        b.addEventListener('click',()=>{_pg=pg;_render();});
        pagerEl.appendChild(b);
      };
      pb('«',1,_pg===1);pb('‹',_pg-1,_pg===1);
      const lo=Math.max(1,_pg-2),hi=Math.min(pages,_pg+2);
      if(lo>1){pb('1',1,false,false);if(lo>2){const s=document.createElement('span');s.textContent='…';s.style.cssText='padding:0 4px;color:#94a3b8';pagerEl.appendChild(s);}}
      for(let p=lo;p<=hi;p++)pb(p,p,false,p===_pg);
      if(hi<pages){if(hi<pages-1){const s=document.createElement('span');s.textContent='…';s.style.cssText='padding:0 4px;color:#94a3b8';pagerEl.appendChild(s);}pb(pages,pages,false,false);}
      pb('›',_pg+1,_pg===pages);pb('»',pages,_pg===pages);
    }
  }
  self.setExtraFilter=function(fn){cfg.extraFilter=fn;_pg=1;_render();};
  self.getFiltered=function(){return _filtered;};
  self.getCols=function(){return _cols;};
  self.rerender=function(){_render();};
  self.updateRows=function(newRows){cfg.data.rows=newRows;_pg=1;if(self._built)_render();};
}

// ── Cell renderers ────────────────────────────────────────────────────────
function macCell(col,v,row,td){
  if(NUM_COLS.has(col))td.className='r';
  if(v===''||v==null)return '';
  if(col==='Device Type'){const c=DTYPE_CLR[v]||'#94a3b8';return `<span class="bdg" style="background:${c}22;color:${c}">${_esc(v)}</span>`;}
  if(col==='Uplink'){const c=UPL_CLR[v]||'#94a3b8';return `<span class="bdg" style="background:${c}22;color:${c}">${_esc(v)}</span>`;}
  if(col==='Interface Status'){
    const s=String(v).toLowerCase();
    const c=s.includes('admin')?'#94a3b8':s==='up'?'#16a34a':'#ef4444';
    return `<span style="color:${c};font-weight:700">●</span> <span style="font-size:.74rem">${_esc(v)}</span>`;}
  if(col==='Has IP'){const c=v==='Yes'?'#16a34a':'#ef4444';return `<span class="bdg" style="background:${c}22;color:${c}">${_esc(v)}</span>`;}
  if(col==='Type Conflict'&&v){td.className='conflict-cell';return '⚠ '+_esc(v);}
  if(col==='interface description')td.className='wrap';
  return undefined;
}
function macRowCls(row){return row['Type Conflict']?'row-conflict':'';}

function devCell(col,v,row,td){
  if(col==='Assigned Site'){
    const hn=_esc(String(row['Hostname']||''));
    const chipCls='site-chip her-'+_esc(String(v||'').replace(/[^A-Za-z0-9]/g,''));
    if(!v)return `<button class="btn-sm" style="padding:2px 7px;font-size:.72rem" onclick="openAssignModal('${hn}')">+ Assign</button>`;
    return `<span class="${chipCls}">${_esc(String(v))}</span>
      <button class="btn-sm" style="padding:2px 5px;font-size:.72rem;margin-left:4px" onclick="openAssignModal('${hn}')">&#x270E;</button>`;
  }
  if(NUM_COLS.has(col))td.className='r';
  if(v===''||v==null)return '';
  if(col==='Total MACs'){
    const max=Math.max(...D.devices.rows.map(r=>Number(r['Total MACs'])||0));
    const pct=max?Math.round(100*Number(v)/max):0;
    return `<div style="display:flex;align-items:center;gap:7px">
      <div style="flex:1;height:5px;background:#f1f5f9;border-radius:3px;min-width:40px">
        <div style="width:${pct}%;height:100%;background:#2563eb;border-radius:3px"></div></div>
      <span style="font-weight:700;color:#2563eb;font-size:.78rem">${_esc(String(v))}</span></div>`;}
  return undefined;
}

// ── Instantiate ───────────────────────────────────────────────────────────
const MacTV=new TableView({id:'mac-tbl',data:D.macs,
  filterEl:'mac-filters',colMenuEl:'mac-col-menu',
  searchPlaceholder:'MAC / IP / hostname / description…',
  infoEl:'mac-info',pagerEl:'mac-pager',
  cellFn:macCell,rowCls:macRowCls});
const SubTV=new TableView({id:'sub-tbl',data:D.subnets,
  filterEl:'sub-filters',
  searchPlaceholder:'Subnet / interface / switch…',
  infoEl:'sub-info',pagerEl:'sub-pager'});
// ── Inject Assigned Site column into devices data ──────────────────────────────
if(!D.devices.columns.includes('Assigned Site')){
  // Insert after 'Hostname' (position 1) so it's always visible
  const _hostIdx=D.devices.columns.indexOf('Hostname');
  const _insertAt=_hostIdx>=0?_hostIdx+1:1;
  D.devices.columns.splice(_insertAt,0,'Assigned Site');
  D.devices.rows.forEach(r=>{r['Assigned Site']=_dsmState[r['Hostname']]||'';});
  const _scu=[...new Set(Object.values(_dsmState).filter(Boolean))].sort();
  if(_scu.length) D.devices.filters['Assigned Site']=_scu;
}
const DevTV=new TableView({id:'dev-tbl',data:D.devices,
  filterEl:'dev-filters',
  searchPlaceholder:'Hostname / model / serial / building…',
  infoEl:'dev-info',pagerEl:'dev-pager',cellFn:devCell});

// ── Conflict filter ────────────────────────────────────────────────────────
let _conflictMode=false;
function conflictOnly(){
  _conflictMode=!_conflictMode;
  const btn=document.getElementById('btn-conflict');
  btn.classList.toggle('active',_conflictMode);
  MacTV.setExtraFilter(_conflictMode?(row=>!!row['Type Conflict']):null);
}

// ── Column toggle dropdown ─────────────────────────────────────────────────
function toggleColMenu(){document.getElementById('mac-col-toggle').classList.toggle('open');}
document.addEventListener('click',e=>{
  const w=document.getElementById('mac-col-toggle');
  if(w&&!w.contains(e.target))w.classList.remove('open');
});

// ── CSV export ─────────────────────────────────────────────────────────────
function exportCSV(which){
  let rows,cols;
  if(which==='mac'){rows=MacTV.getFiltered();cols=MacTV.getCols();}
  else if(which==='sub'){rows=SubTV.getFiltered();cols=SubTV.getCols();}
  else{rows=DevTV.getFiltered();cols=DevTV.getCols();}
  const ec=v=>{const s=String(v??'');return s.includes(',')||s.includes('"')||s.includes('\n')?'"'+s.replace(/"/g,'""')+'"':s;};
  const csv=[cols.map(ec).join(','),...rows.map(r=>cols.map(c=>ec(r[c]??'')).join(','))].join('\r\n');
  const a=document.createElement('a');
  a.href='data:text/csv;charset=utf-8,'+encodeURIComponent(csv);
  a.download='migration_'+which+'_export.csv'; a.click();
}

// ── Sites Management ─────────────────────────────────────────────────────────────────────────
let _siteEditCode=null; // null = adding new, string = editing existing

function _siteHerColor(her){
  if(her==='RW') return '#065f46';
  if(her==='Rogers') return '#4c1d95';
  if(her==='TBD') return '#92400e';
  return '#1e40af'; // RE default
}

function _initSites(){
  _sitesBuilt=true;
  _injectSiteFilters();
  _updateSiteStats();
  _injectGlobalSiteFilter();
  buildSites();
  // Update badge
  const b=document.getElementById('badge-sites');
  if(b) b.textContent=_sitesState.length.toLocaleString();
}

function _injectSiteFilters(){
  const hers=[...new Set(_sitesState.map(s=>s.heritage).filter(Boolean))].sort();
  const provs=[...new Set(_sitesState.map(s=>s.province).filter(Boolean))].sort();
  const hSel=document.getElementById('sf-heritage');
  const pSel=document.getElementById('sf-province');
  if(hSel){hSel.innerHTML='<option value="">All Heritage</option>'+hers.map(h=>`<option>${_esc(h)}</option>`).join('');}
  if(pSel){pSel.innerHTML='<option value="">All Provinces</option>'+provs.map(p=>`<option>${_esc(p)}</option>`).join('');}
}

function _updateSiteStats(){
  const tot=_sitesState.length;
  const provs=new Set(_sitesState.map(s=>s.province).filter(Boolean)).size;
  const asgn=_sitesState.filter(s=>s.assigned_switch_count>0).length;
  const y24=_sitesState.filter(s=>s.planned_completion&&s.planned_completion.includes('2024')).length;
  const y25=_sitesState.filter(s=>s.planned_completion&&s.planned_completion.includes('2025')).length;
  const y26=_sitesState.filter(s=>s.planned_completion&&s.planned_completion.includes('2026')).length;
  const set=function(id,v){const el=document.getElementById(id);if(el)el.textContent=v;};
  set('ss-total',tot); set('ss-provs',provs); set('ss-assigned',asgn);
  set('ss-2024',y24); set('ss-2025',y25); set('ss-2026',y26);
  const b=document.getElementById('badge-sites');if(b)b.textContent=tot.toLocaleString();
}

function _injectGlobalSiteFilter(){
  // no-op: global filter is now a free-text input, nothing to populate dynamically
}

function buildSites(){
  const her=(document.getElementById('sf-heritage')||{}).value||'';
  const prov=(document.getElementById('sf-province')||{}).value||'';
  const yr=(document.getElementById('sf-year')||{}).value||'';
  const q=((document.getElementById('sf-search')||{}).value||'').toLowerCase();
  const rows=_sitesState.filter(s=>{
    if(her && s.heritage!==her) return false;
    if(prov && s.province!==prov) return false;
    if(yr && !(s.planned_completion||'').includes(yr)) return false;
    if(q){
      const hay=(s.site_code+'|'+s.street_address+'|'+s.city+'|'+s.province).toLowerCase();
      if(!hay.includes(q)) return false;
    }
    return true;
  });
  const tb=document.getElementById('sites-tbody');
  if(!tb) return;
  if(!rows.length){tb.innerHTML='<tr><td colspan="9" style="text-align:center;padding:32px;color:var(--muted)">No sites match the current filters.</td></tr>';return;}
  tb.innerHTML=rows.map(s=>{
    const her=_esc(s.heritage||'');
    const chipCls='site-chip her-'+her.replace(/[^A-Za-z0-9]/g,'');
    const addr=_esc(s.street_address||'');
    const city=_esc(s.city||'');
    const prov=_esc(s.province||'');
    const yr=_esc(s.planned_completion||'');
    const sw=s.assigned_switch_count||0;
    const notes=_esc(s.notes||'');
    const code=_esc(s.site_code);
    return `<tr>
      <td><strong>${code}</strong></td>
      <td><span class="${chipCls}">${her||'&mdash;'}</span></td>
      <td>${addr}</td>
      <td>${city}</td><td>${prov}</td><td>${yr}</td>
      <td style="text-align:center">${sw>0?`<strong style="color:#2563eb">${sw}</strong>`:'<span style="color:var(--muted)">0</span>'}</td>
      <td style="font-size:.78rem;color:var(--muted)">${notes}</td>
      <td style="white-space:nowrap">
        <button class="btn-sm" style="padding:2px 7px;font-size:.72rem" onclick="openSiteModal('${code}')">&#x270E; Edit</button>
      </td>
    </tr>`;
  }).join('');
}

function openSiteModal(siteCode){
  _siteEditCode=siteCode;
  const del=document.getElementById('sm-delete-btn');
  if(del) del.style.display=siteCode?'':'none';
  document.getElementById('site-modal-title').textContent=siteCode?'Edit Site':'Add Site';
  let s=siteCode?_sitesState.find(x=>x.site_code===siteCode):null;
  const set=function(id,v){const el=document.getElementById(id);if(el)el.value=v||'';};
  set('sm-code',s?s.site_code:''); set('sm-heritage',s?s.heritage:'');
  set('sm-address',s?s.street_address:''); set('sm-city',s?s.city:'');
  set('sm-province',s?s.province:''); set('sm-postal',s?s.postal_code:'');
  set('sm-year',s?s.planned_completion:''); set('sm-notes',s?s.notes:'');
  const ll=document.getElementById('sm-latlon');
  if(ll) ll.textContent=s&&s.lat?`⌖ ${s.lat.toFixed(5)}, ${s.lng.toFixed(5)}`: '';
  if(siteCode)document.getElementById('sm-code').setAttribute('readonly','');
  else document.getElementById('sm-code').removeAttribute('readonly');
  document.getElementById('site-modal-overlay').classList.remove('hidden');
}

function closeSiteModal(){
  document.getElementById('site-modal-overlay').classList.add('hidden');
  _siteEditCode=null;
}

function saveSiteModal(){
  const get=function(id){const el=document.getElementById(id);return el?el.value.trim():''};
  const code=_siteEditCode||get('sm-code');
  if(!code){alert('Site Code is required.');return;}
  const obj={
    site_code:code, heritage:get('sm-heritage'),
    street_address:get('sm-address'), city:get('sm-city'),
    province:get('sm-province'), postal_code:get('sm-postal'),
    planned_completion:get('sm-year'), notes:get('sm-notes'),
    lat:null, lng:null
  };
  // Preserve lat/lng from existing if not changed
  const existing=_sitesState.find(s=>s.site_code===code);
  if(existing&&existing.lat){obj.lat=existing.lat;obj.lng=existing.lng;}
  // Update _sitesState
  const idx=_sitesState.findIndex(s=>s.site_code===code);
  if(idx>=0) Object.assign(_sitesState[idx],obj);
  else _sitesState.push(Object.assign({assigned_switch_count:0},obj));
  // Queue for export
  const pi=_sitesPending.upsert.findIndex(u=>u.site_code===code);
  if(pi>=0) _sitesPending.upsert[pi]=obj; else _sitesPending.upsert.push(obj);
  _savePending();
  _injectSiteFilters(); _injectGlobalSiteFilter(); _updateSiteStats(); buildSites();
  closeSiteModal();
}

function deleteSite(){
  if(!_siteEditCode) return;
  if(!confirm(`Delete site ${_siteEditCode}? This cannot be undone without re-running the script.`)) return;
  const code=_siteEditCode;
  const idx=_sitesState.findIndex(s=>s.site_code===code);
  if(idx>=0) _sitesState.splice(idx,1);
  _sitesPending.delete.push(code);
  // Remove from upsert queue if pending
  _sitesPending.upsert=_sitesPending.upsert.filter(u=>u.site_code!==code);
  _savePending();
  _injectSiteFilters(); _injectGlobalSiteFilter(); _updateSiteStats(); buildSites();
  closeSiteModal();
}

function lookupAddress(){
  const addr=(document.getElementById('sm-address')||{}).value||'';
  if(!addr){alert('Enter an address first.');return;}
  const url='https://nominatim.openstreetmap.org/search?format=json&limit=1&q='+encodeURIComponent(addr);
  fetch(url,{headers:{'Accept-Language':'en'}}).then(r=>r.json()).then(data=>{
    if(!data||!data.length){alert('Address not found via Nominatim. Try Google Maps.');return;}
    const r=data[0];
    const ll=document.getElementById('sm-latlon');
    if(ll) ll.textContent=`⌖ ${parseFloat(r.lat).toFixed(5)}, ${parseFloat(r.lon).toFixed(5)}`;
    // Stash lat/lng for save
    if(document.getElementById('sm-latlon')){
      document.getElementById('sm-latlon').dataset.lat=r.lat;
      document.getElementById('sm-latlon').dataset.lng=r.lon;
    }
  }).catch(e=>alert('Nominatim lookup failed: '+e));
}

function openGoogleMaps(){
  const addr=(document.getElementById('sm-address')||{}).value||'';
  window.open('https://www.google.com/maps/search/?api=1&query='+encodeURIComponent(addr),'_blank');
}

let _assignTarget=null;
let _assignAllSites=[];

function _filterAssignSites(q){
  q=(q||'').toLowerCase();
  const sel=document.getElementById('am-site-sel');
  const current=_dsmState[_assignTarget]||'';
  const filtered=_assignAllSites.filter(s=>{
    if(!q)return true;
    return (s.site_code+'|'+(s.city||'')+'|'+(s.province||'')+'|'+(s.street_address||'')).toLowerCase().includes(q);
  });
  const cnt=document.getElementById('am-count');
  if(cnt) cnt.textContent=`(${filtered.length} of ${_assignAllSites.length})`;
  sel.innerHTML='<option value="">\u2014 Unassign \u2014</option>'+filtered.map(s=>{
    const lbl=s.site_code+(s.city?' \u2014 '+s.city:'')+(s.province?' ('+s.province+')':'');
    return `<option value="${_esc(s.site_code)}"${s.site_code===current?' selected':''}>${_esc(lbl)}</option>`;
  }).join('');
}

function openAssignModal(hostname){
  _assignTarget=hostname;
  document.getElementById('am-hostname').textContent=hostname;
  // Sort: currently assigned first, then alphabetical
  const current=_dsmState[hostname]||'';
  _assignAllSites=[..._sitesState].sort((a,b)=>{
    if(a.site_code===current)return -1;
    if(b.site_code===current)return 1;
    return a.site_code.localeCompare(b.site_code);
  });
  // Clear search and populate
  const srch=document.getElementById('am-search');
  if(srch) srch.value='';
  _filterAssignSites('');
  document.getElementById('assign-modal-overlay').classList.remove('hidden');
  // Auto-focus the search box
  setTimeout(()=>{if(srch)srch.focus();},80);
}

function closeAssignModal(){
  document.getElementById('assign-modal-overlay').classList.add('hidden');
  _assignTarget=null;
}

function saveAssignModal(){
  const hn=_assignTarget;
  const newSite=document.getElementById('am-site-sel').value;
  const oldSite=_dsmState[hn]||'';
  if(oldSite===newSite){closeAssignModal();return;}
  // Update counts
  if(oldSite){
    const si=_sitesState.findIndex(s=>s.site_code===oldSite);
    if(si>=0) _sitesState[si].assigned_switch_count=Math.max(0,(_sitesState[si].assigned_switch_count||0)-1);
  }
  if(newSite){
    const si=_sitesState.findIndex(s=>s.site_code===newSite);
    if(si>=0) _sitesState[si].assigned_switch_count=(_sitesState[si].assigned_switch_count||0)+1;
    _dsmState[hn]=newSite;
  } else {
    delete _dsmState[hn];
  }
  _dsmPending[hn]=newSite;
  _savePending();
  _refreshAssignedSiteCol();
  _updateSiteStats(); buildSites();
  closeAssignModal();
}

function _refreshAssignedSiteCol(){
  D.devices.rows.forEach(r=>{r['Assigned Site']=_dsmState[r['Hostname']]||'';});
  if(DevTV._built) DevTV.rerender();
}

function exportSiteChanges(){
  const combined={
    upsert_sites: _sitesPending.upsert,
    delete_sites: _sitesPending.delete,
    device_assignments: _dsmPending
  };
  if(!combined.upsert_sites.length&&!combined.delete_sites.length&&!Object.keys(combined.device_assignments).length){
    alert('No pending changes to export. Make edits first.');return;
  }
  const json=JSON.stringify(combined,null,2);
  const a=document.createElement('a');
  a.href='data:application/json;charset=utf-8,'+encodeURIComponent(json);
  a.download='sites_changes.json'; a.click();
}

function setGlobalSite(query){
  const q=(query||'').trim().toLowerCase();
  // Clear the x-button visibility
  const xinp=document.getElementById('site-global-inp');
  if(xinp) xinp.style.borderColor=q?'#2563eb':'';
  if(!q){
    MacTV.setExtraFilter(null);
    DevTV.setExtraFilter(null);
    return;
  }
  // Match DevTV rows directly against query (hostname, model, building, snmp, assigned site)
  const matchedHostnames=new Set(
    D.devices.rows.filter(r=>{
      const hay=[
        r['Hostname']||'',
        r['Model']||'',
        r['Building']||'',
        r['SNMP Location']||'',
        r['Assigned Site']||'',
        // Also check city/province of the assigned site
        (()=>{const s=_sitesState.find(x=>x.site_code===r['Assigned Site']);return s?(s.city||'')+' '+(s.province||''):'';})()
      ].join('|').toLowerCase();
      return hay.includes(q);
    }).map(r=>r['Hostname'])
  );
  if(!matchedHostnames.size){
    DevTV.setExtraFilter(()=>false);
    MacTV.setExtraFilter(()=>false);
    return;
  }
  DevTV.setExtraFilter(row=>matchedHostnames.has(row['Hostname']));
  MacTV.setExtraFilter(row=>matchedHostnames.has(row['Switch hostname']));
}

function _savePending(){
  try{
    localStorage.setItem('_sitesPending',JSON.stringify(_sitesPending));
    localStorage.setItem('_dsmPending',JSON.stringify(_dsmPending));
  }catch(e){}
}

// \u2500\u2500 Timeline \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
let _tlBuilt=false;
function _buildTimeline(){
  _tlBuilt=true;
  const H=D.history||[];
  const badge=document.getElementById('badge-timeline');
  const resEl=document.getElementById('tl-results');
  if(!H.length){
    if(resEl) resEl.innerHTML='<div class="no-data" style="padding:48px">No audit history yet \u2014 run at least one audit to populate the database.</div>';
    return;
  }
  // Populate run selector dropdowns
  const selA=document.getElementById('tl-run-a');
  const selB=document.getElementById('tl-run-b');
  if(selA&&selB){
    selA.innerHTML=''; selB.innerHTML='';
    H.forEach((r,i)=>{
      const ts=r.run_at.slice(0,16).replace('T',' ');
      const tag=r.run_id==='current'?' \u2605 current':'';
      const lbl=ts+tag+' \u2014 '+r.source_file+' ('+r.ports.length+' ports)';
      [selA,selB].forEach(s=>{
        const o=document.createElement('option');
        o.value=i; o.textContent=lbl; s.appendChild(o);
      });
    });
    // Auto-select: last two runs (B = current, A = one before)
    if(H.length>=2){ selA.value=H.length-2; selB.value=H.length-1; }
    else { selA.value=0; selB.value=0; }
  }
  // Build filter chip buttons once
  const ftEl=document.getElementById('tl-type-filters');
  if(ftEl&&!ftEl.children.length){
    [['all','All'],['added','+ Added'],['removed','\u2212 Removed'],['changed','\u223c Changed']].forEach(([t,l])=>{
      const b=document.createElement('button');
      b.className='btn tl-fbtn'+(t==='all'?' active':'');
      b.dataset.type=t; b.textContent=l;
      b.onclick=()=>_tlFilterBtn(t);
      ftEl.appendChild(b);
    });
  }
  // Auto-run if we have at least 2 different runs
  if(H.length>=2) _doCompare();
}

let _tlActiveFilter='all';
function _tlFilterBtn(type){
  _tlActiveFilter=type;
  document.querySelectorAll('.tl-fbtn').forEach(b=>b.classList.toggle('active',b.dataset.type===type));
  document.querySelectorAll('.tl-row').forEach(r=>{
    r.style.display=(type==='all'||r.dataset.dtype===type)?'':'none';
  });
  // Hide switch headers that have no visible rows
  document.querySelectorAll('.tl-sw-hdr').forEach(hdr=>{
    let nxt=hdr.nextElementSibling;
    let any=false;
    while(nxt&&!nxt.classList.contains('tl-sw-hdr')){
      if(nxt.style.display!=='none'&&nxt.classList.contains('tl-row')) any=true;
      nxt=nxt.nextElementSibling;
    }
    hdr.style.display=(any||type==='all')?'':'none';
  });
}

const _TF=['ip_address','device_type','vlan','status','mac_address','cdp_neighbour'];
const _TL={ip_address:'IP Address',device_type:'Device Type',vlan:'VLAN',
  status:'Interface Status',mac_address:'MAC Address',cdp_neighbour:'CDP Neighbour'};

function _doCompare(){
  const H=D.history||[];
  const idxA=parseInt(document.getElementById('tl-run-a').value);
  const idxB=parseInt(document.getElementById('tl-run-b').value);
  const resEl=document.getElementById('tl-results');
  if(isNaN(idxA)||isNaN(idxB)){return;}
  if(idxA===idxB){
    resEl.innerHTML='<div class="no-data" style="padding:32px">Please select two <em>different</em> runs to compare.</div>';
    return;
  }
  const runA=H[idxA], runB=H[idxB];

  // Build per-switch port maps: { switchName -> { interface -> portObj } }
  const swPA={}, swPB={};
  (runA.ports||[]).forEach(p=>{ if(!swPA[p.switch])swPA[p.switch]={}; swPA[p.switch][p.interface]=p; });
  (runB.ports||[]).forEach(p=>{ if(!swPB[p.switch])swPB[p.switch]={}; swPB[p.switch][p.interface]=p; });

  const swA=new Set(Object.keys(swPA));
  const swB=new Set(Object.keys(swPB));
  const bothSw=[...swA].filter(s=>swB.has(s)).sort();
  const onlyA=[...swA].filter(s=>!swB.has(s)).sort();
  const onlyB=[...swB].filter(s=>!swA.has(s)).sort();

  // Diff only shared switches
  let cntAdd=0,cntDel=0,cntChg=0,cntSame=0;
  const rows=[];
  bothSw.forEach(sw=>{
    const pA=swPA[sw]||{}, pB=swPB[sw]||{};
    const allIf=new Set([...Object.keys(pA),...Object.keys(pB)]);
    allIf.forEach(iface=>{
      const a=pA[iface], b=pB[iface];
      if(!a&&b){ cntAdd++; rows.push({sw,iface,dtype:'added',before:null,after:b,fields:[]}); }
      else if(a&&!b){ cntDel++; rows.push({sw,iface,dtype:'removed',before:a,after:null,fields:[]}); }
      else{
        const chg=_TF.filter(f=>String(a[f]||'')!==String(b[f]||''));
        if(chg.length){ cntChg++; rows.push({sw,iface,dtype:'changed',before:a,after:b,fields:chg}); }
        else cntSame++;
      }
    });
  });

  // Update badge
  const badge=document.getElementById('badge-timeline');
  if(badge) badge.textContent=(cntAdd+cntDel+cntChg)||'\u2713';

  // Summary strip
  const sumEl=document.getElementById('tl-summary');
  if(sumEl){
    sumEl.style.display='flex';
    const dateA=runA.run_at.slice(0,16).replace('T',' ')+(runA.run_id==='current'?' (\u2605)':'');
    const dateB=runB.run_at.slice(0,16).replace('T',' ')+(runB.run_id==='current'?' (\u2605)':'');
    sumEl.innerHTML=`
      <div style="font-size:.78rem;color:var(--muted)">
        <strong style="color:var(--text)">${dateA}</strong>
        <span style="padding:0 8px;color:#cbd5e1">\u2192</span>
        <strong style="color:var(--text)">${dateB}</strong>
      </div>
      <div style="width:1px;background:var(--border);align-self:stretch"></div>
      <span style="font-size:.78rem"><span style="color:var(--green);font-weight:700">+${cntAdd}</span> <span style="color:var(--muted)">added</span></span>
      <span style="font-size:.78rem"><span style="color:var(--red);font-weight:700">\u2212${cntDel}</span> <span style="color:var(--muted)">removed</span></span>
      <span style="font-size:.78rem"><span style="color:var(--orange);font-weight:700">\u223c${cntChg}</span> <span style="color:var(--muted)">changed</span></span>
      <span style="font-size:.78rem"><span style="color:var(--muted);font-weight:700">${cntSame}</span> <span style="color:var(--muted)">unchanged</span></span>
      <span style="font-size:.78rem;margin-left:6px;color:var(--muted)">\u2223 ${bothSw.length} switches in scope</span>`;
  }

  // Scope notice
  const scopeEl=document.getElementById('tl-scope');
  if(scopeEl){
    const parts=[];
    if(onlyA.length) parts.push(`<span style="font-size:.76rem;color:#92400e">\u26a0\ufe0f <strong>${onlyA.length}</strong> switch${onlyA.length!==1?'es':''} only in Run A \u2014 not in B's scope, not assumed removed: <em>${_esc(onlyA.join(', '))}</em></span>`);
    if(onlyB.length) parts.push(`<span style="font-size:.76rem;color:#92400e">\u2728 <strong>${onlyB.length}</strong> new switch${onlyB.length!==1?'es':''} only in Run B \u2014 first time seen: <em>${_esc(onlyB.join(', '))}</em></span>`);
    if(parts.length){
      scopeEl.style.display='block';
      scopeEl.innerHTML='<div class="tl-scope-box">'+parts.join('')+'</div>';
    } else {
      scopeEl.style.display='none';
    }
  }

  // Apply active filter
  _tlActiveFilter='all';
  document.querySelectorAll('.tl-fbtn').forEach(b=>b.classList.toggle('active',b.dataset.type==='all'));

  if(!rows.length){
    resEl.innerHTML=`<div class="no-data" style="padding:48px">\u2705 No differences found across ${bothSw.length} shared switch${bothSw.length!==1?'es':''} between the two selected runs.</div>`;
    return;
  }

  // Helpers
  const CL={added:'var(--green)',removed:'var(--red)',changed:'var(--orange)'};
  function sdot(s){
    const sl=(s||'').toLowerCase();
    const c=sl==='up'?'var(--green)':sl.includes('admin')?'var(--muted)':'var(--red)';
    return `<span style="color:${c};font-size:.85rem">\u25cf</span> `;
  }
  function portSummary(p){
    if(!p) return '<span style="color:var(--muted)">\u2014</span>';
    let h=sdot(p.status)+'<strong>'+_esc(p.device_type||'Unknown')+'</strong> ';
    if(p.mac_address) h+=`<span style="font-size:.7rem;color:var(--muted)">${_esc(p.mac_address)}</span> `;
    if(p.ip_address)  h+=`<span style="background:#eff6ff;color:#2563eb;border-radius:3px;padding:1px 5px;font-size:.68rem">${_esc(p.ip_address)}</span> `;
    if(p.vlan)        h+=`<span style="background:#f0fdf4;color:#16a34a;border-radius:3px;padding:1px 5px;font-size:.68rem;margin-left:2px">VLAN ${_esc(p.vlan)}</span>`;
    return h;
  }

  // Group rows by switch
  const bySw={};
  bothSw.forEach(sw=>{
    const sr=rows.filter(r=>r.sw===sw);
    if(sr.length) bySw[sw]=sr;
  });
  const activeSwitch=bothSw.filter(sw=>bySw[sw]);

  let h='<div class="tbl-wrap"><table style="width:100%;border-collapse:collapse;font-size:.78rem">';
  h+='<thead><tr><th style="width:80px">Type</th><th>Interface</th><th>Before (Run A)</th><th>After (Run B)</th><th>Changed Fields</th></tr></thead><tbody>';

  activeSwitch.forEach(sw=>{
    const sr=bySw[sw];
    const na=sr.filter(r=>r.dtype==='added').length;
    const nd=sr.filter(r=>r.dtype==='removed').length;
    const nc=sr.filter(r=>r.dtype==='changed').length;
    h+=`<tr class="tl-sw-hdr"><td colspan="5"><strong>${_esc(sw)}</strong>`;
    h+=`<span style="margin-left:12px;font-size:.71rem;color:var(--muted)">`;
    if(na) h+=`<span style="color:var(--green);margin-right:8px">+${na} added</span>`;
    if(nd) h+=`<span style="color:var(--red);margin-right:8px">\u2212${nd} removed</span>`;
    if(nc) h+=`<span style="color:var(--orange)">\u223c${nc} changed</span>`;
    h+=`</span></td></tr>`;

    sr.forEach(r=>{
      const c=CL[r.dtype];
      h+=`<tr class="tl-row type-${r.dtype}" data-dtype="${r.dtype}" style="border-left:3px solid ${c}">`;
      h+=`<td><span class="bdg" style="background:${c}18;color:${c};font-size:.63rem">${r.dtype}</span></td>`;
      h+=`<td style="font-weight:600;white-space:nowrap;color:var(--text)">${_esc(r.iface)}</td>`;
      // Before column
      if(r.dtype==='added'){
        h+=`<td style="color:var(--muted)">\u2014</td>`;
      } else {
        h+=`<td>${portSummary(r.before)}`;
        if(r.before&&r.before.description) h+=`<div style="font-size:.7rem;color:var(--muted);margin-top:2px">${_esc(r.before.description)}</div>`;
        if(r.dtype==='changed'){
          r.fields.forEach(f=>{
            h+=`<div style="font-size:.72rem;margin-top:3px"><span style="color:var(--muted)">${_TL[f]}: </span><span style="color:var(--red);text-decoration:line-through">${_esc(String(r.before[f]||''))}</span></div>`;
          });
        }
        h+=`</td>`;
      }
      // After column
      if(r.dtype==='removed'){
        h+=`<td style="color:var(--muted)">\u2014</td>`;
      } else {
        h+=`<td>${portSummary(r.after)}`;
        if(r.after&&r.after.description) h+=`<div style="font-size:.7rem;color:var(--muted);margin-top:2px">${_esc(r.after.description)}</div>`;
        if(r.dtype==='changed'){
          r.fields.forEach(f=>{
            h+=`<div style="font-size:.72rem;margin-top:3px"><span style="color:var(--muted)">${_TL[f]}: </span><span style="color:var(--green);font-weight:700">${_esc(String(r.after[f]||''))}</span></div>`;
          });
        }
        h+=`</td>`;
      }
      // Changed fields summary
      h+=`<td style="font-size:.72rem;color:var(--muted)">${r.fields.length?r.fields.map(f=>_TL[f]).join(', '):'—'}</td>`;
      h+=`</tr>`;
    });
  });
  h+='</tbody></table></div>';
  resEl.innerHTML=h;
}

// \u2500\u2500 Topology \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
let _topoNet=null, _topoAllNodes=null, _topoAllEdges=null, _topoExtVisible=true, _topoIsolated=false;

function _initTopology(){
  const T=D.topology;
  const canvas=document.getElementById('topo-canvas');
  if(!T||!T.nodes||!T.nodes.length){
    canvas.innerHTML='<div class="no-data">No topology data \u2014 CDP neighbour relationships were not found in this audit file.<br>Ensure the audit script captures <code>show cdp neighbors detail</code>.</div>';
    return;
  }
  // Stats strip
  document.getElementById('ts-sw').textContent  = T.nodes.filter(n=>n.building!=='External').length;
  document.getElementById('ts-lk').textContent  = T.edges.length;
  document.getElementById('ts-bld').textContent = Object.keys(T.buildingColors||{}).length;
  // Show toolbar + header buttons
  ['topo-toolbar','topo-hdr-btns'].forEach(id=>{const el=document.getElementById(id);if(el)el.style.display='';});
  // Populate building dropdown
  const bldSel=document.getElementById('topo-bld-filter');
  if(bldSel&&T.buildingColors)
    Object.keys(T.buildingColors).forEach(b=>{const o=document.createElement('option');o.value=b;o.textContent=b;bldSel.appendChild(o);});
  // Scale node sizes by MAC count
  const allNodeData=T.nodes.map(n=>({...n,
    size:n.building==='External'?16:Math.max(20,Math.min(42,20+Math.round((n.macs||0)*0.45))),
  }));
  _topoAllNodes=new vis.DataSet(allNodeData);
  _topoAllEdges=new vis.DataSet(T.edges);
  const opts={
    layout:{improvedLayout:true},
    physics:{solver:'barnesHut',
      barnesHut:{gravitationalConstant:-4000,centralGravity:0.1,springLength:160,springConstant:0.04},
      stabilization:{iterations:400,updateInterval:30}},
    edges:{arrows:{to:{enabled:false}},smooth:{type:'continuous',roundness:0.22},
           font:{size:10,align:'middle',color:'#64748b',strokeWidth:2,strokeColor:'#fff'},
           selectionWidth:2.5},
    nodes:{borderWidth:2,shadow:{enabled:true,size:6,x:2,y:2},
           font:{face:'-apple-system,"Segoe UI",Roboto,sans-serif',strokeWidth:2,strokeColor:'rgba(0,0,0,0.4)'}},
    interaction:{hover:true,tooltipDelay:80,navigationButtons:true,keyboard:true,multiselect:false},
  };
  _topoNet=new vis.Network(canvas,{nodes:_topoAllNodes,edges:_topoAllEdges},opts);

  // Click: show detail + dim non-connected
  _topoNet.on('click',params=>{
    if(!params.nodes.length){
      _topoAllNodes.forEach(n=>_topoAllNodes.update({id:n.id,opacity:1}));
      _topoAllEdges.forEach(e=>_topoAllEdges.update({id:e.id,color:{color:'#64748b',highlight:'#2563eb'},width:2}));
      document.getElementById('topo-detail').innerHTML='<span style="color:var(--muted)">Click a switch to see details</span>';
      return;
    }
    const nid=params.nodes[0];
    _showTopoDetail(nid);
    const connE=_topoNet.getConnectedEdges(nid);
    const connN=new Set([nid]);
    connE.forEach(eid=>{const ed=_topoAllEdges.get(eid);if(ed){connN.add(ed.from);connN.add(ed.to);}});
    _topoAllNodes.forEach(n=>_topoAllNodes.update({id:n.id,opacity:connN.has(n.id)?1:0.18}));
    _topoAllEdges.forEach(e=>_topoAllEdges.update({id:e.id,
      color:{color:connE.includes(e.id)?'#2563eb':'#e2e8f0',highlight:'#2563eb'},
      width:connE.includes(e.id)?2.5:1}));
  });

  // Double-click: isolate neighbours / restore all
  _topoNet.on('doubleClick',params=>{
    if(!params.nodes.length){
      _topoIsolated=false;
      _topoAllNodes.forEach(n=>_topoAllNodes.update({id:n.id,hidden:false,opacity:1}));
      _topoAllEdges.forEach(e=>_topoAllEdges.update({id:e.id,hidden:false,color:{color:'#64748b',highlight:'#2563eb'},width:2}));
      return;
    }
    const nid=params.nodes[0];
    if(_topoIsolated){
      _topoIsolated=false;
      _topoAllNodes.forEach(n=>_topoAllNodes.update({id:n.id,hidden:false,opacity:1}));
      _topoAllEdges.forEach(e=>_topoAllEdges.update({id:e.id,hidden:false,color:{color:'#64748b',highlight:'#2563eb'},width:2}));
    } else {
      _topoIsolated=true;
      const connE=_topoNet.getConnectedEdges(nid);
      const connN=new Set([nid]);
      connE.forEach(eid=>{const ed=_topoAllEdges.get(eid);if(ed){connN.add(ed.from);connN.add(ed.to);}});
      _topoAllNodes.forEach(n=>_topoAllNodes.update({id:n.id,hidden:!connN.has(n.id),opacity:1}));
      _topoAllEdges.forEach(e=>_topoAllEdges.update({id:e.id,hidden:!connE.includes(e.id),
        color:{color:'#2563eb',highlight:'#2563eb'},width:2.5}));
      setTimeout(()=>_topoNet.fit(),60);
    }
    _showTopoDetail(nid);
  });

  // Legend
  const legend=document.getElementById('topo-legend');
  if(legend&&T.buildingColors){
    Object.entries(T.buildingColors).forEach(([bld,col])=>{
      const d=document.createElement('div');d.className='dtype-row';
      d.innerHTML=`<div class="dtype-dot" style="background:${col}"></div><div class="dtype-name">${_esc(bld)}</div>`;
      legend.appendChild(d);
    });
    const ext=document.createElement('div');ext.className='dtype-row';
    ext.innerHTML='<div class="dtype-dot" style="background:#94a3b8"></div><div class="dtype-name" style="color:var(--muted)">External / unaudited</div>';
    legend.appendChild(ext);
  }
}

function _showTopoDetail(nid){
  const T=D.topology;
  const n=T.nodes.find(x=>x.id===nid); if(!n)return;
  const connE=_topoNet.getConnectedEdges(nid);
  const nbrs=[];
  connE.forEach(eid=>{
    const ed=_topoAllEdges.get(eid); if(!ed)return;
    const peer=ed.from===nid?ed.to:ed.from;
    if(peer!==nid) nbrs.push({hostname:peer, iface:ed.label||'\u2014', vlan:ed.vlan||''});
  });
  const uptime=(n.uptime||'').split(',').slice(0,2).join(',').trim();
  const ifHtml=n.totalIf?`
  <div class="topo-iface-bar" style="margin-top:10px">
    <div class="topo-iface-cell"><div class="val" style="color:#10b981">${n.activeIf}</div><div class="lbl">Active</div></div>
    <div class="topo-iface-cell"><div class="val" style="color:#f59e0b">${n.inactiveIf}</div><div class="lbl">Inactive</div></div>
    <div class="topo-iface-cell"><div class="val" style="color:#94a3b8">${n.disabledIf}</div><div class="lbl">Disabled</div></div>
  </div>`:'';
  const dtRows=n.dtypes?Object.entries(n.dtypes).filter(([,v])=>v>0)
    .map(([k,v])=>`<div style="display:flex;justify-content:space-between;font-size:.73rem;padding:2px 0;border-bottom:1px solid #f1f5f9">
      <span style="color:var(--muted)">${_esc(k)}</span><strong>${v}</strong></div>`).join(''):'';
  const nbrHtml=nbrs.length
    ?`<div style="margin-top:10px"><div class="flabel" style="margin-bottom:6px">Connected Switches (${nbrs.length})</div>`+
      nbrs.map(nb=>`<div class="topo-nbr-item" onclick="_topoNet.selectNodes(['${_esc(nb.hostname)}']);_showTopoDetail('${_esc(nb.hostname)}')">
        <div style="width:8px;height:8px;border-radius:50%;background:#2563eb;flex-shrink:0"></div>
        <div><div style="font-weight:600">${_esc(nb.hostname)}</div>
        <div style="font-size:.68rem;color:var(--muted)">${_esc(nb.iface)}${nb.vlan?' \u00b7 VLAN '+_esc(nb.vlan):''}</div></div>
      </div>`).join('')+'</div>'
    :`<div style="margin-top:8px;font-size:.74rem;color:var(--muted)">No directly connected switches found</div>`;
  document.getElementById('topo-detail').innerHTML=`
    <div class="topo-detail-name">${_esc(n.id)}</div>
    <div class="topo-detail-grid">
      <span class="k">Model</span>     <span class="v">${_esc(n.model||'\u2014')}</span>
      <span class="k">Serial</span>    <span class="v">${_esc(n.serial||'\u2014')}</span>
      <span class="k">IOS</span>       <span class="v">${_esc(n.ios||'\u2014')}</span>
      <span class="k">Building</span>  <span class="v">${_esc(n.building||'\u2014')}</span>
      <span class="k">Floor</span>     <span class="v">${_esc(n.floor||'\u2014')}</span>
      <span class="k">Uptime</span>    <span class="v">${_esc(uptime||'\u2014')}</span>
      <span class="k">Total MACs</span><span class="v" style="color:var(--accent)">${n.macs||0}</span>
    </div>
    ${ifHtml}
    ${dtRows?`<div style="margin-top:10px"><div class="flabel" style="margin-bottom:5px">Device Types</div>${dtRows}</div>`:''}
    ${nbrHtml}`;}

function _topoLayout(val){
  if(!_topoNet)return;
  const isHier=(val==='hierarchical'||val==='hierarchicalLR');
  const dir=val==='hierarchicalLR'?'LR':'UD';
  _topoNet.setOptions({
    layout:isHier?{hierarchical:{enabled:true,direction:dir,sortMethod:'directed',levelSeparation:140,nodeSpacing:120}}:{improvedLayout:true,hierarchical:false},
    physics:{enabled:!isHier,solver:'barnesHut',
      barnesHut:{gravitationalConstant:-4000,centralGravity:0.1,springLength:160}},
  });
  if(!isHier) _topoNet.stabilize(200);
}

function _topoBldFilter(bld){
  if(!_topoAllNodes)return;
  _topoAllNodes.forEach(n=>{
    const isExt=n.building==='External';
    const show=!bld||(n.building===bld)||(isExt&&_topoExtVisible);
    _topoAllNodes.update({id:n.id,hidden:!show});
  });
  _topoAllEdges.forEach(e=>{
    const nA=_topoAllNodes.get(e.from),nB=_topoAllNodes.get(e.to);
    _topoAllEdges.update({id:e.id,hidden:!!(nA?.hidden||nB?.hidden)});
  });
  setTimeout(()=>_topoNet&&_topoNet.fit(),80);
}

function _topoToggleExt(){
  _topoExtVisible=!_topoExtVisible;
  const btn=document.getElementById('btn-topo-ext');
  if(btn)btn.classList.toggle('active',!_topoExtVisible);
  if(!_topoAllNodes)return;
  _topoAllNodes.forEach(n=>{if(n.building==='External')_topoAllNodes.update({id:n.id,hidden:!_topoExtVisible});});
  _topoAllEdges.forEach(e=>{
    const nA=_topoAllNodes.get(e.from),nB=_topoAllNodes.get(e.to);
    _topoAllEdges.update({id:e.id,hidden:!!(nA?.hidden||nB?.hidden)});
  });
  setTimeout(()=>_topoNet&&_topoNet.fit(),80);
}

function _topoSearch(q){
  if(!_topoAllNodes)return;
  if(!q){_topoAllNodes.forEach(n=>_topoAllNodes.update({id:n.id,opacity:1}));return;}
  const lq=q.toLowerCase();
  let found=null;
  _topoAllNodes.forEach(n=>{
    const match=(n.id||'').toLowerCase().includes(lq)||
                (n.model||'').toLowerCase().includes(lq)||
                (n.building||'').toLowerCase().includes(lq);
    _topoAllNodes.update({id:n.id,opacity:match?1:0.12});
    if(match&&!found)found=n.id;
  });
  if(found){
    _topoNet.selectNodes([found]);
    _topoNet.focus(found,{scale:1.4,animation:{duration:500,easingFunction:'easeInOutQuad'}});
    _showTopoDetail(found);
  }
}

function _topoExport(){
  if(!_topoNet)return;
  const a=document.createElement('a');
  a.href=_topoNet.canvas.frame.canvas.toDataURL('image/png');
  a.download='network_topology.png'; a.click();
}
</script>
</body>
</html>"""

    html = _T.replace("__DATA_JSON__", data_json)
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(html)



def process_file(input_file: str, output_file: str):

    """Process audit file and generate Excel output."""



    # --- Header banner ------------------------------------------------------

    banner = Text()

    banner.append("  CISCO  ", style=T["badge"])

    banner.append("  Migration Sheet Generator  ", style=T["banner_text"])

    console.print()

    console.print(Panel(banner, border_style=T["border"], padding=(0, 1)))

    console.print()



    # --- Read & split --------------------------------------------------------

    console.print(f"  [{T['label']}]Reading[/{T['label']}]  [{T['value']}]{input_file}[/{T['value']}] ...", end="")

    with open(input_file, 'r') as f:

        content = f.read()

    console.print(f"  [{T['done_word']}]done[/{T['done_word']}]")



    banner_re = re.compile(r'={40,}\s*\n\s*Switch\s*:\s*(\S+)', re.IGNORECASE)

    banner_matches = list(banner_re.finditer(content))



    # Build (hostname, raw_text) pairs — keep the hostname from the banner regex
    # capture group so we don't need to re-parse it later.
    _raw_sections: list = []
    for i, match in enumerate(banner_matches):
        start = match.start()
        end   = banner_matches[i + 1].start() if i + 1 < len(banner_matches) else len(content)
        _raw_sections.append((match.group(1), content[start:end]))

    # --- Duplicate hostname detection ---------------------------------------
    from collections import Counter as _Counter
    _hn_counts = _Counter(hn for hn, _ in _raw_sections)
    _dupes = {hn: cnt for hn, cnt in _hn_counts.items() if cnt > 1}
    if _dupes:
        _dupe_lines = "\n".join(
            f"    [bold yellow]{hn}[/bold yellow]  — appears [bold]{cnt}x[/bold]"
            for hn, cnt in sorted(_dupes.items())
        )
        console.print()
        console.print(
            Panel(
                f"[bold yellow]⚠  Duplicate hostnames detected in the input file[/bold yellow]\n\n"
                f"{_dupe_lines}\n\n"
                f"[dim]  merge  — keep the [bold]last[/bold] occurrence of each duplicate (recommended)\n"
                f"  keep   — include all occurrences as-is (may produce inflated data)\n"
                f"  abort  — stop processing[/dim]",
                title="[yellow] Duplicate Hostnames [/yellow]",
                border_style="yellow",
                padding=(1, 3),
            )
        )
        try:
            _choice = input("  How to handle duplicates? [merge/keep/abort]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            _choice = "merge"

        if _choice in ("a", "abort"):
            console.print("  [dim]Aborted by user.[/dim]")
            raise SystemExit(0)
        elif _choice in ("k", "keep"):
            console.print(
                f"  [dim]Keeping all occurrences — output may contain duplicate rows.[/dim]"
            )
            switch_sections = [text for _, text in _raw_sections]
        else:
            # merge = override: for each hostname keep only its LAST occurrence
            _seen: dict = {}
            for idx, (hn, text) in enumerate(_raw_sections):
                _seen[hn] = (idx, text)   # overwrites earlier entries
            # Rebuild in original order (sorted by original index) so sheet order is preserved
            switch_sections = [text for _, text in sorted(_seen.values(), key=lambda x: x[0])]
            _kept = len(switch_sections)
            _dropped = len(_raw_sections) - _kept
            console.print(
                f"  [dim]Merged — kept last occurrence of each hostname. "
                f"Dropped {_dropped} earlier duplicate section(s).[/dim]"
            )
        console.print()
    else:
        switch_sections = [text for _, text in _raw_sections]

    total_switches = len(switch_sections)

    console.print(f"  [{T['label']}]Found[/{T['label']}]  [{T['found_num']}]{total_switches}[/{T['found_num']}]  [{T['label']}]switch sections[/{T['label']}]")

    console.print()

    # --- Internet check + OUI pre-fetch -------------------------------------
    global _INTERNET_OK
    _INTERNET_OK = _check_internet()
    if _INTERNET_OK:
        console.print(f"  [{T['label']}]Internet[/{T['label']}]  [{T['done_word']}]available[/{T['done_word']}] \u2014 fetching OUI vendor names via api.maclookup.app")
        # Extract every MAC address from the raw content in one pass so we can
        # pre-fetch all unique OUI prefixes before the main processing loop.
        # Cisco dot-notation (xxxx.xxxx.xxxx) and colon/dash forms are matched.
        _mac_re = re.compile(
            r'\b([0-9a-fA-F]{4}\.[0-9a-fA-F]{4}\.[0-9a-fA-F]{4}'
            r'|[0-9a-fA-F]{2}(?::[0-9a-fA-F]{2}){5}'
            r'|[0-9a-fA-F]{2}(?:-[0-9a-fA-F]{2}){5})\b'
        )
        all_macs = _mac_re.findall(content)
        unique_oui_count = len({
            re.sub(r'[.:\-]', '', m).lower()[:6]
            for m in all_macs
            if len(re.sub(r'[.:\-]', '', m)) >= 6
        })
        console.print(f"  [{T['label']}]OUI[/{T['label']}]  [{T['found_num']}]{unique_oui_count}[/{T['found_num']}]  [{T['label']}]unique prefixes \u2014 resolving via bulk API[/{T['label']}]")
        _prefetch_ouis(all_macs, console)
        console.print()
    else:
        console.print(f"  [{T['label']}]Internet[/{T['label']}]  [yellow]unavailable[/yellow] \u2014 using built-in OUI table as fallback")
        console.print()

    # --- Process each switch with animated live table -----------------------

    data_rows          = []   # MAC / port rows       -> Sheet 1 (MAC Addresses)

    subnet_rows        = []   # Connected routes      -> Sheet 2 (Subnets)

    device_summary_rows = []  # One row per device    -> Sheet 3 (Device Summary)

    done_infos         = []



    with Progress(

        SpinnerColumn(spinner_name="dots2", style=T["spinner"]),

        TextColumn(f"[{T['task_text']}]" + "{task.description}"),

        BarColumn(bar_width=30, style=T["bar_empty"], complete_style=T["bar_fill"], finished_style=T["bar_done"]),

        TaskProgressColumn(style=T["pct"]),

        TimeElapsedColumn(),

        console=console,

        transient=False,

    ) as progress:



        task = progress.add_task("Processing switches", total=total_switches)



        for section in switch_sections:

            if not section.strip():

                progress.advance(task)

                continue



            lines  = section.split('\n')

            parsed = parse_switch_section(lines)



            if not parsed["hostname"]:

                progress.advance(task)

                continue



            progress.update(task, description=f"Processing  [{T['found_num']}]{parsed['hostname']}[/{T['found_num']}]")

            time.sleep(0.15)          # brief pause so spinner is visible



            rows, building, floor = _build_switch_rows(parsed)  # building/floor returned to avoid double parse

            data_rows.extend(rows)



            # Collect connected-route rows for the Subnets sheet

            for route in parsed["route_connected"]:

                subnet_rows.append({

                    "Switch Hostname": parsed["hostname"],

                    "Switch Model":   parsed["switch_model"],

                    "Building":       building,

                    "Floor":          floor,

                    "Subnet":         route["subnet"],

                    "Interface":      route["interface"],

                    "Route Type":     route["route_type"],

                })



            mac_total = sum(len(v) for v in parsed["mac_table"].values())

            connected_subnets = sum(1 for r in parsed["route_connected"] if r["route_type"] == "Connected")

            active_ports   = sum(1 for s in parsed["int_stats"].values() if not s["disabled"] and (s["pkts_in"] + s["pkts_out"]) > 0)

            inactive_ports = sum(1 for s in parsed["int_stats"].values() if not s["disabled"] and (s["pkts_in"] + s["pkts_out"]) == 0)

            disabled_ports = sum(1 for s in parsed["int_stats"].values() if s["disabled"])

            total_ports    = len(parsed["int_stats"])

            # Count device types from the already-built rows (deduped per MAC)
            type_counts: Counter = Counter(r["Device Type"] for r in rows)

            device_summary_rows.append({

                "Hostname":             parsed["hostname"],

                "Model":                parsed["switch_model"],

                "Serial Number":        parsed["serial_number"],

                "IOS Version":          parsed["ios_version"],

                "Uptime":               parsed["uptime"],

                "Last Restart":         parsed["last_restart"],

                "SNMP Location":        parsed["snmp_location"],

                "Building":             building,

                "Floor":                floor,

                "Total MACs":           mac_total,

                "Total Subnets (C)":    connected_subnets,

                "Total Interfaces":     total_ports,

                "Active Interfaces":    active_ports,

                "Inactive Interfaces":  inactive_ports,

                "Disabled Interfaces":  disabled_ports,

                # Device-type breakdown (0 when the type is absent)
                "Switches/Routers":     type_counts.get("Switch/Router",    0),

                "Wireless APs":         type_counts.get("Wireless AP",      0),

                "IP Phones":            type_counts.get("IP Phone",         0),

                "Video Conf":           type_counts.get("Video Conf",       0),

                "Printers":             type_counts.get("Printer",          0),

                "IP Cameras":           type_counts.get("IP Camera",        0),

                "Servers":              type_counts.get("Server",           0),

                "Virtual Machines":     type_counts.get("Virtual Machine",  0),

                "Workstations":         type_counts.get("Workstation",      0),

                "Unknown Devices":      type_counts.get("Unknown",          0),

            })



            done_infos.append({

                "hostname":     parsed["hostname"],

                "mac_count":    mac_total,

                "record_count": len(rows),

                "building":     building,

                "floor":        floor,

                "model":        parsed["switch_model"],

            })



            progress.advance(task)



    # --- Print per-switch results table -------------------------------------

    console.print()

    console.print(_make_switch_table(done_infos))



    # --- Write Excel (three sheets) -----------------------------------------

    df            = pd.DataFrame(data_rows)

    df_subnets    = pd.DataFrame(subnet_rows)

    df_devices    = pd.DataFrame(device_summary_rows)



    console.print(f"  [{T['label']}]Writing[/{T['label']}]  [{T['value']}]{output_file}[/{T['value']}] ...", end="")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        df.to_excel(writer,        sheet_name='MAC Addresses', index=False)

        df_subnets.to_excel(writer, sheet_name='Subnets',      index=False)

        df_devices.to_excel(writer, sheet_name='Device Summary', index=False)

        # ── Post-write: bold headers, colour, column widths, freeze row 1 ──
        from openpyxl.styles import Font, PatternFill, Alignment
        _hdr_font_w = Font(bold=True, size=10, color="FFFFFF")
        _hdr_fill   = PatternFill("solid", fgColor="1F3864")
        _hdr_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for _sname in ("MAC Addresses", "Subnets", "Device Summary"):
            _ws = writer.sheets[_sname]
            for _cell in _ws[1]:
                _cell.font      = _hdr_font_w
                _cell.fill      = _hdr_fill
                _cell.alignment = _hdr_align
            _ws.row_dimensions[1].height = 28
            for _col in _ws.columns:
                _max_w = max(
                    (len(str(_c.value)) if _c.value is not None else 0)
                    for _c in _col
                )
                _ws.column_dimensions[_col[0].column_letter].width = min(max(_max_w + 2, 8), 40)
            _ws.freeze_panes = "A2"

    console.print(f"  [{T['done_word']}]done[/{T['done_word']}]")

    # --- HTML5 interactive report -------------------------------------------
    html_file = os.path.splitext(output_file)[0] + ".html"
    console.print(f"  [{T['label']}]Writing[/{T['label']}]  [{T['value']}]{html_file}[/{T['value']}] ...", end="")
    try:
        _write_html_report(df, df_subnets, df_devices, html_file, input_file)
        console.print(f"  [{T['done_word']}]done[/{T['done_word']}]")
    except Exception as _html_err:
        console.print(f"  [yellow]warning: HTML report skipped ({_html_err})[/yellow]")
        html_file = ""



    # --- Summary panel ------------------------------------------------------

    total_records   = len(df)

    unique_switches = df["Switch hostname"].nunique() if total_records > 0 else 0

    total_subnets   = len(df_subnets[df_subnets["Route Type"] == "Connected"]) if not df_subnets.empty else 0



    summary = Table.grid(padding=(0, 3))

    summary.add_column(justify="right", style=T["sum_label"])

    summary.add_column(style="bold")

    summary.add_row("Output file",      f"[{T['sum_file']}]{output_file}[/{T['sum_file']}]")

    if html_file:
        summary.add_row("HTML report",    f"[{T['sum_file']}]{html_file}[/{T['sum_file']}]")

    summary.add_row("Switches",         f"[{T['sum_switches']}]{unique_switches}[/{T['sum_switches']}]")

    summary.add_row("MAC rows",         f"[{T['sum_rows']}]{total_records}[/{T['sum_rows']}]")

    summary.add_row("Subnets (sheet 2)",f"[{T['sum_rows']}]{total_subnets}[/{T['sum_rows']}]")

    summary.add_row("Device Summary",   f"[{T['sum_rows']}]{len(df_devices)}[/{T['sum_rows']}] devices  \u2192  sheet 3")



    console.print()

    console.print(

        Panel(

            summary,

            title=f"[{T['complete_title']}] ✔  Complete [/{T['complete_title']}]",

            border_style=T["complete_border"],

            padding=(1, 4),

        )

    )

    console.print()

    # --- Return stats for audit DB ------------------------------------------
    _conflicts = int((df["Type Conflict"] != "").sum()) if (not df.empty and "Type Conflict" in df.columns) else 0
    _buildings = (
        int(df_devices["Building"].dropna().apply(str)
            .str.strip().replace("", pd.NA).dropna().nunique())
        if (not df_devices.empty and "Building" in df_devices.columns) else 0
    )
    return {
        "html_file":       html_file,
        "total_macs":      total_records,
        "total_switches":  unique_switches,
        "total_subnets":   total_subnets,
        "total_conflicts": _conflicts,
        "total_buildings": _buildings,
        "port_rows":       df.fillna("").to_dict("records") if not df.empty else [],
    }


# ── Local audit database ─────────────────────────────────────────────────────
_DB_NAME = "migration_audit.db"


def _db_path() -> str:
    """Always stores the DB next to this script."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), _DB_NAME)


def _db_backup(keep: int = 30) -> None:
    """Copy the DB into ./dbbackup/ with a timestamp filename.

    Keeps the *keep* most-recent backups and silently removes older ones.
    Does nothing if the DB file does not yet exist (first run).
    """
    src = _db_path()
    if not os.path.exists(src):
        return
    backup_dir = os.path.join(os.path.dirname(src), "dbbackup")
    os.makedirs(backup_dir, exist_ok=True)
    ts  = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(backup_dir, f"migration_audit_{ts}.db")
    import shutil as _shutil
    _shutil.copy2(src, dst)
    # Prune: keep only the newest `keep` backups
    backups = sorted(
        (f for f in os.listdir(backup_dir) if f.startswith("migration_audit_") and f.endswith(".db")),
        reverse=True,
    )
    for old in backups[keep:]:
        try:
            os.remove(os.path.join(backup_dir, old))
        except OSError:
            pass
    console.print(f"  [dim]DB backup:[/dim]  {os.path.basename(dst)}  ({len(backups)} total, keeping {keep})")


def _db_connect() -> sqlite3.Connection:
    conn = sqlite3.connect(_db_path())
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS audit_runs (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            run_at        TEXT    NOT NULL,
            source_file   TEXT    NOT NULL,
            output_xlsx   TEXT    NOT NULL,
            output_html   TEXT,
            total_macs    INTEGER DEFAULT 0,
            total_switches INTEGER DEFAULT 0,
            total_subnets  INTEGER DEFAULT 0,
            total_conflicts INTEGER DEFAULT 0,
            total_buildings INTEGER DEFAULT 0
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS audit_port_snapshots (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id        INTEGER NOT NULL REFERENCES audit_runs(id),
            switch        TEXT,
            interface     TEXT,
            vlan          TEXT,
            mac_address   TEXT,
            ip_address    TEXT,
            device_type   TEXT,
            status        TEXT,
            cdp_neighbour TEXT,
            description   TEXT,
            building      TEXT,
            floor         TEXT
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_snaps_run   ON audit_port_snapshots(run_id)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_snaps_iface ON audit_port_snapshots(switch, interface)")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sites (
            id                 INTEGER PRIMARY KEY AUTOINCREMENT,
            site_code          TEXT    UNIQUE NOT NULL,
            heritage           TEXT    DEFAULT '',
            street_address     TEXT    DEFAULT '',
            city               TEXT    DEFAULT '',
            province           TEXT    DEFAULT '',
            postal_code        TEXT    DEFAULT '',
            lat                REAL,
            lng                REAL,
            planned_completion TEXT    DEFAULT '',
            notes              TEXT    DEFAULT '',
            created_at         TEXT    DEFAULT (datetime('now'))
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS device_site_map (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            switch_hostname TEXT    NOT NULL UNIQUE,
            site_id         INTEGER REFERENCES sites(id) ON DELETE SET NULL,
            assigned_at     TEXT    DEFAULT (datetime('now'))
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_dsmap ON device_site_map(site_id)")
    conn.commit()
    return conn


def _db_check_and_warn(input_file: str) -> None:
    """Print a warning if previous runs exist, showing age and stats."""
    conn = _db_connect()
    try:
        rows = conn.execute(
            "SELECT * FROM audit_runs ORDER BY run_at DESC LIMIT 5"
        ).fetchall()
    finally:
        conn.close()

    if not rows:
        console.print(
            f"  [dim]Audit DB:[/dim]  [green]No previous runs — this is the first audit.[/green]"
        )
        return

    last = rows[0]
    last_dt = datetime.datetime.fromisoformat(last["run_at"])
    now_utc = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
    age     = now_utc - last_dt
    days    = age.days
    hours   = age.seconds // 3600

    if days >= 1:
        age_str = f"{days} day{'s' if days != 1 else ''} ago"
    elif hours >= 1:
        age_str = f"{hours} hour{'s' if hours != 1 else ''} ago"
    else:
        mins = age.seconds // 60
        age_str = f"{max(mins,1)} minute{'s' if mins != 1 else ''} ago"

    console.print()
    console.print(
        Panel(
            f"[bold yellow]⚠  Existing audit data found in local database[/bold yellow]\n\n"
            f"  Last run   : [cyan]{last['run_at'][:16].replace('T',' ')} UTC[/cyan]  ([bold]{age_str}[/bold])\n"
            f"  Source     : [dim]{last['source_file']}[/dim]\n"
            f"  Switches   : [bold]{last['total_switches']}[/bold]   "
            f"MACs: [bold]{last['total_macs']}[/bold]   "
            f"Subnets: [bold]{last['total_subnets']}[/bold]   "
            f"Conflicts: [bold]{last['total_conflicts']}[/bold]\n\n"
            f"[dim]Last 5 runs:[/dim]",
            title="[yellow] Local Audit History [/yellow]",
            border_style="yellow",
            padding=(1, 3),
        )
    )

    # Print last-5 history table
    hist = Table(show_header=True, header_style="bold dim", box=None, padding=(0, 2))
    hist.add_column("#",          style="dim",    width=4)
    hist.add_column("Run (UTC)",   style="cyan",   width=17)
    hist.add_column("Source",      style="dim",    max_width=30)
    hist.add_column("Switches",    justify="right")
    hist.add_column("MACs",        justify="right")
    hist.add_column("Conflicts",   justify="right")
    for i, r in enumerate(rows, 1):
        hist.add_row(
            str(i),
            r["run_at"][:16].replace("T"," "),
            os.path.basename(r["source_file"]),
            str(r["total_switches"]),
            str(r["total_macs"]),
            str(r["total_conflicts"]),
        )
    console.print(hist)
    console.print()

    # Prompt — default is to continue
    try:
        ans = input("  Continue and add new run to history? [Y/n]: ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        ans = "y"
    if ans in ("n", "no"):
        console.print("  [dim]Aborted by user.[/dim]")
        raise SystemExit(0)
    console.print()


def _db_record_run(
    source_file: str,
    output_xlsx: str,
    output_html: str,
    total_macs: int,
    total_switches: int,
    total_subnets: int,
    total_conflicts: int,
    total_buildings: int,
) -> int:
    """Append a completed run to the local audit database. Returns the new run_id."""
    run_at = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S")
    conn = _db_connect()
    try:
        conn.execute(
            """
            INSERT INTO audit_runs
              (run_at, source_file, output_xlsx, output_html,
               total_macs, total_switches, total_subnets,
               total_conflicts, total_buildings)
            VALUES (?,?,?,?,?,?,?,?,?)
            """,
            (
                run_at,
                os.path.abspath(source_file),
                os.path.abspath(output_xlsx),
                os.path.abspath(output_html) if output_html else "",
                total_macs, total_switches, total_subnets,
                total_conflicts, total_buildings,
            ),
        )
        conn.commit()
        run_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    finally:
        conn.close()
    console.print(
        f"  [dim]Audit DB:[/dim]  run #{run_id} recorded \u2192 "
        f"[dim]{os.path.basename(_db_path())}[/dim]"
    )
    return run_id


def _db_store_snapshots(run_id: int, port_rows: list) -> None:
    """Bulk-insert per-port snapshots for a completed audit run."""
    if not port_rows:
        return
    conn = _db_connect()
    try:
        conn.executemany(
            """
            INSERT INTO audit_port_snapshots
              (run_id, switch, interface, vlan, mac_address, ip_address,
               device_type, status, cdp_neighbour, description, building, floor)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            [
                (
                    run_id,
                    str(r.get("Switch hostname", "") or ""),
                    str(r.get("interface number", "") or ""),
                    str(r.get("VLAN", "") or ""),
                    str(r.get("mac address", "") or ""),
                    str(r.get("IP Address", "") or ""),
                    str(r.get("Device Type", "") or ""),
                    str(r.get("Interface Status", "") or ""),
                    str(r.get("CDP Neighbour", "") or ""),
                    str(r.get("interface description", "") or ""),
                    str(r.get("Building", "") or ""),
                    str(r.get("Floor", "") or ""),
                )
                for r in port_rows
            ],
        )
        conn.commit()
    finally:
        conn.close()
    console.print(
        f"  [dim]Audit DB:[/dim]  {len(port_rows):,} port snapshots stored"
    )


def _db_load_history() -> list:
    """Load per-port snapshots for the last 10 audit runs, oldest-first."""
    conn = _db_connect()
    try:
        runs = conn.execute(
            "SELECT id, run_at, source_file FROM audit_runs ORDER BY run_at DESC LIMIT 10"
        ).fetchall()
        if not runs:
            return []
        history = []
        for run in reversed(runs):  # oldest first for natural timeline order
            snaps = conn.execute(
                """
                SELECT switch, interface, vlan, mac_address, ip_address,
                       device_type, status, cdp_neighbour, description, building, floor
                FROM audit_port_snapshots
                WHERE run_id = ?
                ORDER BY switch, interface
                """,
                (run["id"],),
            ).fetchall()
            history.append({
                "run_id":      run["id"],
                "run_at":      run["run_at"],
                "source_file": os.path.basename(run["source_file"]),
                "ports":       [dict(s) for s in snaps],
            })
    finally:
        conn.close()
    return history


# ─── Sites seed data (from Rogers migration plan) ───────────────────────────
_SITES_SEED = [
    # (heritage, site_code, street_address, city, province, postal_code, planned_completion)
    ("RE","HAM143",   "143 Hamilton ON N4A 2N9",                           "Hamilton",         "ON","N4A 2N9","2024"),
    ("RE","WOOD021",  "21 Ridgeway Cir, Woodstock ON N4V 1C9",             "Woodstock",        "ON","N4V 1C9","2024"),
    ("RE","BRANT023", "23 Harris Ave, Brantford ON N3R 7W5",               "Brantford",        "ON","N3R 7W5","2024"),
    ("RE","MARK210",  "Unit 1-2 210 Cochrane Dr, Markham ON L3R 8E6",      "Markham",          "ON","L3R 8E6","2024"),
    ("RE","MISS395",  "3573 Wolfedale Rd, Mississauga ON L5C 1V8",         "Mississauga",      "ON","L5C 1V8","2024"),
    ("RE","NEWM395",  "395 Mulock Dr, Newmarket ON L3Y 4P9",               "Newmarket",        "ON","L3Y 4P9","2024"),
    ("RE","QUEB150",  "150 Bd Rene-Levesque E, Quebec QC G1R 2B2",         "Quebec",           "QC","G1R 2B2","2024"),
    ("RE","LACH317",  "3172 Rue Joseph-Dubreuil, Lachine QC H8T 3H4",      "Lachine",          "QC","H8T 3H4","2024"),
    ("RE","MONT740",  "740 Rue Notre Dame O, Montreal QC H3C 1J2",         "Montreal",         "QC","H3C 1J2","2024"),
    ("RE","MONT400",  "400 Rue Bridge, Montreal QC H3K 2C3",               "Montreal",         "QC","H3K 2C3","2024"),
    ("RE","QUEB832",  "3235 1e Rue Saint-Hubert, Quebec QC J7C 4N1",       "Quebec",           "QC","J7C 4N1","2024"),
    ("RE","BLAN074",  "74 Boul De La Seigneurie, Blainville QC J7C 4N1",   "Blainville",       "QC","J7C 4N1","2024"),
    ("RE","MISS157",  "1575 Argentia Rd, Mississauga ON L5N 1B5",          "Mississauga",      "ON","L5N 1B5","2024"),
    ("RE","SCAR065",  "65 Comstock Rd, Scarborough ON M1L 2G8",            "Scarborough",      "ON","M1L 2G8","2024"),
    ("RE","TORO100",  "1000 Dupont St, Toronto ON M6H 1Z6",                "Toronto",          "ON","M6H 1Z6","2024"),
    ("RE","BRAM013",  "13 Hansen Rd S, Brampton ON L6W 3H6",               "Brampton",         "ON","L6W 3H6","2024"),
    ("RE","KESW204",  "204 Simcoe Ave Unit 12, Keswick ON L4P 1T2",        "Keswick",          "ON","L4P 1T2","2024"),
    ("RE","STTH010",  "10 John St, St Thomas ON N5P 2X3",                  "St Thomas",        "ON","N5P 2X3","2024"),
    ("RE","LOND130",  "130 Dufferin Avenue, London ON",                    "London",           "ON","","2024"),
    ("RW","SAUL023",  "23 Manitou Drive, Sault Ste Marie ON",              "Sault Ste Marie",  "ON","","2024"),
    ("RW","MISS205",  "2050 Flavelle Blvd, Mississauga ON L5K 1Z8",        "Mississauga",      "ON","L5K 1Z8","2024"),
    ("RW","VANC106",  "1067 West Cordova Street, Vancouver BC",             "Vancouver",        "BC","","2024"),
    ("RW","KELO234",  "2340 Hunter Road, Kelowna BC",                      "Kelowna",          "BC","","2024"),
    ("RW","KELO235",  "2530 Hunter Road, Kelowna BC",                      "Kelowna",          "BC","","2024"),
    ("RW","VITC861",  "861 Cloverdale Avenue, Victoria BC",                "Victoria",         "BC","","2024"),
    ("RW","NANM431",  "4316 Boban Drive, Nanaimo BC",                      "Nanaimo",          "BC","","2024"),
    ("RW","CALG630",  "834 3rd Avenue SW, Calgary AB",                     "Calgary",          "AB","","2024"),
    ("RW","CALG240",  "2400-32 Avenue NE, Calgary AB",                     "Calgary",          "AB","","2024"),
    ("RW","CALG362",  "3624-23 Street NE, Calgary AB",                     "Calgary",          "AB","","2024"),
    ("RW","CALG363",  "3636-23 Street NE, Calgary AB",                     "Calgary",          "AB","","2024"),
    ("RW","CALG640",  "4950-47 Street NE, Calgary AB (NDC)",               "Calgary",          "AB","","2024"),
    ("RW","WINN020",  "20 Scurfield Blvd, Winnipeg MB",                    "Winnipeg",         "MB","","2024"),
    ("RW","WINN022",  "22 Scurfield Blvd, Winnipeg MB",                    "Winnipeg",         "MB","","2024"),
    ("RE","TORO333",  "333 Bloo OMP ISB, Toronto ON M4Y 2Y5",              "Toronto",          "ON","M4Y 2Y5","2024/2025"),
    ("RE","TORO334",  "OMP ISB, Toronto ON",                               "Toronto",          "ON","","2025"),
    ("RE","CALG075",  "475 Richmond Rd, Ottawa ON K2A 3Y8",                "Ottawa",           "ON","K2A 3Y8","2025"),
    ("RE","KITC085",  "85 Grand Crest Pl, Kitchener ON N2C 2L6",           "Kitchener",        "ON","N2C 2L6","2025"),
    ("RE","LOND800",  "800 York St, London ON N5W 2S9",                    "London",           "ON","N5W 2S9","2025"),
    ("RE","KANA306",  "306 Legget Dr, Kanata ON K2K 1Y6",                  "Kanata",           "ON","K2K 1Y6","2025"),
    ("RE","MARK045",  "45 Evna Park Dr, Markham ON L3R 1C9",               "Markham",          "ON","L3R 1C9","2025"),
    ("RE","EDMO103",  "3915 Jasper Ave, Edmonton AB T5J 3N6",              "Edmonton",         "AB","T5J 3N6","2025"),
    ("RE","REGI192",  "1920 Broad St Rt 1400, Regina SK S4P 0A5",          "Regina",           "SK","S4P 0A5","2025"),
    ("RE","STJO022",  "22 Austin St, St Johns NL A1B 4C2",                 "St Johns",         "NL","A1B 4C2","2025"),
    ("RE","STJO541",  "541 Kenmount, St Johns NL",                         "St Johns",         "NL","","2025"),
    ("RE","VANC180",  "180 W 2nd Ave, Vancouver BC V5Y 3T9",               "Vancouver",        "BC","V5Y 3T9","2025"),
    ("RE","OTTA181",  "1810 St Laurent Blvd, Ottawa ON K1G 3P2",           "Ottawa",           "ON","K1G 3P2","2025"),
    ("RE","KITC230",  "235 The Boardwalk Unit 2, Kitchener ON N2N 0B1",    "Kitchener",        "ON","N2N 0B1","2025"),
    ("RE","FRED377",  "377 York St, Fredericton NB E3B 3P6",               "Fredericton",      "NB","E3B 3P6","2025"),
    ("RE","MISS095",  "95 Topflight Dr, Mississauga ON L5S 1Y1",           "Mississauga",      "ON","L5S 1Y1","2025"),
    ("RE","MONT120",  "1200 McGill College Ave, Montreal QC H3B 4G7",      "Montreal",         "QC","H3B 4G7","2025"),
    ("RE","SUDB880",  "880 LaSalle Blvd, Sudbury ON",                      "Sudbury",          "ON","","2025"),
    ("RE","MISS688",  "6885 Kennedy Rd, Mississauga ON L5T 2R6",           "Mississauga",      "ON","L5T 2R6","2025"),
    ("RE","SAIN055",  "50 Waterloo St, Saint John NB",                     "Saint John",       "NB","","2025"),
    ("RE","HALI608",  "Young Tower 6080 Younge St, Halifax NS B3K 5L2",    "Halifax",          "NS","B3K 5L2","2024/2025"),
    ("RE","SYDN131",  "1318 Grand Lake Rd, Sydney NS",                     "Sydney",           "NS","","2024/2025"),
    ("RE","GORM04",   "4 Fortecon Dr, Gormley ON L4A 2G8",                 "Gormley",          "ON","L4A 2G8","2025"),
    ("RW","VICT302",  "751 Enterprise Crescent, Victoria BC V8Z 6P7",      "Victoria",         "BC","V8Z 6P7","2025"),
    ("RW","ABBO314",  "31450 Marshall Rd, Abbotsford BC V2T 6B1",          "Abbotsford",       "BC","V2T 6B1","2025"),
    ("RW","SURR104",  "10045 138 St, Surrey BC V3T 4K4",                   "Surrey",           "BC","V3T 4K4","2025"),
    ("RW","COUR159",  "1591 McPhee Ave, Courtenay BC V9N 3A5",             "Courtenay",        "BC","V9N 3A5","2025"),
    ("RW","KAML180",  "180 Briar Ave, Kamloops BC V2B 1C1",                "Kamloops",         "BC","V2B 1C1","2025"),
    ("RW","PORT427",  "4278 8th Ave, Port Alberni BC",                     "Port Alberni",     "BC","","2025"),
    ("RW","SQUA110",  "1103 Magee St, Squamish BC V8B 0E8",                "Squamish",         "BC","V8B 0E8","2025"),
    ("RW","PORT182",  "1820 Kingsway Ave, Port Coquitlam BC",              "Port Coquitlam",   "BC","","2025"),
    ("RW","REDD476",  "4761 62 St, Red Deer AB T4N 2R4",                   "Red Deer",         "AB","T4N 2R4","2025"),
    ("RW","FORT208",  "208 Beacon Hill Dr, Fort McMurray AB T9H 2R1",      "Fort McMurray",    "AB","T9H 2R1","2025"),
    ("RW","CHIL877",  "675 Nowell St, Chilliwack BC V2P 7G7",              "Chilliwack",       "BC","V2P 7G7","2025"),
    ("RW","BURN516",  "5161 Byrne Rd, Burnaby BC V5J 3H6",                 "Burnaby",          "BC","V5J 3H6","2025"),
    ("RW","BURN381",  "3811 N Fraser Wy, Burnaby BC V5J 5J2",              "Burnaby",          "BC","V5J 5J2","2025"),
    ("RW","SASK232",  "2326 Hanselman Ave, Saskatoon SK S7L 5X2",          "Saskatoon",        "SK","S7L 5X2","2025"),
    ("RW","NVAN147",  "1471 Pemberton Ave, North Vancouver BC V7P 2R9",    "North Vancouver",  "BC","V7P 2R9","2025"),
    ("RW","CALG022",  "2 Midpark Blvd SE, Calgary AB",                     "Calgary",          "AB","","2025"),
    ("RW","SURR303",  "3033 King George Blvd 19, Surrey BC V4P 1B8",       "Surrey",           "BC","V4P 1B8","2025"),
    ("Rogers","TBD001","15 Davis Way, Rosser ON",                          "Rosser",           "ON","","2025/2026"),
    ("Rogers","TBD002","21 Port Perry, Ontario L9L 1B5",                   "Port Perry",       "ON","L9L 1B5","2025/2026"),
    ("Rogers","TBD003","227 38 Ave NE, Calgary AB T2E 2M3",               "Calgary",          "AB","T2E 2M3","2025/2026"),
    ("RE","MON800",   "800 rue de La Gauchetiere Ouest, Montreal QC",      "Montreal",         "QC","","2026"),
    ("RE","RICH234",  "22-392 Neekir, Richmond BC",                        "Richmond",         "BC","","2026"),
    ("RE","MISS059",  "59 Ambassador, Mississauga ON",                     "Mississauga",      "ON","","2026"),
    ("RE","VAN244",   "2440 Ash St, Vancouver BC",                         "Vancouver",        "BC","","2026"),
    ("RE","VAUG781",  "7810 Keele St, Vaughan ON",                         "Vaughan",          "ON","","2026"),
    ("RE","BARR001",  "1 Sperling, Barrie ON",                             "Barrie",           "ON","","2026"),
    ("RE","NORT010",  "10 Dyas Road, North York ON",                       "North York",       "ON","","2026"),
    ("RE","OSHA301",  "301 Marwood Drive, Oshawa ON",                      "Oshawa",           "ON","","2026"),
    ("RE","YORK035",  "35 Scarlett Rd, York ON",                           "York",             "ON","","2026"),
    ("RE","HALI707",  "721 Bayers Rd, Halifax NS",                         "Halifax",          "NS","","2026"),
    ("RE","MARK550",  "550 Cochrane Dr, Markham ON",                       "Markham",          "ON","","2026"),
    ("RE","MISS125",  "1256 Crest Lawn, Mississauga ON",                   "Mississauga",      "ON","","2026"),
    ("RE","OTTA360",  "360 Albert Street, Ottawa ON",                      "Ottawa",           "ON","","2026"),
    ("RE","NORT273",  "273 Main Street, North Bay ON",                     "North Bay",        "ON","","2026"),
    ("RE","VANC301",  "301 Industrial Ave, Vancouver BC",                  "Vancouver",        "BC","","2026"),
    ("RE","QUEB210",  "210 Pierre Bertrand Bureau 200, Quebec QC",         "Quebec",           "QC","","2026"),
    ("RE","WINN330",  "330 Portage Ave, Winnipeg MB",                      "Winnipeg",         "MB","","2026"),
    ("RE","KANA436",  "436 Hazeldean, Kanata ON",                          "Kanata",           "ON","","2026"),
    ("RE","PRIN253",  "2537 Queensway Street, Prince George BC",           "Prince George",    "BC","","2026"),
    ("RE","DUNC035",  "35 Queens Road, Duncan BC",                         "Duncan",           "BC","","2026"),
    ("RE","MOOS301",  "201 Moose St, Moose Jaw SK",                        "Moose Jaw",        "SK","","2026"),
    ("RE","CAST195",  "1951 Columbia Avenue, Castlegar BC",                "Castlegar",        "BC","","2026"),
    ("RE","CRAN720",  "720 Kootenay Street North, Cranbrook BC",           "Cranbrook",        "BC","","2026"),
    ("RW","PENT001",  "2132 Fairview Road, Penticton BC",                  "Penticton",        "BC","","2026"),
    ("RW","VERN001",  "2924 28 Avenue, Vernon BC",                         "Vernon",           "BC","","2026"),
    ("RW","THUN001",  "1635 Paquette Road, Thunder Bay ON",                "Thunder Bay",      "ON","","2026"),
    ("RW","CAMP001",  "500 Robron Road, Campbell River BC",                "Campbell River",   "BC","","2026"),
    ("RW","MILE001",  "101 Mile Highway 97, 100 Mile House BC",            "100 Mile House",   "BC","","2026"),
    ("RW","DRYD001",  "75 Queen Street, Dryden ON",                        "Dryden",           "ON","","2026"),
    ("RW","FLIN001",  "35 3rd Avenue, Flin Flon MB",                       "Flin Flon",        "MB","","2026"),
    ("RW","STFR001",  "1037 First Street East, Fort Frances ON",           "Fort Frances",     "ON","","2026"),
    ("RW","NELS001",  "314 Hall Street, Nelson BC",                        "Nelson",           "BC","","2026"),
    ("RW","PALB001",  "3904 4th Avenue West, Prince Albert SK",             "Prince Albert",    "SK","","2026"),
    ("RW","HARB001",  "10 Harbourfront Drive NE, Salmon Arm BC",           "Salmon Arm",       "BC","","2026"),
    ("RW","SELK001",  "186 Main Street, Selkirk MB",                       "Selkirk",          "MB","","2026"),
    ("RW","TERR001",  "1 Headend Road 901 Hwy 17, Terrace Bay ON",         "Terrace Bay",      "ON","","2026"),
    ("RW","PELA001",  "10 Penn Lake Road, Marathon ON",                    "Marathon",         "ON","","2026"),
    ("RW","WAWA001",  "131 Mills Drive, Wawa ON",                          "Wawa",             "ON","","2026"),
    ("RW","LETH001",  "1232 3 Avenue South, Lethbridge AB",                "Lethbridge",       "AB","","2026"),
    ("RW","KENO001",  "102 10th Street, Kenora ON",                        "Kenora",           "ON","","2026"),
    ("RW","LLOY001",  "5402 50 Avenue, Lloydminster AB",                   "Lloydminster",     "AB","","2026"),
    ("RW","CANM001",  "715 Railway Avenue, Canmore AB",                    "Canmore",          "AB","","2026"),
    ("RW","DAWS001",  "1044 1445 102 Avenue, Dawson Creek BC",             "Dawson Creek",     "BC","","2026"),
    ("RW","GRAN001",  "7474 19th Street, Grand Forks BC",                  "Grand Forks",      "BC","","2026"),
    ("RW","HINT001",  "104 Wanyandi Avenue, Hinton AB",                    "Hinton",           "AB","","2026"),
]


def _db_seed_sites() -> None:
    """Seed the sites table from _SITES_SEED if it is empty."""
    conn = _db_connect()
    try:
        cnt = conn.execute("SELECT COUNT(*) FROM sites").fetchone()[0]
        if cnt > 0:
            return
        conn.executemany(
            """INSERT OR IGNORE INTO sites
               (heritage, site_code, street_address, city, province, postal_code, planned_completion)
               VALUES (?,?,?,?,?,?,?)""",
            _SITES_SEED,
        )
        conn.commit()
        seeded = conn.execute("SELECT COUNT(*) FROM sites").fetchone()[0]
        console.print(f"  [dim]Sites:[/dim]  Seeded {seeded} sites into database")
    finally:
        conn.close()


def _db_apply_sites_changes(changes_file: str) -> None:
    """Apply a sites_changes.json produced by the HTML report into the DB."""
    import json as _jchg
    if not os.path.exists(changes_file):
        return
    try:
        with open(changes_file, encoding="utf-8") as fh:
            changes = _jchg.load(fh)
    except Exception as ex:
        console.print(f"  [yellow]Sites:[/yellow]  Could not parse {os.path.basename(changes_file)}: {ex}")
        return
    conn = _db_connect()
    try:
        for s in changes.get("upsert_sites", []):
            conn.execute("""
                INSERT INTO sites (site_code, heritage, street_address, city, province,
                                   postal_code, lat, lng, planned_completion, notes)
                VALUES (:site_code,:heritage,:street_address,:city,:province,
                        :postal_code,:lat,:lng,:planned_completion,:notes)
                ON CONFLICT(site_code) DO UPDATE SET
                    heritage=excluded.heritage, street_address=excluded.street_address,
                    city=excluded.city, province=excluded.province,
                    postal_code=excluded.postal_code, lat=excluded.lat, lng=excluded.lng,
                    planned_completion=excluded.planned_completion, notes=excluded.notes
            """, {"city":"","postal_code":"","lat":None,"lng":None,"notes":"", **s})
        for code in changes.get("delete_sites", []):
            conn.execute("DELETE FROM device_site_map WHERE site_id=(SELECT id FROM sites WHERE site_code=?)", (code,))
            conn.execute("DELETE FROM sites WHERE site_code=?", (code,))
        for hostname, site_code in changes.get("device_assignments", {}).items():
            if site_code:
                row = conn.execute("SELECT id FROM sites WHERE site_code=?", (site_code,)).fetchone()
                if row:
                    conn.execute("""
                        INSERT INTO device_site_map (switch_hostname, site_id)
                        VALUES (?,?)
                        ON CONFLICT(switch_hostname) DO UPDATE SET
                            site_id=excluded.site_id, assigned_at=datetime('now')
                    """, (hostname, row["id"]))
            else:
                conn.execute("DELETE FROM device_site_map WHERE switch_hostname=?", (hostname,))
        conn.commit()
        nu = len(changes.get("upsert_sites", []))
        nd = len(changes.get("delete_sites", []))
        na = len(changes.get("device_assignments", {}))
        console.print(f"  [dim]Sites:[/dim]  Imported {nu} upserts, {nd} deletes, {na} assignments")
    finally:
        conn.close()
    os.rename(changes_file, changes_file + ".imported")


def _db_load_sites() -> list:
    """Load all sites from DB with assigned switch count."""
    conn = _db_connect()
    try:
        rows = conn.execute("""
            SELECT s.id, s.site_code, s.heritage, s.street_address, s.city,
                   s.province, s.postal_code, s.lat, s.lng,
                   s.planned_completion, s.notes,
                   COUNT(d.switch_hostname) AS assigned_switch_count
            FROM sites s
            LEFT JOIN device_site_map d ON d.site_id = s.id
            GROUP BY s.id
            ORDER BY s.planned_completion, s.province, s.site_code
        """).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def _db_load_device_site_map() -> dict:
    """Return dict: switch_hostname -> site_code."""
    conn = _db_connect()
    try:
        rows = conn.execute("""
            SELECT d.switch_hostname, s.site_code
            FROM device_site_map d
            JOIN sites s ON s.id = d.site_id
        """).fetchall()
        return {r["switch_hostname"]: r["site_code"] for r in rows}
    finally:
        conn.close()


def main():

    parser = argparse.ArgumentParser(description='Generate Cisco migration spreadsheet from audit data')

    parser.add_argument('input_file',  help='Input audit file (output.txt)')

    parser.add_argument('output_file', help='Output Excel file')

    args = parser.parse_args()

    # ── Backup the DB before any writes this run ─────────────────────────────
    _db_backup()

    # ── Init sites tables, seed, and apply any pending changes ───────────
    _db_seed_sites()
    _changes_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sites_changes.json")
    _db_apply_sites_changes(_changes_json)

    # ── Check local DB for stale data before processing ──────────────────
    _db_check_and_warn(args.input_file)

    # ── Run the main processing pipeline ─────────────────────────────────
    results = process_file(args.input_file, args.output_file)

    # ── Record this run in the local DB ──────────────────────────────────
    if results:
        _run_id = _db_record_run(
            source_file     = args.input_file,
            output_xlsx     = args.output_file,
            output_html     = results.get("html_file", ""),
            total_macs      = results.get("total_macs", 0),
            total_switches  = results.get("total_switches", 0),
            total_subnets   = results.get("total_subnets", 0),
            total_conflicts = results.get("total_conflicts", 0),
            total_buildings = results.get("total_buildings", 0),
        )
        if _run_id and results.get("port_rows"):
            _db_store_snapshots(_run_id, results["port_rows"])


if __name__ == "__main__":

    main()
