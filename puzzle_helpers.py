"""
Gedeelde hulpfuncties voor kruiswoord- en filippine-puzzelgeneratoren.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys

from PIL import ImageFont

# ---------------------------------------------------------------------------
# Lay-outconstanten
# ---------------------------------------------------------------------------

CELL_SIZE = 28   # pixels per cel
PADDING   = 30   # canvas marge rondom het rooster

# ---------------------------------------------------------------------------
# Klembord
# ---------------------------------------------------------------------------

def copy_to_clipboard(text: str) -> bool:
    """Kopieer tekst naar Windows-klembord via clip. Geeft True terug bij succes."""
    try:
        subprocess.run(
            "clip",
            input=text.encode("utf-16-le"),
            check=True,
            shell=True,
        )
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Prompt tonen + klembord
# ---------------------------------------------------------------------------

def display_prompt(prompt: str) -> None:
    """Toon een reeds gebouwde prompt in de terminal en kopieer naar klembord."""
    print("\n" + "=" * 70)
    print("STAP 2 - Kopieer de onderstaande prompt naar ChatGPT")
    print("=" * 70)
    print(prompt)
    print("=" * 70)

    if copy_to_clipboard(prompt):
        print("\n[OK] De prompt is automatisch naar je klembord gekopieerd.")
        print("     Plak hem in ChatGPT (chatgpt.com) en voer hem uit.\n")
    else:
        print("\n     Kopieer de prompt hierboven handmatig naar ChatGPT.\n")


def show_prompt(categories: str, prompt_template: str) -> None:
    """Vervang {{categories}} in het template en toon de prompt."""
    display_prompt(prompt_template.replace("{{categories}}", categories))


# ---------------------------------------------------------------------------
# JSON-respons inlezen vanuit terminal
# ---------------------------------------------------------------------------

def collect_response() -> str:
    """
    Vraag de gebruiker om de JSON-respons van ChatGPT te plakken.
    Lezen stopt zodra de gebruiker een lege regel invoert NA de JSON-inhoud.
    """
    print("STAP 3 - Plak de JSON-respons van ChatGPT hieronder.")
    print("         Druk na het plakken op Enter en daarna nogmaals Enter (lege regel).\n")

    lines: list[str] = []
    try:
        while True:
            line = input()
            if line == "" and lines:
                break
            lines.append(line)
    except EOFError:
        pass

    return "\n".join(lines).strip()


# ---------------------------------------------------------------------------
# Normaliseren
# ---------------------------------------------------------------------------

def normalize(word: str) -> str:
    """Zet speciale tekens om naar ASCII-hoofdletters en verwijder niet-A-Z tekens."""
    replacements = {
        "\u00c2": "A", "\u00c0": "A", "\u00c1": "A", "\u00c4": "A", "\u00c3": "A",
        "\u00ca": "E", "\u00c8": "E", "\u00c9": "E", "\u00cb": "E",
        "\u00ce": "I", "\u00cc": "I", "\u00cd": "I", "\u00cf": "I",
        "\u00d4": "O", "\u00d2": "O", "\u00d3": "O", "\u00d6": "O", "\u00d5": "O",
        "\u00db": "U", "\u00d9": "U", "\u00da": "U", "\u00dc": "U",
        "\u00e2": "A", "\u00e0": "A", "\u00e1": "A", "\u00e4": "A", "\u00e3": "A",
        "\u00ea": "E", "\u00e8": "E", "\u00e9": "E", "\u00eb": "E",
        "\u00ee": "I", "\u00ec": "I", "\u00ed": "I", "\u00ef": "I",
        "\u00f4": "O", "\u00f2": "O", "\u00f3": "O", "\u00f6": "O", "\u00f5": "O",
        "\u00fb": "U", "\u00f9": "U", "\u00fa": "U", "\u00fc": "U",
        "\u00e7": "C", "\u00c7": "C", "\u00f1": "N", "\u00d1": "N",
    }
    result = word.upper()
    for src, dst in replacements.items():
        result = result.replace(src.upper(), dst)
    result = re.sub(r"[^A-Z]", "", result)
    return result


# ---------------------------------------------------------------------------
# JSON parsen
# ---------------------------------------------------------------------------

def parse_response(raw: str) -> list[tuple[str, str]]:
    """
    Verwerk de JSON-respons van ChatGPT.
    Accepteert ook respons omgeven door ```json ... ``` code-blokken.
    Geeft een gesorteerde (langste eerst) lijst van (woord, omschrijving) terug.
    """
    cleaned = re.sub(r"```(?:json)?\s*", "", raw).strip().rstrip("`").strip()

    try:
        data = json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if not match:
            sys.exit("Geen geldige JSON gevonden in de geplakte respons. Start opnieuw.")
        try:
            data = json.loads(match.group())
        except json.JSONDecodeError as exc:
            sys.exit(f"JSON-fout: {exc}\nGeplakte tekst:\n{cleaned}")

    entries = data.get("woorden", [])
    if not entries:
        sys.exit(
            "Geen 'woorden'-lijst gevonden in de JSON.\n"
            "Zorg dat ChatGPT het gevraagde JSON-formaat teruggeeft."
        )

    words_clues: list[tuple[str, str]] = []
    seen: set[str] = set()
    for item in entries:
        raw_word = str(item.get("woord", "")).strip()
        clue     = str(item.get("omschrijving", "")).strip()
        word     = normalize(raw_word)
        if not word or not clue or len(word) < 3 or word in seen:
            continue
        seen.add(word)
        words_clues.append((word, clue))

    words_clues.sort(key=lambda x: -len(x[0]))
    print(f"\nVerwerkt: {len(words_clues)} woorden uit de ChatGPT-respons.")
    return words_clues


# ---------------------------------------------------------------------------
# Fonts
# ---------------------------------------------------------------------------

def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    """Laad een TrueType-font (Arial/DejaVu) of val terug op het systeemfont."""
    candidates = [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\Arial Bold.ttf" if bold else r"C:\Windows\Fonts\Arial.ttf",
    ]
    for name in ("DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
                 "DejaVuSans.ttf"):
        for base in (
            r"C:\Windows\Fonts",
            os.path.join(os.path.dirname(sys.executable), "Lib", "site-packages",
                         "PIL", "fonts"),
        ):
            candidates.append(os.path.join(base, name))

    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()
