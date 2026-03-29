"""
Filippine-puzzel generator - interactief, zonder API-sleutel.

Gebruik:
    python generate_filippine.py          # normale modus (ChatGPT copy-paste flow)
    python generate_filippine.py --test   # testmodus (hardcoded woorden)

Hoe een Filippine werkt:
  - Woorden staan in horizontale rijen (geen kruisende woorden).
  - Elk vakje toont een codenummer: hetzelfde getal = altijd dezelfde letter.
  - Optioneel: één vaste kolom leest van boven naar beneden een geheimwoord.

Output (in dezelfde map als het script):
  filippine_<YYYYMMDD_HHMMSS>_puzzel.png
  filippine_<YYYYMMDD_HHMMSS>_oplossing.png
"""

from __future__ import annotations

import os
import random
import sys
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw

from puzzle_helpers import (
    CELL_SIZE,
    PADDING,
    collect_response,
    display_prompt,
    font,
    normalize,
    parse_response,
)

# ---------------------------------------------------------------------------
# Layout-constanten
# ---------------------------------------------------------------------------

NUM_COL_W      = CELL_SIZE + 6   # breedte van het zwarte rijnummervakje
GREY_COLOR     = "#CCCCCC"
BLACK_CELL     = "#000000"
WHITE_CELL     = "#FFFFFF"
BORDER_COLOR   = "#222222"
CLUE_LINE_H    = 16
MAX_WIDTH      = 1240
CLUE_FONT_PT   = 10
NUM_FONT_PT    = 7
LET_FONT_PT    = 14
SECRET_FONT_PT = 11

_DEFAULT_GEHEIM_POS = 0   # 0-indexed; standaard = eerste letter

# ---------------------------------------------------------------------------
# ChatGPT-prompts
# ---------------------------------------------------------------------------

_PROMPT_PLAIN = """\
Je bent een puzzelmaker die Nederlandse filippine-raadsels maakt.
Genereer een lijst van precies {{n_words}} woorden en bijbehorende omschrijvingen \
voor een filippine-puzzel over de volgende categorieen: {{categories}}.

Eisen:
- Alle woorden en omschrijvingen zijn in het Nederlands.
- Woorden hebben GEEN spaties of koppeltekens (samengestelde woorden aan elkaar schrijven).
- Geen woorden met bijzondere leestekens (geen e met trema, u met dakje, etc.); \
gebruik gewone ASCII-letters.
- Elk woord heeft minimaal 4 letters.
- Omschrijvingen zijn beknopt (max. 10 woorden) en geschikt voor de doelgroep \
(scouts/zeilers, leeftijd 10-17 jaar).
- Geen duplicaten.

Geef je antwoord UITSLUITEND als geldig JSON in dit formaat, zonder extra tekst:
{{
  "woorden": [
    {{"woord": "VOORBEELD", "omschrijving": "Korte omschrijving"}},
    ...
  ]
}}
"""

_PROMPT_SECRET = """\
Je bent een puzzelmaker die Nederlandse filippine-raadsels maakt.
Genereer een lijst van precies {{n_words}} woorden en bijbehorende omschrijvingen \
voor een filippine-puzzel over de volgende categorieen: {{categories}}.

Er is een geheimwoord: "{{geheimwoord}}". Kies de woorden zo dat de \
{{geheim_pos_display}}e letter (geteld vanaf 1) van elk woord overeenkomt met \
de letters van het geheimwoord, van boven naar beneden:
{{geheim_constraints}}

Eisen:
- Alle woorden en omschrijvingen zijn in het Nederlands.
- Woorden hebben GEEN spaties of koppeltekens (samengestelde woorden aan elkaar schrijven).
- Geen woorden met bijzondere leestekens; gebruik gewone ASCII-letters.
- Elk woord heeft minimaal {{min_len}} letters, zodat de geheimkolom altijd aanwezig is.
- Omschrijvingen zijn beknopt (max. 10 woorden) en geschikt voor de doelgroep \
(scouts/zeilers, leeftijd 10-17 jaar).
- Geen duplicaten.

Geef je antwoord UITSLUITEND als geldig JSON in dit formaat, zonder extra tekst:
{{
  "woorden": [
    {{"woord": "VOORBEELD", "omschrijving": "Korte omschrijving"}},
    ...
  ]
}}
"""


def _build_prompt(categories: str, geheimwoord: str, geheim_pos: int) -> str:
    """Bouw de volledige ChatGPT-prompt op basis van de invoer."""
    if geheimwoord:
        n_words = len(geheimwoord)
        constraints = "\n".join(
            f"- Woord {i + 1}: positie {geheim_pos + 1} = '{ch}'"
            for i, ch in enumerate(geheimwoord)
        )
        prompt = (
            _PROMPT_SECRET
            .replace("{{n_words}}", str(n_words))
            .replace("{{categories}}", categories)
            .replace("{{geheimwoord}}", geheimwoord)
            .replace("{{geheim_pos_display}}", str(geheim_pos + 1))
            .replace("{{geheim_constraints}}", constraints)
            .replace("{{min_len}}", str(geheim_pos + 2))
        )
    else:
        prompt = (
            _PROMPT_PLAIN
            .replace("{{n_words}}", "15")
            .replace("{{categories}}", categories)
        )
    return prompt


# ---------------------------------------------------------------------------
# Interactieve invoer
# ---------------------------------------------------------------------------

def _ask_input() -> tuple[str, str, str, int]:
    """Vraag categorieen, titel, geheimwoord en kolompositie."""
    print("\n=== Filippine Puzzel Generator ===\n")
    categories = input("Categorieen (bijv. 'scouting, zeilen, Friesland'): ").strip()
    if not categories:
        sys.exit("Geen categorieen opgegeven. Script gestopt.")

    title   = input("Puzzeltitel (optioneel, enter om over te slaan): ").strip()
    gw_raw  = input("Geheimwoord (optioneel, enter om over te slaan): ").strip()
    geheimwoord = normalize(gw_raw) if gw_raw else ""

    geheim_pos = _DEFAULT_GEHEIM_POS
    if geheimwoord:
        pos_raw = input(
            f"Positie geheimkolom (1-gebaseerd, standaard={_DEFAULT_GEHEIM_POS + 1}): "
        ).strip()
        if pos_raw.isdigit() and int(pos_raw) >= 1:
            geheim_pos = int(pos_raw) - 1

    return categories, title, geheimwoord, geheim_pos


# ---------------------------------------------------------------------------
# Bouwen: validering + lettercodering
# ---------------------------------------------------------------------------

def build_filippine(
    words_clues: list[tuple[str, str]],
    geheimwoord: str,
    geheim_pos: int,
) -> tuple[list[tuple[str, str]], dict[str, int], dict[int, str]]:
    """
    Valideer woorden (geheimkolom-constraint) en genereer de willekeurige
    lettercodering 1..N.

    Geeft (validated_words_clues, letter_to_code, code_to_letter) terug.
    """
    validated: list[tuple[str, str]] = []
    n_expected = len(geheimwoord) if geheimwoord else len(words_clues)

    for i, (word, clue) in enumerate(words_clues):
        w = normalize(word)

        if len(w) <= geheim_pos:
            print(
                f"  [!] Woord '{w}' overgeslagen: te kort voor geheimkolom "
                f"op positie {geheim_pos + 1}."
            )
            continue

        if geheimwoord and i < len(geheimwoord):
            expected = geheimwoord[i]
            if w[geheim_pos] != expected:
                print(
                    f"  [!] Woord '{w}' voldoet niet aan geheimkolom: "
                    f"positie {geheim_pos + 1} = '{w[geheim_pos]}', "
                    f"maar '{expected}' verwacht (rij {i + 1})."
                )
                continue

        validated.append((w, clue))
        if geheimwoord and len(validated) >= len(geheimwoord):
            break

    if not validated:
        sys.exit("Geen geldige woorden overgebleven na validatie.")

    # Willekeurige lettercodering
    all_letters = sorted({ch for w, _ in validated for ch in w})
    codes = list(range(1, len(all_letters) + 1))
    random.shuffle(codes)
    letter_to_code: dict[str, int] = dict(zip(all_letters, codes))
    code_to_letter: dict[int, str] = {v: k for k, v in letter_to_code.items()}

    return validated, letter_to_code, code_to_letter


# ---------------------------------------------------------------------------
# PNG-rendering
# ---------------------------------------------------------------------------

def _render_png(
    words_clues: list[tuple[str, str]],
    letter_to_code: dict[str, int],
    geheimwoord: str,
    geheim_pos: int,
    show_letters: bool,
    output_path: Path,
    title: str = "",
) -> None:
    n_rows       = len(words_clues)
    max_word_len = max(len(w) for w, _ in words_clues)

    # Scale cell size if the grid would be too wide
    cell      = CELL_SIZE
    available = MAX_WIDTH - 2 * PADDING - NUM_COL_W
    if max_word_len * cell > available:
        cell = max(12, available // max_word_len)

    # Fonts
    f_num    = font(max(6, cell // 4))
    f_let    = font(max(8, int(cell * 0.5)), bold=True)
    f_rownum = font(max(8, int(cell * 0.55)), bold=True)
    f_clue   = font(CLUE_FONT_PT)
    f_head   = font(CLUE_FONT_PT, bold=True)
    f_title  = font(16, bold=True)
    f_secret = font(SECRET_FONT_PT, bold=True)

    # Canvas measurements
    title_h   = (PADDING + 22) if title else 0
    grid_px_h = n_rows * cell
    grid_px_w = NUM_COL_W + max_word_len * cell

    # Measure clue column
    dummy_img  = Image.new("RGB", (1, 1))
    dummy_draw = ImageDraw.Draw(dummy_img)

    def text_w(text: str, f) -> int:
        bb = dummy_draw.textbbox((0, 0), text, font=f)
        return bb[2] - bb[0]

    clue_col_w = MAX_WIDTH - 2 * PADDING
    clue_lines: list[str] = ["AANWIJZINGEN"]
    for i, (_, clue) in enumerate(words_clues):
        label = f"{i + 1}. {clue}"
        wds   = label.split()
        line  = ""
        for w in wds:
            candidate = (line + " " + w).strip()
            if text_w(candidate, f_clue) <= clue_col_w:
                line = candidate
            else:
                if line:
                    clue_lines.append(line)
                line = w
        if line:
            clue_lines.append(line)

    clue_block_h = (len(clue_lines) + 1) * CLUE_LINE_H + PADDING

    # Secret word box (oplossing only)
    secret_h = (cell + CLUE_LINE_H * 2 + PADDING) if (show_letters and geheimwoord) else 0

    canvas_w = max(grid_px_w + 2 * PADDING, MAX_WIDTH)
    canvas_h = title_h + grid_px_h + 2 * PADDING + clue_block_h + secret_h + PADDING

    img  = Image.new("RGB", (canvas_w, canvas_h), "white")
    draw = ImageDraw.Draw(img)

    grid_x0 = PADDING
    grid_y0  = PADDING + title_h

    # ---- Title ---------------------------------------------------------------
    if title:
        bb = draw.textbbox((0, 0), title, font=f_title)
        tw = bb[2] - bb[0]
        draw.text(((canvas_w - tw) // 2, PADDING), title, fill="black", font=f_title)

    # ---- Grid ----------------------------------------------------------------
    for row_i, (word, _) in enumerate(words_clues):
        y = grid_y0 + row_i * cell

        # Black number cell on the left
        num_x0 = grid_x0
        num_x1 = grid_x0 + NUM_COL_W - 1
        draw.rectangle([num_x0, y, num_x1, y + cell - 1],
                       fill=BLACK_CELL, outline=BORDER_COLOR, width=1)
        num_str = str(row_i + 1)
        bb  = draw.textbbox((0, 0), num_str, font=f_rownum)
        tw2 = bb[2] - bb[0]
        th2 = bb[3] - bb[1]
        draw.text(
            (num_x0 + (NUM_COL_W - tw2) // 2, y + (cell - th2) // 2),
            num_str, fill="white", font=f_rownum,
        )

        # Letter cells
        for col_j, ch in enumerate(word):
            x    = grid_x0 + NUM_COL_W + col_j * cell
            fill = GREY_COLOR if col_j == geheim_pos else WHITE_CELL
            draw.rectangle([x, y, x + cell - 1, y + cell - 1],
                           fill=fill, outline=BORDER_COLOR, width=1)

            code_str = str(letter_to_code[ch])

            if show_letters:
                # Centered letter
                bb  = draw.textbbox((0, 0), ch, font=f_let)
                tw3 = bb[2] - bb[0]
                th3 = bb[3] - bb[1]
                draw.text(
                    (x + (cell - tw3) // 2, y + (cell - th3) // 2 + 1),
                    ch, fill="black", font=f_let,
                )
                # Small code in top-left (muted, so letter is prominent)
                draw.text((x + 2, y + 1), code_str, fill="#888888", font=f_num)
            else:
                # Only the code number
                draw.text((x + 2, y + 1), code_str, fill="black", font=f_num)

    # ---- Outer border around the entire grid (number col + letter cells) ----
    draw.rectangle(
        [grid_x0, grid_y0,
         grid_x0 + NUM_COL_W + max_word_len * cell,
         grid_y0 + n_rows * cell],
        outline=BORDER_COLOR, width=1,
    )

    # ---- Clues ---------------------------------------------------------------
    clue_y = grid_y0 + grid_px_h + PADDING
    for i, line in enumerate(clue_lines):
        f = f_head if i == 0 else f_clue
        draw.text((PADDING, clue_y), line, fill="black", font=f)
        clue_y += CLUE_LINE_H

    # ---- Secret word box (oplossing only) ------------------------------------
    if show_letters and geheimwoord:
        secret_y = clue_y + PADDING // 2
        draw.text((PADDING, secret_y), f"Geheimwoord:", fill="black", font=f_secret)
        secret_y += CLUE_LINE_H + 4
        for j, ch in enumerate(geheimwoord):
            x = PADDING + j * cell
            draw.rectangle([x, secret_y, x + cell - 1, secret_y + cell - 1],
                           fill=GREY_COLOR, outline=BORDER_COLOR, width=1)
            bb  = draw.textbbox((0, 0), ch, font=f_let)
            tw2 = bb[2] - bb[0]
            th2 = bb[3] - bb[1]
            draw.text(
                (x + (cell - tw2) // 2, secret_y + (cell - th2) // 2 + 1),
                ch, fill="black", font=f_let,
            )

    img.save(str(output_path), dpi=(150, 150))
    print(f"Opgeslagen: {output_path}")


# ---------------------------------------------------------------------------
# Testwoorden
# Geheimwoord "ZEILWEDSTRIJDEN" (15 letters), positie 2 (0-indexed = 3e letter)
# Categorie: scouting, zeilen, Friesland
# Verificatie: de 3e letter (index 2) van elk woord geeft ZEILWEDSTRIJDEN
# ---------------------------------------------------------------------------

_TEST_GEHEIMWOORD = "ZEILWEDSTRIJDEN"
_TEST_GEHEIM_POS  = 2   # 0-indexed: derde letter van elk woord

_TEST_WORDS: list[tuple[str, str]] = [
    ("BEZAAN",      "Achterste zeil aan de mast van een schip"),            # [2]=Z
    ("KOERS",       "Richting die een schip vaart"),                        # [2]=E
    ("SPINNAKER",   "Groot bolstaand zeil bij ruime wind"),                 # [2]=I
    ("VALLEN",      "Touwen om een zeil te hijsen"),                        # [2]=L
    ("ONWEER",      "Gevaarlijke weersomstandigheid op het water"),         # [2]=W
    ("SNEEK",       "Friese stad bekend om watersport"),                    # [2]=E
    ("WADDEN",      "Ondiep getijdengebied bij Friesland"),                 # [2]=D
    ("MAST",        "Verticale paal waaraan zeilen worden gehangen"),       # [2]=S
    ("OPTOCHT",     "Georganiseerde mars bij scouts"),                      # [2]=T
    ("STROOM",      "Beweging van water in een bepaalde richting"),         # [2]=R
    ("DRIEMASTER",  "Historisch zeilschip met drie masten"),                # [2]=I
    ("KAJUIT",      "Overdekte ruimte op een schip"),                       # [2]=J
    ("ONDERSTROOM", "Verborgen waterbeweging onder de oppervlakte"),        # [2]=D
    ("ROEIEN",      "Boot voortbewegen met riemen"),                        # [2]=E
    ("LANDVAST",    "Touw om een boot aan de wal vast te leggen"),          # [2]=N
]


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    test_mode = "--test" in sys.argv

    if test_mode:
        print("\n=== Filippine Puzzel Generator (TEST-MODUS) ===")
        title       = "Test Filippine"
        geheimwoord = _TEST_GEHEIMWOORD
        geheim_pos  = _TEST_GEHEIM_POS
        words_raw   = _TEST_WORDS
        print(
            f"Testmodus: {len(words_raw)} woorden geladen, "
            f"geheimwoord '{geheimwoord}' op positie {geheim_pos + 1}."
        )
    else:
        categories, title, geheimwoord, geheim_pos = _ask_input()
        prompt   = _build_prompt(categories, geheimwoord, geheim_pos)
        display_prompt(prompt)
        raw_json  = collect_response()
        words_raw = parse_response(raw_json)

    print("\nFilippine bouwen...")
    words_clues, letter_to_code, code_to_letter = build_filippine(
        words_raw, geheimwoord, geheim_pos
    )
    print(f"Geplaatst: {len(words_clues)} woorden")

    if geheimwoord:
        actual = "".join(w[geheim_pos] for w, _ in words_clues)
        print(f"Geheimwoord gecontroleerd: {actual}")

    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = Path(__file__).parent

    puzzel_path    = base / f"filippine_{ts}_puzzel.png"
    oplossing_path = base / f"filippine_{ts}_oplossing.png"

    print("\nPNG puzzel renderen...")
    _render_png(
        words_clues, letter_to_code, geheimwoord, geheim_pos,
        show_letters=False, output_path=puzzel_path, title=title,
    )

    print("PNG oplossing renderen...")
    _render_png(
        words_clues, letter_to_code, geheimwoord, geheim_pos,
        show_letters=True, output_path=oplossing_path, title=title,
    )

    print("\nKlaar.")


if __name__ == "__main__":
    main()
