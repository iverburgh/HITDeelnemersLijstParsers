"""
Kruiswoordraadel generator - interactief, zonder API-sleutel.

Gebruik:
    python generate_kruiswoord.py

Het script:
  1. Vraagt om categorieen en een optionele puzzeltitel
  2. Genereert een kant-en-klare prompt die je in ChatGPT kunt plakken
     (wordt automatisch naar het klembord gekopieerd)
  3. Wacht terwijl jij de JSON-respons van ChatGPT terugplakt
  4. Bouwt het kruiswoordrooster en genereert twee PNG-bestanden

Output (in dezelfde map als het script):
  kruiswoord_<YYYYMMDD_HHMMSS>_puzzel.png
  kruiswoord_<YYYYMMDD_HHMMSS>_oplossing.png

Breedte max 1240px (A4 staand @ 150 dpi).
"""

from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from puzzle_helpers import (
    CELL_SIZE,
    PADDING,
    collect_response,
    copy_to_clipboard,
    font as _font,
    normalize as _normalize,
    parse_response as _parse_response,
    show_prompt as _show_prompt_helper,
)

# ---------------------------------------------------------------------------
# Stap 1 - Interactieve invoer
# ---------------------------------------------------------------------------

def _ask_input() -> tuple[str, str]:
    """Vraag categorieen en optionele titel; geef (categorieen, titel) terug."""
    print("\n=== Kruiswoordraadel Generator ===\n")
    categories = input("Categorieen (bijv. 'scouting, zeilen, Friesland'): ").strip()
    if not categories:
        sys.exit("Geen categorieen opgegeven. Script gestopt.")
    title = input("Puzzeltitel (optioneel, enter om over te slaan): ").strip()
    return categories, title


# ---------------------------------------------------------------------------
# Stap 2 - Prompt bouwen, tonen en kopieren naar klembord
# ---------------------------------------------------------------------------

CHATGPT_PROMPT_TEMPLATE = """\
Je bent een puzzelmaker die Nederlandse kruiswoordraadsels maakt.
Genereer een lijst van 30 tot 40 woorden en bijbehorende omschrijvingen voor een \
kruiswoordraadsel over de volgende categorieen: {{categories}}.

Eisen:
- Alle woorden en omschrijvingen zijn in het Nederlands.
- Woorden hebben GEEN spaties of koppeltekens (samengestelde woorden aan elkaar schrijven).
- Geen woorden met bijzondere leestekens (geen e met trema, u met dakje, etc.); \
gebruik gewone ASCII-letters.
- Zorg voor een gebalanceerde mix van woordlengtes zodat het kruiswoordrooster \
min of meer vierkant uitvalt:
    * 5-8 korte woorden (3-5 letters)
    * 12-18 middelmatige woorden (6-9 letters)
    * 8-12 lange woorden (10-15 letters)
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


def _show_prompt(categories: str) -> None:
    """Toon de ChatGPT-prompt in de terminal en kopieer naar klembord."""
    _show_prompt_helper(categories, CHATGPT_PROMPT_TEMPLATE)


# ---------------------------------------------------------------------------
# Stap 3+4 - Respons inlezen en parsen (via puzzle_helpers)
# ---------------------------------------------------------------------------

_collect_response = collect_response


# ---------------------------------------------------------------------------
# Grid
# ---------------------------------------------------------------------------
GRID_SIZE = 120
HORIZONTAL = "H"
VERTICAL   = "V"


@dataclass
class Placement:
    word: str
    clue: str
    row: int
    col: int
    direction: str
    number: int = 0


@dataclass
class Grid:
    size: int = GRID_SIZE
    cells: list[list[Optional[str]]] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.cells = [[None] * self.size for _ in range(self.size)]

    def get(self, r: int, c: int) -> Optional[str]:
        if 0 <= r < self.size and 0 <= c < self.size:
            return self.cells[r][c]
        return None  # out-of-bounds treated as border

    def set(self, r: int, c: int, ch: str) -> None:
        self.cells[r][c] = ch

    def can_place(self, word: str, row: int, col: int, direction: str) -> bool:
        """Return True when 'word' can be placed without conflicts."""
        dr, dc = (0, 1) if direction == HORIZONTAL else (1, 0)
        length = len(word)

        # Cell before the word must be empty/border
        pr, pc = row - dr, col - dc
        if self.get(pr, pc) is not None:
            return False

        # Cell after the word must be empty/border
        er, ec = row + dr * length, col + dc * length
        if self.get(er, ec) is not None:
            return False

        intersects = 0
        for i, ch in enumerate(word):
            r, c = row + dr * i, col + dc * i
            if r < 0 or r >= self.size or c < 0 or c >= self.size:
                return False
            existing = self.cells[r][c]
            if existing is not None:
                if existing != ch:
                    return False  # letter conflict
                intersects += 1
            else:
                # The perpendicular neighbours must be empty (no parallel words touching)
                if direction == HORIZONTAL:
                    if self.get(r - 1, c) is not None or self.get(r + 1, c) is not None:
                        return False
                else:
                    if self.get(r, c - 1) is not None or self.get(r, c + 1) is not None:
                        return False

        return intersects > 0 or not any(
            self.cells[r2][c2] is not None
            for r2 in range(self.size) for c2 in range(self.size)
        )

    def place(self, word: str, row: int, col: int, direction: str) -> None:
        dr, dc = (0, 1) if direction == HORIZONTAL else (1, 0)
        for i, ch in enumerate(word):
            self.cells[row + dr * i][col + dc * i] = ch

    def count_intersections(self, word: str, row: int, col: int, direction: str) -> int:
        dr, dc = (0, 1) if direction == HORIZONTAL else (1, 0)
        return sum(
            1 for i, ch in enumerate(word)
            if self.cells[row + dr * i][col + dc * i] == ch
        )

    def bounds(self) -> tuple[int, int, int, int]:
        """Return (min_row, min_col, max_row, max_col) of occupied cells."""
        rows = [r for r in range(self.size) for c in range(self.size) if self.cells[r][c]]
        cols = [c for r in range(self.size) for c in range(self.size) if self.cells[r][c]]
        return min(rows), min(cols), max(rows), max(cols)


# ---------------------------------------------------------------------------
# Placement algorithm
# ---------------------------------------------------------------------------

# Ratio constraint applied DURING placement to prevent diagonal drift.
# Only enforced once the grid is large enough to matter.
MAX_RATIO_DURING_PLACEMENT = 1.8
MIN_SPAN_FOR_RATIO_CHECK   = 12


def build_crossword(words_clues: list[tuple[str, str]]) -> tuple[Grid, list[Placement]]:
    grid = Grid()
    placements: list[Placement] = []
    skipped: list[str] = []

    first_word, first_clue = words_clues[0]
    start_row = GRID_SIZE // 2
    start_col = (GRID_SIZE - len(first_word)) // 2
    grid.place(first_word, start_row, start_col, HORIZONTAL)
    placements.append(Placement(first_word, first_clue, start_row, start_col, HORIZONTAL))

    for word, clue in words_clues[1:]:
        best: Optional[tuple[float, int, int, str]] = None  # (score, row, col, direction)

        b        = grid.bounds()          # current bounding box
        center_r = (b[0] + b[2]) / 2
        center_c = (b[1] + b[3]) / 2
        span     = max(b[2] - b[0] + 1, b[3] - b[1] + 1)

        for r in range(GRID_SIZE):
            for c in range(GRID_SIZE):
                cell = grid.cells[r][c]
                if cell is None:
                    continue
                for i, ch in enumerate(word):
                    if ch != cell:
                        continue
                    for direction in (HORIZONTAL, VERTICAL):
                        dr, dc = (0, 1) if direction == HORIZONTAL else (1, 0)
                        row0  = r - dr * i
                        col0  = c - dc * i
                        end_r = row0 + dr * (len(word) - 1)
                        end_c = col0 + dc * (len(word) - 1)
                        if row0 < 1 or col0 < 1 or end_r >= GRID_SIZE - 1 or end_c >= GRID_SIZE - 1:
                            continue

                        # Cheap bounding-box ratio check BEFORE the expensive can_place.
                        # Rejects candidates that would push the grid too far off-square.
                        if span >= MIN_SPAN_FOR_RATIO_CHECK:
                            new_h = max(b[2], end_r) - min(b[0], row0) + 1
                            new_w = max(b[3], end_c) - min(b[1], col0) + 1
                            new_ratio = new_h / new_w if new_h >= new_w else new_w / new_h
                            if new_ratio > MAX_RATIO_DURING_PLACEMENT:
                                continue

                        if grid.can_place(word, row0, col0, direction):
                            intersections = grid.count_intersections(word, row0, col0, direction)

                            # Proximity bonus: prefer placements whose midpoint
                            # stays close to the current bounding-box centroid.
                            mid_r    = (row0 + end_r) / 2
                            mid_c    = (col0 + end_c) / 2
                            dist     = abs(mid_r - center_r) + abs(mid_c - center_c)
                            proximity = -dist / max(span, 1)

                            score = intersections * 10.0 + proximity
                            if best is None or score > best[0]:
                                best = (score, row0, col0, direction)

        if best is None:
            skipped.append(word)
            print(f"  [SKIP] '{word}' kon niet worden geplaatst.", file=sys.stderr)
        else:
            _, row0, col0, direction = best
            grid.place(word, row0, col0, direction)
            placements.append(Placement(word, clue, row0, col0, direction))

    print(f"\nGeplaatst: {len(placements)}/{len(words_clues)}  |  Overgeslagen: {len(skipped)}")
    return grid, placements


# ---------------------------------------------------------------------------
# Ratio validation & correction
# ---------------------------------------------------------------------------

# Post-placement ratio target (tighter than during-placement).
MAX_RATIO = 1.5


def _compute_ratio(grid: Grid) -> float:
    """Return hoogte/breedte (or vice versa) so the result is always >= 1."""
    min_r, min_c, max_r, max_c = grid.bounds()
    h = max_r - min_r + 1
    w = max_c - min_c + 1
    return h / w if h >= w else w / h


def _dominant_direction(grid: Grid) -> str:
    min_r, min_c, max_r, max_c = grid.bounds()
    h = max_r - min_r + 1
    w = max_c - min_c + 1
    return HORIZONTAL if w > h else VERTICAL


def build_crossword_balanced(
    words_clues: list[tuple[str, str]],
) -> tuple[Grid, list[Placement]]:
    """Build crossword and iteratively trim words that cause a bad ratio."""
    grid, placements = build_crossword(words_clues)

    for attempt in range(6):
        ratio = _compute_ratio(grid)
        if ratio <= MAX_RATIO:
            break
        dominant = _dominant_direction(grid)
        print(f"  [Ratio {ratio:.2f}] Te langwerpig – poging {attempt + 1} bijsturen...")

        # Remove 1/3 of the longest words in the dominant direction per iteration.
        long_in_dominant = sorted(
            [p for p in placements if p.direction == dominant],
            key=lambda p: -len(p.word),
        )
        n_remove  = max(1, len(long_in_dominant) // 3)
        remove_ids = {id(p) for p in long_in_dominant[:n_remove]}
        remaining  = [(p.word, p.clue) for p in placements if id(p) not in remove_ids]

        if not remaining:
            break
        grid, placements = build_crossword(remaining)

    ratio = _compute_ratio(grid)
    label = "OK" if ratio <= MAX_RATIO else "beste poging"
    print(f"  [Ratio {ratio:.2f}] Rooster gereed ({label}).")
    return grid, placements


# ---------------------------------------------------------------------------
# Numbering
# ---------------------------------------------------------------------------

def assign_numbers(grid: Grid, placements: list[Placement]) -> tuple[
    dict[tuple[int, int], int],
    list[tuple[int, str, str]],   # across clues
    list[tuple[int, str, str]],   # down clues
]:
    min_r, min_c, max_r, max_c = grid.bounds()

    # Build lookup: (row, col) -> list of placements starting there
    start_map: dict[tuple[int, int], list[Placement]] = {}
    for p in placements:
        key = (p.row, p.col)
        start_map.setdefault(key, []).append(p)

    number_map: dict[tuple[int, int], int] = {}
    across: list[tuple[int, str, str]] = []
    down:   list[tuple[int, str, str]] = []
    counter = 1

    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            if grid.cells[r][c] is None:
                continue
            starts = start_map.get((r, c), [])
            if not starts:
                continue
            num = counter
            number_map[(r, c)] = num
            counter += 1
            for p in starts:
                p.number = num
                if p.direction == HORIZONTAL:
                    across.append((num, p.word, p.clue))
                else:
                    down.append((num, p.word, p.clue))

    return number_map, sorted(across), sorted(down)


# ---------------------------------------------------------------------------
# PNG rendering
# ---------------------------------------------------------------------------

from PIL import Image, ImageDraw, ImageFont  # noqa: E402

# CELL_SIZE and PADDING are imported from puzzle_helpers
MAX_WIDTH    = 1240        # A4 staand @ 150 dpi
CLUE_FONT_PT = 10
NUM_FONT_PT  = 7
LET_FONT_PT  = 14
LINE_H       = 15          # regelafstand aanwijzingen
COL_GAP      = 20          # tussenruimte kolommen aanwijzingen


# _font is imported from puzzle_helpers as _font


def _render_png(
    grid: Grid,
    placements: list[Placement],
    number_map: dict[tuple[int, int], int],
    across: list[tuple[int, str, str]],
    down:   list[tuple[int, str, str]],
    show_letters: bool,
    output_path: Path,
    title: str = "",
) -> None:
    min_r, min_c, max_r, max_c = grid.bounds()
    grid_rows = max_r - min_r + 1
    grid_cols = max_c - min_c + 1

    # Scale cell size if grid is too wide
    cell = CELL_SIZE
    grid_px_w = grid_cols * cell
    available = MAX_WIDTH - 2 * PADDING
    if grid_px_w > available:
        cell = max(12, available // grid_cols)
        grid_px_w = grid_cols * cell

    grid_px_h = grid_rows * cell

    # Fonts
    f_num   = _font(max(6, cell // 4))
    f_let   = _font(max(8, int(cell * 0.5)), bold=True)
    f_clue  = _font(CLUE_FONT_PT)
    f_head  = _font(CLUE_FONT_PT, bold=True)
    f_title = _font(16, bold=True)

    # ---- Build clue columns ------------------------------------------------
    # Measure max clue line width for layout
    dummy_img = Image.new("RGB", (1, 1))
    dummy_draw = ImageDraw.Draw(dummy_img)

    def text_w(text: str, font: ImageFont.FreeTypeFont) -> int:
        bb = dummy_draw.textbbox((0, 0), text, font=font)
        return bb[2] - bb[0]

    def clue_lines(entries: list[tuple[int, str, str]], max_col_w: int) -> list[str]:
        lines = []
        for num, _word, clue in entries:
            label = f"{num}. {clue}"
            # Simple word-wrap
            words = label.split()
            line = ""
            for w in words:
                candidate = (line + " " + w).strip()
                if text_w(candidate, f_clue) <= max_col_w:
                    line = candidate
                else:
                    if line:
                        lines.append(line)
                    line = w
            if line:
                lines.append(line)
        return lines

    col_w = (MAX_WIDTH - 2 * PADDING - COL_GAP) // 2

    across_lines_raw: list[str] = []
    down_lines_raw:   list[str] = []

    # Build with header
    across_lines_raw.append("HORIZONTAAL")
    for num, _word, clue in across:
        across_lines_raw.extend(clue_lines([(num, _word, clue)], col_w))

    down_lines_raw.append("VERTICAAL")
    for num, _word, clue in down:
        down_lines_raw.extend(clue_lines([(num, _word, clue)], col_w))

    n_clue_rows = max(len(across_lines_raw), len(down_lines_raw))
    clue_block_h = (n_clue_rows + 1) * LINE_H + PADDING

    title_h = (PADDING + 22) if title else 0

    # ---- Canvas size -------------------------------------------------------
    canvas_w = max(grid_px_w + 2 * PADDING, MAX_WIDTH)
    canvas_h = title_h + grid_px_h + 2 * PADDING + clue_block_h + PADDING

    img  = Image.new("RGB", (canvas_w, canvas_h), "white")
    draw = ImageDraw.Draw(img)

    grid_x0 = PADDING
    grid_y0 = PADDING + title_h

    # ---- Draw title --------------------------------------------------------
    if title:
        bb = draw.textbbox((0, 0), title, font=f_title)
        tw = bb[2] - bb[0]
        draw.text(((canvas_w - tw) // 2, PADDING), title, fill="black", font=f_title)

    # ---- Draw grid ---------------------------------------------------------
    occupied: set[tuple[int, int]] = set()
    for p in placements:
        dr, dc = (0, 1) if p.direction == HORIZONTAL else (1, 0)
        for i in range(len(p.word)):
            occupied.add((p.row + dr * i, p.col + dc * i))

    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            x = grid_x0 + (c - min_c) * cell
            y = grid_y0 + (r - min_r) * cell
            if (r, c) in occupied:
                # White cell with thin black border
                draw.rectangle([x, y, x + cell - 1, y + cell - 1],
                                fill="white", outline="#222222", width=1)
                # Number in top-left, black
                if (r, c) in number_map:
                    draw.text((x + 2, y + 1), str(number_map[(r, c)]),
                               fill="black", font=f_num)
                # Letter (oplossing only)
                if show_letters and grid.cells[r][c]:
                    ch = grid.cells[r][c]
                    bb = draw.textbbox((0, 0), ch, font=f_let)
                    tw, th = bb[2] - bb[0], bb[3] - bb[1]
                    draw.text((x + (cell - tw) // 2, y + (cell - th) // 2 + 1),
                               ch, fill="black", font=f_let)
            # Non-occupied cells: leave white (no fill), matching reference style

    # ---- Thin outer border around the entire grid bounding box -------------
    border_x0 = grid_x0
    border_y0 = grid_y0
    border_x1 = grid_x0 + grid_cols * cell
    border_y1 = grid_y0 + grid_rows * cell
    draw.rectangle([border_x0, border_y0, border_x1, border_y1],
                   outline="#222222", width=1)

    # ---- Draw clues --------------------------------------------------------
    clue_y = grid_y0 + grid_px_h + PADDING

    def draw_col(lines: list[str], x_start: int) -> None:
        y = clue_y
        for i, line in enumerate(lines):
            font = f_head if i == 0 else f_clue
            draw.text((x_start, y), line, fill="black", font=font)
            y += LINE_H

    draw_col(across_lines_raw, PADDING)
    draw_col(down_lines_raw,   PADDING + col_w + COL_GAP)

    img.save(str(output_path), dpi=(150, 150))
    print(f"Opgeslagen: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Test-woordenlijst (--test vlag; geen ChatGPT-interactie nodig)
# ---------------------------------------------------------------------------
_TEST_WORDS: list[tuple[str, str]] = [
    ("JAKOBSLADDER",  "Touwladder bij scoutingconstructies"),
    ("SNEEKERMEER",   "Groot binnenwater in Friesland"),
    ("TJEUKEMEER",    "Groot Fries meer"),
    ("WATERKAART",    "Soort kaart voor navigatie op water"),
    ("WADDENZEE",     "Water tussen eilanden in het noorden"),
    ("TREKTOCHT",     "Activiteit waarbij je lange afstanden loopt"),
    ("OVERSTAG",      "Draai van de boot door de wind"),
    ("NAVIGEREN",     "Richting bepalen met instrument"),
    ("VAARTOCHT",     "Lange tocht per boot"),
    ("ZEILBOOT",      "Boot zonder motor"),
    ("BEAUFORT",      "Weerschaal voor wind"),
    ("SJORREN",       "Samenwerkingstechniek met touwen en palen"),
    ("SJORHOUT",      "Gereedschap bij houtconstructies"),
    ("LANDVAST",      "Touw om een boot vast te leggen"),
    ("OPTIMIST",      "Klein zeilbootje voor een persoon"),
    ("SCHIPPER",      "Iemand die een boot bestuurt"),
    ("SPELTAK",       "Groep jongeren binnen scouting"),
    ("DROPPING",      "Nachtactiviteit bij scouting"),
    ("REGATTA",       "Met meerdere boten racen"),
    ("CONTOUR",       "Ronde lijn op een kaart"),
    ("SKUTSJE",       "Boot speciaal voor Friese wedstrijden"),
    ("SHELTER",       "Tijdelijk onderkomen in het bos"),
    ("LEIDING",       "Leiding bij scouting"),
    ("AFMEREN",       "Boot vastleggen aan kade"),
    ("KNOOP",         "Punt waar twee touwen samenkomen"),
    ("VAART",         "Smal water tussen land"),
    ("TONDEL",        "Iets waarmee je vuur maakt"),
    ("SCHOT",         "Touw om zeil te bedienen"),
    ("SNEEK",         "Kleine Friese stad met waterpoort"),
    ("GIJPEN",        "Draai met de wind mee"),
    ("BANK",          "Ondiep stuk water"),
    ("TARP",          "Bescherming tegen regen op kamp"),
    ("PAAL",          "Houten paal gebruikt bij pionieren"),
    ("GROU",          "Bekende Friese watersportplaats"),
    ("KUIP",          "Deel van boot waar je zit"),
    ("LOEF",          "Windrichting waar de wind vandaan komt"),
    ("LIJ",           "Windrichting waar de wind naartoe gaat"),
]


def main() -> None:
    test_mode = "--test" in sys.argv

    if test_mode:
        print("\n=== Kruiswoordraadel Generator (TEST-MODUS) ===")
        title       = "Test Puzzel"
        words_clues = sorted(_TEST_WORDS, key=lambda x: -len(x[0]))
        print(f"Testmodus: {len(words_clues)} woorden geladen.")
    else:
        categories, title = _ask_input()
        _show_prompt(categories)
        raw_json    = _collect_response()
        words_clues = _parse_response(raw_json)

    print("\nKruiswoordrooster bouwen en balanceren...")
    grid, placements = build_crossword_balanced(words_clues)
    number_map, across, down = assign_numbers(grid, placements)

    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = Path(__file__).parent

    puzzel_path    = base / f"kruiswoord_{ts}_puzzel.png"
    oplossing_path = base / f"kruiswoord_{ts}_oplossing.png"

    print("\nPNG puzzel renderen…")
    _render_png(grid, placements, number_map, across, down,
                show_letters=False, output_path=puzzel_path, title=title)

    print("PNG oplossing renderen…")
    _render_png(grid, placements, number_map, across, down,
                show_letters=True, output_path=oplossing_path, title=title)

    print("\nKlaar.")


if __name__ == "__main__":
    main()
