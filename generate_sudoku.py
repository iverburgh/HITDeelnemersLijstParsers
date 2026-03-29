"""
Sudoku-puzzel generator - moeilijkheidsgraad 5/10.

Gebruik:
    python generate_sudoku.py          # willekeurige sudoku
    python generate_sudoku.py --test   # vaste seed (reproduceerbaar resultaat)

Hoe een Sudoku werkt:
  - 9x9 vakjes, gegroepeerd als negen 3x3 blokken.
  - In elke rij, kolom en elk 3x3 blok komen de cijfers 1-9 precies eenmaal voor.
  - Een aantal vakjes is vooraf ingevuld; de rest moet worden aangevuld.

Moeilijkheidsgraad 5/10:
  - ~30 gegeven cijfers (hints) van de 81 vakjes.
  - Vereist meer geavanceerde oplosstrategieen dan een eenvoudige sudoku.
  - De puzzel heeft gegarandeerd een eenduidige oplossing.

Output (in dezelfde map als het script):
  sudoku_<YYYYMMDD_HHMMSS>_puzzel.png
  sudoku_<YYYYMMDD_HHMMSS>_oplossing.png
"""

from __future__ import annotations

import random
import sys
from copy import deepcopy
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw

from puzzle_helpers import PADDING, font

# ---------------------------------------------------------------------------
# Constanten
# ---------------------------------------------------------------------------

GRID_N       = 9           # aantal rijen/kolommen
BOX_N        = 3           # blokgrootte (3x3)
CELL_PX      = 56          # pixels per cel
THIN_W       = 1           # randbreedte tussen cellen binnen een blok
THICK_W      = 3           # randbreedte tussen blokken en buitenrand
BORDER_COLOR = "#222222"

DIFFICULTY   = 5           # moeilijkheidsgraad op schaal 1-10
TARGET_CLUES = 30          # gewenst aantal gegeven cijfers

COLOR_GIVEN  = "black"     # vooraf ingevulde cijfers
COLOR_SOLVED = "#1a5fa8"   # aangevuld in de oplossing-PNG

# ---------------------------------------------------------------------------
# Validering
# ---------------------------------------------------------------------------

def _is_valid(grid: list[list[int]], r: int, c: int, n: int) -> bool:
    """Controleer of getal n op positie (r, c) geldig is (rij, kolom, blok)."""
    if n in grid[r]:
        return False
    if any(grid[i][c] == n for i in range(GRID_N)):
        return False
    br, bc = (r // BOX_N) * BOX_N, (c // BOX_N) * BOX_N
    for i in range(br, br + BOX_N):
        for j in range(bc, bc + BOX_N):
            if grid[i][j] == n:
                return False
    return True


# ---------------------------------------------------------------------------
# Oplossing genereren via willekeurige backtracking
# ---------------------------------------------------------------------------

def _fill_grid(grid: list[list[int]]) -> bool:
    for r in range(GRID_N):
        for c in range(GRID_N):
            if grid[r][c] == 0:
                nums = list(range(1, GRID_N + 1))
                random.shuffle(nums)
                for n in nums:
                    if _is_valid(grid, r, c, n):
                        grid[r][c] = n
                        if _fill_grid(grid):
                            return True
                        grid[r][c] = 0
                return False
    return True


def generate_solution() -> list[list[int]]:
    """Genereer een volledige, geldige sudoku-oplossing."""
    grid = [[0] * GRID_N for _ in range(GRID_N)]
    _fill_grid(grid)
    return grid


# ---------------------------------------------------------------------------
# Uniekheidscheck met MRV-heuristiek (snelheid)
# ---------------------------------------------------------------------------

def _count_solutions(grid: list[list[int]], limit: int = 2) -> int:
    """
    Tel het aantal geldige oplossingen; stopt zodra limit bereikt is.
    MRV-heuristiek: vul altijd de cel met de minste opties als eerste in.
    """
    count = [0]

    def _solve(g: list[list[int]]) -> None:
        if count[0] >= limit:
            return

        # Zoek lege cel met minste geldige opties (MRV)
        best_r = best_c = -1
        best_opts: list[int] = []
        best_len = GRID_N + 1
        found_empty = False

        for r in range(GRID_N):
            for c in range(GRID_N):
                if g[r][c] == 0:
                    found_empty = True
                    opts = [n for n in range(1, GRID_N + 1) if _is_valid(g, r, c, n)]
                    if not opts:
                        return  # doodlopende tak
                    if len(opts) < best_len:
                        best_len, best_r, best_c, best_opts = len(opts), r, c, opts

        if not found_empty:
            count[0] += 1
            return

        for n in best_opts:
            g[best_r][best_c] = n
            _solve(g)
            g[best_r][best_c] = 0

    _solve([row[:] for row in grid])
    return count[0]


# ---------------------------------------------------------------------------
# Puzzel maken: verwijder cellen tot target_clues remain
# ---------------------------------------------------------------------------

def make_puzzle(
    solution: list[list[int]],
    target_clues: int = TARGET_CLUES,
) -> list[list[int]]:
    """
    Verwijder willekeurige cellen uit de oplossing, steeds gecontroleerd op
    eenduidigheid. Stopt als het gewenste aantal hints bereikt is.
    """
    puzzle = deepcopy(solution)
    cells = [(r, c) for r in range(GRID_N) for c in range(GRID_N)]
    random.shuffle(cells)

    to_remove = GRID_N * GRID_N - target_clues
    removed = 0

    for r, c in cells:
        if removed >= to_remove:
            break
        val = puzzle[r][c]
        puzzle[r][c] = 0
        if _count_solutions(puzzle) == 1:
            removed += 1
        else:
            puzzle[r][c] = val  # terugzetten: zou uniekheid breken

    actual = sum(puzzle[r][c] != 0 for r in range(GRID_N) for c in range(GRID_N))
    print(f"  Hints: {actual}/81  (doel: {target_clues})")
    return puzzle


# ---------------------------------------------------------------------------
# Rasterlay-out: celposities berekenen
# ---------------------------------------------------------------------------

def _cell_positions() -> tuple[list[int], list[int], int, int]:
    """
    Geeft (x_starts, y_starts, total_width, total_height) terug.
    Verwerkt dunne randen binnen blokken en dikke randen tussen blokken.
    """
    def _offsets() -> list[int]:
        positions: list[int] = []
        pos = THICK_W  # begin na buitenrand
        for i in range(GRID_N):
            positions.append(pos)
            if i < GRID_N - 1:
                gap = THICK_W if (i + 1) % BOX_N == 0 else THIN_W
                pos += CELL_PX + gap
        return positions

    xs = _offsets()
    ys = _offsets()
    total_w = xs[-1] + CELL_PX + THICK_W
    total_h = ys[-1] + CELL_PX + THICK_W
    return xs, ys, total_w, total_h


# ---------------------------------------------------------------------------
# PNG-rendering
# ---------------------------------------------------------------------------

def _render_png(
    puzzle: list[list[int]],
    solution: list[list[int]],
    show_solution: bool,
    output_path: Path,
    title: str = "",
) -> None:
    xs, ys, grid_w, grid_h = _cell_positions()

    f_num   = font(int(CELL_PX * 0.55), bold=True)
    f_title = font(20, bold=True)
    f_info  = font(11)

    title_h = (PADDING + 26) if title else 0
    info_h  = PADDING // 2 + 18 + (18 if show_solution else 0) + PADDING

    canvas_w = grid_w + 2 * PADDING
    canvas_h = title_h + grid_h + 2 * PADDING + info_h

    img  = Image.new("RGB", (canvas_w, canvas_h), "white")
    draw = ImageDraw.Draw(img)

    grid_x0 = PADDING
    grid_y0  = PADDING + title_h

    # ---- Titel ---------------------------------------------------------------
    if title:
        bb = draw.textbbox((0, 0), title, font=f_title)
        tw = bb[2] - bb[0]
        draw.text(((canvas_w - tw) // 2, PADDING), title, fill="black", font=f_title)

    # ---- Rasterachtergrond: de randkleur vormt automatisch alle lijnen -------
    draw.rectangle(
        [grid_x0, grid_y0, grid_x0 + grid_w, grid_y0 + grid_h],
        fill=BORDER_COLOR,
    )

    # ---- Cellen (wit) + cijfers ----------------------------------------------
    for r in range(GRID_N):
        for c in range(GRID_N):
            x = grid_x0 + xs[c]
            y = grid_y0 + ys[r]

            draw.rectangle([x, y, x + CELL_PX - 1, y + CELL_PX - 1], fill="white")

            given = puzzle[r][c] != 0
            value = puzzle[r][c] if not show_solution else solution[r][c]
            if value == 0:
                continue

            color = COLOR_GIVEN if (given or not show_solution) else COLOR_SOLVED
            s  = str(value)
            bb = draw.textbbox((0, 0), s, font=f_num)
            tw = bb[2] - bb[0]
            th = bb[3] - bb[1]
            draw.text(
                (x + (CELL_PX - tw) // 2, y + (CELL_PX - th) // 2 - 1),
                s, fill=color, font=f_num,
            )

    # ---- Info onder raster ---------------------------------------------------
    clues  = sum(puzzle[r][c] != 0 for r in range(GRID_N) for c in range(GRID_N))
    info_y = grid_y0 + grid_h + PADDING // 2
    info_line = f"Moeilijkheidsgraad: {DIFFICULTY}/10   |   Hints: {clues}/81"
    bb = draw.textbbox((0, 0), info_line, font=f_info)
    draw.text(
        ((canvas_w - (bb[2] - bb[0])) // 2, info_y),
        info_line, fill="#555555", font=f_info,
    )

    # ---- Legenda (alleen oplossing-PNG) --------------------------------------
    if show_solution:
        leg_y      = info_y + 18
        leg_given  = "● Gegeven cijfer"
        leg_spacer = "     "
        leg_solved = "● Aangevuld"
        bb_all = draw.textbbox((0, 0), leg_given + leg_spacer + leg_solved, font=f_info)
        bb_gs  = draw.textbbox((0, 0), leg_given + leg_spacer, font=f_info)
        x_leg  = (canvas_w - (bb_all[2] - bb_all[0])) // 2
        draw.text((x_leg, leg_y), leg_given, fill=COLOR_GIVEN, font=f_info)
        draw.text(
            (x_leg + (bb_gs[2] - bb_gs[0]), leg_y),
            leg_solved, fill=COLOR_SOLVED, font=f_info,
        )

    img.save(str(output_path), dpi=(150, 150))
    print(f"Opgeslagen: {output_path}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    test_mode = "--test" in sys.argv

    if test_mode:
        print("\n=== Sudoku Generator (TEST-MODUS, vaste seed=42) ===")
        random.seed(42)
        title = "Test Sudoku"
    else:
        print("\n=== Sudoku Generator ===")
        title = input("Puzzeltitel (optioneel, enter om over te slaan): ").strip()

    print("\nOplossing genereren...")
    solution = generate_solution()

    print(f"Puzzel maken (doel: {TARGET_CLUES} hints, moeilijkheidsgraad {DIFFICULTY}/10)...")
    puzzle = make_puzzle(solution)

    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = Path(__file__).parent

    puzzel_path    = base / f"sudoku_{ts}_puzzel.png"
    oplossing_path = base / f"sudoku_{ts}_oplossing.png"

    print("\nPNG puzzel renderen...")
    _render_png(puzzle, solution, show_solution=False, output_path=puzzel_path, title=title)

    print("PNG oplossing renderen...")
    _render_png(puzzle, solution, show_solution=True, output_path=oplossing_path, title=title)

    print("\nKlaar.")


if __name__ == "__main__":
    main()
