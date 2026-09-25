# -*- coding: utf-8 -*-
"""
Вставка картинок у ексель-файл.

У вибраній директорії шукаємо:
  * ексель-файл, у назві якого тільки цифри (наприклад 12345.xlsx);
  * директорію, у назві якої є "обк" (регістр не важливий) — там лежать картинки.

Далі для кожного рядка (з START_ROW) беремо значення з колонки READ_COL,
шукаємо картинку з такою назвою в папці "обк" і вставляємо її в колонку IMG_COL,
масштабуючи під ширину колонки і підганяючи висоту рядка.

Якщо у вибраній директорії такої пари немає — перевіряються її підпапки
(перший рівень), тож можна клацнути і по батьківській папці.
"""

import argparse
import os
import re
import sys
from pathlib import Path

import xlwings as xw
from PIL import Image as PILImage

# ---------------------------------------------------------------- налаштування

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")
IMG_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".gif")
IMG_DIR_MARKER = "обк"
SHEET_NAME = None      # None — активний аркуш
READ_COL = "A"         # звідки брати назву картинки
IMG_COL = "D"          # куди вставляти картинку
START_ROW = 2

DIGITS_RE = re.compile(r"^\d+$")


# ---------------------------------------------------------------- пошук файлів

def find_excels(folder: Path) -> list[Path]:
    return sorted(
        p for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in EXCEL_EXTS
        and DIGITS_RE.match(p.stem)
    )


def find_img_dir(folder: Path) -> Path | None:
    dirs = sorted(
        p for p in folder.iterdir()
        if p.is_dir() and IMG_DIR_MARKER in p.name.lower()
    )
    if len(dirs) > 1:
        print(f"  ! кілька папок з '{IMG_DIR_MARKER}', беру першу: {dirs[0].name}")
    return dirs[0] if dirs else None


def find_jobs(folder: Path) -> list[tuple[Path, Path]]:
    """Пари (ексель, папка з картинками) у папці або, якщо там нічого, в підпапках."""
    def jobs_in(d: Path):
        img_dir = find_img_dir(d)
        if not img_dir:
            return []
        return [(x, img_dir) for x in find_excels(d)]

    jobs = jobs_in(folder)
    if jobs:
        return jobs
    for sub in sorted(p for p in folder.iterdir() if p.is_dir()):
        jobs.extend(jobs_in(sub))
    return jobs


# ---------------------------------------------------------------- вставка

def insert_images_to_excel(
    app: xw.App,
    excel_path: Path,
    images_dir: Path,
    sheet_name: str | None = SHEET_NAME,
    read_col: str = READ_COL,
    img_col: str = IMG_COL,
    start_row: int = START_ROW,
) -> tuple[int, list[str]]:
    """Повертає (кількість вставлених картинок, список ключів без картинки)."""
    wb = app.books.open(str(excel_path))
    try:
        sht = wb.sheets[sheet_name] if sheet_name else wb.sheets.active

        last_row = sht.range(
            f"{read_col}{sht.cells.last_cell.row}"
        ).end("up").row

        col_width_pts = sht.range(f"{img_col}1").width

        inserted, missing = 0, []
        for row in range(start_row, last_row + 1):
            key = sht.range(f"{read_col}{row}").value
            if not key:
                continue

            if isinstance(key, (int, float)) and float(key).is_integer():
                key = str(int(key))
            else:
                key = str(key).strip()

            img_path = None
            for ext in IMG_EXTS:
                candidate = images_dir / f"{key}{ext}"
                if candidate.exists():
                    img_path = str(candidate.resolve())
                    break

            if not img_path:
                missing.append(key)
                continue

            with PILImage.open(img_path) as pil:
                ow, oh = pil.size

            orig_w_pts = ow * 0.70
            scale = col_width_pts / orig_w_pts
            new_w_pts = col_width_pts
            new_h_pts = oh * 0.70 * scale

            cell = sht.range(f"{img_col}{row}")
            sht.pictures.add(
                img_path,
                left=cell.left,
                top=cell.top,
                width=new_w_pts,
                height=new_h_pts,
            )

            sht.range(f"{row}:{row}").row_height = new_h_pts
            inserted += 1

        wb.save()
        return inserted, missing
    finally:
        wb.close()


# ---------------------------------------------------------------- main

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("folder", nargs="?", default=os.getcwd(),
                        help="директорія для обробки (за замовчуванням — поточна)")
    args = parser.parse_args()

    folder = Path(args.folder.strip('"')).resolve()
    if not folder.is_dir():
        print(f"Директорію не знайдено: {folder}")
        return 1

    jobs = find_jobs(folder)
    if not jobs:
        print(f"У {folder} не знайдено ексель-файлу з цифровою назвою "
              f"і папки з '{IMG_DIR_MARKER}' у назві.")
        return 1

    errors = 0
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        for excel_path, img_dir in jobs:
            print(f"{excel_path.name}  <-  {img_dir.name}")
            try:
                inserted, missing = insert_images_to_excel(app, excel_path, img_dir)
            except Exception as e:
                errors += 1
                print(f"  ПОМИЛКА: {e}")
                continue
            print(f"  вставлено картинок: {inserted}")
            if missing:
                print(f"  без картинки ({len(missing)}): {', '.join(missing)}")
    finally:
        app.quit()

    print("\nГотово." if not errors else f"\nЗавершено з помилками: {errors}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
