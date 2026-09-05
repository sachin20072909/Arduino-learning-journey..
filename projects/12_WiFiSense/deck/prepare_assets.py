"""
prepare_assets.py
=================

Turns the raw AI-generated artwork in ``assets/gen`` into the exact crops,
tones and icon cut-outs that ``build_deck.py`` expects in ``assets/processed``.

The generated art is deliberately moody and very dark, so every asset is
*lifted* with a gamma curve until it reads clearly on the deck's near-black
background, and then re-scrimmed only where headline text sits on top of it.

    python3 deck/prepare_assets.py
    python3 deck/prepare_assets.py --stats      # print per-asset tone report
"""

from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageEnhance, ImageFilter

ROOT = Path(__file__).resolve().parents[1]
GEN = ROOT / "assets" / "gen"
OUT = ROOT / "assets" / "processed"


# ---------------------------------------------------------------- helpers


def load(name: str) -> Image.Image:
    return Image.open(GEN / name).convert("RGB")


def cover(img: Image.Image, aspect: float, max_w: int = 1600) -> Image.Image:
    """Crop to ``aspect`` (w / h) from the centre, then downscale."""
    w, h = img.size
    if w / aspect <= h:                        # crop height
        new_h = w / aspect
        top = (h - new_h) / 2
        img = img.crop((0, top, w, top + new_h))
    else:                                      # crop width
        new_w = h * aspect
        left = (w - new_w) / 2
        img = img.crop((left, 0, left + new_w, h))
    if img.width > max_w:
        img = img.resize((max_w, round(max_w / aspect)), Image.LANCZOS)
    return img


def gamma(img: Image.Image, g: float) -> Image.Image:
    """Gamma lift (g > 1 brightens midtones, g < 1 darkens)."""
    if abs(g - 1.0) < 1e-3:
        return img
    lut = [min(255, round(255 * ((i / 255.0) ** (1.0 / g)))) for i in range(256)]
    return img.point(lut * 3)


def tone(img: Image.Image, g: float = 1.0, contrast: float = 1.0,
         saturate: float = 1.0, blur: float = 0.0) -> Image.Image:
    img = gamma(img, g)
    if abs(contrast - 1.0) > 1e-3:
        img = ImageEnhance.Contrast(img).enhance(contrast)
    if abs(saturate - 1.0) > 1e-3:
        img = ImageEnhance.Color(img).enhance(saturate)
    if blur:
        img = img.filter(ImageFilter.GaussianBlur(blur))
    return img


def horizontal_fade(img: Image.Image, start_alpha: float, end_alpha: float,
                    tint=(4, 6, 12)) -> Image.Image:
    """Dark scrim, strongest on the left (where headline text sits)."""
    img = img.convert("RGBA")
    w, h = img.size
    ramp = np.linspace(start_alpha, end_alpha, w, dtype=np.float32) ** 1.35
    scrim = np.zeros((h, w, 4), dtype=np.uint8)
    scrim[..., 0], scrim[..., 1], scrim[..., 2] = tint
    scrim[..., 3] = (np.tile(ramp, (h, 1)) * 255).astype(np.uint8)
    return Image.alpha_composite(img, Image.fromarray(scrim, "RGBA"))


def vertical_fade(img: Image.Image, top_alpha: float, bottom_alpha: float,
                  tint=(4, 6, 12)) -> Image.Image:
    img = img.convert("RGBA")
    w, h = img.size
    ramp = np.linspace(top_alpha, bottom_alpha, h, dtype=np.float32) ** 1.2
    scrim = np.zeros((h, w, 4), dtype=np.uint8)
    scrim[..., 0], scrim[..., 1], scrim[..., 2] = tint
    scrim[..., 3] = (np.tile(ramp, (w, 1)).T * 255).astype(np.uint8)
    return Image.alpha_composite(img, Image.fromarray(scrim, "RGBA"))


def round_corners(img: Image.Image, radius_px: int) -> Image.Image:
    img = img.convert("RGBA")
    scale = 4
    mask = Image.new("L", (img.width * scale, img.height * scale), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, mask.width - 1, mask.height - 1], radius=radius_px * scale,
        fill=255)
    mask = mask.resize(img.size, Image.LANCZOS)
    out = Image.new("RGBA", img.size, (0, 0, 0, 0))
    out.paste(img, (0, 0), mask)
    return out


def luminance_to_alpha(img: Image.Image, lo: int = 12, hi: int = 78) -> Image.Image:
    """Cut an icon out of its black background: dark → transparent."""
    arr = np.asarray(img.convert("RGBA")).astype(np.float32)
    lum = arr[..., :3].mean(axis=2)
    alpha = np.clip((lum - lo) / (hi - lo), 0.0, 1.0) * 255
    arr[..., 3] = alpha
    return Image.fromarray(arr.astype(np.uint8), "RGBA")


def trim_to_content(img: Image.Image, threshold: int = 26) -> Image.Image:
    mask = np.asarray(img.convert("L"), dtype=np.int16) > threshold
    if not mask.any():
        return img
    rows = np.where(mask.any(axis=1))[0]
    cols = np.where(mask.any(axis=0))[0]
    return img.crop((cols[0], rows[0], cols[-1] + 1, rows[-1] + 1))


def pad_square(img: Image.Image, pad_ratio: float = 0.10,
               canvas: int = 420) -> Image.Image:
    w, h = img.size
    side = int(max(w, h) * (1 + 2 * pad_ratio))
    base = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    base.paste(img, ((side - w) // 2, (side - h) // 2))
    return base.resize((canvas, canvas), Image.LANCZOS)


def save(img: Image.Image, name: str) -> Path:
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / name
    img.save(path, "PNG", optimize=True)
    return path


# ---------------------------------------------------------------- pipeline


def main(stats: bool = False) -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    report = []

    def emit(img: Image.Image, name: str):
        save(img, name)
        if stats:
            arr = np.asarray(img.convert("RGBA")).astype(np.float32)
            lum = arr[..., :3].mean(axis=2)
            a = arr[..., 3] / 255.0
            vis = lum[a > 0.5]
            report.append(
                f"  {name:22s} {str(img.size):12s} "
                f"mean={lum.mean():5.1f} p90={np.percentile(lum, 90):5.0f} "
                f"p99={np.percentile(lum, 99):5.0f} "
                f"opaque={(a > 0.5).mean() * 100:5.1f}%")
        else:
            print(f"  -> {name}  {img.size}")

    # ---- 1. title slide backdrop -------------------------------------
    hero = cover(load("hero_title.png"), 13.333 / 7.5, max_w=1900)
    hero = tone(hero, g=1.05, contrast=1.10, saturate=1.05)
    emit(horizontal_fade(hero, 0.62, 0.14), "s01_hero.png")

    # ---- 2. closing slide backdrop ------------------------------------
    closing = cover(load("hero_title.png"), 13.333 / 7.5, max_w=1700)
    closing = tone(closing, g=1.30, contrast=1.05, saturate=0.95, blur=1.4)
    closing = vertical_fade(horizontal_fade(closing, 0.50, 0.30), 0.20, 0.62)
    emit(closing, "s10_closing.png")

    # ---- 3. architecture slide backdrop --------------------------------
    arch = cover(load("arch_bg.png"), 13.333 / 7.5, max_w=1500)
    arch = tone(arch, g=2.50, contrast=1.30, saturate=1.20, blur=0.6)
    arch = vertical_fade(horizontal_fade(arch, 0.45, 0.20), 0.10, 0.35)
    emit(arch, "s07_arch_bg.png")

    # ---- 4. problem: camera vs. privacy --------------------------------
    prob = cover(load("privacy_vs_camera.png"), 2.0, max_w=1500)
    prob = tone(prob, g=2.30, contrast=1.15, saturate=1.10)
    emit(round_corners(prob, 18), "s02_problem.png")

    # ---- 5. solution: RF band -------------------------------------------
    band = cover(load("rf_waves_human.png"), 11.83 / 2.42, max_w=1800)
    band = tone(band, g=2.10, contrast=1.10, saturate=1.08)
    emit(round_corners(vertical_fade(band, 0.0, 0.45), 16), "s03_band.png")

    # ---- 6. CSI waveform panel ------------------------------------------
    csi = cover(load("csi_waveform.png"), 5.30 / 3.62, max_w=1300)
    csi = tone(csi, g=1.85, contrast=1.15, saturate=1.10)
    emit(round_corners(vertical_fade(csi, 0.0, 0.35), 16), "s04_csi.png")

    # ---- 7. ML panel ------------------------------------------------------
    ml = cover(load("ml_neural.png"), 5.733 / 3.10, max_w=1300)
    ml = tone(ml, g=2.00, contrast=1.12, saturate=1.10)
    emit(round_corners(vertical_fade(ml, 0.0, 0.32), 16), "s06_ml.png")

    # ---- 8. object icons (false-positive slide) ---------------------------
    strip = tone(load("objects_set.png"), g=1.70, contrast=1.15, saturate=1.05)
    n = 5
    col_w = strip.width / n
    for i, name in enumerate(["person", "bag", "chair", "fan", "door"]):
        tile = strip.crop((round(i * col_w), 0,
                           round((i + 1) * col_w), strip.height))
        tile = trim_to_content(tile, threshold=20)
        tile = pad_square(tile, pad_ratio=0.12, canvas=420)
        emit(luminance_to_alpha(tile, lo=14, hi=86), f"icon_{name}.png")

    # ---- 9. use-case tiles -------------------------------------------------
    tile_aspect = 3.747 / 1.10          # image area inside each use-case card
    tiles = [
        ("privacy_vs_camera.png", "tile_privacy", 0.50, 1.00, 2.30),
        ("usecase_elderly.png", "tile_elderly", 0.00, 1.00, 2.10),
        ("usecase_building.png", "tile_building", 0.00, 1.00, 1.45),
        ("arch_bg.png", "tile_restricted", 0.00, 1.00, 2.60),
        ("hero_title.png", "tile_night", 0.00, 1.00, 1.30),
        ("rf_waves_human.png", "tile_occupancy", 0.00, 1.00, 2.00),
    ]
    for src, name, x0, x1, g in tiles:
        img = load(src)
        w, h = img.size
        if x0 > 0 or x1 < 1:
            img = img.crop((round(w * x0), 0, round(w * x1), h))
        t = cover(img, tile_aspect, max_w=1100)
        t = tone(t, g=g, contrast=1.12, saturate=1.08)
        emit(round_corners(t, 14), f"{name}.png")

    # ---- 10. fills used inside architecture boxes --------------------------
    esp = cover(load("esp32_board.png"), 4.076 / 0.876, max_w=900)
    emit(round_corners(tone(esp, g=1.35, contrast=1.10), 10), "box_esp32.png")

    rf = cover(load("rf_waves_human.png"), 4.076 / 0.876, max_w=900)
    emit(round_corners(tone(rf, g=2.10, contrast=1.10), 10), "box_rf.png")

    mlbox = cover(load("ml_neural.png"), 5.809 / 0.876, max_w=1000)
    emit(round_corners(tone(mlbox, g=2.05, contrast=1.10), 10), "box_ml.png")

    # ---- 11. demo-slide hardware strip --------------------------------------
    hw = cover(load("esp32_board.png"), 2.60 / 0.82, max_w=900)
    emit(round_corners(tone(hw, g=1.30, contrast=1.12, saturate=1.05), 10),
         "s09_hardware.png")

    if stats:
        print("tone report (0–255 luminance):")
        print("\n".join(report))
    print("\nassets ready in", OUT.relative_to(ROOT))


if __name__ == "__main__":
    main(stats="--stats" in sys.argv)
