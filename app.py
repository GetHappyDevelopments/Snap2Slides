from __future__ import annotations

import argparse
import ctypes
import os
import sys
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import cv2
import numpy as np
from PIL import Image, ImageTk
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Pt
from rapidocr_onnxruntime import RapidOCR
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


EMU_PER_INCH = 914400
SLIDE_WIDTH_IN = 13.333333
MIN_TEXT_CONFIDENCE = 0.45
BURGUNDY = RGBColor(137, 13, 64)
TEXT_BLACK = RGBColor(20, 20, 20)
APP_ROOT = Path(__file__).resolve().parent
APP_ICON_ICO = APP_ROOT / "assets" / "app_icon.ico"
APP_ICON_PNG = APP_ROOT / "assets" / "app_icon.png"
APP_SPLASH_PNG = APP_ROOT / "assets" / "splash.png"
WINDOWS_APP_ID = "Snap2Slides.Image2PptSlicer"


@dataclass(frozen=True)
class Rect:
    x: int
    y: int
    w: int
    h: int

    @property
    def x2(self) -> int:
        return self.x + self.w

    @property
    def y2(self) -> int:
        return self.y + self.h

    @property
    def area(self) -> int:
        return self.w * self.h

    def clipped(self, max_w: int, max_h: int) -> "Rect":
        x = max(0, min(self.x, max_w - 1))
        y = max(0, min(self.y, max_h - 1))
        x2 = max(x + 1, min(self.x2, max_w))
        y2 = max(y + 1, min(self.y2, max_h))
        return Rect(x, y, x2 - x, y2 - y)

    def padded(self, px: int, max_w: int, max_h: int) -> "Rect":
        return Rect(self.x - px, self.y - px, self.w + px * 2, self.h + px * 2).clipped(max_w, max_h)


@dataclass
class TextLine:
    rect: Rect
    text: str
    confidence: float
    color: RGBColor


@dataclass
class TextBlock:
    rect: Rect
    lines: list[TextLine]
    bullet_lines: list[bool]

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines)

    @property
    def confidence(self) -> float:
        return min(line.confidence for line in self.lines)

    @property
    def color(self) -> RGBColor:
        return self.lines[0].color if self.lines else TEXT_BLACK


@dataclass
class VisualBlock:
    rect: Rect
    image: Image.Image


@dataclass
class ShapeBlock:
    rect: Rect
    fill: RGBColor


@dataclass
class AnalyzedSlide:
    rect: Rect
    image: Image.Image
    texts: list[TextBlock]
    visuals: list[VisualBlock]
    shapes: list[ShapeBlock]


def _bands_from_mask(values: np.ndarray, min_width: int) -> list[tuple[int, int]]:
    bands: list[tuple[int, int]] = []
    start: int | None = None
    for idx, is_band in enumerate(values):
        if is_band and start is None:
            start = idx
        elif not is_band and start is not None:
            if idx - start >= min_width:
                bands.append((start, idx))
            start = None
    if start is not None and len(values) - start >= min_width:
        bands.append((start, len(values)))
    return bands


def _intervals_between_bands(size: int, bands: list[tuple[int, int]], min_size: int) -> list[tuple[int, int]]:
    intervals: list[tuple[int, int]] = []
    cursor = 0
    for start, end in bands:
        if start - cursor >= min_size:
            intervals.append((cursor, start))
        cursor = end
    if size - cursor >= min_size:
        intervals.append((cursor, size))
    return intervals


def _smoothed_axis_bands(gray: np.ndarray, axis: int, threshold: float, min_width: int) -> list[tuple[int, int]]:
    values = gray.mean(axis=axis)
    kernel_width = max(9, min_width * 3)
    kernel = np.ones(kernel_width) / kernel_width
    smoothed = np.convolve(values, kernel, mode="same")
    return _bands_from_mask(smoothed < threshold, min_width)


def detect_slide_rects(image: Image.Image) -> list[Rect]:
    arr = np.array(image.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    h, w = gray.shape

    min_band_w = max(6, w // 300)
    min_band_h = max(6, h // 300)
    col_bands = _smoothed_axis_bands(gray, axis=0, threshold=132, min_width=min_band_w)
    row_bands = _smoothed_axis_bands(gray, axis=1, threshold=132, min_width=min_band_h)

    if len(col_bands) < 2 or len(row_bands) < 2:
        dark = gray < 95
        vertical = dark.mean(axis=0) > 0.25
        horizontal = dark.mean(axis=1) > 0.25
        col_bands = _bands_from_mask(vertical, min_band_w)
        row_bands = _bands_from_mask(horizontal, min_band_h)

    min_slide_w = max(120, w // 8)
    min_slide_h = max(80, h // 8)
    col_intervals = _intervals_between_bands(w, col_bands, min_slide_w)
    row_intervals = _intervals_between_bands(h, row_bands, min_slide_h)

    rects = [
        Rect(x1, y1, x2 - x1, y2 - y1).padded(-2, w, h)
        for y1, y2 in row_intervals
        for x1, x2 in col_intervals
        if (x2 - x1) > min_slide_w and (y2 - y1) > min_slide_h
    ]

    if len(rects) >= 2:
        return rects

    return [Rect(0, 0, w, h)]


def _box_to_rect(points: list[list[float]], max_w: int, max_h: int) -> Rect:
    xs = [int(round(p[0])) for p in points]
    ys = [int(round(p[1])) for p in points]
    return Rect(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)).padded(2, max_w, max_h)


def _union_rect(rects: list[Rect], max_w: int, max_h: int, padding: int = 0) -> Rect:
    x1 = min(rect.x for rect in rects)
    y1 = min(rect.y for rect in rects)
    x2 = max(rect.x2 for rect in rects)
    y2 = max(rect.y2 for rect in rects)
    return Rect(x1, y1, x2 - x1, y2 - y1).padded(padding, max_w, max_h)


def _guess_text_color(arr: np.ndarray, rect: Rect) -> RGBColor:
    crop = arr[rect.y : rect.y2, rect.x : rect.x2]
    if crop.size == 0:
        return TEXT_BLACK
    darkish = crop[np.mean(crop, axis=2) < 150]
    if darkish.size == 0:
        return TEXT_BLACK
    mean = darkish.mean(axis=0)
    if mean[0] > mean[1] * 1.25 and mean[0] > mean[2] * 1.15:
        return BURGUNDY
    return TEXT_BLACK


def _intersection_area(first: Rect, second: Rect) -> int:
    x1 = max(first.x, second.x)
    y1 = max(first.y, second.y)
    x2 = min(first.x2, second.x2)
    y2 = min(first.y2, second.y2)
    if x2 <= x1 or y2 <= y1:
        return 0
    return (x2 - x1) * (y2 - y1)


def _overlap_ratio(first: Rect, second: Rect) -> float:
    smaller = min(first.area, second.area)
    if smaller <= 0:
        return 0.0
    return _intersection_area(first, second) / smaller


def _rect_fill_color(arr: np.ndarray, rect: Rect) -> RGBColor:
    crop = arr[rect.y : rect.y2, rect.x : rect.x2]
    if crop.size == 0:
        return RGBColor(255, 255, 255)
    median = np.median(crop.reshape(-1, 3), axis=0)
    return RGBColor(int(median[0]), int(median[1]), int(median[2]))


def _build_text_mask(width: int, height: int, text_rects: list[Rect], padding: int = 2) -> np.ndarray:
    mask = np.zeros((height, width), dtype=np.uint8)
    for rect in text_rects:
        r = rect.padded(padding, width, height)
        mask[r.y : r.y2, r.x : r.x2] = 255
    return mask


def _build_text_pixel_mask(image: Image.Image, text_rects: list[Rect], padding: int = 6) -> np.ndarray:
    arr = np.array(image.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(arr, cv2.COLOR_RGB2HSV)
    saturation = hsv[:, :, 1]
    mask = np.zeros((image.height, image.width), dtype=np.uint8)

    for rect in text_rects:
        r = rect.padded(padding, image.width, image.height)
        crop_gray = gray[r.y : r.y2, r.x : r.x2]
        crop_sat = saturation[r.y : r.y2, r.x : r.x2]
        if crop_gray.size == 0:
            continue

        median_gray = float(np.median(crop_gray))
        if median_gray < 120:
            local_pixels = (crop_gray > 155) | (np.abs(crop_gray.astype(np.int16) - int(median_gray)) > 55)
        else:
            local_pixels = (crop_gray < 185) | ((crop_gray < 225) & (crop_sat > 35))
        local = local_pixels.astype(np.uint8) * 255
        local = cv2.dilate(local, np.ones((3, 3), np.uint8), iterations=1)
        mask[r.y : r.y2, r.x : r.x2] = np.maximum(mask[r.y : r.y2, r.x : r.x2], local)

    return mask


def _remove_text_from_image(image: Image.Image, text_rects: list[Rect]) -> Image.Image:
    if not text_rects:
        return image

    arr = np.array(image.convert("RGB"))
    mask = _build_text_pixel_mask(image, text_rects, padding=6)
    if not np.any(mask):
        return image

    inpainted = cv2.inpaint(arr, mask, 3, cv2.INPAINT_TELEA)
    return Image.fromarray(inpainted)


def _component_rect_for_seed(mask: np.ndarray, seed_rect: Rect) -> Rect | None:
    num_labels, labels, stats, _centroids = cv2.connectedComponentsWithStats(mask, 8)
    if num_labels <= 1:
        return None

    y1 = max(0, seed_rect.y)
    y2 = min(mask.shape[0], seed_rect.y2)
    x1 = max(0, seed_rect.x)
    x2 = min(mask.shape[1], seed_rect.x2)
    seed_labels = labels[y1:y2, x1:x2]
    label_ids, counts = np.unique(seed_labels[seed_labels > 0], return_counts=True)
    if label_ids.size == 0:
        return None

    label = int(label_ids[int(np.argmax(counts))])
    x, y, w, h, _area = stats[label]
    return Rect(int(x), int(y), int(w), int(h))


def _detect_text_containers(slide_image: Image.Image, text_blocks: list[TextBlock], text_rects: list[Rect]) -> list[ShapeBlock]:
    if not text_blocks:
        return []

    cleaned_image = _remove_text_from_image(slide_image, text_rects)
    arr = np.array(cleaned_image.convert("RGB"))
    h, w = arr.shape[:2]
    shapes: list[ShapeBlock] = []

    for block in text_blocks:
        normalized_text = "".join(ch.lower() for ch in block.text if ch.isalnum())
        if len(normalized_text) <= 5 or normalized_text in {"msg", "6sw", "msq"}:
            continue

        seed = block.rect.padded(max(4, block.rect.h // 4), w, h)
        seed_crop = arr[seed.y : seed.y2, seed.x : seed.x2]
        if seed_crop.size == 0:
            continue

        median = np.median(seed_crop.reshape(-1, 3), axis=0)
        seed_std = float(np.mean(np.std(seed_crop.reshape(-1, 3), axis=0)))
        if seed_std > 58:
            continue

        distance = np.linalg.norm(arr.astype(np.int16) - median.astype(np.int16), axis=2)
        threshold = 34 if seed_std < 22 else 46
        similar = (distance < threshold).astype(np.uint8) * 255
        similar = cv2.morphologyEx(similar, cv2.MORPH_CLOSE, np.ones((7, 7), np.uint8), iterations=2)
        similar = cv2.morphologyEx(similar, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8), iterations=1)

        rect = _component_rect_for_seed(similar, seed)
        if rect is None:
            continue
        rect = rect.padded(1, w, h)
        if rect.area < block.rect.area * 1.08 or rect.area > w * h * 0.86:
            continue
        if rect.w < block.rect.w * 0.9 or rect.h < block.rect.h * 0.9:
            continue

        fill = _rect_fill_color(arr, rect)
        shapes.append(ShapeBlock(rect=rect, fill=fill))

    return _dedupe_shapes(shapes, w, h)


def _detect_visuals(slide_image: Image.Image, text_rects: list[Rect]) -> list[VisualBlock]:
    arr = np.array(slide_image.convert("RGB"))
    h, w = arr.shape[:2]
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(arr, cv2.COLOR_RGB2HSV)
    saturation = hsv[:, :, 1]

    text_mask = _build_text_mask(w, h, text_rects, padding=5)
    visual_source = _remove_text_from_image(slide_image, text_rects)

    # Keep non-white or colorful content, but remove OCR text so photos, logos, icons and colored panels remain.
    not_white = ((gray < 248) | (saturation > 28)).astype(np.uint8) * 255
    not_white[text_mask > 0] = 0

    kernel = np.ones((5, 5), np.uint8)
    cleaned = cv2.morphologyEx(not_white, cv2.MORPH_CLOSE, kernel, iterations=2)
    cleaned = cv2.morphologyEx(cleaned, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8), iterations=1)

    num_labels, labels, stats, _ = cv2.connectedComponentsWithStats(cleaned, 8)
    min_area = max(700, int(w * h * 0.002))
    visuals: list[VisualBlock] = []

    for label in range(1, num_labels):
        x, y, bw, bh, area = stats[label]
        if area < min_area or bw < 18 or bh < 18:
            continue
        rect = Rect(int(x), int(y), int(bw), int(bh)).padded(3, w, h)
        crop = visual_source.crop((rect.x, rect.y, rect.x2, rect.y2))
        visuals.append(VisualBlock(rect=rect, image=crop))

    return _dedupe_visuals(visuals, w, h)


def _dedupe_visuals(visuals: list[VisualBlock], max_w: int, max_h: int) -> list[VisualBlock]:
    kept: list[VisualBlock] = []
    for candidate in sorted(visuals, key=lambda item: item.rect.area, reverse=True):
        if any(_overlap_ratio(candidate.rect, existing.rect) > 0.82 for existing in kept):
            continue
        kept.append(candidate)

    kept.sort(key=lambda item: (item.rect.y, item.rect.x))
    return kept


def _dedupe_shapes(shapes: list[ShapeBlock], max_w: int, max_h: int) -> list[ShapeBlock]:
    kept: list[ShapeBlock] = []
    for candidate in sorted(shapes, key=lambda item: item.rect.area, reverse=True):
        if candidate.rect.area > max_w * max_h * 0.86:
            continue
        if any(_overlap_ratio(candidate.rect, existing.rect) > 0.85 for existing in kept):
            continue
        kept.append(candidate)

    kept.sort(key=lambda item: (item.rect.y, item.rect.x))
    return kept


def _is_bullet_line(text: str) -> bool:
    stripped = text.strip()
    return stripped.startswith(("-", "*", "\u2022", "\u00b7", "\u2013", ">>", "\u00bb"))


def _clean_line_text(text: str) -> tuple[str, bool]:
    stripped = " ".join(text.split())
    is_bullet = _is_bullet_line(stripped)
    if is_bullet:
        stripped = stripped.lstrip("-*\u2022\u00b7\u2013\u00bb> ").strip()
    return stripped, is_bullet


def _has_bullet_marker(arr: np.ndarray, rect: Rect) -> bool:
    h, w = arr.shape[:2]
    x1 = max(0, rect.x - max(12, rect.h * 2))
    x2 = max(0, rect.x - 2)
    y1 = max(0, rect.y - rect.h // 3)
    y2 = min(h, rect.y2 + rect.h // 3)
    if x2 <= x1 or y2 <= y1:
        return False

    crop = arr[y1:y2, x1:x2]
    gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(crop, cv2.COLOR_RGB2HSV)
    sat = hsv[:, :, 1]
    mask = ((gray < 170) | ((gray < 230) & (sat > 45))).astype(np.uint8) * 255
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8), iterations=1)
    num_labels, _labels, stats, _centroids = cv2.connectedComponentsWithStats(mask, 8)

    max_area = max(12, rect.h * rect.h)
    for label in range(1, num_labels):
        _x, _y, bw, bh, area = stats[label]
        if 4 <= area <= max_area and bw <= rect.h * 1.2 and bh <= rect.h * 1.2:
            return True
    return False


def _lines_belong_together(previous: TextLine, current: TextLine) -> bool:
    avg_h = max(1, (previous.rect.h + current.rect.h) / 2)
    vertical_gap = current.rect.y - previous.rect.y2
    if vertical_gap < -avg_h * 0.55 or vertical_gap > avg_h * 1.65:
        return False

    x_overlap = _intersection_area(
        Rect(previous.rect.x, 0, previous.rect.w, 1), Rect(current.rect.x, 0, current.rect.w, 1)
    )
    overlap_ratio = x_overlap / max(1, min(previous.rect.w, current.rect.w))
    left_delta = abs(previous.rect.x - current.rect.x)
    center_delta = abs((previous.rect.x + previous.rect.x2) / 2 - (current.rect.x + current.rect.x2) / 2)

    return overlap_ratio > 0.2 or left_delta <= avg_h * 3.2 or center_delta <= avg_h * 5.0


def _group_text_lines(lines: list[TextLine], slide_image: Image.Image) -> list[TextBlock]:
    if not lines:
        return []

    max_w, max_h = slide_image.width, slide_image.height
    arr = np.array(slide_image.convert("RGB"))
    sorted_lines = sorted(lines, key=lambda line: (line.rect.y, line.rect.x))
    groups: list[list[TextLine]] = []
    for line in sorted_lines:
        if groups and _lines_belong_together(groups[-1][-1], line):
            groups[-1].append(line)
        else:
            groups.append([line])

    blocks: list[TextBlock] = []
    for group in groups:
        cleaned_lines: list[tuple[TextLine, bool]] = []
        for line in group:
            cleaned, is_bullet = _clean_line_text(line.text)
            is_bullet = is_bullet or _has_bullet_marker(arr, line.rect)
            if cleaned:
                cleaned_lines.append((TextLine(line.rect, cleaned, line.confidence, line.color), is_bullet))
        if not cleaned_lines:
            continue

        block_lines: list[TextLine] = []
        bullet_flags: list[bool] = []
        for line, is_bullet in cleaned_lines:
            block_lines.append(line)
            bullet_flags.append(is_bullet)

        rect = _union_rect([line.rect for line in block_lines], max_w, max_h, padding=2)
        blocks.append(TextBlock(rect, block_lines, bullet_flags))

    return blocks


def analyze_image(
    image_path: Path,
    confidence: float = MIN_TEXT_CONFIDENCE,
    include_visuals: bool = True,
    progress: Callable[[str], None] | None = None,
) -> list[AnalyzedSlide]:
    progress = progress or (lambda _: None)
    source = Image.open(image_path).convert("RGB")
    slide_rects = detect_slide_rects(source)
    progress(f"{len(slide_rects)} Folie(n) erkannt.")

    ocr = RapidOCR()
    slides: list[AnalyzedSlide] = []
    for index, rect in enumerate(slide_rects, start=1):
        progress(f"Analysiere Folie {index}/{len(slide_rects)} ...")
        slide_img = source.crop((rect.x, rect.y, rect.x2, rect.y2))
        arr = np.array(slide_img)
        result, _ = ocr(arr)
        lines: list[TextLine] = []

        for item in result or []:
            points, text, score = item
            score_float = float(score)
            cleaned = " ".join(str(text).split())
            if score_float < confidence or not cleaned:
                continue
            text_rect = _box_to_rect(points, slide_img.width, slide_img.height)
            if text_rect.w < 4 or text_rect.h < 4:
                continue
            lines.append(
                TextLine(
                    rect=text_rect,
                    text=cleaned,
                    confidence=score_float,
                    color=_guess_text_color(arr, text_rect),
                )
            )

        texts = _group_text_lines(lines, slide_img)
        visuals = _detect_visuals(slide_img, [line.rect for line in lines]) if include_visuals else []
        shapes = _detect_text_containers(slide_img, texts, [line.rect for line in lines])
        slides.append(AnalyzedSlide(rect=rect, image=slide_img, texts=texts, visuals=visuals, shapes=shapes))
        progress(f"Folie {index}: {len(texts)} Textfeld(er), {len(visuals)} Bildobjekt(e), {len(shapes)} Flaeche(n).")

    return slides


def _add_textbox(slide, block: TextBlock, scale_x: float, scale_y: float) -> None:
    left = int(block.rect.x * scale_x)
    top = int(block.rect.y * scale_y)
    width = int(max(1, block.rect.w * scale_x * 1.12))
    height = int(max(1, block.rect.h * scale_y * 1.45))
    median_line_h = sorted(line.rect.h for line in block.lines)[len(block.lines) // 2]
    font_size = Pt(max(7, min(30, median_line_h * scale_y / EMU_PER_INCH * 72 * 1.05)))

    box = slide.shapes.add_textbox(left, top, width, height)
    box.text_frame.margin_left = 0
    box.text_frame.margin_right = 0
    box.text_frame.margin_top = 0
    box.text_frame.margin_bottom = 0
    box.text_frame.word_wrap = True
    box.text_frame.clear()

    for index, line in enumerate(block.lines):
        p = box.text_frame.paragraphs[0] if index == 0 else box.text_frame.add_paragraph()
        is_bullet = index < len(block.bullet_lines) and block.bullet_lines[index]
        p.text = f"\u2022 {line.text}" if is_bullet else line.text
        p.font.name = "Aptos"
        p.font.size = font_size
        p.font.color.rgb = line.color


def export_pptx(slides: list[AnalyzedSlide], output_path: Path, progress: Callable[[str], None] | None = None) -> None:
    if not slides:
        raise ValueError("Keine Folien zum Exportieren gefunden.")

    progress = progress or (lambda _: None)
    prs = Presentation()
    first = slides[0].image
    slide_height_in = SLIDE_WIDTH_IN * first.height / first.width
    prs.slide_width = int(SLIDE_WIDTH_IN * EMU_PER_INCH)
    prs.slide_height = int(slide_height_in * EMU_PER_INCH)
    blank_layout = prs.slide_layouts[6]

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_root = Path(temp_dir)
        for index, analyzed in enumerate(slides, start=1):
            progress(f"Exportiere Folie {index}/{len(slides)} ...")
            slide = prs.slides.add_slide(blank_layout)

            scale_x = prs.slide_width / analyzed.image.width
            scale_y = prs.slide_height / analyzed.image.height

            for visual_index, visual in enumerate(analyzed.visuals, start=1):
                image_path = temp_root / f"slide_{index:03d}_visual_{visual_index:03d}.png"
                visual.image.save(image_path)
                slide.shapes.add_picture(
                    str(image_path),
                    int(visual.rect.x * scale_x),
                    int(visual.rect.y * scale_y),
                    width=int(visual.rect.w * scale_x),
                    height=int(visual.rect.h * scale_y),
                )

            for shape in analyzed.shapes:
                rect = slide.shapes.add_shape(
                    MSO_SHAPE.RECTANGLE,
                    int(shape.rect.x * scale_x),
                    int(shape.rect.y * scale_y),
                    int(shape.rect.w * scale_x),
                    int(shape.rect.h * scale_y),
                )
                rect.fill.solid()
                rect.fill.fore_color.rgb = shape.fill
                rect.line.fill.background()

            for block in analyzed.texts:
                _add_textbox(slide, block, scale_x, scale_y)

        prs.save(output_path)


def convert_image_to_pptx(
    image_path: Path,
    output_path: Path,
    confidence: float = MIN_TEXT_CONFIDENCE,
    include_visuals: bool = True,
    progress: Callable[[str], None] | None = None,
) -> None:
    slides = analyze_image(image_path, confidence, include_visuals, progress)
    export_pptx(slides, output_path, progress)


def configure_windows_taskbar_icon() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(WINDOWS_APP_ID)
    except Exception:
        pass


def show_startup_splash(root: tk.Tk, minimum_ms: int = 3000) -> None:
    if os.environ.get("SNAP2SLIDES_SKIP_SPLASH") == "1" or not APP_SPLASH_PNG.exists():
        return

    root.withdraw()
    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(background="#020b1c")

    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    max_w = min(980, int(screen_w * 0.82))
    max_h = min(560, int(screen_h * 0.72))

    image = Image.open(APP_SPLASH_PNG).convert("RGB")
    image.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)
    splash._photo = ImageTk.PhotoImage(image)  # type: ignore[attr-defined]
    label = ttk.Label(splash, image=splash._photo, borderwidth=0)  # type: ignore[attr-defined]
    label.pack()

    splash.update_idletasks()
    x = (screen_w - splash.winfo_width()) // 2
    y = (screen_h - splash.winfo_height()) // 2
    splash.geometry(f"+{x}+{y}")
    splash.lift()

    def finish_splash() -> None:
        splash.destroy()
        root.deiconify()
        root.lift()

    root.after(minimum_ms, finish_splash)


class Image2PptApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Image2PPTSlicer")
        self._set_app_icon()
        self.geometry("860x620")
        self.minsize(760, 520)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.confidence = tk.DoubleVar(value=MIN_TEXT_CONFIDENCE)
        self.confidence_label = tk.StringVar()
        self.include_visuals = tk.BooleanVar(value=True)
        self.status_text = tk.StringVar(value="Bereit")
        self._preview_image: ImageTk.PhotoImage | None = None
        self._preview_source: Image.Image | None = None
        self._is_running = False

        self._build_ui()
        self._bind_shortcuts()
        self.input_path.trace_add("write", lambda *_args: self._mark_output_stale())
        self.output_path.trace_add("write", lambda *_args: self._mark_output_stale())
        self._update_confidence_label()

    def _set_app_icon(self) -> None:
        if APP_ICON_ICO.exists():
            try:
                self.iconbitmap(default=str(APP_ICON_ICO))
            except tk.TclError:
                pass
        if APP_ICON_PNG.exists():
            try:
                self._app_icon_photo = tk.PhotoImage(file=str(APP_ICON_PNG))
                self.iconphoto(True, self._app_icon_photo)
            except tk.TclError:
                self._app_icon_photo = None

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        root = ttk.Frame(self, padding=16)
        root.pack(fill=tk.BOTH, expand=True)

        form = ttk.Frame(root)
        form.pack(fill=tk.X)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Bild").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.input_path).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(form, text="Auswaehlen...", command=self._pick_input).grid(row=0, column=2, padx=(8, 0), pady=4)

        ttk.Label(form, text="PowerPoint").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.output_path).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(form, text="Speichern als...", command=self._pick_output).grid(row=1, column=2, padx=(8, 0), pady=4)

        options = ttk.Frame(root)
        options.pack(fill=tk.X, pady=(12, 8))
        ttk.Label(options, text="OCR-Mindestkonfidenz").pack(side=tk.LEFT)
        ttk.Scale(
            options,
            from_=0.2,
            to=0.8,
            variable=self.confidence,
            orient=tk.HORIZONTAL,
            length=220,
            command=lambda _value: self._update_confidence_label(),
        ).pack(side=tk.LEFT, padx=10)
        ttk.Label(options, textvariable=self.confidence_label, width=4).pack(side=tk.LEFT)
        self.visuals_check = ttk.Checkbutton(options, text="Bildobjekte erkennen", variable=self.include_visuals)
        self.visuals_check.pack(side=tk.LEFT, padx=16)
        self.convert_button = ttk.Button(options, text="PPTX erzeugen", command=self._start_conversion)
        self.convert_button.pack(side=tk.RIGHT)
        self.open_output_button = ttk.Button(
            options,
            text="Oeffnen",
            command=self._open_output,
            state=tk.DISABLED,
        )
        self.open_output_button.pack(side=tk.RIGHT, padx=(0, 8))

        preview_frame = ttk.LabelFrame(root, text="Vorschau")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 8))
        self.preview = ttk.Label(preview_frame, anchor=tk.CENTER, text="Noch kein Bild ausgewaehlt")
        self.preview.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.preview.bind("<Configure>", self._resize_preview)

        log_frame = ttk.LabelFrame(root, text="Status")
        log_frame.pack(fill=tk.BOTH, expand=False)
        self.log = tk.Text(log_frame, height=8, wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.log.configure(state=tk.DISABLED)

        bottom = ttk.Frame(root)
        bottom.pack(fill=tk.X, pady=(8, 0))
        self.progress = ttk.Progressbar(bottom, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(bottom, textvariable=self.status_text, width=32, anchor=tk.E).pack(side=tk.RIGHT, padx=(12, 0))

    def _bind_shortcuts(self) -> None:
        self.bind("<Control-o>", lambda _event: self._pick_input())
        self.bind("<Control-s>", lambda _event: self._pick_output())
        self.bind("<Return>", lambda _event: self._start_conversion())

    def _update_confidence_label(self) -> None:
        self.confidence_label.set(f"{self.confidence.get():.2f}")

    def _pick_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Bild auswaehlen",
            filetypes=[("Bilder", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("Alle Dateien", "*.*")],
        )
        if not path:
            return
        self.input_path.set(path)
        if not self.output_path.get():
            self.output_path.set(str(Path(path).with_suffix(".pptx")))
        self._load_preview(Path(path))

    def _pick_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="PowerPoint speichern",
            defaultextension=".pptx",
            filetypes=[("PowerPoint", "*.pptx")],
        )
        if path:
            self.output_path.set(path)

    def _load_preview(self, path: Path) -> None:
        try:
            self._preview_source = Image.open(path).convert("RGB")
            self._resize_preview()
            self.status_text.set(f"Bild geladen: {self._preview_source.width} x {self._preview_source.height} px")
        except Exception as exc:
            self._preview_source = None
            self.preview.configure(image="", text="Vorschau nicht verfuegbar")
            messagebox.showerror("Vorschau fehlgeschlagen", str(exc))

    def _resize_preview(self, _event: tk.Event | None = None) -> None:
        if self._preview_source is None:
            return
        width = max(160, self.preview.winfo_width() - 16)
        height = max(120, self.preview.winfo_height() - 16)
        image = self._preview_source.copy()
        image.thumbnail((width, height))
        self._preview_image = ImageTk.PhotoImage(image)
        self.preview.configure(image=self._preview_image, text="")

    def _append_log(self, message: str) -> None:
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)
        self.status_text.set(message)

    def _thread_log(self, message: str) -> None:
        self.after(0, self._append_log, message)

    def _start_conversion(self) -> None:
        if self._is_running:
            return
        input_text = self.input_path.get().strip()
        output_text = self.output_path.get().strip()
        if not input_text:
            messagebox.showerror("Eingabe fehlt", "Bitte ein Bild auswaehlen.")
            return
        if not output_text:
            messagebox.showerror("Ausgabe fehlt", "Bitte einen PowerPoint-Ausgabepfad angeben.")
            return

        input_path = Path(input_text)
        output_path = Path(output_text)
        if not input_path.is_file():
            messagebox.showerror("Eingabe fehlt", "Bitte ein vorhandenes Bild auswaehlen.")
            return
        if output_path.suffix.lower() != ".pptx":
            output_path = output_path.with_suffix(".pptx")
            self.output_path.set(str(output_path))
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("Ausgabe nicht moeglich", f"Der Zielordner konnte nicht angelegt werden:\n{exc}")
            return

        self._set_running(True)
        self.open_output_button.configure(state=tk.DISABLED)
        self._append_log("Starte Umwandlung ...")
        worker = threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path, float(self.confidence.get()), bool(self.include_visuals.get())),
            daemon=True,
        )
        worker.start()

    def _run_conversion(self, input_path: Path, output_path: Path, confidence: float, include_visuals: bool) -> None:
        try:
            convert_image_to_pptx(input_path, output_path, confidence, include_visuals, self._thread_log)
        except Exception as exc:
            self.after(0, messagebox.showerror, "Umwandlung fehlgeschlagen", str(exc))
            self._thread_log(f"Fehler: {exc}")
            self.after(0, self._set_running, False)
            return
        self._thread_log(f"Fertig: {output_path}")
        self.after(0, self._conversion_finished)
        self.after(0, messagebox.showinfo, "Fertig", f"PowerPoint wurde erzeugt:\n{output_path}")

    def _set_running(self, is_running: bool) -> None:
        self._is_running = is_running
        state = tk.DISABLED if is_running else tk.NORMAL
        self.convert_button.configure(state=state)
        self.visuals_check.configure(state=state)
        if is_running:
            self.progress.start(12)
        else:
            self.progress.stop()

    def _conversion_finished(self) -> None:
        self._set_running(False)
        if Path(self.output_path.get()).exists():
            self.open_output_button.configure(state=tk.NORMAL)

    def _mark_output_stale(self) -> None:
        if hasattr(self, "open_output_button"):
            self.open_output_button.configure(state=tk.DISABLED)

    def _open_output(self) -> None:
        path = Path(self.output_path.get())
        if not path.exists():
            messagebox.showerror("Datei fehlt", "Die PowerPoint-Datei wurde nicht gefunden.")
            return
        os.startfile(path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a slide image grid to an editable PPTX.")
    parser.add_argument("--input", "-i", type=Path)
    parser.add_argument("--output", "-o", type=Path)
    parser.add_argument("--confidence", type=float, default=MIN_TEXT_CONFIDENCE)
    parser.add_argument("--no-visuals", action="store_true")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.input and args.output:
        convert_image_to_pptx(
            args.input,
            args.output,
            confidence=args.confidence,
            include_visuals=not args.no_visuals,
            progress=print,
        )
        return

    configure_windows_taskbar_icon()
    app = Image2PptApp()
    show_startup_splash(app)
    app.mainloop()


if __name__ == "__main__":
    main()
