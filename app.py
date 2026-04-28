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
from pptx.util import Inches, Pt
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
class TextBlock:
    rect: Rect
    text: str
    confidence: float
    color: RGBColor


@dataclass
class VisualBlock:
    rect: Rect
    image: Image.Image


@dataclass
class AnalyzedSlide:
    rect: Rect
    image: Image.Image
    texts: list[TextBlock]
    visuals: list[VisualBlock]


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


def _detect_visuals(slide_image: Image.Image, text_rects: list[Rect]) -> list[VisualBlock]:
    arr = np.array(slide_image.convert("RGB"))
    h, w = arr.shape[:2]
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    hsv = cv2.cvtColor(arr, cv2.COLOR_RGB2HSV)
    saturation = hsv[:, :, 1]

    text_mask = np.zeros((h, w), dtype=np.uint8)
    for rect in text_rects:
        r = rect.padded(5, w, h)
        text_mask[r.y : r.y2, r.x : r.x2] = 255

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
        crop = slide_image.crop((rect.x, rect.y, rect.x2, rect.y2))
        visuals.append(VisualBlock(rect=rect, image=crop))

    visuals.sort(key=lambda item: item.rect.area, reverse=True)
    return _remove_contained_visuals(visuals)


def _remove_contained_visuals(visuals: list[VisualBlock]) -> list[VisualBlock]:
    kept: list[VisualBlock] = []
    for candidate in visuals:
        contained = False
        for existing in kept:
            cx1, cy1, cx2, cy2 = candidate.rect.x, candidate.rect.y, candidate.rect.x2, candidate.rect.y2
            ex1, ey1, ex2, ey2 = existing.rect.x, existing.rect.y, existing.rect.x2, existing.rect.y2
            if cx1 >= ex1 and cy1 >= ey1 and cx2 <= ex2 and cy2 <= ey2:
                contained = True
                break
        if not contained:
            kept.append(candidate)
    return kept


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
        texts: list[TextBlock] = []

        for item in result or []:
            points, text, score = item
            score_float = float(score)
            cleaned = " ".join(str(text).split())
            if score_float < confidence or not cleaned:
                continue
            text_rect = _box_to_rect(points, slide_img.width, slide_img.height)
            if text_rect.w < 4 or text_rect.h < 4:
                continue
            texts.append(
                TextBlock(
                    rect=text_rect,
                    text=cleaned,
                    confidence=score_float,
                    color=_guess_text_color(arr, text_rect),
                )
            )

        visuals = _detect_visuals(slide_img, [t.rect for t in texts]) if include_visuals else []
        slides.append(AnalyzedSlide(rect=rect, image=slide_img, texts=texts, visuals=visuals))
        progress(f"Folie {index}: {len(texts)} Textfeld(er), {len(visuals)} Bildobjekt(e).")

    return slides


def _add_textbox(slide, block: TextBlock, scale_x: float, scale_y: float) -> None:
    left = int(block.rect.x * scale_x)
    top = int(block.rect.y * scale_y)
    width = int(max(1, block.rect.w * scale_x * 1.08))
    height = int(max(1, block.rect.h * scale_y * 1.35))

    box = slide.shapes.add_textbox(left, top, width, height)
    box.text_frame.margin_left = 0
    box.text_frame.margin_right = 0
    box.text_frame.margin_top = 0
    box.text_frame.margin_bottom = 0
    box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]
    p.text = block.text
    p.font.name = "Aptos"
    p.font.size = Pt(max(7, min(28, height / EMU_PER_INCH * 72 * 0.72)))
    p.font.color.rgb = block.color


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
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
            bg.line.fill.background()

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
