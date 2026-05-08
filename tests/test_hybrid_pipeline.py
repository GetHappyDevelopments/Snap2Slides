from __future__ import annotations

import zipfile
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw

import app


def _make_slide_grid(path: Path) -> None:
    image = Image.new("RGB", (760, 250), (238, 238, 238))
    draw = ImageDraw.Draw(image)
    for left in (24, 394):
        draw.rectangle((left, 24, left + 320, 24 + 180), fill=(255, 255, 255), outline=(0, 0, 0), width=6)
        draw.rectangle((left + 42, 74, left + 142, 124), fill=(220, 40, 80))
        draw.line((left + 176, 142, left + 278, 142), fill=(20, 80, 210), width=5)
        draw.text((left + 44, 38), "Title", fill=(20, 20, 20))
    image.save(path)


class _FakeOCR:
    def __call__(self, image_array):
        return [
            (
                [[120, 42], [240, 42], [240, 84], [120, 84]],
                "Title",
                0.92,
            )
        ], None


def test_slide_detection_sorts_grid_and_border_trim_removes_frame(tmp_path, monkeypatch):
    input_path = tmp_path / "grid.png"
    _make_slide_grid(input_path)
    monkeypatch.setattr(app, "RapidOCR", lambda: _FakeOCR())

    slides = app.analyze_image(
        input_path,
        confidence=0.5,
        config=app.ReconstructionConfig(border_trim_px=10, min_slide_area=12000),
        debug_dir=tmp_path / "debug",
    )

    assert len(slides) == 2
    assert slides[0].source_rect.x < slides[1].source_rect.x
    assert slides[0].image.getpixel((0, 0)) != (0, 0, 0)
    assert slides[0].texts[0].text == "Title"
    assert any(vector.kind in {"rect", "line"} for vector in slides[0].vectors or [])
    assert (tmp_path / "debug" / "slide_001.svg").exists()
    assert (tmp_path / "debug" / "reconstruction_report.json").exists()


def test_hybrid_vector_and_pixel_modes_create_valid_pptx(tmp_path):
    slide_image = Image.new("RGB", (320, 180), (255, 255, 255))
    photo = Image.fromarray(np.random.default_rng(42).integers(0, 255, (50, 70, 3), dtype=np.uint8))
    analyzed = app.AnalyzedSlide(
        rect=app.Rect(0, 0, 320, 180),
        image=slide_image,
        texts=[
            app.TextBlock(
                rect=app.Rect(36, 24, 92, 20),
                lines=[app.TextLine(app.Rect(36, 24, 92, 20), "Editable text", 0.95, app.TEXT_BLACK)],
                bullet_lines=[False],
            )
        ],
        visuals=[app.VisualBlock(app.Rect(200, 64, 70, 50), photo, "photo")],
        shapes=[app.ShapeBlock(app.Rect(30, 70, 80, 42), app.RGBColor(230, 230, 250))],
        vectors=[
            app.VectorElement("rect", app.Rect(132, 70, 50, 38), 0.88, fill=app.RGBColor(220, 40, 80)),
            app.VectorElement(
                "line",
                app.Rect(30, 135, 120, 4),
                0.90,
                stroke=app.RGBColor(20, 80, 210),
                stroke_width=3,
                points=[(30, 137), (150, 137)],
            ),
        ],
        background=app.RGBColor(255, 255, 255),
    )

    for mode in ("pixel", "hybrid", "vector"):
        output = tmp_path / f"{mode}.pptx"
        app.export_pptx([analyzed], output, mode=mode, debug_dir=tmp_path / f"{mode}_debug")
        valid, errors = app.validate_pptx(output)
        assert valid, errors
        with zipfile.ZipFile(output) as archive:
            assert "ppt/presentation.xml" in archive.namelist()


def test_text_blocks_are_sorted_by_reading_order():
    slide_image = Image.new("RGB", (420, 240), (255, 255, 255))
    lines = [
        app.TextLine(app.Rect(230, 92, 80, 18), "Right body", 0.95, app.TEXT_BLACK),
        app.TextLine(app.Rect(34, 92, 80, 18), "Left body", 0.95, app.TEXT_BLACK),
        app.TextLine(app.Rect(34, 24, 120, 20), "Title", 0.95, app.TEXT_BLACK),
    ]

    blocks = app._group_text_lines(lines, slide_image)

    assert [block.text for block in blocks] == ["Title", "Left body", "Right body"]
    assert [block.order for block in blocks] == [0, 1, 2]


def test_large_images_icons_and_white_regions_are_separated():
    slide_image = Image.new("RGB", (400, 240), (255, 255, 255))
    draw = ImageDraw.Draw(slide_image)
    rng = np.random.default_rng(7)
    photo = Image.fromarray(rng.integers(0, 255, (92, 112, 3), dtype=np.uint8))
    slide_image.paste(photo, (240, 54))
    draw.rectangle((38, 80, 62, 104), fill=(185, 20, 120))
    draw.rectangle((82, 80, 106, 104), fill=(185, 20, 120))
    draw.rectangle((0, 0, 399, 239), outline=(255, 255, 255), width=1)

    visual_candidates = app._detect_visuals(slide_image, [])
    large_images = app._detect_large_image_regions(slide_image, [], visual_candidates)
    icons = app._detect_icon_regions(slide_image, [visual.rect for visual in large_images])

    assert large_images
    assert all(not app._is_mostly_white_region(np.array(slide_image), visual.rect) for visual in large_images)
    assert any(visual.kind == "icon" for visual in icons)
    assert not any(app._overlap_ratio(icon.rect, large.rect) > 0.2 for icon in icons for large in large_images)


def test_export_uses_edited_positions_and_validates_pptx(tmp_path):
    slide_image = Image.new("RGB", (320, 180), (255, 255, 255))
    block = app.TextBlock(
        rect=app.Rect(150, 90, 90, 26),
        lines=[app.TextLine(app.Rect(150, 90, 90, 26), "Edited", 0.95, app.TEXT_BLACK)],
        bullet_lines=[False],
    )
    analyzed = app.AnalyzedSlide(
        rect=app.Rect(0, 0, 320, 180),
        image=slide_image,
        texts=[block],
        visuals=[],
        shapes=[],
        vectors=[],
        background=app.RGBColor(255, 255, 255),
    )

    output = tmp_path / "edited.pptx"
    app.export_pptx([analyzed], output, mode="hybrid")

    valid, errors = app.validate_pptx(output)
    assert valid, errors
    with zipfile.ZipFile(output) as archive:
        slide_xml = archive.read("ppt/slides/slide1.xml").decode("utf-8")
    assert "Edited" in slide_xml


def test_merge_text_blocks_keeps_reading_order_and_union_rect():
    upper = app.TextBlock(
        rect=app.Rect(40, 24, 90, 18),
        lines=[app.TextLine(app.Rect(40, 24, 90, 18), "Upper", 0.95, app.TEXT_BLACK)],
        bullet_lines=[False],
        order=0,
    )
    lower_left = app.TextBlock(
        rect=app.Rect(40, 62, 70, 18),
        lines=[app.TextLine(app.Rect(40, 62, 70, 18), "Left", 0.95, app.TEXT_BLACK)],
        bullet_lines=[False],
        order=1,
    )
    lower_right = app.TextBlock(
        rect=app.Rect(180, 62, 80, 18),
        lines=[app.TextLine(app.Rect(180, 62, 80, 18), "Right", 0.95, app.TEXT_BLACK)],
        bullet_lines=[False],
        order=2,
    )

    merged = app._merge_text_blocks([lower_right, upper, lower_left], 320, 180)

    assert merged.text == "Upper\nLeft\nRight"
    assert merged.rect.x <= 40
    assert merged.rect.y <= 24
    assert merged.rect.x2 >= 260
    assert merged.rect.y2 >= 80
