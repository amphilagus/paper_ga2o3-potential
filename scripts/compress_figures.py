#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["pillow"]
# ///
"""Compress images in the figures directory with size-aware settings."""

from __future__ import annotations

import argparse
import shutil
from dataclasses import dataclass
from pathlib import Path

from PIL import Image, ImageOps


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}


@dataclass(frozen=True)
class CompressionProfile:
    name: str
    max_side: int | None
    colors: int | None
    jpeg_quality: int
    webp_quality: int
    png_level: int


def profile_for_size(size_bytes: int) -> CompressionProfile:
    size_mb = size_bytes / (1024 * 1024)

    if size_mb >= 12:
        return CompressionProfile("very-large", 1800, 160, 78, 74, 9)
    if size_mb >= 6:
        return CompressionProfile("large", 2000, 192, 82, 78, 9)
    if size_mb >= 2:
        return CompressionProfile("medium", 2400, 256, 86, 82, 9)
    if size_mb >= 0.8:
        return CompressionProfile("small", None, None, 90, 86, 9)
    return CompressionProfile("tiny", None, None, 94, 90, 6)


def iter_images(input_dir: Path, recursive: bool) -> list[Path]:
    pattern = "**/*" if recursive else "*"
    return sorted(
        p
        for p in input_dir.glob(pattern)
        if p.is_file() and p.suffix.lower() in IMAGE_EXTENSIONS
    )


def resize_if_needed(image: Image.Image, max_side: int | None) -> Image.Image:
    if max_side is None:
        return image

    width, height = image.size
    longest_side = max(width, height)
    if longest_side <= max_side:
        return image

    scale = max_side / longest_side
    new_size = (round(width * scale), round(height * scale))
    return image.resize(new_size, Image.Resampling.LANCZOS)


def quantize_png_if_needed(image: Image.Image, colors: int | None) -> Image.Image:
    if colors is None:
        return image

    if image.mode in {"RGBA", "LA"}:
        return image.convert("RGBA").quantize(colors=colors, method=Image.Quantize.FASTOCTREE)

    return image.convert("RGB").quantize(colors=colors, method=Image.Quantize.MEDIANCUT)


def save_image(image: Image.Image, output_path: Path, profile: CompressionProfile) -> None:
    suffix = output_path.suffix.lower()

    if suffix == ".png":
        image = quantize_png_if_needed(image, profile.colors)
        image.save(output_path, optimize=True, compress_level=profile.png_level)
        return

    if suffix in {".jpg", ".jpeg"}:
        image.convert("RGB").save(
            output_path,
            optimize=True,
            progressive=True,
            quality=profile.jpeg_quality,
        )
        return

    if suffix == ".webp":
        image.save(output_path, method=6, quality=profile.webp_quality)
        return

    image.save(output_path)


def compress_one(input_path: Path, output_path: Path, dry_run: bool) -> tuple[int, int, str]:
    original_size = input_path.stat().st_size
    profile = profile_for_size(original_size)

    if dry_run:
        return original_size, original_size, profile.name

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(input_path) as image:
        image = ImageOps.exif_transpose(image)
        image = resize_if_needed(image, profile.max_side)
        save_image(image, output_path, profile)

    compressed_size = output_path.stat().st_size

    # Some already-optimized small figures can grow after recompression.
    if compressed_size >= original_size and input_path.resolve() != output_path.resolve():
        shutil.copy2(input_path, output_path)
        compressed_size = output_path.stat().st_size

    return original_size, compressed_size, profile.name


def format_size(size_bytes: int) -> str:
    return f"{size_bytes / (1024 * 1024):.2f} MB"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compress figures with stronger settings for larger images and gentler "
            "settings for smaller images."
        )
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("figures"),
        help="Directory containing original images. Default: figures",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("figures_compressed"),
        help="Directory for compressed images. Default: figures_compressed",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Include images in subdirectories.",
    )
    parser.add_argument(
        "--in-place",
        action="store_true",
        help="Overwrite original images. Use carefully.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show which profile would be used without writing files.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_dir = args.input_dir
    output_dir = input_dir if args.in_place else args.output_dir

    if not input_dir.exists():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    images = iter_images(input_dir, args.recursive)
    if not images:
        raise SystemExit(f"No images found in {input_dir}")

    total_before = 0
    total_after = 0

    for input_path in images:
        relative_path = input_path.relative_to(input_dir)
        output_path = output_dir / relative_path

        before, after, profile_name = compress_one(input_path, output_path, args.dry_run)
        total_before += before
        total_after += after

        if args.dry_run:
            print(f"[{profile_name:10}] {format_size(before):>8}  {input_path}")
            continue

        saved = 100 * (1 - after / before) if before else 0
        print(
            f"[{profile_name:10}] {format_size(before):>8} -> "
            f"{format_size(after):>8}  saved {saved:5.1f}%  {output_path}"
        )

    if not args.dry_run:
        saved = 100 * (1 - total_after / total_before) if total_before else 0
        print(
            f"\nTotal: {format_size(total_before)} -> {format_size(total_after)} "
            f"(saved {saved:.1f}%)"
        )


if __name__ == "__main__":
    main()
