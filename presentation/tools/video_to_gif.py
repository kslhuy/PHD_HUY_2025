"""Convert videos to looping GIFs; no arguments opens a multiple-file picker."""

from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

try:
    import cv2
    from PIL import Image
except ImportError:
    raise SystemExit(
        "Missing packages. Run: py -3 -m pip install opencv-python Pillow"
    )


ASSETS = Path(__file__).resolve().parents[1] / "assets"


def positive(value: str) -> float:
    number = float(value)
    if not math.isfinite(number) or number <= 0:
        raise argparse.ArgumentTypeError("Must be a positive, finite number.")
    return number


def pick_videos() -> list[Path]:
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        paths = filedialog.askopenfilenames(
            parent=root,
            title="Select videos to convert to GIF (Ctrl/Shift for multiple)",
            filetypes=[
                ("Video files", "*.mp4 *.mov *.mkv *.avi *.webm *.m4v *.wmv"),
                ("All files", "*.*"),
            ],
        )
        return [Path(path) for path in paths]
    finally:
        root.destroy()


def resized_rgb(frame, width: int) -> Image.Image:
    height, original_width = frame.shape[:2]
    target_width = min(width, original_width)
    target_height = max(1, round(height * target_width / original_width))
    if target_width != original_width:
        frame = cv2.resize(
            frame, (target_width, target_height), interpolation=cv2.INTER_AREA
        )
    return Image.fromarray(cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))


def palette_for_video(cap, frame_count: int, colors: int) -> Image.Image:
    """Learn one palette across the clip to reduce flicker and file size."""
    samples = []
    count = min(24, frame_count)
    for i in range(count):
        index = round(i * (frame_count - 1) / max(1, count - 1))
        cap.set(cv2.CAP_PROP_POS_FRAMES, index)
        ok, frame = cap.read()
        if ok:
            samples.append(resized_rgb(frame, 320))
    if not samples:
        raise ValueError("No readable video frames.")
    sample_width, sample_height = samples[0].size
    sheet = Image.new("RGB", (sample_width * 4, sample_height * math.ceil(len(samples) / 4)))
    for i, sample in enumerate(samples):
        sheet.paste(sample, ((i % 4) * sample_width, (i // 4) * sample_height))
    return sheet.quantize(colors=colors, method=Image.Quantize.MEDIANCUT)


def free_path(path: Path) -> Path:
    """Preserve earlier exports if the same video is converted again."""
    candidate = path
    number = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{number}{path.suffix}")
        number += 1
    return candidate


def convert(source: Path, args) -> Path:
    source = source.expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    cap = cv2.VideoCapture(str(source))
    if not cap.isOpened():
        cap.release()
        raise ValueError(f"Cannot open video: {source}")
    frames = []
    try:
        source_fps = cap.get(cv2.CAP_PROP_FPS)
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
        if not math.isfinite(source_fps) or source_fps <= 0 or frame_count <= 0:
            raise ValueError("Cannot determine the video's frame rate or duration.")
        source_seconds = frame_count / source_fps
        # GIF stores time in hundredths of a second. Round the TOTAL duration
        # once, and distribute frame delays without accumulating timing error.
        total_cs = max(2, round(source_seconds / args.speed * 100))
        count = max(1, min(round(total_cs / 100 * args.fps), total_cs // 2))
        boundaries = [round(i * total_cs / count) for i in range(count + 1)]
        delays = [(boundaries[i + 1] - boundaries[i]) * 10 for i in range(count)]
        indices = [min(frame_count - 1, round(t / 100 * args.speed * source_fps))
                   for t in boundaries[:-1]]
        print(f"  Source: {source_seconds:.2f} s; GIF: {total_cs / 100:.2f} s; "
              f"speed: {args.speed:g}x; frames: {count}", flush=True)
        print(f"  Colors: {args.colors}; palette: {args.palette}; dither: {args.dither}", flush=True)
        palette = None
        if args.palette == "global":
            print("  Learning colors...", flush=True)
            palette = palette_for_video(cap, frame_count, args.colors)
        dither = (Image.Dither.FLOYDSTEINBERG if args.dither == "floyd"
                  else Image.Dither.NONE)
        # Reopen to guarantee a clean start after seeking for palette samples.
        cap.release()
        cap = cv2.VideoCapture(str(source))
        current = -1
        previous_image = None
        for i, target in enumerate(indices):
            if target != current:
                while current < target:
                    if not cap.grab():
                        raise ValueError(f"Video decoding ended early at frame {current + 1}.")
                    current += 1
                ok, frame = cap.retrieve()
                if not ok:
                    raise ValueError(f"Cannot decode frame {target}.")
                rgb = resized_rgb(frame, args.width)
                frame_palette = palette
                if frame_palette is None:
                    # Learn at output resolution so thin colored plot lines
                    # are represented instead of disappearing in thumbnails.
                    frame_palette = rgb.quantize(
                        colors=args.colors, method=Image.Quantize.MEDIANCUT
                    )
                previous_image = rgb.quantize(palette=frame_palette, dither=dither)
            frames.append(previous_image.copy())
            if (i + 1) % 50 == 0:
                print(f"  Converted {i + 1}/{count} frames...", flush=True)
    finally:
        cap.release()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    output = free_path(args.output_dir / f"{source.stem}_{args.speed:g}x.gif")
    print("  Saving GIF...", flush=True)
    try:
        frames[0].save(
            output, save_all=True, append_images=frames[1:], duration=delays,
            loop=0, optimize=True, disposal=1,
        )
        # Decode every saved frame and confirm timing, dimensions, and looping.
        with Image.open(output) as check:
            actual_ms = 0
            size = check.size
            if check.info.get("loop") != 0:
                raise ValueError("GIF loop verification failed.")
            for i in range(check.n_frames):
                check.seek(i)
                check.load()
                actual_ms += check.info.get("duration", 0)
            if actual_ms != sum(delays):
                raise ValueError("GIF duration verification failed.")
        original_bytes, gif_bytes = source.stat().st_size, output.stat().st_size
        reduction = (1 - gif_bytes / original_bytes) * 100
        comparison = (f"{reduction:.1f}% smaller" if reduction >= 0
                      else f"{-reduction:.1f}% larger")
        print(f"  Saved: {output}\n  {size[0]} x {size[1]}; "
              f"{actual_ms / 1000:.2f} s; {gif_bytes / 1_000_000:.2f} MB "
              f"({comparison} than source)", flush=True)
        if gif_bytes >= original_bytes:
            print("  Tip: lower --width, --fps or --colors for a smaller GIF.")
    except Exception:
        if output.exists():
            output.unlink()
        raise
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("videos", nargs="*", type=Path)
    parser.add_argument("--speed", type=positive, default=5, help="Playback speed multiplier (default: 5).")
    parser.add_argument("--width", type=int, default=800, help="Maximum width in pixels (default: 800).")
    parser.add_argument("--fps", type=positive, default=8, help="Output frames per second (default: 8).")
    parser.add_argument("--colors", type=int, default=256, help="Palette size, 2 to 256 (default: 256).")
    parser.add_argument("--palette", choices=("frame", "global"), default="frame",
                        help="Learn colors per frame or across the clip (default: frame).")
    parser.add_argument("--dither", choices=("floyd", "none"), default="floyd",
                        help="Smooth color transitions with dithering (default: floyd).")
    parser.add_argument("--output-dir", type=Path, default=ASSETS)
    args = parser.parse_args()
    if args.width < 1 or not 2 <= args.colors <= 256 or args.fps > 50:
        parser.error("width must be positive; colors must be 2..256; fps must be <= 50.")
    videos = args.videos or pick_videos()
    if not videos:
        print("No videos selected.")
        return 0
    failures = 0
    for source in videos:
        print(f"\nConverting: {source.name}", flush=True)
        try:
            convert(source, args)
        except Exception as error:
            failures += 1
            print(f"  ERROR: {error}", file=sys.stderr, flush=True)
    print(f"\nDone: {len(videos) - failures}/{len(videos)} videos. Output: {args.output_dir}")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
