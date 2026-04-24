"""
Video Trimmer Tool
==================
Trims multiple videos based on an Excel sheet containing video names, start times, and end times.

Excel format expected:
    Column A: Video Name      (e.g., अपरिमेय संख्याएँ_Part 1)
    Column B: Starting        (e.g., 00:03:36)
    Column C: Endling         (e.g., 00:06:10)

Usage:
    python video_trimmer.py

Requirements:
    pip install pandas openpyxl
    FFmpeg installed and accessible in PATH (download from https://ffmpeg.org/download.html)
"""

import os
import sys
import subprocess
import pandas as pd
from pathlib import Path


# ============================================================
# CONFIGURATION - EDIT THESE PATHS ACCORDING TO YOUR SETUP
# ============================================================
EXCEL_FILE    = r"C:\Users\Aamir\Downloads\Trim video\Video Trimming.xlsx"
VIDEO_FOLDER  = r"C:\Users\Aamir\Downloads\Trim video"
OUTPUT_FOLDER = r"C:\Users\Aamir\Downloads\Trim video\Trimmed"
VIDEO_EXTENSIONS = [".mp4", ".mkv", ".avi", ".mov", ".flv", ".wmv"]

# Encoding mode:
#   "fast"     -> Stream copy (no re-encode). Very fast but cuts on nearest keyframe.
#   "accurate" -> Re-encode. Slower but exact timing.
ENCODING_MODE = "accurate"
# ============================================================


def check_ffmpeg():
    """Verify that FFmpeg is installed and available."""
    try:
        subprocess.run(
            ["ffmpeg", "-version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("ERROR: FFmpeg is not installed or not in PATH.")
        print("Download it from: https://ffmpeg.org/download.html")
        print("After installing, add the 'bin' folder to your system PATH.")
        return False


def time_to_seconds(t):
    """Convert HH:MM:SS (or datetime.time / pandas Timedelta) to seconds string for ffmpeg."""
    if pd.isna(t):
        return None
    if hasattr(t, "hour") and hasattr(t, "minute") and hasattr(t, "second"):
        return f"{t.hour:02d}:{t.minute:02d}:{t.second:02d}"
    if isinstance(t, pd.Timedelta):
        total = int(t.total_seconds())
        h, rem = divmod(total, 3600)
        m, s = divmod(rem, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"
    return str(t).strip()


def find_video_file(video_name, folder):
    """Find a video file in `folder` whose stem matches video_name (case-insensitive)."""
    folder = Path(folder)
    target = video_name.strip().lower()
    for file in folder.iterdir():
        if file.is_file() and file.suffix.lower() in VIDEO_EXTENSIONS:
            if file.stem.lower() == target:
                return file
    for file in folder.iterdir():
        if file.is_file() and file.suffix.lower() in VIDEO_EXTENSIONS:
            if target in file.stem.lower() or file.stem.lower().startswith(target[:20]):
                return file
    return None


def trim_video(input_path, start_time, end_time, output_path, mode="accurate"):
    """Trim a video using FFmpeg."""
    if mode == "fast":
        cmd = [
            "ffmpeg", "-y",
            "-ss", start_time,
            "-to", end_time,
            "-i", str(input_path),
            "-c", "copy",
            "-avoid_negative_ts", "make_zero",
            str(output_path),
        ]
    else:
        cmd = [
            "ffmpeg", "-y",
            "-i", str(input_path),
            "-ss", start_time,
            "-to", end_time,
            "-c:v", "libx264",
            "-c:a", "aac",
            "-preset", "medium",
            "-crf", "23",
            str(output_path),
        ]

    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    if result.returncode != 0:
        print(f"   FFmpeg error: {result.stderr.decode('utf-8', errors='ignore')[-500:]}")
        return False
    return True


def main():
    print("=" * 60)
    print("Video Trimmer Tool")
    print("=" * 60)

    if not check_ffmpeg():
        sys.exit(1)

    if not os.path.isfile(EXCEL_FILE):
        print(f"ERROR: Excel file not found: {EXCEL_FILE}")
        sys.exit(1)
    if not os.path.isdir(VIDEO_FOLDER):
        print(f"ERROR: Video folder not found: {VIDEO_FOLDER}")
        sys.exit(1)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    print(f"\nReading Excel: {EXCEL_FILE}")
    df = pd.read_excel(EXCEL_FILE)

    df.columns = [str(c).strip().lower() for c in df.columns]
    col_map = {}
    for c in df.columns:
        if "video" in c or "name" in c:
            col_map[c] = "video_name"
        elif "start" in c:
            col_map[c] = "start"
        elif "end" in c or "endl" in c:
            col_map[c] = "end"
    df = df.rename(columns=col_map)

    required = {"video_name", "start", "end"}
    if not required.issubset(df.columns):
        print(f"ERROR: Excel must contain columns for Video Name, Starting, Ending.")
        print(f"Found columns: {list(df.columns)}")
        sys.exit(1)

    df = df.dropna(subset=["video_name", "start", "end"]).reset_index(drop=True)
    print(f"Found {len(df)} trim task(s).\n")

    success, failed = 0, 0
    for idx, row in df.iterrows():
        name = str(row["video_name"]).strip()
        start = time_to_seconds(row["start"])
        end = time_to_seconds(row["end"])

        print(f"[{idx+1}/{len(df)}] {name}")
        print(f"   Start: {start}   End: {end}")

        video_file = find_video_file(name, VIDEO_FOLDER)
        if not video_file:
            print(f"   SKIPPED: video file not found in folder.\n")
            failed += 1
            continue

        print(f"   Found: {video_file.name}")
        output_file = Path(OUTPUT_FOLDER) / f"{video_file.stem}_trimmed{video_file.suffix}"

        if trim_video(video_file, start, end, output_file, mode=ENCODING_MODE):
            print(f"   Saved: {output_file}\n")
            success += 1
        else:
            print(f"   FAILED\n")
            failed += 1

    print("=" * 60)
    print(f"Done. Success: {success}   Failed: {failed}")
    print(f"Output folder: {OUTPUT_FOLDER}")
    print("=" * 60)


if __name__ == "__main__":
    main()
