"""Streamlit Video Trimmer — bulk-trim videos using an Excel sheet of timestamps."""
import gc
import io
import subprocess
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

MAX_FILE_MB = 200
MAX_VIDEOS_PER_SESSION = 10
MAX_TOTAL_MB = 1500

st.set_page_config(page_title="Video Trimmer", page_icon="🎬", layout="centered")
st.title("🎬 Video Trimmer")
st.caption("Bulk-trim videos using an Excel sheet of start/end timestamps.")

st.info(
    f"**Limits per session:** up to **{MAX_VIDEOS_PER_SESSION} videos**, "
    f"**{MAX_FILE_MB} MB per file**, **{MAX_TOTAL_MB} MB total**. "
    "Memory is cleared after each session."
)

with st.expander("📋 Excel format — click to see", expanded=False):
    st.markdown(
        """
Your Excel file should have three columns:

| Video Name | Start | End |
|---|---|---|
| my_video_1 | 00:03:36 | 00:06:10 |
| my_video_2 | 00:10:00 | 00:12:30 |

- **Video Name** must match the uploaded video's filename (without the extension).
- **Start / End** in `HH:MM:SS` format (Excel time cells work too).
"""
    )


def _init_state():
    st.session_state.setdefault("session_id", 0)
    st.session_state.setdefault("zip_bytes", None)
    st.session_state.setdefault("result_summary", None)


def _reset_session():
    """Free memory and bump the session id so file uploaders clear."""
    st.session_state.zip_bytes = None
    st.session_state.result_summary = None
    st.session_state.session_id += 1
    gc.collect()


_init_state()
sid = st.session_state.session_id

excel_file = st.file_uploader(
    "1️⃣  Excel sheet",
    type=["xlsx", "xls"],
    key=f"excel_{sid}",
)
video_files = st.file_uploader(
    "2️⃣  Video files",
    type=["mp4", "mkv", "mov", "avi", "flv", "wmv"],
    accept_multiple_files=True,
    key=f"videos_{sid}",
)

mode = st.radio(
    "3️⃣  Trim mode",
    options=["accurate", "fast"],
    horizontal=True,
    format_func=lambda x: (
        "Accurate (re-encode — exact timing)"
        if x == "accurate"
        else "Fast (stream copy — keyframe aligned)"
    ),
    key=f"mode_{sid}",
)


def validate_uploads(videos):
    if not videos:
        return None
    if len(videos) > MAX_VIDEOS_PER_SESSION:
        return (
            f"Too many videos: {len(videos)}. "
            f"Limit is {MAX_VIDEOS_PER_SESSION} per session."
        )
    oversized = [v.name for v in videos if v.size > MAX_FILE_MB * 1024 * 1024]
    if oversized:
        return f"These files exceed {MAX_FILE_MB} MB: {', '.join(oversized)}"
    total_mb = sum(v.size for v in videos) / (1024 * 1024)
    if total_mb > MAX_TOTAL_MB:
        return (
            f"Total upload size is {total_mb:.0f} MB. "
            f"Limit is {MAX_TOTAL_MB} MB per session."
        )
    return None


if video_files:
    total_mb = sum(v.size for v in video_files) / (1024 * 1024)
    st.caption(
        f"📦 {len(video_files)} file(s) · {total_mb:.1f} MB / {MAX_TOTAL_MB} MB used"
    )

validation_error = validate_uploads(video_files)
if validation_error:
    st.error(validation_error)


def time_to_str(t):
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


def trim(input_path, start, end, output_path, mode):
    if mode == "fast":
        cmd = [
            "ffmpeg", "-y",
            "-ss", start, "-to", end,
            "-i", str(input_path),
            "-c", "copy",
            "-avoid_negative_ts", "make_zero",
            str(output_path),
        ]
    else:
        cmd = [
            "ffmpeg", "-y",
            "-i", str(input_path),
            "-ss", start, "-to", end,
            "-c:v", "libx264", "-c:a", "aac",
            "-preset", "medium", "-crf", "23",
            str(output_path),
        ]
    result = subprocess.run(cmd, capture_output=True)
    tail = result.stderr.decode("utf-8", errors="ignore")[-500:]
    return result.returncode == 0, tail


can_run = (
    excel_file is not None
    and video_files
    and validation_error is None
    and st.session_state.zip_bytes is None
)

run = st.button(
    "▶️  Trim videos",
    disabled=not can_run,
    type="primary",
    use_container_width=True,
)

if run:
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)

        video_map = {}
        for v in video_files:
            vpath = tmp / v.name
            vpath.write_bytes(v.getbuffer())
            video_map[Path(v.name).stem.lower()] = vpath

        df = pd.read_excel(excel_file)
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
            st.error(
                f"Excel is missing required columns. Found: {list(df.columns)}. "
                "Need columns for Video Name, Start, End."
            )
            st.stop()

        df = df.dropna(subset=["video_name", "start", "end"]).reset_index(drop=True)
        if len(df) == 0:
            st.warning("No rows with all three fields filled in.")
            st.stop()

        out_dir = tmp / "Trimmed"
        out_dir.mkdir(exist_ok=True)

        progress = st.progress(0.0, text="Starting…")
        log = st.container()
        success, failed = 0, []

        for i, row in df.iterrows():
            name = str(row["video_name"]).strip()
            start = time_to_str(row["start"])
            end = time_to_str(row["end"])

            progress.progress(i / len(df), text=f"[{i+1}/{len(df)}] {name}")

            key = name.lower()
            vpath = video_map.get(key)
            if not vpath:
                for k, p in video_map.items():
                    if key in k or k.startswith(key[:20]):
                        vpath = p
                        break

            if not vpath:
                failed.append(f"{name}: no matching uploaded video")
                log.warning(f"⚠️  **{name}** — no matching video in uploads")
                continue

            out_file = out_dir / f"{vpath.stem}_trimmed{vpath.suffix}"
            ok, err = trim(vpath, start, end, out_file, mode)
            if ok:
                success += 1
                log.success(f"✅  **{name}** → `{out_file.name}`")
            else:
                failed.append(f"{name}: {err}")
                log.error(f"❌  **{name}** — FFmpeg error")

        progress.progress(1.0, text="Done")

        if success > 0:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in out_dir.iterdir():
                    zf.write(f, f.name)
            zip_buf.seek(0)
            st.session_state.zip_bytes = zip_buf.getvalue()

        st.session_state.result_summary = {"success": success, "failed": failed}

    video_map.clear()
    del df
    gc.collect()
    st.rerun()


if st.session_state.zip_bytes is not None:
    summary = st.session_state.result_summary or {}
    st.success(
        f"Finished. Success: **{summary.get('success', 0)}** · "
        f"Failed: **{len(summary.get('failed', []))}**"
    )

    st.download_button(
        "⬇️  Download trimmed videos (ZIP)",
        data=st.session_state.zip_bytes,
        file_name="trimmed_videos.zip",
        mime="application/zip",
        use_container_width=True,
    )

    if summary.get("failed"):
        with st.expander("Show failures"):
            for f in summary["failed"]:
                st.text(f)

    st.divider()
    if st.button(
        "🧹 Clear & start new session",
        on_click=_reset_session,
        use_container_width=True,
    ):
        st.rerun()
