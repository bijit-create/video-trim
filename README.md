# 🎬 Video Trimmer

A Streamlit web app for bulk-trimming videos using an Excel sheet of start/end timestamps. Upload your videos + spreadsheet, click trim, download all clips as a ZIP.

## Features

- Upload multiple videos at once
- Supply start/end times via Excel sheet (`HH:MM:SS`)
- Two modes: **accurate** (re-encode, exact timing) or **fast** (stream copy, keyframe-aligned)
- One-click ZIP download of all trimmed clips
- Memory is freed between sessions so the app stays responsive

## Limits (Streamlit Community Cloud free tier)

| Limit | Value |
|---|---|
| Max file size | 200 MB |
| Max videos per session | 10 |
| Max total upload per session | 1.5 GB |

These fit within the free-tier runtime (2.7 GB RAM, 2 CPU cores). For heavier workloads, deploy on a paid host or run locally.

## Excel format

| Video Name | Start | End |
|---|---|---|
| my_video_1 | 00:03:36 | 00:06:10 |
| my_video_2 | 00:10:00 | 00:12:30 |

The **Video Name** must match the uploaded file's name (without the extension). Matching is case-insensitive.

## Run locally

Prerequisites: Python 3.10+ and [FFmpeg](https://ffmpeg.org/download.html) installed and on your PATH.

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub (public).
2. Go to https://share.streamlit.io and sign in with GitHub.
3. Click **New app**, select this repo, set the main file to `app.py`, and deploy.
4. The `packages.txt` ensures FFmpeg is installed on the server automatically.

## License

MIT
