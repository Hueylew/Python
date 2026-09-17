# Video Merger (native macOS app)

A native SwiftUI rebuild of the shell-script `Video Merger.app`. Same three
jobs, same bundled ffmpeg — but now a real window you can **drag files onto**,
with a **live progress bar**, that **stays open when a job finishes**.

## Why it was rebuilt rather than patched

The old app's executable was a bash script driving `osascript` dialogs. None of
the three things above are possible in that shape:

* a bash-script bundle has no run loop, so it never receives the Apple Event
  that a drop onto the app sends — nothing could be dropped on it;
* `display dialog` is modal and static, so there was nowhere to draw progress
  (the HEVC mode worked around this by opening a Terminal window);
* it was `LSUIElement`, had no window, and exited as soon as the work was done.

## Interface

One window:

* **Header** – app icon and what the app does.
* **Mode** – *Merge video files* / *Merge a DVD folder* / *Convert to MP4 (HEVC)*.
  Dropping a DVD folder switches to DVD mode by itself.
* **Centre** – the file list (top to bottom is the play order; drag a row, or
  use Move Up / Move Down) or, in DVD mode, the titles found on the disc.
* **Progress** – a determinate bar with the percentage, the current file,
  elapsed time, an estimate of what's left, and ffmpeg's encoding speed, plus
  **Cancel**. Cancelling deletes the half-written output.
* **Result** – every file written, with **Show in Finder** and **Open**. It
  stays on screen, and the window is ready for the next job.

The whole window is a drop target, and files can also be dropped on the Dock
icon or sent with *Open With*. Closing the window does not quit the app.

## What it does

| Mode | What happens |
|---|---|
| **Merge video files** | joins the clips in list order by stream copy (lossless, fast). If the clips don't share a format ffmpeg refuses, and the app offers a high-quality re-encode (libx264 CRF 18) instead. |
| **Merge a DVD folder** | finds `VIDEO_TS` (in the folder you picked, or up to three levels down), groups `VTS_<title>_<part>.VOB` into titles — skipping `_0`, the menu — and joins the chosen title's parts in order to one `.mpg`. Falls back to a byte-exact join if ffmpeg can't demux it. |
| **Convert to MP4 (HEVC)** | re-encodes each file to `<name>_x265.mp4` beside the original, numbered if that name is taken. *Quality* is libx265 CRF 20; *Speed* is Apple's hardware HEVC encoder. |

Merging and DVD joining ask where to save. Converting doesn't — each file is
written next to its original, so a save panel per file would only be in the way.

### How the progress bar knows

`ffprobe` reads each input's duration up front; ffmpeg is then run with
`-progress pipe:1`, whose `out_time_us` counts microseconds of finished output.
Measuring one against the other gives a true percentage rather than a bar that
merely spins. A file whose duration can't be read contributes nothing to the
total, and if *nothing* is known the bar falls back to indeterminate.

A one-second heartbeat redraws the window independently of that feed, so the
elapsed clock keeps moving even when ffmpeg says nothing. If the feed stays
quiet for more than five seconds the window says it is finalising — otherwise a
frozen bar is indistinguishable from a crash.

### Saving to a NAS

MP4 keeps its index (`moov`) at the end unless `+faststart` is used, and ffmpeg
implements that by reading the finished file back and writing the whole thing
out again. On a local disk that is quick. Over a network share it turns one
transfer into three — a 9GB merge onto an AFP-mounted NAS spent most of half an
hour doing it, reporting no progress throughout.

The index only needs to be at the front for progressive streaming straight off a
web server; every player that opens a file seeks to the end and finds it there.
So when the destination is not a local volume, `+faststart` is skipped and the
result says so. Local saves are unchanged.

## Requirements

* macOS 13+ (built and tested on macOS 26, Apple Silicon)
* Nothing else — `ffmpeg` and `ffprobe` are bundled inside the `.app`. A system
  install on `PATH` is used only if the bundled copies are missing.

## Build

```bash
cd VideoMergerApp
./build.sh
```

Produces `build/Video Merger.app` (ad-hoc signed, arm64 native, ~81MB). To
install it over the old one:

```bash
rm -rf "/Applications/Video Merger.app" && cp -R "build/Video Merger.app" /Applications/
```

First launch: right-click → **Open** once if Gatekeeper complains (ad-hoc
signature). macOS will ask for access to Downloads/Desktop/Documents the first
time the app touches those folders — allow it.

The build prefers a self-contained ffmpeg it finds in an installed copy of the
app (the old script bundled one) and caches it at `build/.ffmpeg-static`, so
replacing the installed app doesn't change what the next build produces.
Otherwise it copies Homebrew's ffmpeg and rewrites its library paths to point
inside the bundle.

## Source layout

| File | Role |
|---|---|
| `Sources/App.swift` | SwiftUI window, drop handling, view model |
| `Sources/Merger.swift` | the three jobs, and ffmpeg's progress feed |
| `Sources/Media.swift` | durations, DVD title detection, formatting |
| `Sources/Shell.swift` | subprocess runner with streaming stdout, cancellation, tool discovery |
| `make-icon.swift` | renders `AppIcon.icns` (two clips joining into one) |
| `build.sh` | renders the icon, compiles with `swiftc`, bundles ffmpeg, assembles and signs the `.app` |
