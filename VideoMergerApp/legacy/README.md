# The previous Video Merger

`video-merger.sh` is the entire previous app: its `.app` bundle's executable was
this bash script, driving the UI through `osascript` dialogs and merging with a
bundled ffmpeg. It lived only inside `/Applications/Video Merger.app` and was
never in this repository, so it is kept here rather than lost when that bundle
was replaced by the native rebuild.

It is the source the three modes in `../Sources/Merger.swift` were ported from —
the ffmpeg arguments, the VOB grouping rules and the re-encode fallback all
match it deliberately. It is reference material, not something that is built or
run any more.
