#!/bin/bash
# Video Merger - losslessly stitch clips end-to-end.
#   * Video files mode: join MP4/MOV/M4V (stream copy; optional HQ re-encode).
#   * DVD mode: pick a VIDEO_TS folder, join a title's VOB parts to one .mpg.
# The app's executable is this shell script; UI is via osascript, merging via
# a bundled ffmpeg (or a system ffmpeg if present). Written for macOS bash 3.2.
set -u

DIR="$(cd "$(dirname "$0")" && pwd)"                 # .../Contents/MacOS
APP="$(cd "$DIR/../.." && pwd)"                       # .../VideoMerger.app
RES="$(cd "$DIR/../Resources" && pwd)"                # .../Contents/Resources
BUNDLED="$RES/ffmpeg"

/usr/bin/xattr -dr com.apple.quarantine "$APP" >/dev/null 2>&1 || true
[ -f "$BUNDLED" ] && chmod +x "$BUNDLED" >/dev/null 2>&1 || true

FFMPEG=""
for c in "$BUNDLED" /opt/homebrew/bin/ffmpeg /usr/local/bin/ffmpeg "$(command -v ffmpeg 2>/dev/null || true)"; do
  if [ -n "${c:-}" ] && [ -x "$c" ]; then FFMPEG="$c"; break; fi
done
if [ -z "$FFMPEG" ]; then
  /usr/bin/osascript -e 'display alert "Video Merger" message "Could not find the ffmpeg engine inside the app." as critical' >/dev/null 2>&1
  exit 1
fi

osa() { /usr/bin/osascript "$@" 2>/dev/null; }
note() { osa -e "display notification \"$1\" with title \"Video Merger\"" >/dev/null 2>&1; }
crit() { osa -e "display alert \"Video Merger\" message \"$1\" as critical" >/dev/null 2>&1; }
done_dialog() { # $1 = output path
  local btn
  btn="$(osa -e "display alert \"Done\" message \"Saved to:
$1\" buttons {\"OK\", \"Show in Finder\"} default button \"Show in Finder\"" -e 'button returned of result')"
  [ "$btn" = "Show in Finder" ] && open -R "$1"
}

# ------------------------------------------------------------------ MODE
MODE="$(osa -e 'set m to choose from list {"Merge video files (MP4, MOV, M4V)", "Merge a DVD folder (VIDEO_TS)", "Convert to MP4 (x265 / HEVC)"} with prompt "What do you want to do?" default items {"Merge video files (MP4, MOV, M4V)"} without empty selection allowed
if m is false then return "CANCEL"
return item 1 of m')"
[ -z "$MODE" ] && exit 0
[ "$MODE" = "CANCEL" ] && exit 0

# =================================================================== VIDEO FILES
if [ "${MODE#Merge video files}" != "$MODE" ]; then
  PLAN="$(osa <<'AS'
on run
	set inFiles to choose file with prompt "Select the video files to merge (you'll set the order next):" of type {"mp4", "m4v", "mov"} with multiple selections allowed
	set total to count of inFiles
	if total < 2 then
		display alert "Video Merger" message "Please pick at least two clips to join." as warning
		error number -128
	end if
	set paths to {}
	set displays to {}
	set usedFlags to {}
	repeat with i from 1 to total
		set p to POSIX path of (item i of inFiles)
		set end of paths to p
		set AppleScript's text item delimiters to "/"
		set baseName to last text item of p
		set AppleScript's text item delimiters to ""
		set end of displays to (i as string) & " - " & baseName
		set end of usedFlags to false
	end repeat
	set orderedPaths to {}
	repeat with stepN from 1 to (total - 1)
		set remain to {}
		repeat with i from 1 to total
			if item i of usedFlags is false then set end of remain to item i of displays
		end repeat
		set choice to choose from list remain with prompt "Choose clip #" & stepN & " of " & total & " (top to bottom = play order):" default items {item 1 of remain} without empty selection allowed
		if choice is false then error number -128
		set chosenDisp to item 1 of choice
		set AppleScript's text item delimiters to " - "
		set idx to (text item 1 of chosenDisp) as integer
		set AppleScript's text item delimiters to ""
		set item idx of usedFlags to true
		set end of orderedPaths to item idx of paths
	end repeat
	repeat with i from 1 to total
		if item i of usedFlags is false then set end of orderedPaths to item i of paths
	end repeat
	set outFile to choose file name with prompt "Save the merged video as:" default name "merged.mp4"
	set outPath to POSIX path of outFile
	if outPath does not end with ".mp4" and outPath does not end with ".mov" and outPath does not end with ".m4v" then set outPath to outPath & ".mp4"
	set outText to outPath
	repeat with p in orderedPaths
		set outText to outText & linefeed & (p as text)
	end repeat
	return outText
end run
AS
)"
  [ $? -ne 0 ] && exit 0
  [ -z "$PLAN" ] && exit 0

  OUT="$(printf '%s\n' "$PLAN" | sed -n '1p')"
  INPUTS=()
  while IFS= read -r line; do
    [ -n "$line" ] && INPUTS+=("$line")
  done < <(printf '%s\n' "$PLAN" | sed -n '2,$p')
  [ "${#INPUTS[@]}" -lt 2 ] && { crit "Need at least two clips to merge."; exit 0; }

  WORK="$(mktemp -d /tmp/videomerger.XXXXXX)"; LIST="$WORK/list.txt"; ERR="$WORK/err.txt"; : > "$LIST"
  for p in "${INPUTS[@]}"; do
    esc="${p//\'/\'\\\'\'}"
    printf "file '%s'\n" "$esc" >> "$LIST"
  done
  note "Merging your clips..."
  if "$FFMPEG" -y -hide_banner -loglevel error -f concat -safe 0 -i "$LIST" -c copy -movflags +faststart "$OUT" 2>"$ERR"; then
    done_dialog "$OUT"; rm -rf "$WORK"; exit 0
  fi
  CHOICE="$(osa -e 'display dialog "These clips don'"'"'t share the same format, so they can'"'"'t be joined without re-encoding.

Re-encoding produces one clean file at high quality (visually near-identical), but it is slower and technically re-compresses the video." buttons {"Cancel", "Re-encode (high quality)"} default button "Re-encode (high quality)" with title "Video Merger" with icon caution' -e 'button returned of result')"
  if [ "$CHOICE" = "Re-encode (high quality)" ]; then
    note "Re-encoding... this can take a while."
    if "$FFMPEG" -y -hide_banner -loglevel error -f concat -safe 0 -i "$LIST" \
        -c:v libx264 -crf 18 -preset medium -pix_fmt yuv420p -c:a aac -b:a 256k -movflags +faststart "$OUT" 2>"$ERR"; then
      done_dialog "$OUT"; rm -rf "$WORK"; exit 0
    fi
  fi
  MSG="$(tail -n 6 "$ERR" 2>/dev/null | tr '"' "'")"
  crit "Could not merge:
$MSG"; rm -rf "$WORK"; exit 1
fi

# =================================================================== CONVERT -> MP4 (HEVC/x265)
if [ "${MODE#Convert}" != "$MODE" ]; then
  FILES="$(osa -e 'set fs to choose file with prompt "Select the video file(s) to convert to MP4 (x265):" of type {"mp4", "m4v", "mov", "m2ts", "mts", "ts", "mkv", "mpg", "mpeg", "vob", "avi", "webm"} with multiple selections allowed
set out to ""
repeat with f in fs
	set out to out & POSIX path of f & linefeed
end repeat
return out')"
  [ $? -ne 0 ] && exit 0
  [ -z "$FILES" ] && exit 0
  INPUTS=()
  while IFS= read -r line; do
    [ -n "$line" ] && INPUTS+=("$line")
  done < <(printf '%s\n' "$FILES")
  [ "${#INPUTS[@]}" -lt 1 ] && exit 0

  ENC="$(osa -e 'set c to choose from list {"Quality - best compression (libx265, slower)", "Speed - hardware accelerated (Apple HEVC, faster)"} with prompt "How should it encode to HEVC / x265?" default items {"Quality - best compression (libx265, slower)"} without empty selection allowed
if c is false then return "CANCEL"
return item 1 of c')"
  [ -z "$ENC" ] && exit 0
  [ "$ENC" = "CANCEL" ] && exit 0
  if [ "${ENC#Speed}" != "$ENC" ]; then MODEENC="speed"; else MODEENC="quality"; fi

  # Build a Terminal .command script so the (potentially long) encode shows live progress.
  CMD="$(mktemp /tmp/videomerger_encode_XXXXXX)"; CMD="${CMD}.command"
  {
    echo '#!/bin/bash'
    echo 'clear'
    echo 'echo "Video Merger  -  HEVC / x265 conversion"'
    printf 'echo "Mode: %s"\n' "$MODEENC"
    echo 'echo "-----------------------------------------"'
    echo 'fail=0'
  } > "$CMD"
  for in in "${INPUTS[@]}"; do
    dir="$(dirname "$in")"; base="$(basename "$in")"; stem="${base%.*}"
    out="$dir/${stem}_x265.mp4"; n=1
    while [ -e "$out" ]; do out="$dir/${stem}_x265_$n.mp4"; n=$((n+1)); done
    qff="$(printf %q "$FFMPEG")"; qin="$(printf %q "$in")"; qout="$(printf %q "$out")"; qbase="$(printf %q "$base")"
    {
      printf 'echo; echo "Converting: %s"\n' "$qbase"
      if [ "$MODEENC" = "speed" ]; then
        printf '%s -y -hide_banner -stats -i %s -map 0:v:0 -map 0:a? -c:v hevc_videotoolbox -q:v 60 -tag:v hvc1 -c:a aac -b:a 192k -movflags +faststart %s' "$qff" "$qin" "$qout"
        printf ' || %s -y -hide_banner -stats -i %s -map 0:v:0 -map 0:a? -c:v hevc_videotoolbox -b:v 10M -tag:v hvc1 -c:a aac -b:a 192k -movflags +faststart %s' "$qff" "$qin" "$qout"
        printf ' || fail=1\n'
      else
        printf '%s -y -hide_banner -stats -i %s -map 0:v:0 -map 0:a? -c:v libx265 -crf 20 -preset medium -tag:v hvc1 -c:a aac -b:a 192k -movflags +faststart %s || fail=1\n' "$qff" "$qin" "$qout"
      fi
      printf 'open -R %s 2>/dev/null || true\n' "$qout"
    } >> "$CMD"
  done
  {
    echo 'echo; echo "-----------------------------------------"'
    echo 'if [ "$fail" = 0 ]; then echo "All conversions finished."; else echo "Finished, but one or more files had errors (see above)."; fi'
    echo 'echo "You can close this window."'
    # remove the temp script itself after it runs
    printf 'rm -f %q\n' "$CMD"
  } >> "$CMD"
  chmod +x "$CMD"
  open -a Terminal "$CMD"
  exit 0
fi

# =================================================================== DVD (VIDEO_TS)
FOLDER="$(osa -e 'set f to choose folder with prompt "Select your DVD folder or its VIDEO_TS folder:"
return POSIX path of f')"
[ -z "$FOLDER" ] && exit 0
FOLDER="${FOLDER%/}"

# Locate the VIDEO_TS directory.
VTS=""
child="$(find "$FOLDER" -maxdepth 1 -type d -iname 'VIDEO_TS' 2>/dev/null | head -1)"
if [ -n "$child" ]; then
  VTS="$child"
elif printf '%s' "$(basename "$FOLDER")" | grep -qi '^VIDEO_TS$'; then
  VTS="$FOLDER"
else
  child="$(find "$FOLDER" -maxdepth 3 -type d -iname 'VIDEO_TS' 2>/dev/null | head -1)"
  [ -n "$child" ] && VTS="$child" || VTS="$FOLDER"
fi

# Gather content VOB parts (VTS_tt_pp.VOB, pp 1..9; pp 0 = menu, skipped).
WORK="$(mktemp -d /tmp/videomerger.XXXXXX)"
TDIR="$WORK/titles"; mkdir -p "$TDIR"
found=0
while IFS= read -r -d '' f; do
  b="$(basename "$f")"
  u="$(printf '%s' "$b" | tr '[:lower:]' '[:upper:]')"
  tt="${u#VTS_}"; tt="${tt%%_*}"           # title set number
  pp="${u##*_}"; pp="${pp%.VOB}"           # part number
  case "$tt" in ''|*[!0-9]*) continue ;; esac
  case "$pp" in ''|*[!0-9]*) continue ;; esac
  sz="$(stat -f%z "$f" 2>/dev/null || echo 0)"
  printf '%s\t%s\t%s\n' "$pp" "$sz" "$f" >> "$TDIR/$tt"
  found=1
done < <(find "$VTS" -maxdepth 1 -type f -iname 'VTS_*_[1-9].VOB' -print0 2>/dev/null)

if [ "$found" -ne 1 ]; then
  crit "No DVD video (VTS_*.VOB) files were found in:
$VTS

Make sure you picked a ripped DVD folder that contains a VIDEO_TS folder."
  rm -rf "$WORK"; exit 0
fi

# Build a menu of titles: "tt|parts|bytes|GBtext". Track the largest.
CHOICES=""; MAXTT=""; MAXBYTES=-1
for tfile in "$TDIR"/*; do
  tt="$(basename "$tfile")"
  parts="$(wc -l < "$tfile" | tr -d ' ')"
  bytes="$(awk -F'\t' '{s+=$2} END{printf "%.0f", s}' "$tfile")"
  gb="$(awk -v b="$bytes" 'BEGIN{printf "%.2f", b/1073741824}')"
  label="Title $tt  -  $parts part(s)  -  ${gb} GB"
  CHOICES="$CHOICES$label
"
  if [ "$bytes" -gt "$MAXBYTES" ] 2>/dev/null; then MAXBYTES="$bytes"; MAXTT="$tt"; fi
done
CHOICES="${CHOICES%$'\n'}"

# Pick a title (auto if only one).
NUM_TITLES="$(ls -1 "$TDIR" | wc -l | tr -d ' ')"
if [ "$NUM_TITLES" -gt 1 ]; then
  DEFLABEL="$(printf '%s' "$CHOICES" | grep "^Title $MAXTT  " | head -1)"
  SEL="$(osa -e "set c to choose from list (paragraphs of \"$(printf '%s' "$CHOICES" | sed 's/\"/\\\"/g')\") with prompt \"This DVD has more than one title. Choose the one to merge (largest is usually the main feature):\" default items {\"$DEFLABEL\"} without empty selection allowed
if c is false then return \"CANCEL\"
return item 1 of c")"
  [ -z "$SEL" ] && { rm -rf "$WORK"; exit 0; }
  [ "$SEL" = "CANCEL" ] && { rm -rf "$WORK"; exit 0; }
  CHTT="$(printf '%s' "$SEL" | sed -E 's/^Title ([0-9]+).*/\1/')"
else
  CHTT="$MAXTT"
fi
TFILE="$TDIR/$CHTT"
[ -f "$TFILE" ] || { crit "Could not read the selected title."; rm -rf "$WORK"; exit 0; }

# Order this title's parts by part number and build the concat argument.
JOINED=""; PARTS=()
while IFS= read -r line; do
  f="${line#*$'\t'*$'\t'}"          # third field (path)
  PARTS+=("$f")
  if [ -z "$JOINED" ]; then JOINED="$f"; else JOINED="$JOINED|$f"; fi
done < <(sort -t"$(printf '\t')" -k1,1n "$TFILE")

# Output file.
OUT="$(osa -e 'set o to choose file name with prompt "Save the merged DVD video as:" default name "dvd_movie.mpg"
set p to POSIX path of o
if p does not end with ".mpg" and p does not end with ".mpeg" and p does not end with ".vob" then set p to p & ".mpg"
return p')"
[ -z "$OUT" ] && { rm -rf "$WORK"; exit 0; }

note "Joining DVD title $CHTT (${#PARTS[@]} part(s))..."
ERR="$WORK/err.txt"
# Lossless: ffmpeg concat protocol, stream copy. Fall back to byte-exact cat.
if "$FFMPEG" -y -hide_banner -loglevel error -i "concat:$JOINED" -c copy "$OUT" 2>"$ERR"; then
  done_dialog "$OUT"; rm -rf "$WORK"; exit 0
fi
if cat "${PARTS[@]}" > "$OUT" 2>>"$ERR"; then
  done_dialog "$OUT"; rm -rf "$WORK"; exit 0
fi
MSG="$(tail -n 6 "$ERR" 2>/dev/null | tr '"' "'")"
crit "Could not join the DVD title:
$MSG"
rm -rf "$WORK"; exit 1
