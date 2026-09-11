"""Core duplicate-video detection logic (no GUI dependencies).

Two tiers of duplicate detection:
  - Exact: files whose content hashes match byte-for-byte.
  - Possible: files with similar duration whose sampled frames look alike
    (perceptual hash), catching re-encodes/transcodes that exact hashing misses.
"""
from __future__ import annotations

import hashlib
import io
import os
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

from PIL import Image

VIDEO_EXTENSIONS = {
    ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm",
    ".m4v", ".mpg", ".mpeg", ".ts", ".m2ts",
}

# Perceptual-hash comparison tuning.
DURATION_TOLERANCE_SEC = 2.0
SAMPLE_POSITIONS = (0.25, 0.5, 0.75)
HASH_SIZE = 8
MAX_HAMMING_PER_FRAME = 10  # out of 64 bits; higher = looser match

# ffprobe/ffmpeg/hashing all block on a subprocess or on disk, so threads
# (not processes) are enough to overlap them. Measured on a USB video library:
# 1→23.5s, 2→16.5s, 4→14.9s, 8→15.2s — reading big files is the bottleneck, so
# past ~4 concurrent readers the drive just thrashes. Raise it if your media
# lives on a fast internal SSD.
MAX_WORKERS = min(4, (os.cpu_count() or 4))


@dataclass
class VideoFile:
    path: Path
    size: int
    duration: Optional[float] = None
    width: Optional[int] = None
    height: Optional[int] = None
    codec: Optional[str] = None
    mtime: float = 0.0

    @property
    def resolution(self) -> str:
        if self.width and self.height:
            return f"{self.width}x{self.height}"
        return "?"

    @property
    def pixels(self) -> int:
        return (self.width or 0) * (self.height or 0)

    @property
    def bitrate_bps(self) -> Optional[float]:
        if not self.duration or self.duration <= 0:
            return None
        return self.size * 8 / self.duration


def quality_key(vf: VideoFile) -> tuple:
    """Rank a copy by how much picture information it retains. Resolution
    dominates; at equal resolution more bits per second means less compression
    damage; file size is the final tiebreak."""
    return (vf.pixels, vf.bitrate_bps or 0.0, vf.size)


@dataclass
class DuplicateGroup:
    kind: str  # "exact" or "possible"
    files: list[VideoFile] = field(default_factory=list)

    @property
    def best_copy(self) -> VideoFile:
        """The copy worth keeping — highest quality, not merely largest."""
        return max(self.files, key=quality_key)

    @property
    def keep_reason(self) -> str:
        best = self.best_copy
        others = [f for f in self.files if f is not best]
        if not others:
            return ""
        if any(best.pixels > f.pixels for f in others):
            return "highest resolution"
        if any((best.bitrate_bps or 0) > (f.bitrate_bps or 0) for f in others):
            return "highest bitrate"
        return "largest file"

    @property
    def wasted_bytes(self) -> int:
        if len(self.files) < 2:
            return 0
        best = self.best_copy
        return sum(f.size for f in self.files if f is not best)


ProgressCallback = Callable[[int, int, str], None]


def find_video_files(folders: list[Path]) -> list[Path]:
    seen: set[Path] = set()
    results: list[Path] = []
    for folder in folders:
        if not folder.is_dir():
            continue
        for path in folder.rglob("*"):
            if path.name.startswith("._"):
                continue  # macOS AppleDouble sidecar, not real video data
            if path.is_file() and path.suffix.lower() in VIDEO_EXTENSIONS:
                resolved = path.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    results.append(path)
    return results


def probe_video(path: Path) -> tuple[Optional[float], Optional[int], Optional[int], Optional[str]]:
    try:
        result = subprocess.run(
            [
                "ffprobe", "-v", "quiet",
                "-print_format", "json",
                "-show_format", "-show_streams",
                "-select_streams", "v:0",
                str(path),
            ],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode != 0:
            return None, None, None, None
        import json
        data = json.loads(result.stdout)
        streams = data.get("streams", [])
        fmt = data.get("format", {})
        duration = None
        if fmt.get("duration"):
            duration = float(fmt["duration"])
        elif streams and streams[0].get("duration"):
            duration = float(streams[0]["duration"])
        width = streams[0].get("width") if streams else None
        height = streams[0].get("height") if streams else None
        codec = streams[0].get("codec_name") if streams else None
        return duration, width, height, codec
    except (subprocess.TimeoutExpired, ValueError, json.JSONDecodeError, OSError):
        return None, None, None, None


def build_video_file(path: Path) -> VideoFile:
    stat = path.stat()
    duration, width, height, codec = probe_video(path)
    return VideoFile(
        path=path, size=stat.st_size, duration=duration,
        width=width, height=height, codec=codec, mtime=stat.st_mtime,
    )


def _safe_build(path: Path) -> Optional[VideoFile]:
    try:
        return build_video_file(path)
    except OSError:
        return None


def hash_file_contents(path: Path, chunk_size: int = 1024 * 1024) -> str:
    hasher = hashlib.sha256()
    with open(path, "rb") as f:
        while chunk := f.read(chunk_size):
            hasher.update(chunk)
    return hasher.hexdigest()


class _UnionFind:
    def __init__(self, n: int):
        self.parent = list(range(n))

    def find(self, x: int) -> int:
        while self.parent[x] != x:
            self.parent[x] = self.parent[self.parent[x]]
            x = self.parent[x]
        return x

    def union(self, a: int, b: int) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[ra] = rb


@dataclass
class _ByteCluster:
    """Files that are byte-identical to each other (or a lone file with no
    exact match). Used as the unit of perceptual comparison, since identical
    bytes always produce an identical perceptual hash — no need to re-sample
    frames for every member of an exact-duplicate group."""
    files: list[VideoFile]
    signature: Optional[list[int]] = None

    @property
    def duration(self) -> Optional[float]:
        return self.files[0].duration


def _safe_hash(path: Path) -> Optional[str]:
    try:
        return hash_file_contents(path)
    except OSError:
        return None


def _group_by_exact_hash(
    files: list[VideoFile], progress: Optional[ProgressCallback] = None
) -> list[_ByteCluster]:
    by_size: dict[int, list[VideoFile]] = {}
    for vf in files:
        by_size.setdefault(vf.size, []).append(vf)

    # only files sharing a size can possibly be byte-identical
    need_hash = [vf for group in by_size.values() if len(group) > 1 for vf in group]
    digests: dict[Path, Optional[str]] = {}
    if need_hash:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = {pool.submit(_safe_hash, vf.path): vf for vf in need_hash}
            for done, future in enumerate(as_completed(futures), 1):
                vf = futures[future]
                digests[vf.path] = future.result()
                if progress:
                    progress(done, len(need_hash), f"Hashing {vf.path.name}")

    clusters: list[_ByteCluster] = []
    for group in by_size.values():
        if len(group) == 1:
            clusters.append(_ByteCluster(files=group))
            continue
        by_hash: dict[str, list[VideoFile]] = {}
        for vf in group:
            digest = digests.get(vf.path)
            if digest is None:  # unreadable — keep it, but on its own
                clusters.append(_ByteCluster(files=[vf]))
            else:
                by_hash.setdefault(digest, []).append(vf)
        for members in by_hash.values():
            clusters.append(_ByteCluster(files=members))

    return clusters


def _dhash(image: Image.Image, hash_size: int = HASH_SIZE) -> int:
    image = image.convert("L").resize((hash_size + 1, hash_size), Image.LANCZOS)
    pixels = list(image.getdata())
    bits = 0
    for row in range(hash_size):
        for col in range(hash_size):
            idx = row * (hash_size + 1) + col
            bits = (bits << 1) | int(pixels[idx] > pixels[idx + 1])
    return bits


def _hamming(a: int, b: int) -> int:
    return bin(a ^ b).count("1")


def extract_frame_hash(path: Path, timestamp: float) -> Optional[int]:
    try:
        result = subprocess.run(
            [
                "ffmpeg", "-v", "quiet", "-ss", str(timestamp), "-i", str(path),
                "-frames:v", "1", "-f", "image2", "-vcodec", "png",
                "-vf", "scale=64:64", "pipe:1",
            ],
            capture_output=True, timeout=30,
        )
        if result.returncode != 0 or not result.stdout:
            return None
        image = Image.open(io.BytesIO(result.stdout))
        return _dhash(image)
    except (subprocess.TimeoutExpired, OSError):
        return None


def compute_signature(vf: VideoFile) -> Optional[list[int]]:
    if not vf.duration or vf.duration <= 0:
        return None
    hashes = []
    for frac in SAMPLE_POSITIONS:
        ts = max(0.0, vf.duration * frac)
        h = extract_frame_hash(vf.path, ts)
        if h is None:
            return None
        hashes.append(h)
    return hashes


def _compute_cluster_signatures(
    clusters: list[_ByteCluster], progress: Optional[ProgressCallback] = None
) -> None:
    """Sample frames for one representative file per cluster. Byte-identical
    files always share the same frames, so a whole exact-duplicate cluster
    only needs one ffmpeg pass, not one per member."""
    dated = [c for c in clusters if c.duration]
    if not dated:
        return
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futures = {pool.submit(compute_signature, c.files[0]): c for c in dated}
        for done, future in enumerate(as_completed(futures), 1):
            cluster = futures[future]
            cluster.signature = future.result()
            if progress:
                progress(done, len(dated), f"Sampling frames: {cluster.files[0].path.name}")


def _merge_by_similarity(clusters: list[_ByteCluster]) -> list[DuplicateGroup]:
    """Merge byte-clusters whose representative frames look alike, so a file
    that's an exact duplicate of A and another that's merely a re-encode of A
    end up in the same reported group instead of the re-encode being dropped."""
    signed = sorted((c for c in clusters if c.signature), key=lambda c: c.duration)
    n = len(signed)
    uf = _UnionFind(n)

    for i in range(n):
        for j in range(i + 1, n):
            if signed[j].duration - signed[i].duration > DURATION_TOLERANCE_SEC:
                break  # sorted by duration — nothing further can be in range
            dist = sum(_hamming(x, y) for x, y in zip(signed[i].signature, signed[j].signature))
            avg = dist / len(signed[i].signature)
            if avg <= MAX_HAMMING_PER_FRAME:
                uf.union(i, j)

    members_by_root: dict[int, list[int]] = {}
    for i in range(n):
        members_by_root.setdefault(uf.find(i), []).append(i)

    groups: list[DuplicateGroup] = []
    for indices in members_by_root.values():
        clusters_here = [signed[i] for i in indices]
        files = [vf for c in clusters_here for vf in c.files]
        if len(files) < 2:
            continue  # a lone file with no exact or perceptual match
        # "exact" only when nothing but a single byte-identical group is involved
        kind = "exact" if len(clusters_here) == 1 else "possible"
        files.sort(key=quality_key, reverse=True)  # best copy first
        groups.append(DuplicateGroup(kind=kind, files=files))

    # clusters with no duration/signature (ffprobe/ffmpeg failed) can still
    # be reported if they were themselves an exact-duplicate group
    for c in clusters:
        if c.signature is None and len(c.files) > 1:
            groups.append(DuplicateGroup(kind="exact", files=c.files))

    return groups


def scan(
    folders: list[Path], progress: Optional[ProgressCallback] = None
) -> list[DuplicateGroup]:
    def stage(label: str):
        def cb(done, total, message):
            if progress:
                progress(done, total, f"{label}: {message}")
        return cb

    if progress:
        progress(0, 0, "Step 1/4 — Finding video files...")
    paths = find_video_files(folders)

    files: list[VideoFile] = []
    if paths:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futures = {pool.submit(_safe_build, p): p for p in paths}
            for done, future in enumerate(as_completed(futures), 1):
                vf = future.result()
                if vf is not None:
                    files.append(vf)
                if progress:
                    name = futures[future].name
                    progress(done, len(paths), f"Step 2/4 — Reading metadata: {name}")
    # threads finish out of order; sort so results are reproducible run to run
    files.sort(key=lambda f: f.path)

    clusters = _group_by_exact_hash(files, stage("Step 3/4 — Checking exact duplicates"))
    _compute_cluster_signatures(clusters, stage("Step 4/4 — Checking for re-encoded duplicates"))

    groups = _merge_by_similarity(clusters)
    groups.sort(key=lambda g: g.wasted_bytes, reverse=True)  # biggest wins first
    return groups
