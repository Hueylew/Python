import Foundation

// ── Files the app will work on ───────────────────────────────────────────────

/// Extensions the merge list accepts. Merging by stream copy really only works
/// for the MP4 family, which is why the original only offered those three — but
/// anything ffmpeg reads can be dropped, since converting accepts far more and
/// a mismatched merge falls back to re-encoding anyway.
let mergeableExtensions: Set<String> = ["mp4", "m4v", "mov"]
let convertibleExtensions: Set<String> = [
    "mp4", "m4v", "mov", "m2ts", "mts", "ts", "mkv", "mpg", "mpeg", "vob", "avi", "webm",
]

struct MediaItem: Identifiable, Hashable {
    let id = UUID()
    let url: URL
    var duration: Double = 0      // seconds; 0 when ffprobe couldn't say
    var bytes: Int64 = 0

    var name: String { url.lastPathComponent }
    var ext: String { url.pathExtension.lowercased() }

    static func == (a: MediaItem, b: MediaItem) -> Bool { a.id == b.id }
    func hash(into h: inout Hasher) { h.combine(id) }
}

/// Read a file's duration and size. Duration drives the progress bar, so a file
/// ffprobe can't read still gets added — it just contributes nothing to the
/// total, and the bar falls back to indeterminate if nothing is known at all.
func inspect(_ url: URL) -> MediaItem {
    var item = MediaItem(url: url)
    item.bytes = (try? url.resourceValues(forKeys: [.fileSizeKey]))
        .flatMap { $0.fileSize }.map(Int64.init) ?? 0
    guard let ffprobe = Tools.ffprobe else { return item }
    let text = runCapturing(ffprobe, [
        "-v", "error", "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1", url.path,
    ])
    item.duration = Double(text.trimmingCharacters(in: .whitespacesAndNewlines)) ?? 0
    return item
}

// ── DVD folders ──────────────────────────────────────────────────────────────

struct DVDTitle: Identifiable, Hashable {
    let id = UUID()
    let number: Int
    let parts: [URL]
    let bytes: Int64

    var label: String {
        "Title \(number) — \(parts.count) part\(parts.count == 1 ? "" : "s") — \(humanSize(bytes))"
    }

    static func == (a: DVDTitle, b: DVDTitle) -> Bool { a.id == b.id }
    func hash(into h: inout Hasher) { h.combine(id) }
}

/// Find the VIDEO_TS directory for a folder someone picked or dropped: the
/// folder itself, a child, or a grandchild — people drop the disc folder, the
/// VIDEO_TS inside it, or a folder of rips, and all three should work.
func findVideoTS(in folder: URL) -> URL? {
    let fm = FileManager.default
    if folder.lastPathComponent.uppercased() == "VIDEO_TS" { return folder }

    func directories(of dir: URL) -> [URL] {
        (try? fm.contentsOfDirectory(at: dir, includingPropertiesForKeys: [.isDirectoryKey],
                                     options: [.skipsHiddenFiles])) ?? []
    }
    var frontier = [folder]
    for _ in 0..<3 {
        var next: [URL] = []
        for dir in frontier {
            for child in directories(of: dir) {
                guard (try? child.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory == true
                else { continue }
                if child.lastPathComponent.uppercased() == "VIDEO_TS" { return child }
                next.append(child)
            }
        }
        frontier = next
        if frontier.isEmpty { break }
    }
    // A folder of loose VOBs with no VIDEO_TS wrapper still merges fine.
    return hasVOBs(folder) ? folder : nil
}

private func hasVOBs(_ dir: URL) -> Bool {
    let files = (try? FileManager.default.contentsOfDirectory(atPath: dir.path)) ?? []
    return files.contains { $0.uppercased().hasSuffix(".VOB") }
}

/// Group a VIDEO_TS folder's content VOBs into titles.
///
/// `VTS_<title>_<part>.VOB` with part 1-9 is content; part 0 is the menu, so it
/// is skipped — exactly what the shell version did. Parts come back in play
/// order, and the titles themselves largest-first, since the biggest title is
/// almost always the main feature.
func dvdTitles(in videoTS: URL) -> [DVDTitle] {
    let fm = FileManager.default
    let names = (try? fm.contentsOfDirectory(atPath: videoTS.path)) ?? []
    var groups: [Int: [(part: Int, url: URL, bytes: Int64)]] = [:]

    for name in names {
        let upper = name.uppercased()
        guard upper.hasPrefix("VTS_"), upper.hasSuffix(".VOB") else { continue }
        let stem = String(upper.dropFirst(4).dropLast(4))       // "01_2"
        let fields = stem.split(separator: "_")
        guard fields.count == 2,
              let title = Int(fields[0]), let part = Int(fields[1]), (1...9).contains(part)
        else { continue }
        let url = videoTS.appendingPathComponent(name)
        let bytes = (try? url.resourceValues(forKeys: [.fileSizeKey]))
            .flatMap { $0.fileSize }.map(Int64.init) ?? 0
        groups[title, default: []].append((part, url, bytes))
    }

    return groups.map { number, parts in
        let ordered = parts.sorted { $0.part < $1.part }
        return DVDTitle(number: number,
                        parts: ordered.map(\.url),
                        bytes: ordered.reduce(0) { $0 + $1.bytes })
    }
    .sorted { $0.bytes > $1.bytes }
}

// ── Formatting ───────────────────────────────────────────────────────────────

func humanSize(_ bytes: Int64) -> String {
    guard bytes > 0 else { return "—" }
    let units = ["B", "KB", "MB", "GB", "TB"]
    var value = Double(bytes), index = 0
    while value >= 1024, index < units.count - 1 { value /= 1024; index += 1 }
    return String(format: index == 0 ? "%.0f %@" : "%.1f %@", value, units[index])
}

/// Throughput, e.g. "38.2 MB/s". Network shares are the reason this is worth
/// showing: it is the difference between "slow" and "stalled".
func humanRate(_ bytesPerSecond: Double) -> String {
    guard bytesPerSecond > 1 else { return "" }
    return humanSize(Int64(bytesPerSecond)) + "/s"
}

func humanDuration(_ seconds: Double) -> String {
    guard seconds > 0, seconds.isFinite else { return "—" }
    let total = Int(seconds.rounded())
    let h = total / 3600, m = (total % 3600) / 60, s = total % 60
    return h > 0 ? String(format: "%d:%02d:%02d", h, m, s)
                 : String(format: "%d:%02d", m, s)
}
