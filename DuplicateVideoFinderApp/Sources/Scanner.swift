import CryptoKit
import Foundation

// ── Tuning ───────────────────────────────────────────────────────────────────
let videoExtensions: Set<String> = [
    "mp4", "mkv", "avi", "mov", "wmv", "flv", "webm",
    "m4v", "mpg", "mpeg", "ts", "m2ts",
]

let durationToleranceSec = 2.0
let samplePositions = [0.25, 0.5, 0.75]
let maxHammingPerFrame = 10  // out of 64 bits; higher = looser match

// Reading big files is the bottleneck, so past ~4 concurrent readers a typical
// external drive just thrashes. Measured 1→23.5s, 2→16.5s, 4→14.9s, 8→15.2s.
let maxWorkers = min(4, ProcessInfo.processInfo.activeProcessorCount)

enum Tools {
    static let ffmpeg = findTool("ffmpeg")
    static let ffprobe = findTool("ffprobe")
    static var missing: [String] {
        var m: [String] = []
        if ffmpeg == nil { m.append("ffmpeg") }
        if ffprobe == nil { m.append("ffprobe") }
        return m
    }
}

// ── Model ────────────────────────────────────────────────────────────────────
struct VideoFile: Identifiable, Hashable {
    let id = UUID()
    let url: URL
    let size: Int64
    var duration: Double?
    var width: Int?
    var height: Int?
    var codec: String?

    var name: String { url.lastPathComponent }
    var folder: String { url.deletingLastPathComponent().path }
    var pixels: Int { (width ?? 0) * (height ?? 0) }

    var bitrate: Double? {
        guard let d = duration, d > 0 else { return nil }
        return Double(size) * 8 / d
    }

    var resolution: String {
        guard let w = width, let h = height, w > 0, h > 0 else { return "?" }
        return "\(w)x\(h)"
    }

    /// Bitrate in ~1.4% buckets. Bucketing rather than comparing raw doubles
    /// keeps the ordering transitive (required by sort) while stopping a
    /// meaningless fraction-of-a-percent difference from outranking a clearly
    /// better filename.
    var bitrateBucket: Int {
        guard let b = bitrate, b > 0 else { return 0 }
        return Int((log2(b) * 50).rounded())
    }

    /// Rough "would a human recognise this name?" score. A downloader's UUID or
    /// hash scores 0; a name built from real words scores higher. Only ever used
    /// to break ties between copies of equal picture quality.
    var nameScore: Int {
        let stem = url.deletingPathExtension().lastPathComponent
        if UUID(uuidString: stem) != nil { return 0 }

        var score = 0
        for token in stem.split(whereSeparator: { !$0.isLetter && !$0.isNumber }) {
            // a long hex run is a hash, not a word — skip it whole, or its
            // accidental "bee"/"aff" fragments would score like real words
            if token.count >= 12, token.allSatisfy(\.isHexDigit) { continue }

            // score the letter runs inside, so "pantyhose18963" still counts
            for word in token.split(whereSeparator: { !$0.isLetter }) {
                guard word.count >= 3 else { continue }
                // vowels are what rule out tokens like "XWPXAHDM"
                let vowels = word.filter { "aeiouAEIOU".contains($0) }.count
                guard Double(vowels) / Double(word.count) >= 0.2 else { continue }
                score += word.count
            }
        }
        return score
    }

    /// Rank by how much picture information a copy retains. Resolution
    /// dominates; then bits per second, since that's real compression damage.
    /// Only once those tie does the filename decide — a name is trivially
    /// fixable, lost picture detail is not.
    var qualityKey: (Int, Int, Int, Int64) { (pixels, bitrateBucket, nameScore, size) }
}

func betterQuality(_ a: VideoFile, _ b: VideoFile) -> Bool {
    let x = a.qualityKey, y = b.qualityKey
    if x.0 != y.0 { return x.0 > y.0 }
    if x.1 != y.1 { return x.1 > y.1 }
    if x.2 != y.2 { return x.2 > y.2 }
    return x.3 > y.3
}

struct DuplicateGroup: Identifiable {
    let id = UUID()
    let kind: Kind
    var files: [VideoFile]   // sorted best-first

    enum Kind { case exact, possible
        var label: String { self == .exact ? "Exact duplicate" : "Possible duplicate" }
    }

    var bestCopy: VideoFile { files[0] }

    var keepReason: String {
        let best = bestCopy
        let others = files.dropFirst()
        if others.contains(where: { best.pixels > $0.pixels }) { return "highest resolution" }
        if others.contains(where: { best.bitrateBucket > $0.bitrateBucket }) { return "highest bitrate" }
        if others.contains(where: { best.nameScore > $0.nameScore }) { return "clearest filename" }
        return "largest file"
    }

    var wastedBytes: Int64 { files.dropFirst().reduce(0) { $0 + $1.size } }
}

// ── Cancellation ─────────────────────────────────────────────────────────────
final class CancelToken: @unchecked Sendable {
    private let lock = NSLock()
    private var flag = false
    var isCancelled: Bool {
        lock.lock(); defer { lock.unlock() }
        return flag
    }
    func cancel() { lock.lock(); flag = true; lock.unlock() }
}

// ── Concurrency helper ───────────────────────────────────────────────────────
/// Map `items` in parallel, at most `limit` at a time, reporting progress as
/// each finishes. Results keep the input order.
func concurrentMap<T, R>(_ items: [T], limit: Int, cancel: CancelToken,
                         onProgress: @escaping (Int, Int, T) -> Void,
                         transform: @escaping (T) -> R) -> [R?] {
    var results = [R?](repeating: nil, count: items.count)
    let lock = NSLock()
    var done = 0
    let sem = DispatchSemaphore(value: limit)
    let group = DispatchGroup()

    for (i, item) in items.enumerated() {
        if cancel.isCancelled { break }
        sem.wait()
        group.enter()
        DispatchQueue.global(qos: .userInitiated).async {
            defer { sem.signal(); group.leave() }
            if cancel.isCancelled { return }
            let r = transform(item)
            lock.lock()
            results[i] = r
            done += 1
            let d = done
            lock.unlock()
            onProgress(d, items.count, item)
        }
    }
    group.wait()
    return results
}

// ── File discovery ───────────────────────────────────────────────────────────
func findVideoFiles(in folders: [URL]) -> [URL] {
    let fm = FileManager.default
    var seen = Set<String>()
    var results: [URL] = []

    for folder in folders {
        guard let e = fm.enumerator(at: folder,
                                    includingPropertiesForKeys: [.isRegularFileKey],
                                    options: [.skipsHiddenFiles]) else { continue }
        for case let url as URL in e {
            let name = url.lastPathComponent
            if name.hasPrefix("._") { continue }  // macOS AppleDouble sidecar
            guard videoExtensions.contains(url.pathExtension.lowercased()) else { continue }
            let values = try? url.resourceValues(forKeys: [.isRegularFileKey])
            guard values?.isRegularFile == true else { continue }
            let key = (try? url.resourceValues(forKeys: [.canonicalPathKey]).canonicalPath) ?? url.path
            if seen.insert(key).inserted { results.append(url) }
        }
    }
    return results
}

// ── Metadata ─────────────────────────────────────────────────────────────────
func buildVideoFile(_ url: URL) -> VideoFile? {
    guard let attrs = try? FileManager.default.attributesOfItem(atPath: url.path),
          let size = attrs[.size] as? Int64 else { return nil }
    var vf = VideoFile(url: url, size: size)

    guard let ffprobe = Tools.ffprobe else { return vf }
    let r = runProcess(ffprobe, [
        "-v", "quiet", "-print_format", "json",
        "-show_format", "-show_streams", "-select_streams", "v:0", url.path,
    ], timeout: 30)
    guard r.status == 0,
          let json = try? JSONSerialization.jsonObject(with: r.out) as? [String: Any]
    else { return vf }

    let streams = json["streams"] as? [[String: Any]] ?? []
    let format = json["format"] as? [String: Any] ?? [:]

    if let d = format["duration"] as? String, let v = Double(d) { vf.duration = v }
    else if let d = streams.first?["duration"] as? String, let v = Double(d) { vf.duration = v }

    if let s = streams.first {
        vf.width = s["width"] as? Int
        vf.height = s["height"] as? Int
        vf.codec = s["codec_name"] as? String
    }
    return vf
}

// ── Exact hashing ────────────────────────────────────────────────────────────
func hashFileContents(_ url: URL) -> String? {
    guard let handle = try? FileHandle(forReadingFrom: url) else { return nil }
    defer { try? handle.close() }
    var hasher = SHA256()
    while let chunk = try? handle.read(upToCount: 1 << 20), !chunk.isEmpty {
        hasher.update(data: chunk)
    }
    return hasher.finalize().map { String(format: "%02x", $0) }.joined()
}

// ── Perceptual hashing ───────────────────────────────────────────────────────
/// Ask ffmpeg for one frame already scaled to 9x8 greyscale — that's exactly the
/// 72 bytes a difference-hash needs, so no image decoding is required.
func frameHash(_ url: URL, at seconds: Double) -> UInt64? {
    guard let ffmpeg = Tools.ffmpeg else { return nil }
    let r = runProcess(ffmpeg, [
        "-v", "quiet", "-ss", String(seconds), "-i", url.path,
        "-frames:v", "1", "-vf", "scale=9:8", "-pix_fmt", "gray",
        "-f", "rawvideo", "-",
    ], timeout: 30)
    guard r.status == 0, r.out.count >= 72 else { return nil }

    let px = [UInt8](r.out.prefix(72))
    var bits: UInt64 = 0
    for row in 0..<8 {
        for col in 0..<8 {
            let i = row * 9 + col
            bits = (bits << 1) | (px[i] > px[i + 1] ? 1 : 0)
        }
    }
    return bits
}

func computeSignature(_ vf: VideoFile) -> [UInt64]? {
    guard let duration = vf.duration, duration > 0 else { return nil }
    var hashes: [UInt64] = []
    for frac in samplePositions {
        guard let h = frameHash(vf.url, at: max(0, duration * frac)) else { return nil }
        hashes.append(h)
    }
    return hashes
}

func hamming(_ a: [UInt64], _ b: [UInt64]) -> Int {
    zip(a, b).reduce(0) { $0 + ($1.0 ^ $1.1).nonzeroBitCount }
}

// ── Grouping ─────────────────────────────────────────────────────────────────
/// Files that are byte-identical to each other (or a lone file with no exact
/// match). This is the unit of perceptual comparison: identical bytes always
/// produce an identical frame hash, so one ffmpeg pass covers the whole cluster.
final class ByteCluster {
    var files: [VideoFile]
    var signature: [UInt64]?
    init(files: [VideoFile]) { self.files = files }
    var duration: Double? { files.first?.duration }
}

struct UnionFind {
    private var parent: [Int]
    init(_ n: Int) { parent = Array(0..<n) }
    mutating func find(_ x: Int) -> Int {
        var x = x
        while parent[x] != x { parent[x] = parent[parent[x]]; x = parent[x] }
        return x
    }
    mutating func union(_ a: Int, _ b: Int) {
        let ra = find(a), rb = find(b)
        if ra != rb { parent[ra] = rb }
    }
}

func groupByExactHash(_ files: [VideoFile], cancel: CancelToken,
                      progress: @escaping (Int, Int, String) -> Void) -> [ByteCluster] {
    var bySize: [Int64: [VideoFile]] = [:]
    for f in files { bySize[f.size, default: []].append(f) }

    // only files sharing a size can possibly be byte-identical
    let needHash = bySize.values.filter { $0.count > 1 }.flatMap { $0 }
    var digests: [UUID: String] = [:]
    if !needHash.isEmpty {
        let hashed = concurrentMap(needHash, limit: maxWorkers, cancel: cancel,
                                   onProgress: { d, t, f in progress(d, t, "Hashing \(f.name)") },
                                   transform: { hashFileContents($0.url) })
        for (i, h) in hashed.enumerated() {
            if let h = h ?? nil { digests[needHash[i].id] = h }
        }
    }

    var clusters: [ByteCluster] = []
    for group in bySize.values {
        if group.count == 1 { clusters.append(ByteCluster(files: group)); continue }
        var byHash: [String: [VideoFile]] = [:]
        for f in group {
            if let d = digests[f.id] {
                byHash[d, default: []].append(f)
            } else {
                clusters.append(ByteCluster(files: [f]))  // unreadable — keep it alone
            }
        }
        for members in byHash.values { clusters.append(ByteCluster(files: members)) }
    }
    return clusters
}

func computeClusterSignatures(_ clusters: [ByteCluster], cancel: CancelToken,
                              progress: @escaping (Int, Int, String) -> Void) {
    let dated = clusters.filter { $0.duration != nil }
    guard !dated.isEmpty else { return }
    let sigs = concurrentMap(dated, limit: maxWorkers, cancel: cancel,
                             onProgress: { d, t, c in
                                 progress(d, t, "Sampling frames: \(c.files[0].name)")
                             },
                             transform: { computeSignature($0.files[0]) })
    for (i, s) in sigs.enumerated() { dated[i].signature = s ?? nil }
}

/// Merge byte-clusters whose frames look alike, so a file that's an exact
/// duplicate of A and another that's merely a re-encode of A land in the same
/// group rather than the re-encode being dropped.
func mergeBySimilarity(_ clusters: [ByteCluster]) -> [DuplicateGroup] {
    let signed = clusters.filter { $0.signature != nil }
        .sorted { ($0.duration ?? 0) < ($1.duration ?? 0) }
    var uf = UnionFind(signed.count)

    for i in 0..<signed.count {
        for j in (i + 1)..<signed.count {
            let di = signed[i].duration ?? 0, dj = signed[j].duration ?? 0
            if dj - di > durationToleranceSec { break }  // sorted — nothing further is in range
            guard let a = signed[i].signature, let b = signed[j].signature else { continue }
            if Double(hamming(a, b)) / Double(a.count) <= Double(maxHammingPerFrame) {
                uf.union(i, j)
            }
        }
    }

    var byRoot: [Int: [Int]] = [:]
    for i in 0..<signed.count { byRoot[uf.find(i), default: []].append(i) }

    var groups: [DuplicateGroup] = []
    for indices in byRoot.values {
        let here = indices.map { signed[$0] }
        var files = here.flatMap { $0.files }
        guard files.count > 1 else { continue }
        files.sort(by: betterQuality)
        // "exact" only when a single byte-identical cluster is involved
        groups.append(DuplicateGroup(kind: here.count == 1 ? .exact : .possible, files: files))
    }

    // clusters we couldn't sample can still be reported if byte-identical
    for c in clusters where c.signature == nil && c.files.count > 1 {
        groups.append(DuplicateGroup(kind: .exact, files: c.files.sorted(by: betterQuality)))
    }
    return groups
}

// ── Entry point ──────────────────────────────────────────────────────────────
func scan(folders: [URL], cancel: CancelToken,
          progress: @escaping (Int, Int, String) -> Void) -> [DuplicateGroup] {
    progress(0, 0, "Step 1/4 — Finding video files…")
    let paths = findVideoFiles(in: folders)
    if cancel.isCancelled { return [] }

    var files: [VideoFile] = []
    if !paths.isEmpty {
        let built = concurrentMap(paths, limit: maxWorkers, cancel: cancel,
                                  onProgress: { d, t, u in
                                      progress(d, t, "Step 2/4 — Reading metadata: \(u.lastPathComponent)")
                                  },
                                  transform: { buildVideoFile($0) })
        files = built.compactMap { $0 ?? nil }
    }
    if cancel.isCancelled { return [] }
    files.sort { $0.url.path < $1.url.path }  // threads finish out of order

    let clusters = groupByExactHash(files, cancel: cancel) { d, t, m in
        progress(d, t, "Step 3/4 — Checking exact duplicates: \(m)")
    }
    if cancel.isCancelled { return [] }

    computeClusterSignatures(clusters, cancel: cancel) { d, t, m in
        progress(d, t, "Step 4/4 — Checking for re-encoded duplicates: \(m)")
    }
    if cancel.isCancelled { return [] }

    var groups = mergeBySimilarity(clusters)
    groups.sort { $0.wastedBytes > $1.wastedBytes }  // biggest wins first
    return groups
}

// ── Formatting ───────────────────────────────────────────────────────────────
func humanSize(_ bytes: Int64) -> String {
    var v = Double(bytes)
    for unit in ["B", "KB", "MB", "GB", "TB"] {
        if v < 1024 { return String(format: "%.1f %@", v, unit) }
        v /= 1024
    }
    return String(format: "%.1f PB", v)
}

func humanBitrate(_ bps: Double?) -> String {
    guard let b = bps, b > 0 else { return "?" }
    if b >= 1_000_000 { return String(format: "%.1f Mbps", b / 1_000_000) }
    return String(format: "%.0f kbps", b / 1_000)
}

func humanDuration(_ seconds: Double?) -> String {
    guard let s = seconds, s > 0 else { return "?" }
    let t = Int(s)
    let (h, m, sec) = (t / 3600, (t % 3600) / 60, t % 60)
    return h > 0 ? String(format: "%d:%02d:%02d", h, m, sec)
                 : String(format: "%d:%02d", m, sec)
}
