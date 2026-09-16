import CryptoKit
import Foundation

// ── Tuning ───────────────────────────────────────────────────────────────────
let videoExtensions: Set<String> = [
    "mp4", "mkv", "avi", "mov", "wmv", "flv", "webm",
    "m4v", "mpg", "mpeg", "ts", "m2ts",
]

let durationToleranceSec = 2.0
let samplePositions = [0.25, 0.5, 0.75]

/// A frame is hashed as a 17x16 grey thumbnail compared left-to-right, giving
/// 256 bits. The obvious 9x8 / 64-bit version is too coarse to be safe: on a
/// folder of different clips from one camera session, distinct videos scored
/// 4-23 while genuine re-encodes score 0-5 — the two populations overlap, so no
/// threshold separates them. At 256 bits the same pairs scored 43-95 against 0
/// for a true duplicate, which is a gap with nothing in it. It costs nothing:
/// the same single ffmpeg call, returning 272 bytes instead of 72.
let hashWordsPerFrame = 4                  // 4 x UInt64 = 256 bits
let maxHammingPerFrame = 24                // out of 256; observed false floor was 43

// Two readers, not four, for a local disk.
//
// Measured on two USB drives, and the second one is emphatic. On a library of
// 207 files: 4 workers took 249.9s with 44 timeouts, 2 took 38.2s with 2, and
// 1 took 40.9s with none. The timeouts were never the drive sleeping on its
// own schedule — concurrent readers were thrashing the heads and manufacturing
// the stalls. An earlier drive preferred 4 (14.9s against 16.5s at 2), so the
// cost of choosing 2 there is about a tenth, against a six-fold gain here.
//
// A network share is latency-bound rather than seek-bound — most of its wall
// clock is round trips, not platter movement — so it still wants requests in
// flight. That figure has not been measured as carefully as this one.
func workerCount(for folders: [URL]) -> Int {
    // Escape hatch for tuning against a particular drive — the right number is
    // a property of the hardware, and a spinning disk that thrashes wants far
    // fewer readers than the default assumes.
    if let override = ProcessInfo.processInfo.environment["DVF_WORKERS"],
       let n = Int(override), n > 0 { return n }

    let cores = ProcessInfo.processInfo.activeProcessorCount
    let anyRemote = folders.contains { url in
        ((try? url.resourceValues(forKeys: [.volumeIsLocalKey]))?.volumeIsLocal ?? true) == false
    }
    // Network work is latency-bound so it wants requests in flight, but each
    // worker still decodes a frame, so going past core count just thrashes the
    // machine. One single-threaded decoder per core is the honest ceiling.
    return anyRemote ? cores : min(2, cores)
}

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
    var mtime: Double = 0        // with size, this is the cache key
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
        // scored on the canonical name, so a copy doesn't out-rank the original
        // it was made from just by carrying the extra word "copy"
        let stem = canonicalStem
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

    /// Markers a tool adds when it can't reuse a name. Verified against what
    /// this Mac actually does: Finder's Duplicate gives "clip copy" then
    /// "clip copy 2", Safari appends "-1", browsers use "(1)".
    ///
    /// A bare trailing space-number ("Episode 2", "Part 1 of 3") is deliberately
    /// not treated as a counter — too many real titles end that way.
    private static let copyMarkers = [
        #"\s*\(\d{1,3}\)$"#,                        // "clip (2)"
        #"(?i)\s*[ _-]cop(y|ies)([ _-]?\d{1,3})?$"#, // "clip copy", "clip copy 3"
        #"[-_]\d{1,3}$"#,                           // "clip-1" (Safari), "clip_2"
    ]
    private static let datePrefix = #"^\d{4}[-_.]?\d{2}[-_.]?\d{2}([-_. ]+|$)"#

    var looksDerived: Bool {
        let stem = url.deletingPathExtension().lastPathComponent
            .trimmingCharacters(in: .whitespaces)
        for pattern in Self.copyMarkers
        where stem.range(of: pattern, options: .regularExpression) != nil { return true }
        return stem.range(of: Self.datePrefix, options: .regularExpression) != nil
    }

    /// The name with any copy marker or date prefix taken off, so "clip copy"
    /// reads as well as "clip" rather than scoring higher for containing the
    /// extra word "copy" — which would otherwise rank the copy above the
    /// original it was made from.
    var canonicalStem: String {
        var stem = url.deletingPathExtension().lastPathComponent
            .trimmingCharacters(in: .whitespaces)
        if let r = stem.range(of: Self.datePrefix, options: .regularExpression) {
            stem.removeSubrange(r)
        }
        var changed = true
        while changed {
            changed = false
            for pattern in Self.copyMarkers {
                if let r = stem.range(of: pattern, options: .regularExpression) {
                    stem.removeSubrange(r)
                    changed = true
                }
            }
        }
        return stem.trimmingCharacters(in: .whitespaces)
    }

    /// 1 for a name that looks like the original, 0 for one that looks derived.
    var originality: Int { looksDerived ? 0 : 1 }

    /// Rank by how much picture information a copy retains. Resolution
    /// dominates; then bits per second, since that's real compression damage.
    /// Only once those tie does the name decide — readability first, then
    /// whether it looks like the original rather than a copy of one. A name is
    /// trivially fixable; lost picture detail is not.
    var qualityKey: (Int, Int, Int, Int, Int64) {
        (pixels, bitrateBucket, nameScore, originality, size)
    }
}

func betterQuality(_ a: VideoFile, _ b: VideoFile) -> Bool {
    let x = a.qualityKey, y = b.qualityKey
    if x.0 != y.0 { return x.0 > y.0 }
    if x.1 != y.1 { return x.1 > y.1 }
    if x.2 != y.2 { return x.2 > y.2 }
    if x.3 != y.3 { return x.3 > y.3 }
    return x.4 > y.4
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
        if others.contains(where: { best.originality > $0.originality }) {
            return "the original, not a numbered or dated copy"
        }
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
        // .utility, not .userInitiated: a background scan should lose to
        // whatever the person is actually doing, rather than competing with it
        DispatchQueue.global(qos: .utility).async {
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
    var results: [URL] = []
    var seen = Set<String>()

    // Resolve overlap once per folder rather than per file: asking each file for
    // its canonical path is a fresh filesystem round trip, which on a network
    // share costs more than everything else in this step put together.
    var roots: [URL] = []
    for folder in folders.map({ $0.resolvingSymlinksInPath().standardizedFileURL }) {
        let path = folder.path.hasSuffix("/") ? folder.path : folder.path + "/"
        // drop anything already covered by a folder we're keeping
        if roots.contains(where: { path.hasPrefix($0.path.hasSuffix("/") ? $0.path : $0.path + "/") }) {
            continue
        }
        roots.removeAll { $0.path.hasPrefix(path) }
        roots.append(folder)
    }

    for folder in roots {
        guard let e = fm.enumerator(at: folder,
                                    includingPropertiesForKeys: [.isRegularFileKey, .fileSizeKey],
                                    options: [.skipsHiddenFiles]) else { continue }
        for case let url as URL in e {
            if url.lastPathComponent.hasPrefix("._") { continue }  // AppleDouble sidecar
            guard videoExtensions.contains(url.pathExtension.lowercased()) else { continue }
            // prefetched above, so this costs nothing
            guard (try? url.resourceValues(forKeys: [.isRegularFileKey]))?.isRegularFile == true
            else { continue }
            if seen.insert(url.standardizedFileURL.path).inserted { results.append(url) }
        }
    }
    return results
}

// ── Metadata ─────────────────────────────────────────────────────────────────
func buildVideoFile(_ url: URL, cache: ScanCache? = nil) -> VideoFile? {
    guard let stamp = fileStamp(url) else { return nil }
    var vf = VideoFile(url: url, size: stamp.size, mtime: stamp.mtime)

    // Unchanged since last time? Then ffprobe already told us everything.
    if let e = cache?.entry(for: url.path, size: stamp.size, mtime: stamp.mtime), e.probed {
        vf.duration = e.duration
        vf.width = e.width
        vf.height = e.height
        vf.codec = e.codec
        return vf
    }

    // Nothing to probe with: don't record a verdict we'd be stuck with if
    // ffmpeg gets installed later.
    guard let ffprobe = Tools.ffprobe else { return vf }

    let r = runProcessRetrying(ffprobe, [
        "-v", "quiet", "-threads", "1", "-print_format", "json",
        "-show_format", "-show_streams", "-select_streams", "v:0", url.path,
    ], timeout: 30)

    if r.status == 0,
       let json = try? JSONSerialization.jsonObject(with: r.out) as? [String: Any] {
        let streams = json["streams"] as? [[String: Any]] ?? []
        let format = json["format"] as? [String: Any] ?? [:]

        if let d = format["duration"] as? String, let v = Double(d) { vf.duration = v }
        else if let d = streams.first?["duration"] as? String, let v = Double(d) { vf.duration = v }

        if let s = streams.first {
            vf.width = s["width"] as? Int
            vf.height = s["height"] as? Int
            vf.codec = s["codec_name"] as? String
        }
    }

    // Record the outcome even when ffprobe found nothing, so a file it can't
    // read isn't re-probed on every future scan.
    cache?.storeMetadata(path: url.path, size: stamp.size, mtime: stamp.mtime,
                         duration: vf.duration, width: vf.width,
                         height: vf.height, codec: vf.codec)
    return vf
}

// ── Exact hashing ────────────────────────────────────────────────────────────
/// Hash a small slice from each end of the file. Cheap enough to be free even
/// over a network share, and enough to separate same-size files that aren't
/// actually the same. Only ever used to decide who deserves a real full hash.
func edgeFingerprint(_ url: URL, window: Int = 64 * 1024) -> String? {
    guard let handle = try? FileHandle(forReadingFrom: url) else { return nil }
    defer { try? handle.close() }

    var hasher = SHA256()
    guard let head = try? handle.read(upToCount: window) else { return nil }
    hasher.update(data: head)

    if let size = try? handle.seekToEnd(), size > UInt64(window) {
        let tailStart = size - UInt64(window)
        if (try? handle.seek(toOffset: tailStart)) != nil,
           let tail = try? handle.read(upToCount: window) {
            hasher.update(data: tail)
        }
    }
    return hasher.finalize().map { String(format: "%02x", $0) }.joined()
}

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
func frameHash(_ url: URL, at seconds: Double) -> [UInt64]? {
    guard let ffmpeg = Tools.ffmpeg else { return nil }
    let (w, h) = (17, 16)                      // one extra column to compare against
    // -threads 1: we want one frame, and letting ffmpeg spin up its usual
    // decode threads costs ~2.7x the CPU for no gain. Measured on 4K HEVC:
    // 0.33s real / 0.76s CPU by default, against 0.27s / 0.28s with one thread.
    // Multiplied by every worker, that difference is what made the machine drag.
    let r = runProcessRetrying(ffmpeg, [
        "-v", "quiet", "-threads", "1", "-ss", String(seconds), "-i", url.path,
        "-frames:v", "1", "-vf", "scale=\(w):\(h)", "-pix_fmt", "gray",
        "-f", "rawvideo", "-",
    ], timeout: 30)
    guard r.status == 0, r.out.count >= w * h else { return nil }

    let px = [UInt8](r.out.prefix(w * h))
    var words = [UInt64](repeating: 0, count: hashWordsPerFrame)
    var bit = 0
    for row in 0..<h {
        for col in 0..<(w - 1) {
            let i = row * w + col
            if px[i] > px[i + 1] { words[bit / 64] |= (1 << UInt64(bit % 64)) }
            bit += 1
        }
    }
    return words
}

/// One frame of a cluster's signature, served from the cache when the file
/// hasn't changed since it was last sampled.
func clusterFrame(_ c: ByteCluster, stage: Int, cache: ScanCache?) -> [UInt64]? {
    let vf = c.files[0]
    if let e = cache?.entry(for: vf.url.path, size: vf.size, mtime: vf.mtime),
       let cached = e.signature, cached.count >= (stage + 1) * hashWordsPerFrame {
        let start = stage * hashWordsPerFrame
        return Array(cached[start ..< start + hashWordsPerFrame])
    }
    guard let duration = vf.duration, duration > 0 else { return nil }
    return frameHash(vf.url, at: max(0, duration * samplePositions[stage]))
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
    /// Frames sampled so far, in `samplePositions` order, each one a 256-bit
    /// hash. Filled a pass at a time — most clusters never need all of them.
    var frames: [[UInt64]] = []
    /// Sampling failed, so this cluster can't take part in a frame comparison.
    var failed = false

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

func groupByExactHash(_ files: [VideoFile], workers: Int, cancel: CancelToken,
                      progress: @escaping (Int, Int, String) -> Void) -> [ByteCluster] {
    var bySize: [Int64: [VideoFile]] = [:]
    for f in files { bySize[f.size, default: []].append(f) }

    // only files sharing a size can possibly be byte-identical
    let sameSize = bySize.values.filter { $0.count > 1 }.flatMap { $0 }
    var digests: [UUID: String] = [:]

    if !sameSize.isEmpty {
        // First pass: a few KB from each end. Files that differ show it here,
        // which spares us streaming whole videos across the network to find out.
        let edges = concurrentMap(sameSize, limit: workers, cancel: cancel,
                                  onProgress: { d, t, f in progress(d, t, "Checking \(f.name)") },
                                  transform: { edgeFingerprint($0.url) })
        var byEdge: [String: [VideoFile]] = [:]
        for (i, e) in edges.enumerated() {
            let f = sameSize[i]
            // unreadable ends: fall back to a full hash rather than guess
            let key = (e ?? nil) ?? "unreadable-\(f.id)"
            byEdge["\(f.size)-\(key)", default: []].append(f)
        }

        // Second pass: confirm the survivors properly. Matching ends are not
        // proof, and these get deleted, so nothing is called exact on a sample.
        let needHash = byEdge.values.filter { $0.count > 1 }.flatMap { $0 }
        if !needHash.isEmpty {
            let hashed = concurrentMap(needHash, limit: workers, cancel: cancel,
                                       onProgress: { d, t, f in progress(d, t, "Hashing \(f.name)") },
                                       transform: { hashFileContents($0.url) })
            for (i, h) in hashed.enumerated() {
                if let h = h ?? nil { digests[needHash[i].id] = h }
            }
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

/// Only files whose duration is within tolerance of some other file can ever be
/// matched by `mergeBySimilarity`. Everything else is decided before a single
/// frame is read, so sampling it would be wasted work — and on a big library
/// that's the overwhelming majority of the files.
func clustersWorthSampling(_ clusters: [ByteCluster]) -> [ByteCluster] {
    let dated = clusters.filter { $0.duration != nil }
        .sorted { ($0.duration ?? 0) < ($1.duration ?? 0) }
    guard dated.count > 1 else { return [] }

    var worth: [ByteCluster] = []
    for (i, c) in dated.enumerated() {
        let d = c.duration ?? 0
        // sorted, so the closest durations are the immediate neighbours
        let nearPrev = i > 0 && d - (dated[i - 1].duration ?? 0) <= durationToleranceSec
        let nearNext = i + 1 < dated.count && (dated[i + 1].duration ?? 0) - d <= durationToleranceSec
        if nearPrev || nearNext { worth.append(c) }
    }
    return worth
}

/// Sample one frame at a time across the whole candidate set, discarding pairs
/// that already can't match before paying for the next frame.
///
/// The verdict is unchanged: a pair still has to finish with its average
/// distance within tolerance. It's only the order of work that differs. Because
/// the remaining frames can never *reduce* a running total, any pair already
/// past the budget after one frame is dead, and a cluster with no surviving
/// partner never gets sampled again — which is most of them, since unrelated
/// videos sit near the halfway mark of a 64-bit hash.
func matchClusters(_ clusters: [ByteCluster], workers: Int, cache: ScanCache?,
                   cancel: CancelToken,
                   progress: @escaping (Int, Int, String) -> Void) -> [DuplicateGroup] {
    let pool = clustersWorthSampling(clusters)   // already sorted by duration
    let budget = maxHammingPerFrame * samplePositions.count

    // Every pair close enough in duration to be worth a look. This is quadratic
    // in how many clips share a length, so on a big folder it is both slow and
    // memory-hungry — and it used to run silently, which is what made the app
    // look frozen. Report as it goes.
    var pairs: [(i: Int, j: Int, dist: Int)] = []
    for i in 0..<pool.count {
        if cancel.isCancelled { return [] }
        if i % 200 == 0 {
            progress(i, pool.count, "Pairing up candidates (\(pairs.count) so far)")
        }
        for j in (i + 1)..<pool.count {
            if (pool[j].duration ?? 0) - (pool[i].duration ?? 0) > durationToleranceSec { break }
            pairs.append((i, j, 0))
        }
    }
    ScanLog.shared.note("  step 4: \(clusters.count) cluster(s), \(pool.count) worth sampling, "
                        + "\(pairs.count) candidate pair(s)")

    for stage in 0..<samplePositions.count {
        if pairs.isEmpty || cancel.isCancelled { break }

        // only clusters still in a live pair need this frame
        var needed = Set<Int>()
        for p in pairs {
            if pool[p.i].frames.count <= stage { needed.insert(p.i) }
            if pool[p.j].frames.count <= stage { needed.insert(p.j) }
        }

        if !needed.isEmpty {
            let todo = needed.sorted()
            let got = concurrentMap(todo, limit: workers, cancel: cancel,
                                    onProgress: { d, t, idx in
                                        progress(d, t, "Pass \(stage + 1)/\(samplePositions.count): "
                                                 + pool[idx].files[0].name)
                                    },
                                    transform: { clusterFrame(pool[$0], stage: stage, cache: cache) })
            for (k, value) in got.enumerated() {
                let c = pool[todo[k]]
                if let h = value ?? nil {
                    c.frames.append(h)
                    let vf = c.files[0]
                    cache?.storeSignature(path: vf.url.path, size: vf.size,
                                          mtime: vf.mtime, signature: c.frames.flatMap { $0 })
                } else {
                    c.failed = true
                    let vf = c.files[0]
                    // remember the failure, so a broken file isn't retried every scan
                    cache?.storeSignature(path: vf.url.path, size: vf.size,
                                          mtime: vf.mtime, signature: nil)
                }
            }
        }

        let before = pairs.count
        pairs = pairs.compactMap { p in
            let a = pool[p.i], b = pool[p.j]
            guard !a.failed, !b.failed,
                  a.frames.count > stage, b.frames.count > stage else { return nil }
            let d = p.dist + hamming(a.frames[stage], b.frames[stage])
            return d > budget ? nil : (p.i, p.j, d)
        }
        // how well one-frame-first is actually working on this folder
        ScanLog.shared.note("  pass \(stage + 1): sampled \(needed.count) file(s), "
                            + "pairs \(before) → \(pairs.count)")
    }

    // Group only where every member matches every other member.
    //
    // Merging transitively (A~B, B~C therefore A+B+C) quietly breaks the
    // duration rule: each link can sit inside the tolerance while the ends are
    // far outside it. A folder of similar-looking clips then chains into one
    // enormous group spanning durations that were never compared. Requiring
    // mutual agreement keeps a group meaning what it says — everything in here
    // is a duplicate of everything else in here.
    var neighbours: [Int: Set<Int>] = [:]
    for p in pairs {
        neighbours[p.i, default: []].insert(p.j)
        neighbours[p.j, default: []].insert(p.i)
    }

    var groups: [DuplicateGroup] = []
    var accounted = Set<ObjectIdentifier>()
    var taken = Set<Int>()

    // strongest hub first, so the most-connected file anchors its group
    let anchors = neighbours.keys.sorted {
        let (a, b) = (neighbours[$0]?.count ?? 0, neighbours[$1]?.count ?? 0)
        return a == b ? $0 < $1 : a > b
    }

    for (seen, anchor) in anchors.enumerated() where !taken.contains(anchor) {
        if seen % 200 == 0 { progress(seen, anchors.count, "Forming groups") }
        var members = [anchor]
        // closest duration first, so the tightest matches win a contested file
        let candidates = (neighbours[anchor] ?? []).subtracting(taken).sorted {
            let base = pool[anchor].duration ?? 0
            let da = abs((pool[$0].duration ?? 0) - base)
            let db = abs((pool[$1].duration ?? 0) - base)
            return da == db ? $0 < $1 : da < db
        }
        for c in candidates where members.allSatisfy({ neighbours[$0]?.contains(c) == true }) {
            members.append(c)
        }
        guard members.count > 1 else { continue }

        taken.formUnion(members)
        let here = members.map { pool[$0] }
        var files = here.flatMap { $0.files }
        files.sort(by: betterQuality)
        // "exact" only when a single byte-identical cluster is involved
        groups.append(DuplicateGroup(kind: here.count == 1 ? .exact : .possible, files: files))
        for c in here { accounted.insert(ObjectIdentifier(c)) }
    }

    // byte-identical clusters that never took part in a frame match — including
    // everything the duration filter skipped — are still duplicates
    for c in clusters where c.files.count > 1 && !accounted.contains(ObjectIdentifier(c)) {
        groups.append(DuplicateGroup(kind: .exact, files: c.files.sorted(by: betterQuality)))
    }
    return groups
}

// ── Entry point ──────────────────────────────────────────────────────────────
func scan(folders: [URL], cancel: CancelToken, cache: ScanCache? = ScanCache.load(),
          progress: @escaping (Int, Int, String) -> Void) -> [DuplicateGroup] {
    let workers = workerCount(for: folders)
    let scanStarted = Date()
    let log = ScanLog.shared
    log.resetCounts()
    log.beginScan()
    log.note("──────── scan starting ────────")

    // Advisory only here — the app refuses outright, but a scan driven from a
    // tool should still record that it was competing for the same disks, since
    // that inflates every timing in this file.
    let gotLock = ScanLock.acquire(folders: folders)
    if !gotLock, let other = ScanLock.holder() {
        log.note(String(format: "  WARNING: another scan (pid %d) has been running %.0fs "
                        + "on %@ — every timing below is inflated by the contention",
                        other.pid, other.age, other.folders))
    }
    defer {
        if gotLock { ScanLock.release() }
        log.flush()
    }
    for f in folders {
        let local = (try? f.resourceValues(forKeys: [.volumeIsLocalKey]))?.volumeIsLocal ?? true
        log.note("  folder: \(f.path)  [\(local ? "local" : "network")]")
    }
    log.note("  workers: \(workers)   cores: \(ProcessInfo.processInfo.activeProcessorCount)"
             + "   cache: \(cache == nil ? "off" : "on")")

    var phase = Date()
    func mark(_ name: String, _ extra: String = "") {
        log.note(String(format: "  %@ took %.1fs %@", name, -phase.timeIntervalSinceNow, extra))
        phase = Date()
    }

    progress(0, 0, "Step 1/4 — Finding video files…")
    let paths = findVideoFiles(in: folders)
    mark("step 1 find files", "— \(paths.count) files")
    if cancel.isCancelled { log.note("  cancelled"); log.flush(); return [] }

    var files: [VideoFile] = []
    if !paths.isEmpty {
        let built = concurrentMap(paths, limit: workers, cancel: cancel,
                                  onProgress: { d, t, u in
                                      progress(d, t, "Step 2/4 — Reading metadata: \(u.lastPathComponent)")
                                  },
                                  transform: { buildVideoFile($0, cache: cache) })
        files = built.compactMap { $0 ?? nil }
    }
    // A file with no duration can't be compared against anything, so it drops
    // out of detection entirely. That has to be visible: "no duplicates found"
    // means something quite different if a chunk of the folder was never read.
    let unreadable = files.filter { ($0.duration ?? 0) <= 0 }
    mark("step 2 metadata", "— \(files.count) files, \(unreadable.count) with no duration")
    if !unreadable.isEmpty {
        log.note("  WARNING: \(unreadable.count) file(s) could not be read and were "
                 + "excluded from duplicate detection:")
        for f in unreadable.prefix(20) { log.note("    \(f.name)") }
        if unreadable.count > 20 { log.note("    …and \(unreadable.count - 20) more") }
    }
    if cancel.isCancelled { cache?.save(); log.note("  cancelled"); log.flush(); return [] }
    files.sort { $0.url.path < $1.url.path }  // threads finish out of order

    let clusters = groupByExactHash(files, workers: workers, cancel: cancel) { d, t, m in
        progress(d, t, "Step 3/4 — Checking exact duplicates: \(m)")
    }
    mark("step 3 exact duplicates")
    if cancel.isCancelled { cache?.save(); log.note("  cancelled"); log.flush(); return [] }

    var groups = matchClusters(clusters, workers: workers, cache: cache, cancel: cancel) { d, t, m in
        progress(d, t, "Step 4/4 — Checking for re-encoded duplicates: \(m)")
    }
    mark("step 4 re-encoded duplicates")
    // keep whatever we learned, even if the user cancels part-way
    cache?.save()
    if cancel.isCancelled { log.note("  cancelled"); log.flush(); return [] }

    groups.sort { $0.wastedBytes > $1.wastedBytes }  // biggest wins first

    let t = log.tallies
    log.note(String(format: "  TOTAL %.1fs — %d group(s); %d slow op(s), %d timeout(s), %d failure(s)",
                    -scanStarted.timeIntervalSinceNow, groups.count,
                    t.slow, t.timeouts, t.failures))
    if t.timeouts > 0 {
        log.note("  NOTE: timeouts usually mean the drive stalled or spun down, "
                 + "not that the file is bad")
    }
    log.note("")
    log.flush()
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
