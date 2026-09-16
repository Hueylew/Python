import Foundation

/// Remembers what we learned about a file last time, so re-scanning a library
/// that hasn't changed doesn't pay for ffprobe and frame sampling all over
/// again. Keyed on size + modification date, so editing or replacing a file
/// invalidates its entry on its own — there is nothing to clear by hand.
struct CacheEntry: Codable {
    var size: Int64
    var mtime: Double

    var duration: Double?
    var width: Int?
    var height: Int?
    var codec: String?
    var probed: Bool = false        // ffprobe ran, even if it told us nothing

    var signature: [UInt64]?
    var sampled: Bool = false       // frame sampling ran, even if it failed

    var lastSeen: Double = Date().timeIntervalSince1970

    func matches(size: Int64, mtime: Double) -> Bool {
        self.size == size && abs(self.mtime - mtime) < 0.000_01
    }
}

final class ScanCache: @unchecked Sendable {
    private var entries: [String: CacheEntry]
    private let lock = NSLock()
    private var dirty = false

    /// Keep the cache from growing without limit across many scans.
    private let maxEntries = 200_000

    static let defaultURL: URL = {
        let base = FileManager.default.urls(for: .applicationSupportDirectory,
                                            in: .userDomainMask).first
            ?? URL(fileURLWithPath: NSTemporaryDirectory())
        // v2: frame hashes went from 64-bit to 256-bit. Reading a v1 signature
        // as though it were v2 would silently compare the wrong bits, so the
        // old file is left behind rather than migrated.
        return base.appendingPathComponent("Duplicate Video Finder/cache-v2.json")
    }()

    let fileURL: URL

    init(fileURL: URL = ScanCache.defaultURL, entries: [String: CacheEntry] = [:]) {
        self.fileURL = fileURL
        self.entries = entries
    }

    static func load(from url: URL = ScanCache.defaultURL) -> ScanCache {
        guard let data = try? Data(contentsOf: url),
              let decoded = try? JSONDecoder().decode([String: CacheEntry].self, from: data)
        else { return ScanCache(fileURL: url) }
        return ScanCache(fileURL: url, entries: decoded)
    }

    var count: Int { lock.lock(); defer { lock.unlock() }; return entries.count }

    // ── reads ────────────────────────────────────────────────────────────────
    func entry(for path: String, size: Int64, mtime: Double) -> CacheEntry? {
        lock.lock(); defer { lock.unlock() }
        guard let e = entries[path], e.matches(size: size, mtime: mtime) else { return nil }
        return e
    }

    // ── writes ───────────────────────────────────────────────────────────────
    func storeMetadata(path: String, size: Int64, mtime: Double,
                       duration: Double?, width: Int?, height: Int?, codec: String?) {
        lock.lock(); defer { lock.unlock() }
        var e = entries[path].flatMap { $0.matches(size: size, mtime: mtime) ? $0 : nil }
            ?? CacheEntry(size: size, mtime: mtime)
        e.size = size; e.mtime = mtime
        e.duration = duration; e.width = width; e.height = height; e.codec = codec
        e.probed = true
        e.lastSeen = Date().timeIntervalSince1970
        entries[path] = e
        dirty = true
    }

    func storeSignature(path: String, size: Int64, mtime: Double, signature: [UInt64]?) {
        lock.lock(); defer { lock.unlock() }
        // only attach to an entry describing this same file
        guard var e = entries[path], e.matches(size: size, mtime: mtime) else { return }
        e.signature = signature
        e.sampled = true
        e.lastSeen = Date().timeIntervalSince1970
        entries[path] = e
        dirty = true
    }

    /// Forget everything. Used once a batch of duplicates has actually been
    /// dealt with: entries for files that no longer exist are dead weight, and
    /// the next scan of those folders should start from what's really there.
    func clear() {
        lock.lock()
        entries.removeAll()
        dirty = false
        lock.unlock()
        try? FileManager.default.removeItem(at: fileURL)
    }

    func save() {
        lock.lock()
        guard dirty else { lock.unlock(); return }
        if entries.count > maxEntries {
            // drop the least recently seen first
            let keep = entries.sorted { $0.value.lastSeen > $1.value.lastSeen }.prefix(maxEntries)
            entries = Dictionary(uniqueKeysWithValues: keep.map { ($0.key, $0.value) })
        }
        let snapshot = entries
        dirty = false
        lock.unlock()

        try? FileManager.default.createDirectory(at: fileURL.deletingLastPathComponent(),
                                                 withIntermediateDirectories: true)
        if let data = try? JSONEncoder().encode(snapshot) {
            try? data.write(to: fileURL, options: .atomic)
        }
    }
}

/// Size and modification date in one stat, which is what the cache is keyed on.
func fileStamp(_ url: URL) -> (size: Int64, mtime: Double)? {
    guard let a = try? FileManager.default.attributesOfItem(atPath: url.path),
          let size = a[.size] as? Int64 else { return nil }
    let mtime = (a[.modificationDate] as? Date)?.timeIntervalSince1970 ?? 0
    return (size, mtime)
}
