import Foundation

/// A scan log written to ~/Library/Logs/Duplicate Video Finder/scan.log.
///
/// Aimed at the questions that are hard to answer while watching a scan: where
/// the time actually went, whether a pause was a slow drive or our own work,
/// whether the cache is doing anything. It records phase timings, anything that
/// took unusually long, and every failure — not every operation, which would
/// cost more than it tells us.
final class ScanLog: @unchecked Sendable {
    static let shared = ScanLog()

    /// Only operations slower than this get a line of their own. A frame
    /// extract is ~0.3s locally, so anything past a couple of seconds is a
    /// stalled read, a spinning-up drive, or a very large file.
    static let slowThreshold: TimeInterval = 2.0

    private let lock = NSLock()
    private var buffer: [String] = []
    private var slowCount = 0
    private var timeoutCount = 0
    private var failureCount = 0

    private let maxBytes = 2 * 1024 * 1024        // keep one previous log alongside

    let fileURL: URL = {
        let base = FileManager.default.urls(for: .libraryDirectory, in: .userDomainMask).first
            ?? URL(fileURLWithPath: NSHomeDirectory() + "/Library")
        return base.appendingPathComponent("Logs/Duplicate Video Finder/scan.log")
    }()

    private static let stamp: DateFormatter = {
        let f = DateFormatter()
        f.dateFormat = "yyyy-MM-dd HH:mm:ss.SSS"
        return f
    }()

    // ── writing ──────────────────────────────────────────────────────────────
    /// Identifies which scan wrote a line. Two scans appending to one file
    /// interleave, and their timestamps then look out of order — without this
    /// the log is genuinely hard to read when anything overlaps.
    private var tag = "----"

    func beginScan() {
        lock.lock()
        tag = String(format: "%04x", UInt16.random(in: 0...UInt16.max))
        lock.unlock()
    }

    func note(_ message: String) {
        lock.lock(); let id = tag; lock.unlock()
        let line = "\(Self.stamp.string(from: Date()))  [\(id)] \(message)"
        lock.lock()
        buffer.append(line)
        let full = buffer.count >= 200
        lock.unlock()
        if full { flush() }
    }

    /// Record an operation only when it was slow enough to be worth knowing.
    func timed(_ operation: String, seconds: TimeInterval, detail: String) {
        guard seconds >= Self.slowThreshold else { return }
        lock.lock(); slowCount += 1; lock.unlock()
        note(String(format: "SLOW  %@ took %.1fs — %@", operation, seconds, detail))
    }

    func timedOut(_ operation: String, detail: String) {
        lock.lock(); timeoutCount += 1; lock.unlock()
        note("TIMEOUT  \(operation) — \(detail)")
    }

    func failed(_ operation: String, detail: String) {
        lock.lock(); failureCount += 1; lock.unlock()
        note("FAIL  \(operation) — \(detail)")
    }

    /// Files the scan could not read, and so never actually checked. This has
    /// to travel back to the window: "no duplicates found" means something very
    /// different if part of the folder was skipped.
    private var excludedFiles: [(name: String, reason: String)] = []

    func excluded(_ name: String, reason: String) {
        lock.lock(); excludedFiles.append((name, reason)); lock.unlock()
    }

    var excluded: [(name: String, reason: String)] {
        lock.lock(); defer { lock.unlock() }
        return excludedFiles
    }

    /// Counts since the last `resetCounts()`, for the end-of-scan summary.
    var tallies: (slow: Int, timeouts: Int, failures: Int) {
        lock.lock(); defer { lock.unlock() }
        return (slowCount, timeoutCount, failureCount)
    }

    func resetCounts() {
        lock.lock()
        slowCount = 0; timeoutCount = 0; failureCount = 0
        excludedFiles.removeAll()
        lock.unlock()
    }

    func flush() {
        lock.lock()
        guard !buffer.isEmpty else { lock.unlock(); return }
        let text = buffer.joined(separator: "\n") + "\n"
        buffer.removeAll(keepingCapacity: true)
        lock.unlock()

        let fm = FileManager.default
        try? fm.createDirectory(at: fileURL.deletingLastPathComponent(),
                                withIntermediateDirectories: true)
        rotateIfTooBig()

        // O_APPEND, so concurrent scans can't overwrite each other. Seeking to
        // the end and then writing is two steps, and a second process writing
        // between them silently loses the first one's lines — which is exactly
        // what happened the one time two scans overlapped.
        let fd = open(fileURL.path, O_WRONLY | O_CREAT | O_APPEND, 0o644)
        guard fd >= 0 else { return }
        defer { close(fd) }
        _ = text.withCString { write(fd, $0, strlen($0)) }
    }

    /// Keep the current log and one previous, so it can't grow without limit.
    private func rotateIfTooBig() {
        let fm = FileManager.default
        guard let size = try? fm.attributesOfItem(atPath: fileURL.path)[.size] as? Int64,
              size > maxBytes else { return }
        let previous = fileURL.deletingPathExtension().appendingPathExtension("previous.log")
        try? fm.removeItem(at: previous)
        try? fm.moveItem(at: fileURL, to: previous)
    }
}

/// Time a block and log it if it was slow. Returns whatever the block returns.
@inline(__always)
func logIfSlow<T>(_ operation: String, _ detail: @autoclosure () -> String,
                  _ body: () -> T) -> T {
    let start = Date()
    let result = body()
    ScanLog.shared.timed(operation, seconds: -start.timeIntervalSinceNow, detail: detail())
    return result
}
