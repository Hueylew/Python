import Foundation

/// A note on disk saying "a scan is in progress", so two of them don't end up
/// fighting over the same drive.
///
/// This matters more than it sounds. Two scans of one external disk means twice
/// the concurrent readers on a mechanism that was already the bottleneck: a
/// real pair of overlapping scans turned a ~170s job into 443s and 571s, with
/// 67 timeouts between them, purely from contention.
///
/// The holder's process id is recorded, so a lock left behind by a crash is
/// spotted and ignored rather than blocking every future scan.
enum ScanLock {
    struct Holder {
        let pid: Int32
        let started: Date
        let folders: String

        var age: TimeInterval { -started.timeIntervalSinceNow }
        var isThisProcess: Bool { pid == ProcessInfo.processInfo.processIdentifier }
    }

    static let url: URL = {
        let base = FileManager.default.urls(for: .applicationSupportDirectory,
                                            in: .userDomainMask).first
            ?? URL(fileURLWithPath: NSTemporaryDirectory())
        return base.appendingPathComponent("Duplicate Video Finder/scan.lock")
    }()

    /// Whoever currently holds it, or nil if free — or held by a process that
    /// is no longer running.
    static func holder() -> Holder? {
        guard let text = try? String(contentsOf: url, encoding: .utf8) else { return nil }
        let parts = text.split(separator: "\n", omittingEmptySubsequences: false)
        guard parts.count >= 3, let pid = Int32(parts[0]),
              let epoch = Double(parts[1]) else { return nil }

        // kill(pid, 0) just asks whether the process exists
        guard kill(pid, 0) == 0 || errno == EPERM else {
            try? FileManager.default.removeItem(at: url)   // stale, from a crash
            return nil
        }
        return Holder(pid: pid, started: Date(timeIntervalSince1970: epoch),
                      folders: String(parts[2]))
    }

    /// Take the lock, or report that someone else already has it.
    ///
    /// Created with O_EXCL so the check and the claim are one step. Testing
    /// with two scans launched in the same millisecond, a plain
    /// look-then-write let both through — which is precisely the case the lock
    /// exists to stop.
    @discardableResult
    static func acquire(folders: [URL]) -> Bool {
        try? FileManager.default.createDirectory(at: url.deletingLastPathComponent(),
                                                 withIntermediateDirectories: true)
        _ = holder()      // clears the file first if a crashed run left it behind

        let fd = open(url.path, O_CREAT | O_EXCL | O_WRONLY, 0o644)
        guard fd >= 0 else { return false }   // someone got there first
        defer { close(fd) }

        let names = folders.map { $0.lastPathComponent }.joined(separator: ", ")
        let body = "\(ProcessInfo.processInfo.processIdentifier)\n"
            + "\(Date().timeIntervalSince1970)\n\(names)"
        _ = body.withCString { write(fd, $0, strlen($0)) }
        return true
    }

    /// Only drop the lock if it's still ours — never tread on another scan's.
    static func release() {
        guard let h = holder(), h.isThisProcess else { return }
        try? FileManager.default.removeItem(at: url)
    }
}
