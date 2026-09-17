import Foundation

// ── Running external tools ───────────────────────────────────────────────────

/// What came back from an external tool.
struct RunResult {
    var status: Int32
    var err: String
    var cancelled: Bool

    var ok: Bool { status == 0 && !cancelled }
}

/// Holds whichever process is running right now, so the window's Cancel button
/// has something to stop. ffmpeg exits cleanly on SIGTERM, leaving a truncated
/// output file behind — the caller deletes it.
final class JobControl {
    private let lock = NSLock()
    private var current: Process?
    private var stopped = false

    var isCancelled: Bool { lock.withLock { stopped } }

    func cancel() {
        lock.withLock {
            stopped = true
            current?.terminate()
        }
    }

    /// Take ownership of a process. Returns false if Cancel was pressed while
    /// we were getting here, in which case the process must not start.
    func adopt(_ proc: Process) -> Bool {
        lock.withLock {
            guard !stopped else { return false }
            current = proc
            return true
        }
    }

    func release() { lock.withLock { current = nil } }
}

/// Run a tool, handing each line of stdout to `onOutLine` as it arrives.
///
/// stdout is streamed rather than collected because that is where ffmpeg's
/// `-progress` feed comes out: waiting for the process to exit would mean a
/// progress bar that only ever reads 0% or 100%. stderr is collected instead —
/// it is only wanted when something fails.
@discardableResult
func runProcess(_ launchPath: String, _ args: [String],
                control: JobControl? = nil,
                onOutLine: ((String) -> Void)? = nil) -> RunResult {
    let proc = Process()
    proc.executableURL = URL(fileURLWithPath: launchPath)
    proc.arguments = args

    let outPipe = Pipe(), errPipe = Pipe()
    proc.standardOutput = outPipe
    proc.standardError = errPipe
    // ffmpeg reads stdin for interactive keypresses; inheriting ours lets it
    // block there forever when it is run from a GUI app with no terminal.
    proc.standardInput = FileHandle.nullDevice

    let ioQueue = DispatchQueue(label: "merger.proc.io")
    var errData = Data()
    var pending = Data()

    outPipe.fileHandleForReading.readabilityHandler = { h in
        let chunk = h.availableData
        guard !chunk.isEmpty else { return }
        ioQueue.sync {
            pending.append(chunk)
            while let nl = pending.firstIndex(of: 0x0A) {
                let line = pending[pending.startIndex..<nl]
                pending.removeSubrange(pending.startIndex...nl)
                if let text = String(data: line, encoding: .utf8) {
                    onOutLine?(text.trimmingCharacters(in: .whitespacesAndNewlines))
                }
            }
        }
    }
    errPipe.fileHandleForReading.readabilityHandler = { h in
        let d = h.availableData
        if !d.isEmpty { ioQueue.sync { errData.append(d) } }
    }

    guard control?.adopt(proc) ?? true else {
        return RunResult(status: -1, err: "", cancelled: true)
    }
    do {
        try proc.run()
    } catch {
        control?.release()
        return RunResult(status: -1,
                         err: "could not start \(launchPath): \(error.localizedDescription)",
                         cancelled: false)
    }
    proc.waitUntilExit()
    control?.release()

    outPipe.fileHandleForReading.readabilityHandler = nil
    errPipe.fileHandleForReading.readabilityHandler = nil
    if let rest = try? outPipe.fileHandleForReading.readToEnd(), !rest.isEmpty {
        ioQueue.sync { pending.append(rest) }
    }
    if let rest = try? errPipe.fileHandleForReading.readToEnd(), !rest.isEmpty {
        ioQueue.sync { errData.append(rest) }
    }

    return ioQueue.sync {
        if let text = String(data: pending, encoding: .utf8), !text.isEmpty {
            onOutLine?(text.trimmingCharacters(in: .whitespacesAndNewlines))
        }
        return RunResult(status: proc.terminationStatus,
                         err: String(data: errData, encoding: .utf8) ?? "",
                         cancelled: control?.isCancelled ?? false)
    }
}

/// Run a tool and just collect its stdout — for short questions like ffprobe's.
func runCapturing(_ launchPath: String, _ args: [String]) -> String {
    var lines: [String] = []
    let queue = DispatchQueue(label: "merger.capture")
    runProcess(launchPath, args) { line in queue.sync { lines.append(line) } }
    return queue.sync { lines.joined(separator: "\n") }
}

/// Locate ffmpeg/ffprobe. The copies bundled inside the .app win, so the app
/// keeps working on a Mac with no Homebrew; a system install is the fallback.
func findTool(_ name: String) -> String? {
    let fm = FileManager.default
    if let res = Bundle.main.resourceURL {
        // "bin/ffmpeg" is this build's layout; the bare name is where the old
        // shell-script version of this app kept it.
        for candidate in ["bin/\(name)", name] {
            let path = res.appendingPathComponent(candidate).path
            if fm.isExecutableFile(atPath: path) { return path }
        }
    }
    var dirs = ["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin", "/bin"]
    if let path = ProcessInfo.processInfo.environment["PATH"] {
        dirs = path.split(separator: ":").map(String.init) + dirs
    }
    for d in dirs where fm.isExecutableFile(atPath: d + "/" + name) { return d + "/" + name }
    return nil
}

enum Tools {
    static let ffmpeg = findTool("ffmpeg")
    static let ffprobe = findTool("ffprobe")
}
