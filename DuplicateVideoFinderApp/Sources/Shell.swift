import Foundation

// ── Result of running an external process ─────────────────────────────────────
struct RunResult {
    var status: Int32
    var out: Data
    var err: String
    var timedOut: Bool

    var outText: String { String(data: out, encoding: .utf8) ?? "" }
}

/// Run an external command, capturing stdout as raw bytes (frame data is binary),
/// with an optional hard timeout.
func runProcess(_ launchPath: String, _ args: [String], timeout: TimeInterval? = nil) -> RunResult {
    let proc = Process()
    proc.executableURL = URL(fileURLWithPath: launchPath)
    proc.arguments = args
    // scanning is background work — it shouldn't out-compete the foreground app
    proc.qualityOfService = .utility

    let outPipe = Pipe()
    let errPipe = Pipe()
    proc.standardOutput = outPipe
    proc.standardError = errPipe
    // ffmpeg reads stdin for interactive keypresses. Inheriting ours lets it
    // block there forever, which showed up as a 25s stall (the full timeout,
    // terminate and kill ladder) on a file that decodes in half a second.
    proc.standardInput = FileHandle.nullDevice

    let ioQueue = DispatchQueue(label: "proc.io")
    var outData = Data()
    var errData = Data()
    outPipe.fileHandleForReading.readabilityHandler = { h in
        let d = h.availableData
        if !d.isEmpty { ioQueue.sync { outData.append(d) } }
    }
    errPipe.fileHandleForReading.readabilityHandler = { h in
        let d = h.availableData
        if !d.isEmpty { ioQueue.sync { errData.append(d) } }
    }

    do {
        try proc.run()
    } catch {
        return RunResult(status: -1, out: Data(),
                         err: "failed to launch \(launchPath): \(error)", timedOut: false)
    }

    let started = Date()
    var timedOut = false
    if let timeout = timeout {
        let sem = DispatchSemaphore(value: 0)
        DispatchQueue.global().async { proc.waitUntilExit(); sem.signal() }
        if sem.wait(timeout: .now() + timeout) == .timedOut {
            timedOut = true
            proc.terminate()
            if sem.wait(timeout: .now() + 5) == .timedOut {
                if proc.isRunning { kill(proc.processIdentifier, SIGKILL) }
                // Wait for the kill to actually land. Asking a still-running
                // task for its terminationStatus raises an ObjC exception that
                // Swift cannot catch, which would take the whole app down.
                _ = sem.wait(timeout: .now() + 5)
            }
        }
    } else {
        proc.waitUntilExit()
    }

    outPipe.fileHandleForReading.readabilityHandler = nil
    errPipe.fileHandleForReading.readabilityHandler = nil
    if let rest = try? outPipe.fileHandleForReading.readToEnd() { ioQueue.sync { outData.append(rest) } }
    if let rest = try? errPipe.fileHandleForReading.readToEnd() { ioQueue.sync { errData.append(rest) } }

    // Belt and braces: if it somehow outlived the kill, report a failure rather
    // than asking for a status that would raise and terminate the process.
    let status: Int32 = proc.isRunning ? -1 : proc.terminationStatus

    // A stalled read is indistinguishable from a slow one while you're watching
    // it, so record which file and how long — that's what separates a drive
    // spinning up from the app doing real work.
    let tool = (launchPath as NSString).lastPathComponent
    let target = (args.last(where: { $0.hasPrefix("/") }) as NSString?)?.lastPathComponent
        ?? args.last ?? ""
    if timedOut {
        ScanLog.shared.timedOut(tool, detail: target)
    } else {
        ScanLog.shared.timed(tool, seconds: -started.timeIntervalSinceNow, detail: target)
        if status != 0 { ScanLog.shared.failed("\(tool) exit \(status)", detail: target) }
    }

    return ioQueue.sync {
        RunResult(status: status, out: outData,
                  err: String(data: errData, encoding: .utf8) ?? "", timedOut: timedOut)
    }
}

/// Locate ffmpeg/ffprobe. The copies bundled inside the .app win, so the app
/// keeps working on a Mac with no Homebrew; a system install is the fallback.
func findTool(_ name: String) -> String? {
    let fm = FileManager.default
    if let res = Bundle.main.resourceURL {
        let bundled = res.appendingPathComponent("bin/\(name)").path
        if fm.isExecutableFile(atPath: bundled) { return bundled }
    }
    var dirs = ["/opt/homebrew/bin", "/usr/local/bin", "/usr/bin", "/bin"]
    if let path = ProcessInfo.processInfo.environment["PATH"] {
        dirs = path.split(separator: ":").map(String.init) + dirs
    }
    for d in dirs {
        let p = d + "/" + name
        if fm.isExecutableFile(atPath: p) { return p }
    }
    return nil
}
