import Foundation

// ── What a job reports back while it runs ────────────────────────────────────

struct JobProgress {
    /// 0...1, or nil when nothing known can be measured against.
    var fraction: Double?
    /// What is happening right now, e.g. "Joining 3 clips" or a filename.
    var detail: String
    /// ffmpeg's encoding speed, e.g. "18.4x". Empty while stream-copying fast
    /// enough that it means nothing.
    var speed: String = ""
}

enum JobOutcome {
    case finished(outputs: [URL])
    /// Stream copy refused the inputs. The window offers to re-encode, which
    /// runs the same job again with `allowReencode` set.
    case needsReencode
    case failed(String)
    case cancelled
}

enum HEVCEncoder: String, CaseIterable, Identifiable {
    case quality, speed
    var id: String { rawValue }

    var label: String {
        switch self {
        case .quality: return "Quality — best compression (libx265, slower)"
        case .speed:   return "Speed — hardware accelerated (Apple HEVC, faster)"
        }
    }
}

enum Job {
    case merge(items: [MediaItem], output: URL, allowReencode: Bool)
    case dvd(title: DVDTitle, output: URL)
    case convert(items: [MediaItem], encoder: HEVCEncoder)
}

// ── Reading ffmpeg's progress feed ───────────────────────────────────────────

/// Turns ffmpeg's `-progress` key=value stream into a fraction of the whole job.
///
/// `out_time_us` counts microseconds of finished output, so measuring it against
/// the total duration of the inputs gives a real percentage rather than a bar
/// that just spins. `completed` carries the seconds already finished by earlier
/// files in a multi-file job.
private final class ProgressFeed {
    private let totalSeconds: Double
    private let report: (JobProgress) -> Void
    private var completed: Double = 0
    private var detail: String = ""
    private var speed: String = ""

    init(totalSeconds: Double, report: @escaping (JobProgress) -> Void) {
        self.totalSeconds = totalSeconds
        self.report = report
    }

    func startFile(_ name: String) {
        detail = name
        emit(current: 0)
    }

    func finishFile(seconds: Double) {
        completed += seconds
        emit(current: 0)
    }

    func consume(_ line: String) {
        let parts = line.split(separator: "=", maxSplits: 1)
        guard parts.count == 2 else { return }
        let key = parts[0].trimmingCharacters(in: .whitespaces)
        let value = parts[1].trimmingCharacters(in: .whitespaces)
        switch key {
        case "out_time_us", "out_time_ms":
            // Despite the name, ffmpeg reports out_time_ms in microseconds too.
            if let micro = Double(value), micro >= 0 { emit(current: micro / 1_000_000) }
        case "speed":
            // "N/A" until the first frames land — not worth showing.
            speed = value.hasSuffix("x") ? value : ""
        default:
            break
        }
    }

    private func emit(current: Double) {
        var fraction: Double?
        if totalSeconds > 0 {
            fraction = min(max((completed + current) / totalSeconds, 0), 1)
        }
        report(JobProgress(fraction: fraction, detail: detail, speed: speed))
    }
}

// ── Running a job ────────────────────────────────────────────────────────────

/// Do the work. Blocking — call it off the main thread.
func run(job: Job, control: JobControl, report: @escaping (JobProgress) -> Void) -> JobOutcome {
    guard let ffmpeg = Tools.ffmpeg else {
        return .failed("Could not find the ffmpeg engine inside the app.")
    }
    switch job {
    case let .merge(items, output, allowReencode):
        return runMerge(ffmpeg, items, output, allowReencode, control, report)
    case let .dvd(title, output):
        return runDVD(ffmpeg, title, output, control, report)
    case let .convert(items, encoder):
        return runConvert(ffmpeg, items, encoder, control, report)
    }
}

/// Arguments every run shares: never prompt, never print a banner, only report
/// errors — and stream machine-readable progress on stdout instead of the
/// human-readable stats ffmpeg writes to stderr.
private let commonArgs = [
    "-y", "-hide_banner", "-loglevel", "error", "-nostats", "-progress", "pipe:1",
]

// ── Merging video files ──────────────────────────────────────────────────────

private func runMerge(_ ffmpeg: String, _ items: [MediaItem], _ output: URL,
                      _ allowReencode: Bool, _ control: JobControl,
                      _ report: @escaping (JobProgress) -> Void) -> JobOutcome {
    let work = FileManager.default.temporaryDirectory
        .appendingPathComponent("videomerger-\(UUID().uuidString)")
    try? FileManager.default.createDirectory(at: work, withIntermediateDirectories: true)
    defer { try? FileManager.default.removeItem(at: work) }

    let listFile = work.appendingPathComponent("list.txt")
    // The concat demuxer's own quoting: single-quote each path, and escape any
    // single quote inside it the awkward way concat expects.
    let list = items.map { item -> String in
        "file '" + item.url.path.replacingOccurrences(of: "'", with: #"'\''"#) + "'"
    }.joined(separator: "\n") + "\n"
    do {
        try list.write(to: listFile, atomically: true, encoding: .utf8)
    } catch {
        return .failed("Could not prepare the merge list: \(error.localizedDescription)")
    }

    let total = items.reduce(0) { $0 + $1.duration }
    let feed = ProgressFeed(totalSeconds: total, report: report)
    let what = allowReencode ? "Re-encoding \(items.count) clips into one"
                             : "Joining \(items.count) clips"
    feed.startFile(what)

    let encodeArgs = allowReencode
        ? ["-c:v", "libx264", "-crf", "18", "-preset", "medium", "-pix_fmt", "yuv420p",
           "-c:a", "aac", "-b:a", "256k"]
        : ["-c", "copy"]
    let args = commonArgs
        + ["-f", "concat", "-safe", "0", "-i", listFile.path]
        + encodeArgs
        + ["-movflags", "+faststart", output.path]

    let result = runProcess(ffmpeg, args, control: control) { feed.consume($0) }
    if result.cancelled { try? FileManager.default.removeItem(at: output); return .cancelled }
    if result.ok { return .finished(outputs: [output]) }

    try? FileManager.default.removeItem(at: output)
    // A plain stream copy fails when the clips don't share a format. That is an
    // offer to re-encode rather than an error — unless re-encoding is what just
    // failed, in which case something is genuinely wrong.
    if !allowReencode { return .needsReencode }
    return .failed(lastLines(result.err))
}

// ── Merging a DVD title ──────────────────────────────────────────────────────

private func runDVD(_ ffmpeg: String, _ title: DVDTitle, _ output: URL,
                    _ control: JobControl,
                    _ report: @escaping (JobProgress) -> Void) -> JobOutcome {
    let total = title.parts.reduce(0.0) { $0 + inspect($1).duration }
    let feed = ProgressFeed(totalSeconds: total, report: report)
    feed.startFile("Joining title \(title.number) — \(title.parts.count) parts")

    // VOBs are MPEG program streams, which the concat *protocol* joins byte-wise
    // before demuxing — the right tool here, where the concat demuxer isn't.
    let joined = title.parts.map(\.path).joined(separator: "|")
    let args = commonArgs + ["-i", "concat:\(joined)", "-c", "copy", output.path]

    let result = runProcess(ffmpeg, args, control: control) { feed.consume($0) }
    if result.cancelled { try? FileManager.default.removeItem(at: output); return .cancelled }
    if result.ok { return .finished(outputs: [output]) }

    // Fall back to a byte-exact join, as the shell version did: a DVD with a
    // damaged stream often still plays once the parts are simply concatenated.
    try? FileManager.default.removeItem(at: output)
    report(JobProgress(fraction: nil, detail: "Joining parts directly…"))
    if concatenateBytes(title.parts, to: output, control: control,
                        totalBytes: title.bytes, report: report) {
        return .finished(outputs: [output])
    }
    if control.isCancelled { try? FileManager.default.removeItem(at: output); return .cancelled }
    try? FileManager.default.removeItem(at: output)
    return .failed(lastLines(result.err))
}

/// Byte-for-byte append of every part into one file, in 8MB chunks so a 7GB
/// title doesn't have to fit in memory and the bar keeps moving.
private func concatenateBytes(_ parts: [URL], to output: URL, control: JobControl,
                              totalBytes: Int64,
                              report: @escaping (JobProgress) -> Void) -> Bool {
    let fm = FileManager.default
    fm.createFile(atPath: output.path, contents: nil)
    guard let sink = try? FileHandle(forWritingTo: output) else { return false }
    defer { try? sink.close() }

    var written: Int64 = 0
    for part in parts {
        guard let source = try? FileHandle(forReadingFrom: part) else { return false }
        defer { try? source.close() }
        while true {
            if control.isCancelled { return false }
            guard let chunk = try? source.read(upToCount: 8 << 20), !chunk.isEmpty else { break }
            do { try sink.write(contentsOf: chunk) } catch { return false }
            written += Int64(chunk.count)
            report(JobProgress(
                fraction: totalBytes > 0 ? Double(written) / Double(totalBytes) : nil,
                detail: "Joining parts directly — \(part.lastPathComponent)"))
        }
    }
    return true
}

// ── Converting to HEVC ───────────────────────────────────────────────────────

private func runConvert(_ ffmpeg: String, _ items: [MediaItem], _ encoder: HEVCEncoder,
                        _ control: JobControl,
                        _ report: @escaping (JobProgress) -> Void) -> JobOutcome {
    let total = items.reduce(0) { $0 + $1.duration }
    let feed = ProgressFeed(totalSeconds: total, report: report)
    var outputs: [URL] = []
    var failures: [String] = []

    for (index, item) in items.enumerated() {
        if control.isCancelled { return .cancelled }
        let output = uniqueOutput(for: item.url)
        feed.startFile("\(index + 1) of \(items.count) — \(item.name)")

        var result = runProcess(ffmpeg, convertArgs(item.url, output, encoder, fallback: false),
                                control: control) { feed.consume($0) }
        // videotoolbox rejects -q:v on some inputs; the shell version retried
        // with a bitrate target, and so does this.
        if !result.ok, !result.cancelled, encoder == .speed {
            try? FileManager.default.removeItem(at: output)
            result = runProcess(ffmpeg, convertArgs(item.url, output, encoder, fallback: true),
                                control: control) { feed.consume($0) }
        }

        if result.cancelled { try? FileManager.default.removeItem(at: output); return .cancelled }
        if result.ok {
            outputs.append(output)
        } else {
            try? FileManager.default.removeItem(at: output)
            failures.append("\(item.name): \(lastLines(result.err, limit: 2))")
        }
        feed.finishFile(seconds: item.duration)
    }

    if outputs.isEmpty {
        return .failed(failures.isEmpty ? "Nothing was converted."
                                        : failures.joined(separator: "\n"))
    }
    if !failures.isEmpty {
        return .failed("Converted \(outputs.count) of \(items.count). Failed:\n"
                       + failures.joined(separator: "\n"))
    }
    return .finished(outputs: outputs)
}

private func convertArgs(_ input: URL, _ output: URL,
                         _ encoder: HEVCEncoder, fallback: Bool) -> [String] {
    // -map 0:a? keeps the audio when there is some and doesn't fail when there
    // isn't. hvc1 is the tag QuickTime needs to play HEVC in an MP4.
    let common = commonArgs + ["-i", input.path, "-map", "0:v:0", "-map", "0:a?"]
    let audio = ["-tag:v", "hvc1", "-c:a", "aac", "-b:a", "192k",
                 "-movflags", "+faststart", output.path]
    switch encoder {
    case .quality:
        return common + ["-c:v", "libx265", "-crf", "20", "-preset", "medium"] + audio
    case .speed:
        let rate = fallback ? ["-b:v", "10M"] : ["-q:v", "60"]
        return common + ["-c:v", "hevc_videotoolbox"] + rate + audio
    }
}

/// `clip.mov` becomes `clip_x265.mp4` beside it, numbered if that already exists.
private func uniqueOutput(for input: URL) -> URL {
    let folder = input.deletingLastPathComponent()
    let stem = input.deletingPathExtension().lastPathComponent
    var candidate = folder.appendingPathComponent("\(stem)_x265.mp4")
    var n = 1
    while FileManager.default.fileExists(atPath: candidate.path) {
        candidate = folder.appendingPathComponent("\(stem)_x265_\(n).mp4")
        n += 1
    }
    return candidate
}

/// ffmpeg's failures are explained in its last few lines; the rest is noise.
private func lastLines(_ text: String, limit: Int = 6) -> String {
    let lines = text.split(whereSeparator: \.isNewline).map(String.init)
    guard !lines.isEmpty else { return "ffmpeg failed without saying why." }
    return lines.suffix(limit).joined(separator: "\n")
}
