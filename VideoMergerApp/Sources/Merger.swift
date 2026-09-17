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
    /// Bytes of output written so far, across every file in the job.
    var bytesWritten: Int64 = 0
    /// Bytes per second, averaged over the last few seconds. 0 until there are
    /// enough samples to mean anything.
    var rate: Double = 0
    /// Names the network drives this job is moving data across, e.g.
    /// "Copying to Media over the network". Empty when it is all local, which
    /// is what makes the line worth showing at all: a merge that is shuttling
    /// gigabytes over AFP otherwise looks exactly like one on the internal SSD.
    var transfer: String = ""
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
    private let transfer: String
    private let report: (JobProgress) -> Void
    private var completed: Double = 0
    private var detail: String = ""
    private var speed: String = ""
    /// Bytes written by files already finished, so a multi-file job keeps a
    /// running total rather than restarting at each one.
    private var doneBytes: Int64 = 0
    private var fileBytes: Int64 = 0
    private var samples: [(at: Date, total: Int64)] = []
    private var rate: Double = 0

    init(totalSeconds: Double, transfer: String = "",
         report: @escaping (JobProgress) -> Void) {
        self.totalSeconds = totalSeconds
        self.transfer = transfer
        self.report = report
    }

    func startFile(_ name: String) {
        detail = name
        emit(current: 0)
    }

    func finishFile(seconds: Double) {
        completed += seconds
        doneBytes += fileBytes
        fileBytes = 0
        emit(current: 0)
    }

    /// Bytes reported by an external copier rather than by ffmpeg.
    func setBytes(_ written: Int64) { track(written) }

    private func track(_ written: Int64) {
        fileBytes = written
        let now = Date()
        let total = doneBytes + written
        samples.append((now, total))
        // A four-second window: long enough to ride out a stalled network
        // write, short enough to still read as "now".
        samples.removeAll { now.timeIntervalSince($0.at) > 4 }
        if let first = samples.first, samples.count > 1 {
            let seconds = now.timeIntervalSince(first.at)
            if seconds > 0.5 { rate = Double(total - first.total) / seconds }
        }
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
        case "total_size":
            // How much output exists so far. On a stream copy to a share this
            // is literally the number of bytes that have crossed the network.
            if let written = Int64(value), written >= 0 { track(written) }
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
        report(JobProgress(fraction: fraction, detail: detail, speed: speed,
                           bytesWritten: doneBytes + fileBytes, rate: rate,
                           transfer: transfer))
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

/// True when the file will be written to a network volume.
func isOnNetworkVolume(_ url: URL) -> Bool {
    // The output doesn't exist yet, so ask the folder it will be created in.
    let probe = FileManager.default.fileExists(atPath: url.path)
        ? url : url.deletingLastPathComponent()
    guard let local = (try? probe.resourceValues(forKeys: [.volumeIsLocalKey]))?.volumeIsLocal
    else { return false }       // can't tell — behave as we always did
    return !local
}

/// The network volume a path sits on, or nil when it is on a local disk.
private func networkVolume(of url: URL) -> String? {
    let fm = FileManager.default
    let probe = fm.fileExists(atPath: url.path) ? url : url.deletingLastPathComponent()
    guard let local = (try? probe.resourceValues(forKeys: [.volumeIsLocalKey]))?.volumeIsLocal,
          !local else { return nil }
    return (try? probe.resourceValues(forKeys: [.volumeNameKey]))?.volumeName
        ?? "a network drive"
}

/// Says which network drives a job will move data across, for the window to
/// show while it runs. Empty when everything is on local disks — there is
/// nothing to warn about then, and a line that always appears stops being read.
private func transferNote(from inputs: [URL], to output: URL) -> String {
    let sources = Set(inputs.compactMap(networkVolume(of:))).sorted()
    let destination = networkVolume(of: output)

    func list(_ names: [String]) -> String {
        names.count <= 1 ? (names.first ?? "")
                         : names.dropLast().joined(separator: ", ") + " and " + names.last!
    }

    switch (sources.isEmpty, destination) {
    case (true, nil):
        return ""
    case (true, let to?):
        return "Copying to \(to) over the network"
    case (false, nil):
        return "Copying from \(list(sources)) over the network"
    case (false, let to?):
        // Both ends on the same share still crosses the network twice: every
        // byte is read down and written back up.
        return sources == [to]
            ? "Copying within \(to) over the network"
            : "Copying from \(list(sources)) to \(to) over the network"
    }
}

/// `+faststart` moves an MP4's index to the front, which ffmpeg does by reading
/// the finished file back and writing the whole thing out again. Locally that's
/// quick. On a share it turns one transfer into three — a 9GB merge onto a NAS
/// spent most of half an hour here, reporting no progress the entire time.
///
/// The index only has to be at the front for progressive streaming straight off
/// a web server. Every player that opens a file — and everything that reads from
/// a NAS — simply seeks to the end and finds it there. So it isn't worth 3x the
/// network, and we skip it when the destination isn't a local disk.
private func faststartArgs(for output: URL) -> [String] {
    isOnNetworkVolume(output) ? [] : ["-movflags", "+faststart"]
}

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
    let feed = ProgressFeed(totalSeconds: total,
                            transfer: transferNote(from: items.map(\.url), to: output),
                            report: report)
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
        + faststartArgs(for: output)
        + [output.path]

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
    let feed = ProgressFeed(totalSeconds: total,
                            transfer: transferNote(from: title.parts, to: output),
                            report: report)
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
                        totalBytes: title.bytes,
                        transfer: transferNote(from: title.parts, to: output),
                        report: report) {
        return .finished(outputs: [output])
    }
    if control.isCancelled { try? FileManager.default.removeItem(at: output); return .cancelled }
    try? FileManager.default.removeItem(at: output)
    return .failed(lastLines(result.err))
}

/// Byte-for-byte append of every part into one file, in 8MB chunks so a 7GB
/// title doesn't have to fit in memory and the bar keeps moving.
private func concatenateBytes(_ parts: [URL], to output: URL, control: JobControl,
                              totalBytes: Int64, transfer: String,
                              report: @escaping (JobProgress) -> Void) -> Bool {
    let fm = FileManager.default
    fm.createFile(atPath: output.path, contents: nil)
    guard let sink = try? FileHandle(forWritingTo: output) else { return false }
    defer { try? sink.close() }

    var written: Int64 = 0
    var samples: [(at: Date, total: Int64)] = []
    var rate: Double = 0

    for part in parts {
        guard let source = try? FileHandle(forReadingFrom: part) else { return false }
        defer { try? source.close() }
        while true {
            if control.isCancelled { return false }
            guard let chunk = try? source.read(upToCount: 8 << 20), !chunk.isEmpty else { break }
            do { try sink.write(contentsOf: chunk) } catch { return false }
            written += Int64(chunk.count)

            let now = Date()
            samples.append((now, written))
            samples.removeAll { now.timeIntervalSince($0.at) > 4 }
            if let first = samples.first, samples.count > 1 {
                let seconds = now.timeIntervalSince(first.at)
                if seconds > 0.5 { rate = Double(written - first.total) / seconds }
            }

            report(JobProgress(
                fraction: totalBytes > 0 ? Double(written) / Double(totalBytes) : nil,
                detail: "Joining parts directly — \(part.lastPathComponent)",
                bytesWritten: written, rate: rate, transfer: transfer))
        }
    }
    return true
}

// ── Converting to HEVC ───────────────────────────────────────────────────────

private func runConvert(_ ffmpeg: String, _ items: [MediaItem], _ encoder: HEVCEncoder,
                        _ control: JobControl,
                        _ report: @escaping (JobProgress) -> Void) -> JobOutcome {
    let total = items.reduce(0) { $0 + $1.duration }
    // Converting writes beside each original, so the route is the same for all
    // of them — the first one describes the job.
    let feed = ProgressFeed(totalSeconds: total,
                            transfer: items.first.map {
                                transferNote(from: [$0.url], to: uniqueOutput(for: $0.url))
                            } ?? "",
                            report: report)
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
    let audio = ["-tag:v", "hvc1", "-c:a", "aac", "-b:a", "192k"]
        + faststartArgs(for: output) + [output.path]
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
