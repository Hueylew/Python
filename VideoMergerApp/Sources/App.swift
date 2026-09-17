import AppKit
import SwiftUI
import UniformTypeIdentifiers

enum Mode: String, CaseIterable, Identifiable {
    case merge, dvd, convert
    var id: String { rawValue }

    var label: String {
        switch self {
        case .merge:   return "Merge video files"
        case .dvd:     return "Merge a DVD folder"
        case .convert: return "Convert to MP4 (HEVC)"
        }
    }
}

// ── View model ───────────────────────────────────────────────────────────────

@MainActor
final class MergerModel: ObservableObject {
    /// One model for the whole app, so files dropped on the Dock icon reach the
    /// same list the window is showing.
    static let shared = MergerModel()

    @Published var mode: Mode = .merge
    @Published var items: [MediaItem] = []
    @Published var selection = Set<MediaItem.ID>()

    @Published var dvdFolder: URL?
    @Published var titles: [DVDTitle] = []
    @Published var selectedTitle: DVDTitle.ID?

    @Published var encoder: HEVCEncoder = .quality

    @Published var running = false
    @Published var progress: JobProgress?
    @Published var startedAt: Date?
    @Published var statusText = "Drag video files onto the window to get started."
    @Published var results: [URL] = []
    @Published var alert: AlertBox?
    @Published var reencodeOffer: ReencodeOffer?

    private var control = JobControl()

    struct AlertBox: Identifiable {
        let id = UUID()
        let title: String
        let message: String
    }

    /// A merge that only stream-copying refused, held while we ask whether to
    /// spend the time re-encoding it instead.
    struct ReencodeOffer: Identifiable {
        let id = UUID()
        let items: [MediaItem]
        let output: URL
    }

    var totalDuration: Double { items.reduce(0) { $0 + $1.duration } }
    var chosenTitle: DVDTitle? { titles.first { $0.id == selectedTitle } }

    var canStart: Bool {
        guard !running else { return false }
        switch mode {
        case .merge:   return items.count >= 2
        case .convert: return !items.isEmpty
        case .dvd:     return chosenTitle != nil
        }
    }

    var startLabel: String {
        switch mode {
        case .merge:   return "Merge \(items.count) Clips"
        case .convert: return items.count == 1 ? "Convert" : "Convert \(items.count) Files"
        case .dvd:     return "Merge DVD Title"
        }
    }

    // ── Taking files in ──────────────────────────────────────────────────────

    func addFiles() {
        let panel = NSOpenPanel()
        panel.canChooseFiles = true
        panel.canChooseDirectories = true
        panel.allowsMultipleSelection = true
        panel.prompt = "Add"
        panel.message = mode == .dvd ? "Choose the DVD folder or its VIDEO_TS folder"
                                     : "Choose the video files to work on"
        guard panel.runModal() == .OK else { return }
        accept(panel.urls)
    }

    /// Everything dropped on the window or the Dock icon lands here.
    ///
    /// A DVD folder switches to DVD mode by itself rather than being silently
    /// ignored, since a folder of VOBs can't mean anything else.
    func accept(_ urls: [URL]) {
        guard !urls.isEmpty else { return }
        if running {
            statusText = "Busy — wait for the current job to finish before adding more."
            return
        }

        // A dropped DVD wins: it can only mean one thing, and it is the mode
        // people are most likely to get wrong by hand.
        for url in urls where isDirectory(url) {
            if let videoTS = findVideoTS(in: url), !dvdTitles(in: videoTS).isEmpty {
                loadDVD(url)
                return
            }
        }

        let files = urls.flatMap(videoFiles(under:))
        guard !files.isEmpty else {
            statusText = urls.count == 1
                ? "\(urls[0].lastPathComponent) isn't a video file or a DVD folder."
                : "Nothing there was a video file or a DVD folder."
            return
        }
        if mode == .dvd { mode = .merge }
        add(files)
    }

    private func add(_ urls: [URL]) {
        let known = Set(items.map(\.url.standardizedFileURL))
        let fresh = urls.map(\.standardizedFileURL)
            .filter { !known.contains($0) }
        guard !fresh.isEmpty else {
            statusText = "Already on the list."
            return
        }

        // Show them at once, then fill in durations: ffprobe takes a moment per
        // file, and a list that appears instantly beats an accurate one later.
        let placeholders = fresh.map { url -> MediaItem in
            var item = MediaItem(url: url)
            item.bytes = (try? url.resourceValues(forKeys: [.fileSizeKey]))
                .flatMap { $0.fileSize }.map(Int64.init) ?? 0
            return item
        }
        items.append(contentsOf: placeholders)
        results = []
        statusText = "\(items.count) file\(items.count == 1 ? "" : "s") ready — "
            + "drag to reorder, top to bottom is the play order."

        Task.detached(priority: .userInitiated) {
            for placeholder in placeholders {
                let measured = inspect(placeholder.url)
                await MainActor.run {
                    guard let index = self.items.firstIndex(where: { $0.id == placeholder.id })
                    else { return }
                    self.items[index].duration = measured.duration
                    self.items[index].bytes = measured.bytes
                }
            }
        }
    }

    private func loadDVD(_ folder: URL) {
        guard let videoTS = findVideoTS(in: folder) else {
            alert = AlertBox(title: "No DVD video found",
                             message: "No VTS_*.VOB files were found in \(folder.path).\n\n"
                                 + "Pick a ripped DVD folder that contains a VIDEO_TS folder.")
            return
        }
        let found = dvdTitles(in: videoTS)
        guard !found.isEmpty else {
            alert = AlertBox(title: "No DVD video found",
                             message: "\(videoTS.path) has no VTS_*.VOB content files.")
            return
        }
        mode = .dvd
        dvdFolder = videoTS
        titles = found
        selectedTitle = found.first?.id      // largest, usually the main feature
        results = []
        statusText = found.count == 1
            ? "Found one title in \(videoTS.lastPathComponent)."
            : "Found \(found.count) titles — the largest is usually the main feature."
    }

    private func isDirectory(_ url: URL) -> Bool {
        (try? url.resourceValues(forKeys: [.isDirectoryKey]))?.isDirectory ?? false
    }

    /// A dropped file is itself; a dropped folder is the videos directly inside
    /// it, in name order, so dropping a folder of numbered clips does the
    /// obvious thing.
    private func videoFiles(under url: URL) -> [URL] {
        guard isDirectory(url) else {
            return convertibleExtensions.contains(url.pathExtension.lowercased()) ? [url] : []
        }
        let contents = (try? FileManager.default.contentsOfDirectory(
            at: url, includingPropertiesForKeys: [.isDirectoryKey],
            options: [.skipsHiddenFiles])) ?? []
        return contents
            .filter { !isDirectory($0) }
            .filter { convertibleExtensions.contains($0.pathExtension.lowercased()) }
            .sorted { $0.lastPathComponent.localizedStandardCompare($1.lastPathComponent) == .orderedAscending }
    }

    // ── Editing the list ─────────────────────────────────────────────────────

    func move(from offsets: IndexSet, to destination: Int) {
        items.move(fromOffsets: offsets, toOffset: destination)
    }

    func moveSelection(by delta: Int) {
        let indices = items.indices.filter { selection.contains(items[$0].id) }
        guard !indices.isEmpty else { return }
        for index in delta < 0 ? indices : indices.reversed() {
            let target = index + delta
            guard items.indices.contains(target),
                  !selection.contains(items[target].id) else { continue }
            items.swapAt(index, target)
        }
    }

    func removeSelected() {
        items.removeAll { selection.contains($0.id) }
        selection.removeAll()
    }

    func clearAll() {
        items = []
        selection = []
        titles = []
        dvdFolder = nil
        selectedTitle = nil
        results = []
        statusText = "Drag video files onto the window to get started."
    }

    // ── Running ──────────────────────────────────────────────────────────────

    func start() {
        guard canStart else { return }
        guard Tools.ffmpeg != nil else {
            alert = AlertBox(title: "ffmpeg missing",
                             message: "Could not find the ffmpeg engine inside the app.")
            return
        }
        switch mode {
        case .merge:
            guard let output = askOutput(defaultName: "merged.mp4",
                                         allowed: ["mp4", "mov", "m4v"]) else { return }
            begin(.merge(items: items, output: output, allowReencode: false))
        case .dvd:
            guard let title = chosenTitle,
                  let output = askOutput(defaultName: "dvd_movie.mpg",
                                         allowed: ["mpg", "mpeg", "vob"]) else { return }
            begin(.dvd(title: title, output: output))
        case .convert:
            // Each file is written next to its original, so there is nothing to
            // ask: one save panel per file would be worse than no panel at all.
            begin(.convert(items: items, encoder: encoder))
        }
    }

    private func askOutput(defaultName: String, allowed: [String]) -> URL? {
        let panel = NSSavePanel()
        panel.nameFieldStringValue = defaultName
        panel.message = "Save the merged video as:"
        panel.canCreateDirectories = true
        panel.allowedContentTypes = allowed.compactMap { UTType(filenameExtension: $0) }
        guard panel.runModal() == .OK, var url = panel.url else { return nil }
        if !allowed.contains(url.pathExtension.lowercased()) {
            url = url.appendingPathExtension(allowed[0])
        }
        return url
    }

    private func begin(_ job: Job) {
        running = true
        results = []
        startedAt = Date()
        progress = JobProgress(fraction: nil, detail: "Starting…")
        statusText = "Working…"

        let control = JobControl()
        self.control = control

        Task.detached(priority: .userInitiated) {
            let outcome = run(job: job, control: control) { update in
                Task { @MainActor in self.progress = update }
            }
            await MainActor.run { self.finish(job, outcome) }
        }
    }

    func cancel() {
        control.cancel()
        statusText = "Cancelling…"
    }

    private func finish(_ job: Job, _ outcome: JobOutcome) {
        running = false
        progress = nil
        let elapsed = startedAt.map { -$0.timeIntervalSinceNow } ?? 0
        startedAt = nil

        switch outcome {
        case let .finished(outputs):
            results = outputs
            let what = outputs.count == 1 ? outputs[0].lastPathComponent
                                          : "\(outputs.count) files"
            statusText = String(format: "Done in %@ — %@", humanDuration(elapsed), what)
            // The window stays open and ready for the next job, so nudge the
            // Dock instead: a long encode is watched from another app, if at all.
            if !NSApp.isActive { NSApp.requestUserAttention(.informationalRequest) }
            NSSound(named: "Glass")?.play()

        case .needsReencode:
            if case let .merge(items, output, _) = job {
                reencodeOffer = ReencodeOffer(items: items, output: output)
                statusText = "These clips don't share a format."
            } else {
                statusText = "Could not merge."
            }

        case let .failed(message):
            statusText = "Failed."
            alert = AlertBox(title: "Could not finish", message: message)

        case .cancelled:
            statusText = "Cancelled — nothing was written."
        }
    }

    func acceptReencode() {
        guard let offer = reencodeOffer else { return }
        reencodeOffer = nil
        begin(.merge(items: offer.items, output: offer.output, allowReencode: true))
    }

    // ── Results ──────────────────────────────────────────────────────────────

    func reveal(_ url: URL) { NSWorkspace.shared.activateFileViewerSelecting([url]) }
    func open(_ url: URL) { NSWorkspace.shared.open(url) }
}

// ── Main view ────────────────────────────────────────────────────────────────

struct ContentView: View {
    @ObservedObject private var model = MergerModel.shared
    @State private var isDropTarget = false

    var body: some View {
        VStack(alignment: .leading, spacing: 14) {
            header
            Picker("", selection: $model.mode) {
                ForEach(Mode.allCases) { Text($0.label).tag($0) }
            }
            .pickerStyle(.segmented)
            .labelsHidden()
            .disabled(model.running)

            content
            if model.mode == .convert { encoderPicker }
            progressBar
            resultsRow
            footer
        }
        .padding(16)
        .frame(minWidth: 680, minHeight: 540)
        // The whole window is the drop target, so there is nothing to aim at.
        .onDrop(of: [.fileURL], isTargeted: $isDropTarget) { providers in
            acceptDrop(providers)
        }
        .overlay {
            RoundedRectangle(cornerRadius: 10)
                .strokeBorder(Color.accentColor, lineWidth: 3)
                .opacity(isDropTarget ? 1 : 0)
                .allowsHitTesting(false)
                .animation(.easeOut(duration: 0.12), value: isDropTarget)
        }
        .alert(item: $model.alert) { box in
            Alert(title: Text(box.title), message: Text(box.message),
                  dismissButton: .default(Text("OK")))
        }
        .alert(item: $model.reencodeOffer) { _ in
            Alert(
                title: Text("These clips don't share the same format"),
                message: Text("They can't be joined without re-encoding.\n\n"
                    + "Re-encoding produces one clean file at high quality "
                    + "(visually near-identical), but it is slower and technically "
                    + "re-compresses the video."),
                primaryButton: .default(Text("Re-encode (high quality)")) {
                    model.acceptReencode()
                },
                secondaryButton: .cancel())
        }
    }

    /// Dropped URLs arrive one callback at a time and in no particular order,
    /// but for a merge the order they were dragged in is the play order — so
    /// collect them into their original slots before handing them over.
    private func acceptDrop(_ providers: [NSItemProvider]) -> Bool {
        // loadObject(ofClass: URL.self), not loadItem(forTypeIdentifier:) — the
        // latter's completion never fires for a dropped file URL, so the drop
        // silently does nothing.
        let usable = providers.filter { $0.canLoadObject(ofClass: URL.self) }
        guard !usable.isEmpty else { return false }

        let lock = NSLock()
        var slots = [URL?](repeating: nil, count: usable.count)
        let group = DispatchGroup()

        for (index, provider) in usable.enumerated() {
            group.enter()
            _ = provider.loadObject(ofClass: URL.self) { url, _ in
                if let url, FileManager.default.fileExists(atPath: url.path) {
                    lock.withLock { slots[index] = url }
                }
                group.leave()
            }
        }
        group.notify(queue: .main) {
            let urls = lock.withLock { slots.compactMap { $0 } }
            MergerModel.shared.accept(urls)
        }
        return true
    }

    private var header: some View {
        HStack(spacing: 10) {
            Image(nsImage: NSApp.applicationIconImage)
                .resizable().frame(width: 34, height: 34)
            VStack(alignment: .leading, spacing: 1) {
                Text("Video Merger").font(.system(size: 19, weight: .bold))
                Text("Stitch clips end-to-end, join a DVD, or convert to HEVC.")
                    .foregroundStyle(.secondary).font(.system(size: 12))
            }
        }
    }

    @ViewBuilder private var content: some View {
        switch model.mode {
        case .merge, .convert: fileList
        case .dvd: dvdList
        }
    }

    private var fileList: some View {
        GroupBox(model.mode == .merge ? "Clips to join — top to bottom is the play order"
                                      : "Files to convert") {
            HStack(alignment: .top, spacing: 10) {
                List(selection: $model.selection) {
                    ForEach(Array(model.items.enumerated()), id: \.element.id) { index, item in
                        HStack(spacing: 8) {
                            Text("\(index + 1).")
                                .foregroundStyle(.secondary)
                                .frame(width: 24, alignment: .trailing)
                            Text(item.name).lineLimit(1).truncationMode(.middle)
                            Spacer(minLength: 8)
                            Text(humanDuration(item.duration))
                                .foregroundStyle(.secondary).frame(width: 64, alignment: .trailing)
                            Text(humanSize(item.bytes))
                                .foregroundStyle(.secondary).frame(width: 72, alignment: .trailing)
                        }
                        .font(.system(size: 12)).monospacedDigit()
                        .tag(item.id)
                    }
                    .onMove { model.move(from: $0, to: $1) }
                }
                .frame(minHeight: 190)
                .overlay {
                    if model.items.isEmpty { emptyDropHint }
                }

                VStack(spacing: 6) {
                    Button("Add Files…") { model.addFiles() }
                        .keyboardShortcut("o", modifiers: .command)
                    Button("Move Up") { model.moveSelection(by: -1) }
                        .disabled(model.selection.isEmpty || model.mode == .convert)
                    Button("Move Down") { model.moveSelection(by: 1) }
                        .disabled(model.selection.isEmpty || model.mode == .convert)
                    Button("Remove") { model.removeSelected() }
                        .disabled(model.selection.isEmpty)
                    Button("Clear") { model.clearAll() }
                        .disabled(model.items.isEmpty)
                }
                .frame(width: 118)
                .disabled(model.running)
            }
            .padding(4)
        }
    }

    private var emptyDropHint: some View {
        VStack(spacing: 6) {
            Image(systemName: isDropTarget ? "arrow.down.doc.fill" : "film.stack")
                .font(.system(size: 28))
                .foregroundStyle(isDropTarget ? Color.accentColor : Color.secondary)
            Text(isDropTarget ? "Drop to add"
                              : "Drag video files here — or a folder, or a DVD folder")
                .font(.system(size: 12))
                .foregroundStyle(isDropTarget ? Color.accentColor : Color.secondary)
        }
    }

    private var dvdList: some View {
        GroupBox("DVD titles" + (model.dvdFolder.map { " — \($0.path)" } ?? "")) {
            VStack(alignment: .leading, spacing: 8) {
                if model.titles.isEmpty {
                    VStack(spacing: 6) {
                        Image(systemName: "opticaldisc")
                            .font(.system(size: 28)).foregroundStyle(.secondary)
                        Text("Drag a ripped DVD folder (or its VIDEO_TS) onto the window.")
                            .font(.system(size: 12)).foregroundStyle(.secondary)
                        Button("Choose DVD Folder…") { model.addFiles() }
                    }
                    .frame(maxWidth: .infinity, minHeight: 190)
                } else {
                    List(model.titles, selection: $model.selectedTitle) { title in
                        Text(title.label).font(.system(size: 12)).monospacedDigit().tag(title.id)
                    }
                    .frame(minHeight: 160)
                    HStack {
                        Text("The largest title is usually the main feature.")
                            .font(.system(size: 11)).foregroundStyle(.secondary)
                        Spacer()
                        Button("Choose Another DVD…") { model.addFiles() }
                    }
                }
            }
            .padding(4)
            .disabled(model.running)
        }
    }

    private var encoderPicker: some View {
        Picker("Encoder:", selection: $model.encoder) {
            ForEach(HEVCEncoder.allCases) { Text($0.label).tag($0) }
        }
        .disabled(model.running)
        .font(.system(size: 12))
    }

    /// Visible for the whole run: a determinate bar when the input durations are
    /// known, and an indeterminate one when they aren't, rather than nothing.
    @ViewBuilder private var progressBar: some View {
        if let progress = model.progress {
            VStack(alignment: .leading, spacing: 4) {
                HStack(spacing: 10) {
                    if let fraction = progress.fraction {
                        ProgressView(value: fraction).frame(maxWidth: .infinity)
                        Text("\(Int(fraction * 100))%")
                            .font(.system(size: 12, weight: .semibold)).monospacedDigit()
                            .frame(width: 42, alignment: .trailing)
                    } else {
                        ProgressView().progressViewStyle(.linear).frame(maxWidth: .infinity)
                    }
                    Button("Cancel") { model.cancel() }
                        .keyboardShortcut(".", modifiers: .command)
                }
                HStack {
                    Text(progress.detail).lineLimit(1).truncationMode(.middle)
                    Spacer()
                    Text(timingText(progress))
                        .monospacedDigit()
                }
                .font(.system(size: 11)).foregroundStyle(.secondary)
            }
            .padding(10)
            .background(RoundedRectangle(cornerRadius: 8).fill(Color.secondary.opacity(0.10)))
        }
    }

    private func timingText(_ progress: JobProgress) -> String {
        var parts: [String] = []
        if let started = model.startedAt {
            let elapsed = -started.timeIntervalSinceNow
            parts.append("\(humanDuration(elapsed)) elapsed")
            // Only guess at a finish time once there is enough of a run to
            // extrapolate from — an estimate off the first second is noise.
            if let fraction = progress.fraction, fraction > 0.02, elapsed > 2 {
                parts.append("~\(humanDuration(elapsed / fraction - elapsed)) left")
            }
        }
        if !progress.speed.isEmpty { parts.append(progress.speed) }
        return parts.joined(separator: "  ·  ")
    }

    @ViewBuilder private var resultsRow: some View {
        if !model.results.isEmpty {
            GroupBox {
                VStack(alignment: .leading, spacing: 4) {
                    ForEach(model.results, id: \.self) { url in
                        HStack(spacing: 8) {
                            Image(systemName: "checkmark.circle.fill").foregroundStyle(.green)
                            Text(url.path).font(.system(size: 12))
                                .lineLimit(1).truncationMode(.head)
                            Spacer(minLength: 8)
                            Button("Show in Finder") { model.reveal(url) }.buttonStyle(.link)
                            Button("Open") { model.open(url) }.buttonStyle(.link)
                        }
                        .font(.system(size: 11))
                    }
                }
                .padding(4)
            }
        }
    }

    private var footer: some View {
        HStack(spacing: 12) {
            Text(model.statusText)
                .font(.system(size: 12)).foregroundStyle(.secondary)
                .lineLimit(2).fixedSize(horizontal: false, vertical: true)
            Spacer(minLength: 8)
            if model.mode != .dvd, !model.items.isEmpty {
                Text(humanDuration(model.totalDuration) + " total")
                    .font(.system(size: 11)).foregroundStyle(.secondary).monospacedDigit()
            }
            Button {
                model.start()
            } label: {
                Label(model.startLabel, systemImage: "film.stack")
            }
            .keyboardShortcut(.return, modifiers: .command)
            .disabled(!model.canStart)
        }
    }
}

// ── App ──────────────────────────────────────────────────────────────────────

final class AppDelegate: NSObject, NSApplicationDelegate {
    /// Files dropped onto the Dock icon, or opened with "Open With".
    func application(_ sender: NSApplication, open urls: [URL]) {
        Task { @MainActor in MergerModel.shared.accept(urls) }
    }

    /// Finishing a job is not a reason to quit, and neither is closing the
    /// window — the point of the rebuild is that it is still there afterwards.
    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { false }

    func applicationShouldHandleReopen(_ sender: NSApplication, hasVisibleWindows: Bool) -> Bool {
        true
    }
}

@main
struct VideoMergerApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) private var delegate

    var body: some Scene {
        WindowGroup("Video Merger") {
            ContentView()
        }
        .windowResizability(.contentMinSize)
        .commands {
            CommandGroup(replacing: .newItem) {}   // nothing sensible to open anew
        }
    }
}
