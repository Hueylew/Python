import AppKit
import SwiftUI

// ── View model ───────────────────────────────────────────────────────────────
@MainActor
final class ScanModel: ObservableObject {
    @Published var folders: [URL] = []
    @Published var groups: [DuplicateGroup] = []
    @Published var marked: Set<UUID> = []
    @Published var recommended: Set<UUID> = []

    @Published var scanning = false
    @Published var statusText = "Ready."
    @Published var progressDone = 0
    @Published var progressTotal = 0
    @Published var didScan = false

    @Published var alert: AlertBox?

    private var cancelToken = CancelToken()

    struct AlertBox: Identifiable {
        let id = UUID()
        let title: String
        let message: String
    }

    var totalWasted: Int64 { groups.reduce(0) { $0 + $1.wastedBytes } }

    var markedFiles: [VideoFile] {
        groups.flatMap { $0.files }.filter { marked.contains($0.id) }
    }

    // ── Folders ──────────────────────────────────────────────────────────────
    func addFolder() {
        let panel = NSOpenPanel()
        panel.canChooseDirectories = true
        panel.canChooseFiles = false
        panel.allowsMultipleSelection = true
        panel.prompt = "Add"
        panel.message = "Choose folders to search for duplicate videos"
        guard panel.runModal() == .OK else { return }
        for url in panel.urls where !folders.contains(url) { folders.append(url) }
    }

    func removeFolders(_ selection: Set<URL>) {
        folders.removeAll { selection.contains($0) }
    }

    // ── Scanning ─────────────────────────────────────────────────────────────
    func startScan() {
        guard !scanning else { return }
        guard !folders.isEmpty else {
            alert = AlertBox(title: "No folders",
                             message: "Add at least one folder to search first.")
            return
        }
        if !Tools.missing.isEmpty {
            alert = AlertBox(title: "Missing tools",
                             message: "Could not find: \(Tools.missing.joined(separator: ", ")).")
            return
        }

        scanning = true
        didScan = false
        groups = []
        marked = []
        recommended = []
        progressDone = 0
        progressTotal = 0
        statusText = "Starting scan…"

        let token = CancelToken()
        cancelToken = token
        let targets = folders

        Task.detached(priority: .userInitiated) {
            let found = scan(folders: targets, cancel: token) { done, total, message in
                Task { @MainActor in
                    self.progressDone = done
                    self.progressTotal = total
                    self.statusText = message
                }
            }
            await MainActor.run { self.finishScan(found, cancelled: token.isCancelled) }
        }
    }

    func cancelScan() {
        cancelToken.cancel()
        statusText = "Cancelling…"
    }

    private func finishScan(_ found: [DuplicateGroup], cancelled: Bool) {
        scanning = false
        progressDone = 0
        progressTotal = 0
        groups = found
        didScan = !cancelled

        for g in found {
            recommended.insert(g.bestCopy.id)
            for f in g.files.dropFirst() { marked.insert(f.id) }
        }

        statusText = cancelled
            ? "Scan cancelled."
            : "Scan complete — \(found.count) duplicate group(s) found."
    }

    // ── Marking / acting ─────────────────────────────────────────────────────
    func toggleMark(_ file: VideoFile) {
        if marked.contains(file.id) { marked.remove(file.id) } else { marked.insert(file.id) }
    }

    func play(_ file: VideoFile) {
        guard FileManager.default.fileExists(atPath: file.url.path) else {
            alert = AlertBox(title: "File missing",
                             message: "\(file.url.path) no longer exists.")
            return
        }
        NSWorkspace.shared.open(file.url)
    }

    func revealInFinder(_ file: VideoFile) {
        NSWorkspace.shared.activateFileViewerSelecting([file.url])
    }

    func trashMarked() {
        let targets = markedFiles
        guard !targets.isEmpty else {
            alert = AlertBox(title: "Nothing marked",
                             message: "No files are marked for deletion.")
            return
        }

        let total = targets.reduce(Int64(0)) { $0 + $1.size }
        let confirm = NSAlert()
        confirm.messageText = "Move \(targets.count) file(s) to the Trash?"
        confirm.informativeText = "That frees \(humanSize(total)). "
            + "They go to the Trash, so you can still get them back."
        confirm.addButton(withTitle: "Move to Trash")
        confirm.addButton(withTitle: "Cancel")
        confirm.alertStyle = .warning
        guard confirm.runModal() == .alertFirstButtonReturn else { return }

        var removed = Set<UUID>()
        var failures: [String] = []
        var noTrash: [VideoFile] = []   // volume has no Trash (network/exFAT shares)

        for f in targets {
            do {
                try FileManager.default.trashItem(at: f.url, resultingItemURL: nil)
                removed.insert(f.id)
            } catch let e as NSError
                where e.domain == NSCocoaErrorDomain && e.code == NSFeatureUnsupportedError {
                noTrash.append(f)
            } catch {
                failures.append("\(f.name): \(error.localizedDescription)")
            }
        }

        var note: String?
        if !noTrash.isEmpty {
            let outcome = handleVolumeWithoutTrash(noTrash)
            removed.formUnion(outcome.removed)
            failures.append(contentsOf: outcome.failures)
            note = outcome.note
        }

        purgeFromResults(removed)

        var lines: [String] = []
        if removed.count > 0 { lines.append("\(removed.count) file(s) removed.") }
        if let note { lines.append(note) }
        if !failures.isEmpty {
            lines.append("")
            lines.append("Could not remove \(failures.count):")
            lines.append(contentsOf: failures.prefix(15))
            if failures.count > 15 { lines.append("…and \(failures.count - 15) more.") }
        }
        alert = AlertBox(title: failures.isEmpty ? "Done" : "Finished with problems",
                         message: lines.joined(separator: "\n"))
    }

    /// Some volumes — network shares, and many exFAT/FAT disks — have no Trash
    /// at all, so `trashItem` refuses. Rather than silently deleting for real,
    /// offer to move the files aside into a folder on that same volume: it's an
    /// instant rename rather than a copy, and it stays undoable.
    private func handleVolumeWithoutTrash(
        _ files: [VideoFile]
    ) -> (removed: Set<UUID>, failures: [String], note: String?) {
        let volumes = Set(files.map { volumeName(for: $0.url) }).sorted()
        let volumeList = volumes.joined(separator: ", ")
        let folderName = "Duplicate Video Finder — To Delete"

        let choice = NSAlert()
        choice.messageText = "\(files.count) file(s) can't go to the Trash"
        choice.informativeText = """
            The volume \(volumeList) has no Trash — that's normal for network \
            shares and most exFAT drives.

            Moving them into a “\(folderName)” folder on that same volume is \
            instant (nothing is copied) and you can still change your mind. \
            Delete that folder yourself once you're happy.
            """
        choice.addButton(withTitle: "Move to Folder on \(volumes.count == 1 ? volumeList : "Each Volume")")
        choice.addButton(withTitle: "Leave Them Alone")
        choice.addButton(withTitle: "Delete Permanently")
        choice.alertStyle = .warning

        switch choice.runModal() {
        case .alertFirstButtonReturn:
            return relocate(files, intoFolderNamed: folderName)
        case .alertThirdButtonReturn:
            let sure = NSAlert()
            sure.messageText = "Permanently delete \(files.count) file(s)?"
            sure.informativeText = "This cannot be undone — they do not go to the Trash."
            sure.addButton(withTitle: "Cancel")
            sure.addButton(withTitle: "Delete Permanently")
            sure.alertStyle = .critical
            guard sure.runModal() == .alertSecondButtonReturn else {
                return ([], [], "Left \(files.count) file(s) where they were.")
            }
            return deleteForever(files)
        default:
            return ([], [], "Left \(files.count) file(s) where they were.")
        }
    }

    private func relocate(
        _ files: [VideoFile], intoFolderNamed folderName: String
    ) -> (removed: Set<UUID>, failures: [String], note: String?) {
        var removed = Set<UUID>()
        var failures: [String] = []
        var destinations = Set<String>()
        let fm = FileManager.default

        for f in files {
            guard let root = volumeRoot(for: f.url) else {
                failures.append("\(f.name): could not work out which volume it's on")
                continue
            }
            let folder = root.appendingPathComponent(folderName, isDirectory: true)
            do {
                try fm.createDirectory(at: folder, withIntermediateDirectories: true)
                try fm.moveItem(at: f.url, to: uniqueDestination(in: folder, for: f.url))
                removed.insert(f.id)
                destinations.insert(folder.path)
            } catch {
                failures.append("\(f.name): \(error.localizedDescription)")
            }
        }

        let note = destinations.isEmpty ? nil
            : "Moved aside into:\n" + destinations.sorted().joined(separator: "\n")
        return (removed, failures, note)
    }

    private func deleteForever(
        _ files: [VideoFile]
    ) -> (removed: Set<UUID>, failures: [String], note: String?) {
        var removed = Set<UUID>()
        var failures: [String] = []
        for f in files {
            do {
                try FileManager.default.removeItem(at: f.url)
                removed.insert(f.id)
            } catch {
                failures.append("\(f.name): \(error.localizedDescription)")
            }
        }
        return (removed, failures, removed.isEmpty ? nil
                : "\(removed.count) file(s) deleted permanently.")
    }

    /// Never overwrite something already sitting in the destination folder.
    private func uniqueDestination(in folder: URL, for source: URL) -> URL {
        let fm = FileManager.default
        let ext = source.pathExtension
        let stem = source.deletingPathExtension().lastPathComponent
        var candidate = folder.appendingPathComponent(source.lastPathComponent)
        var n = 2
        while fm.fileExists(atPath: candidate.path) {
            let name = ext.isEmpty ? "\(stem) (\(n))" : "\(stem) (\(n)).\(ext)"
            candidate = folder.appendingPathComponent(name)
            n += 1
        }
        return candidate
    }

    private func volumeRoot(for url: URL) -> URL? {
        (try? url.resourceValues(forKeys: [.volumeURLKey]))?.volume
    }

    private func volumeName(for url: URL) -> String {
        (try? url.resourceValues(forKeys: [.volumeNameKey]))?.volumeName ?? "that drive"
    }

    private func purgeFromResults(_ ids: Set<UUID>) {
        guard !ids.isEmpty else { return }
        for i in groups.indices { groups[i].files.removeAll { ids.contains($0.id) } }
        groups.removeAll { $0.files.count < 2 }
        marked.subtract(ids)
    }
}

// ── Row ──────────────────────────────────────────────────────────────────────
struct FileRow: View {
    let file: VideoFile
    let isMarked: Bool
    let isRecommended: Bool
    let onToggle: () -> Void
    let onPlay: () -> Void
    let onReveal: () -> Void

    var body: some View {
        HStack(spacing: 8) {
            Button(action: onToggle) {
                Image(systemName: isMarked ? "checkmark.square.fill" : "square")
                    .foregroundStyle(isMarked ? .red : .secondary)
            }
            .buttonStyle(.plain)
            .help(isMarked ? "Marked for deletion — click to keep" : "Click to mark for deletion")

            Image(systemName: isRecommended ? "star.fill" : "star")
                .foregroundStyle(isRecommended ? .green : .clear)
                .font(.caption)

            Text(file.name)
                .lineLimit(1)
                .truncationMode(.middle)
                .frame(minWidth: 180, alignment: .leading)

            Spacer(minLength: 4)

            Text(humanSize(file.size)).frame(width: 76, alignment: .trailing)
            Text(file.resolution).frame(width: 88, alignment: .trailing)
            Text(humanBitrate(file.bitrate)).frame(width: 86, alignment: .trailing)
            Text(humanDuration(file.duration)).frame(width: 62, alignment: .trailing)
            Text(file.codec ?? "?").frame(width: 52, alignment: .trailing)
            Text(file.folder)
                .lineLimit(1)
                .truncationMode(.head)
                .foregroundStyle(.secondary)
                .frame(width: 200, alignment: .leading)
        }
        .font(.system(size: 12))
        .monospacedDigit()
        .padding(.vertical, 2)
        .contentShape(Rectangle())
        .onTapGesture(count: 2) { onPlay() }
        .contextMenu {
            Button("Play") { onPlay() }
            Button("Reveal in Finder") { onReveal() }
            Divider()
            Button(isMarked ? "Keep this one" : "Mark for deletion") { onToggle() }
        }
    }
}

// ── Main view ────────────────────────────────────────────────────────────────
struct ContentView: View {
    @StateObject private var model = ScanModel()
    @State private var folderSelection = Set<URL>()

    var body: some View {
        VStack(alignment: .leading, spacing: 14) {
            header
            foldersCard
            scanControls
            resultsCard
            footer
        }
        .padding(16)
        .frame(minWidth: 940, minHeight: 620)
        .alert(item: $model.alert) { box in
            Alert(title: Text(box.title), message: Text(box.message),
                  dismissButton: .default(Text("OK")))
        }
    }

    private var header: some View {
        VStack(alignment: .leading, spacing: 2) {
            Text("Duplicate Video Finder").font(.system(size: 20, weight: .bold))
            Text("Find duplicate videos across folders — even renamed or re-encoded copies.")
                .foregroundStyle(.secondary).font(.system(size: 12))
        }
    }

    private var foldersCard: some View {
        GroupBox("Folders to search") {
            HStack(alignment: .top, spacing: 10) {
                List(selection: $folderSelection) {
                    ForEach(model.folders, id: \.self) { url in
                        Text(url.path).font(.system(size: 12)).tag(url)
                    }
                }
                .frame(height: 78)
                .overlay {
                    if model.folders.isEmpty {
                        Text("No folders added yet")
                            .foregroundStyle(.tertiary).font(.system(size: 12))
                    }
                }

                VStack(spacing: 6) {
                    Button("Add Folder…") { model.addFolder() }
                        .keyboardShortcut("o", modifiers: .command)
                    Button("Remove Selected") {
                        model.removeFolders(folderSelection)
                        folderSelection.removeAll()
                    }
                    .disabled(folderSelection.isEmpty)
                }
                .frame(width: 130)
            }
            .padding(4)
        }
    }

    private var scanControls: some View {
        VStack(alignment: .leading, spacing: 6) {
            HStack(spacing: 12) {
                if model.scanning {
                    Button("Cancel") { model.cancelScan() }
                        .keyboardShortcut(".", modifiers: .command)
                } else {
                    Button {
                        model.startScan()
                    } label: {
                        Label("Scan for Duplicates", systemImage: "magnifyingglass")
                    }
                    .keyboardShortcut(.return, modifiers: .command)
                    .disabled(model.folders.isEmpty)
                }

                if model.progressTotal > 0 {
                    ProgressView(value: Double(model.progressDone),
                                 total: Double(model.progressTotal))
                        .frame(maxWidth: .infinity)
                    Text("\(model.progressDone)/\(model.progressTotal) · "
                         + "\(Int(Double(model.progressDone) / Double(model.progressTotal) * 100))%")
                        .font(.system(size: 11, weight: .semibold)).monospacedDigit()
                } else if model.scanning {
                    ProgressView().progressViewStyle(.linear).frame(maxWidth: .infinity)
                } else {
                    Spacer()
                }
            }
            Text(model.statusText)
                .font(.system(size: 11)).foregroundStyle(.secondary)
                .lineLimit(1).truncationMode(.middle)
        }
    }

    private var resultsCard: some View {
        GroupBox("Results") {
            if model.groups.isEmpty {
                VStack {
                    Spacer()
                    Text(model.didScan ? "No duplicates found."
                                       : "Add folders, then scan. Double-click any file to play it.")
                        .foregroundStyle(.secondary)
                    Spacer()
                }
                .frame(maxWidth: .infinity, maxHeight: .infinity)
            } else {
                List {
                    ForEach(Array(model.groups.enumerated()), id: \.element.id) { index, group in
                        Section {
                            ForEach(group.files) { file in
                                FileRow(file: file,
                                        isMarked: model.marked.contains(file.id),
                                        isRecommended: model.recommended.contains(file.id),
                                        onToggle: { model.toggleMark(file) },
                                        onPlay: { model.play(file) },
                                        onReveal: { model.revealInFinder(file) })
                                    .listRowBackground(rowColour(file))
                            }
                        } header: {
                            Text("Group \(index + 1) — \(group.kind.label) — recover "
                                 + "\(humanSize(group.wastedBytes)) — keep ★ \(group.keepReason)")
                                .font(.system(size: 11, weight: .semibold))
                        }
                    }
                }
                .listStyle(.inset(alternatesRowBackgrounds: false))
            }
        }
        .frame(maxHeight: .infinity)
    }

    private func rowColour(_ file: VideoFile) -> Color {
        if model.marked.contains(file.id) { return Color.red.opacity(0.13) }
        if model.recommended.contains(file.id) { return Color.green.opacity(0.13) }
        return Color.clear
    }

    private var footer: some View {
        HStack {
            Text(summaryText).font(.system(size: 12)).foregroundStyle(.secondary)
            Spacer()
            Button(role: .destructive) {
                model.trashMarked()
            } label: {
                Label("Move Marked to Trash", systemImage: "trash")
            }
            .disabled(model.markedFiles.isEmpty)
        }
    }

    private var summaryText: String {
        if model.groups.isEmpty { return model.didScan ? "No duplicates found." : "No scan run yet." }
        return "\(model.groups.count) group(s) — up to \(humanSize(model.totalWasted)) "
             + "recoverable if you keep one copy per group."
    }
}

// ── App ──────────────────────────────────────────────────────────────────────
@main
struct DuplicateVideoFinderApp: App {
    var body: some Scene {
        WindowGroup("Duplicate Video Finder") {
            ContentView()
        }
        .windowResizability(.contentMinSize)
    }
}
