"""Modern Tkinter GUI (ttkbootstrap) for the duplicate video finder."""
from __future__ import annotations

import queue
import shutil
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as ttk
from send2trash import send2trash
from ttkbootstrap.constants import DANGER, OUTLINE, PRIMARY, SECONDARY, SUCCESS

import scanner

MARK_ON = "✓"
MARK_OFF = ""
SPINNER_FRAMES = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"


def system_is_dark_mode() -> bool:
    try:
        result = subprocess.run(
            ["defaults", "read", "-g", "AppleInterfaceStyle"],
            capture_output=True, text=True, timeout=2,
        )
        return result.stdout.strip().lower() == "dark"
    except (OSError, subprocess.TimeoutExpired):
        return False


def human_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} PB"


def human_bitrate(bps: float | None) -> str:
    if not bps:
        return "?"
    if bps >= 1_000_000:
        return f"{bps / 1_000_000:.1f} Mbps"
    return f"{bps / 1_000:.0f} kbps"


def human_duration(seconds: float | None) -> str:
    if not seconds:
        return "?"
    seconds = int(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:d}:{m:02d}:{s:02d}" if h else f"{m:d}:{s:02d}"


class App:
    def __init__(self, root: ttk.Window):
        self.root = root
        self.root.title("Duplicate Video Finder")
        self.root.geometry("1060x700")
        self.root.minsize(820, 520)

        self.folders: list[Path] = []
        self.progress_queue: queue.Queue = queue.Queue()
        self.file_marks: dict[str, bool] = {}
        self.group_files: dict[str, scanner.VideoFile] = {}
        self.recommended: set[str] = set()  # row ids the scanner suggests keeping
        self.scanning = False
        self.spinner_i = 0

        self._build_ui()
        self._check_tools()

    def _nudge_redraw(self):
        # macOS Tk sometimes fails to repaint the window after a native
        # dialog (file picker / alert) closes; a no-op resize forces it.
        self.root.update_idletasks()
        w, h = self.root.winfo_width(), self.root.winfo_height()
        self.root.geometry(f"{w}x{h + 1}")
        self.root.update_idletasks()
        self.root.geometry(f"{w}x{h}")

    def _info(self, title, message):
        messagebox.showinfo(title, message)
        self._nudge_redraw()

    def _warn(self, title, message):
        messagebox.showwarning(title, message)
        self._nudge_redraw()

    def _error(self, title, message):
        messagebox.showerror(title, message)
        self._nudge_redraw()

    def _confirm(self, title, message) -> bool:
        result = messagebox.askyesno(title, message)
        self._nudge_redraw()
        return result

    def _check_tools(self):
        missing = [t for t in ("ffmpeg", "ffprobe") if not shutil.which(t)]
        if missing:
            self._warn(
                "Missing tools",
                "Could not find: " + ", ".join(missing) + ".\n\n"
                "Exact-duplicate detection will still work, but detecting "
                "re-encoded/renamed duplicates needs ffmpeg (install with "
                "'brew install ffmpeg').",
            )

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)

        # -- header ---------------------------------------------------------
        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(0, 12))
        ttk.Label(header, text="Duplicate Video Finder", font=("SF Pro Display", 20, "bold")).pack(anchor="w")
        ttk.Label(
            header,
            text="Find duplicate videos across folders — even renamed or re-encoded copies.",
            bootstyle=SECONDARY,
        ).pack(anchor="w")

        # -- folders card -----------------------------------------------------
        folders_card = ttk.Labelframe(outer, text="Folders to search", padding=12)
        folders_card.pack(fill="x", pady=(0, 12))

        folder_row = ttk.Frame(folders_card)
        folder_row.pack(fill="x")

        list_frame = ttk.Frame(folder_row)
        list_frame.pack(side="left", fill="both", expand=True)
        self.folder_list = tk.Listbox(
            list_frame, height=4, relief="flat", highlightthickness=1,
            borderwidth=0, activestyle="none",
        )
        self._theme_listbox()
        self.folder_list.pack(side="left", fill="both", expand=True)
        scroll = tk.Scrollbar(list_frame, command=self.folder_list.yview)
        scroll.pack(side="left", fill="y")
        self.folder_list.config(yscrollcommand=scroll.set)

        btn_col = ttk.Frame(folder_row)
        btn_col.pack(side="left", padx=(10, 0), fill="y")
        ttk.Button(btn_col, text="+ Add Folder…", bootstyle=PRIMARY, command=self.add_folder).pack(fill="x")
        ttk.Button(
            btn_col, text="Remove Selected", bootstyle=(SECONDARY, OUTLINE), command=self.remove_folder
        ).pack(fill="x", pady=(6, 0))

        # -- scan controls ----------------------------------------------------
        scan_card = ttk.Frame(outer)
        scan_card.pack(fill="x", pady=(0, 12))

        scan_row = ttk.Frame(scan_card)
        scan_row.pack(fill="x")
        self.scan_btn = ttk.Button(
            scan_row, text="🔍  Scan for Duplicates", bootstyle=SUCCESS, command=self.start_scan
        )
        self.scan_btn.pack(side="left", ipadx=8, ipady=4)
        self.percent_label = ttk.Label(scan_row, text="", font=("SF Pro Display", 11, "bold"))
        self.percent_label.pack(side="right")

        self.progress = ttk.Progressbar(scan_card, mode="determinate", bootstyle=SUCCESS)
        self.progress.pack(fill="x", pady=(10, 4))

        self.status_label = ttk.Label(scan_card, text="Ready.", bootstyle=SECONDARY)
        self.status_label.pack(anchor="w")

        # -- results ------------------------------------------------------------
        results_card = ttk.Labelframe(outer, text="Results", padding=8)
        results_card.pack(fill="both", expand=True)

        style = ttk.Style()
        style.configure("Treeview", rowheight=26)

        tree_frame = ttk.Frame(results_card)
        tree_frame.pack(fill="both", expand=True)

        columns = ("mark", "size", "resolution", "bitrate", "duration", "codec", "folder")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", bootstyle=PRIMARY)
        self.tree.heading("#0", text="File  (double-click to play)")
        self.tree.heading("mark", text="Delete?")
        self.tree.heading("size", text="Size")
        self.tree.heading("resolution", text="Resolution")
        self.tree.heading("bitrate", text="Bitrate")
        self.tree.heading("duration", text="Duration")
        self.tree.heading("codec", text="Codec")
        self.tree.heading("folder", text="Location")

        # widths sum to just under the default window width so nothing clips
        self.tree.column("#0", width=360)  # fits the group summary line
        self.tree.column("mark", width=60, anchor="center")
        self.tree.column("size", width=80, anchor="e")
        self.tree.column("resolution", width=90, anchor="center")
        self.tree.column("bitrate", width=85, anchor="e")
        self.tree.column("duration", width=70, anchor="center")
        self.tree.column("codec", width=60, anchor="center")
        self.tree.column("folder", width=180)

        # explicit foregrounds: these pale backgrounds need dark text in both themes
        self.tree.tag_configure("marked", background="#f8d7da", foreground="#721c24")
        self.tree.tag_configure("keep", background="#d4edda", foreground="#155724")
        self.tree.tag_configure("kept", background="", foreground="")
        self.tree.tag_configure("group_exact", font=("SF Pro Display", 10, "bold"))
        self.tree.tag_configure("group_possible", font=("SF Pro Display", 10, "bold"))

        tree_scroll = tk.Scrollbar(tree_frame, command=self.tree.yview)
        self.tree.config(yscrollcommand=tree_scroll.set)
        tree_scroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-Button-1>", self.on_tree_double_click)

        # -- footer ---------------------------------------------------------
        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(12, 0))
        self.summary_label = ttk.Label(bottom, text="No scan run yet.", bootstyle=SECONDARY)
        self.summary_label.pack(side="left")
        ttk.Button(
            bottom, text="🗑  Send Marked Files to Trash", bootstyle=DANGER, command=self.delete_marked
        ).pack(side="right", ipadx=6, ipady=2)

    def _theme_listbox(self):
        style = ttk.Style()
        colors = style.colors
        self.folder_list.configure(
            bg=colors.inputbg, fg=colors.inputfg,
            selectbackground=colors.primary, selectforeground=colors.selectfg,
            highlightbackground=colors.border, highlightcolor=colors.primary,
        )

    # -- folder management -------------------------------------------------
    def add_folder(self):
        chosen = filedialog.askdirectory(title="Select a folder to search")
        self._nudge_redraw()
        if chosen:
            path = Path(chosen)
            if path not in self.folders:
                self.folders.append(path)
                self.folder_list.insert("end", str(path))

    def remove_folder(self):
        for idx in reversed(self.folder_list.curselection()):
            del self.folders[idx]
            self.folder_list.delete(idx)

    # -- scanning ------------------------------------------------------------
    def start_scan(self):
        if self.scanning:
            return
        if not self.folders:
            self._info("No folders", "Add at least one folder to search first.")
            return

        self.scanning = True
        self.scan_btn.config(state="disabled", text="Scanning…")
        self.tree.delete(*self.tree.get_children())
        self.file_marks.clear()
        self.group_files.clear()
        self.recommended.clear()
        self.percent_label.config(text="")
        self.status_label.config(text="Starting scan...")
        self.progress.config(mode="indeterminate")
        self.progress.start(12)

        def progress_cb(done, total, message):
            self.progress_queue.put(("progress", done, total, message))

        def worker():
            try:
                groups = scanner.scan(list(self.folders), progress=progress_cb)
                self.progress_queue.put(("done", groups))
            except Exception as exc:  # surface any scan failure to the UI
                self.progress_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(100, self._poll_queue)

    def _poll_queue(self):
        try:
            while True:
                item = self.progress_queue.get_nowait()
                if item[0] == "progress":
                    _, done, total, message = item
                    if total > 0:
                        if str(self.progress["mode"]) != "determinate":
                            self.progress.stop()
                            self.progress.config(mode="determinate")
                        self.progress.config(value=done, maximum=total)
                        self.percent_label.config(text=f"{done}/{total} · {done / total * 100:.0f}%")
                    else:
                        if str(self.progress["mode"]) != "indeterminate":
                            self.progress.config(mode="indeterminate")
                            self.progress.start(12)
                        self.percent_label.config(text="")
                    self.status_label.config(text=message)
                elif item[0] == "done":
                    self._on_scan_complete(item[1])
                    return
                elif item[0] == "error":
                    self.scanning = False
                    self.progress.stop()
                    self.scan_btn.config(state="normal", text="🔍  Scan for Duplicates")
                    self._error("Scan failed", item[1])
                    return
        except queue.Empty:
            pass
        if self.scanning:
            self.spinner_i = (self.spinner_i + 1) % len(SPINNER_FRAMES)
            spinner = SPINNER_FRAMES[self.spinner_i]
            current = self.status_label.cget("text")
            base = current[2:] if current[:1] in SPINNER_FRAMES else current
            self.status_label.config(text=f"{spinner} {base}")
            self.root.after(100, self._poll_queue)

    def _on_scan_complete(self, groups: list):
        self.scanning = False
        self.progress.stop()
        self.progress.config(mode="determinate", value=0)
        self.percent_label.config(text="")
        self.scan_btn.config(state="normal", text="🔍  Scan for Duplicates")
        self.status_label.config(text=f"Scan complete — {len(groups)} duplicate group(s) found.")

        total_wasted = 0
        for gi, group in enumerate(groups, 1):
            label = "Exact duplicate" if group.kind == "exact" else "Possible duplicate"
            wasted = group.wasted_bytes
            total_wasted += wasted
            group_id = f"group{gi}"
            tag = "group_exact" if group.kind == "exact" else "group_possible"
            best = group.best_copy
            self.tree.insert(
                "", "end", iid=group_id, open=True, tags=(tag,),
                text=f"Group {gi} — {label} — recover {human_size(wasted)} "
                     f"— keep ★ {group.keep_reason}",
                values=("", "", "", "", "", "", ""),
            )
            for vf in group.files:
                file_id = f"{group_id}_{vf.path}"
                recommended = vf is best
                self.file_marks[file_id] = not recommended
                self.group_files[file_id] = vf
                if recommended:
                    self.recommended.add(file_id)
                self.tree.insert(
                    group_id, "end", iid=file_id,
                    text=("★ " if recommended else "    ") + vf.path.name,
                    tags=("keep",) if recommended else ("marked",),
                    values=(
                        MARK_OFF if recommended else MARK_ON,
                        human_size(vf.size),
                        vf.resolution,
                        human_bitrate(vf.bitrate_bps),
                        human_duration(vf.duration),
                        vf.codec or "?",
                        str(vf.path.parent),
                    ),
                )

        if not groups:
            self.summary_label.config(text="No duplicates found.")
        else:
            self.summary_label.config(
                text=f"{len(groups)} group(s) — up to {human_size(total_wasted)} "
                     "recoverable if you keep one copy per group."
            )

    # -- marking / deleting ---------------------------------------------------
    def on_tree_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not row_id or row_id not in self.file_marks:
            return
        if col != "#1":  # the "mark" column
            return
        marked = not self.file_marks[row_id]
        self.file_marks[row_id] = marked
        values = list(self.tree.item(row_id, "values"))
        values[0] = MARK_ON if marked else MARK_OFF
        if marked:
            tag = "marked"
        else:
            tag = "keep" if row_id in self.recommended else "kept"
        self.tree.item(row_id, values=values, tags=(tag,))

    def on_tree_double_click(self, event):
        """Open the clicked video in the system default player so you can
        eyeball whether two files really are the same content."""
        row_id = self.tree.identify_row(event.y)
        if not row_id or row_id not in self.group_files:
            return
        if self.tree.identify_column(event.x) == "#1":
            return  # that column is the delete checkbox, not a play target
        path = self.group_files[row_id].path
        if not path.exists():
            self._error("File missing", f"{path} no longer exists.")
            return
        try:
            subprocess.run(["open", str(path)], check=True)
        except (OSError, subprocess.CalledProcessError) as exc:
            self._error("Could not open file", f"{path.name}\n\n{exc}")

    def delete_marked(self):
        marked_ids = [rid for rid, marked in self.file_marks.items() if marked]
        if not marked_ids:
            self._info("Nothing marked", "No files are marked for deletion.")
            return

        total_size = sum(self.group_files[rid].size for rid in marked_ids)
        if not self._confirm(
            "Confirm",
            f"Send {len(marked_ids)} file(s) totaling {human_size(total_size)} to the Trash?",
        ):
            return

        failures = []
        for rid in marked_ids:
            vf = self.group_files[rid]
            try:
                send2trash(str(vf.path))
                group_id = self.tree.parent(rid)
                self.tree.delete(rid)
                del self.file_marks[rid]
                del self.group_files[rid]
                self.recommended.discard(rid)
                if group_id and len(self.tree.get_children(group_id)) <= 1:
                    for child in self.tree.get_children(group_id):
                        self.file_marks.pop(child, None)
                        self.group_files.pop(child, None)
                        self.recommended.discard(child)
                    self.tree.delete(group_id)
            except OSError as exc:
                failures.append(f"{vf.path.name}: {exc}")

        if failures:
            self._error("Some files could not be deleted", "\n".join(failures))
        else:
            self._info("Done", "Marked files were moved to the Trash.")


def main():
    themename = "darkly" if system_is_dark_mode() else "flatly"
    root = ttk.Window(themename=themename)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
