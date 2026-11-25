"""FileFinder: type part of a file name, search every drive. Double-click a result to open it."""
import os, string, subprocess, threading, queue, datetime, tkinter as tk
from tkinter import ttk

SKIP = {"$recycle.bin", "system volume information", "windows", "node_modules", ".git", "__pycache__"}


def drives():
    return [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]


def search(term, roots, stop, out):
    """Walk every root; put (path, is_dir) for names containing term (case-insensitive)."""
    term = term.lower()
    for root in roots:
        for cur, dirs, files in os.walk(root, onerror=lambda e: None):
            if stop.is_set():
                return
            dirs[:] = [d for d in dirs if d.lower() not in SKIP]
            for name in dirs:
                if term in name.lower():
                    out.put((os.path.join(cur, name), True))
            for name in files:
                if term in name.lower():
                    out.put((os.path.join(cur, name), False))


def human(n):
    for u in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.0f} {u}"
        n /= 1024
    return f"{n:.1f} TB"


class App:
    def __init__(self, root):
        self.root, self.stop, self.q, self.count = root, threading.Event(), queue.Queue(), 0
        root.title("FileFinder")
        root.geometry("1000x560")
        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Find:").pack(side="left")
        self.entry = ttk.Entry(top, font=("Segoe UI", 12))
        self.entry.pack(side="left", fill="x", expand=True, padx=6)
        self.entry.bind("<Return>", lambda e: self.start())
        root.bind("<Escape>", lambda e: root.destroy())
        ttk.Button(top, text="Search", command=self.start).pack(side="left")
        ttk.Button(top, text="Stop", command=self.stop.set).pack(side="left", padx=4)

        cols = ("name", "type", "size", "modified", "created", "path")
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        for c, w in zip(cols, (220, 70, 80, 130, 130, 450)):
            self.tree.heading(c, text=c.title())
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=8)
        self.tree.bind("<Double-1>", lambda e: self.open_file())
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.show_path())

        bottom = ttk.Frame(root, padding=8)
        bottom.pack(fill="x")
        self.path = tk.Label(bottom, text="", fg="#0b62d6", cursor="hand2", anchor="w", font=("Segoe UI", 10, "underline"))
        self.path.pack(side="left", fill="x", expand=True)
        self.path.bind("<Button-1>", lambda e: self.open_folder())
        ttk.Button(bottom, text="Open", command=self.open_file).pack(side="left")
        ttk.Button(bottom, text="Show in folder", command=self.open_folder).pack(side="left", padx=4)
        ttk.Button(bottom, text="Copy path", command=self.copy_path).pack(side="left")
        self.status = ttk.Label(root, text=f"Drives: {', '.join(drives())}", padding=(8, 0, 8, 6))
        self.status.pack(fill="x")
        self.entry.focus()

    def start(self):
        term = self.entry.get().strip()
        if not term:
            return
        self.stop.set()
        self.stop, self.q, self.count = threading.Event(), queue.Queue(), 0
        self.tree.delete(*self.tree.get_children())
        self.status.config(text=f"Searching for “{term}”…")
        t = threading.Thread(target=search, args=(term, drives(), self.stop, self.q), daemon=True)
        t.start()
        self.poll(t, self.q)

    def poll(self, t, q):
        if q is not self.q:  # a newer search replaced this one
            return
        while not q.empty():
            p, is_dir = q.get()
            try:
                st = os.stat(p)
                fmt = lambda ts: datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
                kind = "Folder" if is_dir else (os.path.splitext(p)[1].lstrip(".").upper() or "File")
                self.tree.insert("", "end", iid=p, values=(os.path.basename(p), kind, "" if is_dir else human(st.st_size), fmt(st.st_mtime), fmt(st.st_ctime), p))
                self.count += 1
            except OSError:
                pass
        done = not t.is_alive() and q.empty()
        self.status.config(text=f"{self.count} found" + ("" if done else " — still searching…"))
        if not done:
            self.root.after(150, self.poll, t, q)

    def selected(self):
        sel = self.tree.selection()
        return sel[0] if sel else None

    def show_path(self):
        self.path.config(text=self.selected() or "")

    def open_file(self):
        if p := self.selected():
            os.startfile(p)

    def open_folder(self):
        if p := self.selected():
            subprocess.Popen(["explorer", "/select,", p])

    def copy_path(self):
        if p := self.selected():
            self.root.clipboard_clear()
            self.root.clipboard_append(p)
            self.status.config(text=f"Copied: {p}")


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
