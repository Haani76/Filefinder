"""FileFinder: type part of a file name, search every drive. Double-click a result to open it."""
import os, re, string, subprocess, threading, queue, datetime, zipfile, tkinter as tk
from tkinter import ttk

try:  # optional: image and PDF previews
    from PIL import Image, ImageTk
except ImportError:
    Image = None
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

SKIP = {"$recycle.bin", "system volume information", "windows", "node_modules", ".git", "__pycache__"}
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tif", ".tiff"}
MAX_TEXT = 20000  # characters shown in the preview


def preview(path):
    """Return ("image", PIL.Image) or ("text", str) describing the file at path."""
    if os.path.isdir(path):
        try:
            items = sorted(os.listdir(path))
        except OSError as e:
            return "text", f"Can't open folder: {e}"
        return "text", f"Folder with {len(items)} items:\n\n" + "\n".join(items[:500])
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in IMAGES and Image:
            img = Image.open(path)
            img.thumbnail((520, 520))
            return "image", img
        if ext == ".pdf" and PdfReader:
            pages = PdfReader(path).pages
            return "text", f"PDF, {len(pages)} pages\n\n" + "\n".join((p.extract_text() or "") for p in pages[:5])[:MAX_TEXT]
        if ext in (".docx", ".pptx", ".xlsx"):
            with zipfile.ZipFile(path) as z:
                xml = " ".join(z.read(n).decode("utf8", "ignore") for n in z.namelist()
                               if re.match(r"(word/document|ppt/slides/slide\d+|xl/sharedStrings)\.xml", n))
            return "text", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", xml)).strip()[:MAX_TEXT]
        with open(path, "rb") as f:
            raw = f.read(MAX_TEXT)
        if b"\0" in raw:
            return "text", "No preview for this file type."
        return "text", raw.decode("utf8", "replace")
    except Exception as e:  # corrupt/locked files shouldn't crash the app
        return "text", f"Can't preview: {e}"


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
        root.geometry("1400x640")
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
        panes = ttk.PanedWindow(root, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8)
        self.tree = ttk.Treeview(panes, columns=cols, show="headings")
        for c, w in zip(cols, (200, 60, 70, 120, 120, 300)):
            self.tree.heading(c, text=c.title())
            self.tree.column(c, width=w, anchor="w")
        side = ttk.Frame(panes)
        panes.add(self.tree, weight=3)
        panes.add(side, weight=2)
        self.image = ttk.Label(side, anchor="center")
        self.text = tk.Text(side, wrap="word", font=("Consolas", 10), state="disabled")
        self.text.pack(fill="both", expand=True)
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
        p = self.selected()
        self.path.config(text=p or "")
        if not p:
            return
        kind, data = preview(p)
        if kind == "image":
            self.photo = ImageTk.PhotoImage(data)  # keep a reference or Tk drops it
            self.image.config(image=self.photo)
            self.text.pack_forget()
            self.image.pack(fill="both", expand=True)
        else:
            self.image.pack_forget()
            self.text.pack(fill="both", expand=True)
            self.text.config(state="normal")
            self.text.delete("1.0", "end")
            self.text.insert("1.0", data)
            self.text.config(state="disabled")

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
