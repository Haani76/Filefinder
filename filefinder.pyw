"""FileFinder: search every drive by file name (and optionally file contents). Double-click a result to open it."""
import os, re, string, subprocess, threading, queue, datetime, zipfile, hashlib, tempfile, tkinter as tk
from tkinter import ttk

try:  # optional: image and PDF previews
    from PIL import Image, ImageTk
except ImportError:
    Image = None
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None
try:  # optional: draws real PDF pages
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None
try:  # optional: lets installed Microsoft Office convert documents for preview
    import pythoncom, win32com.client
except ImportError:
    win32com = None

SKIP = {"$recycle.bin", "system volume information", "windows", "node_modules", ".git", "__pycache__"}
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tif", ".tiff"}
MAX_TEXT = 20000  # characters shown in the preview
OFFICE = {".docx", ".pptx", ".xlsx"}
TEXT = {".txt", ".md", ".csv", ".tsv", ".json", ".xml", ".html", ".htm", ".css", ".js", ".jsx", ".ts", ".tsx",
        ".py", ".pyw", ".java", ".c", ".h", ".cpp", ".cs", ".go", ".rs", ".php", ".rb", ".sql", ".sh", ".bat",
        ".ps1", ".ini", ".cfg", ".conf", ".yml", ".yaml", ".toml", ".log", ".tex", ".rtf", ".srt"}
CONTENT_MAX_BYTES = 20 * 1024 * 1024  # skip bigger files in content search; they make it crawl


def office_text(path):
    with zipfile.ZipFile(path) as z:
        xml = " ".join(z.read(n).decode("utf8", "ignore") for n in z.namelist()
                       if re.match(r"(word/document|ppt/slides/slide\d+|xl/sharedStrings)\.xml", n))
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", xml)).strip()


def file_text(path):
    """Full searchable text of a text/PDF/Office file, or None if the type isn't searchable. May raise."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf" and PdfReader:
        return "\n".join((p.extract_text() or "") for p in PdfReader(path).pages)
    if ext in OFFICE:
        return office_text(path)
    if ext in TEXT:
        with open(path, "rb") as f:
            return f.read().decode("utf8", "replace")
    return None


WORD, POWERPOINT, EXCEL = {".doc", ".docx", ".rtf", ".odt"}, {".ppt", ".pptx", ".odp"}, {".xls", ".xlsx", ".ods"}
MAX_PAGES = 10  # pages drawn in the preview
CACHE = os.path.join(tempfile.gettempdir(), "FileFinderPreview")
RENDER_LOCK = threading.Lock()  # one Office/pdfium job at a time; neither is safe to run in parallel


def office_to_pdf(path):
    """Have installed Office export the document to a cached PDF; returns its path. May raise."""
    st = os.stat(path)
    os.makedirs(CACHE, exist_ok=True)
    out = os.path.join(CACHE, hashlib.md5(f"{path}|{st.st_mtime}|{st.st_size}".encode()).hexdigest() + ".pdf")
    if os.path.exists(out):
        return out
    tmp = out[:-4] + ".tmp.pdf"
    ext = os.path.splitext(path)[1].lower()
    pythoncom.CoInitialize()
    app = None
    try:
        if ext in WORD:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible, app.DisplayAlerts = False, 0
            # positional: FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument
            # (a dummy password makes protected files fail instead of popping a hidden prompt)
            doc = app.Documents.Open(path, False, True, False, "x")
            doc.ExportAsFixedFormat(tmp, 17)
            doc.Close(False)
            docs = app.Documents
        elif ext in POWERPOINT:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            pres = app.Presentations.Open(path, True, False, False)  # ReadOnly, Untitled, WithWindow
            pres.SaveAs(tmp, 32)
            pres.Close()
            docs = app.Presentations
        else:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible, app.DisplayAlerts = False, False
            wb = app.Workbooks.Open(path, 0, True, 5, "x")  # UpdateLinks, ReadOnly, Format, Password
            wb.ExportAsFixedFormat(0, tmp)
            wb.Close(False)
            docs = app.Workbooks
        os.replace(tmp, out)
        return out
    finally:
        # PowerPoint is single-instance: never quit it if the user has their own files open
        if app is not None and docs_count(locals().get("docs")) == 0:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def docs_count(docs):
    try:
        return docs.Count
    except Exception:
        return 0


def render_pdf(path, width):
    """Draw the first MAX_PAGES pages at the given pixel width; returns (images, total_pages)."""
    pdf = pdfium.PdfDocument(path)
    try:
        imgs = []
        for i in range(min(len(pdf), MAX_PAGES)):
            page = pdf[i]
            imgs.append(page.render(scale=width / page.get_width()).to_pil())
        return imgs, len(pdf)
    finally:
        pdf.close()


def preview(path, width=520, still_wanted=lambda: True):
    """Return ("pages", ([PIL images], total)) for things shown as they look, or ("text", str)."""
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
            img.thumbnail((width, width * 3))
            return "pages", ([img], 1)
        if ext == ".pdf" and pdfium:
            with RENDER_LOCK:
                return "pages", render_pdf(path, width)
        if ext in WORD | POWERPOINT | EXCEL and win32com and pdfium:
            with RENDER_LOCK:
                if not still_wanted():  # user already clicked something else; skip the slow Office step
                    return "text", ""
                try:
                    return "pages", render_pdf(office_to_pdf(path), width)
                except Exception as e:
                    if ext not in OFFICE:
                        raise
                    return "text", f"(Office couldn't render this file: {e})\n\n" + office_text(path)[:MAX_TEXT]
        if ext == ".pdf" and PdfReader:
            pages = PdfReader(path).pages
            return "text", f"PDF, {len(pages)} pages\n\n" + "\n".join((p.extract_text() or "") for p in pages[:5])[:MAX_TEXT]
        if ext in OFFICE:
            return "text", office_text(path)[:MAX_TEXT]
        with open(path, "rb") as f:
            raw = f.read(MAX_TEXT)
        if b"\0" in raw:
            return "text", "No preview for this file type."
        return "text", raw.decode("utf8", "replace")
    except Exception as e:  # corrupt/locked files shouldn't crash the app
        return "text", f"Can't preview: {e}"


def windows_dark():
    """True when Windows apps are set to dark mode."""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


THEMES = {  # bg, panel (fields/lists), fg, muted, accent/link, selection
    True: dict(bg="#1e1f22", panel="#2b2d31", fg="#e6e6e6", muted="#9aa0a6", link="#6cb4ff", sel="#3d5a80"),
    False: dict(bg="#f3f3f3", panel="#ffffff", fg="#1b1b1b", muted="#5f6368", link="#0b62d6", sel="#cce4ff"),
}


def drives():
    return [f"{d}:\\" for d in string.ascii_uppercase if os.path.exists(f"{d}:\\")]


def search(term, roots, stop, out, contents=False):
    """Walk every root; put (path, is_dir, "name"|"content") for matches (case-insensitive)."""
    term = term.lower()
    for root in roots:
        for cur, dirs, files in os.walk(root, onerror=lambda e: None):
            if stop.is_set():
                return
            dirs[:] = [d for d in dirs if d.lower() not in SKIP]
            for name in dirs:
                if term in name.lower():
                    out.put((os.path.join(cur, name), True, "name"))
            for name in files:
                p = os.path.join(cur, name)
                if term in name.lower():
                    out.put((p, False, "name"))
                elif contents:
                    if stop.is_set():
                        return
                    try:
                        if os.path.getsize(p) <= CONTENT_MAX_BYTES:
                            t = file_text(p)
                            if t and term in t.lower():
                                out.put((p, False, "content"))
                    except Exception:  # locked/corrupt files: skip, keep searching
                        pass


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
        ttk.Button(top, text="Stop", command=lambda: self.stop.set()).pack(side="left", padx=4)
        self.contents = tk.BooleanVar()
        ttk.Checkbutton(top, text="Search inside files (slower)", variable=self.contents).pack(side="left", padx=4)
        self.dark = tk.BooleanVar(value=windows_dark())
        ttk.Checkbutton(top, text="Dark mode", variable=self.dark, command=self.apply_theme).pack(side="left", padx=4)

        cols = ("name", "match", "type", "size", "modified", "created", "path")
        panes = ttk.PanedWindow(root, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8)
        self.tree = ttk.Treeview(panes, columns=cols, show="headings")
        for c, w in zip(cols, (200, 65, 60, 70, 120, 120, 300)):
            self.tree.heading(c, text=c.title())
            self.tree.column(c, width=w, anchor="w")
        side = ttk.Frame(panes)
        panes.add(self.tree, weight=3)
        panes.add(side, weight=2)
        self.pages = ttk.Frame(side)
        self.canvas = tk.Canvas(self.pages, highlightthickness=0)
        bar = ttk.Scrollbar(self.pages, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(-e.delta // 120, "units"))
        self.photos, self.ptoken = [], 0
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
        self.apply_theme()
        self.entry.focus()

    def apply_theme(self):
        dark = self.dark.get()
        c = THEMES[dark]
        s = ttk.Style()
        s.theme_use("clam")  # the only built-in ttk theme that honours custom colours
        s.configure(".", background=c["bg"], foreground=c["fg"], fieldbackground=c["panel"],
                    bordercolor=c["panel"], lightcolor=c["bg"], darkcolor=c["bg"], troughcolor=c["bg"],
                    insertcolor=c["fg"], selectbackground=c["sel"], selectforeground=c["fg"])
        s.configure("TButton", background=c["panel"], padding=(10, 4))
        s.map("TButton", background=[("active", c["sel"])])
        s.map("TCheckbutton", background=[("active", c["bg"])], indicatorbackground=[("!selected", c["panel"])])
        s.configure("Treeview", background=c["panel"], fieldbackground=c["panel"], foreground=c["fg"], rowheight=24)
        s.map("Treeview", background=[("selected", c["sel"])], foreground=[("selected", c["fg"])])
        s.configure("Treeview.Heading", background=c["bg"], foreground=c["muted"])
        s.map("Treeview.Heading", background=[("active", c["panel"])])
        self.root.configure(bg=c["bg"])
        self.text.configure(bg=c["panel"], fg=c["fg"], insertbackground=c["fg"], relief="flat")
        self.path.configure(bg=c["bg"], fg=c["link"])
        self.canvas.configure(bg=c["bg"])
        self.theme = c
        try:  # dark title bar on Windows 10/11
            import ctypes
            self.root.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(ctypes.c_int(dark)), 4)
            self.root.withdraw(); self.root.deiconify()  # force the title bar to repaint
        except Exception:
            pass

    def start(self):
        term = self.entry.get().strip()
        if not term:
            return
        self.stop.set()
        self.stop, self.q, self.count = threading.Event(), queue.Queue(), 0
        self.tree.delete(*self.tree.get_children())
        self.status.config(text=f"Searching for “{term}”…")
        t = threading.Thread(target=search, args=(term, drives(), self.stop, self.q, self.contents.get()), daemon=True)
        t.start()
        self.poll(t, self.q)

    def poll(self, t, q):
        if q is not self.q:  # a newer search replaced this one
            return
        while not q.empty():
            p, is_dir, how = q.get()
            try:
                st = os.stat(p)
                fmt = lambda ts: datetime.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
                kind = "Folder" if is_dir else (os.path.splitext(p)[1].lstrip(".").upper() or "File")
                self.tree.insert("", "end", iid=p, values=(os.path.basename(p), how.title(), kind, "" if is_dir else human(st.st_size), fmt(st.st_mtime), fmt(st.st_ctime), p))
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
        self.ptoken += 1
        token, box = self.ptoken, {}
        width = max(self.canvas.winfo_width(), 400) - 24
        self.show_text("Loading preview…")
        # rendering (especially via Office) can take seconds, so do it off the UI thread
        threading.Thread(target=lambda: box.update(r=preview(p, width, lambda: token == self.ptoken)), daemon=True).start()
        self.wait_preview(token, box)

    def wait_preview(self, token, box):
        if token != self.ptoken:  # a newer click took over
            return
        if "r" not in box:
            self.root.after(100, self.wait_preview, token, box)
            return
        kind, data = box["r"]
        if kind == "text":
            self.show_text(data)
            return
        imgs, total = data
        self.text.pack_forget()
        self.pages.pack(fill="both", expand=True)
        self.canvas.delete("all")
        self.photos = [ImageTk.PhotoImage(i) for i in imgs]  # keep references or Tk drops them
        y = 8
        for ph in self.photos:
            self.canvas.create_image(8, y, image=ph, anchor="nw")
            y += ph.height() + 12
        if total > len(imgs):
            self.canvas.create_text(8, y, anchor="nw", fill=self.theme["muted"],
                                    text=f"Showing {len(imgs)} of {total} pages. Open the file to see the rest.")
            y += 30
        self.canvas.configure(scrollregion=(0, 0, 0, y))
        self.canvas.yview_moveto(0)

    def show_text(self, data):
        self.pages.pack_forget()
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
