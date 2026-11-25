# FileFinder

A small search box for Windows. Type part of a file or folder name and it searches every drive, every folder, every file.

- Results show name, type, size, modified and created dates, and the full path.
- Double-click a result to open it; click the path to show it in Explorer; **Copy path** copies it.
- Enter searches, Esc closes. Skips `Windows`, `$Recycle.Bin`, `node_modules` and `.git`.

## Run

Needs Python 3.8+ (tkinter is included).

```
pythonw filefinder.pyw
```

Tip: make a desktop shortcut to `pythonw.exe "path\to\filefinder.pyw"` and give it a shortcut key (e.g. Ctrl+Alt+F).
