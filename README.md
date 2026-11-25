# FileFinder

A small search box for Windows. Type part of a file or folder name and it searches every drive, every folder, every file.

- Results show name, type, size, modified and created dates, and the full path.
- Double-click a result to open it; click the path to show it in Explorer; **Copy path** copies it.
- Click a result to preview it on the right (drag the divider to resize):
  - Text and code files: the first 20,000 characters
  - Images (PNG, JPG, GIF, BMP, WebP, ICO, TIFF): a thumbnail
  - PDFs: page count and text of the first 5 pages
  - Word, PowerPoint, Excel (`.docx`, `.pptx`, `.xlsx`): plain text
  - Folders: the list of items inside
- Enter searches, Esc closes. Skips `Windows`, `$Recycle.Bin`, `node_modules` and `.git`.

## Run

Needs Python 3.8+ (tkinter is included). For image and PDF previews, also install:

```
pip install pillow pypdf
```

Without them the app still works; those files just show "No preview".

```
pythonw filefinder.pyw
```

Tip: make a desktop shortcut to `pythonw.exe "path\to\filefinder.pyw"` and give it a shortcut key (e.g. Ctrl+Alt+F).
