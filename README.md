# FileFinder

A small search box for Windows. Type part of a file or folder name and it searches every drive, every folder, every file.

- Tick **Search inside files** to also match text inside files: text and code files, PDFs, and Word/PowerPoint/Excel (`.docx`, `.pptx`, `.xlsx`). Files over 20 MB are skipped. It's much slower than a name search; the **Match** column says whether a result matched by name or content.
- Results show name, type, size, modified and created dates, and the full path.
- Double-click a result to open it; click the path to show it in Explorer; **Copy path** copies it.
- Click a result to preview it on the right (drag the divider to resize):
  - Text and code files: the first 20,000 characters
  - Images (PNG, JPG, GIF, BMP, WebP, ICO, TIFF): the image
  - PDFs: the real pages, as they look (first 10 pages; scroll with the mouse wheel)
  - Word, PowerPoint, Excel (`.doc/.docx/.rtf/.odt`, `.ppt/.pptx/.odp`, `.xls/.xlsx/.ods`): the real pages, as they look. Your installed Microsoft Office converts them in the background, so the first preview of a file takes a few seconds; after that it's cached. Without Office, `.docx/.pptx/.xlsx` fall back to plain text.
  - Folders: the list of items inside
  - In page/image previews: **mouse wheel zooms** in and out around the cursor (20%–600%), **drag** to move around, use the horizontal/vertical scrollbars to seek, **double-click** to fit the width again.
- **Dark mode**: starts in your Windows app theme; toggle it with the checkbox.
- Enter searches, Esc closes. Skips `Windows`, `$Recycle.Bin`, `node_modules` and `.git`.

## Run

Needs Python 3.8+ (tkinter is included). For page previews and searching inside PDFs, also install:

```
pip install pillow pypdf pypdfium2 pywin32
```

Without them the app still works; previews fall back to plain text or "No preview".

```
pythonw filefinder.pyw
```

Tip: make a desktop shortcut to `pythonw.exe "path\to\filefinder.pyw"` and give it a shortcut key (e.g. Ctrl+Alt+F).
