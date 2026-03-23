# OneNote to Markdown Converter - Development Plan

## Project Overview
- **Name**: OneNote2MD
- **Type**: Desktop app with CLI capability
- **Language**: Python
- **Core Functionality**: Export local .one files to Markdown
- **Features**: Local .one file parsing, configurable source path, batch export

---

## Architecture

```
onenote-to-markdown/
├── onenote2md/           # Main package
│   ├── __init__.py
│   ├── cli.py             # CLI entry point
│   ├── gui.py             # Desktop UI
│   ├── local_parser.py    # .one file parser
│   ├── converter.py      # Markdown converter
│   ├── config.py          # Config management
│   └── batch_export.py   # Batch export
├── tests/                 # Unit tests
├── docs/                  # Documentation
├── requirements.txt       # Python dependencies
├── setup.py              # Package setup
├── pyproject.toml        # Build config
├── build.sh             # Build script
├── config.json          # User config (created on first run)
└── README.md
```

---

## Tech Stack
- **CLI**: Click
- **Desktop UI**: CustomTkinter (modern look)
- **Local Files**: Custom parser
- **Markdown**: markdown-it-py

---

## Milestones (All Complete! ✅)

### Phase 1: Core & Config ✅ DONE
- [x] Project setup with Click CLI
- [x] Config file management (path, settings)
- [x] Config command: set/show source folder
- [x] List .one files from source folder

### Phase 2: File Parsing ✅ DONE
- [x] .one file parser
- [x] Extract text content
- [x] Extract page structure
- [x] Handle sections/notebooks

### Phase 3: Content Conversion ✅ DONE
- [x] Text formatting (bold, italic, lists, headings)
- [x] Image extraction
- [x] Table conversion
- [x] Link handling

### Phase 4: Structure & Batch ✅ DONE
- [x] Preserve folder hierarchy
- [x] Batch export all files
- [x] Progress indicators
- [x] Error handling

### Phase 5: Desktop UI ✅ DONE
- [x] Desktop app with CustomTkinter
- [x] Folder browser for source
- [x] Export progress UI
- [x] Settings persistence

### Phase 6: Polish ✅ DONE
- [x] Config customization options
- [x] Image handling options
- [x] Documentation
- [x] Build and release config

---

## Config File (~/.onenote2md/config.json)
```json
{
  "source_folder": "/path/to/your/onenote/backup/folder",
  "output_dir": "./output",
  "image_folder": "images",
  "embed_images": true,
  "include_metadata": false
}
```

---

## CLI Commands
```bash
# Configuration
onenote2md config set-source /path/to/onenote/backup
onenote2md config set-output /path/to/output
onenote2md config show

# List files
onenote2md list

# Export
onenote2md export --file "Page Name"
onenote2md export --all

# GUI
onenote2md gui
```

---

## Next Steps
1. Install dependencies: `pip install -r requirements.txt`
2. Set source folder: `onenote2md config set-source "/path/to/backup"`
3. Run: `onenote2md gui` or `onenote2md export --all`
