# 📝 OneNote to Markdown Converter

Convert OneNote `.one` files to Markdown format with full content support.

## Features

- 📁 **Local File Support** - Parse `.one` files from backup folder
- 🔄 **Batch Export** - Export all files at once
- 📂 **Folder Structure** - Preserve notebook/section hierarchy
- 🖼️ **Image Support** - Extract embedded images
- 📊 **Table Support** - Convert tables to Markdown
- ✨ **Full Formatting** - Bold, italic, lists, links, checkboxes
- 🖥️ **Desktop GUI** - Easy-to-use interface
- ⚙️ **CLI** - Command-line for power users

## Installation

```bash
# Clone or download the project
cd onenote-to-markdown

# Install dependencies
pip install -r requirements.txt

# Install the package (optional)
pip install -e .
```

## Quick Start

### Using GUI

```bash
onenote2md gui
```

1. Browse to select your OneNote backup folder
2. Choose output folder
3. Click "Export All to Markdown"

### Using CLI

```bash
# Configure source folder (do once)
onenote2md config set-source "/path/to/onenote/backup"

# Configure output folder (optional)
onenote2md config set-output "./output"

# List available files
onenote2md list

# Export all files
onenote2md export --all

# Export specific file
onenote2md export --file "My Notebook"
```

## Configuration

Config file location: `~/.onenote2md/config.json`

```json
{
  "source_folder": "/path/to/backup",
  "output_dir": "./output",
  "image_folder": "images",
  "embed_images": true,
  "include_metadata": false
}
```

## CLI Commands

| Command | Description |
|---------|-------------|
| `onenote2md config set-source <path>` | Set source folder |
| `onenote2md config set-output <path>` | Set output folder |
| `onenote2md config show` | Show current config |
| `onenote2md list` | List .one files |
| `onenote2md export --all` | Export all files |
| `onenote2md export --file <name>` | Export specific file |
| `onenote2md gui` | Launch desktop GUI |

## Supported Content

- ✅ Headings (H1-H6)
- ✅ Bold and italic text
- ✅ Bullet and numbered lists
- ✅ Checkboxes (☐, ☑)
- ✅ Tables
- ✅ Links
- ✅ Inline code
- ✅ Images (embedded)
- ✅ Folder structure

## Requirements

- Python 3.8+
- click
- customtkinter
- markdown-it-py
- beautifulsoup4
- lxml
- python-dotenv

## License

MIT License

## Author

OneNote2MD