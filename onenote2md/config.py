import json
import os
from pathlib import Path

CONFIG_DIR = Path.home() / ".onenote2md"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_CONFIG = {
    "source_folder": "",
    "output_dir": "./output",
    "image_folder": "images",
    "embed_images": True,
    "include_metadata": False
}

def ensure_config_dir():
    """Create config directory if it doesn't exist."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

def load_config() -> dict:
    """Load config from file, return defaults if not exists."""
    ensure_config_dir()
    if not CONFIG_FILE.exists():
        save_config(DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()
    
    with open(CONFIG_FILE, 'r') as f:
        config = json.load(f)
        # Merge with defaults for any missing keys
        for key, value in DEFAULT_CONFIG.items():
            if key not in config:
                config[key] = value
        return config

def save_config(config: dict):
    """Save config to file."""
    ensure_config_dir()
    with open(CONFIG_FILE, 'w') as f:
        json.dump(config, f, indent=2)

def get_source_folder() -> str:
    """Get the configured source folder."""
    config = load_config()
    return config.get("source_folder", "")

def set_source_folder(path: str):
    """Set the source folder path."""
    config = load_config()
    config["source_folder"] = path
    save_config(config)

def get_output_dir() -> str:
    """Get the output directory."""
    config = load_config()
    return config.get("output_dir", "./output")