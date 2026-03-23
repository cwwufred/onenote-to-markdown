"""
Desktop GUI for OneNote to Markdown converter.
Uses CustomTkinter for modern, cross-platform UI.
"""

import customtkinter as ctk
import os
import threading
from pathlib import Path
from tkinter import filedialog
from onenote2md import config
from onenote2md.local_parser import list_one_files
from onenote2md.batch_export import batch_export

# Configure CustomTkinter appearance
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class OneNote2MDApp(ctk.CTk):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        
        self.title("📝 OneNote to Markdown Converter")
        self.geometry("900x700")
        self.minsize(700, 500)
        
        # Load config
        self.cfg = config.load_config()
        
        # State
        self.one_files = []
        self.export_thread = None
        self.is_exporting = False
        
        # Setup UI
        self.setup_ui()
        
    def setup_ui(self):
        """Create the UI components."""
        
        # Main container
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # ===== Title =====
        title = ctk.CTkLabel(
            main_frame,
            text="📝 OneNote to Markdown Converter",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title.pack(pady=(0, 20))
        
        # ===== Settings Section =====
        settings_frame = ctk.CTkFrame(main_frame)
        settings_frame.pack(fill="x", pady=(0, 15))
        
        # Source folder
        source_label = ctk.CTkLabel(settings_frame, text="📁 Source Folder", font=ctk.CTkFont(size=14, weight="bold"))
        source_label.pack(anchor="w", padx=15, pady=(10, 5))
        
        source_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        source_row.pack(fill="x", padx=15, pady=5)
        
        self.source_entry = ctk.CTkEntry(source_row, width=500, placeholder_text="Select folder containing .one files...")
        self.source_entry.insert(0, self.cfg.get("source_folder", ""))
        self.source_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(source_row, text="Browse", width=100, command=self.browse_source).pack(side="right")
        
        # Output folder
        output_label = ctk.CTkLabel(settings_frame, text="📂 Output Folder", font=ctk.CTkFont(size=14, weight="bold"))
        output_label.pack(anchor="w", padx=15, pady=(10, 5))
        
        output_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        output_row.pack(fill="x", padx=15, pady=5)
        
        self.output_entry = ctk.CTkEntry(output_row, width=500, placeholder_text="Output directory for Markdown files...")
        self.output_entry.insert(0, self.cfg.get("output_dir", "./output"))
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(output_row, text="Browse", width=100, command=self.browse_output).pack(side="right")
        
        # Options
        options_frame = ctk.CTkFrame(main_frame)
        options_frame.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(options_frame, text="⚙️ Options", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        options_row = ctk.CTkFrame(options_frame, fg_color="transparent")
        options_row.pack(fill="x", padx=15, pady=5)
        
        self.structure_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(options_row, text="Preserve folder structure", variable=self.structure_var).pack(side="left", padx=10)
        
        self.embed_images_var = ctk.BooleanVar(value=self.cfg.get("embed_images", True))
        ctk.CTkCheckBox(options_row, text="Embed images", variable=self.embed_images_var).pack(side="left", padx=10)
        
        # ===== Files Section =====
        files_frame = ctk.CTkFrame(main_frame)
        files_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        files_header = ctk.CTkFrame(files_frame, fg_color="transparent")
        files_header.pack(fill="x", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(files_header, text="📄 OneNote Files", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        
        ctk.CTkButton(
            files_header, 
            text="🔄 Refresh", 
            width=80,
            command=self.refresh_files
        ).pack(side="right")
        
        # File list with scrollbar
        list_frame = ctk.CTkFrame(files_frame, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        self.files_listbox = ctk.CTkTextbox(list_frame, font=ctk.CTkFont(size=12))
        self.files_listbox.pack(side="left", fill="both", expand=True)
        
        # Scrollbar
        scrollbar = ctk.CTkScrollbar(list_frame, command=self.files_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # ===== Export Section =====
        export_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        export_frame.pack(fill="x")
        
        # Progress bar
        self.progress = ctk.CTkProgressBar(export_frame)
        self.progress.pack(fill="x", pady=(0, 10))
        self.progress.set(0)
        
        # Status and buttons
        button_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        button_row.pack(fill="x")
        
        self.status_label = ctk.CTkLabel(button_row, text="Ready", font=ctk.CTkFont(size=12))
        self.status_label.pack(side="left")
        
        self.export_btn = ctk.CTkButton(
            button_row,
            text="📥 Export All to Markdown",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self.start_export
        )
        self.export_btn.pack(side="right")
        
        # Load files on startup
        self.after(100, self.refresh_files)
        
    def browse_source(self):
        """Browse for source folder."""
        folder = filedialog.askdirectory(title="Select OneNote Backup Folder")
        if folder:
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, folder)
            self.cfg["source_folder"] = folder
            config.save_config(self.cfg)
            self.refresh_files()
            
    def browse_output(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)
            self.cfg["output_dir"] = folder
            config.save_config(self.cfg)
            
    def refresh_files(self):
        """Refresh the file list."""
        source = self.source_entry.get()
        
        self.files_listbox.delete("1.0", "end")
        
        if not source:
            self.files_listbox.insert("1.0", "⚠️ Please configure source folder above")
            return
            
        if not Path(source).exists():
            self.files_listbox.insert("1.0", f"⚠️ Folder not found: {source}")
            return
            
        self.one_files = list_one_files(source)
        
        if not self.one_files:
            self.files_listbox.insert("1.0", "⚠️ No .one files found in this folder")
            return
            
        # Group by folder
        folders = {}
        for f in sorted(self.one_files):
            parent = f.parent.name
            if parent not in folders:
                folders[parent] = []
            folders[parent].append(f.name)
            
        for folder, files in folders.items():
            self.files_listbox.insert("end", f"📁 {folder}/\n")
            for f in files:
                self.files_listbox.insert("end", f"   📄 {f}\n")
                
        self.files_listbox.insert("end", f"\n📊 Total: {len(self.one_files)} file(s)")
        
    def start_export(self):
        """Start export in background thread."""
        if self.is_exporting:
            return
            
        source = self.source_entry.get()
        output = self.output_entry.get()
        
        if not source:
            self.set_status("❌ Please configure source folder", "red")
            return
            
        if not self.one_files:
            self.set_status("⚠️ No files to export", "orange")
            return
            
        # Update UI
        self.is_exporting = True
        self.export_btn.configure(state="disabled", text="⏳ Exporting...")
        self.progress.set(0)
        
        # Start export in background
        self.export_thread = threading.Thread(
            target=self.run_export,
            args=(source, output)
        )
        self.export_thread.start()
        
    def run_export(self, source: str, output: str):
        """Run export (called from background thread)."""
        try:
            # Configure batch exporter with progress callback
            results = batch_export(
                source,
                output,
                preserve_structure=self.structure_var.get(),
                progress_callback=self.update_progress
            )
            
            # Update UI with results
            successful = sum(1 for r in results if r.success)
            failed = len(results) - successful
            
            if failed > 0:
                self.set_status(f"⚠️ Exported {successful} file(s), {failed} failed", "orange")
            else:
                self.set_status(f"✅ Exported {successful} file(s) successfully!", "green")
                
        except Exception as e:
            self.set_status(f"❌ Export failed: {e}", "red")
            
        finally:
            self.is_exporting = False
            self.after(0, self.reset_export_button)
            
    def update_progress(self, current: int, total: int, filename: str):
        """Update progress bar (called from background thread)."""
        progress = current / total if total > 0 else 0
        self.after(0, lambda: self.progress.set(progress))
        self.after(0, lambda: self.set_status(f"📥 Exporting: {filename}", "blue"))
        
    def set_status(self, message: str, color: str = "gray"):
        """Set status message."""
        color_map = {
            "gray": "#808080",
            "red": "#FF5555",
            "orange": "#FFA500",
            "green": "#50FA7B",
            "blue": "#8BE9FD"
        }
        self.status_label.configure(text=message, text_color=color_map.get(color, "#808080"))
        
    def reset_export_button(self):
        """Reset export button after completion."""
        self.export_btn.configure(state="normal", text="📥 Export All to Markdown")
        self.progress.set(1)
        

def main():
    """Launch the GUI."""
    app = OneNote2MDApp()
    app.mainloop()


if __name__ == "__main__":
    main()