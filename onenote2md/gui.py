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
        self.selected_files = []
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
        source_label = ctk.CTkLabel(settings_frame, text="📁 Source Folder / Files", font=ctk.CTkFont(size=14, weight="bold"))
        source_label.pack(anchor="w", padx=15, pady=(10, 5))
        
        source_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        source_row.pack(fill="x", padx=15, pady=5)
        
        self.source_entry = ctk.CTkEntry(source_row, width=400, placeholder_text="Select folder or files...")
        self.source_entry.insert(0, self.cfg.get("source_folder", ""))
        self.source_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(source_row, text="Browse Folder", width=120, command=self.browse_source_folder).pack(side="left", padx=5)
        ctk.CTkButton(source_row, text="Select Files", width=120, command=self.browse_source_files).pack(side="left", padx=5)
        
        # Output folder
        output_label = ctk.CTkLabel(settings_frame, text="📂 Output Folder", font=ctk.CTkFont(size=14, weight="bold"))
        output_label.pack(anchor="w", padx=15, pady=(10, 5))
        
        output_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        output_row.pack(fill="x", padx=15, pady=5)
        
        self.output_entry = ctk.CTkEntry(output_row, width=500, placeholder_text="Output directory for Markdown files...")
        self.output_entry.insert(0, self.cfg.get("output_dir", "./output"))
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(output_row, text="Browse", width=80, command=self.browse_output).pack(side="right")
        
        # ===== Files Section =====
        files_frame = ctk.CTkFrame(main_frame)
        files_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        files_header = ctk.CTkFrame(files_frame, fg_color="transparent")
        files_header.pack(fill="x", padx=15, pady=(10, 5))
        
        ctk.CTkLabel(files_header, text="📄 Selected Files", font=ctk.CTkFont(size=14, weight="bold")).pack(side="left")
        
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
        
        self.status_label = ctk.CTkLabel(button_row, text="Ready - Select files or folder to start", font=ctk.CTkFont(size=12))
        self.status_label.pack(side="left")
        
        self.export_btn = ctk.CTkButton(
            button_row,
            text="📥 Export to Markdown",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            command=self.start_export
        )
        self.export_btn.pack(side="right")
        
        # Load files on startup
        self.after(100, self.refresh_files)
        
    def browse_source_folder(self):
        """Browse for source folder."""
        folder = filedialog.askdirectory(title="Select OneNote Backup Folder")
        if folder:
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, folder)
            self.cfg["source_folder"] = folder
            config.save_config(self.cfg)
            self.selected_files = []  # Clear selected files
            self.refresh_files()
            
    def browse_source_files(self):
        """Browse for specific .one or PDF files."""
        files = filedialog.askopenfilenames(
            title="Select OneNote/PDF Files",
            filetypes=[
                ("OneNote files", "*.one"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        if files:
            self.selected_files = list(files)
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, f"{len(files)} file(s) selected")
            self.refresh_selected_files()
            
    def browse_output(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)
            self.cfg["output_dir"] = folder
            config.save_config(self.cfg)
            
    def refresh_files(self):
        """Refresh the file list from folder."""
        source = self.source_entry.get()
        
        self.files_listbox.delete("1.0", "end")
        
        # Check if we have selected files
        if self.selected_files:
            self.refresh_selected_files()
            return
            
        if not source:
            self.files_listbox.insert("1.0", "⚠️ Please select a folder or files above")
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
        
    def refresh_selected_files(self):
        """Show selected files."""
        self.files_listbox.delete("1.0", "end")
        
        if not self.selected_files:
            self.files_listbox.insert("1.0", "⚠️ No files selected")
            return
            
        for f in self.selected_files:
            fname = Path(f).name
            self.files_listbox.insert("end", f"📄 {fname}\n")
            
        self.files_listbox.insert("end", f"\n📊 Total: {len(self.selected_files)} file(s) selected")
        
    def start_export(self):
        """Start export in background thread."""
        if self.is_exporting:
            return
            
        output = self.output_entry.get()
        
        if not output:
            self.set_status("❌ Please configure output folder", "red")
            return
            
        # Determine source files
        source = self.source_entry.get()
        
        if self.selected_files:
            # Use selected files
            files_to_export = self.selected_files
        elif source and Path(source).exists():
            # Use folder
            files_to_export = list_one_files(source)
        else:
            self.set_status("⚠️ No files to export", "orange")
            return
            
        if not files_to_export:
            self.set_status("⚠️ No files to export", "orange")
            return
            
        # Update UI
        self.is_exporting = True
        self.export_btn.configure(state="disabled", text="⏳ Exporting...")
        self.progress.set(0)
        
        # Start export in background
        self.export_thread = threading.Thread(
            target=self.run_export,
            args=(files_to_export, output)
        )
        self.export_thread.start()
        
    def run_export(self, files, output: str):
        """Run export (called from background thread)."""
        try:
            # Create output directory
            os.makedirs(output, exist_ok=True)
            
            total = len(files)
            
            for idx, f in enumerate(files, 1):
                progress = idx / total
                self.after(0, lambda p=progress: self.progress.set(p))
                self.after(0, lambda i=idx, t=total, n=Path(f).name: self.set_status(f"📥 Exporting: {n} ({i}/{t})", "blue"))
                
                # Check file type
                file_ext = Path(f).suffix.lower()
                
                if file_ext == '.one':
                    # Convert .one to PDF first
                    self.after(0, lambda n=Path(f).name: self.set_status(f"🔄 Converting to PDF: {n}", "orange"))
                    
                    try:
                        from onenote2md.one_to_pdf import convert_one_to_pdf
                        pdf_path = convert_one_to_pdf(f, output)
                        
                        if pdf_path and os.path.exists(pdf_path):
                            self.after(0, lambda p=pdf_path: self.set_status(f"📄 Converting PDF to Markdown: {Path(p).name}", "blue"))
                            from onenote2md.pdf_converter import pdf_to_markdown
                            pdf_to_markdown(pdf_path, output)
                        else:
                            # Fallback: try direct conversion
                            from onenote2md.local_parser import parse_one_file
                            from onenote2md.converter import convert_to_markdown
                            notebook = parse_one_file(f)
                            if notebook:
                                convert_to_markdown(notebook, output)
                    except Exception as e:
                        print(f"PDF conversion error: {e}")
                        # Fallback to direct conversion
                        try:
                            from onenote2md.local_parser import parse_one_file
                            from onenote2md.converter import convert_to_markdown
                            notebook = parse_one_file(f)
                            if notebook:
                                convert_to_markdown(notebook, output)
                        except Exception as e2:
                            print(f"Direct conversion also failed: {e2}")
                            
                elif file_ext == '.pdf':
                    # Convert PDF to Markdown
                    from onenote2md.pdf_converter import pdf_to_markdown
                    pdf_to_markdown(f, output)
                    
                else:
                    # Try direct conversion
                    from onenote2md.local_parser import parse_one_file
                    from onenote2md.converter import convert_to_markdown
                    notebook = parse_one_file(f)
                    if notebook:
                        convert_to_markdown(notebook, output)
                    
            self.after(0, lambda: self.set_status(f"✅ Exported {total} file(s) successfully!", "green"))
                
        except Exception as e:
            self.after(0, lambda err=str(e): self.set_status(f"❌ Export failed: {err}", "red"))
            
        finally:
            self.is_exporting = False
            self.after(0, self.reset_export_button)
            
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
        self.export_btn.configure(state="normal", text="📥 Export to Markdown")
        self.progress.set(1)
        

def main():
    """Launch the GUI."""
    app = OneNote2MDApp()
    app.mainloop()


if __name__ == "__main__":
    main()