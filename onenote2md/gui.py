"""
OneNote to Markdown Converter - All-in-one GUI
"""

import customtkinter as ctk
import os
import sys
import subprocess
import threading
from pathlib import Path
from tkinter import filedialog

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
        self.cfg = self.load_config()
        
        # State
        self.selected_files = []
        self.is_exporting = False
        
        # Setup UI
        self.setup_ui()
        
    def load_config(self):
        """Load config from file."""
        import json
        config_dir = Path.home() / ".onenote2md"
        config_file = config_dir / "config.json"
        
        default_config = {
            "source_folder": "",
            "output_dir": "./output",
            "image_folder": "images",
            "embed_images": True
        }
        
        if config_file.exists():
            with open(config_file, 'r') as f:
                config = json.load(f)
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
        else:
            config_dir.mkdir(parents=True, exist_ok=True)
            with open(config_file, 'w') as f:
                json.dump(default_config, f, indent=2)
            return default_config
            
    def save_config(self, config):
        """Save config to file."""
        import json
        config_dir = Path.home() / ".onenote2md"
        config_dir.mkdir(parents=True, exist_ok=True)
        config_file = config_dir / "config.json"
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=2)
        
    def setup_ui(self):
        """Create the UI components."""
        
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title = ctk.CTkLabel(main_frame, text="📝 OneNote to Markdown Converter", font=ctk.CTkFont(size=28, weight="bold"))
        title.pack(pady=(0, 20))
        
        # Info
        info = ctk.CTkLabel(main_frame, text="💡 Select .one files → Converts to PDF → Converts to Markdown", font=ctk.CTkFont(size=12))
        info.pack(pady=(0, 10))
        
        # Settings
        settings_frame = ctk.CTkFrame(main_frame)
        settings_frame.pack(fill="x", pady=(0, 15))
        
        # Source
        ctk.CTkLabel(settings_frame, text="📁 Select Files", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        source_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        source_row.pack(fill="x", padx=15, pady=5)
        
        self.source_entry = ctk.CTkEntry(source_row, width=400, placeholder_text="Select files to convert...")
        self.source_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(source_row, text="Select Files", width=120, command=self.browse_source_files).pack(side="left", padx=5)
        
        # Output
        ctk.CTkLabel(settings_frame, text="📂 Output Folder", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        output_row = ctk.CTkFrame(settings_frame, fg_color="transparent")
        output_row.pack(fill="x", padx=15, pady=5)
        
        self.output_entry = ctk.CTkEntry(output_row, width=500, placeholder_text="Output directory...")
        self.output_entry.insert(0, self.cfg.get("output_dir", "./output"))
        self.output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        ctk.CTkButton(output_row, text="Browse", width=80, command=self.browse_output).pack(side="right")
        
        # Files list
        files_frame = ctk.CTkFrame(main_frame)
        files_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        ctk.CTkLabel(files_frame, text="📄 Selected Files", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=15, pady=(10, 5))
        
        list_frame = ctk.CTkFrame(files_frame, fg_color="transparent")
        list_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        self.files_listbox = ctk.CTkTextbox(list_frame, font=ctk.CTkFont(size=12))
        self.files_listbox.pack(side="left", fill="both", expand=True)
        
        scrollbar = ctk.CTkScrollbar(list_frame, command=self.files_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Export
        export_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        export_frame.pack(fill="x")
        
        self.progress = ctk.CTkProgressBar(export_frame)
        self.progress.pack(fill="x", pady=(0, 10))
        self.progress.set(0)
        
        button_row = ctk.CTkFrame(export_frame, fg_color="transparent")
        button_row.pack(fill="x")
        
        self.status_label = ctk.CTkLabel(button_row, text="Ready - Select files to convert", font=ctk.CTkFont(size=12))
        self.status_label.pack(side="left")
        
        self.export_btn = ctk.CTkButton(button_row, text="📥 Convert to Markdown", font=ctk.CTkFont(size=14, weight="bold"), height=40, command=self.start_export)
        self.export_btn.pack(side="right")
        
    def browse_source_files(self):
        """Browse for files."""
        files = filedialog.askopenfilenames(title="Select Files", filetypes=[("OneNote files", "*.one"), ("PDF files", "*.pdf"), ("All files", "*.*")])
        if files:
            self.selected_files = list(files)
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, f"{len(files)} file(s) selected")
            self.refresh_files()
            
    def browse_output(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)
            self.cfg["output_dir"] = folder
            self.save_config(self.cfg)
            
    def refresh_files(self):
        """Show selected files."""
        self.files_listbox.delete("1.0", "end")
        if not self.selected_files:
            self.files_listbox.insert("1.0", "⚠️ No files selected")
            return
        for f in self.selected_files:
            self.files_listbox.insert("end", f"📄 {Path(f).name}\n")
        self.files_listbox.insert("end", f"\n📊 Total: {len(self.selected_files)} file(s)")
        
    def start_export(self):
        """Start export."""
        if self.is_exporting:
            return
            
        output = self.output_entry.get()
        
        if not output:
            self.set_status("❌ Please select output folder", "red")
            return
            
        if not self.selected_files:
            self.set_status("⚠️ No files selected", "orange")
            return
            
        self.is_exporting = True
        self.export_btn.configure(state="disabled", text="⏳ Converting...")
        self.progress.set(0)
        
        self.export_thread = threading.Thread(target=self.run_export, args=(output,))
        self.export_thread.start()
        
    def run_export(self, output):
        """Run export."""
        try:
            os.makedirs(output, exist_ok=True)
            total = len(self.selected_files)
            
            for idx, f in enumerate(self.selected_files, 1):
                progress = idx / total
                self.after(0, lambda p=progress: self.progress.set(p))
                
                ext = Path(f).suffix.lower()
                
                if ext == '.pdf':
                    self.after(0, lambda i=idx, t=total, n=Path(f).name: self.set_status(f"📄 Converting PDF: {n} ({i}/{t})", "blue"))
                    self.convert_pdf_to_md(f, output)
                else:
                    # .one file - convert via PDF
                    self.after(0, lambda i=idx, t=total, n=Path(f).name: self.set_status(f"🔄 Converting .one → PDF: {n} ({i}/{t})", "orange"))
                    pdf_path = self.convert_one_to_pdf(f, output)
                    
                    if pdf_path and os.path.exists(pdf_path):
                        self.after(0, lambda p=pdf_path: self.set_status(f"📄 Converting PDF → Markdown: {Path(p).name}", "blue"))
                        self.convert_pdf_to_md(pdf_path, output)
                    else:
                        # Fallback: direct conversion
                        self.after(0, lambda n=Path(f).name: self.set_status(f"⚠️ PDF failed, trying direct: {n}", "orange"))
                        self.convert_one_direct(f, output)
                    
            self.after(0, lambda: self.set_status(f"✅ Exported {total} file(s)!", "green"))
                
        except Exception as e:
            self.after(0, lambda err=str(e): self.set_status(f"❌ Error: {err}", "red"))
        finally:
            self.is_exporting = False
            self.after(0, self.reset_export)
            
    def convert_one_to_pdf(self, one_file, output_dir) -> str:
        """Convert .one file to PDF using OneNote."""
        
        pdf_path = os.path.join(output_dir, Path(one_file).stem + ".pdf")
        
        try:
            # Method 1: Try OneNote COM via PowerShell
            ps_script = f'''
$oneNote = New-Object -ComObject OneNote.Application
$oneNote.OpenHierarchy("{one_file.replace("\\", "\\\\")}", $false)
$oneNote.Publish("{one_file.replace("\\", "\\\\")}", "{pdf_path.replace("\\", "\\\\")}", 2)
'''
            
            result = subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                text=True,
                timeout=60
            )
            
            if result.returncode == 0 and os.path.exists(pdf_path):
                return pdf_path
                
        except Exception as e:
            print(f"COM method failed: {e}")
            
        return None
        
    def convert_pdf_to_md(self, pdf_path, output_dir):
        """Convert PDF to Markdown."""
        try:
            text = ""
            
            # Try pdfminer
            try:
                from pdfminer.high_level import extract_text
                text = extract_text(pdf_path)
            except:
                pass
                
            # Fallback to PyPDF2
            if not text or len(text) < 50:
                try:
                    import PyPDF2
                    with open(pdf_path, 'rb') as pf:
                        reader = PyPDF2.PdfReader(pf)
                        pages_text = []
                        for p in reader.pages:
                            txt = p.extract_text()
                            if txt:
                                pages_text.append(txt)
                        text = '\n\n'.join(pages_text)
                except:
                    pass
                    
            if not text or len(text) < 50:
                text = "[Could not extract text from PDF]"
                
            # Convert to Markdown with better formatting
            md = self.text_to_markdown(text)
            
            # Save
            out_path = Path(output_dir) / f"{Path(pdf_path).stem}.md"
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(md)
                
        except Exception as e:
            print(f"PDF to MD error: {e}")
            # Save error message
            out_path = Path(output_dir) / f"{Path(pdf_path).stem}.md"
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(f"# {Path(pdf_path).stem}\n\n[Conversion error: {str(e)}]")
                
    def text_to_markdown(self, text: str) -> str:
        """Convert text to better formatted Markdown."""
        lines = text.split('\n')
        md_lines = []
        
        # Track context
        in_table = False
        table_lines = []
        
        for line in lines:
            line = line.strip()
            
            if not line:
                if in_table:
                    # End table
                    md_lines.extend(table_lines)
                    md_lines.append('')
                    table_lines = []
                    in_table = False
                md_lines.append('')
                continue
                
            # Detect tables (lines with multiple spaces or |)
            if '|' in line and line.count('|') >= 2:
                if not in_table:
                    in_table = True
                    table_lines = []
                    
                # Format table row
                cells = [c.strip() for c in line.split('|')]
                cells = [c for c in cells if c]  # Remove empty
                
                if all(c.replace('-', '').replace(':', '') == '' for c in cells):
                    # Skip separator line
                    continue
                    
                table_lines.append(line)
                
            elif in_table:
                # End table before processing this line
                md_lines.extend(table_lines)
                md_lines.append('')
                table_lines = []
                in_table = False
                
                # Process as normal line
                md_lines.append(self.format_line(line))
            else:
                # Normal line - apply formatting
                md_lines.append(self.format_line(line))
                
        # Close any open table
        if in_table and table_lines:
            md_lines.extend(table_lines)
            
        # Build markdown
        md = f"# {Path(pdf_path).stem if 'pdf_path' in locals() else 'Document'}\n\n"
        md += '\n'.join(md_lines)
        
        return md
        
    def format_line(self, line: str) -> str:
        """Format a line with basic markdown."""
        
        # Detect headings (short lines, all caps, or numbered sections)
        if len(line) < 60 and line.isupper():
            return f"## {line}"
        elif len(line) < 50 and not line.endswith(('.', ',', ':', ';', ')')):
            return f"### {line}"
            
        # Bold text **text**
        import re
        line = re.sub(r'\*\*(.+?)\*\*', r'**\1**', line)
        
        # Italic *text*
        line = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'*\1*', line)
        
        # Links [text](url)
        line = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'[\1](\2)', line)
        
        # Clean up extra spaces
        line = re.sub(r'\s+', ' ', line)
        
        return line
        
    def convert_one_direct(self, one_file, output_dir):
        """Fallback: Direct conversion of .one file."""
        try:
            with open(one_file, 'rb') as f:
                data = f.read()
                
            # Try to extract text
            text = ""
            for enc in ['utf-16-le', 'utf-8', 'latin-1']:
                try:
                    text = data.decode(enc, errors='ignore')
                    if len(text) > 100:
                        break
                except:
                    continue
                    
            if not text:
                text = "[Could not extract content from .one file]"
                
            # Clean
            text = text.replace('\x00', '')
            lines = []
            for line in text.split('\n'):
                line = line.strip()
                if line and len(line) > 2:
                    lines.append(line)
                    
            content = '\n\n'.join(lines[:500])
            
            md = f"# {Path(one_file).stem}\n\n{content}\n"
            
            out_path = Path(output_dir) / f"{Path(one_file).stem}.md"
            with open(out_path, 'w', encoding='utf-8') as f:
                f.write(md)
                
        except Exception as e:
            print(f"Direct conversion error: {e}")
            
    def set_status(self, message, color):
        """Set status message."""
        colors = {"gray": "#808080", "red": "#FF5555", "orange": "#FFA500", "green": "#50FA7B", "blue": "#8BE9FD"}
        self.status_label.configure(text=message, text_color=colors.get(color, "#808080"))
        
    def reset_export(self):
        """Reset export button."""
        self.export_btn.configure(state="normal", text="📥 Convert to Markdown")
        self.progress.set(1)
        

def main():
    app = OneNote2MDApp()
    app.mainloop()


if __name__ == "__main__":
    main()
