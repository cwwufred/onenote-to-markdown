"""OneNote to Markdown Converter - GUI"""

import customtkinter as ctk
import os
import subprocess
import threading
from pathlib import Path
from tkinter import filedialog

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class OneNote2MDApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("OneNote to Markdown")
        self.geometry("900x700")
        self.cfg = self.load_config()
        self.selected_files = []
        self.is_exporting = False
        self.setup_ui()

    def load_config(self):
        import json
        config_dir = Path.home() / ".onenote2md"
        config_file = config_dir / "config.json"
        default = {"source_folder": "", "output_dir": "./output"}
        if config_file.exists():
            with open(config_file, 'r') as f:
                cfg = json.load(f)
                return {**default, **cfg}
        return default

    def save_config(self, cfg):
        import json
        config_dir = Path.home() / ".onenote2md"
        config_dir.mkdir(parents=True, exist_ok=True)
        with open(config_dir / "config.json", 'w') as f:
            json.dump(cfg, f)

    def setup_ui(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(main, text="OneNote to Markdown", font=("Arial", 24, "bold")).pack(pady=(0, 10))
        ctk.CTkLabel(main, text="Select .one files -> PDF -> Markdown", font=("Arial", 12)).pack(pady=(0, 10))

        sf = ctk.CTkFrame(main)
        sf.pack(fill="x", pady=10)
        ctk.CTkLabel(sf, text="Files:", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        self.source_entry = ctk.CTkEntry(sf, width=400)
        self.source_entry.pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkButton(sf, text="Browse", command=self.browse).pack(side="right", padx=10)

        of = ctk.CTkFrame(main)
        of.pack(fill="x", pady=10)
        ctk.CTkLabel(of, text="Output:", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        self.output_entry = ctk.CTkEntry(of, width=400)
        self.output_entry.insert(0, self.cfg.get("output_dir", "./output"))
        self.output_entry.pack(side="left", padx=10, fill="x", expand=True)
        ctk.CTkButton(of, text="Browse", command=self.browse_out).pack(side="right", padx=10)

        ff = ctk.CTkFrame(main)
        ff.pack(fill="both", expand=True, pady=10)
        ctk.CTkLabel(ff, text="Selected:", font=("Arial", 14, "bold")).pack(anchor="w", padx=10, pady=5)
        self.files_list = ctk.CTkTextbox(ff)
        self.files_list.pack(fill="both", expand=True, padx=10, pady=5)

        self.progress = ctk.CTkProgressBar(main)
        self.progress.pack(fill="x", pady=10)
        self.progress.set(0)

        self.status = ctk.CTkLabel(main, text="Ready", font=("Arial", 12))
        self.status.pack(pady=5)

        self.export_btn = ctk.CTkButton(main, text="Convert to Markdown", font=("Arial", 14, "bold"), height=40, command=self.start_export)
        self.export_btn.pack(pady=10)

    def browse(self):
        files = filedialog.askopenfilenames(title="Select Files", filetypes=[("OneNote", "*.one"), ("PDF", "*.pdf")])
        if files:
            self.selected_files = list(files)
            self.source_entry.delete(0, "end")
            self.source_entry.insert(0, f"{len(files)} files")
            self.refresh_list()

    def browse_out(self):
        folder = filedialog.askdirectory(title="Output Folder")
        if folder:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, folder)
            self.cfg["output_dir"] = folder
            self.save_config(self.cfg)

    def refresh_list(self):
        self.files_list.delete("1.0", "end")
        for f in self.selected_files:
            self.files_list.insert("end", f"{Path(f).name}\n")

    def start_export(self):
        if self.is_exporting or not self.selected_files:
            return
        self.is_exporting = True
        self.export_btn.configure(state="disabled", text="Converting...")
        threading.Thread(target=self.run_export).start()

    def run_export(self):
        output = self.output_entry.get()
        os.makedirs(output, exist_ok=True)
        total = len(self.selected_files)
        
        for i, f in enumerate(self.selected_files, 1):
            self.after(0, lambda p=i/total: self.progress.set(p))
            ext = Path(f).suffix.lower()
            
            if ext == ".pdf":
                self.after(0, lambda n=Path(f).name: self.status.configure(text=f"Converting PDF: {n}"))
                self.convert_pdf(f, output)
            else:
                self.after(0, lambda n=Path(f).name: self.status.configure(text=f"Converting .one to PDF: {n}"))
                pdf = self.convert_one_to_pdf(f, output)
                if pdf and os.path.exists(pdf):
                    self.after(0, lambda n=Path(pdf).name: self.status.configure(text=f"Converting PDF to MD: {n}"))
                    self.convert_pdf(pdf, output)
                else:
                    self.convert_one_direct(f, output)
            # Keep the intermediate PDF file
            if ext == ".one" and pdf and os.path.exists(pdf):
                pass  # PDF is already in output folder

        self.after(0, lambda: self.status.configure(text="Done!"))
        self.after(0, lambda: self.export_btn.configure(state="normal", text="Convert to Markdown"))
        self.after(0, lambda: self.progress.set(1))
        self.is_exporting = False

    def convert_one_to_pdf(self, one_file, output_dir):
        pdf_path = os.path.join(output_dir, Path(one_file).stem + ".pdf")
        try:
            one_escaped = one_file.replace("\\", "\\\\")
            pdf_escaped = pdf_path.replace("\\", "\\\\")
            ps = f'''
$one = New-Object -ComObject OneNote.Application
$one.OpenHierarchy("{one_escaped}", $false)
$one.Publish("{one_escaped}", "{pdf_escaped}", 2)
'''
            subprocess.run(["powershell", "-Command", ps], capture_output=True, timeout=60)
            return pdf_path if os.path.exists(pdf_path) else None
        except:
            return None

    def convert_pdf(self, pdf_path, output_dir):
        text = ""
        try:
            from pdfminer.high_level import extract_text
            text = extract_text(pdf_path)
        except:
            try:
                import PyPDF2
                with open(pdf_path, 'rb') as pf:
                    reader = PyPDF2.PdfReader(pf)
                    text = "\n\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
            except:
                pass
        if not text:
            text = "[Could not extract]"
        
        md = f"# {Path(pdf_path).stem}\n\n" + self.format_text(text)
        out = Path(output_dir) / f"{Path(pdf_path).stem}.md"
        with open(out, 'w', encoding='utf-8') as f:
            f.write(md)

    def format_text(self, text):
        import re
        lines = []
        for line in text.split('\n'):
            line = line.strip()
            if not line:
                lines.append("")
            elif len(line) < 50 and line.isupper():
                lines.append(f"## {line}")
            elif len(line) < 40 and not line.endswith(('.', ',')):
                lines.append(f"### {line}")
            else:
                lines.append(line)
        return '\n'.join(lines)

    def convert_one_direct(self, one_file, output_dir):
        try:
            with open(one_file, 'rb') as f:
                data = f.read()
            for enc in ['utf-16-le', 'utf-8', 'latin-1']:
                try:
                    text = data.decode(enc, errors='ignore')
                    if len(text) > 100:
                        break
                except:
                    continue
            text = '\n'.join([l.strip() for l in text.split('\n') if l.strip() and len(l.strip()) > 2])
            md = f"# {Path(one_file).stem}\n\n{text[:10000]}\n"
            out = Path(output_dir) / f"{Path(one_file).stem}.md"
            with open(out, 'w', encoding='utf-8') as f:
                f.write(md)
        except:
            pass


if __name__ == "__main__":
    app = OneNote2MDApp()
    app.mainloop()
