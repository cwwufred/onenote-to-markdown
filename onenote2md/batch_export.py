"""
Batch export with folder structure preservation and progress tracking.
"""

import os
import sys
from pathlib import Path
from typing import List, Dict, Optional
from dataclasses import dataclass
from datetime import datetime


@dataclass
class ExportResult:
    """Result of an export operation."""
    file: Path
    success: bool
    output_path: Optional[Path] = None
    error: Optional[str] = None
    pages: int = 0


class BatchExporter:
    """Handle batch export with folder structure and progress."""
    
    def __init__(self, output_dir: str = "./output", preserve_structure: bool = True):
        self.output_dir = Path(output_dir)
        self.preserve_structure = preserve_structure
        self.results: List[ExportResult] = []
        
    def export_folder(self, source_folder: str, 
                     progress_callback=None) -> List[ExportResult]:
        """
        Export all .one files from source folder.
        
        Args:
            source_folder: Path to folder containing .one files
            progress_callback: Optional callback function for progress updates
            
        Returns:
            List of ExportResult objects
        """
        source = Path(source_folder)
        
        if not source.exists():
            raise FileNotFoundError(f"Source folder not found: {source_folder}")
            
        # Find all .one files
        one_files = list(source.rglob("*.one"))
        
        if not one_files:
            print("⚠️ No .one files found")
            return []
            
        self.results = []
        total = len(one_files)
        
        print(f"📥 Starting batch export: {total} file(s)\n")
        
        for idx, one_file in enumerate(sorted(one_files), 1):
            # Update progress
            if progress_callback:
                progress_callback(idx, total, one_file.name)
            else:
                self._default_progress(idx, total, one_file.name)
                
            # Determine output path
            if self.preserve_structure:
                # Calculate relative path from source
                rel_path = one_file.relative_to(source)
                output_path = self.output_dir / rel_path.with_suffix('.md')
                
                # Create parent directories
                output_path.parent.mkdir(parents=True, exist_ok=True)
            else:
                output_path = self.output_dir / f"{one_file.stem}.md"
                
            # Parse and convert
            result = self._export_file(one_file, output_path)
            self.results.append(result)
            
            # Show result
            if result.success:
                print(f"  ✅ [{idx}/{total}] {one_file.name} → {result.output_path.name}")
            else:
                print(f"  ❌ [{idx}/{total}] {one_file.name} - {result.error}")
                
        # Summary
        successful = sum(1 for r in self.results if r.success)
        failed = total - successful
        
        print(f"\n📊 Export complete:")
        print(f"   ✅ Success: {successful}")
        print(f"   ❌ Failed: {failed}")
        print(f"   📂 Output: {self.output_dir}")
        
        return self.results
        
    def _export_file(self, one_file: Path, output_path: Path) -> ExportResult:
        """Export a single .one file."""
        try:
            # Import here to avoid circular imports
            from onenote2md.local_parser import parse_one_file
            from onenote2md.converter import convert_to_markdown
            
            # Parse file
            notebook = parse_one_file(str(one_file))
            
            if not notebook:
                return ExportResult(
                    file=one_file,
                    success=False,
                    error="Failed to parse file"
                )
                
            # Count pages
            page_count = sum(len(s.pages) for s in notebook.sections)
            
            # Convert and save
            result_path = convert_to_markdown(notebook, str(output_path.parent))
            
            # Rename if needed
            if result_path != output_path:
                result_path.rename(output_path)
                
            return ExportResult(
                file=one_file,
                success=True,
                output_path=output_path,
                pages=page_count
            )
            
        except Exception as e:
            return ExportResult(
                file=one_file,
                success=False,
                error=str(e)
            )
            
    def _default_progress(self, current: int, total: int, filename: str):
        """Default progress display."""
        percent = int((current / total) * 100)
        bar_len = 30
        filled = int(bar_len * current / total)
        bar = '█' * filled + '░' * (bar_len - filled)
        
        sys.stdout.write(f"\r  Progress: [{bar}] {percent}% ({current}/{total})")
        sys.stdout.flush()
        
    def get_summary(self) -> Dict:
        """Get export summary."""
        return {
            'total': len(self.results),
            'successful': sum(1 for r in self.results if r.success),
            'failed': sum(1 for r in self.results if not r.success),
            'total_pages': sum(r.pages for r in self.results if r.success),
            'output_dir': str(self.output_dir)
        }


def batch_export(source_folder: str, output_dir: str = "./output",
                preserve_structure: bool = True,
                progress_callback=None) -> List[ExportResult]:
    """
    Convenience function for batch export.
    
    Args:
        source_folder: Path to folder with .one files
        output_dir: Output directory
        preserve_structure: Preserve folder hierarchy
        progress_callback: Optional progress callback
        
    Returns:
        List of ExportResult objects
    """
    exporter = BatchExporter(output_dir, preserve_structure)
    return exporter.export_folder(source_folder, progress_callback)