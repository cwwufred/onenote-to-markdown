"""
OneNote .one file parser.

Uses olefile to properly parse CFBF (Compound File Binary Format) .one files.
"""

from pathlib import Path
from typing import List, Optional
from dataclasses import dataclass
import re


@dataclass
class OneNotePage:
    """Represents a OneNote page."""
    title: str
    content: str
    id: str
    last_modified: str = ""


@dataclass
class OneNoteSection:
    """Represents a OneNote section (like a tab)."""
    name: str
    pages: List[OneNotePage]
    id: str


@dataclass
class OneNoteNotebook:
    """Represents a OneNote notebook."""
    name: str
    sections: List[OneNoteSection]
    id: str


class OneNoteParser:
    """Parser for .one files using olefile."""
    
    def __init__(self):
        self.files = {}
        
    def parse_file(self, file_path: str) -> Optional[OneNoteNotebook]:
        """Parse a .one file and return notebook structure."""
        try:
            import olefile
            return self._parse_with_olefile(file_path)
        except ImportError:
            return self._parse_fallback(file_path)
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return self._create_error_notebook(file_path, str(e))
    
    def _parse_with_olefile(self, file_path: str) -> OneNoteNotebook:
        """Parse using olefile."""
        
        try:
            ole = olefile.OleFileIO(file_path)
            
            # List all streams in the file
            streams = ole.listdir()
            
            sections = []
            all_text = []
            
            # Read each stream
            for stream in streams:
                try:
                    stream_path = '/'.join(stream)
                    
                    # Skip system streams
                    if stream[0].startswith('\x01'):
                        continue
                        
                    # Read stream data
                    data = ole.openstream(stream).read()
                    
                    # Try to extract text from the stream
                    text = self._extract_text_from_data(data)
                    if text and len(text) > 10:
                        all_text.append(f"=== {stream_path} ===\n{text}")
                        
                except Exception as e:
                    continue
                    
            ole.close()
            
            if all_text:
                # Create sections from extracted text
                content = '\n\n'.join(all_text[:20])  # Limit to first 20 streams
                
                sections = [OneNoteSection(
                    name="All Content",
                    pages=[OneNotePage(
                        title="Notebook Content",
                        content=content[:100000],
                        id="page-1"
                    )],
                    id="section-1"
                )]
                
                return OneNoteNotebook(
                    name=Path(file_path).stem,
                    sections=sections,
                    id="notebook-1"
                )
                
        except Exception as e:
            print(f"OLE parsing error: {e}")
            
        return self._parse_fallback(file_path)
    
    def _extract_text_from_data(self, data: bytes) -> str:
        """Extract readable text from binary data."""
        
        # Try UTF-16 LE (common in OneNote)
        try:
            text = data.decode('utf-16-le', errors='ignore')
            text = self._clean_text(text)
            if len(text) > 50:
                return text
        except:
            pass
            
        # Try UTF-8
        try:
            text = data.decode('utf-8', errors='ignore')
            text = self._clean_text(text)
            if len(text) > 50:
                return text
        except:
            pass
            
        # Try Latin-1
        try:
            text = data.decode('latin-1', errors='ignore')
            text = self._clean_text(text)
            if len(text) > 50:
                return text
        except:
            pass
            
        return ""
    
    def _clean_text(self, text: str) -> str:
        """Clean up extracted text."""
        
        # Remove excessive whitespace
        lines = []
        prev_empty = False
        
        for line in text.split('\n'):
            line = line.strip()
            if line:
                lines.append(line)
                prev_empty = False
            elif not prev_empty:
                lines.append('')
                prev_empty = True
                
        return '\n'.join(lines)
    
    def _parse_fallback(self, file_path: str) -> OneNoteNotebook:
        """Fallback parsing using basic binary reading."""
        
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
                
            # Try to find any readable text
            text = self._extract_text_from_data(data)
            
            if text and len(text) > 50:
                # Split into pages
                pages = self._split_into_pages(text)
                
                return OneNoteNotebook(
                    name=Path(file_path).stem,
                    sections=[OneNoteSection(
                        name="Notebook",
                        pages=pages,
                        id="section-1"
                    )],
                    id="notebook-1"
                )
                
        except Exception as e:
            print(f"Fallback parsing error: {e}")
            
        return self._create_error_notebook(file_path, "Could not parse file content")
    
    def _split_into_pages(self, text: str) -> List[OneNotePage]:
        """Split text into pages."""
        
        # Try to find section/page markers
        # Look for common delimiters
        
        # Split by double newlines into paragraphs
        paragraphs = re.split(r'\n\s*\n', text)
        
        pages = []
        current_content = []
        current_title = "Page 1"
        page_num = 1
        
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue
                
            # Check if this looks like a title (short, no ending punctuation)
            if len(para) < 80 and not para.endswith('.') and not para.endswith(','):
                if current_content:
                    # Save previous page
                    pages.append(OneNotePage(
                        title=current_title,
                        content='\n\n'.join(current_content)[:5000],
                        id=f"page-{page_num}"
                    ))
                    page_num += 1
                    current_content = []
                    
                current_title = para[:50]
            else:
                current_content.append(para)
                
        # Add last page
        if current_content:
            pages.append(OneNotePage(
                title=current_title,
                content='\n\n'.join(current_content)[:5000],
                id=f"page-{page_num}"
            ))
            
        if not pages:
            pages = [OneNotePage(
                title="Page 1",
                content=text[:10000],
                id="page-1"
            )]
            
        return pages
    
    def _create_error_notebook(self, file_path: str, error: str) -> OneNoteNotebook:
        """Create error notebook."""
        return OneNoteNotebook(
            name=Path(file_path).stem,
            sections=[OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(
                    title="Page 1",
                    content=f"[Note: .one file parsing is limited. For complete export, please use OneNote's built-in Export to PDF or HTML, then convert to Markdown.]",
                    id="page-1"
                )],
                id="section-1"
            )],
            id="notebook-1"
        )


def parse_one_file(file_path: str) -> Optional[OneNoteNotebook]:
    """Parse a .one file."""
    parser = OneNoteParser()
    return parser.parse_file(file_path)


def list_one_files(folder: str) -> List[Path]:
    """List all .one files in a folder."""
    return list(Path(folder).rglob("*.one"))