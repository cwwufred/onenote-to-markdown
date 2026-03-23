"""
OneNote .one file parser.

Extracts content from .one files using binary/text analysis.
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
    """Parser for .one files."""
    
    def __init__(self):
        self.files = {}
        
    def parse_file(self, file_path: str) -> Optional[OneNoteNotebook]:
        """Parse a .one file and return notebook structure."""
        try:
            return self._parse_binary(file_path)
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return self._create_error_notebook(file_path, str(e))
    
    def _parse_binary(self, file_path: str) -> OneNoteNotebook:
        """Parse .one file by extracting text content."""
        
        with open(file_path, 'rb') as f:
            data = f.read()
        
        # Try to find any text content
        # Extract printable text sequences
        text_content = self._extract_text_from_binary(data)
        
        if not text_content:
            return self._create_error_notebook(file_path, "No readable content found")
        
        # Try to find page/section structure
        sections = self._find_structure(text_content)
        
        if not sections:
            # Create one big section with all content
            sections = [OneNoteSection(
                name="All Pages",
                pages=[OneNotePage(
                    title="Page 1",
                    content=text_content[:50000],
                    id="page-1"
                )],
                id="section-1"
            )]
            
        return OneNoteNotebook(
            name=Path(file_path).stem,
            sections=sections,
            id="notebook-1"
        )
    
    def _extract_text_from_binary(self, data: bytes) -> str:
        """Extract readable text from binary data."""
        
        # Try different encodings
        for encoding in ['utf-16-le', 'utf-16', 'utf-8', 'latin-1']:
            try:
                text = data.decode(encoding, errors='ignore')
                # Check if it has meaningful content
                if len(text) > 100:
                    # Clean up the text
                    return self._clean_text(text)
            except:
                continue
                
        return ""
    
    def _clean_text(self, text: str) -> str:
        """Clean up extracted text."""
        
        # Remove null bytes and excessive whitespace
        text = text.replace('\x00', '')
        
        # Find sequences of readable characters
        # Look for text between common XML/HTML tags
        lines = []
        current_line = []
        
        for char in text:
            if char.isprintable() or char in '\n\r\t':
                current_line.append(char)
            else:
                if len(''.join(current_line).strip()) > 2:
                    lines.append(''.join(current_line))
                current_line = []
                
        if current_line:
            lines.append(''.join(current_line))
            
        # Join meaningful lines
        cleaned = []
        for line in lines:
            line = line.strip()
            if line and len(line) > 2:
                cleaned.append(line)
                
        return '\n'.join(cleaned[:500])  # Limit lines
    
    def _find_structure(self, text: str) -> List[OneNoteSection]:
        """Try to find page/section structure in text."""
        
        sections = []
        
        # Look for common OneNote page markers
        # These patterns are common in OneNote files
        
        # Try to split by common delimiters
        # Look for titles or section headers
        
        # Split by double newlines to find paragraphs
        paragraphs = re.split(r'\n\s*\n', text)
        
        pages = []
        current_page_content = []
        page_id = 1
        
        for para in paragraphs:
            para = para.strip()
            if not para:
                continue
                
            # Check if this looks like a title (short, no punctuation)
            if len(para) < 100 and not para.endswith('.') and not para.endswith(','):
                # Save previous page if exists
                if current_page_content:
                    pages.append(OneNotePage(
                        title=f"Page {page_id}",
                        content='\n'.join(current_page_content)[:10000],
                        id=f"page-{page_id}"
                    ))
                    page_id += 1
                    current_page_content = []
                    
            current_page_content.append(para)
            
        # Add last page
        if current_page_content:
            pages.append(OneNotePage(
                title=f"Page {page_id}",
                content='\n'.join(current_page_content)[:10000],
                id=f"page-{page_id}"
            ))
            
        if pages:
            sections.append(OneNoteSection(
                name="Notebook",
                pages=pages,
                id="section-1"
            ))
            
        return sections
    
    def _create_error_notebook(self, file_path: str, error: str) -> OneNoteNotebook:
        """Create a notebook with error information."""
        return OneNoteNotebook(
            name=Path(file_path).stem,
            sections=[OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(
                    title="Page 1",
                    content=f"[Could not parse file properly. The .one file may be in a newer format or encrypted.]",
                    id="page-1"
                )],
                id="section-1"
            )],
            id="notebook-1"
        )


def parse_one_file(file_path: str) -> Optional[OneNoteNotebook]:
    """Convenience function to parse a .one file."""
    parser = OneNoteParser()
    return parser.parse_file(file_path)


def list_one_files(folder: str) -> List[Path]:
    """List all .one files in a folder."""
    return list(Path(folder).rglob("*.one"))