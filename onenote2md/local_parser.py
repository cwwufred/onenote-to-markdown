"""
OneNote .one file parser.

The .one file format is based on Compound File Binary Format (CFBF).
This parser extracts content from OneNote files.
"""

import struct
import xml.etree.ElementTree as ET
from pathlib import Path
from dataclasses import dataclass
from typing import List, Optional
import io


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
    
    # OneNote file signature and structure constants
    ONE_SIGNATURE = b'\xE4\x52\x5C\x7B\x8C\xD5\xD8\x11'
    
    def __init__(self):
        self.files = {}
        
    def parse_file(self, file_path: str) -> Optional[OneNoteNotebook]:
        """
        Parse a .one file and return notebook structure.
        
        Args:
            file_path: Path to .one file
            
        Returns:
            OneNoteNotebook object or None if parsing fails
        """
        try:
            with open(file_path, 'rb') as f:
                data = f.read()
                
            # Try to extract XML content from the file
            # .one files store content in XML format within the CFBF structure
            return self._parse_one_content(data, Path(file_path).stem)
            
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return None
            
    def _parse_one_content(self, data: bytes, file_name: str) -> OneNoteNotebook:
        """
        Extract content from .one file data.
        
        The .one file contains XML data that can be parsed.
        """
        # Try to find and parse XML content in the file
        # OneNote files contain compressed/encrypted XML data
        
        # For now, we'll create a basic structure with extracted text
        # Real implementation would use onenote-parser library
        
        # Try to find XML-like content
        try:
            # Look for XML content in the file
            xml_content = self._extract_xml(data)
            if xml_content:
                return self._parse_xml_to_notebook(xml_content, file_name)
        except:
            pass
            
        # Fallback: create basic notebook with raw data info
        return OneNoteNotebook(
            name=file_name,
            sections=[
                OneNoteSection(
                    name="Quick Notes",
                    pages=[
                        OneNotePage(
                            title="Page 1",
                            content=f"[Content extraction requires additional parsing - file: {file_name}]",
                            id="page-1"
                        )
                    ],
                    id="section-1"
                )
            ],
            id="notebook-1"
        )
        
    def _extract_xml(self, data: bytes) -> Optional[str]:
        """Try to extract XML content from file data."""
        # Look for XML-like patterns in the data
        # This is a simplified approach
        
        try:
            # Try UTF-16 decode
            text = data.decode('utf-16', errors='ignore')
            if '<?xml' in text or '<OneNote' in text:
                return text
        except:
            pass
            
        try:
            # Try UTF-8
            text = data.decode('utf-8', errors='ignore')
            if '<?xml' in text or '<OneNote' in text:
                return text
        except:
            pass
            
        return None
        
    def _parse_xml_to_notebook(self, xml_content: str, file_name: str) -> OneNoteNotebook:
        """Parse XML content to notebook structure."""
        sections = []
        
        try:
            root = ET.fromstring(xml_content)
            
            # Extract notebook name
            nb_name = root.get('name', file_name)
            
            # Find sections (simplified)
            for section in root.findall('.//Section') or root.findall('.//s'):
                sec_name = section.get('name', section.get('n', 'Untitled Section'))
                pages = []
                
                for page in section.findall('.//Page') or section.findall('.//p'):
                    page_title = page.get('name', page.get('n', 'Untitled Page'))
                    page_content = self._extract_page_content(page)
                    
                    pages.append(OneNotePage(
                        title=page_title,
                        content=page_content,
                        id=page.get('id', f"page-{len(pages)}")
                    ))
                    
                if pages:
                    sections.append(OneNoteSection(
                        name=sec_name,
                        pages=pages,
                        id=section.get('id', f"section-{len(sections)}")
                    ))
                    
        except ET.ParseError:
            # XML parsing failed, create basic structure
            pass
            
        if not sections:
            sections.append(OneNoteSection(
                name="Quick Notes",
                pages=[
                    OneNotePage(
                        title="Page 1",
                        content="[No content parsed - file may be encrypted or use unsupported format]",
                        id="page-1"
                    )
                ],
                id="section-1"
            ))
            
        return OneNoteNotebook(
            name=file_name,
            sections=sections,
            id="notebook-1"
        )
        
    def _extract_page_content(self, page_elem) -> str:
        """Extract text content from a page element."""
        texts = []
        
        # Get text from various elements
        for text in page_elem.findall('.//Text') or page_elem.findall('.//t'):
            if text.text:
                texts.append(text.text)
                
        return '\n'.join(texts) if texts else "[No text content]"


def parse_one_file(file_path: str) -> Optional[OneNoteNotebook]:
    """
    Convenience function to parse a .one file.
    
    Args:
        file_path: Path to .one file
        
    Returns:
        OneNoteNotebook object or None
    """
    parser = OneNoteParser()
    return parser.parse_file(file_path)


def list_one_files(folder: str) -> List[Path]:
    """
    List all .one files in a folder.
    
    Args:
        folder: Path to folder
        
    Returns:
        List of Path objects for .one files
    """
    return list(Path(folder).rglob("*.one"))