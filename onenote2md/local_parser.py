"""
OneNote .one file parser.

Uses ZIP/XML extraction to get all content from .one files.
"""

from pathlib import Path
from typing import List, Optional
from dataclasses import dataclass
import zipfile
import re
import xml.etree.ElementTree as ET


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
    """Parser for .one files using ZIP/XML extraction."""
    
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
            return self._parse_via_zip(file_path)
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return self._create_error_notebook(file_path, str(e))
    
    def _parse_via_zip(self, file_path: str) -> OneNoteNotebook:
        """
        Parse .one file by treating it as a ZIP archive.
        .one files are Compound File Binary Format (CFBF), but many have embedded ZIP structure.
        """
        sections = []
        
        try:
            # Try to open as ZIP first
            with zipfile.ZipFile(file_path, 'r') as z:
                names = z.namelist()
                
                # Find all XML files
                xml_files = [n for n in names if n.endswith('.xml')]
                
                # Try to find section and page structure
                section_pages = {}  # section_name -> [pages]
                
                for xml_file in xml_files:
                    try:
                        content = z.read(xml_file).decode('utf-8', errors='ignore')
                        
                        # Parse the XML
                        root = ET.fromstring(content)
                        
                        # Get element type
                        tag = root.tag.lower()
                        
                        # Check for sections
                        if 'section' in tag or 'sectiongroup' in tag:
                            section_name = root.get('name', root.get('nickname', xml_file))
                            section_id = root.get('id', xml_file)
                            
                            # Get all pages in this section
                            pages = self._extract_pages_from_element(root)
                            
                            if pages:
                                if section_name not in section_pages:
                                    section_pages[section_name] = []
                                section_pages[section_name].extend(pages)
                                
                        # Also check for direct pages
                        if 'page' in tag:
                            page = self._parse_page_element(root, xml_file)
                            if page:
                                section_name = "Pages"
                                if section_name not in section_pages:
                                    section_pages[section_name] = []
                                section_pages[section_name].append(page)
                                
                    except ET.ParseError:
                        continue
                    except Exception:
                        continue
                
                # Convert to sections
                for sec_name, pages in section_pages.items():
                    if pages:
                        sections.append(OneNoteSection(
                            name=sec_name,
                            pages=pages,
                            id=f"section-{len(sections)}"
                        ))
                
        except zipfile.BadZipFile:
            # Not a ZIP file, try CFBF parsing
            return self._parse_cfbf(file_path)
        except Exception as e:
            print(f"ZIP parsing error: {e}")
            
        # If no sections found, try to extract any content
        if not sections:
            sections = self._extract_any_content(file_path)
            
        if not sections:
            sections.append(OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(
                    title="Page 1",
                    content="[Could not parse content]",
                    id="page-1"
                )],
                id="section-1"
            ))
            
        return OneNoteNotebook(
            name=Path(file_path).stem,
            sections=sections,
            id="notebook-1"
        )
    
    def _extract_pages_from_element(self, element) -> List[OneNotePage]:
        """Extract all pages from an XML element."""
        pages = []
        
        # Find all page elements
        for page_elem in element.iter():
            tag = page_elem.tag.lower()
            if 'page' in tag and page_elem != element:
                page = self._parse_page_element_from_xml(page_elem)
                if page:
                    pages.append(page)
                    
        return pages
    
    def _parse_page_element_from_xml(self, page_elem) -> Optional[OneNotePage]:
        """Parse a page element from XML."""
        try:
            # Get page info
            page_id = page_elem.get('id', page_elem.get('objectID', ''))
            page_name = page_elem.get('name', page_elem.get('title', ''))
            
            if not page_name:
                # Try to get from title element
                for child in page_elem:
                    if 'title' in child.tag.lower():
                        page_name = child.text or ''
                        break
                        
            if not page_name:
                page_name = f"Page {page_id[:8]}" if page_id else "Untitled"
                
            # Extract content
            content = self._extract_text_from_element(page_elem)
            
            return OneNotePage(
                title=page_name[:100],
                content=content[:50000],  # Limit content size
                id=page_id[:50] if page_id else f"page-{len(pages)}"
            )
        except Exception as e:
            return None
    
    def _extract_text_from_element(self, element) -> str:
        """Extract all text content from an XML element."""
        texts = []
        
        for child in element.iter():
            # Get text content
            if child.text and child.text.strip():
                texts.append(child.text.strip())
            if child.tail and child.tail.strip():
                texts.append(child.tail.strip())
                
        return '\n'.join(texts)
    
    def _parse_page_element(self, root, xml_file: str) -> Optional[OneNotePage]:
        """Parse a page element."""
        try:
            page_id = root.get('id', root.get('objectID', xml_file))
            page_name = root.get('name', root.get('title', ''))
            
            if not page_name:
                # Try to find title in children
                for child in root:
                    if 'title' in child.tag.lower():
                        page_name = child.text or ''
                        break
                        
            if not page_name:
                page_name = "Untitled Page"
                
            content = self._extract_text_from_element(root)
            
            return OneNotePage(
                title=page_name[:100],
                content=content[:50000],
                id=page_id[:50] if page_id else xml_file
            )
        except Exception:
            return None
    
    def _extract_any_content(self, file_path: str) -> List[OneNoteSection]:
        """Extract any content from the file as last resort."""
        sections = []
        
        try:
            # Try reading as binary and finding any text
            with open(file_path, 'rb') as f:
                data = f.read()
                
            # Try UTF-16 decode
            try:
                text = data.decode('utf-16', errors='ignore')
            except:
                text = data.decode('utf-8', errors='ignore')
                
            # Find XML-like content
            xml_matches = re.findall(r'<[^>]+>', text)
            
            if xml_matches:
                # Create one big page with all content
                all_text = ' '.join(xml_matches[:1000])  # Limit
                sections.append(OneNoteSection(
                    name="All Content",
                    pages=[OneNotePage(
                        title="Extracted Content",
                        content=all_text[:50000],
                        id="page-1"
                    )],
                    id="section-1"
                ))
                
        except Exception as e:
            print(f"Any content extraction error: {e}")
            
        return sections
    
    def _parse_cfbf(self, file_path: str) -> OneNoteNotebook:
        """Try to parse as Compound File Binary Format."""
        # For now, return error notebook
        return self._create_error_notebook(file_path, "CFBF parsing not implemented")
    
    def _create_error_notebook(self, file_path: str, error: str) -> OneNoteNotebook:
        """Create a notebook with error information."""
        return OneNoteNotebook(
            name=Path(file_path).stem,
            sections=[OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(
                    title="Page 1",
                    content=f"[Could not parse file: {error}]",
                    id="page-1"
                )],
                id="section-1"
            )],
            id="notebook-1"
        )


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