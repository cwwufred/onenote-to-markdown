"""
OneNote .one file parser.

Uses onenote-parser library to properly extract all content.
"""

from pathlib import Path
from typing import List, Optional
from dataclasses import dataclass


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
    """Parser for .one files using onenote-parser library."""
    
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
            # Try to use onenote-parser library
            from onenote_parser import OneNoteParser as ONParser
            
            oneparser = ONParser(filepath=file_path)
            oneparser.parse()
            
            # Convert to our structure
            return self._convert_from_onenote_parser(oneparser, Path(file_path).stem)
            
        except ImportError:
            # Fallback to basic parsing
            return self._parse_file_fallback(file_path)
        except Exception as e:
            print(f"Error parsing {file_path}: {e}")
            return self._parse_file_fallback(file_path)
    
    def _convert_from_onenote_parser(self, oneparser, file_name: str) -> OneNoteNotebook:
        """Convert onenote-parser result to our structure."""
        sections = []
        
        try:
            # Try to get sections from the parsed data
            if hasattr(oneparser, 'sections'):
                for section in oneparser.sections:
                    pages = []
                    
                    # Get section name
                    sec_name = getattr(section, 'name', 'Untitled Section')
                    sec_id = getattr(section, 'id', f"section-{len(sections)}")
                    
                    # Try to get pages
                    if hasattr(section, 'pages'):
                        for page in section.pages:
                            page_title = getattr(page, 'title', 'Untitled Page')
                            page_id = getattr(page, 'id', f"page-{len(pages)}")
                            page_content = getattr(page, 'content', '') or getattr(page, 'text', '')
                            
                            pages.append(OneNotePage(
                                title=page_title,
                                content=str(page_content) if page_content else '',
                                id=page_id
                            ))
                    
                    if pages:
                        sections.append(OneNoteSection(
                            name=sec_name,
                            pages=pages,
                            id=sec_id
                        ))
                        
        except Exception as e:
            print(f"Conversion error: {e}")
            
        # If no sections found, try alternative approach
        if not sections:
            sections = self._extract_all_pages(oneparser, file_name)
            
        if not sections:
            sections.append(OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(title="Page 1", content="[Content extraction failed]", id="page-1")],
                id="section-1"
            ))
            
        return OneNoteNotebook(
            name=file_name,
            sections=sections,
            id="notebook-1"
        )
    
    def _extract_all_pages(self, oneparser, file_name: str) -> List[OneNoteSection]:
        """Try to extract all pages recursively."""
        sections = []
        
        try:
            # Get all objects from the parser
            if hasattr(oneparser, 'objects'):
                page_count = 0
                current_pages = []
                
                for obj in oneparser.objects:
                    # Check if it's a page
                    if hasattr(obj, 'type'):
                        obj_type = str(obj.type).lower()
                        if 'page' in obj_type:
                            title = getattr(obj, 'title', f"Page {page_count + 1}")
                            content = getattr(obj, 'content', '') or getattr(obj, 'text', '')
                            
                            current_pages.append(OneNotePage(
                                title=str(title) if title else f"Page {page_count + 1}",
                                content=str(content) if content else '',
                                id=f"page-{page_count}"
                            ))
                            page_count += 1
                            
                if current_pages:
                    sections.append(OneNoteSection(
                        name="All Pages",
                        pages=current_pages,
                        id="section-1"
                    ))
                    
        except Exception as e:
            print(f"Recursive extraction error: {e}")
            
        return sections
        
    def _parse_file_fallback(self, file_path: str) -> Optional[OneNoteNotebook]:
        """
        Fallback parser using basic CFBF parsing.
        """
        try:
            from onenote_parser import parse
            
            # Try the parse function directly
            result = parse(file_path)
            
            if result:
                return self._convert_parsed_result(result, Path(file_path).stem)
                
        except Exception as e:
            print(f"Fallback parsing error: {e}")
            
        # Ultimate fallback - try to read raw XML
        return self._parse_raw_xml(file_path)
    
    def _convert_parsed_result(self, result, file_name: str) -> OneNoteNotebook:
        """Convert parsed result to our structure."""
        sections = []
        
        try:
            # Handle different result formats
            if isinstance(result, dict):
                # Try common keys
                for key in ['sections', 'SectionSet', 'sections_list']:
                    if key in result:
                        items = result[key]
                        for idx, item in enumerate(items):
                            pages = []
                            
                            if isinstance(item, dict):
                                sec_name = item.get('name', f"Section {idx + 1}")
                                sec_id = item.get('id', f"section-{idx}")
                                
                                # Get pages
                                page_items = item.get('pages', [])
                                for pidx, p in enumerate(page_items):
                                    if isinstance(p, dict):
                                        pages.append(OneNotePage(
                                            title=p.get('title', f"Page {pidx + 1}"),
                                            content=p.get('content', '') or p.get('text', ''),
                                            id=p.get('id', f"page-{pidx}")
                                        ))
                                        
                            if pages:
                                sections.append(OneNoteSection(
                                    name=sec_name,
                                    pages=pages,
                                    id=sec_id
                                ))
                                
        except Exception as e:
            print(f"Result conversion error: {e}")
            
        if not sections:
            sections.append(OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(title="Page 1", content="[No content parsed]", id="page-1")],
                id="section-1"
            ))
            
        return OneNoteNotebook(
            name=file_name,
            sections=sections,
            id="notebook-1"
        )
    
    def _parse_raw_xml(self, file_path: str) -> OneNoteNotebook:
        """Last resort - try to extract any XML content."""
        import zipfile
        
        try:
            # .one files can be opened as ZIP to get XML
            with zipfile.ZipFile(file_path, 'r') as z:
                # List all files in the archive
                names = z.namelist()
                
                pages = []
                page_id = 0
                
                # Look for any XML files with content
                for name in names:
                    if name.endswith('.xml'):
                        try:
                            content = z.read(name).decode('utf-8', errors='ignore')
                            
                            # Try to extract text content
                            import re
                            # Find text between tags
                            texts = re.findall(r'>([^<]+)<', content)
                            
                            if texts and len(texts) > 5:  # Has meaningful content
                                title = name.split('/')[-1].replace('.xml', '') or f"Page {page_id + 1}"
                                full_text = '\n'.join([t.strip() for t in texts if t.strip()])
                                
                                pages.append(OneNotePage(
                                    title=title[:50],  # Limit title length
                                    content=full_text[:5000],  # Limit content
                                    id=f"page-{page_id}"
                                ))
                                page_id += 1
                                
                        except:
                            continue
                            
                if pages:
                    return OneNoteNotebook(
                        name=Path(file_path).stem,
                        sections=[OneNoteSection(
                            name="All Content",
                            pages=pages,
                            id="section-1"
                        )],
                        id="notebook-1"
                    )
                    
        except Exception as e:
            print(f"XML extraction error: {e}")
            
        # Complete failure - return minimal structure
        return OneNoteNotebook(
            name=Path(file_path).stem,
            sections=[OneNoteSection(
                name="Quick Notes",
                pages=[OneNotePage(
                    title="Page 1",
                    content=f"[Could not parse file: {file_path}]",
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