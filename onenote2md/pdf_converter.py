"""
PDF to Markdown converter.
Uses pdfminer.six for text extraction.
"""

from pathlib import Path
from typing import List, Optional
import re


def convert_pdf_to_text(pdf_path: str) -> str:
    """
    Extract text from PDF file.
    
    Args:
        pdf_path: Path to PDF file
        
    Returns:
        Extracted text content
    """
    try:
        from pdfminer.high_level import extract_text
        return extract_text(pdf_path)
    except ImportError:
        return _extract_pdf_fallback(pdf_path)


def _extract_pdf_fallback(pdf_path: str) -> str:
    """Fallback PDF extraction using basic methods."""
    try:
        import PyPDF2
        
        text = []
        with open(pdf_path, 'rb') as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text.append(page.extract_text())
                
        return '\n\n'.join(text)
    except:
        return "[Could not extract text from PDF. Install pdfminer.six or PyPDF2]"


def pdf_to_markdown(pdf_path: str, output_dir: str = "./output") -> str:
    """
    Convert PDF to Markdown file.
    
    Args:
        pdf_path: Path to PDF file
        output_dir: Output directory
        
    Returns:
        Path to created Markdown file
    """
    # Extract text
    text = convert_pdf_to_text(pdf_path)
    
    # Convert to Markdown
    markdown = _text_to_markdown(text)
    
    # Save to file
    output_path = Path(output_dir) / f"{Path(pdf_path).stem}.md"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(markdown)
        
    return str(output_path)


def _text_to_markdown(text: str) -> str:
    """Convert plain text to Markdown format."""
    
    lines = text.split('\n')
    markdown_lines = []
    
    for line in lines:
        line = line.strip()
        
        if not line:
            markdown_lines.append('')
            continue
            
        # Detect headings (short lines in all caps or short lines)
        if len(line) < 60 and line.isupper():
            markdown_lines.append(f"## {line}")
        elif len(line) < 50 and not line.endswith(('.', ',', ':', ';')):
            markdown_lines.append(f"### {line}")
        else:
            # Regular paragraph
            markdown_lines.append(line)
            
    return '\n'.join(markdown_lines)


def list_pdf_files(folder: str) -> List[Path]:
    """List all PDF files in a folder."""
    return list(Path(folder).rglob("*.pdf"))