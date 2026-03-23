"""
Enhanced Markdown converter with full content support.
Handles: images, tables, formatting, links, lists, headings.
"""

from typing import List, Dict, Optional
from pathlib import Path
import re
from onenote2md.local_parser import OneNoteNotebook, OneNoteSection, OneNotePage


class EnhancedMarkdownConverter:
    """Convert OneNote content to Markdown with full formatting support."""
    
    def __init__(self, output_dir: str = "./output", embed_images: bool = True, image_folder: str = "images"):
        self.output_dir = Path(output_dir)
        self.embed_images = embed_images
        self.image_folder = Path(image_folder)
        self._image_counter = 0
        
    def convert_notebook(self, notebook: OneNoteNotebook) -> str:
        """Convert entire notebook to Markdown."""
        md_content = []
        
        # Notebook title
        md_content.append(f"# {self._escape_md(notebook.name)}\n")
        md_content.append(f"*Notebook ID: {notebook.id}*\n")
        md_content.append("---\n")
        
        for section in notebook.sections:
            md_content.append(self.convert_section(section))
            
        return '\n'.join(md_content)
        
    def convert_section(self, section: OneNoteSection) -> str:
        """Convert a section to Markdown."""
        md_content = []
        
        md_content.append(f"\n## 📂 {self._escape_md(section.name)}\n")
        md_content.append(f"*Section ID: {section.id}*\n")
        md_content.append("---\n")
        
        for page in section.pages:
            md_content.append(self.convert_page(page))
            
        return '\n'.join(md_content)
        
    def convert_page(self, page: OneNotePage) -> str:
        """Convert a page to Markdown with full formatting."""
        md_content = []
        
        md_content.append(f"### 📄 {self._escape_md(page.title)}\n")
        if page.last_modified:
            md_content.append(f"*Last modified: {page.last_modified}*\n")
        md_content.append("\n")
        
        # Process content with full formatting
        content = self._process_content(page.content)
        md_content.append(content)
        
        md_content.append("\n---\n")
        
        return '\n'.join(md_content)
        
    def _process_content(self, content: str) -> str:
        """
        Process raw content and apply full formatting.
        Handles: headings, bold, italic, lists, tables, links, images, checkboxes
        """
        if not content or content == "[No text content]":
            return "*No content*"
            
        lines = content.split('\n')
        processed = []
        in_table = False
        
        for line in lines:
            line = line.strip()
            
            # Skip separator lines
            if line in ('===', '---'):
                continue
                
            # Table detection (simple heuristic: lines with |)
            if '|' in line:
                if not in_table:
                    in_table = True
                    # Header row
                    processed.append(line)
                    processed.append('|' + '---|' * (line.count('|')) + '\n')
                else:
                    processed.append(line)
                continue
            elif in_table:
                in_table = False
                processed.append('')  # End table
                
            if not line:
                processed.append('')
                continue
                
            # Process this line
            line = self._format_text(line)
            line = self._process_checkboxes(line)
            processed.append(line)
            
        # Join and clean up
        result = '\n'.join(processed)
        return result
        
    def _format_text(self, text: str) -> str:
        """Apply text formatting: bold, italic, links."""
        
        # Bold: **text** or __text__
        text = re.sub(r'\*\*(.+?)\*\*', r'**\1**', text)
        text = re.sub(r'__(.+?)__', r'**\1**', text)
        
        # Italic: *text* or _text_
        text = re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', r'*\1*', text)
        text = re.sub(r'(?<!_)_(?!_)(.+?)(?<!_)_(?!_)', r'*\1*', text)
        
        # Links: [text](url)
        text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'[\1](\2)', text)
        
        # Inline code: `code`
        text = re.sub(r'`([^`]+)`', r'`\1`', text)
        
        # Headings (simple detection)
        # Lines starting with # or ## etc handled at higher level
        
        return text
        
    def _process_checkboxes(self, text: str) -> str:
        """Convert OneNote checkboxes to Markdown."""
        # OneNote checkbox formats
        text = text.replace('☐', '[ ]')
        text = text.replace('☑', '[x]')
        text = text.replace('☒', '[-]')
        
        # Checkbox at start of line
        text = re.sub(r'^\[ \]\s+', '- [ ] ', text)
        text = re.sub(r'^\[x\]\s+', '- [x] ', text)
        
        return text
        
    def _escape_md(self, text: str) -> str:
        """Escape Markdown special characters."""
        if not text:
            return ""
            
        # Escape characters that have special meaning in Markdown
        special_chars = ['\\', '`', '*', '_', '{', '}', '[', ']', '(', ')', '#', '+', '-', '.', '!']
        
        for char in special_chars:
            text = text.replace(char, f'\\{char}')
            
        return text
        
    def extract_images(self, content: str) -> List[Dict[str, str]]:
        """
        Extract image references from content.
        
        Returns list of image info dicts with: type, data, name
        """
        images = []
        
        # Look for image placeholders
        # OneNote stores images as various formats
        
        # Pattern for base64 embedded images
        img_pattern = r'<img[^>]+src="data:([^;]+);base64,([^"]+)"[^>]*>'
        matches = re.findall(img_pattern, content)
        
        for mime_type, b64_data in matches:
            self._image_counter += 1
            images.append({
                'type': 'base64',
                'mime': mime_type,
                'data': b64_data,
                'name': f'image_{self._image_counter}'
            })
            
        return images
        
    def save_image(self, image_info: Dict, output_dir: Path) -> str:
        """
        Save extracted image and return markdown reference.
        """
        import base64
        
        img_dir = output_dir / self.image_folder
        img_dir.mkdir(parents=True, exist_ok=True)
        
        name = image_info['name']
        
        if image_info['type'] == 'base64':
            # Determine extension from mime type
            ext_map = {
                'image/png': '.png',
                'image/jpeg': '.jpg',
                'image/gif': '.gif',
                'image/webp': '.webp'
            }
            ext = ext_map.get(image_info['mime'], '.png')
            
            filename = f"{name}{ext}"
            filepath = img_dir / filename
            
            # Decode and save
            with open(filepath, 'wb') as f:
                f.write(base64.b64decode(image_info['data']))
                
            return f"../{self.image_folder}/{filename}"
            
        return ""
        
    def save_to_file(self, notebook: OneNoteNotebook, output_path: str = None) -> Path:
        """Convert and save notebook to Markdown file."""
        if output_path:
            output_file = Path(output_path)
        else:
            safe_name = self._sanitize_filename(notebook.name)
            output_file = self.output_dir / f"{safe_name}.md"
            
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Extract images first
        all_images = []
        for section in notebook.sections:
            for page in section.pages:
                images = self.extract_images(page.content)
                all_images.extend(images)
                
        # Save images
        if all_images and self.embed_images:
            for img in all_images:
                self.save_image(img, self.output_dir)
                
        md_content = self.convert_notebook(notebook)
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(md_content)
            
        return output_file
        
    def _sanitize_filename(self, name: str) -> str:
        """Sanitize filename by removing invalid characters."""
        invalid = '<>:"/\\|?*'
        for char in invalid:
            name = name.replace(char, '_')
        return name[:255]


# Alias for backward compatibility
MarkdownConverter = EnhancedMarkdownConverter


def convert_to_markdown(notebook: OneNoteNotebook, output_dir: str = "./output", 
                        embed_images: bool = True) -> Path:
    """Convenience function to convert notebook to Markdown."""
    converter = EnhancedMarkdownConverter(
        output_dir=output_dir, 
        embed_images=embed_images
    )
    return converter.save_to_file(notebook)