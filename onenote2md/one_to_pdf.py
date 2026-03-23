"""
OneNote automation - convert .one files to PDF using COM automation.
"""

import os
import subprocess
from pathlib import Path
from typing import List, Optional


def convert_one_to_pdf(input_file: str, output_dir: str = None) -> Optional[str]:
    """
    Convert .one file to PDF using OneNote COM automation.
    
    Args:
        input_file: Path to .one file
        output_dir: Output directory (default: same as input)
    
    Returns:
        Path to created PDF file, or None if failed
    """
    
    if output_dir is None:
        output_dir = str(Path(input_file).parent)
    
    pdf_path = os.path.join(output_dir, Path(input_file).stem + ".pdf")
    
    try:
        # Try Windows COM automation
        return _convert_via_com(input_file, pdf_path)
    except Exception as e:
        print(f"COM conversion failed: {e}")
        
    try:
        # Try command-line approach
        return _convert_via_cli(input_file, pdf_path)
    except Exception as e:
        print(f"CLI conversion failed: {e}")
        
    return None


def _convert_via_com(input_file: str, pdf_path: str) -> str:
    """Convert using Windows COM automation."""
    
    # Create PowerShell script for OneNote COM
    ps_script = f'''
$oneNote = New-Object -ComObject OneNote.Application
$oneNote.OpenHierarchy("{input_file.Replace("\\", "\\\\")}", $false, $false)
$oneNote.Publish("{input_file.Replace("\\", "\\\\")}", "{pdf_path.replace("\\", "\\\\")}", 2)  # 2 = pfPDF
$oneNote.Windows.CurrentWindow.Close($false)
'''
    
    # Run PowerShell script
    result = subprocess.run(
        ['powershell', '-Command', ps_script],
        capture_output=True,
        text=True
    )
    
    if result.returncode == 0 and os.path.exists(pdf_path):
        return pdf_path
        
    raise Exception("COM conversion failed")


def _convert_via_cli(input_file: str, pdf_path: str) -> str:
    """Convert using OneNote command-line."""
    
    # Try using OneNote's built-in command line export
    # OneNote.exe /export /format:pdf /output:"<output>" "<input>"
    
    onenote_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE",
        r"C:\Program Files\Microsoft Office 15\root\Office15\ONENOTE.EXE",
    ]
    
    for onenote_path in onenote_paths:
        if os.path.exists(onenote_path):
            cmd = [
                onenote_path,
                '/export',
                '/format:pdf',
                f'/output:"{pdf_path}"',
                f'"{input_file}"'
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                return pdf_path
                
    raise Exception("No OneNote installation found")


def batch_convert_one_to_pdf(files: List[str], output_dir: str) -> List[str]:
    """
    Batch convert .one files to PDF.
    
    Args:
        files: List of .one file paths
        output_dir: Output directory
    
    Returns:
        List of created PDF paths
    """
    
    os.makedirs(output_dir, exist_ok=True)
    
    pdf_files = []
    
    for file in files:
        print(f"Converting: {Path(file).name}")
        
        pdf_path = convert_one_to_pdf(file, output_dir)
        
        if pdf_path and os.path.exists(pdf_path):
            pdf_files.append(pdf_path)
            print(f"  ✅ Created: {pdf_path}")
        else:
            print(f"  ❌ Failed: {file}")
            
    return pdf_files


def is_onenote_available() -> bool:
    """Check if OneNote is installed."""
    
    onenote_paths = [
        r"C:\Program Files\Microsoft Office\root\Office16\ONENOTE.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\ONENOTE.EXE",
        r"C:\Program Files\Microsoft Office 15\root\Office15\ONENOTE.EXE",
        r"C:\Program Files (x86)\Microsoft Office 15\root\Office15\ONENOTE.EXE",
    ]
    
    for path in onenote_paths:
        if os.path.exists(path):
            return True
            
    return False