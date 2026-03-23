import click
import os
from pathlib import Path
from onenote2md import config
from onenote2md.local_parser import parse_one_file, list_one_files
from onenote2md.converter import convert_to_markdown
from onenote2md.batch_export import batch_export
from onenote2md.pdf_converter import pdf_to_markdown, list_pdf_files

@click.group()
def main():
    """OneNote to Markdown converter."""
    pass

@main.group()
def config_cmd():
    """Manage configuration."""
    pass

@config_cmd.command('set-source')
@click.argument('path', type=click.Path(exists=True))
def set_source(path):
    """Set the source folder containing .one files."""
    abs_path = os.path.abspath(path)
    config.set_source_folder(abs_path)
    click.echo(f"✅ Source folder set to: {abs_path}")

@config_cmd.command('set-output')
@click.argument('path', type=click.Path())
def set_output(path):
    """Set the output folder for exported files."""
    abs_path = os.path.abspath(path)
    cfg = config.load_config()
    cfg['output_dir'] = abs_path
    config.save_config(cfg)
    click.echo(f"✅ Output folder set to: {abs_path}")

@config_cmd.command('show')
def show_config():
    """Show current configuration."""
    cfg = config.load_config()
    click.echo("\n📁 Current Configuration:\n")
    click.echo(f"  Source Folder: {cfg.get('source_folder') or '(not set)'}")
    click.echo(f"  Output Dir:     {cfg.get('output_dir')}")
    click.echo("")

@main.command('list')
@click.option('--type', type=click.Choice(['one', 'pdf', 'all']), default='one', help='File type to list')
def list_files(type):
    """List files in source folder."""
    source = config.get_source_folder()
    
    if not source:
        click.echo("❌ Source folder not set. Run:")
        click.echo("   onenote2md config set-source /path/to/folder")
        return
    
    if type in ['one', 'all']:
        one_files = list_one_files(source)
        if one_files:
            click.echo(f"\n📚 Found {len(one_files)} .one file(s):\n")
            for f in sorted(one_files):
                click.echo(f"  📄 {f.name}")
                
    if type in ['pdf', 'all']:
        pdf_files = list_pdf_files(source)
        if pdf_files:
            click.echo(f"\n📄 Found {len(pdf_files)} PDF file(s):\n")
            for f in sorted(pdf_files):
                click.echo(f"  📑 {f.name}")

@main.command('export')
@click.option('--one', '-o', 'export_one', is_flag=True, help='Export .one files')
@click.option('--pdf', '-p', 'export_pdf', is_flag=True, help='Export PDF files')
@click.option('--all', '-a', 'export_all', is_flag=True, help='Export all files')
@click.option('--file', '-f', 'file_name', help='Export specific file by name')
@click.option('--output', '-d', 'output_dir', help='Output directory')
def export(export_one, export_pdf, export_all, file_name, output_dir):
    """Export files to Markdown."""
    source = config.get_source_folder()
    
    if not source:
        click.echo("❌ Source folder not set. Run:")
        click.echo("   onenote2md config set-source /path/to/folder")
        return
    
    # Determine output directory
    out_dir = output_dir if output_dir else config.get_output_dir()
    os.makedirs(out_dir, exist_ok=True)
    
    exported_count = 0
    
    # Export .one files
    if export_one or export_all:
        one_files = list_one_files(source)
        if file_name:
            one_files = [f for f in one_files if file_name.lower() in f.name.lower()]
            
        if one_files:
            click.echo(f"\n📥 Exporting {len(one_files)} .one file(s)...\n")
            for f in one_files:
                try:
                    notebook = parse_one_file(str(f))
                    if notebook:
                        convert_to_markdown(notebook, out_dir)
                        click.echo(f"  ✅ {f.name}")
                        exported_count += 1
                except Exception as e:
                    click.echo(f"  ❌ {f.name}: {e}")
    
    # Export PDFs
    if export_pdf or export_all:
        pdf_files = list_pdf_files(source)
        if file_name:
            pdf_files = [f for f in pdf_files if file_name.lower() in f.name.lower()]
            
        if pdf_files:
            click.echo(f"\n📥 Exporting {len(pdf_files)} PDF file(s)...\n")
            for f in pdf_files:
                try:
                    pdf_to_markdown(str(f), out_dir)
                    click.echo(f"  ✅ {f.name}")
                    exported_count += 1
                except Exception as e:
                    click.echo(f"  ❌ {f.name}: {e}")
    
    if exported_count > 0:
        click.echo(f"\n✅ Exported {exported_count} file(s)")
        click.echo(f"📂 Output: {out_dir}")
    else:
        click.echo("⚠️ No files exported")

@main.command('gui')
def launch_gui():
    """Launch the desktop GUI."""
    try:
        from onenote2md.gui import main as gui_main
        gui_main()
    except ImportError as e:
        click.echo(f"❌ GUI not available: {e}")

if __name__ == '__main__':
    main()