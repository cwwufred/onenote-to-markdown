import click
import os
from pathlib import Path
from onenote2md import config
from onenote2md.local_parser import parse_one_file, list_one_files
from onenote2md.converter import convert_to_markdown
from onenote2md.batch_export import batch_export

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
    click.echo(f"  Image Folder:   {cfg.get('image_folder')}")
    click.echo(f"  Embed Images:   {cfg.get('embed_images')}")
    click.echo("")

@main.command('list')
def list_files():
    """List all .one files in source folder."""
    source = config.get_source_folder()
    
    if not source:
        click.echo("❌ Source folder not set. Run:")
        click.echo("   onenote2md config set-source /path/to/onenote/backup")
        return
    
    one_files = list_one_files(source)
    
    if not one_files:
        click.echo(f"⚠️ No .one files found in: {source}")
        return
    
    # Show folder structure
    click.echo(f"\n📚 Found {len(one_files)} .one file(s):\n")
    
    # Group by parent folder
    folders = {}
    for f in sorted(one_files):
        parent = f.parent.name
        if parent not in folders:
            folders[parent] = []
        folders[parent].append(f.name)
        
    for folder, files in folders.items():
        click.echo(f"  📁 {folder}/")
        for f in files:
            click.echo(f"      📄 {f}")
    click.echo("")

@main.command('export')
@click.option('--file', '-f', 'file_name', help='Export specific file by name')
@click.option('--all', '-a', 'export_all', is_flag=True, help='Export all files')
@click.option('--output', '-o', 'output_dir', help='Output directory')
@click.option('--structure/--no-structure', default=True, help='Preserve folder structure')
def export(file_name, export_all, output_dir, structure):
    """Export .one files to Markdown."""
    source = config.get_source_folder()
    
    if not source:
        click.echo("❌ Source folder not set. Run:")
        click.echo("   onenote2md config set-source /path/to/onenote/backup")
        return
    
    # Determine output directory
    out_dir = output_dir if output_dir else config.get_output_dir()
    
    if export_all:
        # Batch export all files
        results = batch_export(source, out_dir, preserve_structure=structure)
        
        # Show summary
        successful = sum(1 for r in results if r.success)
        failed = len(results) - successful
        
        if failed > 0:
            click.echo(f"\n⚠️ {failed} file(s) failed:")
            for r in results:
                if not r.success:
                    click.echo(f"  - {r.file.name}: {r.error}")
    else:
        # Single file export
        one_files = list_one_files(source)
        
        if not one_files:
            click.echo("⚠️ No .one files found")
            return
        
        if file_name:
            one_files = [f for f in one_files if file_name.lower() in f.name.lower()]
            if not one_files:
                click.echo(f"⚠️ No file matching '{file_name}' found")
                return
        else:
            click.echo("Use --file <name> to specify a file, or --all for all files")
            click.echo("\nAvailable files:")
            for f in sorted(one_files):
                click.echo(f"  📄 {f.name}")
            return
        
        click.echo(f"\n📥 Exporting {len(one_files)} file(s)...\n")
        
        for f in one_files:
            click.echo(f"  Processing: {f.name}...")
            
            try:
                notebook = parse_one_file(str(f))
                if notebook:
                    output_path = convert_to_markdown(notebook, out_dir)
                    click.echo(f"    ✅ Saved to: {output_path}")
                else:
                    click.echo(f"    ⚠️ Could not parse file")
            except Exception as e:
                click.echo(f"    ❌ Error: {e}")
                
        click.echo(f"\n📂 Output: {out_dir}")

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