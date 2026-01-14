"""
Run extraction for a specific project.

Usage:
    python run_extraction.py <project_number>
    python run_extraction.py 01
    python run_extraction.py all
"""

import sys
import importlib
from pathlib import Path

PROJECTS_DIR = Path(__file__).parent / 'projects'


def list_available_projects():
    """List all available project extraction scripts."""
    scripts = sorted(PROJECTS_DIR.glob('p[0-9][0-9]_*.py'))
    return scripts


def run_project(project_num: str):
    """Run extraction for a specific project."""
    # Find matching script
    scripts = list(PROJECTS_DIR.glob(f'p{project_num}_*.py'))
    
    if not scripts:
        print(f"No extraction script found for project {project_num}")
        print("\nAvailable projects:")
        for script in list_available_projects():
            print(f"  - {script.stem}")
        return False
    
    script = scripts[0]
    print(f"Running: {script.name}\n")
    
    # Import and run
    module_name = f"projects.{script.stem}"
    spec = importlib.util.spec_from_file_location(module_name, script)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    
    if hasattr(module, 'run'):
        module.run()
        return True
    else:
        print(f"Error: {script.name} has no run() function")
        return False


def run_all():
    """Run extraction for all projects."""
    scripts = list_available_projects()
    print(f"Running extraction for {len(scripts)} projects\n")
    
    for script in scripts:
        project_num = script.stem[1:3]  # Extract "01" from "p01_..."
        print(f"\n{'#' * 60}")
        run_project(project_num)
    
    print(f"\n{'#' * 60}")
    print("All extractions complete!")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        print("\nAvailable projects:")
        for script in list_available_projects():
            print(f"  {script.stem[1:3]} - {script.stem[4:]}")
        return
    
    arg = sys.argv[1]
    
    if arg.lower() == 'all':
        run_all()
    else:
        # Zero-pad if needed
        project_num = arg.zfill(2)
        run_project(project_num)


if __name__ == '__main__':
    main()
