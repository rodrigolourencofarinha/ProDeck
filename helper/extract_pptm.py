import os
import shutil
import zipfile
import tempfile
import argparse
from pathlib import Path
import subprocess
import sys
import time
import gc

def install_oletools():
    """Install oletools if not already installed"""
    try:
        import oletools
        print("✓ oletools already installed")
    except ImportError:
        print("Installing oletools...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "oletools"])
        print("✓ oletools installed successfully")

def safe_remove(path, max_attempts=5, delay=1):
    """Safely remove file/folder with retry logic"""
    for attempt in range(max_attempts):
        try:
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
            return True
        except (PermissionError, OSError) as e:
            if attempt < max_attempts - 1:
                print(f"  ⏳ Attempt {attempt + 1}/{max_attempts}: Waiting for file access...")
                time.sleep(delay)
                gc.collect()  # Force garbage collection
            else:
                print(f"  ⚠ Warning: Could not remove {path}: {e}")
                return False
    return False

def copy_vba_project(vba_project_path, temp_dir):
    """Copy vbaProject.bin to a new location to avoid locking issues"""
    vba_copy_path = temp_dir / "vbaProject_copy.bin"
    try:
        shutil.copy2(vba_project_path, vba_copy_path)
        return vba_copy_path
    except Exception as e:
        print(f"  ⚠ Warning: Could not copy vbaProject.bin: {e}")
        return vba_project_path

def extract_pptm_vba(pptm_path, output_folder):
    """
    Extract and organize VBA components from PPTM file
    
    Args:
        pptm_path (str): Path to the .pptm file
        output_folder (str): Path where organized files will be saved
    """
    
    # Convert paths to Path objects
    pptm_path = Path(pptm_path)
    output_folder = Path(output_folder)
    
    # Validate input file
    if not pptm_path.exists():
        raise FileNotFoundError(f"PPTM file not found: {pptm_path}")
    
    if not pptm_path.suffix.lower() == '.pptm':
        raise ValueError("File must have .pptm extension")
    
    print("🔍 Checking if PowerPoint is running...")
    # Check if PowerPoint processes are running
    try:
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq POWERPNT.EXE'], 
                              capture_output=True, text=True, shell=True)
        if 'POWERPNT.EXE' in result.stdout:
            print("  ⚠ Warning: PowerPoint is running. Consider closing it to avoid file conflicts.")
            input("  Press Enter to continue anyway, or Ctrl+C to cancel...")
    except:
        pass
    
    # Create output directory structure
    output_folder.mkdir(parents=True, exist_ok=True)
    modules_folder = output_folder / "modules"
    forms_folder = output_folder / "forms"
    classes_folder = output_folder / "classes"
    
    # Clean structured folders safely
    for folder in [modules_folder, forms_folder, classes_folder]:
        if folder.exists():
            safe_remove(folder)
        folder.mkdir(parents=True)

    # Clean top-level ribbon folders
    for ribbon_name in ["customUI", "customUI14"]:
        ribbon_path = output_folder / ribbon_name
        if ribbon_path.exists():
            print(f"🧹 Removing old ribbon folder: {ribbon_name}")
            safe_remove(ribbon_path)
        print(f"📁 Created output structure in: {output_folder}")
        
    # Step 1: Copy PPTM and save as ZIP
    temp_dir = Path(tempfile.mkdtemp())
    zip_path = temp_dir / f"{pptm_path.stem}.zip"
    
    vba_parser = None  # Initialize parser variable
    
    try:
        print("📋 Copying PPTM to ZIP...")
        shutil.copy2(pptm_path, zip_path)
        
        # Step 2: Extract ZIP
        print("📦 Extracting ZIP contents...")
        extract_dir = temp_dir / "extracted"
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # Step 2.5: Copy .frx files from extracted folder to forms_folder
        forms_zip_folder = extract_dir / "ppt" / "forms"
        if forms_zip_folder.exists():
            for frx_file in forms_zip_folder.glob("*.frx"):
                dest = forms_folder / frx_file.name
                shutil.copy2(frx_file, dest)
                print(f"    📦 Copied form binary: {frx_file.name}")
        else:
            print("  ⚠ No embedded .frx forms found in archive")

        
        # Step 3: Copy customUI14 (Ribbon)
        print("🎀 Processing Ribbon (customUI14)...")
        custom_ui_path = extract_dir / "customUI"
        custom_ui14_path = extract_dir / "customUI14"
        
        # Check for both customUI and customUI14
        ribbon_found = False
        for ui_path in [custom_ui14_path, custom_ui_path]:
            if ui_path.exists():
                print(f"  ✓ Found ribbon UI: {ui_path.name}")
                dest_path = output_folder / ui_path.name  # direct to output folder
                shutil.copytree(ui_path, dest_path, dirs_exist_ok=True)
                ribbon_found = True
        
        if not ribbon_found:
            print("  ⚠ No customUI/customUI14 folder found")
        
        # Step 4: Extract VBA bin
        print("🔧 Processing VBA Project...")
        vba_project_path = extract_dir / "ppt" / "vbaProject.bin"
        
        if not vba_project_path.exists():
            print("  ⚠ No vbaProject.bin found - file may not contain VBA code")
            return
        
        print("  ✓ Found vbaProject.bin")
        
        # Copy vbaProject.bin to avoid locking issues
        vba_copy_path = copy_vba_project(vba_project_path, temp_dir)
        
        # Use oletools to extract VBA code
        try:
            from oletools.olevba import VBA_Parser
            
            print("  🔍 Parsing VBA code...")
            vba_parser = VBA_Parser(str(vba_copy_path))
            
            if vba_parser.detect_vba_macros():
                print("  ✓ VBA macros detected")
                
                # Extract each VBA module
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    if vba_code and vba_filename:
                        print(f"    📄 Processing: {vba_filename}")
                        
                        # Clean the filename and remove existing extensions
                        clean_filename = vba_filename.replace('/', '_').replace('\\', '_')
                        
                        # Remove existing extensions to avoid double extensions
                        name_without_ext = os.path.splitext(clean_filename)[0]
                        
                        # Skip non-VBA files
                        if vba_filename.endswith('.txt') or 'P-code' in vba_filename:
                            print(f"      ⚠ Skipping non-VBA file: {vba_filename}")
                            continue
                        
                        # -------- REPLACE the old “Determine file type …” block with this --------
                        # Decide the destination by reading the *content* first, then the extension
                        first_line = vba_code.lstrip().splitlines()[0].upper() if vba_code.strip() else ''
                        is_class_header = first_line.startswith("VERSION") and "CLASS" in first_line

                        if (
                                is_class_header                                     # module header says it is a class
                                or vba_filename.lower().endswith('.cls')            # already has .cls extension
                                or vba_filename in {'ThisPresentation',
                                                    'ThisWorkbook',
                                                    'ThisDocument'}                 # built‑in class modules
                        ):
                            dest_folder = classes_folder
                            file_ext = '.cls'

                        elif vba_filename.lower().endswith('.frm'):
                            dest_folder = forms_folder
                            file_ext = '.frm'

                        elif vba_filename.lower().endswith('.frx'):
                            dest_folder = forms_folder
                            file_ext = '.frx'

                        else:
                            # everything else is treated as a standard module
                            dest_folder = modules_folder
                            file_ext = '.bas'
                        # -------------------------------------------------------------------------
                        # Save the VBA code
                        output_file = dest_folder / f"{name_without_ext}{file_ext}"
                        
                        # Add header comment
                        header = f"' Extracted from: {vba_filename}\n' Source: {pptm_path.name}\n\n"
                        
                        with open(output_file, 'w', encoding='utf-8', errors='ignore') as f:
                            f.write(header + vba_code)
                        
                        print(f"      ✓ Saved: {output_file.name}")
            else:
                print("  ⚠ No VBA macros found in vbaProject.bin")
        
        except ImportError:
            print("  ❌ oletools not available. Installing...")
            install_oletools()
            print("  ℹ Please run the script again after oletools installation")
            return
        except Exception as e:
            print(f"  ❌ Error parsing VBA: {str(e)}")
        finally:
            # Close the VBA parser to release file handles
            if vba_parser:
                try:
                    vba_parser.close()
                except:
                    pass
            vba_parser = None
            gc.collect()  # Force garbage collection
        
        # Step 6: Clean up temporary files
        print("🧹 Cleaning up temporary files...")
        
    finally:
        # Ensure VBA parser is closed
        if vba_parser:
            try:
                vba_parser.close()
            except:
                pass
        
        # Force garbage collection before cleanup
        gc.collect()
        time.sleep(0.5)  # Give Windows time to release file handles
        
        # Always clean up temp directory with retry logic
        if temp_dir.exists():
            print("  🗑 Removing temporary files...")
            safe_remove(temp_dir)
    
    # Summary
    print("\n" + "="*50)
    print("📊 EXTRACTION SUMMARY")
    print("="*50)
    
    def count_files(folder):
        return len(list(folder.glob('*'))) if folder.exists() else 0
    
    print(f"📁 Output folder: {output_folder}")
    print(f"🔧 Modules (.bas): {count_files(modules_folder)} files")
    print(f"📝 Classes (.cls): {count_files(classes_folder)} files")
    print(f"🖼 Forms (.frm): {count_files(forms_folder)} files")
    ribbon_ui_paths = [output_folder / "customUI", output_folder / "customUI14"]
    ribbon_count = sum(count_files(path) for path in ribbon_ui_paths if path.exists())
    print(f"🎀 Ribbon files: {ribbon_count} files")
    print("\n✅ Extraction completed successfully!")

def main():
    """Main function with example usage"""

    script_dir   = Path(__file__).resolve().parent
    project_root = script_dir.parent
    output_dir  = project_root 

    # Example usage - modify these paths
    pptm_file = r"C:\path\to\your\presentation.pptm"  # Change this path
    # output_dir = r"C:\path\to\output\folder"          # Change this path
    
    print("🚀 PPTM VBA Extractor")
    print("="*30)
    
    # Check if oletools is installed
    install_oletools()
    
    try:
        # Get user input if paths are not set
        if not os.path.exists(pptm_file):
            pptm_file = input("Enter path to PPTM file: ").strip('"')
        
        if not output_dir or output_dir == r"C:\path\to\output\folder":
            output_dir = input("Enter output folder path: ").strip('"')
        
        # Run extraction
        extract_pptm_vba(pptm_file, output_dir)
        
    except KeyboardInterrupt:
        print("\n❌ Operation cancelled by user")
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        print("Please check your file paths and try again.")

if __name__ == "__main__":
    main()