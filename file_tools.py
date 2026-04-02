import os
import subprocess
import glob
from datetime import datetime
try:
    from spire.presentation import Presentation, FileFormat
except ImportError:
    print("spire.presentation not installed, please run: pip install Spire.Presentation")
    Presentation = None

try:
    from pypdf import PdfWriter
except ImportError:
    print("pypdf not installed, please run: pip install pypdf")
    PdfWriter = None

def get_timestamp():
    return datetime.now().strftime("%Y%md_%H%M%S")

def ensure_directories():
    folders = ["Merge_pptx", "Convert_pptx_pdf", "Merge_pdf"]
    for folder in folders:
        if not os.path.exists(folder):
            os.makedirs(folder)
            print(f"Created directory: {folder}")

def merge_pptx_files():
    folder = "Merge_pptx"
    files = glob.glob(os.path.join(folder, "*.pptx"))
    
    if not files:
        print(f"No .pptx files found in '{folder}'")
        return
        
    print(f"Found {len(files)} files to merge.")
    
    if Presentation is None:
        print("Cannot merge: Spire.Presentation is not available.")
        return
        
    print(f"Loading base file: {files[0]}")
    try:
        target_pres = Presentation()
        target_pres.LoadFromFile(files[0])
        
        for file_path in files[1:]:
            print(f"Merging {file_path}...")
            temp_pres = Presentation()
            temp_pres.LoadFromFile(file_path)
            
            for i in range(temp_pres.Slides.Count):
                target_pres.Slides.AppendBySlide(temp_pres.Slides[i])
            temp_pres.Dispose()
            
        output_filename = f"Merged_Output_{get_timestamp()}.pptx"
        output_path = os.path.join(folder, output_filename)
        target_pres.SaveToFile(output_path, FileFormat.Pptx2013)
        target_pres.Dispose()
        print(f"Successfully merged files into {output_path}")
    except Exception as e:
        print(f"An error occurred while merging PPTX: {e}")

def convert_pptx_to_pdf():
    folder = "Convert_pptx_pdf"
    files = glob.glob(os.path.join(folder, "*.pptx"))
    
    if not files:
        print(f"No .pptx files found in '{folder}'")
        return
        
    print(f"Found {len(files)} files to convert.")
    
    for file_path in files:
        print(f"Converting {file_path} to PDF...")
        try:
            # We use libreoffice headless mode to convert the file
            # the format is: libreoffice --headless --convert-to pdf "filename.pptx" --outdir "output_folder"
            result = subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                file_path,
                "--outdir", folder
            ], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            
            if result.returncode == 0:
                print(f"Successfully converted {os.path.basename(file_path)} to PDF.")
            else:
                print(f"Error converting {file_path}:")
                print(result.stderr)
        except FileNotFoundError:
            print("Error: LibreOffice is not installed or not in the PATH. "
                  "Please install LibreOffice ('sudo apt install libreoffice' on Debian/Ubuntu) "
                  "to use this feature.")
            return
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

def merge_pdf_files():
    folder = "Merge_pdf"
    files = glob.glob(os.path.join(folder, "*.pdf"))
    
    if not files:
        print(f"No .pdf files found in '{folder}'")
        return
        
    print(f"Found {len(files)} files to merge.")
    
    if PdfWriter is None:
        print("Cannot merge: pypdf is not available.")
        return

    try:
        merger = PdfWriter()
        for pdf_file in files:
            print(f"Appending {pdf_file}...")
            merger.append(pdf_file)
            
        output_filename = f"Merged_Output_{get_timestamp()}.pdf"
        output_path = os.path.join(folder, output_filename)
        
        merger.write(output_path)
        merger.close()
        print(f"Successfully merged files into {output_path}")
    except Exception as e:
        print(f"An error occurred while merging PDF files: {e}")

def main():
    print("--- PPTX and PDF Utility Tool ---")
    ensure_directories()
    
    while True:
        print("\nPlease choose an option:")
        print("1. Merge all PPTX files in 'Merge_pptx' folder")
        print("2. Convert all PPTX files to PDF in 'Convert_pptx_pdf' folder")
        print("3. Merge all PDF files in 'Merge_pdf' folder")
        print("4. Exit")
        choice = input("\nEnter your choice (1/2/3/4): ").strip()
        
        if choice == '1':
            merge_pptx_files()
        elif choice == '2':
            convert_pptx_to_pdf()
        elif choice == '3':
            merge_pdf_files()
        elif choice == '4':
            print("Exiting...")
            break
        else:
            print("Invalid choice, please enter 1, 2, 3, or 4.")

if __name__ == "__main__":
    main()
