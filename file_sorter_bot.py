import os
import shutil

# LIBRARIES
try:
    from pypdf import PdfReader
except ImportError:
    print("CRITICAL: Please install 'pypdf' via Pip.")
    exit()

# --- CONFIGURATION ---
FOLDER_ROOT = "/storage/emulated/0/_0Printing Workflow"
INBOX_FOLDER = os.path.join(FOLDER_ROOT, "_Inbox_Dump")

# TARGET FOLDERS
TARGET_STD = os.path.join(FOLDER_ROOT, "Standard_PDF")
TARGET_SLIDES = os.path.join(FOLDER_ROOT, "Slides_PDF")
TARGET_PPT = os.path.join(FOLDER_ROOT, "PPTX_Slides")

# --- FUNCTIONS ---

def get_pdf_orientation(filepath):
    """
    Returns 'landscape' or 'portrait'.
    """
    try:
        reader = PdfReader(filepath)
        # Check the first page
        if len(reader.pages) > 0:
            page = reader.pages[0]
            # Handle different pdf versions for width/height
            try:
                w = float(page.mediabox.width)
                h = float(page.mediabox.height)
            except AttributeError:
                w = float(page.mediaBox.getWidth())
                h = float(page.mediaBox.getHeight())
            
            if w > h:
                return 'landscape' # Likely a Slide
            else:
                return 'portrait' # Likely a Standard Doc
    except Exception as e:
        print(f"   ⚠️ Could not read PDF: {filepath} ({e})")
        return 'unknown'
    return 'unknown'

def sort_files():
    print("--- 🧹 SORTER BOT STARTED 🧹 ---")
    
    # 1. Check if Inbox exists
    if not os.path.exists(INBOX_FOLDER):
        os.makedirs(INBOX_FOLDER)
        print(f"❌ '_Inbox_Dump' folder was missing. I created it.")
        print(f"📂 Put your files in: {INBOX_FOLDER}")
        print("   Then run this script again.")
        return

    # 2. Check Target Folders
    for folder in [TARGET_STD, TARGET_SLIDES, TARGET_PPT]:
        if not os.path.exists(folder):
            os.makedirs(folder)

    # 3. Scan Files
    files = os.listdir(INBOX_FOLDER)
    if not files:
        print("💤 Inbox is empty. Nothing to sort.")
        return

    count_moved = 0
    
    print(f"🔎 Found {len(files)} files. Sorting now...\n")

    for filename in files:
        filepath = os.path.join(INBOX_FOLDER, filename)
        
        # Skip directories, only move files
        if os.path.isdir(filepath):
            continue

        file_ext = filename.lower().split('.')[-1]
        
        # --- LOGIC: POWERPOINT ---
        if file_ext in ['pptx', 'ppt']:
            shutil.move(filepath, os.path.join(TARGET_PPT, filename))
            print(f"   [PPT] Moved: {filename}")
            count_moved += 1

        # --- LOGIC: PDF ---
        elif file_ext == 'pdf':
            orientation = get_pdf_orientation(filepath)
            
            if orientation == 'landscape':
                # Move to Slides
                shutil.move(filepath, os.path.join(TARGET_SLIDES, filename))
                print(f"   [PDF-SLIDE] Moved: {filename}")
                count_moved += 1
            elif orientation == 'portrait':
                # Move to Standard
                shutil.move(filepath, os.path.join(TARGET_STD, filename))
                print(f"   [PDF-STD] Moved: {filename}")
                count_moved += 1
            else:
                print(f"   [SKIP] Could not determine PDF type: {filename}")

        # --- LOGIC: OTHERS ---
        else:
            print(f"   [SKIP] Unknown file type: {filename}")

    print(f"\n✅ Done! Moved {count_moved} files.")
    print("👉 Now run your Main Script to calculate prices.")

if __name__ == "__main__":
    sort_files()