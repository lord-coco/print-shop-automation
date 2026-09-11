import os
import shutil
from datetime import datetime

# ==========================================
# 🛠️ LIBRARIES
# ==========================================
try:
    from pypdf import PdfReader, PdfWriter, PageObject
except ImportError:
    print("CRITICAL ERROR: Please install 'pypdf' in Pydroid pip menu.")
    exit()

try:
    from pptx import Presentation
except ImportError:
    print("CRITICAL ERROR: Please install 'python-pptx' in Pydroid pip menu.")
    exit()

# ==========================================
# ⚙️ CONFIGURATION
# ==========================================
PRICE_PER_SHEET_SELLING = 35
PRICE_PER_SHEET_COST = 25
FOLDER_ROOT = "/storage/emulated/0/_0Printing Workflow" 

# ==========================================
# 🔧 FUNCTIONS
# ==========================================

def get_next_even(n):
    return n if n % 2 == 0 else n + 1

def get_next_multiple_of_8(n):
    if n % 8 == 0:
        return n
    return ((n // 8) + 1) * 8

def get_page_dims(page):
    try:
        return float(page.mediabox.width), float(page.mediabox.height)
    except AttributeError:
        return float(page.mediaBox.getWidth()), float(page.mediaBox.getHeight())

# ==========================================
# 🚀 MAIN PROCESS
# ==========================================
def process_job():
    # 1. ASK FOR CLIENT NAME
    print("--- 🖨️ NEW PRINT JOB 🖨️ ---")
    client_name_raw = input("Enter Client Name: ").strip()
    if not client_name_raw:
        client_name_raw = "Unknown_Client"
    
    # Sanitize name
    client_name = "".join(c for c in client_name_raw if c.isalnum() or c in (' ', '_', '-')).strip()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    readable_date = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 2. CREATE CLIENT OUTPUT FOLDER
    client_folder = os.path.join(FOLDER_ROOT, client_name)
    if not os.path.exists(client_folder):
        os.makedirs(client_folder)
        
    ppt_output_folder = os.path.join(client_folder, "PPT_Processed")
    if not os.path.exists(ppt_output_folder):
        os.makedirs(ppt_output_folder)

    # New Folder for Unmerged Slides (if user says No to merge)
    slides_output_folder = os.path.join(client_folder, "Slides_Processed")

    print(f"\n📂 Working in folder: {client_name}...")

    # Counters
    count_std_files = 0
    sheets_standard = 0
    
    count_slide_files = 0
    sheets_pdf_slides = 0
    
    count_pptx_files = 0
    sheets_pptx = 0
    
    log_details = []
    
    # Archiving Lists
    files_to_archive_standard = []
    files_to_archive_slides = []
    files_to_archive_pptx = []

    # Setup Mergers
    merger_standard = PdfWriter()
    merger_slides = PdfWriter()
    has_standard = False
    has_slides = False

    # --- PROCESSING STANDARD PDFS (ALWAYS MERGE) ---
    folder_standard = os.path.join(FOLDER_ROOT, "Standard_PDF")
    if os.path.exists(folder_standard):
        for filename in sorted(os.listdir(folder_standard)):
            if filename.lower().endswith(".pdf"):
                filepath = os.path.join(folder_standard, filename)
                try:
                    reader = PdfReader(filepath)
                    page_count = len(reader.pages)
                    padded_count = get_next_even(page_count)
                    
                    merger_standard.append(reader)
                    if padded_count > page_count:
                        last_page = reader.pages[-1]
                        w, h = get_page_dims(last_page)
                        merger_standard.add_blank_page(width=w, height=h)
                    
                    s = padded_count / 2
                    sheets_standard += s
                    count_std_files += 1
                    
                    log_details.append(f"[STD] {filename}: {int(s)} papers")
                    has_standard = True
                    files_to_archive_standard.append(filename)
                except Exception as e:
                    print(f"ERROR {filename}: {e}")

    # --- PROCESSING SLIDE PDFS (PREPARE DATA ONLY) ---
    folder_slides = os.path.join(FOLDER_ROOT, "Slides_PDF")
    if os.path.exists(folder_slides):
        for filename in sorted(os.listdir(folder_slides)):
            if filename.lower().endswith(".pdf"):
                filepath = os.path.join(folder_slides, filename)
                try:
                    reader = PdfReader(filepath)
                    page_count = len(reader.pages)
                    padded_count = get_next_multiple_of_8(page_count)
                    
                    # Add to memory but don't write yet
                    merger_slides.append(reader)
                    pages_to_add = padded_count - page_count
                    if pages_to_add > 0:
                        last_page = reader.pages[-1]
                        w, h = get_page_dims(last_page)
                        for _ in range(pages_to_add):
                            merger_slides.add_blank_page(width=w, height=h)
                            
                    s = padded_count / 8
                    sheets_pdf_slides += s
                    count_slide_files += 1
                    
                    log_details.append(f"[SLD-PDF] {filename}: {int(s)} papers")
                    has_slides = True
                    files_to_archive_slides.append(filename)
                except Exception as e:
                    print(f"ERROR {filename}: {e}")

    # --- PROCESSING PPTX AND PPT SLIDES ---
    folder_pptx = os.path.join(FOLDER_ROOT, "PPTX_Slides")
    if os.path.exists(folder_pptx):
        for filename in sorted(os.listdir(folder_pptx)):
            filepath = os.path.join(folder_pptx, filename)
            
            # CASE 1: Modern .pptx (Automatic)
            if filename.lower().endswith(".pptx"):
                try:
                    prs = Presentation(filepath)
                    slide_count = len(prs.slides)
                    padded_count = get_next_multiple_of_8(slide_count)
                    
                    s = padded_count / 8
                    sheets_pptx += s
                    count_pptx_files += 1
                    
                    log_details.append(f"[PPTX] {filename}: {int(s)} papers")
                    
                    shutil.copy2(filepath, os.path.join(ppt_output_folder, filename))
                    files_to_archive_pptx.append(filename)
                    
                except Exception as e:
                    print(f"ERROR {filename}: {e}")
            
            # CASE 2: Old .ppt (Manual Input Required)
            elif filename.lower().endswith(".ppt"):
                print(f"\n⚠️  Found old format file: {filename}")
                try:
                    user_input = input(f"   How many slides in '{filename}'? (Enter number): ")
                    slide_count = int(user_input)
                except ValueError:
                    print("   Invalid number entered. Assuming 0.")
                    slide_count = 0
                
                padded_count = get_next_multiple_of_8(slide_count)
                s = padded_count / 8
                sheets_pptx += s
                count_pptx_files += 1
                
                log_details.append(f"[PPT-OLD] {filename}: {slide_count} slides -> {int(s)} papers")
                
                shutil.copy2(filepath, os.path.join(ppt_output_folder, filename))
                files_to_archive_pptx.append(filename)

    # --- CALCULATIONS ---
    total_files = count_std_files + count_slide_files + count_pptx_files
    total_sheets = sheets_standard + sheets_pdf_slides + sheets_pptx
    
    price_standard = sheets_standard * PRICE_PER_SHEET_SELLING
    price_pdf_slides = sheets_pdf_slides * PRICE_PER_SHEET_SELLING
    price_pptx = sheets_pptx * PRICE_PER_SHEET_SELLING
    
    total_revenue = total_sheets * PRICE_PER_SHEET_SELLING
    total_cost = total_sheets * PRICE_PER_SHEET_COST
    profit = total_revenue - total_cost

    # --- DECISION TIME: MERGE SLIDES? ---
    merge_slides_decision = False
    if has_slides:
        print(f"\n⚠️  PDF Slides Found: {count_slide_files} files.")
        user_choice = input("   Do you want to merge them into one file? (y/n): ").lower().strip()
        if user_choice == 'y':
            merge_slides_decision = True
        else:
            merge_slides_decision = False
            print("   👉 Skipping merge. Copying individual files instead.")

    # --- SAVE MERGED FILES ---
    if has_standard:
        out_name = f"{client_name}_Standard_{timestamp}.pdf"
        merger_standard.write(os.path.join(client_folder, out_name))
        merger_standard.close()
        print(f"\n✅ Merged Standard PDF Saved.")

    if has_slides:
        if merge_slides_decision:
            # Merge Mode
            out_name = f"{client_name}_Slides_{timestamp}.pdf"
            merger_slides.write(os.path.join(client_folder, out_name))
            merger_slides.close()
            print(f"✅ Merged Slides PDF Saved.")
        else:
            # Copy Mode
            if not os.path.exists(slides_output_folder):
                os.makedirs(slides_output_folder)
            
            for f in files_to_archive_slides:
                src = os.path.join(folder_slides, f)
                dst = os.path.join(slides_output_folder, f)
                shutil.copy2(src, dst)
            print(f"✅ Individual Slide PDFs copied to 'Slides_Processed' folder.")
            merger_slides.close()

    # --- GENERATE INVOICE TEXT ---
    invoice_text = f"""🧾 *PAYMENT BREAKDOWN*
👤 *Customer:* {client_name_raw}
📅 *Date:* {readable_date}

1️⃣ *Standard Printing*
• {count_std_files} Documents
• {int(sheets_standard)} A4 Papers
💵 *₦{int(price_standard):,}*

2️⃣ *PDF Slides (4/pg)*
• {count_slide_files} Documents
• {int(sheets_pdf_slides)} A4 Papers
💵 *₦{int(price_pdf_slides):,}*

3️⃣ *PPTX/PPT Slides (4/pg)*
• {count_pptx_files} Documents
• {int(sheets_pptx)} A4 Papers
💵 *₦{int(price_pptx):,}*

--------------------------------
📚 *Total Documents:* {total_files}
📄 *Total A4 Papers:* {int(total_sheets)}
💰 *GRAND TOTAL: ₦{int(total_revenue):,}*
================================
"""

    internal_record = f"""
--------------------------------
INTERNAL RECORD (Private)
--------------------------------
Cost: ₦{int(total_cost):,} | Profit: ₦{int(profit):,}

Breakdown:
""" + "\n".join(log_details)

    print(invoice_text)
    
    # Save Invoice to CLIENT Folder
    inv_filename = f"{client_name}_Invoice_{timestamp}.txt"
    invoice_path = os.path.join(client_folder, inv_filename)
    
    with open(invoice_path, "w") as f:
        f.write(invoice_text + internal_record)
    print(f"✅ Payment Breakdown saved to client folder.")

    # --- ARCHIVING / CLEANUP ---
    print("\n🧹 Cleaning up workspace...")
    history_root = os.path.join(FOLDER_ROOT, "_Completed_Jobs_History")
    job_history_folder = os.path.join(history_root, f"{client_name}_{timestamp}")
    
    # Create main history folder
    if not os.path.exists(job_history_folder):
        os.makedirs(job_history_folder)
        
    # Create Subfolders in History
    hist_std = os.path.join(job_history_folder, "Standard_PDF")
    hist_sld = os.path.join(job_history_folder, "Slides_PDF")
    hist_ppt = os.path.join(job_history_folder, "PPTX_Slides")
    
    os.makedirs(hist_std, exist_ok=True)
    os.makedirs(hist_sld, exist_ok=True)
    os.makedirs(hist_ppt, exist_ok=True)

    # Move Files to their specific subfolders
    for f in files_to_archive_standard:
        shutil.move(os.path.join(folder_standard, f), os.path.join(hist_std, f))
        
    for f in files_to_archive_slides:
        shutil.move(os.path.join(folder_slides, f), os.path.join(hist_sld, f))
        
    for f in files_to_archive_pptx:
        shutil.move(os.path.join(folder_pptx, f), os.path.join(hist_ppt, f))
        
    # --- NEW ADDITION: ARCHIVE THE INVOICE TOO ---
    try:
        shutil.copy2(invoice_path, job_history_folder)
        print(f"✅ Invoice archived to History folder.")
    except Exception as e:
        print(f"⚠️ Could not archive invoice: {e}")

    print(f"✅ Source files moved to History (Sorted by type).")
    print(f"📂 Open folder: {client_name} for the Breakdown!")

if __name__ == "__main__":
    if not os.path.exists(FOLDER_ROOT):
        print(f"Folder not found: {FOLDER_ROOT}")
    else:
        process_job()