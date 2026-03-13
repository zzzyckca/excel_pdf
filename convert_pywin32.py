import win32com.client
import time
import os
import glob
import multiprocessing
import traceback
import pythoncom

# ==========================================
# CONFIGURATION
# ==========================================
# Define your input and output folders here.
INPUT_DIR = r"c:\Users\yckde\Documents\GitHub\excel_pdf\test"
OUTPUT_DIR = r"c:\Users\yckde\Documents\GitHub\excel_pdf\pdf"
# ==========================================

def convert_file(args):
    """
    Worker function to convert a single Excel file to PDF.
    Runs in a dedicated background process.
    """
    file_path, output_folder = args
    
    # Initialize COM for this specific thread
    pythoncom.CoInitialize()
    
    # Needs to be DispatchEx to launch a new, separate Excel process for safety
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
    except Exception as e:
        pythoncom.CoUninitialize()
        return False, file_path, f"Failed to start Excel: {e}", 0
        
    # Strict background execution settings
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.EnableEvents = False
    excel.Interactive = False
    
    abs_file_path = os.path.abspath(file_path)
    file_name_with_ext = os.path.basename(abs_file_path)
    file_name = os.path.splitext(file_name_with_ext)[0]
    output_pdf_path = os.path.abspath(os.path.join(output_folder, f"{file_name}.pdf"))
    
    start_time = time.time()
    wb = None
    success = False
    error_msg = ""
    
    try:
        # Open the workbook (UpdateLinks=0 to prevent popups, ReadOnly=True)
        wb = excel.Workbooks.Open(abs_file_path, UpdateLinks=0, ReadOnly=True)
        
        # 0 corresponds to xlTypePDF
        # Exporting the entire workbook
        wb.ExportAsFixedFormat(0, output_pdf_path)
        success = True
        
    except Exception as e:
        error_msg = str(e)
    finally:
        if wb:
            try:
                # False means do not save changes
                wb.Close(False)
            except:
                pass
        
        try:
            excel.Quit()
        except:
            pass
            
        pythoncom.CoUninitialize()
        
    end_time = time.time()
    elapsed = end_time - start_time
    
    return success, file_name_with_ext, error_msg, elapsed

def main():
    print("-" * 50)
    print("Excel to PDF Batch Converter (Multiprocessing)")
    print("-" * 50)

    # Use global constants
    input_dir = INPUT_DIR
    output_dir = OUTPUT_DIR
    
    if not os.path.exists(input_dir):
        print(f"ERROR: Input folder '{input_dir}' does not exist.")
        return
        
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    excel_files = glob.glob(os.path.join(input_dir, "*.xls*"))
    
    if not excel_files:
        print(f"No Excel files found in {input_dir}")
        return
        
    print(f"Found {len(excel_files)} Excel file(s).")
    
    # Determine number of processes (use max cores - 1 to leave room for OS, or at least 1)
    num_processes = max(1, multiprocessing.cpu_count() - 1)
    print(f"Starting conversion using {num_processes} parallel process(es)...")
    print("-" * 60)
    
    # Prepare arguments for the worker function
    pool_args = [(f, output_dir) for f in excel_files]
    
    success_count = 0
    error_count = 0
    failed_files = []
    
    overall_start_time = time.time()
    
    # Run the multiprocessing pool
    with multiprocessing.Pool(processes=num_processes) as pool:
        # pool.imap_unordered yields results as soon as they are ready
        for result in pool.imap_unordered(convert_file, pool_args):
            success, filename, err_msg, elapsed = result
            if success:
                print(f"[✓] SUCCESS : {filename} (in {elapsed:.2f}s)")
                success_count += 1
            else:
                print(f"[x] ERROR   : {filename} - {err_msg}")
                error_count += 1
                failed_files.append((filename, err_msg))

    overall_end_time = time.time()
    total_elapsed = overall_end_time - overall_start_time
    
    # Summary Report
    print("\n" + "=" * 60)
    print("SUMMARY REPORT")
    print("=" * 60)
    print(f"Total Files Processed : {len(excel_files)}")
    print(f"Successful Conversions: {success_count}")
    print(f"Failed Conversions    : {error_count}")
    print(f"Total Time Elapsed    : {total_elapsed / 60:.2f} minutes ({total_elapsed:.2f} seconds)")
    
    if failed_files:
        print("\nFailed Files Details:")
        for f, err in failed_files:
            print(f" - {f}: {err}")
    print("=" * 60)

if __name__ == "__main__":
    # Ensure Windows multiprocessing compatibility
    multiprocessing.freeze_support()
    main()
