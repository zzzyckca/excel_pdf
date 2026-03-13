import win32com.client
import time
import os
import glob
import multiprocessing
import traceback
import pythoncom
import subprocess
from pathlib import Path
from datetime import datetime
import pandas as pd
import sys

# ==========================================
# CONFIGURATION
# ==========================================
# Define your input folder here. PDFs will be saved in the same directory.
INPUT_DIR = r"c:\Users\yckde\Documents\GitHub\excel_pdf\test"
# Define the folder where summary reports will be saved
REPORT_DIR = r"c:\Users\yckde\Documents\GitHub\excel_pdf\reports"
# Configure which drive letter to map the path to bypass long path limits
MAPPED_DRIVE_LETTER = "M:"
# Set the maximum number of parallel processes. Leave as "" or None to automatically use (Total CPU Cores - 1).
MAX_PROCESSES = ""
# ==========================================

def map_drive(input_dir):
    """
    Maps a folder to MAPPED_DRIVE_LETTER using 'net use'.
    Returns the mapped path (e.g., 'M:\\') or None if it failed.
    """
    input_dir = os.path.abspath(input_dir)
    
    # Proactively unmap the drive in case it was left over from a previous crash
    subprocess.run(f'net use {MAPPED_DRIVE_LETTER} /delete /y', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    print(f"\n[*] Mapping temporary drive '{MAPPED_DRIVE_LETTER}' to '{input_dir}' using 'net use'...")
    try:
        cmd = f'net use {MAPPED_DRIVE_LETTER} "{input_dir}" /persistent:no'
        subprocess.run(cmd, shell=True, check=True, stdout=subprocess.DEVNULL)
        return MAPPED_DRIVE_LETTER + "\\"
    except subprocess.CalledProcessError as e:
        print(f"[!] Failed to map drive with net use: {e}. Proceeding with original path.")
        return None

def unmap_drive(input_dir):
    """Unmaps the MAPPED_DRIVE_LETTER and the input_dir network connection."""
    print(f"[*] Unmapping temporary drive '{MAPPED_DRIVE_LETTER}'...")
    try:
        subprocess.run(f'net use {MAPPED_DRIVE_LETTER} /delete /y', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print(f"[!] Failed to unmap drive {MAPPED_DRIVE_LETTER}: {e}")
        
    print(f"[*] Disconnecting network path '{input_dir}'...")
    try:
        subprocess.run(f'net use "{input_dir}" /delete /y', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        print(f"[!] Failed to disconnect path: {e}")

def convert_file(file_path):
    """
    Worker function to convert a single Excel file to PDF.
    Runs in a dedicated background process.
    """
    
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
    file_name_with_ext = os.path.basename(file_path)
    file_name = os.path.splitext(file_name_with_ext)[0]
    
    # Generate output path
    output_pdf_dir = os.path.dirname(os.path.abspath(file_path))
    output_pdf_path = os.path.join(output_pdf_dir, f"{file_name}.pdf")
    
    start_time = time.time()
    start_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
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
    end_dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    elapsed = end_time - start_time
    
    return success, file_name_with_ext, abs_file_path, error_msg, start_dt, end_dt, elapsed

def main():
    # Force UTF-8 encoding in the Windows Console to prevent charmap UnicodeEncodeErrors with filenames
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        
    print("-" * 50)
    print("Excel to PDF Batch Converter (Multiprocessing)")
    print("-" * 50)

    # Use global constants
    input_dir = INPUT_DIR
    
    if not os.path.exists(input_dir):
        print(f"ERROR: Input folder '{input_dir}' does not exist.")
        return
        
    mapped_successfully = False
    
    # Map the drive dynamically right at the start before any processing
    mapped_path = map_drive(input_dir)
    if mapped_path:
        input_dir = mapped_path
        mapped_successfully = True

    try:
        success_count = 0
        error_count = 0
        failed_files = []
        summary_data = []
    
        overall_start_time = time.time()
        
        mapped_path_obj = Path(input_dir)
        
        # Search the files inside the globally mapped path
        excel_files = []
        for ext in ['*.xls', '*.xlsx', '*.xlsm']:
            excel_files.extend([str(p) for p in mapped_path_obj.rglob(ext)])
            
        if not excel_files:
            print(f"No Excel files found in {input_dir}")
            return
            
        print(f"Found {len(excel_files)} Excel file(s).")
        
        # Determine number of processes
        if MAX_PROCESSES != "" and MAX_PROCESSES is not None and str(MAX_PROCESSES).strip() != "":
            try:
                num_processes = max(1, int(MAX_PROCESSES))
                print(f"Starting conversion using manually specified {num_processes} parallel process(es)...")
            except ValueError:
                num_processes = max(1, multiprocessing.cpu_count() - 1)
                print(f"[!] Invalid MAX_PROCESSES value. Auto-detecting {num_processes} parallel process(es)...")
        else:
            num_processes = max(1, multiprocessing.cpu_count() - 1)
            print(f"Starting conversion using auto-detected {num_processes} parallel process(es) (Cores - 1)...")
            
        print("-" * 60)
        
        # Run the multiprocessing pool
        with multiprocessing.Pool(processes=num_processes) as pool:
            # pool.imap_unordered yields results as soon as they are ready
            for result in pool.imap_unordered(convert_file, excel_files):
                success, filename, abs_path, err_msg, start_dt, end_dt, elapsed = result
                
                # Translate the mapped drive back to the original path for the report
                original_full_path = abs_path.replace(MAPPED_DRIVE_LETTER, INPUT_DIR, 1)
                
                # Add to pandas collection dictionary
                row_status = "SUCCESS" if success else "FAILED"
                summary_data.append({
                    "file_name": filename,
                    "file_full_path": original_full_path,
                    "status": row_status,
                    "start_time": start_dt,
                    "completed/failure_time": end_dt,
                    "number of seconds": round(elapsed, 2)
                })
                
                if success:
                    print(f"[OK] SUCCESS : {filename} (in {elapsed:.2f}s)")
                    success_count += 1
                else:
                    print(f"[FAIL] ERROR   : {filename} - {err_msg}")
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
        
        # Export to Pandas DataFrame and Excel
        if summary_data:
            print("\nExporting Summary Report to Excel...")
            df = pd.DataFrame(summary_data)
            
            # Create the report directory if it does not exist
            if not os.path.exists(REPORT_DIR):
                os.makedirs(REPORT_DIR)
                
            # Generate the report filename based on the layout: report_YYYY_MM_DD_HH_MM.xlsx
            start_dt_obj = datetime.fromtimestamp(overall_start_time)
            report_filename = start_dt_obj.strftime("report_%Y_%m_%d_%H_%M.xlsx")
            
            # Save it explicitly to the REPORT_DIR
            report_path = os.path.join(REPORT_DIR, report_filename)
            
            try:
                df.to_excel(report_path, index=False)
                print(f"[OK] Saved Excel report to: {report_path}")
            except Exception as e:
                print(f"[FAIL] Failed to save Excel report: {e}")

    finally:
        # Guarantee the drive is unmapped at the very end regardless of success or crash
        if mapped_successfully:
            unmap_drive(INPUT_DIR)

if __name__ == "__main__":
    # Ensure Windows multiprocessing compatibility
    multiprocessing.freeze_support()
    main()
