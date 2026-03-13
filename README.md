# Excel to PDF Batch Converter (Multiprocessing Edition)

This Python script (`convert_pywin32.py`) provides a robust, high-performance solution for batch converting large numbers of Excel files (`.xlsx`, `.xls`) to PDF documents. It is specifically designed to run on Windows systems (such as Windows Server 2022) and uses Python's multiprocessing capabilities to dramatically speed up conversion times.

## Key Features

- **Blazing Fast Multiprocessing:** Automatically detects the number of CPU cores (virtual processors) on your machine and runs concurrent Excel instances. For example, on an 8-core server, it processes 7 Excel files simultaneously, leaving 1 core free for system stability.
- **Strict Background Execution:** Uses advanced COM configuration (`DispatchEx`) to launch completely isolated, hidden instances of Excel. It strictly enforces background rules to prevent open pop-ups, alerts, or macros from interrupting the batch process.
- **Entire Workbook Export:** Ensures that every visible worksheet within an Excel file is exported into a single, cohesive PDF document.
- **Comprehensive Summary Report:** Consolidates all success and error messages into a clean summary printed to the console upon completion.

## System Requirements

- **Operating System:** Windows (e.g., Windows 10, Windows 11, Windows Server 2019/2022).
- **Microsoft Excel:** A licensed version of Microsoft Excel *must* be installed on the machine. (The script interacts directly with the actual Excel desktop application under the hood).
- **Python:** Python 3.7 or higher.
- **Python Packages:** You must install the `pywin32` library to communicate with the Excel application.
  ```bash
  pip install pywin32
  ```

## Usage

1. Move the `convert_pywin32.py` script to your local machine or Windows Server.
2. Open `convert_pywin32.py` in your chosen text editor or IDE.
3. Modify the `INPUT_DIR` path located at the very top of the script in the Configuration section to point to your folder containing Excel files. The script will automatically save the generated PDFs in the same folder alongside their original Excel files:
   ```python
   # ==========================================
   # CONFIGURATION
   # ==========================================
   # Define your input folder here. PDFs will be saved in the same directory.
   INPUT_DIR = r"c:\Users\yckde\Documents\GitHub\excel_pdf\test"
   # ==========================================
   ```
4. Run the script from the command prompt, PowerShell, or your IDE:
   ```bash
   python convert_pywin32.py
   ```

## How It Works Under the Hood

The script utilizes a combination of built-in Python multiprocessing and direct COM (Component Object Model) interactions to achieve true parallel processing of Excel files.

### 1. Process Management (`multiprocessing`)
The script uses `multiprocessing.Pool` to manage concurrent workers. The pool size is determined dynamically:
`num_processes = max(1, multiprocessing.cpu_count() - 1)`
This ensures maximum throughput (e.g., 7 workers on an 8-core machine) while dedicating 1 core to system stability. `pool.imap_unordered` is used to stream results back as soon as each individual file finishes, rather than waiting for the entire batch to complete before displaying feedback.

### 2. COM Initialization (`pythoncom`)
Because the Component Object Model (COM) in Windows is thread-aware and process-bound, each worker process *must* initialize its own COM apartment.
`pythoncom.CoInitialize()`
Without this step, `multiprocessing` worker threads will fail to communicate with the OS when attempting to launch Excel.

### 3. Isolated Excel Instances (`win32com.client.DispatchEx`)
Unlike simple automation tasks that attach to an existing, open Excel window, batch processing requires isolation.
`excel = win32com.client.DispatchEx('Excel.Application')`
`DispatchEx` forces Windows to spawn a completely separate `EXCEL.EXE` process in memory for *each* parallel worker. If one file crashes or contains a corrupted macro, it does not bring down the other files currently processing in parallel.

### 4. Strict Background Execution
To prevent hidden Excel processes from hanging (waiting for user input on dialog boxes), the script ruthlessly disables the UI and standard warnings:
```python
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False
excel.Interactive = False
```

### 5. Advanced Long-Path Bypass (`net use`)
Windows Excel COM Automation has a hardcoded internal limit of 218 characters for file paths. To bypass this and support deeply nested network folders, this script automatically maps your `INPUT_DIR` to a temporary `M:\` drive using `net use` at startup, processes all files safely through this short path, and then unmaps the drive when finished.

To ensure stability across crashes, the script implements proactive networking cleanup:
- Before mapping, it forcibly executes `net use M: /delete /y` to clear out any abandoned drives from previous runs.
- After processing, it not only deletes the `M:` drive but also explicitly runs `net use "Input Path" /delete /y` to sever lingering background UNC connections.

### 6. PDF Export Generation
The script uses the native Excel function `ExportAsFixedFormat` with `Type=0` (`xlTypePDF`). The `0` argument is critical, as it bypasses page-by-page printing and signals Excel to dump the *entire workbook* (all visible sheets) sequentially into a single PDF output file.

### 7. Pandas Export & Summary Report
While processing, the script gathers detailed telemetry (start time, end time, elapsed seconds, and error reasons). Once all files are processed, it uses `pandas` to compile this data into a structured DataFrame. This data is then exported as a `conversion_summary_report.xlsx` file inside the root of your `INPUT_DIR` for easy auditing.

### 8. Memory Cleanup
A crucial final step is `excel.Quit()` combined with `pythoncom.CoUninitialize()`. This ensures that once the PDF is generated, the detached `EXCEL.EXE` background process is immediately killed, freeing up RAM for the next file in the multiprocessing queue.
