# Mini Project: Automated File Handler and Info Extractor

This mini-project watches a **receiver folder** for pairs of files (an audio file and a PDF/Word file) and automatically:

1. **Extracts** important information from the document (patient details, date, transcriber, etc.).  
2. **Organizes** the files (audio + doc) into a structured folder tree.  
3. **Records** the details in a CSV manifest for easy reference.  
4. **Handles** any errors (improper file types, missing info) by moving the files to a separate `error_folder`.

## Overview

- **`info_extractor.py`**  
  This module contains the function `get_requested_info(file_path: str) -> dict`, which:  
  - Takes the path to a Word (`.doc` / `.docx`) or converted PDF file.  
  - Parses the document to extract fields like **Patient Name**, **Creation Date**, **Transcription Date**, **Transcriber**, **Exam Type**, and **Doctor**.  
  - Returns a dictionary with those details.

- **`file_handler.py`**  
  This script uses `get_requested_info(...)` to process every pair of files (one doc, one audio) that lands in the **receiver folder**. Here’s what it does in detail:  
  1. **Monitors** the `receiver_folder` in a loop (e.g., every 5 seconds).  
  2. **Collects** pairs of files: exactly one audio file and one doc/PDF.  
  3. **Converts** PDFs to `.docx` (using `pdf2docx`) if necessary.  
  4. **Extracts** metadata by calling `get_requested_info(...)`.  
  5. **Builds** a structured directory path under `file_tree` based on the extracted info, e.g.:  
     ```
     file_tree/
       └─ [Transcriber]/
           └─ [Doctor]/
               └─ [Exam Type]/
                   └─ YYYY/
                       └─ MM/
                           └─ DD/
     ```  
  6. **Generates** a unique numeric ID for each pair and renames both the doc and audio files to something like `patient_examtype_ID.ext`.  
  7. **Moves** the newly named files into the correct folder path.  
  8. **Logs** an entry in the `manifest.csv`, storing:  
     - ID  
     - Patient Name  
     - Creation Date  
     - Transcription Date  
     - Transcriber  
     - Exam Type  
     - Doctor  
     - Destination folder path  
  9. **Cleans** the `receiver_folder` after successful processing.  

  If **any** step fails (e.g., missing fields, no doc file, multiple audio files, etc.), the script places the files into an **`error_folder`** and logs a Windows toast notification (optional) with the error reason.

## How to Use

1. **Install Dependencies**  
   - Python 3.x  
   - `pip install python-docx pdf2docx win10toast` (and any additional libraries if needed).  

2. **Set Up Folders**  
   - Create or confirm the existence of:  
     - `receiver_folder` (where the user drops pairs of files)  
     - `file_tree` (where organized files are saved)  
     - `error_folder` (where problematic files go)  
   - Ensure these paths are set correctly in `file_handler.py`.  

3. **Run `file_handler.py`**  
   - It will automatically create missing folders and a `manifest.csv` if not present.  
   - The script then goes into a **loop** watching for new pairs of files in `receiver_folder`.  

4. **Dropping Files**  
   - The user simply **places** one **audio** file (e.g., `.mp3`, `.wav`) **and** one **doc** (`.doc`, `.docx`) **or** `pdf` into the `receiver_folder`.  
   - Within a few seconds, the script picks them up and processes them:  
     - If all required info is extracted successfully, they get moved and logged.  
     - If anything is wrong (e.g., user drops 3 files, or 2 docs, or the doc is missing required fields), the script moves them to `error_folder` and notifies the user.

5. **Result**  
   - Properly processed files end up in an organized structure under `file_tree`.  
   - Each pair is uniquely identified by an auto-incrementing ID in the CSV, for easy reference.  
   - `error_folder` contains any files that could not be processed automatically.

## Notes and Troubleshooting

- **PDF to DOCX**  
  - Requires `pdf2docx` to be installed.  
  - If PDF conversion fails or is not desired, you can remove that part or handle `.pdf` in the error flow.

- **Document Format**  
  - The extraction logic in `info_extractor.py` depends on consistent formatting of fields like `Paciente:`, `Fecha:`, etc. If documents vary wildly, you may need to adjust the regex patterns.

- **Windows Toast Notifications**  
  - Uses `win10toast`. If you’re not on Windows or you don’t want pop-ups, you can remove that logic or replace it with simple console logs.

- **CSV Manifest**  
  - Located at `manifest.csv` in your chosen directory (`C:\Users\Usuario\Desktop\data_base\manifest.csv`, for instance).  
  - Column structure:  
    ```
    ID, patient_name, creation_date, transcription_date, transcriber, exam_type, doctor, folder_address
    ```

With this setup, you have an **automated pipeline** that watches a folder, **extracts** metadata from a doc/PDF, organizes files, and **logs** everything in a central CSV.
