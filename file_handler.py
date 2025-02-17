import os
import time
import shutil
import csv
from pathlib import Path
from pdf2docx import Converter
from win10toast import ToastNotifier

toaster = ToastNotifier()

# Import your function from your info_extractor module
from info_extractor import get_requested_info

file_tree = r"C:\Users\Usuario\Desktop\data_base\file_tree"
manifest = r"C:\Users\Usuario\Desktop\data_base\manifest.csv"
receiver_folder = r"C:\Users\Usuario\Desktop\receiver_folder"
error_folder = r"C:\Users\Usuario\Desktop\error_folder"

AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".ogg", ".flac"}
DOC_EXTENSIONS = {".doc", ".docx", ".pdf"}

# Fields we need from get_requested_info
REQUIRED_FIELDS = [
    "Patient Name",
    "Creation Date",
    "Transcription Date",
    "Transcriber",
    "Exam Type",
    "Doctor",
]

# ------------------- Helper Functions ------------------- #

def ensure_folder_exists(folder_path: str) -> None:
    """
    Create the folder if it doesn't exist.
    """
    if not os.path.isdir(folder_path):
        os.makedirs(folder_path, exist_ok=True)

def convert_pdf_to_docx(pdf_path: str, docx_path: str) -> None:
    """
    Convert a PDF file to DOCX using the pdf2docx library.
    """
    cv = Converter(pdf_path)
    cv.convert(docx_path)
    cv.close()

def ensure_manifest_exists(csv_path: str) -> None:
    """
    Create the manifest CSV with headers if it doesn’t exist yet
    or if it’s empty. The header columns:
      ID, patient_name, creation_date, transcription_date,
      transcriber, exam_type, doctor, folder_address
    """
    create_new_header = False

    if not os.path.isfile(csv_path):
        # File doesn't exist
        create_new_header = True
    else:
        # If file exists but is empty, also create new header
        if os.path.getsize(csv_path) == 0:
            create_new_header = True

    if create_new_header:
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([
                "ID",
                "patient_name",
                "creation_date",
                "transcription_date",
                "transcriber",
                "exam_type",
                "doctor",
                "folder_address"
            ])

def get_next_id(csv_path: str) -> int:
    """
    Read the CSV manifest and return the next available integer ID.
    If the file doesn't exist or is empty, start at 1.
    """
    if not os.path.isfile(csv_path):
        return 1
    if os.path.getsize(csv_path) == 0:
        return 1

    with open(csv_path, "r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        header = next(reader, None)  # Skip the header row
        if not header:
            return 1
        
        max_id = 0
        for row in reader:
            if not row:
                continue
            try:
                row_id = int(row[0])  # First column is ID
                if row_id > max_id:
                    max_id = row_id
            except ValueError:
                pass
        return max_id + 1

def append_to_manifest(csv_path: str, new_id: int, info_dict: dict, folder_address: str) -> None:
    """
    Append a new row to the CSV manifest, storing:
      ID, patient_name, creation_date, transcription_date,
      transcriber, exam_type, doctor, folder_address
    """
    with open(csv_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            new_id,
            info_dict["Patient Name"],
            info_dict["Creation Date"],
            info_dict["Transcription Date"],
            info_dict["Transcriber"],
            info_dict["Exam Type"],
            info_dict["Doctor"],
            folder_address
        ])

def create_target_folder(base_folder: str, info_dict: dict) -> str:
    """
    Build and create the folder tree:
      Transcriber / Doctor / Exam Type / YYYY / MM / DD
    Return the absolute path to that folder.
    """
    transcriber = info_dict["Transcriber"]
    doctor = info_dict["Doctor"]
    exam_type = info_dict["Exam Type"]
    
    # Expect "DD/MM/YYYY"
    date_parts = info_dict["Transcription Date"].split("/")
    if len(date_parts) != 3:
        # If no valid date, placeholders
        year, month, day = "0000", "00", "00"
    else:
        day, month, year = date_parts
    
    folder_path = os.path.join(
        base_folder,
        transcriber,
        doctor,
        exam_type,
        year,
        month,
        day
    )
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def move_files_to_error(files_to_move: list, error_dest: str, error_message: str = "") -> None:
    """
    Move all listed files to the error folder and display a Windows 
    toast notification describing what went wrong (if supported).
    """
    for file_path in files_to_move:
        file_name = os.path.basename(file_path)
        shutil.move(file_path, os.path.join(error_dest, file_name))

    # Once done, show a toast notification with your error message
    if error_message:
        toaster.show_toast(
            "File Processing Error",
            error_message,
            duration=10,  # number of seconds the notification stays on screen
            threaded=True
        )

def process_two_files(doc_path: str, audio_path: str) -> None:
    """
    Process the doc file and audio file:
      1) Convert PDF -> DOCX if needed
      2) Extract info
      3) Validate fields
      4) Get next ID
      5) Rename files (including ID)
      6) Create folder structure
      7) Move files
      8) Append to CSV manifest
    """
    from unicodedata import normalize, category
    import re

    def safe_filename_part(s: str) -> str:
        # Remove diacritics, then keep only alnum + underscores
        s_clean = "".join(
            c for c in normalize("NFKD", s) if category(c) != "Mn"
        )
        s_clean = s_clean.replace(" ", "_").lower()
        return re.sub(r"[^a-z0-9_]+", "", s_clean)
    
    doc_ext = Path(doc_path).suffix.lower()
    docx_temp_path = doc_path

    # Convert PDF -> DOCX if needed
    if doc_ext == ".pdf":
        docx_temp_path = doc_path.replace(".pdf", ".docx")
        convert_pdf_to_docx(doc_path, docx_temp_path)
        # Remove original PDF if conversion succeeded
        if os.path.isfile(docx_temp_path):
            os.remove(doc_path)

    # Extract info
    info_dict = get_requested_info(docx_temp_path)

    # Validate required fields
    for field in REQUIRED_FIELDS:
        if not info_dict.get(field, "").strip():
            raise ValueError(f"Missing required field: {field}")

    # Get next ID for naming
    new_id = get_next_id(manifest)

    # Create the final folder structure
    target_folder = create_target_folder(file_tree, info_dict)

    # Build new filenames: "patient_examtype_ID.ext"
    patient_clean = safe_filename_part(info_dict["Patient Name"])
    exam_clean = safe_filename_part(info_dict["Exam Type"])

    doc_final_name = f"{patient_clean}_{exam_clean}_{new_id}{Path(docx_temp_path).suffix}"
    audio_ext = Path(audio_path).suffix.lower()
    audio_final_name = f"{patient_clean}_{exam_clean}_{new_id}{audio_ext}"

    doc_final_path = os.path.join(target_folder, doc_final_name)
    audio_final_path = os.path.join(target_folder, audio_final_name)

    # Move them
    shutil.move(docx_temp_path, doc_final_path)
    shutil.move(audio_path, audio_final_path)

    # Append to manifest
    folder_address = target_folder
    append_to_manifest(manifest, new_id, info_dict, folder_address)

# ------------------- Main Loop ------------------- #

def main():
    print("Starting file handler. Monitoring folder:", receiver_folder)

    # Ensure all required directories exist
    ensure_folder_exists(file_tree)
    ensure_folder_exists(error_folder)
    ensure_folder_exists(receiver_folder)

    # Ensure the manifest is present with headers
    ensure_manifest_exists(manifest)

    while True:
        time.sleep(5)  # Wait between checks

        all_files = [os.path.join(receiver_folder, f) for f in os.listdir(receiver_folder)]
        if not all_files:
            continue

        # Expect exactly 2 files: 1 audio + 1 doc/pdf
        if len(all_files) != 2:
            move_files_to_error(all_files, error_folder)
            continue

        doc_file = None
        audio_file = None

        for fpath in all_files:
            ext = Path(fpath).suffix.lower()
            if ext in AUDIO_EXTENSIONS and audio_file is None:
                audio_file = fpath
            elif ext in DOC_EXTENSIONS and doc_file is None:
                doc_file = fpath
            else:
                doc_file = None
                audio_file = None
                break

        if not doc_file or not audio_file:
            # Not the expected combo
            move_files_to_error(all_files, error_folder)
            continue

        try:
            process_two_files(doc_file, audio_file)
        except Exception as e:
            print("Error processing files:", e)
            # Move them to error folder
            move_files_to_error(all_files, error_folder)

if __name__ == "__main__":
    main()
