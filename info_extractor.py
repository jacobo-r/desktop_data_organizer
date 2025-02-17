import os
import re
import docx
import unicodedata

# This script will extract the data needed from a word file in a specified folder

# ------------------ CONFIG / DICTIONARIES ------------------ #
TRANSCRIBERS = [ 
    "GALVIS MORALES JENIFFER",
    "OROZCO BARTOLO OSBALDO",
    "OPSINA ARANGO GLORIA NEIVER",
    "RESTREPO CORREA JHONATAN",
    "RODRIGUEZ SERNA PAOLA ANDREA",
    "TAFUR GONZALES SAMIR YANED",
    "UTIMA PINEDA MARIA AMPARO"
]

EXAM_TYPES = {
    "RADIOGRAFIA": ["RADIOGRAFIA", "RX"],
    "ECOGRAFIA": ["ECOGRAFIA", "ECO", "ECOS", "ECOGRAFIAS"],
    "TOMOGRAFIA": ["TOMOGRAFIA", "TAC", "TOMOGRAFIAS"],
    "ANGIOTOMOGRAFIA": ["ANGIOTOMOGRAFIA", "ANGIOTAC"],
    "RESONANCIA": ["RESONANCIA", "RM", "RMN", "RESONANCIAS"],
    "ANGIORESONANCIA": ["ANGIORESONANCIA"],
    "UROGRAFIA": ["UROGRAFIA"],
    "URETROCISTOGRAFIA": ["URETROCISTOGRAFIA"],
    "NEFROSTOMIA": ["NEFROSTOMIA"],
    "CAVOGRAFIA": ["CAVOGRAFIA"],
    "DRENAJE": ["DRENAJE"],
    "MAMOGRAFIA": ["MAMOGRAFIA", "MAMO"],
    "BIOPSIA": ["BIOPSIA"],
    "ANESTESIA": ["ANESTESIA"],
    "IMPLANTE": ["IMPLANTE"],
    "OCLUSION": ["OCLUSION"],
    "TORACENTESIS": ["TORACENTESIS"],
    "COLANGIORESONANCIA": ["COLANGIORESONANCIA"],
    "FLEBOGRAFIA": ["FLEBOGRAFIA"],
    "PARACENTESIS": ["PARACENTESIS"],
    "ELASTOGRAFIA": ["ELASTOGRAFIA"],
    "RETIRO": ["RETIRO"],
    "ANGIOPLASTIA": ["ANGIOPLASTIA"],
    "VENOGRAFIA": ["VENOGRAFIA"],
    "FARINGOGRAFIA": ["FARINGOGRAFIA"],
    "COLANGIOGRAFIA": ["COLANGIOGRAFIA"],
    "PANANGIOGRAFIA": ["PANANGIOGRAFIA"],
    "PIELOGRAFIA": ["PIELOGRAFIA"],
    "PERICARDIOCENTESIS": ["PERICARDIOCENTESIS"],
    "ARTRORESONANCIA": ["ARTRORESONANCIA"],
    "ARTERIOGRAFIA": ["ARTERIOGRAFIA"],
    "COLECISTOSTOMIA": ["COLECISTOSTOMIA"],
    "FLUOROSCOPIA": ["FLUOROSCOPIA", "FLURO"],
    "MARCAPASO": ["MARCAPASO"],
    "FISTULOGRAFIA": ["FISTULOGRAFIA"],
    "URETROGRAFIA": ["URETROGRAFIA"]
}

doctor_map = {
    "VICTOR HUGO RUIZ GRANADA":       ["RUIZ"],
    "JUAN CARLOS CORREA PUERTA":      ["CORREA"],
    "YESID CARDOZO VELEZ":            ["YESID", "CARDOZO"],
    "SANDRA LUCIA LOPEZ SIERRA":      ["SANDRA", "DANDRA"],
    "OSCAR ANDRES ALVAREZ GOMEZ":     ["OSCAR", "ALVAREZ"],
    "AGUSTO LEON ARIAS ZULUAGA":      ["ARIAS"],
    "JORGE AUGUSTO PULGARIN OSORIO":  ["PULGARIN"],
    "LYNDA IVETTE CARVAJAL ACOSTA":   ["LYNDA", "CARVAJAL"],
    "ALONSO GOMEZ GARCIA":            ["GOMEZ", "GARCIA", "ALONSO"],
    "JOSE FERNANDO VILLABONA GARCIA": ["VILLABONA"],
    "FRANKLIN LEONARDO HANNA QUESADA":["HANNA"],
    "LUIS ALBERTO ROJAS":             ["ROJAS"],
    "CESAR YEPES":                    ["CESAR", "YEPES"],
    "LOREANNYS LORETYS OSPINO ORTIZ": ["LOREANNYS", "OSPINO", "LOREANYS", "LOREANY"],
    "JUAN MANUEL TORO SANCHEZ":       ["TORO"],
    "CARLOS FELIPE HURTADO ARIAS":    ["HURTADO", "FELIPE", "HURTAO"]
}

# ------------------ HELPERS ------------------ #
def remove_accents(s: str) -> str:
    return ''.join(
        c for c in unicodedata.normalize('NFKD', s)
        if unicodedata.category(c) != 'Mn'
    )

def find_transcriber_any_token(transcripcion: str) -> str:
    """
    Return the *first* transcriber from TRANSCRIBERS whose *any* name token
    appears in 'transcripcion' (case-insensitive, accent-insensitive).
    """
    trans_norm = remove_accents(transcripcion).upper()
    for full_name in TRANSCRIBERS:
        name_norm = remove_accents(full_name).upper()
        tokens = name_norm.split()
        for token in tokens:
            if token in trans_norm:
                return full_name
    return ""

def find_transcription_date(transcripcion: str) -> str:
    """
    Extract the first DD/MM/YYYY from 'transcripcion'. Return empty if not found.
    """
    match = re.search(r"\b(\d{2}/\d{2}/\d{4})\b", transcripcion)
    return match.group(1) if match else ""

def find_exam_type(procedimiento: str) -> str:
    """
    Return the exam type from EXAM_TYPES if any of its keywords 
    appear in the 'procedimiento' text. Otherwise, return empty.
    """
    proc_norm = remove_accents(procedimiento).upper()
    for exam_key, keywords in EXAM_TYPES.items():
        for kw in keywords:
            kw_norm = remove_accents(kw).upper()
            if kw_norm in proc_norm:
                return exam_key
    return ""

# Regex patterns for capturing fields in doc text
regex_patterns = {
    "paciente":        r"^paciente\s*:\s*(.*)$",
    "documento":       r"^documento\s*:\s*(.*)$",
    "entidad":         r"^entidad\s*:\s*(.*)$",
    "procedimiento":   r"^procedimiento\s*:\s*(.*)$",
    "fecha":           r"^fecha\s*:\s*(.*)$",
    "nro_remision":    r"^nro\s+remisi(?:o|ó)n\s*:\s*(.*)$",
    "transcripcion":   r"^transcripci(?:o|ó)n\s*:\s*(.*)$"
}

def extract_field(paragraph_text: str, fields: dict) -> dict:
    normalized = remove_accents(paragraph_text).lower().strip()
    for field, pattern in regex_patterns.items():
        match = re.match(pattern, normalized)
        if match:
            fields[field] = match.group(1).strip()
    return fields

def identify_doctor(paragraph_text: str) -> str:
    normalized = remove_accents(paragraph_text).lower()
    for doctor_name, keywords in doctor_map.items():
        for kw in keywords:
            kw_norm = remove_accents(kw).lower()
            if kw_norm in normalized:
                return doctor_name
    return ""

def parse_docx_file(file_path: str) -> dict:
    doc = docx.Document(file_path)

    # Initialize the fields
    fields = {
        "paciente": "",
        "documento": "",
        "entidad": "",
        "procedimiento": "",
        "fecha": "",
        "nro_remision": "",
        "transcripcion": "",
        "content_after_bars": "",
        "doctor": ""
    }

    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]

    # 1) Extract top fields
    last_field_index = -1
    for i, paragraph in enumerate(paragraphs):
        before = fields.copy()
        fields = extract_field(paragraph, fields)
        if fields != before:
            last_field_index = i

    # 2) Collect main body after last recognized field
    body_lines = []
    for j in range(last_field_index + 1, len(paragraphs)):
        p_text = paragraphs[j].strip()
        # Stop if we see signature indicators
        if re.search(r'^(atte|atentamente|dra\.|dr\.)', remove_accents(p_text).lower()):
            break
        body_lines.append(p_text)

    fields["content_after_bars"] = "\n".join(body_lines).strip()

    # 3) Identify doctor from bottom up
    for paragraph in reversed(paragraphs):
        doctor_found = identify_doctor(paragraph)
        if doctor_found:
            fields["doctor"] = doctor_found
            break

    return fields

def print_requested_fields(info: dict) -> None:
    """
    Prints the 6 requested items:
      1. Patient Name
      2. Creation Date (fecha)
      3. Transcription Date (from transcripcion)
      4. Transcriber (by single-token match)
      5. Exam Type (from procedimiento)
      6. Doctor
    """
    paciente        = info.get("paciente", "").strip()
    fecha_creacion  = info.get("fecha", "").strip()
    transcripcion   = info.get("transcripcion", "").strip()
    procedimiento   = info.get("procedimiento", "").strip()
    doctor          = info.get("doctor", "").strip()

    patient_name        = paciente
    creation_date       = fecha_creacion
    transcription_date  = find_transcription_date(transcripcion)
    transcriber         = find_transcriber_any_token(transcripcion)
    exam_type           = find_exam_type(procedimiento)

    print(f"Patient Name: {patient_name}")
    print(f"Creation Date: {creation_date}")
    print(f"Transcription Date: {transcription_date}")
    print(f"Transcriber: {transcriber}")
    print(f"Exam Type: {exam_type}")
    print(f"Doctor: {doctor}")
    print()  # blank line

def get_requested_info(file_path: str) -> dict:
    """
    Parse the specified Word document and return a dictionary with:
      - 'Patient Name'
      - 'Creation Date'
      - 'Transcription Date'
      - 'Transcriber'
      - 'Exam Type'
      - 'Doctor'

    This function relies on:
      1) parse_docx_file(file_path) -> dict   (already defined in your script)
      2) find_transcription_date(transcripcion: str) -> str
      3) find_transcriber_any_token(transcripcion: str) -> str
      4) find_exam_type(procedimiento: str) -> str

    which are also defined in your script.
    """
    # Parse the document fields (you already have parse_docx_file in your script)
    info = parse_docx_file(file_path)

    # Extract the fields we need
    patient_name = info.get("paciente", "").strip()
    creation_date = info.get("fecha", "").strip()
    transcripcion_text = info.get("transcripcion", "").strip()
    procedimiento_text = info.get("procedimiento", "").strip()
    doctor = info.get("doctor", "").strip()

    # Use your existing helper functions:
    transcription_date = find_transcription_date(transcripcion_text)
    transcriber = find_transcriber_any_token(transcripcion_text)
    exam_type = find_exam_type(procedimiento_text)

    # Return them in a single dictionary
    return {
        "Patient Name": patient_name,
        "Creation Date": creation_date,
        "Transcription Date": transcription_date,
        "Transcriber": transcriber,
        "Exam Type": exam_type,
        "Doctor": doctor
    }


# ------------------ MAIN ------------------ #
if __name__ == "__main__":
    folder_path = r"C:\Users\Usuario\Desktop\test"  # Adjust path as needed

    for fname in os.listdir(folder_path):
        if fname.lower().endswith(".docx"):
            file_path = os.path.join(folder_path, fname)
            info = parse_docx_file(file_path)

            # Print doc name (optional)
            print(f"--- {fname} ---")
            print_requested_fields(info)


    # we have extracted the info from file
    # we want to rename the file to patientname_study_date #handle dupes
    # we want to automatically put this into a CSV (for the startup)
    # we want to put this into an organized file structure (for the transcriptor)
    