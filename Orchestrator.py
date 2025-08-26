import os
from dotenv import load_dotenv
from Converters.DocxToJson import DocxToJson
from Converters.JsonToDocx import JsonToDocx
from JsonEditorUI import JsonEditorUI

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(BASE_DIR, ".env")
load_dotenv(dotenv_path=ENV_PATH)

def convert_docx_roundtrip(docx_path, output_dir="Output"):
    print("Step 1: DOCX → XML → JSON")
    d2j = DocxToJson(docx_path, output_dir)
    d2j.extract_docx()
    json_path = d2j.convert_to_json()

    print("Step 2: JSON → XML → DOCX")
    j2d = JsonToDocx(json_path, output_dir)
    j2d.json_to_xml()
    reconstructed_docx = j2d.xml_to_docx()

    d2j.cleanup()
    j2d.cleanup()

    print(f"Roundtrip complete! Output: {reconstructed_docx}")
    return reconstructed_docx

def open_json_editor(json_path):
    """Launch Tkinter editor for JSON dataBinding fields"""
    editor = JsonEditorUI(json_path)
    editor.run()

# Example usage

# Roundtrip (DOCX -> JSON -> DOCX)
convert_docx_roundtrip("Data/DOCX Files/Master Approval Letter.docx", "Output/DOCX Files/Master Approval Letter")

# TKINTER UI
# json_path = "Output/DOCX Files/Master Approval Letter/Document.json"
# open_json_editor(json_path)