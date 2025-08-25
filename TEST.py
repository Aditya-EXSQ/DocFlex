import subprocess
from docx_utils.flatten import opc_to_flat_opc
import os
from dotenv import load_dotenv

load_dotenv()

SOFFICE_PATH = os.getenv("SOFFICE_PATH")

# Step 1: DOCX → Flat OPC XML
opc_to_flat_opc("Data\DOCX Files\Master Approval Letter.docx", "flat_opc.xml")
print("Step 1: flat_opc.xml generated.")

# Step 2: Flat OPC XML → DOCX via LibreOffice CLI
# Make sure LibreOffice is installed and 'soffice' is in your PATH

cmd = [
    SOFFICE_PATH,
    "--headless",
    "--convert-to", "docx",
    "flat_opc.xml",
    "--outdir", "."
]

result = subprocess.run(cmd, capture_output=True, text=True)

if result.returncode != 0:
    print("LibreOffice conversion failed:\n", result.stderr)
else:
    output_path = os.path.splitext("flat_opc.xml")[0] + ".docx"
    print(f"Step 2: Conversion successful! Output saved to: {output_path}")
