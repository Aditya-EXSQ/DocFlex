import subprocess
import os
import sys
import shutil
from dotenv import load_dotenv

load_dotenv()

SOFFICE_PATH = os.getenv("SOFFICE_PATH")

def _unique_cleaned_path(input_path, output_dir=None):
    base = os.path.splitext(os.path.basename(input_path))[0]
    if output_dir is None:
        output_dir = os.path.dirname(input_path) or "."
    cand = os.path.join(output_dir, f"{base} - cleaned.docx")
    if not os.path.exists(cand):
        return cand
    i = 1
    while True:
        cand_i = os.path.join(output_dir, f"{base} - cleaned ({i}).docx")
        if not os.path.exists(cand_i):
            return cand_i
        i += 1

def flatten_docx_copy(input_path, output_dir=None):
    """
    1) Make a copy: <name> - cleaned.docx
    2) Run DOCX -> DOC -> DOCX on the copy only
    3) Return path to the cleaned copy
    """
    if output_dir is None:
        output_dir = os.path.dirname(input_path) or "."

    # 1) make the copy first
    cleaned_path = _unique_cleaned_path(input_path, output_dir)
    shutil.copy2(input_path, cleaned_path)

    base_clean = os.path.splitext(os.path.basename(cleaned_path))[0]
    intermediate_doc = os.path.join(output_dir, base_clean + ".doc")

    try:
        # 2a) DOCX (copy) -> DOC
        subprocess.run(
            [SOFFICE_PATH, "--headless", "--convert-to", "doc", cleaned_path, "--outdir", output_dir],
            check=True
        )

        # 2b) DOC -> DOCX (overwrites the copied docx with flattened content)
        subprocess.run(
            [SOFFICE_PATH, "--headless", "--convert-to", "docx", intermediate_doc, "--outdir", output_dir],
            check=True
        )
    except FileNotFoundError:
        raise RuntimeError(f"LibreOffice not found at: {SOFFICE_PATH}")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"LibreOffice conversion failed: {e}") from e
    finally:
        # cleanup intermediate .doc if present
        if os.path.exists(intermediate_doc):
            try:
                os.remove(intermediate_doc)
            except OSError:
                pass

    # The final flattened file is the same cleaned_path name
    if not os.path.exists(cleaned_path):
        # some LibreOffice builds create a new file; ensure we point to it
        generated = os.path.join(output_dir, base_clean + ".docx")
        if os.path.exists(generated):
            # replace our copied placeholder with the generated file
            try:
                if os.path.exists(cleaned_path):
                    os.remove(cleaned_path)
            except OSError:
                pass
            os.replace(generated, cleaned_path)

    return cleaned_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python CleanDOCX.py <input.docx> [output_dir]")
        sys.exit(1)

    inp = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else None
    result = flatten_docx_copy(inp, outdir)
    print(f"Cleaned copy created: {result}")
