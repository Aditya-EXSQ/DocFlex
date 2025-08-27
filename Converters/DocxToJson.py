"""
Converters/DocxToJson.py

Converts a DOCX into a JSON representation:
 - Each ZIP entry stored as base64
 - XML entries include _editable with run/placeholder metadata
 - Now also extracts visible <w:t> text into runs[].text
 - Adds all_text field for easier searching
"""

import os, zipfile, base64, json, tempfile, shutil, xml.etree.ElementTree as ET


class DocxToJson:
    def __init__(self, docx_path: str, output_dir: str = "Output"):
        self.docx_path = os.path.abspath(docx_path)
        self.output_dir = os.path.abspath(output_dir)
        os.makedirs(self.output_dir, exist_ok=True)
        self._tmpdir = None

    def extract_docx(self):
        self._tmpdir = tempfile.mkdtemp(prefix="docx2json_")
        with zipfile.ZipFile(self.docx_path, "r") as z:
            z.extractall(self._tmpdir)
        return self._tmpdir

    def convert_to_json(self, out_name: str = "Document.json"):
        out_path = os.path.join(self.output_dir, out_name)
        result = {}

        with zipfile.ZipFile(self.docx_path, "r") as z:
            for zi in z.infolist():
                raw = z.read(zi.filename)
                entry = {
                    "content_b64": base64.b64encode(raw).decode("utf-8"),
                    "zipinfo": {
                        "filename": zi.filename,
                        "date_time": list(zi.date_time),
                        "compress_type": zi.compress_type,
                        "external_attr": zi.external_attr,
                    },
                }

                # Extract editable runs if XML file
                if zi.filename.lower().endswith(".xml"):
                    try:
                        text = raw.decode("utf-8")
                        root = ET.fromstring(text)

                        runs, placeholders = [], []
                        run_id = 0

                        all_text = []  # collect all <w:t> content

                        for el in root.iter():
                            tag = el.tag
                            if isinstance(tag, str) and (tag.endswith("}t") or tag == "t"):
                                t_text = el.text if el.text else ""
                                runs.append(
                                    {
                                        "run_id": run_id,
                                        "text": t_text,
                                    }
                                )
                                run_id += 1
                                all_text.append(t_text)

                                # detect placeholders like {{NAME}} or MemFirstName
                                if "{" in t_text or "Mem" in t_text:
                                    placeholders.append(
                                        {
                                            "placeholder": t_text,
                                            "runs": [{"run_id": run_id - 1}],
                                        }
                                    )

                        entry["_editable"] = {
                            "runs": runs,
                            "all_text": "".join(all_text),
                            "placeholders": placeholders,
                        }

                    except Exception as e:
                        # skip if not XML or parse fails
                        pass

                result[zi.filename] = entry

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        return out_path

    def cleanup(self):
        if self._tmpdir and os.path.exists(self._tmpdir):
            shutil.rmtree(self._tmpdir, ignore_errors=True)
        self._tmpdir = None
