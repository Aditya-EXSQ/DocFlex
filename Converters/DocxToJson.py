"""
Converters/DocxToJson.py

- Stores each ZIP entry as base64 (lossless).
- For XML parts, extracts all <w:t> in document order into _editable.runs.
- Builds _editable.all_text (concatenated visible text).
- Detects placeholders that MAY span multiple runs, e.g. "{MemFirstName}" split
  as "{", "MemFirstName", "}" across 3 runs; maps them back to per-run slices.

This keeps the physical runs intact for perfect round-trip, but exposes
logical placeholders as single editable items in _editable.placeholders.
"""

import os, zipfile, base64, json, tempfile, shutil, re
import xml.etree.ElementTree as ET


PLACEHOLDER_RE = re.compile(r"\{\{.*?\}\}|\{[^{}\n\r]+\}")

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

    def _xml_get_runs(self, raw: bytes):
        """
        Return (runs, all_text) where:
          runs = [ {run_id, text} ... ] in document order of <*:t> nodes
          all_text = ''.join(run.text for run in runs)
        """
        try:
            text = raw.decode("utf-8", errors="replace")
            root = ET.fromstring(text)
        except Exception:
            return [], ""

        runs = []
        run_id = 0
        for el in root.iter():
            tag = el.tag
            if isinstance(tag, str) and (tag.endswith("}t") or tag == "t"):
                t_text = el.text if el.text else ""
                runs.append({"run_id": run_id, "text": t_text})
                run_id += 1

        all_text = "".join(r["text"] for r in runs)
        return runs, all_text

    def _map_placeholders(self, runs, all_text):
        """
        Find placeholders in concatenated all_text and map them back to per-run slices.
        Returns: list of { placeholder, runs: [ {run_id, start_in_run, end_in_run} ... ] }
        """
        # Precompute run positions in all_text
        positions = []
        pos = 0
        for r in runs:
            start = pos
            end = pos + len(r["text"])
            positions.append((start, end))  # absolute [start, end)
            pos = end

        placeholders = []
        for m in PLACEHOLDER_RE.finditer(all_text):
            p_start, p_end = m.span()
            p_text = m.group(0)
            span = []

            # Map this [p_start, p_end) back into one or more runs
            for idx, (rs, re_) in enumerate(positions):
                if re_ <= p_start:
                    continue
                if rs >= p_end:
                    break
                # overlap with run idx
                s_in_run = max(0, p_start - rs)
                e_in_run = min(re_ - rs, p_end - rs)
                if s_in_run < e_in_run:
                    span.append({
                        "run_id": runs[idx]["run_id"],
                        "start_in_run": int(s_in_run),
                        "end_in_run": int(e_in_run),
                    })

            if span:
                placeholders.append({
                    "placeholder": p_text,
                    "runs": span
                })

        return placeholders

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

                # For XML files, extract editable data
                if zi.filename.lower().endswith(".xml"):
                    runs, all_text = self._xml_get_runs(raw)
                    if runs:
                        placeholders = self._map_placeholders(runs, all_text)
                        entry["_editable"] = {
                            "runs": runs,                 # physical runs (unchanged)
                            "all_text": all_text,         # concatenated visible text
                            "placeholders": placeholders  # logical placeholders across runs
                        }

                result[zi.filename] = entry

        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        return out_path

    def cleanup(self):
        if self._tmpdir and os.path.exists(self._tmpdir):
            shutil.rmtree(self._tmpdir, ignore_errors=True)
        self._tmpdir = None
