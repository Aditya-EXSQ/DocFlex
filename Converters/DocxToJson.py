"""
Converters/DocxToJson.py

Class-based, editable + lossless DOCX -> JSON exporter.

Usage:
    from Converters.DocxToJson import DocxToJson
    d2j = DocxToJson("input.docx", "Output/SomeFolder")
    d2j.extract_docx()                 # optional (keeps a temp extraction)
    json_path = d2j.convert_to_json()  # returns path to created JSON
    d2j.cleanup()
"""
import zipfile
import base64
import re
import json
import os
import tempfile
import shutil

_T_BYTES_RE = re.compile(
    br'(<(?P<tag>(?:[A-Za-z0-9_]+:)?t)(?P<attrs>[^>]*)>)(?P<text>.*?)(</(?P=tag)>)',
    flags=re.DOTALL
)


class DocxToJson:
    def __init__(self, docx_path: str, output_dir: str = "Output"):
        self.docx_path = os.path.abspath(docx_path)
        self.output_dir = os.path.abspath(output_dir)
        os.makedirs(self.output_dir, exist_ok=True)

        self._tmpdir = None
        self._zip_infolist = None
        self._zip_entries = {}  # filename -> bytes

    def extract_docx(self):
        """Extracts the docx contents into a temp dir and caches bytes in memory."""
        if not os.path.exists(self.docx_path):
            raise FileNotFoundError(self.docx_path)
        self._tmpdir = tempfile.mkdtemp(prefix="docx2json_")
        with zipfile.ZipFile(self.docx_path, "r") as z:
            self._zip_infolist = z.infolist()
            for info in self._zip_infolist:
                fn = info.filename
                raw = z.read(fn)
                # write to disk (helpful for debugging / optional)
                out_path = os.path.join(self._tmpdir, fn)
                os.makedirs(os.path.dirname(out_path), exist_ok=True)
                with open(out_path, "wb") as f:
                    f.write(raw)
                # cache bytes and zipinfo minimal metadata
                self._zip_entries[fn] = {"bytes": raw, "zipinfo": {
                    "compress_type": info.compress_type,
                    "date_time": list(info.date_time),
                    "external_attr": info.external_attr,
                    "create_system": info.create_system
                }}

        return self._tmpdir

    def _make_editable_for_xml(self, raw_bytes: bytes):
        """
        Build 'runs' and 'placeholders' data structures for a single XML bytes blob.
        Returns dict: {"runs": [...], "placeholders": [...]}
        """
        runs = []
        for i, m in enumerate(_T_BYTES_RE.finditer(raw_bytes)):
            text_bytes = m.group('text') or b''
            try:
                text_decoded = text_bytes.decode('utf-8')
            except Exception:
                text_decoded = text_bytes.decode('utf-8', errors='replace')
            start, end = m.span()
            runs.append({
                "run_id": i,
                "text": text_decoded,
                "start": start,
                "end": end,
                "snippet_b64": base64.b64encode(raw_bytes[start:end]).decode('ascii')
            })

        concatenated = ''.join(r['text'] for r in runs)
        placeholders = []
        # detect placeholders like {{name}} or {name}
        for pm in re.finditer(r'\{\{.*?\}\}|\{[^{}]+\}', concatenated):
            ph_text = pm.group(0)
            ph_start = pm.start()
            ph_end = pm.end()
            mapping = []
            pos = 0
            for r in runs:
                rlen = len(r['text'])
                if pos + rlen <= ph_start:
                    pos += rlen
                    continue
                if pos >= ph_end:
                    break
                s_in_run = max(0, ph_start - pos)
                e_in_run = min(rlen, ph_end - pos)
                mapping.append({
                    "run_id": r['run_id'],
                    "start_in_run": s_in_run,
                    "end_in_run": e_in_run
                })
                pos += rlen
            placeholders.append({
                "placeholder": ph_text,
                "concat_offset": ph_start,
                "runs": mapping
            })

        return {"runs": runs, "placeholders": placeholders}

    def convert_to_json(self, json_name: str = "Document.json"):
        """
        Converts the cached docx into an editable JSON.
        If extract_docx() wasn't called earlier, this will read the zip directly.
        Returns full path to the JSON file.
        """
        if not self._zip_entries:
            # read zip entries directly
            with zipfile.ZipFile(self.docx_path, "r") as z:
                for info in z.infolist():
                    fn = info.filename
                    raw = z.read(fn)
                    self._zip_entries[fn] = {"bytes": raw, "zipinfo": {
                        "compress_type": info.compress_type,
                        "date_time": list(info.date_time),
                        "external_attr": info.external_attr,
                        "create_system": info.create_system
                    }}

        result = {}
        # Maintain the order of entries as in original zip (if available)
        entry_order = list(self._zip_entries.keys())
        for name in entry_order:
            raw = self._zip_entries[name]["bytes"]
            zi = self._zip_entries[name]["zipinfo"]
            entry = {
                "content_b64": base64.b64encode(raw).decode('ascii'),
                "zipinfo": zi
            }
            if name.lower().endswith('.xml'):
                try:
                    entry["_editable"] = self._make_editable_for_xml(raw)
                except Exception:
                    # parsing failed; skip editable block
                    pass
            result[name] = entry

        json_path = os.path.join(self.output_dir, json_name)
        os.makedirs(os.path.dirname(json_path) or '.', exist_ok=True)
        with open(json_path, 'w', encoding='utf-8') as jf:
            json.dump(result, jf, indent=2, ensure_ascii=False)

        return json_path

    def cleanup(self):
        """Remove temporary extraction directory and clear caches."""
        if self._tmpdir and os.path.exists(self._tmpdir):
            shutil.rmtree(self._tmpdir, ignore_errors=True)
        self._tmpdir = None
        self._zip_entries = {}
        self._zip_infolist = None
