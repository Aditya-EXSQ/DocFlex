"""
Converters/JsonToDocx.py

Class-based, lossless JSON -> DOCX re-constructor with editable placeholder application.

Usage:
    from Converters.JsonToDocx import JsonToDocx
    j2d = JsonToDocx("path/to/Document.json", "Output/SomeFolder")
    tmp_xml_dir = j2d.json_to_xml()     # writes modified XMLs to a temp dir (applies _edits if present)
    docx_path = j2d.xml_to_docx()      # creates a reconstructed docx and returns its path
    j2d.cleanup()
"""
import json
import base64
import zipfile
import os
import tempfile
import shutil
import datetime

def _xml_escape_bytes(s: str) -> bytes:
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    s = s.replace('"', '&quot;').replace("'", '&apos;')
    return s.encode('utf-8')


class JsonToDocx:
    def __init__(self, json_path: str, output_dir: str = "Output"):
        self.json_path = os.path.abspath(json_path)
        self.output_dir = os.path.abspath(output_dir)
        os.makedirs(self.output_dir, exist_ok=True)
        self._tmpdir = None
        self._data = None  # loaded JSON structure

    def json_to_xml(self):
        """
        Loads JSON and writes each entry into a temporary directory.
        If JSON contains a top-level "_edits" mapping or per-file edits passed via
        the JSON, this will apply them to XML files (splicing raw bytes).
        Returns path to created temp directory containing the XML and other files.
        """
        with open(self.json_path, 'r', encoding='utf-8') as f:
            self._data = json.load(f)

        # Create tmpdir for extracted/modified files
        self._tmpdir = tempfile.mkdtemp(prefix="json2docx_")
        # default edits mapping (top-level _edits) or empty
        edits = self._data.get("_edits", {})

        # For each entry, decode bytes and possibly apply edits for XML parts; write file to tmpdir
        for name, entry in self._data.items():
            if not isinstance(entry, dict) or 'content_b64' not in entry:
                continue
            raw = base64.b64decode(entry['content_b64'])
            target_path = os.path.join(self._tmpdir, name)
            os.makedirs(os.path.dirname(target_path), exist_ok=True)

            # If XML and editable, and there are edits for this entry, apply them losslessly
            if name.lower().endswith('.xml') and "_editable" in entry:
                file_edits = edits.get(name, {})
                if file_edits:
                    editable = entry["_editable"]
                    runs = editable.get("runs", [])
                    placeholders = editable.get("placeholders", [])
                    # Build run modifications map: run_id -> new inner bytes
                    run_mods = {}
                    for ph in placeholders:
                        ph_text = ph['placeholder']
                        if ph_text not in file_edits:
                            continue
                        new_value = file_edits[ph_text]
                        run_span = ph.get("runs", [])
                        if not run_span:
                            continue
                        first_run_id = run_span[0]['run_id']
                        new_bytes = _xml_escape_bytes(new_value)
                        run_mods[first_run_id] = new_bytes
                        for sub in run_span[1:]:
                            run_mods[sub['run_id']] = b''

                    if run_mods:
                        new_raw = raw
                        search_pos = 0
                        for r in runs:
                            rid = r['run_id']
                            if rid not in run_mods:
                                continue
                            orig_snip = base64.b64decode(r['snippet_b64'])
                            idx = new_raw.find(orig_snip, search_pos)
                            if idx == -1:
                                idx = new_raw.find(orig_snip)
                                if idx == -1:
                                    raise RuntimeError(f"Could not locate run snippet for run {rid} in {name}.")
                            len_orig = len(orig_snip)
                            orig_rel = orig_snip
                            # find boundaries inside orig_snip
                            open_end_rel = orig_rel.find(b'>')
                            close_start_rel = orig_rel.rfind(b'</')
                            if open_end_rel == -1 or close_start_rel == -1:
                                raise RuntimeError("Malformed run snippet while reconstructing.")
                            opening_tag = orig_rel[:open_end_rel+1]
                            closing_tag = orig_rel[close_start_rel:]
                            new_inner = run_mods[rid]
                            new_snip = opening_tag + new_inner + closing_tag
                            new_raw = new_raw[:idx] + new_snip + new_raw[idx + len_orig:]
                            search_pos = idx + len(new_snip)
                        raw = new_raw

            # Write raw (possibly modified) bytes to file
            with open(target_path, "wb") as wf:
                wf.write(raw)

        return self._tmpdir

    def xml_to_docx(self, out_name: str = None):
        """
        Packages files from the temp dir (created by json_to_xml) into a DOCX.
        Preserves zipinfo metadata from the JSON where available.
        Returns path to the reconstructed DOCX.
        """
        if self._data is None:
            # if json_to_xml wasn't called, just load data
            with open(self.json_path, 'r', encoding='utf-8') as f:
                self._data = json.load(f)

        if out_name is None:
            base = os.path.splitext(os.path.basename(self.json_path))[0]
            out_name = f"{base}_reconstructed.docx"
        out_path = os.path.join(self.output_dir, out_name)
        # remove if exists
        if os.path.exists(out_path):
            os.remove(out_path)

        # We'll write entries in the same order they appear in the JSON (which mirrors original zip order)
        with zipfile.ZipFile(out_path, 'w') as outz:
            for name, entry in self._data.items():
                if not isinstance(entry, dict) or 'content_b64' not in entry:
                    continue
                # Attempt to read the file from temp dir if we created one in json_to_xml
                file_path = None
                if self._tmpdir:
                    candidate = os.path.join(self._tmpdir, name)
                    if os.path.exists(candidate):
                        file_path = candidate
                # otherwise fall back to content_b64
                if file_path:
                    with open(file_path, 'rb') as rf:
                        raw = rf.read()
                else:
                    raw = base64.b64decode(entry['content_b64'])

                zi = zipfile.ZipInfo(filename=name)
                zinfo = entry.get('zipinfo', {})
                dt = zinfo.get("date_time")
                if dt and isinstance(dt, list) and len(dt) == 6:
                    zi.date_time = tuple(dt)
                else:
                    zi.date_time = datetime.datetime.now().timetuple()[:6]
                if "compress_type" in zinfo:
                    zi.compress_type = zinfo.get("compress_type")
                if "external_attr" in zinfo:
                    zi.external_attr = zinfo.get("external_attr", 0)
                outz.writestr(zi, raw)

        return out_path

    def cleanup(self):
        if self._tmpdir and os.path.exists(self._tmpdir):
            shutil.rmtree(self._tmpdir, ignore_errors=True)
        self._tmpdir = None
        # do not remove the JSON file — orchestrator owns that
