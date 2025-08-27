"""
Converters/JsonToDocx.py (ElementTree-based)

Parses XML entries and updates <*:t> elements based on JSON _edits mapping.
This guarantees well-formed XML after edits, avoiding corruption.
"""

import json, base64, zipfile, os, tempfile, shutil, datetime
import xml.etree.ElementTree as ET


class JsonToDocx:
    def __init__(self, json_path: str, output_dir: str = "Output"):
        self.json_path = os.path.abspath(json_path)
        self.output_dir = os.path.abspath(output_dir)
        os.makedirs(self.output_dir, exist_ok=True)
        self._tmpdir = None
        self._data = None

    def json_to_xml(self):
        with open(self.json_path, "r", encoding="utf-8") as f:
            self._data = json.load(f)
        self._tmpdir = tempfile.mkdtemp(prefix="json2docx_")
        edits = self._data.get("_edits", {})

        for name, entry in self._data.items():
            if not isinstance(entry, dict) or "content_b64" not in entry:
                continue
            raw = base64.b64decode(entry["content_b64"])
            target_path = os.path.join(self._tmpdir, name)
            os.makedirs(os.path.dirname(target_path), exist_ok=True)

            if name.lower().endswith(".xml") and "_editable" in entry:
                file_edits = edits.get(name, {})
                if file_edits:
                    try:
                        root = ET.fromstring(raw.decode("utf-8", errors="replace"))
                    except Exception:
                        with open(target_path, "wb") as wf:
                            wf.write(raw)
                        continue

                    # collect <t> nodes
                    t_elems = []
                    for el in root.iter():
                        if isinstance(el.tag, str) and (el.tag.endswith("}t") or el.tag == "t"):
                            t_elems.append(el)

                    runs = entry["_editable"].get("runs", [])
                    runid_to_elem = {}
                    for r, el in zip(runs, t_elems):
                        runid_to_elem[r["run_id"]] = el

                    placeholders = entry["_editable"].get("placeholders", [])
                    for ph in placeholders:
                        ph_text = ph.get("placeholder")
                        if ph_text not in file_edits:
                            continue
                        new_val = file_edits[ph_text]
                        run_span = ph.get("runs", [])
                        if not run_span:
                            continue
                        first_run = run_span[0]["run_id"]
                        first_el = runid_to_elem.get(first_run)
                        if first_el is not None:
                            first_el.text = str(new_val)
                        for sub in run_span[1:]:
                            sub_el = runid_to_elem.get(sub["run_id"])
                            if sub_el is not None:
                                sub_el.text = ""

                    new_raw = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    with open(target_path, "wb") as wf:
                        wf.write(new_raw)
                    continue

            # default: just write raw
            with open(target_path, "wb") as wf:
                wf.write(raw)

        return self._tmpdir

    def xml_to_docx(self, out_name: str = None):
        if self._data is None:
            with open(self.json_path, "r", encoding="utf-8") as f:
                self._data = json.load(f)
        if out_name is None:
            base = os.path.splitext(os.path.basename(self.json_path))[0]
            out_name = f"{base}_reconstructed.docx"
        out_path = os.path.join(self.output_dir, out_name)
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, "w") as outz:
            for name, entry in self._data.items():
                if not isinstance(entry, dict) or "content_b64" not in entry:
                    continue
                file_path = None
                if self._tmpdir:
                    candidate = os.path.join(self._tmpdir, name)
                    if os.path.exists(candidate):
                        file_path = candidate
                if file_path:
                    with open(file_path, "rb") as rf:
                        raw = rf.read()
                else:
                    raw = base64.b64decode(entry["content_b64"])
                zi = zipfile.ZipInfo(filename=name)
                zinfo = entry.get("zipinfo", {})
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
