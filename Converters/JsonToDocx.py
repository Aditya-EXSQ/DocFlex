"""
Converters/JsonToDocx.py (ElementTree-based with span-accurate edits + auto-detect)

- Applies edits using _editable.placeholders[].runs spans:
    Each span item gives (run_id, start_in_run, end_in_run) so we can splice
    inside the <w:t> text accurately even when a placeholder covers PART of a run.

- Two ways to edit:
    A) Explicit:   _edits["word/document.xml"]["{MemFirstName}"] = "Aditya"
    B) Direct JSON edits: change runs[].text; auto-detect compares the current
       text over each placeholder span vs the original placeholder string.

- Multiple occurrences of the same placeholder text will all be replaced
  when using the _edits mapping keyed by the placeholder string.

NOTE: We keep physical runs intact for lossless round-trip; only <w:t>.text
values are updated.
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

    def _collect_t_elems(self, raw: bytes):
        root = ET.fromstring(raw.decode("utf-8", errors="replace"))
        t_elems = []
        for el in root.iter():
            tag = el.tag
            if isinstance(tag, str) and (tag.endswith("}t") or tag == "t"):
                t_elems.append(el)
        return root, t_elems

    def _auto_detect_edits_from_runs(self, entry):
        """
        Inspect runs[].text vs placeholder span text; if different, emit an edit
        mapping from original placeholder string -> new span text.
        """
        edits = {}
        editable = entry.get("_editable") or {}
        runs = editable.get("runs", [])
        placeholders = editable.get("placeholders", [])
        if not runs or not placeholders:
            return edits

        # build lookup by run_id
        run_text_by_id = {r["run_id"]: (r.get("text") or "") for r in runs}

        for ph in placeholders:
            orig_ph = ph.get("placeholder", "")
            span = ph.get("runs", [])
            # reconstruct current text across the span from runs (respects partial-run slices)
            current_parts = []
            for seg in span:
                rid = seg["run_id"]
                s = int(seg.get("start_in_run", 0))
                e = int(seg.get("end_in_run", 0))
                t = run_text_by_id.get(rid, "")
                current_parts.append(t[s:e])
            current_text = "".join(current_parts)

            if current_text != orig_ph:
                edits[orig_ph] = current_text

        return edits

    def json_to_xml(self):
        with open(self.json_path, "r", encoding="utf-8") as f:
            self._data = json.load(f)
        self._tmpdir = tempfile.mkdtemp(prefix="json2docx_")

        # Start with explicit _edits; augment with auto-detected edits.
        explicit_edits = self._data.get("_edits", {})

        for name, entry in self._data.items():
            if not isinstance(entry, dict) or "content_b64" not in entry:
                continue

            raw = base64.b64decode(entry["content_b64"])
            target_path = os.path.join(self._tmpdir, name)
            os.makedirs(os.path.dirname(target_path), exist_ok=True)

            if name.lower().endswith(".xml") and "_editable" in entry:
                # Merge explicit edits with auto-detected ones (direct JSON run edits)
                file_edits = dict(explicit_edits.get(name, {}))
                auto = self._auto_detect_edits_from_runs(entry)
                # Auto-detected may overwrite explicit if same key; you can reverse if desired
                file_edits.update(auto)

                if file_edits:
                    try:
                        root, t_elems = self._collect_t_elems(raw)
                    except Exception:
                        # If parsing fails, write original bytes
                        with open(target_path, "wb") as wf:
                            wf.write(raw)
                        continue

                    # Map run_id -> element in doc order
                    runs = entry["_editable"].get("runs", [])
                    runid_to_elem = {}
                    for r, el in zip(runs, t_elems):
                        runid_to_elem[r["run_id"]] = el

                    # Apply edits for each placeholder occurrence using spans
                    placeholders = entry["_editable"].get("placeholders", [])
                    for ph in placeholders:
                        orig_ph = ph.get("placeholder")
                        if orig_ph not in file_edits:
                            continue
                        replacement = str(file_edits[orig_ph])
                        span = ph.get("runs", [])
                        if not span:
                            continue

                        # Gather elements and slice bounds aligned to the span order
                        elems = []
                        for seg in span:
                            rid = seg["run_id"]
                            s = int(seg.get("start_in_run", 0))
                            e = int(seg.get("end_in_run", 0))
                            el = runid_to_elem.get(rid)
                            if el is None:
                                elems = []
                                break
                            # Normalize None text to empty
                            el.text = el.text or ""
                            elems.append((el, s, e))

                        if not elems:
                            continue

                        if len(elems) == 1:
                            el, s, e = elems[0]
                            el.text = (el.text[:s]) + replacement + (el.text[e:])
                        else:
                            # First element: prefix + replacement
                            first_el, s0, e0 = elems[0]
                            first_el.text = (first_el.text[:s0]) + replacement

                            # Middle elements: remove covered span but preserve any non-covered text (rare)
                            for el, s, e in elems[1:-1]:
                                el.text = (el.text[:s]) + (el.text[e:])

                            # Last element: keep suffix after end
                            last_el, sl, el_ = elems[-1]
                            last_el.text = last_el.text[el_:]

                    # Serialize back
                    new_raw = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    with open(target_path, "wb") as wf:
                        wf.write(new_raw)
                    continue

            # Default: write original bytes
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
