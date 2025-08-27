"""
JsonEditorUI.py

Provides class JsonEditorUI(json_path=None) which opens a Tkinter UI to view/edit placeholders
stored in an editable JSON produced by DocxToJson. When the user saves, the JSON file is written
with a top-level "_edits" mapping that JsonToDocx will pick up automatically.

Usage (from orchestrator):
    editor = JsonEditorUI(json_path)
    editor.run()   # blocks until UI closed; JSON at json_path will contain _edits (if saved)

If json_path is None, user is prompted to open a JSON file.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json, os, importlib.util, traceback

# Try to locate JsonToDocx helper (so the UI can optionally reconstruct)
POSSIBLE_JSON2DOCX = [
    os.path.join(os.getcwd(), "JsonToDocx.py"),
    os.path.join(os.getcwd(), "Converters", "JsonToDocx.py"),
    "/mnt/data/JsonToDocx.py",
    "/mnt/data/Converters/JsonToDocx.py"
]

JSON2DOCX_PATH = None
for p in POSSIBLE_JSON2DOCX:
    if os.path.exists(p):
        JSON2DOCX_PATH = p
        break

def load_module_from_path(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

class JsonEditorUI:
    def __init__(self, json_path=None, auto_save_on_close=True):
        """
        json_path: path to editable JSON produced by DocxToJson
        auto_save_on_close: if True the editor writes _edits back into json_path when closed
        """
        self.json_path = json_path
        self.auto_save_on_close = auto_save_on_close

        self.data = None
        self.edits = {}   # staged edits: { "<xml_name>": { "<placeholder>": "<newval>" } }
        self.root = None
        self.docx_reconstructor = None
        if JSON2DOCX_PATH:
            try:
                self.docx_reconstructor = load_module_from_path(JSON2DOCX_PATH, "JsonToDocx_for_ui")
            except Exception:
                self.docx_reconstructor = None

    # -------------------------
    # Public
    # -------------------------
    def run(self):
        """
        Show the UI and block until closed.
        After close, if auto_save_on_close True, writes _edits into json_path.
        """
        self._build_ui()
        self.root.mainloop()
        # on close: write edits to JSON if requested
        if self.auto_save_on_close and self.json_path and self.data is not None:
            self.data["_edits"] = self.edits
            backup = self.json_path + ".bak"
            try:
                os.replace(self.json_path, backup)
            except Exception:
                # if replace fails, try to leave backup aside
                pass
            with open(self.json_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)

    # -------------------------
    # Internal UI implementation
    # -------------------------
    def _build_ui(self):
        self.root = tk.Tk()
        self.root.title("DOCX JSON Placeholder Editor")
        self.root.geometry("980x600")

        frm = ttk.Frame(self.root, padding=8)
        frm.pack(fill="both", expand=True)

        top = ttk.Frame(frm)
        top.pack(fill="x", pady=(0,8))
        ttk.Button(top, text="Open editable JSON", command=self._ask_open_json).pack(side="left", padx=4)
        ttk.Button(top, text="Save edits to JSON", command=self._save_json_with_edits).pack(side="left", padx=4)
        ttk.Button(top, text="Reconstruct DOCX (from current JSON)", command=self._reconstruct_docx).pack(side="left", padx=4)
        ttk.Button(top, text="Close", command=self._on_close).pack(side="right", padx=4)

        # content panes
        paned = ttk.Panedwindow(frm, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned, width=240)
        paned.add(left, weight=1)
        ttk.Label(left, text="XML parts").pack(anchor="w")
        self.list_xml = tk.Listbox(left, exportselection=False)
        self.list_xml.pack(fill="both", expand=True, padx=4, pady=4)
        self.list_xml.bind("<<ListboxSelect>>", self._on_xml_select)

        mid = ttk.Frame(paned, width=320)
        paned.add(mid, weight=2)
        ttk.Label(mid, text="Placeholders").pack(anchor="w")
        self.list_place = tk.Listbox(mid, selectmode="extended")
        self.list_place.pack(fill="both", expand=True, padx=4, pady=4)
        self.list_place.bind("<<ListboxSelect>>", self._on_placeholder_select)

        right = ttk.Frame(paned, width=360)
        paned.add(right, weight=2)
        ttk.Label(right, text="Edit value for selected placeholder(s)").pack(anchor="w")
        self.selected_label = ttk.Label(right, text="", wraplength=320, foreground="blue")
        self.selected_label.pack(anchor="w", padx=4, pady=(2,6))
        self.entry_new = tk.Text(right, height=6)
        self.entry_new.pack(fill="x", padx=4)
        btnf = ttk.Frame(right)
        btnf.pack(fill="x", pady=6)
        ttk.Button(btnf, text="Apply to selection", command=self._apply_to_selection).pack(side="left", padx=4)
        ttk.Button(btnf, text="Clear selection edits", command=self._clear_selection_edits).pack(side="left", padx=4)
        ttk.Label(right, text="Staged edits:").pack(anchor="w")
        self.staged = tk.Text(right, height=12, state="disabled")
        self.staged.pack(fill="both", expand=True, padx=4, pady=4)

        # status bar
        self.status = ttk.Label(self.root, text="Ready", relief="sunken", anchor="w")
        self.status.pack(fill="x", side="bottom")

        # If a json_path was provided, try to load it
        if self.json_path:
            try:
                self._load_json(self.json_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load JSON {self.json_path}:\n{e}")

        # Bind close event to save behavior
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _set_status(self, s):
        if self.status:
            self.status.config(text=s)
            self.root.update_idletasks()

    def _ask_open_json(self):
        path = filedialog.askopenfilename(title="Open editable JSON", filetypes=[("JSON","*.json"),("All files","*.*")])
        if not path:
            return
        try:
            self._load_json(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open JSON: {e}")

    def _load_json(self, path):
        with open(path, 'r', encoding='utf-8') as f:
            self.data = json.load(f)
        self.json_path = path
        self.edits = {}   # clear staged edits
        # populate xml parts
        xmls = [k for k in self.data.keys() if k.lower().endswith('.xml')]
        self.list_xml.delete(0, tk.END)
        for x in xmls:
            self.list_xml.insert(tk.END, x)
        self.list_place.delete(0, tk.END)
        self._set_status(f"Loaded {os.path.basename(path)}: {len(xmls)} XML parts")

    def _on_xml_select(self, evt):
        sel = self.list_xml.curselection()
        if not sel:
            return
        xml_name = self.list_xml.get(sel[0])
        entry = self.data.get(xml_name, {})
        self.list_place.delete(0, tk.END)
        if "_editable" in entry:
            for p in entry["_editable"].get("placeholders", []):
                self.list_place.insert(tk.END, p["placeholder"])
            self._set_status(f"{len(entry['_editable'].get('placeholders', []))} placeholders in {xml_name}")
        else:
            self._set_status(f"No placeholders in {xml_name}")

    def _on_placeholder_select(self, evt):
        sels = self.list_place.curselection()
        if not sels:
            self.selected_label.config(text="")
            return
        items = [self.list_place.get(i) for i in sels]
        self.selected_label.config(text=", ".join(items))
        default = items[0].strip("{}")
        self.entry_new.delete("1.0", tk.END)
        self.entry_new.insert("1.0", default)

    def _apply_to_selection(self):
        sel_xml = self.list_xml.curselection()
        if not sel_xml:
            messagebox.showwarning("No XML", "Select XML part first")
            return
        xml_name = self.list_xml.get(sel_xml[0])
        sel_place = self.list_place.curselection()
        if not sel_place:
            messagebox.showwarning("No placeholder", "Select at least one placeholder")
            return
        new_val = self.entry_new.get("1.0", tk.END).rstrip("\n")
        if xml_name not in self.edits:
            self.edits[xml_name] = {}
        for i in sel_place:
            ph = self.list_place.get(i)
            self.edits[xml_name][ph] = new_val
        self._update_staged()
        self._set_status(f"Staged {len(sel_place)} edits for {xml_name}")

    def _clear_selection_edits(self):
        sel_xml = self.list_xml.curselection()
        if not sel_xml:
            return
        xml_name = self.list_xml.get(sel_xml[0])
        sel_place = self.list_place.curselection()
        if not sel_place:
            return
        if xml_name in self.edits:
            for i in sel_place:
                ph = self.list_place.get(i)
                self.edits[xml_name].pop(ph, None)
            if not self.edits[xml_name]:
                self.edits.pop(xml_name, None)
        self._update_staged()
        self._set_status("Cleared selection edits")

    def _update_staged(self):
        self.staged.configure(state="normal")
        self.staged.delete("1.0", tk.END)
        for xml_name, mp in self.edits.items():
            self.staged.insert(tk.END, f"== {xml_name} ==\n")
            for ph, nv in mp.items():
                self.staged.insert(tk.END, f"{ph}  ->  {nv}\n")
            self.staged.insert(tk.END, "\n")
        self.staged.configure(state="disabled")

    def _save_json_with_edits(self):
        if not self.json_path or self.data is None:
            messagebox.showwarning("No JSON", "Open an editable JSON first.")
            return
        self.data["_edits"] = self.edits
        backup = self.json_path + ".bak"
        try:
            os.replace(self.json_path, backup)
        except Exception:
            # ignore
            pass
        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)
        messagebox.showinfo("Saved", f"Saved edits into {self.json_path} (backup: {backup})")
        self._set_status("Saved edits to JSON")

    def _reconstruct_docx(self):
        if self.json_path is None:
            messagebox.showwarning("No JSON", "Open an editable JSON first.")
            return
        # write _edits into the data object in memory before reconstruct
        self.data["_edits"] = self.edits
        # prompt user for output path
        out = filedialog.asksaveasfilename(title="Save reconstructed DOCX as...", defaultextension=".docx", filetypes=[("Word docx","*.docx")])
        if not out:
            return
        # attempt to use JsonToDocx helper if available
        if self.docx_reconstructor is None:
            messagebox.showwarning("Missing helper", "JsonToDocx helper not found in expected paths. Reconstruction disabled.")
            return
        try:
            # write the temporary JSON to disk so JsonToDocx can read it
            with open(self.json_path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
            # call helper: editable_json_to_docx(json_path, out, edits=self.edits)
            # JsonToDocx may be either a module with function editable_json_to_docx
            if hasattr(self.docx_reconstructor, "editable_json_to_docx"):
                self.docx_reconstructor.editable_json_to_docx(self.json_path, out, edits=self.edits)
            else:
                # Or module exposing class JsonToDocx as in Converters/JsonToDocx.py
                if hasattr(self.docx_reconstructor, "JsonToDocx"):
                    j2d = self.docx_reconstructor.JsonToDocx(self.json_path, os.path.dirname(out))
                    j2d.json_to_xml()
                    j2d.xml_to_docx(os.path.basename(out))
                    j2d.cleanup()
            messagebox.showinfo("Done", f"Reconstructed DOCX saved to:\n{out}")
            self._set_status(f"Reconstructed: {out}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to reconstruct DOCX: {e}")
            self._set_status("Reconstruction error")

    def _on_close(self):
        # write edits back (if configured) and then destroy UI
        if self.auto_save_on_close and self.json_path and self.data is not None:
            try:
                self.data["_edits"] = self.edits
                backup = self.json_path + ".bak"
                try:
                    os.replace(self.json_path, backup)
                except Exception:
                    pass
                with open(self.json_path, "w", encoding="utf-8") as f:
                    json.dump(self.data, f, indent=2, ensure_ascii=False)
                self._set_status("Saved edits to JSON")
            except Exception:
                pass
        self.root.destroy()
