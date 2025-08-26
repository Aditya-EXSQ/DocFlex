import json
import tkinter as tk
from tkinter import ttk, messagebox

class JsonEditorUI:
    def __init__(self, json_path):
        self.json_path = json_path
        self.data = self._load_json()
        self.root = tk.Tk()
        self.root.title("DOCX DataBinding JSON Editor")

        self.entries = []  # keep references to Entry widgets and JSON nodes
        self._build_ui()

    def _load_json(self):
        with open(self.json_path, "r", encoding="utf-8") as f:
            return json.load(f)

    def _find_bindings(self, node, path=""):
        """
        Recursively search for w:sdt nodes with dataBinding.
        Return list of (alias/xpath, node_ref_to_wt_string).
        """
        results = []

        if isinstance(node, dict):
            if "w:sdt" in node:
                sdt_nodes = node["w:sdt"]
                if not isinstance(sdt_nodes, list):
                    sdt_nodes = [sdt_nodes]

                for sdt in sdt_nodes:
                    try:
                        # Check if this sdt has a dataBinding
                        if "w:sdtPr" in sdt and "w:dataBinding" in sdt["w:sdtPr"]:
                            alias = sdt["w:sdtPr"].get("w:alias", {}).get("@w:val", None)
                            xpath = sdt["w:sdtPr"]["w:dataBinding"].get("@w:xpath", "")
                            label = alias or xpath

                            # Find placeholder text (w:t)
                            wt_nodes = []
                            sdt_content = sdt.get("w:sdtContent", {})
                            if "w:r" in sdt_content:
                                runs = sdt_content["w:r"]
                                if not isinstance(runs, list):
                                    runs = [runs]
                                for run in runs:
                                    if "w:t" in run:
                                        wt_nodes.append(run)

                            for run in wt_nodes:
                                results.append((label, run))
                    except Exception as e:
                        print("Error parsing sdt:", e)

            # Recurse further
            for key, val in node.items():
                results.extend(self._find_bindings(val, path + "/" + key))

        elif isinstance(node, list):
            for item in node:
                results.extend(self._find_bindings(item, path))

        return results

    def _build_ui(self):
        bindings = self._find_bindings(self.data)

        if not bindings:
            tk.Label(self.root, text="No dataBinding fields found.").pack(padx=10, pady=10)
            return

        canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        frame = ttk.Frame(canvas)

        frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for idx, (label, run_node) in enumerate(bindings):
            text_val = run_node.get("w:t", "")

            lbl = ttk.Label(frame, text=f"{label}:")
            lbl.grid(row=idx, column=0, padx=5, pady=5, sticky="e")

            entry = ttk.Entry(frame, width=50)
            entry.insert(0, text_val)
            entry.grid(row=idx, column=1, padx=5, pady=5)

            self.entries.append((entry, run_node))

        save_btn = ttk.Button(self.root, text="Save Changes", command=self._save_changes)
        save_btn.pack(pady=10)

    def _save_changes(self):
        for entry, run_node in self.entries:
            new_text = entry.get()
            run_node["w:t"] = new_text

        with open(self.json_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

        messagebox.showinfo("Saved", f"Changes saved to {self.json_path}")

    def run(self):
        self.root.mainloop()
