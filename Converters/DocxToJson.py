import os, shutil, zipfile, json, xmltodict

class DocxToJson:
    def __init__(self, docx_path, output_dir="Output"):
        self.docx_path = docx_path
        self.output_dir = output_dir
        self.xml_dir = os.path.join(output_dir, "XML_Extracted")
        self.json_path = os.path.join(output_dir, "Document.json")

    def extract_docx(self):
        if os.path.exists(self.xml_dir):
            shutil.rmtree(self.xml_dir)
        os.makedirs(self.xml_dir)

        with zipfile.ZipFile(self.docx_path, 'r') as docx_zip:
            docx_zip.extractall(self.xml_dir)

        return self.xml_dir

    def convert_to_json(self):
        docx_structure = {}

        for root, _, files in os.walk(self.xml_dir):
            for file in files:
                file_path = os.path.join(root, file)
                rel_path = os.path.relpath(file_path, self.xml_dir)

                if not (file.endswith('.xml') or file.endswith('.rels')):
                    try:
                        with open(file_path, 'rb') as bin_file:
                            content = bin_file.read()
                        docx_structure[rel_path] = {"_binary": True, "content": content.hex()}
                    except Exception as e:
                        print(f"Warning: Could not read binary file {rel_path}: {e}")
                    continue

                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as xml_file:
                        xml_content = xml_file.read().strip()

                    if not xml_content:
                        docx_structure[rel_path] = {"_empty": True}
                        continue

                    docx_structure[rel_path] = xmltodict.parse(xml_content)

                except Exception as e:
                    print(f"Warning: Could not parse XML file {rel_path}: {e}")
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        docx_structure[rel_path] = {"_raw": True, "content": f.read()}

        with open(self.json_path, "w", encoding="utf-8") as json_file:
            json.dump(docx_structure, json_file, indent=2, ensure_ascii=False)

        return self.json_path

    def cleanup(self):
        """Delete intermediate XML files if DELETE_XML=true"""
        if os.getenv("DELETE_XML", "false").lower() == "true":
            if os.path.exists(self.xml_dir):
                shutil.rmtree(self.xml_dir)
                print(f"✅ Deleted intermediate XML folder: {self.xml_dir}")
