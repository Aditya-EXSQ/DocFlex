import os, shutil, zipfile, json, xmltodict

class JsonToDocx:
    def __init__(self, json_path, output_dir="Output"):
        self.json_path = json_path
        self.output_dir = output_dir
        self.xml_dir = os.path.join(output_dir, "XML_Reconstructed")
        self.docx_path = os.path.join(output_dir, "Reconstructed.docx")

    def json_to_xml(self):
        if os.path.exists(self.xml_dir):
            shutil.rmtree(self.xml_dir)
        os.makedirs(self.xml_dir)

        with open(self.json_path, 'r', encoding='utf-8') as json_file:
            docx_structure = json.load(json_file)

        for rel_path, content_data in docx_structure.items():
            file_path = os.path.join(self.xml_dir, rel_path)
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            if isinstance(content_data, dict):
                if '_binary' in content_data:
                    with open(file_path, 'wb') as bin_file:
                        bin_file.write(bytes.fromhex(content_data['content']))
                elif '_raw' in content_data:
                    with open(file_path, 'w', encoding='utf-8') as text_file:
                        text_file.write(content_data['content'])
                elif '_empty' in content_data:
                    open(file_path, 'w').close()
                else:
                    try:
                        xml_content = xmltodict.unparse(content_data, pretty=True)
                        with open(file_path, 'w', encoding='utf-8') as xml_file:
                            xml_file.write(xml_content)
                    except Exception as e:
                        print(f"Warning: Could not unparse XML {rel_path}: {e}")
                        with open(file_path, 'w', encoding='utf-8') as f:
                            json.dump(content_data, f, indent=2, ensure_ascii=False)
            else:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(str(content_data))

        return self.xml_dir

    def xml_to_docx(self):
        if os.path.exists(self.docx_path):
            os.remove(self.docx_path)

        with zipfile.ZipFile(self.docx_path, 'w', zipfile.ZIP_DEFLATED) as docx_zip:
            for root, _, files in os.walk(self.xml_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    rel_path = os.path.relpath(file_path, self.xml_dir)
                    docx_zip.write(file_path, rel_path)

        return self.docx_path

    def cleanup(self):
        """Delete reconstructed XML files if DELETE_XML=true"""
        if os.getenv("DELETE_XML", "false").lower() == "true":
            if os.path.exists(self.xml_dir):
                shutil.rmtree(self.xml_dir)
                print(f"✅ Deleted intermediate XML folder: {self.xml_dir}")
